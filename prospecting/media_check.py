"""見込み客のオウンドメディアを外から計測する。

入力: 候補CSV（列: url, segment, source_keyword）。url は記事URLでもトップでもよい。
出力: 計測結果CSV（1ドメイン1行）。

計測すること（公開情報だけ。robots.txt を守り、本文は保存しない）:
- メディアの場所（/column/ /blog/ など）と記事数
- 公開日の分布（最終更新日、直近90日・365日の公開本数）
- 問い合わせ導線の有無、問い合わせページの「営業お断り」表記
- サイト名・説明（業種の判定に使う）

使い方:
  python3 prospecting/media_check.py seeds/2026-10-22.csv out/measured.csv
"""

from __future__ import annotations

import csv
import gzip
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timezone
from urllib import robotparser
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

UA = "AIMediaKun-MediaCheck/0.1 (MIRAIKU; public-page survey, respects robots.txt)"
TIMEOUT = 15
PER_DOMAIN_INTERVAL = 1.0  # 同一ドメインへの間隔（秒）
MAX_REQUESTS_PER_DOMAIN = 30
MAX_SITEMAPS = 15

# メディアらしいパスの区切り。news・topics（お知らせ）は記事メディアとして数えない
MEDIA_SEGMENTS = [
    "column",
    "columns",
    "blog",
    "media",
    "knowledge",
    "useful",
    "tips",
    "magazine",
    "journal",
    "library",
    "oyakudachi",
    "otoku",
    "contents",
    "article",
    "articles",
    "archives",
    "learn",
    "guide",
    "howto",
    "colum",
    "kiji",
    "jirei-column",
]
# この下の階層はメディアとみなさない（IR資料・採用・会社案内など）
NOT_MEDIA_TOP = {"ir", "recruit", "company", "corporate", "about"}
CONTACT_HINT = re.compile(
    r"contact|inquiry|toiawase|otoiawase|お問い?合わせ|お問合せ|ご相談|資料請求",
    re.IGNORECASE,
)
NO_SALES = re.compile(
    r"営業(目的|活動|メール|のご連絡|電話|・勧誘|・セールス|や勧誘)?[^。]{0,30}(お断り|ご遠慮|控え|禁止|お受けしておりません|対応いたしかね)"
)
DATE_PATTERNS = [
    re.compile(r"(20\d{2})[./\-年](\d{1,2})[./\-月](\d{1,2})"),
]


@dataclass
class Result:
    domain: str
    seed_url: str
    segment: str
    source_keyword: str
    site_title: str = ""
    site_description: str = ""
    media_path: str = ""
    article_count: int = 0
    date_source: str = ""  # sitemap / rss / index_page / none
    first_date: str = ""
    last_date: str = ""
    days_since_last: int | None = None
    posts_90d: int = 0
    posts_365d: int = 0
    contact_url: str = ""
    no_sales_notice: bool = False
    robots_blocked: bool = False
    error: str = ""
    requests_used: int = 0
    notes: list[str] = field(default_factory=list)


class Fetcher:
    """ドメイン単位で間隔と回数を守る取得器"""

    def __init__(self, base: str):
        self.base = base
        self.session = requests.Session()
        self.session.headers["User-Agent"] = UA
        self.count = 0
        self.conn_errors = 0
        self.last = 0.0
        self.lock = threading.Lock()
        self.rp = robotparser.RobotFileParser()
        self.robots_text = ""
        try:
            r = self._raw(urljoin(base, "/robots.txt"))
            if r is not None and r.status_code == 200:
                self.robots_text = r.text
                self.rp.parse(r.text.splitlines())
            else:
                self.rp.parse([])
        except Exception:
            self.rp.parse([])

    def _raw(self, url: str):
        with self.lock:
            if self.count >= MAX_REQUESTS_PER_DOMAIN:
                return None
            wait = PER_DOMAIN_INTERVAL - (time.time() - self.last)
            if wait > 0:
                time.sleep(wait)
            self.count += 1
            self.last = time.time()
        return self.session.get(url, timeout=TIMEOUT, allow_redirects=True)

    def get(self, url: str):
        if not self.rp.can_fetch(UA, url):
            return None
        r = None
        for attempt in range(2):
            try:
                r = self._raw(url)
                break
            except requests.RequestException:
                self.conn_errors += 1
                time.sleep(2 * (attempt + 1))
        if r is None:
            return None
        if r is None or r.status_code != 200:
            return None
        return r

    def sitemaps_from_robots(self) -> list[str]:
        return re.findall(r"(?im)^\s*sitemap:\s*(\S+)", self.robots_text)


def parse_date(s: str) -> date | None:
    d = _parse_date(s)
    # 1970-01-01 などの壊れた日付は捨てる
    return d if d and d.year >= 2000 else None


def _parse_date(s: str) -> date | None:
    s = (s or "").strip()
    if not s:
        return None
    for fmt in (
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M%z",
        "%Y-%m-%d",
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S %Z",
    ):
        try:
            return datetime.strptime(
                s.replace("Z", "+00:00") if "%z" in fmt else s, fmt
            ).date()
        except ValueError:
            continue
    m = DATE_PATTERNS[0].search(s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def read_xml(resp) -> BeautifulSoup:
    content = resp.content
    if resp.url.endswith(".gz") or content[:2] == b"\x1f\x8b":
        try:
            content = gzip.decompress(content)
        except OSError:
            pass
    return BeautifulSoup(content, "xml")


def collect_sitemap_entries(
    f: Fetcher, base: str, seed_media: str = ""
) -> list[tuple[str, date | None, str]]:
    """sitemap を辿って (url, lastmod) を集める。記事系の子 sitemap を優先する。
    seed_media があれば、その下に別に置かれた sitemap（/media/ だけ別の WordPress など）も見る"""
    queue = f.sitemaps_from_robots() or [
        urljoin(base, "/sitemap.xml"),
        urljoin(base, "/sitemap_index.xml"),
        urljoin(base, "/wp-sitemap.xml"),
    ]
    if seed_media:
        queue = [
            urljoin(base, seed_media + "sitemap.xml"),
            urljoin(base, seed_media + "sitemap_index.xml"),
        ] + queue
    seen, entries, visited = set(), [], 0
    while queue and visited < MAX_SITEMAPS:
        sm = queue.pop(0)
        if sm in seen:
            continue
        seen.add(sm)
        r = f.get(sm)
        visited += 1
        if r is None:
            continue
        soup = read_xml(r)
        children = [
            loc.get_text(strip=True)
            for s in soup.find_all("sitemap")
            if (loc := s.find("loc"))
        ]
        if children:
            # 投稿・コラム系を先に、固定ページ・タクソノミー・画像は後回し
            def prio(u: str) -> int:
                lu = u.lower()
                if any(k in lu for k in ("post", "column", "blog", "article", "media")):
                    return 0
                if any(
                    k in lu
                    for k in (
                        "page",
                        "category",
                        "tag",
                        "taxonomy",
                        "image",
                        "author",
                        "user",
                    )
                ):
                    return 2
                return 1

            queue.extend(sorted(children, key=prio))
            continue
        for u in soup.find_all("url"):
            loc = u.find("loc")
            if not loc:
                continue
            lm = u.find("lastmod")
            entries.append(
                (
                    loc.get_text(strip=True),
                    parse_date(lm.get_text()) if lm else None,
                    sm,
                )
            )
    return entries


def seed_media_path(seed_url: str) -> str:
    """検索で見つかったページの URL にあるメディアのパス（例: /media/column/3502.html → /media/）"""
    segs = [s for s in urlparse(seed_url).path.split("/") if s]
    for i, s in enumerate(segs[:-1]):
        if s.lower() in NOT_MEDIA_TOP:
            return ""
        if s.lower() in MEDIA_SEGMENTS:
            return "/" + "/".join(segs[: i + 1]) + "/"
    return ""


def find_media_path(urls: list[str], seed_url: str) -> str:
    """記事が最も多いメディアらしいパスの先頭（例: /column/）"""
    counts: dict[str, int] = {}
    for u in urls:
        segs = [s for s in urlparse(u).path.split("/") if s]
        if len(segs) < 2:
            continue
        for i, s in enumerate(segs[:-1]):
            if s.lower() in NOT_MEDIA_TOP:
                break
            if s.lower() in MEDIA_SEGMENTS:
                key = "/" + "/".join(segs[: i + 1]) + "/"
                counts[key] = counts.get(key, 0) + 1
                break
    # 検索で見つかったページのメディアを優先する（古いブログと今のコラムが並ぶサイトで古い方を拾わないため）
    key = seed_media_path(seed_url)
    if key and counts.get(key, 0) >= 3:
        return key
    if counts:
        return max(counts, key=counts.get)
    # /column029.html のようにファイル名で並ぶ型
    flat = [
        u
        for u in urls
        if re.search(
            r"/(column|blog|colum)[\w-]*\.html?$", urlparse(u).path, re.IGNORECASE
        )
    ]
    if len(flat) >= 3:
        return "/column*.html"
    segs = [s for s in urlparse(seed_url).path.split("/") if s]
    for i, s in enumerate(segs[:-1]):
        if s.lower() in MEDIA_SEGMENTS:
            return "/" + "/".join(segs[: i + 1]) + "/"
    return ""


def dates_from_rss(f: Fetcher, base: str, media_path: str) -> list[date]:
    for path in ([media_path + "feed/"] if media_path else []) + [
        "/feed/",
        "/rss/",
        "/feed",
    ]:
        r = f.get(urljoin(base, path))
        if (
            r is None
            or b"<rss" not in r.content[:2000]
            and b"<feed" not in r.content[:2000]
        ):
            continue
        soup = read_xml(r)
        ds = [
            parse_date(t.get_text())
            for t in soup.find_all(["pubDate", "published", "updated"])
        ]
        ds = [d for d in ds if d]
        if ds:
            return ds
    return []


def dates_from_page(f: Fetcher, url: str) -> list[date]:
    r = f.get(url)
    if r is None:
        return []
    text = BeautifulSoup(r.content, "lxml").get_text(" ")
    out = []
    for m in DATE_PATTERNS[0].finditer(text):
        try:
            d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            continue
        if date(2010, 1, 1) <= d <= date.today():
            out.append(d)
    return out


def check_contact(
    f: Fetcher, base: str, home_html: BeautifulSoup | None, res: Result
) -> None:
    link = None
    if home_html is not None:
        for a in home_html.find_all("a", href=True):
            label = a.get_text(" ", strip=True) + " " + a["href"]
            if CONTACT_HINT.search(label) and not a["href"].startswith(
                ("mailto:", "tel:", "#")
            ):
                link = urljoin(base, a["href"])
                if urlparse(link).netloc == urlparse(base).netloc:
                    break
    if not link:
        return
    res.contact_url = link
    r = f.get(link)
    if r is not None:
        text = BeautifulSoup(r.content, "lxml").get_text(" ")
        res.no_sales_notice = bool(NO_SALES.search(text))


def measure(row: dict) -> Result:
    seed = row["url"].strip()
    p = urlparse(seed if "://" in seed else "https://" + seed)
    base = f"{p.scheme}://{p.netloc}"
    res = Result(
        domain=p.netloc.lower().removeprefix("www."),
        seed_url=seed,
        segment=row.get("segment", ""),
        source_keyword=row.get("source_keyword", ""),
    )
    try:
        f = Fetcher(base)
        if not f.rp.can_fetch(UA, base + "/"):
            res.robots_blocked = True
            res.error = "robots.txt で全体が拒否"
            return res
        home = f.get(base + "/")
        home_soup = BeautifulSoup(home.content, "lxml") if home is not None else None
        if home_soup is not None:
            t = home_soup.find("meta", property="og:site_name") or home_soup.find(
                "title"
            )
            res.site_title = (
                t.get("content")
                if t and t.name == "meta"
                else (t.get_text(strip=True) if t else "")
            )[:120]
            d = home_soup.find("meta", attrs={"name": "description"})
            res.site_description = (d.get("content", "") if d else "")[:200]

        entries = collect_sitemap_entries(f, base, seed_media_path(seed))
        urls = [u for u, _, _ in entries]
        res.media_path = find_media_path(urls, seed)
        if res.media_path == "/column*.html":
            in_media = [
                (u, d)
                for u, d, _ in entries
                if re.search(
                    r"/(column|blog|colum)[\w-]*\.html?$",
                    urlparse(u).path,
                    re.IGNORECASE,
                )
            ]
        elif res.media_path:
            in_media = [
                (u, d)
                for u, d, _ in entries
                if urlparse(u).path.startswith(res.media_path)
                and urlparse(u).path.rstrip("/") != res.media_path.rstrip("/")
                and "/page/" not in u
                and "/category/" not in u
                and "/tag/" not in u
            ]
        else:
            # WordPress でパスに区切りが無い投稿（/記事スラッグ/）は投稿用 sitemap から数える
            in_media = [
                (u, d)
                for u, d, sm in entries
                if re.search(r"post-sitemap|posts-post|sitemap-posts|sitemap_post", sm)
            ]
            if in_media:
                res.media_path = "（投稿）"
        res.article_count = len(in_media)
        dates = [d for _, d in in_media if d]
        # lastmod が一斉更新されている（同じ日付に偏る）ときは信用しない
        if dates:
            top = max(dates.count(d) for d in set(dates))
            if len(dates) >= 10 and top / len(dates) > 0.5:
                res.notes.append("sitemapの日付が一斉更新のため不採用")
                dates = []
        rss = dates_from_rss(
            f, base, res.media_path if res.media_path.startswith("/") else ""
        )
        if rss:
            # RSS は直近の公開日なので最終更新日はこちらを優先し、分布は sitemap と合わせる
            dates = dates + rss
            res.date_source = "rss+sitemap" if len(dates) > len(rss) else "rss"
        elif dates:
            res.date_source = "sitemap"
            res.notes.append(
                "sitemapのlastmodは更新日のため、公開日より新しく出ることがある"
            )
        elif res.media_path.startswith("/"):
            dates = dates_from_page(f, urljoin(base, res.media_path))
            res.date_source = "index_page" if dates else "none"
        else:
            res.date_source = "none"
        if not res.article_count and res.date_source in ("rss", "index_page"):
            res.notes.append("記事数はsitemapから取れず不明")
        if dates:
            today = date.today()
            res.first_date = min(dates).isoformat()
            last = max(dates)
            res.last_date = last.isoformat()
            res.days_since_last = (today - last).days
            res.posts_90d = sum(1 for d in dates if (today - d).days <= 90)
            res.posts_365d = sum(1 for d in dates if (today - d).days <= 365)
        check_contact(f, base, home_soup, res)
        res.requests_used = f.count
        if home is None and f.conn_errors:
            res.error = "接続できず（相手側または経路の制限。手で確認）"
    except Exception as e:  # 1社の失敗で全体を止めない
        res.error = f"{type(e).__name__}: {e}"[:200]
    return res


def main(src: str, dst: str) -> None:
    with open(src, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    # 同一ドメインは最初の1件だけ
    uniq, seen = [], set()
    for r in rows:
        d = (
            urlparse(r["url"] if "://" in r["url"] else "https://" + r["url"])
            .netloc.lower()
            .removeprefix("www.")
        )
        if d and d not in seen:
            seen.add(d)
            uniq.append(r)
    with ThreadPoolExecutor(max_workers=8) as ex:
        results = list(ex.map(measure, uniq))
    fields = list(asdict(results[0]).keys()) if results else []
    with open(dst, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields)
        w.writeheader()
        for r in results:
            d = asdict(r)
            d["notes"] = " / ".join(d["notes"])
            w.writerow(d)
    print(
        f"{len(results)} domains measured at {datetime.now(timezone.utc).isoformat()}"
    )


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
