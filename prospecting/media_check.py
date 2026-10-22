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
from dataclasses import fields as dataclass_fields
from datetime import date, datetime, timezone
from pathlib import Path
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
AD_TAGS = {
    "Google広告": re.compile(r"AW-\d{6,}|googleadservices|googlesyndication"),
    "Yahoo広告": re.compile(r"yjads|s\.yimg\.jp/images/listing/tool/cv|ytag\("),
    "Meta広告": re.compile(r"connect\.facebook\.net/[^\"']*/fbevents|fbq\("),
    "LINE広告": re.compile(r"tr\.line\.me|_lt\("),
}
MA_TAGS = {
    "HubSpot": re.compile(r"js\.hs-scripts\.com|js\.hsforms\.net|hs-analytics"),
    "Marketo": re.compile(r"munchkin|mktoForms"),
    "Pardot": re.compile(r"pi\.pardot\.com|pardot"),
    "SATORI": re.compile(r"satori\.marketing|satori\.segs"),
    "b→dash": re.compile(r"bdash-cloud|b-dash"),
    "List Finder": re.compile(r"list-finder"),
    "SHANON": re.compile(r"shanon-services"),
}
CASES_HINT = re.compile(
    r"(case|cases|jirei|works|achievement|voice)(/|\.html|$)|導入事例|お客様の声|実績",
    re.IGNORECASE,
)
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
    stale_2y: int = (
        0  # sitemap の lastmod が 2 年以上前の記事数（リライトされていない記事）
    )
    old_year_titles: int = 0  # 「2024年版」のような過ぎた年号が入った記事タイトルの数
    old_year_examples: str = ""
    news_ratio: float | None = (
        None  # RSS の題名のうちお知らせ・求人らしいものの割合（記事でない投稿の見分け）
    )
    site_pages: int = 0  # sitemap に載っているページ数（HP の厚さ）
    home_links: int = (
        0  # トップページからの同じサイト内へのリンク数（sitemap が無いときの厚さ）
    )
    has_cases: bool = False  # 導入事例・実績のページがある
    ad_tags: str = ""  # 広告のタグ（Web 予算がある証拠）
    ma_tags: str = ""  # MA・CRM のタグ
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


MEDIA_LABEL = re.compile(
    r"コラム|ブログ|お役立ち|ナレッジ|オウンドメディア|基礎知識|用語集|ノウハウ|blog|column",
    re.IGNORECASE,
)


def home_media_link(soup: BeautifulSoup, base: str) -> str:
    """トップページのリンクから、記事メディアらしい URL を 1 つ選ぶ"""
    host = urlparse(base).netloc.removeprefix("www.")
    votes: dict[str, int] = {}
    for a in soup.find_all("a", href=True):
        u = urljoin(base, a["href"]).split("#")[0]
        pu = urlparse(u)
        if pu.netloc.removeprefix("www.") != host:
            continue
        segs = [x for x in pu.path.split("/") if x]
        if not segs or segs[0].lower() in NOT_MEDIA_TOP:
            continue
        label = a.get_text(" ", strip=True)
        if segs[0].lower() in MEDIA_SEGMENTS or (
            MEDIA_LABEL.search(label) and len(label) <= 20
        ):
            key = f"{pu.scheme}://{pu.netloc}/{segs[0]}/x"
            votes[key] = votes.get(key, 0) + 1
    return max(votes, key=votes.get) if votes else ""


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


OLD_YEAR = re.compile(
    r"[【\[（(]\s*(20[12]\d)年(?:度)?(?:版|最新版?|最新情報)?\s*[】\]）)]"
    r"|(20[12]\d)年(?:度)?(?:最新版?|版)(?!との|から|の改|と比)"
)


NEWS_TITLE = re.compile(
    r"お知らせ|急募|時給|日給|月給|【長期】|【短期】|勤務地|募集中|採用情報|休業|年末年始|夏季|営業日|営業時間|掲載|開催|出展|プレスリリース|"
    r"就任|移転|リニューアル|更新しました|公開しました|開設|受賞|取得しました|のご案内|メンテナンス"
)


def media_titles(f: Fetcher, base: str, media_path: str) -> tuple[list[str], list[str]]:
    """メディアの一覧ページと RSS から (全タイトル, RSS のタイトル) を集める"""
    titles: list[str] = []
    rss_titles: list[str] = []
    urls = [urljoin(base, media_path)] if media_path.startswith("/") else []
    urls.append(
        urljoin(base, (media_path if media_path.startswith("/") else "/") + "feed/")
    )
    if media_path.startswith("/"):
        urls.append(urljoin(base, "/feed/"))
    for u in urls:
        if u == urljoin(base, "/feed/") and len(urls) > 2 and rss_titles:
            continue  # メディアの RSS が取れたらサイト全体の RSS は見ない
        r = f.get(u)
        if r is None:
            continue
        soup = (
            read_xml(r)
            if b"<rss" in r.content[:2000]
            else BeautifulSoup(r.content, "lxml")
        )
        tags = (
            soup.find_all("title")
            if b"<rss" in r.content[:2000]
            else soup.find_all(["a", "h2", "h3"])
        )
        is_rss = b"<rss" in r.content[:2000]
        for t in tags[1:] if is_rss else tags:  # RSS の先頭はサイト名
            txt = t.get_text(" ", strip=True)
            if 10 <= len(txt) <= 80:
                titles.append(txt)
            if is_rss and txt:
                rss_titles.append(txt)
    return titles, rss_titles


def old_year_titles(titles: list[str]) -> list[str]:
    """過ぎた年号の入った記事タイトルを拾う"""
    this_year = date.today().year
    out = []
    for t in dict.fromkeys(titles):
        m = OLD_YEAR.search(t)
        # 「【2024年版】」「2023年度最新」など、今年より前の年号を見出しに掲げた記事
        if m and int(m.group(1) or m.group(2)) < this_year:
            out.append(t)
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
                (t.get("content") or "")
                if t and t.name == "meta"
                else (t.get_text(strip=True) if t else "")
            )[:120]
            d = home_soup.find("meta", attrs={"name": "description"})
            res.site_description = (d.get("content", "") if d else "")[:200]
            html = home.text
            res.ad_tags = ",".join(k for k, rx in AD_TAGS.items() if rx.search(html))
            res.ma_tags = ",".join(k for k, rx in MA_TAGS.items() if rx.search(html))
            host = urlparse(base).netloc.removeprefix("www.")
            inner = {
                urljoin(base, a["href"]).split("#")[0]
                for a in home_soup.find_all("a", href=True)
                if urlparse(urljoin(base, a["href"])).netloc.removeprefix("www.")
                == host
            }
            res.home_links = len(inner)
            # 会社のトップ URL しか分からないときは、メニューのブログ・コラムへのリンクを手がかりにする
            if not seed_media_path(seed):
                nav = home_media_link(home_soup, base)
                if nav:
                    seed = nav
            res.has_cases = any(
                CASES_HINT.search(a.get_text(" ", strip=True) + " " + a["href"])
                for a in home_soup.find_all("a", href=True)
            )

        entries = collect_sitemap_entries(f, base, seed_media_path(seed))
        urls = [u for u, _, _ in entries]
        res.site_pages = len(set(urls))
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
        sitemap_dates = list(dates)
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
            # 本数は記事ごとに 1 つの日付（sitemap）で数え、RSS と二重に数えない
            per_article = sitemap_dates or dates
            res.posts_90d = sum(1 for d in per_article if (today - d).days <= 90)
            res.posts_365d = sum(1 for d in per_article if (today - d).days <= 365)
            res.stale_2y = sum(1 for d in sitemap_dates if (today - d).days > 730)
        titles, rss_titles = media_titles(f, base, res.media_path)
        if rss_titles:
            res.news_ratio = round(
                sum(1 for t in rss_titles if NEWS_TITLE.search(t)) / len(rss_titles), 2
            )
        olds = old_year_titles(titles)
        res.old_year_titles = len(olds)
        res.old_year_examples = " / ".join(olds[:3])[:200]
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
    # 途中で止まっても続きから測れるよう、出力済みのドメインは飛ばし、1 社ごとに追記する
    done: set[str] = set()
    if Path(dst).exists():
        with open(dst, newline="", encoding="utf-8") as fh:
            done = {r["domain"] for r in csv.DictReader(fh)}
    # 同一ドメインは最初の1件だけ
    uniq, seen = [], set(done)
    for r in rows:
        d = (
            urlparse(r["url"] if "://" in r["url"] else "https://" + r["url"])
            .netloc.lower()
            .removeprefix("www.")
        )
        if d and d not in seen:
            seen.add(d)
            uniq.append(r)
    fields = [f.name for f in dataclass_fields(Result)]
    new_file = not done
    n = 0
    with (
        open(dst, "a", newline="", encoding="utf-8") as fh,
        ThreadPoolExecutor(max_workers=16) as ex,
    ):
        w = csv.DictWriter(fh, fieldnames=fields)
        if new_file:
            w.writeheader()
        for res in ex.map(measure, uniq):
            d = asdict(res)
            d["notes"] = " / ".join(d["notes"])
            w.writerow(d)
            fh.flush()
            n += 1
            if n % 100 == 0:
                print(n, "/", len(uniq), flush=True)
    print(f"{n} domains measured at {datetime.now(timezone.utc).isoformat()}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
