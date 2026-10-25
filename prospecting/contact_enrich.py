"""架電・フォーム送信の準備として、各社の電話番号と問い合わせフォームの中身を調べる。

使い方:
  python3 prospecting/contact_enrich.py targets.csv enriched.csv

入力の列: domain, contact_url（無ければ空）。出力は 1 社 1 行。途中で止まっても続きから調べる。
robots.txt に従い、1 社あたりトップ・問い合わせページなど数ページだけを読む。フォームは読むだけで送信しない。
"""

from __future__ import annotations

import csv
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass
from dataclasses import fields as dataclass_fields
from pathlib import Path
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup

from media_check import CONTACT_HINT, NO_SALES, Fetcher

PHONE = re.compile(
    r"0\d{1,4}[-‐－ー（(）)\s]{1,2}\d{1,4}[-‐－ー）)\s]{1,2}\d{3,4}|0120[-‐－]?\d{2,3}[-‐－]?\d{3}"
)
FAX_NEAR = re.compile(r"FAX|Fax|fax|ＦＡＸ|ファックス")
RECRUIT = re.compile(
    r"recruit|saiyo|saiyou|entry|career|job|staff|求人|採用|応募|エントリー|登録", re.I
)
CONTACT_PATHS = [
    "/contact/",
    "/inquiry/",
    "/toiawase/",
    "/otoiawase/",
    "/contact.html",
    "/form/",
]
CAPTCHA = {
    "reCAPTCHA": re.compile(r"google\.com/recaptcha|grecaptcha|g-recaptcha"),
    "hCaptcha": re.compile(r"hcaptcha\.com|h-captcha"),
    "Turnstile": re.compile(r"challenges\.cloudflare\.com/turnstile|cf-turnstile"),
}
EXTERNAL_FORMS = {
    "Googleフォーム": re.compile(r"docs\.google\.com/forms|forms\.gle"),
    "HubSpot": re.compile(r"js\.hsforms\.net|hsforms\.com|hs-form"),
    "formrun": re.compile(r"form\.run|formrun"),
    "Tayori": re.compile(r"tayori\.com"),
    "kintone": re.compile(r"kintoneapp\.com|form\.kintoneapp"),
    "Marketo": re.compile(r"mktoForm"),
    "Pardot": re.compile(r"go\.pardot\.com|pi\.pardot"),
    "SATORI": re.compile(r"satori\.marketing|satori\.segs"),
    "Microsoft Forms": re.compile(r"forms\.office\.com"),
}
FIELD_KIND = [
    ("会社名", re.compile(r"会社|社名|企業|法人|company|organization|corp", re.I)),
    ("氏名", re.compile(r"氏名|お名前|名前|ご担当|担当者|name", re.I)),
    ("フリガナ", re.compile(r"フリガナ|ふりがな|カナ|kana", re.I)),
    ("メール", re.compile(r"メール|mail", re.I)),
    ("電話", re.compile(r"電話|tel|phone", re.I)),
    ("部署", re.compile(r"部署|役職|department|position", re.I)),
    ("住所", re.compile(r"住所|郵便|address|zip|pref", re.I)),
    (
        "種別",
        re.compile(r"種別|種類|項目|区分|内容を選|ご用件|category|subject|type", re.I),
    ),
    (
        "本文",
        re.compile(r"内容|本文|ご質問|ご相談|message|inquiry|comment|body|詳細", re.I),
    ),
]


@dataclass
class Contact:
    domain: str
    phone: str = ""
    phone_source: str = ""
    contact_url: str = ""
    form_found: bool = False
    form_kind: str = ""  # 自社フォーム / 外部（HubSpot など） / メールのみ / 見つからず
    form_fields: str = ""  # 入力欄の種類
    required_count: int = 0
    has_select: bool = False  # 用件の選択肢がある（送信時に選ぶ必要）
    captcha: str = ""
    confirm_step: bool = False  # 確認画面を挟みそう
    no_sales_notice: bool = False
    privacy_consent: bool = False  # 同意チェックがある
    mail: str = ""
    error: str = ""


def pick_phone(soup: BeautifulSoup, text: str) -> str:
    for a in soup.find_all("a", href=True):
        if a["href"].startswith("tel:"):
            num = re.sub(r"[^\d]", "", a["href"][4:])
            if 10 <= len(num) <= 11:
                return num
    for m in PHONE.finditer(text):
        before = text[max(0, m.start() - 12) : m.start()]
        if FAX_NEAR.search(before):
            continue
        num = re.sub(r"[^\d]", "", m.group(0))
        if 10 <= len(num) <= 11:
            return num
    return ""


def fmt_phone(num: str) -> str:
    if num.startswith(("0120", "0800")):
        return f"{num[:4]}-{num[4:7]}-{num[7:]}"
    if num.startswith(("03", "06")) and len(num) == 10:
        return f"{num[:2]}-{num[2:6]}-{num[6:]}"
    if num.startswith(("090", "080", "070", "050")):
        return f"{num[:3]}-{num[3:7]}-{num[7:]}"
    return f"{num[:4]}-{num[4:6]}-{num[6:]}" if len(num) == 10 else num


def analyze_form(html: str, soup: BeautifulSoup, c: Contact) -> None:
    c.captcha = ",".join(k for k, rx in CAPTCHA.items() if rx.search(html))
    ext = [k for k, rx in EXTERNAL_FORMS.items() if rx.search(html)]
    forms = [
        f
        for f in soup.find_all("form")
        if f.find("textarea") or len(f.find_all(["input", "select"])) >= 4
    ]
    forms = [
        f
        for f in forms
        if not (f.get("role") == "search" or "search" in (f.get("class") or []))
    ]
    if forms:
        f = max(forms, key=lambda x: len(x.find_all(["input", "textarea", "select"])))
        c.form_found = True
        c.form_kind = "外部（" + ",".join(ext) + "）" if ext else "自社フォーム"
        kinds, req = [], 0
        for el in f.find_all(["input", "textarea", "select"]):
            if el.name == "input" and el.get("type") in (
                "hidden",
                "submit",
                "button",
                "image",
            ):
                continue
            label = ""
            if el.get("id"):
                lab = f.find("label", attrs={"for": el["id"]})
                label = lab.get_text(" ", strip=True) if lab else ""
            key = " ".join([el.get("name", ""), el.get("placeholder", ""), label])
            for kind, rx in FIELD_KIND:
                if rx.search(key):
                    kinds.append(kind)
                    break
            if el.has_attr("required") or el.get("aria-required") == "true":
                req += 1
            if el.name == "select" or (
                el.name == "input" and el.get("type") == "radio"
            ):
                c.has_select = True
            if (
                el.name == "input"
                and el.get("type") == "checkbox"
                and re.search(
                    r"同意|privacy|個人情報|agree", key + f.get_text(" ")[:0], re.I
                )
            ):
                c.privacy_consent = True
        c.form_fields = ",".join(dict.fromkeys(kinds))
        c.required_count = req
        ftext = f.get_text(" ", strip=True)
        c.confirm_step = bool(re.search(r"確認画面|確認する|入力内容を確認", ftext))
        if not c.privacy_consent:
            c.privacy_consent = bool(re.search(r"同意", ftext)) and bool(
                f.find("input", attrs={"type": "checkbox"})
            )
    elif ext:
        c.form_found = True
        c.form_kind = "外部（" + ",".join(ext) + "）"
    elif soup.find("iframe", src=True):
        srcs = " ".join(i["src"] for i in soup.find_all("iframe", src=True))
        ext = [k for k, rx in EXTERNAL_FORMS.items() if rx.search(srcs)]
        if ext:
            c.form_found = True
            c.form_kind = "外部（" + ",".join(ext) + "）"
    if not c.form_found:
        m = re.search(r"mailto:([\w.+-]+@[\w.-]+)", html)
        if m:
            c.mail = m.group(1)
            c.form_kind = "メールのみ"


def enrich(row: dict) -> Contact:
    dom = row["domain"]
    c = Contact(domain=dom)
    base = "https://" + dom
    try:
        f = Fetcher(base)
        home = f.get(base + "/")
        if home is None:
            base = "https://www." + dom
            f = Fetcher(base)
            home = f.get(base + "/")
        home_soup = BeautifulSoup(home.content, "lxml") if home is not None else None
        if home_soup is not None:
            c.phone = pick_phone(home_soup, home_soup.get_text(" "))
            c.phone_source = "トップ" if c.phone else ""
        url = row.get("contact_url") or ""
        if RECRUIT.search(url):
            url = ""  # 求職者向けの応募・登録フォームは使わない
        if not url and home_soup is not None:
            for a in home_soup.find_all("a", href=True):
                label = a.get_text(" ", strip=True) + " " + a["href"]
                if (
                    CONTACT_HINT.search(label)
                    and not RECRUIT.search(label)
                    and not a["href"].startswith(("mailto:", "tel:", "#"))
                ):
                    url = urljoin(base, a["href"])
                    break
        tried = [url] if url else []
        tried += [urljoin(base, p) for p in CONTACT_PATHS if urljoin(base, p) != url]
        for u in tried[:4]:
            r = f.get(u)
            if r is None or "html" not in r.headers.get("Content-Type", "html"):
                continue
            soup = BeautifulSoup(r.content, "lxml")
            html = r.text
            c.contact_url = u
            analyze_form(html, soup, c)
            text = soup.get_text(" ")
            c.no_sales_notice = bool(NO_SALES.search(text))
            if not c.phone:
                c.phone = pick_phone(soup, text)
                c.phone_source = "問い合わせページ" if c.phone else ""
            if c.form_found:
                break
        if not c.form_found and not c.form_kind:
            c.form_kind = "見つからず"
        c.phone = fmt_phone(c.phone) if c.phone else ""
    except Exception as e:  # 1社の失敗で全体を止めない
        c.error = f"{type(e).__name__}: {e}"[:200]
    return c


def main(src: str, dst: str) -> None:
    rows = list(csv.DictReader(open(src, encoding="utf-8")))
    done: set[str] = set()
    if Path(dst).exists():
        done = {r["domain"] for r in csv.DictReader(open(dst, encoding="utf-8"))}
    todo = [r for r in rows if r["domain"] not in done]
    fields = [x.name for x in dataclass_fields(Contact)]
    with (
        open(dst, "a", newline="", encoding="utf-8") as fh,
        ThreadPoolExecutor(16) as ex,
    ):
        w = csv.DictWriter(fh, fieldnames=fields)
        if not done:
            w.writeheader()
        for n, c in enumerate(ex.map(enrich, todo), 1):
            w.writerow(asdict(c))
            fh.flush()
            if n % 50 == 0:
                print(n, "/", len(todo), flush=True)
    print("done", len(todo))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
