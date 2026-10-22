"""日本人材派遣協会（JASSA）の会員派遣会社検索から、会員の社名と HP の URL を集める。

使い方:
  python3 prospecting/sources/jassa.py prospecting/universe/jassa.csv

公開されている会員検索（robots.txt で許可された /search/）だけを、1 秒に 1 回の間隔で読む。
出力の列は media_check.py の seeds と同じ（url, segment, source_keyword）に社名・所在地を足したもの。
"""

from __future__ import annotations

import csv
import re
import sys
import time

import requests
from bs4 import BeautifulSoup

UA = "AIMediaKun-MediaCheck/0.1 (MIRAIKU; public-page survey, respects robots.txt)"
BASE = "https://www.jassa.or.jp/search/list/"
SOURCE = "名簿:日本人材派遣協会"


def main(dst: str) -> None:
    s = requests.Session()
    s.headers["User-Agent"] = UA
    r = s.get(BASE, params={"mode": "search", "syllabary": "11"}, timeout=30)
    pages = int(re.search(r'name="page_max" value="(\d+)"', r.text).group(1))
    ids = list(dict.fromkeys(re.findall(r"companyId=(\d+)", r.text)))
    for pg in range(2, pages + 1):
        time.sleep(1)
        r = s.get(BASE, params={"mode": "paged", "now_page": pg}, timeout=30)
        ids += [i for i in dict.fromkeys(re.findall(r"companyId=(\d+)", r.text)) if i not in ids]
    print(f"{len(ids)} 社", file=sys.stderr)

    with open(dst, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["url", "segment", "source_keyword", "company", "address"])
        for n, cid in enumerate(ids, 1):
            time.sleep(1)
            try:
                r = s.get(BASE + "page/", params={"companyId": cid}, timeout=30)
            except requests.RequestException:
                continue
            main_el = BeautifulSoup(r.text, "lxml").find("main")
            if main_el is None:
                continue
            urls = [
                a["href"]
                for a in main_el.find_all("a", href=True)
                if a["href"].startswith("http") and "jassa.or.jp" not in a["href"]
                and "smp.ne.jp" not in a["href"] and "youtube.com" not in a["href"]
            ]
            text = main_el.get_text(" ", strip=True)
            name = re.search(r"会社名\s*(\S+)", text)
            addr = re.search(r"(?:所在地|住所)\s*(\S+\s?\S*)", text)
            if urls:
                w.writerow([urls[0], "人材派遣・紹介", SOURCE, name.group(1) if name else "", addr.group(1) if addr else ""])
            if n % 50 == 0:
                print(n, file=sys.stderr)
                fh.flush()


if __name__ == "__main__":
    main(sys.argv[1])
