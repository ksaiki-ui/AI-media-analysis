"""gBizINFO の法人基本情報（一括ダウンロードの CSV）から、狙う業種・規模の会社と HP の URL を取り出す。

使い方:
  python3 prospecting/sources/gbizinfo.py Kihonjoho_UTF-8.zip prospecting/universe/gbizinfo.csv

一括ダウンロードには gBizINFO のアクセストークンが要る（https://info.gbiz.go.jp/hojin/DownloadTop）。
トークンも ZIP（約 300MB）もリポジトリには置かない。

選び方:
- 閉鎖していない、Web サイトの URL がある、従業員 10〜300 人
- 株式会社・有限会社・合同会社・士業法人（学校法人・社会福祉法人・医療法人などは外す）
- 事業種目（日本標準産業分類の大分類）が G 情報通信 / L 専門・技術サービス / O 教育・学習支援、
  または R サービス業のうち社名・事業概要が人材・BPO・研修・コンサル系のもの
- 事業概要が広告・Web 制作・SEO・SaaS 中心の会社は外す（自分で書ける側・競合）
"""

from __future__ import annotations

import csv
import io
import re
import sys
import zipfile

R_KW = re.compile(
    r"人材|派遣|紹介|アウトソーシング|BPO|コール|コンタクト|研修|コンサル|事務代行|経理|給与|労務|データ入力|翻訳|リサーチ|調査"
)
SELF_SERVE = re.compile(
    r"広告|Web制作|ウェブ制作|ホームページ制作|WEB制作|SEO|マーケティング支援|コンテンツ制作|"
    r"SaaS|クラウドサービスの提供|放送|新聞|出版|メディア事業"
)
NAME_OK = re.compile(r"株式会社|有限会社|合同会社|税理士法人|社会保険労務士法人|司法書士法人|行政書士法人|弁護士法人|監査法人")
NAME_NG = re.compile(r"学校法人|社会福祉法人|医療法人|財団法人|社団法人|協同組合|独立行政法人")


def segment(row: dict) -> str:
    majors = {x.split(":")[0] for x in row["事業種目"].split(",") if x}
    text = row["商号または名称"] + " " + row["事業概要"]
    if "G" in majors:
        return "情報通信（受託開発・SIer）"
    if "L" in majors:
        return "専門サービス（コンサル・士業）"
    if "O" in majors:
        return "教育・研修"
    if "R" in majors and R_KW.search(text):
        return "事業サービス（人材・BPO）"
    return ""


def main(src: str, dst: str) -> None:
    z = zipfile.ZipFile(src)
    reader = csv.DictReader(io.TextIOWrapper(z.open(z.namelist()[0]), encoding="utf-8-sig"))
    n = 0
    with open(dst, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["url", "segment", "source_keyword", "company", "address", "employees", "corporate_number"])
        for row in reader:
            if row["登記記録の閉鎖等年月日"] or not row["WebサイトURL"].startswith("http"):
                continue
            try:
                emp = int(row["従業員数"] or 0)
            except ValueError:
                continue
            name = row["商号または名称"]
            if not (10 <= emp <= 300) or not NAME_OK.search(name) or NAME_NG.search(name):
                continue
            seg = segment(row)
            if not seg or SELF_SERVE.search(row["事業概要"]):
                continue
            w.writerow([row["WebサイトURL"].strip(), seg, "名簿:gBizINFO", name, row["都道府県"] + row["市区町村（郡）"], emp, row["法人番号"]])
            n += 1
    print(n, "社")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
