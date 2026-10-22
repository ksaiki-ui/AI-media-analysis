"""計測結果を採点し、架電・フォーム用のターゲットリスト（xlsx）にする。

使い方:
  python3 prospecting/build_list.py out/measured.csv out/target_list.xlsx [既存リスト.xlsx]

既存リストを渡すと、そこに載っているドメインは除いて新しい分だけを出す（同じ会社に二重に当たらないため）。
文面に入れる事実は計測値だけ。推測で数字や評価を足さない。
"""

from __future__ import annotations

import csv
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
EXCLUDE = {
    line.strip().lower()
    for line in (HERE / "exclude_domains.txt").read_text(encoding="utf-8").splitlines()
    if line.strip() and not line.startswith("#")
}

STATUS = ["未着手", "架電済・不在", "架電済・断り", "フォーム送信済", "アポ獲得", "対象外"]


def to_int(v: str) -> int | None:
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def judge(r: dict) -> tuple[str, str]:
    """(ランク, 理由)。A=最優先 / B / C / 除外"""
    dom = r["domain"]
    if any(dom == e or dom.endswith("." + e) for e in EXCLUDE):
        return "除外", "除外リスト（大手・ポータル・代理店など）"
    if r["error"]:
        return "要手確認", r["error"]
    if not r["media_path"]:
        return "除外", "記事メディアを検出できず"
    days = to_int(r["days_since_last"])
    n = to_int(r["article_count"]) or 0
    if days is None:
        return "要手確認", "公開日を取得できず"
    if days < 60:
        return "除外", f"更新が活発（最終{days}日前）"
    if n and n < 5:
        return "除外", f"記事が{n}本と少ない"
    has_contact = bool(r["contact_url"])
    if days >= 180 and (n >= 10 or n == 0) and has_contact:
        return "A", f"{days}日更新なし・記事{n or '不明'}本"
    if days >= 90:
        return "B", f"{days}日更新なし" + ("" if has_contact else "・問い合わせ導線を検出できず")
    return "C", f"{days}日更新なし（やや最近）"


def ym(iso: str) -> str:
    d = date.fromisoformat(iso)
    return f"{d.year}年{d.month}月"


def site_name(r: dict) -> str:
    t = (r["site_title"] or r["domain"]).split("|")[0].split("｜")[0].split(" - ")[0].strip()
    return t.strip(" -–—|｜")[:40]


def hook(r: dict) -> str:
    n = to_int(r["article_count"]) or 0
    months = (to_int(r["days_since_last"]) or 0) // 30
    count = f"{n}本の記事を公開されていますが、" if n else ""
    return (
        f"{site_name(r)}様のホームページの記事を拝見しました。{count}最後の公開は{ym(r['last_date'])}で、"
        f"{months}ヶ月ほど更新が止まっているようでしたのでご連絡しました。"
    )


def form_text(r: dict) -> str:
    return (
        f"{site_name(r)}\nご担当者様\n\n"
        "突然のご連絡失礼いたします。株式会社MIRAIKUでオウンドメディアの運用支援をしております。\n\n"
        f"{hook(r)}\n\n"
        "御社のコラムと同じ領域で、競合サイトがどのテーマで集客しているか、御社がまだ取れていないテーマは"
        "どこかを、1枚の簡易診断にまとめてお送りできます。費用はかかりません。\n"
        "ご興味があれば、このフォームのご返信か下記までお知らせください。30分ほどのオンラインで"
        "内容をご説明します。\n\n"
        "※検索順位をお約束するものではありません。\n"
        "※今後のご連絡が不要でしたら、その旨ご返信ください。以後お送りしません。\n"
    )


def load_existing(path: str | None) -> set[str]:
    if not path or not Path(path).exists():
        return set()
    wb = load_workbook(path, read_only=True)
    out = set()
    for ws in wb.worksheets:
        header = None
        for row in ws.iter_rows(values_only=True):
            if header is None:
                header = list(row)
                continue
            if "ドメイン" in header:
                v = row[header.index("ドメイン")]
                if v:
                    out.add(str(v).lower())
    return out


COLUMNS = [
    "ランク", "理由", "領域", "サイト名", "ドメイン", "メディアURL", "記事数", "最終公開（推定）", "更新なし日数",
    "直近1年の公開・更新数", "日付の出どころ", "問い合わせURL", "フォーム送信", "架電の一言目", "フォーム文面",
    "見つけたキーワード", "計測日", "ステータス", "担当", "接触日", "結果メモ",
]


def main(src: str, dst: str, existing: str | None = None) -> None:
    rows = list(csv.DictReader(open(src, encoding="utf-8")))
    done = load_existing(existing)
    today = date.today().isoformat()
    out = []
    for r in rows:
        if r["domain"] in done:
            continue
        rank, why = judge(r)
        base = f"https://{r['domain']}"
        media = base + r["media_path"] if r["media_path"].startswith("/") and "*" not in r["media_path"] else base
        ok_target = rank in ("A", "B", "C")
        out.append({
            "ランク": rank,
            "理由": why,
            "領域": r["segment"],
            "サイト名": site_name(r),
            "ドメイン": r["domain"],
            "メディアURL": media,
            "記事数": to_int(r["article_count"]) or None,
            "最終公開（推定）": r["last_date"] or None,
            "更新なし日数": to_int(r["days_since_last"]),
            "直近1年の公開・更新数": to_int(r["posts_365d"]),
            "日付の出どころ": r["date_source"],
            "問い合わせURL": r["contact_url"],
            "フォーム送信": "不可（営業お断り）" if r["no_sales_notice"] == "True" else ("可" if r["contact_url"] else "導線なし"),
            "架電の一言目": hook(r) if ok_target else "",
            "フォーム文面": form_text(r) if ok_target and r["no_sales_notice"] != "True" else "",
            "見つけたキーワード": r["source_keyword"],
            "計測日": today,
            "ステータス": "未着手" if ok_target else "対象外",
            "担当": "",
            "接触日": "",
            "結果メモ": "",
        })
    order = {"A": 0, "B": 1, "C": 2, "要手確認": 3, "除外": 4}
    out.sort(key=lambda x: (order[x["ランク"]], -(x["更新なし日数"] or 0)))

    wb = Workbook()
    ws = wb.active
    ws.title = "リスト"
    ws.append(COLUMNS)
    for o in out:
        ws.append([o[c] for c in COLUMNS])
    head = PatternFill("solid", fgColor="1B3A2B")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head
        c.alignment = Alignment(vertical="center", wrap_text=True)
    widths = {"A": 8, "B": 26, "C": 12, "D": 28, "E": 24, "F": 34, "G": 8, "H": 12, "I": 10, "J": 10, "K": 12,
              "L": 34, "M": 14, "N": 60, "O": 70, "P": 22, "Q": 11, "R": 14, "S": 10, "T": 11, "U": 30}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=c.column_letter in ("B", "N", "O", "U"))
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUS) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"R2:R{max(2, ws.max_row)}")
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions

    s = wb.create_sheet("集計")
    s.append(["ランク", "件数"])
    for k in order:
        s.append([k, sum(1 for o in out if o["ランク"] == k)])
    s.append([])
    s.append(["ランクの基準"])
    s.append(["A", "更新なし180日以上・記事10本以上（不明含む）・問い合わせ導線あり"])
    s.append(["B", "更新なし90日以上（A以外）"])
    s.append(["C", "更新なし60〜89日"])
    s.append(["除外", "更新が活発（60日未満）／記事メディアなし／記事5本未満／除外リスト"])
    s.append(["要手確認", "接続できない・日付が取れない。人が開いて確認する"])
    s.column_dimensions["B"].width = 70
    wb.save(dst)
    counts = {k: sum(1 for o in out if o["ランク"] == k) for k in order}
    print(dst, counts)


if __name__ == "__main__":
    main(*sys.argv[1:4])
