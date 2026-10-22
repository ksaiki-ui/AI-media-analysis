"""計測結果を採点し、架電・フォーム用のターゲットリスト（xlsx）にする。

使い方:
  python3 prospecting/build_list.py out/measured.csv out/target_list.xlsx [既存リスト.xlsx]

既存リストを渡すと、そこに載っているドメインは除いて新しい分だけを出す（同じ会社に二重に当たらないため）。
文面に入れる事実は計測値だけ。推測で数字や評価を足さない。
"""

from __future__ import annotations

import csv
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
EXCLUDE = {
    line.strip().lower()
    for line in (HERE / "exclude_domains.txt").read_text(encoding="utf-8").splitlines()
    if line.strip() and not line.startswith("#")
}

STATUS = [
    "未着手",
    "架電済・不在",
    "架電済・断り",
    "フォーム送信済",
    "アポ獲得",
    "対象外",
]


def to_int(v: str) -> int | None:
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


# 自分で書ける側・競合になる側（SaaS ベンダー、広告・制作・SEO 会社）。タイトルと説明文で判定する
SELF_SERVE = re.compile(
    r"SaaS|クラウドサービス|マーケティング支援|SEO対策|SEOコンサル|広告代理|Web制作|ホームページ制作|"
    r"コンテンツマーケティング|オウンドメディア(制作|運用|支援)|記事制作|ライティング代行"
)

PATTERNS = ("停止", "リライト遅れ", "低頻度で継続")


def judge(r: dict) -> tuple[str, str, str]:
    """(パターン, 優先度, 理由)。パターンは 停止 / リライト遅れ / 低頻度で継続 / 除外 / 要手確認"""
    dom = r["domain"]
    if any(dom == e or dom.endswith("." + e) for e in EXCLUDE):
        return "除外", "", "除外リスト（大手・ポータル・代理店など）"
    if SELF_SERVE.search(r["site_title"] + " " + r["site_description"]):
        return "除外", "", "SaaS・広告・制作・SEOの会社（自分で書ける側）"
    if r["error"]:
        return "要手確認", "", r["error"]
    if not r["media_path"]:
        return "除外", "", "記事メディアを検出できず"
    days = to_int(r["days_since_last"])
    n = to_int(r["article_count"]) or 0
    y365 = to_int(r["posts_365d"]) or 0
    y90 = to_int(r["posts_90d"]) or 0
    stale = to_int(r.get("stale_2y")) or 0
    olds = to_int(r.get("old_year_titles")) or 0
    if days is None:
        return "要手確認", "", "公開日を取得できず"
    if n and n < 5:
        return "除外", "", f"記事が{n}本と少ない"
    has_contact = bool(r["contact_url"])
    prio = "高" if has_contact and (n >= 10 or n == 0) else "中"
    if days >= 180:
        return "停止", prio, f"{days}日更新なし・記事{n or '不明'}本"
    # 書く体制がある会社は外す（sitemap の更新日も数えるので、リライトの多い会社もここで外れる）
    if y90 >= 6:
        return "除外", "", f"更新が活発（直近90日で{y90}本）"
    rewrite = olds >= 2 or (n >= 30 and stale / n >= 0.5)
    if rewrite:
        why = []
        if olds >= 2:
            why.append(f"過ぎた年号の記事{olds}本")
        if n >= 30 and stale / n >= 0.5:
            why.append(f"2年以上手つかず{stale}/{n}本")
        return "リライト遅れ", prio, "・".join(why)
    if days <= 90 and 3 <= y365 <= 36:
        return "低頻度で継続", prio, f"直近1年で{y365}本・最終{days}日前"
    if days > 90:
        return "停止", "中", f"{days}日更新なし（停滞気味）"
    return "除外", "", f"直近1年で{y365}本（低頻度の範囲外）"


def ym(iso: str) -> str:
    d = date.fromisoformat(iso)
    return f"{d.year}年{d.month}月"


def site_name(r: dict) -> str:
    t = (
        (r["site_title"] or r["domain"])
        .split("|")[0]
        .split("｜")[0]
        .split(" - ")[0]
        .strip()
    )
    return t.strip(" -–—|｜")[:40]


def hook(r: dict, pattern: str) -> str:
    n = to_int(r["article_count"]) or 0
    head = f"{site_name(r)}様のホームページの記事を拝見しました。"
    if pattern == "停止":
        months = (to_int(r["days_since_last"]) or 0) // 30
        count = f"{n}本の記事を公開されていますが、" if n else ""
        return (
            f"{head}{count}最後の公開は{ym(r['last_date'])}で、"
            f"{months}ヶ月ほど更新が止まっているようでしたのでご連絡しました。"
        )
    if pattern == "リライト遅れ":
        olds = to_int(r.get("old_year_titles")) or 0
        stale = to_int(r.get("stale_2y")) or 0
        if olds >= 2:
            ex = r.get("old_year_examples", "").split(" / ")[0]
            return f"{head}「{ex}」のように、年号の入った記事がそのままになっているようでしたのでご連絡しました。"
        return f"{head}{n}本の記事のうち{stale}本ほどが2年以上更新されていないようでしたのでご連絡しました。"
    y365 = to_int(r["posts_365d"]) or 0
    return f"{head}直近1年で{y365}本ほど、継続して記事を出されているようでしたのでご連絡しました。"


OFFER = {
    "停止": "止まっている記事づくりを、少ない手間で再開する方法をご提案できます。",
    "リライト遅れ": "既存の記事のうち、どれから直すと効果が出やすいかの優先順位をご提案できます。",
    "低頻度で継続": "今の体制のまま、記事を出すペースを上げる方法をご提案できます。",
}


def form_text(r: dict, pattern: str) -> str:
    return (
        f"{site_name(r)}\nご担当者様\n\n"
        "突然のご連絡失礼いたします。株式会社MIRAIKUでオウンドメディアの運用支援をしております。\n\n"
        f"{hook(r, pattern)}\n\n"
        "御社の記事と同じ領域で、競合サイトがどのテーマで集客しているか、御社がまだ取れていないテーマは"
        f"どこかを、1枚の簡易診断にまとめてお送りできます。費用はかかりません。{OFFER[pattern]}\n"
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
    "パターン",
    "優先度",
    "理由",
    "領域",
    "サイト名",
    "ドメイン",
    "メディアURL",
    "記事数",
    "最終公開（推定）",
    "更新なし日数",
    "直近1年の公開・更新数",
    "2年以上手つかずの記事",
    "過ぎた年号の記事",
    "日付の出どころ",
    "問い合わせURL",
    "フォーム送信",
    "架電の一言目",
    "フォーム文面",
    "見つけたキーワード",
    "計測日",
    "ステータス",
    "担当",
    "接触日",
    "結果メモ",
]


def main(src: str, dst: str, existing: str | None = None) -> None:
    rows = list(csv.DictReader(open(src, encoding="utf-8")))
    done = load_existing(existing)
    today = date.today().isoformat()
    out = []
    for r in rows:
        if r["domain"] in done:
            continue
        pattern, prio, why = judge(r)
        base = f"https://{r['domain']}"
        media = (
            base + r["media_path"]
            if r["media_path"].startswith("/") and "*" not in r["media_path"]
            else base
        )
        ok_target = pattern in PATTERNS
        out.append(
            {
                "パターン": pattern,
                "優先度": prio,
                "理由": why,
                "領域": r["segment"],
                "サイト名": site_name(r),
                "ドメイン": r["domain"],
                "メディアURL": media,
                "記事数": to_int(r["article_count"]) or None,
                "最終公開（推定）": r["last_date"] or None,
                "更新なし日数": to_int(r["days_since_last"]),
                "直近1年の公開・更新数": to_int(r["posts_365d"]),
                "2年以上手つかずの記事": to_int(r.get("stale_2y")),
                "過ぎた年号の記事": r.get("old_year_examples", ""),
                "日付の出どころ": r["date_source"],
                "問い合わせURL": r["contact_url"],
                "フォーム送信": "不可（営業お断り）"
                if r["no_sales_notice"] == "True"
                else ("可" if r["contact_url"] else "導線なし"),
                "架電の一言目": hook(r, pattern) if ok_target else "",
                "フォーム文面": form_text(r, pattern)
                if ok_target and r["no_sales_notice"] != "True"
                else "",
                "見つけたキーワード": r["source_keyword"],
                "計測日": today,
                "ステータス": "未着手" if ok_target else "対象外",
                "担当": "",
                "接触日": "",
                "結果メモ": "",
            }
        )
    order = {"低頻度で継続": 0, "リライト遅れ": 1, "停止": 2, "要手確認": 3, "除外": 4}
    out.sort(
        key=lambda x: (
            order[x["パターン"]],
            x["優先度"] != "高",
            -(x["更新なし日数"] or 0),
        )
    )

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
    widths = {
        "パターン": 12,
        "優先度": 6,
        "理由": 26,
        "領域": 12,
        "サイト名": 28,
        "ドメイン": 24,
        "メディアURL": 34,
        "問い合わせURL": 34,
        "架電の一言目": 60,
        "フォーム文面": 70,
        "過ぎた年号の記事": 40,
        "見つけたキーワード": 22,
        "ステータス": 14,
        "結果メモ": 30,
    }
    letter = {c: get_column_letter(i + 1) for i, c in enumerate(COLUMNS)}
    for c in COLUMNS:
        ws.column_dimensions[letter[c]].width = widths.get(c, 11)
    wrap = {
        letter[c]
        for c in (
            "理由",
            "架電の一言目",
            "フォーム文面",
            "過ぎた年号の記事",
            "結果メモ",
        )
    }
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=c.column_letter in wrap)
    dv = DataValidation(
        type="list", formula1='"' + ",".join(STATUS) + '"', allow_blank=True
    )
    ws.add_data_validation(dv)
    col = letter["ステータス"]
    dv.add(f"{col}2:{col}{max(2, ws.max_row)}")
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = ws.dimensions

    s = wb.create_sheet("集計")
    s.append(["パターン", "件数", "うち優先度・高"])
    for k in order:
        s.append(
            [
                k,
                sum(1 for o in out if o["パターン"] == k),
                sum(1 for o in out if o["パターン"] == k and o["優先度"] == "高"),
            ]
        )
    s.append([])
    s.append(["パターンの基準（上から順に当てはめる）"])
    s.append(
        ["停止", "最終公開から180日以上（90〜179日は「停滞気味」として優先度・中）"]
    )
    s.append(
        [
            "リライト遅れ",
            "「2024年版」など過ぎた年号の記事が2本以上、または記事30本以上で半分以上が2年以上手つかず",
        ]
    )
    s.append(
        [
            "低頻度で継続",
            "最終公開が90日以内で、直近1年の公開・更新が3〜36本（直近90日は5本まで）",
        ]
    )
    s.append(["優先度・高", "問い合わせ導線あり、かつ記事10本以上（本数不明を含む）"])
    s.append(
        [
            "除外",
            "除外リスト／SaaS・広告・制作・SEOの会社／記事メディアなし／記事5本未満／直近90日で6本以上（活発）",
        ]
    )
    s.append(["要手確認", "接続できない・日付が取れない。人が開いて確認する"])
    s.column_dimensions["A"].width = 18
    s.column_dimensions["B"].width = 80
    wb.save(dst)
    counts = {k: sum(1 for o in out if o["パターン"] == k) for k in order}
    print(dst, counts)


if __name__ == "__main__":
    main(*sys.argv[1:4])
