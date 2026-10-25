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
    r"コンテンツマーケティング|オウンドメディア(制作|運用|支援)|記事制作|ライティング代行|"
    r"PR会社|ＰＲ会社|広報支援|マーケティング会社|デジタルマーケティング|Marketing"
)
# 個人向けの事業（BtoB の狙いから外れる）と、採用サイト（会社の HP ではない）
NOT_B2B = re.compile(
    r"学習塾|進学塾|個別指導|保育園|こども園|幼稚園|英会話スクール|教室|中途採用|新卒採用|採用サイト|RECRUIT|Recruit"
)

# 区分（HP の厚さ × 記事の量 × 更新の状態）。上ほど有望
PATTERNS = ("少量×更新中", "厚めHP×記事なし", "少量×停止", "大量×停止")
BASE_SCORE = {
    "少量×更新中": 40,
    "厚めHP×記事なし": 30,
    "少量×停止": 30,
    "大量×停止": 30,
}
MANY = 30  # 記事がこれ以上なら「大量」


def flag(v: str) -> bool:
    return str(v) == "True"


def thick_hp(r: dict) -> tuple[bool, str]:
    """HP が厚めか。ページ数・トップのリンク数・事例ページ・広告/MA のタグで決める"""
    pages = to_int(r.get("site_pages")) or 0
    links = to_int(r.get("home_links")) or 0
    tags = ",".join(t for t in (r.get("ad_tags", ""), r.get("ma_tags", "")) if t)
    why = []
    if tags:
        why.append(f"タグ:{tags}")
    if pages >= 30:
        why.append(f"{pages}ページ")
    elif links >= 40:
        why.append(f"トップのリンク{links}")
    if flag(r.get("has_cases")):
        why.append("事例あり")
    thick = (
        bool(tags)
        or pages >= 30
        or links >= 40
        or (flag(r.get("has_cases")) and (pages >= 15 or links >= 25))
    )
    return thick, "・".join(why)


def judge(r: dict) -> tuple[str, str, str]:
    """(区分, 優先度, 理由)。区分は PATTERNS か 大量×更新中 / 簡易HPのみ / 除外 / 要手確認"""
    dom = r["domain"]
    if any(dom == e or dom.endswith("." + e) for e in EXCLUDE):
        return "除外", "", "除外リスト（大手・ポータル・代理店など）"
    if SELF_SERVE.search(r["site_title"] + " " + r["site_description"]):
        return "除外", "", "SaaS・広告・制作・SEOの会社（自分で書ける側）"
    if NOT_B2B.search(r["site_title"] + " " + r["site_description"]):
        return "除外", "", "個人向けの事業、または採用サイト"
    if r["error"]:
        return "要手確認", "", r["error"]
    n = to_int(r["article_count"]) or 0
    days = to_int(r["days_since_last"])
    thick, hp_why = thick_hp(r)
    has_media = bool(r["media_path"]) and (n >= 3 or (n == 0 and days is not None))
    news = r.get("news_ratio")
    news_only = news not in (None, "") and float(news) >= 0.5
    if has_media and news_only:
        has_media = False  # 投稿はあるが、お知らせ・求人ばかりで記事ではない
    if not has_media:
        if thick:
            pattern = "厚めHP×記事なし"
            why = (
                "投稿はお知らせ・求人のみ" if news_only else "記事メディアなし"
            ) + f"・HPは厚め（{hp_why}）"
        else:
            return "簡易HPのみ", "", "記事メディアなし・HPも簡易"
    elif days is None:
        return "要手確認", "", "記事はあるが公開日を取得できず"
    else:
        many = n >= MANY
        live = days <= 90
        size = f"記事{n}本" if n else "記事数不明"
        if many and live:
            return "大量×更新中", "", f"{size}・最終{days}日前（体制あり）"
        pattern = ("大量" if many else "少量") + ("×更新中" if live else "×停止")
        why = f"{size}・最終{days}日前" + (f"・{hp_why}" if hp_why else "")
    score = BASE_SCORE[pattern]
    if r.get("ad_tags") or r.get("ma_tags"):
        score += 20
    if flag(r.get("has_cases")):
        score += 10
    if r["contact_url"]:
        score += 10
    prio = "高" if score >= 60 else ("中" if score >= 40 else "低")
    return pattern, prio, why


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
    head = f"{site_name(r)}様のホームページを拝見しました。"
    if pattern == "厚めHP×記事なし":
        return (
            f"{head}サービスや会社のご紹介は充実されていますが、コラムやブログのような記事は"
            "まだお持ちでないようでしたので、ご連絡しました。"
        )
    count = f"{n}本ほど記事を公開されていて、" if n else "記事を公開されていて、"
    if pattern.endswith("更新中"):
        return f"{head}{count}最近も{ym(r['last_date'])}に更新されているのを拝見し、ご連絡しました。"
    months = (to_int(r["days_since_last"]) or 0) // 30
    return f"{head}{count}最後の公開が{ym(r['last_date'])}で、{months}ヶ月ほど止まっているようでしたのでご連絡しました。"


OFFER = {
    "少量×更新中": "今の体制のまま、記事を出すペースを上げる方法をご提案できます。",
    "厚めHP×記事なし": "記事を始めるなら、どのテーマから書くと問い合わせにつながりやすいかをご提案できます。",
    "少量×停止": "止まっている記事づくりを、少ない手間で再開する方法をご提案できます。",
    "大量×停止": "これまでの記事を活かしながら、少ない手間で更新を再開する方法をご提案できます。",
}


def form_text(r: dict, pattern: str) -> str:
    return (
        f"{site_name(r)}\nご担当者様\n\n"
        "突然のご連絡失礼いたします。株式会社MIRAIKUでオウンドメディアの運用支援をしております。\n\n"
        f"{hook(r, pattern)}\n\n"
        "御社と同じ領域で、競合サイトがどのテーマで集客しているか、御社がまだ取れていないテーマは"
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
    "区分",
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
    "HPのページ数",
    "広告・MAタグ",
    "事例ページ",
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
                "区分": pattern,
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
                "HPのページ数": to_int(r.get("site_pages"))
                or to_int(r.get("home_links")),
                "広告・MAタグ": ",".join(
                    t for t in (r.get("ad_tags", ""), r.get("ma_tags", "")) if t
                ),
                "事例ページ": "あり" if flag(r.get("has_cases")) else "",
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
    order = {p: i for i, p in enumerate(PATTERNS)} | {
        "要手確認": 4,
        "大量×更新中": 5,
        "簡易HPのみ": 6,
        "除外": 7,
    }
    prank = {"高": 0, "中": 1, "低": 2, "": 3}
    out.sort(
        key=lambda x: (
            prank[x["優先度"]] if x["区分"] in PATTERNS else 9,
            order[x["区分"]],
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
        "区分": 14,
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
    s.append(["区分", "件数", "優先度・高", "優先度・中", "優先度・低"])
    for k in order:
        rows_k = [o for o in out if o["区分"] == k]
        s.append(
            [k, len(rows_k)]
            + [sum(1 for o in rows_k if o["優先度"] == p) for p in ("高", "中", "低")]
        )
    s.append([])
    s.append(["区分の基準"])
    s.append(
        [
            "少量×更新中",
            f"記事{MANY}本未満・最終公開90日以内。細々と続けている／始めたばかり",
        ]
    )
    s.append(
        [
            "厚めHP×記事なし",
            "記事メディアなし。HPは30ページ以上、広告・MAのタグ、事例ページのどれかあり（Web予算あり・未着手）",
        ]
    )
    s.append(
        ["少量×停止", f"記事{MANY}本未満・最終公開から91日以上。挑戦したが続かなかった"]
    )
    s.append(
        ["大量×停止", f"記事{MANY}本以上・最終公開から91日以上。過去に力を入れていた"]
    )
    s.append(
        ["大量×更新中（対象外）", f"記事{MANY}本以上・最終公開90日以内。体制・外注あり"]
    )
    s.append(["簡易HPのみ（対象外）", "記事メディアなし・HPも簡易"])
    s.append(
        [
            "優先度",
            "区分の基礎点（少量×更新中40、他30）＋広告/MAタグ20＋事例ページ10＋問い合わせ導線10。60以上=高、40以上=中",
        ]
    )
    s.column_dimensions["A"].width = 22
    s.column_dimensions["B"].width = 80
    wb.save(dst)
    counts = {k: sum(1 for o in out if o["区分"] == k) for k in order}
    print(dst, counts)


if __name__ == "__main__":
    main(*sys.argv[1:4])
