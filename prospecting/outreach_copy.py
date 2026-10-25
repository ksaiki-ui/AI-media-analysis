"""問い合わせフォーム・メール・架電で使う文面を、共通の型と会社ごとの段落から組み立てる。

文面は 5 つの塊でできている。

1. 宛名と名乗り（共通）
2. 会社ごとの段落（区分と計測した事実から作る。ここだけが会社ごとに変わる）
3. 業種ごとの一文（どんな読者の記事が問い合わせの入口になりやすいか）
4. 申し出と次の一歩（共通。無料の簡易診断と 30 分のオンライン）
5. LP と資料のリンク、順位を約束しない旨、配信停止の一文、署名（共通）

入れる事実は計測した公開情報（記事数・最後の公開月・直近 1 年の本数・事例ページ・過ぎた年号の題名）だけ。
広告や MA のツール名は、相手に「見張られている」と受け取られやすいので文面に入れない（架電のメモにだけ残す）。
"""

from __future__ import annotations

import re
from datetime import date
from urllib.parse import urlencode

SIGNATURE = "――――――――\n株式会社MIRAIKU\n{担当者名}\n{メール}\n{電話}"
OPT_OUT = "※今後のご連絡が不要でしたら、その旨ご返信ください。以後お送りしません。"
NO_PROMISE = "※検索順位をお約束するものではありません。"

# 業種ごとに「問い合わせの入口になりやすい記事の読者」
READER = {
    "情報通信（受託開発・SIer）": "発注先を探している企業の情報システム部門や事業部門の方が、費用感や進め方を調べるときに読む記事",
    "専門サービス（コンサル・士業）": "経営者や管理部門の方が、制度や手続きを調べる段階で読む記事",
    "事業サービス（人材・BPO）": "外注を検討している企業のご担当者が、費用や進め方を比べるときに読む記事",
    "人材派遣・紹介": "求職者向けではなく、人を採りたい企業の人事担当の方が読む記事",
    "教育・研修": "研修を企画する人事・教育担当の方が、テーマや進め方を調べるときに読む記事",
}

# 区分ごとの「どう役に立てるか」
OFFER = {
    "少量×更新中": "今の体制のまま、記事を出すペースを上げる方法をご提案できます。",
    "厚めHP×記事なし": "記事を始めるなら、どのテーマから書くと問い合わせにつながりやすいかをご提案できます。",
    "少量×停止": "止まっている記事づくりを、少ない手間で再開する方法をご提案できます。",
    "大量×停止": "これまでの記事を活かしながら、手直しと新しい記事を少ない手間で回す方法をご提案できます。",
}

SUBJECT = {
    "少量×更新中": "{社名}様の記事づくりの簡易診断（無料）のご案内",
    "厚めHP×記事なし": "{社名}様のホームページに記事を足す場合の簡易診断（無料）のご案内",
    "少量×停止": "{社名}様の記事の再開に向けた簡易診断（無料）のご案内",
    "大量×停止": "{社名}様の既存記事の活かし方の簡易診断（無料）のご案内",
}

# 題名の年号が過ぎていて手直しの余地が分かりやすいもの（お知らせ・休業案内などは使わない）
DATED_TITLE = re.compile(r"【[^】]*(?:19|20)\d{2}年[^】]*(?:版|最新)[^】]*】")


def _int(v) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def _ym(v) -> str:
    if not v:
        return ""
    d = v if isinstance(v, date) else date.fromisoformat(str(v)[:10])
    return f"{d.year}年{d.month}月"


def dated_example(v) -> str:
    """「過ぎた年号の記事」の列から、手直しの例に使える題名を 1 つ（30 字まで）"""
    for t in re.split(r" / |\n", str(v or "")):
        t = re.sub(r"^\d{4}[/.-]\d{1,2}[/.-]\d{1,2}\s*", "", t).strip()
        if DATED_TITLE.search(t):
            return t if len(t) <= 30 else t[:29] + "…"
    return ""


def custom_paragraph(d: dict) -> str:
    """会社ごとの段落。d は target_list の 1 行（列名は build_list.COLUMNS）"""
    seg = d["区分"]
    n = _int(d.get("記事数"))
    last = _ym(d.get("最終公開（推定）"))
    months = _int(d.get("更新なし日数")) // 30
    year = _int(d.get("直近1年の公開・更新数"))
    count = f"{n}本ほどの" if n else ""
    if seg == "少量×更新中":
        pace = f"直近1年で{year}本ほどのペースで" if year >= 2 else ""
        recent = f"（最近では{last}にも公開されていました）" if last else ""
        return (
            f"御社のホームページで、{count}記事を{pace}続けていらっしゃるのを拝見しました{recent}。"
            "少ない人数で続けていると、テーマ選びと下書きに時間を取られ、ペースを上げにくいというお話をよく伺います。"
        )
    if seg == "厚めHP×記事なし":
        cases = "や導入事例" if d.get("事例ページ") else ""
        return (
            f"御社のホームページを拝見し、サービスのご紹介{cases}まで丁寧に作り込まれている一方で、"
            "コラムやブログのような記事はまだお持ちでないようでした。"
            "今のホームページは御社をご存じの方には十分に伝わりますが、"
            "まだ御社を知らない方が検索から訪れる入口は少ない状態かと思います。"
        )
    stopped = f"{last}から{months}ヶ月ほど" if last and months else "しばらく"
    if seg == "少量×停止":
        return (
            f"御社のホームページで{count}記事を公開されたあと、{stopped}更新が止まっているようでした。"
            "始めてみたものの、通常の業務の合間では続けにくかった、というお話をよく伺います。"
        )
    ex = dated_example(d.get("過ぎた年号の記事"))
    dated = f"「{ex}」のように年号の入った記事もあり、" if ex else ""
    return (
        f"御社のホームページで{count}記事を積み上げてこられたのに、{stopped}公開が止まっているようでした。"
        f"{dated}これだけの蓄積があれば、新しく書く前に、既存の記事を今の情報に直すところから始める方法もあります。"
    )


def call_opener(d: dict, name: str) -> str:
    """架電で担当の方につながったときの一言目（30 秒で言える長さ）"""
    seg = d["区分"]
    n = _int(d.get("記事数"))
    last = _ym(d.get("最終公開（推定）"))
    months = _int(d.get("更新なし日数")) // 30
    head = f"{name}様のホームページを拝見してお電話しました。"
    if seg == "厚めHP×記事なし":
        body = "サービスのご紹介は充実されていますが、コラムのような記事はまだお持ちでないようでしたので。"
    elif seg == "少量×更新中":
        body = (
            f"{n}本ほど記事を続けていらっしゃって、{last}にも公開されていたので。"
            if n
            else "記事を続けていらっしゃるのを拝見したので。"
        )
    else:
        body = (
            f"{n}本ほど記事を出されていて、{last}から{months}ヶ月ほど止まっているようでしたので。"
            if n and last
            else "記事の更新が止まっているようでしたので。"
        )
    return (
        f"{head}{body}"
        "同じ領域の競合サイトが何のテーマで集客しているかを、1枚の簡易診断に無料でまとめています。"
        "お送りしてもよろしいですか。"
    )


# 架電の台本（区分ごとの問いかけと、よくある返しへの答え）
CALL_QUESTION = {
    "少量×更新中": "今は記事をどなたが書かれていますか。月に何本くらいのペースを目指されていますか。",
    "厚めHP×記事なし": "問い合わせは今、紹介や広告が中心ですか。検索から来る問い合わせを増やすお考えはありますか。",
    "少量×停止": "記事が止まったのは、書く人の時間が取れなかったからでしょうか。それとも成果が見えなかったからでしょうか。",
    "大量×停止": "以前はどなたが記事を担当されていましたか。いまある記事の手直しはどうされていますか。",
}
CALL_REBUTTAL = [
    (
        "間に合っています／不要です",
        "承知しました。診断は1枚の資料をお送りするだけなので、もし後でご関心が出たら、とだけお伝えさせてください。",
    ),
    (
        "担当がいない／分からない",
        "ホームページやマーケティングを見ている方はどなたになりますか。お名前だけ伺えれば、資料をお送りします。",
    ),
    (
        "AIの記事は品質が心配",
        "AIが下書きを作り、事実の点検と御社のご確認を通してから公開する流れです。診断では書き方より先に、狙うテーマをご覧いただけます。",
    ),
    (
        "費用は",
        "月5万円で記事5本からです。まずは無料の診断で、効果が見込めるかをご判断いただければと思います。",
    ),
    (
        "資料だけ送って",
        "承知しました。メールアドレスを伺えますか。診断と一緒にお送りします。",
    ),
]


def with_utm(url: str, no: int, seg: str, medium: str, campaign: str) -> str:
    """どの会社がリンクを開いたかを数えるため、LP の URL に番号を付ける"""
    if not url.startswith("http") or "claude.ai/" in url:
        return url  # 計測できない場所（共有前の Artifact など）には付けない
    q = urlencode(
        {
            "utm_source": "outreach",
            "utm_medium": medium,
            "utm_campaign": campaign,
            "utm_content": f"{no}-{seg}",
        }
    )
    return url + ("&" if "?" in url else "?") + q


def links_block(links: dict, no: int, seg: str, medium: str) -> str:
    camp = links.get("campaign", "outreach")
    lp = with_utm(links.get("lp", "【LPのURL】"), no, seg, medium, camp)
    deck = links.get("deck", "【サービス資料のURL】")
    return f"▼サービスの概要\n{lp}\n▼サービス資料（PDF）\n{deck}"


STAFFING = re.compile(
    r"人材|スタッフ|派遣|キャリア|ワーク|ジョブ|staff|career|work|job", re.IGNORECASE
)


def reader_line(d: dict, name: str, summary: str = "") -> str:
    """業種の一文。人材派遣協会の名簿には人材以外が本業の会社もいるので、社名や事業概要で人材と分かるときだけ使う"""
    field = d.get("領域", "")
    if field == "人材派遣・紹介" and not STAFFING.search(
        f"{name} {d.get('ドメイン', '')} {summary}"
    ):
        return ""
    reader = READER.get(field, "")
    return (
        f"御社の業種では、{reader}が問い合わせの入口になりやすいと考えています。"
        if reader
        else ""
    )


def form_message(
    d: dict,
    name: str,
    links: dict,
    no: int,
    sender: dict | None = None,
    summary: str = "",
) -> str:
    """問い合わせフォーム・メールの本文（目安 600 字前後）。summary は gBizINFO の事業概要"""
    seg = d["区分"]
    sig = SIGNATURE.format(
        **(
            sender
            or {"担当者名": "【担当者名】", "メール": "【メール】", "電話": "【電話】"}
        )
    )
    return (
        f"{name}\nご担当者様\n\n"
        "突然のご連絡失礼いたします。株式会社MIRAIKUで、BtoB企業のオウンドメディアの運用を支援しております。\n\n"
        f"{custom_paragraph(d)}\n\n"
        f"{reader_line(d, name, summary)}"
        "同じ領域の競合サイトがどのテーマで集客しているか、御社がまだ取れていないテーマはどこかを、"
        f"1枚の簡易診断にまとめて無料でお送りできます。{OFFER[seg]}\n"
        "ご興味があれば、このフォームへのご返信か下記までお知らせください。30分ほどのオンラインでご説明します。\n\n"
        f"{links_block(links, no, seg, 'form')}\n\n"
        f"{NO_PROMISE}\n{OPT_OUT}\n\n{sig}\n"
    )


def short_message(d: dict, name: str, links: dict, no: int) -> str:
    """文字数の上限が厳しいフォーム向け（目安 300 字以内。署名は送り主の欄で足りる前提）"""
    seg = d["区分"]
    first = custom_paragraph(d).split("。")[0] + "。"
    camp = links.get("campaign", "outreach")
    lp = with_utm(links.get("lp", "【LPのURL】"), no, seg, "form", camp)
    return (
        f"{name} ご担当者様\n株式会社MIRAIKUです。{first}"
        "同じ領域の競合サイトが集客しているテーマと、御社がまだ取れていないテーマを、1枚の簡易診断にまとめて無料でお送りできます。"
        f"ご興味があればご返信ください。\n概要: {lp}\n"
        "※順位はお約束しません。不要でしたらご返信ください。以後お送りしません。"
    )


def subject(d: dict, name: str) -> str:
    return SUBJECT[d["区分"]].format(社名=name)


def design_rows() -> list[list[str]]:
    """xlsx の「文面の設計」シートに載せる説明"""
    rows = [
        ["塊", "中身", "共通／会社ごと"],
        [
            "1. 宛名と名乗り",
            "社名・ご担当者様／株式会社MIRAIKUでBtoB企業のオウンドメディアの運用を支援",
            "共通",
        ],
        [
            "2. 会社ごとの段落",
            "区分と計測した事実（記事数・最後の公開月・直近1年の本数・事例ページ・過ぎた年号の題名）から作る",
            "会社ごと",
        ],
        [
            "3. 業種の一文",
            "その業種で問い合わせの入口になりやすい記事の読者",
            "業種ごと",
        ],
        [
            "4. 申し出",
            "競合サイトのテーマと御社の取れていないテーマの簡易診断（無料・1枚）＋区分ごとの提案の一文＋30分のオンライン",
            "共通＋区分ごと",
        ],
        [
            "5. リンクと注意書き",
            "LP（番号付きで誰が開いたか分かる）・サービス資料・順位を約束しない・配信停止・署名",
            "共通",
        ],
        [],
        ["区分", "会社ごとの段落の狙い", "提案の一文"],
    ]
    for seg, offer in OFFER.items():
        aim = {
            "少量×更新中": "続けている努力を認め、ペースを上げにくい悩みに寄せる",
            "厚めHP×記事なし": "HPの作り込みを認め、検索からの入口が無いことに気づいてもらう",
            "少量×停止": "続けにくかった事情に寄せ、再開の手間が小さいことを示す",
            "大量×停止": "蓄積を認め、新規より手直しから始められることを示す",
        }[seg]
        rows.append([seg, aim, offer])
    rows += [
        [],
        ["架電の問いかけ", "", ""],
        *[[seg, q, ""] for seg, q in CALL_QUESTION.items()],
        [],
        ["よくある返し", "答え方", ""],
        *[[k, v, ""] for k, v in CALL_REBUTTAL],
        [],
        [
            "書かないこと",
            "広告・MAのツール名（見張られている印象）、検索順位の約束、計測していない数字や評価",
            "",
        ],
    ]
    return rows
