"""가격내림 사이트 시안 빌더: 비교 엑셀(하락 시트)에서 정적 페이지를 만든다.

사용법: python3 build.py <비교.xlsx> <수집일 YYYY-MM-DD> <출력 폴더(docs)> [사진주소.xlsx]
시안 단계라 입력은 비교 엑셀의 '하락' 시트. 운영 때는 같은 틀에 매일 데이터를 넣는다.
"""
import sys, re, json, hashlib, html, datetime, pathlib, openpyxl

CATS = ["생활용품", "라면·즉석식품", "생수·음료", "건강식품", "반려동물", "유아(기저귀+분유)"]
WD = "월화수목금토일"
DISCLOSURE = "이 포스팅은 쿠팡 파트너스 활동의 일환으로, 이에 따른 일정액의 수수료를 제공받습니다."
SITE = "장바규니규조대"
STYLE_V = "0"  # main()에서 style.css 내용 해시로 바꿔, 옛 CSS가 폰에 남지 않게 한다


def won(n):
    return f"{int(n):,}원"


def pkey(u):
    # 상품 ID + itemId(옵션)로 같은 상품을 찾는다
    m = re.search(r"/products/(\d+)", u or "")
    i = re.search(r"itemId=(\d+)", u or "")
    return (m.group(1) if m else "", i.group(1) if i else "")


def load_photos(xlsx):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    out = {}
    for ws in wb:
        if ws.title == "요약":
            continue
        for r in list(ws.iter_rows(values_only=True))[1:]:
            if r[8] and r[9]:
                out[pkey(r[8])] = r[9].replace("230x230ex", "492x492ex")  # 큰 크기도 열리는 것 확인(2026-10-19)
    return out


def load_links():
    # 산큐백이 파트너스에서 만든 링크: links.json {"상품ID:itemId": 파트너스 주소}. 없으면 쿠팡 원본 주소(시험용, 수익 안 됨)
    p = pathlib.Path(__file__).with_name("links.json")
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def load(xlsx, photos=None):
    ws = openpyxl.load_workbook(xlsx, data_only=True)["하락"]
    rows = list(ws.iter_rows(values_only=True))[1:]
    links = load_links()
    items = {c: [] for c in CATS}
    for r in rows:
        cat = r[2].split(" ", 1)[1]
        items[cat].append(dict(name=r[1], old=r[3], new=r[4], pct=r[6], reviews=r[8], url=links.get(":".join(pkey(r[9])), r[9]), photo=(photos or {}).get(pkey(r[9]))))
    for c in CATS:
        items[c].sort(key=lambda x: x["pct"])  # 더 많이 내린 순(음수가 작을수록 위)
    return items


def card(it, rank, hero=False):
    e = html.escape
    cls = "card hero" if hero else "card"
    photo = (f'<img src="{e(it["photo"])}" alt="{e(it["name"])}" loading="lazy">' if it.get("photo")
             else "<span>상품 사진</span>")
    return f"""<article class="{cls}">
  <a class="photo" href="{e(it['url'])}" target="_blank" rel="sponsored noopener" aria-label="{e(it['name'])}">{photo}<b class="rank">{rank}</b></a>
  <div class="body">
    <h3>{e(it['name'])}</h3>
    <p class="price"><s>어제 {won(it['old'])}</s> <i>→</i> <strong>오늘 {won(it['new'])}</strong></p>
    <p class="meta"><span class="pct">{round(it['pct'])}%</span><span class="rev">리뷰 {it['reviews']:,}</span></p>
    <a class="btn" href="{e(it['url'])}" target="_blank" rel="sponsored noopener">쿠팡 현재 가격 보기</a>
  </div>
</article>"""


def section(i, cat, its):
    top = its[:5]
    out = [f'<section id="c{i}"><h2>{html.escape(cat)}</h2>']
    if not top:
        out.append('<p class="empty">오늘은 어제보다 내린 상품이 없어요.</p>')
    else:
        out.append(card(top[0], 1, hero=True))
        if len(top) > 1:
            out.append('<div class="grid">' + "".join(card(t, n + 2) for n, t in enumerate(top[1:])) + "</div>")
    if len(its) > 5:
        lines = "".join(
            f"<li><span>{html.escape(t['name'])}</span><b>{round(t['pct'])}%</b></li>" for t in its[5:]
        )
        out.append(f'<details><summary>이 카테고리에서 내린 상품 전체 {len(its)}개</summary><ul>{lines}</ul></details>')
    out.append("</section>")
    return "".join(out)


def page(date, items, title):
    chips = "".join(
        f'<a href="#c{i}">{html.escape(c)}<small>{len(items[c])}</small></a>' for i, c in enumerate(CATS) if items[c]
    )
    secs = "".join(section(i, c, items[c]) for i, c in enumerate(CATS) if items[c])
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light"><title>{html.escape(title)}</title>
<link rel="stylesheet" href="../style.css?v={STYLE_V}"></head><body>
<header class="top"><a class="home" href="../index.html">{SITE}</a>
<h1>{html.escape(title)}</h1>
<p class="note">{date} 오전 7시대 수집 기준 · 와우회원 기준 · 가격은 바뀔 수 있어요</p></header>
<nav class="chips">{chips}</nav><main>{secs}</main>
<footer><p><a href="../index.html">홈으로</a></p></footer><div class="ad-bar">{DISCLOSURE}</div></body></html>"""


def short(d, t):
    # 홈 목록에서는 연도를 빼서 한 줄에 들어가게 한다 (날짜 페이지 제목은 그대로)
    return t.replace(d[:4] + "-", "", 1).replace("-", "/", 1)


def home(days):
    # days: (날짜, 제목, 내린 상품 수, 최대 하락률(%)) 최신이 앞
    li = "".join(
        f'<li><a href="{d}/index.html"><span>{"<em>NEW</em>" if i == 0 else ""}{short(d, t)}</span>'
        f'<small>{"최대 " + str(round(mx)) + "%" if n else "내린 상품 없음"}</small><b>›</b></a></li>'
        for i, (d, t, n, mx) in enumerate(days)
    )
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light"><title>{SITE}</title>
<link rel="stylesheet" href="style.css?v={STYLE_V}"></head><body>
<header class="top home-top"><div class="logo">로고</div><h1>{SITE}</h1></header>
<main><ul class="days">{li}</ul></main>
<footer></footer><div class="ad-bar">{DISCLOSURE}</div></body></html>"""


def main(xlsx, date, out, photo_xlsx=None):
    global STYLE_V
    out = pathlib.Path(out)
    css = out / "style.css"
    if css.exists():
        STYLE_V = hashlib.md5(css.read_bytes()).hexdigest()[:8]
    d = datetime.date.fromisoformat(date)
    title = f"{date}({WD[d.weekday()]}) 쿠팡 오늘의 가격내림 생필품"
    items = load(xlsx, load_photos(photo_xlsx) if photo_xlsx else None)
    (out / date).mkdir(parents=True, exist_ok=True)
    (out / date / "index.html").write_text(page(date, items, title), encoding="utf-8")
    n = sum(len(v) for v in items.values())
    mx = min(x["pct"] for v in items.values() for x in v)
    (out / "index.html").write_text(home([(date, title, n, mx)]), encoding="utf-8")


if __name__ == "__main__":
    main(*sys.argv[1:5])
