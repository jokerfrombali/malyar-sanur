# -*- coding: utf-8 -*-
"""Статический сайт «Маляр на Бали» из книги SEO. Выход: docs/ (GitHub Pages)."""
import json, os, re, shutil, html
import openpyxl
from districts import DISTRICTS, GRID, guide

ROOT = os.path.dirname(os.path.abspath(__file__))
CFG = json.load(open(os.path.join(ROOT, "site_config.json"), encoding="utf-8"))
BOOK = os.path.join(ROOT, "SEO-малярные-работы-Бали-2026-09-25.xlsx")
OUT = os.path.join(ROOT, "docs")
B = CFG["base_path"].rstrip("/")
E = html.escape

wb = openpyxl.load_workbook(BOOK, read_only=True)
pages = {r[0]: dict(id=r[0], url=r[1], parent=r[2], type=r[3], h1=r[5], title=r[6], desc=r[8])
         for r in wb["03 Структура"].iter_rows(min_row=2, values_only=True)}
sem = list(wb["09 Семантика"].iter_rows(min_row=2, values_only=True))
QW = ("как ", "чем ", "почему ", "что ", "можно ли ", "сколько ", "какой ", "какая ", "какую ", "нужно ли ", "когда ")

# Тексты услуг: только проверяемое, без цен, сроков-обещаний и выдуманных кейсов
TXT = {
 "S01": ("Крашу стены и потолки в виллах, квартирах и домах на Бали: от одной комнаты до всего дома. Перед покраской проверяю основание на влагу и плесень — во влажном климате это главная причина, по которой свежая краска пузырится и отслаивается.",
         ["Осмотр и замер влажности стен", "Защита мебели и пола", "Очистка, обработка антисептиком при необходимости", "Шпаклёвка трещин и шлифовка", "Грунт и два слоя краски", "Уборка после работы"]),
 "S02": ("Покраска фасадов вилл у моря требует другого подхода: соль, солнце и дожди разрушают покрытие быстрее, чем в глубине острова. Подбираю фасадные краски под прибрежные условия и планирую работы на сухой сезон.",
         ["Осмотр трещин, отслоений и зелёного налёта", "Мойка и биоцидная обработка", "Ремонт трещин эластичными составами", "Грунт и фасадная краска", "Покраска цоколя, парапетов, металлических элементов"]),
 "S03": ("Не всегда нужно перекрашивать всю стену. Подкрашиваю следы после арендаторов, царапины, пятна после протечек и сколы — с подбором цвета, чтобы не было заметных пятен.",
         ["Подбор и колеровка цвета по образцу", "Локальная подготовка и грунт", "Подкраска с растушёвкой", "Честная рекомендация, если стену лучше перекрасить целиком"]),
 "S04": ("Плесень на Бали — частая проблема: тропическая влажность и сезон дождей. Удаляю плесень со стен, потолков, дерева и швов, нахожу причину и защищаю поверхность, чтобы она не вернулась через месяц.",
         ["Поиск источника влаги", "Механическая очистка", "Обработка фунгицидом", "Просушка", "Антигрибковый грунт и краска", "Рекомендации по вентиляции"]),
 "S05": ("Антигрибковая обработка — профилактика и обязательный этап перед покраской поражённых поверхностей: стены, потолки, деревянные балки, шкафы, мебель.",
         ["Выбор состава под поверхность", "Нанесение и выдержка", "Повторная обработка при глубоком поражении", "Финишное покрытие"]),
 "S06": ("Покрываю лаком двери, мебель, столешницы, лестницы, полы и уличное дерево. Во влажном воздухе лак ведёт себя иначе: мутнеет, дольше сохнет, отслаивается при ошибках подготовки — поэтому работаю по погоде и с промежуточной шлифовкой.",
         ["Снятие старого покрытия", "Шлифовка", "Грунт/порозаполнитель", "Нанесение лака в несколько слоёв", "Промежуточная шлифовка", "Лак для улицы с UV-защитой"]),
 "S07": ("Масло и пропитка — для тика, террас, уличной мебели и дерева, которое должно дышать. Защищают от влаги, солнца и насекомых и легко обновляются без полной шлифовки.",
         ["Очистка и шлифовка", "Антисептическая пропитка", "Масло или масло-воск в 2–3 слоя", "План обновления покрытия"]),
 "S08": ("Восстанавливаю деревянную мебель и двери: снимаю старое покрытие, убираю плесень и пятна, заделываю дефекты и заново покрываю маслом или лаком.",
         ["Оценка состояния", "Снятие старого покрытия", "Ремонт сколов и трещин", "Тонировка при необходимости", "Финишное покрытие"]),
 "S09": ("Деревянная терраса у бассейна или сада сереет и темнеет от солнца, воды и плесени. Шлифую, очищаю и покрываю маслом для террас.",
         ["Мойка и отбеливание дерева", "Шлифовка", "Обработка от плесени", "Масло для террас", "Рекомендации по уходу"]),
 "S10": ("У моря металл ржавеет быстро. Крашу ворота, заборы, перила и решётки с полной подготовкой: удаление ржавчины, антикоррозийный грунт, эмаль.",
         ["Удаление ржавчины", "Преобразователь и антикоррозийный грунт", "Эмаль в 2 слоя", "Покраска сложных элементов кистью"]),
 "S11": ("Качество покраски определяется подготовкой. Заделываю трещины, ремонтирую штукатурку, шпаклюю и выравниваю стены перед покраской.",
         ["Расшивка и заделка трещин", "Ремонт отбитой штукатурки", "Шпаклёвка", "Шлифовка", "Грунтование"]),
 "S12": ("Влагозащитные покрытия и гидрофобизация — для стен, которые мокнут в дождь, цоколей, парапетов и камня. Важно: покрытие не заменяет устранение протечки, сначала ищем источник воды.",
         ["Диагностика источника влаги", "Гидрофобизатор для камня и бетона", "Влагозащитные краски", "Обработка швов и примыканий"]),
}

IMG = {"S01": "roller", "S02": "facade", "S03": "ladder", "S04": "bali", "S05": "roller", "S06": "varnish",
       "S07": "wood", "S08": "restore", "S09": "deck", "S10": "ladder", "S11": "hero", "S12": "facade"}
ALT = {"hero": "Мастер красит наружную стену дома", "facade": "Покраска белого фасада здания", "ladder": "Мастер на стремянке у фасада",
       "restore": "Мастер работает с деревом", "bali": "Строитель на объекте на Бали", "wood": "Балийский мастер работает с деревом",
       "roller": "Валик с краской на стене", "varnish": "Покрытие деревянной доски кистью", "villa": "Вилла с бассейном на Бали",
       "deck": "Деревянная терраса после дождя"}
D_IMG = {"sanur": "facade", "changu": "villa", "pererenan": "deck", "seminyak": "ladder", "kuta": "hero",
         "dzhimbaran": "facade", "nusa-dua": "villa", "uluvatu": "roller", "ubud": "wood", "denpasar": "restore"}
WA_SVG = '<svg class="wa-ico" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2zm5.3 14.1c-.2.6-1.3 1.2-1.8 1.2-.5.1-1 .2-3.3-.7-2.8-1.1-4.6-4-4.7-4.2-.1-.2-1.1-1.5-1.1-2.9s.7-2 1-2.3c.2-.3.5-.3.7-.3h.5c.2 0 .4 0 .6.5l.8 2c.1.2.1.4 0 .5l-.4.6-.4.4c-.1.2-.3.3-.1.6.2.3.8 1.3 1.7 2.1 1.2 1 2.1 1.4 2.4 1.5.3.1.5.1.6-.1l.9-1c.2-.3.4-.2.6-.1l1.9.9c.3.1.5.2.5.3.1.2.1.8-.2 1.3z"/></svg>'
SITE = CFG["site_url"].rstrip("/")

def wa_href():
    if CFG["whatsapp"]:
        return "https://wa.me/" + re.sub(r"[^0-9]", "", CFG["whatsapp"])
    if CFG["telegram"]:
        return "https://t.me/" + CFG["telegram"].lstrip("@")
    return f"{B}/kontakty/"

def cta(cls=""):
    return f'<a class="btn btn-wa {cls}" href="{wa_href()}">{WA_SVG}Написать мастеру</a>'

def img(name, eager=False):
    load = 'fetchpriority="high"' if eager else 'loading="lazy"'
    return f'<img src="{B}/img/{name}.webp" alt="{E(ALT[name])}" {load} decoding="async">'

def bali(h):  # страницы услуг: Бали вместо Санура
    return h.replace("в Сануре", "на Бали").replace("(Бали)", "").strip()

def short(t, n=110):
    return t if len(t) <= n else t[:n].rsplit(" ", 1)[0] + "…"

CSS = """:root{--bg:#FBF8F3;--surface:#fff;--ink:#1B2B27;--muted:#5B6B66;--line:#E8E2D8;--brand:#1F5E4F;--brand-2:#E9F2EE;--accent:#F2A65A;--accent-2:#FCEBD8;--wa:#25D366;--radius:20px}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font:17px/1.6 Inter,system-ui,sans-serif;-webkit-font-smoothing:antialiased}
h1,h2,h3{font-family:Fraunces,Georgia,serif;font-weight:600;line-height:1.15;letter-spacing:-.01em;margin:0 0 .5em}
h1{font-size:clamp(2rem,5vw,3.4rem)}h2{font-size:clamp(1.5rem,3vw,2.2rem)}h3{font-size:1.18rem}
a{color:var(--brand)}p{margin:0 0 1em}img{max-width:100%;display:block}
.wrap{max-width:1120px;margin:0 auto;padding-left:20px;padding-right:20px}
header{position:sticky;top:0;z-index:40;background:rgba(251,248,243,.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--line)}
.nav{display:flex;align-items:center;justify-content:space-between;height:68px;gap:16px}
.logo{font-family:Fraunces,serif;font-weight:600;font-size:1.3rem;color:var(--ink);text-decoration:none;white-space:nowrap}.logo span{color:var(--accent)}
.menu{display:flex;align-items:center;gap:22px}.menu>a,.dd>summary{color:var(--muted);text-decoration:none;font-size:.95rem;cursor:pointer;list-style:none}
.menu>a:hover,.dd>summary:hover{color:var(--ink)}.dd{position:relative}.dd>summary::-webkit-details-marker{display:none}.dd>summary::after{content:" ▾";font-size:.75em}
.dd-panel{position:absolute;top:34px;left:-16px;background:var(--surface);border:1px solid var(--line);border-radius:16px;box-shadow:0 16px 40px rgba(27,43,39,.12);padding:10px;display:grid;grid-template-columns:1fr 1fr;gap:2px;min-width:440px}
.dd-panel a{padding:8px 12px;border-radius:10px;text-decoration:none;color:var(--ink);font-size:.93rem}.dd-panel a:hover{background:var(--brand-2)}
.burger{display:none}
.btn{display:inline-flex;align-items:center;gap:10px;padding:14px 22px;border-radius:999px;font-weight:600;text-decoration:none;font-size:1rem;transition:transform .15s,box-shadow .15s;white-space:nowrap}
.btn:hover{transform:translateY(-1px)}.btn-wa{background:var(--wa);color:#fff;box-shadow:0 6px 20px rgba(37,211,102,.3)}
.btn-ghost{background:transparent;color:var(--ink);border:1.5px solid var(--line)}.btn-sm{padding:9px 16px;font-size:.9rem}
.wa-ico{width:20px;height:20px;fill:currentColor;flex:none}
.hero{padding-top:56px;padding-bottom:48px;display:grid;grid-template-columns:1.05fr .95fr;gap:48px;align-items:center}
.eyebrow{display:inline-block;background:var(--brand-2);color:var(--brand);padding:6px 14px;border-radius:999px;font-size:.85rem;font-weight:600;margin-bottom:18px}
.lead{font-size:1.15rem;color:var(--muted);max-width:36em}
.cta-row{display:flex;gap:12px;flex-wrap:wrap;margin:24px 0 18px}
.trust{display:flex;gap:10px 22px;flex-wrap:wrap;color:var(--muted);font-size:.92rem;list-style:none;padding:0;margin:0}
.trust li::before{content:"✓ ";color:var(--brand);font-weight:700}
.photo{position:relative;border-radius:28px;overflow:hidden;aspect-ratio:4/5;background:var(--accent-2)}.photo img{width:100%;height:100%;object-fit:cover}
.chip{position:absolute;left:18px;bottom:18px;background:#fff;border-radius:14px;padding:12px 16px;box-shadow:0 10px 30px rgba(0,0,0,.1);font-size:.9rem;max-width:80%}
.chip b{display:block;font-family:Fraunces,serif;font-size:1.05rem}
section{padding:56px 0}.alt-bg{background:var(--surface);border-top:1px solid var(--line);border-bottom:1px solid var(--line)}
.sec-head{max-width:680px;margin-bottom:30px}.sec-head p{color:var(--muted)}
.grid{display:grid;gap:18px;grid-template-columns:repeat(auto-fill,minmax(250px,1fr))}
.card{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);overflow:hidden;text-decoration:none;color:inherit;display:flex;flex-direction:column;transition:border-color .15s,box-shadow .15s,transform .15s}
a.card:hover{border-color:#cfd9d4;box-shadow:0 12px 30px rgba(27,43,39,.08);transform:translateY(-2px)}
.card img{aspect-ratio:16/10;object-fit:cover;width:100%}.card .body{padding:18px 20px 22px}.card h3{margin-bottom:6px}.card p{color:var(--muted);font-size:.95rem;margin:0}
.pills{display:flex;flex-wrap:wrap;gap:10px;padding:0;list-style:none;margin:0}.pills a,.pills span{display:inline-block;background:var(--brand-2);color:var(--brand);border-radius:999px;padding:9px 16px;font-weight:600;font-size:.93rem;text-decoration:none}
.pills a:hover{background:#d8eae3}
.split{display:grid;grid-template-columns:1fr 1fr;gap:48px;align-items:center}.split .photo{aspect-ratio:4/3}
.steps{counter-reset:s;display:grid;grid-template-columns:repeat(4,1fr);gap:18px;padding:0;list-style:none;margin:0}
.steps li{counter-increment:s;background:var(--surface);border-radius:var(--radius);padding:22px;border:1px solid var(--line)}
.steps li::before{content:counter(s);display:grid;place-items:center;width:36px;height:36px;border-radius:50%;background:var(--accent-2);color:#b0632a;font-weight:700;margin-bottom:12px}
.steps b{display:block;margin-bottom:4px}.steps.one{grid-template-columns:1fr}
.checks{list-style:none;padding:0;display:grid;gap:10px;margin:0}.checks li{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:14px 18px}
.checks li::before{content:"✓";color:var(--brand);font-weight:700;margin-right:10px}
details.faq{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:14px 18px;margin-bottom:10px}details.faq summary{cursor:pointer;font-weight:600}
details.faq p{margin:.6em 0 0;color:var(--muted)}
.crumbs{font-size:.88rem;color:var(--muted);padding-top:18px}.crumbs a{color:var(--muted)}
.band{background:var(--brand);color:#fff;border-radius:28px;padding:44px;display:grid;grid-template-columns:1.4fr 1fr;gap:24px;align-items:center}
.band h2{color:#fff}.band p{color:#d5e6df;margin:0}.band .cta-row{margin:0;justify-content:flex-end}
.page-hero{display:grid;grid-template-columns:1.1fr .9fr;gap:40px;align-items:center;padding-top:24px;padding-bottom:16px}.page-hero .photo{aspect-ratio:4/3}
.prose{max-width:760px}.prose h2{margin-top:1.6em}
.toc{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:18px 22px;margin:18px 0}.toc ol{margin:0;padding-left:20px}
.answer{background:var(--brand-2);border-radius:16px;padding:18px 22px;margin:14px 0 6px}
footer{border-top:1px solid var(--line);padding:40px 0 100px;color:var(--muted);font-size:.9rem}
.fcols{display:grid;grid-template-columns:repeat(3,1fr);gap:24px;margin-bottom:24px}.fcols a{display:block;color:var(--muted);text-decoration:none;padding:3px 0}.fcols b{color:var(--ink);display:block;margin-bottom:6px}
.ver{opacity:.6}.mbar{display:none}
.promises{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:16px}.promise{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);padding:22px}
.promise b{font-family:Fraunces,serif;font-size:1.15rem;display:block;margin-bottom:6px}.promise b::before{content:"✓ ";color:var(--brand)}.promise p{margin:0;color:var(--muted);font-size:.95rem}
.qf{display:grid;grid-template-columns:1fr 1fr;gap:40px;align-items:center;background:var(--accent-2);border-radius:28px;padding:40px}
.qform{display:grid;gap:14px}.qform label{display:grid;gap:6px;font-weight:600;font-size:.92rem}
.qform select,.qform input{font:inherit;padding:13px 14px;border:1.5px solid var(--line);border-radius:12px;background:#fff;color:var(--ink);width:100%}
.qform .btn{justify-content:center;border:0;cursor:pointer}
.prices{width:100%;border-collapse:collapse;background:var(--surface);border-radius:16px;overflow:hidden}.prices td{padding:14px 18px;border-bottom:1px solid var(--line)}
.dur{display:inline-block;background:var(--accent-2);color:#8a4b1c;border-radius:999px;padding:6px 14px;font-weight:600;font-size:.9rem;margin:0 0 14px}
@media (max-width:900px){
.menu{display:none}.burger{display:block}.burger>summary{list-style:none;cursor:pointer;width:44px;height:44px;display:grid;place-items:center;border:1.5px solid var(--line);border-radius:12px;font-size:1.3rem}
.burger>summary::-webkit-details-marker{display:none}
.burger[open]>summary{background:var(--brand-2)}
.mpanel{position:fixed;left:0;right:0;top:100%;height:calc(100vh - 68px);height:calc(100dvh - 68px);background:var(--bg);overflow:auto;padding:18px 20px 110px;border-top:1px solid var(--line)}
.mpanel b{display:block;margin:16px 0 6px;font-family:Fraunces,serif;font-size:1.1rem}.mpanel a{display:block;padding:9px 0;color:var(--ink);text-decoration:none;border-bottom:1px solid var(--line)}
.head-cta{display:none}body:has(.burger[open]){overflow:hidden}body:has(.burger[open]) .mbar{display:none}
.hero,.split,.page-hero,.band,.qf{grid-template-columns:1fr}.qf{padding:24px}.hero{padding-top:26px}.steps{grid-template-columns:1fr 1fr}.band{padding:28px}.band .cta-row{justify-content:flex-start}
.photo{aspect-ratio:4/3}.fcols{grid-template-columns:1fr 1fr}
.mbar{display:flex;position:fixed;left:12px;right:12px;bottom:12px;z-index:50}.mbar .btn{width:100%;justify-content:center}}
@media (max-width:520px){.steps{grid-template-columns:1fr}body{font-size:16px}.fcols{grid-template-columns:1fr}}"""

def nav_html():
    svc_links = "".join(f'<a href="{B}{p["url"]}">{E(bali(p["h1"]).replace(" на Бали", ""))}</a>' for p in svcs)
    svc_links += "".join(f'<a href="{B}/uslugi/{o["slug"]}/">{E(o["h1"].replace(" на Бали", ""))}</a>' for o in OFFERS)
    d_links = "".join(f'<a href="{B}/rayony/{d["slug"]}/">{E(d["ru"])}</a>' for d in DISTRICTS)
    desk = (f'<nav class="menu" aria-label="Главное меню"><details class="dd"><summary>Малярные работы</summary><div class="dd-panel">{svc_links}</div></details>'
            f'<details class="dd"><summary>Районы</summary><div class="dd-panel">{d_links}</div></details>'
            f'<a href="{B}/ceny/">Цены</a><a href="{B}/stati/">Статьи</a><a href="{B}/kontakty/">Контакты</a></nav>')
    mob = (f'<details class="burger"><summary aria-label="Открыть меню">☰</summary><div class="mpanel">'
           f'<a href="{B}/">Главная</a><b>Малярные работы</b>{svc_links}<b>Районы Бали</b>{d_links}'
           f'<b>Ещё</b><a href="{B}/ceny/">Цены</a><a href="{B}/stati/">Статьи</a><a href="{B}/raboty/">Работы</a><a href="{B}/kontakty/">Контакты</a></div></details>')
    return desk, mob

def footer_html():
    col = lambda t, items: f'<div><b>{t}</b>' + "".join(f'<a href="{B}{u}">{E(n)}</a>' for n, u in items) + "</div>"
    return (f'<footer><div class="wrap"><div class="fcols">'
            + col("Малярные работы", [(bali(p["h1"]), p["url"]) for p in svcs])
            + col("Районы", [(f"Маляр {d['loc']}", f"/rayony/{d['slug']}/") for d in DISTRICTS])
            + col("Полезное", [("Все районы", "/rayony/"), ("Статьи", "/stati/"), ("Цены", "/ceny/"), ("Работы", "/raboty/"), ("Контакты", "/kontakty/")])
            + f'</div>{E(CFG["name"])} — малярные работы, удаление плесени и покрытие дерева на Бали.<br>'
              'Фотографии — иллюстрации (Pexels, Unsplash), не работы мастера. <span class="ver">Версия ' + E(CFG.get("version", "")) + '</span></div></footer>')

def crumbs_schema(items):
    return {"@context": "https://schema.org", "@type": "BreadcrumbList", "itemListElement": [
        dict({"@type": "ListItem", "position": i + 1, "name": n}, **({"item": SITE + u} if u else {})) for i, (n, u) in enumerate(items)]}

def faq_schema(qa):
    return {"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": [
        {"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in qa]}

def faq_block(qa, title="Частые вопросы"):
    return (f'<section class="alt-bg"><div class="wrap"><div class="sec-head"><h2>{title}</h2></div>' +
            "".join(f'<details class="faq"><summary>{E(q)}</summary><p>{E(a)}</p></details>' for q, a in qa) + "</div></section>")

def band(title="Пришлите фото — оценю работу", text="Фото поверхности и примерная площадь. Отвечу, что нужно сделать и сколько это стоит."):
    return f'<section><div class="wrap band"><div><h2>{E(title)}</h2><p>{E(text)}</p></div><div class="cta-row">{cta()}</div></div></section>'

PAGES_META = []  # для sitemap и llms.txt

def page(path, title, desc, body, crumbs=None, schema=None, og=None, kind="page"):
    canon = SITE + path
    cr = ""
    if crumbs:
        cr = '<nav class="wrap crumbs" aria-label="Хлебные крошки">' + " › ".join(f'<a href="{B}{u}">{E(t)}</a>' if u else E(t) for t, u in crumbs) + "</nav>"
        schema = (schema or []) + [crumbs_schema([(t, u) for t, u in crumbs])]
    sch = "".join(f'<script type="application/ld+json">{json.dumps(x, ensure_ascii=False)}</script>' for x in (schema or []))
    desk, mob = nav_html()
    doc = f"""<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{E(title)}</title><meta name="description" content="{E(desc)}"><link rel="canonical" href="{canon}">
<meta property="og:type" content="website"><meta property="og:locale" content="ru_RU"><meta property="og:title" content="{E(title)}"><meta property="og:description" content="{E(desc)}">
<meta property="og:url" content="{canon}"><meta property="og:image" content="{SITE}/img/{og or 'hero'}.webp"><meta name="theme-color" content="#FBF8F3">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600&family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="{B}/style.css">{sch}</head><body>
<header><div class="wrap nav"><a class="logo" href="{B}/">Маляр<span>·</span>Бали</a>{desk}
<span class="head-cta">{cta("btn-sm")}</span>{mob}</div></header>
<main>{cr}{body}</main>{footer_html()}
<div class="mbar">{cta()}</div></body></html>"""
    d = os.path.join(OUT, path.strip("/"))
    os.makedirs(d, exist_ok=True)
    open(os.path.join(d, "index.html"), "w", encoding="utf-8").write(doc)
    PAGES_META.append((canon, title, desc, kind))
    return canon

def faq_for(pid, n=5):
    return [r[2] for r in sem if r[7] == pid and r[3] == "ru" and str(r[2]).startswith(QW) and "санур" not in r[2]][:n]

AREA_SERVED = [{"@type": "Place", "name": f"{d['en']}, Bali, Indonesia"} for d in DISTRICTS]
biz = {"@context": "https://schema.org", "@type": "HousePainter", "@id": SITE + "/#business", "name": CFG["name"], "url": SITE + "/",
       "image": SITE + "/img/hero.webp", "description": "Малярные работы на Бали: покраска стен и фасадов, удаление плесени, покрытие дерева лаком и маслом.",
       "areaServed": AREA_SERVED, "knowsLanguage": ["ru", "en"],
       "hasOfferCatalog": {"@type": "OfferCatalog", "name": "Малярные работы", "itemListElement": []}}
if CFG["phone"]:
    biz["telephone"] = CFG["phone"]
if CFG["gbp_url"]:
    biz["sameAs"] = [CFG["gbp_url"]]
website = {"@context": "https://schema.org", "@type": "WebSite", "name": CFG["name"], "url": SITE + "/", "inLanguage": "ru"}


# ---------------- v2.0: блоки доверия, форма, предложения ----------------
DUR = {"S01": "комната — обычно 1–2 дня", "S02": "фасад виллы — обычно 4–10 дней по погоде", "S03": "подкраска — обычно в пределах дня",
       "S04": "обработка + просушка — обычно 2–4 дня до покраски", "S05": "обработка — 1 день, плюс выдержка состава", "S06": "дверь или мебель — 2–4 дня с сушкой слоёв",
       "S07": "терраса или мебель — 1–3 дня", "S08": "предмет мебели — 3–7 дней", "S09": "терраса — 2–4 дня без дождя", "S10": "ворота или перила — 1–3 дня",
       "S11": "зависит от площади — 1–3 дня до покраски", "S12": "1–3 дня в сухую погоду"}

def rating_badge():
    if CFG.get("google_rating") and CFG.get("google_reviews"):
        return f'<li>★ {E(str(CFG["google_rating"]))} в Google · {CFG["google_reviews"]} отзывов</li>'
    return ""

def promises_block():
    items = list(CFG.get("promises") or [])
    if CFG.get("warranty"):
        items.append(["Гарантия", CFG["warranty"]])
    if CFG.get("response_time"):
        items.append(["Быстрый ответ", CFG["response_time"]])
    if not items:
        return ""
    cards = "".join(f'<div class="promise"><b>{E(t)}</b><p>{E(x)}</p></div>' for t, x in items)
    return (f'<section><div class="wrap"><div class="sec-head"><h2>Мои обещания</h2><p>Чего чаще всего боятся, когда пускают мастера в дом: '
            f'доплат, грязи, пропавшего мастера. Вот как я работаю.</p></div><div class="promises">{cards}</div></div></section>')

def master_block():
    if not (CFG.get("master_name") and CFG.get("master_story")):
        return ""
    pic = (f'<div class="photo"><img src="{B}/img/{E(CFG["master_photo"])}" alt="{E(CFG["master_name"])}, маляр на Бали" loading="lazy"></div>'
           if CFG.get("master_photo") else "")
    return (f'<section class="alt-bg"><div class="wrap split">{pic}<div><span class="eyebrow">Кто приедет</span>'
            f'<h2>{E(CFG["master_name"])}</h2><p>{E(CFG["master_story"])}</p></div></div></section>')

def prices_block():
    pr = CFG.get("prices") or []
    if not pr:
        return ""
    rows = "".join(f'<tr><td>{E(x["name"])}</td><td>от {E(x["from"])} {E(x.get("unit", ""))}</td></tr>' for x in pr)
    return (f'<section class="alt-bg"><div class="wrap"><div class="sec-head"><h2>Ориентиры цен</h2><p>Точная смета — по фото. Подготовка включена.</p></div>'
            f'<table class="prices">{rows}</table></div></section>')

def quick_form():
    d_opts = "".join(f'<option>{E(d["ru"])}</option>' for d in DISTRICTS) + "<option>Другой район</option>"
    s_opts = "".join(f'<option>{E(bali(p["h1"]).replace(" на Бали", ""))}</option>' for p in svcs)
    wa = re.sub(r"[^0-9]", "", CFG.get("whatsapp") or "")
    return f"""<section id="zayavka"><div class="wrap qf"><div><h2>Оценка по фото</h2>
<p class="lead">Выберите район и работу — откроется WhatsApp с готовым сообщением. Останется приложить 2–3 фото.</p>
<ul class="trust"><li>Без выезда</li><li>Бесплатно</li><li>Ни к чему не обязывает</li></ul></div>
<form class="qform" onsubmit="return qsend(this)"><label>Район<select name="d">{d_opts}</select></label>
<label>Что сделать<select name="s">{s_opts}</select></label>
<label>Комментарий<input name="c" placeholder="например: 2 спальни, плесень на потолке"></label>
<button class="btn btn-wa" type="submit">{WA_SVG}Отправить в WhatsApp</button></form></div></section>
<script>function qsend(f){{var t="Здравствуйте! Район: "+f.d.value+". Работа: "+f.s.value+". "+f.c.value+" Фото пришлю следом.";
var n="{wa}";if(n){{location.href="https://wa.me/"+n+"?text="+encodeURIComponent(t)}}else{{location.href="{B}/kontakty/"}}return false}}</script>"""

OFFERS = [
 dict(slug="podgotovka-villy-k-zaezdu", h1="Подготовка виллы к заезду гостей на Бали", img="villa",
      lead="Для владельцев и управляющих: между заездами освежаю стены, подкрашиваю следы, обновляю лак на дверях и мебели, убираю плесень. Цель — вилла выглядит как на фото в объявлении, а простой минимальный.",
      items=["Осмотр по фото или на месте до выезда гостей", "Список работ с приоритетами: что сделать сейчас, что в следующее окно",
             "Подкраска стен, дверей, плинтусов", "Удаление пятен плесени и обработка", "Лак и масло на уличной мебели и террасе", "Фотоотчёт для владельца"],
      faq=[("Успеете между заездами?", "Объём планируем под ваше окно между гостями. Если всё не помещается, делим на этапы."),
           ("Можно, если я не на Бали?", "Да. Согласование и фотоотчёт в WhatsApp, доступ — через управляющего.")]),
 dict(slug="obsluzhivanie-villy", h1="Обслуживание виллы по графику", img="deck",
      lead="Во влажном климате проще следить за покрытиями регулярно, чем раз в несколько лет делать большой ремонт. Осмотр по графику, мелкая подкраска, обработка плесени, обновление масла и лака.",
      items=["Осмотр стен, потолков, дерева и металла", "Подкраска и обработка очагов плесени", "Обновление масла на террасе и уличной мебели",
             "Проверка ржавчины на воротах и перилах", "Отчёт с фото и планом на следующий визит"],
      faq=[("Как часто нужен осмотр?", "Зависит от района и дома; у моря и в Убуде чаще. Частоту предложу после первого осмотра."),
           ("Это дешевле ремонта?", "Обычно да: мелкие дефекты не успевают превратиться в отслоения и гниль.")]),
]

GUIDE_HIRE = dict(slug="kak-vybrat-malyara-na-bali", h1="Как выбрать маляра на Бали и не пожалеть",
    title="Как выбрать маляра на Бали: чек-лист без рисков", desc="Что проверить у мастера на Бали до начала работ: смета, материалы, подготовка, предоплата, сроки, отчёты.",
    sections=[("Попросите смету письменно", "Объём, материалы, цена и что входит в подготовку. Устная договорённость — главный источник споров и доплат."),
              ("Уточните, входит ли подготовка", "Очистка, обработка плесени, шпаклёвка трещин и грунт. Без них краска во влажном климате отслаивается быстро."),
              ("Спросите про материалы", "Какая краска и лак, подходят ли они для влажности и солнца. Хороший мастер называет конкретные продукты."),
              ("Договоритесь о предоплате", "Разумно платить аванс на материалы с чеками, а работу — по этапам или по завершении."),
              ("Сроки и погода", "Наружные работы зависят от дождей. Спросите, как мастер планирует дни и что будет, если пойдёт дождь."),
              ("Отчёты и связь", "Если вы не на месте, договоритесь о фото этапов и о том, кто открывает виллу."),
              ("Посмотрите реальные работы", "Фото объектов на Бали, отзывы в Google с именами — лучше, чем красивые стоковые картинки."),
              ("Приёмка", "Осмотрите работу при дневном свете и сбоку: пропуски, потёки, следы на полу и мебели. Замечания — сразу, в WhatsApp с фото.")])

# ---------------- генерация ----------------
if os.path.exists(OUT):
    shutil.rmtree(OUT)
os.makedirs(os.path.join(OUT, "img"))
for f in os.listdir(os.path.join(ROOT, "img_src")):
    if f.endswith(".webp"):
        shutil.copy(os.path.join(ROOT, "img_src", f), os.path.join(OUT, "img", f))
open(os.path.join(OUT, "style.css"), "w", encoding="utf-8").write(CSS)
open(os.path.join(OUT, ".nojekyll"), "w").write("")
svcs = [p for p in pages.values() if p["type"] == "услуга"]
SVC = {p["id"]: p for p in svcs}
for p in svcs:
    biz["hasOfferCatalog"]["itemListElement"].append({"@type": "Offer", "itemOffered": {"@type": "Service", "name": bali(p["h1"]), "url": SITE + p["url"]}})

def svc_cards(items=None):
    return '<div class="grid">' + "".join(
        f'<a class="card" href="{B}{p["url"]}">{img(IMG[p["id"]])}<div class="body"><h3>{E(bali(p["h1"]).replace(" на Бали", ""))}</h3><p>{E(short(TXT[p["id"]][0]))}</p></div></a>'
        for p in (items or svcs)) + "</div>"

def district_cards(items=None):
    return '<div class="grid">' + "".join(
        f'<a class="card" href="{B}/rayony/{d["slug"]}/">{img(D_IMG[d["slug"]])}<div class="body"><h3>Маляр {E(d["loc"])}</h3><p>{E(short(d["intro"], 105))}</p></div></a>'
        for d in (items or DISTRICTS)) + "</div>"

def grid_url(d, sid):
    return f"/rayony/{d['slug']}/{SVC[sid]['url'].strip('/').split('/')[-1]}/"

STEPS = [("Фото в WhatsApp", "Пришлите фото и примерную площадь."), ("Осмотр", "Приезжаю, смотрю основание и влажность."),
         ("Смета", "Объём, материалы, сроки — до начала работ."), ("Работа и уборка", "Делаю, убираю, показываю результат.")]
HOME_FAQ = [("Как найти маляра на Бали?", "Напишите в WhatsApp, пришлите фото поверхности и укажите район. Мастер ответит, что нужно сделать, и договорится об осмотре."),
            ("В каких районах Бали работает мастер?", "Санур, Чангу, Перенан, Семиньяк, Кута, Джимбаран, Нуса Дуа, Улувату, Убуд и Денпасар."),
            ("Сколько стоит покраска на Бали?", "Цена зависит от площади, подготовки и материалов. Точную смету мастер даёт после фото или осмотра."),
            ("Можно ли просто закрасить плесень?", "Нет: плесень проступит через краску. Нужны очистка, обработка фунгицидом, просушка и антигрибковая краска."),
            ("Когда лучше красить фасад на Бали?", "В сухой сезон, примерно с апреля по октябрь, в дни без дождя.")]

# Главная
urls = []
urls.append(page("/", "Маляр на Бали — покраска, удаление плесени, лак для дерева", "Маляр на Бали: покраска стен и фасадов вилл, удаление плесени, покрытие дерева лаком и маслом. Санур, Чангу, Убуд, Семиньяк, Улувату и другие районы.", f"""
<div class="wrap hero"><div>
<span class="eyebrow">Русскоязычный маляр · 10 районов Бали</span>
<h1>Маляр на Бали: покраска, плесень, лак для дерева</h1>
<p class="lead">Аккуратно, по смете и с фотоотчётом. Крашу стены и фасады вилл, убираю плесень, защищаю дерево — с учётом влажности и морского воздуха, чтобы покрытие не облезло через сезон.</p>
<div class="cta-row">{cta()}<a class="btn btn-ghost" href="#zayavka">Оценка по фото</a></div>
<ul class="trust">{rating_badge()}<li>Смета до начала работ</li><li>Укрываю и убираю</li><li>Фотоотчёт в WhatsApp</li></ul>
</div><div class="photo">{img("hero", eager=True)}<div class="chip"><b>Бали</b>Покраска · плесень · дерево</div></div></div>

{promises_block()}{master_block()}
<section class="alt-bg" id="uslugi"><div class="wrap"><div class="sec-head"><h2>Малярные работы</h2><p>От подкраски одной стены до фасада виллы и террасы у бассейна.</p></div>{svc_cards()}</div></section>

<section id="rayony"><div class="wrap"><div class="sec-head"><h2>Маляр по районам Бали</h2><p>У каждого района свой климат и свои дома: в Убуде больше плесени, на Улувату — солнца, у побережья — соли. Выберите район.</p></div>{district_cards()}</div></section>

<section class="alt-bg"><div class="wrap split"><div class="photo">{img("facade")}</div><div>
<h2>Почему на Бали покрытия служат меньше</h2>
<p>Влажность, сезон дождей, солнце и морской воздух разрушают краску, лак и металл быстрее, чем в умеренном климате.</p>
<ul class="checks"><li>Проверяю влажность стены до покраски</li><li>Обрабатываю плесень, а не закрашиваю её</li><li>Подбираю краски и лаки под тропики</li><li>Планирую фасады на сухой сезон</li></ul>
</div></div></section>

<section><div class="wrap"><div class="sec-head"><h2>Как проходит работа</h2></div><ol class="steps">{"".join(f"<li><b>{E(a)}</b>{E(b)}</li>" for a, b in STEPS)}</ol></div></section>
<section class="alt-bg"><div class="wrap"><div class="sec-head"><h2>Для владельцев и управляющих вилл</h2><p>Когда важны сроки между заездами и контроль издалека.</p></div>
<div class="grid">{"".join(f'<a class="card" href="{B}/uslugi/{o["slug"]}/">{img(o["img"])}<div class="body"><h3>{E(o["h1"])}</h3><p>{E(short(o["lead"]))}</p></div></a>' for o in OFFERS)}
<a class="card" href="{B}/stati/{GUIDE_HIRE["slug"]}/">{img("ladder")}<div class="body"><h3>{E(GUIDE_HIRE["h1"])}</h3><p>{E(GUIDE_HIRE["desc"])}</p></div></a></div></div></section>
{prices_block()}
{quick_form()}

<section class="alt-bg"><div class="wrap"><div class="sec-head"><h2>Статьи по районам</h2><p>Что учитывать при покраске и уходе за домом в разных частях острова.</p></div>
<ul class="pills">{"".join(f'<li><a href="{B}/stati/{guide(d)["slug"]}/">{E(d["ru"])}</a></li>' for d in DISTRICTS)}</ul></div></section>
{faq_block(HOME_FAQ)}
{band()}
""", schema=[biz, website, faq_schema(HOME_FAQ)], kind="home"))

# Хабы
urls.append(page("/uslugi/", "Малярные работы на Бали — все услуги", "Все малярные работы на Бали: покраска стен и фасадов, плесень, лак и масло для дерева, металл, подготовка стен.",
                 f'<section><div class="wrap"><div class="sec-head"><h1>Малярные работы на Бали</h1><p class="lead">Выберите работу — на странице этапы, материалы, частые вопросы и районы.</p></div>{svc_cards()}</div></section>',
                 crumbs=[("Главная", "/"), ("Малярные работы", "")], kind="hub"))
urls.append(page("/rayony/", "Маляр по районам Бали — Санур, Чангу, Убуд и другие", "Маляр в 10 районах Бали: Санур, Чангу, Перенан, Семиньяк, Кута, Джимбаран, Нуса Дуа, Улувату, Убуд, Денпасар.",
                 f'<section><div class="wrap"><div class="sec-head"><h1>Маляр по районам Бали</h1><p class="lead">Климат и дома в районах разные — от этого зависят подготовка, материалы и сроки.</p></div>{district_cards()}</div></section>',
                 crumbs=[("Главная", "/"), ("Районы", "")], kind="hub"))

# Страницы услуг (Бали)
for p in svcs:
    intro, steps = TXT[p["id"]]
    h1 = bali(p["h1"])
    qa = [(q[0].upper() + q[1:] + "?", "Зависит от состояния поверхности и материала. Пришлите фото — мастер подскажет, что нужно в вашем случае.") for q in faq_for(p["id"])]
    cat = dict(GRID).get(p["id"])
    d_links = "".join(f'<li><a href="{B}{grid_url(d, p["id"]) if cat else "/rayony/" + d["slug"] + "/"}">{E(d["ru"])}</a></li>' for d in DISTRICTS)
    svc_schema = {"@context": "https://schema.org", "@type": "Service", "name": h1, "serviceType": h1, "areaServed": AREA_SERVED,
                  "provider": {"@id": SITE + "/#business"}, "url": SITE + p["url"]}
    others = [o for o in svcs if o["id"] != p["id"]][:3]
    urls.append(page(p["url"], f"{h1} — мастер, расчёт по фото", p["desc"].replace("в Сануре", "на Бали"), f"""
<div class="wrap page-hero"><div><span class="eyebrow">Бали · расчёт по фото</span><h1>{E(h1)}</h1><span class="dur">Срок: {E(DUR[p["id"]])}</span><p class="lead">{E(intro)}</p>
<div class="cta-row">{cta()}</div></div><div class="photo">{img(IMG[p["id"]], eager=True)}</div></div>
<section><div class="wrap split"><div><h2>Что входит в работу</h2><ol class="steps one">{"".join(f"<li>{E(s)}</li>" for s in steps)}</ol></div>
<div><h2>Стоимость</h2><p>Цена зависит от площади, состояния поверхности и материалов. Точную смету мастер даёт после фото или осмотра.</p>
<ul class="checks"><li>Смета до начала работ</li><li>Материалы под влажный климат</li><li>Уборка после работы</li></ul></div></div></section>
<section class="alt-bg"><div class="wrap"><div class="sec-head"><h2>{E(h1.replace(" на Бали", ""))} по районам</h2></div><ul class="pills">{d_links}</ul></div></section>
{faq_block(qa) if qa else ""}
{band("Нужна оценка?", "Пришлите фото поверхности, район и примерную площадь.")}
<section><div class="wrap"><div class="sec-head"><h2>Другие малярные работы</h2></div>{svc_cards(others)}</div></section>
""", crumbs=[("Главная", "/"), ("Малярные работы", "/uslugi/"), (h1, "")], schema=[svc_schema] + ([faq_schema(qa)] if qa else []), og=IMG[p["id"]], kind="service"))

# Районы, сетка, гайды
for i, d in enumerate(DISTRICTS):
    near = [DISTRICTS[(i + k) % len(DISTRICTS)] for k in (1, 2, 3)]
    g = guide(d)
    d_schema = {"@context": "https://schema.org", "@type": "Service", "name": f"Малярные работы {d['loc']}", "serviceType": "Малярные работы",
                "areaServed": {"@type": "Place", "name": f"{d['en']}, Bali, Indonesia"}, "provider": {"@id": SITE + "/#business"}, "url": f"{SITE}/rayony/{d['slug']}/"}
    grid_cards = '<div class="grid">' + "".join(
        f'<a class="card" href="{B}{grid_url(d, sid)}">{img(IMG[sid])}<div class="body"><h3>{E(bali(SVC[sid]["h1"]).replace(" на Бали", ""))} {E(d["loc"])}</h3><p>{E(short(d["notes"][cat], 100))}</p></div></a>'
        for sid, cat in GRID) + "</div>"
    urls.append(page(f"/rayony/{d['slug']}/", f"Маляр {d['loc']} — покраска, плесень, лак для дерева", f"Маляр {d['loc']}: покраска стен и фасадов, удаление плесени, лак и масло для дерева с учётом климата {d['gen']}.", f"""
<div class="wrap page-hero"><div><span class="eyebrow">{E(d["ru"])} · Бали</span><h1>Маляр {E(d["loc"])}</h1>
<div class="answer">Малярные работы {E(d["loc"])}: покраска стен и фасадов, удаление плесени, лак и масло для дерева. Пришлите фото и адрес — мастер оценит объём.</div>
<p class="lead">{E(d["intro"])}</p><div class="cta-row">{cta()}</div></div><div class="photo">{img(D_IMG[d["slug"]], eager=True)}</div></div>
<section class="alt-bg"><div class="wrap"><div class="sec-head"><h2>Работы {E(d["loc"])}</h2><p>С учётом местного климата и типичных домов: {E(d["houses"])}.</p></div>{grid_cards}</div></section>
<section><div class="wrap split"><div><h2>Что разрушает покрытия {E(d["loc"])}</h2><ul class="checks">{"".join(f"<li>{E(r[0].upper() + r[1:])}</li>" for r in d["risks"])}</ul>
<p style="margin-top:18px"><a href="{B}/stati/{g["slug"]}/">Подробно: {E(g["h1"])}</a></p></div>
<div><h2>Другие работы</h2><ul class="pills">{"".join(f'<li><a href="{B}{p["url"]}">{E(bali(p["h1"]).replace(" на Бали", ""))}</a></li>' for p in svcs if p["id"] not in dict(GRID))}</ul></div></div></section>
{faq_block(d["faq"])}
<section><div class="wrap"><div class="sec-head"><h2>Соседние районы</h2></div><ul class="pills">{"".join(f'<li><a href="{B}/rayony/{n["slug"]}/">Маляр {E(n["loc"])}</a></li>' for n in near)}<li><a href="{B}/rayony/">Все районы</a></li></ul></div></section>
{band()}
""", crumbs=[("Главная", "/"), ("Районы", "/rayony/"), (d["ru"], "")], schema=[d_schema, faq_schema(d["faq"])], og=D_IMG[d["slug"]], kind="district"))

    for sid, cat in GRID:
        s = SVC[sid]
        name = bali(s["h1"]).replace(" на Бали", "")
        h1 = f"{name} {d['loc']}"
        intro, steps = TXT[sid]
        qa = [(f"Выполняете {name.lower()} {d['loc']}?", f"Да, {d['ru']} — один из районов выезда. Пришлите фото и адрес для оценки."),
              (f"Что важно учесть {d['loc']}?", d["notes"][cat])]
        others = [(sid2, c2) for sid2, c2 in GRID if sid2 != sid]
        g_schema = {"@context": "https://schema.org", "@type": "Service", "name": h1, "serviceType": name,
                    "areaServed": {"@type": "Place", "name": f"{d['en']}, Bali, Indonesia"}, "provider": {"@id": SITE + "/#business"}, "url": SITE + grid_url(d, sid)}
        urls.append(page(grid_url(d, sid), f"{h1} — мастер, расчёт по фото", f"{h1}: {short(d['notes'][cat], 120)}", f"""
<div class="wrap page-hero"><div><span class="eyebrow">{E(d["ru"])} · {E(name)}</span><h1>{E(h1)}</h1>
<div class="answer">{E(d["notes"][cat])}</div><p class="lead">{E(intro)}</p><div class="cta-row">{cta()}</div></div><div class="photo">{img(IMG[sid], eager=True)}</div></div>
<section><div class="wrap split"><div><h2>Этапы работы</h2><ol class="steps one">{"".join(f"<li>{E(x)}</li>" for x in steps)}</ol></div>
<div><h2>Особенности района</h2><p>{E(d["intro"])}</p><ul class="checks">{"".join(f"<li>{E(r[0].upper() + r[1:])}</li>" for r in d["risks"][:3])}</ul></div></div></section>
{faq_block(qa)}
<section><div class="wrap split"><div><h2>Другие работы {E(d["loc"])}</h2><ul class="pills">{"".join(f'<li><a href="{B}{grid_url(d, s2)}">{E(bali(SVC[s2]["h1"]).replace(" на Бали", ""))}</a></li>' for s2, _ in others)}<li><a href="{B}/rayony/{d["slug"]}/">Все работы {E(d["loc"])}</a></li></ul></div>
<div><h2>{E(name)} в других районах</h2><ul class="pills">{"".join(f'<li><a href="{B}{grid_url(n, sid)}">{E(n["ru"])}</a></li>' for n in near)}<li><a href="{B}{s["url"]}">Вся информация об услуге</a></li></ul></div></div></section>
{band()}
""", crumbs=[("Главная", "/"), ("Районы", "/rayony/"), (d["ru"], f"/rayony/{d['slug']}/"), (name, "")], schema=[g_schema, faq_schema(qa)], og=IMG[sid], kind="grid"))

    art_schema = {"@context": "https://schema.org", "@type": "Article", "headline": g["h1"], "description": g["desc"], "inLanguage": "ru",
                  "about": {"@type": "Place", "name": f"{d['en']}, Bali"}, "author": {"@id": SITE + "/#business"}, "publisher": {"@id": SITE + "/#business"},
                  "image": f"{SITE}/img/{D_IMG[d['slug']]}.webp", "mainEntityOfPage": f"{SITE}/stati/{g['slug']}/"}
    toc = "".join(f'<li><a href="#s{j}">{E(t)}</a></li>' for j, (t, _) in enumerate(g["sections"]))
    body = "".join(f'<h2 id="s{j}">{E(t)}</h2><p>{E(x)}</p>' for j, (t, x) in enumerate(g["sections"]))
    urls.append(page(f"/stati/{g['slug']}/", g["title"], g["desc"], f"""
<article class="wrap prose"><h1>{E(g["h1"])}</h1>
<div class="answer"><b>Коротко:</b> {E(d["intro"])} Главные риски: {E(", ".join(d["risks"][:3]))}.</div>
<div class="toc"><b>Содержание</b><ol>{toc}</ol></div>{body}
<h2>Работы {E(d["loc"])}</h2><ul class="pills">{"".join(f'<li><a href="{B}{grid_url(d, s2)}">{E(bali(SVC[s2]["h1"]).replace(" на Бали", ""))}</a></li>' for s2, _ in GRID)}<li><a href="{B}/rayony/{d["slug"]}/">Маляр {E(d["loc"])}</a></li></ul>
</article>{band()}
""", crumbs=[("Главная", "/"), ("Статьи", "/stati/"), (d["ru"], "")], schema=[art_schema], og=D_IMG[d["slug"]], kind="article"))

for o in OFFERS:
    o_schema = {"@context": "https://schema.org", "@type": "Service", "name": o["h1"], "areaServed": AREA_SERVED, "provider": {"@id": SITE + "/#business"}, "url": f"{SITE}/uslugi/{o['slug']}/"}
    urls.append(page(f"/uslugi/{o['slug']}/", o["h1"] + " — маляр", short(o["lead"], 150), f"""
<div class="wrap page-hero"><div><span class="eyebrow">Для владельцев и управляющих</span><h1>{E(o["h1"])}</h1><p class="lead">{E(o["lead"])}</p>
<div class="cta-row">{cta()}</div></div><div class="photo">{img(o["img"], eager=True)}</div></div>
<section><div class="wrap split"><div><h2>Что входит</h2><ul class="checks">{"".join(f"<li>{E(x)}</li>" for x in o["items"])}</ul></div>
<div><h2>Районы</h2><ul class="pills">{"".join(f'<li><a href="{B}/rayony/{d["slug"]}/">{E(d["ru"])}</a></li>' for d in DISTRICTS)}</ul></div></div></section>
{promises_block()}{faq_block(o["faq"])}{quick_form()}
""", crumbs=[("Главная", "/"), ("Малярные работы", "/uslugi/"), (o["h1"], "")], schema=[o_schema, faq_schema(o["faq"])], og=o["img"], kind="service"))

gh = GUIDE_HIRE
gh_schema = {"@context": "https://schema.org", "@type": "Article", "headline": gh["h1"], "description": gh["desc"], "inLanguage": "ru",
             "author": {"@id": SITE + "/#business"}, "publisher": {"@id": SITE + "/#business"}, "image": SITE + "/img/ladder.webp", "mainEntityOfPage": f"{SITE}/stati/{gh['slug']}/"}
urls.append(page(f"/stati/{gh['slug']}/", gh["title"], gh["desc"], f"""
<article class="wrap prose"><h1>{E(gh["h1"])}</h1>
<div class="answer"><b>Коротко:</b> письменная смета, подготовка в цене, понятные материалы, аванс только на материалы, фото этапов и приёмка при дневном свете.</div>
<div class="toc"><b>Чек-лист</b><ol>{"".join(f'<li><a href="#h{j}">{E(t)}</a></li>' for j, (t, _) in enumerate(gh["sections"]))}</ol></div>
{"".join(f'<h2 id="h{j}">{E(t)}</h2><p>{E(x)}</p>' for j, (t, x) in enumerate(gh["sections"]))}
</article>{promises_block()}{quick_form()}
""", crumbs=[("Главная", "/"), ("Статьи", "/stati/"), ("Как выбрать маляра", "")], schema=[gh_schema], og="ladder", kind="article"))

urls.append(page("/stati/", "Статьи о покраске и уходе за домом на Бали", "Гайды по районам Бали: климат, типичные дома, что разрушает краску и дерево, когда планировать работы.",
                 f'<section><div class="wrap"><div class="sec-head"><h1>Статьи о покраске на Бали</h1><p class="lead">Гайды по районам: климат, дома и типичные проблемы покрытий.</p></div><div class="grid">' + f'<a class="card" href="{B}/stati/{GUIDE_HIRE["slug"]}/">{img("ladder")}<div class="body"><h3>{E(GUIDE_HIRE["h1"])}</h3><p>{E(GUIDE_HIRE["desc"])}</p></div></a>' +
                 "".join(f'<a class="card" href="{B}/stati/{guide(d)["slug"]}/">{img(D_IMG[d["slug"]])}<div class="body"><h3>{E(guide(d)["h1"])}</h3><p>{E(guide(d)["desc"])}</p></div></a>' for d in DISTRICTS) +
                 "</div></div></section>", crumbs=[("Главная", "/"), ("Статьи", "")], kind="hub"))

for pid, h1, title, desc, body, pic in [
        ("P090", "Цены на малярные работы на Бали", "Цены на покраску, обработку от плесени и лакировку — Бали", "Из чего складывается цена малярных работ на Бали. Точная смета — по фото и осмотру.",
         "Цена складывается из площади, подготовки (очистка, обработка от плесени, шпаклёвка), материалов и доступа (высота, леса). Пришлите фото и район — пришлю расчёт.", "roller"),
        ("P091", "Примеры работ", "Примеры малярных работ на Бали — до и после", "Фото объектов на Бали: покраска, плесень, дерево.",
         "Здесь появятся фотографии реальных объектов: до и после.", "villa"),
        ("P092", "Контакты мастера", "Контакты — маляр на Бали", "Связаться с маляром на Бали: WhatsApp, Telegram. Районы выезда.",
         "Работаю в 10 районах Бали. Напишите, пришлите фото и укажите район — отвечу и договоримся об осмотре.", "bali")]:
    p = pages[pid]
    urls.append(page(p["url"], title, desc,
                     f'<div class="wrap page-hero"><div><h1>{E(h1)}</h1><p class="lead">{E(body)}</p><div class="cta-row">{cta()}</div></div><div class="photo">{img(pic, eager=True)}</div></div>',
                     crumbs=[("Главная", "/"), (h1, "")], og=pic))

# llms.txt — карта сайта для ИИ-ассистентов
L = [f"# {CFG['name']}", "", "> Маляр на Бали: покраска стен и фасадов вилл, удаление плесени, покрытие дерева лаком и маслом. "
     "Районы: " + ", ".join(d["ru"] for d in DISTRICTS) + ". Язык: русский, английский. Расчёт по фото.", ""]
for kind, head in [("home", "Главная"), ("service", "Малярные работы"), ("district", "Районы"), ("grid", "Работы по районам"), ("article", "Статьи"), ("hub", "Разделы"), ("page", "Прочее")]:
    items = [m for m in PAGES_META if m[3] == kind]
    if items:
        L += [f"## {head}", ""] + [f"- [{t}]({u}): {d}" for u, t, d, _ in items] + [""]
open(os.path.join(OUT, "llms.txt"), "w", encoding="utf-8").write("\n".join(L))

open(os.path.join(OUT, "sitemap.xml"), "w", encoding="utf-8").write(
    '<?xml version="1.0" encoding="UTF-8"?><urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' +
    "".join(f"<url><loc>{u}</loc></url>" for u in urls) + "</urlset>")
open(os.path.join(OUT, "robots.txt"), "w").write(f"User-agent: *\nAllow: /\nSitemap: {CFG['site_url']}/sitemap.xml\n")
open(os.path.join(OUT, "404.html"), "w", encoding="utf-8").write(f'<!doctype html><meta charset="utf-8"><title>Страница не найдена</title><p>Страница не найдена. <a href="{B}/">На главную</a></p>')
print("pages:", len(urls))
