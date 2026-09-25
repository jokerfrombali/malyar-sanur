# -*- coding: utf-8 -*-
"""Статический сайт «Маляр в Сануре» из книги SEO. Выход: docs/ (GitHub Pages)."""
import json, os, re, shutil, html
import openpyxl

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
 "S01": ("Крашу стены и потолки в виллах, квартирах и домах Санура: от одной комнаты до всего дома. Перед покраской проверяю основание на влагу и плесень — во влажном климате это главная причина, по которой свежая краска пузырится и отслаивается.",
         ["Осмотр и замер влажности стен", "Защита мебели и пола", "Очистка, обработка антисептиком при необходимости", "Шпаклёвка трещин и шлифовка", "Грунт и два слоя краски", "Уборка после работы"]),
 "S02": ("Покраска фасадов вилл у моря требует другого подхода: соль, солнце и дожди разрушают покрытие быстрее, чем в глубине острова. Подбираю фасадные краски под прибрежные условия и планирую работы на сухой сезон.",
         ["Осмотр трещин, отслоений и зелёного налёта", "Мойка и биоцидная обработка", "Ремонт трещин эластичными составами", "Грунт и фасадная краска", "Покраска цоколя, парапетов, металлических элементов"]),
 "S03": ("Не всегда нужно перекрашивать всю стену. Подкрашиваю следы после арендаторов, царапины, пятна после протечек и сколы — с подбором цвета, чтобы не было заметных пятен.",
         ["Подбор и колеровка цвета по образцу", "Локальная подготовка и грунт", "Подкраска с растушёвкой", "Честная рекомендация, если стену лучше перекрасить целиком"]),
 "S04": ("Плесень в Сануре — частая проблема: морская влажность и сезон дождей. Удаляю плесень со стен, потолков, дерева и швов, нахожу причину и защищаю поверхность, чтобы она не вернулась через месяц.",
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

def contact_html():
    parts = []
    if CFG["whatsapp"]:
        parts.append(f'<a class="btn" href="https://wa.me/{re.sub(r"[^0-9]", "", CFG["whatsapp"])}">WhatsApp</a>')
    if CFG["telegram"]:
        parts.append(f'<a class="btn alt" href="https://t.me/{CFG["telegram"].lstrip("@")}">Telegram</a>')
    if CFG["phone"]:
        parts.append(f'<a class="btn alt" href="tel:{re.sub(r"[^0-9+]", "", CFG["phone"])}">Позвонить</a>')
    return " ".join(parts) or '<a class="btn" href="%s/kontakty/">Связаться</a>' % B

CSS = """:root{--bg:#fbfaf7;--fg:#1d2327;--mut:#5b6670;--acc:#0f6e6e;--acc2:#e9f3f2;--card:#fff;--bd:#e3e1dc}
@media (prefers-color-scheme:dark){:root{--bg:#141719;--fg:#e8eaeb;--mut:#a3adb5;--acc:#5cc2bb;--acc2:#1d2a2a;--card:#1b1f22;--bd:#2c3236}}
*{box-sizing:border-box}body{margin:0;font:17px/1.6 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--fg)}
a{color:var(--acc)}header,main,footer{max-width:960px;margin:0 auto;padding:0 16px}
header{display:flex;justify-content:space-between;align-items:center;padding-top:14px;padding-bottom:14px;border-bottom:1px solid var(--bd);flex-wrap:wrap;gap:8px}
header .logo{font-weight:700;text-decoration:none;color:var(--fg)}nav a{margin-left:14px;text-decoration:none;font-size:15px}
h1{font-size:clamp(26px,5vw,38px);line-height:1.2;margin:28px 0 12px}h2{margin-top:32px}
.lead{font-size:19px;color:var(--mut)}.btn{display:inline-block;background:var(--acc);color:#fff;padding:11px 18px;border-radius:8px;text-decoration:none;font-weight:600;margin:4px 6px 4px 0}
.btn.alt{background:var(--acc2);color:var(--acc)}.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px}
.card{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:16px}.card h3{margin:0 0 6px;font-size:18px}.card p{margin:0;color:var(--mut);font-size:15px}
.card a{text-decoration:none}ol,ul{padding-left:22px}details{border-bottom:1px solid var(--bd);padding:10px 0}summary{cursor:pointer;font-weight:600}
.crumbs{font-size:14px;color:var(--mut);margin-top:16px}.cta{background:var(--acc2);border-radius:12px;padding:20px;margin:36px 0}
footer{border-top:1px solid var(--bd);margin-top:48px;padding-top:18px;padding-bottom:40px;color:var(--mut);font-size:14px}
.sticky{position:fixed;right:16px;bottom:16px}"""

def page(path, title, desc, body, crumbs=None, schema=None):
    canon = CFG["site_url"].rstrip("/") + path
    cr = ""
    if crumbs:
        cr = '<div class="crumbs">' + " › ".join(f'<a href="{B}{u}">{E(t)}</a>' if u else E(t) for t, u in crumbs) + "</div>"
    sch = "".join(f'<script type="application/ld+json">{json.dumps(x, ensure_ascii=False)}</script>' for x in (schema or []))
    doc = f"""<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{E(title)}</title><meta name="description" content="{E(desc)}"><link rel="canonical" href="{canon}">
<link rel="stylesheet" href="{B}/style.css">{sch}</head><body>
<header><a class="logo" href="{B}/">{E(CFG["name"])}</a><nav><a href="{B}/uslugi/">Услуги</a><a href="{B}/ceny/">Цены</a><a href="{B}/raboty/">Работы</a><a href="{B}/kontakty/">Контакты</a></nav></header>
<main>{cr}{body}</main>
<footer>{E(CFG["name"])} — малярные работы, удаление плесени и покрытие дерева в Сануре, Бали.</footer>
<div class="sticky">{contact_html().split(" ")[0] if (CFG["whatsapp"] or CFG["telegram"]) else ""}</div></body></html>"""
    d = os.path.join(OUT, path.strip("/"))
    os.makedirs(d, exist_ok=True)
    open(os.path.join(d, "index.html"), "w", encoding="utf-8").write(doc)
    return canon

def faq_for(pid, n=5):
    qs = []
    for r in sem:
        if r[7] == pid and r[3] == "ru" and str(r[2]).startswith(QW) and "санур" not in r[2]:
            qs.append(r[2])
    return qs[:n]

biz = {"@context": "https://schema.org", "@type": "HousePainter", "name": CFG["name"], "url": CFG["site_url"] + "/",
       "areaServed": {"@type": "Place", "name": "Sanur, Bali, Indonesia"}}
if CFG["phone"]:
    biz["telephone"] = CFG["phone"]
if CFG["gbp_url"]:
    biz["sameAs"] = [CFG["gbp_url"]]

if os.path.exists(OUT):
    shutil.rmtree(OUT)
os.makedirs(OUT)
open(os.path.join(OUT, "style.css"), "w", encoding="utf-8").write(CSS)
open(os.path.join(OUT, ".nojekyll"), "w").write("")
urls = []
svcs = [p for p in pages.values() if p["type"] == "услуга"]

def svc_cards():
    return '<div class="grid">' + "".join(
        f'<div class="card"><a href="{B}{p["url"]}"><h3>{E(p["h1"])}</h3></a><p>{E(TXT[p["id"]][0][:120])}…</p></div>' for p in svcs) + "</div>"

home = pages["P000"]
urls.append(page("/", home["title"], home["desc"], f"""
<h1>{E(home["h1"])}</h1>
<p class="lead">Покраска стен и фасадов вилл, удаление плесени, покрытие дерева лаком и маслом. Работаю только в Сануре — быстро приезжаю на осмотр и знаю, как влажность и морской воздух влияют на покрытия.</p>
<p>{contact_html()}</p><p>Пришлите фото и примерную площадь — оценю объём работ и стоимость.</p>
<h2>Услуги</h2>{svc_cards()}
<h2>Почему у моря покрытия служат меньше</h2>
<p>В Сануре дома стоят рядом с океаном: соль в воздухе ускоряет коррозию металла, солнце разрушает лак на уличном дереве, а высокая влажность и сезон дождей дают плесень на стенах и потолках. Поэтому я начинаю с диагностики причины, а не с закрашивания следствий.</p>
<h2>Район работы</h2><p>Санур: Санур Кайя, Санур Кауф, Семаванг, Синду, Мертасари, Батуджимбар.</p>
""", schema=[biz]))
urls.append(page("/uslugi/", "Услуги маляра в Сануре", "Все малярные услуги в Сануре: покраска, плесень, лак и масло для дерева, металл.",
                 f"<h1>Услуги маляра в Сануре</h1>{svc_cards()}", crumbs=[("Главная", "/"), ("Услуги", "")]))
for p in svcs:
    intro, steps = TXT[p["id"]]
    faq = faq_for(p["id"])
    faq_html = ("<h2>Частые вопросы</h2>" + "".join(f"<details><summary>{E(q[0].upper() + q[1:])}?</summary><p>Ответ зависит от состояния поверхности — пришлите фото, и я подскажу, что нужно именно в вашем случае.</p></details>" for q in faq)) if faq else ""
    svc_schema = {"@context": "https://schema.org", "@type": "Service", "name": p["h1"], "areaServed": "Sanur, Bali", "provider": {"@type": "HousePainter", "name": CFG["name"]}}
    crumbs_schema = {"@context": "https://schema.org", "@type": "BreadcrumbList", "itemListElement": [
        {"@type": "ListItem", "position": 1, "name": "Главная", "item": CFG["site_url"] + "/"},
        {"@type": "ListItem", "position": 2, "name": "Услуги", "item": CFG["site_url"] + "/uslugi/"},
        {"@type": "ListItem", "position": 3, "name": p["h1"]}]}
    urls.append(page(p["url"], p["title"], p["desc"], f"""
<h1>{E(p["h1"])}</h1><p class="lead">{E(intro)}</p><p>{contact_html()}</p>
<h2>Что входит в работу</h2><ol>{"".join(f"<li>{E(s)}</li>" for s in steps)}</ol>
<h2>Стоимость</h2><p>Цена зависит от площади, состояния поверхности и материалов. Точную смету даю после фото или осмотра — выезд в пределах Санура.</p>
{faq_html}
<div class="cta"><strong>Нужна оценка?</strong><p>Пришлите фото поверхности и примерную площадь.</p>{contact_html()}</div>
<h2>Другие услуги</h2><ul>{"".join(f'<li><a href="{B}{o["url"]}">{E(o["h1"])}</a></li>' for o in svcs if o["id"] != p["id"])}</ul>
""", crumbs=[("Главная", "/"), ("Услуги", "/uslugi/"), (p["h1"], "")], schema=[svc_schema, crumbs_schema]))
for pid, body in [("P090", "<p>Цены формируются по смете: площадь, подготовка (очистка, обработка от плесени, шпаклёвка), материалы и доступ (высота, леса). Пришлите фото — пришлю расчёт.</p>"),
                  ("P091", "<p>Раздел наполняется фотографиями реальных объектов в Сануре: до и после.</p>"),
                  ("P092", "<p>Район работы: Санур. Напишите — отвечу и договоримся об осмотре.</p>")]:
    p = pages[pid]
    urls.append(page(p["url"], p["title"], p["desc"], f"<h1>{E(p['h1'])}</h1>{body}<p>{contact_html()}</p>", crumbs=[("Главная", "/"), (p["h1"], "")]))

open(os.path.join(OUT, "sitemap.xml"), "w", encoding="utf-8").write(
    '<?xml version="1.0" encoding="UTF-8"?><urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">' +
    "".join(f"<url><loc>{u}</loc></url>" for u in urls) + "</urlset>")
open(os.path.join(OUT, "robots.txt"), "w").write(f"User-agent: *\nAllow: /\nSitemap: {CFG['site_url']}/sitemap.xml\n")
open(os.path.join(OUT, "404.html"), "w", encoding="utf-8").write(f'<!doctype html><meta charset="utf-8"><title>Страница не найдена</title><p>Страница не найдена. <a href="{B}/">На главную</a></p>')
print("pages:", len(urls))
