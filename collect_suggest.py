# -*- coding: utf-8 -*-
"""Сбор Google Autocomplete (gl=id, hl=ru) по фразам ядра. Кеш: suggest_cache.json (возобновляемый)."""
import json, os, time, urllib.parse, urllib.request, sys
import openpyxl

CACHE = r"D:\Маляр\suggest_cache.json"
BOOK = r"D:\Маляр\SEO-малярные-работы-Бали-2026-09-25.xlsx"
cache = json.load(open(CACHE, encoding="utf-8")) if os.path.exists(CACHE) else {}

def fetch(q, hl):
    key = f"{hl}|{q}"
    if key in cache:
        return
    url = "https://suggestqueries.google.com/complete/search?client=firefox&gl=id&hl=%s&q=%s" % (hl, urllib.parse.quote(q))
    for attempt in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=15) as r:
                data = json.loads(r.read().decode("utf-8", "replace"))
            cache[key] = {"ts": time.strftime("%Y-%m-%d %H:%M"), "s": data[1]}
            return
        except Exception as e:
            time.sleep(3 * (attempt + 1))
    cache[key] = {"ts": time.strftime("%Y-%m-%d %H:%M"), "s": None, "err": "fail"}

wb = openpyxl.load_workbook(BOOK, read_only=True)
ws = wb["09 Семантика"]
seeds = []
for row in ws.iter_rows(min_row=2, values_only=True):
    qid, raw, norm, lang, *_ , status = row[:9]
    if row[-1] == "SRC07":
        continue  # подсказки второго уровня не расширяем
    if row[8] in ("принято условно", "требует проверки") and row[3] == "ru":
        seeds.append((norm, "ru"))
    elif row[3] == "en":
        seeds.append((norm, "en"))
markers = ["плесень", "грибок на", "покраска", "покрасить", "лак для", "покрыть лаком", "масло для дерева", "пропитка для дерева",
           "краска для", "маляр", "сырость", "тик", "терраса дерево", "реставрация мебели", "шпаклевка", "ржавчина", "фасад", "потолок"]
ALPH = "абвгдежзиклмнопрстуфхцчшэя"
for m in markers:
    for c in ALPH:
        seeds.append((f"{m} {c}", "ru"))
    seeds.append((f"{m} ", "ru"))
seeds.append(("маляр бали", "ru")); seeds.append(("painter bali", "en"))
for m in ["маляр санур", "покраска санур", "плесень санур", "мастер санур", "ремонт санур", "вилла санур"]:
    seeds.append((m + " ", "ru"))
for m in ["painter sanur", "painting sanur", "mold sanur", "handyman sanur", "varnish sanur", "villa maintenance sanur", "house painter denpasar"]:
    seeds.append((m + " ", "en"))
    for c in "abcdefghijklmnoprstuvw":
        seeds.append((f"{m} {c}", "en"))
todo = [s for s in dict.fromkeys(seeds) if f"{s[1]}|{s[0]}" not in cache]
print("seeds", len(seeds), "todo", len(todo), flush=True)
from concurrent.futures import ThreadPoolExecutor
def job(a):
    fetch(*a); time.sleep(0.3)
with ThreadPoolExecutor(6) as ex:
    for i, _ in enumerate(ex.map(job, todo), 1):
        if i % 100 == 0:
            json.dump(dict(cache), open(CACHE, "w", encoding="utf-8"), ensure_ascii=False)
            print(i, sum(1 for v in list(cache.values()) if v.get("s") is None), flush=True)
json.dump(cache, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False)
print("done", len(cache), sum(1 for v in cache.values() if v.get("s") is None))
