# Маляр в Сануре

Статический сайт мастера малярных работ в Сануре (Бали) для GitHub Pages (папка `docs/`).

- `site_config.json` — имя, WhatsApp, Telegram, телефон, ссылка на Google Business Profile.
- `build_site.py` — генерирует `docs/` из SEO-книги (книга хранится локально, в репозиторий не входит).
- `build_core.py`, `collect_suggest.py` — сбор семантического ядра и подсказок Google.

Пересборка: `python build_core.py && python build_site.py`.
