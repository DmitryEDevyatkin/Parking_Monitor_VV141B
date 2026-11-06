#!/usr/bin/env python3
"""
Мониторинг машиномест в МФК «Варшавские ворота» (РГ-Девелопмент).

Каждый запуск:
  1. забирает список машиномест с сайта rg-dev.ru (тот же источник, что и карта паркинга);
  2. сравнивает его с прошлым запуском (файл state.json);
  3. записывает все изменения в журнал history.csv — это общая база, она только пополняется;
  4. пересобирает реестр мест places.csv и отчёт parking.xlsx;
  5. если что-то изменилось — пишет в Telegram и прикладывает parking.xlsx.

Статусы на сайте: 1 = «Свободно», 2 = «Забронировано».
Проданные места из списка пропадают — бот помнит их по журналу.

Нужны две переменные окружения:
  TELEGRAM_TOKEN    — токен бота от @BotFather
  TELEGRAM_CHAT_ID  — ваш числовой id в Telegram

Команды:
  python monitor.py            — обычная проверка
  python monitor.py --test     — проверить связь с Telegram
  python monitor.py --rebuild  — пересобрать places.csv и parking.xlsx из history.csv (без сайта)

Для Excel-отчёта нужна библиотека openpyxl (pip install openpyxl).
Без неё бот работает, просто не создаёт parking.xlsx.
"""

import csv
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import datetime, timedelta, timezone

# ---------- Настройки ----------
COMPLEX_SLUG = "varshavskie-vorota"
BASE = "https://rg-dev.ru"
PAGE_URL = f"{BASE}/parking/{COMPLEX_SLUG}/"
SALES_PHONE = "+7 (495) 106-90-57"

STATE_FILE = os.environ.get("STATE_FILE", "state.json")
HISTORY_FILE = os.environ.get("HISTORY_FILE", "history.csv")
PLACES_FILE = os.environ.get("PLACES_FILE", "places.csv")
XLSX_FILE = os.environ.get("XLSX_FILE", "parking.xlsx")

TG_TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
TG_CHAT = os.environ.get("TELEGRAM_CHAT_ID", "").strip()

MSK = timezone(timedelta(hours=3))
DAILY_REPORT_HOUR = 10   # ежедневная сводка «бот жив» — после 10:00 МСК, раз в день
FAIL_ALERT_AFTER = 3     # предупредить, если сайт не отвечает N запусков подряд
SOLD_LIST_LIMIT = 20     # сколько проданных мест показывать в сообщении (остальные — в файле)

FREE = 1
BOOKED = 2
ST_FREE = "в продаже"
ST_BOOKED = "в резерве"
ST_GONE = "нет в списке"

TS_FMT = "%d.%m.%Y %H:%M"

# Названия событий в журнале
EV_FIRST = "Первое наблюдение"
EV_NEW_FREE = "Вывод в продажу"
EV_NEW_BOOKED = "Новое в резерве"
EV_NEW_OTHER = "Новое в списке"
EV_RETURNED = "Вернулось в список"
EV_FREE_AGAIN = "Снова свободно"
EV_BOOKED = "Забронировано"
EV_STATUS = "Изменение статуса"
EV_GONE = "Пропало из списка"
EV_PRICE = "Изменение цены"
EV_RECONCILE = "Сверка статуса"

HISTORY_COLUMNS = [
    "Дата и время", "Место", "Уровень", "Площадь, м²", "Событие",
    "Было", "Стало", "Цена, ₽", "Прежняя цена, ₽", "Источник", "Комментарий",
]
PLACES_COLUMNS = [
    "Место", "Уровень", "Площадь, м²", "Статус", "Цена (текущая или последняя), ₽",
    "Первая цена, ₽", "Впервые замечено", "Последнее изменение", "Пропало из списка",
    "Сколько раз бронировали", "История",
]
# --------------------------------


def log(msg):
    print(f"[{datetime.now(MSK):%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def now_msk():
    return datetime.now(MSK).replace(tzinfo=None, second=0, microsecond=0)


def status_text(code):
    return {FREE: ST_FREE, BOOKED: ST_BOOKED}.get(code, f"статус {code}")


def to_int(value):
    try:
        return int(float(str(value).replace(" ", "").replace(",", ".")))
    except (TypeError, ValueError):
        return None


def to_float(value):
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return None


def parse_ts(text):
    try:
        return datetime.strptime(text.strip(), TS_FMT)
    except (ValueError, AttributeError):
        return None


# ---------- Сайт ----------

def get_json(path):
    req = urllib.request.Request(
        BASE + path,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
            "Accept": "application/json",
        },
    )
    last_error = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception as e:  # сеть, таймаут, неожиданный ответ
            last_error = e
            time.sleep(5 * (attempt + 1))
    raise RuntimeError(f"{path}: {last_error}")


def fetch_places():
    """Возвращает {id места: {...}} по всем корпусам комплекса."""
    complex_data = get_json(f"/api/parking_complexes/{COMPLEX_SLUG}/")
    places = {}
    for building in complex_data.get("building_set", []):
        data = get_json(f"/api/parking_buildings/{building['id']}/")
        for floor in data.get("floor_set", []):
            for p in floor.get("parking_place_set") or []:
                places[p["id"]] = {
                    "number": str(p.get("number", "?")),
                    "floor": floor.get("number"),
                    "building": str(building.get("name", "")),
                    "area": p.get("total_square"),
                    "price": to_int(p.get("actual_price") or p.get("price")),
                    "status": p.get("status"),
                    "action": p.get("action_title") or "",
                }
    return places


# ---------- Журнал (база) ----------

def load_history():
    try:
        with open(HISTORY_FILE, encoding="utf-8-sig", newline="") as f:
            return [row for row in csv.DictReader(f) if row.get("Место")]
    except FileNotFoundError:
        return []


def save_history(rows):
    with open(HISTORY_FILE, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HISTORY_COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def make_row(ts, p, event, was, became, old_price=None, source="бот", comment=""):
    return {
        "Дата и время": ts.strftime(TS_FMT),
        "Место": str(p.get("number", "?")),
        "Уровень": "" if p.get("floor") is None else str(p.get("floor")),
        "Площадь, м²": "" if p.get("area") is None else str(p.get("area")),
        "Событие": event,
        "Было": was or "",
        "Стало": became or "",
        "Цена, ₽": "" if not p.get("price") else str(p.get("price")),
        "Прежняя цена, ₽": "" if not old_price else str(old_price),
        "Источник": source,
        "Комментарий": comment,
    }


SHORT = {ST_FREE: "продажа", ST_BOOKED: "бронь", ST_GONE: "пропало"}


def fmt_mln(n):
    return f"{n / 1_000_000:.2f}".rstrip("0").rstrip(".").replace(".", ",") + " млн"


def build_registry(history):
    """Сводит журнал в реестр: {номер места: {...}}. Порядок строк журнала = хронология."""
    reg = {}
    for row in history:
        n = row["Место"]
        ts = parse_ts(row["Дата и время"])
        if ts is None:
            continue
        r = reg.get(n)
        if r is None:
            r = reg[n] = {
                "number": n, "floor": None, "area": None, "price": None, "first_price": None,
                "status": None, "status_since": None, "first_seen": ts, "last_change": ts,
                "gone_at": None, "gone_from": None, "bookings": 0, "trail": [], "periods": [],
                "approx": False,
            }
        if row.get("Уровень", "") != "":
            r["floor"] = to_int(row["Уровень"])
        if row.get("Площадь, м²", "") != "":
            r["area"] = to_float(row["Площадь, м²"])
        price = to_int(row.get("Цена, ₽"))
        if price:
            if r["first_price"] is None:
                r["first_price"] = price
            r["price"] = price
        if "приблиз" in (row.get("Комментарий", "") + row.get("Источник", "")).lower():
            r["approx"] = True

        became = row.get("Стало", "")
        if row["Событие"] == EV_PRICE:
            old = to_int(row.get("Прежняя цена, ₽"))
            if old and price:
                r["trail"].append(f"{ts:%d.%m} цена {fmt_mln(old)} → {fmt_mln(price)}")
        elif became and became != r["status"]:
            if r["status"] is not None and r["status_since"] is not None:
                r["periods"].append((r["status"], r["status_since"], ts, became))
            if became == ST_BOOKED:
                r["bookings"] += 1
            if became == ST_GONE:
                r["gone_at"] = ts
                r["gone_from"] = r["status"] or row.get("Было", "")
            else:
                r["gone_at"] = None
                r["gone_from"] = None
            r["status"] = became
            r["status_since"] = ts
            r["trail"].append(f"{ts:%d.%m %H:%M} {SHORT.get(became, became)}")
        r["last_change"] = ts
    return reg


def place_label(r):
    s = r["status"]
    if s == ST_GONE:
        if r["gone_from"] == ST_BOOKED:
            return "продано (предположительно)"
        if r["gone_from"] == ST_FREE:
            return "снято с продажи или продано"
        return "нет в списке"
    return s or "?"


def reconcile(registry, places, ts):
    """Строки-поправки, если база разошлась с сайтом (первый запуск, пропущенные проверки)."""
    rows = []
    current = {p["number"]: p for p in places.values()}
    for n, p in current.items():
        st = status_text(p["status"])
        r = registry.get(n)
        if r is None:
            rows.append(make_row(ts, p, EV_FIRST, "", st))
        elif r["status"] != st:
            rows.append(make_row(ts, p, EV_RECONCILE, r["status"], st, source="сверка",
                                 comment="база расходилась с сайтом, точное время изменения неизвестно"))
    for n, r in registry.items():
        if n not in current and r["status"] not in (ST_GONE, None):
            rows.append(make_row(ts, r, EV_GONE, r["status"], ST_GONE, source="сверка",
                                 comment="обнаружено при сверке, точное время неизвестно"))
    return rows


# ---------- Изменения ----------

def find_changes(old, new, registry):
    events = []

    def ev(icon, text, kind, p, was, became, old_price=None):
        events.append({"icon": icon, "text": text, "kind": kind, "place": p,
                       "was": was, "became": became, "old_price": old_price})

    for pid, p in new.items():
        st = status_text(p["status"])
        o = old.get(pid)
        if o is None:
            hist = registry.get(p["number"])
            if hist and hist["status"] == ST_GONE:
                ev("🟢" if p["status"] == FREE else "🟡", f"Вернулось в список, {st}",
                   EV_RETURNED, p, ST_GONE, st)
            elif p["status"] == FREE:
                ev("🟢", "Новое место в продаже", EV_NEW_FREE, p, "", st)
            elif p["status"] == BOOKED:
                ev("🟡", "Новое место в списке, в резерве", EV_NEW_BOOKED, p, "", st)
            else:
                ev("⚪", f"Новое место в списке, {st}", EV_NEW_OTHER, p, "", st)
            continue
        if o["status"] != p["status"]:
            was = status_text(o["status"])
            if p["status"] == FREE:
                ev("🟢", "Снова свободно", EV_FREE_AGAIN, p, was, st)
            elif p["status"] == BOOKED:
                ev("🟡", "Забронировано", EV_BOOKED, p, was, st)
            else:
                ev("⚪", f"Статус: {was} → {st}", EV_STATUS, p, was, st)
        if o.get("price") and p.get("price") and o["price"] != p["price"]:
            ev("💰", f"Цена: {fmt_price(o['price'])} → {fmt_price(p['price'])}",
               EV_PRICE, p, st, st, old_price=o["price"])
    for pid, o in old.items():
        if pid not in new:
            ev("🔴", "Пропало из списка — продано или снято с продажи",
               EV_GONE, o, status_text(o["status"]), ST_GONE)
    order = {"🟢": 0, "💰": 1, "🟡": 2, "🔴": 3, "⚪": 4}
    events.sort(key=lambda e: order.get(e["icon"], 9))
    return events


# ---------- Оформление сообщений ----------

def fmt_price(n):
    return f"{n:,}".replace(",", " ") + " ₽" if n else "цена не указана"


def place_line(p):
    parts = [f"№{p['number']}", f"{p['floor']} ур."]
    if p.get("area"):
        parts.append(f"{p['area']} м²")
    parts.append(fmt_price(p.get("price")))
    line = " · ".join(parts)
    if p.get("action"):
        line += f" ({p['action']})"
    return line


def sort_key(p):
    return (p.get("floor") or 0, str(p.get("number")))


def gone_sorted(registry):
    gone = [r for r in registry.values() if r["status"] == ST_GONE]
    return sorted(gone, key=lambda r: r["gone_at"] or datetime.min, reverse=True)


def summary(places, registry):
    free = sorted((p for p in places.values() if p["status"] == FREE), key=sort_key)
    booked = sorted((p for p in places.values() if p["status"] == BOOKED), key=sort_key)
    other = sorted((p for p in places.values() if p["status"] not in (FREE, BOOKED)), key=sort_key)
    lines = []
    if free:
        lines.append(f"🟢 В продаже ({len(free)}):")
        lines += [f"  • {place_line(p)}" for p in free]
    else:
        lines.append("🟢 В продаже: свободных мест нет.")
    lines.append("")
    if booked:
        lines.append(f"🟡 В резерве ({len(booked)}):")
        lines += [f"  • {place_line(p)}" for p in booked]
    else:
        lines.append("🟡 В резерве: нет.")
    if other:
        lines.append("")
        lines.append(f"⚪ С другим статусом ({len(other)}):")
        lines += [f"  • {place_line(p)} — {status_text(p['status'])}" for p in other]

    gone = gone_sorted(registry)
    lines.append("")
    if gone:
        lines.append(f"🔴 Продано, по нашим данным ({len(gone)}):")
        for r in gone[:SOLD_LIST_LIMIT]:
            when = f" — пропало {r['gone_at']:%d.%m %H:%M}" if r["gone_at"] else ""
            note = "" if r["gone_from"] == ST_BOOKED else " (было свободно — возможно, снято с продажи)"
            lines.append(f"  • {place_line(r)}{when}{note}")
        if len(gone) > SOLD_LIST_LIMIT:
            lines.append(f"  …и ещё {len(gone) - SOLD_LIST_LIMIT} — в файле parking.xlsx")
    else:
        lines.append("🔴 Продано, по нашим данным: пока нет.")

    lines.append("")
    lines.append(f"Карта: {PAGE_URL}")
    lines.append(f"Отдел продаж: {SALES_PHONE}")
    return lines


def build_message(title, events, places, registry):
    lines = [title, ""]
    for e in events:
        lines.append(f"{e['icon']} {e['text']}: {place_line(e['place'])}")
    if events:
        lines.append("")
    lines += summary(places, registry)
    return "\n".join(lines)


# ---------- Реестр и Excel ----------

def registry_rows(registry):
    order = {ST_FREE: 0, ST_BOOKED: 1}
    items = sorted(
        registry.values(),
        key=lambda r: (order.get(r["status"], 2),
                       -(r["gone_at"] or datetime.min).timestamp() if r["status"] == ST_GONE else 0,
                       r["floor"] or 0, r["number"]),
    )
    rows = []
    for r in items:
        rows.append({
            "Место": r["number"],
            "Уровень": r["floor"],
            "Площадь, м²": r["area"],
            "Статус": place_label(r),
            "Цена (текущая или последняя), ₽": r["price"],
            "Первая цена, ₽": r["first_price"],
            "Впервые замечено": r["first_seen"],
            "Последнее изменение": r["last_change"],
            "Пропало из списка": r["gone_at"],
            "Сколько раз бронировали": r["bookings"],
            "История": " → ".join(r["trail"]),
        })
    return rows


def save_places_csv(registry):
    with open(PLACES_FILE, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=PLACES_COLUMNS)
        w.writeheader()
        for row in registry_rows(registry):
            w.writerow({k: (v.strftime(TS_FMT) if isinstance(v, datetime) else ("" if v is None else v))
                        for k, v in row.items()})


def fmt_duration(td):
    minutes = int(td.total_seconds() // 60)
    days, rest = divmod(minutes, 1440)
    hours, mins = divmod(rest, 60)
    if days:
        return f"{days} дн {hours} ч"
    if hours:
        return f"{hours} ч {mins} мин"
    return f"{mins} мин"


def compute_stats(registry, history):
    regs = list(registry.values())
    gone = [r for r in regs if r["status"] == ST_GONE]
    sold = [r for r in gone if r["gone_from"] == ST_BOOKED]
    withdrawn = [r for r in gone if r["gone_from"] != ST_BOOKED]
    booked_now = [r for r in regs if r["status"] == ST_BOOKED]

    free_to_booked, booked_to_gone = [], []
    for r in regs:
        for status, start, end, nxt in r["periods"]:
            if status == ST_FREE and nxt == ST_BOOKED:
                free_to_booked.append(end - start)
            if status == ST_BOOKED and nxt == ST_GONE:
                booked_to_gone.append(end - start)

    def avg(lst):
        return sum(lst, timedelta()) / len(lst) if lst else None

    def avg_price(lst):
        prices = [r["price"] for r in lst if r["price"]]
        return round(sum(prices) / len(prices)) if prices else None

    sold_prices = [r["price"] for r in sold if r["price"]]
    first_ts = min((r["first_seen"] for r in regs), default=None)
    return {
        "total": len(regs),
        "free": sum(r["status"] == ST_FREE for r in regs),
        "booked": len(booked_now),
        "sold": len(sold),
        "withdrawn": len(withdrawn),
        "avg_sold": avg_price(sold),
        "min_sold": min(sold_prices) if sold_prices else None,
        "max_sold": max(sold_prices) if sold_prices else None,
        "avg_booked": avg_price(booked_now),
        "free_to_booked": (avg(free_to_booked), len(free_to_booked)),
        "booked_to_gone": (avg(booked_to_gone), len(booked_to_gone)),
        "since": first_ts,
        "approx": any(r["approx"] for r in regs),
    }


def write_xlsx(registry, history, path=XLSX_FILE):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        log("Библиотека openpyxl не установлена — parking.xlsx не создан (pip install openpyxl).")
        return False

    PRICE = '#,##0" ₽"'
    DATE = "DD.MM.YYYY HH:MM"
    bold = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="DDE3EA")
    fills = {
        ST_FREE: PatternFill("solid", fgColor="D9F2D9"),
        ST_BOOKED: PatternFill("solid", fgColor="FFF2CC"),
        "gone": PatternFill("solid", fgColor="F8D7DA"),
    }

    def header(ws, cols, widths):
        ws.append(cols)
        for i, w in enumerate(widths, 1):
            c = ws.cell(row=1, column=i)
            c.font = bold
            c.fill = head_fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    wb = Workbook()

    # --- Реестр ---
    ws = wb.active
    ws.title = "Реестр"
    header(ws, PLACES_COLUMNS, [8, 10, 11, 28, 18, 14, 16, 16, 16, 14, 80])
    for row in registry_rows(registry):
        ws.append([row[c] for c in PLACES_COLUMNS])
        i = ws.max_row
        for col in (5, 6):
            ws.cell(row=i, column=col).number_format = PRICE
        for col in (7, 8, 9):
            ws.cell(row=i, column=col).number_format = DATE
        status = row["Статус"]
        fill = fills.get(status) or (fills["gone"] if status not in (ST_FREE, ST_BOOKED) else None)
        if fill:
            ws.cell(row=i, column=4).fill = fill
    ws.auto_filter.ref = ws.dimensions

    # --- Журнал ---
    wj = wb.create_sheet("Журнал")
    header(wj, HISTORY_COLUMNS, [16, 8, 8, 10, 22, 14, 14, 14, 14, 26, 50])
    for row in history:
        values = []
        for c in HISTORY_COLUMNS:
            v = row.get(c, "")
            if c == "Дата и время":
                v = parse_ts(v) or v
            elif c in ("Цена, ₽", "Прежняя цена, ₽", "Уровень"):
                v = to_int(v) if v != "" else None
            elif c == "Площадь, м²":
                v = to_float(v) if v != "" else None
            values.append(v)
        wj.append(values)
        i = wj.max_row
        wj.cell(row=i, column=1).number_format = DATE
        wj.cell(row=i, column=8).number_format = PRICE
        wj.cell(row=i, column=9).number_format = PRICE
    wj.auto_filter.ref = wj.dimensions

    # --- Статистика ---
    st = compute_stats(registry, history)
    wsx = wb.create_sheet("Статистика")
    wsx.column_dimensions["A"].width = 46
    for col in "BCDE":
        wsx.column_dimensions[col].width = 18

    def line(label, value=None, fmt=None, note=None):
        wsx.append([label, value, note])
        if fmt:
            wsx.cell(row=wsx.max_row, column=2).number_format = fmt

    def section(title):
        wsx.append([])
        wsx.append([title])
        wsx.cell(row=wsx.max_row, column=1).font = bold

    wsx.append(["Статистика по паркингу «Варшавские ворота»"])
    wsx.cell(row=1, column=1).font = Font(bold=True, size=13)
    line("Отчёт сформирован", now_msk(), DATE)
    line("Наблюдения ведутся с", st["since"], DATE)

    section("Места")
    line("Всего мест замечено за всё время", st["total"])
    line("В продаже сейчас", st["free"])
    line("В резерве сейчас", st["booked"])
    line("Продано (пропало из резерва)", st["sold"])
    line("Снято с продажи или продано (пропало из свободных)", st["withdrawn"])

    section("Цены")
    line("Средняя цена проданных", st["avg_sold"], PRICE)
    line("Минимальная цена проданных", st["min_sold"], PRICE)
    line("Максимальная цена проданных", st["max_sold"], PRICE)
    line("Средняя цена мест в резерве", st["avg_booked"], PRICE)

    section("Скорость")
    d, n = st["free_to_booked"]
    line("В среднем от появления в продаже до брони", fmt_duration(d) if d else "нет данных",
         note=f"случаев: {n}")
    d, n = st["booked_to_gone"]
    line("В среднем от брони до исчезновения (продажи)", fmt_duration(d) if d else "нет данных",
         note=f"случаев: {n}")
    if st["approx"]:
        wsx.append(["Часть событий восстановлена по отчётам, их время приблизительное."])
        wsx.cell(row=wsx.max_row, column=1).font = Font(italic=True, color="666666")

    section("По уровням")
    wsx.append(["Уровень", "Всего", "В продаже", "В резерве", "Продано / снято"])
    for c in range(1, 6):
        wsx.cell(row=wsx.max_row, column=c).font = bold
    floors = sorted({r["floor"] for r in registry.values()}, key=lambda x: (x is None, x))
    for fl in floors:
        rs = [r for r in registry.values() if r["floor"] == fl]
        wsx.append([fl, len(rs), sum(r["status"] == ST_FREE for r in rs),
                    sum(r["status"] == ST_BOOKED for r in rs),
                    sum(r["status"] == ST_GONE for r in rs)])

    section("По дням")
    wsx.append(["Дата", "Вышло в продажу / освободилось", "Забронировано", "Пропало из списка"])
    for c in range(1, 5):
        wsx.cell(row=wsx.max_row, column=c).font = bold
    days = {}
    for row in history:
        ts = parse_ts(row["Дата и время"])
        if not ts:
            continue
        d = days.setdefault(ts.date(), [0, 0, 0])
        became, was = row.get("Стало"), row.get("Было")
        if became == ST_FREE and was != ST_FREE and row["Событие"] != EV_FIRST:
            d[0] += 1
        elif became == ST_BOOKED and was != ST_BOOKED and row["Событие"] != EV_FIRST:
            d[1] += 1
        elif became == ST_GONE and was != ST_GONE:
            d[2] += 1
    for day in sorted(days):
        wsx.append([day.strftime("%d.%m.%Y")] + days[day])

    wb.save(path)
    return True


def save_all(history, registry):
    save_history(history)
    save_places_csv(registry)
    return write_xlsx(registry, history)


# ---------- Telegram ----------

def _tg_request(method, data, headers=None):
    if not TG_TOKEN or not TG_CHAT:
        return None  # не настроен
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{TG_TOKEN}/{method}", data=data, headers=headers or {}
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            ok = json.loads(resp.read().decode()).get("ok", False)
    except urllib.error.HTTPError as e:
        try:
            reason = json.loads(e.read().decode()).get("description", "")
        except Exception:
            reason = ""
        hint = ""
        if "chat not found" in reason:
            hint = (" → Проверьте TELEGRAM_CHAT_ID (нужен ВАШ id от @userinfobot, только цифры) "
                    "и что вы нажали «Запустить» в своём боте.")
        elif e.code == 401:
            hint = " → Неверный TELEGRAM_TOKEN, скопируйте его у @BotFather заново."
        elif e.code == 403:
            hint = " → Бот заблокирован или не запущен: откройте своего бота и нажмите «Запустить»."
        log(f"Ошибка Telegram ({method}): HTTP {e.code} {reason}{hint}")
        return False
    except Exception as e:
        log(f"Ошибка Telegram ({method}): {e}")
        return False
    if not ok:
        log(f"Telegram отклонил запрос {method} — проверьте токен и chat_id.")
    return ok


def send_telegram(text):
    if len(text) > 4000:
        text = text[:3990] + "\n…"
    data = urllib.parse.urlencode({
        "chat_id": TG_CHAT, "text": text, "disable_web_page_preview": "true",
    }).encode()
    result = _tg_request("sendMessage", data)
    if result is None:
        log("Telegram не настроен (нет TELEGRAM_TOKEN / TELEGRAM_CHAT_ID). Сообщение:\n" + text)
        return True
    return result


def send_document(path, caption=""):
    if not os.path.exists(path):
        return False
    filename = f"parking_VV_{now_msk():%Y-%m-%d_%H%M}.xlsx"
    if not TG_TOKEN or not TG_CHAT:
        log(f"Telegram не настроен — файл {filename} не отправлен.")
        return True
    boundary = uuid.uuid4().hex
    with open(path, "rb") as f:
        content = f.read()
    body = b""
    for name, value in (("chat_id", TG_CHAT), ("caption", caption)):
        body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{name}\"\r\n\r\n"
                 f"{value}\r\n").encode()
    body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"document\"; "
             f"filename=\"{filename}\"\r\nContent-Type: application/vnd.openxmlformats-"
             f"officedocument.spreadsheetml.sheet\r\n\r\n").encode() + content + b"\r\n"
    body += f"--{boundary}--\r\n".encode()
    return bool(_tg_request("sendDocument", body,
                            {"Content-Type": f"multipart/form-data; boundary={boundary}"}))


# ---------- Состояние ----------

def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1, sort_keys=True)


# ---------- Основной запуск ----------

def main():
    try:  # чтобы эмодзи и кириллица не ломали вывод в консоль Windows
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    if "--test" in sys.argv:
        ok = send_telegram("✅ Тест: бот мониторинга машиномест «Варшавские ворота» на связи.")
        log("Тестовое сообщение отправлено." if ok else "Не удалось отправить тестовое сообщение.")
        return 0 if ok else 1

    if "--rebuild" in sys.argv:
        history = load_history()
        registry = build_registry(history)
        save_places_csv(registry)
        write_xlsx(registry, history)
        log(f"Пересобрано из журнала: событий {len(history)}, мест {len(registry)}.")
        return 0

    state = load_state()
    now = now_msk()
    today = now.strftime("%Y-%m-%d")
    old_places = state.get("places")

    try:
        places = fetch_places()
        if not places and old_places and len(old_places) > 3:
            raise RuntimeError("сайт вернул пустой список — похоже на сбой, пропускаю")
    except Exception as e:
        state["fails"] = state.get("fails", 0) + 1
        log(f"Не удалось получить данные ({state['fails']} раз подряд): {e}")
        if state["fails"] == FAIL_ALERT_AFTER:
            send_telegram(
                f"⚠️ Не получается получить данные с сайта уже {FAIL_ALERT_AFTER} проверки подряд.\n"
                f"Ошибка: {e}\nПока связь не восстановится, уведомлений об изменениях не будет."
            )
        save_state(state)
        return 0

    # Защита от неполного ответа сайта: если разом пропала большая часть мест,
    # не верим с первого раза — ждём, пока картина повторится 3 проверки подряд.
    if old_places:
        vanished = [pid for pid in old_places if pid not in places]
        if len(vanished) >= 4 and len(vanished) > len(old_places) / 2:
            state["suspect"] = state.get("suspect", 0) + 1
            if state["suspect"] < 3:
                log(f"Пропало сразу {len(vanished)} из {len(old_places)} мест — похоже на сбой сайта, "
                    f"жду подтверждения ({state['suspect']}/3).")
                save_state(state)
                return 0
        state["suspect"] = 0

    if state.get("fails", 0) >= FAIL_ALERT_AFTER:
        send_telegram("✅ Связь с сайтом восстановлена, мониторинг работает.")
    state["fails"] = 0

    # --- Журнал ---
    history = load_history()
    registry = build_registry(history)
    events = find_changes(old_places, places, registry) if old_places is not None else []
    new_rows = [make_row(now, e["place"], e["kind"], e["was"], e["became"], e["old_price"])
                for e in events]
    registry = build_registry(history + new_rows)
    new_rows += reconcile(registry, places, now)
    history = history + new_rows
    registry = build_registry(history)

    # --- Сообщение ---
    message = None
    if old_places is None:
        message = build_message("🅿️ Мониторинг машиномест «Варшавские ворота» запущен",
                                [], places, registry)
    elif events:
        message = build_message("🅿️ Изменения на паркинге «Варшавские ворота»",
                                events, places, registry)
    elif not state.get("db_ready"):
        message = build_message("📊 База машиномест подключена: теперь бот помнит все места, "
                                "включая проданные. Excel-отчёт — в файле ниже.",
                                [], places, registry)
    elif now.hour >= DAILY_REPORT_HOUR and state.get("last_report_date") != today:
        message = build_message("☀️ Ежедневная сводка: изменений нет, бот работает",
                                [], places, registry)

    if message:
        if not send_telegram(message):
            # Ничего не запоминаем — в следующий раз попробуем снова, события не потеряются
            save_state(state)
            return 1
        state["last_report_date"] = today
        log("Сообщение отправлено.")

    if message or new_rows:
        xlsx_ok = save_all(history, registry)
        if message and xlsx_ok:
            send_document(XLSX_FILE, "Реестр мест, журнал событий и статистика")
    if not message:
        log(f"Изменений нет. Мест в списке: {len(places)}, "
            f"свободно: {sum(p['status'] == FREE for p in places.values())}.")

    state["db_ready"] = True
    state["places"] = places
    save_state(state)
    return 0


if __name__ == "__main__":
    sys.exit(main())
