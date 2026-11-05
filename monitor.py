#!/usr/bin/env python3
"""
Мониторинг машиномест в МФК «Варшавские ворота» (РГ-Девелопмент).

Каждый запуск:
  1. забирает список машиномест с сайта rg-dev.ru (тот же источник, что и карта паркинга);
  2. сравнивает его с прошлым запуском (файл state.json);
  3. если что-то изменилось — пишет в Telegram.

Статусы на сайте: 1 = «Свободно», 2 = «Забронировано».
Проданные места из списка пропадают.

Нужны две переменные окружения:
  TELEGRAM_TOKEN    — токен бота от @BotFather
  TELEGRAM_CHAT_ID  — ваш числовой id в Telegram

Проверка связи с Telegram:  python monitor.py --test
Работает на стандартном Python 3.8+, ничего устанавливать не нужно.
"""

import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

# ---------- Настройки ----------
COMPLEX_SLUG = "varshavskie-vorota"
BASE = "https://rg-dev.ru"
PAGE_URL = f"{BASE}/parking/{COMPLEX_SLUG}/"
SALES_PHONE = "+7 (495) 106-90-57"

STATE_FILE = os.environ.get("STATE_FILE", "state.json")
TG_TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
TG_CHAT = os.environ.get("TELEGRAM_CHAT_ID", "").strip()

MSK = timezone(timedelta(hours=3))
DAILY_REPORT_HOUR = 10   # ежедневная сводка «бот жив» — после 10:00 МСК, раз в день
FAIL_ALERT_AFTER = 3     # предупредить, если сайт не отвечает N запусков подряд

STATUS_NAMES = {1: "свободно", 2: "забронировано"}
FREE = 1
BOOKED = 2
# --------------------------------


def log(msg):
    print(f"[{datetime.now(MSK):%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


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


def to_int(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


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


# ---------- Оформление сообщений ----------

def fmt_price(n):
    return f"{n:,}".replace(",", " ") + " ₽" if n else "цена не указана"


def status_name(s):
    return STATUS_NAMES.get(s, f"статус {s}")


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
    return (p.get("floor") or 0, p.get("number"))


def summary(places):
    free = sorted((p for p in places.values() if p["status"] == FREE), key=sort_key)
    booked = [p for p in places.values() if p["status"] == BOOKED]
    lines = []
    if free:
        lines.append(f"Свободно сейчас ({len(free)}):")
        lines += [f"  • {place_line(p)}" for p in free]
    else:
        lines.append("Свободных мест сейчас нет.")
    lines.append(f"Забронировано: {len(booked)}")
    other = len(places) - len(free) - len(booked)
    if other:
        lines.append(f"С другим статусом: {other}")
    lines.append("")
    lines.append(f"Карта: {PAGE_URL}")
    lines.append(f"Отдел продаж: {SALES_PHONE}")
    return lines


def find_changes(old, new):
    events = []
    for pid, p in new.items():
        o = old.get(pid)
        if o is None:
            if p["status"] == FREE:
                events.append(("🟢", "Новое место в продаже", p))
            else:
                events.append(("🟡", f"Новое место в списке, {status_name(p['status'])}", p))
            continue
        if o["status"] != p["status"]:
            if p["status"] == FREE:
                events.append(("🟢", "Снова свободно", p))
            elif p["status"] == BOOKED:
                events.append(("🟡", "Забронировано", p))
            else:
                events.append(("⚪", f"Статус: {status_name(o['status'])} → {status_name(p['status'])}", p))
        if o.get("price") and p.get("price") and o["price"] != p["price"]:
            events.append(("💰", f"Цена: {fmt_price(o['price'])} → {fmt_price(p['price'])}", p))
    for pid, o in old.items():
        if pid not in new:
            events.append(("🔴", "Пропало из списка — продано или снято с продажи", o))
    # Самое важное (свободные) — первым
    order = {"🟢": 0, "💰": 1, "🟡": 2, "🔴": 3, "⚪": 4}
    events.sort(key=lambda e: order.get(e[0], 9))
    return events


def build_message(title, events, places):
    lines = [title, ""]
    for icon, text, p in events:
        lines.append(f"{icon} {text}: {place_line(p)}")
    if events:
        lines.append("")
    lines += summary(places)
    return "\n".join(lines)


# ---------- Telegram ----------

def send_telegram(text):
    if not TG_TOKEN or not TG_CHAT:
        log("Telegram не настроен (нет TELEGRAM_TOKEN / TELEGRAM_CHAT_ID). Сообщение:\n" + text)
        return True
    if len(text) > 4000:
        text = text[:3990] + "\n…"
    data = urllib.parse.urlencode({
        "chat_id": TG_CHAT,
        "text": text,
        "disable_web_page_preview": "true",
    }).encode()
    try:
        with urllib.request.urlopen(
            f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage", data=data, timeout=30
        ) as resp:
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
        log(f"Ошибка отправки в Telegram: HTTP {e.code} {reason}{hint}")
        return False
    except Exception as e:
        log(f"Ошибка отправки в Telegram: {e}")
        return False
    if not ok:
        log("Telegram отклонил сообщение — проверьте токен и chat_id.")
    return ok


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

    state = load_state()
    now = datetime.now(MSK)
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

    if state.get("fails", 0) >= FAIL_ALERT_AFTER:
        send_telegram("✅ Связь с сайтом восстановлена, мониторинг работает.")
    state["fails"] = 0

    message = None
    if old_places is None:
        message = build_message("🅿️ Мониторинг машиномест «Варшавские ворота» запущен", [], places)
    else:
        events = find_changes(old_places, places)
        if events:
            message = build_message("🅿️ Изменения на паркинге «Варшавские ворота»", events, places)
        elif now.hour >= DAILY_REPORT_HOUR and state.get("last_report_date") != today:
            message = build_message("☀️ Ежедневная сводка: изменений нет, бот работает", [], places)

    if message:
        if not send_telegram(message):
            # Не запоминаем новое состояние — в следующий раз попробуем отправить снова
            save_state(state)
            return 1
        state["last_report_date"] = today
        log("Сообщение отправлено.")
    else:
        log(f"Изменений нет. Мест в списке: {len(places)}, "
            f"свободно: {sum(p['status'] == FREE for p in places.values())}.")

    state["places"] = places
    save_state(state)
    return 0


if __name__ == "__main__":
    sys.exit(main())
