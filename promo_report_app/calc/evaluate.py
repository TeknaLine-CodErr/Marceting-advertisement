# -*- coding: utf-8 -*-
"""Сведение каталога ЦО и файла акции в итоговый справочный отчёт по товарам.

Правила (заданы пользователем):
  1. Если цена по акции (Рассчитанная цена для участия, N) >= нашей цены (K из ЦО) —
     товар участвует в акции по своей цене (снижать не нужно). Статус: зелёный.
  2. Если цена по акции < нашей цены — товар пришлось бы продавать по цене акции.
     Считаем полный П/У по этой цене (расчёт из calc.pricing) и смотрим % от себестоимости:
       >= green_min      -> зелёный (одобрен)
       [red_max, green_min) -> жёлтый (пограничный, на проверку менеджеру)
       < red_max          -> красный (не одобряем)
"""

from dataclasses import dataclass

from calc import pricing
from config import (
    DEFAULT_RED_MAX,
    DEFAULT_GREEN_MIN,
    DEFAULT_STOCK_DAYS_HIGH,
    DEFAULT_LOW_SALES_PER_WEEK,
)


@dataclass
class Thresholds:
    red_max: float = DEFAULT_RED_MAX
    green_min: float = DEFAULT_GREEN_MIN
    stock_days_high: float = DEFAULT_STOCK_DAYS_HIGH
    low_sales_per_week: float = DEFAULT_LOW_SALES_PER_WEEK


STATUS_LABELS = {
    "green": "Одобрен",
    "yellow": "Пограничный — на проверку",
    "red": "Отклонён",
    "no_co": "Нет данных в ЦО",
    "no_price_data": "Нет цены в акции",
    "error": "Ошибка расчёта",
}

STOCK_RECOMMENDATION_LABELS = {
    "high_stock": "Много остатков",
    "low_sales": "Низкие продажи",
    "normal": "Оборачиваемость в норме",
}


def _is_positive_number(v):
    return isinstance(v, (int, float)) and v > 0


def _num_or_none(v):
    return v if isinstance(v, (int, float)) else None


STOCK_ADD_LABELS = {True: "Да", False: "Нет"}


def stock_recommendation(ordered_qty_7d, stock_ozon, stock_own, thresholds):
    """Справочная (не финансовая) подсказка по остаткам/продажам за 7 дней. Возвращает
    (tag, add_flag, metric_comment):
      - остатков нет вообще / нет данных о продажах -> (None, None, "") — подсказки нет
      - "дни запаса" = суммарный остаток / средний расход в день высокие -> ("high_stock", True, "...")
      - заказов за неделю мало (в т.ч. 0) -> ("low_sales", True, "...")
      - иначе -> ("normal", False, "") — оборачиваемость в норме, добавлять не обязательно
    Не влияет на статус одобрения по марже — только справочная колонка."""
    stock_ozon_n = _num_or_none(stock_ozon) or 0
    stock_own_n = _num_or_none(stock_own) or 0
    total_stock = stock_ozon_n + stock_own_n
    ordered = _num_or_none(ordered_qty_7d)

    if total_stock <= 0 or ordered is None:
        return None, None, ""

    avg_daily = ordered / 7
    days_of_stock = (total_stock / avg_daily) if avg_daily > 0 else None

    if days_of_stock is not None and days_of_stock >= thresholds.stock_days_high:
        return "high_stock", True, f"~{days_of_stock:.0f} дн. запаса при {ordered:.0f} шт/нед"
    if ordered <= thresholds.low_sales_per_week:
        return "low_sales", True, f"{ordered:.0f} шт/нед при остатке {total_stock:.0f} шт"
    return "normal", False, ""


def build_report(co_catalog, promo_catalog, thresholds=None):
    thresholds = thresholds or Thresholds()
    report = []

    for article, promo in promo_catalog.items():
        co = co_catalog.get(article)

        row = {
            "article": article,
            "name": (promo.get("name") or (co.get("name") if co else None)),
            "category": (promo.get("category") or (co.get("category") if co else None)),
            "scheme": co["scheme"] if co else None,
            "source_file": co["source_file"] if co else None,
            "price_co": co["price"] if co else None,
            "price_promo_max": promo.get("calculated_max_price"),
            "min_price": promo.get("min_price"),
            "price_used": None,
            "pct_of_cost": None,
            "net_profit": None,
            "status": None,
            "status_label": None,
            "reason": None,
            "price_change_needed": False,
            "price_change_note": "",
            "ordered_qty_7d": promo.get("ordered_qty_7d"),
            "stock_ozon": promo.get("stock_ozon"),
            "stock_own": promo.get("stock_own"),
        }

        stock_tag, stock_note = stock_recommendation(
            row["ordered_qty_7d"], row["stock_ozon"], row["stock_own"], thresholds
        )
        row["stock_recommendation"] = stock_tag
        row["stock_recommendation_label"] = STOCK_RECOMMENDATION_LABELS.get(stock_tag, "")
        row["stock_note"] = stock_note

        if co is None:
            row["status"] = "no_co"
            row["reason"] = "Артикул не найден ни в одном загруженном файле ЦО"
            row["status_label"] = STATUS_LABELS[row["status"]]
            report.append(row)
            continue

        price_co = co["price"]
        price_promo_max = promo.get("calculated_max_price")

        if not _is_positive_number(price_co):
            row["status"] = "no_co"
            row["reason"] = "В ЦО не указана корректная цена товара"
            row["status_label"] = STATUS_LABELS[row["status"]]
            report.append(row)
            continue

        if not _is_positive_number(price_promo_max):
            row["status"] = "no_price_data"
            row["reason"] = "В файле акции не рассчитана максимальная цена участия"
            row["status_label"] = STATUS_LABELS[row["status"]]
            report.append(row)
            continue

        own_price_mode = price_promo_max >= price_co
        used_price = price_co if own_price_mode else price_promo_max

        calc = pricing.calculate(co, used_price)
        row["price_used"] = used_price

        if calc is None:
            row["status"] = "error"
            row["reason"] = "Не удалось определить ставку комиссии Ozon для категории товара"
            row["status_label"] = STATUS_LABELS[row["status"]]
            report.append(row)
            continue

        pct = calc["pct_of_cost"]
        row["pct_of_cost"] = pct
        row["net_profit"] = calc["net_profit"]
        row["calc"] = calc

        if own_price_mode:
            row["status"] = "green"
            row["reason"] = "Цена акции не ниже нашей цены — участвуем по своей цене без снижения"
        elif pct is None:
            row["status"] = "error"
            row["reason"] = "Не удалось посчитать % от себестоимости (нулевая себестоимость в ЦО)"
        elif pct >= thresholds.green_min:
            row["status"] = "green"
            row["reason"] = (
                f"Цену нужно снизить до {used_price:.0f} ₽, маржа {pct:.0%} от себестоимости — приемлемо"
            )
        elif pct >= thresholds.red_max:
            row["status"] = "yellow"
            row["reason"] = (
                f"Цену нужно снизить до {used_price:.0f} ₽, маржа {pct:.0%} от себестоимости — "
                f"пограничная зона, на проверку менеджеру"
            )
        else:
            row["status"] = "red"
            row["reason"] = (
                f"Цену нужно снизить до {used_price:.0f} ₽, маржа {pct:.0%} от себестоимости — не одобряем"
            )

        # Товар одобрен для участия (зелёный/жёлтый) только по сниженной цене -> в ЦО нужно
        # вручную поменять "Наша Цена" на цену акции, иначе там останется старая, более высокая
        if row["status"] in ("green", "yellow") and not own_price_mode:
            row["price_change_needed"] = True
            row["price_change_note"] = f"Да: {price_co:.0f} → {used_price:.0f} ₽"
        else:
            row["price_change_needed"] = False
            row["price_change_note"] = ""

        row["status_label"] = STATUS_LABELS[row["status"]]
        report.append(row)

    return report
