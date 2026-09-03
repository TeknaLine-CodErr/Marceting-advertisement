# -*- coding: utf-8 -*-
"""Пересчёт полного расчёта затрат из ЦО по произвольной цене (см. wb_formulas2.txt).

Логика 1-в-1 повторяет формулы листов СП / FBS / Неликвид:
  Комиссия Ozon в руб.      = Цена * ставка_комиссии(категория, Цена)   [ставка по тарифной сетке]
  Эквайринг, руб.           = Цена * Q1
  Доставка до места выдачи  = MIN(Цена * W1, 25)
  Общая логистика           = Базовая логистика + Обратная логистика + Отправление ППЗ/ПВЗ + Доставка
  Непредвиденные расходы    = Цена * AD1
  ИТОГО РАСХОДОВ            = Общий себес + Комиссия + Эквайринг + Приемка + Общая логистика
                               + Кросс-докинг + Транспорт + Пакет + Коробка + Непредвиденные
                               + Отзывы + Хранение + Реклама на выкуп
  Налоги                    = Цена * AM1 * 0.8
  Чистая прибыль            = Цена - ИТОГО РАСХОДОВ - Налоги + Скидка_за_локальные * %выкупа
                               (на листах без локальной скидки это слагаемое из ЦО всегда 0,
                               поэтому его можно прибавлять универсально для любой схемы)
  % от себестоимости        = Чистая прибыль / (Себестоимость + Себестоимость крепежа)
"""


def _num(v, default=0):
    return v if isinstance(v, (int, float)) else default


def commission_rate_for_price(category_rates, commission_label, price, fallback_rate):
    """Возвращает (ставка, use_fallback). Тарифная сетка берётся из листа Категории того же
    файла ЦО, откуда пришёл товар, по набору колонок (commission_label = "FBO"/"FBS"),
    заранее определённому для этого листа-схемы в readers.co_reader. Если сетка недоступна
    или сетка/категория не найдены — используем закэшированную в ЦО ставку как приближение."""
    triples = (category_rates or {}).get(commission_label) if commission_label else None
    if triples:
        for low, high, rate in triples:
            if low is None or high is None or rate is None:
                continue
            if low <= price <= high:
                return rate, False
    return fallback_rate, True


def calculate(row, price):
    """row: одна запись каталога ЦО (см. readers.co_reader). price: цена, по которой считаем П/У.

    Возвращает dict с расчётом или None, если по товару нет ставки комиссии вообще
    (ни по сетке, ни в кэше ЦО)."""
    rate, rate_is_fallback = commission_rate_for_price(
        row.get("category_rates"), row.get("commission_label"), price, row.get("commission_rate")
    )
    if not isinstance(rate, (int, float)):
        return None

    consts = row["row1_consts"]
    cost_total = _num(row.get("cost_total"))
    cost_base = _num(row.get("cost_base"))
    cost_fastener = _num(row.get("cost_fastener"))
    logistics_base = _num(row.get("logistics_base"))
    logistics_return = _num(row.get("logistics_return"))
    local_adjustment = _num(row.get("local_adjustment"))
    shipment_fee = _num(row.get("shipment_fee"))
    acceptance = _num(row.get("acceptance"))
    cross_docking = _num(row.get("cross_docking"))
    transport = _num(row.get("transport"))
    package = _num(row.get("package"))
    box = _num(row.get("box"))
    reviews = _num(row.get("reviews"))
    storage = _num(row.get("storage"))
    buyout_pct = _num(row.get("buyout_pct"))
    ad_cost_buyout = _num(row.get("ad_cost_per_buyout"))

    acquiring_rate = _num(consts.get("acquiring_rate"))
    delivery_rate = _num(consts.get("delivery_rate"))
    contingency_rate = _num(consts.get("contingency_rate"))
    tax_rate = _num(consts.get("tax_rate"))

    commission_rub = price * rate
    acquiring = price * acquiring_rate
    delivery = min(price * delivery_rate, 25)
    logistics_total = logistics_base + logistics_return + shipment_fee + delivery
    contingency = price * contingency_rate

    total_costs = (
        cost_total + commission_rub + acquiring + acceptance + logistics_total
        + cross_docking + transport + package + box + contingency + reviews + storage
        + ad_cost_buyout
    )
    taxes = price * tax_rate * 0.8

    net_profit = price - total_costs - taxes + local_adjustment * buyout_pct

    cost_base_full = cost_base + cost_fastener
    pct_of_cost = net_profit / cost_base_full if cost_base_full else None
    pct_of_revenue = net_profit / price if price else None
    marketplace_share = (total_costs - cost_base - cross_docking) / price if price else None

    return {
        "commission_rate": rate,
        "commission_rate_is_fallback": rate_is_fallback,
        "commission_rub": commission_rub,
        "acquiring": acquiring,
        "delivery": delivery,
        "logistics_total": logistics_total,
        "contingency": contingency,
        "taxes": taxes,
        "total_costs": total_costs,
        "net_profit": net_profit,
        "pct_of_cost": pct_of_cost,
        "pct_of_revenue": pct_of_revenue,
        "marketplace_share": marketplace_share,
        "price_used": price,
    }
