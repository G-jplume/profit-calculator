"""
跨境电商新品利润核算计算引擎
基于原始 Excel 利润核算表的计算逻辑，支持 AMZ FBM / AMZ FBA / WF 三大业务线。
"""
import math
from dataclasses import dataclass, field
from typing import Optional


# ============================================================
# FedEx 基础运费表 (折扣结算价, Zone 2-8)
# 数据来源: 良仓-基础运费fedex homedelivery Sheet
# ============================================================
# 格式: {重量(lb): {zone: 价格}}
# 简化: 各仓结算价基本一致，使用统一折扣率 0.35 (4折)
def _build_fedex_base_rates():
    raw = {
        1: [11.32, 11.65, 12.69, 13.27, 13.72, 13.93, 14.17],
        2: [12.08, 13.32, 14.53, 14.84, 15.49, 16.13, 16.40],
        3: [12.57, 14.03, 15.16, 15.93, 16.58, 17.20, 18.05],
        4: [12.94, 14.11, 15.78, 16.81, 17.27, 18.47, 19.34],
        5: [13.29, 14.76, 16.17, 17.56, 18.26, 19.33, 20.48],
        6: [13.39, 14.81, 16.33, 17.63, 18.28, 19.34, 20.49],
        7: [14.15, 15.15, 16.76, 18.19, 18.61, 19.86, 21.26],
        8: [14.56, 15.70, 17.34, 18.71, 19.36, 20.66, 22.19],
        9: [14.78, 15.92, 17.42, 18.87, 19.77, 21.52, 23.37],
        10: [14.98, 16.09, 17.59, 19.39, 20.00, 22.54, 24.91],
        11: [15.92, 16.56, 18.17, 19.71, 20.91, 24.66, 26.79],
        12: [16.18, 17.26, 18.35, 19.99, 21.64, 25.65, 28.04],
        13: [16.19, 17.34, 18.56, 20.37, 22.31, 27.22, 29.31],
        14: [16.86, 17.81, 18.75, 20.87, 23.62, 29.16, 31.87],
        15: [17.35, 18.31, 19.30, 21.55, 24.48, 30.31, 33.32],
        20: [19.25, 20.45, 21.72, 24.50, 28.10, 35.10, 39.05],
        25: [21.15, 22.60, 24.15, 27.46, 31.72, 39.90, 44.79],
        30: [23.06, 24.75, 26.58, 30.42, 35.34, 44.69, 50.53],
        35: [24.96, 26.90, 29.00, 33.38, 38.96, 49.49, 56.26],
        40: [26.87, 29.04, 31.43, 36.34, 42.58, 54.29, 62.00],
        45: [28.77, 31.19, 33.86, 39.30, 46.20, 59.09, 67.74],
        50: [30.68, 33.34, 36.29, 42.26, 49.82, 63.89, 73.47],
        55: [32.58, 35.49, 38.72, 45.22, 53.44, 68.69, 79.21],
        60: [34.49, 37.63, 41.15, 48.18, 57.06, 73.49, 84.95],
        65: [36.39, 39.78, 43.58, 51.14, 60.68, 78.29, 90.69],
        70: [38.30, 41.93, 46.01, 54.10, 64.30, 83.09, 96.42],
        75: [40.20, 44.08, 48.44, 57.06, 67.92, 87.89, 102.16],
        80: [42.11, 46.22, 50.87, 60.02, 71.54, 92.69, 107.90],
        85: [44.01, 48.37, 53.30, 62.98, 75.16, 97.49, 113.63],
        90: [45.92, 50.52, 55.73, 65.94, 78.78, 102.29, 119.37],
        95: [47.82, 52.66, 58.16, 68.90, 82.40, 107.09, 125.11],
        100: [49.73, 54.81, 60.59, 71.86, 86.02, 111.89, 130.84],
        110: [53.54, 59.07, 65.45, 77.78, 93.26, 121.49, 142.32],
        120: [57.36, 63.34, 70.31, 83.70, 100.50, 131.09, 153.80],
        130: [61.17, 67.60, 75.17, 89.62, 107.74, 140.69, 165.28],
        140: [64.99, 71.86, 80.03, 95.54, 114.98, 150.29, 176.75],
        150: [68.80, 76.13, 84.89, 101.46, 122.22, 159.89, 188.23],
    }
    discount = 0.35
    result = {}
    for lb, prices in raw.items():
        zones = {}
        for i, p in enumerate(prices):
            zones[i + 2] = round(p * discount, 2)
        result[lb] = zones
    return result


FEDEX_BASE_RATES = _build_fedex_base_rates()


def get_fedex_base_rate(weight_lb: float, zone: int) -> float:
    weight = int(math.ceil(weight_lb))
    if weight > 150:
        weight = 150
    if weight < 1:
        weight = 1
    key = weight
    while key not in FEDEX_BASE_RATES and key > 1:
        key -= 1
    if key not in FEDEX_BASE_RATES:
        return 0.0
    return FEDEX_BASE_RATES[key].get(zone, FEDEX_BASE_RATES[key].get(5, 0.0))


# ============================================================
# 出库费 (各仓通用 - 按重量分级)
# 数据来源: 良仓-出库费 Sheet
# ============================================================
OUTBOUND_FEE = {
    "CLASS_A": {"max_lb": 4.4, "fee": 1.0},
    "CLASS_B": {"max_lb": 22, "fee": 1.8},
    "CLASS_C": {"max_lb": 50, "fee": 2.5},
    "CLASS_D": {"max_lb": 70, "fee": 2.5},
    "CLASS_E": {"max_lb": 100, "fee": 3.0},
    "CLASS_F": {"max_lb": 150, "fee": 4.0},
    "CLASS_G": {"max_lb": 200, "fee": 7.5},
    "CLASS_H": {"max_lb": 9999, "fee": 12.0},
}

SELF_LABEL_FEE = 2.0


def get_outbound_fee(weight_lb: float) -> float:
    for cls in OUTBOUND_FEE.values():
        if weight_lb <= cls["max_lb"]:
            return cls["fee"]
    return 12.0


# ============================================================
# FedEx 附加费 (2026年1月更新)
# 数据来源: 良仓-附加费 Home Delivery Sheet
# ============================================================
SURCHARGES = {
    "residential": 2.71,
    "signature": 3.60,
    "fuel_rate": 0.175,
    "ahs_dimension": {2: 4.57, "3-4": 5.07, "5-6": 5.89, "7-8": 6.30},
    "ahs_weight": {2: 7.12, "3-4": 7.77, "5-6": 8.64, "7-8": 9.08},
    "ahs_packaging": {2: 4.10, "3-4": 4.75, "5-6": 5.40, "7-8": 6.05},
    "oversize": {2: 33.80, "3-4": 33.80, "5-6": 33.80, "7-8": 33.80},
}

# 旺季附加费 (Demand Surcharge)
PEAK_SURCHARGES = {
    "off_season": {
        "ahs": 1.38,
        "oversize": 15.80,
        "unauthorized": 0,
        "ground_home": 0,
    },
    "peak_season": {
        "ahs": 5.00,
        "oversize": 50.00,
        "unauthorized": 500.00,
        "ground_home": 0.55,
    },
}


def get_ahs_dimension_fee(zone: int) -> float:
    if zone == 2:
        return SURCHARGES["ahs_dimension"][2]
    elif zone in (3, 4):
        return SURCHARGES["ahs_dimension"]["3-4"]
    elif zone in (5, 6):
        return SURCHARGES["ahs_dimension"]["5-6"]
    else:
        return SURCHARGES["ahs_dimension"]["7-8"]


def get_ahs_weight_fee(zone: int) -> float:
    if zone == 2:
        return SURCHARGES["ahs_weight"][2]
    elif zone in (3, 4):
        return SURCHARGES["ahs_weight"]["3-4"]
    elif zone in (5, 6):
        return SURCHARGES["ahs_weight"]["5-6"]
    else:
        return SURCHARGES["ahs_weight"]["7-8"]


# ============================================================
# CG仓拣货和发货价格
# 数据来源: CG仓拣货和发货价格 Sheet
# ============================================================
CG_RATES = {
    "Bin Small": {"base": 12.35, "per_lb": 0.17},
    "Bin Large": {"base": 15.10, "per_lb": 0.17},
    "Bin Heavy": {"base": 15.42, "per_lb": 0.17},
    "Standard Small": {"base": 17.05, "per_lb": 0.20},
    "Standard Medium": {"base": 23.78, "per_lb": 0.24},
    "Standard Large": {"base": 49.83, "per_lb": 0.31},
    "Standard Oversize": {"base": 52.43, "per_lb": 0.34},
}


# ============================================================
# 6大海外仓名称
# ============================================================
WAREHOUSES = ["良仓", "大方广", "诺一奥", "无忧达", "派速捷", "优派"]


# ============================================================
# 数据类
# ============================================================
@dataclass
class ProductInput:
    name: str = ""
    origin: str = "中国"
    zone: int = 5
    box_count: int = 1
    price: float = 159.99
    purchase_price_cny: float = 450.0
    length_cm: float = 170.0
    width_cm: float = 50.0
    height_cm: float = 12.0
    weight_kg: float = 36.0
    exchange_rate: float = 7.0


@dataclass
class CostBreakdown:
    purchase_cost: float = 0.0
    shipping_cost: float = 0.0
    logistics_cost: float = 0.0
    commission: float = 0.0
    ad_cost: float = 0.0
    return_cost: float = 0.0
    storage_cost: float = 0.0

    @property
    def total(self) -> float:
        return (
            self.purchase_cost
            + self.shipping_cost
            + self.logistics_cost
            + self.commission
            + self.ad_cost
            + self.return_cost
            + self.storage_cost
        )


# ============================================================
# 尺寸转换和判定逻辑
# ============================================================
def cm_to_inch(cm: float) -> float:
    return cm * 0.3937


def calc_girth_inch(length_in: float, width_in: float, height_in: float) -> float:
    sides = sorted([length_in, width_in, height_in], reverse=True)
    longest = math.ceil(sides[0])
    second = math.ceil(sides[1])
    shortest = math.ceil(sides[2])
    return (longest + second * 2 + shortest * 2) - longest


def calc_weight_lbs(weight_kg: float) -> float:
    return weight_kg * 2.20462


def calc_volume_weight_lbs(l_in: float, w_in: float, h_in: float) -> float:
    return (l_in * w_in * h_in) / 139


def calc_chargeable_weight(weight_kg: float, l_cm: float, w_cm: float, h_cm: float) -> float:
    l_in = cm_to_inch(l_cm)
    w_in = cm_to_inch(w_cm)
    h_in = cm_to_inch(h_cm)
    actual_lbs = calc_weight_lbs(weight_kg)
    vol_lbs = calc_volume_weight_lbs(l_in, w_in, h_in)
    return max(math.ceil(actual_lbs), math.ceil(vol_lbs))


def determine_package_size(l_in: float, w_in: float, h_in: float, weight_lbs: float) -> str:
    longest = math.ceil(max(l_in, w_in, h_in))
    sides = [math.ceil(l_in), math.ceil(w_in), math.ceil(h_in)]
    sides_sorted = sorted(sides, reverse=True)
    girth = (sides_sorted[0] + sides_sorted[1] * 2 + sides_sorted[2] * 2)

    if weight_lbs <= 1 and longest <= 15 and sides_sorted[1] <= 12 and sides_sorted[2] <= 0.75:
        return "小号标准尺寸"
    elif weight_lbs <= 20 and longest <= 18 and sides_sorted[1] <= 14 and sides_sorted[2] <= 8:
        return "大号标准尺寸"
    elif longest <= 96 and girth <= 130:
        return "大号大件"
    else:
        return "超大尺寸"


def determine_cg_category(l_in: float, w_in: float, h_in: float, weight_lbs: float) -> str:
    longest = math.ceil(max(l_in, w_in, h_in))
    sides_sorted = sorted([math.ceil(l_in), math.ceil(w_in), math.ceil(h_in)], reverse=True)
    girth = sides_sorted[0] + 2 * (sides_sorted[1] + sides_sorted[2])

    if longest <= 19 and sides_sorted[1] <= 12 and sides_sorted[2] <= 6 and weight_lbs <= 25:
        return "Bin Small"
    elif longest <= 26 and sides_sorted[1] <= 17 and sides_sorted[2] <= 14 and weight_lbs <= 25:
        return "Bin Large"
    elif longest <= 26 and sides_sorted[1] <= 17 and sides_sorted[2] <= 14 and weight_lbs > 25:
        return "Bin Heavy"
    elif longest <= 108 and girth <= 130 and weight_lbs <= 70:
        return "Standard Small"
    elif longest <= 108 and girth <= 165 and weight_lbs <= 150:
        return "Standard Medium"
    elif longest <= 108 and girth <= 165 and weight_lbs <= 150:
        return "Standard Large"
    else:
        return "Standard Oversize"


# ============================================================
# 物流费计算 (单仓)
# ============================================================
def calc_logistics_cost(
    weight_lbs: float,
    l_cm: float,
    w_cm: float,
    h_cm: float,
    zone: int,
    is_peak: bool = False,
) -> dict:
    l_in = cm_to_inch(l_cm)
    w_in = cm_to_inch(w_cm)
    h_in = cm_to_inch(h_cm)

    chargeable_weight = max(
        math.ceil(calc_weight_lbs(weight_lbs / 2.20462)),
        math.ceil(calc_volume_weight_lbs(l_in, w_in, h_in)),
    )

    # 基础运费
    base_rate = get_fedex_base_rate(chargeable_weight, zone)

    # 燃油附加费
    fuel_surcharge = base_rate * SURCHARGES["fuel_rate"]

    # 住宅附加费
    residential = SURCHARGES["residential"]

    # AHS判定
    longest_in = math.ceil(max(l_in, w_in, h_in))
    girth = calc_girth_inch(l_in, w_in, h_in)
    ahs_dim = 0
    ahs_wt = 0
    if longest_in > 48 or girth > 105:
        ahs_dim = get_ahs_dimension_fee(zone)
    if chargeable_weight > 50:
        ahs_wt = get_ahs_weight_fee(zone)

    # Oversize判定
    oversize = 0
    if longest_in > 96 or girth > 128:
        oversize = SURCHARGES["oversize"].get(zone, 33.80)

    # 旺季附加费
    peak = PEAK_SURCHARGES["peak_season" if is_peak else "off_season"]
    peak_ahs = peak["ahs"]
    peak_oversize = peak["oversize"] if oversize > 0 else 0
    peak_ground = peak["ground_home"]

    # 出库费
    outbound = get_outbound_fee(chargeable_weight)
    self_label = SELF_LABEL_FEE

    total = (
        base_rate
        + fuel_surcharge
        + residential
        + ahs_dim
        + ahs_wt
        + oversize
        + (peak_ahs if ahs_dim > 0 or ahs_wt > 0 else 0)
        + (peak_oversize)
        + peak_ground
        + outbound
        + self_label
    )

    return {
        "base_rate": round(base_rate, 2),
        "fuel_surcharge": round(fuel_surcharge, 2),
        "residential": round(residential, 2),
        "ahs_dimension": round(ahs_dim, 2),
        "ahs_weight": round(ahs_wt, 2),
        "oversize": round(oversize, 2),
        "peak_surcharge": round(peak_ahs + peak_oversize + peak_ground, 2),
        "outbound_fee": round(outbound, 2),
        "self_label_fee": round(self_label, 2),
        "total": round(total, 2),
        "chargeable_weight": chargeable_weight,
    }


# ============================================================
# 头程计算
# ============================================================
def calc_shipping_cost(
    origin: str,
    l_cm: float,
    w_cm: float,
    h_cm: float,
    weight_kg: float,
    box_count: int,
    exchange_rate: float = 7.0,
) -> dict:
    box_volume = (l_cm * w_cm * h_cm) / 1_000_000
    total_volume = box_volume * box_count
    total_weight = weight_kg * box_count

    # 装柜量 (67方/柜 或 19500kg/柜)
    vol_per_container = 67 / total_volume if total_volume > 0 else 999
    wt_per_container = 19500 / total_weight if total_weight > 0 else 999
    actual_shipment = max(1, int(min(vol_per_container, wt_per_container)))

    # 各产地费用
    if origin == "中国":
        domestic_truck = 5000 / exchange_rate
        sea_freight = 2500
        customs_tariff = 6840
        us_truck = 1500
        total_head = domestic_truck + sea_freight + customs_tariff + us_truck
    elif origin in ("越南", "泰国", "越南/泰国"):
        vn_truck = 0
        vn_local = 690
        sea_freight = 2500
        customs_tariff = 6840
        us_truck = 1500
        total_head = vn_truck + vn_local + sea_freight + customs_tariff + us_truck
    elif origin == "马来西亚":
        my_truck = 0
        customs_clearance = 115
        sea_freight = 2500
        customs_tariff = 6840
        us_truck = 1500
        total_head = my_truck + customs_clearance + sea_freight + customs_tariff + us_truck
    else:
        domestic_truck = 5000 / exchange_rate
        sea_freight = 2500
        customs_tariff = 6840
        us_truck = 1500
        total_head = domestic_truck + sea_freight + customs_tariff + us_truck

    per_unit_head = total_head / actual_shipment if actual_shipment > 0 else total_head

    return {
        "box_volume_m3": round(box_volume, 4),
        "total_volume_m3": round(total_volume, 4),
        "total_weight_kg": round(total_weight, 2),
        "container_qty": actual_shipment,
        "total_head_cost": round(total_head, 2),
        "per_unit_head_cost": round(per_unit_head, 2),
    }


# ============================================================
# 佣金计算
# ============================================================
def calc_commission(price: float) -> float:
    if price < 200:
        return round(price * 0.15, 2)
    else:
        return round(200 * 0.15 + (price - 200) * 0.10, 2)


# ============================================================
# 完整利润计算 (AMZ FBM)
# ============================================================
def calc_amz_fbm_profit(p: ProductInput) -> dict:
    l_in = cm_to_inch(p.length_cm)
    w_in = cm_to_inch(p.width_cm)
    h_in = cm_to_inch(p.height_cm)

    if p.origin == "中国":
        purchase_cost = p.purchase_price_cny / p.exchange_rate
    else:
        purchase_cost = p.purchase_price_cny * 1.1

    purchase_with_tax = purchase_cost * 1.1

    ship = calc_shipping_cost(
        p.origin, p.length_cm, p.width_cm, p.height_cm, p.weight_kg, p.box_count, p.exchange_rate
    )
    head_cost_per_unit = ship["per_unit_head_cost"]

    # 6仓比价 (淡季)
    wh_results_off = {}
    for wh in WAREHOUSES:
        logi = calc_logistics_cost(
            p.weight_kg * 2.20462, p.length_cm, p.width_cm, p.height_cm, p.zone, is_peak=False
        )
        wh_results_off[wh] = logi

    # 6仓比价 (旺季)
    wh_results_peak = {}
    for wh in WAREHOUSES:
        logi = calc_logistics_cost(
            p.weight_kg * 2.20462, p.length_cm, p.width_cm, p.height_cm, p.zone, is_peak=True
        )
        wh_results_peak[wh] = logi

    best_wh_off = min(wh_results_off, key=lambda k: wh_results_off[k]["total"])
    best_wh_peak = min(wh_results_peak, key=lambda k: wh_results_peak[k]["total"])

    logistics_off = wh_results_off[best_wh_off]["total"]
    logistics_peak = wh_results_peak[best_wh_peak]["total"]

    commission = calc_commission(p.price)
    ad_cost = round(p.price * 0.08, 2)
    return_cost = round(p.price * 0.03, 2)
    storage_cost = round(p.price * 0.01, 2)

    cost_off = CostBreakdown(
        purchase_cost=round(purchase_with_tax, 2),
        shipping_cost=round(head_cost_per_unit, 2),
        logistics_cost=round(logistics_off, 2),
        commission=round(commission, 2),
        ad_cost=ad_cost,
        return_cost=return_cost,
        storage_cost=storage_cost,
    )

    cost_peak = CostBreakdown(
        purchase_cost=round(purchase_with_tax, 2),
        shipping_cost=round(head_cost_per_unit, 2),
        logistics_cost=round(logistics_peak, 2),
        commission=round(commission, 2),
        ad_cost=ad_cost,
        return_cost=return_cost,
        storage_cost=storage_cost,
    )

    profit_off = round(p.price - cost_off.total, 2)
    profit_peak = round(p.price - cost_peak.total, 2)
    margin_off = round(profit_off / p.price * 100, 1)
    margin_peak = round(profit_peak / p.price * 100, 1)

    return {
        "business_type": "AMZ FBM",
        "cost_off_season": cost_off,
        "cost_peak_season": cost_peak,
        "profit_off_season": profit_off,
        "profit_peak_season": profit_peak,
        "margin_off_season": margin_off,
        "margin_peak_season": margin_peak,
        "best_warehouse_off": best_wh_off,
        "best_warehouse_peak": best_wh_peak,
        "warehouse_comparison_off": wh_results_off,
        "warehouse_comparison_peak": wh_results_peak,
        "shipping_info": ship,
        "ad_return_storage_ratio": round((0.08 + 0.03 + 0.01) * 100, 1),
    }


# ============================================================
# 完整利润计算 (AMZ FBA)
# ============================================================
def calc_amz_fba_profit(p: ProductInput) -> dict:
    l_in = cm_to_inch(p.length_cm)
    w_in = cm_to_inch(p.width_cm)
    h_in = cm_to_inch(p.height_cm)

    purchase_cost = p.purchase_price_cny / p.exchange_rate

    # FBA尺寸分段
    pkg_size = determine_package_size(l_in, w_in, h_in, calc_weight_lbs(p.weight_kg))
    chargeable_weight = calc_chargeable_weight(p.weight_kg, p.length_cm, p.width_cm, p.height_cm)

    # FBA物流费 (简化费率表)
    fba_fee_map = {
        "小号标准尺寸": 3.87,
        "大号标准尺寸": 6.92 + max(0, chargeable_weight - 3) * 0.32,
        "大号大件": 9.61 + max(0, chargeable_weight - 3) * 0.42,
        "超大尺寸": 25.00 + max(0, chargeable_weight - 3) * 0.65,
    }
    fba_fee = fba_fee_map.get(pkg_size, 25.0)

    # 头程 (AGL散货)
    box_volume = (p.length_cm * p.width_cm * p.height_cm) / 1_000_000
    vol_per_container = 67 / (box_volume * p.box_count) if box_volume > 0 else 999
    wt_per_container = 19500 / (p.weight_kg * p.box_count)
    actual_shipment = max(1, int(min(vol_per_container, wt_per_container)))

    agl_cbm_rate = 619.3
    agl_per_box = (
        agl_cbm_rate * box_volume / p.exchange_rate
        + (100 + 597.4 + 350 + 15 * box_volume * 100) / p.exchange_rate / 100
        + 2600 / actual_shipment / p.exchange_rate
    )

    commission = calc_commission(p.price)
    ad_cost = round(p.price * 0.10, 2)
    return_cost = round(p.price * 0.035, 2)
    storage_cost = round(p.price * 0.015, 2)

    total_cost = (
        purchase_cost + agl_per_box + fba_fee + commission + ad_cost + return_cost + storage_cost
    )

    profit = round(p.price - total_cost, 2)
    margin = round(profit / p.price * 100, 1)

    return {
        "business_type": "AMZ FBA",
        "purchase_cost": round(purchase_cost, 2),
        "shipping_cost": round(agl_per_box, 2),
        "fba_fee": round(fba_fee, 2),
        "commission": round(commission, 2),
        "ad_cost": ad_cost,
        "return_cost": return_cost,
        "storage_cost": storage_cost,
        "total_cost": round(total_cost, 2),
        "profit": profit,
        "margin": margin,
        "package_size": pkg_size,
        "chargeable_weight": chargeable_weight,
        "actual_shipment": actual_shipment,
        "ad_return_storage_ratio": round((0.10 + 0.035 + 0.015) * 100, 1),
    }


# ============================================================
# 完整利润计算 (Walmart WF)
# ============================================================
def calc_wf_profit(p: ProductInput) -> dict:
    l_in = cm_to_inch(p.length_cm)
    w_in = cm_to_inch(p.width_cm)
    h_in = cm_to_inch(p.height_cm)

    if p.origin == "中国":
        purchase_cost = p.purchase_price_cny / p.exchange_rate
    else:
        purchase_cost = p.purchase_price_cny * 1.1

    ship = calc_shipping_cost(
        p.origin, p.length_cm, p.width_cm, p.height_cm, p.weight_kg, p.box_count, p.exchange_rate
    )
    head_cost_per_unit = ship["per_unit_head_cost"]

    # 出库操作费
    outbound_per_box = 5.0
    outbound_total = outbound_per_box * p.box_count

    # CG仓分类
    weight_lbs = calc_weight_lbs(p.weight_kg)
    cg_category = determine_cg_category(l_in, w_in, h_in, weight_lbs)
    cg_rate = CG_RATES.get(cg_category, CG_RATES["Standard Medium"])
    cg_cost = cg_rate["base"] + cg_rate["per_lb"] * max(0, weight_lbs - 1)

    cost = purchase_cost + head_cost_per_unit + outbound_total
    out_price = cost / 0.78
    front_price_low = out_price * 1.7
    front_price_high = out_price * 2.1

    return {
        "business_type": "Walmart WF",
        "purchase_cost": round(purchase_cost, 2),
        "shipping_cost": round(head_cost_per_unit, 2),
        "outbound_fee": round(outbound_total, 2),
        "total_cost": round(cost, 2),
        "out_price": round(out_price, 2),
        "front_price_low": round(front_price_low, 2),
        "front_price_high": round(front_price_high, 2),
        "cg_category": cg_category,
        "cg_cost": round(cg_cost, 2),
        "deduction_rate": 0.22,
        "shipping_info": ship,
    }
