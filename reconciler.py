# -*- coding: utf-8 -*-
"""
对账比对逻辑模块
阶段一：采购对账（送货单 vs 采购入库单）
阶段二：财务对账（进货单 vs 送货单 vs 退货单）
"""

import re
from collections import defaultdict

# 金额允许误差
AMOUNT_TOLERANCE = 0.01

# 仓库映射：送货单分店 → 入库单仓库名
STORE_MAP = {
    "总店": "集盒超市总店仓库",
    "中职店": "集盒超市中职店仓库",
    "高职店": "集盒超市高职店仓库",
}


def _normalize_store(store_str):
    """标准化仓库名称，统一为入库单格式"""
    if not store_str:
        return ""
    s = str(store_str).strip()
    if "集盒超市" in s:
        return s
    for short, full in STORE_MAP.items():
        if short in s:
            return full
    return s


def _make_key(po, barcode, store=None):
    """生成匹配键"""
    po = str(po).strip() if po else ""
    barcode = str(barcode).strip() if barcode else ""
    if store:
        store = _normalize_store(store)
        return (po, barcode, store)
    return (po, barcode)


def _extract_floats(text):
    """从字符串中提取所有浮点数（移植自原始系统 extract_floats_from_string）"""
    if not text:
        return []
    try:
        cleaned_parts = [p.strip('√✓✔/Vv') for p in str(text).split()]
        return [float(p) for p in cleaned_parts if p and _is_float(p)]
    except Exception:
        return []


def _is_float(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def _search_barcode_in_row(raw_cells, target_barcode):
    """在行的所有单元格中搜索条码（移植自原始系统 flexible_match_with_like）"""
    target = str(target_barcode).strip()
    if not target:
        return False
    for cell in raw_cells:
        cell_str = re.sub(r'[\s\n\r]', '', str(cell))
        if cell_str == target:
            return True
        if target in cell_str:
            return True
    return False


def _search_amount_in_row(raw_cells, target_amount, tolerance=AMOUNT_TOLERANCE):
    """
    在行的所有单元格中搜索匹配的金额（移植自原始系统）。
    支持直接匹配和两个金额组合匹配。
    """
    if target_amount is None:
        return False
    target = abs(float(target_amount))
    all_floats = []
    for cell in raw_cells:
        floats = _extract_floats(str(cell))
        all_floats.extend(floats)

    # 直接匹配
    for f in all_floats:
        if abs(abs(f) - target) <= tolerance:
            return True

    # 组合匹配：两个金额相加
    from itertools import combinations
    for comb in combinations(all_floats, 2):
        if abs(abs(sum(comb)) - target) <= tolerance:
            return True

    return False


def _diff_value(val_a, val_b, tolerance=AMOUNT_TOLERANCE):
    """比较两个数值，返回差值和是否一致"""
    if val_a is None and val_b is None:
        return 0, True
    if val_a is None or val_b is None:
        return None, False
    diff = round(val_a - val_b, 4)
    match = abs(diff) <= tolerance
    return diff, match


def _extract_file_amount(supplier_name, files_dict):
    """从文件名中提取预期金额（如'承运18,807.5'）"""
    for key in ["goods_receipt", "return_receipt", "pdf"]:
        path = files_dict.get(key)
        if not path:
            continue
        import os
        fname = os.path.basename(path)
        # 匹配文件名中的金额数字
        # 模式：供应商名后面的数字，如 "进货单承运18,807.5.xlsx" 或 "承运18807.5.pdf"
        match = re.search(r'[\d,]+\.?\d*', fname.replace(supplier_name, "", 1))
        if match:
            amount_str = match.group(0).replace(",", "")
            try:
                return float(amount_str)
            except ValueError:
                pass
    return None


# ═══════════════════════════════════════════════════════════
# 阶段一：采购对账（送货单 vs 采购入库单）
# ═══════════════════════════════════════════════════════════
def reconcile_purchase(delivery_items, receipt_items):
    """
    采购对账：送货单(OCR) vs 采购入库单

    Args:
        delivery_items: OCR解析的送货单条目列表
        receipt_items: 采购入库单明细条目列表

    Returns:
        dict: {
            "matched": list,       # 一致项
            "diff": list,          # 差异项
            "unmatched_delivery": list,  # 送货单有、入库单无
            "unmatched_receipt": list,   # 入库单有、送货单无
            "summary": dict,       # 汇总统计
        }
    """
    matched = []
    diff = []
    unmatched_delivery = []
    unmatched_receipt = []

    # 构建入库单索引：多级索引
    receipt_by_full = {}      # (PO, barcode, store)
    receipt_by_po_barcode = {}  # (PO, barcode)
    receipt_by_barcode_store = {}  # (barcode, store)
    receipt_by_barcode = {}   # (barcode,)
    for item in receipt_items:
        po = item["po_number"]
        bc = item["barcode"]
        st = item.get("store", "")
        receipt_by_full.setdefault((po, bc, st), []).append(item)
        receipt_by_po_barcode.setdefault((po, bc), []).append(item)
        receipt_by_barcode_store.setdefault((bc, st), []).append(item)
        receipt_by_barcode.setdefault((bc,), []).append(item)

    consumed_ids = set()

    def _fuzzy_barcode_match(bc):
        """修复2: 模糊条码匹配 — 12位补位/14位截取/编辑距离1"""
        all_barcodes = set(r["barcode"] for r in receipt_items)
        cleaned = re.sub(r'[\s\n\r]', '', bc)
        # 12位 → 补0尝试
        if len(cleaned) == 12:
            for pos in range(13):
                candidate = cleaned[:pos] + '0' + cleaned[pos:]
                if candidate in all_barcodes:
                    return candidate
        # 14位 → 逐位删除
        if len(cleaned) == 14:
            for pos in range(14):
                candidate = cleaned[:pos] + cleaned[pos+1:]
                if candidate in all_barcodes:
                    return candidate
        # 编辑距离=1（替换单个字符）
        if len(cleaned) == 13:
            for pos in range(13):
                for digit in '0123456789':
                    if digit != cleaned[pos]:
                        candidate = cleaned[:pos] + digit + cleaned[pos+1:]
                        if candidate in all_barcodes:
                            return candidate
        return None

    def _find_receipt(d_item):
        """多级匹配 + 模糊条码 + 行内搜索"""
        store = _normalize_store(d_item.get("store", ""))
        po = d_item.get("po_number", "")
        bc = d_item.get("barcode", "")

        # 修复3: PO精确匹配优先 — 有PO号时只走PO级别
        if po:
            for candidates in [receipt_by_full.get((po, bc, store), []),
                               receipt_by_po_barcode.get((po, bc), [])]:
                for item in candidates:
                    if id(item) not in consumed_ids:
                        consumed_ids.add(id(item))
                        return item, "精确"

        # 无PO号或PO匹配失败 → 条码级别
        for candidates in [receipt_by_barcode_store.get((bc, store), []),
                           receipt_by_barcode.get((bc,), [])]:
            for item in candidates:
                if id(item) not in consumed_ids:
                    consumed_ids.add(id(item))
                    return item, "条码匹配(无PO)"

        # 修复2: 模糊条码匹配
        fuzzy_bc = _fuzzy_barcode_match(bc)
        if fuzzy_bc:
            for candidates in [receipt_by_barcode.get((fuzzy_bc,), [])]:
                for item in candidates:
                    if id(item) not in consumed_ids:
                        consumed_ids.add(id(item))
                        return item, f"模糊匹配({bc}→{fuzzy_bc})"

        # 行内搜索
        raw_cells = d_item.get("raw_cells", [])
        if raw_cells:
            for r_item in receipt_items:
                if id(r_item) in consumed_ids:
                    continue
                r_bc = r_item.get("barcode", "")
                r_amt = r_item.get("amount")
                if r_bc and _search_barcode_in_row(raw_cells, r_bc):
                    if r_amt and _search_amount_in_row(raw_cells, r_amt):
                        consumed_ids.add(id(r_item))
                        return r_item, "行内搜索"

        return None, ""

    # 遍历送货单逐条匹配
    for d_item in delivery_items:
        store = _normalize_store(d_item.get("store", ""))

        result_item, match_method = _find_receipt(d_item)
        r_item = result_item

        if not r_item:
            unmatched_delivery.append({
                "source": "送货单",
                "po_number": d_item.get("po_number", ""),
                "barcode": d_item.get("barcode", ""),
                "name": d_item.get("name", ""),
                "store": store,
                "qty": d_item.get("qty"),
                "price": d_item.get("price"),
                "amount": d_item.get("amount"),
                "reason": "入库单中未找到匹配",
                "ocr_warning": d_item.get("ocr_warning", ""),
            })
            continue

        # 比对字段
        qty_diff, qty_match = _diff_value(d_item.get("qty"), r_item.get("actual_qty"), 0)
        price_diff, price_match = _diff_value(d_item.get("price"), r_item.get("price"), AMOUNT_TOLERANCE)
        amount_diff, amount_match = _diff_value(d_item.get("amount"), r_item.get("amount"), AMOUNT_TOLERANCE)

        # 补充匹配方法到 ocr_warning
        match_note = d_item.get("ocr_warning", "")
        if match_method and match_method != "精确":
            match_note += ("; " if match_note else "") + match_method

        record = {
            "po_number": d_item.get("po_number", ""),
            "barcode": d_item.get("barcode", ""),
            "name_delivery": d_item.get("name", ""),
            "name_receipt": r_item.get("name", ""),
            "store": store,
            "qty_delivery": d_item.get("qty"),
            "qty_receipt": r_item.get("actual_qty"),
            "qty_diff": qty_diff,
            "price_delivery": d_item.get("price"),
            "price_receipt": r_item.get("price"),
            "price_diff": price_diff,
            "amount_delivery": d_item.get("amount"),
            "amount_receipt": r_item.get("amount"),
            "amount_diff": amount_diff,
            "ocr_warning": match_note,
            "phase": "采购对账",
        }

        # 判定一致性
        is_amount_match_only = (amount_match and (not qty_match or not price_match))

        # 金额近似匹配（太古等箱价四舍五入场景）：金额差≤2%且绝对值≤1.5
        is_amount_approx = False
        if not amount_match and d_item.get("amount") and r_item.get("amount"):
            abs_diff = abs(d_item["amount"] - r_item["amount"])
            rel_diff = abs_diff / max(abs(r_item["amount"]), 0.01)
            if abs_diff <= 1.5 and rel_diff <= 0.02:
                is_amount_approx = True

        # 单价微差：数量一致、单价差≤0.1 → 视为一致
        is_price_minor = (qty_match and not price_match
                          and price_diff is not None and abs(price_diff) <= 0.1)

        # 修复1: 跨PO差异 — 条码匹配但数量/金额全不同，且PO不同或OCR无PO
        is_cross_po = (not amount_match and not qty_match
                       and r_item.get("barcode") == d_item.get("barcode")
                       and (not d_item.get("po_number")
                            or d_item.get("po_number") != r_item.get("po_number")))

        if qty_match and price_match and amount_match:
            record["result"] = "一致"
            matched.append(record)
        elif is_amount_approx and (not qty_match or not price_match):
            # 金额近似一致（差≤2%），数量/单价可能因箱vs散件不同
            record["result"] = "一致"
            record["ocr_warning"] = (record.get("ocr_warning", "") +
                ("; " if record.get("ocr_warning") else "") +
                f"金额近似: 送货={d_item.get('amount')} 入库={r_item.get('amount')} 差={amount_diff}")
            matched.append(record)
        elif is_price_minor:
            record["result"] = "一致"
            record["ocr_warning"] = (record.get("ocr_warning", "") +
                ("; " if record.get("ocr_warning") else "") +
                f"单价微差(≤0.1): 送货={d_item.get('price')} 入库={r_item.get('price')}")
            matched.append(record)
        elif is_amount_match_only:
            note = "金额一致"
            if not qty_match and d_item.get("qty") and r_item.get("actual_qty") and d_item["qty"] != 0:
                ratio = r_item["actual_qty"] / d_item["qty"]
                note += f",单位换算: 送货({d_item['qty']})×{ratio:.0f}=入库({r_item['actual_qty']})"
            if not price_match:
                note += f",单价口径不同: 送货={d_item.get('price')} 入库={r_item.get('price')}"
            record["result"] = "一致"
            record["ocr_warning"] = (record.get("ocr_warning", "") +
                ("; " if record.get("ocr_warning") else "") + note)
            matched.append(record)
        elif is_cross_po:
            record["result"] = "一致"
            record["ocr_warning"] = (record.get("ocr_warning", "") +
                ("; " if record.get("ocr_warning") else "") +
                f"跨PO差异: OCR_PO={d_item.get('po_number','')} 入库_PO={r_item.get('po_number','')}")
            matched.append(record)
        else:
            diffs = []
            severity = "黄色"
            if not qty_match:
                diffs.append("数量")
                if qty_diff is not None and abs(qty_diff) > 1:
                    severity = "红色"
            if not price_match:
                diffs.append("单价")
                severity = "红色"
            if not amount_match:
                diffs.append("金额")
                severity = "红色"

            record["result"] = "差异"
            record["diff_fields"] = ", ".join(diffs)
            record["severity"] = severity
            diff.append(record)

    # 查找入库单中未匹配的条目
    for r_item in receipt_items:
        if id(r_item) not in consumed_ids:
            # 跳过退货类型
            if r_item.get("biz_type") == "采购退货":
                continue
            unmatched_receipt.append({
                "source": "采购入库单",
                "po_number": r_item.get("po_number", ""),
                "barcode": r_item.get("barcode", ""),
                "name": r_item.get("name", ""),
                "store": r_item.get("store", ""),
                "qty": r_item.get("actual_qty"),
                "price": r_item.get("price"),
                "amount": r_item.get("amount"),
                "reason": "送货单中未找到匹配",
            })

    total = len(matched) + len(diff) + len(unmatched_delivery) + len(unmatched_receipt)
    summary = {
        "phase": "采购对账",
        "total_delivery": len(delivery_items),
        "total_receipt": len(receipt_items),
        "matched": len(matched),
        "diff": len(diff),
        "unmatched_delivery": len(unmatched_delivery),
        "unmatched_receipt": len(unmatched_receipt),
        "match_rate": f"{len(matched)/max(len(delivery_items),1)*100:.1f}%",
    }

    return {
        "matched": matched,
        "diff": diff,
        "unmatched_delivery": unmatched_delivery,
        "unmatched_receipt": unmatched_receipt,
        "summary": summary,
    }


# ═══════════════════════════════════════════════════════════
# 阶段二：财务对账（进货单 vs 送货单 vs 退货单）
# ═══════════════════════════════════════════════════════════
def reconcile_finance(goods_items, delivery_items, return_items, supplier_name="", files_dict=None):
    """
    财务对账：进货单 vs 送货单 vs 退货单

    Args:
        goods_items: 进货单条目列表
        delivery_items: OCR解析的送货单条目列表
        return_items: 退货单条目列表
        supplier_name: 供应商名称
        files_dict: 文件路径映射（用于提取文件名金额）

    Returns:
        dict: {
            "matched": list,
            "diff": list,
            "unmatched_goods": list,
            "unmatched_delivery": list,
            "returns": list,
            "amount_check": dict,
            "summary": dict,
        }
    """
    matched = []
    diff = []
    unmatched_goods = []
    unmatched_delivery = []
    returns = []

    # 构建送货单多级索引
    delivery_by_po_barcode = {}
    delivery_by_barcode = {}
    for item in delivery_items:
        po = item.get("po_number", "")
        bc = item.get("barcode", "")
        if po:
            delivery_by_po_barcode.setdefault((po, bc), []).append(item)
        delivery_by_barcode.setdefault((bc,), []).append(item)

    # 分笔匹配：追踪每条OCR记录的剩余数量，支持同一条码多笔入账
    remaining_d_qty = {}  # id(item) -> 剩余数量（None=未消耗）

    def _d_available(item):
        """检查OCR条目是否仍有剩余数量可分配"""
        item_id = id(item)
        if item_id not in remaining_d_qty:
            return True
        return remaining_d_qty[item_id] > 0

    def _consume_d(item, needed_qty):
        """从OCR条目消耗指定数量，返回 (用于比对的条目, 分笔备注)"""
        item_id = id(item)
        orig_qty = item.get("qty", 0) or 0
        if item_id not in remaining_d_qty:
            remaining_d_qty[item_id] = orig_qty

        remaining = remaining_d_qty[item_id]
        price = item.get("price", 0) or 0

        if needed_qty > 0 and remaining > needed_qty:
            # 部分消耗：拆分OCR数量，只分配所需数量
            remaining_d_qty[item_id] = round(remaining - needed_qty, 4)
            virtual = dict(item)
            virtual["qty"] = needed_qty
            virtual["amount"] = round(needed_qty * price, 2)
            note = f"分笔匹配: OCR原始{orig_qty}件, 本笔分配{needed_qty}件, 剩余{round(remaining - needed_qty, 4)}件"
            return virtual, note
        else:
            # 完全消耗：分配所有剩余数量
            remaining_d_qty[item_id] = 0
            if remaining < orig_qty:
                # 之前已被部分消耗，返回剩余数量的虚拟条目
                virtual = dict(item)
                virtual["qty"] = remaining
                virtual["amount"] = round(remaining * price, 2)
                note = f"分笔匹配: OCR原始{orig_qty}件, 本笔分配{remaining}件(全部剩余)"
                return virtual, note
            else:
                # 首次消耗且完全匹配
                return item, ""

    def _find_delivery(g_item, po_only=False):
        """多级匹配 + 分笔消耗 + 行内搜索 查找送货单条目
        Args:
            po_only: True=仅PO+条码匹配（第一轮），False=含条码回退和行内搜索（第二轮）
        """
        po = g_item.get("po_number", "")
        bc = g_item.get("barcode", "")
        g_qty = g_item.get("qty", 0) or 0
        g_amt = g_item.get("amount")

        # PO+条码精确匹配（优先级最高）
        if po:
            for item in delivery_by_po_barcode.get((po, bc), []):
                if _d_available(item):
                    return _consume_d(item, g_qty)

        if po_only:
            return None, ""

        # 条码回退匹配（仅第二轮）
        for item in delivery_by_barcode.get((bc,), []):
            if _d_available(item):
                return _consume_d(item, g_qty)

        # 行内搜索 — 在送货单的每行原始单元格中搜索进货单条码+金额
        if bc and g_amt:
            for d_item in delivery_items:
                if not _d_available(d_item):
                    continue
                raw_cells = d_item.get("raw_cells", [])
                if raw_cells and _search_barcode_in_row(raw_cells, bc):
                    if _search_amount_in_row(raw_cells, g_amt):
                        return _consume_d(d_item, g_qty)

        return None, ""

    # 两轮匹配：第一轮PO+条码精确匹配，第二轮条码回退+行内搜索
    # 防止条码回退抢占本该属于PO精确匹配的OCR数量
    pending_goods = []  # 第一轮未匹配的进货单条目
    for g_item in goods_items:
        d_item, split_note = _find_delivery(g_item, po_only=True)
        if not d_item:
            pending_goods.append(g_item)
            continue
        # 第一轮匹配成功，直接处理（复用下方比对逻辑）
        pending_goods.append(("__matched__", g_item, d_item, split_note))

    # 第二轮：对未匹配条目做条码回退+行内搜索
    final_goods = []
    for entry in pending_goods:
        if isinstance(entry, tuple) and entry[0] == "__matched__":
            final_goods.append(entry[1:])  # (g_item, d_item, split_note)
        else:
            g_item = entry
            d_item, split_note = _find_delivery(g_item, po_only=False)
            final_goods.append((g_item, d_item, split_note))

    for g_item, d_item, split_note in final_goods:
        if not d_item:
            unmatched_goods.append({
                "source": "进货单",
                "po_number": g_item.get("po_number", ""),
                "barcode": g_item.get("barcode", ""),
                "name": g_item.get("name", ""),
                "receipt_no": g_item.get("receipt_no", ""),
                "qty": g_item.get("qty"),
                "price": g_item.get("price"),
                "amount": g_item.get("amount"),
                "reason": "送货单中未找到匹配",
            })
            continue

        # 比对字段
        qty_diff, qty_match = _diff_value(g_item.get("qty"), d_item.get("qty"), 0)
        price_diff, price_match = _diff_value(g_item.get("price"), d_item.get("price"), AMOUNT_TOLERANCE)
        amount_diff, amount_match = _diff_value(g_item.get("amount"), d_item.get("amount"), AMOUNT_TOLERANCE)

        # 合并OCR警告和分笔备注
        ocr_warn = d_item.get("ocr_warning", "")
        if split_note:
            ocr_warn += ("; " if ocr_warn else "") + split_note

        record = {
            "po_number": g_item.get("po_number", ""),
            "barcode": g_item.get("barcode", ""),
            "name_goods": g_item.get("name", ""),
            "name_delivery": d_item.get("name", ""),
            "receipt_no": g_item.get("receipt_no", ""),
            "qty_goods": g_item.get("qty"),
            "qty_delivery": d_item.get("qty"),
            "qty_diff": qty_diff,
            "price_goods": g_item.get("price"),
            "price_delivery": d_item.get("price"),
            "price_diff": price_diff,
            "amount_goods": g_item.get("amount"),
            "amount_delivery": d_item.get("amount"),
            "amount_diff": amount_diff,
            "ocr_warning": ocr_warn,
            "phase": "财务对账",
        }

        # 判定：金额一致但数量或单价不同 → 视为一致
        is_amount_match_only = (amount_match and (not qty_match or not price_match))

        # 金额近似匹配（太古等箱价四舍五入场景）：金额差≤2%且绝对值≤1.5
        is_amount_approx_f = False
        if not amount_match and g_item.get("amount") and d_item.get("amount"):
            abs_diff = abs(g_item["amount"] - d_item["amount"])
            rel_diff = abs_diff / max(abs(g_item["amount"]), 0.01)
            if abs_diff <= 1.5 and rel_diff <= 0.02:
                is_amount_approx_f = True

        if qty_match and price_match and amount_match:
            record["result"] = "一致"
            matched.append(record)
        elif is_amount_approx_f and (not qty_match or not price_match):
            record["result"] = "一致"
            record["ocr_warning"] = (record.get("ocr_warning", "") +
                ("; " if record.get("ocr_warning") else "") +
                f"金额近似: 进货={g_item.get('amount')} 送货={d_item.get('amount')} 差={amount_diff}")
            matched.append(record)
        elif is_amount_match_only:
            note = "金额一致"
            if not qty_match and g_item.get("qty") and d_item.get("qty") and d_item["qty"] != 0:
                ratio = g_item["qty"] / d_item["qty"]
                note += f",单位换算: 送货({d_item['qty']})×{ratio:.0f}=进货({g_item['qty']})"
            if not price_match:
                note += f",单价口径不同: 送货={d_item.get('price')} 进货={g_item.get('price')}"
            record["result"] = "一致"
            record["ocr_warning"] = (record.get("ocr_warning", "") +
                ("; " if record.get("ocr_warning") else "") + note)
            matched.append(record)
        else:
            diffs = []
            severity = "黄色"
            if not qty_match:
                diffs.append("数量")
                if qty_diff is not None and abs(qty_diff) > 1:
                    severity = "红色"
            if not price_match:
                diffs.append("单价")
                severity = "红色"
            if not amount_match:
                diffs.append("金额")
                severity = "红色"

            record["result"] = "差异"
            record["diff_fields"] = ", ".join(diffs)
            record["severity"] = severity
            diff.append(record)

    # 查找送货单中未匹配的条目（完全未消耗或有剩余数量的）
    for d_item in delivery_items:
        item_id = id(d_item)
        if item_id not in remaining_d_qty:
            # 完全未被匹配过
            unmatched_delivery.append({
                "source": "送货单",
                "po_number": d_item.get("po_number", ""),
                "barcode": d_item.get("barcode", ""),
                "name": d_item.get("name", ""),
                "store": d_item.get("store", ""),
                "qty": d_item.get("qty"),
                "price": d_item.get("price"),
                "amount": d_item.get("amount"),
                "reason": "进货单中未找到匹配",
                "ocr_warning": d_item.get("ocr_warning", ""),
            })
        elif remaining_d_qty[item_id] > 0:
            # 部分消耗后仍有剩余
            leftover = remaining_d_qty[item_id]
            price = d_item.get("price", 0) or 0
            unmatched_delivery.append({
                "source": "送货单",
                "po_number": d_item.get("po_number", ""),
                "barcode": d_item.get("barcode", ""),
                "name": d_item.get("name", ""),
                "store": d_item.get("store", ""),
                "qty": leftover,
                "price": price,
                "amount": round(leftover * price, 2),
                "reason": f"分笔后剩余{leftover}件未匹配",
                "ocr_warning": d_item.get("ocr_warning", ""),
            })

    # 退货单处理
    for r_item in return_items:
        returns.append({
            "po_number": r_item.get("po_number", ""),
            "barcode": r_item.get("barcode", ""),
            "name": r_item.get("name", ""),
            "receipt_no": r_item.get("receipt_no", ""),
            "qty": r_item.get("qty"),
            "price": r_item.get("price"),
            "amount": r_item.get("amount"),
        })

    # 金额汇总校验
    goods_total = sum(i.get("amount") or 0 for i in goods_items)
    delivery_total = sum(i.get("amount") or 0 for i in delivery_items)
    return_total = sum(abs(i.get("amount") or 0) for i in return_items)
    actual_payable = goods_total - return_total

    # 从文件名提取预期金额
    file_amount = _extract_file_amount(supplier_name, files_dict or {})

    amount_check = {
        "goods_total": round(goods_total, 2),
        "delivery_total": round(delivery_total, 2),
        "return_total": round(return_total, 2),
        "actual_payable": round(actual_payable, 2),
        "file_amount": file_amount,
        "goods_vs_delivery_diff": round(goods_total - delivery_total, 2),
        "payable_vs_file_diff": round(actual_payable - file_amount, 2) if file_amount else None,
        "goods_delivery_match": abs(goods_total - delivery_total) <= AMOUNT_TOLERANCE,
        "file_amount_match": abs(actual_payable - file_amount) <= 0.1 if file_amount else None,
    }

    summary = {
        "phase": "财务对账",
        "total_goods": len(goods_items),
        "total_delivery": len(delivery_items),
        "total_returns": len(return_items),
        "matched": len(matched),
        "diff": len(diff),
        "unmatched_goods": len(unmatched_goods),
        "unmatched_delivery": len(unmatched_delivery),
        "match_rate": f"{len(matched)/max(len(goods_items),1)*100:.1f}%",
        "goods_total": round(goods_total, 2),
        "delivery_total": round(delivery_total, 2),
        "return_total": round(return_total, 2),
        "actual_payable": round(actual_payable, 2),
    }

    return {
        "matched": matched,
        "diff": diff,
        "unmatched_goods": unmatched_goods,
        "unmatched_delivery": unmatched_delivery,
        "returns": returns,
        "amount_check": amount_check,
        "summary": summary,
    }


# ═══════════════════════════════════════════════════════════
# 完整对账入口
# ═══════════════════════════════════════════════════════════
def reconcile_supplier(delivery_items, excel_data, supplier_name=""):
    """
    对单个供应商执行完整两阶段对账

    Args:
        delivery_items: OCR解析的送货单条目
        excel_data: read_all_supplier_excel()返回的数据
        supplier_name: 供应商名称

    Returns:
        dict: {
            "phase1": 采购对账结果,
            "phase2": 财务对账结果,
            "supplier_name": str,
        }
    """
    # 阶段一：采购对账
    phase1 = reconcile_purchase(
        delivery_items,
        excel_data["purchase_receipt_detail"]
    )

    # 阶段二：财务对账
    phase2 = reconcile_finance(
        excel_data["goods_receipt"],
        delivery_items,
        excel_data["return_receipt"],
        supplier_name=supplier_name,
        files_dict=excel_data.get("files", {}),
    )

    return {
        "phase1": phase1,
        "phase2": phase2,
        "supplier_name": supplier_name,
    }


# ─── 测试入口 ───
if __name__ == "__main__":
    import sys, os
    sys.stdout.reconfigure(encoding='utf-8')

    # 导入依赖模块
    from excel_reader import read_all_supplier_excel
    from ocr_parser import parse_supplier_pdf

    base = r"F:\claude开发项目\atutoordermatching\3.16-对账汇总"
    supplier = "承运"
    supplier_dir = os.path.join(base, supplier)

    print("=" * 60)
    print(f"对账测试: {supplier}")
    print("=" * 60)

    # 读取Excel数据
    print("\n[1] 读取Excel数据...")
    excel_data = read_all_supplier_excel(supplier_dir)
    print(f"  采购订单: {len(excel_data['purchase_order'])} 条")
    print(f"  采购入库单: {len(excel_data['purchase_receipt_detail'])} 条")
    print(f"  进货单: {len(excel_data['goods_receipt'])} 条")
    print(f"  退货单: {len(excel_data['return_receipt'])} 条")

    # 读取OCR数据（使用缓存避免重复调用API）
    cache_file = os.path.join(supplier_dir, "_ocr_cache.json")
    if os.path.exists(cache_file):
        import json
        print(f"\n[2] 从缓存加载OCR数据: {cache_file}")
        with open(cache_file, 'r', encoding='utf-8') as f:
            delivery_items = json.load(f)
    else:
        print(f"\n[2] OCR识别送货单...")
        delivery_items = parse_supplier_pdf(supplier_dir)
        # 保存缓存
        import json
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(delivery_items, f, ensure_ascii=False, indent=2)
        print(f"  已缓存到 {cache_file}")

    print(f"  送货单: {len(delivery_items)} 条")

    # 执行对账
    print(f"\n[3] 执行对账比对...")
    result = reconcile_supplier(delivery_items, excel_data, supplier)

    # 输出阶段一结果
    p1 = result["phase1"]
    print(f"\n{'─'*60}")
    print(f"阶段一：采购对账（送货单 vs 采购入库单）")
    print(f"{'─'*60}")
    print(f"  送货单条目: {p1['summary']['total_delivery']}")
    print(f"  入库单条目: {p1['summary']['total_receipt']}")
    print(f"  一致: {p1['summary']['matched']} 条")
    print(f"  差异: {p1['summary']['diff']} 条")
    print(f"  送货单未匹配: {p1['summary']['unmatched_delivery']} 条")
    print(f"  入库单未匹配: {p1['summary']['unmatched_receipt']} 条")
    print(f"  一致率: {p1['summary']['match_rate']}")

    if p1["diff"]:
        print(f"\n  差异项前5条:")
        for d in p1["diff"][:5]:
            print(f"    [{d['severity']}] {d['barcode']} {d['name_delivery'][:15]}")
            print(f"      差异字段: {d['diff_fields']}")
            if d.get("qty_diff") is not None and d["qty_diff"] != 0:
                print(f"      数量: 送货={d['qty_delivery']} 入库={d['qty_receipt']} 差={d['qty_diff']}")
            if d.get("price_diff") is not None and d["price_diff"] != 0:
                print(f"      单价: 送货={d['price_delivery']} 入库={d['price_receipt']} 差={d['price_diff']}")

    if p1["unmatched_delivery"]:
        print(f"\n  送货单未匹配前5条:")
        for u in p1["unmatched_delivery"][:5]:
            print(f"    {u['barcode']} {u['name'][:20]} 数量={u['qty']} "
                  f"{'[OCR异常]' if u.get('ocr_warning') else ''}")

    if p1["unmatched_receipt"]:
        print(f"\n  入库单未匹配前5条:")
        for u in p1["unmatched_receipt"][:5]:
            print(f"    {u['barcode']} {u['name'][:20]} 数量={u['qty']}")

    # 输出阶段二结果
    p2 = result["phase2"]
    print(f"\n{'─'*60}")
    print(f"阶段二：财务对账（进货单 vs 送货单 vs 退货单）")
    print(f"{'─'*60}")
    print(f"  进货单条目: {p2['summary']['total_goods']}")
    print(f"  送货单条目: {p2['summary']['total_delivery']}")
    print(f"  退货单条目: {p2['summary']['total_returns']}")
    print(f"  一致: {p2['summary']['matched']} 条")
    print(f"  差异: {p2['summary']['diff']} 条")
    print(f"  进货单未匹配: {p2['summary']['unmatched_goods']} 条")
    print(f"  送货单未匹配: {p2['summary']['unmatched_delivery']} 条")
    print(f"  一致率: {p2['summary']['match_rate']}")

    ac = p2["amount_check"]
    print(f"\n  金额校验:")
    print(f"    进货单总额: {ac['goods_total']}")
    print(f"    送货单总额: {ac['delivery_total']}")
    print(f"    退货单总额: {ac['return_total']}")
    print(f"    实际应付: {ac['actual_payable']}")
    print(f"    文件名金额: {ac['file_amount']}")
    print(f"    进货-送货差额: {ac['goods_vs_delivery_diff']}")
    print(f"    应付-文件名差额: {ac['payable_vs_file_diff']}")

    if p2["diff"]:
        print(f"\n  差异项前5条:")
        for d in p2["diff"][:5]:
            print(f"    [{d['severity']}] {d['barcode']} {d['name_goods'][:15]}")
            print(f"      差异字段: {d['diff_fields']}")
            if d.get("qty_diff") is not None and d["qty_diff"] != 0:
                print(f"      数量: 进货={d['qty_goods']} 送货={d['qty_delivery']} 差={d['qty_diff']}")
            if d.get("price_diff") is not None and d["price_diff"] != 0:
                print(f"      单价: 进货={d['price_goods']} 送货={d['price_delivery']} 差={d['price_diff']}")
