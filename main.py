# -*- coding: utf-8 -*-
"""
超市采购对账自动化系统 - 主程序入口
遍历所有供应商文件夹，执行OCR识别、数据比对、生成报告
"""

import os
import sys
import json
import time
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

from ocr_parser import parse_supplier_pdf
from excel_reader import read_all_supplier_excel, identify_files
from reconciler import reconcile_supplier, reconcile_purchase, reconcile_finance
from report_generator import generate_supplier_report, generate_summary_report


# ─── 配置 ───
from config import BASE_DATA_DIR, OUTPUT_DIR as _CFG_OUTPUT_DIR
BASE_DIR = os.environ.get("RECONCILE_BASE_DIR", BASE_DATA_DIR)
OUTPUT_DIR = _CFG_OUTPUT_DIR
DATE_STR = datetime.now().strftime("%Y%m%d")

# 使用V3表格识别的供应商（UseNewModel=true，识别效果更好）
V3_SUPPLIERS = {"优亿家", "炬博"}


def _empty_phase(phase_name):
    """生成空的对账阶段结果（用于跳过的阶段）"""
    if phase_name == "采购对账":
        return {
            "matched": [], "diff": [], "unmatched_delivery": [], "unmatched_receipt": [],
            "summary": {"phase": "采购对账", "total_delivery": 0, "total_receipt": 0,
                        "matched": 0, "diff": 0, "unmatched_delivery": 0, "unmatched_receipt": 0,
                        "match_rate": "-"}
        }
    else:
        return {
            "matched": [], "diff": [], "unmatched_goods": [], "unmatched_delivery": [], "returns": [],
            "amount_check": {"goods_total": 0, "delivery_total": 0, "return_total": 0,
                             "actual_payable": 0, "file_amount": 0, "goods_vs_delivery_diff": 0,
                             "payable_vs_file_diff": 0, "goods_delivery_match": True, "file_amount_match": True},
            "summary": {"phase": "财务对账", "total_goods": 0, "total_delivery": 0, "total_returns": 0,
                        "matched": 0, "diff": 0, "unmatched_goods": 0, "unmatched_delivery": 0,
                        "match_rate": "-", "goods_total": 0, "delivery_total": 0, "return_total": 0,
                        "actual_payable": 0}
        }


def process_supplier(supplier_dir, use_cache=True, mode="both"):
    """
    处理单个供应商的完整对账流程

    Args:
        supplier_dir: 供应商文件夹路径
        use_cache: 是否使用OCR缓存
        mode: "purchase"=只采购对账, "finance"=只财务对账, "both"=两个都跑

    Returns:
        dict: 对账结果，或None（如果出错）
    """
    supplier_name = Path(supplier_dir).name
    print(f"\n{'═'*60}")
    print(f"  处理供应商: {supplier_name}")
    print(f"{'═'*60}")

    # 1. 识别文件
    files = identify_files(supplier_dir)

    # 根据对账模式决定必需文件
    if mode == "purchase":
        required = ["purchase_receipt", "pdf"]
    elif mode == "finance":
        required = ["goods_receipt", "pdf"]
    else:
        required = ["purchase_receipt", "goods_receipt", "pdf"]

    missing = [k for k in required if not files[k]]
    if missing:
        print(f"  ⚠ 缺少文件: {missing}，跳过")
        return None

    file_status = []
    if files["purchase_order"]: file_status.append("采购订单✓")
    if files["purchase_receipt"]: file_status.append("采购入库单✓")
    if files["goods_receipt"]: file_status.append("进货单✓")
    if files["pdf"]: file_status.append("PDF✓")
    if files["return_receipt"]: file_status.append("退货单✓")
    print(f"  文件: {' '.join(file_status)}")

    # 2. 读取Excel数据
    print(f"  读取Excel...")
    try:
        excel_data = read_all_supplier_excel(supplier_dir)
    except Exception as e:
        print(f"  ✗ Excel读取失败: {e}")
        return None

    print(f"    采购订单={len(excel_data['purchase_order'])} "
          f"入库单={len(excel_data['purchase_receipt_detail'])} "
          f"进货单={len(excel_data['goods_receipt'])} "
          f"退货单={len(excel_data['return_receipt'])}")

    # 3. OCR识别送货单
    use_v3 = supplier_name in V3_SUPPLIERS
    cache_file = os.path.join(supplier_dir, "_ocr_cache_v3.json" if use_v3 else "_ocr_cache.json")
    if use_cache and os.path.exists(cache_file):
        print(f"  从缓存加载OCR数据{'[V3]' if use_v3 else ''}...")
        with open(cache_file, 'r', encoding='utf-8') as f:
            delivery_items = json.load(f)
        print(f"    送货单={len(delivery_items)} 条（缓存）")
    else:
        print(f"  OCR识别送货单{'[V3模式]' if use_v3 else ''}...")
        try:
            delivery_items = parse_supplier_pdf(supplier_dir, use_v3=use_v3)
        except Exception as e:
            print(f"  ✗ OCR识别失败: {e}")
            return None
        # 保存缓存
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(delivery_items, f, ensure_ascii=False, indent=2)
        print(f"    送货单={len(delivery_items)} 条（已缓存）")

    if not delivery_items:
        print(f"  ⚠ 送货单无数据，跳过")
        return None

    # 4. 执行对账比对
    print(f"  执行对账比对（模式: {mode}）...")

    if mode == "both":
        result = reconcile_supplier(delivery_items, excel_data, supplier_name)
    elif mode == "purchase":
        phase1 = reconcile_purchase(delivery_items, excel_data["purchase_receipt_detail"])
        result = {"phase1": phase1, "phase2": _empty_phase("财务对账"), "supplier_name": supplier_name}
    elif mode == "finance":
        phase2 = reconcile_finance(
            excel_data["goods_receipt"], delivery_items, excel_data["return_receipt"],
            supplier_name=supplier_name, files_dict=excel_data.get("files", {}))
        result = {"phase1": _empty_phase("采购对账"), "phase2": phase2, "supplier_name": supplier_name}
    else:
        result = reconcile_supplier(delivery_items, excel_data, supplier_name)

    # 5. 输出摘要
    if mode in ("both", "purchase"):
        p1 = result["phase1"]["summary"]
        print(f"  采购对账: 一致={p1['matched']} 差异={p1['diff']} "
              f"未匹配={p1['unmatched_delivery']+p1['unmatched_receipt']} "
              f"一致率={p1['match_rate']}")
    if mode in ("both", "finance"):
        p2 = result["phase2"]["summary"]
        print(f"  财务对账: 一致={p2['matched']} 差异={p2['diff']} "
              f"未匹配={p2['unmatched_goods']+p2['unmatched_delivery']} "
              f"一致率={p2['match_rate']}")

    ac = result["phase2"]["amount_check"]
    if mode in ("both", "finance"):
        print(f"  金额: 进货={ac['goods_total']} 送货={ac['delivery_total']} "
              f"退货={ac['return_total']} 应付={ac['actual_payable']}")

    # 6. 生成报告
    print(f"  生成对账报告...")
    report_path = generate_supplier_report(result, OUTPUT_DIR, DATE_STR)
    print(f"    → {os.path.basename(report_path)}")

    return result


def main():
    """主入口：遍历所有供应商"""
    print("=" * 60)
    print("  超市采购对账自动化系统")
    print(f"  数据目录: {BASE_DIR}")
    print(f"  输出目录: {OUTPUT_DIR}")
    print(f"  对账日期: {DATE_STR}")
    print("=" * 60)

    # 检查是否使用缓存
    use_cache = "--no-cache" not in sys.argv
    if not use_cache:
        print("  注意: --no-cache 模式，将重新进行OCR识别")

    # 对账模式
    mode = os.environ.get("RECONCILE_MODE", "both")
    mode_names = {"purchase": "采购对账", "finance": "财务对账", "both": "完整对账"}
    print(f"  对账模式: {mode_names.get(mode, mode)}")

    # 获取所有供应商文件夹
    suppliers = sorted([
        d for d in Path(BASE_DIR).iterdir()
        if d.is_dir() and not d.name.startswith("_")
    ])

    # 支持通过环境变量筛选供应商（GUI传入）
    selected = os.environ.get("RECONCILE_SUPPLIERS", "")
    if selected:
        selected_set = set(selected.split(","))
        suppliers = [s for s in suppliers if s.name in selected_set]

    print(f"\n  发现 {len(suppliers)} 个供应商: {[s.name for s in suppliers]}")

    # 逐个处理
    all_results = []
    failed = []
    start_time = time.time()

    for supplier_dir in suppliers:
        try:
            result = process_supplier(str(supplier_dir), use_cache=use_cache, mode=mode)
            if result:
                all_results.append(result)
            else:
                failed.append(supplier_dir.name)
        except Exception as e:
            print(f"  ✗ 处理失败: {e}")
            failed.append(supplier_dir.name)

    elapsed = time.time() - start_time

    # 生成汇总报告
    if all_results:
        print(f"\n{'═'*60}")
        print(f"  生成汇总报告...")
        summary_path = generate_summary_report(all_results, OUTPUT_DIR, DATE_STR)
        print(f"  → {os.path.basename(summary_path)}")

    # 输出总结
    print(f"\n{'═'*60}")
    print(f"  处理完成")
    print(f"{'═'*60}")
    print(f"  成功: {len(all_results)} 个供应商")
    if failed:
        print(f"  失败: {len(failed)} 个 ({', '.join(failed)})")
    print(f"  耗时: {elapsed:.1f} 秒")
    print(f"  报告目录: {OUTPUT_DIR}")

    # 汇总所有供应商的一致率
    if all_results:
        if mode in ("both", "purchase"):
            total_p1_items = sum(r["phase1"]["summary"]["total_delivery"] for r in all_results)
            total_p1_matched = sum(r["phase1"]["summary"]["matched"] for r in all_results)
            print(f"\n  总体采购对账一致率: {total_p1_matched}/{total_p1_items} "
                  f"({total_p1_matched/max(total_p1_items,1)*100:.1f}%)")
        if mode in ("both", "finance"):
            total_p2_items = sum(r["phase2"]["summary"]["total_goods"] for r in all_results)
            total_p2_matched = sum(r["phase2"]["summary"]["matched"] for r in all_results)
            print(f"  总体财务对账一致率: {total_p2_matched}/{total_p2_items} "
                  f"({total_p2_matched/max(total_p2_items,1)*100:.1f}%)")


if __name__ == "__main__":
    main()
