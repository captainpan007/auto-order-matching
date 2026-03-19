# -*- coding: utf-8 -*-
"""
PDF送货单 OCR 解析模块
使用腾讯云 表格识别 API 将供应商送货单PDF解析为结构化数据
"""

import json
import base64
import re
import os
import time
from pathlib import Path

import fitz  # PyMuPDF - PDF转图片
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.ocr.v20181119 import ocr_client, models

# ─── 腾讯云 OCR 配置 ───
# 优先从环境变量读取，其次从 SecretKey.csv 读取
def _load_credentials():
    sid = os.environ.get("TENCENT_SECRET_ID", "")
    skey = os.environ.get("TENCENT_SECRET_KEY", "")
    if sid and skey:
        return sid, skey
    csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SecretKey.csv")
    if os.path.exists(csv_path):
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            lines = f.read().strip().split('\n')
            if len(lines) >= 2:
                parts = lines[1].split(',')
                if len(parts) >= 2:
                    return parts[0].strip(), parts[1].strip()
    raise RuntimeError("未找到腾讯云密钥，请设置环境变量 TENCENT_SECRET_ID/TENCENT_SECRET_KEY 或提供 SecretKey.csv")

SECRET_ID, SECRET_KEY = _load_credentials()
REGION = "ap-guangzhou"


def _get_ocr_client():
    """创建腾讯云 OCR 客户端"""
    cred = credential.Credential(SECRET_ID, SECRET_KEY)
    http_profile = HttpProfile()
    http_profile.endpoint = "ocr.tencentcloudapi.com"
    client_profile = ClientProfile()
    client_profile.httpProfile = http_profile
    return ocr_client.OcrClient(cred, REGION, client_profile)


def _pdf_to_images(pdf_path, dpi=200, max_size_bytes=9_500_000):
    """将PDF每页转为PNG图片的base64列表，自动降DPI防止超限"""
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        current_dpi = dpi
        while current_dpi >= 100:
            zoom = current_dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.tobytes("png")
            # base64 大约增加 33%
            if len(img_bytes) * 1.34 <= max_size_bytes:
                break
            current_dpi -= 30  # 逐步降低DPI
        img_b64 = base64.b64encode(img_bytes).decode("utf-8")
        images.append((page_num + 1, img_b64))
    doc.close()
    return images


def _ocr_table_from_image(client, img_b64):
    """调用腾讯云表格识别API，返回识别结果"""
    req = models.RecognizeTableAccurateOCRRequest()
    req.ImageBase64 = img_b64
    resp = client.RecognizeTableAccurateOCR(req)
    return json.loads(resp.to_json_string())


def _ocr_table_v3_from_pdf(client, pdf_b64, page_number):
    """调用表格识别V3（UseNewModel=true），直接传PDF逐页识别"""
    req = models.RecognizeTableAccurateOCRRequest()
    # 通过 from_json_string 传参，避免 SDK 序列化参数名错误
    params = json.dumps({
        "ImageBase64": pdf_b64,
        "PdfPageNumber": page_number,
        "UseNewModel": True,
    })
    req.from_json_string(params)
    resp = client.RecognizeTableAccurateOCR(req)
    return json.loads(resp.to_json_string())


def _ocr_general_from_image(client, img_b64):
    """调用通用精准OCR，用于提取表头信息（PO号、分店等）"""
    req = models.GeneralAccurateOCRRequest()
    req.ImageBase64 = img_b64
    resp = client.GeneralAccurateOCR(req)
    return json.loads(resp.to_json_string())


def _ocr_general_from_pdf(client, pdf_b64, page_number):
    """调用通用精准OCR，直接传PDF逐页识别"""
    req = models.GeneralAccurateOCRRequest()
    req.ImageBase64 = pdf_b64
    req.PdfPageNumber = page_number
    resp = client.GeneralAccurateOCR(req)
    return json.loads(resp.to_json_string())


def _extract_header_info(general_result):
    """从通用OCR结果中提取表头信息：PO号、分店、日期、单号"""
    texts = [item["DetectedText"] for item in general_result.get("TextDetections", [])]
    full_text = " ".join(texts)
    return _extract_header_from_text(full_text, general_result.get("TextDetections", []))


def _extract_header_from_table_result(table_result):
    """从表格OCR结果的header/context单元格中提取表头信息，替代GeneralAccurateOCR"""
    all_texts = []
    for table in table_result.get("TableDetections", []):
        for cell in table.get("Cells", []):
            cell_type = cell.get("Type", "")
            text = cell.get("Text", "").strip()
            if text and cell_type in ("header", "context", "body"):
                all_texts.append(text)
    full_text = " ".join(all_texts)
    # 构造伪TextDetections（无置信度，统一100）
    pseudo_detections = [{"DetectedText": t, "Confidence": 100} for t in all_texts]
    return _extract_header_from_text(full_text, pseudo_detections)


def _extract_header_from_text(full_text, text_detections):
    """从文本中提取PO号、分店、日期等表头信息（通用逻辑）"""
    # 提取PO号 — 支持多种格式
    po_number = ""
    po_candidates = []
    for item in text_detections:
        t = item["DetectedText"].replace(" ", "") if isinstance(item, dict) else str(item)
        conf = item.get("Confidence", 100) if isinstance(item, dict) else 100
        matches = re.findall(r'[Pp][Oo0][-—.\s]*(\d{4})[-—.\s]*(\d{2})[-—.\s]*(\d{4})', t)
        for y, m, n in matches:
            po_candidates.append((conf, f"PO-{y}-{m}-{n}"))
        if not matches:
            matches2 = re.findall(r'[Pp][Oo0][-—.\s]*(\d{3})[-—.\s]*(\d{2})[-—.\s]*(\d{4})', t)
            for y, m, n in matches2:
                year = "2026" if y.startswith("202") else "20" + y
                po_candidates.append((conf, f"PO-{year}-{m}-{n}"))

    if po_candidates:
        po_candidates.sort(key=lambda x: x[0], reverse=True)
        po_number = po_candidates[0][1]

    # 提取分店
    store = ""
    if "中职" in full_text:
        store = "中职店"
    elif "高职" in full_text:
        store = "高职店"
    elif "总店" in full_text:
        store = "总店"

    # 提取日期
    date_match = re.search(r'(\d{4}[-/]\d{2}[-/]\d{2})', full_text)
    date_str = date_match.group(1) if date_match else ""

    # 提取送货单号
    xc_match = re.search(r'XC\d+[-—]\d+', full_text.replace(" ", ""))
    xc_number = xc_match.group(0) if xc_match else ""

    return {
        "po_number": po_number,
        "store": store,
        "date": date_str,
        "xc_number": xc_number,
    }


def _validate_barcode(code_str):
    """校验条码：13位纯数字"""
    if not code_str:
        return code_str, False
    # 清理空格、换行和特殊字符
    cleaned = re.sub(r'[\s\n\r\-\.，,]', '', str(code_str).strip())
    if re.match(r'^\d{13}$', cleaned):
        return cleaned, True
    # 尝试修复：如果是12位或14位，可能OCR多/少识别了一位
    if re.match(r'^\d{12,14}$', cleaned):
        return cleaned, False  # 标记为异常但保留
    return cleaned, False


def _parse_number(val):
    """解析数字，处理OCR可能的误识别"""
    if val is None or val == "":
        return None
    s = str(val).strip().replace(" ", "").replace("，", "").replace(",", "")
    # 处理可能的负号误识别
    s = s.replace("一", "-").replace("—", "-")
    # 去除验收勾选符号 √ ✓ V 等
    s = re.sub(r'[√✓✔Vv]', '', s)
    # 处理"整箱/零散"格式（如 "2/0" "1/1"）→ 取整箱数
    slash_match = re.match(r'^(\d+)\s*/\s*(\d+)$', s)
    if slash_match:
        return float(slash_match.group(1))
    # 处理带中文单位的数量（如 "1箱" "2件" "25包" "1瓶"）→ 取数字
    unit_match = re.match(r'^[-]?(\d+\.?\d*)\s*[箱件包瓶条袋桶盒托只罐]', s)
    if unit_match:
        num = float(unit_match.group(1))
        return -num if s.startswith('-') else num
    # 去除换行及其后的内容（如 "无\n26.00" → 取最后的数字）
    if '\n' in s:
        parts = s.split('\n')
        for p in reversed(parts):
            p = p.strip()
            try:
                return float(p)
            except ValueError:
                continue
        return None
    # 去除非数字前缀后缀（但保留负号和小数点）
    s = re.sub(r'[^\d.\-]', '', s)
    if not s or s in ('.', '-', '-.'):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_table_cells(table_result):
    """从表格识别结果中解析行列数据，同时返回每行原始单元格列表"""
    rows_data = {}
    if not table_result:
        return [], []

    table_infos = table_result.get("TableDetections", [])
    all_rows = []
    all_raw_rows = []  # 每行的原始单元格文本列表（不依赖列名）

    for table_info in table_infos:
        cells = table_info.get("Cells", [])
        if not cells:
            continue
        # 按行号分组
        row_dict = {}
        for cell in cells:
            row_idx = cell.get("RowTl", 0)
            col_idx = cell.get("ColTl", 0)
            text = cell.get("Text", "").strip()
            if row_idx not in row_dict:
                row_dict[row_idx] = {}
            row_dict[row_idx][col_idx] = text

        # 转为有序列表
        for row_idx in sorted(row_dict.keys()):
            cols = row_dict[row_idx]
            max_col = max(cols.keys()) if cols else 0
            row = [cols.get(c, "") for c in range(max_col + 1)]
            all_rows.append(row)
            # 原始单元格列表：所有非空文本
            raw_cells = [cols.get(c, "") for c in sorted(cols.keys())]
            all_raw_rows.append(raw_cells)

    return all_rows, all_raw_rows


def _match_columns(header_row):
    """
    根据表头行识别列索引映射。
    支持所有供应商的送货单格式：
      承运: 序号 | 条码+品名 | 单位 | 单价 | 送货数 | 送货金额 | 仓库 | 箱数 | 摘要
      承运: 序号 | 条码 | 品名 | 单位 | 单价 | 送货数 | 送货金额 | 仓库 | 箱数 | 摘要
      优亿家: 商品名称 | 条形码 | 规格 | 数量 | 价格 | 金额 | 备注
      优品味: 序号 | 条码 | 品名 | 规格 | 件数 | 零数 | 单价 | 金额 | 备注
      翔瑞: 序号 | 条形码 | 商品名称 | 商品单位 | 订单数量 | 销售价 | 订单金额 | 备注
      炬博: 商品条码 | 商品名称 | 规格 | 整件数 | 零散数 | 单价 | 金额 | 单支总数量 | 备注
      鑫凯: 商品编码 | 条形码 | 商品名称 | 规格 | 单位 | 数量 | 单价 | 金额 | 备注
    """
    mapping = {}
    barcode_name_merged = False

    # 预处理：拆分含换行的单元格，展开为 (索引, 文本) 列表
    expanded = []
    for i, h in enumerate(header_row):
        parts = str(h).split('\n')
        for part in parts:
            expanded.append((i, part.replace(" ", "").strip()))

    for i, h_clean in expanded:
        if not h_clean:
            continue
        if ("序号" in h_clean or h_clean == "序") and "seq" not in mapping:
            mapping["seq"] = i
        elif ("条码" in h_clean or "店内码" in h_clean or "条形码" in h_clean) and \
             ("品名" in h_clean or "规格" in h_clean):
            mapping["barcode_name"] = i
            barcode_name_merged = True
        elif "商品条码" in h_clean:
            mapping["barcode"] = i
        elif "条形码" in h_clean or "条码" in h_clean or "店内码" in h_clean:
            mapping["barcode"] = i
        elif "存货编码" in h_clean:
            # 乐元/双广/旺红的用友格式送货单，存货编码=条码
            if "barcode" not in mapping:
                mapping["barcode"] = i
        elif "商品编码" in h_clean:
            if "barcode" not in mapping:
                mapping["internal_code"] = i
        elif "存货名称" in h_clean or "商品名称" in h_clean:
            mapping["name"] = i
        elif "品名" in h_clean:
            if "name" not in mapping:
                mapping["name"] = i
        elif "商品单位" in h_clean or "单位" in h_clean:
            mapping["unit"] = i
        elif "散价" in h_clean:
            mapping["price"] = i
        elif "件价" in h_clean:
            if "price" not in mapping:
                mapping["case_price"] = i
        elif "小单价" in h_clean or "销售价" in h_clean or "单价" in h_clean or "价格" in h_clean:
            if "price" not in mapping:
                mapping["price"] = i
        elif "单支总数量" in h_clean:
            # 炬博的"单支总数量"是实际散件总数（最优先）
            mapping["qty"] = i
        elif "散数" in h_clean or "零散数" in h_clean or "零数" in h_clean:
            # 散数/零散数/零数 = 实际散件数量
            if "qty" not in mapping:
                mapping["qty"] = i
        elif "订单数量" in h_clean or "送货数" in h_clean or "数量" in h_clean:
            if "qty" not in mapping:
                mapping["qty"] = i
        elif "整件数" in h_clean or "件数" in h_clean:
            mapping["box_count"] = i
        elif "折后金额" in h_clean or "订单金额" in h_clean or "送货金额" in h_clean:
            mapping["amount"] = i
        elif "金额" in h_clean and "amount" not in mapping and "零售" not in h_clean:
            mapping["amount"] = i

    mapping["_merged"] = barcode_name_merged
    return mapping


def _split_barcode_name(cell_text):
    """从合并的单元格中分离条码和品名"""
    cell_text = str(cell_text).strip()
    # 模式1: "6970527350152 爱一百老式绿豆饼45g1*50"（空格分隔）
    # 模式2: "6970527350152爱一百老式绿豆饼45g1*50"（无空格）
    # 模式3: "6970527350152\n爱一百老式绿豆饼45g1*50"（换行分隔）
    match = re.match(r'^(\d{12,14})\s*(.*)', cell_text.replace("\n", " "))
    if match:
        return match.group(1).strip(), match.group(2).strip()
    # 尝试在文本中找13位数字
    match = re.search(r'(\d{13})', cell_text)
    if match:
        barcode = match.group(1)
        name = cell_text.replace(barcode, "").strip()
        return barcode, name
    return cell_text, ""


def _get_cell(row, mapping, key):
    """安全获取行中指定列的值"""
    idx = mapping.get(key)
    if idx is not None and idx < len(row):
        return row[idx]
    return ""


def _detect_remark_column(header_row):
    """查找备注/摘要列的索引"""
    for i, h in enumerate(header_row):
        h_clean = str(h).replace(" ", "").replace("\n", "")
        if h_clean in ("摘要", "备注", "备注栏", "说明"):
            return i
    # 取最后一列作为备注列候选
    return len(header_row) - 1 if header_row else None


def _parse_remark(remark_text):
    """
    解析备注栏中的手写标注
    返回: (extra_qty, is_cancelled, remark_note)
      extra_qty: 多送数量（情况1："＋5" → 5）
      is_cancelled: 是否缺货划叉（情况2："×" → True）
      remark_note: 原始备注文本
    """
    if not remark_text:
        return 0, False, ""
    text = str(remark_text).strip()
    if not text:
        return 0, False, ""

    extra_qty = 0
    is_cancelled = False

    # 情况2：缺货划叉 — ×、✕、✗ 直接匹配；X/x 只匹配独立出现（排除 GX6 等规格文本）
    if re.search(r'[×✕✗叉]', text):
        is_cancelled = True
    elif re.search(r'(?<![A-Za-z0-9])[Xx](?![A-Za-z0-9])', text):
        is_cancelled = True

    # 情况1：多送备注 — "+数字" 或 "＋数字"
    plus_match = re.search(r'[+＋]\s*(\d+)', text)
    if plus_match:
        extra_qty = int(plus_match.group(1))

    return extra_qty, is_cancelled, text


def _extract_supplement_rows(rows, header_idx):
    """
    情况3：提取表格合计行之后的手写补充行
    格式：13位条形码 + 商品名称 + 数量
    """
    supplements = []
    past_summary = False

    for row in rows[header_idx + 1:]:
        row_text = " ".join(str(c) for c in row if c)
        if "合计" in row_text:
            past_summary = True
            continue
        if not past_summary:
            continue
        if not row_text.strip():
            continue

        # 在合计行之后，寻找包含13位条码的行
        barcode_match = re.search(r'(\d{13})', row_text)
        if barcode_match:
            barcode = barcode_match.group(1)
            # 提取数量 — 行末尾的独立数字
            remaining = row_text.replace(barcode, "").strip()
            qty_match = re.search(r'(\d+)\s*$', remaining)
            qty = int(qty_match.group(1)) if qty_match else None
            # 提取品名 — 条码和数量之间的文字
            name = remaining
            if qty_match:
                name = remaining[:qty_match.start()].strip()

            supplements.append({
                "barcode": barcode,
                "name": name,
                "qty": qty,
                "is_supplement": True,
            })

    return supplements


def _infer_columns_from_data(rows):
    """
    无表头时，通过分析数据行自动推断列结构。
    策略：找第一行含13位条码的数据行，根据条码位置和列数推断映射。

    已知的无表头格式：
      炬博: 条码 | 品名 | 规格 | 整件数 | 零散数 | 单价 | 金额 | 单支总数量 | 备注 (9列)
      太古: 条码 | 品名 | 规格 | 编号 | 单位 | 数量 | 单价 | 金额 | 产地 | 日期 ... (10+列)
    """
    for row in rows:
        if not row:
            continue
        row_text = " ".join(str(c) for c in row if c)
        if "合计" in row_text:
            continue

        # 查找含13位条码的列
        barcode_col = None
        for i, cell in enumerate(row):
            cleaned = re.sub(r'[\s\n\r]', '', str(cell)) if cell else ""
            if re.match(r'^\d{13}$', cleaned):
                barcode_col = i
                break

        if barcode_col is None:
            continue

        num_cols = len(row)
        mapping = {"barcode": barcode_col, "_merged": False}

        # 品名通常紧跟条码后面
        if barcode_col + 1 < num_cols:
            mapping["name"] = barcode_col + 1

        # 根据总列数推断其余列
        if num_cols >= 8 and barcode_col == 0:
            # 炬博格式: 条码(0) | 品名(1) | 规格(2) | 整件数(3) | 零散数(4) | 单价(5) | 金额(6) | 单支总数(7) | 备注(8)
            # 优先用"单支总数量"作为qty
            mapping["box_count"] = 3
            mapping["price"] = 5
            mapping["amount"] = 6
            if num_cols >= 9:
                mapping["qty"] = 7  # 单支总数量
            else:
                mapping["qty"] = 4  # 零散数

        elif num_cols >= 7 and barcode_col == 0:
            # 可能的简化格式
            mapping["price"] = num_cols - 3
            mapping["amount"] = num_cols - 2
            mapping["qty"] = 3

        elif barcode_col == 0 and num_cols >= 10:
            # 太古格式: 条码(0) | 品名(1) | 规格(2) | 编号(3) | 单位(4) | 数量(5) | 单价(6) | 金额(7) | ...
            mapping["unit"] = 4
            mapping["qty"] = 5
            mapping["price"] = 6
            mapping["amount"] = 7

        return mapping

    return None


def _is_taigu_format(rows):
    """检测是否为太古（可口可乐）送货单格式"""
    for row in rows[:5]:
        row_text = " ".join(str(c) for c in row)
        if "预销" in row_text or "可口可乐" in row_text or "可乐" in row_text:
            return True
        # 检测数量列的"箱 N/"格式
        for cell in row:
            if re.search(r'箱\s*\d+/', str(cell)):
                return True
    return False


def _parse_taigu_qty(cell_text):
    """解析太古数量格式: '箱 1/0' → 1, '箱\n2/0' → 2, '箱 1/' → 1, '1/0' → 1"""
    s = str(cell_text).replace('\n', ' ').strip()
    # 找 N/ 或 N/M 格式，取斜杠前的整数
    match = re.search(r'(\d+)\s*/', s)
    if match:
        return int(match.group(1))
    # 纯数字
    match = re.search(r'(\d+)', s)
    if match:
        return int(match.group(1))
    return None


def _parse_taigu_price(cell_text):
    """解析太古单价: '34.00' → 34.0, '0 38.00' → 38.0, '箱 103.50' → 103.5"""
    s = str(cell_text).replace('\n', ' ').strip()
    # 提取所有浮点数，取最后一个（最可能是真正的价格）
    numbers = re.findall(r'\d+\.?\d*', s)
    if numbers:
        return float(numbers[-1])
    return None


def _extract_taigu_items(rows, header_info):
    """太古（可口可乐）专用解析器"""
    items = []

    for row in rows:
        if not row:
            continue
        row_text = " ".join(str(c) for c in row)
        if "合计" in row_text or "预销" in row_text or "业代" in row_text:
            continue
        if not row_text.strip():
            continue

        # 在行内找13位条码
        barcode = None
        barcode_col = None
        for i, cell in enumerate(row):
            cleaned = re.sub(r'[\s\n\r]', '', str(cell))
            if re.match(r'^\d{13}$', cleaned):
                barcode = cleaned
                barcode_col = i
                break

        if not barcode:
            continue

        # 品名: 条码后面第一个非空文本列
        name = ""
        for i in range(barcode_col + 1, len(row)):
            cell_str = str(row[i]).strip()
            if cell_str and not re.match(r'^[\d./\s]*$', cell_str) and len(cell_str) > 2:
                name = cell_str.replace('\n', ' ')
                break

        # 数量和单价: 在行内搜索"箱 N/"格式和价格
        qty_boxes = None
        price_per_box = None

        for cell in row:
            cell_str = str(cell)
            # 找数量（箱 N/ 格式）
            if qty_boxes is None and re.search(r'箱|/', cell_str):
                parsed_qty = _parse_taigu_qty(cell_str)
                if parsed_qty is not None and parsed_qty > 0:
                    qty_boxes = parsed_qty

        # 找单价: 行内的纯数字列（通常在最后几列）
        number_cells = []
        for i, cell in enumerate(row):
            p = _parse_taigu_price(str(cell))
            if p is not None and p > 0:
                number_cells.append((i, p))

        # 单价通常是倒数第二或第三个数字列
        if len(number_cells) >= 2:
            price_per_box = number_cells[-2][1]
        elif len(number_cells) >= 1:
            price_per_box = number_cells[-1][1]

        # 金额 = 箱数 × 箱价
        amount = None
        if qty_boxes and price_per_box:
            amount = round(qty_boxes * price_per_box, 2)

        # 如果没有解析到数量，尝试用金额/价格反算
        if qty_boxes is None and price_per_box and len(number_cells) >= 2:
            # 最后一个数字可能是金额
            possible_amount = number_cells[-1][1]
            if possible_amount > price_per_box:
                amount = possible_amount
                qty_boxes = round(amount / price_per_box)

        warning = "太古专用解析"
        if qty_boxes is None:
            warning += "; 数量未识别"

        items.append({
            "po_number": header_info.get("po_number", ""),
            "store": header_info.get("store", ""),
            "date": header_info.get("date", ""),
            "barcode": barcode,
            "barcode_valid": True,
            "name": name,
            "unit": "箱",
            "price": price_per_box,
            "qty": qty_boxes,
            "amount": amount,
            "is_cancelled": False,
            "extra_qty": 0,
            "ocr_warning": warning,
            "raw_cells": list(row),
        })

    return items


def _fill_values_from_raw_cells(raw_cells, barcode):
    """
    当列映射失败（qty/price/amount全None）时，从原始单元格行内搜索补全。
    策略：找到条码所在位置，以条码为基准向右提取数字。
    最后一个数字=金额，倒数第二=单价，倒数第三=数量。
    """
    if not raw_cells or not barcode:
        return None

    # 找条码所在的列索引
    barcode_col = None
    for i, cell in enumerate(raw_cells):
        cleaned = re.sub(r'[\s\n\r]', '', str(cell))
        if barcode in cleaned:
            barcode_col = i
            break

    if barcode_col is None:
        return None

    # 从条码右侧收集所有可解析为数字的单元格
    numbers = []
    name_parts = []
    for i in range(barcode_col + 1, len(raw_cells)):
        cell = str(raw_cells[i]).strip()
        if not cell:
            continue
        val = _parse_number(cell)
        if val is not None:
            numbers.append((i, val))
        else:
            # 非数字的可能是商品名称或规格
            if not numbers:  # 数字出现之前的文本可能是品名
                name_parts.append(cell)

    if not numbers:
        return None

    result = {}
    # 也尝试从条码左侧找品名
    if barcode_col > 0:
        left_text = str(raw_cells[barcode_col - 1]).strip()
        if left_text and not _parse_number(left_text):
            result["name"] = left_text

    if len(numbers) >= 3:
        result["qty"] = numbers[-3][1]
        result["price"] = numbers[-2][1]
        result["amount"] = numbers[-1][1]
    elif len(numbers) >= 2:
        result["price"] = numbers[-2][1]
        result["amount"] = numbers[-1][1]
    elif len(numbers) >= 1:
        result["amount"] = numbers[-1][1]

    # 如果数量或金额为负数，疑似退货行，跳过
    if result.get("qty") is not None and result["qty"] < 0:
        return None
    if result.get("amount") is not None and result["amount"] < 0:
        return None

    return result if result.get("amount") is not None else None


def _extract_items_from_rows(rows, header_info, raw_rows=None):
    """从解析的行数据中提取商品条目，处理手写标注"""
    items = []
    if not rows:
        return items

    # 太古（可口可乐）专用解析
    if _is_taigu_format(rows):
        return _extract_taigu_items(rows, header_info)

    # 查找表头行
    header_idx = None
    col_mapping = {}
    for i, row in enumerate(rows):
        row_text = " ".join(str(c) for c in row)
        if "条码" in row_text or "店内码" in row_text or "条形码" in row_text \
                or "商品条码" in row_text \
                or "存货编码" in row_text \
                or ("商品名称" in row_text and "金额" in row_text) \
                or ("商品编码" in row_text and "金额" in row_text):
            col_mapping = _match_columns(row)
            header_idx = i
            break

    if header_idx is None:
        # 无标准表头 — 尝试通过数据行自动推断列结构
        col_mapping = _infer_columns_from_data(rows)
        if col_mapping:
            header_idx = -1  # 从第0行开始处理
        else:
            # 最后兜底：条码品名合并的固定列顺序
            col_mapping = {
                "seq": 0, "barcode_name": 1, "unit": 2,
                "price": 3, "qty": 4, "amount": 5, "_merged": True
            }
            header_idx = -1

    is_merged = col_mapping.get("_merged", False)

    # 查找备注列
    if header_idx >= 0:
        remark_col = _detect_remark_column(rows[header_idx])
    else:
        remark_col = None

    # 从表头下一行开始提取数据
    data_start = header_idx + 1
    for row_i, row in enumerate(rows[data_start:], start=data_start):
        row_text = " ".join(str(c) for c in row)
        # 跳过合计行和空行
        if "合计" in row_text or not row_text.strip():
            continue

        # 提取条码和品名
        if is_merged:
            raw_cell = _get_cell(row, col_mapping, "barcode_name")
            raw_barcode_str, name = _split_barcode_name(raw_cell)
        else:
            raw_barcode_str = _get_cell(row, col_mapping, "barcode")
            name = _get_cell(row, col_mapping, "name")

        barcode, barcode_valid = _validate_barcode(raw_barcode_str)

        # 跳过没有条码的行
        if not barcode or not re.search(r'\d{6,}', barcode):
            continue

        unit = _get_cell(row, col_mapping, "unit")
        price = _parse_number(_get_cell(row, col_mapping, "price"))
        qty = _parse_number(_get_cell(row, col_mapping, "qty"))
        amount = _parse_number(_get_cell(row, col_mapping, "amount"))

        # 提取备注栏内容
        remark_text = ""
        if remark_col is not None and remark_col < len(row):
            remark_text = str(row[remark_col]) if row[remark_col] else ""
        # 也检查行内所有单元格是否有+或×标记
        all_cells_text = " ".join(str(c) for c in row if c)
        if not remark_text and re.search(r'[+＋][0-9]|[×✕✗Xx]', all_cells_text):
            remark_text = all_cells_text

        # 解析手写标注
        extra_qty, is_cancelled, remark_note = _parse_remark(remark_text)

        # 情况2：缺货划叉 → 数量置0
        if is_cancelled:
            qty = 0
            amount = 0

        # 情况1：多送备注 → 数量加上extra
        if extra_qty > 0 and qty is not None:
            qty = qty + extra_qty
            # 重算金额
            if price:
                amount = round(qty * price, 2)

        # 容错：当 qty/price/amount 全为 None 时，用行内搜索从 raw_cells 补全
        row_search_filled = False
        if qty is None and price is None and amount is None:
            row_raw_tmp = raw_rows[row_i] if raw_rows and row_i < len(raw_rows) else list(row)
            filled = _fill_values_from_raw_cells(row_raw_tmp, barcode)
            if filled:
                qty = filled.get("qty")
                price = filled.get("price")
                amount = filled.get("amount")
                if filled.get("name"):
                    name = filled["name"]
                row_search_filled = True

        # 容错：当 qty 为 None 但 price 和 amount 都有值时，反算数量
        qty_calculated = False
        if qty is None and price and amount and price > 0:
            qty = round(amount / price)
            qty_calculated = True

        warning = ""
        if not barcode_valid:
            warning = f"条码非标准13位: {raw_barcode_str}"
        if row_search_filled:
            warning += ("; " if warning else "") + "行内搜索补全"
        if qty_calculated:
            warning += ("; " if warning else "") + f"数量由金额/单价反算={qty}"
        if is_cancelled:
            warning += ("; " if warning else "") + "缺货划叉×"
        if extra_qty > 0:
            warning += ("; " if warning else "") + f"多送+{extra_qty}"

        # 负数行自动识别为退货（乐元/双广/旺红用友格式）
        is_return_item = False
        if qty is not None and qty < 0:
            is_return_item = True
            warning += ("; " if warning else "") + "退货行(负数)"
        elif amount is not None and amount < 0:
            is_return_item = True
            warning += ("; " if warning else "") + "退货行(负数金额)"

        # 保存原始单元格（用于行内搜索匹配）
        row_raw = raw_rows[row_i] if raw_rows and row_i < len(raw_rows) else list(row)

        items.append({
            "po_number": header_info.get("po_number", ""),
            "store": header_info.get("store", ""),
            "date": header_info.get("date", ""),
            "barcode": barcode,
            "barcode_valid": barcode_valid,
            "name": name,
            "unit": unit,
            "price": price,
            "qty": qty,
            "amount": amount,
            "is_cancelled": is_cancelled,
            "is_return": is_return_item,
            "extra_qty": extra_qty,
            "ocr_warning": warning,
            "raw_cells": row_raw,
        })

    # 情况3：提取合计行后的手写补充行
    supplements = _extract_supplement_rows(rows, header_idx if header_idx >= 0 else 0)
    for supp in supplements:
        items.append({
            "po_number": header_info.get("po_number", ""),
            "store": header_info.get("store", ""),
            "date": header_info.get("date", ""),
            "barcode": supp["barcode"],
            "barcode_valid": True,
            "name": supp.get("name", ""),
            "unit": "",
            "price": None,
            "qty": supp.get("qty"),
            "amount": None,
            "is_cancelled": False,
            "extra_qty": 0,
            "ocr_warning": "手写补充行",
        })

    return items


def parse_delivery_pdf(pdf_path, use_v3=False):
    """
    解析单个供应商的PDF送货单

    Args:
        pdf_path: PDF文件路径
        use_v3: 是否使用V3模式（UseNewModel=true，PDF直传）

    Returns:
        list[dict]: 结构化的送货单条目列表
    """
    print(f"  正在解析PDF: {os.path.basename(pdf_path)} {'[V3模式]' if use_v3 else ''}")

    client = _get_ocr_client()
    all_items = []

    if use_v3:
        # V3模式：表格OCR用PDF直传+UseNewModel，通用OCR仍用图片
        images = _pdf_to_images(pdf_path, dpi=200)
        total_pages = len(images)
        print(f"  共 {total_pages} 页")

        with open(pdf_path, 'rb') as f:
            pdf_b64 = base64.b64encode(f.read()).decode('utf-8')

        for page_num, img_b64 in images:
            print(f"  正在OCR识别第 {page_num}/{total_pages} 页...")

            # 表格OCR V3（用PDF直传）— 同时从表格结果提取表头
            raw_rows = []
            try:
                table_result = _ocr_table_v3_from_pdf(client, pdf_b64, page_num)
                rows, raw_rows = _parse_table_cells(table_result)
                header_info = _extract_header_from_table_result(table_result)
                print(f"    表头: PO={header_info['po_number']}, 分店={header_info['store']}, 日期={header_info['date']}")
                print(f"    表格识别: {len(rows)} 行")
            except Exception as e:
                print(f"    表格OCR失败: {e}")
                rows = []
                header_info = {"po_number": "", "store": "", "date": "", "xc_number": ""}

            # 提取商品条目
            items = _extract_items_from_rows(rows, header_info, raw_rows=raw_rows)
            print(f"    提取条目: {len(items)} 条")
            all_items.extend(items)

            time.sleep(0.6)  # V3限频

    else:
        # 原有模式：PDF转图片
        images = _pdf_to_images(pdf_path, dpi=200)
        print(f"  共 {len(images)} 页")

        for page_num, img_b64 in images:
            print(f"  正在OCR识别第 {page_num}/{len(images)} 页...")

            # 表格OCR提取商品数据 — 同时从表格结果提取表头
            raw_rows = []
            try:
                table_result = _ocr_table_from_image(client, img_b64)
                rows, raw_rows = _parse_table_cells(table_result)
                header_info = _extract_header_from_table_result(table_result)
                print(f"    表头: PO={header_info['po_number']}, 分店={header_info['store']}, 日期={header_info['date']}")
                print(f"    表格识别: {len(rows)} 行")
            except Exception as e:
                print(f"    表格OCR失败: {e}")
                rows = []
                header_info = {"po_number": "", "store": "", "date": "", "xc_number": ""}

            # 提取商品条目
            items = _extract_items_from_rows(rows, header_info, raw_rows=raw_rows)
            print(f"    提取条目: {len(items)} 条")
            all_items.extend(items)

            time.sleep(0.3)

    print(f"  PDF解析完成，共提取 {len(all_items)} 条商品记录")
    return all_items


def parse_supplier_pdf(supplier_dir, use_v3=False):
    """
    解析指定供应商文件夹下的所有PDF文件

    Args:
        supplier_dir: 供应商文件夹路径
        use_v3: 是否使用V3模式

    Returns:
        list[dict]: 所有PDF的结构化条目合并列表
    """
    pdf_files = list(Path(supplier_dir).glob("*.pdf"))
    if not pdf_files:
        print(f"  警告: {supplier_dir} 下未找到PDF文件")
        return []

    all_items = []
    for pdf_file in pdf_files:
        items = parse_delivery_pdf(str(pdf_file), use_v3=use_v3)
        all_items.extend(items)

    return all_items


# ─── 测试入口 ───
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8')

    test_dir = r"F:\claude开发项目\atutoordermatching\3.16-对账汇总\承运"
    print("=" * 60)
    print("测试：承运 PDF 送货单 OCR 解析")
    print("=" * 60)

    items = parse_supplier_pdf(test_dir)

    print("\n" + "=" * 60)
    print(f"解析结果汇总: 共 {len(items)} 条记录")
    print("=" * 60)

    # 按PO号分组统计
    po_groups = {}
    for item in items:
        po = item["po_number"] or "(未识别PO号)"
        po_groups.setdefault(po, []).append(item)

    for po, group in sorted(po_groups.items()):
        print(f"\nPO号: {po} ({len(group)} 条)")
        print(f"  分店: {group[0]['store']}")
        for item in group[:3]:  # 每组显示前3条
            valid_mark = "✓" if item["barcode_valid"] else "✗"
            print(f"  [{valid_mark}] {item['barcode']} | {item['name'][:20]} | "
                  f"数量={item['qty']} 单价={item['price']} 金额={item['amount']}")
        if len(group) > 3:
            print(f"  ... 还有 {len(group) - 3} 条")

    # 统计手写标注
    cancelled = [i for i in items if i.get("is_cancelled")]
    extra = [i for i in items if i.get("extra_qty", 0) > 0]
    supplements = [i for i in items if "手写补充行" in i.get("ocr_warning", "")]

    if cancelled:
        print(f"\n缺货划叉条目 ({len(cancelled)} 条):")
        for c in cancelled:
            print(f"  {c['barcode']} {c['name'][:20]} [缺货×]")

    if extra:
        print(f"\n多送备注条目 ({len(extra)} 条):")
        for e in extra:
            print(f"  {e['barcode']} {e['name'][:20]} +{e['extra_qty']} → 最终数量={e['qty']}")

    if supplements:
        print(f"\n手写补充行 ({len(supplements)} 条):")
        for s in supplements:
            print(f"  {s['barcode']} {s['name'][:20]} 数量={s['qty']}")

    if not cancelled and not extra and not supplements:
        print("\n未检测到手写标注")

    # 统计异常
    warnings = [item for item in items if item["ocr_warning"]]
    if warnings:
        print(f"\nOCR异常/标注条目 ({len(warnings)} 条):")
        for w in warnings:
            print(f"  {w['barcode']} - {w['ocr_warning']}")
    else:
        print("\n无OCR异常条目")
