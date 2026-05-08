#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PDF海关单据提取脚本（表格版）
从 Tax type 表格中提取：
- 关税：[A00] Customs duties 的 Amount Assessed
- VAT税：[B00] VAT (PVA) 的 Amount Assessed
同时提取 MRN（进口条码编码）
"""

import re
import argparse
from pathlib import Path
from typing import Optional, Dict, Union

import pandas as pd
import pdfplumber


class CustomsTableParser:
    """专门解析表格中的关税和VAT"""

    def __init__(self, debug: bool = False):
        self.debug = debug

    # ---------- MRN 提取（与之前相同）----------
    def extract_mrn(self, text: str) -> Optional[str]:
        patterns = [
            r'MRN\s*:\s*([A-Z0-9]{10,30})',
            r'MRN\s+([A-Z0-9]{10,30})',
            r'Master\s+Reference\s+Number\s*:\s*([A-Z0-9]{10,30})'
        ]
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return None

    # ---------- 表格提取核心 ----------
    def extract_tax_amounts(self, text: str) -> Dict[str, Optional[float]]:
        """
        从文本中查找类似以下结构的行：
        [A00] Customs duties    523.78    523.78
        [B00] VAT (PVA)         1011.72   0.00
        返回 {"duty": 523.78, "vat": 1011.72}
        """
        duty = None
        vat = None

        # 将文本按行分割
        lines = text.split('\n')
        
        # 用于调试：打印所有包含 [A00] 或 [B00] 的行
        if self.debug:
            print("\n===== 包含税种代码的行 =====")
            for line in lines:
                if re.search(r'\[A00\]|\[B00\]|Customs|VAT', line, re.IGNORECASE):
                    print(repr(line))

        for line in lines:
            # 匹配 [A00] Customs duties 行
            # 允许中间有任意空白，然后捕获第一个数字（Amount Assessed）
            # 格式示例： "[A00] Customs duties    523.78    523.78"
            match_duty = re.search(
                r'\[A00\]\s+Customs\s+duties\s+(\d+\.\d{2})', 
                line, 
                re.IGNORECASE
            )
            if match_duty:
                duty = float(match_duty.group(1))
                if self.debug:
                    print(f"找到关税: {duty}")
                continue

            # 匹配 [B00] VAT (PVA) 行
            # 注意用户提供文本中可能是 [BOO]（字母O），我们同时支持 [B00] 和 [BOO]
            match_vat = re.search(
                r'\[B0?0\]\s+VAT\s+\(PVA\)\s+(\d+\.\d{2})', 
                line, 
                re.IGNORECASE
            )
            if match_vat:
                vat = float(match_vat.group(1))
                if self.debug:
                    print(f"找到VAT: {vat}")
                continue

        return {"duty": duty, "vat": vat}

    def parse_pdf(self, pdf_path: Union[str, Path]) -> Dict[str, Optional[Union[str, float]]]:
        pdf_path = Path(pdf_path)
        full_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"

        if not full_text.strip():
            return {"mrn": None, "duty": None, "vat": None}

        if self.debug:
            print("\n====== PDF 提取的纯文本（前1500字符）======")
            print(full_text[:1500])
            print("===========================================\n")

        mrn = self.extract_mrn(full_text)
        tax_amounts = self.extract_tax_amounts(full_text)

        return {
            "mrn": mrn,
            "duty": tax_amounts["duty"],
            "vat": tax_amounts["vat"]
        }


def save_to_excel(result: Dict, output_path: Path, pdf_name: str):
    df = pd.DataFrame([{
        "进口条码编码": result.get("mrn", ""),
        "关税": result.get("duty", ""),
        "vat税": result.get("vat", ""),
        "来源文件": pdf_name
    }])
    if output_path.suffix.lower() == '.xlsx':
        df.to_excel(output_path, index=False, engine='openpyxl')
    else:
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"结果已保存至: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="从海关PDF表格中提取MRN、关税和VAT")
    parser.add_argument("--pdf", required=True, help="输入的PDF文件路径")
    parser.add_argument("--output", required=True, help="输出的Excel或CSV文件路径")
    parser.add_argument("--debug", action="store_true", help="打印调试信息")
    args = parser.parse_args()

    parser_obj = CustomsTableParser(debug=args.debug)
    result = parser_obj.parse_pdf(args.pdf)

    print("\n========== 提取结果 ==========")
    print(f"进口条码编码 (MRN): {result['mrn']}")
    print(f"关税 ([A00] Amount Assessed): {result['duty']}")
    print(f"VAT税 ([B00] Amount Assessed): {result['vat']}")
    print("==============================\n")

    if result['duty'] is None and result['vat'] is None:
        print("⚠️ 未提取到关税和VAT，请使用 --debug 模式查看提取的文本内容。")

    save_to_excel(result, Path(args.output), Path(args.pdf).name)


if __name__ == "__main__":
    main()