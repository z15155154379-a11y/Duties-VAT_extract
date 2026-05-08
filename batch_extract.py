#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
批量提取海关PDF中的关键字段（MRN、关税、VAT）并汇总到Excel表格

功能：
    遍历指定文件夹中的所有PDF文件，
    提取每个文件的：
        - 进口条码编码（MRN）
        - 关税（[A00] Customs duties 的 Amount Assessed）
        - VAT税（[B00] VAT (PVA) 的 Amount Assessed）
    将所有结果保存在一个Excel文件中，一行对应一个PDF文件。

依赖库：
    pip install pandas pdfplumber openpyxl

使用方法：
    python batch_extract.py --folder ./pdf_folder --output 汇总结果.xlsx
    python batch_extract.py --folder ./pdf_folder --output 结果.xlsx --debug
"""

import re
import argparse
from pathlib import Path
from typing import Optional, Dict, List, Union

import pandas as pd
import pdfplumber


class CustomsTableParser:
    """单个PDF解析器（与之前版本相同）"""

    def __init__(self, debug: bool = False):
        self.debug = debug

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

    def extract_tax_amounts(self, text: str) -> Dict[str, Optional[float]]:
        """从文本表格中提取关税和VAT的Amount Assessed值"""
        duty = None
        vat = None
        lines = text.split('\n')

        if self.debug:
            print("\n===== 包含税种代码的行 =====")
            for line in lines:
                if re.search(r'\[A00\]|\[B00\]|Customs|VAT', line, re.IGNORECASE):
                    print(repr(line))

        for line in lines:
            # 匹配 [A00] Customs duties 行，提取第一个数字（Amount Assessed）
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

            # 匹配 [B00] 或 [BOO] VAT (PVA) 行
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
        """解析单个PDF文件，返回 {mrn, duty, vat}"""
        pdf_path = Path(pdf_path)
        full_text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + "\n"
        except Exception as e:
            print(f"  错误：无法读取 {pdf_path.name} - {e}")
            return {"mrn": None, "duty": None, "vat": None}

        if not full_text.strip():
            print(f"  警告：{pdf_path.name} 未提取到文本（可能是扫描件）")
            return {"mrn": None, "duty": None, "vat": None}

        if self.debug:
            print(f"\n====== {pdf_path.name} 提取的文本片段 ======")
            print(full_text[:1500])
            print("===========================================\n")

        mrn = self.extract_mrn(full_text)
        tax = self.extract_tax_amounts(full_text)
        return {"mrn": mrn, "duty": tax["duty"], "vat": tax["vat"]}


def batch_process(folder_path: Path, output_path: Path, debug: bool = False) -> None:
    """
    批量处理文件夹内所有PDF文件，汇总结果保存到Excel
    """
    # 获取所有PDF文件（非递归，仅当前目录）
    pdf_files = list(folder_path.glob("*.pdf"))
    if not pdf_files:
        print(f"⚠️ 文件夹 {folder_path} 中没有找到任何PDF文件。")
        return

    print(f"找到 {len(pdf_files)} 个PDF文件，开始批量提取...")
    parser = CustomsTableParser(debug=debug)
    results = []

    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"正在处理 ({i}/{len(pdf_files)}): {pdf_file.name}")
        data = parser.parse_pdf(pdf_file)
        results.append({
            "文件名": pdf_file.name,
            "进口条码编码": data.get("mrn", ""),
            "关税": data.get("duty", ""),
            "vat税": data.get("vat", "")
        })

    # 保存到Excel
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"\n✅ 批量提取完成！共处理 {len(results)} 个文件，结果已保存至: {output_path}")

    # 统计提取成功数量
    mrn_ok = sum(1 for r in results if r["进口条码编码"])
    duty_ok = sum(1 for r in results if r["关税"] is not None)
    vat_ok = sum(1 for r in results if r["vat税"] is not None)
    print(f"统计：MRN成功 {mrn_ok}/{len(results)}，关税成功 {duty_ok}/{len(results)}，VAT成功 {vat_ok}/{len(results)}")


def main():
    parser = argparse.ArgumentParser(description="批量提取PDF文件夹中的MRN、关税和VAT，汇总到Excel")
    parser.add_argument("--folder", required=True, help="存放PDF文件的文件夹路径")
    parser.add_argument("--output", required=True, help="输出的Excel文件路径（如 汇总结果.xlsx）")
    parser.add_argument("--debug", action="store_true", help="打印调试信息（显示每个PDF的文本片段）")
    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.exists() or not folder.is_dir():
        print(f"错误：文件夹不存在或不是目录 - {args.folder}")
        return

    batch_process(folder, Path(args.output), debug=args.debug)


if __name__ == "__main__":
    main()