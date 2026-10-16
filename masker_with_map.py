"""masker_with_map.py
在 Windows 上批量脱敏 docx/xlsx/pptx，并在输出目录生成映射文件（CSV），
可用于保存原始->脱敏对应关系，也支持用该映射进行反向还原（unmask）。

用法：
  python masker_with_map.py mask <in_folder> <out_folder>    # 执行脱敏并生成 mapping_YYYYmmdd_HHMMSS.csv
  python masker_with_map.py unmask <in_folder> <out_folder> <mapping_csv>  # 使用 mapping 恢复（仅对包含 masked token 的文件有效）

备注：
 - 本脚本仍仅处理 .docx .xlsx .pptx 文件（Office Open XML）。.doc/.xls/.ppt 需先另存为新格式。
 - 脱敏采用占位 token（例如 [EMAIL_1], [PHONE_2], [ID_3]），相同原文会映射为相同 token（全局一致），并记录在 mapping CSV 中。
 - mapping CSV 格式: masked,original,type,first_seen_in_file
"""

import os
import re
import sys
import csv
import argparse
from datetime import datetime
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

# 正则：邮箱、手机号（中国 11 位）、身份证（15 或 18）
EMAIL_RE = re.compile(r'([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)')
PHONE_RE = re.compile(r'(\b1[3-9]\d{9}\b)')
ID_RE = re.compile(r'\b(\d{15}|\d{17}[0-9Xx])\b')

# token 格式模板
TEMPLATES = {
    'EMAIL': '[EMAIL_{:05d}]',
    'PHONE': '[PHONE_{:05d}]',
    'ID': '[ID_{:05d}]'
}

class Mapper:
    def __init__(self):
        # maps original -> masked
        self.maps = {
            'EMAIL': {},
            'PHONE': {},
            'ID': {}
        }
        # counters per type
        self.cnts = {
            'EMAIL': 0,
            'PHONE': 0,
            'ID': 0
        }
        # record first seen file for mapping
        self.first_seen = {}

    def get_mask(self, typ, original, source_file=None):
        d = self.maps[typ]
        if original in d:
            return d[original]
        self.cnts[typ] += 1
        token = TEMPLATES[typ].format(self.cnts[typ])
        d[original] = token
        if token not in self.first_seen:
            self.first_seen[token] = source_file or ''
        return token

    def save_csv(self, csv_path):
        # csv columns: masked,original,type,first_seen_in_file
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            w.writerow(['masked', 'original', 'type', 'first_seen_in_file'])
            for typ in ['EMAIL', 'PHONE', 'ID']:
                for orig, masked in self.maps[typ].items():
                    w.writerow([masked, orig, typ, self.first_seen.get(masked, '')])

    @classmethod
    def load_from_csv(cls, csv_path):
        m = cls()
        if not os.path.exists(csv_path):
            raise FileNotFoundError(csv_path)
        with open(csv_path, 'r', encoding='utf-8') as f:
            r = csv.DictReader(f)
            for row in r:
                masked = row['masked']
                original = row['original']
                typ = row['type']
                m.maps.setdefault(typ, {})[original] = masked
                # try to parse counter from masked token to keep counters growing
                mo = re.search(r'_(\d+)\]', masked)
                if mo:
                    n = int(mo.group(1))
                    m.cnts[typ] = max(m.cnts.get(typ, 0), n)
                m.first_seen[masked] = row.get('first_seen_in_file', '')
        return m


# 替换函数，使用 mapper
def mask_text(text, mapper, source_file=None):
    if not text:
        return text
    # emails
    def e_repl(m):
        orig = m.group(1)
        tok = mapper.get_mask('EMAIL', orig, source_file)
        return tok
    text = EMAIL_RE.sub(e_repl, text)
    # phones
    def p_repl(m):
        orig = m.group(1)
        tok = mapper.get_mask('PHONE', orig, source_file)
        return tok
    text = PHONE_RE.sub(p_repl, text)
    # ids
    def id_repl(m):
        orig = m.group(1)
        tok = mapper.get_mask('ID', orig, source_file)
        return tok
    text = ID_RE.sub(id_repl, text)
    return text

# 反向替换：把 masked token 还原为 original
def unmask_text(text, mapper):
    if not text:
        return text
    # build masked->original map
    rev = {}
    for typ in mapper.maps:
        for orig, masked in mapper.maps[typ].items():
            rev[masked] = orig
    # token 格式类似 [EMAIL_00001]
    if not rev:
        return text
    # 简单地替换所有已知 token
    # 为避免部分匹配问题，按 token 长度降序替换
    tokens = sorted(rev.keys(), key=len, reverse=True)
    for t in tokens:
        if t in text:
            text = text.replace(t, rev[t])
    return text


def process_docx(in_path, out_path, mapper=None, mode='mask'):
    doc = Document(in_path)
    for p in doc.paragraphs:
        for run in p.runs:
            if mode == 'mask':
                run.text = mask_text(run.text, mapper, in_path)
            else:
                run.text = unmask_text(run.text, mapper)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if mode == 'mask':
                    cell.text = mask_text(cell.text, mapper, in_path)
                else:
                    cell.text = unmask_text(cell.text, mapper)
    doc.save(out_path)


def process_xlsx(in_path, out_path, mapper=None, mode='mask'):
    wb = load_workbook(in_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if mode == 'mask':
                        cell.value = mask_text(cell.value, mapper, in_path)
                    else:
                        cell.value = unmask_text(cell.value, mapper)
    wb.save(out_path)


def process_pptx(in_path, out_path, mapper=None, mode='mask'):
    prs = Presentation(in_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if mode == 'mask':
                        run.text = mask_text(run.text, mapper, in_path)
                    else:
                        run.text = unmask_text(run.text, mapper)
    prs.save(out_path)


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)


def walk_and_process(in_folder, out_folder, mapper, mode='mask'):
    ensure_dir(out_folder)
    exts = {'.docx': process_docx, '.xlsx': process_xlsx, '.pptx': process_pptx}
    for root, _, files in os.walk(in_folder):
        rel = os.path.relpath(root, in_folder)
        out_sub = os.path.join(out_folder, rel) if rel != '.' else out_folder
        ensure_dir(out_sub)
        for f in files:
            name, ext = os.path.splitext(f)
            ext = ext.lower()
            if ext in exts:
                in_path = os.path.join(root, f)
                out_path = os.path.join(out_sub, f"{name}{('_masked' if mode=='mask' else '_unmasked')}{ext}")
                try:
                    print(f"Processing {in_path} -> {out_path}")
                    exts[ext](in_path, out_path, mapper=mapper, mode=mode)
                except Exception as e:
                    print(f"Failed {in_path}: {e}")


def main():
    parser = argparse.ArgumentParser(description='Batch mask/unmask docx/xlsx/pptx with mapping CSV')
    subparsers = parser.add_subparsers(dest='cmd')

    p_mask = subparsers.add_parser('mask')
    p_mask.add_argument('in_folder')
    p_mask.add_argument('out_folder')

    p_unmask = subparsers.add_parser('unmask')
    p_unmask.add_argument('in_folder')
    p_unmask.add_argument('out_folder')
    p_unmask.add_argument('mapping_csv')

    args = parser.parse_args()
    if args.cmd == 'mask':
        in_folder = args.in_folder
        out_folder = args.out_folder
        ensure_dir(out_folder)
        mapper = Mapper()
        walk_and_process(in_folder, out_folder, mapper, mode='mask')
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        csv_name = f'mapping_{ts}.csv'
        csv_path = os.path.join(out_folder, csv_name)
        mapper.save_csv(csv_path)
        print(f"Mapping saved to: {csv_path}")
    elif args.cmd == 'unmask':
        in_folder = args.in_folder
        out_folder = args.out_folder
        mapping_csv = args.mapping_csv
        ensure_dir(out_folder)
        mapper = Mapper.load_from_csv(mapping_csv)
        walk_and_process(in_folder, out_folder, mapper, mode='unmask')
        print('Unmask completed')
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
