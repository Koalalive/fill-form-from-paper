#!/usr/bin/env python3
"""Pre-flight: validate template before filling.
Usage: python preflight.py <template.docx>"""
import sys, re
from docx import Document

def preflight(template_path):
    doc = Document(template_path)
    issues = []
    expected_labels = ['project','section','content']
    paras = doc.paragraphs
    empty_paras = sum(1 for p in paras if not p.text.strip())
    print(f'Paragraphs: {len(paras)} ({empty_paras} empty)')
    print(f'Tables: {len(doc.tables)}')
    s = doc.sections[0]
    w_mm = s.page_width / 360000 * 25.4
    h_mm = s.page_height / 360000 * 25.4
    if abs(w_mm - 210) > 5 or abs(h_mm - 297) > 5:
        issues.append(f'Page size {w_mm:.0f}x{h_mm:.0f}mm (expected A4)')
    if issues:
        for i in issues: print(f'  ISSUE: {i}')
        return False
    print('Status: READY')
    return True

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python preflight.py <template.docx>')
        sys.exit(1)
    sys.exit(0 if preflight(sys.argv[1]) else 1)
