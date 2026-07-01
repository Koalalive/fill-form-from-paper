#!/usr/bin/env python3
"""Audit bold runs - flag any exceeding 1 sentence or 100 chars."""
from docx import Document
import glob, os, re

def audit(docx_path):
    doc = Document(docx_path)
    issues = []
    for i, p in enumerate(doc.paragraphs):
        for j, run in enumerate(p.runs):
            if not run.font.bold: continue
            rt = run.text
            sents = [s for s in rt.split('。') if s.strip()]
            if len(sents) > 1 or len(rt) > 100:
                issues.append((i, j, len(sents), len(rt), rt[:50]))
    if issues:
        print(f'{len(issues)} oversize bold runs:')
        for pi, ri, sc, ch, txt in issues:
            print(f'  P[{pi}] run[{ri}]: {sc} sentences, {ch} chars')
    else:
        print('CLEAN')
    return len(issues) == 0

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        files = glob.glob('*_v*.docx')
        if files:
            files.sort(key=os.path.getmtime, reverse=True)
            audit(files[0])
    else:
        audit(sys.argv[1])
