#!/usr/bin/env python3
"""Rollback to previous version.
Usage: python rollback.py <output_dir> --list | --to N"""
import sys, os, glob, re

def rollback(out_dir, target_v=None, list_only=False):
    files = glob.glob(os.path.join(out_dir, '*_v*.docx'))
    versions = []
    for f in files:
        m = re.search(r'_v(\d+)\.docx$', f)
        if m: versions.append((int(m.group(1)), f))
    versions.sort(key=lambda x: x[0])
    if list_only or target_v is None:
        for v, path in versions:
            size_kb = os.path.getsize(path) / 1024
            print(f'  v{v:3d}  {size_kb:6.0f}KB  {os.path.basename(path)}')
        return
    for v, path in versions:
        if v == target_v:
            print(f'Rollback to: {os.path.basename(path)}')
            return path
    print(f'Version v{target_v} not found')

if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('dir')
    p.add_argument('--list','-l',action='store_true')
    p.add_argument('--to',type=int)
    args = p.parse_args()
    rollback(args.dir, args.to, args.list)
