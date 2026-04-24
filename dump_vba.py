"""Dump the contents of Module1 and Module2 for review."""
from pathlib import Path
from oletools.olevba import VBA_Parser

path = Path(r"C:\Users\amers\Downloads\Revenue Model\Revenue Model - 04062026 (Internal).xlsm")
vp = VBA_Parser(str(path))

targets = {"Module1.bas", "Module2.bas"}

for (filename, stream_path, vba_filename, vba_code) in vp.extract_macros():
    if vba_filename in targets:
        print(f"================ {vba_filename} ================")
        print(vba_code)
        print()
