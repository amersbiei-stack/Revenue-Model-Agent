"""Extract VBA modules from the workbook and list every Sub/Function defined."""
import re
from pathlib import Path
from oletools.olevba import VBA_Parser

path = Path(r"C:\Users\amers\Downloads\Revenue Model\Revenue Model - 04062026 (Internal).xlsm")
vp = VBA_Parser(str(path))

if not vp.detect_vba_macros():
    print("No VBA macros found.")
    raise SystemExit(0)

helpers_expected = [
    "Step2_RollFormulas_TwoMonthsBack_SourceToTarget_Reliable",
    "GetFileDateFromWorkbookName_MMDDYYYY",
    "GetRowBlocks",
    "FindMonthColumnInRow4_Robust",
    "RollFormulasByBlocks_DragFix",
    "RestoreAppState",
    "TimedPopup",
]

defs_found = {}   # name -> (module, kind, signature line)
all_modules = {}  # module name -> full source

# Regex to match Sub/Function definitions
proc_re = re.compile(
    r"^\s*(Public|Private|Friend)?\s*(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)

for (filename, stream_path, vba_filename, vba_code) in vp.extract_macros():
    # vba_filename looks like "Modules/Module1" etc.
    module = vba_filename
    all_modules[module] = vba_code
    for line in vba_code.splitlines():
        m = proc_re.match(line)
        if m:
            kind = m.group(3).strip()
            name = m.group(5)
            defs_found.setdefault(name, (module, kind, line.strip()))

print(f"=== VBA modules found ({len(all_modules)}) ===")
for m in all_modules:
    loc = len(all_modules[m].splitlines())
    print(f"  {m}  ({loc} lines)")

print(f"\n=== All Subs/Functions ({len(defs_found)}) ===")
for name in sorted(defs_found):
    module, kind, sig = defs_found[name]
    print(f"  [{kind}]  {name}   <{module}>")

print("\n=== Helper dependency check ===")
for h in helpers_expected:
    if h in defs_found:
        module, kind, sig = defs_found[h]
        print(f"  OK    {h}   ({kind} in {module})")
    else:
        # try case-insensitive fallback
        matches = [n for n in defs_found if n.lower() == h.lower()]
        if matches:
            nm = matches[0]
            module, kind, sig = defs_found[nm]
            print(f"  CASE  {h}  -> found as {nm} ({kind} in {module})")
        else:
            print(f"  MISS  {h}   <<< not defined anywhere")
