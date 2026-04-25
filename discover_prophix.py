"""Discovery script for Prophix Analyzer CV2's COM API.

Usage:
  1. Open Excel + Revenue Model workbook MANUALLY (double-click the file).
  2. Wait for Prophix Analyzer CV2 to fully load — confirm by clicking the
     Insert ribbon and seeing Analyzer / Contributor / Analyzer in the
     Prophix group.
  3. Open a Command Prompt or PowerShell in this project folder.
  4. Run:  py discover_prophix.py
  5. Paste the entire output back to the agent.

The script attaches to the running Excel via COM (does NOT launch its own),
finds every Prophix-related entry in Application.COMAddIns, and dumps the
methods/properties exposed on each add-in's automation Object so we can
decide whether to drive it programmatically.
"""
from __future__ import annotations

import sys
import traceback


def _print(msg: str = "") -> None:
    print(msg, flush=True)


def _safe_str(v) -> str:
    try:
        return str(v)
    except Exception:
        return "<unprintable>"


def _enumerate_object_members(obj, label: str) -> None:
    _print(f"\n--- Members of {label} ---")
    if obj is None:
        _print("  (Object is None)")
        return

    # 1. Plain Python dir()
    try:
        members = sorted(set(dir(obj)))
    except Exception as e:
        members = []
        _print(f"  dir() failed: {e}")
    if members:
        _print(f"  dir() returned {len(members)} names:")
        for m in members:
            if m.startswith("_"):
                continue
            _print(f"    {m}")

    # 2. COM type info via _oleobj_ if available
    try:
        oleobj = getattr(obj, "_oleobj_", None)
        if oleobj is not None:
            try:
                ti = oleobj.GetTypeInfo()
                ta = ti.GetTypeAttr()
                _print(f"\n  COM TypeInfo (idl funcs={ta.cFuncs}, vars={ta.cVars}):")
                for i in range(ta.cFuncs):
                    fd = ti.GetFuncDesc(i)
                    name = ti.GetNames(fd.memid)[0]
                    nargs = fd.cParams
                    _print(f"    func [{i:>3}] memid=0x{fd.memid:08x} {name}(args={nargs})")
                for i in range(ta.cVars):
                    vd = ti.GetVarDesc(i)
                    name = ti.GetNames(vd.memid)[0]
                    _print(f"    prop [{i:>3}] memid=0x{vd.memid:08x} {name}")
            except Exception as e:
                _print(f"  GetTypeInfo() failed: {e}")
        else:
            _print("  (no _oleobj_ attribute)")
    except Exception as e:
        _print(f"  COM type-info inspection failed: {e}")


def main() -> int:
    try:
        import win32com.client
    except ImportError:
        _print("ERROR: pywin32 is not installed. Run:  py -m pip install -r requirements.txt")
        return 2

    _print("Attaching to running Excel via COM (GetActiveObject)...")
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        _print(f"ERROR: could not attach to a running Excel.Application instance.\n"
               f"       Make sure Excel is open with the workbook (double-click "
               f"the .xlsm), then re-run.\n"
               f"       Underlying error: {e}")
        return 3

    try:
        _print(f"Excel version: {app.Version}  build={getattr(app, 'Build', '?')}")
        _print(f"Excel.Workbooks.Count = {app.Workbooks.Count}")
        for i in range(1, app.Workbooks.Count + 1):
            try:
                wb = app.Workbooks.Item(i)
                _print(f"  WB[{i}] {wb.Name}  -> {wb.FullName}")
            except Exception as e:
                _print(f"  WB[{i}] inspection failed: {e}")
    except Exception as e:
        _print(f"WARNING: could not enumerate Workbooks: {e}")

    # ---- COM Add-Ins ----
    try:
        addins = app.COMAddIns
        count = addins.Count
    except Exception as e:
        _print(f"ERROR: COMAddIns inaccessible: {e}")
        return 4

    _print(f"\nCOM add-ins available: {count}")
    prophix_addins = []
    for i in range(1, count + 1):
        try:
            a = addins.Item(i)
            progid = _safe_str(getattr(a, "ProgID", ""))
            desc = _safe_str(getattr(a, "Description", ""))
            connect = bool(getattr(a, "Connect", False))
            guid = _safe_str(getattr(a, "Guid", ""))
            _print(f"  [{i}] ProgID={progid}")
            _print(f"      Desc  ={desc}")
            _print(f"      Guid  ={guid}")
            _print(f"      Connect={connect}")
            if "prophix" in progid.lower() or "prophix" in desc.lower():
                prophix_addins.append((progid, a))
        except Exception as e:
            _print(f"  [{i}] inspection failed: {e}")

    if not prophix_addins:
        _print("\n!!! No Prophix entries in COMAddIns !!!")
        _print("    Make sure Excel was launched MANUALLY (double-click the .xlsm),")
        _print("    not via this script or the agent. The Prophix Analyzer pane")
        _print("    or its ribbon buttons should be visible before running this.")
        return 5

    # ---- Inspect each Prophix add-in's automation object ----
    for progid, addin in prophix_addins:
        _print(f"\n========== {progid} ==========")
        try:
            obj = addin.Object
        except Exception as e:
            _print(f"  addin.Object access failed: {e}")
            continue
        if obj is None:
            _print("  addin.Object is None — no automation interface exposed.")
            continue
        _print(f"  addin.Object type: {type(obj)}")
        _enumerate_object_members(obj, f"{progid}.Object")

        # Heuristic probe: try common refresh-related method names without
        # arguments. Just to see which ones EXIST (we won't actually run them
        # blindly — the AttributeError vs other errors tells us a lot).
        probe_names = [
            "Refresh", "RefreshAll", "RefreshAllSheets", "RefreshSheet",
            "RefreshAllData", "Recalculate", "ProcessRefresh", "Process",
            "RefreshAllReports", "RefreshAllAnalyses", "RefreshAnalyzer",
            "Update", "UpdateAll", "ShowPane", "ShowTaskPane",
            "OpenAnalyzer", "OpenPane",
        ]
        _print(f"\n  Method-existence probe (no calls made):")
        for name in probe_names:
            present = hasattr(obj, name)
            _print(f"    {'YES' if present else ' no'}  obj.{name}")

    _print("\nDone. Copy this entire output back to the agent so we can decide.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(1)
