"""Compatibility wrapper for file dialogs in frozen builds.

Tries tkinter.filedialog first. If unavailable, falls back to Tcl/Tk dialog
commands directly so the app can still open/save/browse files.
"""

from __future__ import annotations

import tkinter as tk
from typing import Any, Dict, Iterable, Tuple

try:
    from tkinter import filedialog as _native_filedialog  # type: ignore
except Exception:  # pragma: no cover
    _native_filedialog = None


def _to_tcl_filetypes(filetypes: Iterable[Tuple[str, str]]):
    normalized = []
    for item in filetypes:
        if not isinstance(item, (tuple, list)) or len(item) < 2:
            continue
        normalized.append((str(item[0]), str(item[1])))
    return tuple(normalized)


def _build_dialog_args(options: Dict[str, Any]):
    args = []
    for key, value in options.items():
        if value in (None, ""):
            continue

        option_name = f"-{key}"
        if key == "filetypes":
            value = _to_tcl_filetypes(value)
            if not value:
                continue
        elif key == "mustexist":
            value = 1 if bool(value) else 0
        else:
            value = str(value)

        args.extend((option_name, value))

    return args


def _run_tcl_dialog(command: str, **options):
    parent = options.pop("parent", None)
    created_root = None

    if parent is not None:
        tkapp = parent.tk
    else:
        root = tk._default_root
        if root is None:
            created_root = tk.Tk()
            created_root.withdraw()
            root = created_root
        tkapp = root.tk

    try:
        result = tkapp.call(command, *_build_dialog_args(options))
        return str(result) if result else ""
    finally:
        if created_root is not None:
            created_root.destroy()


def askopenfilename(**options):
    if _native_filedialog is not None:
        return _native_filedialog.askopenfilename(**options)
    return _run_tcl_dialog("tk_getOpenFile", **options)


def asksaveasfilename(**options):
    if _native_filedialog is not None:
        return _native_filedialog.asksaveasfilename(**options)
    return _run_tcl_dialog("tk_getSaveFile", **options)


def askdirectory(**options):
    if _native_filedialog is not None:
        return _native_filedialog.askdirectory(**options)
    return _run_tcl_dialog("tk_chooseDirectory", **options)
