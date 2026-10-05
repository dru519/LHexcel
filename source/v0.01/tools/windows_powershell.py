"""Prepare a deterministic Windows PowerShell 5.1 child environment."""
from __future__ import annotations

import os
from collections.abc import Mapping
from pathlib import Path


def windows_powershell_environment(base: Mapping[str, str] | None = None) -> dict[str, str]:
    """Remove PowerShell Core module roots that can shadow Windows inbox modules."""
    environment = dict(os.environ if base is None else base)
    module_roots: list[str] = []
    seen: set[str] = set()
    for raw in environment.get("PSModulePath", "").split(os.pathsep):
        value = raw.strip()
        if not value or r"\windowspowershell\modules" not in value.casefold():
            continue
        key = value.casefold().rstrip("\\/")
        if key not in seen:
            module_roots.append(value)
            seen.add(key)
    system_root = environment.get("SystemRoot", r"C:\Windows")
    inbox = str(Path(system_root) / "System32" / "WindowsPowerShell" / "v1.0" / "Modules")
    inbox_key = inbox.casefold().rstrip("\\/")
    if inbox_key not in seen:
        module_roots.append(inbox)
    environment["PSModulePath"] = os.pathsep.join(module_roots)
    return environment
