from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping


FEATURE_CONTRACT_PATH = "contracts/feature-contract.json"
PRODUCT_MANIFEST_PATH = "build/manifests/Product.json"
EXPECTED_HANGUL_RELEASE = {"NX-HANGUL-TABLE-SEND", "NX-HANGUL-PICTURE-SEND"}
FORBIDDEN_REMOVED_IDS = {
    "NX-HWPX-PREVIEW", "NX-HWPX-EXPORT-OPEN", "NX-DATA-ROSTER-CREATE", "NX-DATA-ROSTER-UPDATE",
}
ALLOWED_URL_LITERALS = {
    "http://schemas.microsoft.com/office/2009/07/customui",
}
AI_UI_PATH = "src/vba/features/ai/NxAiUi.bas"
AI_FORM_PATH = "src/vba/features/ai/ui/FNxAi.vba"
AI_HOMEPAGES = {
    "ChatGPT": "https://chatgpt.com/",
    "Claude": "https://claude.ai/",
    "Gemini": "https://gemini.google.com/",
}
TRANSPORT_PATTERNS = {
    "WINHTTP": re.compile(r"\bwinhttp(?:\.winhttprequest(?:\.\d+\.\d+)?)?\b", re.I),
    "XMLHTTP": re.compile(r"\b(?:server)?xmlhttp(?:\.\d+\.\d+)?\b", re.I),
    "URLDOWNLOADTOFILE": re.compile(r"\burldownloadtofile\b", re.I),
    "INTERNETOPEN": re.compile(r"\binternetopen(?:url)?\b", re.I),
    "POWERSHELL_WEB": re.compile(r"\binvoke-(?:webrequest|restmethod)\b", re.I),
    "DOTNET_HTTP": re.compile(r"\b(?:system\.net\.)?(?:webclient|httpclient)\b", re.I),
    "BITS": re.compile(r"\bstart-bitstransfer\b", re.I),
    "CURL_WGET": re.compile(r"(?<![a-z0-9_])(?:curl(?:\.exe)?|wget(?:\.exe)?)(?![a-z0-9_])", re.I),
}
URL_PATTERN = re.compile(r"https?://[^\s\"'<>]+", re.I)


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def violation(code: str, detail: str, path: str | None = None) -> dict[str, str]:
    result = {"code": code, "detail": detail}
    if path is not None:
        result["path"] = path
    return result


def ai_homepage_flow_is_closed(sources: Mapping[str, str]) -> bool:
    """Attest the two narrow URL/opening procedures, not arbitrary browser delivery."""
    ui = sources.get(AI_UI_PATH, "")
    form = sources.get(AI_FORM_PATH, "")

    def procedure_lines(source: str, signature: str, ending: str) -> list[str]:
        matches = re.findall(
            rf"^{re.escape(signature)}\r?\n.*?^{re.escape(ending)}\s*$",
            source,
            re.M | re.S,
        )
        if len(matches) != 1:
            return []
        return [line.strip() for line in matches[0].splitlines() if line.strip()]

    url_signature = "Public Function NxAiTargetUrl(ByVal targetName As String) As String"
    expected_url = [
        url_signature,
        "Select Case targetName",
        *[f'Case "{name}": NxAiTargetUrl = "{url}"' for name, url in AI_HOMEPAGES.items()],
        'Case Else: NxRaiseContractError "대상 AI를 선택하세요."',
        "End Select",
        "End Function",
    ]
    open_signature = "Public Function NxAiOpenTargetAfterCopy(ByVal targetName As String, ByVal copyState As NxFrameState) As Boolean"
    expected_open = [
        open_signature,
        "If copyState <> NxFrameSuccess Then Exit Function",
        "On Error GoTo Failed",
        "ThisWorkbook.FollowHyperlink Address:=NxAiTargetUrl(targetName), NewWindow:=True",
        "NxAiOpenTargetAfterCopy = True",
        "Failed:",
        "Err.Clear",
        "End Function",
    ]
    execute = procedure_lines(form, "Private Sub cmdExecute_Click()", "End Sub")
    copy_then_open = [
        "copyState = NxAiShowExecutionPlan(mSession)",
        "If copyState = NxFrameSuccess Then",
        "If NxAiOpenTargetAfterCopy(mSession.TargetAI, copyState) Then",
    ]
    all_source = "\n".join(sources.values())
    return (
        procedure_lines(ui, url_signature, "End Function") == expected_url
        and procedure_lines(ui, open_signature, "End Function") == expected_open
        and URL_PATTERN.findall(ui) == list(AI_HOMEPAGES.values())
        and "\n".join(copy_then_open) in "\n".join(execute)
        and sum(line.startswith("copyState =") for line in execute) == 1
        and len(re.findall(r"\bNxAiOpenTargetAfterCopy\b", all_source, re.I)) == 3
    )


def audit_documents(
    feature_contract: Mapping[str, Any],
    product_manifest: Mapping[str, Any],
    module_sources: Mapping[str, bytes],
) -> dict[str, Any]:
    violations: list[dict[str, str]] = []
    features = {
        str(row.get("id", "")): row
        for row in feature_contract.get("features", [])
        if isinstance(row, Mapping)
    }
    lineage = feature_contract.get("lineage", {})
    release_ids = [str(value) for value in lineage.get("release_feature_ids", [])]
    product_scope = product_manifest.get("release_scope", {})
    product_release_ids = [str(value) for value in product_scope.get("included_feature_ids", [])]
    product_excluded_ids = {str(value) for value in product_scope.get("excluded_feature_ids", [])}
    retired_ids = [str(value) for value in lineage.get("retired_feature_ids", [])]

    release_surface_exact = (
        release_ids == product_release_ids
        and len(release_ids) == len(set(release_ids))
        and int(product_scope.get("feature_count", -1)) == len(release_ids)
        and all(feature_id in features for feature_id in release_ids)
        and set(release_ids) == set(features)
        and not product_excluded_ids
    )
    if not release_surface_exact:
        violations.append(
            violation(
                "RELEASE_SURFACE",
                "feature-contract lineage and Product release scope must be one exact ordered surface",
            )
        )

    release_network_safe = True
    release_ai_local = True
    for feature_id in release_ids:
        feature = features.get(feature_id)
        privacy = feature.get("privacy", {}) if isinstance(feature, Mapping) else {}
        if privacy.get("network_transfer") is not False:
            release_network_safe = False
            violations.append(
                violation(
                    "RELEASE_NETWORK_TRANSFER",
                    "release feature must declare network_transfer=false",
                    feature_id,
                )
            )
        if privacy.get("ai_processing") is not False:
            release_ai_local = False
            violations.append(
                violation(
                    "RELEASE_AI_PROCESSING",
                    "release feature must declare ai_processing=false",
                    feature_id,
                )
            )

    hangul_release_ids = {
        feature_id
        for feature_id in release_ids
        if isinstance(features.get(feature_id), Mapping)
        and features[feature_id].get("owner") == "hangul"
    }
    hangul_surface_exact = hangul_release_ids == EXPECTED_HANGUL_RELEASE
    if not hangul_surface_exact:
        violations.append(
            violation(
                "HANGUL_RELEASE_SURFACE",
                "아래한글 표 전송하기 must be the only direct Hangul release feature",
            )
        )

    removed_surface_absent = (
        not retired_ids
        and FORBIDDEN_REMOVED_IDS.isdisjoint(features)
        and FORBIDDEN_REMOVED_IDS.isdisjoint(product_excluded_ids)
    )
    if not removed_surface_absent:
        violations.append(
            violation(
                "REMOVED_FEATURE_SURFACE",
                "removed HWPX and roster features must be absent from product contracts",
            )
        )

    module_rows = product_manifest.get("modules", [])
    expected_paths = [str(row.get("path", "")) for row in module_rows if isinstance(row, Mapping)]
    manifest_hash_bound = len(expected_paths) == len(set(expected_paths)) and set(expected_paths) == set(module_sources)
    decoded_sources: dict[str, str] = {}
    if not manifest_hash_bound:
        violations.append(
            violation(
                "PRODUCT_MODULE_SET",
                "Product manifest module set and inspected source set must match exactly",
            )
        )
    for row in module_rows:
        if not isinstance(row, Mapping):
            manifest_hash_bound = False
            continue
        path = str(row.get("path", ""))
        raw = module_sources.get(path)
        if raw is None:
            manifest_hash_bound = False
            violations.append(violation("PRODUCT_MODULE_MISSING", "Product module is missing", path))
            continue
        actual_sha = sha256_bytes(raw)
        if actual_sha != row.get("sha256"):
            manifest_hash_bound = False
            violations.append(
                violation("PRODUCT_MODULE_HASH", "Product module SHA-256 binding changed", path)
            )
        try:
            decoded_sources[path] = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            manifest_hash_bound = False
            violations.append(
                violation("PRODUCT_MODULE_ENCODING", "Product module is not strict UTF-8", path)
            )

    homepage_flow_closed = ai_homepage_flow_is_closed(decoded_sources)
    no_network_transport = True
    for path, source in decoded_sources.items():
        for label, pattern in TRANSPORT_PATTERNS.items():
            if pattern.search(source):
                no_network_transport = False
                violations.append(
                    violation(
                        "NETWORK_TRANSPORT_SOURCE",
                        f"forbidden network transport token detected: {label}",
                        path,
                    )
                )
        # Documentation-only comment lines cannot initiate a transfer.
        # Keep strings and inline code intact, including apostrophes inside strings.
        executable_url_source = "\n".join(
            line for line in source.splitlines()
            if not line.lstrip().startswith("'")
            and not re.match(r"^\s*Rem(?:\s|$)", line, re.I)
        )
        for literal in URL_PATTERN.findall(executable_url_source):
            approved_homepage = (
                path == AI_UI_PATH
                and homepage_flow_closed
                and literal in AI_HOMEPAGES.values()
            )
            if literal not in ALLOWED_URL_LITERALS and not approved_homepage:
                no_network_transport = False
                violations.append(
                    violation(
                        "NETWORK_URL_LITERAL",
                        "HTTP(S) literal is not a namespace or an attested fixed AI homepage",
                        path,
                    )
                )

    ai_paths = [path for path in decoded_sources if path.startswith("src/vba/features/ai/")]
    ai_source = "\n".join(decoded_sources[path] for path in sorted(ai_paths))
    ai_clipboard_only = (
        ai_source.count('DeliveryMode = "CLIPBOARD_ONLY"') >= 2
        and '"ai_prompt_clipboard"' in ai_source
        and '"local_clipboard"' in ai_source
        and homepage_flow_closed
        and len(re.findall(r"\bFollowHyperlink\b", ai_source, re.I)) == 1
        and all(
            not re.search(rf"\b{token}(?:A|W)?\b", ai_source, re.I)
            for token in ("NxBrowserOpen", "NxExternalProcess", "ShellExecute", "Shell")
        )
    )
    if not ai_clipboard_only:
        violations.append(
            violation(
                "AI_DELIVERY_SURFACE",
                "AI data delivery must be clipboard-only; only the closed homepage flow after successful copy is allowed",
            )
        )

    checks = {
        "release_surface_exact": release_surface_exact,
        "release_network_transfer_false": release_network_safe,
        "release_ai_processing_false": release_ai_local,
        "hangul_table_send_only": hangul_surface_exact,
        "removed_features_absent": removed_surface_absent,
        "product_modules_hash_bound": manifest_hash_bound,
        "network_transport_tokens_absent": no_network_transport,
        "ai_clipboard_only": ai_clipboard_only,
    }
    violations.sort(key=lambda row: (row["code"], row.get("path", ""), row["detail"]))
    return {
        "schema_version": 2,
        "status": "PASS" if all(checks.values()) and not violations else "FAIL",
        "release_feature_count": len(release_ids),
        "release_feature_ids": release_ids,
        "hangul_release_feature_ids": sorted(hangul_release_ids),
        "removed_feature_ids": sorted(FORBIDDEN_REMOVED_IDS),
        "scanned_module_count": len(decoded_sources),
        "checks": checks,
        "violations": violations,
    }


def load_json(path: Path) -> tuple[dict[str, Any], bytes]:
    raw = path.read_bytes()
    return json.loads(raw.decode("utf-8-sig")), raw


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    root = Path(args.root).resolve()
    feature_path = root / FEATURE_CONTRACT_PATH
    product_path = root / PRODUCT_MANIFEST_PATH
    feature_contract, feature_raw = load_json(feature_path)
    product_manifest, product_raw = load_json(product_path)
    module_sources = {
        str(row["path"]): (root / str(row["path"])).read_bytes()
        for row in product_manifest.get("modules", [])
    }
    evidence = audit_documents(feature_contract, product_manifest, module_sources)
    evidence["feature_contract_sha256"] = sha256_bytes(feature_raw)
    evidence["product_manifest_sha256"] = sha256_bytes(product_raw)
    evidence["created_utc"] = datetime.now(timezone.utc).isoformat()

    output = Path(args.output)
    if not output.is_absolute():
        output = root / output
    output = output.resolve()
    output.relative_to(root)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.tmp")
    temporary.write_text(
        json.dumps(evidence, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(output)
    print(f"status={evidence['status']}")
    print(f"output={output}")
    return 0 if evidence["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
