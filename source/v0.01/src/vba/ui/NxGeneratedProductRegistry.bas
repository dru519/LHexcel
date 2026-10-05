Attribute VB_Name = "NxGeneratedProductRegistry"
Option Explicit
' GENERATED FILE - edit contracts/feature-contract.json and ribbon-shell-contract.json
' Feature: NX-AI-SUMMARY
' Feature: NX-AI-CLEAN
' Feature: NX-AI-FORMULA
' Feature: NX-AI-WRITE
' Feature: NX-AI-IMAGE
' Feature: NX-TPL-REGISTER-SHEET
' Feature: NX-TPL-LIST
' Feature: NX-TPL-LOAD
' Feature: NX-TPL-RENAME
' Feature: NX-TPL-DELETE
' Feature: NX-DATA-NORMALIZE
' Feature: NX-DATA-UNIQUE-COUNT
' Feature: NX-DATA-DUPLICATE-LIST
' Feature: NX-DATA-AGE
' Feature: NX-DATA-KOREAN-MONEY
' Feature: NX-DATA-PRIVACY-MASK
' Feature: NX-DATA-PRIVACY-SCAN
' Feature: NX-DATA-FOCUS-CELL
' Feature: NX-DATA-COPY-VISIBLE
' Feature: NX-DATA-PASTE-VISIBLE-VALUES
' Feature: NX-HANGUL-TABLE-SEND
' Feature: NX-HANGUL-PICTURE-SEND
' Feature: NX-DRAW-TITLE-TABLE
' Feature: NX-DRAW-BUSINESS-TABLE
' Feature: NX-DRAW-ROLE-STYLE
' Feature: NX-DRAW-CLEAR-INNER
' Feature: NX-DRAW-FIT-PICTURE
' Feature: NX-DRAW-INSERT-PICTURE
' Feature: NX-FILE-FOLDER-CREATE
' Feature: NX-FILE-CONSOLIDATE
' Feature: NX-FILE-MANNER-SAVE
' Feature: NX-FILE-RANGE-PNG
' Feature: NX-FILE-CHART-PNG
' Feature: NX-FILE-SHEET-COPY-SAVE
' Feature: NX-FILE-RANGE-COPY-SAVE
' Feature: NX-FILE-WORKBOOK-COMPARE
' Feature: NX-FILE-BATCH-RENAME
' Feature: NX-FILE-SHEET-BATCH-RENAME
' Feature: NX-FILE-PDF-CURRENT-SHEET
' Feature: NX-FILE-PDF-EACH-SHEET
' Feature: NX-FILE-PDF-SELECTED-COMBINED
' Feature: NX-FILE-PDF-ALL-COMBINED
' Feature: NX-FILE-PDF-SETTINGS
' Feature: NX-UTIL-CALCULATOR
' Feature: NX-UTIL-SYMBOLS
' Feature: NX-UTIL-NAVIGATOR
' Feature: NX-UTIL-DOCUMENT-NAVIGATOR
' Feature: NX-FILE-SHEET-COMPARE
' Feature: NX-FILE-FILE-COMPARE
Public Function NxCreateGeneratedProductRegistry() As CNxFeatureRegistry
    Dim registry As CNxFeatureRegistry
    Set registry = CreateBaseRegistry()
    NxAiRegisterFeatures registry
    NxTemplateRegisterFeatures registry
    NxDataRegisterFeatures registry
    NxHangulRegisterFeatures registry
    NxDrawingRegisterFeatures registry
    NxFileRegisterFeatures registry
    NxCalculatorRegisterFeatures registry
    NxSymbolsRegisterFeatures registry
    NxNavigatorRegisterFeatures registry
    Set NxCreateGeneratedProductRegistry = registry
End Function

Public Function NxProductSurfaceGroupCount() As Long
    NxProductSurfaceGroupCount = 12
End Function

Public Function NxProductSurfaceFeatureCount() As Long
    NxProductSurfaceFeatureCount = 49
End Function

Public Function NxGeneratedReleaseFeatureIds() As Variant
    NxGeneratedReleaseFeatureIds = Array("NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE", "NX-TPL-REGISTER-SHEET", "NX-TPL-LIST", "NX-TPL-LOAD", _
        "NX-TPL-RENAME", "NX-TPL-DELETE", "NX-DATA-NORMALIZE", "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST", "NX-DATA-AGE", "NX-DATA-KOREAN-MONEY", "NX-DATA-PRIVACY-MASK", _
        "NX-DATA-PRIVACY-SCAN", "NX-DATA-FOCUS-CELL", "NX-DATA-COPY-VISIBLE", "NX-DATA-PASTE-VISIBLE-VALUES", "NX-HANGUL-TABLE-SEND", "NX-HANGUL-PICTURE-SEND", "NX-DRAW-TITLE-TABLE", "NX-DRAW-BUSINESS-TABLE", _
        "NX-DRAW-ROLE-STYLE", "NX-DRAW-CLEAR-INNER", "NX-DRAW-FIT-PICTURE", "NX-DRAW-INSERT-PICTURE", "NX-FILE-FOLDER-CREATE", "NX-FILE-CONSOLIDATE", "NX-FILE-MANNER-SAVE", "NX-FILE-RANGE-PNG", _
        "NX-FILE-CHART-PNG", "NX-FILE-SHEET-COPY-SAVE", "NX-FILE-RANGE-COPY-SAVE", "NX-FILE-WORKBOOK-COMPARE", "NX-FILE-BATCH-RENAME", "NX-FILE-SHEET-BATCH-RENAME", "NX-FILE-PDF-CURRENT-SHEET", "NX-FILE-PDF-EACH-SHEET", _
        "NX-FILE-PDF-SELECTED-COMBINED", "NX-FILE-PDF-ALL-COMBINED", "NX-FILE-PDF-SETTINGS", "NX-UTIL-CALCULATOR", "NX-UTIL-SYMBOLS", "NX-UTIL-NAVIGATOR", "NX-UTIL-DOCUMENT-NAVIGATOR", "NX-FILE-SHEET-COMPARE", _
        "NX-FILE-FILE-COMPARE")
End Function
