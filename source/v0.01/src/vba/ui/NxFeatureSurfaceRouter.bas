Attribute VB_Name = "NxFeatureSurfaceRouter"
Option Explicit

Private mDataNormalizeForm As FNxDataNormalize
Private mDrawTableForm As FNxDrawTable
Private mPictureInsertForm As FNxPictureInsert
Private mRoleStyleForm As FNxRoleStyle

Public Sub NxRouteFeatureSurface(ByVal definition As CNxFeatureDefinition)
    NxDistributionEnsureExecutable
    If definition Is Nothing Then NxRaiseContractError "Feature definition is required"
    If Not definition.IsSealed Then NxRaiseContractError "Feature definition must be sealed"
    Select Case definition.LaunchSurface
        Case "direct": NxRunDirectFeature definition
        Case "dialog": NxOpenFeatureDialog definition
        Case "hybrid": NxRouteHybridFeature definition
        Case "workbench": NxAiOpenInput definition.FeatureId
        Case Else: NxRaiseContractError "Unknown feature launch surface"
    End Select
End Sub

Public Sub NxRunDirectFeature(ByVal definition As CNxFeatureDefinition)
    NxDistributionEnsureExecutable
    Select Case definition.FeatureId
        Case NX_FEATURE_HANGUL_TABLE_SEND
            NxHangulSendSelection
        Case NX_FEATURE_DRAW_CLEAR_INNER
            NxDrawingRunDirect definition.FeatureId
        Case NX_FEATURE_FILE_MANNER_SAVE
            NxMannerSaveFromRibbon
        Case NX_FEATURE_FILE_RANGE_PNG
            NxFileSaveRangeFromRibbon
        Case NX_FEATURE_FILE_SHEET_COPY_SAVE, NX_FEATURE_FILE_RANGE_COPY_SAVE
            NxFileRunCopySaveFromRibbon definition.FeatureId
        Case NX_FEATURE_DATA_COPY_VISIBLE
            NxVisibleCellsCopyFromRibbon
        Case NX_FEATURE_DATA_PASTE_VISIBLE_VALUES
            NxVisibleCellsPasteValuesFromRibbon
        Case NX_FEATURE_DATA_AGE, NX_FEATURE_DATA_KOREAN_MONEY, _
             NX_FEATURE_DATA_PRIVACY_MASK, NX_FEATURE_DATA_PRIVACY_SCAN
            NxDataRunSpecialFromRibbon definition.FeatureId
        Case NX_FEATURE_UTIL_CALCULATOR
            NxCalculatorOpen
        Case NX_FEATURE_UTIL_SYMBOLS
            NxSymbolsOpen
        Case NX_FEATURE_UTIL_NAVIGATOR
            If Not NxHostShowNavigator() Then NxNavigatorOpenVbaFallback
        Case NX_FEATURE_UTIL_DOCUMENT_NAVIGATOR
            NxDocumentNavigatorOpen
        Case Else
            NxRaiseContractError "Unknown direct feature"
    End Select
End Sub

Public Sub NxOpenFeatureDialog(ByVal definition As CNxFeatureDefinition)
    NxDistributionEnsureExecutable
    Dim context As CNxFeatureDialogContext
    Set context = NxCreateFeatureDialogContext(definition)
    Select Case definition.FeatureId
        Case NX_FEATURE_HANGUL_PICTURE_SEND
            NxHangulOpenPictures
        Case NX_FEATURE_AI_SUMMARY, NX_FEATURE_AI_CLEAN, NX_FEATURE_AI_FORMULA, NX_FEATURE_AI_WRITE, NX_FEATURE_AI_IMAGE
            NxAiOpenInput definition.FeatureId
        Case NX_FEATURE_TPL_REGISTER_SHEET, NX_FEATURE_TPL_LIST, _
             NX_FEATURE_TPL_LOAD, NX_FEATURE_TPL_RENAME, NX_FEATURE_TPL_DELETE
            NxTemplateOpenFeatureDialog definition.FeatureId
        Case NX_FEATURE_DATA_NORMALIZE
            Set mDataNormalizeForm = New FNxDataNormalize
            mDataNormalizeForm.BindFeatureContext context
            mDataNormalizeForm.Show vbModeless
        Case NX_FEATURE_DATA_AGE, NX_FEATURE_DATA_KOREAN_MONEY, NX_FEATURE_DATA_PRIVACY_MASK
            NxDataRunSpecialFromRibbon definition.FeatureId
        Case NX_FEATURE_DATA_UNIQUE_COUNT, NX_FEATURE_DATA_DUPLICATE_LIST
            NxDataOpenFeatureDialog definition.FeatureId
        Case NX_FEATURE_DRAW_TITLE_TABLE, NX_FEATURE_DRAW_BUSINESS_TABLE
            Set mDrawTableForm = New FNxDrawTable
            mDrawTableForm.BindFeatureContext context
            mDrawTableForm.Show vbModeless
        Case NX_FEATURE_DRAW_ROLE_STYLE
            Set mRoleStyleForm = New FNxRoleStyle
            mRoleStyleForm.BindFeatureContext context
            mRoleStyleForm.Show vbModeless
        Case NX_FEATURE_DRAW_INSERT_PICTURE, NX_FEATURE_DRAW_FIT_PICTURE
            Set mPictureInsertForm = New FNxPictureInsert
            mPictureInsertForm.BindFeatureContext context
            mPictureInsertForm.Show vbModeless
        Case NX_FEATURE_FILE_FOLDER_CREATE
            NxFolderCreateOpen context
        Case NX_FEATURE_FILE_CONSOLIDATE, _
             NX_FEATURE_FILE_CHART_PNG, _
             NX_FEATURE_FILE_WORKBOOK_COMPARE, "NX-FILE-SHEET-COMPARE", "NX-FILE-FILE-COMPARE", NX_FEATURE_FILE_BATCH_RENAME, _
             NX_FEATURE_FILE_SHEET_BATCH_RENAME, NX_FEATURE_FILE_PDF_CURRENT_SHEET, _
             NX_FEATURE_FILE_PDF_EACH_SHEET, NX_FEATURE_FILE_PDF_SELECTED_COMBINED, _
             NX_FEATURE_FILE_PDF_ALL_COMBINED, NX_FEATURE_FILE_PDF_SETTINGS
            NxFileOpenFeatureDialog definition.FeatureId
        Case Else
            NxRaiseContractError "Unknown feature dialog"
    End Select
End Sub

Public Sub NxOpenDateConversion()
    Dim context As CNxFeatureDialogContext
    Set context = NxCreateFeatureDialogContext(NxDataFeatureDefinition(NX_FEATURE_DATA_NORMALIZE))
    Set mDataNormalizeForm = New FNxDataNormalize
    mDataNormalizeForm.BindFeatureContext context
    mDataNormalizeForm.UseDateConversion
    mDataNormalizeForm.Show vbModeless
End Sub

Public Sub NxRouteHybridFeature(ByVal definition As CNxFeatureDefinition)
    If definition.FeatureId <> NX_FEATURE_DATA_FOCUS_CELL Then NxRaiseContractError "Unknown hybrid feature"
    If definition.StateResolver <> "NxFocusIsEnabled" Then NxRaiseContractError "Unknown hybrid state resolver"
    NxFocusToggle
End Sub

Private Function NxCreateFeatureDialogContext(ByVal definition As CNxFeatureDefinition) As CNxFeatureDialogContext
    Dim context As New CNxFeatureDialogContext
    context.Configure definition.FeatureId, definition.DialogId, definition.DialogVariant, definition.LaunchSurface = "workbench"
    Set NxCreateFeatureDialogContext = context
End Function
