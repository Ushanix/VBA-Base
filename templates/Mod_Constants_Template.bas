Option Explicit

' ============================================
' Module   : Mod_Constants
' Layer    : Common / Config
' Purpose  : Centralized constants for <ProjectName>
' Version  : 1.0.0
' Created  : <YYYY-MM-DD>
' Note     : Copy this template and replace placeholders
'            for each new project based on VBA-Base
' ============================================

' ============================================
' Tbl Marker Prefix
' ============================================
Public Const TBL_MARKER_PREFIX As String = "Tbl:"

' ============================================
' Tbl Marker Names (PascalCase)
' ============================================

' --- UI_Dashboard ---
Public Const TBL_INDEX_HEADER As String = "IndexHeader"
Public Const TBL_UI_OPERATIONS As String = "UI_Operations"
Public Const TBL_UI_STATUS As String = "UI_Status"
Public Const TBL_UI_SHEET_INDEX As String = "UI_SheetIndex"

' --- Add your project-specific markers below ---
' Public Const TBL_XXX As String = "XXX"

' ============================================
' Sheet Name Prefixes
' ============================================
Public Const PREFIX_TEMPLATE As String = "TPL_"
Public Const PREFIX_DEFINITION As String = "DEF_"
Public Const PREFIX_UI As String = "UI_"
Public Const PREFIX_LOG As String = "LOG_"

' --- Add your project-specific prefix ---
' Public Const PREFIX_<DOMAIN> As String = "<PREFIX>_"

' ============================================
' Fixed Sheet Names
' ============================================
Public Const SHEET_UI_DASHBOARD As String = "UI_Dashboard"
Public Const SHEET_DEF_PARAMETER As String = "DEF_Parameter"
Public Const SHEET_LOG_UPDATE_HISTORY As String = "LOG_UpdateHistory"

' --- Add your project-specific sheet names ---
' Public Const SHEET_XXX As String = "XXX"

' ============================================
' Parameter Keys (from DEF_Parameter)
' ============================================
Public Const PARAM_VAULT_ROOT As String = "VAULT_ROOT"
Public Const PARAM_OUTPUT_ROOT As String = "OUTPUT_ROOT"
Public Const PARAM_OUTPUT_MODE As String = "OUTPUT_MODE"
Public Const PARAM_TEMPLATE_ROOT As String = "TEMPLATE_ROOT"
Public Const PARAM_DATA_EXPORT_PATH As String = "DATA_EXPORT_PATH"

' --- Add your project-specific parameters ---
' Public Const PARAM_XXX As String = "XXX"

' ============================================
' Default Values
' ============================================
Public Const DEFAULT_SORT_ORDER As Long = 9999
Public Const DEFAULT_STATUS As String = "plan"

' ============================================
' Key-Value Table Names
' ============================================
Public Function GetKeyValueTables() As Variant
    GetKeyValueTables = Array( _
        "IndexHeader", "UI_Status", "DEF_Parameter")
    ' Add your project-specific KV tables to the array
End Function
