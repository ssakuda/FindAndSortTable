Option Explicit

' =============================
'  USER CONFIGURATION SETTINGS
' =============================

' --------------------------------
'  BOM/Weldment Cut List Creation
' --------------------------------
'Templates to use
Public Const BOM_TEMPLATE_PATH As String = "C:/Program Files/SOLIDWORKS Corp/SOLIDWORKS/lang/english/bom-standard.sldbomtbt"
Public Const WCL_TEMPLATE_PATH As String = "C:/Program Files/SOLIDWORKS Corp/SOLIDWORKS/lang/english/cut list.sldwldtbt"

'Placement on the sheet relative to the anchor point
'Available options
'(1) swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft
'(2) swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight
'(3) swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomLeft
'(4) swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight
Public Const BOM_ANCHOR_POS As Integer = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft
Public Const WCL_ANCHOR_POS As Integer = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft

' ----------------------------
'  BOM/Weldment Cut List Sort
' ----------------------------
'If you need custom sort rules, set to true
'Use SortBomCustom and SortCutlistCustom to write custom rules
Public Const USE_CUSTOM_BOM_SORT As Boolean = False
Public Const USE_CUSTOM_WCL_SORT As Boolean = False

'Set up to three columns to sort a BOM
'Leftmost column is 0
'Unused sort columns should be set to -1 (Do not delete any constant or this macro will fail to run)
Public Const SORT_BOM_COL_1 As Integer = 1
Public Const SORT_BOM_COL_2 As Integer = -1
Public Const SORT_BOM_COL_3 As Integer = -1

'Set which column to sort a cutlist
'Leftmost column is 0
Public Const SORT_WCL_COL As Integer = 2

' ============================
'  END CONFIGURATION SETTINGS
' ============================