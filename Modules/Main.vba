' ==========================================================
' Copyright(C) 2026 Shelley Sakuda
' MIT License
' Repository    https://github.com/ssakuda/FindAndSortTable
' Contact       ssakuda+github@gmail.com
' ==========================================================

Option Explicit

Sub FindAndSortTable()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim swView As View
    Dim swCurrentTableAnno As TableAnnotation
    Dim sortConfig As SortSettings
    Dim config As String
    Dim response As Boolean

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    'Checks drawing
    If Not IsValidDrawing(swModel) Then
        Exit Sub
    End If
     
    'Looks for existing BOM or weldment cut list on page
    Set swCurrentTableAnno = GetExistingTable(swModel)
    
    'If no BOM or cut list is found, attempts to insert and sort one on the page
    If swCurrentTableAnno Is Nothing Then
        Call InsertSortedTable
        Exit Sub
    End If
    
    'There is a more specific function to check if something is checked out in Vault,
    'but it is not available in VBA
    If swModel.IsOpenedReadOnly Then
        MsgBox "Drawing is not checked out.  Make sure to check out file to save changes."
    End If

    'If a BOM or cut list is found, sorts it
    Set sortConfig = ConfigureSortSettings(swCurrentTableAnno)
    response = SortTable(swCurrentTableAnno, sortConfig)

    If Not response Then
        MsgBox "Unable to sort table."
    End If
End Sub

Sub InsertSortedTable()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim swCurrentTableAnno As TableAnnotation
    Dim swView As View
    Dim sortConfig As SortSettings
    Dim insertConfig As InsertSettings
    Dim response As Boolean

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    'Checks drawing
    If Not IsValidDrawing(swModel) Then
        Exit Sub
    End If

    'There is a more specific function to check if something is checked out in Vault
    'But it is not available in VBA
    If swModel.IsOpenedReadOnly Then
        MsgBox "Drawing is not checked out.  Make sure to check out file to save changes."
    End If

    'Searches for a view to use for a BOM or cutlist
    Set swView = GetTableView(swModel)
    
    If swView Is Nothing Then
        MsgBox "No model view found on this page.", vbCritical
        Exit Sub
    End If
    
    'Gets location of default templates for fallback
    Set insertConfig = ConfigureInsertSettings(swView, swApp)
    Set swCurrentTableAnno = InsertTable(swView, insertConfig)
    
    If swCurrentTableAnno Is Nothing Then
        MsgBox "Unable to insert table.", vbCritical
        Exit Sub
    End If

    'Sorts inserted BOM or cutlist
    Set sortConfig = ConfigureSortSettings(swCurrentTableAnno)
    response = SortTable(swCurrentTableAnno, sortConfig)

    If Not response Then
        MsgBox "Unable to sort table."
    End If
End Sub

Public Function InsertTable(swView As View, insertConfig As InsertSettings) As TableAnnotation
    Dim swTableDoc As ModelDoc2

    Set swTableDoc = swView.ReferencedDocument

    Select Case swTableDoc.GetType
        Case swDocumentTypes_e.swDocASSEMBLY
            Set InsertTable = CreateAsmBom(swView, insertConfig)
        Case swDocumentTypes_e.swDocPART
            'Checks if part has a weldment and inserts BOM or cutlist accordingly
            If swTableDoc.IsWeldment Then
                Set InsertTable = CreateCutList(swView, insertConfig)
            Else
                'Creates a BOM but pops up a warning.
                MsgBox "Part is not a weldment. Check to see if it needs a BOM table."
               
                'No sort applied as it's a single item
                Set InsertTable = CreatePartBom(swView, insertConfig)
            End If
        Case Else
            Set InsertTable = Nothing
    End Select
End Function

Private Function SortTable(swCurrentTableAnno As TableAnnotation, sortConfig As SortSettings) As Boolean
    Dim swBomSortOptions As BomTableSortData
    
    Select Case swCurrentTableAnno.Type
        Case swTableAnnotation_BillOfMaterials
            Set swBomSortOptions = swCurrentTableAnno.GetBomTableSortData
            Call ConfigureBomSort(swBomSortOptions, sortConfig)
            
            'Applies custom sort as necessary
            If sortConfig.UseCustomSort Then
                SortTable = SortBomCustom(swCurrentTableAnno, sortConfig)
                Exit Function
            End If
            
            SortTable = SortBom(swCurrentTableAnno, sortConfig)
        Case swTableAnnotation_WeldmentCutList
            'Applies custom sort as necessary
            If sortConfig.UseCustomSort Then
                SortTable = SortCutListCustom(swCurrentTableAnno, sortConfig)
                Exit Function
            End If

            SortTable = SortCutList(swCurrentTableAnno, sortConfig)
        Case Else
            SortTable = False
    End Select
End Function