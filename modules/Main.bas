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
    Dim swCurrentTableAnno As TableAnnotation
    Dim sortConfig As New SortSettings
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
    Call sortConfig.ConfigureSortSettings(swCurrentTableAnno)
    response = sortConfig.SortTable()

    If Not response Then
        MsgBox "Unable to sort table."
    End If
End Sub

Sub InsertSortedTable()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim swView As View
    Dim swCurrentTableAnno As TableAnnotation
    Dim insertConfig As New InsertSettings
    Dim sortConfig As New SortSettings
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

    'Searches for a view to use for a BOM or cut list
    Set swView = GetTableView(swModel)
    
    If swView Is Nothing Then
        MsgBox "No model view found on this page.", vbCritical
        Exit Sub
    End If
    
    'Gets location of default templates for fallback
    Call insertConfig.ConfigureInsertSettings(swView, swApp)
    Set swCurrentTableAnno = insertConfig.InsertTable()
    
    If swCurrentTableAnno Is Nothing Then
        MsgBox "Unable to insert table.", vbCritical
        Exit Sub
    End If

    'Sorts inserted BOM or cut list
    Call sortConfig.ConfigureSortSettings(swCurrentTableAnno)
    response = sortConfig.SortTable()

    If Not response Then
        MsgBox "Unable to sort table."
    End If
End Sub