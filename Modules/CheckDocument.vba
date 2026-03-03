Option Explicit

Public Function IsValidDrawing(swModel As ModelDoc2) As Boolean
    'Checks for an open document
    If swModel Is Nothing Then
        MsgBox "No active document found.", vbExclamation
        IsValidDrawing = False
     Exit Function
    End If
    
    'Checks open document type
    If swModel.GetType <> swDocumentTypes_e.swDocDRAWING Then
        MsgBox "This macro is only For use with drawings.", vbCritical
        IsValidDrawing = False
     Exit Function
    End If

    IsValidDrawing = True
End Function

'Looks for a BOM or weldment cutlist on the current sheet
'Function assumes that there is only one per page
Public Function GetExistingTable(swDraw As DrawingDoc) As TableAnnotation
    Dim swView As View
    Dim hasBom As Boolean
    Dim currentTable As TableAnnotation
    Dim isBom As Boolean
    
    isBom = False
    
    Set swView = swDraw.GetFirstView() 'First view is the sheet itself
    Set currentTable = swView.GetFirstTableAnnotation()
    
    Do While Not currentTable Is Nothing And isBom
        If (currentTable.Type = swTableAnnotation_BillOfMaterials Or _
            currentTable.Type = swTableAnnotation_WeldmentCutList) Then
            isBom = True
        Else
            currentTable = currentTable.GetNext
        End If
    Loop

    Set GetExistingTable = currentTable
End Function

Public Function GetTableView(swDraw As DrawingDoc) As View
    Dim swSheet As Sheet
    Dim swView As View

    'Looks for an active view on the current sheet
    'Otherwise it will pick the view inserted first
    Set swSheet = swDraw.GetCurrentSheet()
    Set swView = swDraw.ActiveDrawingView
    
    If swView Is Nothing Then
        Set swView = swDraw.GetFirstView().GetNextView() 'First view is the sheet itself
    End If
    
    'It is possible to find nothing
    Set GetTableView = swView
End Function