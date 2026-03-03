Option Explicit

Private pCustomSortStatus As Boolean
Private pBomSortOrder(2) As Long
Private pWclSortOrder As Integer

Public Property Get UseCustomSort() As Boolean
    UseCustomSort = pCustomSortStatus
End Property

Public Property Let UseCustomSort(ByVal newStatus As Boolean)
    pCustomSortStatus = newStatus
End Property

' -------------------------
'  BOM-specific Properties
' -------------------------
Public Property Get BomColumnsToSort() As Variant
    BomColumnsToSort = pBomSortOrder
End Property

Public Property Let BomColumnsToSort(ByVal newOrder As Variant)
    Dim i As Long

    For i = 0 To 2
        pBomSortOrder(i) = CLng(newOrder(i))
    Next i
End Property

' ------------------------------
'  Cut List-specific Properties
' ------------------------------
Public Property Get CutListColumnToSort() As Integer
    CutListColumnToSort = pWclSortOrder
End Property

Public Property Let CutListColumnToSort(ByVal newOrder As Integer)
    pWclSortOrder = newOrder
End Property