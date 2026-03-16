Option Explicit

Function CreateAsmBom(swView As View, insertConfig As InsertSettings) As BomTableAnnotation
    Dim config As String

    config = swView.ReferencedConfiguration
    Set CreateAsmBom = swView.InsertBomTable4( _
                        True, _
                        0, _
                        0, _
                        insertConfig.AnchorType, _
                        swBomType_TopLevelOnly, _
                        config, _
                        insertConfig.TemplatePath, _
                        False, _
                        swIndentedBOMNotSet, _
                        False)
End Function

Function CreatePartBom(swView As View, insertConfig As InsertSettings) As BomTableAnnotation
    Dim config As String

    config = swView.ReferencedConfiguration
    Set CreatePartBom = swView.InsertBomTable4( _
                        True, _
                        0, _
                        0, _
                        insertConfig.AnchorType, _
                        swBomType_PartsOnly, _
                        config, _
                        insertConfig.TemplatePath, _
                        False, _
                        swIndentedBOMNotSet, _
                        False)
End Function

Function SortBom(swBomAnno As BomTableAnnotation, sortConfig As SortSettings) As Boolean
    Dim swBomSortOptions As BomTableSortData

    Set swBomSortOptions = swBomAnno.GetBomTableSortData()
    Call sortConfig.ConfigureBomSort(swBomSortOptions)
    SortBom = swBomAnno.Sort(swBomSortOptions)
End Function

Function SortBomCustom(swBomTableAnno As TableAnnotation, sortConfig As SortSettings) As Boolean
    Dim swSortData As BomTableSortData
    Dim colPartNoIndex As Integer
    Dim colSortIndex As Integer
    Dim rowCount As Integer
    Dim isSorted As Boolean
    Dim isDeleted As Boolean

    'BOM Order
    'Priority 1: Manufactured part numbers - begins with a letter
    'Priority 2: Stock purchased part numbers - begins with a number

    colPartNoIndex = sortBomCol1
    colSortIndex = swBomTableAnno.ColumnCount

    'Adds a sort column to the BOM
    'Against standard VBA convention, 0 is false, 1 is true here
    'So we will run type coercion on this function
    isSorted = (swBomTableAnno.InsertColumn2( _
                   swTableItemInsertPosition_After, _
                   colSortIndex - 1, _
                   "Priority", _
                   swInsertColumn_SingleLineTight) <> 0)
    
    If Not isSorted Then
        MsgBox "An error occured with inserting sort column.", vbCritical
        SortBomCustom = False
        Exit Function
    End If

    'Assign a priority value to every row based on its part number
    rowCount = swBomTableAnno.rowCount - 1

    Dim i As Integer
    For i = 1 To rowCount
        Dim partNo As String

        partNo = swBomTableAnno.Text2(i, colPartNoIndex, False)

        'Checks if the part is a stock part number
        If IsNumeric(Left(partNo, 1)) Then
            swBomTableAnno.Text2(i, colSortIndex, False) = 2
        Else
            swBomTableAnno.Text2(i, colSortIndex, False) = 1
        End If
    Next i

    'Overrides default sort configuration
    'Sort BOM by priority value, Then part number
    Set swSortData = swBomTableAnno.GetBomTableSortData()

    swSortData.Ascending(0) = True
    swSortData.ColumnIndex(0) = colSortIndex
    swSortData.Ascending(1) = True
    swSortData.ColumnIndex(1) = colPartNoIndex
    swSortData.Ascending(2) = True
    swSortData.ColumnIndex(2) = -1  'Unused sort option
    swSortData.DoNotChangeItemNumber = False
    swSortData.ItemGroups = swBomTableSortItemGroup_None
    swSortData.SaveCurrentSortParameters = True
    swSortData.SortMethod = swBomTableSortMethod_Numeric

    SortBomCustom = swBomTableAnno.Sort(swSortData)

    'Delete sort column
    isDeleted = (swBomTableAnno.DeleteColumn2(colSortIndex, False) <> 0)
    
    If Not isDeleted Then
        MsgBox "Sort column was unable to be deleted. Please remove manually."
    End If
End Function