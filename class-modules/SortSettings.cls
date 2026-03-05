Option Explicit

'------------
' Properties
'------------
Private pCustomSortStatus As Boolean
Private pBomSortOrder(0 To 2) As Long
Private pWclSortOrder As Integer
Private pTableAnnotation As TableAnnotation

'----------------
' Initialization
'----------------
Private Sub Class_Initialize()
    pBomSortOrder(0) = 1    'Unable to use SetBomOrder here
    pBomSortOrder(1) = -1
    pBomSortOrder(2) = -1
    Me.WclColumnToSort = 2
End Sub

Public Sub ConfigureSortSettings(swTableAnno As TableAnnotation)
    Set Me.TableAnnotation = swTableAnno

    Select Case swTableAnno.Type
        Case swTableAnnotationType_e.swTableAnnotation_BillOfMaterials
            Me.UseCustomSort = useCustomBomSort

            Call SetBomOrder(swTableAnno)
        Case swTableAnnotationType_e.swTableAnnotation_WeldmentCutList
            Me.UseCustomSort = useCustomWclSort

            If sortWclCol < 0 Or sortWclCol >= swTableAnno.ColumnCount Then
                MsgBox "User-provided sort column does not exist. Default will be set."

                Me.WclColumnToSort = 1
            End If

            Me.WclColumnToSort = sortWclCol
        Case Else
            MsgBox "Table type not recognized."
    End Select
End Sub

'-----------------
' Get/Set Methods
'-----------------
Public Property Get UseCustomSort() As Boolean
    UseCustomSort = pCustomSortStatus
End Property

Public Property Let UseCustomSort(ByVal newStatus As Boolean)
    pCustomSortStatus = newStatus
End Property

Public Property Get TableAnnotation() As TableAnnotation
    Set TableAnnotation = pTableAnnotation
End Property

Public Property Set TableAnnotation(ByVal newTableAnnotation As TableAnnotation)
    Set pTableAnnotation = newTableAnnotation
End Property

Public Property Get BomColumnsToSort() As Variant
    BomColumnsToSort = pBomSortOrder
End Property

'Since the array needs go through extensive validation, this is not a normal set
Public Sub SetBomOrder(swTableAnno As TableAnnotation)
    Dim i As Long
    Dim sortCol As Integer
    Dim orderArr(0 To 2) As Variant
    Dim defaultArr(0 To 2) As Variant
    Dim msgArr(0 To 2) As String
    Dim useMessage As Boolean

    orderArr(0) = CInt(sortBomCol1)
    orderArr(1) = CInt(sortBomCol2)
    orderArr(2) = CInt(sortBomCol3)

    'Default array to deal with issues
    defaultArr(0) = 1
    defaultArr(1) = -1
    defaultArr(2) = -1

    useMessage = False

    For i = 0 To 2
        If Not IsValidColumn(orderArr(i)) Then
            useMessage = True
            orderArr(i) = defaultArr(i)
        End If

        If orderArr(i) <> -1 Then
            msgArr(i) = CStr(orderArr(i))
        Else
            msgArr(i) = "unused"
        End If
    Next i
    
    If useMessage Then
        MsgBox _
            "There was an error in sort configuration. Please check settings applied" & vbCrLf & _
            "First sort: column " & msgArr(0) & vbCrLf & _
            "Second sort: column " & msgArr(1) & vbCrLf & _
            "Third sort: column " & msgArr(2), _
            vbExclamation
    End If
End Sub

Public Property Get WclColumnToSort() As Integer
    WclColumnToSort = pWclSortOrder
End Property

Public Property Let WclColumnToSort(ByVal newOrder As Integer)
    pWclSortOrder = newOrder
End Property

'-------------------
' Utility Functions
'-------------------
Public Sub ConfigureBomSort(swBomSortOptions As BomTableSortData)
    swBomSortOptions.Ascending(0) = True
    swBomSortOptions.ColumnIndex(0) = Me.BomColumnsToSort(0)
    swBomSortOptions.Ascending(1) = True
    swBomSortOptions.ColumnIndex(1) = Me.BomColumnsToSort(1)
    swBomSortOptions.Ascending(2) = True
    swBomSortOptions.ColumnIndex(2) = Me.BomColumnsToSort(2)
    swBomSortOptions.DoNotChangeItemNumber = False
    swBomSortOptions.ItemGroups = swBomTableSortItemGroup_None
    swBomSortOptions.SaveCurrentSortParameters = True
    swBomSortOptions.SortMethod = swBomTableSortMethod_Numeric
End Sub

Private Function IsValidColumn(sortCol As Variant) As Boolean
    If sortCol < -1 Or sortCol >= Me.TableAnnotation.ColumnCount Then
        IsValidColumn = False
        Exit Function
    End If

    IsValidColumn = True
End Function

Public Function SortTable() As Boolean
    Dim swBomSortOptions As BomTableSortData
    
    Select Case Me.TableAnnotation.Type
        Case swTableAnnotation_BillOfMaterials
            Set swBomSortOptions = Me.TableAnnotation.GetBomTableSortData
            Call ConfigureBomSort(swBomSortOptions)
            
            'Applies custom sort as necessary
            If Me.UseCustomSort Then
                SortTable = SortBomCustom(Me.TableAnnotation, Me)
                Exit Function
            End If
            
            SortTable = SortBom(Me.TableAnnotation, Me)
        Case swTableAnnotation_WeldmentCutList
            'Applies custom sort as necessary
            If Me.UseCustomSort Then
                SortTable = SortCutListCustom(Me.TableAnnotation, Me)
                Exit Function
            End If

            SortTable = SortCutList(Me.TableAnnotation, Me)
        Case Else
            SortTable = False
    End Select
End Function