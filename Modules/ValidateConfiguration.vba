'Checks user input for validity
Function ConfigureInsertSettings(swView As View, swApp As SldWorks.SldWorks) As InsertSettings
    Dim insertConfig As New InsertSettings
    Dim swTableDoc As ModelDoc2

    Set swTableDoc = swView.ReferencedDocument

    'If no valid template is provided, attempts to apply the default table provided with the SolidWorks install.
    'If this fails, SolidWorks will still insert a table, but it will be barebones.
    Select Case swTableDoc.GetType
        Case swDocumentTypes_e.swDocASSEMBLY
            insertConfig.AnchorType = bomAnchorPos

            If IsValidPath(bomTemplatePath) Then
                insertConfig.TemplatePath = bomTemplatePath
            Else
                insertConfig.TemplatePath = UpdateTemplatePath("bom", swApp)
            End If
        Case swDocumentTypes_e.swDocPART
            If swTableDoc.IsWeldment Then
                insertConfig.AnchorType = wclAnchorPos

                If IsValidPath(wclTemplatePath) Then
                    insertConfig.TemplatePath = wclTemplatePath
                Else
                    insertConfig.TemplatePath = UpdateTemplatePath("wcl", swApp)
                End If
            Else
                insertConfig.AnchorType = bomAnchorPos
                
                If IsValidPath(bomTemplatePath) Then
                    insertConfig.TemplatePath = bomTemplatePath
                Else
                    insertConfig.TemplatePath = UpdateTemplatePath("bom", swApp)
                End If
            End If
        Case Else
            insertConfig.AnchorType = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft
            insertConfig.TemplatePath = ""
    End Select

    Set ConfigureInsertSettings = insertConfig
End Function

Function ConfigureSortSettings(swTableAnno As TableAnnotation) As SortSettings
    Dim sortConfig As New SortSettings

    Select Case swTableAnno.Type
        Case swTableAnnotationType_e.swTableAnnotation_BillOfMaterials
            sortConfig.UseCustomSort = useCustomBomSort

            sortConfig.BomColumnsToSort = ValidateArray(swTableAnno)
        Case swTableAnnotationType_e.swTableAnnotation_WeldmentCutList
            sortConfig.UseCustomSort = useCustomWclSort

            If sortWclCol < 0 Or sortWclCol >= swTableAnno.ColumnCount Then
                MsgBox "User-provided sort column does not exist. Default will be set."

                sortConfig.CutListColumnToSort = 1
            End If

            sortConfig.CutListColumnToSort = sortWclCol
        Case Else
            MsgBox "Table type not recognized."
    End Select
    
    Set ConfigureSortSettings = sortConfig
End Function

Sub ConfigureBomSort(swBomSortOptions As BomTableSortData, sortConfig As SortSettings)
    swBomSortOptions.Ascending(0) = True
    swBomSortOptions.ColumnIndex(0) = sortConfig.BomColumnsToSort(0)
    swBomSortOptions.Ascending(1) = True
    swBomSortOptions.ColumnIndex(1) = sortConfig.BomColumnsToSort(1)
    swBomSortOptions.Ascending(2) = True
    swBomSortOptions.ColumnIndex(2) = sortConfig.BomColumnsToSort(2)
    swBomSortOptions.DoNotChangeItemNumber = False
    swBomSortOptions.ItemGroups = swBomTableSortItemGroup_None
    swBomSortOptions.SaveCurrentSortParameters = True
    swBomSortOptions.SortMethod = swBomTableSortMethod_Numeric
End Sub

Private Function IsValidPath(inputPath As String) As Boolean
    Dir ("") 'Resets search

    If inputPath = "" Or Dir(inputPath) = "" Then
        IsValidPath = False
        Exit Function
    End If
    
    IsValidPath = True
End Function

Private Function UpdateTemplatePath(tableType As String, swApp As SldWorks.SldWorks) As String
    Dim defaultLang As String
    Dim defaultDataFolder As String
    Dim defaultTemplateFolder As String
    
    defaultLang = swApp.GetCurrentLanguage()
    defaultDataFolder = swApp.GetDataFolder(True)

    defaultTemplateFolder = Left(defaultDataFolder, InStr(1, defaultDataFolder, "\data")) & _
                                defaultLang & "\"

    Select Case tableType
    Case "bom"
        MsgBox "User-provided BOM template not found. A basic template will be applied."

        UpdateTemplatePath = defaultTemplateFolder & "bom-standard.sldbomtbt"
    Case "wcl"
        MsgBox "User-provided cut list template not found. A basic template will be applied."

        UpdateTemplatePath = defaultTemplateFolder & "cut list.sldwldtbt"
    Case Else
        UpdateTemplatePath = ""
    End Select
End Function

Function ValidateArray(swTableAnno As TableAnnotation) As Variant
    Dim i As Long
    Dim sortCol As Integer
    Dim orderArr(0 To 2) As Variant
    Dim errArr(0 To 2) As Variant
    Dim msgArr(0 To 2) As String
    Dim useMessage As Boolean

    orderArr(0) = CInt(sortBomCol1)
    orderArr(1) = CInt(sortBomCol2)
    orderArr(2) = CInt(sortBomCol3)

    'Default array to deal with issues
    errArr(0) = 1
    errArr(1) = -1
    errArr(2) = -1

    useMessage = False

    For i = 0 To 2
        If Not IsValidColumn(orderArr(i), swTableAnno) Then
            useMessage = True
            orderArr(i) = errArr(i)
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

    ValidateArray = orderArr
End Function

Function IsValidColumn(sortCol As Variant, swTableAnno As TableAnnotation) As Boolean
    If sortCol < -1 Or sortCol >= swTableAnno.ColumnCount Then
        IsValidColumn = False
        Exit Function
    End If

    IsValidColumn = True
End Function