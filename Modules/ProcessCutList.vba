Function CreateCutList(swView As View, insertConfig As InsertSettings) As WeldmentCutListAnnotation
    Dim config As String

    config = GetWeldConfig(swView.ReferencedConfiguration)
    Set CreateCutList = swView.InsertWeldmentTable( _
                        True, _
                        0, _
                        0, _
                        insertConfig.AnchorType, _
                        config, _
                        insertConfig.TemplatePath)
End Function

Function SortCutList(swWclAnnotation As WeldmentCutListAnnotation, sortConfig As SortSettings) As Boolean
    SortCutList = swWclAnnotation.Sort(sortConfig.CutListColumnToSort, True)
End Function

Private Function GetWeldConfig(currentConfig As String) As String
    Dim pos As Integer
    
    
    'Config name could have <As Machined> and should be changed to <As Welded>
    pos = InStr(currentConfig, "<")
    If InStr(currentConfig, "<As Welded>") > 0 Then
        GetWeldConfig = currentConfig
        Exit Function
    ElseIf pos = 0 Then
        MsgBox "Unknown error occurred.", vbCritical
        Exit Function
    End If
    
    GetWeldConfig = Left(currentConfig, pos - 1) & "<As Welded>"
End Function

Function SortCutListCustom(swBomTableAnno As TableAnnotation, sortConfig As SortSettings) As Boolean
    SortWclCustom = False

    MsgBox "No sort rules have been added."
End Function