Private pTemplatePath As String
Private pAnchorType As Integer

Public Property Get TemplatePath() As String
    TemplatePath = pTemplatePath
End Property

Public Property Let TemplatePath(ByVal newPath As String)
    pTemplatePath = newPath
End Property

Public Property Get AnchorType() As Integer
    AnchorType = pAnchorType
End Property

Public Property Let AnchorType(ByVal newAnchor As Integer)
    If newAnchor > 0 And newAnchor <= 4 Then
        pAnchorType = newAnchor
    Else
        MsgBox "Selected anchor point is invalid. Will be set to top-left."
        
        pAnchorType = 1
    End If
End Property
