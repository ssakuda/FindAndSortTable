Option Explicit

'------------
' Properties
'------------
Private pTableView As View
Private pTemplatePath As String
Private pAnchorType As Integer

'----------------
' Initialization
'----------------
Private Sub Class_Initialize()
    'Default values
    Me.AnchorType = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft
    pTemplatePath = ""    'Unable to use SetTemplatePath here
End Sub

Public Sub ConfigureInsertSettings(swView As View, swApp As SldWorks.SldWorks)
    Dim swTableDoc As ModelDoc2
    Dim modifiedPath As String

    Set Me.TableView = swView
    Set swTableDoc = swView.ReferencedDocument

    'If no valid template is provided, attempts to apply the default table provided with the SolidWorks install.
    'If this fails, SolidWorks will still insert a table, but it will be barebones.
    Select Case swTableDoc.GetType
        Case swDocumentTypes_e.swDocASSEMBLY
            Me.AnchorType = bomAnchorPos
            Call Me.SetTemplatePath( _
                    bomTemplatePath, _
                    swTableAnnotationType_e.swTableAnnotation_BillOfMaterials, _
                    swApp)
        Case swDocumentTypes_e.swDocPART
            If swTableDoc.IsWeldment Then
                Me.AnchorType = wclAnchorPos
                Call Me.SetTemplatePath( _
                    wclTemplatePath, _
                    swTableAnnotationType_e.swTableAnnotation_WeldmentCutList, _
                    swApp)
                Exit Sub
            End If
                
                Me.AnchorType = bomAnchorPos
                Call Me.SetTemplatePath( _
                        bomTemplatePath, _
                        swTableAnnotationType_e.swTableAnnotation_BillOfMaterials, _
                        swApp)
        Case Else
            Me.AnchorType = swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopLeft
            Call Me.SetTemplatePath( _
                    "", _
                    -1, _
                    swApp)
    End Select
End Sub

'-----------------
' Get/Set Methods
'-----------------
Public Property Get TableView() As View
    Set TableView = pTableView
End Property

Public Property Set TableView(ByVal newTableView As View)
    Set pTableView = newTableView
End Property

Public Property Get TemplatePath() As String
    TemplatePath = pTemplatePath
End Property

'Since the path should go through IsValidPath for validation, this is not a normal set
Public Sub SetTemplatePath(ByVal newPath As String, tableType As swTableAnnotationType_e, swApp As SldWorks.SldWorks)
    If Not IsValidPath(newPath) Then
        pTemplatePath = GetDefaultTemplatePath(tableType, swApp)
        Exit Sub
    End If

    pTemplatePath = newPath
End Sub

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

'-------------------
' Utility Functions
'-------------------
Private Function IsValidPath(inputPath As String) As Boolean
    Dir ("") 'Resets search

    If inputPath = "" Or Dir(inputPath) = "" Then
        IsValidPath = False
        Exit Function
    End If
    
    IsValidPath = True
End Function

Private Function GetDefaultTemplatePath(tableType As swTableAnnotationType_e, swApp As SldWorks.SldWorks) As String
    Dim defaultLang As String
    Dim defaultDataFolder As String
    Dim defaultTemplateFolder As String
    
    defaultLang = swApp.GetCurrentLanguage()
    defaultDataFolder = swApp.GetDataFolder(True)
    defaultTemplateFolder = Left(defaultDataFolder, InStr(1, defaultDataFolder, "\data")) & _
                                defaultLang & "\"

    Select Case tableType
        Case swTableAnnotationType_e.swTableAnnotation_BillOfMaterials
            MsgBox "User-provided BOM template not found. A basic template will be applied."
    
            UpdateTemplatePath = defaultTemplateFolder & "bom-standard.sldbomtbt"
        Case swTableAnnotationType_e.swTableAnnotation_WeldmentCutList
            MsgBox "User-provided cut list template not found. A basic template will be applied."
    
            UpdateTemplatePath = defaultTemplateFolder & "cut list.sldwldtbt"
        Case Else
            UpdateTemplatePath = ""
        End Select
End Function

Public Function InsertTable() As TableAnnotation
    Dim swTableDoc As ModelDoc2

    Set swTableDoc = Me.TableView.ReferencedDocument
    Select Case swTableDoc.GetType
        Case swDocumentTypes_e.swDocASSEMBLY
            Set InsertTable = CreateAsmBom(Me.TableView, Me)
        Case swDocumentTypes_e.swDocPART
            'Checks if part has a weldment and inserts BOM or cutlist accordingly
            If swTableDoc.IsWeldment Then
                Set InsertTable = CreateCutList(Me.TableView, Me)
            Else
                'Creates a BOM but pops up a warning.
                MsgBox "Part is not a weldment. Check to see if it needs a BOM table."
               
                'No sort applied as it's a single item
                Set InsertTable = CreatePartBom(Me.TableView, Me)
            End If
        Case Else
            Set InsertTable = Nothing
    End Select
End Function