VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjDOM As DOMDocument

Private Sub Class_Terminate()
    If Not (mobjDOM Is Nothing) Then
        Set mobjDOM = Nothing
    End If
End Sub

Public Sub LoadLanguageFile(ByVal LanguageID As String)
On Error GoTo ErrLoad
    Set mobjDOM = New DOMDocument
    If Not mobjDOM.Load(App.Path & "\Languages\" & LanguageID & ".xml") Then
        Set mobjDOM = Nothing
        If Not ErLoad Then ErLoad = True: GoTo ErrLoad
    End If
Exit Sub
ErrLoad:
MsgBox "Language file " & LanguageID & " not loaded." & vbCrLf & _
Err.Description, vbCritical, "Error loading xml file"
    
End Sub

Public Function GetLabel(ByVal Section As String, ByVal Tag As String) As String
On Error GoTo ErrLoad
    If Not (mobjDOM Is Nothing) Then
        GetLabel = mobjDOM.selectSingleNode("Labels//" & Section & "//" & Tag).Text
    Else
        If Not ErLoad Then ErLoad = True: GoTo ErrLoad
    End If
Exit Function
ErrLoad:
MsgBox "Language file not loaded." & vbCrLf & "Key: Labels//" & Section & "//" & Tag, vbCritical, "Error loading xml file"
End Function

Private Sub Class_Initialize()
'
End Sub

