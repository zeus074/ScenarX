VERSION 5.00
Begin VB.Form FrmLanguage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selezione la lingua"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3675
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton C_ok 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2475
      TabIndex        =   1
      Top             =   225
      Width           =   1215
   End
   Begin VB.ComboBox Cmb_Lang 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   2040
   End
End
Attribute VB_Name = "FrmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldLng As String

Private Sub C_ok_Click()
Language = LCase(Cmb_Lang.Text)
Unload Me
End Sub

Private Sub Cmb_Lang_Click()
If OldLng = LCase(Cmb_Lang.Text) Then Exit Sub
OldLng = Language: C_ok.Enabled = True
End Sub

Private Sub Cmb_Lang_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim Nm As String
File1.Path = App.Path & "\Languages\"
File1.Pattern = "*.xml"
If File1.ListCount = 0 Then Exit Sub
For X = 1 To File1.ListCount
Nm = File1.List(X - 1): Nm = Left(Nm, Len(Nm) - 4)
Cmb_Lang.AddItem Nm
If LCase(Nm) = Language Then OldLng = Language: Cmb_Lang.ListIndex = X - 1
Next X
End Sub
