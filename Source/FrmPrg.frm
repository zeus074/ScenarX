VERSION 5.00
Begin VB.Form FrmPrg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scenario"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox CselAll 
      Caption         =   "Seleziona tutti"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   30
      TabIndex        =   2
      Top             =   75
      Width           =   2490
   End
   Begin VB.CommandButton CAct 
      Caption         =   "Programma"
      Enabled         =   0   'False
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   4050
      Width           =   2490
   End
   Begin VB.ListBox Lsce 
      Appearance      =   0  'Flat
      Height          =   3630
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   300
      Width           =   2640
   End
End
Attribute VB_Name = "FrmPrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoMod As Boolean

Private Sub CAct_Click()
Dim Act As Boolean
Ret = ""
For X = 0 To 15
If Lsce.Selected(X) = True Then
Act = True
If Ret = "" Then Ret = Ret & X + 1 Else Ret = Ret & "," & X + 1
End If
Next X

If Act Then Unload Me Else Ret = ""
End Sub

Private Sub CselAll_Click()
If CselAll.Value <> 1 Then Exit Sub
NoMod = True: For X = 0 To 15: Lsce.Selected(X) = True: Next X: NoMod = False
CAct.Enabled = True
End Sub

Private Sub Form_Activate()
'Lsce.ListIndex = 0
End Sub

Private Sub Form_Load()
Ret = "": CAct.Caption = txtProgram2: Me.Caption = txtSce
CselAll.Caption = txtSelAll
End Sub

Private Sub Lsce_Click()
If NoMod Then Exit Sub
CAct.Enabled = False
Dim Act As Boolean
For X = 0 To 15
If Lsce.Selected(X) <> True Then CselAll.Value = 0 Else Act = True
Next X
If Act Then CAct.Enabled = True
End Sub
