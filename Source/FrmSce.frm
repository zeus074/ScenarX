VERSION 5.00
Begin VB.Form FrmSce 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scenario"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAct 
      Caption         =   "Attiva"
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   3225
      Width           =   2490
   End
   Begin VB.ListBox Lsce 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2640
   End
End
Attribute VB_Name = "FrmSce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CAct_Click()
If Lsce.ListIndex < 0 Then Exit Sub
RET = Lsce.ListIndex + 1: Unload Me
End Sub

Private Sub Form_Activate()
Lsce.ListIndex = 0
End Sub

Private Sub Form_Load()
RET = "": CAct.Caption = txtActivate: Me.Caption = txtSce
End Sub
