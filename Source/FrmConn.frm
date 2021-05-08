VERSION 5.00
Begin VB.Form FrmConn 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comunicazione"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   Icon            =   "FrmConn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Cok 
      Caption         =   "OK"
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      Top             =   825
      Width           =   2865
   End
   Begin VB.TextBox Tretry 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   450
      Width           =   315
   End
   Begin VB.TextBox Twait 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1050
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "1"
      Top             =   150
      Width           =   315
   End
   Begin VB.Image ImgMap 
      Height          =   480
      Left            =   75
      Picture         =   "FrmConn.frx":617A
      Top             =   75
      Width           =   480
   End
   Begin VB.Label LbTries 
      BackStyle       =   0  'Transparent
      Caption         =   "numero di tentativi controllo ack (sec.)"
      Height          =   240
      Left            =   1500
      TabIndex        =   1
      Top             =   450
      Width           =   2790
   End
   Begin VB.Label Lbset 
      BackStyle       =   0  'Transparent
      Caption         =   "Attesa ack/nack (sec.)"
      Height          =   240
      Left            =   1500
      TabIndex        =   0
      Top             =   150
      Width           =   2865
   End
End
Attribute VB_Name = "FrmConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cok_Click()
TimWait = Twait: TimRetry = Tretry
SaveSetting App.EXEName, "Option", "Wait", TimWait
SaveSetting App.EXEName, "Option", "Retry", TimRetry
Unload Me
End Sub

Private Sub Form_Load()
Twait = TimWait
Tretry = TimRetry
Lbset = txtWack
LbTries = txtNtries
End Sub

Private Sub Tretry_Change()
If Tretry > 3 Then Tretry = 3
If Tretry < 1 Then Tretry = 1
End Sub

Private Sub Tretry_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub Twait_Change()
If Twait > 20 Then Twait = 20
If Twait < 1 Then Twait = 1
End Sub

Private Sub Twait_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
        KeyAscii = 0
End Select
End Sub
