VERSION 5.00
Begin VB.Form FrmUpd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controllo aggiornamenti"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "FrmUpd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tchk 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   75
   End
   Begin VB.Image Img_agg 
      Height          =   480
      Index           =   2
      Left            =   75
      Picture         =   "FrmUpd.frx":038A
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Llink2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   675
      MouseIcon       =   "FrmUpd.frx":0983
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1275
      Width           =   3915
   End
   Begin VB.Label Llink 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   675
      MouseIcon       =   "FrmUpd.frx":0C8D
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   750
      Width           =   3915
   End
   Begin VB.Label Lchk 
      BackStyle       =   0  'Transparent
      Caption         =   "Controllo in corso ..."
      Height          =   390
      Left            =   675
      TabIndex        =   0
      Top             =   225
      Width           =   3465
   End
   Begin VB.Image Img_agg 
      Height          =   480
      Index           =   1
      Left            =   75
      Picture         =   "FrmUpd.frx":0F97
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img_agg 
      Height          =   480
      Index           =   0
      Left            =   75
      Picture         =   "FrmUpd.frx":14C5
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "FrmUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute& Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Sub Form_Load()
Lchk = txtUpdChk
Tchk.Enabled = True
Llink2 = txtUpdlink
End Sub

Private Function ChkUpdate()
On Local Error GoTo ErrUpd
Dim inet As Object, b() As Byte, Versione As String
Const icByteArray = 1
ChkUpdate = ""

Set inet = CreateObject("InetCtls.Inet")
b() = inet.OpenURL("http://www.medioformato.it/my-btproject/scenarx/sc_ver.txt", icByteArray)
Versione = StrConv(b, vbUnicode)
If Left(Versione, 1) <> "<" Then
Dim Tp As Variant
Tp = Split(Versione, ",")
ChkUpdate = Tp(0)
nVER = Tp(0): nLINK = Tp(1)
Else
nVER = "ERR": ChkUpdate = "ERR"
End If
Set inet = Nothing
Exit Function
ErrUpd:
nVER = "ERR": ChkUpdate = "ERR"
Resume OK
OK:
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Llink.FontUnderline = True Then Llink.FontUnderline = False
If Llink2.FontUnderline = True Then Llink2.FontUnderline = False
End Sub

Private Sub Llink_Click()
ShellExecute FrmUpd.hWnd, vbNullString, nLINK, vbNullString, vbNullString, 1
End Sub

Private Sub Llink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Llink.FontUnderline = False Then Llink.FontUnderline = True
End Sub

Private Sub Llink2_Click()
ShellExecute FrmUpd.hWnd, vbNullString, "www.medioformato.it/my-btproject/scenarx/", vbNullString, vbNullString, 1
End Sub

Private Sub Llink2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Llink2.FontUnderline = False Then Llink2.FontUnderline = True
End Sub

Private Sub Tchk_Timer()
Tchk.Enabled = False
Img_agg(0).Visible = False: Img_agg(1).Visible = True
If ChkUpdate <> "" Then
If nVER = "ERR" Then GoTo ErrUpd
Dim S1, S2 As Integer, TMP As Variant

TMP = Replace(VER, ".", "")
S1 = Val(TMP)
TMP = Replace(nVER, ".", "")
S2 = Val(TMP)
    If S1 < S2 Then
    Lchk = txtUpdReady & " : " & nVER & vbCrLf & _
    txtUpdAct & " " & VER & " --> " & txtUpdNew & " " & nVER
    Llink = txtUpdDwn & " " & nLINK
    DoEvents
    Else
    Lchk = txtUpdNo  ' & " " & nVER
    Llink = "": Llink.Enabled = False
    End If
End If
Img_agg(1).Visible = False: Img_agg(0).Visible = True

Exit Sub
ErrUpd:
Img_agg(1).Visible = False: Img_agg(2).Visible = True
Lchk = txtUpdErr
End Sub
