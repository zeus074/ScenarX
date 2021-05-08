Attribute VB_Name = "ModGen"
Option Explicit
Global ll As Long

Global glConnectOK As Boolean ' flag di connessione TCP / TCP connection flag
Global glMainTop As Single
Global glMainLeft As Single
Global glBusy As Boolean ' flag di busy applicativo / busy software flag
Global BufferInTCP As String
Global glSsck As Boolean
Global ErLoad  As Boolean
Global PrgAct  As Boolean
Global InProgram  As Boolean
Global NoSave As Boolean

' costanti applicativo / software constants
Global Const glcACKTCP = "*#*1##"
Global Const glcNACKTCP = "*#*0##"

Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Lindex As Integer
Public DelayS, DelayT, DelayG, DelayC, DelayM As Byte
Public txtNoname, txtNew, Language, Ret As String
Public VER, nVER, nLINK As String
Public txtFrmOpen, txtCmd(6), txtOpt, txtInsert, txtActivate, txtNoOpen As String
Public txtDelete, txtProgram, txtManProg, txtDirty, txtDirtyCap, txtNoCon As String
Public txtStatDis, txtStatCon, txtSck, txtSsck, txtWaitCon, txtNoAck As String
Public txtErrCom, txtAllCom, txtErrEnd, txtOkEnd, txtErrProg, txtInProg As String
Public txtClearAll, txtClear, txtDelCap, txtClear2, TxtOpen, txtNf, txtNf2 As String
Public txtSend, txtFullSce, txtSave, txtGen, txtGr, txtAmb, txtPnt, txtSource As String
Public txtSpeed, txtSpeed2, txtLevel, txtTemper, txtProg, txtSce, txtVol, txtCom As String
Public txtBusL, txtBusP, txtFrmLng, txtDeleted, txtProgrammed, txtStart, txtStop As String
Public txtUpdReady, txtUpdAct, txtUpdNew, txtUpdDwn, txtUpdNo, txtUpdErr, txtUpdlink As String
Public txtUpdChk, txtProgram2, txtSelAll, txtClearAll2, txtDelAll, txtDel, txtChk As String
Public txtWack, txtNtries, txtVirtErr, txtVirtErr2, txtVirtErr3, txtVirtErr4, txtDuplicate, txtDuplicate2 As String
Public LoadVer, txtOldVer, txtFileLoaded As String
Public TimWait, TimRetry As Double
Public gobjLanguage As ClsLanguage

Public Sub SetLanguage(ByVal Plng As String)
Dim objForm         As iLanguage
Dim frmForm         As Form
    If gobjLanguage Is Nothing Then
        Set gobjLanguage = New ClsLanguage
    End If
    gobjLanguage.LoadLanguageFile Plng
    For Each frmForm In Forms
        If TypeOf frmForm Is iLanguage Then
            Set objForm = frmForm
            objForm.Updated
            Set objForm = Nothing
        End If
    Next
SaveSetting App.EXEName, "Option", "Language", Plng: Language = Plng
End Sub

Public Sub Log(testo As String)
If frmMain.Clog.Value <> 1 Then Exit Sub

If Dir$(App.Path & "\log.txt") = "" Then
Open App.Path & "\log.txt" For Output As #1
Print #1, testo
Close #1
Else
Open App.Path & "\log.txt" For Append As #1
Print #1, testo
Close #1
End If
End Sub
