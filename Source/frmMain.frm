VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   0
   ClientWidth     =   3465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pctPlacca 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   0
      Top             =   0
      Width           =   3465
      Begin VB.PictureBox pctEsci 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         Picture         =   "frmMain.frx":1B342
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   135
         TabIndex        =   4
         ToolTipText     =   "Chiude la finestra"
         Top             =   1845
         Width           =   2025
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   2850
         Top             =   2820
      End
      Begin VB.PictureBox pctTasto 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Index           =   5
         Left            =   2085
         Picture         =   "frmMain.frx":1C83C
         ScaleHeight     =   1350
         ScaleWidth      =   675
         TabIndex        =   3
         Top             =   510
         Width           =   675
      End
      Begin VB.PictureBox pctTasto 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Index           =   3
         Left            =   1395
         Picture         =   "frmMain.frx":1CC92
         ScaleHeight     =   1350
         ScaleWidth      =   675
         TabIndex        =   2
         Top             =   510
         Width           =   675
      End
      Begin VB.PictureBox pctTasto 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Index           =   1
         Left            =   705
         Picture         =   "frmMain.frx":1D0E8
         ScaleHeight     =   1350
         ScaleWidth      =   675
         TabIndex        =   1
         Top             =   510
         Width           =   675
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   720
         Top             =   450
         Width           =   2040
      End
      Begin VB.Label lblWait 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2370
         TabIndex        =   5
         Top             =   2880
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim lcTO As Long ' contatore timeout visualizzazione placca / timeout counter to display the platen
Dim ButtonPress As Long ' indicatore pulsante premuto / pressed button indicator
Dim OffSetX As Long ' offset X
Dim OffSetY As Long ' offset Y
'Purpose:clip immagine placca / platen image clip
Public Sub ClipPlacca()
Dim x As Single
Dim Y As Single
Dim RetVal As Long
Dim PosIni As Long
Dim PosEnd As Long
Dim Rgn_Source As Long
Dim Rgn_Clip As Long
Dim Rgn_Dest As Long
Dim FlagFirstloop As Boolean
Dim k As Long
Dim Red As Long
Dim Green As Long
Dim Blue As Long
Dim MaxClip As Long
Dim MinClip As Long

MaxClip = 200
MinClip = 185

Const bck = &HFF0000
On Error Resume Next
For Y = 0 To pctPlacca.ScaleHeight
    PosIni = 0
    PosEnd = 0
    For x = 0 To pctPlacca.ScaleWidth + 1
        RetVal = pctPlacca.Point(x, Y)
        Red = (RetVal And &HFF0000) / &H10000
        Green = (RetVal And &HFF00&) / &H100&
        Blue = (RetVal And &HFF&) / &H1&
        If (Red >= MinClip And Red <= MaxClip) And (Green >= MinClip And Green <= MaxClip) And (Blue >= MinClip And Blue <= MaxClip) Then
            RetVal = bck
        End If
        If RetVal = bck And PosIni = 0 Then
            PosIni = x
        ElseIf RetVal <> bck And (PosIni <> 0 Or x = 1) Then
            PosEnd = x
            Call BuildPolygon(PosIni - 1, PosEnd + 1, Y, k)
            Dim YU As Long
            Dim yD As Long
            YU = Y - 1
            yD = Y
            PosIni = PosIni - 1
            CountPoly = CountPoly + 1
            PosEnd = 0
            PosIni = 0
            k = k + 4
        End If
    Next x
    If PosIni <> 0 And PosEnd = 0 Then
        PosEnd = pctPlacca.ScaleWidth
    End If
    If PosIni <> 0 Then
        
    
    End If
Next Y
ReDim NumPoly(CountPoly - 1)

For k = 0 To CountPoly - 1
    NumPoly(k) = 4
Next k
    
Rgn_TOT = CreatePolyPolygonRgn(PloyPoint(0), NumPoly(0), CountPoly, WINDING)
Call ClipPicture
End Sub
'Purpose:routine di ausilio clip immagine placca / routine in aid of the platen image clip
Public Sub BuildPolygon(PosIni As Long, PosEnd As Long, Y As Single, k As Long)

ReDim Preserve PloyPoint(k + 3)
PloyPoint(k).x = PosIni
PloyPoint(k).Y = Y - 1
PloyPoint(k + 1).x = PosEnd
PloyPoint(k + 1).Y = Y - 1
PloyPoint(k + 2).x = PosEnd
PloyPoint(k + 2).Y = Y + 1
PloyPoint(k + 3).x = PosIni
PloyPoint(k + 3).Y = Y + 1
End Sub
'Purpose:inizializzazione form / form initialization
Private Sub Form_Load()
Dim k As Long
On Error Resume Next

Call LoadConfig(glVirtualSw)
Call ClipPlacca
If glVirtualSw.Placca <> "" Then
    Me.pctPlacca.Picture = LoadPicture(App.Path & "\system\" & glVirtualSw.Placca)
End If

DoEvents
Me.Hide
End Sub
'Purpose:salvataggio dati posizione / saving position records
Private Sub Form_Unload(Cancel As Integer)
Dim RetVal As Long
Dim Buffer As String

Buffer = CStr(frmMain.Top)
glMainTop = frmMain.Top
RetVal = WritePrivateProfileString("GENERAL", "TOP", Buffer, App.Path & "\" & glFileCFG)

Buffer = CStr(frmMain.Left)
glMainLeft = frmMain.Left
RetVal = WritePrivateProfileString("GENERAL", "LEFT", Buffer, App.Path & "\" & glFileCFG)

End Sub
'Purpose:routine di ausilio clip immagine placca / routine in aid of the platen image clip
Public Sub ClipPicture()
Dim Rgn_Source  As Long
Dim Rgn_Dest As Long
Dim ll As Long
On Error Resume Next
Rgn_Source = CreateRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight)
Rgn_Dest = CreateRectRgn(0, 0, 1, 1)
ll = CombineRgn(Rgn_Dest, Rgn_Source, Rgn_TOT, RGN_DIFF)
ll = SetWindowRgn(Me.hwnd, Rgn_Dest, True)

ll = UpdateWindow(Me.hwnd)

End Sub
'Purpose:movimentazione immagine esci / exit image
Public Sub pctEsci_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If glBusy = True Then
    Beep
    Exit Sub
End If
If Button = 1 Then
    pctEsci.Top = pctEsci.Top + 2
End If
End Sub
'Purpose:movimentazione immagine esci / exit image
Private Sub pctEsci_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
pctEsci.SetFocus
End Sub
'Purpose:movimentazione immagine esci /exit image
Public Sub pctEsci_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim RetVal As Long
Dim Buffer As String

If glBusy = True Then
    Beep
    Exit Sub
End If
If Button = 1 Then
    On Error Resume Next
    Call sndPlaySound(glVirtualSw.Wav, SND_ASYNC)
    pctEsci.Top = 123
    Me.Hide
    frmMenu.Winsock1.Close
    Call InsLog("Disconnessione da " & frmMenu.Winsock1.RemoteHostIP)
    frmMain.Timer1.Enabled = False
    lcTO = 0

    Buffer = CStr(frmMain.Top)
    glMainTop = frmMain.Top
    RetVal = WritePrivateProfileString("GENERAL", "TOP", Buffer, App.Path & "\" & glFileCFG)
    
    Buffer = CStr(frmMain.Left)
    glMainLeft = frmMain.Left
    RetVal = WritePrivateProfileString("GENERAL", "LEFT", Buffer, App.Path & "\" & glFileCFG)

End If
End Sub
'Purpose:movimentazione placca / platen movement
Private Sub pctPlacca_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
OffSetX = x
OffSetY = Y
ButtonPress = True
End Sub
'Purpose:movimentazione placca / platen movement
Private Sub pctPlacca_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If ButtonPress = True Then
    If glBusy = True Then
        ButtonPress = False
    Else
        frmMain.Left = frmMain.Left + (x - OffSetX) * Screen.TwipsPerPixelX
        frmMain.Top = frmMain.Top + (Y - OffSetY) * Screen.TwipsPerPixelY
    End If
End If
End Sub
'Purpose:movimentazione placca / platen movement
Private Sub pctPlacca_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
ButtonPress = False
glMainLeft = frmMain.Left
glMainTop = frmMain.Top
End Sub
'Purpose:movimentazione tasto / button movement
Private Sub pctTasto_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim PixelY As Long

If glBusy = True Then
    Exit Sub
End If
PixelY = Y / Screen.TwipsPerPixelY
If Button = 1 Then
    If PixelY <= pctTasto(Index).Height / 3 Then
        ' superiore
        If glVirtualSw.Tasti(CStr(Index)).Abilitato = True Then
            pctTasto(Index).Top = pctTasto(Index).Top + 2
        End If
    ElseIf PixelY >= pctTasto(Index).Height * 2 / 3 Then
        ' inferiore
        If glVirtualSw.Tasti(CStr(Index + 1)).Abilitato = True Then
            pctTasto(Index).Top = pctTasto(Index).Top - 2
        End If
    End If
End If
End Sub
'Purpose:movimentazione tasto / button movement
Private Sub pctTasto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim PixelY As Long


PixelY = Y / Screen.TwipsPerPixelY
On Error Resume Next
If PixelY <= pctTasto(Index).Height / 3 Then
    If glVirtualSw.Tasti(CStr(Index)).Abilitato = True Then
        ' superiore
        pctTasto(Index).SetFocus
        pctTasto(Index).ToolTipText = glVirtualSw.Tasti(CStr(Index)).ToolTip
    End If
ElseIf PixelY >= pctTasto(Index).Height * 2 / 3 Then
    ' inferiore
    If glVirtualSw.Tasti(CStr(Index + 1)).Abilitato = True Then
        pctTasto(Index).SetFocus
        pctTasto(Index).ToolTipText = glVirtualSw.Tasti(CStr(Index + 1)).ToolTip
    End If
Else
    pctTasto(Index).ToolTipText = ""
End If

End Sub
'Purpose:movimentazione tasto e invio comando / button movement and command sending
Private Sub pctTasto_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim PixelY As Long
Dim RetVal As Long

If glBusy = True Then
    Exit Sub
End If
PixelY = Y / Screen.TwipsPerPixelY

If Button = 1 Then
    On Error Resume Next
    If PixelY <= pctTasto(Index).Height / 3 Then
        If glVirtualSw.Tasti(CStr(Index)).Abilitato = True Then
            ' superiore
            pctTasto(1).Refresh
            DoEvents
            pctTasto(3).Refresh
            DoEvents
            pctTasto(5).Refresh
            DoEvents
            Call sndPlaySound(glVirtualSw.Wav, SND_ASYNC)
            glBusy = True
            DoEvents
            RetVal = SendTCP(CStr(Index), glcButtonSup)
            If RetVal = False Then
                Call MsgBox("Impossibile eseguire il comando. Mancata risposta da server", vbCritical, "Attenzione")
                InsLog "Impossibile eseguire il comando. Mancata risposta da server"
            End If
            glBusy = False
        End If
    ElseIf PixelY >= pctTasto(Index).Height * 2 / 3 Then
        ' inferiore
        If glVirtualSw.Tasti(CStr(Index + 1)).Abilitato = True Then
            pctTasto(1).Refresh
            DoEvents
            pctTasto(3).Refresh
            DoEvents
            pctTasto(5).Refresh
            DoEvents
            Call sndPlaySound(glVirtualSw.Wav, SND_ASYNC)
            glBusy = True
            RetVal = SendTCP(CStr(Index + 1), glcButtonInf)
            If RetVal = False Then
                MsgBox "Impossibile eseguire il comando. Mancata risposta da server", vbCritical, "Attenzione"
                InsLog "Impossibile eseguire il comando. Mancata risposta da server"
            End If
            glBusy = False
        End If
    End If
End If
pctTasto(Index).Top = 34
End Sub
'Purpose:invio comando OPEN a remote host / sending OPEN command to a remote host
Public Function SendTCP(ButtonKey As String, ButtonType As Long) As Boolean
Dim Buffer As String
Dim TimerIni As Long
Dim Retry As Long
Dim RetVal As Long
Dim SessionOpen As Boolean

On Error Resume Next
If frmMenu.Winsock1.state <> sckConnected Then
    ' remote host non connesso, eseguo la connessione / remote host not connected, I do the connection
    RetVal = frmMenu.ConnectTCP(SessionOpen)
    If RetVal = False Then
        frmMain.lblWait.Caption = "Connessione non disponibile"
        Exit Function
    Else
        If SessionOpen = False Then
            ' remote host non riconosciuto (mancata ricezione ACK) / remote host not recognized (ACK not received)
            If frmMenu.Winsock1.state = sckConnected Then
                frmMenu.Winsock1.Close
            End If
            frmMain.lblWait.Caption = "Connessione non disponibile"
            Exit Function
        Else
            ' invio al dispositivo il tipo di servizio / sending service type to the device
            RetVal = frmMenu.BuildConnessione
            If RetVal = False Then
                If frmMenu.Winsock1.State = sckConnected Then
                    frmMenu.Winsock1.Close
                End If
                frmMain.lblWait.Caption = "Connessione non disponibile"
                Exit Function
            Else
                frmMain.lblWait.Caption = ""
            End If
        End If
    End If
End If

' caricamento buffer da trasmettere / loading buffer for trasmission
Buffer = glVirtualSw.Tasti(ButtonKey).OPENCMD
lcTO = 0
If frmMenu.Winsock1.state = sckConnected Then
    Do
        ' invio dati via TCP per 3 volte e attesa ACK / sending record by TCP for 3 times and waiting ACK
        Call InsLog("Invio " & Buffer)
        BufferInTCP = ""
        frmMenu.Winsock1.SendData Buffer
        TimerIni = Timer
        ' attesa ACK/NACK per un timeout di 10 secondi / waiting 10 seconds for ACK/NACK
        Do
            If Timer < TimerIni Then
                TimerIni = Timer
            End If
            DoEvents
            If InStr(BufferInTCP, glcACKTCP) > 0 Then
                ' ricevuto ACK / ACK received
                SendTCP = True
                Call InsLog("Ricevuto ACK")
                Exit Function
            ElseIf InStr(BufferInTCP, glcNACKTCP) > 0 Then
                ' ricevuto NACK / NACK received
                Call InsLog("Ricevuto NACK")
                BufferInTCP = ""
                Exit Do
            End If
            DoEvents
            If frmMenu.Winsock1.state <> sckConnected Then
                Exit Function
            End If
        Loop Until Timer > TimerIni + 10
        BufferInTCP = ""
        Retry = Retry + 1
        If Retry >= 3 Then
            ' nessuna ricezione ACK dopo 3 tentativi / ACK not received after 3 tries
            Call InsLog("Mancata ricezione ACK dopo 3 retry")
            SendTCP = False
            Exit Function
        End If
    Loop Until Retry > 3
End If
End Function
'Purpose:conteggio timeout di visualizzazione placca / counting time for platen display
Private Sub Timer1_Timer()
lcTO = lcTO + 1
If lcTO >= 6 Then
    If glBusy = False Then
        Call pctEsci_MouseUp(1, 0, 0, 0)
    End If
End If
End Sub
