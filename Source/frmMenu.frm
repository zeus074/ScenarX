VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scenar-X"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   8895
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Clog 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCE2E9&
      Caption         =   "Log"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7950
      MaskColor       =   &H00DCE2E9&
      TabIndex        =   56
      Top             =   4860
      Width           =   765
   End
   Begin VB.CheckBox Cvirt 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCE2E9&
      Caption         =   "modello ""Virt."""
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6150
      MaskColor       =   &H00DCE2E9&
      TabIndex        =   55
      Top             =   4860
      Width           =   1515
   End
   Begin VB.CommandButton Cdelsce 
      Caption         =   "Elimina scenario 1"
      Height          =   315
      Left            =   7050
      TabIndex        =   46
      Top             =   5250
      Width           =   1665
   End
   Begin VB.CommandButton Cunlock 
      Caption         =   "Sblocca"
      Height          =   315
      Left            =   3150
      TabIndex        =   45
      Top             =   5625
      Width           =   915
   End
   Begin VB.CommandButton Clock 
      Caption         =   "Blocca"
      Height          =   315
      Left            =   3150
      TabIndex        =   44
      Top             =   5250
      Width           =   915
   End
   Begin VB.TextBox TPLscen 
      Height          =   285
      Left            =   4050
      MaxLength       =   1
      TabIndex        =   43
      Text            =   "1"
      Top             =   4875
      Width           =   240
   End
   Begin VB.TextBox TAscen 
      Height          =   285
      Left            =   3375
      MaxLength       =   1
      TabIndex        =   42
      Text            =   "0"
      Top             =   4875
      Width           =   240
   End
   Begin VB.CommandButton Cdelsceall 
      Caption         =   "Elimina Tutti"
      Height          =   315
      Left            =   7050
      TabIndex        =   41
      Top             =   5625
      Width           =   1665
   End
   Begin VB.CommandButton CrunSce 
      Caption         =   "Attiva scenario 1"
      Height          =   315
      Left            =   4650
      TabIndex        =   40
      Top             =   5250
      Width           =   1665
   End
   Begin VB.CommandButton CstartPrg 
      Caption         =   "Start"
      Height          =   315
      Left            =   5700
      TabIndex        =   39
      Top             =   7245
      Width           =   690
   End
   Begin VB.CommandButton CstopPrg 
      Caption         =   "Stop"
      Height          =   315
      Left            =   6450
      TabIndex        =   38
      Top             =   7245
      Width           =   690
   End
   Begin VB.ComboBox CbusSce 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4425
      TabIndex        =   37
      Top             =   4875
      Width           =   1590
   End
   Begin VB.CommandButton CrunSceN 
      Caption         =   "Attiva scenario ..."
      Height          =   315
      Left            =   4650
      TabIndex        =   36
      Top             =   5625
      Width           =   1665
   End
   Begin VB.TextBox Trit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3225
      MaxLength       =   2
      TabIndex        =   35
      Text            =   "2"
      Top             =   6225
      Width           =   390
   End
   Begin VB.TextBox Trit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3225
      MaxLength       =   2
      TabIndex        =   34
      Text            =   "5"
      Top             =   6450
      Width           =   390
   End
   Begin VB.TextBox Trit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3225
      MaxLength       =   2
      TabIndex        =   33
      Text            =   "5"
      Top             =   6675
      Width           =   390
   End
   Begin VB.TextBox Trit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3225
      MaxLength       =   2
      TabIndex        =   32
      Text            =   "1"
      Top             =   6900
      Width           =   390
   End
   Begin VB.CheckBox CDelPrg 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFE9DD&
      Caption         =   "Elimina prima"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7350
      MaskColor       =   &H00DCE2E9&
      TabIndex        =   31
      Top             =   6200
      Width           =   1365
   End
   Begin VB.CommandButton CprogMore 
      Caption         =   "Programma lo scenario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   7350
      TabIndex        =   30
      Top             =   6525
      Width           =   1365
   End
   Begin VB.ComboBox Ccomando 
      Height          =   315
      Left            =   2775
      TabIndex        =   22
      Top             =   2400
      Width           =   2190
   End
   Begin VB.ComboBox Cdove 
      Height          =   315
      Left            =   2775
      TabIndex        =   21
      Top             =   3000
      Width           =   2190
   End
   Begin VB.CommandButton Cadd 
      Caption         =   "< Inserisci su scenario 1"
      Height          =   315
      Left            =   2775
      TabIndex        =   19
      Top             =   3975
      Width           =   1965
   End
   Begin VB.TextBox Topen 
      Height          =   315
      Left            =   2775
      TabIndex        =   18
      Top             =   3600
      Width           =   4665
   End
   Begin VB.TextBox Topt2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6825
      TabIndex        =   17
      Top             =   2400
      Width           =   1590
   End
   Begin VB.CommandButton Csend 
      Caption         =   "Invia -> Scs"
      Height          =   315
      Left            =   7650
      TabIndex        =   16
      Top             =   3600
      Width           =   1065
   End
   Begin VB.ComboBox Copt 
      Height          =   315
      Left            =   5100
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.ComboBox Cbus 
      Height          =   315
      Left            =   5100
      TabIndex        =   14
      Top             =   3000
      Width           =   1590
   End
   Begin VB.TextBox txtIPServer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   825
      MaxLength       =   15
      TabIndex        =   10
      Top             =   300
      Width           =   1590
   End
   Begin VB.TextBox txtPortServer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      MaxLength       =   5
      TabIndex        =   9
      Top             =   675
      Width           =   990
   End
   Begin VB.Timer TbarT 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8175
      Top             =   -150
   End
   Begin MSComctlLib.ProgressBar BarT 
      Height          =   240
      Left            =   6440
      TabIndex        =   8
      Top             =   7800
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox TnSce 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3D3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2625
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1290
      Width           =   3840
   End
   Begin VB.ComboBox Cnum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Text            =   "1"
      Top             =   1275
      Width           =   765
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   240
      Left            =   7500
      TabIndex        =   4
      Top             =   7800
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ListBox LstCMD 
      Height          =   5910
      Left            =   75
      TabIndex        =   0
      Top             =   1650
      Width           =   2490
   End
   Begin MSComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   7740
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11218
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1817
            MinWidth        =   1817
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8475
      Top             =   -75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.1.0.102"
      RemotePort      =   10101
   End
   Begin VB.TextBox Topt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   20
      Top             =   2400
      Width           =   1590
   End
   Begin VB.Label LbSce 
      BackStyle       =   0  'Transparent
      Caption         =   "Centralina Scenari"
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   2775
      TabIndex        =   54
      Top             =   4575
      Width           =   5940
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PL:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   3750
      TabIndex        =   53
      Top             =   4875
      Width           =   240
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   3150
      TabIndex        =   52
      Top             =   4875
      Width           =   150
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   1
      Left            =   2850
      Picture         =   "frmMenu.frx":0CCA
      Top             =   4875
      Width           =   240
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   5
      Left            =   6750
      Picture         =   "frmMenu.frx":1054
      Top             =   5250
      Width           =   240
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   6
      Left            =   2775
      Picture         =   "frmMenu.frx":13DE
      Top             =   7245
      Width           =   240
   End
   Begin VB.Label Lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmazione manuale scenario 1"
      ForeColor       =   &H00004000&
      Height          =   390
      Index           =   10
      Left            =   3000
      TabIndex        =   51
      Top             =   7200
      Width           =   1890
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   3
      Left            =   2850
      Picture         =   "frmMenu.frx":1768
      Top             =   5325
      Width           =   240
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   4
      Left            =   4350
      Picture         =   "frmMenu.frx":1AF2
      Top             =   5325
      Width           =   240
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   7
      Left            =   2850
      Picture         =   "frmMenu.frx":1E7C
      Top             =   6225
      Width           =   240
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo dopo comando temporizzato (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   3675
      TabIndex        =   50
      Top             =   6450
      Width           =   3465
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo dopo comando gruppo (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   3675
      TabIndex        =   49
      Top             =   6675
      Width           =   3465
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo inizio programmazione (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   3675
      TabIndex        =   48
      Top             =   6225
      Width           =   3465
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo dopo altro comando (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   3675
      TabIndex        =   47
      Top             =   6900
      Width           =   3465
   End
   Begin VB.Label LbOpen 
      BackStyle       =   0  'Transparent
      Caption         =   "Comando OPEN"
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   2775
      TabIndex        =   29
      Top             =   1950
      Width           =   3015
   End
   Begin VB.Label Lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Cosa"
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   4
      Left            =   2775
      TabIndex        =   28
      Top             =   2175
      Width           =   990
   End
   Begin VB.Label Lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Opzioni"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   5
      Left            =   5100
      TabIndex        =   27
      Top             =   2175
      Width           =   1590
   End
   Begin VB.Label Lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Dove"
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   7
      Left            =   2775
      TabIndex        =   26
      Top             =   2775
      Width           =   990
   End
   Begin VB.Label Lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Comando open"
      ForeColor       =   &H00004080&
      Height          =   240
      Index           =   8
      Left            =   2775
      TabIndex        =   25
      Top             =   3375
      Width           =   1590
   End
   Begin VB.Label Lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Opzioni"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   6
      Left            =   6825
      TabIndex        =   24
      Top             =   2175
      Width           =   1590
   End
   Begin VB.Label Lb 
      BackStyle       =   0  'Transparent
      Caption         =   "Bus"
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   9
      Left            =   5100
      TabIndex        =   23
      Top             =   2775
      Width           =   990
   End
   Begin VB.Label LscCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0 Comandi"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   150
      TabIndex        =   13
      Top             =   7575
      Width           =   2415
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   450
      TabIndex        =   12
      Top             =   300
      Width           =   225
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porta Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   11
      Top             =   675
      Width           =   1125
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   2
      Left            =   150
      Picture         =   "frmMenu.frx":2206
      Top             =   300
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   2925
      X2              =   8550
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Label Lfile 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuovo"
      Height          =   240
      Left            =   2925
      TabIndex        =   7
      Top             =   1650
      Width           =   5190
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   0
      Left            =   2625
      Picture         =   "frmMenu.frx":2590
      Top             =   1650
      Width           =   240
   End
   Begin VB.Shape Shsel 
      BorderColor     =   &H00C000C0&
      FillColor       =   &H0000FFFF&
      Height          =   390
      Left            =   4035
      Shape           =   5  'Rounded Square
      Top             =   645
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Lbltasto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0FF&
      Height          =   240
      Left            =   4800
      TabIndex        =   3
      Top             =   195
      Width           =   2190
   End
   Begin VB.Image Iico 
      Height          =   480
      Index           =   5
      Left            =   6600
      MouseIcon       =   "frmMenu.frx":291A
      MousePointer    =   99  'Custom
      Picture         =   "frmMenu.frx":2C24
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Iico 
      Height          =   480
      Index           =   4
      Left            =   6075
      MouseIcon       =   "frmMenu.frx":38EE
      MousePointer    =   99  'Custom
      Picture         =   "frmMenu.frx":3BF8
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Iico 
      Height          =   480
      Index           =   3
      Left            =   5550
      MouseIcon       =   "frmMenu.frx":48C2
      MousePointer    =   99  'Custom
      Picture         =   "frmMenu.frx":4BCC
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Iico 
      Height          =   480
      Index           =   2
      Left            =   5025
      MouseIcon       =   "frmMenu.frx":5896
      MousePointer    =   99  'Custom
      Picture         =   "frmMenu.frx":5BA0
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Iico 
      Height          =   480
      Index           =   1
      Left            =   4500
      MouseIcon       =   "frmMenu.frx":686A
      MousePointer    =   99  'Custom
      Picture         =   "frmMenu.frx":6B74
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Iico 
      Height          =   480
      Index           =   0
      Left            =   3975
      MouseIcon       =   "frmMenu.frx":783E
      MousePointer    =   99  'Custom
      Picture         =   "frmMenu.frx":7B48
      Top             =   600
      Width           =   480
   End
   Begin VB.Label LblCMD 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista comandi scenario"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   100
      TabIndex        =   2
      Top             =   1350
      Width           =   2490
   End
   Begin VB.Image ImgLogo 
      Height          =   1245
      Left            =   0
      Picture         =   "frmMenu.frx":8812
      Top             =   0
      Width           =   8895
   End
   Begin VB.Image Ibg 
      Height          =   6525
      Left            =   0
      Picture         =   "frmMenu.frx":11DA3
      Top             =   1260
      Width           =   8895
   End
   Begin VB.Menu mFle 
      Caption         =   "File"
      Begin VB.Menu mNuovo 
         Caption         =   "Nuovo"
      End
      Begin VB.Menu mApri 
         Caption         =   "Apri"
      End
      Begin VB.Menu mrecent 
         Caption         =   "Recente"
         Begin VB.Menu m1 
            Caption         =   "1"
         End
      End
      Begin VB.Menu mSalva 
         Caption         =   "Salva"
      End
      Begin VB.Menu mnuVuoto1 
         Caption         =   "-"
      End
      Begin VB.Menu mEsci 
         Caption         =   "Esci"
      End
   End
   Begin VB.Menu Mopt 
      Caption         =   "Opzioni"
      Begin VB.Menu mChkDupl 
         Caption         =   "Avverti indirizzi duplicati"
      End
      Begin VB.Menu mSep 
         Caption         =   "-"
      End
      Begin VB.Menu mCom 
         Caption         =   "Comunicazione"
      End
   End
   Begin VB.Menu mLanguage 
      Caption         =   "Lingua"
   End
   Begin VB.Menu mHlp 
      Caption         =   "?"
      Begin VB.Menu mChkUpd 
         Caption         =   "Controlla aggiornamenti"
      End
      Begin VB.Menu mbar 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu mHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* SCENAR-X V 1.1       Scritto da Gaggiottini Mirco                                    *
'*--------------------------------------------------------------------------------------*
'* protocollo di comunicazione di: bt_emanuele                                          *
'****************************************************************************************

Option Explicit
Dim Illum(31), Autom(108), Antif(31), Dove(108), Termo(15), Diff(6) As String
Dim DoveT(1100), DoveD(118), Aux(10), Cito(3) As String
Dim InProg As Boolean
Dim OldIP, OldPort, Scen(32, 100) As String, nSce(32) As String
Dim nFile As String, TMP As Variant
Dim X, Y, Z, Sel, Vel, Liv As Integer
Dim OldSce As Byte
Dim Rtv, Dirty As Boolean
Dim MN As Boolean
Private Declare Function ShellExecute& Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Implements iLanguage

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
Private Const LB_ITEMFROMPOINT = &H1A9
Dim ndx As Long

' notifica al dispositivo del tipo di servizio richiesto / this notifies the device of requested service type
Public Function BuildConnessione() As Boolean
Dim TimerIni As Double
Dim Buffer As String
Dim BufferAns As String
Dim RetVal As Long

' invio tipo servizio / sending service type
If glSsck Then Buffer = "*99*9##" Else Buffer = "*99*0##"
If Winsock1.State = sckConnected Then
    
    BufferInTCP = ""
    Log Time & "> * connessione eth " & Winsock1.RemoteHostIP & ":" & Winsock1.RemotePort & " *"
    Log Time & "> * invio " & Buffer & " *"
    Winsock1.SendData Buffer
Riprova:
    TimerIni = Timer
    Do
        If Timer < TimerIni Then TimerIni = Timer
        DoEvents
        If InStr(BufferInTCP, glcACKTCP) > 0 Then
            ' OK
            BuildConnessione = True: Log Time & "> * connesso *"
            'Call InsLog("Ricevuto:" & BufferInTCP)
            Exit Function
        ElseIf InStr(BufferInTCP, "*#") > 0 And Right(BufferInTCP, 1) = "#" Then
        'ricerca password
            If PassOpen = "" Then
            BuildConnessione = True: Log Time & "> * attenzione: richiesta password *"
            MsgBox "Error: Open-password required", vbExclamation
            BuildConnessione = False: Exit Function
            Else
            Winsock1.SendData OpenPwd(BufferInTCP, PassOpen)
            GoTo Riprova
            End If
        End If
        DoEvents
        If Winsock1.State <> sckConnected Then
            Exit Function
        End If
    Loop Until Timer > TimerIni + 10
End If
End Function

'Purpose:invio comando OPEN a remote host / sending OPEN command to a remote host
Public Function SendTCP(OpenCMD As String) As String
Dim Buffer As String
Dim TimerIni As Long
Dim Retry As Long
Dim RetVal As Long
Dim SessionOpen As Boolean
Dim lcTO As Integer
SendTCP = False

On Error Resume Next
If Winsock1.State <> sckConnected Then
    ' remote host non connesso, eseguo la connessione / remote host not connected, I do the connection
    RetVal = ConnectTCP(SessionOpen)
    If RetVal = False Then
        StBar.Panels(1).Text = txtNoCon
        SendTCP = False
        Exit Function
    Else
        If SessionOpen = False Then
            ' remote host non riconosciuto (mancata ricezione ACK) / remote host not recognized (ACK not received)
            If Winsock1.State = sckConnected Then
                Winsock1.Close
            End If
            StBar.Panels(1).Text = txtNoCon
            SendTCP = False
            Exit Function
        Else
            ' invio al dispositivo il tipo di servizio / sending service type to the device
            RetVal = BuildConnessione
            If RetVal = False Then
                If Winsock1.State = sckConnected Then
                    Winsock1.Close
                End If
                StBar.Panels(1).Text = txtNoCon
                SendTCP = False
                Exit Function
            Else
                
            End If
        End If
    End If
End If

' caricamento buffer da trasmettere / loading buffer for trasmission
Buffer = OpenCMD
lcTO = 0
If Winsock1.State = sckConnected Then
    Do
        ' invio dati via TCP per 3 volte e attesa ACK / sending record by TCP for 3 times and waiting ACK

        BufferInTCP = ""
        Winsock1.SendData Buffer
        TimerIni = Timer
        ' attesa ACK/NACK fino al timeout / waiting until timeout for ACK/NACK
        Do
            If Timer < TimerIni Then
                TimerIni = Timer
            End If
            DoEvents
            If InStr(BufferInTCP, glcACKTCP) > 0 Then
                ' ricevuto ACK / ACK received
                SendTCP = True
                StBar.Panels(1).Text = txtSend & " (" & Buffer & ")"
                
                Exit Function
            ElseIf InStr(BufferInTCP, glcNACKTCP) > 0 Then
                ' ricevuto NACK / NACK received
               SendTCP = False
                BufferInTCP = ""
                Exit Do
            End If
            DoEvents
            If Winsock1.State <> sckConnected Then
            SendTCP = False
                Exit Function
            End If
        Loop Until Timer > TimerIni + TimWait
        BufferInTCP = ""
        Retry = Retry + 1
        If Retry >= TimRetry Then
            ' nessuna ricezione ACK dopo x tentativi / ACK not received after x tries
            StBar.Panels(1).Text = txtNoAck & "  " & Buffer
            SendTCP = False
            Exit Function
        End If
    Loop Until Retry > TimRetry
End If
End Function

Private Sub Cadd_Click()
Dim Spl, Spl2 As Variant, AlIns As Boolean
If LstCMD.ListCount > 99 Then MsgBox txtFullSce, vbExclamation, App.EXEName: Exit Sub

If Left(Topen, 1) <> "*" Or Right(Topen, 2) <> "##" Then MsgBox txtNoOpen, vbExclamation, App.EXEName: Exit Sub

'controllo se già inserito
If Not mChkDupl.Checked Then GoTo NoChk
Spl = Split(Topen, "*")
If UBound(Spl) < 3 Then MsgBox txtNoOpen, vbExclamation, App.EXEName: Exit Sub
For X = 0 To LstCMD.ListCount - 1
Spl2 = Split(LstCMD.List(X), "*")
Select Case Spl(1)
Case 1, 2, 9, 16, "#16"
If Left(Spl(1), 1) <> "#" Then
If Spl(1) = Spl2(1) And Spl(3) = Spl2(3) Then LstCMD.ListIndex = X: AlIns = True: Exit For
Else
If Spl(1) = Spl2(1) And Spl(2) = Spl2(2) Then LstCMD.ListIndex = X: AlIns = True: Exit For
End If
'If Left(Spl(3), 1) = "#" Then DelayM = DelayG
End Select
Next X

    If AlIns Then
    TMP = MsgBox(txtDuplicate, vbQuestion + vbYesNo, txtDuplicate2)
    If TMP = vbNo Then Exit Sub
    End If
    
NoChk:
Dirty = True
LstCMD.AddItem Topen
LscCount = LstCMD.ListCount & " " & txtCom
Scen(Cnum.Text, LstCMD.ListCount) = Topen
End Sub

Private Sub Cbus_Click()
If Cbus.ListIndex < 0 Or Cbus.ListIndex < 0 Or Ccomando.ListIndex < 0 Then Exit Sub
GeneraOpen
End Sub

Private Sub Cbus_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CbusSce_Click()
SaveSetting App.EXEName, "Option", "BusSce", CbusSce.ListIndex
End Sub

Private Sub CheckAll()
'controllo lista programmazione
Dim RetVal As Boolean
Dim Tmr As Double
If glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = False
Screen.MousePointer = vbHourglass

For X = 0 To LstCMD.ListCount - 1
LstCMD.Selected(X) = True
RetVal = SendTCP(LstCMD.List(X))
If RetVal = False Then StBar.Panels(1) = txtErrCom & " " & LstCMD.List(X): Screen.MousePointer = vbDefault: Exit Sub

Tmr = Timer
LOP:
DoEvents
If Timer < Tmr + 0.5 Then GoTo LOP
DoEvents

Next X

Screen.MousePointer = vbDefault
MsgBox txtAllCom, vbInformation
End Sub

Private Sub CbusSce_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Ccomando_Click()
If (Ccomando.ListIndex = 0 Or Ccomando.ListIndex = 1) And Sel = 1 Then
Topt.Enabled = True: Lb(5) = txtSpeed & " 0-255 (0=" & txtSpeed2 & ")": Lb(5).ForeColor = &H8000&
Lb(6) = txtOpt: Lb(6).ForeColor = &H808080
ElseIf (Ccomando.ListIndex = 29 Or Ccomando.ListIndex = 30) And Sel = 1 Then
Topt.Enabled = True: Lb(5) = txtLevel & " 1-100": Lb(5).ForeColor = &H8000&
Topt2.Enabled = True: Lb(6) = txtSpeed & " 0-255 (0=" & txtSpeed2 & ")": Lb(6).ForeColor = &H8000&
Else
Topt.Enabled = False: Topt = "": Topt2.Enabled = False: Topt2 = "": Lb(5) = txtOpt: Lb(6) = txtOpt
Lb(5).ForeColor = &H808080: Lb(6).ForeColor = &H808080
End If

If Sel = 4 Then
    If Ccomando.ListIndex = 3 Or Ccomando.ListIndex = 14 Then
    Copt.Clear: For X = 0 To 74: Copt.AddItem Format(3 + (X * 0.5), "#.0 °C"): Next X
    Copt.Visible = True: Lb(5) = txtTemper & " 3-40°C": Lb(5).ForeColor = &H8000&: Topen = ""
    ElseIf Ccomando.ListIndex = 4 Or Ccomando.ListIndex = 5 Then
    Copt.Clear: For X = 1 To 16: Copt.AddItem txtSce & " " & X: Next X: Topen = ""
    Copt.Visible = True: Lb(5) = txtSce & " (1-16)": Lb(5).ForeColor = &H8000&
    ElseIf Ccomando.ListIndex = 7 Or Ccomando.ListIndex = 8 Then
    Copt.Clear: For X = 1 To 3: Copt.AddItem txtProg & " " & X: Next X: Topen = ""
    Copt.Visible = True: Lb(5) = txtProg & " (1-3)": Lb(5).ForeColor = &H8000&
    ElseIf Ccomando.ListIndex < 3 And (Ccomando.ListIndex > 5 And Ccomando.ListIndex < 14) Then
    Lb(5) = txtOpt: Lb(5).ForeColor = &H808080
    End If
End If

If Ccomando.ListIndex < 10 And Sel = 4 Then Cdove.Enabled = False: Lb(7).ForeColor = &H808080: Cdove.Text = "" Else Cdove.Enabled = True: Lb(7).ForeColor = &H8000&

If Sel = 16 Then
    If (Ccomando.ListIndex = 4 Or Ccomando.ListIndex = 5) And Not MN Then
    'Dove
    MN = True
    Cdove.Clear: For X = 0 To 8: Cdove.AddItem txtAmb & " " & X + 1: Next X
    Topen = ""
    ElseIf (Ccomando.ListIndex < 4 Or Ccomando.ListIndex > 5) And MN Then
    'Dove
    MN = False: Cdove.Clear: For X = 0 To 118: Cdove.AddItem DoveD(X): Next X
    Topen = ""
    End If
    
    If Ccomando.ListIndex = 6 Then
    Topt.Enabled = True: Lb(5) = txtVol & " 0-31": Lb(5).ForeColor = &H8000&
    Else
    Lb(5) = txtOpt: Lb(5).ForeColor = &H808080
    End If
End If
GeneraOpen
End Sub

Private Sub CDelPrg_Click()
SaveSetting App.EXEName, "Option", "Delprg", CDelPrg.Value
End Sub

Private Sub Cdelsce_Click()
Dim CmdDove As String
TMP = MsgBox(txtClear & Cnum.Text & " " & txtClear2, vbQuestion + vbYesNo, txtDelCap)
If TMP = vbYes Then
If Not glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = True
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")

'Controllo se esite
If Cvirt.Value = 1 Then
Rtv = SendTCP("*#1001*" & CmdDove & "*1##")
If Rtv <> True Then If noVirt = vbNo Then Exit Sub Else NoSave = True: Cvirt.Value = 0
End If

'INVIO frame cancella
Rtv = SendTCP("*0*42#" & Cnum.Text & "*" & CmdDove & "##")
If Rtv Then MsgBox txtDeleted, vbInformation
End If
End Sub

Private Function noVirt()
TMP = MsgBox(txtVirtErr & " (A:" & TAscen & " PL:" & TPLscen & " " & CbusSce.Text & _
") " & txtVirtErr2 & vbCrLf & txtVirtErr3, vbQuestion + vbYesNo, txtVirtErr4)
noVirt = TMP
End Function

Private Sub Cdelsceall_Click()
Dim CmdDove As String
TMP = MsgBox(txtClearAll, vbQuestion + vbYesNo, txtDelCap)
If TMP = vbYes Then
If Not glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = True
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")

'Controllo se esite
If Cvirt.Value = 1 Then
Rtv = SendTCP("*#1001*" & CmdDove & "*1##")
If Rtv <> True Then If noVirt = vbNo Then Exit Sub Else NoSave = True: Cvirt.Value = 0
End If

'invio frame
Rtv = SendTCP("*0*42*" & CmdDove & "##")
End If
If Rtv Then MsgBox txtDeleted, vbInformation
End Sub

Private Sub Cdove_Click()
If Cbus.ListIndex < 0 Or Cbus.ListIndex < 0 Then Exit Sub
GeneraOpen
End Sub

Private Sub GeneraOpen()
'*****************************************************************************************
'* GENERA COMANDO OPENWEBNET
'*****************************************************************************************
If Ccomando.ListIndex < 0 Or (Cdove.ListIndex < 0 And Sel <> 4) Then Exit Sub
Dim CmdDove, cmdCosa As String

Select Case Sel

Case 1, 2
'Illuminazione - Automazione
If Cdove.ListIndex < 10 Then
CmdDove = Cdove.ListIndex
ElseIf Cdove.ListIndex < 91 Then
CmdDove = Right(Cdove.Text, 2)
Else
CmdDove = "#" & Cdove.ListIndex - 90
End If
If Cbus.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(Cbus.ListIndex, "00")

If Ccomando.ListIndex < 19 Then cmdCosa = Ccomando.ListIndex Else cmdCosa = Ccomando.ListIndex + 1

If Ccomando.ListIndex = 0 Or Ccomando.ListIndex = 1 Then
    If Topt <> "" Then
    If Val(Topt) < 0 Then Topt = 0
    If Val(Topt) > 255 Then Topt = 255
    Vel = Val(Topt)
    Topen = "*" & Sel & "*" & cmdCosa & "#" & Vel & "*" & CmdDove & "##"
    Else
    Topen = "*" & Sel & "*" & cmdCosa & "*" & CmdDove & "##"
    End If
    
ElseIf Ccomando.ListIndex = 29 Or Ccomando.ListIndex = 30 Then
    If Topt <> "" And Topt2 <> "" Then
    If Val(Topt2) < 0 Then Topt = 0
    If Val(Topt2) > 255 Then Topt = 255
    Vel = Val(Topt2)
    If Val(Topt) < 1 Then Topt = 1
    If Val(Topt) > 100 Then Topt = 100
    Liv = Val(Topt)
    Topen = "*" & Sel & "*" & cmdCosa & "#" & Liv & "#" & Vel & "*" & CmdDove & "##"
    Else
    Topen = "*" & Sel & "*" & cmdCosa & "*" & CmdDove & "##"
    End If
Else
Topen = "*" & Sel & "*" & cmdCosa & "*" & CmdDove & "##"
End If

Case 4
'Termoregolazione
'centrale
If Ccomando.ListIndex <= 1 Then Topen = "*" & Sel & "*" & (2 - Ccomando.ListIndex) & "02*#0##"
If Ccomando.ListIndex = 2 Then Topen = "*" & Sel & "*303*#0##"
If Ccomando.ListIndex = 3 And Copt.ListIndex >= 0 Then Topen = "*#" & Sel & "*#0*#14*0" & Format(30 + (Copt.ListIndex) * 5, "000") & "*3##"
If (Ccomando.ListIndex = 4 Or Ccomando.ListIndex = 5) And Copt.ListIndex >= 0 Then Topen = "*" & Sel & "*" & Ccomando.ListIndex - 3 & "2" & Format(Copt.ListIndex + 1, "00") & "*#0##"
If Ccomando.ListIndex = 6 Then Topen = "*" & Sel & "*3200*#0##"
If (Ccomando.ListIndex = 7 Or Ccomando.ListIndex = 8) And Copt.ListIndex >= 0 Then Topen = "*" & Sel & "*" & Ccomando.ListIndex - 6 & "1" & Format(Copt.ListIndex + 1, "00") & "*#0##"
If Ccomando.ListIndex = 9 Then Topen = "*" & Sel & "*3100*#0##"
'zona
If Cdove.ListIndex < 0 Or Cdove.Text = "" Then Exit Sub
If Ccomando.ListIndex = 10 Or Ccomando.ListIndex = 11 Then Topen = "*" & Sel & "*" & (12 - Ccomando.ListIndex) & "02*#" & Cdove.ListIndex + 1 & "##"
If Ccomando.ListIndex = 12 Then Topen = "*" & Sel & "*303*#" & Cdove.ListIndex + 1 & "##"
If Ccomando.ListIndex = 13 Then Topen = "*" & Sel & "*311*#" & Cdove.ListIndex + 1 & "##"
If Ccomando.ListIndex = 14 And Copt.ListIndex >= 0 Then Topen = "*#" & Sel & "*#" & Cdove.ListIndex + 1 & "*#14*0" & Format(30 + (Copt.ListIndex) * 5, "000") & "*3##"
If Ccomando.ListIndex = 15 Then Topen = "*" & Sel & "*40*" & Cdove.ListIndex + 1 & "##"

Case 6
CmdDove = 4000 + Cdove.ListIndex
If Cbus.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Cbus.ListIndex
If Ccomando.ListIndex = 0 Then cmdCosa = 10
If Ccomando.ListIndex = 1 Then cmdCosa = 12
If Ccomando.ListIndex = 2 Then cmdCosa = 11
Topen = "*" & Sel & "*" & cmdCosa & "*" & CmdDove & "##"

Case 9
CmdDove = Cdove.ListIndex + 1
If Cbus.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Cbus.ListIndex
cmdCosa = Ccomando.ListIndex
Topen = "*" & Sel & "*" & cmdCosa & "*" & CmdDove & "##"

Case 16
'Diff. sonora
If Cdove.ListIndex = 0 Then
CmdDove = 0
ElseIf Cdove.ListIndex < 10 Then
CmdDove = "#" & Cdove.ListIndex
Else
CmdDove = Format(Cdove.ListIndex - 9, "00")
End If
If Cbus.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Cbus.ListIndex

If Ccomando.ListIndex < 2 Then
Topen = "*" & Sel & "*" & (Ccomando.ListIndex * 3) & "*" & CmdDove & "##"
ElseIf Ccomando.ListIndex < 4 Then
Topen = "*" & Sel & "*" & (10 + (Ccomando.ListIndex - 2) * 3) & "*" & CmdDove & "##"
ElseIf Ccomando.ListIndex < 6 Then
Topen = "*" & Sel & "*" & (20 + (Ccomando.ListIndex - 4) * 3) & "*1" & Cdove.ListIndex + 1 & "0##"
ElseIf Ccomando.ListIndex = 6 Then
If Topt = "" Then Topt = 0
If Val(Topt) > 31 Then Topt = 31
Topen = "*#" & Sel & "*" & CmdDove & "*#1*" & Val(Topt) & "*##"

End If

End Select
End Sub

Private Sub Clog_Click()
SaveSetting App.EXEName, "Option", "Log", Clog.Value
End Sub

Private Sub Cnum_Click()
If PrgAct Then Exit Sub
'If Cnum.Text = OldSce Then Exit Sub
OldSce = Cnum.Text
If nSce(Cnum.Text) = "" Then TnSce = txtNoname Else TnSce = nSce(Cnum.Text)
'Cprog.Caption = txtProgram & " " & Cnum.Text
Cdelsce.Caption = txtDelete & " " & Cnum.Text
Cadd.Caption = "< " & txtInsert & " " & Cnum.Text
CrunSce.Caption = txtActivate & " " & Cnum.Text
Lb(10).Caption = txtManProg & " " & Cnum.Text

LstCMD.Clear
For X = 1 To 100
    If Scen(Cnum.Text, X) = "" Then
    Exit For
    Else
    LstCMD.AddItem Scen(Cnum.Text, X)
    End If
Next X

LscCount = LstCMD.ListCount & " " & txtCom
'If LstCMD.ListIndex >= 0 Then CDel.Enabled = True Else CDel.Enabled = False
'If LstCMD.ListCount > 0 Then Cchk.Enabled = True: CdelAll.Enabled = True Else Cchk.Enabled = False: CdelAll.Enabled = False

End Sub

Private Sub Cnum_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Copt_Click()
If Cbus.ListIndex < 0 Or Cbus.ListIndex < 0 Or Ccomando.ListIndex < 0 Then Exit Sub
GeneraOpen
End Sub

Private Sub Copt_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Programma(Optional Scenari As String)

On Local Error GoTo ErrProg
Dim DoProg As Boolean

'Inizio programmazione / Start programming
Dim RetVal As Boolean
Dim Tmr As Double, Spl, Sp As Variant
Dim CmdDove As String, sAct As Integer

Log ""
Log Time & "> Inizio programmazione"

If Not glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = True: InProgram = True
Screen.MousePointer = vbHourglass
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")

If Scenari <> "" Then Sp = Split(Scenari, ",")

Y = 0
Cnt:
If Scenari <> "" Then
sAct = Sp(Y): Cnum.ListIndex = sAct - 1
If LstCMD.ListCount <= 0 Then Log Time & "> Scenario " & sAct & " senza comandi": GoTo NoProg
Else
sAct = Cnum.Text
End If

'Controllo se esite
If Cvirt.Value = 1 Then
Rtv = SendTCP("*#1001*" & CmdDove & "*1##")
Log Time & "> virt abilitato: controllo modulo scenari."
If Rtv <> True Then If noVirt = vbNo Then Screen.MousePointer = vbDefault: InProgram = False: Exit Sub Else NoSave = True: Cvirt.Value = 0: Log Time & "> modulo scenari non trovato continuo. (*#1001*" & CmdDove & "*1##)"
End If

'Cancella precedente / Delete previous
If CDelPrg.Value = 1 Then
Log Time & "> cancellazione scenario abilitato: cancello scenario " & sAct & ". (*0*42#" & sAct & "*" & CmdDove & "##)"
RetVal = SendTCP("*0*42#" & sAct & "*" & CmdDove & "##")

If RetVal <> True Then StBar.Panels(1) = txtErrProg: Screen.MousePointer = vbDefault: InProgram = False: Log Time & "> Errore cancellazione scenario. EXIT.": Exit Sub
StBar.Panels(1) = txtDeleted & " #" & sAct
Tmr = Timer
LOP0:
DoEvents
If Timer < Tmr + 1 Then GoTo LOP0
End If

' in programmazione
DoProg = True
Log Time & "> Start porgrammazione (*0*40#" & sAct & "*" & CmdDove & "##)"
RetVal = SendTCP("*0*40#" & sAct & "*" & CmdDove & "##")
DoEvents

If RetVal <> True Then StBar.Panels(1) = txtErrProg: Screen.MousePointer = vbDefault: InProgram = False: Log Time & "> Errore start programmazione": Exit Sub

StBar.Panels(1) = txtStart
BarT.Visible = True: BarT.Value = 0: BarT.Max = DelayS * 60: TbarT.Enabled = True
Tmr = Timer
LOP:
DoEvents
If Timer < Tmr + DelayS Then GoTo LOP
TbarT.Enabled = False: BarT.Value = 0

Bar1.Visible = True: Bar1 = 0: Bar1.Max = LstCMD.ListCount

'invio lista
For X = 0 To LstCMD.ListCount - 1
Bar1 = X + 1
Log Time & "> Invio comando " & X + 1 & ". (" & LstCMD.List(X) & ")"
SendTCP LstCMD.List(X)

DelayM = DelayC
Spl = Split(LstCMD.List(X), "*")
If Spl(1) <= 2 Then
If Spl(1) = 1 And (Spl(2) >= 11 And Spl(2) <= 29) Then DelayM = DelayT
If Left(Spl(3), 1) = "#" Then DelayM = DelayG
End If

BarT.Max = DelayM * 60: TbarT.Enabled = True
Tmr = Timer
LOP2:
DoEvents
If Timer < Tmr + DelayM Then GoTo LOP2
DoEvents
TbarT.Enabled = False: BarT.Value = 0
Next X

StBar.Panels(1) = txtStop
BarT.Max = 60: TbarT.Enabled = True
Tmr = Timer
LOP3:
DoEvents
If Timer < Tmr + 1 Then GoTo LOP3
Bar1.Visible = False: TbarT.Enabled = False: BarT.Visible = False

'Fine programmazione
Log Time & "> Fine programmazione. (*0*41#" & sAct & "*" & CmdDove & "##)"
SendTCP "*0*41#" & sAct & "*" & CmdDove & "##"
NoProg:
If Scenari <> "" Then If Y < (UBound(Sp)) Then Y = Y + 1: GoTo Cnt

Screen.MousePointer = vbDefault
If DoProg Then MsgBox txtProgrammed, vbInformation, App.EXEName
StBar.Panels(1) = "": InProgram = False
Exit Sub

ErrProg:
Screen.MousePointer = vbDefault: InProg = False
MsgBox Err.Description, vbCritical, App.EXEName
End Sub

Private Sub CprogMore_Click()
For X = 1 To 16
'If nSce(x) = "" Then nSce(x) = txtNoname
If nSce(X) = txtNoname Then nSce(X) = ""
FrmPrg.Lsce.AddItem Format(X, "00") & " " & nSce(X)
Next X
FrmPrg.Lsce.Selected(Cnum.Text - 1) = True
FrmPrg.Lsce.ListIndex = Cnum.Text - 1

FrmPrg.Show vbModal, Me
If Ret = "" Then Exit Sub

Call Programma(Ret)

End Sub

Private Sub CrunSce_Click()
Dim CmdDove As String

glSsck = False
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")
SendTCP "*0*" & Cnum.Text & "*" & CmdDove & "##"
End Sub

Private Sub CrunSceN_Click()
For X = 1 To 16
'If nSce(x) = "" Then nSce(x) = txtNoname
If nSce(X) = txtNoname Then nSce(X) = ""
FrmSce.Lsce.AddItem Format(X, "00") & " " & nSce(X)
Next X
FrmSce.Show vbModal, Me
If Ret = "" Then Exit Sub

Dim CmdDove As String
glSsck = False
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")
SendTCP "*0*" & Ret & "*" & CmdDove & "##"

End Sub

Private Sub Csend_Click()
If Left(Topen, 1) <> "*" Or Right(Topen, 2) <> "##" Then MsgBox txtNoOpen, vbExclamation, App.EXEName: Exit Sub

If InProg Then
If Not glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = True
Else
glSsck = False
End If

SendTCP Topen
End Sub

Private Sub CstartPrg_Click()
Dim RetVal As Boolean
Dim CmdDove As String
If Not glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = True

CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")

'Controllo se esite
If Cvirt.Value = 1 Then
Rtv = SendTCP("*#1001*" & CmdDove & "*1##")
If Rtv <> True Then If noVirt = vbNo Then Exit Sub Else NoSave = True: Cvirt.Value = 0
End If

RetVal = SendTCP("*0*40#" & Cnum.Text & "*" & CmdDove & "##")
If RetVal <> True Then StBar.Panels(1) = txtErrProg: Exit Sub
StBar.Panels(1) = txtInProg & " " & Cnum.Text & " ..."
InProg = True
End Sub

Private Sub CstopPrg_Click()
Dim RetVal As Boolean
Dim CmdDove As String

If Not glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = True
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")

'Controllo se esite
If Cvirt.Value = 1 Then
Rtv = SendTCP("*#1001*" & CmdDove & "*1##")
If Rtv <> True Then If noVirt = vbNo Then Exit Sub Else NoSave = True: Cvirt.Value = 0
End If

RetVal = SendTCP("*0*41#" & Cnum.Text & "*" & CmdDove & "##")

If RetVal <> True Then StBar.Panels(1) = txtErrEnd: Exit Sub
StBar.Panels(1) = txtOkEnd & " " & Cnum.Text
InProg = False
End Sub

Private Sub Clock_Click()
'blocca
Dim CmdDove As String
glSsck = True
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")

'Controllo se esite
If Cvirt.Value = 1 Then
Rtv = SendTCP("*#1001*" & CmdDove & "*1##")
If Rtv <> True Then If noVirt = vbNo Then Exit Sub Else NoSave = True: Cvirt.Value = 0
End If
SendTCP "*0*43*" & CmdDove & "##"
End Sub

Private Sub Cunlock_Click()
'sblocca
Dim CmdDove As String
glSsck = True
CmdDove = TAscen & TPLscen
If CbusSce.ListIndex > 0 Then CmdDove = CmdDove & "#4#" & Format(CbusSce.ListIndex, "00")

'Controllo se esite
If Cvirt.Value = 1 Then
Rtv = SendTCP("*#1001*" & CmdDove & "*1##")
If Rtv <> True Then If noVirt = vbNo Then Exit Sub Else NoSave = True: Cvirt.Value = 0
End If

SendTCP "*0*44*" & CmdDove & "##"
End Sub

Private Sub Cvirt_Click()
On Local Error Resume Next
If NoSave Then NoSave = False: Exit Sub
SaveSetting App.EXEName, "Option", "Virt", Cvirt.Value

End Sub

Private Sub Form_Unload(Cancel As Integer)
If InProgram Then Cancel = True: Exit Sub
If Winsock1.State = sckConnected Then Winsock1.Close
End
End Sub

Private Sub iLanguage_Updated()
LoadLanguageStrings
End Sub

Private Sub LoadLanguageStrings()
On Local Error GoTo ErrDes

With gobjLanguage
'Menù
    If .GetLabel("MainMnu", "New") = "" Then GoTo NoLingua
    mNuovo.Caption = .GetLabel("MainMnu", "New")
    mApri.Caption = .GetLabel("MainMnu", "Open")
    mrecent.Caption = .GetLabel("MainMnu", "Recent")
    mSalva.Caption = .GetLabel("MainMnu", "Save")
    mEsci.Caption = .GetLabel("MainMnu", "Exit")
    Mopt.Caption = .GetLabel("MainMnu", "Option")
    mChkDupl.Caption = .GetLabel("MainMnu", "Duplicate")
    mCom.Caption = .GetLabel("MainMnu", "Comunication")
    mLanguage.Caption = .GetLabel("MainMnu", "Language")
    mChkUpd.Caption = .GetLabel("MainMnu", "Update")
    
'label
    txtFrmLng = .GetLabel("Text", "LngCap")
    Lb(1).Caption = .GetLabel("Text", "SrvPort")
    Lb(4).Caption = .GetLabel("Text", "What")
    txtOpt = .GetLabel("Text", "Option")
    Lb(7).Caption = .GetLabel("Text", "Where")
    Lb(8).Caption = .GetLabel("Text", "OpenCmd")
    LbSce.Caption = .GetLabel("Text", "FrScen")
    LblCMD.Caption = .GetLabel("Text", "List")
    txtNoname = .GetLabel("Text", "NoName")
    txtNew = .GetLabel("Text", "New")
    txtFrmOpen = .GetLabel("Text", "FrOpen")
    txtCmd(1) = .GetLabel("Text", "Lighting")
    txtCmd(2) = .GetLabel("Text", "Automation")
    txtCmd(3) = .GetLabel("Text", "Thermo")
    txtCmd(4) = .GetLabel("Text", "Video")
    txtCmd(5) = .GetLabel("Text", "Aux")
    txtCmd(6) = .GetLabel("Text", "Sound")
    txtManProg = .GetLabel("Text", "Manualprg")
    Lrit(0) = .GetLabel("Text", "StartDelay")
    Lrit(1) = .GetLabel("Text", "TimedDelay")
    Lrit(2) = .GetLabel("Text", "GroupDelay")
    Lrit(3) = .GetLabel("Text", "CommandDelay")
    Cvirt.Caption = .GetLabel("Text", "Virt")
    txtUpdChk = .GetLabel("Text", "UpdChk")
    txtUpdReady = .GetLabel("Text", "UpdReady")
    txtUpdAct = .GetLabel("Text", "UpdAct")
    txtUpdNew = .GetLabel("Text", "UpdNew")
    txtUpdDwn = .GetLabel("Text", "UpdDwn")
    txtUpdNo = .GetLabel("Text", "NoUpd")
    txtUpdErr = .GetLabel("Text", "UpdErr")
    txtUpdlink = .GetLabel("Text", "Link")
    txtSelAll = .GetLabel("Text", "SctAll")
    txtWack = .GetLabel("Text", "WaitAck")
    txtNtries = .GetLabel("Text", "Retry")
    txtFileLoaded = .GetLabel("Text", "FileLoad")
    
'Button
    Csend.Caption = .GetLabel("Button", "SendCmd")
    txtInsert = .GetLabel("Button", "Insert")
    Clock.Caption = .GetLabel("Button", "Lock")
    Cunlock.Caption = .GetLabel("Button", "Unlock")
    txtActivate = .GetLabel("Button", "Activate")
    CrunSceN.Caption = .GetLabel("Button", "Activate2")
    txtDelete = .GetLabel("Button", "Delete")
    Cdelsceall.Caption = .GetLabel("Button", "Deleteall")
    txtProgram = .GetLabel("Button", "Program")
    txtDel = .GetLabel("Button", "Remove")
    txtDelAll = .GetLabel("Button", "Removeall")
    txtChk = .GetLabel("Button", "Chk")
    CDelPrg.Caption = .GetLabel("Button", "DelProg")
    txtStart = .GetLabel("Button", "Start")
    txtStop = .GetLabel("Button", "Stop")
    CstartPrg.Caption = txtStart: CstopPrg.Caption = txtStop
    txtProgram2 = .GetLabel("Button", "Program2")
    CprogMore.Caption = txtProgram
    
'message
    txtDirty = .GetLabel("Message", "Dirty")
    txtDirtyCap = .GetLabel("Message", "DirtyCap")
    txtStatCon = .GetLabel("Message", "StatCon")
    txtStatDis = .GetLabel("Message", "StatDis")
    txtSck = .GetLabel("Message", "Sck")
    txtSsck = .GetLabel("Message", "Ssck")
    txtWaitCon = .GetLabel("Message", "WaitCon")
    txtNoAck = .GetLabel("Message", "NoAck")
    txtNoCon = .GetLabel("Message", "NoCon")
    txtNoOpen = .GetLabel("Message", "IncOpen")
    txtErrCom = .GetLabel("Message", "ErrCmd")
    txtAllCom = .GetLabel("Message", "CmdOk")
    txtErrEnd = .GetLabel("Message", "EndProgErr")
    txtOkEnd = .GetLabel("Message", "EndProgOk")
    txtErrProg = .GetLabel("Message", "ErrProg")
    txtInProg = .GetLabel("Message", "InProg")
    txtClearAll = .GetLabel("Message", "DelAll")
    txtClearAll2 = .GetLabel("Message", "DelAll2")
    txtClear = .GetLabel("Message", "DelScen")
    txtDelCap = .GetLabel("Message", "DelCap")
    txtClear2 = .GetLabel("Message", "DelScen2")
    TxtOpen = .GetLabel("Message", "Open")
    txtNf = .GetLabel("Message", "NotFound")
    txtNf2 = .GetLabel("Message", "NotFound2")
    txtSend = .GetLabel("Message", "Sent")
    txtDeleted = .GetLabel("Message", "Deleted")
    txtProgrammed = .GetLabel("Message", "Programmed")
    txtFullSce = .GetLabel("Message", "FullScen")
    txtSave = .GetLabel("Message", "Save")
    txtCom = .GetLabel("Message", "Commands")
    txtVirtErr = .GetLabel("Message", "VirtErr")
    txtVirtErr2 = .GetLabel("Message", "VirtErr2")
    txtVirtErr3 = .GetLabel("Message", "VirtErr3")
    txtVirtErr4 = .GetLabel("Message", "VirtErr4")
    txtDuplicate = .GetLabel("Message", "Duplicate")
    txtDuplicate2 = .GetLabel("Message", "Duplicate2")
    txtOldVer = .GetLabel("Message", "OldVer")
    
    'scs
    txtGen = .GetLabel("scs", "Gen")
    txtAmb = .GetLabel("scs", "Amb")
    txtGr = .GetLabel("scs", "Gr")
    txtPnt = .GetLabel("scs", "Point")
    txtSource = .GetLabel("scs", "Src")
    txtBusP = .GetLabel("scs", "Bus")
    txtBusL = .GetLabel("scs", "Bus2")
    txtSpeed = .GetLabel("scs", "Speed")
    txtSpeed2 = .GetLabel("scs", "Speed2")
    txtLevel = .GetLabel("scs", "Level")
    txtTemper = .GetLabel("scs", "Temperature")
    txtProg = .GetLabel("scs", "Program")
    txtSce = .GetLabel("scs", "Scenario")
    txtVol = .GetLabel("scs", "Volume")

    'CHI (1 illuminazione)
    For X = 0 To 31
    If X <> 19 Then Illum(X) = .GetLabel("scs", "Lig" & X)
    Next X
    
    'CHI (2 automazione)
    For X = 0 To 2
    Autom(X) = .GetLabel("scs", "Aut" & X)
    Next X
    
    'CHI (4 termoregolazione)
    For X = 0 To 15
    Termo(X) = .GetLabel("scs", "Tem" & X)
    Next X

    'CHI (citofonia)
    For X = 0 To 2
    Cito(X) = .GetLabel("scs", "Vid" & X)
    Next X

    'CHI (9 ausiliari)
    For X = 0 To 10
    Aux(X) = .GetLabel("scs", "Aux" & X)
    Next X

    'CHI (16 Diff. Sonora)
    For X = 0 To 6
    Diff(X) = .GetLabel("scs", "Snd" & X)
    Next X
    
    GoTo OKlingua
NoLingua:

OKlingua:
'Dove automazione
Cdove.Clear
Dove(0) = txtGen: Cdove.AddItem Dove(0)
For X = 1 To 9: Dove(X) = txtAmb & " " & X: Cdove.AddItem Dove(X): Next X
For X = 11 To 19: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 21 To 29: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 31 To 39: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 41 To 49: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 51 To 59: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 61 To 69: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 71 To 79: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 81 To 89: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 91 To 99: Dove(X) = txtPnt & " " & X: Cdove.AddItem Dove(X): Next X
For X = 100 To 108: Dove(X) = txtGr & " " & X - 99: Cdove.AddItem Dove(X): Next X

'Dove Diff. Sonora
DoveD(0) = txtGen & " Ampli."
For X = 1 To 9: DoveD(X) = "Ampli. Amb. " & X: Next X
For X = 10 To 108: DoveD(X) = "Ampli. " & X - 9: Next X
DoveD(109) = "Gen. " & txtSource
For X = 110 To 118: DoveD(X) = txtSource & " " & X - 109: Next X

'Dove bus
Cbus.Clear: CbusSce.Clear
Cbus.AddItem txtBusP: CbusSce.AddItem txtBusP
For X = 1 To 9: Cbus.AddItem txtBusL & " I=" & X: CbusSce.AddItem txtBusL & " I=" & X: Next X
Cbus.ListIndex = 0: CbusSce.ListIndex = GetSetting(App.EXEName, "Option", "BusSce", 0)
TnSce = txtNoname
Lfile = txtNew
Lb(5).Caption = txtOpt: Lb(6).Caption = txtOpt
CrunSce.Caption = txtActivate & " " & Cnum.Text
Cdelsce.Caption = txtDelete & " " & Cnum.Text
'Cprog.Caption = txtProgram & " " & Cnum.Text
Lb(10).Caption = txtManProg & " " & Cnum.Text


End With
Sel = 0:  Iico_Click (Lindex)

Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, App.EXEName
End Sub

Private Sub Ibg_Click()
'
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim K As Long
Dim RetVal As Long
Dim Buffer As String
Dim sBuffer As Long
Dim FileToSave As String

VER = App.Major & "." & App.Minor & "." & App.Revision
Me.Caption = "Scenar-X  V." & VER & "B"
'PassOpen = 197482

' carico la configurazione generale / load general setting
txtIPServer = GetSetting(App.EXEName, "Option", "ServerIP", "127.0.0.1")
txtPortServer = GetSetting(App.EXEName, "Option", "ServerPORT", "20000")
TAscen = GetSetting(App.EXEName, "Option", "Asce", "0")
TPLscen = GetSetting(App.EXEName, "Option", "PLsce", "1")
Trit(0) = GetSetting(App.EXEName, "Option", "DelayS", "2")
Trit(1) = GetSetting(App.EXEName, "Option", "DelayT", "5")
Trit(2) = GetSetting(App.EXEName, "Option", "DelayG", "5")
Trit(3) = GetSetting(App.EXEName, "Option", "DelayC", "1")
TimWait = GetSetting(App.EXEName, "Option", "Wait", 10)
TimRetry = GetSetting(App.EXEName, "Option", "Retry", 3)
NoSave = True: Cvirt.Value = GetSetting(App.EXEName, "Option", "Virt", 1)
If GetSetting(App.EXEName, "Option", "Duplicate", 1) = True Then mChkDupl.Checked = True

DelayS = Trit(0): DelayT = Trit(1): DelayG = Trit(2): DelayC = Trit(3)
CDelPrg.Value = GetSetting(App.EXEName, "Option", "Delprg", "0")
Clog.Value = GetSetting(App.EXEName, "Option", "Virt", 0)

OldIP = txtIPServer
Winsock1.RemoteHost = txtIPServer
Winsock1.RemotePort = txtPortServer

m1.Caption = GetSetting(App.EXEName, "File", "last1", "")
Language = GetSetting(App.EXEName, "Option", "Language", "italiano")
SetLanguage Language ' "English"

For X = 1 To 16: Cnum.AddItem X: Next X: Cnum.ListIndex = 0
For X = 1 To 16
For Y = 1 To 100: Scen(X, Y) = "": Next Y
Next X

Copt.Top = Topt.Top

Shsel.Visible = True: Sel = 1
LbOpen.Caption = txtFrmOpen & " (" & txtCmd(1) & ")"

'altra istanza
    Dim CaptionForm As String
    If App.PrevInstance Then
        CaptionForm = Me.Caption
        Me.Caption = ""
        AppActivate CaptionForm
        End
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lbltasto <> "" Then Lbltasto = ""
End Sub

Private Sub Iico_Click(Index As Integer)
Lindex = Index: Lb(7).ForeColor = &H8000&
Topt.Enabled = False: Topt = "": Topt2.Enabled = False: Topt2 = "": Lb(5) = txtOpt: Lb(6) = txtOpt
Lb(5).ForeColor = &H808080: Lb(6).ForeColor = &H808080
Shsel.Left = Iico(Index).Left + 55: Shsel.Visible = True
Cdove.Text = "": Topen = "": MN = False: Cdove.Enabled = True
LbOpen.Caption = txtFrmOpen & " (" & txtCmd(Index + 1) & ")"

Select Case Index

Case 0
If Sel > 2 Then
'Dove
Cdove.Clear: Cdove.AddItem Dove(0)
For X = 1 To 9: Cdove.AddItem Dove(X): Next X
For X = 11 To 19: Cdove.AddItem Dove(X): Next X
For X = 21 To 29: Cdove.AddItem Dove(X): Next X
For X = 31 To 39: Cdove.AddItem Dove(X): Next X
For X = 41 To 49: Cdove.AddItem Dove(X): Next X
For X = 51 To 59: Cdove.AddItem Dove(X): Next X
For X = 61 To 69: Cdove.AddItem Dove(X): Next X
For X = 71 To 79: Cdove.AddItem Dove(X): Next X
For X = 81 To 89: Cdove.AddItem Dove(X): Next X
For X = 91 To 99: Cdove.AddItem Dove(X): Next X
For X = 100 To 108: Cdove.AddItem Dove(X): Next X
End If
Ccomando.Clear: Sel = 1: Copt.Visible = False
For X = 0 To 18: Ccomando.AddItem Illum(X): Next X
For X = 20 To 31: Ccomando.AddItem Illum(X): Next X

Case 1
If Sel > 2 Then
'Dove
Cdove.Clear: Cdove.AddItem Dove(0)
For X = 1 To 9: Cdove.AddItem Dove(X): Next X
For X = 11 To 19: Cdove.AddItem Dove(X): Next X
For X = 21 To 29: Cdove.AddItem Dove(X): Next X
For X = 31 To 39: Cdove.AddItem Dove(X): Next X
For X = 41 To 49: Cdove.AddItem Dove(X): Next X
For X = 51 To 59: Cdove.AddItem Dove(X): Next X
For X = 61 To 69: Cdove.AddItem Dove(X): Next X
For X = 71 To 79: Cdove.AddItem Dove(X): Next X
For X = 81 To 89: Cdove.AddItem Dove(X): Next X
For X = 91 To 99: Cdove.AddItem Dove(X): Next X
For X = 100 To 108: Cdove.AddItem Dove(X): Next X
End If
Ccomando.Clear: Sel = 2: Copt.Visible = False
For X = 0 To 2: Ccomando.AddItem Autom(X): Next X

Case 2
If Sel <> 4 Then
'Dove termoregolazione
Cdove.Clear
For X = 1 To 99: Cdove.AddItem "Zona " & X: Next X
End If
Ccomando.Clear: Sel = 4: 'Copt.Visible = True
For X = 0 To 15: Ccomando.AddItem Termo(X): Next X

Case 3
If Sel <> 6 Then
Ccomando.Clear: For X = 0 To 2: Ccomando.AddItem Cito(X): Next X
End If
Cdove.Clear: Sel = 6: Copt.Visible = False
For X = 0 To 95: Cdove.AddItem 4000 + X: Next X

Case 4
If Sel <> 9 Then
'Dove ausiliari
Cdove.Clear: For X = 1 To 9: Cdove.AddItem "Aux N." & X: Next X
End If
Ccomando.Clear: Sel = 9: Copt.Visible = False
For X = 0 To 10: Ccomando.AddItem Aux(X): Next X

Case 5
If Sel <> 16 Then
'Dove Diff. Sonora
Ccomando.Clear: For X = 0 To 6: Ccomando.AddItem Diff(X): Next X
'Dove
Cdove.Clear
For X = 0 To 118: Cdove.AddItem DoveD(X): Next X
End If
Sel = 16: Copt.Visible = False

End Select
End Sub

Private Sub Iico_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Lbltasto = txtCmd(Index + 1)
End Sub

Private Sub ImgLogo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lbltasto <> "" Then Lbltasto = ""
End Sub

Private Sub LstCMD_Click()
'If LstCMD.ListIndex >= 0 Then CDel.Enabled = True Else CDel.Enabled = False
'If LstCMD.ListCount > 0 Then Cchk.Enabled = True: CdelAll.Enabled = True Else Cchk.Enabled = False: CdelAll.Enabled = False
End Sub

Public Function ListBoxHit(ListBox As ListBox, ByVal X As Single, ByVal Y As Single) As Long
  
    ndx = SendMessage(LstCMD.hWnd, LB_ITEMFROMPOINT, 0, (Y \ Screen.TwipsPerPixelY) * 65536 + (X \ Screen.TwipsPerPixelX))
    If (ndx And &H10000) = &H10000 Then
        ListBoxHit = -1
    Else
        ListBoxHit = ndx
    End If
End Function

Private Sub LstCMD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrDesc

If Button = 2 Then
  ndx = ListBoxHit(LstCMD, X, Y)
  If Not ndx = -1 Then LstCMD.Selected(ndx) = True ' Else Exit Sub
If LstCMD.ListCount = 0 Then Exit Sub

Dim mii As MENUITEMINFO, Ret As Boolean
Dim CurPos As POINT_TYPE
Dim nuovoPopupMenu, MenuSel, RetVal As Long

nuovoPopupMenu = CreatePopupMenu()
mii.cbSize = Len(mii)
mii.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE

If Not ndx = -1 Then
With mii
.fType = MFT_STRING
.wID = 3
.dwTypeData = txtDel
.cch = Len(.dwTypeData)
End With
RetVal = InsertMenuItem(nuovoPopupMenu, 0, 1, mii)

With mii
.fType = MFT_STRING
.wID = 4
.dwTypeData = txtDelAll
.cch = Len(.dwTypeData)
End With
RetVal = InsertMenuItem(nuovoPopupMenu, 1, 1, mii)

With mii
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 2
End With
RetVal = InsertMenuItem(nuovoPopupMenu, 2, 1, mii)

With mii
.fType = MFT_STRING
'.fState = MFS_ENABLED Or MFS_DEFAULT 'assegnamo a questa voce un identificatore
.wID = 1
.dwTypeData = Csend.Caption
.cch = Len(.dwTypeData)
End With
RetVal = InsertMenuItem(nuovoPopupMenu, 3, 1, mii)
End If

With mii
.fType = MFT_STRING
.wID = 5
.dwTypeData = txtChk
.cch = Len(.dwTypeData)
End With
RetVal = InsertMenuItem(nuovoPopupMenu, 4, 1, mii)

RetVal = GetCursorPos(CurPos)
MenuSel = TrackPopupMenu(nuovoPopupMenu, TPM_TOPALIGN Or TPM_LEFTALIGN Or _
TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_LEFTBUTTON, CurPos.X, CurPos.Y, 0, _
LstCMD.hWnd, 0)
'distruggiamo il Popup-Menu
RetVal = DestroyMenu(nuovoPopupMenu)

Select Case MenuSel
Case 1
'test
If InProg Then
If Not glSsck And Winsock1.State = sckConnected Then Winsock1.Close
glSsck = True
Else
glSsck = False
End If

Ret = SendTCP(LstCMD.List(LstCMD.ListIndex))
If Ret = False Then StBar.Panels(1) = txtErrCom & " " & LstCMD.List(LstCMD.ListIndex)

Case 3
'rimuovi
Dirty = True
For X = LstCMD.ListIndex + 1 To 99: Scen(Cnum.Text, X) = Scen(Cnum.Text, X + 1): Next X
LstCMD.RemoveItem (LstCMD.ListIndex)
'If LstCMD.ListCount > 0 Then Cchk.Enabled = True Else Cchk.Enabled = False
LscCount = LstCMD.ListCount & " " & txtCom

Case 4
'rimuovi tutti
TMP = MsgBox(txtClearAll2, vbQuestion + vbYesNo, txtDelCap)
If TMP = vbNo Then Exit Sub

Dirty = True
LstCMD.Clear: ' Cchk.Enabled = False
For Y = 1 To 100: Scen(Cnum.Text, Y) = "": Next Y
LscCount = LstCMD.ListCount & " " & txtCom

Case 5
CheckAll

End Select
End If
Exit Sub
ErrDesc:
MsgBox Err.Description, vbCritical, App.EXEName
End Sub

Private Sub m1_Click()
If m1.Caption <> "" Then nFile = m1.Caption: Call mApri_Click
End Sub

Private Sub mAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mApri_Click()
On Local Error GoTo ErrDes
If Dirty Then
TMP = MsgBox(txtDirty, vbQuestion + vbYesNo, txtDirtyCap)
If TMP = vbNo Then Exit Sub
End If

Dim FL As String, L As String * 25
LoadVer = ""

'If Command$ <> "" Then FL = Replace(Command$, Chr$(34), ""): GoTo Apri
If nFile <> "" Then FL = nFile: GoTo Apri
FL = DialogFile(Me.hWnd, 1, TxtOpen & "", Lfile & ".scn", "File scenari" & Chr(0) & "*.scn", App.Path, "Scenari")
If FL = "" Then Exit Sub
Apri:

If Dir$(FL) = "" Then nFile = "": MsgBox txtNf & " " & FL & vbCrLf & txtNf2, vbExclamation, App.EXEName: Exit Sub
TMP = Right(FL, Len(FL) - InStrRev(FL, "\"))
Lfile = Left(TMP, Len(TMP) - 4)
nFile = ""
Bar1 = 0: Bar1.Max = 1600
'Programma
TMP = GetPrivateProfileString("Programma", "Ver", "", L, 20, FL)
If LTrim(Left$(L, TMP)) <> "" Then LoadVer = LTrim(Left$(L, TMP))

'Server
TMP = GetPrivateProfileString("Server", "IP", "", L, 20, FL)
If LTrim(Left$(L, TMP)) <> "" Then txtIPServer = LTrim(Left$(L, TMP))
TMP = GetPrivateProfileString("Server", "Port", "", L, 6, FL)
If LTrim(Left$(L, TMP)) <> "" Then txtPortServer = LTrim(Left$(L, TMP))
'Centralina
TMP = GetPrivateProfileString("Modulo", "A", "", L, 2, FL)
If LTrim(Left$(L, TMP)) <> "" Then TAscen = Val(LTrim(Left$(L, TMP)))
TMP = GetPrivateProfileString("Modulo", "PL", "", L, 2, FL)
If LTrim(Left$(L, TMP)) <> "" Then TPLscen = Val(LTrim(Left$(L, TMP)))
TMP = GetPrivateProfileString("Modulo", "Bus", "", L, 2, FL)
If LTrim(Left$(L, TMP)) <> "" Then CbusSce.ListIndex = Val(LTrim(Left$(L, TMP)))
TMP = GetPrivateProfileString("Modulo", "Virt", "", L, 2, FL)
If LTrim(Left$(L, TMP)) <> "" Then If Val(LTrim(Left$(L, TMP))) = 1 Then Cvirt.Value = 1 Else Cvirt.Value = 0

For X = 1 To 16
TMP = GetPrivateProfileString("Scenario " & X, "Nome", "", L, 31, FL)
nSce(X) = LTrim(Left$(L, TMP))
For Y = 1 To 100
Bar1 = Bar1 + 1
TMP = GetPrivateProfileString("Scenario " & X, Str$(Y), "", L, 25, FL)
Scen(X, Y) = LTrim(Left$(L, TMP))
DoEvents
Next Y
Next X

Bar1 = 0: Dirty = False: Cnum.ListIndex = 0: LstCMD.Clear: OldSce = Cnum.Text
If nSce(Cnum.Text) = "" Then TnSce = txtNoname Else TnSce = nSce(Cnum.Text)

For X = 1 To 100
    If Scen(Cnum.Text, X) = "" Then
    Exit For
    Else
    LstCMD.AddItem Scen(Cnum.Text, X)
    End If
Next X
m1.Caption = FL
SaveSetting App.EXEName, "File", "last1", FL

LscCount = LstCMD.ListCount & " " & txtCom
StBar.Panels(1).Text = txtFileLoaded & ": """ & Lfile & """ " & LoadVer

'controllo compatibilità
If LoadVer <> "" Then
TMP = Replace(LoadVer, ".", "")
Else
MsgBox txtOldVer, vbInformation, App.EXEName
End If

Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, App.EXEName
End Sub

Private Sub mChkDupl_Click()
If NoSave Then NoSave = False: Exit Sub
If mChkDupl.Checked Then mChkDupl.Checked = False Else mChkDupl.Checked = True
SaveSetting App.EXEName, "Option", "Duplicate", mChkDupl.Checked
End Sub

Private Sub mChkUpd_Click()
FrmUpd.Caption = mChkUpd.Caption
FrmUpd.Show vbModal, Me
End Sub

Private Sub mCom_Click()
FrmConn.Caption = mCom.Caption
FrmConn.Show vbModal, Me
End Sub

Private Sub mEsci_Click()
If Not InProg Then Unload Me
End Sub

Private Sub mHelp_Click()
ShellExecute Me.hWnd, "open", App.Path & "\ScenarX.pdf", ByVal 0&, ByVal 0&, ByVal 0&
End Sub

Private Sub mLanguage_Click()
If nSce(Cnum.Text) = txtNoname Then nSce(Cnum.Text) = ""
FrmLanguage.Caption = txtFrmLng
FrmLanguage.Show vbModal, Me
SetLanguage Language
Call Cnum_Click
End Sub

Private Sub mNuovo_Click()
If Dirty Then
TMP = MsgBox(txtDirty, vbQuestion + vbYesNo, txtDirtyCap)
If TMP = vbNo Then Exit Sub
End If

Lfile = txtNew
LstCMD.Clear
For X = 1 To 16
For Y = 1 To 100: Scen(X, Y) = "": nSce(X) = "": Next Y
Next X
Cnum.ListIndex = 0: Dirty = False: OldSce = Cnum.Text: TnSce = txtNoname
LscCount = LstCMD.ListCount & " " & txtCom
End Sub

Private Sub mSalva_Click()
On Local Error GoTo ErrDes
Dim FL As String, CR As String
FL = DialogFile(Me.hWnd, 0, txtSave & "", Lfile & ".scn", "File scenari" & Chr(0) & "*.scn", App.Path, "Scenari")
If FL = "" Then Exit Sub
TMP = Right(FL, Len(FL) - InStrRev(FL, "\"))
Lfile = Left(TMP, Len(TMP) - 4)
'Lfile = Sel
Bar1 = 0: Bar1.Max = 1600
'programma
CR = VER
WritePrivateProfileString "Programma", "Ver", CR, FL
'Server
WritePrivateProfileString "Server", "IP", txtIPServer.Text, FL
WritePrivateProfileString "Server", "Port", txtPortServer.Text, FL
'Centralina
WritePrivateProfileString "Modulo", "A", TAscen.Text, FL
WritePrivateProfileString "Modulo", "PL", TPLscen.Text, FL
CR = Trim(Str$(CbusSce.ListIndex))
WritePrivateProfileString "Modulo", "Bus", CR, FL
CR = Trim(Str$(Cvirt.Value))
WritePrivateProfileString "Modulo", "Virt", CR, FL
For X = 1 To 16
WritePrivateProfileString "Scenario" & Str$(X), "Nome", nSce(X), FL
For Y = 1 To 100
Bar1 = Bar1 + 1
WritePrivateProfileString "Scenario" & Str$(X), Str$(Y), Scen(X, Y), FL
DoEvents
Next Y
Next X
Bar1 = 0: Dirty = False
Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, App.EXEName
End Sub

Private Sub TAscen_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub TAscen_Validate(Cancel As Boolean)
SaveSetting App.EXEName, "Option", "Asce", TAscen
End Sub

Private Sub TbarT_Timer()
If BarT.Value < BarT.Max Then BarT.Value = BarT.Value + 1 Else TbarT.Enabled = False
End Sub

Private Sub TnSce_Change()
nSce(Cnum.Text) = TnSce
End Sub

Private Sub TnSce_Click()
TnSce.SelStart = 0: TnSce.SelLength = Len(TnSce)
End Sub

Private Sub Topen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Topt_Change()
If Cbus.ListIndex < 0 Or Cbus.ListIndex < 0 Or Ccomando.ListIndex < 0 Then Exit Sub
GeneraOpen
End Sub

Private Sub Topt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}": GeneraOpen: Exit Sub
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Topt2_Change()
If Cbus.ListIndex < 0 Or Cbus.ListIndex < 0 Or Ccomando.ListIndex < 0 Then Exit Sub
GeneraOpen
End Sub

Private Sub Topt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}": GeneraOpen: Exit Sub
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub TPLscen_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub TPLscen_Validate(Cancel As Boolean)
SaveSetting App.EXEName, "Option", "PLsce", TPLscen
End Sub

Private Sub Trit_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub Trit_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 0
If Trit(Index) < 2 Then Trit(Index) = 2
Case 1, 2
If Trit(Index) < 5 Then Trit(Index) = 5
Case 3
If Trit(Index) < 1 Then Trit(Index) = 1
End Select

DelayS = Trit(0): DelayT = Trit(1): DelayG = Trit(2): DelayC = Trit(3)
SaveSetting App.EXEName, "Option", "DelayS", DelayS
SaveSetting App.EXEName, "Option", "DelayT", DelayT
SaveSetting App.EXEName, "Option", "DelayG", DelayG
SaveSetting App.EXEName, "Option", "DelayC", DelayC
End Sub

Private Sub txtIPServer_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8, 46
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtIPServer_Validate(Cancel As Boolean)
If txtPortServer = "" Then txtIPServer = OldIP: Exit Sub
'validazione indirizzo
Dim Tmp2 As Variant
Tmp2 = Split(txtIPServer, ".")
If UBound(Tmp2) <> 3 Then txtIPServer = OldIP: Exit Sub
For X = 0 To 3: If Tmp2(X) > 255 Then Tmp2(X) = 255
Next X
txtIPServer = Tmp2(0) & "." & Tmp2(1) & "." & Tmp2(2) & "." & Tmp2(3)
OldIP = txtIPServer
SaveSetting App.EXEName, "Option", "ServerIP", txtIPServer

End Sub

'Purpose:routine per filtro su numero di porta TCP del dispositivo remoto / routine to filter the remote device port number
Private Sub txtPortServer_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57, 8
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtPortServer_Validate(Cancel As Boolean)
SaveSetting App.EXEName, "Option", "ServerPORT", txtPortServer
End Sub

'Purpose:evento di chiusura socket TCP / event to close TCP socket
Private Sub Winsock1_Close()
StBar.Panels(1).Text = txtStatDis
glConnectOK = False
End Sub

'Purpose:evento di connessione socket TCP / evento to connect TCP socket
Private Sub Winsock1_Connect()
If glSsck Then
StBar.Panels(1).Text = txtStatCon & " " & Winsock1.RemoteHostIP & ":" & Winsock1.RemotePort & " [" & txtSsck & "]"
Else
StBar.Panels(1).Text = txtStatCon & " " & Winsock1.RemoteHostIP & ":" & Winsock1.RemotePort & " [" & txtSck & "]"
End If
glConnectOK = True
End Sub
'Purpose:evento di ricezone dati socket TCP / evento to receive datas from TCP socket
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
If Winsock1.State = sckConnected Then
    Call Winsock1.GetData(Buffer)
    BufferInTCP = BufferInTCP & Buffer
End If
End Sub

'Purpose:evento di errore socket TCP / error event from TCP socket
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal helpFile As String, ByVal helpContext As Long, CancelDisplay As Boolean)
'Call InsLog("Errore TCP/IP " & Description)
glConnectOK = False
End Sub

'Purpose:connession TCP e validazione remote host (attesa ACK OPEN) / TCP connection and remote host validation (waiting for OPEN ACK)
Public Function ConnectTCP(SessionOpen As Boolean) As Boolean
Dim TimerIni As Double

glBusy = True
StBar.Panels(1).Text = txtWaitCon
If Winsock1.State <> sckClosed Then
    Winsock1.Close
End If
BufferInTCP = ""
' tentativo di connessione a remote host / try to connect to a remote host
Winsock1.RemoteHost = txtIPServer
Winsock1.RemotePort = txtPortServer
Winsock1.Connect
' attesa di 3 secondi per la connessione / waiting 3 seconds for connection
TimerIni = Timer
Do
    If Timer < TimerIni Then
        TimerIni = Timer
    End If
    DoEvents
    If glConnectOK = True Then
        ' connessione eseguita / connection done
        Exit Do
    End If
    DoEvents
Loop Until Timer > TimerIni + 3
ConnectTCP = glConnectOK
' attesa ACK OPEN / waiting for OPEN ACK
TimerIni = Timer
Do
    If Timer < TimerIni Then
        TimerIni = Timer
    End If
    DoEvents
    If InStr(BufferInTCP, glcACKTCP) > 0 Then
        ' ricevuto ACK da remote host / ACK received from remote host
        SessionOpen = True
        Exit Do
    End If
    DoEvents
Loop Until Timer > TimerIni + 3

End Function
