VERSION 5.00
Begin VB.Form FrmOpt 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   17
      Top             =   1125
      Width           =   990
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
      TabIndex        =   16
      Top             =   750
      Width           =   1590
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
      Left            =   525
      MaxLength       =   2
      TabIndex        =   11
      Text            =   "1"
      Top             =   3600
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
      Left            =   525
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "5"
      Top             =   3375
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
      Left            =   525
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "5"
      Top             =   3150
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
      Index           =   0
      Left            =   525
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "2"
      Top             =   2925
      Width           =   390
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
      Left            =   1125
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "1"
      Top             =   1725
      Width           =   315
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
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "1"
      Top             =   2025
      Width           =   315
   End
   Begin VB.CommandButton Cok 
      Caption         =   "OK"
      Height          =   315
      Left            =   1125
      TabIndex        =   3
      Top             =   2400
      Width           =   2865
   End
   Begin VB.ComboBox Cmb_Lang 
      Height          =   315
      Left            =   825
      TabIndex        =   2
      Top             =   375
      Width           =   2040
   End
   Begin VB.CommandButton C_ok 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3075
      TabIndex        =   1
      Top             =   375
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4275
      TabIndex        =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   2
      Left            =   150
      Picture         =   "FrmOpt.frx":0000
      Top             =   750
      Width           =   240
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
      TabIndex        =   19
      Top             =   1125
      Width           =   1125
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
      TabIndex        =   18
      Top             =   750
      Width           =   225
   End
   Begin VB.Shape Shape3 
      Height          =   1065
      Left            =   75
      Top             =   2850
      Width           =   4515
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo dopo altro comando (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   975
      TabIndex        =   15
      Top             =   3600
      Width           =   3465
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo inizio programmazione (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   975
      TabIndex        =   14
      Top             =   2925
      Width           =   3465
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo dopo comando gruppo (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   975
      TabIndex        =   13
      Top             =   3375
      Width           =   3465
   End
   Begin VB.Label Lrit 
      BackStyle       =   0  'Transparent
      Caption         =   "Ritardo dopo comando temporizzato (sec.)"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   975
      TabIndex        =   12
      Top             =   3150
      Width           =   3465
   End
   Begin VB.Image ImgIco 
      Height          =   240
      Index           =   7
      Left            =   150
      Picture         =   "FrmOpt.frx":038A
      Top             =   2925
      Width           =   240
   End
   Begin VB.Shape Shape2 
      Height          =   1140
      Left            =   75
      Top             =   1650
      Width           =   4515
   End
   Begin VB.Shape Shape1 
      Height          =   1365
      Left            =   75
      Top             =   150
      Width           =   4515
   End
   Begin VB.Label Lbset 
      BackStyle       =   0  'Transparent
      Caption         =   "Attesa ack/nack (sec.)"
      Height          =   240
      Left            =   1575
      TabIndex        =   7
      Top             =   1725
      Width           =   2865
   End
   Begin VB.Label LbTries 
      BackStyle       =   0  'Transparent
      Caption         =   "numero di tentativi controllo ack (sec.)"
      Height          =   240
      Left            =   1575
      TabIndex        =   6
      Top             =   2025
      Width           =   2790
   End
   Begin VB.Image ImgMap 
      Height          =   480
      Left            =   150
      Picture         =   "FrmOpt.frx":0714
      Top             =   1650
      Width           =   480
   End
End
Attribute VB_Name = "FrmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
