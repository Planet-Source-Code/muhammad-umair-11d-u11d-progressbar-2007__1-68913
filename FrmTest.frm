VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "U11D PROGRESSBAR™ BY UMAIR_11D®© CORPORATION."
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   3  'Windows Default
   Begin U11DProgressBar.ProgressBar ProgressBar11 
      Height          =   870
      Left            =   1185
      TabIndex        =   16
      Top             =   4080
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   1535
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   255
      Text            =   "CUSTOME FONT AND SIZE AND TEXT EFFECT"
      TextEffectColor =   65280
      TextEffect      =   4
   End
   Begin U11DProgressBar.ProgressBar ProgressBar10 
      Height          =   300
      Left            =   2175
      TabIndex        =   15
      Top             =   3750
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   10
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 10 CUSTOME"
      TextEffectColor =   16777215
      TextEffect      =   5
      PBSCustomeColor2=   16777152
      PBSCustomeColor1=   12632319
   End
   Begin U11DProgressBar.ProgressBar ProgressBar9 
      Height          =   300
      Left            =   2055
      TabIndex        =   14
      Top             =   3435
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   9
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 9"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar8 
      Height          =   300
      Left            =   1935
      TabIndex        =   13
      Top             =   3105
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   8
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 8"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar7 
      Height          =   300
      Left            =   1335
      TabIndex        =   12
      Top             =   1470
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   3
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 3"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar6 
      Height          =   300
      Left            =   1455
      TabIndex        =   11
      Top             =   1785
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   4
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 4"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar5 
      Height          =   300
      Left            =   1575
      TabIndex        =   10
      Top             =   2115
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   5
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 5"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar4 
      Height          =   300
      Left            =   1695
      TabIndex        =   9
      Top             =   2445
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   6
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 6"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar3 
      Height          =   300
      Left            =   1815
      TabIndex        =   8
      Top             =   2775
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   7
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 7"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar2 
      Height          =   300
      Left            =   1215
      TabIndex        =   7
      Top             =   1155
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      Theme           =   2
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 2"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   1095
      TabIndex        =   6
      Top             =   840
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   529
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "THEME 1"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   5085
      Width           =   1215
   End
   Begin U11DProgressBar.ProgressBar P 
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   529
      Max             =   500
      Value           =   250
      Enabled         =   0   'False
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "DISBLED"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin VB.Timer T 
      Interval        =   1
      Left            =   495
      Top             =   5595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Chage Theme"
      Height          =   495
      Left            =   105
      TabIndex        =   0
      Top             =   5085
      Width           =   1215
   End
   Begin U11DProgressBar.ProgressBar P 
      Height          =   300
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   450
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   529
      Max             =   500
      Value           =   250
      TextStyle       =   3
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "ENABLED"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar P 
      Height          =   3960
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   855
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   6985
      Orientations    =   2
      Max             =   500
      Value           =   250
      Enabled         =   0   'False
      TextStyle       =   2
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "U11D ProgressBar"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin U11DProgressBar.ProgressBar P 
      Height          =   3960
      Index           =   3
      Left            =   525
      TabIndex        =   4
      Top             =   855
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   6985
      Orientations    =   2
      Max             =   500
      Value           =   250
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "U11D ProgressBar"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim THM As Long
Private Sub Command1_Click()
If THM >= 10 Then
THM = 1
Else
THM = THM + 1
End If
P(0).Theme = THM
P(1).Theme = THM
P(2).Theme = THM
P(3).Theme = THM
End Sub

Private Sub Command2_Click()
P(0).About
End Sub

Private Sub Form_Load()
THM = 1
P(0).Theme = THM
P(1).Theme = THM
P(2).Theme = THM
P(3).Theme = THM
End Sub

Private Sub T_Timer()
If P(0).Value >= P(0).Max Then
P(0).Value = 0
Else
P(0).Value = P(0).Value + 1
End If
P(1).Value = P(0).Value
P(2).Value = P(0).Value
P(3).Value = P(0).Value
End Sub
