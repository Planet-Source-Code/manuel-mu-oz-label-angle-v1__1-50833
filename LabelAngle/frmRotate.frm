VERSION 5.00
Begin VB.Form FrmRotate 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label Angle v1.0"
   ClientHeight    =   5355
   ClientLeft      =   3015
   ClientTop       =   1680
   ClientWidth     =   7455
   ForeColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Vote on PSC"
      Height          =   405
      Left            =   5295
      TabIndex        =   10
      Top             =   4770
      Width           =   1920
   End
   Begin Project1.LabelAngle LabelAngle2 
      Height          =   795
      Left            =   2895
      TabIndex        =   2
      Top             =   2400
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1402
      Angle           =   186
      Caption         =   "LabelAngle  v 1.0"
      ColorStyle      =   0
      BackColor       =   65280
      Caption         =   "LabelAngle  v 1.0"
      ForeColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PosX            =   3506
      PosY            =   433
   End
   Begin Project1.LabelAngle LabelAngle5 
      Height          =   1170
      Left            =   1095
      TabIndex        =   8
      Top             =   1485
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   2064
      Angle           =   45
      Caption         =   "AutoSize"
      ColorStyle      =   1
      BackColor       =   16744576
      Caption         =   "AutoSize"
      ForeColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      PosX            =   0
      PosY            =   870
   End
   Begin Project1.LabelAngle LabelAngle4 
      Height          =   4410
      Left            =   6720
      TabIndex        =   4
      Top             =   195
      Width           =   540
      _ExtentX        =   370
      _ExtentY        =   1588
      Caption         =   "(c) Lito 2.003     A Coruña -  Spain"
      AutoSize        =   0   'False
      ColorStyle      =   0
      BackColor       =   8438015
      Caption         =   "(c) Lito 2.003     A Coruña -  Spain"
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      PosX            =   150
      PosY            =   3400
   End
   Begin Project1.LabelAngle LabelAngle1 
      Height          =   3675
      Left            =   180
      TabIndex        =   1
      Top             =   465
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   6482
      Caption         =   "LabelAngle v1.0"
      Caption         =   "LabelAngle v1.0"
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      PosX            =   0
      PosY            =   3675
   End
   Begin Project1.LabelAngle LabelAngle3 
      Height          =   1350
      Left            =   4305
      TabIndex        =   3
      Top             =   375
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   2381
      Angle           =   20
      Caption         =   "LabelAngle"
      Caption         =   "LabelAngle"
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ottawa"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PosX            =   0
      PosY            =   795
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmRotate.frx":0000
      Height          =   660
      Left            =   2895
      TabIndex        =   9
      Top             =   1845
      Width           =   3450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label normal"
      BeginProperty Font 
         Name            =   "Ottawa"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   585
      Left            =   1140
      TabIndex        =   0
      Top             =   825
      Width           =   2760
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   $"frmRotate.frx":0092
      ForeColor       =   &H00000080&
      Height          =   900
      Left            =   1005
      TabIndex        =   7
      Top             =   165
      Width           =   5565
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B3C49F&
      Caption         =   "This a stable version that needs some improvements:     *Borders Style.  *Auto adjust text,   etc."
      ForeColor       =   &H00008000&
      Height          =   585
      Left            =   1005
      TabIndex        =   5
      Top             =   3990
      Width           =   5565
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   $"frmRotate.frx":01BF
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   1005
      TabIndex        =   6
      Top             =   2985
      Width           =   5565
   End
End
Attribute VB_Name = "FrmRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
frmVote.Show 1
End Sub

