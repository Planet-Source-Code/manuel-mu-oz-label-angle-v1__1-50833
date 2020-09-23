VERSION 5.00
Begin VB.Form frmVote 
   BackColor       =   &H00D1BDB6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vote For This Project at PSC"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdVote 
      Caption         =   "Vote"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton chkVote 
      BackColor       =   &H00808080&
      Caption         =   "Exellent"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton chkVote 
      BackColor       =   &H00808080&
      Caption         =   "Good"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.OptionButton chkVote 
      BackColor       =   &H00808080&
      Caption         =   "Average"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.OptionButton chkVote 
      BackColor       =   &H00808080&
      Caption         =   "Below Average"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton chkVote 
      BackColor       =   &H00808080&
      Caption         =   "Poor"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   120
      Top             =   615
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "What do you think about this code? "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CodeID As Long = "50833"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdVote_Click()
    Dim Vote As Long, i As Integer
    For i = 1 To 5
      If chkVote(i).Value Then Vote = i
    Next
    
    ShellExecute hWnd, "Open", "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?optCodeRatingValue=" & Vote & "&txtCodeId=" & CodeID & "&lngWId=1", 0&, "C:\", 1
    MsgBox "Thanks for voting in this application.", vbInformation
    Unload Me
    
End Sub
