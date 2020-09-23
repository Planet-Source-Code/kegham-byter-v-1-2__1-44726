VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00D8E8EF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8EF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   3720
   End
   Begin VB.Timer Timer3 
      Interval        =   1500
      Left            =   960
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   1920
      Top             =   3720
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "About the coder and must read please ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4665
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "we can do what we want"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":185C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8EF&
      BackStyle       =   0  'Transparent
      Caption         =   "crackme proudley present bYTER v 1.2  file compare process engine and vb patch source code generator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "bYETR is the new age cracking tech vbdotlb® 1998 - 2003"
      ForeColor       =   &H00C19222&
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   4695
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Timer1 = False
Me.Timer2 = False
Me.Timer3 = False
Me.Timer4 = False
Me.Hide
Form1.Show
End Sub

Private Sub Timer1_Timer()
Command1.Caption = "3"
Timer2.Enabled = True
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
Command1.Caption = "2"
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Command1.Caption = "1"
Timer3.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
Command1.Caption = "O k e y"
Timer4.Enabled = False
Command1.Enabled = True

End Sub
