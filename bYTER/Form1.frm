VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00E39F68&
   BorderStyle     =   0  'None
   ClientHeight    =   6435
   ClientLeft      =   -15
   ClientTop       =   -75
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00CBE0E9&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000008&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E1EDF2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   4080
      Width           =   6615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E1EDF2&
      Caption         =   "Vbpatcher help"
      Height          =   300
      Left            =   5520
      MouseIcon       =   "Form1.frx":57E2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Only if you are not familier with these stuff :)"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E1EDF2&
      Caption         =   "Hex converter"
      Height          =   300
      Left            =   5520
      MouseIcon       =   "Form1.frx":5934
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Hex Asci calculator ( Must use when you compile the vb patcher source code )"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E1EDF2&
      Caption         =   "Open"
      Height          =   660
      Left            =   5040
      MouseIcon       =   "Form1.frx":5A86
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":5BD8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "This will open the original and the modified once"
      Top             =   840
      Width           =   615
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E1EDF2&
      ForeColor       =   &H00000000&
      Height          =   1005
      ItemData        =   "Form1.frx":855A
      Left            =   120
      List            =   "Form1.frx":855C
      MultiSelect     =   2  'Extended
      TabIndex        =   25
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E1EDF2&
      Caption         =   "Compare bytes"
      Height          =   300
      Left            =   5520
      MouseIcon       =   "Form1.frx":855E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Compare the differences between the original and the modified files ( In bytes )"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E1EDF2&
      Caption         =   "Write patcher"
      Height          =   300
      Left            =   5520
      MouseIcon       =   "Form1.frx":86B0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "After comparing the differences click here to generate the vb patcher source code"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E1EDF2&
      Caption         =   "Both files information here"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   6615
      Begin VB.Frame Frame4 
         BackColor       =   &H00E1EDF2&
         Caption         =   "Controlls"
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   5040
         TabIndex        =   26
         Top             =   270
         Width           =   1455
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Height          =   1575
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E1EDF2&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   1755
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6450
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E1EDF2&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E1EDF2&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E1EDF2&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1440
            Width           =   1215
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E1EDF2&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1005
            ItemData        =   "Form1.frx":8802
            Left            =   3600
            List            =   "Form1.frx":8804
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E1EDF2&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1005
            ItemData        =   "Form1.frx":8806
            Left            =   2280
            List            =   "Form1.frx":8808
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E1EDF2&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1005
            ItemData        =   "Form1.frx":880A
            Left            =   120
            List            =   "Form1.frx":880C
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset table"
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
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Original byte(s)"
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
            Height          =   195
            Left            =   2280
            TabIndex        =   23
            Top             =   0
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New byte(s)"
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
            Height          =   195
            Left            =   3600
            TabIndex        =   22
            Top             =   0
            Width           =   885
         End
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E1EDF2&
      Caption         =   "&Cracked"
      Height          =   300
      Left            =   5760
      MouseIcon       =   "Form1.frx":880E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to open the modified file"
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E1EDF2&
      Caption         =   "&Original"
      Height          =   300
      Left            =   5760
      MouseIcon       =   "Form1.frx":8960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to open the Original file"
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E1EDF2&
      Caption         =   "Choose the original(old) and the cracked(new) file to compare"
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
      Height          =   1215
      Left            =   345
      TabIndex        =   9
      Top             =   510
      Width           =   6615
      Begin VB.TextBox txtOriginal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1EDF2&
         Enabled         =   0   'False
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
         Height          =   300
         Left            =   840
         TabIndex        =   11
         Top             =   360
         Width           =   3780
      End
      Begin VB.TextBox txtPatched 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1EDF2&
         Enabled         =   0   'False
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
         Height          =   300
         Left            =   840
         TabIndex        =   10
         Top             =   720
         Width           =   3780
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   105
         TabIndex        =   27
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Original:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cracked:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   675
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6180
      MouseIcon       =   "Form1.frx":8AB2
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   75
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      MouseIcon       =   "Form1.frx":8C04
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   34
      Top             =   75
      Width           =   585
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait comparing bytes differences  ...  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C19222&
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   5640
      Width           =   6615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "bYTER VERSION 1.2"
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
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   90
      Width           =   1590
   End
   Begin VB.Label Label13 
      BackColor       =   &H00D8E8EF&
      BackStyle       =   0  'Transparent
      Caption         =   "Coding experience over years bring me to make this utility to make your job more faster ..."
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
      Height          =   240
      Left            =   360
      TabIndex        =   30
      Top             =   6120
      Width           =   6615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7065
      MouseIcon       =   "Form1.frx":8D56
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   425
      Left            =   -390
      MousePointer    =   5  'Size
      Picture         =   "Form1.frx":8EA8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   7320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Visit web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6285
      MouseIcon       =   "Form1.frx":A6F8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5895
      Width           =   690
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Email me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   360
      MouseIcon       =   "Form1.frx":A84A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5850
      Width           =   690
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H00CBE0E9&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E39F68&
      BorderWidth     =   6
      FillColor       =   &H00E1EDF2&
      FillStyle       =   0  'Solid
      Height          =   6465
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   7395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private ResultRegion As Long

Private Const RectXRound As Integer = 28
Private Const RectYRound As Integer = 28


Dim OldX As Integer, OldY As Integer, MoveIt As Boolean

Private Sub Command1_Click()
txtOriginal.Text = ""
Command3.Enabled = True

Close


txtOriginal.Text = Open_File(Me.hwnd)

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HCBE0E9
End Sub

Private Sub Command2_Click()
txtPatched.Text = ""
Command3.Enabled = True
Close

txtPatched.Text = Open_File1(Me.hwnd)
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HCBE0E9
End Sub

Private Sub Command3_Click()
On Error Resume Next

If txtOriginal.Text = txtPatched.Text Or Len(txtOriginal.Text) < 3 Or Len(txtPatched.Text) < 3 Then

MsgBox "File comparison error,  check the entries please" & vbCrLf & "--------------------------------------------------------------------------" & vbCrLf & "Reason1:  Both files are the same size in bytes" & vbCrLf & "Reason2:  Original and cracked entries empty", vbInformation, "Try again please ( Caused the folowing reasons )"

List1.Clear
List2.Clear
List3.Clear
List4.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Exit Sub

End If
' Down os the main problem __

List1.Clear
List2.Clear
List3.Clear
List4.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Label12.Visible = True
MousePointer = 11
ByteCompare txtOriginal.Text, txtPatched.Text, List1, List2, List3
addthebytes

If txtPatched.Text = "" Then

If txtOriginal.Text = "" Then

End If
End If
Label12.Visible = False
MousePointer = 0
If List3.ListCount = 0 Then
Exit Sub
Else
If List3.ListCount > 0 Then

conf = MsgBox("Files has been compared do you want to generate a vb patcher", vbYesNo + vbInformation, "bYTER has checked the differences of the 2 files")
If conf = vbYes Then
manualgenerate

End If
End If
End If
Command3.Enabled = False
Exit Sub

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.BackColor = &HCBE0E9
End Sub

Private Sub Command4_Click()
Form2.Show

End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.BackColor = &HCBE0E9
End Sub

Private Sub Command5_Click()
frmchar.Show
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BackColor = &HCBE0E9

End Sub

Private Sub Command6_Click()
On Error Resume Next
Text4.Text = ""

If List4.ListCount = 0 Then
Text4.Text = ""
MsgBox "Cannot write vb patcher till the time that lists are empty from bytes", vbInformation, "Add original and cracked then Compare both please ..."
Exit Sub

Else

Dim nFiles As Integer
Path = App.Path & "/patcher.ini"
Close #1

    Open App.Path & "/patcher.ini" For Output As #1
    Print #1, "'----------   Cut this line not included ----------"
    Print #1, "Private Sub Command1_Click()"
    Print #1, "On error resume next"
    Print #1, "Open" & " " & """" & txtOriginal.Text & """" & " " & "For  Binary As #1"
    Print #1, List4.List(s)
    For nFiles = 1 To List4.ListCount - 1
    List4.ListIndex = nFiles
    Print #1, List4.Text
    Next nFiles
    Print #1, "close #1"
    Print #1, "End Sub"
    Print #1, "'----------   Cut this line not included ----------"
    Print #1, "' --- Do not forget to  replace the bytes with the hex calulator please ---'"

    
    Close #1
 Label12.Visible = True
 
   Label12.Caption = "Patcher location: " & App.Path & "\patcher.ini"

grabthecode


End If

List4.Enabled = True
Exit Sub


' & vbCrLf & "Do not forget please to calculate the replaced bytes with the hex converter" & Path, vbInformation, "Finished"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.BackColor = &HCBE0E9
End Sub

Private Sub Command7_Click()
On Error Resume Next
Close

Command3.Enabled = True

txtOriginal.Text = ""
txtPatched.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

List1.Clear
List2.Clear
List3.Clear
List4.Clear

txtOriginal.Text = Open_File(Me.hwnd)
txtOriginal.Enabled = True

txtPatched.Text = Open_File1(Me.hwnd)
txtPatched.Enabled = True


If txtOriginal.Text = txtPatched.Text Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
List1.Clear
List2.Clear
List3.Clear

Exit Sub

End If
List2.Clear
List3.Clear
Text2.Text = ""
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
MousePointer = 11

Label12.Visible = True
ByteCompare txtOriginal.Text, txtPatched.Text, List1, List2, List3
addthebytes

Label12.Visible = False
MousePointer = 0
Command3.Enabled = False
Exit Sub
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.BackColor = &HCBE0E9

End Sub



Private Sub Form_Load()
On Error Resume Next
Label12.Visible = False
Form1.Width = 7423

 Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(1, 1, 0, 0), True)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.BackColor = &HD8E8EF
Command2.BackColor = &HD8E8EF
Command7.BackColor = &HD8E8EF
Command5.BackColor = &HD8E8EF
Command6.BackColor = &HD8E8EF
Command4.BackColor = &HD8E8EF
Command3.BackColor = &HD8E8EF
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
        OldX = X
        OldY = Y
        MoveIt = True
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label15.FontBold = False
Label15.ForeColor = vbWhite
Label8.FontBold = False
Label8.ForeColor = vbWhite
Form1.Label7.ForeColor = vbWhite
Form1.Label7.FontBold = False
If MoveIt = True Then
    Form1.Top = Form1.Top + Y - OldY
    Form1.Left = Form1.Left + X - OldX
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveIt = False
End Sub

Private Sub Label10_Click()

Shell "start http://www.geocities.com/vbdotlb", vbHide

End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveIt = True Then
    Form1.Top = Form1.Top + Y - OldY
    Form1.Left = Form1.Left + X - OldX
End If
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveIt = False
End Sub



Private Sub Label14_Click()
frmSplash.Show

End Sub

Private Sub Label15_Click()
Form1.WindowState = vbMinimized

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.FontBold = True
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HD8E8EF

Command2.BackColor = &HD8E8EF
Command7.BackColor = &HD8E8EF
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.BackColor = &HD8E8EF
Command6.BackColor = &HD8E8EF
Command4.BackColor = &HD8E8EF
Command3.BackColor = &HD8E8EF
End Sub

Private Sub Label7_Click()
frmSplash.Show

End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label7.FontBold = True

End Sub

Private Sub Label8_Click()


On Error Resume Next

GotoVal = Me.Height / 20    'Form unload animation

For gointo = 20 To GotoVal
DoEvents
Me.Height = Me.Height - 75
      
If Me.Height <= 10 Then GoTo horiz
    Next gointo
horiz:
Me.Height = 50
GotoVal = Me.Width / 15
For gointo = 20 To GotoVal
DoEvents
    Me.Width = Me.Width - 75
    
        If Me.Width <= 100 Then End
        Next gointo


End



End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.FontBold = True


End Sub

Private Sub Label9_Click()
Shell "start mailto:kegham_d@hotmail.com", vbHide

End Sub

Private Sub List1_Click()
On Error Resume Next

Text1.Text = ""
Text1.Text = List1.Text
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex

End Sub

Private Sub List2_Click()
On Error Resume Next

Text2.Text = ""
Text2.Text = List2.Text
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex

End Sub

Private Sub List3_Click()
On Error Resume Next

Text3.Text = ""
Text3.Text = List3.Text
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex

End Sub

Sub addthebytes()
List4.Clear

For A = 0 To List1.ListCount

List4.AddItem "Put #1, " & "&H" & List1.List(A) & "," & List3.List(A)       'List1.List(a)

If Len(List4.List(A)) = 11 Then
List4.RemoveItem A
End If
Next


'Text4.Text = "Put #1, " & "&H" & (List1.List(a)) & "," & "" & List3.List(0)

End Sub

Sub grabthecode()
On Error Resume Next
If List3.ListCount < 1 Then

Text4.Text = ""
Label12.Visible = False
Exit Sub
Else

Open App.Path & "/patcher.ini" For Input As #1

Do Until EOF(1) 'Till end of file
Dim tmp

Line Input #1, tmp
Text4.Text = Text4.Text & vbCrLf & tmp
Loop
Clipboard.SetText (Text4.Text)
End If

End Sub

Sub manualgenerate()
On Error Resume Next
Text4.Text = ""

If List4.ListCount = 0 Then
Text4.Text = ""
MsgBox "Cannot write vb patcher till the time that lists are empty from bytes", vbInformation, "Compare first please ..."
Exit Sub

Else

Dim nFiles As Integer
Path = App.Path & "/patcher.ini"
Close #1

    Open App.Path & "/patcher.ini" For Output As #1
    Print #1, "'----------   Cut this line not included ----------"
    Print #1, "Private Sub Command1_Click()"
    Print #1, "On error resume next"
    Print #1, "Open" & " " & """" & txtOriginal.Text & """" & " " & "For  Binary As #1"
    Print #1, List4.List(s)
    For nFiles = 1 To List4.ListCount - 1
    List4.ListIndex = nFiles
    Print #1, List4.Text
    Next nFiles
    Print #1, "close #1"
    Print #1, "End Sub"
    Print #1, "'----------   Cut this line not included ----------"
    Print #1, "' --- Do not forget to  replace the bytes with the hex calulator please ---'"

    
    Close #1
 Label12.Visible = True
 
   Label12.Caption = "Patcher location: " & App.Path & "\patcher.ini"

grabthecode


End If

List4.Enabled = True
Exit Sub


End Sub

'------ Form shaping function down

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim Corraction As Integer
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim i As Integer
    
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    
    For i = shpBorder.LBound To shpBorder.UBound
        Select Case shpBorder(i).Shape
            Case 0: 'rectangle & square
                ObjectRegion = CreateRectRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
            Case 1: 'circle
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateRectRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case 4:  'round square
                ObjectRegion = CreateRoundRectRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                        RectXRound, RectYRound)
            Case 5: 'round square
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateRoundRectRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY, _
                            RectXRound, RectYRound)
                End If
            Case 3: 'circle
                If shpBorder(i).Width > shpBorder(i).Height Then
                    Corraction = (shpBorder(i).Width - shpBorder(i).Height) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpBorder(i).Left + Corraction) / Screen.TwipsPerPixelX + OffsetX, _
                            shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left - Corraction + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                Else
                    Corraction = (shpBorder(i).Height - shpBorder(i).Width) / 2
                        
                    ObjectRegion = CreateEllipticRgn( _
                            (shpBorder(i).Left) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top + Corraction) / Screen.TwipsPerPixelY + OffsetY, _
                            (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                            (shpBorder(i).Top - Corraction + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
                End If
            Case Else:  'oval
                shpBorder(i).Shape = 2
                ObjectRegion = CreateEllipticRgn( _
                        shpBorder(i).Left / Screen.TwipsPerPixelX + OffsetX, _
                        shpBorder(i).Top / Screen.TwipsPerPixelY + OffsetY, _
                        (shpBorder(i).Left + shpBorder(i).Width) / Screen.TwipsPerPixelX + OffsetX, _
                        (shpBorder(i).Top + shpBorder(i).Height) / Screen.TwipsPerPixelY + OffsetY)
        End Select
        nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
        nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
        DeleteObject ObjectRegion
    Next i
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function



