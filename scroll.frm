VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   Icon            =   "scroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "500"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   " This is a scrolling taskbar. Can't you see it scroll? moo to joo. mo moo moo moo moo moo taking up space *** "
      Top             =   120
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3480
      Top             =   600
   End
   Begin VB.Label Label1 
      Caption         =   "© David Crosbie xb@start.com.au"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'code copyright David Crosbie 2001-2002
'
'to open url
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' variables
Dim TitString
Dim TitLetter
Private Sub Command1_Click()
'changes the speed and text of titlebar
TitString = Text1
Timer1.Interval = Text2
End Sub

Private Sub Form_Load()
'sets the default string t obe displayed.
TitString = " This is a scrolling taskbar. Can't you see it scroll? moo to joo. mo moo moo moo moo moo taking up space *** "
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'hi-lights button
Label1.ForeColor = vbBlack
End Sub

Private Sub Label1_Click()
'opens my website
ShellExecute hWnd, "open", "http://all.at/dco", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
'scrolls the text
TitLetter = Left(TitString, 1)
TitString = Right(TitString, Len(TitString) - 1)
TitString = TitString & TitLetter
Me.Caption = TitString
End Sub
