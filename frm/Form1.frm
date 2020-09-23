VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   11670
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      Height          =   750
      Left            =   11760
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   840
      TabIndex        =   7
      Top             =   4020
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture6 
      Height          =   615
      Left            =   11820
      Picture         =   "Form1.frx":1E72
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture5 
      Height          =   315
      Left            =   11760
      Picture         =   "Form1.frx":24B4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture4 
      Height          =   315
      Left            =   11760
      Picture         =   "Form1.frx":27F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture3 
      Height          =   315
      Left            =   11760
      Picture         =   "Form1.frx":2B38
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture2 
      Height          =   315
      Left            =   11760
      Picture         =   "Form1.frx":2E7A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   315
      Left            =   11760
      Picture         =   "Form1.frx":31BC
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin Projekt1.cJournal cJournal1 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10455
      _ExtentX        =   17595
      _ExtentY        =   4577
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cJournal1_DblClickEntry(cStartDate As Date, cEndDate As Date, cText As String, cGroup As String)
MsgBox "You DoubleClick the Entry form " & cStartDate & " to " & cEndDate & " " & cText & " in Group " & cGroup
frmDayView.cDayview1.DrawADate 30, "06:30:00", False, "07:00:00", "16:00:00", cStartDate
frmDayView.Show 1

End Sub

Private Sub cJournal1_DragDropDate(oldStartDate As Date, oldEndDate As Date, NewStartDate As Date, NewEndDate As Date, cText As String)
MsgBox "The Date has moved from:" & oldStartDate & vbCrLf & "The Date has moved to:" & NewStartDate
End Sub

Private Sub cJournal1_HooverValue(Value As Date)
Me.Caption = Value
End Sub


Private Sub Form_Load()

cJournal1.AddUserEvent CDate(Now), Picture2, "test", "JR", CDate(Now + 1), vbYellow, "08:00:00", "10:30:00"
cJournal1.AddUserEvent CDate(Now + 1), Picture2, "test", "JR", CDate(Now + 2), vbRed, "08:00:00", "10:30:00"
cJournal1.AddUserEvent CDate(Now + 5), Picture1, "test", "JR", CDate(Now + 10), &H80000001, "08:00:00", "10:30:00"
cJournal1.AddUserEvent CDate(Now + 15), Picture5, "test", "JR", CDate(Now + 15), vbBlue, "08:00:00", "10:30:00"

cJournal1.AddUserEvent CDate(Now), Picture1, "testsd", "HP", CDate(Now), vbCyan, "08:00:00", "10:30:00"
cJournal1.AddUserEvent CDate(Now + 1), Picture4, "testsd", "HP", CDate(Now + 8), vbGreen, "08:00:00", "10:30:00"
cJournal1.AddUserEvent CDate(Now + 10), Picture4, "testsd", "HP", CDate(Now + 15), vbYellow, "08:00:00", "10:30:00"

cJournal1.cInit CDate(Now), CDate(Now + 30), True
'cDayView1.AddUserEvent CDate(Now), Picture2, "test", "JR", CDate(Now + 1), &H80000001, "08:25:00", "10:30:00"
'Me.cDayView1.DrawADate 15, "06:45:00", True
End Sub

Private Sub Form_Resize()
cJournal1.Move 0, 0, Me.Width - 100, Me.Height - 400
End Sub
