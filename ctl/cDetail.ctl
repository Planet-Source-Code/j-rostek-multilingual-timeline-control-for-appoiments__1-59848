VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl cDetail 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11220
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5100
   ScaleWidth      =   11220
   Begin VB.PictureBox pic 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   8940
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   10620
      Picture         =   "cDetail.ctx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   10320
      Picture         =   "cDetail.ctx":01CE
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Frame 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      Begin VB.PictureBox OpenClose 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         Picture         =   "cDetail.ctx":039C
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   1
         Top             =   60
         Width           =   165
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   2
         Top             =   45
         Width           =   10095
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   1200
      TabIndex        =   6
      Top             =   1380
      Visible         =   0   'False
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Icon"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Text"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Group"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "EndDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Colour"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "BeginTime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "EndTime"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "cDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'##################################################################################################
'#### Date Variables ##############################################################################
'##################################################################################################
Dim DisplayDays As Long
Dim StartDate As Date
Dim EndDate As Date
Dim isOpen As Boolean
'##################################################################################################
'#### Event Open or Close the View ################################################################
'##################################################################################################
Public Event OpenClose()
'##################################################################################################
'#### Hover over a Date Event #####################################################################
'##################################################################################################
Public Event HooverValue(Value As Date)
Dim DaysAtCurrentMonth As Long
Dim myLastHoverDate As Date
'##################################################################################################
'#### Event for Doubleclick an Emty Date###########################################################
'##################################################################################################
Public Event DblClick(Value As Date)
'##################################################################################################
'#### Event for Doubleclick an Filled Date ########################################################
'##################################################################################################
Public Event DblClickEntry(cStartDate As Date, cEndDate As Date, cText As String, cGroup As String)
'##################################################################################################
'#### Event for Drag and Drop Date ################################################################
'##################################################################################################
Public Event DragDate(oldStartDate As Date, oldEndDate As Date, NewStartDate As Date, NewEndDate As Date, cText As String)
'##################################################################################################
'#### For Drag and Drop Dates #####################################################################
'##################################################################################################
Private m_Dragging As Boolean
Private m_StartX As Single
Private m_StartY As Single

Private Sub DragMouseDown(ByVal ctl As Control, ByVal X As Single, ByVal Y As Single)
'##################################################################################################
'#### Start Date Drag #############################################################################
'##################################################################################################
  m_Dragging = True
  m_StartX = X
  m_StartY = Y
End Sub
Private Sub DragMouseMove(ByVal ctl As Control, ByVal X As Single, ByVal Y As Single)
'##################################################################################################
'#### Continue for Drag a Date ####################################################################
'##################################################################################################
  Dim new_x As Single
  Dim new_y As Single
  If Not m_Dragging Then Exit Sub
  new_x = ctl.Left + (X - m_StartX)
  If new_x < 0 Then
      new_x = 0
  ElseIf new_x > ScaleWidth - ctl.Width Then
      new_x = ScaleWidth - ctl.Width
  End If
  new_x = (Int((new_x / DaysAtCurrentMonth)) * DaysAtCurrentMonth) + 30
  ctl.Move new_x
End Sub
Private Sub DragMouseUp(ByVal ctl As Control, ByVal X As Single, ByVal Y As Single)
'##################################################################################################
'#### End the Date Drag ###########################################################################
'##################################################################################################
    If m_Dragging Then m_Dragging = False
End Sub
Public Property Get GetEntryName() As String
'##################################################################################################
'#### Get the Entry Global Name ###################################################################
'##################################################################################################
  GetEntryName = lbl.Caption
End Property

Private Sub Frame_DblClick()
OpenClose_Click
End Sub

Private Sub lbl_DblClick()
OpenClose_Click
End Sub

Private Sub OpenClose_Click()
'##################################################################################################
'#### Displays Date Entrys or not #################################################################
'##################################################################################################
If isOpen = False Then
  myMaxHeight = 630
  OpenClose.Picture = Picture2.Picture
  DisplayControl StartDate, EndDate, False
  For Each Control In UserControl
    If Control.Name = "pic" Then
      If Control.Visible = True Then myMaxHeight = Control.Top + Control.Height + 100
    End If
  Next
  UserControl.Height = myMaxHeight
  For Each Control In UserControl
    If Control.Name = "pic" Then Control.Visible = False
  Next
  DisplayControl StartDate, EndDate, True
  For n = 1 To ListView1.ListItems.Count
    AddEntry CDate(ListView1.ListItems(n).Text), CDate(ListView1.ListItems(n).SubItems(5)), pic(ListView1.ListItems(n).SubItems(3)), ListView1.ListItems(n).SubItems(2)
  Next
  DoEvents
  DisplayControl StartDate, EndDate, False
  For Each Control In UserControl
    If Control.Name = "pic" Then
      If Control.Index > 0 Then Control.Visible = True
    End If
  Next
  isOpen = True
Else
  OpenClose.Picture = Picture1.Picture
  UserControl.Height = Frame.Height
  isOpen = False
End If
Me.CRefresh
RaiseEvent OpenClose
End Sub

Private Sub pic_DblClick(Index As Integer)
  picStart = Format(Left(pic(Index).Tag, InStr(pic(Index).Tag, "|") - 1), "Short Date")
  picEnd = Format(Right(pic(Index).Tag, Len(Left(pic(Index).Tag, InStr(pic(Index).Tag, "|"))) - 1), "Short Date")
  For g = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(g).SubItems(4) = lbl.Caption And Format(ListView1.ListItems(g).Text, "Short Date") = Format(CDate(picStart), "Short Date") And Format(ListView1.ListItems(g).SubItems(5), "Short Date") = Format(CDate(picEnd), "Short Date") And ListView1.ListItems(g).SubItems(3) = Index Then
    RaiseEvent DblClickEntry(CDate(picStart), CDate(picEnd), ListView1.ListItems(g).SubItems(2), lbl.Caption)
    Exit For
    End If
  Next
End Sub

Private Sub pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'##################################################################################################
'#### A Date Entry will now moved if the Date full showing ########################################
'##################################################################################################
If Button = 1 And pic(Index).Left > 0 And pic(Index).Left + pic(Index).Width < UserControl.Width Then
  DragMouseDown pic(Index), X, Y
End If
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'##################################################################################################
'#### A Date Entry is now moving ##################################################################
'##################################################################################################
DragMouseMove pic(Index), X, Y
End Sub

Private Sub pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'##################################################################################################
'#### A Date Entry has moved to an other Date #####################################################
'##################################################################################################
  DragMouseUp pic(Index), X, Y
  picStart = Left(pic(Index).Tag, InStr(pic(Index).Tag, "|") - 1)
  picEnd = Right(pic(Index).Tag, Len(Left(pic(Index).Tag, InStr(pic(Index).Tag, "|"))) - 1)
  theNewLeftDate = CalculateValue(pic(Index).Left, pic(Index).Top)
  theNewStartDate = CDate(theNewLeftDate) '+ 1
  DaysForTheEntry = DateDiff("d", CDate(picStart), CDate(picEnd))
  theNewEndDate = DateAdd("d", CDbl(DaysForTheEntry), theNewStartDate)
  If Format(CDate(picStart), "Short Date") = Format(CDate(theNewStartDate), "Short Date") Then Exit Sub
  For g = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(g).SubItems(4) = lbl.Caption And Format(ListView1.ListItems(g).Text, "Short Date") = Format(CDate(picStart), "Short Date") And Format(ListView1.ListItems(g).SubItems(5), "Short Date") = Format(CDate(picEnd), "Short Date") And ListView1.ListItems(g).SubItems(3) = Index Then
      ListView1.ListItems(g).Text = CDate(theNewStartDate)
      ListView1.ListItems(g).SubItems(5) = CDate(theNewEndDate)
      pic(Index).Tag = Format(CDate(theNewStartDate), "Short Date") & "|" & Format(CDate(theNewEndDate), "Short Date")

      ' Fill with Entrys
      UserControl.Height = myMaxHeight
      For Each Control In UserControl
        If Control.Name = "pic" Then Control.Visible = False
      Next
      For n = 1 To ListView1.ListItems.Count
        AddEntry CDate(ListView1.ListItems(n).Text), CDate(ListView1.ListItems(n).SubItems(5)), pic(ListView1.ListItems(n).SubItems(3)), ListView1.ListItems(n).SubItems(2)
      Next
      DisplayControl StartDate, EndDate, False
      For Each Control In UserControl
        If Control.Name = "pic" Then
          If Control.Index > 0 Then Control.Visible = True
        End If
      Next

      RaiseEvent DragDate(CDate(picStart), CDate(picEnd), CDate(theNewStartDate), CDate(theNewEndDate), ListView1.ListItems(g).SubItems(2))
    End If
  Next
  
End Sub

Private Sub UserControl_DblClick()
'##################################################################################################
'#### Event for Doubleclick a Empty Date ##########################################################
'##################################################################################################
RaiseEvent DblClick(myLastHoverDate)
End Sub

Private Sub UserControl_Initialize()
'##################################################################################################
'#### Draw a Gradient to the Explorerbar ##########################################################
'##################################################################################################
Dim Gradiente As New clsGradient
Gradiente.Angle = -90
Gradiente.Color1 = &H8000000F
Gradiente.Color2 = &HFFFFFF
Gradiente.Draw Frame
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
myLastHoverDate = CalculateValue(X, Y)
RaiseEvent HooverValue(myLastHoverDate)
End Sub
Private Function CalculateValue(X As Single, Y As Single) As Long
'##################################################################################################
'#### Get the Date under the MousePosition ########################################################
'##################################################################################################
On Error Resume Next
    Dim myValue As Long
    mvarStartValue = 1
    myValue = Int(X / (DaysAtCurrentMonth / 10))
    myValue = myValue + mvarStartValue * 10
    myTemp = Replace(myValue, " ", "")
    myValue = Left(myTemp, Len(myTemp) - 1)
    CalculateValue = CDate(DateAdd("d", myValue - 2, StartDate))
End Function
Private Sub UserControl_Resize()
CRefresh
End Sub

Sub CRefresh()
'##################################################################################################
'#### Set Controls Align and more #################################################################
'##################################################################################################
lbl.Move lbl.Left, lbl.Top, UserControl.Width
Frame.Move -100, 0, UserControl.Width + 200
Dim Gradiente As New clsGradient
Gradiente.Angle = -90
Gradiente.Color1 = &H8000000F
Gradiente.Color2 = &HFFFFFF
Gradiente.Draw Frame
Set Gradiente = Nothing
End Sub
Public Sub DisplayControl(cStartDate As Date, cEndDate As Date, IncludeWeeks As Boolean)
'##################################################################################################
'#### Visble the Pictures if the Date not Display #################################################
'##################################################################################################
For g = 1 To pic.Count - 1
  pic(g).Visible = False
Next
'##################################################################################################
'#### First Picture Top ###########################################################################
'##################################################################################################
pic(0).Top = 660
'##################################################################################################
'#### Clear the Control for Drawing lines #########################################################
'##################################################################################################
Cls
'##################################################################################################
'#### Setting up Start and End Date################################################################
'##################################################################################################
StartDate = cStartDate
EndDate = cEndDate
If DateDiff("d", cStartDate, EndDate) < 0 Then
  EndDate = DateAdd("d", 1, cEndDate)
End If
'##################################################################################################
'#### Display more as one Day #####################################################################
'##################################################################################################
If DateDiff("d", cStartDate, EndDate) > 0 Then
  strdate = cStartDate
  theDays = DateDiff("d", strdate, EndDate)
  CurrentDay = Day(CDate(strdate))
  DaysAtCurrentMonth = UserControl.Width / (theDays + 1)
  For g = 0 To theDays
    CurrentX = (g * DaysAtCurrentMonth) '- 30
    Line (g * DaysAtCurrentMonth, 315)-(g * DaysAtCurrentMonth, ScaleHeight)
    CurrentY = 500
    CurrentX = (g * DaysAtCurrentMonth) - 30
    myMax = (g * DaysAtCurrentMonth) - 30
    strdate = DateAdd("d", 1, strdate)
    CurrentDay = Day(CDate(strdate))
  Next
End If

'##################################################################################################
'#### Add Entrys from List ########################################################################
'##################################################################################################
If isOpen = True Then
  For n = 1 To ListView1.ListItems.Count
    AddEntry CDate(ListView1.ListItems(n).Text), CDate(ListView1.ListItems(n).SubItems(5)), pic(ListView1.ListItems(n).SubItems(3)), ListView1.ListItems(n).SubItems(2)
  Next
End If
DisplayDays = theDays
theOldHeight = UserControl.Height
If UserControl.Height <> theOldHeight Then RaiseEvent OpenClose
End Sub
Private Sub AddEntry(cEntryStartDate As Date, cEntryEndDate As Date, cPicture As PictureBox, cText As String)
'################### Displaying Number of Days as Pic #############################################
'##################################################################################################
'#### Add Entrys from List one by one and Ckec her Position########################################
'##################################################################################################
If DateDiff("d", cEntryStartDate, EndDate) < 0 Then Exit Sub
DispDays = DateDiff("d", StartDate, EndDate)
DispDaysLegth = DateDiff("d", cEntryStartDate, cEntryEndDate) + 1
theStart = DateDiff("d", StartDate, cEntryStartDate)
theLegthDate = DateDiff("d", StartDate, cEntryEndDate)
If CLng(theLegthDate) >= 0 Then
  MyTop = 350
  ControlHeight = 615
   For Each Control In UserControl
    If Control.Name = "pic" Then
      If Control.Visible = True Then
        picStart = Left(Control.Tag, InStr(Control.Tag, "|") - 1)
        picEnd = Right(Control.Tag, Len(Left(Control.Tag, InStr(Control.Tag, "|"))) - 1)
        theLegthFromOtherDate = DateDiff("d", CDate(picStart), CDate(picEnd))
        mdate = CDate(picStart)
        For m = 0 To theLegthFromOtherDate
          If Format(CDate(mdate), "Short Date") = cEntryStartDate Or Format(CDate(mdate), "Short Date") = cEntryEndDate Then
            If Control.Top + Control.Height > MyTop Then MyTop = Control.Top + Control.Height
            Exit For
          End If
          mdate = DateAdd("d", 1, CDate(mdate))
        Next
        mdate = CDate(picStart)
        For m = 0 To DispDaysLegth
          If CDate(mdate) = cEntryStartDate Or CDate(mdate) = cEntryEndDate Then
            If Control.Top + Control.Height > MyTop Then MyTop = Control.Top + Control.Height
            Exit For
          End If
          mdate = DateAdd("d", 1, CDate(mdate))
        Next
      End If
    End If
   Next
   cLeft = CLng(theStart) * DaysAtCurrentMonth + 60
   cWidth = (CLng(DispDaysLegth) * DaysAtCurrentMonth) - 60
   If cWidth <= 0 Then cWidth = 10
  cPicture.Move cLeft, MyTop, cWidth
  cPicture.Visible = True
     cPicture.Cls
  cPicture.CurrentX = 250
  cPicture.CurrentY = 0
  cPicture.Print cText
  cPicture.ZOrder 0
  For Each Control In UserControl
    If Control.Name = "pic" Then
      If Control.Visible = True Then
        If Control.Top + Control.Height > MyControlHeight Then MyControlHeight = Control.Top + Control.Height + 50
      End If
    End If
  Next
  UserControl.Height = MyControlHeight
End If
End Sub
Sub SetListview(cTheDate As Date, cPicture As PictureBox, cText As String, cGroup As String, cToDate As Date, cColour As ColorConstants, cBeginTime As String, cEndTime As String)
'##################################################################################################
'#### Add Entrys from List and Load Pictures dynamicly ############################################
'##################################################################################################
lbl.Caption = cGroup
ListView1.ListItems.Add , , cTheDate
ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = cPicture
ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = cText
ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = ListView1.ListItems.Count
ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = cGroup
ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = cToDate
ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = cColour
ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = cBeginTime
ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = cEndTime
Load pic(pic.Count)
pic(pic.Count - 1).BackColor = cColour
pic(pic.Count - 1).Picture = cPicture
pic(pic.Count - 1).Tag = cTheDate & "|" & cToDate
End Sub
