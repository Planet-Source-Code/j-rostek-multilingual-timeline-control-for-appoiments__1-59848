VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl cJournal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5730
   ScaleWidth      =   10215
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   900
      TabIndex        =   0
      Top             =   2940
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
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   9660
      SmallChange     =   100
      TabIndex        =   8
      Top             =   2460
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRight 
      Height          =   495
      Left            =   540
      Picture         =   "cJournal.ctx":0000
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   360
      Width           =   435
   End
   Begin VB.CommandButton cmdLeft 
      Height          =   495
      Left            =   0
      Picture         =   "cJournal.ctx":01CE
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   360
      Width           =   435
   End
   Begin VB.CommandButton pictMinus 
      Height          =   255
      Left            =   60
      Picture         =   "cJournal.ctx":039C
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton pictPlus 
      Height          =   255
      Left            =   420
      Picture         =   "cJournal.ctx":056A
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   9915
      TabIndex        =   2
      Top             =   960
      Width           =   9915
      Begin Projekt1.cDetail cDetail1 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   "JR"
         Top             =   0
         Visible         =   0   'False
         Width           =   7395
         _extentx        =   13044
         _extenty        =   556
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2400
      ScaleHeight     =   375
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "cJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'##################################################################################################
'#### Hover over a Date Event #####################################################################
'##################################################################################################
Public Event HooverValue(Value As Date)
Public Event DoubleClick(Value As Date)
'##################################################################################################
'#### Day Width (Display Width) ###################################################################
'##################################################################################################
Dim DaysAtCurrentMonth As Long
'##################################################################################################
'#### Systemsettings for Dates get from the System not from the Program ###########################
'##################################################################################################
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SDATE = &H1D               '  date separator
Private Const LOCALE_SABBREVDAYNAME1 = &H31     '  abbreviated name for Monday
Private Const LOCALE_SABBREVDAYNAME2 = &H32     '  abbreviated name for Tuesday
Private Const LOCALE_SABBREVDAYNAME3 = &H33     '  abbreviated name for Wednesday
Private Const LOCALE_SABBREVDAYNAME4 = &H34     '  abbreviated name for Thursday
Private Const LOCALE_SABBREVDAYNAME5 = &H35     '  abbreviated name for Friday
Private Const LOCALE_SABBREVDAYNAME6 = &H36     '  abbreviated name for Saturday
Private Const LOCALE_SABBREVDAYNAME7 = &H37     '  abbreviated name for Sunday
Private Const LOCALE_SMONTHNAME1 = &H38         '  long name for January
Private Const LOCALE_SMONTHNAME2 = &H39         '  long name for February
Private Const LOCALE_SMONTHNAME3 = &H3A         '  long name for March
Private Const LOCALE_SMONTHNAME4 = &H3B         '  long name for April
Private Const LOCALE_SMONTHNAME5 = &H3C         '  long name for May
Private Const LOCALE_SMONTHNAME6 = &H3D         '  long name for June
Private Const LOCALE_SMONTHNAME7 = &H3E         '  long name for July
Private Const LOCALE_SMONTHNAME8 = &H3F         '  long name for August
Private Const LOCALE_SMONTHNAME9 = &H40         '  long name for September
Private Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Private Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Private Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'##################################################################################################
'#### Event for Doubleclick an Filled Date ########################################################
'##################################################################################################
Public Event DblClickEntry(cStartDate As Date, cEndDate As Date, cText As String, cGroup As String)
'##################################################################################################
'#### Event Drag and Drop a Date ##################################################################
'##################################################################################################
Public Event DragDropDate(oldStartDate As Date, oldEndDate As Date, NewStartDate As Date, NewEndDate As Date, cText As String)
'##################################################################################################
'#### Display a really flat button ################################################################
'##################################################################################################
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private WindowRect As Rect
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'##################################################################################################
'#### Date Variables ##############################################################################
'##################################################################################################
Dim DisplayDays As Long
Dim StartDate As Date
Dim EndDate As Date

Sub AddUserEvent(cTheDate As Date, cPicture As PictureBox, cText As String, cGroup As String, cToDate As Date, cColour As ColorConstants, cBeginTime As String, cEndTime As String)
'##################################################################################################
'#### Add UserEvent to the Control ################################################################
'##################################################################################################
ListView1.ListItems.Add , , cTheDate
ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = cPicture
ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = cText
ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = ListView1.ListItems.Count
ListView1.ListItems(ListView1.ListItems.Count).SubItems(4) = cGroup
ListView1.ListItems(ListView1.ListItems.Count).SubItems(5) = cToDate
ListView1.ListItems(ListView1.ListItems.Count).SubItems(6) = cColour
ListView1.ListItems(ListView1.ListItems.Count).SubItems(7) = cBeginTime
ListView1.ListItems(ListView1.ListItems.Count).SubItems(8) = cEndTime
Dim cGroupExists As Boolean
For Each Control In UserControl
  If Control.Name = "cDetail1" Then
    If Control.Tag = cGroup Then cGroupExists = True
  End If
Next
If cGroupExists = False Then
  Load cDetail1(cDetail1.Count)
  cDetail1(cDetail1.Count - 1).Tag = cGroup
  cDetail1(cDetail1.Count - 1).Top = cDetail1(cDetail1.Count - 1).Top + cDetail1(cDetail1.Count - 1).Height
  cDetail1(cDetail1.Count - 1).Visible = True
End If
For Each Control In UserControl
  If Control.Name = "cDetail1" Then
    If Control.Tag = cGroup Then
      Control.SetListview cTheDate, cPicture, cText, cGroup, cToDate, cColour, cBeginTime, cEndTime
    End If
  End If
Next
cDetail1_OpenClose cDetail1.Count - 1
End Sub
Sub cInit(cStartDate As Date, cEndDate As Date, IncludeWeeks As Boolean)
DisplayControl CDate(cStartDate), CDate(cEndDate), IncludeWeeks
End Sub
Public Function fConvKW(ByVal ZuKonvertDat As Date) As String
'##################################################################################################
'#### Get the Correct Calaendar Week ##############################################################
'##################################################################################################
  If Weekday(ZuKonvertDat) = vbMonday Then ZuKonvertDat = ZuKonvertDat + 1
  fConvKW = DatePart("ww", ZuKonvertDat, vbMonday, vbFirstJan1)
  fConvKW = fConvKW
End Function


Private Sub DisplayControl(cStartDate As Date, cEndDate As Date, IncludeWeeks As Boolean)

For g = 1 To pic.Count - 1
  pic(g).Visible = False
Next
pictMinus.Top = 0
pictMinus.Left = 0
pictPlus.Top = 0
pictPlus.Left = UserControl.Width - pictPlus.Width


Cls

StartDate = cStartDate
EndDate = cEndDate
If DateDiff("d", cStartDate, EndDate) < 0 Then
  EndDate = DateAdd("d", 1, cEndDate)
End If
strdate = cStartDate
theDays = DateDiff("d", strdate, EndDate)
CurrentDay = 1
DaysAtCurrentMonth = UserControl.Width / (theDays + 1)
theWeek = fConvKW(CDate(strdate)) - 1

For g = 0 To theDays
  
If theMonth <> Month(strdate) Then
      CurrentY = 0
      CurrentX = ((g + 1) * DaysAtCurrentMonth) + 30
      Print GetLongMonthName(Month(strdate)) & " " & Year(strdate)
      Line ((g) * DaysAtCurrentMonth, 0)-((g) * DaysAtCurrentMonth, 250)
      theMonth = Month(strdate)
End If
  
  If IncludeWeeks = True And DisplayDays <= 40 Then
    If theWeek <> fConvKW(CDate(strdate)) - 1 Then
      CurrentY = 250
      CurrentX = ((g) * DaysAtCurrentMonth) + 30
      Print "KW " & fConvKW(CDate(strdate)) - 1
      theWeek = fConvKW(CDate(strdate)) - 1
      Line ((g) * DaysAtCurrentMonth, 300)-((g) * DaysAtCurrentMonth, 900)
    End If
  End If
  CurrentX = (g * DaysAtCurrentMonth) '- 30
  Line (g * DaysAtCurrentMonth, 500)-(g * DaysAtCurrentMonth, 900)
  CurrentY = 500
  CurrentX = (g * DaysAtCurrentMonth) - 30
  myMax = (g * DaysAtCurrentMonth) - 30
  If DisplayDays <= 45 Then Print Day(CDate(strdate))

      CurrentY = 700
      CurrentX = (g * DaysAtCurrentMonth) + 30
      If DisplayDays <= 40 Then
      Select Case Weekday(strdate)
        Case 1
          Print Left(GetShortNameDay7, 2) 'Sunday
        Case 2
          Print Left(GetShortNameDay1, 2) 'Monday
        Case 3
          Print Left(GetShortNameDay2, 2) 'Thuesday
        Case 4
          Print Left(GetShortNameDay3, 2) 'Wednesday
        Case 5
          Print Left(GetShortNameDay4, 2) 'Thuesday
        Case 6
          Print Left(GetShortNameDay5, 2) 'Friday
        Case 7
          Print Left(GetShortNameDay6, 2) 'Saturday
        End Select
      End If
        strdate = DateAdd("d", 1, strdate)
        CurrentDay = Day(CDate(strdate))
Next
myDetailsHeight = 0
For n = 0 To cDetail1.Count - 1
cDetail1(n).Visible = True
UserControl.cDetail1(n).DisplayControl CDate(StartDate), CDate(EndDate), False
cDetail1(n).CRefresh
myDetailsHeight = myDetailsHeight + cDetail1(n).Height
Next
DisplayDays = theDays


End Sub
Function MonthEnd(vbdate As String) As Boolean
  On Error GoTo erroh
    MonthEnd = Day(CDate(vbdate))
    MonthEnd = False
    Exit Function
erroh:
MonthEnd = True
End Function

Private Sub RenderControl(StartDate As Date, EndDate As Date, IncludeWeeks As Boolean)
Cls
theDays = Format(DateSerial(Year(StartDate), Month(StartDate) + 1, 1 - 1), "Dd") - 1
DaysAtCurrentMonth = UserControl.Width / (theDays + 1)
theWeek = DatePart("ww", "1" & GetDateSeparator & Month(StartDate) & GetDateSeparator & Year(StartDate))
For g = 0 To theDays
  If IncludeWeeks = True Then
    If theWeek <> fConvKW(CDate(g + 1 & GetDateSeparator & Month(StartDate) & GetDateSeparator & Year(StartDate))) Then
      CurrentY = 250
      CurrentX = ((g) * DaysAtCurrentMonth) + 30
      Print "KW " & fConvKW(CDate(g + 1 & GetDateSeparator & Month(StartDate) & GetDateSeparator & Year(StartDate)))
      theWeek = fConvKW(CDate(g + 1 & GetDateSeparator & Month(StartDate) & GetDateSeparator & Year(StartDate)))
      Line ((g) * DaysAtCurrentMonth, 300)-((g) * DaysAtCurrentMonth, ScaleHeight)
    End If
  End If
    CurrentX = (g * DaysAtCurrentMonth) '- 30
    Line (g * DaysAtCurrentMonth, 500)-(g * DaysAtCurrentMonth, ScaleHeight)

  
  
  CurrentY = 500
  CurrentX = (g * DaysAtCurrentMonth) - 30
  myMax = (g * DaysAtCurrentMonth) - 30
  Print g + 1

      CurrentY = 700
      CurrentX = (g * DaysAtCurrentMonth) + 30
      Select Case Weekday(g + 1 & GetDateSeparator & Month(StartDate) & GetDateSeparator & Year(StartDate))
        Case 1
          Print Left(GetShortNameDay7, 2) 'Sunday
        Case 2
          Print Left(GetShortNameDay1, 2) 'Monday
        Case 3
          Print Left(GetShortNameDay2, 2) 'Thuesday
        Case 4
          Print Left(GetShortNameDay3, 2) 'Wednesday
        Case 5
          Print Left(GetShortNameDay4, 2) 'Thuesday
        Case 6
          Print Left(GetShortNameDay5, 2) 'Friday
        Case 7
          Print Left(GetShortNameDay6, 2) 'Saturday
        End Select
Next
CurrentY = 0
CurrentX = myMax / 2
Print GetLongMonthName(Month(StartDate)) & " " & Year(StartDate)
End Sub
Private Function GetDateSeparator() As String
'################### Get SystemSetting for other Language Systems as German #######################
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, buffer, 99)
   GetDateSeparator = LPSTRToVBString(buffer)
End Function
Public Function GetShortNameDay1() As String
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME1, buffer, 99)
   GetShortNameDay1 = LPSTRToVBString(buffer)
End Function
Public Function GetShortNameDay2() As String
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME2, buffer, 99)
   GetShortNameDay2 = LPSTRToVBString(buffer)
End Function
Public Function GetShortNameDay3() As String
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME3, buffer, 99)
   GetShortNameDay3 = LPSTRToVBString(buffer)
End Function
Public Function GetShortNameDay4() As String
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME4, buffer, 99)
   GetShortNameDay4 = LPSTRToVBString(buffer)
End Function
Public Function GetShortNameDay5() As String
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME5, buffer, 99)
   GetShortNameDay5 = LPSTRToVBString(buffer)
End Function
Public Function GetShortNameDay6() As String
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME6, buffer, 99)
   GetShortNameDay6 = LPSTRToVBString(buffer)
End Function
Public Function GetShortNameDay7() As String
   Dim buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME7, buffer, 99)
   GetShortNameDay7 = LPSTRToVBString(buffer)
End Function
Public Function GetLongMonthName(cMonth As Integer) As String
   Dim buffer As String * 100
   Dim dl&
   Select Case cMonth
    Case 12
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME12, buffer, 99)
    Case 11
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME11, buffer, 99)
    Case 10
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME10, buffer, 99)
    Case 9
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME9, buffer, 99)
    Case 8
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME8, buffer, 99)
    Case 7
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME7, buffer, 99)
    Case 6
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME6, buffer, 99)
    Case 5
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME5, buffer, 99)
    Case 4
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME4, buffer, 99)
    Case 3
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME3, buffer, 99)
    Case 2
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME2, buffer, 99)
    Case 1
      dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME1, buffer, 99)
   
   End Select
   GetLongMonthName = LPSTRToVBString(buffer)
End Function

Private Function LPSTRToVBString$(ByVal s$)
'################### Stringparsing ################################################################
   Dim nullpos&
   nullpos& = InStr(s$, Chr$(0))
   If nullpos > 0 Then
      LPSTRToVBString = Left$(s$, nullpos - 1)
   Else
      LPSTRToVBString = ""
   End If
End Function

Private Function CalculateValue(X As Single, Y As Single) As Long
'################### Stringparsing ################################################################
On Error Resume Next
    Dim myValue As Long
    mvarStartValue = 1
    myValue = Int(X / (DaysAtCurrentMonth / 10))
    myValue = myValue + mvarStartValue * 10
    myTemp = Replace(myValue, " ", "")
    myValue = Left(myTemp, Len(myTemp) - 1)
    CalculateValue = CDate(DateAdd("d", myValue - 2, StartDate))
End Function


Private Sub cDetail1_DblClick(Index As Integer, Value As Date)
MsgBox "Neu am '" & Value & "'" & vbCrLf & "für User '" & cDetail1(Index).GetEntryName, vbInformation + vbSystemModal + vbMsgBoxRight

End Sub


Private Sub cDetail1_DblClickEntry(Index As Integer, cStartDate As Date, cEndDate As Date, cText As String, cGroup As String)
RaiseEvent DblClickEntry(cStartDate, cEndDate, cText, cGroup)
End Sub

Private Sub cDetail1_DragDate(Index As Integer, oldStartDate As Date, oldEndDate As Date, NewStartDate As Date, NewEndDate As Date, cText As String)
For g = 1 To ListView1.ListItems.Count
  If ListView1.ListItems(g).SubItems(4) = cDetail1(Index).GetEntryName And Format(ListView1.ListItems(g).Text, "Short Date") = Format(CDate(oldStartDate), "Short Date") And Format(ListView1.ListItems(g).SubItems(5), "Short Date") = Format(CDate(oldEndDate), "Short Date") And ListView1.ListItems(g).SubItems(2) = cText Then
    ListView1.ListItems(g).Text = NewStartDate
    ListView1.ListItems(g).SubItems(5) = NewEndDate
  End If
Next
cDetail1_OpenClose Index
RaiseEvent DragDropDate(oldStartDate, oldEndDate, NewStartDate, NewEndDate, cText)
End Sub

Private Sub cDetail1_HooverValue(Index As Integer, Value As Date)
RaiseEvent HooverValue(Value)
End Sub

Private Sub cDetail1_OpenClose(Index As Integer)
DoEvents
For g = 1 To cDetail1.Count - 1
  cDetail1(g).Move cDetail1(g).Left, cDetail1(g - 1).Top + cDetail1(g - 1).Height ' + 1000
Next


myDetailsHeight = 0
For g = 0 To cDetail1.Count - 1
  myDetailsHeight = myDetailsHeight + cDetail1(g).Height
Next
  If myDetailsHeight > Picture1.Height Then
    VScroll1.Top = Picture1.Top
    VScroll1.Height = Picture1.Height
    VScroll1.Left = (Picture1.Left + Picture1.Width) - VScroll1.Width
    VScroll1.Max = (myDetailsHeight - Picture1.Top) - 960
    VScroll1.Visible = True
  Else
    theFirtTop = 0
    For g = 0 To cDetail1.Count - 1
      cDetail1(g).Move cDetail1(g).Left, theFirtTop
      theFirtTop = theFirtTop + cDetail1(g).Height
    Next
    VScroll1.Visible = False
    VScroll1.Value = 0
  End If
End Sub

Private Sub cmdLeft_Click()
StartDate = DateAdd("d", -1, StartDate)
EndDate = DateAdd("d", DisplayDays, StartDate)
DisplayControl StartDate, EndDate, True
cDetail1_OpenClose 0
End Sub

Private Sub cmdRight_Click()
StartDate = DateAdd("d", 1, StartDate)
EndDate = DateAdd("d", DisplayDays, StartDate)
DisplayControl StartDate, EndDate, True
cDetail1_OpenClose 0
End Sub

Private Sub pic_DblClick(Index As Integer)
EndDate = pic(Index).Tag
DisplayControl CDate(DateAdd("d", -1, pic(Index).Tag)), CDate(pic(Index).Tag), True
End Sub

Private Sub pictMinus_Click()
DisplayControl CDate(StartDate), CDate(DateAdd("d", -1, EndDate)), True
End Sub

Private Sub pictPlus_Click()
DisplayControl CDate(StartDate), CDate(DateAdd("d", 1, EndDate)), True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent HooverValue(CalculateValue(X, Y))
End Sub

Private Sub UserControl_Resize()
For n = 0 To cDetail1.Count - 1
cDetail1(n).Width = UserControl.Width
Next
    Const SnippOff As Long = 3
    Dim hRgn As Long
    With WindowRect
        .Left = SnippOff
        .Top = SnippOff
        .Right = ScaleX(pictPlus.Width, ScaleMode, vbPixels) - SnippOff
        .Bottom = ScaleY(pictPlus.Height, ScaleMode, vbPixels) - SnippOff
        hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
        SetWindowRgn pictPlus.hWnd, hRgn, True
        DeleteObject hRgn
    End With
    With WindowRect
        .Left = SnippOff
        .Top = SnippOff
        .Right = ScaleX(pictMinus.Width, ScaleMode, vbPixels) - SnippOff
        .Bottom = ScaleY(pictMinus.Height, ScaleMode, vbPixels) - SnippOff
        hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
        SetWindowRgn pictMinus.hWnd, hRgn, True
        DeleteObject hRgn
    End With

    With WindowRect
        .Left = SnippOff
        .Top = SnippOff
        .Right = ScaleX(cmdLeft.Width, ScaleMode, vbPixels) - SnippOff
        .Bottom = ScaleY(cmdLeft.Height, ScaleMode, vbPixels) - SnippOff
        hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
        SetWindowRgn cmdLeft.hWnd, hRgn, True
        DeleteObject hRgn
    End With
    With WindowRect
        .Left = SnippOff
        .Top = SnippOff
        .Right = ScaleX(cmdRight.Width, ScaleMode, vbPixels) - SnippOff
        .Bottom = ScaleY(cmdRight.Height, ScaleMode, vbPixels) - SnippOff
        hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
        SetWindowRgn cmdRight.hWnd, hRgn, True
        DeleteObject hRgn
    End With


pictMinus.Move 0, 0
pictPlus.Move 0, UserControl.Width - pictPlus.Width
cmdLeft.Move 0, UserControl.Height - cmdLeft.Height
'cmdRight.Move UserControl.Width - cmdRight.Width, UserControl.Height - cmdRight.Height
cmdRight.Move 440, UserControl.Height - cmdRight.Height
Picture1.Move 0, 960, UserControl.Width, UserControl.Height - 1500
DisplayControl CDate(StartDate), CDate(EndDate), True
End Sub

Private Sub VScroll1_Change()
ScrollEntrysDown
End Sub
Private Sub ScrollEntrysDown()
If VScroll1.Visible = True Then
    cDetail1(0).Move cDetail1(0).Left, (-VScroll1.Value)
    cDetail1_OpenClose 0
End If
End Sub

Private Sub VScroll1_Scroll()
ScrollEntrysDown
End Sub
