VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl cDayview 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   2280
   Begin VB.PictureBox cAppoiment 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   660
      Picture         =   "Dayview.ctx":0000
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid cDayGrid 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2355
      _Version        =   393216
      Rows            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColorSel    =   -2147483647
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   3
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "cDayview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub DrawADate(cMinuteSteps As Integer, cDisplayFirstTime As String, AMPM As Boolean, cSetWorktimeFrom As String, cSetWorktimeTo As String, cTheDate As Date)
'##################################################################################################
'#### Draw Times on the Picture Begin with DisplayedTime and Set the Minutes Steps ################
'##################################################################################################
cResize
cDisplayAMPM = AMPM
cInternalMinuteSteps = cMinuteSteps
If cMinuteSteps = 0 Then cMinuteSteps = 60
theTime = "00:00:00" 'cDisplayFirstTime
NoWorkTimeColour = RGB(244, 238, 166)
WorkTimeColour = RGB(251, 249, 217)
mySteps = 0
For g = 60 To 86400 Step (cMinuteSteps * 60)
    mySteps = mySteps + 1
    cDayGrid.Rows = mySteps
    cDayGrid.Row = mySteps - 1
    cDayGrid.Col = 0
    cDayGrid.CellBackColor = RGB(218, 218, 218)
    cDayGrid.Text = Format(theTime, "hh:mm")
    cDayGrid.ColWidth(0) = 700
    cDayGrid.CellAlignment = 0
    cDayGrid.Col = 1
    cDayGrid.CellBackColor = NoWorkTimeColour
    cDayGrid.Col = 0
    theTime = DateAdd("s", cMinuteSteps * 60, theTime)
Next
cDayGrid.ColWidth(1) = UserControl.Width - 1000

cDayGrid.ScrollTrack = True
For g = 1 To cDayGrid.Rows - 1
  If cDayGrid.TextMatrix(g, 0) = Format(cDisplayFirstTime, "hh:mm") Then
    cDayGrid.TopRow = g
    myTopRow = g
  End If
  cGridTime = Replace(Format(cDayGrid.TextMatrix(g, 0), "hh:mm"), ":", "")
  cSetWorktimeFromR = Replace(Format(cSetWorktimeFrom, "hh:mm"), ":", "")
  cSetWorktimeToR = Replace(Format(cSetWorktimeTo, "hh:mm"), ":", "")
  If cGridTime >= cSetWorktimeFromR And cGridTime <= cSetWorktimeToR Then
    cDayGrid.Col = 1
    cDayGrid.Row = g
    cDayGrid.CellBackColor = WorkTimeColour
  End If
Next
cDayGrid.TopRow = myTopRow
RenderAppoiments

'##################################################################################################
'#### Set VScroll Values for the Correct Scrolling ################################################
'##################################################################################################
End Sub
Private Sub RenderAppoiments()
For g = 0 To cAppoiment.Count - 1
  If cAppoiment(g).Tag <> "" Then
    cStartTime = Replace(Format(Left(cAppoiment(g).Tag, InStr(cAppoiment(g).Tag, "|") - 1), "hh:mm"), ":", "")
    cEndTime = Replace(Format(Right(cAppoiment(g).Tag, Len(Left(cAppoiment(g).Tag, InStr(cAppoiment(g).Tag, "|"))) - 1), "hh:mm"), ":", "")
    
     
      cTheCounter = 0
      For n = 1 To cDayGrid.Rows - 1
        cGridTime = Replace(Format(cDayGrid.TextMatrix(n, 0), "hh:mm"), ":", "")
        If cGridTime >= cStartTime And cGridTime <= cEndTime Then
          cTheCounter = cTheCounter + 1
          If cTheCounter = 1 Then
            thePicTop = cDayGrid.RowPos(n)
          Else
            thePicBottom = cDayGrid.RowPos(n)
          End If
        End If
      Next
     
      cAppoiment(g).Move 700, thePicTop, cDayGrid.ColWidth(1) - 200, thePicBottom - thePicTop + 240
  End If
Next
End Sub
Sub AddAppoiment(cStartTime As String, cEndTime As String, cPicture As PictureBox, cColour As OLE_COLOR)
'If cAppoiment.Count <> 0 Then
Load cAppoiment(cAppoiment.Count)
cAppoiment(cAppoiment.Count - 1).Tag = cStartTime & "|" & cEndTime
cAppoiment(cAppoiment.Count - 1).BackColor = cColour
cAppoiment(cAppoiment.Count - 1).Picture = cPicture
cAppoiment(cAppoiment.Count - 1).Visible = True
cAppoiment(cAppoiment.Count - 1).Move 0, 0
cAppoiment(cAppoiment.Count - 1).ZOrder 0
RenderAppoiments
End Sub
Sub UnloadAppoiments()
For g = cAppoiment.Count - 1 To 1 Step -1
  Unload cAppoiment(g)
Next
End Sub

Private Sub cResize()
cDayGrid.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Private Sub cAppoiment_Click(Index As Integer)
cAppoiment(Index).ZOrder 0
End Sub

Private Sub cAppoiment_DblClick(Index As Integer)
MsgBox "Clicked on Appoint"
End Sub


Private Sub cDayGrid_Scroll()
RenderAppoiments
End Sub


