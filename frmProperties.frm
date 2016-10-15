VERSION 5.00
Begin VB.Form frmEditEvent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Untitled - Event"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRecurrence 
      Caption         =   "Recurrence..."
      Height          =   375
      Left            =   9240
      TabIndex        =   22
      Top             =   2220
      Width           =   1575
   End
   Begin VB.CommandButton btnCustomProperties 
      Caption         =   "Custom Properties ..."
      Height          =   375
      Left            =   9000
      TabIndex        =   20
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CheckBox chkMeeting 
      Caption         =   "Meeting"
      Height          =   195
      Left            =   8280
      TabIndex        =   19
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CheckBox chkPrivate 
      Caption         =   "&Private"
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtDescription 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2760
      Width           =   10695
   End
   Begin VB.ComboBox cmbShowTimeAs 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CheckBox chkAllDayEvent 
      Caption         =   "All da&y event"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cmbEndTime 
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cmbEndDate 
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox cmbStartTime 
      Height          =   315
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cmbStartDate 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox cmbLabel 
      Height          =   315
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   6375
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   9735
   End
   Begin VB.Label ctrlColor 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   10580
      TabIndex        =   21
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblShowTimeAs 
      Caption         =   "Sho&w time as:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2325
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   10800
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   10800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblEndTime 
      Caption         =   "En&d time:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1605
      Width           =   855
   End
   Begin VB.Label lblStartTime 
      Caption         =   "Start time:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label lblLabel 
      Caption         =   "La&bel:"
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   645
      Width           =   615
   End
   Begin VB.Label lblLocation 
      Caption         =   "&Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   645
      Width           =   855
   End
   Begin VB.Label lblSubject 
      Caption         =   "Sub&ject:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   285
      Width           =   855
   End
End
Attribute VB_Name = "frmEditEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const CB_SETDROPPEDWIDTH = &H160

Dim m_pEditingEvent As CalendarEvent
Dim m_bAddEvent As Boolean
Public AllDayOverride As Boolean

Private Sub btnCustomProperties_Click()
  '  If m_pEditingEvent Is Nothing Then
  '      Exit Sub
  '  End If
  '
  '  frmCustomEventProperties.SetEvent m_pEditingEvent
  '
  '  frmCustomEventProperties.Show vbModal, Me
End Sub

Private Sub btnRecurrence_Click()
    'Set frmEditRecurrence.m_pMasterEvent = m_pEditingEvent.CloneEvent
    'frmEditRecurrence.Show vbModal
   '
   ' Dim bRecurrenceStateChanged As Boolean
   ' bRecurrenceStateChanged = m_pEditingEvent.RecurrenceState <> frmEditRecurrence.m_pMasterEvent.RecurrenceState
   '
   ' Set m_pEditingEvent = frmEditRecurrence.m_pMasterEvent

    'If frmEditRecurrence.m_bUpdateFromEvent Or bRecurrenceStateChanged Then
    '    UpdateControlsFromEvent
    'End If

End Sub

Private Sub chkAllDayEvent_Click()

    cmbEndTime.Visible = IIf(chkAllDayEvent.Value = 1, False, True)
    cmbStartTime.Visible = IIf(chkAllDayEvent.Value = 1, False, True)

End Sub

Private Sub cmbEndTime_Click()
    Dim Index As Long
    Index = InStr(1, cmbEndTime.Text, "(")
    If Index > 0 Then
        cmbEndTime.Text = Left(cmbEndTime.Text, Index - 2)
    End If
    
    
End Sub


Private Sub cmbLabel_Click()
    Dim pLabel As CalendarEventLabel
    Dim nLabelID As Long
    
    nLabelID = cmbLabel.ItemData(cmbLabel.ListIndex)
    
    Set pLabel = frmInbox.wndCalendarControl.DataProvider.LabelList.Find(nLabelID)
    If Not pLabel Is Nothing Then
        ctrlColor.BackColor = pLabel.Color
    End If
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function DateFromString(DatePart As String, TimePart As String) As Date
    Dim dtDatePart As Date, dtTimePart As Date
    dtDatePart = DatePart
    dtTimePart = TimePart
    DateFromString = dtDatePart + dtTimePart
End Function

Function IsDateValid(DatePart As String) As Boolean
    IsDateValid = False
    On Error GoTo Error
    Dim dtDate As Date

    dtDate = DatePart
    IsDateValid = True
Error:
End Function

Private Function CheckDates() As Boolean
    CheckDates = True
    If (Not IsDateValid(cmbStartDate.Text)) Then
        cmbStartDate.SetFocus
        CheckDates = False
        Exit Function
    End If
    If (Not IsDateValid(cmbStartTime.Text)) Then
        cmbStartTime.SetFocus
        CheckDates = False
        Exit Function
    End If
    If (Not IsDateValid(cmbEndDate.Text)) Then
        cmbEndDate.SetFocus
        CheckDates = False
        Exit Function
    End If
    If (Not IsDateValid(cmbEndTime.Text)) Then
        cmbEndTime.SetFocus
        CheckDates = False
        Exit Function
    End If
End Function

Private Sub cmdOk_Click()

    If (Not CheckDates()) Then Exit Sub

    Dim StartTime As Date, EndTime As Date
    StartTime = DateFromString(cmbStartDate.Text, cmbStartTime.Text)
    EndTime = DateFromString(cmbEndDate.Text, cmbEndTime.Text)
    
    If chkAllDayEvent.Value = 1 Then
        If DateDiff("s", TimeValue(EndTime), 0) = 0 Then
            EndTime = EndTime + 1
        End If
    End If
    
    If m_pEditingEvent.RecurrenceState <> xtpCalendarRecurrenceMaster Then
        m_pEditingEvent.StartTime = StartTime
        m_pEditingEvent.EndTime = EndTime
    End If
    
    m_pEditingEvent.Subject = txtSubject.Text
    m_pEditingEvent.Location = txtLocation.Text
    m_pEditingEvent.Body = txtDescription
    m_pEditingEvent.AllDayEvent = chkAllDayEvent.Value = 1
    m_pEditingEvent.Label = cmbLabel.ItemData(cmbLabel.ListIndex)
    m_pEditingEvent.BusyStatus = cmbShowTimeAs.ListIndex
    
    m_pEditingEvent.PrivateFlag = chkPrivate.Value = 1
    m_pEditingEvent.MeetingFlag = chkMeeting.Value = 1
    
    If m_bAddEvent Then
        frmInbox.wndCalendarControl.DataProvider.AddEvent m_pEditingEvent
    Else
        frmInbox.wndCalendarControl.DataProvider.ChangeEvent m_pEditingEvent
    End If
    
    frmInbox.wndCalendarControl.Populate

    Unload Me
End Sub

Private Sub UpdateEndTimeCombo()
    On Error GoTo Error
    
    Dim i As Long
    For i = 1 To cmbEndTime.ListCount - 1
        cmbEndTime.RemoveItem 0
    Next i
    
    Dim BeginTime As Date
    BeginTime = TimeValue(cmbStartTime.Text)
    
    cmbEndTime.AddItem BeginTime & " (0 minutes)"
    cmbEndTime.AddItem TimeValue(BeginTime + 1 / 24 / 2) & " (30 minutes)"
    cmbEndTime.AddItem TimeValue(BeginTime + 1 / 24) & " (1 hour)"
    
    For i = 3 To 47
        cmbEndTime.AddItem TimeValue(BeginTime + i / 24 / 2) & " (" & i / 2 & " hours)"
    Next i
    
    Call SendMessage(cmbEndTime.hwnd, CB_SETDROPPEDWIDTH, 200, 0)
    
    
Error:
    
End Sub

Private Sub Form_Load()

    '===============================
    Dim pLabel As CalendarEventLabel
    
    For Each pLabel In frmInbox.wndCalendarControl.DataProvider.LabelList
        cmbLabel.AddItem pLabel.Name
        cmbLabel.ItemData(cmbLabel.NewIndex) = pLabel.LabelID
    Next
        
    cmbShowTimeAs.AddItem "Free"
    cmbShowTimeAs.AddItem "Tentative"
    cmbShowTimeAs.AddItem "Busy"
    cmbShowTimeAs.AddItem "Out of Office"
    
    If Not m_bAddEvent Then
        If Not m_pEditingEvent Is Nothing Then
            If m_pEditingEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
                m_pEditingEvent.MakeAsRException
            End If
        
            UpdateControlsFromEvent
        End If
    End If
    
End Sub

Public Sub SetStartEnd(BeginSelection As Date, EndSelection As Date, AllDay As Boolean)
    Dim StartDate As Date, StartTime As Date, EndDate As Date, EndTime As Date

    StartDate = DateValue(BeginSelection)
    StartTime = TimeValue(BeginSelection)

    EndDate = DateValue(EndSelection)
    EndTime = TimeValue(EndSelection)

    If AllDay Then
        cmbEndTime.Visible = False
        cmbStartTime.Visible = False
    
        If DateDiff("s", EndTime, 0) = 0 Then
            EndDate = EndDate - 1
        End If
    End If
    
    cmbStartDate.Text = StartDate
    cmbStartTime.Text = StartTime
    
    UpdateEndTimeCombo

    cmbEndDate.Text = EndDate
    cmbEndTime.Text = EndTime
    
 
End Sub


Public Sub NewEvent()
    Set m_pEditingEvent = frmInbox.wndCalendarControl.DataProvider.CreateEvent
    m_bAddEvent = True
    
    Dim BeginSelection As Date, EndSelection As Date, AllDay As Boolean
    frmInbox.wndCalendarControl.ActiveView.getSelection BeginSelection, EndSelection, AllDay

    SetStartEnd BeginSelection, EndSelection, AllDay
    
    If AllDayOverride Then
        AllDay = True
    End If
    
    chkAllDayEvent.Value = IIf(AllDay, 1, 0)
    
    txtSubject = "New Event"

    cmbShowTimeAs.ListIndex = IIf(AllDay, 0, 2)
    cmbLabel.ListIndex = 0
    
    AllDayOverride = False
    
End Sub

Public Sub ModifyEvent(ModEvent As CalendarEvent)
    Set m_pEditingEvent = ModEvent
    m_bAddEvent = False
    
'    txtSubject.Text = m_pEditingEvent.Subject
'    txtDescription.Text = m_pEditingEvent.Description
'    txtLocation.Text = m_pEditingEvent.Location
'
'    chkAllDayEvent.Value = IIf(m_pEditingEvent.AllDayEvent, 1, 0)
'
'    Dim i As Long
'    For i = 0 To cmbLabel.ListCount - 1
'        If cmbLabel.ItemData(i) = m_pEditingEvent.Label Then
'            cmbLabel.ListIndex = i
'            Exit For
'        End If
'    Next
'
'    cmbShowTimeAs.ListIndex = m_pEditingEvent.BusyStatus
'
'    chkPrivate.Value = IIf(m_pEditingEvent.PrivateFlag, 1, 0)
'    chkMeeting.Value = IIf(m_pEditingEvent.MeetingFlag, 1, 0)
'
'    SetStartEnd m_pEditingEvent.StartTime, m_pEditingEvent.EndTime, m_pEditingEvent.AllDayEvent
'
    If (m_pEditingEvent.Subject <> "") Then Me.Caption = m_pEditingEvent.Subject & " - Event"
    
End Sub

Public Sub UpdateControlsFromEvent()

    txtSubject = m_pEditingEvent.Subject
    txtDescription = m_pEditingEvent.Body
    txtLocation = m_pEditingEvent.Location
    
    chkAllDayEvent = IIf(m_pEditingEvent.AllDayEvent, 1, 0)
    
    Dim i As Long
    For i = 0 To cmbLabel.ListCount - 1
        If cmbLabel.ItemData(i) = m_pEditingEvent.Label Then
            cmbLabel.ListIndex = i
            Exit For
        End If
    Next
    
    cmbShowTimeAs.ListIndex = m_pEditingEvent.BusyStatus
    
    chkPrivate.Value = IIf(m_pEditingEvent.PrivateFlag, 1, 0)
    chkMeeting.Value = IIf(m_pEditingEvent.MeetingFlag, 1, 0)
    
    SetStartEnd m_pEditingEvent.StartTime, m_pEditingEvent.EndTime, m_pEditingEvent.AllDayEvent
    
    If (m_pEditingEvent.Subject <> "") Then
        Me.Caption = m_pEditingEvent.Subject & " - Event"
    End If

    Dim bDatesVisible As Boolean
    bDatesVisible = m_pEditingEvent.RecurrenceState <> xtpCalendarRecurrenceMaster
       
    lblStartTime.Visible = bDatesVisible
    lblEndTime.Visible = bDatesVisible
    cmbStartDate.Visible = bDatesVisible
    cmbStartTime.Visible = bDatesVisible
    cmbEndDate.Visible = bDatesVisible
    cmbEndTime.Visible = bDatesVisible
    chkAllDayEvent.Visible = bDatesVisible
        
    If bDatesVisible Then
        chkAllDayEvent_Click
    End If
    
End Sub


