VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#17.2#0"; "Codejock.Calendar.v17.2.0.ocx"
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   3315
   ClientLeft      =   5370
   ClientTop       =   4335
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin XtremeCalendarControl.DatePicker wndDatePicker 
      Height          =   3120
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3570
      _Version        =   1114114
      _ExtentX        =   6297
      _ExtentY        =   5503
      _StockProps     =   64
      AutoSize        =   0   'False
      ShowTodayButton =   0   'False
      ShowNoneButton  =   0   'False
      Show3DBorder    =   2
      YearsTriangle   =   -1  'True
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Event Selec(vFecha As Date)

'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal Wparam As Long, ByVal lParam As Long) As Long
Const CB_SHOWDROPDOWN = &H14F

Dim g_sDPpopUpCBValue As String

Public Fecha As Date

Dim TiempoClick As Single
Dim FechaSel As Date


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_DblClick()
    wndDatePicker_KeyDown 13, 0
End Sub

Private Sub Form_Load()
   ' wndDatePicker.VisualTheme = xtpCalendarThemeOffice2003
   
    wndDatePicker.AskDayMetrics = True
    
    TiempoClick = Timer
    If Fecha > "01/01/1901" Then
        FechaSel = Fecha
        wndDatePicker.Select Fecha
        wndDatePicker.EnsureVisibleSelection
    Else
        FechaSel = CDate(Now)
    End If
        
    Me.Icon = frmppal.Icon
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim nWidth As Long, nHeight As Long
    nWidth = Me.ScaleWidth - 200
    nHeight = Me.ScaleHeight - wndDatePicker.top - 100
    If nWidth > 0 And nHeight > 0 Then
        wndDatePicker.Move 100, wndDatePicker.top, nWidth, nHeight
    End If
End Sub


Private Sub popupTimer_Timer()

End Sub

Private Sub wndDatePicker_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If WeekDay(Day) = vbSunday Then
      '  Set Metrics.Font = Me.Font
       ' Metrics.ForeColor = vbRed
        Metrics.Font.Bold = True
    End If
End Sub

'Private Sub popupTimer_Timer()
'    popupTimer.Enabled = False
'
'    cmbPopUp.Text = g_sDPpopUpCBValue
'    g_sDPpopUpCBValue = "?-?-?"
'
'End Sub
'
Private Sub wndDatePicker_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim cerrar As Boolean
    
    cerrar = False
    If KeyCode = 13 Or KeyCode = 144 Or KeyCode = 127 Then 'enter
        Fecha = Me.wndDatePicker.Selection(0).DateBegin
        cerrar = True
    ElseIf KeyCode = 27 Then
        Fecha = "01/01/1900"
        cerrar = True
    End If
    
    If cerrar Then
        If Fecha <> "01/01/1900" Then RaiseEvent Selec(Fecha)
        Unload Me
    End If
End Sub


Private Sub wndDatePicker_MonthChanged()
  '  St op
End Sub

Private Sub wndDatePicker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Caption = Button
End Sub

Private Sub wndDatePicker_SelectionChanged()
Dim F As Date
Dim Reajusta As Boolean
Dim J As Integer
    
    
    On Error GoTo ewndDatePicker_SelectionChanged
    
'    CreateObject("WScript.Shell").SendKeys "{ENTER}"
    If Timer - TiempoClick > 0.3 Then
        'Me.wndDatePicker.Selection(0).DateBegin
        If wndDatePicker.Selection(0).DateBegin <> FechaSel Then
            Reajusta = False
            If Month(wndDatePicker.Selection(0).DateBegin) <> Month(FechaSel) Then
                Reajusta = True
            Else
                If Year(wndDatePicker.Selection(0).DateBegin) <> Year(FechaSel) Then Reajusta = True
            End If
            If Reajusta Then
                'Hay que cambiar la ventana "activa"
                
                
                J = WeekDay(wndDatePicker.Selection(0).DateBegin, vbMonday)
                F = wndDatePicker.Selection(0).DateBegin
                If Day(wndDatePicker.Selection(0).DateBegin) < 15 Then
                    
                
                    J = 7 - J
                    
                    If J > 0 Then F = DateAdd("d", J, F)
                        
                    
                        
                
                Else
                    'Veo el domingo de la semana que contiene el dia seleccionado
                    J = 7 - J
                    'wndDatePicker.FirstVisibleDay = F
                    If J > 0 Then F = DateAdd("d", -J, wndDatePicker.Selection(0).DateBegin)
                        
                    
                        
                    'Estamoas ajustando por arriba
                    'Caben 42 dias
                    'F = DateAdd("d", 42, wndDatePicker.Selection(0).DateBegin)
                End If
                wndDatePicker.EnsureVisible F
            End If
        Else
            wndDatePicker.EnsureVisibleSelection
        End If
        'wndDatePicker.Select wndDatePicker.HitTest
        FechaSel = wndDatePicker.Selection(0).DateBegin
        
       
        
        TiempoClick = Timer
    Else
        CreateObject("WScript.Shell").SendKeys "{F16}"
        'St op
    End If
    
    
    
    Exit Sub
ewndDatePicker_SelectionChanged:
    MuestraError Err.Number, , Err.Description
End Sub

