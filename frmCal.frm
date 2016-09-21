VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form frmCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCal 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCal 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _Version        =   524288
      _ExtentX        =   9340
      _ExtentY        =   5953
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1998
      Month           =   10
      Day             =   7
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   2
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Fecha As Date
Private Dia As Integer
Private Mes As Integer
Private Año As Integer
Public Event Selec(vFecha As Date)

Private Sub Calendar1_Click()
    Fecha = Calendar1.Value
End Sub

Private Sub Calendar1_DblClick()
    Calendar1_Click
    cmdCal_Click (0)
End Sub

Private Sub cmdCal_Click(Index As Integer)
    Select Case Index
        Case 0
            '-- Aceptar
            RaiseEvent Selec(Fecha)
        Case 1
    End Select
    Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dia = Val(Format(Fecha, "dd"))
    Mes = Val(Format(Fecha, "mm"))
    Año = Val(Format(Fecha, "yyyy"))
    Calendar1.Day = 1
    Calendar1.Year = Año
    Calendar1.Month = Mes
    Calendar1.Day = Dia

    Calendar1.FirstDay = 1
End Sub
