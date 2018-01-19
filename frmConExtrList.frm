VERSION 5.00
Begin VB.Form frmConExtrList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelProcess 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   44
      Top             =   6870
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frameConceptoDer 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6705
      Left            =   7110
      TabIndex        =   23
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox chkMovim 
         Caption         =   "Sólo Cuentas con movimiento en el ejercicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   210
         TabIndex        =   40
         Top             =   4290
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Formato Extendido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   3960
         Value           =   1  'Checked
         Width           =   3405
      End
      Begin VB.TextBox txtPag2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1110
         TabIndex        =   8
         Tag             =   "imgConcepto"
         Top             =   2940
         Width           =   1305
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2490
         Width           =   2835
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmConExtrList.frx":0000
         Left            =   1110
         List            =   "frmConExtrList.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   2835
      End
      Begin VB.CheckBox chkTotalAsiento 
         Caption         =   "Salto de Página por Cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   3510
         Value           =   1  'Checked
         Width           =   3405
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1110
         TabIndex        =   5
         Tag             =   "imgConcepto"
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "1ªPágina"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   38
         Top             =   2940
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo en Cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   37
         Top             =   2220
         Width           =   2040
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   810
         Picture         =   "frmConExtrList.frx":0004
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Formato del Extracto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   36
         Top             =   1530
         Width           =   2040
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   35
         Top             =   990
         Width           =   690
      End
   End
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1050
         Width           =   4215
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1470
         Width           =   4215
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1230
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1050
         Width           =   1275
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
      End
      Begin VB.TextBox txtNIF 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "imgConcepto"
         Top             =   3360
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   2640
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   2220
         Width           =   1305
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   990
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   990
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "N.I.F."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Top             =   3420
         Width           =   960
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   930
         Picture         =   "frmConExtrList.frx":008F
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   930
         Picture         =   "frmConExtrList.frx":011A
         Top             =   2250
         Width           =   240
      End
      Begin VB.Label lblAsiento 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2550
         TabIndex        =   34
         Top             =   990
         Width           =   4095
      End
      Begin VB.Label lblAsiento 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2550
         TabIndex        =   33
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   28
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10290
      TabIndex        =   13
      Top             =   6870
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8730
      TabIndex        =   11
      Top             =   6870
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   6810
      Width           =   1335
   End
   Begin VB.Frame FrameTipoSalida 
      Caption         =   "Tipo de salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   14
      Top             =   4020
      Width           =   6915
      Begin VB.CommandButton PushButtonImpr 
         Caption         =   "Propiedades"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5190
         TabIndex        =   26
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   25
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   24
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1200
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   720
         Width           =   3345
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "eMail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "PDF"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Archivo csv"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1515
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label lblProgre 
      Caption         =   "Progreso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   43
      Top             =   6960
      Width           =   6195
   End
End
Attribute VB_Name = "frmConExtrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************

Public Legalizacion As String

Public Cuenta As String
Public Descripcion As String
Public FecDesde As String
Public FecHasta As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1

Private SQL As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String

Dim PararProceso As Boolean

Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmpconextcab", "codusu = " & vUsu.Codigo) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
    
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If
    
    'LEGALIZACION
    If Legalizacion <> "" Then
        CadenaDesdeOtroForm = "OK"
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdCancelProcess_Click()
    PararProceso = True
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Legalizacion <> "" Then
            optTipoSal(2).Value = True
            
            chkTotalAsiento.Value = False
            chkMovim.Value = False
            Me.Check1.Value = False
            
            cmdAccion_Click (1)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Form_Load()
    PrimeraVez = True

    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Extracto de Cuentas"
    lblProgre.Caption = ""
    For i = 0 To 1
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    PrimeraVez = True
     
    CargaCombo
     
    If Cuenta <> "" Then
        txtCuentas(0).Text = Cuenta
        txtCuentas(1).Text = Cuenta
        txtNCuentas(0).Text = Descripcion
        txtNCuentas(1).Text = Descripcion
    End If
    
    txtFecha(0).Text = FecDesde 'vParam.fechaini
    txtFecha(1).Text = FecHasta 'vParam.fechafin
     
    txtFecha(2).Text = Format(Now, "dd/mm/yyyy")

    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
 
    txtPag2(0).Text = 1
   
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    If Legalizacion <> "" Then
        txtFecha(0).Text = RecuperaValor(Legalizacion, 2) ' fecha inicio
        txtFecha(1).Text = RecuperaValor(Legalizacion, 3) ' fecha fin
        txtFecha(2).Text = RecuperaValor(Legalizacion, 1) ' fecha de informe
    End If
    
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCuentas(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCuentas(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCuentas_Click(Index As Integer)

    IndCodigo = Index

    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing

End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2
        IndCodigo = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
        
    End Select
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
        frmppal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmppal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmppal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmppal.cd1.FilterIndex = 1
    frmppal.cd1.ShowSave
    If frmppal.cd1.FileTitle <> "" Then
        If Dir(frmppal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmppal.cd1.FileName
    End If
End Sub

Private Sub PushButtonImpr_Click()
    frmppal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub


Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub


Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgCuentas"
        imgCuentas_Click Indice
    Case "imgFecha"
        imgFec_Click Indice
    End Select
    
End Sub

Private Sub txtCuentas_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    If txtCuentas(Index).Text = "" Then
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCuentas(Index).Text) Then
        If InStr(1, txtCuentas(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCuentas(Index).Text, vbExclamation
        txtCuentas(Index).Text = ""
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 0, 1 'Cuentas
            'lblCuentas(Index).Caption = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", txtCuentas(Index), "T")
            
            RC = txtCuentas(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, SQL) Then
                txtCuentas(Index) = RC
                txtNCuentas(Index).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
                PonFoco txtCuentas(Index)
            End If
            
            If Index = 0 Then Hasta = 1
            If Hasta >= 1 Then
                txtCuentas(Hasta).Text = txtCuentas(Index).Text
                txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
            End If
    End Select

End Sub


Private Sub AccionesCSV()
Dim Sql2 As String
Dim NF As Integer
Dim Rs As ADODB.Recordset

    '********************************************
    '********************************************
    '
    '   Este CSV no lo podemos generar desde
    '   el "public", ya que hay que el saldo
    '   inicial NO es cero si la fecha
    '   no es inicio    ejercicio
    '
    '
    '********************************************
    '********************************************
    On Error GoTo eAccionesCSV

    'Monto el SQL
    
    NF = -1
    NumRegElim = 0
    'GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    Set miRsAux = New ADODB.Recordset
    SQL = "select * from tmpconextcab where codusu =" & vUsu.Codigo & " ORDER BY cta"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Rs = New ADODB.Recordset
    While Not miRsAux.EOF
        SQL = "Select hlinapu.codmacta as Código, cuentas.nommacta Denominación, hlinapu.fechaent Fecha, hlinapu.numasien Asiento, hlinapu.numdocum Documento, hlinapu.codconce Concepto, "
        SQL = SQL & " hlinapu.ampconce Ampliación, hlinapu.ctacontr Contrapartida, hlinapu.codccost CC,"
        'Acumulado periodo anterior
        
        SQL = SQL & " hlinapu.timported Debe, hlinapu.timporteh Haber, "
        SQL = SQL & "if (@cuenta=hlinapu.codmacta, @saldo:= @saldo + (coalesce(hlinapu.timported,0) - coalesce(hlinapu.timporteh,0)), "
        SQL = SQL & " @saldo:= " & DBSet(miRsAux!acumantt, "N") & " + (coalesce(hlinapu.timported,0) - coalesce(hlinapu.timporteh,0)) ) as Saldo "
        
        If Combo1.ListIndex <> 1 Then SQL = SQL & ",if(punteada=1,'SI','') Punteada "
         
        SQL = SQL & ", if(@cuenta:=hlinapu.codmacta, '','')  as '' " ' hacemos la asignacion de la variable pero no queremos mostrarla
        SQL = SQL & " FROM (hlinapu  INNER JOIN tmpconextcab ON hlinapu.codmacta = tmpconextcab.cta and tmpconextcab.codusu = " & vUsu.Codigo & ")"
        SQL = SQL & " INNER JOIN cuentas ON hlinapu.codmacta = cuentas.codmacta,"
        SQL = SQL & " (select @saldo:=0 , @cuenta:=null) r "
        SQL = SQL & " where " & cadselect
        SQL = SQL & " AND  tmpconextcab.cta =" & DBSet(miRsAux!Cta, "T")
        SQL = SQL & " ORDER BY 1,3,4   "
        
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            If NumRegElim = 0 Then
                If Dir(txtTipoSalida(1).Text, vbArchive) <> "" Then
                    If MsgBox("El fichero ya existe. ¿Sobreescribir?", vbQuestion + vbYesNo) <> vbYes Then
                        miRsAux.Close
                        Rs.Close
                        GoTo eAccionesCSV
                    End If
                End If
                NumRegElim = NumRegElim + 1
                'Primer registro
                NF = FreeFile
                Open txtTipoSalida(1).Text For Output As #NF
                       
                       
                cad = ""
                For i = 0 To Rs.Fields.Count - 1
                    cad = cad & ";""" & Rs.Fields(i).Name & """"
                Next i
                Print #NF, Mid(cad, 2)
                    
            End If
 
            
            cad = ""
            For i = 0 To Rs.Fields.Count - 1
                cad = cad & ";""" & DBLet(Rs.Fields(i).Value, "T") & """"
            Next i
            Print #NF, Mid(cad, 2)
    
            Rs.MoveNext
        Wend
        Rs.Close
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If NumRegElim > 0 Then MsgBox "Fichero creado con éxito", vbInformation
    
    
eAccionesCSV:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    If NF > 0 Then Close #NF
    Set Rs = Nothing
    Set miRsAux = Nothing
End Sub






Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
    cadParam = cadParam & "pNIF=""" & txtNIF.Text & """|"
    
    If txtPag2(0).Text <> "" Then
        cadParam = cadParam & "pNumPag=" & txtPag2(0).Text - 1 & "|"
    Else
        cadParam = cadParam & "pNumPag=0|"
    End If
    cadParam = cadParam & "pSalto=" & chkTotalAsiento.Value & "|"
    ' Normal
    Select Case Combo1.ListIndex
        Case 0
            cadParam = cadParam & "pOpcion=0|"
        Case 1, 2, 3
            cadParam = cadParam & "pOpcion=1|"
    End Select
    
    Select Case Combo1.ListIndex
        Case 0
            cadParam = cadParam & "pTipo=""Normal""|"
        Case 1
            cadParam = cadParam & "pTipo=""Punteada y Pendiente""|"
        Case 2
            cadParam = cadParam & "pTipo=""Sólo punteada""|"
        Case 3
            cadParam = cadParam & "pTipo=""Sólo pendiente""|"
    End Select
    
    numParam = numParam + 6
    
    
    If Check1.Value = 1 Then
        indRPT = "0303-01" '"ConsExtracExt.rpt"
    Else
        indRPT = "0303-00" '"ConsExtrac.rpt"
    End If

    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu



    cadFormula = "{tmpconextcab.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text, (Legalizacion <> "")) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 60
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("hlinapu.fechaent", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    If Not PonerDesdeHasta("hlinapu.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
    
    Select Case Combo1.ListIndex
        Case 2 ' solo punteada
            If Not AnyadirAFormula(cadselect, "hlinapu.punteada = 1") Then Exit Function
        
        Case 3 ' sin puntear
            If Not AnyadirAFormula(cadselect, "hlinapu.punteada = 0") Then Exit Function
    End Select
    
    If Me.cmdAccion(0).visible Then Me.cmdAccion(0).Enabled = False
    Me.cmdAccion(1).Enabled = False
    Me.cmdCancelar.visible = False
    Me.cmdCancelProcess.visible = True
    
    'Proceso
    MontaSQL = CargaTemporal2
    
    
    PararProceso = False
    If Me.cmdAccion(0).visible Then Me.cmdAccion(0).Enabled = True
    cmdAccion(1).Enabled = True
    Me.cmdCancelar.visible = True
    Me.cmdCancelProcess.visible = False
    
    lblProgre.Caption = ""
End Function

Private Function CargaTemporal2() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim B As Boolean

    On Error GoTo eCargaTemporal
    
    CargaTemporal2 = False

    lblProgre.Caption = "Prepara campos"
    lblProgre.Refresh
    ' Tabla auxiliar en donde vamos a guardar las cuentas a mostrar
    SQL = "delete from tmpconextcab where codusu = " & vUsu.Codigo
    Conn.Execute SQL
            
    ' Tabla auxiliar en donde vamos a guardar las lineas de apuntes
    SQL = "delete from tmpconext where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    SQL = "select hlinapu.codmacta, sum(coalesce(timported,0)), sum(coalesce(timporteh,0)), "
    SQL = SQL & "cuentas.nommacta from hlinapu inner join cuentas on hlinapu.codmacta = cuentas.codmacta where (1=1) "
            
    If txtNIF.Text <> "" Then
        SQL = SQL & " and cuentas.nifdatos = " & DBSet(txtNIF, "T")
    End If
            
    If cadselect <> "" Then SQL = SQL & " and " & cadselect
    
    ' si esta marcada solo cuentas con movimientos en el ejercicio actual
    If chkMovim.Value = 1 Then
        SQL = SQL & " and hlinapu.codmacta in (select distinct codmacta from hlinapu where fechaent >= " & DBSet(vParam.fechaini, "F") & " and codconce < 900) "
    End If
            
    SQL = SQL & " group by 1,4 "
            
    Select Case Combo2.ListIndex
        Case 0 'Todas
        
        Case 1 'saldo <> 0
            SQL = SQL & " having sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) <> 0"
        Case 2 'saldo = 0
            SQL = SQL & " having sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) = 0"
    End Select

    B = True

    Set Rs = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF And B
        lblProgre.Caption = Rs!codmacta & " " & Rs!Nommacta
        lblProgre.Refresh
        B = (CargaDatosConExt(Rs!codmacta, txtFecha(0).Text, txtFecha(1).Text, cadselect, Rs.Fields(3).Value) = 0)
   
   
        DoEvents
        If PararProceso Then B = False   'Parara el proceso
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    CargaTemporal2 = B
    Screen.MousePointer = vbDefault
    Exit Function
    
eCargaTemporal:
    Err.Clear
End Function


Private Function CargaTemporal() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset

    On Error GoTo eCargaTemporal
    
    CargaTemporal = False
    
    ' Tabla auxiliar en donde vamos a guardar las cuentas a mostrar
    SQL = "delete from tmpconextcab where codusu = " & vUsu.Codigo
    Conn.Execute SQL
            
    ' Tabla auxiliar en donde vamos a guardar las lineas de apuntes
    SQL = "delete from tmpconext where codusu = " & vUsu.Codigo
    Conn.Execute SQL
            
            
    SQL = "insert into tmpconextcab (codusu, cta, acumperD, acumperH, cuenta) select " & vUsu.Codigo & ", hlinapu.codmacta, sum(coalesce(timported,0)), sum(coalesce(timporteh,0)), "
    SQL = SQL & "concat(hlinapu.codmacta, ' - ', cuentas.nommacta) from hlinapu inner join cuentas on hlinapu.codmacta = cuentas.codmacta where (1=1) "
            
    If txtNIF.Text <> "" Then
        SQL = SQL & " and hlinapu.nifdatos = " & DBSet(txtNIF, "T")
    End If
            
    If cadselect <> "" Then SQL = SQL & " and " & cadselect
            
    SQL = SQL & " group by 1,2, 5 "
            
    Select Case Combo2.ListIndex
        Case 0 'Todas
        
        Case 1 'saldo <> 0
            SQL = SQL & " having sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) <> 0"
        Case 2 'saldo = 0
            SQL = SQL & " having sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) = 0"
    End Select
            
    Conn.Execute SQL
    
    ' Solo para el caso de que no sea salida a csv cargamos las lineas y acumulados que necesitemos
    If optTipoSal(1).Value = 0 Then
        ' ACUMULADOS
        SQL = "select cta from tmpconextcab where codusu = " & vUsu.Codigo
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            'Acumulados anteriores al periodo
            Sql2 = "select sum(coalesce(timported,0)), sum(coalesce(timporteh,0)) from hlinapu where codmacta = " & DBSet(Rs!Cta, "T")
            Sql2 = Sql2 & " and fechaent >= " & DBSet(vParam.fechaini, "F") & " and fechaent < " & DBSet(txtFecha(0).Text, "F")
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            'Acumulados totales
            Sql3 = "select sum(coalesce(timported,0)), sum(coalesce(timporteh,0)) from hlinapu where codmacta = " & DBSet(Rs!Cta, "T")
            Sql3 = Sql3 & " and fechaent >= " & DBSet(vParam.fechaini, "F") & " and fechaent <= " & DBSet(vParam.fechafin, "F")
            
            Set Rs3 = New ADODB.Recordset
            Rs3.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Sql2 = "update tmpconextcab set acumantD = " & DBSet(Rs2.Fields(0).Value, "N")
            Sql2 = Sql2 & " , acumantH = " & DBSet(Rs2.Fields(1).Value, "N")
            Sql2 = Sql2 & " , acumantT = " & DBSet(DBLet(Rs2.Fields(0).Value, "N") - DBLet(Rs2.Fields(1).Value, "N"), "N")
            Sql2 = Sql2 & " , acumtotD = " & DBSet(Rs3.Fields(0).Value, "N")
            Sql2 = Sql2 & " , acumtotH = " & DBSet(Rs3.Fields(1).Value, "N")
            Sql2 = Sql2 & " , acumtotT = " & DBSet(DBLet(Rs3.Fields(0).Value, "N") - DBLet(Rs3.Fields(1).Value, "N"), "N")
            Sql2 = Sql2 & " where codusu = " & DBSet(vUsu.Codigo, "N") & " and cuenta = " & DBSet(Rs!Cta, "T")
            
            Conn.Execute Sql2
            
            'Lineas de apuntes
            Sql2 = "insert into tmpconext (codusu,cta,numdiari,Pos,fechaent,numasien,linliapu,nomdocum,ampconce,timporteD,timporteH,saldo,Punteada,contra,ccost) "
            Sql2 = Sql2 & " select " & vUsu.Codigo & ",codmacta, numdiari, @Pos:=@Pos + 1, fechaent, numasien, linliapu, numdocum, ampconce, timported, timporteh, "
            Sql2 = Sql2 & " @saldo:= @saldo + (coalesce(hlinapu.timported,0) - coalesce(hlinapu.timporteh,0)) s1, "
            Sql2 = Sql2 & " if(punteada=1,'SI','') punteada, ctacontr, codccost "
            Sql2 = Sql2 & " from hlinapu, (select @saldo:=" & DBSet(DBLet(Rs2.Fields(0).Value, "N") - DBLet(Rs2.Fields(1).Value, "N"), "N") & ") r, (select @Pos:=0) p "
            Sql2 = Sql2 & " where codmacta = " & DBSet(Rs!Cta, "T")
            If cadselect <> "" Then Sql2 = Sql2 & " and " & cadselect
            
            Select Case Combo1.ListIndex
                Case 2 ' solo punteada
                    Sql2 = Sql2 & " and punteada = 1 "
                Case 3 ' solo pendiente
                    Sql2 = Sql2 & " and punteada = 0 "
            End Select
            
            Sql2 = Sql2 & " order by 1,codmacta,fechaent,numdiari,numasien,linliapu "
            
            Conn.Execute Sql2
            
            
            Set Rs2 = Nothing
            Set Rs3 = Nothing
            
            
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    End If

    
    CargaTemporal = True
    
    Exit Function
    
eCargaTemporal:
    MuestraError Err.Number, "Carga Temporal", Err.Description
End Function

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda "imgFecha", Index
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtFecha(2).Text = "" Then
        MsgBox "Debe introducir un valor para la Fecha del listado.", vbExclamation
        PonFoco txtFecha(2)
        Exit Function
    End If

    DatosOK = True

End Function


Private Sub CargaCombo()
    
    Combo1.Clear

    'formato del extracto
    Combo1.AddItem "Normal"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Punteada y pdte"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Solo punteada"
    Combo1.ItemData(Combo1.NewIndex) = 2
    Combo1.AddItem "Solo pendiente"
    Combo1.ItemData(Combo1.NewIndex) = 3
    
    'saldo en la cuenta
    Combo2.Clear
    Combo2.AddItem "Todas"
    Combo2.ItemData(Combo2.NewIndex) = 0
    Combo2.AddItem "Saldo distinto CERO"
    Combo2.ItemData(Combo2.NewIndex) = 1
    Combo2.AddItem "Saldo igual a CERO"
    Combo2.ItemData(Combo2.NewIndex) = 2
  
End Sub

Private Sub txtNIF_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtPag2_GotFocus(Index As Integer)
    ConseguirFoco txtPag2(Index), 3
End Sub

Private Sub txtPag2_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub txtPag2_LostFocus(Index As Integer)
    
    txtPag2(Index).Text = Trim(txtPag2(Index).Text)
    If txtPag2(Index).Text <> "" Then
        If Not IsNumeric(txtPag2(Index).Text) Then
            MsgBox "Numero de página incorrecto: " & txtPag2(Index).Text, vbExclamation
            txtPag2(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
