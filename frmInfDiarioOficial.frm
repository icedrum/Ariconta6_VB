VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfDiarioOficial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   5055
      Left            =   7110
      TabIndex        =   12
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtNumRes 
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
         Index           =   4
         Left            =   1890
         TabIndex        =   41
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtNumRes 
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
         Index           =   3
         Left            =   1890
         TabIndex        =   40
         Top             =   2250
         Width           =   1455
      End
      Begin VB.TextBox txtNumRes 
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
         Left            =   1890
         TabIndex        =   39
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtNumRes 
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
         Left            =   1890
         TabIndex        =   38
         Top             =   1170
         Width           =   1455
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   1890
         TabIndex        =   34
         Top             =   690
         Width           =   1485
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   180
         TabIndex        =   24
         Top             =   3120
         Width           =   4185
         Begin VB.CheckBox Check1 
            Caption         =   "�ltimo:  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Value           =   1  'Checked
            Width           =   1155
         End
         Begin VB.CheckBox Check1 
            Caption         =   "9� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   120
            TabIndex        =   33
            Top             =   1380
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "8� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   2760
            TabIndex        =   32
            Top             =   990
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "7� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   1380
            TabIndex        =   31
            Top             =   990
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "6� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   30
            Top             =   990
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "5� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   2760
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "4� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   1380
            TabIndex        =   28
            Top             =   600
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "3� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2� nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   2760
            TabIndex        =   26
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1er nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1380
            TabIndex        =   25
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   19
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Acumulado Haber"
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
         Index           =   3
         Left            =   150
         TabIndex        =   44
         Top             =   2820
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Acumulado Debe"
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
         Index           =   2
         Left            =   150
         TabIndex        =   43
         Top             =   2340
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Asiento"
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
         Index           =   0
         Left            =   150
         TabIndex        =   42
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Nro.P�gina"
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
         TabIndex        =   36
         Top             =   1740
         Width           =   1080
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
         Top             =   750
         Width           =   690
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   7
         Left            =   1440
         Picture         =   "frmInfDiarioOficial.frx":0000
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selecci�n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtAno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3210
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmInfDiarioOficial.frx":008B
         Left            =   1170
         List            =   "frmInfDiarioOficial.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1290
         Width           =   1935
      End
      Begin VB.TextBox txtAno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3210
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmInfDiarioOficial.frx":008F
         Left            =   1170
         List            =   "frmInfDiarioOficial.frx":0091
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   810
         Width           =   1935
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
         Left            =   180
         TabIndex        =   18
         Top             =   1260
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
         Left            =   180
         TabIndex        =   17
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Mes / A�o"
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
         Left            =   180
         TabIndex        =   16
         Top             =   540
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   10320
      TabIndex        =   2
      Top             =   5190
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
      TabIndex        =   0
      Top             =   5190
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
      TabIndex        =   1
      Top             =   5190
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
      TabIndex        =   3
      Top             =   2400
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
         TabIndex        =   15
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   14
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   13
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancelarAccion 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   37
      Top             =   5190
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Index           =   25
      Left            =   1650
      TabIndex        =   45
      Top             =   5220
      Width           =   5565
   End
End
Attribute VB_Name = "frmInfDiarioOficial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1306

Public Opcion As Byte
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

Public Legalizacion As String ' "fecha informe|fechainicio|fechafin|nrodigitos"


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon  As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private frmCtas As frmCtasAgrupadas

Private SQL As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim RS As ADODB.Recordset

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean


Dim HanPulsadoSalir As Boolean
Dim Importe As Currency
Dim CONT As Long

Public Sub InicializarVbles(A�adireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If A�adireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub


Private Sub Check1_Click(Index As Integer)
Dim Valor As Byte
Dim I As Integer
    
    Valor = Check1(Index).Value
    
    If Valor = 1 Then
        For I = 1 To Check1.Count
            If Check1(I).Visible Then
                If I <> Index Then Check1(I).Value = 0
            End If
        Next I
'        Check1(Index).Value = Valor
    End If


End Sub

Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    PulsadoCancelar = False
    Me.cmdCancelarAccion.Visible = True
    Me.cmdCancelarAccion.Enabled = True
    
    Me.cmdCancelar.Visible = False
    Me.cmdCancelar.Enabled = False
        
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    

'++
    Screen.MousePointer = vbHourglass
    If GenerarLibroResumen Then
        
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
    
    End If
    
    
    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True
    
    
    Screen.MousePointer = vbDefault
    Label2(25).Caption = ""

End Sub

Private Sub cmdCancelar_Click()
    If Me.cmdCancelarAccion.Visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Legalizacion <> "" Then
            Me.optTipoSal(2).Value = True
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
    
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Diario Oficial"

    
    PrimeraVez = True
     
    CargarComboFecha
     
    PonerNiveles
     
    If Legalizacion <> "" Then
        txtFecha(7).Text = RecuperaValor(Legalizacion, 1)
    
        'Fecha inicial
        cmbFecha(0).ListIndex = Month(RecuperaValor(Legalizacion, 2)) - 1
        cmbFecha(1).ListIndex = Month(RecuperaValor(Legalizacion, 3)) - 1
    
        txtAno(0).Text = Year(RecuperaValor(Legalizacion, 2)) - 1
        txtAno(1).Text = Year(RecuperaValor(Legalizacion, 3)) - 1
            
        For I = 1 To 10
            If RecuperaValor(Legalizacion, 4) = Check1(I).Tag Then Check1(I).Value = 1
        Next I
    
    Else
        'Fecha informe
        txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
        'Fecha inicial
        cmbFecha(0).ListIndex = Month(vParam.fechaini) - 1
        cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1
        
        ' SE OFERTA EL EJERCICIO ANTERIOR AL ACTUAL
        txtAno(0).Text = Year(vParam.fechaini) - 1
        txtAno(1).Text = Year(vParam.fechafin) - 1
    End If
   
    PosicionarCombo cmbFecha(0), cmbFecha(0).ListIndex
    PosicionarCombo cmbFecha(1), cmbFecha(1).ListIndex
        
   
   
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
    
    
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub




Private Sub Image1_Click(Index As Integer)

    Select Case Index
        Case 0 'cuentas agrupadas
            Set frmCtas = New frmCtasAgrupadas
            frmCtas.Show vbModal
            Set frmCtas = Nothing
    End Select

End Sub



Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 7
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
        frmPpal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmPpal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmPpal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmPpal.cd1.FilterIndex = 1
    frmPpal.cd1.ShowSave
    If frmPpal.cd1.FileTitle <> "" Then
        If Dir(frmPpal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmPpal.cd1.FileName
    End If
End Sub

Private Sub PushButtonImpr_Click()
    frmPpal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Set frmCtas = New frmCtasAgrupadas
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
    
End Sub

Private Sub AccionesCSV()
    
    SQL = "select fecha, asiento, cuenta, titulo, concepto, coalesce(debe,0) debe, coalesce(haber,0) haber from tmpdirioresum where codusu = " & vUsu.Codigo
    SQL = SQL & " order by clave "

        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    
    cadParam = cadParam & "pFecha=""" & txtFecha(7).Text & """|"
    
    'Numero de p�gina
    If txtNumRes(1).Text <> "" Then
        cadParam = cadParam & "pNumPag=" & txtNumRes(1).Text - 1 & "|"
    Else
        cadParam = cadParam & "pNumPag=0|"
    End If
    numParam = numParam + 2
    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & txtAno(0).Text & " a " & cmbFecha(1).Text & " " & txtAno(1).Text & """|"
    numParam = numParam + 1
    
    
    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "1306-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"SumasySaldos.rpt"

    cadFormula = "{tmpdirioresum.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
    
    MontaSQL = True
           
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
    
    If txtAno(0).Text = "" Or txtAno(1).Text = "" Then
        MsgBox "Introduce las fechas(a�os) de consulta", vbExclamation
        Exit Function
    End If
    If Me.cmbFecha(0).ListIndex < 0 Then
       MsgBox "Seleccione mes consulta desde", vbExclamation
       Exit Function
    End If
    If Me.cmbFecha(1).ListIndex < 0 Then
       MsgBox "Seleccione mes consulta hasta", vbExclamation
       Exit Function
    End If
    
    If Not ComparaFechasCombos(0, 1, 0, 1) Then Exit Function
    If txtFecha(7).Text <> "" Then
        If Not IsDate(txtFecha(7).Text) Then
            MsgBox "Fecha impresi�n incorrecta", vbExclamation
            txtFecha(7).SetFocus
        End If
    End If
    
    
    If Abs(Val(txtAno(0).Text) - Val(txtAno(1).Text)) > 2 Then
        MsgBox "Fechas pertenecen a ejercicios distintos.", vbExclamation
        Exit Function
    End If


    'Fechas
    'Trabajaresmos contra ejercicios cerrados
    'Si el mes es mayor o igual k el de inicio, significa k la feha
    'de inicio de aquel ejercicio fue la misma k ahora pero de aquel a�o
    'si no significa k fue la misma de ahora pero del a�o anterior
    I = cmbFecha(0).ListIndex + 1
    If I >= Month(vParam.fechaini) Then
        CONT = Val(txtAno(0).Text)
    Else
        CONT = Val(txtAno(0).Text) - 1
    End If
    cad = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & CONT
    FechaIncioEjercicio = CDate(cad)
    
    I = cmbFecha(1).ListIndex + 1
    If I <= Month(vParam.fechafin) Then
        CONT = Val(txtAno(1).Text)
    Else
        CONT = Val(txtAno(1).Text) + 1
    End If
    cad = Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & CONT
    FechaFinEjercicio = CDate(cad)

    
    
    
    
    
    'Veamos si pertenecen a un mismo a�o
    If Abs(DateDiff("d", FechaFinEjercicio, FechaIncioEjercicio)) > 365 Then
        MsgBox "Las fechas son incorrectas. Abarca mas de un ejercicio", vbExclamation
        Exit Function
    End If


    'AHora, si ha puesto importes, entonces veremos
    'Si :  -importes correctos.
    '      -si exite importe, que no sea mes inicio ejerecicio
    txtNumRes(3).Text = Trim(txtNumRes(3).Text)
    txtNumRes(4).Text = Trim(txtNumRes(4).Text)
    If txtNumRes(3).Text <> "" Or txtNumRes(4).Text <> "" Then
       If cmbFecha(0).ListIndex + 1 = Month(FechaIncioEjercicio) Then
            MsgBox "No puede poner importes para el mes de inicio de ejerecicio", vbExclamation
            Exit Function
        End If
    End If
    
    'Solo un nivel seleccionado
    CONT = 0
    For I = 1 To 10
        If Check1(I).Visible = True Then
            If Check1(I).Value Then CONT = CONT + 1
        End If
    Next I
    If CONT <> 1 Then
        MsgBox "Seleccione uno, y solo uno, de los niveles para mostrar el informe", vbExclamation
        Exit Function
    End If
    



    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer


QueCombosFechaCargar "0|1|"

For I = 1 To vEmpresa.numnivel - 1
    J = DigitosNivel(I)
    Check1(I).Visible = True
    Check1(I).Caption = "Digitos: " & J
Next I



End Sub




Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    cad = RecuperaValor(Lista, L)
    If cad <> "" Then
        I = Val(cad)
        With cmbFecha(I)
            .Clear
            For CONT = 1 To 12
                RC = "25/" & CONT & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next CONT
        End With
    End If
    L = L + 1
Loop Until cad = ""
End Sub


Private Function ComparaFechasCombos(Indice1 As Integer, Indice2 As Integer, InCombo1 As Integer, InCombo2 As Integer) As Boolean
    ComparaFechasCombos = False
    If txtAno(Indice1).Text <> "" And txtAno(Indice2).Text <> "" Then
        If Val(txtAno(Indice1).Text) > Val(txtAno(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        Else
            If Val(txtAno(Indice1).Text) = Val(txtAno(Indice2).Text) Then
                If Me.cmbFecha(InCombo1).ListIndex > Me.cmbFecha(InCombo2).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    ComparaFechasCombos = True
End Function


'Siempre k la fecha no este en fecha siguiente
Private Function HayAsientoCierre(Mes As Byte, Anyo As Integer, Optional Contabilidad As String) As Boolean
Dim C As String
    HayAsientoCierre = False
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierre = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            C = "Select count(*) from " & Contabilidad
            C = C & " hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then
                    If RS.Fields(0) > 0 Then HayAsientoCierre = True
                End If
            End If
            RS.Close
        End If
    End If
End Function




Private Function TieneCuentasEnTmpBalance(DigitosNivel As String) As Boolean
Dim RS As ADODB.Recordset
Dim C As String

    Set RS = New ADODB.Recordset
    TieneCuentasEnTmpBalance = False
    C = Mid("__________", 1, CInt(DigitosNivel))
    C = "Select count(*) from tmpbalancesumas  where cta like '" & C & "'"
    C = C & " AND codusu = " & vUsu.Codigo
    RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then TieneCuentasEnTmpBalance = True
        End If
    End If
    RS.Close
End Function

Private Sub PonerNiveles()
Dim I As Integer
Dim J As Integer


    Frame2.Visible = True
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        cad = "Digitos: " & J
        Check1(I).Visible = True
        Check1(I).Tag = J
        Me.Check1(I).Caption = cad
        Me.Check1(I).Value = 0
    Next I
    Check1(10).Tag = vEmpresa.DigitosUltimoNivel
    Check1(10).Visible = True
    Me.Check1(10).Value = 0
    For I = vEmpresa.numnivel To 9
        Check1(I).Visible = False
    Next I
    
    
End Sub



Private Sub txtNumRes_GotFocus(Index As Integer)
    ConseguirFoco txtNumRes(Index), 3
End Sub

Private Sub txtNumRes_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtNumRes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtNumRes_LostFocus(Index As Integer)

txtNumRes(Index).Text = Trim(txtNumRes(Index).Text)
If txtNumRes(Index).Text = "" Then Exit Sub
If Not IsNumeric(txtNumRes(Index).Text) Then
    MsgBox "El campo tiene que ser num�rico: " & txtNumRes(Index).Text, vbExclamation
    txtNumRes(Index).Text = ""
    txtNumRes(Index).SetFocus
    Exit Sub
Else
    If Index = 3 Or Index = 4 Then PonerFormatoDecimal txtNumRes(Index), 1
End If
End Sub


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtTitulo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Function GenerarLibroResumen() As Boolean
Dim I2 As Currency

    On Error GoTo EGenerarLibroResumen
    GenerarLibroResumen = False
    
    'Eliminamos registros tmp
    SQL = "Delete FROM tmpdirioresum where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
        
    'Comprobamos k nivel
    For I = 1 To Me.Check1.Count
        If Check1(I).Visible Then
            If Check1(I).Value Then
                CONT = I
                Exit For
            End If
        End If
    Next I
    
    
    I = CONT
    FijaValoresLibroResumen FechaIncioEjercicio, FechaFinEjercicio, I, False, txtNumRes(0).Text
    
    Importe = 0
    I2 = 0
    If txtAno(0).Text = txtAno(1).Text Then
        I = CInt(Val(txtAno(0).Text))
        For CONT = cmbFecha(0).ListIndex + 1 To cmbFecha(1).ListIndex + 1
           Label2(25).Caption = "Fecha: " & CONT & " / " & I
           Label2(25).Refresh
           
           DoEvents
           If PulsadoCancelar Then Exit Function
           
           
           'Si ha puesto ACUMULADOS ANTERIORES
           If CONT = cmbFecha(0).ListIndex + 1 Then
                If txtNumRes(3).Text <> "" Then Importe = CCur(TransformaPuntosComas(txtNumRes(3).Text))
                If txtNumRes(4).Text <> "" Then I2 = CCur(TransformaPuntosComas(txtNumRes(4).Text))
           End If
           ProcesaLibroResumen CONT, I, Importe, I2
           Importe = 0
           I2 = 0
        Next CONT
    Else
        'A�os partidos
        'El primer tramo de hasta fin de a�os
        I = CInt(Val(txtAno(0).Text))
        For CONT = cmbFecha(0).ListIndex + 1 To 12
           Label2(25).Caption = "Fecha: " & CONT & " / " & I
           Label2(25).Refresh
           
           DoEvents
           If PulsadoCancelar Then Exit Function
           
           If CONT = cmbFecha(0).ListIndex + 1 Then
                If txtNumRes(3).Text <> "" Then Importe = CCur(txtNumRes(3).Text)
                If txtNumRes(4).Text <> "" Then I2 = CCur(txtNumRes(4).Text)
           End If
           ProcesaLibroResumen CONT, I, Importe, I2
           Importe = 0: I2 = 0
        Next CONT
        'A�os siguiente
        I = CInt(Val(txtAno(1).Text))
        For CONT = 1 To cmbFecha(1).ListIndex + 1
           Label2(25).Caption = "Fecha: " & CONT & " / " & I
           Label2(25).Refresh
           
           DoEvents
           If PulsadoCancelar Then Exit Function
           
           ProcesaLibroResumen CONT, I, Importe, I2
        Next CONT
    End If
    
    'Vemos si ha generado datos
    Set miRsAux = New ADODB.Recordset
    SQL = "Select count(*) from tmpdirioresum where codusu =" & vUsu.Codigo
    CONT = 0
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then CONT = miRsAux.Fields(0)
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If CONT = 0 Then
        MsgBox "Ningun dato generado para estos valores.", vbExclamation
        Exit Function
    End If
    
    Label2(25).Caption = ""
    Label2(25).Refresh
    GenerarLibroResumen = True
    Exit Function
EGenerarLibroResumen:
    MuestraError Err.Number, "Generar libro resumen"
End Function

