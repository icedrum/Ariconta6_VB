VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfCtaExplo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   12960
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
      Height          =   5505
      Left            =   7110
      TabIndex        =   20
      Top             =   0
      Width           =   5745
      Begin VB.CheckBox chkPorcentajes 
         Caption         =   "Listado con porcentajes"
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
         Left            =   750
         TabIndex        =   49
         Top             =   3660
         Width           =   4665
      End
      Begin VB.CheckBox chkComparativo 
         Caption         =   "Comparar con ejercicio anterior"
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
         Left            =   240
         TabIndex        =   48
         Top             =   3330
         Width           =   4665
      End
      Begin VB.Frame Frame1 
         Caption         =   "Existencias Acumuladas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1305
         Left            =   120
         TabIndex        =   43
         Top             =   4020
         Width           =   5475
         Begin VB.TextBox txtExplo 
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
            Left            =   1350
            TabIndex        =   3
            Top             =   510
            Width           =   1845
         End
         Begin VB.TextBox txtExplo 
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
            Left            =   3240
            TabIndex        =   4
            Top             =   510
            Width           =   1875
         End
         Begin VB.TextBox txtExplo 
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
            Left            =   1350
            TabIndex        =   5
            Top             =   870
            Width           =   1845
         End
         Begin VB.TextBox txtExplo 
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
            Left            =   3240
            TabIndex        =   6
            Top             =   870
            Width           =   1875
         End
         Begin VB.Label Label9 
            Caption         =   "Acumuladas"
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
            Left            =   90
            TabIndex        =   47
            Top             =   570
            Width           =   1155
         End
         Begin VB.Label Label9 
            Caption         =   "Mes"
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
            Left            =   90
            TabIndex        =   46
            Top             =   900
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Iniciales(Debe)"
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
            Left            =   1470
            TabIndex        =   45
            Top             =   270
            Width           =   1605
         End
         Begin VB.Label Label9 
            Caption         =   "Finales(Haber)"
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
            Left            =   3480
            TabIndex        =   44
            Top             =   270
            Width           =   1545
         End
      End
      Begin VB.CheckBox chkExplotacion 
         Caption         =   "Imprimir acumulados y movimientos del mes"
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
         Left            =   240
         TabIndex        =   42
         Top             =   3000
         Width           =   4665
      End
      Begin VB.Frame FrameCtasExplo 
         Height          =   1575
         Left            =   120
         TabIndex        =   32
         Top             =   1290
         Width           =   5535
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "9º nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   9
            Left            =   1530
            TabIndex        =   41
            Top             =   1140
            Width           =   1425
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "8º nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   8
            Left            =   120
            TabIndex        =   40
            Top             =   1140
            Width           =   1365
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "7º nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   7
            Left            =   4260
            TabIndex        =   39
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "6º nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   2940
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "5º nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   1530
            TabIndex        =   37
            Top             =   720
            Width           =   1305
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "4º nivel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1275
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "3º nivel"
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
            Left            =   4260
            TabIndex        =   35
            Top             =   270
            Width           =   1215
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "2º nivel"
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
            Left            =   2940
            TabIndex        =   34
            Top             =   270
            Width           =   1245
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Left            =   1530
            TabIndex        =   33
            Top             =   270
            Width           =   1305
         End
         Begin VB.CheckBox chkCtaExplo 
            Caption         =   "Último:  "
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
            TabIndex        =   7
            Top             =   240
            Value           =   1  'Checked
            Width           =   1365
         End
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
         Left            =   2040
         TabIndex        =   2
         Top             =   780
         Width           =   1485
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   5100
         TabIndex        =   28
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
         TabIndex        =   29
         Top             =   780
         Width           =   690
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   7
         Left            =   1680
         Picture         =   "frmInfCtaExplo.frx":0000
         Top             =   810
         Width           =   240
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
      Height          =   2805
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6915
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
         Index           =   2
         ItemData        =   "frmInfCtaExplo.frx":008B
         Left            =   1740
         List            =   "frmInfCtaExplo.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1230
         Width           =   1575
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
         Height          =   360
         Index           =   4
         Left            =   810
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1230
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
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
         Index           =   0
         Left            =   840
         TabIndex        =   31
         Top             =   900
         Width           =   840
      End
      Begin VB.Label lblCuentas 
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
         Index           =   6
         Left            =   2520
         TabIndex        =   27
         Top             =   1110
         Width           =   4095
      End
      Begin VB.Label lblCuentas 
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
         Index           =   7
         Left            =   2520
         TabIndex        =   26
         Top             =   1500
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
         TabIndex        =   25
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Mes"
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
         Left            =   1770
         TabIndex        =   24
         Top             =   900
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
      Left            =   11610
      TabIndex        =   10
      Top             =   5670
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
      Left            =   10020
      TabIndex        =   8
      Top             =   5670
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
      TabIndex        =   9
      Top             =   5610
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
      Height          =   2625
      Left            =   120
      TabIndex        =   11
      Top             =   2880
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
         TabIndex        =   23
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   22
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   21
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1830
      TabIndex        =   30
      Top             =   5640
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
End
Attribute VB_Name = "frmInfCtaExplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 307

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
Private WithEvents frmCon  As frmConceptos
Attribute frmCon.VB_VarHelpID = -1

Private SQL As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim RS As ADODB.Recordset
Dim vFecha As String

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean

Public Legalizacion As String   'Datos para la legalizacion

Dim HanPulsadoSalir As Boolean

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



Private Sub chkComparativo_Click()
    chkExplotacion.Enabled = (chkComparativo.Value = 0)
    Frame1.Enabled = (chkComparativo.Value = 0)
    chkPorcentajes.Enabled = (chkComparativo.Value = 1)
    If Not Frame1.Enabled Then
        For I = 0 To txtExplo.Count - 1
            txtExplo(I).Text = ""
        Next I
        chkExplotacion.Value = 0
    Else
        chkPorcentajes.Value = 0
    End If
End Sub

Private Sub chkCtaExplo_Click(Index As Integer)
    If chkCtaExplo(Index).Value = 1 Then
        For I = 1 To 10
            If I <> Index Then chkCtaExplo(I).Value = 0
        Next I
    End If
End Sub

Private Sub chkExplotacion_Click()
    If chkExplotacion.Value = 1 Then
        chkComparativo.Enabled = False
        chkComparativo.Value = 0
        chkPorcentajes.Enabled = False
        chkPorcentajes.Value = 0
    Else
        chkComparativo.Enabled = True
        chkPorcentajes.Enabled = (chkComparativo.Value = 1)
    End If
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
'++

    If Not MontaSQL Then Exit Sub
    
    If Me.chkComparativo.Value = 0 Then
        If Not HayRegParaInforme("hlinapu", cadselect) Then Exit Sub
    Else
        If Not HayRegParaInforme("tmpbalancesumas", "codusu = " & vUsu.Codigo) Then Exit Sub
    End If
    
    
    
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
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
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
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Cuenta de Explotación"

    
    PrimeraVez = True
     
    CargarComboFecha
     
    I = 0
    txtAno(4).Text = Year(DateAdd("yyyy", I, Now))
    cmbFecha(2).ListIndex = Month(DateAdd("yyyy", I, Now)) - 1
     
    txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
    
    chkPorcentajes.Enabled = (chkComparativo.Value = 1)
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
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




Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
    
End Sub




Private Sub AccionesCSV()
Dim SQL2 As String
Dim Tipo As Byte
Dim cad As String

    On Error GoTo eAccionesCSV


    ' cuenta de explotacion normal
    If chkComparativo.Value = 0 Then

        pb2.Visible = True
        CargarProgres pb2, 6
        
        'Cargamos una temporal para poder exportar a fichero
        SQL = "delete from tmpbalancesumas where codusu = " & vUsu.Codigo
        Conn.Execute SQL
    
        IncrementarProgres pb2, 1
        
        cad = "01/" & cmbFecha(2).ListIndex + 1 & "/" & txtAno(4).Text
        
        SQL = "insert into tmpbalancesumas (codusu,cta,nomcta,acumAntD,acumAntH) "
        SQL = SQL & " select " & vUsu.Codigo & ", hlinapu.codmacta Cuenta , nommacta Titulo, sum(coalesce(timported,0)), sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu left join cuentas on hlinapu.codmacta = cuentas.codmacta where hlinapu.codconce<>970 AND mid(hlinapu.codmacta,1,1) IN ('6','7') and hlinapu.fechaent >= '" & Format(Me.txtAno(4).Text, "0000") & "-" & Format(Month(vParam.fechaini), "00") & "-" & Format(Day(vParam.fechaini), "00") & "'"
        SQL = SQL & " and hlinapu.fechaent < " & DBSet(cad, "F")
        SQL = SQL & " group by 1,2,3 "
        SQL = SQL & " order by 1,2,3 "
            
        Conn.Execute SQL
        
        IncrementarProgres pb2, 1
        
        
        I = DiasMes(cmbFecha(2).ListIndex + 1, txtAno(4).Text)
        
        SQL = "update tmpbalancesumas set "
        SQL = SQL & " acumperd = (select sum(coalesce(timported,0)) from hlinapu where fechaent between " & DBSet(cad, "F") & " and "
        SQL = SQL & " '" & txtAno(4).Text & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(I, "00") & "' and hlinapu.codmacta = tmpbalancesumas.cta)"
        
        Conn.Execute SQL
        
        IncrementarProgres pb2, 1
        
        SQL = "update tmpbalancesumas set "
        SQL = SQL & " acumperh = (select sum(coalesce(timporteh,0)) from hlinapu where fechaent between " & DBSet(cad, "F") & " and "
        SQL = SQL & " '" & txtAno(4).Text & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(I, "00") & "' and hlinapu.codmacta = tmpbalancesumas.cta)"
        
        Conn.Execute SQL
        
        IncrementarProgres pb2, 1
        
        'existencias iniciales
        If txtExplo(0).Text <> "" Or txtExplo(2).Text <> "" Then
            SQL = "insert into tmpbalancesumas (codusu,cta,nomcta,acumAntD,acumAntH,acumPerD,acumPerH) values ( "
            SQL = SQL & vUsu.Codigo & ",'0000000000','Existencias Iniciales'," & DBSet(ComprobarCero(txtExplo(0).Text), "N") & ",0," & DBSet(ComprobarCero(txtExplo(2).Text), "N") & ",0) "
        
            Conn.Execute SQL
        End If
        
        IncrementarProgres pb2, 1
        
        'existencias finales
        If txtExplo(1).Text <> "" Or txtExplo(3).Text <> "" Then
            SQL = "insert into tmpbalancesumas (codusu,cta,nomcta,acumAntD,acumAntH,acumPerD,acumPerH) values ( "
            SQL = SQL & vUsu.Codigo & ",'9999999999','Existencias Finales',0," & DBSet(ComprobarCero(txtExplo(1).Text), "N") & ",0," & DBSet(ComprobarCero(txtExplo(3).Text), "N") & ") "
        
            Conn.Execute SQL
        End If
        
        IncrementarProgres pb2, 1
        
        If chkExplotacion.Value = 1 Then
            SQL = "select if(cta in ('0000000000','9999999999'),'',cta) Cuenta, nomcta Titulo, coalesce(acumantd,0) AcumAntD, coalesce(acumanth,0) AcumAntH, coalesce(acumperd,0) AcumPerD, coalesce(acumperh,0) AcumPerH from tmpbalancesumas "
        Else
            SQL = "select if(cta in ('0000000000','9999999999'),'',cta) Cuenta, nomcta Titulo, (coalesce(acumantd,0) + coalesce(acumperd,0)) SaldoD, (coalesce(acumanth,0) + coalesce(acumperH,0)) SaldoH from tmpbalancesumas "
        End If
        SQL = SQL & " where codusu = " & vUsu.Codigo
        SQL = SQL & " and (coalesce(acumantd,0) + coalesce(acumanth,0) + coalesce(acumperd,0) + coalesce(acumperh,0)) <> 0"
        SQL = SQL & " order by cta "
    
        pb2.Visible = False
    
    
    Else ' cuenta de explotacion comparativa
        If CargarTablaTemporal(vFecha) Then
            SQL = "select aaaaa.cta CtaPasivo, aaaaa.nomcta Titulo, aaaaa.totald '" & Format(CInt(txtAno(4).Text) - 1, "0000") & "', aaaaa.totalh '" & Format(CInt(txtAno(4).Text), "0000") & "'"
            SQL = SQL & ", bbbbb.cta CtaPasivo, bbbbb.nomcta Titulo, bbbbb.totald '" & Format(CInt(txtAno(4).Text) - 1, "0000") & "', bbbbb.totalh '" & Format(CInt(txtAno(4).Text), "0000") & "'"
            SQL = SQL & " from tmpbalancesumas aaaaa, tmpbalancesumas bbbbb "
            SQL = SQL & " where aaaaa.codusu = " & vUsu.Codigo & " and bbbbb.codusu = " & vUsu.Codigo
            SQL = SQL & " order by aaaaa.cta, bbbbb.cta "
            
            If Me.chkPorcentajes.Value = 0 Then
                SQL = "select aaaaa.cta Cuenta, aaaaa.nomcta Titulo, aaaaa.totald '" & Format(CInt(txtAno(4).Text) - 1, "0000") & "', aaaaa.totalh '" & Format(CInt(txtAno(4).Text), "0000") & "'"
            Else
                SQL = "select aaaaa.cta Cuenta, aaaaa.nomcta Titulo, aaaaa.totald '" & Format(CInt(txtAno(4).Text) - 1, "0000") & "', aaaaa.totalh '" & Format(CInt(txtAno(4).Text), "0000") & "', round(aaaaa.totalh / aaaaa.totald * 100,2) -100 Porcentaje"
            End If
            SQL = SQL & " from tmpbalancesumas aaaaa "
            SQL = SQL & " where aaaaa.codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1"
        End If
    End If
    
        
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
    Exit Sub
    
eAccionesCSV:
    pb2.Visible = False
    MuestraError Err.Number, "Carga de Temporal", Err.Description
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    
    vMostrarTree = False
    conSubRPT = False
        
    'Parametros
    cadParam = cadParam & "Digitos=" & CONT & "|"
    cad = "01/" & cmbFecha(2).ListIndex + 1 & "/" & txtAno(4).Text
    cadParam = cadParam & "FechaPeriodo=""" & cad & """|"
    
    numParam = numParam + 2
    'Existencias iniciales y finales
    cad = "InicioAcumulada=" & DBSet(txtExplo(0).Text, "N") & "|InicioPeriodo=" & DBSet(txtExplo(2).Text, "N")
    cadParam = cadParam & cad & "|"
    cad = "FinAcumulada=" & DBSet(txtExplo(1).Text, "N") & "|FinPeriodo=" & DBSet(txtExplo(3).Text, "N")
    cadParam = cadParam & cad & "|"
    numParam = numParam + 4
    
    
    cadParam = cadParam & "pTipo=" & chkExplotacion.Value & "|"
    numParam = numParam + 1
    cadParam = cadParam & "pPeriodo=""Mes cálculo: " & UCase(Mid(cmbFecha(2).Text, 1, 1)) & Mid(cmbFecha(2).Text, 2, Len(cmbFecha(2).Text)) & "     Año: " & txtAno(4).Text & """|"
    numParam = numParam + 1
    
    If Me.chkComparativo = 1 Then
        cadParam = cadParam & "Anyo1=""" & Format(CInt(txtAno(4).Text) - 1, "0000") & """|"
        cadParam = cadParam & "Anyo2=""" & Format(txtAno(4).Text, "0000") & """|"
        
        cadParam = cadParam & "pPorcen=" & chkPorcentajes.Value & "|"
        cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
        numParam = numParam + 4
        
        indRPT = "0307-01" '"CtaExplotacionComp.rpt"
        
        If Me.chkPorcentajes.Value = 1 Then indRPT = "0307-02" ' "CtaExplotacionComp1.rpt"
        
        cadFormula = "{tmpbalancesumas.codusu} = " & vUsu.Codigo
    Else
        indRPT = "0307-00" '"CtaExplotacion.rpt"
    End If

    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault

End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String
Dim Anyo As String

    MontaSQL = False
    
    I = cmbFecha(2).ListIndex + 1
    If I >= Month(vParam.fechaini) Then
        Anyo = Val(txtAno(4).Text)
    Else
        Anyo = Val(txtAno(4).Text) - 1
    End If
    vFecha = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & Anyo
    
    cadFormula = "{hlinapu.codconce}<>970 AND mid({hlinapu.codmacta},1,1) IN [" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & "]" ''6','7']"
    cadselect = "hlinapu.codconce<>970 AND mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & ")" ''6','7')"
    'Montamos la fecha de inicio del periodo solicitado
    cadFormula = cadFormula & " AND {hlinapu.fechaent} >= Date (" & Me.txtAno(4).Text & "," & Month(vParam.fechaini) & "," & Day(vParam.fechaini) & ") "
    cadselect = cadselect & " AND hlinapu.fechaent >= '" & Format(Me.txtAno(4).Text, "0000") & "-" & Format(Month(vParam.fechaini), "00") & "-" & Format(Day(vParam.fechaini), "00") & "'"
    I = DiasMes(cmbFecha(2).ListIndex + 1, CInt(txtAno(4).Text))
    cadFormula = cadFormula & " AND {hlinapu.fechaent} <= Date (" & Me.txtAno(4).Text & ", " & cmbFecha(2).ListIndex + 1 & "," & I & ")  "
    cadselect = cadselect & " AND hlinapu.fechaent <= '" & Format(Me.txtAno(4).Text, "0000") & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(I, "00") & "'"
    
    If chkComparativo.Value = 1 Then
       Screen.MousePointer = vbHourglass
       
       CargarTablaTemporal (vFecha)
       
       Screen.MousePointer = vbDefault
    End If
    
    
    MontaSQL = True
           
End Function

Private Function CargarTablaTemporal(Fecha As String) As Boolean
Dim SQL As String
Dim FechaI As Date
Dim FechaF As Date
Dim FechaIAnt As Date
Dim FechaFAnt As Date
Dim I As Integer
Dim CADENA As String
Dim Digitos As Integer

    On Error GoTo eCargarTablaTemporal
        
        
    CargarTablaTemporal = False

    pb2.Visible = True
    CargarProgres pb2, 7



    'Cargamos las fechas de calculo
    FechaI = CDate(Fecha)
    I = DiasMes(cmbFecha(2).ListIndex + 1, CInt(txtAno(4).Text))
    FechaF = CDate(I & "/" & cmbFecha(2).ListIndex + 1 & "/" & Me.txtAno(4).Text)
    
    FechaIAnt = DateAdd("yyyy", -1, FechaI)
    FechaFAnt = DateAdd("yyyy", -1, FechaF)
    
    For I = 1 To 10
        If Me.chkCtaExplo(I).Visible Then
            If Me.chkCtaExplo(I).Value = 1 Then
                If I < 10 Then
                    Digitos = DigitosNivel(I)
                Else
                    Digitos = vEmpresa.DigitosUltimoNivel
                End If
            End If
        End If
    Next I
    
    CADENA = String(Digitos, "_")
    
    'Borramos los temporales
    SQL = "DELETE from tmpbalancesumas where codusu= " & vUsu.Codigo
    Conn.Execute SQL

    IncrementarProgres pb2, 1

    'Metemos las cuentas
    SQL = "insert into tmpbalancesumas (codusu, cta, nomcta, totald, totalh) "
    SQL = SQL & "select distinct " & vUsu.Codigo & ", mid(hlinapu.codmacta,1," & Digitos & ") , cuentas.nommacta, 0, 0 "
    SQL = SQL & " from hlinapu inner join cuentas on mid(hlinapu.codmacta,1," & Digitos & ") = cuentas.codmacta where hlinapu.codconce<>970 AND mid(hlinapu.codmacta,1,1) IN ('6','7') "
    SQL = SQL & " and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaF, "F")
    
    
    Conn.Execute SQL
    
    IncrementarProgres pb2, 1
    
    'Actualizamos acumulados del periodo actual
    'activo
    SQL = "update tmpbalancesumas set "
    SQL = SQL & " totalh = (select sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) from hlinapu where hlinapu.codconce<>970  and fechaent between " & DBSet(FechaI, "F") & " and " & DBSet(FechaF, "F")
    SQL = SQL & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
    SQL = SQL & " where mid(tmpbalancesumas.cta,1,1) = '6' "
    SQL = SQL & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
    
    Conn.Execute SQL
    
    IncrementarProgres pb2, 1
    
    'pasivo
    SQL = "update tmpbalancesumas set "
    SQL = SQL & " totalh = (select sum(coalesce(timporteh,0)) - sum(coalesce(timported,0)) from hlinapu where hlinapu.codconce<>970  and fechaent between " & DBSet(FechaI, "F") & " and " & DBSet(FechaF, "F")
    SQL = SQL & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
    SQL = SQL & " where mid(tmpbalancesumas.cta,1,1) = '7' "
    SQL = SQL & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
    
    Conn.Execute SQL
    
    IncrementarProgres pb2, 1
    
    
    'Actualizamos acumulados del periodo anterior
    'activo
    SQL = "update tmpbalancesumas set "
    SQL = SQL & " totald = (select sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) from hlinapu where hlinapu.codconce<>970  and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaFAnt, "F")
    SQL = SQL & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
    SQL = SQL & " where mid(tmpbalancesumas.cta,1,1) = '6' "
    SQL = SQL & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
    
    Conn.Execute SQL
    
    IncrementarProgres pb2, 1
        
    'pasivo
    SQL = "update tmpbalancesumas set "
    SQL = SQL & " totald = (select sum(coalesce(timporteh,0)) - sum(coalesce(timported,0)) from hlinapu where hlinapu.codconce<>970  and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaFAnt, "F")
    SQL = SQL & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
    SQL = SQL & " where mid(tmpbalancesumas.cta,1,1) = '7' "
    SQL = SQL & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
    
    Conn.Execute SQL
        
    IncrementarProgres pb2, 1
    
    'borramos las cuentas que no tienen movimientos en ese periodo
    SQL = "delete from tmpbalancesumas where codusu = " & vUsu.Codigo & " and totald is null and totalh is null"
    Conn.Execute SQL
    
    IncrementarProgres pb2, 1
    
    CargarTablaTemporal = True
        
    pb2.Visible = False
    Exit Function
        
        
eCargarTablaTemporal:
    pb2.Visible = False
    MuestraError Err.Number, "Cargando tabla temporal", Err.Description
End Function


Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAno_LostFocus(Index As Integer)
    txtAno(Index).Text = Trim(txtAno(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtAno(Index).Text = Trim(txtAno(Index).Text)
    
    If Not IsNumeric(txtAno(Index).Text) Then
        MsgBox "El Año debe ser numérico: " & txtAno(Index).Text, vbExclamation
        txtAno(Index).Text = ""
        Exit Sub
    End If

End Sub

Private Sub txtExplo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

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
    
    'Mes de cálculo
    If cmbFecha(2).ListIndex < 0 Then
        MsgBox "Seleccion un mes para el cálculo.", vbExclamation
        Exit Function
    End If
    
    ' Uno y solo uno de los niveles tiene que estar marcado
    cad = ""
    For I = 1 To 10
        If Me.chkCtaExplo(I).Visible Then
            If Me.chkCtaExplo(I).Value = 1 Then
                If I < 10 Then
                    CONT = DigitosNivel(I)
                Else
                    CONT = vEmpresa.DigitosUltimoNivel
                End If
                cad = cad & "1"
            End If
        End If
    Next I
    If Len(cad) <> 1 Then
        MsgBox "Seleccione uno(y solo uno) de los niveles para el informe.", vbExclamation
        Exit Function
    End If
    
    If txtAno(4).Text = "" Then
        MsgBox "Ponga el año para el listado.", vbExclamation
        Exit Function
    End If

    If vParam.grupoord <> "" And vParam.Automocion <> "" Then
        If CDate("01/" & cmbFecha(2).ListIndex + 1 & "/" & txtAno(4).Text) > vParam.fechafin Then
            'Ha seleccionado a uno o dos digitos
            If chkCtaExplo(1).Value = 1 Or chkCtaExplo(2).Value = 1 Then
                MsgBox "La cuenta de exclusión del grupoord de la analítica no esta incluída en el balance", vbExclamation
            End If
        End If
    End If

    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer

    QueCombosFechaCargar "2|"
    
    
    'Y ademas deshabilitamos los niveles no utilizados por la aplicacion
    For I = vEmpresa.numnivel To 9
        Me.chkCtaExplo(I).Visible = False
    Next I
    
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        chkCtaExplo(I).Visible = True
        chkCtaExplo(I).Caption = "Digitos: " & J
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



Private Sub txtExplo_GotFocus(Index As Integer)
    PonFoco txtExplo(Index)
End Sub

Private Sub txtExplo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtExplo_LostFocus(Index As Integer)
    txtExplo(Index).Text = Trim(txtExplo(Index).Text)
    If txtExplo(Index).Text = "" Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoDecimal txtExplo(Index), 3

End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
