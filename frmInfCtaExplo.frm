VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfCtaExplo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   13050
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
      TabIndex        =   33
      Top             =   0
      Width           =   5865
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   46
         Top             =   4020
         Width           =   5655
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
         TabIndex        =   13
         Top             =   3000
         Width           =   4665
      End
      Begin VB.Frame FrameCtasExplo 
         Height          =   1575
         Left            =   120
         TabIndex        =   45
         Top             =   1290
         Width           =   5655
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   9
            Left            =   1530
            TabIndex        =   12
            Top             =   1140
            Width           =   1425
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   8
            Left            =   120
            TabIndex        =   11
            Top             =   1140
            Width           =   1425
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   7
            Left            =   4260
            TabIndex        =   10
            Top             =   720
            Width           =   1305
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   6
            Left            =   2940
            TabIndex        =   9
            Top             =   720
            Width           =   1425
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   5
            Left            =   1530
            TabIndex        =   8
            Top             =   720
            Width           =   1425
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   4
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   1425
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   3
            Left            =   4260
            TabIndex        =   6
            Top             =   270
            Width           =   1305
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   2
            Left            =   2940
            TabIndex        =   5
            Top             =   270
            Width           =   1425
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
            Height          =   300
            Index           =   1
            Left            =   1530
            TabIndex        =   4
            Top             =   270
            Width           =   1425
         End
         Begin VB.CheckBox chkCtaExplo 
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
            Height          =   300
            Index           =   10
            Left            =   120
            TabIndex        =   3
            Top             =   270
            Value           =   1  'Checked
            Width           =   1425
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
         TabIndex        =   41
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
         TabIndex        =   42
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
      Height          =   2805
      Left            =   120
      TabIndex        =   32
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
         Index           =   0
         ItemData        =   "frmInfCtaExplo.frx":008B
         Left            =   600
         List            =   "frmInfCtaExplo.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1230
         Width           =   1095
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
         Index           =   2
         ItemData        =   "frmInfCtaExplo.frx":008F
         Left            =   1740
         List            =   "frmInfCtaExplo.frx":0091
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
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1230
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "A�o"
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
         Left            =   600
         TabIndex        =   44
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
         TabIndex        =   40
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
         Left            =   2640
         TabIndex        =   39
         Top             =   2400
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
         TabIndex        =   38
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
         TabIndex        =   37
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
      Left            =   11730
      TabIndex        =   22
      Top             =   5640
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
      Left            =   10140
      TabIndex        =   20
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
      TabIndex        =   21
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
      TabIndex        =   24
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
         TabIndex        =   36
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   35
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   34
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1830
      TabIndex        =   43
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
Dim Cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim Rs As ADODB.Recordset
Dim vFecha As String

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean

Public Legalizacion As String   'Datos para la legalizacion
Dim DeTantosEnTantos As Integer
Dim HanPulsadoSalir As Boolean

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



Private Sub chkComparativo_Click()
    'Febrero 2022
    'Puede hacer de todo
    'chkExplotacion.Enabled = (chkComparativo.Value = 0)
    Frame1.Enabled = (chkComparativo.Value = 0)
    'chkPorcentajes.Enabled = (chkComparativo.Value = 1)
    If Not Frame1.Enabled Then
        For i = 0 To txtExplo.Count - 1
            txtExplo(i).Text = ""
        Next i
        'chkExplotacion.Value = 0
    Else
        'chkPorcentajes.Value = 0
    End If
End Sub

Private Sub chkCtaExplo_Click(Index As Integer)
    If chkCtaExplo(Index).Value = 1 Then
        For i = 1 To 10
            If i <> Index Then chkCtaExplo(i).Value = 0
        Next i
    End If
End Sub

Private Sub chkExplotacion_Click()

    'Febrero 2022.  Puede hacer de todo
    'If chkExplotacion.Value = 1 Then
    '    chkComparativo.Enabled = False
    '    chkComparativo.Value = 0
    '    chkPorcentajes.Enabled = False
    '    chkPorcentajes.Value = 0
    'Else
    '    chkComparativo.Enabled = True
    '    chkPorcentajes.Enabled = (chkComparativo.Value = 1)
    'End If
End Sub

Private Sub cmbFecha_Change(Index As Integer)
    If Index = 0 Then
        txtAno(4).Text = cmbFecha(0).Text
    End If
End Sub

Private Sub cmbFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub cmbFecha_LostFocus(Index As Integer)
    If Index = 2 Then PonerFocoBtn Me.cmdAccion(1)
End Sub

Private Sub cmbFecha_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        txtAno(4).Text = cmbFecha(0).Text
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
    
    
   Screen.MousePointer = vbHourglass
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
    
    lblCuentas(7).Caption = ""
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    lblCuentas(7).Caption = ""
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
    Me.Icon = frmppal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Cuenta de Explotaci�n"

    
    PrimeraVez = True
     
    CargarComboFecha
    DeTantosEnTantos = -1 'La primera vez que lance un informe
    CargarComboA�o cmbFecha(0)
    PosicionarCombo cmbFecha(0), Year(DateAdd("yyyy", 0, Now))
    i = 0
    txtAno(4).Text = Year(DateAdd("yyyy", i, Now))
    cmbFecha(2).ListIndex = Month(DateAdd("yyyy", i, Now)) - 1
     
    txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
    
    'chkPorcentajes.Enabled = (chkComparativo.Value = 1)
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
        
    lblCuentas(7).Caption = ""
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




Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
    
End Sub




Private Sub AccionesCSV()
Dim Sql2 As String
Dim Tipo As Byte
Dim Cad As String

    On Error GoTo eAccionesCSV


    ' cuenta de explotacion normal
    If chkComparativo.Value = 0 Then


        'Sin comparativo
        'NORMAL
        pb2.visible = True
        CargarProgres pb2, 8
        
        'Cargamos una temporal para poder exportar a fichero
        SQL = "delete from tmpbalancesumas where codusu = " & vUsu.Codigo
        Conn.Execute SQL
    
        IncrementarProgres pb2, 1
        
        Cad = "01/" & Format(cmbFecha(2).ListIndex + 1, "00") & "/" & txtAno(4).Text
        
        SQL = "insert into tmpbalancesumas (codusu,cta,nomcta,acumAntD,acumAntH) "
        SQL = SQL & " select " & vUsu.Codigo & ", hlinapu.codmacta Cuenta , nommacta Titulo, sum(coalesce(timported,0)), sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu left join cuentas on hlinapu.codmacta = cuentas.codmacta where hlinapu.codconce<>960 AND mid(hlinapu.codmacta,1,1) IN ('6','7')"    'PONIA 970
          i = cmbFecha(2).ListIndex + 1
        If i >= Month(vParam.fechaini) Then
            i = Val(txtAno(4).Text)
        Else
            i = Val(txtAno(4).Text) - 1
        End If
        
        SQL = SQL & " and hlinapu.fechaent >= '" & Format(i, "0000") & "-" & Format(Month(vParam.fechaini), "00") & "-" & Format(Day(vParam.fechaini), "00") & "'"
        SQL = SQL & " and hlinapu.fechaent < " & DBSet(Cad, "F")
        SQL = SQL & " group by 1,2,3 "
        
        SQL = SQL & " order by 1,2,3 "
            
        Conn.Execute SQL
        
        IncrementarProgres pb2, 1
        
        'Otro ERROR. EN el update de abajo, solo updatea los que hayan.
        'Si hubiera alguna cuenta que tuviera movimientos del periodo , pero NO anteriores.... EEERROR
        
        i = DiasMes(cmbFecha(2).ListIndex + 1, txtAno(4).Text)
        
        
        Set miRsAux = New ADODB.Recordset
'
'        SQL = "update tmpbalancesumas set "
'        SQL = SQL & " acumperd = (select sum(coalesce(timported,0)) from hlinapu where fechaent between " & DBSet(Cad, "F") & " and "
'        SQL = SQL & " '" & txtAno(4).Text & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(i, "00") & "' and hlinapu.codmacta = tmpbalancesumas.cta AND hlinapu.codconce<>960)"
'        SQL = SQL & " WHERE codusu = " & vUsu.Codigo
'
'
'        Conn.Execute SQL
'
'
'        SQL = SQL & " acumperd = (select sum(coalesce(timported,0)) from hlinapu where fechaent between " & DBSet(Cad, "F") & " and "
'        SQL = SQL & " '" & txtAno(4).Text & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(i, "00") & "' and hlinapu.codmacta = tmpbalancesumas.cta AND hlinapu.codconce<>960)"
'        SQL = SQL & " WHERE codusu = " & vUsu.Codigo
'
        
        
        SQL = "select codmacta,sum(coalesce(timported,0)) de,sum(coalesce(timporteh,0)) ha from hlinapu,tmpbalancesumas where hlinapu.codmacta = tmpbalancesumas.cta AND fechaent between " & DBSet(Cad, "F") & " and "
        SQL = SQL & " '" & txtAno(4).Text & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(i, "00") & "'  AND hlinapu.codconce<>960 AND  codusu = " & vUsu.Codigo
        SQL = SQL & " GROUP BY codmacta"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""

        While Not miRsAux.EOF
            SQL = SQL & "UPDATE tmpbalancesumas set acumperd=" & DBSet(miRsAux!De, "N") & " , acumperh=" & DBSet(miRsAux!Ha, "N") & " WHERE codusu =" & vUsu.Codigo & " AND cta =" & DBSet(miRsAux!codmacta, "T") & "; "
            Conn.Execute SQL
            
            SQL = ""
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        

'        SQL = "update tmpbalancesumas set "
'        SQL = SQL & " acumperh = (select sum(coalesce(timporteh,0)) from hlinapu where fechaent between " & DBSet(Cad, "F") & " and "
'        SQL = SQL & " '" & txtAno(4).Text & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(i, "00") & "' and hlinapu.codmacta = tmpbalancesumas.cta  AND hlinapu.codconce<>96)"
'        SQL = SQL & " WHERE codusu = " & vUsu.Codigo
'
'        Conn.Execute SQL


        SQL = "UPDATE tmpbalancesumas set acumperh=0 where acumperh is null and codusu = " & vUsu.Codigo
        Conn.Execute SQL
        SQL = "UPDATE tmpbalancesumas set acumperd=0 where acumperd is null and codusu = " & vUsu.Codigo
        Conn.Execute SQL
        IncrementarProgres pb2, 1
        
        'Para subsanar el error anterior, de que ctas del periodo que NO esten en anteriores
        IncrementarProgres pb2, 1
        
        SQL = "insert into tmpbalancesumas (codusu,cta,nomcta,acumAntD,acumAntH,acumPerD ,acumPerH) "
        SQL = SQL & " select " & vUsu.Codigo & ", hlinapu.codmacta Cuenta , nommacta Titulo,0,0, sum(coalesce(timported,0)), sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu left join cuentas on hlinapu.codmacta = cuentas.codmacta where hlinapu.codconce<>960 AND mid(hlinapu.codmacta,1,1) IN ('6','7')"   'ponia 970
        SQL = SQL & " AND fechaent between " & DBSet(Cad, "F") & " and "
        SQL = SQL & " '" & txtAno(4).Text & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(i, "00") & "'"
        SQL = SQL & " and not hlinapu.codmacta In (select cta from tmpbalancesumas WHERE codusu =" & vUsu.Codigo & " )"
        SQL = SQL & " group by 1,2,3 "
        SQL = SQL & " order by 1,2,3 "
        Conn.Execute SQL
        
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
        
        'Diciembre 2017
        If Not Me.chkCtaExplo(10).Value = 1 Then
            'NO es a ultimi nivel
            SQL = ""
            For i = 1 To 9
                If chkCtaExplo(i).Value Then
                    SQL = i
                    Exit For
                End If
            Next
            If SQL = "" Then Err.Raise 513, , "Error obteniendo nivel"
            
            i = CInt(SQL)
            
            SQL = "select mid(cta,1," & i & "),'' nommacta,"
            SQL = SQL & "  sum(coalesce(acumantd,0)) , sum(coalesce(acumanth,0))"
            SQL = SQL & " , sum(coalesce(acumPerD, 0)), sum(coalesce(acumPerH,0)),codusu"
            SQL = SQL & " from tmpbalancesumas  where codusu = " & vUsu.Codigo
            SQL = SQL & " group by 1"
            SQL = "INSERT INTO tmpbalancesumas(cta,nomcta,acumAntD,acumAntH,acumPerD,acumPerH,codusu) " & SQL
            Conn.Execute SQL
            
            'Borramos ultimo nivel
            SQL = "DELETE from tmpbalancesumas where length(cta)=" & vEmpresa.DigitosUltimoNivel & " AND codusu =" & vUsu.Codigo
            Conn.Execute SQL
            espera 0.25
            
            'Updateamos titulo
            SQL = "UPDATE tmpbalancesumas,cuentas set nomcta=nommacta WHERE codusu = " & vUsu.Codigo
            SQL = SQL & " AND cta = codmacta"
            Conn.Execute SQL
            
            
            
            
            
            
        End If
        
        'VA POR SALDOS
        
        
        
        IncrementarProgres pb2, 1
        SQL = "update  tmpbalancesumas set TotalD =(acumAntD +acumPerD)-(acumAntH +acumPerH),TotalH=0"
        SQL = SQL & " Where (acumAntD + acumPerD) > (acumAntH + acumPerH) AND codusu =" & vUsu.Codigo
        Conn.Execute SQL

        SQL = "update  tmpbalancesumas set Totalh =(acumAnth +acumPerh)-(acumAntd +acumPerd),TotalD=0"
        SQL = SQL & " Where (acumAntD + acumPerD) < (acumAntH + acumPerH) AND codusu =" & vUsu.Codigo
        Conn.Execute SQL
        espera 0.3
        
        SQL = "UPDATE tmpbalancesumas set TotalD= 0 WHERE TotalD is null and codusu = " & vUsu.Codigo
        Conn.Execute SQL
        SQL = "UPDATE tmpbalancesumas set Totalh=0 where Totalh is null and codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
        
        
        'Para la cadena consulta
        If chkExplotacion.Value = 1 Then
            SQL = "select if(cta in ('0000000000','9999999999'),'',cta) Cuenta, nomcta Titulo,"
            SQL = SQL & " acumantd,  acumanth AcumAntH, acumperd AcumPerD, acumperh AcumPerH, totalD, totalH from tmpbalancesumas "
        Else
            SQL = "select if(cta in ('0000000000','9999999999'),'',cta) Cuenta, nomcta Titulo, totalD SaldoD, totalH SaldoH from tmpbalancesumas "
        End If
        SQL = SQL & " where codusu = " & vUsu.Codigo
        SQL = SQL & " and (coalesce(acumantd,0) + coalesce(acumanth,0) + coalesce(acumperd,0) + coalesce(acumperh,0)) <> 0"
        SQL = SQL & " order by cta "
    
        pb2.visible = False
    
    
    Else
        
                
            i = CInt(txtAno(4).Text)
            cadParam = ""
            If Me.chkExplotacion.Value = 0 Then cadParam = "hasta_"
            cadParam = cadParam & Format(cmbFecha(2).ListIndex + 1, "00") & "_"
            SQL = "select cta,nomcta,totald " & cadParam & CStr(i - 1) & ", totalh " & cadParam & i
            If Me.chkPorcentajes.Value = 1 Then SQL = SQL & ", if(totald =0,0, round((((totalh-totald) / totald))*100,2) )    porcen"
            SQL = SQL & " FROM tmpbalancesumas WHERE codusu=" & vUsu.Codigo & " ORDER BY cta"
            
    End If
    
        
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
    Exit Sub
    
eAccionesCSV:
    pb2.visible = False
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
    Cad = "01/" & cmbFecha(2).ListIndex + 1 & "/" & txtAno(4).Text
    cadParam = cadParam & "FechaPeriodo=""" & Cad & """|"
    
    numParam = numParam + 2
    'Existencias iniciales y finales
    Cad = "InicioAcumulada=" & DBSet(txtExplo(0).Text, "N") & "|InicioPeriodo=" & DBSet(txtExplo(2).Text, "N")
    cadParam = cadParam & Cad & "|"
    Cad = "FinAcumulada=" & DBSet(txtExplo(1).Text, "N") & "|FinPeriodo=" & DBSet(txtExplo(3).Text, "N")
    cadParam = cadParam & Cad & "|"
    numParam = numParam + 4
    
    
    
    
    
    cadParam = cadParam & "pTipo=" & chkExplotacion.Value & "|"
    numParam = numParam + 1
    cadParam = cadParam & "pPeriodo=""Mes c�lculo: " & UCase(Mid(cmbFecha(2).Text, 1, 1)) & Mid(cmbFecha(2).Text, 2, Len(cmbFecha(2).Text)) & "     A�o: " & txtAno(4).Text & """|"
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
        
        
        cadParam = cadParam & "SoloPeriodo=" & chkExplotacion.Value & "|"
        numParam = numParam + 4
        
        
    Else
        indRPT = "0307-00" '"CtaExplotacion.rpt"
    End If

    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 27
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault

End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim Anyo As String

    MontaSQL = False
    
    i = cmbFecha(2).ListIndex + 1
    If i >= Month(vParam.fechaini) Then
        Anyo = Val(txtAno(4).Text)
    Else
        Anyo = Val(txtAno(4).Text) - 1
    End If
    vFecha = Format(Day(vParam.fechaini), "00") & "/" & Month(vParam.fechaini) & "/" & Anyo
    
    cadFormula = "{hlinapu.codconce}<>960 AND mid({hlinapu.codmacta},1,1) IN [" & DBSet(vParam.GrupoGto, "T") & "," & DBSet(vParam.GrupoVta, "T") & "]" ''6','7']"
    cadselect = "hlinapu.codconce<>960 AND mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.GrupoGto, "T") & "," & DBSet(vParam.GrupoVta, "T") & ")" ''6','7')"
    
    
    'Montamos la fecha de inicio del periodo solicitado
    'Estaba este
    'cadFormula = cadFormula & " AND {hlinapu.fechaent} >= Date (" & Me.txtAno(4).Text & "," & Month(vParam.fechaini) & "," & Day(vParam.fechaini) & ")     "
    'cadselect = cadselect & " AND hlinapu.fechaent >= '" & Format(Me.txtAno(4).Text, "0000") & "-" & Format(Month(vParam.fechaini), "00") & "-" & Format(Day(vParam.fechaini), "00") & "'"
    
    
    'Ponemos este
    cadFormula = cadFormula & " AND {hlinapu.fechaent} >= Date (" & Year(vFecha) & "," & Month(vParam.fechaini) & "," & Day(vParam.fechaini) & ")     "
    cadselect = cadselect & " AND hlinapu.fechaent >= '" & Format(Year(vFecha), "0000") & "-" & Format(Month(vParam.fechaini), "00") & "-" & Format(Day(vParam.fechaini), "00") & "'"
    
    
    
    i = DiasMes(cmbFecha(2).ListIndex + 1, CInt(txtAno(4).Text))
    cadFormula = cadFormula & " AND {hlinapu.fechaent} <= Date (" & Me.txtAno(4).Text & ", " & cmbFecha(2).ListIndex + 1 & "," & i & ")  "
    cadselect = cadselect & " AND hlinapu.fechaent <= '" & Format(Me.txtAno(4).Text, "0000") & "-" & Format(cmbFecha(2).ListIndex + 1, "00") & "-" & Format(i, "00") & "'"
    
    
    
    
    If chkComparativo.Value = 1 Then
       Screen.MousePointer = vbHourglass
       
       CargarTablaTemporal (vFecha)
       
       Screen.MousePointer = vbDefault
    End If
    
    
    MontaSQL = True
           
End Function

'Private Function CargarTablaTemporalMonica(Fecha As String) As Boolean
'Dim Sql As String
'Dim FechaI As Date
'Dim FechaF As Date
'Dim FechaIAnt As Date
'Dim FechaFAnt As Date
'Dim I As Integer
'Dim CADENA As String
'Dim Digitos As Integer
'Dim ConceptoPerdidasyGanancias As Integer
'
'
'    ConceptoPerdidasyGanancias = 960 'ponia 970 en el sqql de abjao 970
'
'    On Error GoTo eCargarTablaTemporal
'
'
'    CargarTablaTemporal = False
'
'    pb2.visible = True
'    CargarProgres pb2, 7
'
'
'
'    'Cargamos las fechas de calculo
'    FechaI = CDate(Fecha)
'    I = DiasMes(cmbFecha(2).ListIndex + 1, CInt(txtAno(4).Text))
'    FechaF = CDate(I & "/" & cmbFecha(2).ListIndex + 1 & "/" & Me.txtAno(4).Text)
'
'    FechaIAnt = DateAdd("yyyy", -1, FechaI)
'    FechaFAnt = DateAdd("yyyy", -1, FechaF)
'
'    For I = 1 To 10
'        If Me.chkCtaExplo(I).visible Then
'            If Me.chkCtaExplo(I).Value = 1 Then
'                If I < 10 Then
'                    Digitos = DigitosNivel(I)
'                Else
'                    Digitos = vEmpresa.DigitosUltimoNivel
'                End If
'            End If
'        End If
'    Next I
'
'    CADENA = String(Digitos, "_")
'
'    'Borramos los temporales
'    Sql = "DELETE from tmpbalancesumas where codusu= " & vUsu.Codigo
'    Conn.Execute Sql
'
'
'
'
'    IncrementarProgres pb2, 1
'
'    'Metemos las cuentas
'    Sql = "insert into tmpbalancesumas (codusu, cta, nomcta, totald, totalh) "
'    Sql = Sql & "select distinct " & vUsu.Codigo & ", mid(hlinapu.codmacta,1," & Digitos & ") , cuentas.nommacta, 0, 0 "
'    Sql = Sql & " from hlinapu inner join cuentas on mid(hlinapu.codmacta,1," & Digitos & ") = cuentas.codmacta where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
'    Sql = Sql & "  AND mid(hlinapu.codmacta,1,1) IN ('6','7') "
'    Sql = Sql & " and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaF, "F")
'
'
'    Conn.Execute Sql
'
'    IncrementarProgres pb2, 1
'
'
'
'    'Actualizamos acumulados del periodo actual
'    'activo
'
'    Sql = "update tmpbalancesumas set "
'    Sql = Sql & " totalh = (select sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
'    Sql = Sql & " and fechaent between " & DBSet(FechaI, "F") & " and " & DBSet(FechaF, "F")
'    Sql = Sql & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
'    Sql = Sql & " where mid(tmpbalancesumas.cta,1,1) = '6' "
'    Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
'
'    Conn.Execute Sql
'
'    IncrementarProgres pb2, 1
'
'    'pasivo
'    Sql = "update tmpbalancesumas set "
'    Sql = Sql & " totalh = (select sum(coalesce(timporteh,0)) - sum(coalesce(timported,0)) from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
'    Sql = Sql & "   and fechaent between " & DBSet(FechaI, "F") & " and " & DBSet(FechaF, "F")
'    Sql = Sql & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
'    Sql = Sql & " where mid(tmpbalancesumas.cta,1,1) = '7' "
'    Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
'
'    Conn.Execute Sql
'
'    IncrementarProgres pb2, 1
'
'
'    'Actualizamos acumulados del periodo anterior
'    'activo
'    Sql = "update tmpbalancesumas set "
'    Sql = Sql & " totald = (select sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
'    Sql = Sql & "   and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaFAnt, "F")
'
'
'
'
'
'    Sql = Sql & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
'    Sql = Sql & " where mid(tmpbalancesumas.cta,1,1) = '6' "
'    Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
'
'    Conn.Execute Sql
'
'    IncrementarProgres pb2, 1
'
'    'pasivo
'    Sql = "update tmpbalancesumas set "
'    Sql = Sql & " totald = (select sum(coalesce(timporteh,0)) - sum(coalesce(timported,0)) from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
'    Sql = Sql & " and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaFAnt, "F")
'    Sql = Sql & " and mid(hlinapu.codmacta,1," & Digitos & ") = tmpbalancesumas.cta) "
'    Sql = Sql & " where mid(tmpbalancesumas.cta,1,1) = '7' "
'    Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N") '& ")"
'
'    Conn.Execute Sql
'
'    IncrementarProgres pb2, 1
'
'    'borramos las cuentas que no tienen movimientos en ese periodo
'    Sql = "delete from tmpbalancesumas where codusu = " & vUsu.Codigo & " and totald is null and totalh is null"
'    Conn.Execute Sql
'
'    IncrementarProgres pb2, 1
'
'    CargarTablaTemporal = True
'
'    pb2.visible = False
'    Exit Function
'
'
'eCargarTablaTemporal:
'    pb2.visible = False
'    MuestraError Err.Number, "Cargando tabla temporal", Err.Description
'End Function




Private Function CargarTablaTemporal(Fecha As String) As Boolean
Dim SQL As String
Dim FechaI As Date
Dim FechaF As Date
Dim FechaIAnt As Date
Dim FechaFAnt As Date
Dim i As Integer
Dim Cadena As String
Dim Digitos As Integer
Dim ConceptoPerdidasyGanancias As Integer

Dim col6 As Collection
Dim col7 As Collection
            
            
        
            
    lblCuentas(7).Caption = "Preparando datos"
    lblCuentas(7).Refresh
    
    ConceptoPerdidasyGanancias = 960 'ponia 970 en el sqql de abjao 970
    
    On Error GoTo eCargarTablaTemporal
        
        
    CargarTablaTemporal = False




    'Cargamos las fechas de calculo
    
    If chkExplotacion.Value = 1 Then
        'Febrero 2022
        ' QUiere los datos solo del mes.
        FechaI = CDate("01/" & Format(cmbFecha(2).ListIndex + 1, "00") & "/" & Me.txtAno(4).Text)
    Else
        'Lo que habia
        FechaI = CDate(Fecha)
    End If
    i = DiasMes(cmbFecha(2).ListIndex + 1, CInt(txtAno(4).Text))
    FechaF = CDate(i & "/" & cmbFecha(2).ListIndex + 1 & "/" & Me.txtAno(4).Text)
    
    FechaIAnt = DateAdd("yyyy", -1, FechaI)
    FechaFAnt = DateAdd("yyyy", -1, FechaF)
    
    For i = 1 To 10
        If Me.chkCtaExplo(i).visible Then
            If Me.chkCtaExplo(i).Value = 1 Then
                If i < 10 Then
                    Digitos = DigitosNivel(i)
                Else
                    Digitos = vEmpresa.DigitosUltimoNivel
                End If
            End If
        End If
    Next i
    
    
    
    'Borramos los temporales
    SQL = "DELETE from tmpbalancesumas where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "DELETE from tmpevolsal where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    
    Set miRsAux = New ADODB.Recordset


    
    
    If DeTantosEnTantos < 0 Then
        DeTantosEnTantos = 10
        
        'Falta paremtro
        'Falta paremtro  FALTA#########################################################
        If vEmpresa.NIF = "F46221503" Or vEmpresa.NIF = "B98877806" Then
            DeTantosEnTantos = 1 'teletaxi
        Else
            DeTantosEnTantos = 10
            lblCuentas(7).Caption = "Establecer valores iniciales"
            lblCuentas(7).Refresh
            SQL = "select count(*)  from hlinapu WHERE hlinapu.codconce<>960 AND "
            SQL = SQL & " fechaent >=" & DBSet(vParam.fechaini, "F") & " and " & DBSet(FechaF, "F")
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                NumRegElim = DBLet(miRsAux.Fields(0), "N")
                If NumRegElim > 100000 Then
                    DeTantosEnTantos = 1
                ElseIf NumRegElim > 75000 Then
                        DeTantosEnTantos = 3
                ElseIf NumRegElim > 50000 Then
                    DeTantosEnTantos = 5
                End If
            End If
            miRsAux.Close
        End If
    End If
    
    'Octubre 2021
    lblCuentas(7).Caption = "Preparando datos 6%"
    lblCuentas(7).Refresh
    'Iremos desde fhaianterior a fecf
    SQL = "select codmacta from hlinapu where codmacta like '6%' "
    SQL = SQL & " and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaF, "F") & " GROUP BY codmacta"
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set col6 = New Collection
    i = 1
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & ", " & DBSet(miRsAux!codmacta, "T")
        i = i + 1
        miRsAux.MoveNext
        If miRsAux.EOF Then i = DeTantosEnTantos + 1
        
        If i > DeTantosEnTantos Then
            col6.Add Mid(SQL, 2)
            SQL = ""
            i = 1
        End If
    Wend
    miRsAux.Close
    
    
    
    lblCuentas(7).Caption = "Preparando datos 7%"
    lblCuentas(7).Refresh
    
    SQL = "select codmacta from hlinapu where codmacta like '7%' "
    SQL = SQL & " and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaF, "F") & " GROUP BY codmacta"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set col7 = New Collection
    i = 1
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & ", " & DBSet(miRsAux!codmacta, "T")
        i = i + 1
        miRsAux.MoveNext
        If miRsAux.EOF Then i = DeTantosEnTantos + 1
        
        If i > DeTantosEnTantos Then
            col7.Add Mid(SQL, 2)
            SQL = ""
            i = 1
        End If
    Wend
    miRsAux.Close
    
    
    
    pb2.visible = True
    CargarProgres pb2, (col6.Count * 2) + (col7.Count * 2) + 4 'cada mes tiene 4 procesos
    

    For i = 1 To col6.Count
        
        
        IncrementarProgres pb2, 1
        lblCuentas(7).Caption = "Cuentas 6 ant (" & i & "/" & col6.Count & ")"
        lblCuentas(7).Refresh
        SQL = "insert into tmpevolsal(codusu,codmacta,apertura,importemes1 ,importemes2, importemes3  ) "
        SQL = SQL & "select " & vUsu.Codigo & ", codmacta , 1, sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) "
        'Feb2022
        SQL = SQL & ", sum(coalesce(timported,0)) , sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
        SQL = SQL & " and fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaFAnt, "F")
        SQL = SQL & "  AND hlinapu.codmacta in (" & col6.Item(i) & ")"
        SQL = SQL & " GROUP BY codmacta"
        Conn.Execute SQL

        IncrementarProgres pb2, 1
        lblCuentas(7).Caption = "Cuentas 6 act(" & i & "/" & col6.Count & ")"
        lblCuentas(7).Refresh
        SQL = "insert into tmpevolsal(codusu,codmacta,apertura,importemes1) "
        SQL = SQL & "select " & vUsu.Codigo & ", codmacta , 2, sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
        SQL = SQL & " and fechaent between " & DBSet(FechaI, "F") & " and " & DBSet(FechaF, "F")
        SQL = SQL & "  AND hlinapu.codmacta in (" & col6.Item(i) & ")"
        SQL = SQL & " GROUP BY codmacta"
        Conn.Execute SQL

        espera 0.1
    
    
        If (i Mod 25) = 0 Then DoEvents
    
    Next

    DoEvents

    For i = 1 To col7.Count
        lblCuentas(7).Caption = "Cuentas 7 ant (" & i & "/" & col7.Count & ")"
        lblCuentas(7).Refresh
        IncrementarProgres pb2, 1
        SQL = "insert into tmpevolsal(codusu,codmacta,apertura,importemes1) "
        SQL = SQL & "select " & vUsu.Codigo & ", codmacta , 1, sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) "
        'SQL = SQL & ", sum(coalesce(timported,0)) , sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
        SQL = SQL & " AND fechaent between " & DBSet(FechaIAnt, "F") & " and " & DBSet(FechaFAnt, "F")
        SQL = SQL & "  AND hlinapu.codmacta in (" & col7.Item(i) & ")"
        SQL = SQL & " GROUP BY codmacta"
        Conn.Execute SQL
     
        
        lblCuentas(7).Caption = "Cuentas 7 act(" & i & "/" & col7.Count & ")"
        lblCuentas(7).Refresh
        IncrementarProgres pb2, 1
        SQL = "insert into tmpevolsal(codusu,codmacta,apertura,importemes1)"   'importemes2, importemes3) "
        SQL = SQL & "select " & vUsu.Codigo & ", codmacta , 2, sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) "
        'SQL = SQL & ", sum(coalesce(timported,0)) , sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu where hlinapu.codconce<>" & ConceptoPerdidasyGanancias
        SQL = SQL & " AND fechaent between " & DBSet(FechaI, "F") & " and " & DBSet(FechaF, "F")
        SQL = SQL & "  AND hlinapu.codmacta in (" & col7.Item(i) & ")"
        SQL = SQL & " GROUP BY codmacta"
        Conn.Execute SQL
        
        If (i Mod 25) = 0 Then DoEvents
    
    Next i
    
    Screen.MousePointer = vbHourglass
    DoEvents
    espera 0.5
    

    
    
    'Saldos grupo 7 al reves
    IncrementarProgres pb2, 1
    SQL = "update tmpevolsal set importemes1=-importemes1 where codusu = " & vUsu.Codigo & " AND codmacta like '7%' "
    Conn.Execute SQL
    
    
    'Si no es ultimo nivel
    'creo los datos y borro el utlimo nivel
    IncrementarProgres pb2, 1
    If Digitos <> vEmpresa.DigitosUltimoNivel Then
        SQL = "insert into tmpevolsal(codmacta,apertura,codusu,importemes1) "
        SQL = SQL & " select substring(codmacta,1," & Digitos & "),apertura,codusu,sum(importemes1) from tmpevolsal  where codusu =" & vUsu.Codigo & "  group by 1,2"
        Conn.Execute SQL
    
        espera 0.25
        
        SQL = " delete from tmpevolsal  where codusu =" & vUsu.Codigo & "  AND length(codmacta) =" & vEmpresa.DigitosUltimoNivel
        Conn.Execute SQL
    End If
    
        
        
    
    
    
   
    lblCuentas(7).Caption = "Insertar para comparativo"
    lblCuentas(7).Refresh
    
    If chkExplotacion.Value = 1 Then
        'Febrero 2022
        ' Ha pedido comparativo movimientos mes. Para que las cuentas sean las mismas en todo los informes, sacaremos incluso las que tenga CERO
        
        ' Para que salgan todas en el orden que toca, siempre las mismas,haremos lo siguiente
        '  insertaremos desde cuenta, todas las codmacta del grupo, haciendo los IGNORES
        SQL = "insert IGNORE into tmpevolsal(codmacta,apertura,codusu,importemes1) "
        SQL = SQL & " select substring(codmacta,1," & Digitos & "),1 apertura," & vUsu.Codigo & " codusu, 0 from cuentas  where apudirec='S' and codmacta like '6%'"
        Conn.Execute SQL
        SQL = "insert IGNORE into tmpevolsal(codmacta,apertura,codusu,importemes1) "
        SQL = SQL & " select substring(codmacta,1," & Digitos & "),2 apertura," & vUsu.Codigo & " codusu, 0 from cuentas  where apudirec='S' and codmacta like '6%'"
        Conn.Execute SQL
    
    
        SQL = "insert IGNORE into tmpevolsal(codmacta,apertura,codusu,importemes1) "
        SQL = SQL & " select substring(codmacta,1," & Digitos & "),1 apertura," & vUsu.Codigo & " codusu, 0 from cuentas  where apudirec='S' and codmacta like '7%'"
        Conn.Execute SQL
        
        
        SQL = "insert IGNORE into tmpevolsal(codmacta,apertura,codusu,importemes1) "
        SQL = SQL & " select substring(codmacta,1," & Digitos & "),2 apertura," & vUsu.Codigo & " codusu, 0 from cuentas  where apudirec='S' and codmacta like '7%'"
        Conn.Execute SQL
    
    End If
    
    
    'Metemos en la tabla del informe
    lblCuentas(7).Caption = "Insertar tmp impresion"
    lblCuentas(7).Refresh
    IncrementarProgres pb2, 1
    SQL = "INSERT INTO tmpbalancesumas(codusu,cta,nomcta,TotalD,TotalH)"   'acumperd,acumperh,acumantd,acumanth)"
    SQL = SQL & "SELECT codusu,tmpevolsal.codmacta,if(cuentas.nommacta is null,'ERROR',cuentas.nommacta),sum(if(apertura<2,importemes1,0)),"
    SQL = SQL & " sum(if(apertura>=2,importemes1,0)) "
'    If chkExplotacion.Value = 1 Then
'        SQL = SQL & " ,sum(if(apertura<2,importemes2,0)), sum(if(apertura>=2,importemes2,0))"
'        SQL = SQL & " ,sum(if(apertura<2,importemes3,0)), sum(if(apertura>=2,importemes3,0))"
'    Else
'        SQL = SQL & " ,0,0,0,0"
'    End If
'
    SQL = SQL & " from tmpevolsal left join cuentas on tmpevolsal.codmacta=cuentas.codmacta"
    SQL = SQL & " where codusu =" & vUsu.Codigo & " group by tmpevolsal.codmacta"
    Conn.Execute SQL
    
    
        
    IncrementarProgres pb2, 1
    
    'borramos las cuentas que no tienen movimientos en ese periodo
    'Febrero 2022   Solo si no selecciona movimientos mes
    If chkExplotacion.Value = 0 Then
        SQL = "delete from tmpbalancesumas where codusu = " & vUsu.Codigo & " and totald is null and totalh is null"
        Conn.Execute SQL
        
    End If
    IncrementarProgres pb2, 1
    lblCuentas(7).Caption = "Abriendo"
    
    
    CargarTablaTemporal = True
        
    pb2.visible = False
    Set col6 = Nothing
    Set col7 = Nothing
    
    Exit Function
        
        
eCargarTablaTemporal:
    pb2.visible = False
    MuestraError Err.Number, "Cargando tabla temporal", Err.Description
    lblCuentas(7).Caption = ""
    
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
        MsgBox "El A�o debe ser num�rico: " & txtAno(Index).Text, vbExclamation
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
    
    'Mes de c�lculo
    If cmbFecha(2).ListIndex < 0 Then
        MsgBox "Seleccion un mes para el c�lculo.", vbExclamation
        Exit Function
    End If
    
    ' Uno y solo uno de los niveles tiene que estar marcado
    Cad = ""
    For i = 1 To 10
        If Me.chkCtaExplo(i).visible Then
            If Me.chkCtaExplo(i).Value = 1 Then
                If i < 10 Then
                    CONT = DigitosNivel(i)
                Else
                    CONT = vEmpresa.DigitosUltimoNivel
                End If
                Cad = Cad & "1"
            End If
        End If
    Next i
    If Len(Cad) <> 1 Then
        MsgBox "Seleccione uno(y solo uno) de los niveles para el informe.", vbExclamation
        Exit Function
    End If
    
    If txtAno(4).Text = "" Then
        MsgBox "Ponga el a�o para el listado.", vbExclamation
        Exit Function
    End If

    If vParam.GrupoOrd <> "" And vParam.Automocion <> "" Then
        If CDate("01/" & cmbFecha(2).ListIndex + 1 & "/" & txtAno(4).Text) > vParam.fechafin Then
            'Ha seleccionado a uno o dos digitos
            If chkCtaExplo(1).Value = 1 Or chkCtaExplo(2).Value = 1 Then
                MsgBox "La cuenta de exclusi�n del grupoord de la anal�tica no esta inclu�da en el balance", vbExclamation
            End If
        End If
    End If


    If chkPorcentajes.Value = 1 Then
        If chkComparativo.Value = 0 Then
            MsgBox "Porcentajes s�lo v�lidos para comparativo", vbExclamation
            Exit Function
        End If
    End If

    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer

    QueCombosFechaCargar "2|"
    
    
    'Y ademas deshabilitamos los niveles no utilizados por la aplicacion
    For i = vEmpresa.numnivel To 9
        Me.chkCtaExplo(i).visible = False
    Next i
    
    For i = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(i)
        chkCtaExplo(i).visible = True
        chkCtaExplo(i).Caption = "Digitos: " & J
    Next i


End Sub


Private Sub CargarComboA�o(cmb As ComboBox)
Dim J As Integer
    
    cmb.Clear
    
    'Y ademas deshabilitamos los niveles no utilizados por la aplicacion
    For J = 2000 To Year(vParam.fechafin) + 1
        cmb.AddItem J
        cmb.ItemData(cmb.NewIndex) = J
    Next J
    
End Sub





Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    Cad = RecuperaValor(Lista, L)
    If Cad <> "" Then
        i = Val(Cad)
        With cmbFecha(i)
            .Clear
            For CONT = 1 To 12
                RC = "25/" & CONT & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next CONT
        End With
    End If
    L = L + 1
Loop Until Cad = ""
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
            Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If Not IsNull(Rs.Fields(0)) Then
                    If Rs.Fields(0) > 0 Then HayAsientoCierre = True
                End If
            End If
            Rs.Close
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
