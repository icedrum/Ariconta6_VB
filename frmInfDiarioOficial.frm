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
      TabIndex        =   33
      Top             =   0
      Width           =   4455
      Begin VB.Frame FrameDigitos 
         Height          =   1815
         Left            =   120
         TabIndex        =   41
         Top             =   3120
         Width           =   4245
         Begin VB.CheckBox Check1 
            Caption         =   "�ltimo"
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
            TabIndex        =   11
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
            TabIndex        =   20
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
            Left            =   2880
            TabIndex        =   19
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
            Left            =   1440
            TabIndex        =   18
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
            TabIndex        =   17
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
            Left            =   2880
            TabIndex        =   16
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
            Left            =   1440
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   600
            Width           =   1275
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
            Left            =   2880
            TabIndex        =   13
            Top             =   240
            Width           =   1305
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
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkRenumerar 
         Alignment       =   1  'Right Justify
         Caption         =   "Renumerar"
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
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   1965
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
         Index           =   4
         Left            =   1890
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   690
         Width           =   1485
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   40
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
      Begin VB.CheckBox chkInfApuntes 
         Alignment       =   1  'Right Justify
         Caption         =   "Inf. extendida"
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
         Left            =   120
         TabIndex        =   49
         ToolTipText     =   "Datos creacion. usuario-Aplicacion"
         Top             =   3840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1965
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   43
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
         TabIndex        =   42
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
      TabIndex        =   32
      Top             =   0
      Width           =   6915
      Begin VB.CheckBox chkDiarioOficial 
         Caption         =   "Diario oficial resumen"
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
         TabIndex        =   4
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3135
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
         ItemData        =   "frmInfDiarioOficial.frx":008B
         Left            =   2790
         List            =   "frmInfDiarioOficial.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   810
         Width           =   1215
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
         Index           =   3
         ItemData        =   "frmInfDiarioOficial.frx":008F
         Left            =   2790
         List            =   "frmInfDiarioOficial.frx":0091
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1260
         Width           =   1215
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
         ItemData        =   "frmInfDiarioOficial.frx":0093
         Left            =   1170
         List            =   "frmInfDiarioOficial.frx":0095
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1260
         Width           =   1575
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
         ItemData        =   "frmInfDiarioOficial.frx":0097
         Left            =   1170
         List            =   "frmInfDiarioOficial.frx":0099
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   810
         Width           =   1575
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
         TabIndex        =   39
         Top             =   1290
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
         TabIndex        =   38
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
         TabIndex        =   37
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   23
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
      TabIndex        =   24
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
      TabIndex        =   44
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
      TabIndex        =   48
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

Private Sql As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim Rs As ADODB.Recordset

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
            If Check1(I).visible Then
                If I <> Index Then Check1(I).Value = 0
            End If
        Next I
'        Check1(Index).Value = Valor
    End If


End Sub


Private Sub chkDiarioOficial_Click()
    FrameDigitos.visible = chkDiarioOficial.Value = 1
    Me.chkRenumerar.visible = chkDiarioOficial.Value <> 1
    Me.Caption = "Diario Oficial "
    If Me.chkDiarioOficial.Value = 1 Then Me.Caption = Me.Caption & "(RESUMEN)"
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida

End Sub


Private Sub chkDiarioOficial_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkRenumerar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim B As Boolean

    If Not DatosOK Then Exit Sub
    
    PulsadoCancelar = False
    Me.cmdCancelarAccion.visible = True
    Me.cmdCancelarAccion.Enabled = True
    
    Me.cmdCancelar.visible = False
    Me.cmdCancelar.Enabled = False
        
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    

'++
    Screen.MousePointer = vbHourglass
    If Me.chkDiarioOficial.Value = 0 Then
        B = GeneraDiarioOficial
    Else
        B = GenerarLibroResumen
    End If
    Label2(25).Caption = ""
    If B Then
        
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
    
    
    Me.cmdCancelarAccion.visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.visible = True
    Me.cmdCancelar.Enabled = True
    
    
    Screen.MousePointer = vbDefault
    Label2(25).Caption = ""

End Sub

Private Sub cmdCancelar_Click()
    If Me.cmdCancelarAccion.visible Then Exit Sub
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
    
    Me.Icon = frmppal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Diario Oficial (RESUMEN)"

        
    PrimeraVez = True
     
     
    chkDiarioOficial.Value = 1
    CargarComboFecha
     
    PonerNiveles
     
    If Legalizacion <> "" Then
        txtFecha(7).Text = RecuperaValor(Legalizacion, 1)
    
        'Fecha inicial
        cmbFecha(0).ListIndex = Month(RecuperaValor(Legalizacion, 2)) - 1
        cmbFecha(1).ListIndex = Month(RecuperaValor(Legalizacion, 3)) - 1
    
        cmbFecha(2).Text = Year(RecuperaValor(Legalizacion, 2))
        cmbFecha(3).Text = Year(RecuperaValor(Legalizacion, 3))
            
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
        cmbFecha(2).Text = Year(vParam.fechaini) - 1
        cmbFecha(3).Text = Year(vParam.fechafin) - 1
    End If
   
    PosicionarCombo cmbFecha(0), cmbFecha(0).ListIndex
    PosicionarCombo cmbFecha(1), cmbFecha(1).ListIndex
        
    PosicionarCombo cmbFecha(2), cmbFecha(2).ListIndex
    PosicionarCombo cmbFecha(3), cmbFecha(3).ListIndex
   
   
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.visible = False
    
    
    
    
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
    
    chkInfApuntes.visible = Index = 1
    
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
    
    If Me.chkDiarioOficial.Value = 1 Then
        Sql = "select fecha, asiento, cuenta, titulo, concepto, coalesce(debe,0) debe, coalesce(haber,0) haber from tmpdirioresum where codusu = " & vUsu.Codigo
        Sql = Sql & " order by clave "
    Else
        Sql = "select  fecha1 fecha,texto1 asiento,texto2 linea ,texto3,observa1,observa2 ampliacion,coalesce(importe1,0) debe,coalesce(importe2,0) haber"
        'Febrero 2016
        If Me.chkInfApuntes.Value = 1 Then Sql = Sql & " ,fecha2 FechaCreacion,trim(substring(texto,1,15)) usuario,substring(texto,16) Origen"
        Sql = Sql & " From tmptesoreriacomun WHERE codusu =" & vUsu.Codigo
        Sql = Sql & " order by codigo "
    
    End If
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    If Me.chkDiarioOficial.Value = 0 Then
        
        'Fechas
        RC = Me.cmbFecha(0).Text & "/" & cmbFecha(2).Text & "  al  " & Me.cmbFecha(1).Text & "/" & cmbFecha(3).Text
        cadParam = cadParam & "Fechas= ""Desde  " & RC & """|"
        
            
        'Fecha de impresion
        cadParam = cadParam & "FechaImp= """ & txtFecha(7).Text & """|"
        'Numero de hoja
        If txtNumRes(1).Text <> "" Then
            I = Val(txtNumRes(1).Text)
        Else
            I = 0
        End If
        cadParam = cadParam & "Numhoja= " & I & "|"
    
    
        'Acumulados anteriores
        If txtNumRes(3).Text <> "" Or txtNumRes(4).Text <> "" Then
            I = 0  'En el informe diremos k si se muestra
        Else
            I = 1
        End If
        cadParam = cadParam & "TieneAcumulados= " & I & "|"
    
        If I = 1 Then
            cadParam = cadParam & "AntD= 0|"
            cadParam = cadParam & "AntH= 0|"
        Else
            cadParam = cadParam & "AntD= " & TransformaComasPuntos(txtNumRes(3).Text) & "|"
            cadParam = cadParam & "AntH= " & TransformaComasPuntos(txtNumRes(4).Text) & "|"
        End If
        
    
        indRPT = "1306-01"
        cadFormula = "{tmptesoreriacomun.codusu}=" & vUsu.Codigo

    
    
    Else
        'RESUMEN. Lo que habia
        
        cadParam = cadParam & "pFecha=""" & txtFecha(7).Text & """|"
        
        'Numero de p�gina
        If txtNumRes(1).Text <> "" Then
            cadParam = cadParam & "pNumPag=" & txtNumRes(1).Text - 1 & "|"
        Else
            cadParam = cadParam & "pNumPag=0|"
        End If
        numParam = numParam + 2
        
        cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & cmbFecha(2).Text & " a " & cmbFecha(1).Text & " " & cmbFecha(3).Text & """|"
        numParam = numParam + 1
        
        
        indRPT = "1306-00"
        cadFormula = "{tmpdirioresum.codusu}=" & vUsu.Codigo

    End If
    
    vMostrarTree = False
    conSubRPT = False
        
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text, (Legalizacion <> "")) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 28
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
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
    
    If cmbFecha(2).Text = "" Or cmbFecha(3).Text = "" Then
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
    
    If Not ComparaFechasCombos(2, 3, 0, 1) Then Exit Function
    If txtFecha(7).Text <> "" Then
        If Not IsDate(txtFecha(7).Text) Then
            MsgBox "Fecha impresi�n incorrecta", vbExclamation
            txtFecha(7).SetFocus
        End If
    End If
    
    
    If Abs(Val(cmbFecha(2).Text) - Val(cmbFecha(3).Text)) > 2 Then
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
        CONT = Val(cmbFecha(2).Text)
    Else
        CONT = Val(cmbFecha(2).Text) - 1
    End If
    cad = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & CONT
    FechaIncioEjercicio = CDate(cad)
    
    I = cmbFecha(1).ListIndex + 1
    If I <= Month(vParam.fechafin) Then
        CONT = Val(cmbFecha(3).Text)
    Else
        CONT = Val(cmbFecha(3).Text) + 1
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
    If Me.chkDiarioOficial.Value = 1 Then
        CONT = 0
        For I = 1 To 10
            If Check1(I).visible = True Then
                If Check1(I).Value Then CONT = CONT + 1
            End If
        Next I
        If CONT <> 1 Then
            MsgBox "Seleccione uno, y solo uno, de los niveles para mostrar el informe", vbExclamation
            Exit Function
        End If
        
    Else
        If Me.chkRenumerar.Value = 1 Then
            If txtNumRes(0).Text = "" Then
                MsgBox "Tiene que poner el numero del primer asiento", vbExclamation
                Exit Function
            End If
        
            If Val(txtNumRes(0).Text) < 0 Then
                MsgBox "Valores positivos para numero de asiento", vbExclamation
                Exit Function
            End If
        End If
    End If


    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer


QueCombosFechaCargar "0|1|"

For I = 1 To vEmpresa.numnivel - 1
    J = DigitosNivel(I)
    Check1(I).visible = True
    Check1(I).Caption = "Digitos:" & J
Next I

    cmbFecha(2).Clear
    cmbFecha(3).Clear
    
    For I = 1 To 50
        cmbFecha(2).AddItem "20" & Format(I, "00")
        cmbFecha(3).AddItem "20" & Format(I, "00")
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
    If cmbFecha(Indice1).Text <> "" And cmbFecha(Indice2).Text <> "" Then
        If Val(cmbFecha(Indice1).Text) > Val(cmbFecha(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        Else
            If Val(cmbFecha(Indice1).Text) = Val(cmbFecha(Indice2).Text) Then
                If Me.cmbFecha(InCombo1).ListIndex > Me.cmbFecha(InCombo2).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    ComparaFechasCombos = True
End Function




Private Sub PonerNiveles()
Dim I As Integer
Dim J As Integer


    FrameDigitos.visible = True
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        cad = "Digitos: " & J
        Check1(I).visible = True
        Check1(I).Tag = J
        Me.Check1(I).Caption = cad
        Me.Check1(I).Value = 0
    Next I
    Check1(10).Tag = vEmpresa.DigitosUltimoNivel
    Check1(10).visible = True
    Me.Check1(10).Value = 0
    For I = vEmpresa.numnivel To 9
        Check1(I).visible = False
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


Private Function GenerarLibroResumen() As Boolean
Dim I2 As Currency

    On Error GoTo EGenerarLibroResumen
    GenerarLibroResumen = False
    
    'Eliminamos registros tmp
    Sql = "Delete FROM tmpdirioresum where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
        
    'Comprobamos k nivel
    For I = 1 To Me.Check1.Count
        If Check1(I).visible Then
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
    If cmbFecha(2).Text = cmbFecha(3).Text Then
        I = CInt(Val(cmbFecha(2).Text))
        For CONT = cmbFecha(0).ListIndex + 1 To cmbFecha(1).ListIndex + 1
           Label2(25).Caption = "Fecha: " & CONT & " / " & I
           Label2(25).Refresh
           
           DoEvent2
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
        I = CInt(Val(cmbFecha(2).Text))
        For CONT = cmbFecha(0).ListIndex + 1 To 12
           Label2(25).Caption = "Fecha: " & CONT & " / " & I
           Label2(25).Refresh
           
           DoEvent2
           If PulsadoCancelar Then Exit Function
           
           If CONT = cmbFecha(0).ListIndex + 1 Then
                If txtNumRes(3).Text <> "" Then Importe = CCur(txtNumRes(3).Text)
                If txtNumRes(4).Text <> "" Then I2 = CCur(txtNumRes(4).Text)
           End If
           ProcesaLibroResumen CONT, I, Importe, I2
           Importe = 0: I2 = 0
        Next CONT
        'A�os siguiente
        I = CInt(Val(cmbFecha(3).Text))
        For CONT = 1 To cmbFecha(1).ListIndex + 1
           Label2(25).Caption = "Fecha: " & CONT & " / " & I
           Label2(25).Refresh
           
           DoEvent2
           If PulsadoCancelar Then Exit Function
           
           ProcesaLibroResumen CONT, I, Importe, I2
        Next CONT
    End If
    
    'Vemos si ha generado datos
    Set miRsAux = New ADODB.Recordset
    Sql = "Select count(*) from tmpdirioresum where codusu =" & vUsu.Codigo
    CONT = 0
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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






'************************************************************************************************
' GeneraDiarioOficial

Private Function GeneraDiarioOficial() As Boolean
Dim Total As Long
Dim Pos As Long
Dim miCo As Long
Dim Cadena As String

Dim F1 As Date
Dim F2 As Date
Dim FFin As Date
Dim Aux As String


    On Error GoTo EGeneraDiarioOficial

    GeneraDiarioOficial = False
    
    Set Rs = New ADODB.Recordset
    
    'Borramos la temporal
    Label2(25).Caption = "Preparando BD"
    Label2(25).Refresh
    Conn.Execute "Delete from tmptesoreriacomun where codusu = " & vUsu.Codigo
    
    
    
    
    'Parte comun
    If chkRenumerar.Value = 1 Then CONT = Val(Me.txtNumRes(0).Text) - 1
    NumRegElim = 0
    
    F1 = CDate("01/" & Format(cmbFecha(0).ListIndex + 1, "00") & "/" & cmbFecha(2).Text)
    I = DiasMes(Me.cmbFecha(1).ListIndex + 1, CInt(cmbFecha(3).Text))
    FFin = CDate(Format(I, "00") & "/" & Format(cmbFecha(1).ListIndex + 1, "00") & "/" & cmbFecha(3).Text)
    
    
    Do
        DoEvent2
        F2 = DateAdd("m", 1, F1)
        F2 = DateAdd("d", -1, F2)
        If F2 > FFin Then F2 = FFin
            

        
        Label2(25).Caption = Format(F1, "mm/yy")
        Label2(25).Refresh
        
        cad = ""
        Sql = ""
        If Me.chkInfApuntes.Value = 1 Then cad = "hcabapu,": Sql = "hcabapu."
        cad = " FROM " & cad & " hlinapu,cuentas WHERE "
        If Me.chkInfApuntes.Value = 1 Then cad = cad & " hcabapu.numdiari = hlinapu.numdiari AND hcabapu.numasien= hlinapu.numasien  AND hcabapu.fechaent  = hlinapu.fechaent AND "
        cad = cad & " hlinapu.codmacta = cuentas.codmacta"
        cad = cad & " AND " & Sql & "fechaent >=" & DBSet(F1, "F")
        cad = cad & " AND " & Sql & "fechaent <=" & DBSet(F2, "F")
    
    
        Sql = "select hlinapu.fechaent,hlinapu.numasien,linliapu,cuentas.codmacta, cuentas.nommacta,numdocum,"
        Sql = Sql & "ampconce,timported,timporteh, "
        If Me.chkInfApuntes.Value = 0 Then
            Sql = Sql & " null texto, null fecha2"
        Else
            ' el usuario, como solo queda un campo, pondremos 15 carcateres(FIJOS) para el usuario. El resto texto obsfac
            'feccreacion usucreacion desdeaplicacion
            RC = "usucreacion,desdeaplicacion texto ,feccreacion fecha2"
            
            Sql = Sql & RC
        End If
        Sql = Sql & "" & cad
        Sql = Sql & " ORDER BY hlinapu.fechaent,hlinapu.numasien,linliapu"
        
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        DoEvent2
        If PulsadoCancelar Then
            Rs.Close
            Exit Function
        End If
 
    
        'Construimos la mitad de cadena de insercion
        Cadena = "INSERT INTO tmptesoreriacomun(codusu,codigo,fecha1,texto1,texto2,"
        Cadena = Cadena & "texto3,observa1,texto4,observa2,importe1,importe2,Texto,fecha2) VALUES "
        Sql = ""
     
        
        
    
        While Not Rs.EOF
            NumRegElim = NumRegElim + 1
            If chkRenumerar.Value = 1 Then
                'Si k estamos renumerando
               If miCo <> Rs!NumAsien Then
                    CONT = CONT + 1
                    miCo = Rs!NumAsien
                End If
            Else
                CONT = Rs!NumAsien
            End If
            
            If Label2(25).Caption <> Rs!FechaEnt Then
                Label2(25).Caption = Rs!FechaEnt
                Label2(25).Refresh
            End If
            
            cad = Rs!Nommacta
            NombreSQL cad
            cad = ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & Format(Rs!FechaEnt, FormatoFecha) & "','" & Format(CONT, "000000") & "','" & Format(Rs!Linliapu, "0000") & "','" & Rs!codmacta & "','" & cad & "','"
            cad = cad & DevNombreSQL(DBLet(Rs!Numdocum)) & "','" & DevNombreSQL(DBLet(Rs!Ampconce)) & "',"
            If Not IsNull(Rs!timported) Then
                RC = TransformaComasPuntos(CStr(Rs!timported))
                cad = cad & RC & ",NULL"
            Else
                RC = TransformaComasPuntos(CStr(Rs!timporteH))
                cad = cad & "NULL," & RC
            End If
            
            'Febrero 2021. Castelduc
            ' feccreacion usucreacion desdeaplicacion
            If IsNull(Rs!TEXTO) Then
                If Me.chkInfApuntes.Value = 0 Then
                    RC = "null"
                Else
                    RC = Space(20)
                End If
            Else
                'Varios supuetos
                
                RC = DBLet(Rs!TEXTO, "T")
                If UCase(Mid(RC, 1, 8)) = "ARICONTA" Then
                    'Dentro ariconta
                    RC = LCase(Mid(RC, 13))
                    
                    'Casos
                    If InStr(1, RC, "introducci�n de asientos") > 0 Then
                        RC = "Introducci�n de manual asientos"
                    
                    ElseIf InStr(1, RC, "inmoviliz") > 0 Then
                        RC = "Amortizacion"
                    ElseIf InStr(1, RC, "fra.pro reg:") > 0 Then
                        RC = "Fra. proveedor"
                    ElseIf InStr(1, RC, "contab. fra cli") > 0 Then
                        RC = "Fra. cliente"
                    ElseIf InStr(1, RC, "cancelacion cliente n�recepcio") > 0 Then
                        RC = "Tesoreria. Cancelacion recepcion doc."
                    ElseIf InStr(1, RC, "abono remesa:") > 0 Then
                        RC = "Tesoreria. Abono remesa"
                    ElseIf InStr(1, RC, "compensa:") > 0 Then
                        RC = "Tesoreria. Compensacion"
                    ElseIf InStr(1, RC, "devoluci�n remesa") > 0 Then
                        RC = "Tesoreria. Devolucion remesa"
                    Else
                        RC = RC 'NO hacemos nada
                    End If
                Else
                    'AOPlicacion externa. Ariges, aritaxi.....
                    
                    RC = RC
                    
                End If
                Aux = DBLet(Rs!usucreacion, "T")
                If Aux = "" Then Aux = "N/D"
                RC = Mid(Aux & Space(15), 1, 15) & RC
                RC = DBSet(RC, "T")
                
            End If
            cad = cad & "," & RC
            cad = cad & "," & DBSet(Rs!Fecha2, "F", "S") & ")"
            
            Sql = Sql & cad
           
            'Siguiente
            Pos = Pos + 1
            DoEvent2
            If PulsadoCancelar Then
                Rs.Close
                Exit Function
            End If
            
            If Len(Sql) > 20000 Then
                Sql = Mid(Sql, 2)
                Sql = Cadena & Sql
                Conn.Execute Sql
                Sql = ""
            End If
            
        
        
            Rs.MoveNext
        Wend
        Rs.Close
     
    
        If Sql <> "" Then
            Sql = Mid(Sql, 2)
            Sql = Cadena & Sql
            Conn.Execute Sql
            Sql = ""
        End If
         
        F1 = DateAdd("d", 1, F2)
        If F1 > FFin Then I = 1
    Loop Until I = 1
    GeneraDiarioOficial = True
    
EGeneraDiarioOficial:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function



