VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtilidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilidades"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   Icon            =   "frmUtilidades.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "Eliminar"
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
      Left            =   8940
      TabIndex        =   49
      Top             =   6390
      Width           =   1365
   End
   Begin VB.CommandButton cmdBus 
      Caption         =   "Buscar"
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
      Left            =   7530
      TabIndex        =   30
      Top             =   5940
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Can"
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
      Left            =   8940
      TabIndex        =   31
      Top             =   5940
      Width           =   1365
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4755
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame frameCLI 
      Height          =   1455
      Left            =   105
      TabIndex        =   21
      Top             =   5400
      Width           =   7275
      Begin VB.Frame FrameProgresoFac 
         Height          =   1215
         Left            =   90
         TabIndex        =   35
         Top             =   150
         Width           =   6945
         Begin VB.Label Label7 
            Caption         =   "Label7"
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
            Left            =   1680
            TabIndex        =   37
            Top             =   510
            Width           =   4395
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
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
            Left            =   120
            TabIndex        =   36
            Top             =   510
            Width           =   1455
         End
      End
      Begin VB.TextBox txtCLI2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   420
         Width           =   3345
      End
      Begin VB.OptionButton OptContab 
         Caption         =   "Errores Contabilización"
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
         Index           =   0
         Left            =   3780
         TabIndex        =   47
         Top             =   960
         Width           =   2565
      End
      Begin VB.OptionButton OptContab 
         Caption         =   "Saltos de factura"
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
         Left            =   930
         TabIndex        =   46
         Top             =   960
         Value           =   -1  'True
         Width           =   2475
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox Text1 
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
         Left            =   1740
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   420
         Width           =   1365
      End
      Begin VB.TextBox txtCLI 
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
         Left            =   3150
         MaxLength       =   1
         TabIndex        =   24
         Top             =   420
         Width           =   525
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   3720
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1770
         TabIndex        =   26
         Top             =   180
         Width           =   330
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   3
         Left            =   1410
         Picture         =   "frmUtilidades.frx":000C
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   2
         Left            =   2820
         Picture         =   "frmUtilidades.frx":0097
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Serie"
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
         Left            =   3180
         TabIndex        =   25
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame FrameSaldos 
      Caption         =   "Control de Cuadre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   90
      TabIndex        =   39
      Top             =   5370
      Width           =   7275
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text4"
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text4"
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4740
         TabIndex        =   42
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Haber"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2490
         TabIndex        =   41
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label8 
         Caption         =   "Debe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   40
         Top             =   390
         Width           =   825
      End
   End
   Begin VB.Frame frameBus2 
      Height          =   1515
      Left            =   90
      TabIndex        =   14
      Top             =   5340
      Width           =   7275
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   405
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   295
         Left            =   120
         TabIndex        =   16
         Top             =   150
         Width           =   5655
      End
   End
   Begin VB.Frame FrameDHCuenta 
      Height          =   1515
      Left            =   90
      TabIndex        =   32
      Top             =   5340
      Width           =   7275
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   2970
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   630
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   690
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
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
         Left            =   2970
         TabIndex        =   34
         Top             =   390
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
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
         Left            =   690
         TabIndex        =   33
         Top             =   390
         Width           =   1155
      End
   End
   Begin VB.Frame FrameExclusion 
      Height          =   1515
      Left            =   90
      TabIndex        =   11
      Top             =   5340
      Width           =   7305
      Begin VB.CommandButton cmdNuevaExclusion 
         Caption         =   "Insertar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2190
         TabIndex        =   13
         Top             =   570
         Width           =   1185
      End
      Begin VB.CommandButton cmdEliminaExclusion 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3810
         TabIndex        =   12
         Top             =   570
         Width           =   1185
      End
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   405
      Left            =   180
      TabIndex        =   2
      Top             =   5970
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Frame framBusCta 
      Height          =   1515
      Left            =   90
      TabIndex        =   17
      Top             =   5340
      Width           =   7275
      Begin VB.CommandButton cmdCrearCuenta 
         Caption         =   "Crear"
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
         Left            =   5310
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtHuecoCta 
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
         Left            =   2130
         TabIndex        =   19
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblHuecoCta 
         Caption         =   "Label5"
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
         Left            =   180
         TabIndex        =   18
         Top             =   630
         Width           =   1875
      End
   End
   Begin VB.Frame frameBusASiento 
      Height          =   1515
      Left            =   90
      TabIndex        =   5
      Top             =   5340
      Width           =   7275
   End
   Begin VB.Frame FrameDescuadre 
      Caption         =   "Intervalo búsqueda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   90
      TabIndex        =   6
      Top             =   5340
      Width           =   7275
      Begin VB.TextBox Text1 
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
         Left            =   4050
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
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
         Left            =   1170
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   1
         Left            =   3810
         Picture         =   "frmUtilidades.frx":0122
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imfech 
         Height          =   240
         Index           =   0
         Left            =   900
         Picture         =   "frmUtilidades.frx":01AD
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3390
         TabIndex        =   10
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   8
         Top             =   660
         Width           =   555
      End
   End
   Begin VB.Frame FrameAccionesCtas 
      Height          =   1515
      Left            =   90
      TabIndex        =   4
      Top             =   5340
      Width           =   7275
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   300
         TabIndex        =   38
         Top             =   300
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nueva búsqueda"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   0
      Left            =   10020
      Picture         =   "frmUtilidades.frx":0238
      Top             =   300
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Index           =   1
      Left            =   9690
      Picture         =   "frmUtilidades.frx":0382
      Top             =   300
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Cuentas sin movimientos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   5490
      Width           =   4515
   End
End
Attribute VB_Name = "frmUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'//////////////////////////////////////////////////////////////////////
'/*
'/*         Este formulario es para algunos puntos de utilidades.
'/*         Esta a parte pq vamos a poner el boton de parar busqueda
'/*         ya que es una simple busqueda



Public Opcion As Byte
    '0.- Cuentas sin movimiento
    '1.- ASientos descuadrados
    '2.- Agrupacion de cuentas en balance
    '3.- Cuentas as excluir en el consolidado
    '4.- Busqueda huecos cuentas libres
    '5.- Facturas Clientes
    '6.- Facturas proveedores
    '7.- Cuentas libres. Igual que el 4 pero cuando pulse crear, devolvera la cta libre
    

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
    
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
    
    
    
Private Estado As Byte
    '0.- Antes de empezar a buscar
    '1.- Buscando
    '2.- Han parado la busqueda
    '3.- Ha terminado la busqueda y hay datos
Dim Sql As String
Dim Rs As Recordset
Dim NumCuentas As Long
Dim I As Long
Dim ItmX As ListItem
Dim HanPulsadoCancelar As Boolean
Dim PrimeraVez As Boolean


Dim SePuedeCErrar As Boolean

Private Sub cmdBus_Click()
    SePuedeCErrar = False
    HacerBusqueda
    SePuedeCErrar = True
End Sub

Private Sub HacerBusqueda()


    VisualizarSeleccionar False
    

    Select Case Estado
    Case 0
        ListView1.ListItems.Clear
        Select Case Opcion
        Case 0
            NumCuentas = 1
            MontarBusqueda
            If NumCuentas = 0 Then Exit Sub
            'Solo para esto.
            frameBus2.Visible = True
            cmdCancel.Enabled = False

            QuitarHlinApu 0   ' Hco apuntes
            QuitarHlinApu 3   ' Facturas clientes
            QuitarHlinApu 4   ' Facturas proveedores
            QuitarHlinApu 9   ' contrapartida en hco apuntes
            QuitarHlinApu 11   ' presupuestos
            
            
            'Si tiene tesoreria comprobar que no esta en tablas de tesoreria
            If vEmpresa.TieneTesoreria Then
                For I = 23 To 25
                    QuitarHlinApu CByte(I)
                Next I
            End If
            
            
            QuitarOtrasCuentas
            
            
            
            RecordsetRestantes
            frameBus2.Visible = False
            PonerCampos 1
            cmdCancel.Enabled = True
            HanPulsadoCancelar = False
            RecorriendoRecordset
    
            
        Case 4, 7
            If Len(Me.txtHuecoCta.Text) < Me.txtHuecoCta.MaxLength Then
                MsgBox "Escriba el subgrupo completo", vbExclamation
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            cmdCrearCuenta.Visible = False
            CargarRecordSetCtasLibres
            If ListView1.ListItems.Count > 0 Then
                ListView1.ListItems(1).EnsureVisible
                cmdCrearCuenta.Visible = True
            Else
                MsgBox "Ninguna cuenta libre para el subgrupo: " & Me.txtHuecoCta.Text, vbInformation
            End If
            Screen.MousePointer = vbDefault
            PonerCampos 0
            
        Case 5, 6
            Screen.MousePointer = vbHourglass
            
            If OptContab(0).Value Then
            'If Me.chkContabilizacion.Value = 1 Then
                CargaEncabezado 100   '
                Set miRsAux = New ADODB.Recordset
                BuscarContabilizacionFacturas
                Set miRsAux = Nothing
            Else
                CargaEncabezado Opcion
                BuscarFacturasSaltos
            End If
            
            Screen.MousePointer = vbDefault
        Case Else
            'Buscar asiento descuadrado
            Me.FrameDescuadre.Visible = False
            MontaSQLBuscaAsien
            PonerCampos 1
            HanPulsadoCancelar = False
            RecorriendoRecordsetDescuadres
        End Select
    
    Case 2
        'Volvemos donde nos habiamos quedado
        PonerCampos 1
        HanPulsadoCancelar = False
        If Opcion = 0 Then
            RecorriendoRecordset
        Else
            RecorriendoRecordsetDescuadres
        End If
    Case 3
        ListView1.ListItems.Clear
        PonerCampos 0
        If Opcion = 1 Then CadenaDesdeOtroForm = ""
    Case 4
        'Busqueda cta libre
        
    End Select
End Sub

Private Sub cmdCancel_Click()
    Select Case Estado
    Case 0
        SePuedeCErrar = True
        If Opcion = 7 Then CadenaDesdeOtroForm = ""
        If Opcion = 1 Then CadenaDesdeOtroForm = ""
        Unload Me
        
    Case 1
        HanPulsadoCancelar = True
        PonerCampos 0
        
    Case 2
        'Volvemos a poner una nueva busqueda
        IntentaCErrar
        PonerCampos 0
        If Opcion = 1 Then Me.FrameDescuadre.Visible = True
        
    Case 3
        SePuedeCErrar = True
        If Opcion = 7 Then CadenaDesdeOtroForm = ""
        If Opcion = 1 Then CadenaDesdeOtroForm = ""
        Unload Me
        
    End Select
End Sub


Private Sub IntentaCErrar()
On Error Resume Next
    Rs.Close
    Err.Clear
    Set Rs = Nothing
End Sub


Private Sub cmdCrearCuenta_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione una cuenta", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = ""
    If Opcion = 7 Then
        SePuedeCErrar = True
        CadenaDesdeOtroForm = ListView1.SelectedItem.Text
        Unload Me
    Else
        frmCuentas.CodCta = ListView1.SelectedItem.Text
        frmCuentas.vModo = 1
        frmCuentas.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            'Eliminamos el nodo
            If ListView1.SelectedItem.Text = CadenaDesdeOtroForm Then ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    End If
End Sub

Private Sub cmdEliminaExclusion_Click()
    EliminarCta
End Sub



Private Sub EliminarCta()
On Error Resume Next
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    
    Sql = "Va a eliminar de "
    If Opcion = 2 Then
        Sql = Sql & "  la agrupacion de cuentas en balance la cuenta: " & vbCrLf
    Else
        Sql = Sql & "  la exclusion de cuentas en consolidado: " & vbCrLf
    End If
    Sql = Sql & ListView1.SelectedItem.Text & " - " & ListView1.SelectedItem.SubItems(1) & vbCrLf
    Sql = Sql & "Desea continuar ?"
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    Screen.MousePointer = vbHourglass
    Sql = "DELETE FROM "
    If Opcion = 3 Then
        Sql = Sql & "ctaexclusion"
    Else
        Sql = Sql & "ctaagrupadas"
    End If
    Sql = Sql & " WHERE codmacta = '" & ListView1.SelectedItem.Text & "';"
    Conn.Execute Sql
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar cuenta de agrupacion"
    Else
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub CmdEliminar_Click()
    Sql = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            Sql = "SI"
            Exit For
        End If
    Next I
    If Sql = "" Then
        MsgBox "Seleccione alguna cuenta a eliminar", vbExclamation
        Exit Sub
    End If
    Sql = "Va a eliminar las cuentas seleccionadas. ¿ Esta seguro ?"
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        Screen.MousePointer = vbHourglass
        Eliminar
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub cmdNuevaExclusion_Click()
    Set frmC = New frmColCtas
    frmC.ConfigurarBalances = 6
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.Show vbModal
End Sub

Private Sub Command1_Click()
    Checkear True
End Sub

Private Sub Checkear(SiNo As Boolean)
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = SiNo
    Next I
End Sub


Private Sub Command2_Click()
    Checkear False
End Sub

Private Sub Command3_Click()
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            If Not BloqueoManual(True, "Busquedas", "1") Then
                MsgBox "Se esta realizando la busqueda desde otro PC", vbExclamation
                PrimeraVez = True
                SePuedeCErrar = True
                Unload Me
            End If
            PonFocus Text2(0)
        Case 1
            Text1(0).SetFocus
        Case 4, 7
            txtHuecoCta.SetFocus
        Case 5, 6
            Text1(3).SetFocus
            
        End Select
    End If
End Sub

Private Sub VisualizarSeleccionar(SiNo As Boolean)
    imgCheck(0).Enabled = SiNo
    imgCheck(1).Enabled = SiNo
    imgCheck(0).Visible = SiNo
    imgCheck(1).Visible = SiNo

    

End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape And Opcion = 0 Then Unload Me
End Sub
'++

Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    SePuedeCErrar = False
    If Opcion <> 0 Then
        Me.ListView1.Icons = frmPpal.ImageList1
        Me.ListView1.SmallIcons = frmPpal.ImageList2
    End If
    CargaEncabezado Opcion
    PonerCampos 0
    Limpiar Me
    
    Me.frameCLI.Visible = False
    FrameExclusion.Visible = False
    Me.framBusCta.Visible = False
    Me.FrameDescuadre.Visible = False
    frameBus2.Visible = False
    Me.FrameDHCuenta.Visible = False
    frameBusASiento.Visible = False
    Me.FrameAccionesCtas.Visible = False
    Me.FrameSaldos.Visible = False
    
    
    
    VisualizarSeleccionar False
    
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 5
        .Buttons(2).Image = 1
    End With
    
    
    Select Case Opcion
    Case 0
        FrameDHCuenta.Visible = True
        Label1.Caption = "Cuentas sin movimientos"
    Case 1
        Label1.Caption = "Asientos descuadrados"
        Me.FrameDescuadre.Visible = True
        'ofertamos fechas
        Text1(0).Text = vParam.fechaini
        Text1(1).Text = vParam.fechafin
    Case 4, 7
        framBusCta.Visible = True
        Label1.Caption = "Búsquedas número de cuentas libres"
        txtHuecoCta.Text = ""
        cmdCrearCuenta.Visible = False
        PonerDigitosPenultimoNivel
        Me.cmdCrearCuenta.Enabled = vUsu.Nivel < 3
    Case 5, 6
        'Facturas clienes  y Facturas proveedores
        FrameProgresoFac.Left = 120
        FrameProgresoFac.Visible = False
        Me.frameCLI.Visible = True
        Text1(3).Text = vParam.fechaini
        Text1(2).Text = vParam.fechafin
        If Opcion = 5 Then
            Label1.Caption = "Nº Facturas CLIENTE incorrectos"
        Else
            Label1.Caption = "Nº Facturas PROVEEDORES incorrectos"
        End If
        
        imgppal(1).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
        
        
    End Select
    
    'No puede eliminar cuentas
    Me.Toolbar1.Buttons(1).Enabled = vUsu.Nivel < 2
    Me.CmdEliminar.Enabled = vUsu.Nivel < 2
    Me.CmdEliminar.Visible = False
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos(NuevoEstado As Byte)


    Select Case NuevoEstado
    Case 0
        Me.Label2.Caption = ""
        Me.pb1.Visible = False
        Me.cmdCancel.Caption = "Salir"
        Me.cmdBus.Caption = "Iniciar"
    Case 1
        Me.cmdCancel.Caption = "Parar"
        
    Case 2
        Me.cmdCancel.Caption = "Cancelar"
        Me.cmdBus.Caption = "Reanudar"
    Case 3
        Me.cmdBus.Caption = "Buscar"
        Me.cmdCancel.Caption = "Salir"
    End Select
    If Opcion = 0 Then
        Me.FrameAccionesCtas.Visible = NuevoEstado = 3
    Else
        Me.frameBusASiento.Visible = NuevoEstado = 3
        If Opcion = 1 Then
            Me.FrameSaldos.Visible = (NuevoEstado = 3)
            If NuevoEstado = 3 Then
                CargarSaldos
            Else
                LimpiarSaldos
            End If
            Me.FrameDescuadre.Visible = (NuevoEstado = 0)
        End If
    End If
    Me.cmdBus.Enabled = (NuevoEstado <> 1)
    cmdBus.Visible = (Opcion < 2) Or Opcion >= 4 'Cuando es agrupacion no mostramos el inciar
    Estado = NuevoEstado
End Sub


Private Sub CargarSaldos()
Dim Sql2 As String
Dim Rs2 As ADODB.Recordset

    Set Rs2 = New ADODB.Recordset
    
    Sql2 = ""
    'Fecha inicio
    If Text1(0).Text <> "" Then Sql2 = " fechaent >= '" & Format(Text1(0).Text, FormatoFecha) & "'"
    'Fecha fin
    If Text1(1).Text <> "" Then
        If Sql2 <> "" Then Sql2 = Sql2 & " AND "
        Sql2 = Sql2 & " fechaent <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    End If
    If Sql2 <> "" Then Sql2 = " WHERE " & Sql2
    
    Sql2 = "Select sum(coalesce(timported,0)) debe, sum(coalesce(timporteh,0)) haber, sum(coalesce(timported,0)) - sum(coalesce(timporteh,0)) saldo from hlinapu " & Sql2
    Rs2.Open Sql2, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not Rs2.EOF Then
        Text4.Text = Format(DBLet(Rs2.Fields(0).Value, "N"), "###,###,##0.00")
        Text3.Text = Format(DBLet(Rs2.Fields(1).Value, "N"), "###,###,##0.00")
        Text5.Text = Format(DBLet(Rs2.Fields(2).Value, "N"), "###,###,##0.00")
    End If

End Sub


Private Sub LimpiarSaldos()
    Text4.Text = ""
    Text3.Text = ""
    Text5.Text = ""
End Sub

Private Sub CargaEncabezado(LaOpcion As Byte)
Dim clmX As ColumnHeader

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    
    Select Case LaOpcion
    Case 0, 2, 3
        Me.ListView1.Checkboxes = LaOpcion = 0
        '* Estamos en cuentas sin movimiento
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Cuenta"
        clmX.Width = 2000 '1500
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Título"
        clmX.Width = 6500
    Case 1
        Me.ListView1.Checkboxes = False
        '* Estamos en cuentas sin movimiento
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Asiento"
        clmX.Width = 2000
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Fecha"
        clmX.Width = 1400
        'Diario
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Diario"
        clmX.Width = 800
    
        'Debe
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Descuadre"
        clmX.Width = 4000
        clmX.Alignment = lvwColumnRight
    Case 4, 7
        Me.ListView1.Checkboxes = False
        '* Estamos en buscando huecos cuentas
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Cuenta"
        clmX.Width = 2000 '1500
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Comentario"
        clmX.Width = 7000 '4500
    Case 5, 6
        
        Me.ListView1.Checkboxes = False
        '* Facturas
        'Cuenta
        If LaOpcion = 5 Then
            Set clmX = ListView1.ColumnHeaders.Add()
            clmX.Text = "Serie"
            clmX.Width = 900
            I = 3900
            Sql = "Codigo"
        Else
            Set clmX = ListView1.ColumnHeaders.Add()
            clmX.Text = "Serie"
            clmX.Width = 900
            I = 3900
            'i = 4500
            Sql = "Registro"
        End If
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = Sql
        clmX.Width = 1500
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Año"
        clmX.Width = 800
        'Clave2 ...
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Comentario"
        clmX.Width = I + 2000
       
       
    Case 100
        'Esta opcion es para las facturas, la busqueda de las contbilizaciones
        Me.ListView1.Checkboxes = False
        'Cuenta
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Factura"
        clmX.Width = 3200 '2500
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Asiento"
        clmX.Width = 3000 '4500
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Comentario"
        clmX.Width = 3600
        
        
        
        
    End Select
End Sub


Private Sub MontarBusqueda()
    Sql = "DELETE FROM tmpbussinmov"
    Conn.Execute Sql
    Sql = "INSERT INTO tmpbussinmov SELECT codmacta,nommacta from cuentas where apudirec='S'"
    If Text2(0).Text <> "" Then Sql = Sql & " AND codmacta >= '" & Text2(0).Text & "'"
    If Text2(1).Text <> "" Then Sql = Sql & " AND codmacta <= '" & Text2(1).Text & "'"
    Conn.Execute Sql
    
    
    If Text2(0).Text <> "" Or Text2(1).Text <> "" Then
        RecordsetRestantes
        If NumCuentas > 0 Then
            Rs.Close
            Set Rs = Nothing
        Else
            MsgBox "Ningun dato seleccionado", vbExclamation
        End If
    End If
    
End Sub



Private Sub QuitarHlinApu(vOpcion As Byte)
Dim T As Long
Dim t2 As Long
Dim codmacta1 As String
Dim Sql2 As String

    'Opcion
    ' 0 .- Hlinapu
   
    codmacta1 = "codmacta"
    Select Case vOpcion
    Case 0
        Sql = "hlinapu"
    Case 3, 4
        Sql = "factcli"
        If vOpcion = 4 Then Sql = "factpro"
        If vOpcion = 3 Then
            Sql2 = "Facturas de Clientes"
        Else
            Sql2 = "Facturas de Proveedores"
        End If
    Case 9, 10
        'Contrapartida en hlinapu, y hlinapu1
        codmacta1 = "ctacontr"
        Sql = "hlinapu"
        If vOpcion = 10 Then Sql = Sql & "1"
        Sql2 = "Contrapartida en Histórico de Apuntes"
        
    Case 11
        'PResupuestaria
        Sql = "presupuestos"
        Sql2 = "Presupuestos"
        
        
    '-----------------------------
    'TESORERIA
    Case 21
        Sql = "slicaja"
    
    Case 22
        codmacta1 = "ctacaja"
        Sql = "susucaja"
        
    Case 23
        Sql = "Departamentos"
        Sql2 = "Departamentos"
        
    Case 24
        Sql = "cobros"
        Sql2 = "Cobros"
        
    Case 25
        Sql = "pagos"
        Sql2 = "Pagos"
        codmacta1 = "codmacta"
    
    End Select
    Label4.Tag = Sql2
    Label4.Caption = "buscando datos " & Sql2
    pb2.Value = 0
    Me.Refresh
    
    Sql = "Select " & codmacta1 & " from " & Sql
    
    'Si es de hsaldos entonces tenemos k buscar solo en las k sean de ultmo nivel
    If vOpcion = 8 Or vOpcion = 7 Then _
        Sql = Sql & " WHERE codmacta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
    
    Sql = Sql & " group by " & codmacta1
    
    'having
    Sql = Sql & " HAVING NOT (" & codmacta1 & " IS NULL)"
    
    Set Rs = New ADODB.Recordset
    'Primro el contador
    Rs.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    T = 0
    While Not Rs.EOF
        T = T + 1
        Rs.MoveNext
    Wend
    
    If T > 0 Then
        Rs.MoveFirst
        Label4.Caption = Label4.Tag
        Label4.Refresh
        t2 = 0
        T = T + 1
        While Not Rs.EOF
            t2 = t2 + 1
            pb2.Value = ((t2 / T) * 1000)
            Sql = "Delete from tmpbussinmov where codmacta ='" & Rs.Fields(0) & "';"
            Conn.Execute Sql
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Label4.Caption = ""
    Set Rs = Nothing
End Sub


Private Sub RecordsetRestantes()
    Set Rs = New ADODB.Recordset
    Sql = "Select count(*) from tmpbussinmov"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumCuentas = 0
    I = 0
    If Not Rs.EOF Then
        NumCuentas = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    If NumCuentas = 0 Then Exit Sub
    pb1.Visible = True
    Label2.Caption = ""
    pb1.Value = 0
    Me.Refresh
    Sql = "Select * from tmpbussinmov order by codmacta"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
End Sub



Private Sub RecorriendoRecordset()
Dim ExisteReferencia As Boolean

    If NumCuentas = 0 Then Exit Sub
    While Not Rs.EOF
        ExisteReferencia = False
        
        Label2.Caption = Rs.Fields(0) & " - " & Rs.Fields(1)
        Label2.Refresh
        I = I + 1
        
        pb1.Value = Int(((I / NumCuentas)) * 1000)
        
        'Comprobamos en Facturas
        Sql = DevuelveDesdeBD("codmacta", "factcli", "codmacta", Rs.Fields(0), "T")
        If Sql <> "" Then ExisteReferencia = True
        
        'Proveedores
        If Not ExisteReferencia Then
            Sql = DevuelveDesdeBD("codmacta", "factpro", "codmacta", Rs.Fields(0), "T")
            If Sql <> "" Then ExisteReferencia = True
        End If
        
        'Lineas de facturas
        If Not ExisteReferencia Then
            Sql = DevuelveDesdeBD("codmacta", "factcli_lineas", "codmacta", Rs.Fields(0), "T")
            If Sql <> "" Then ExisteReferencia = True
        End If
        
        'Lineas de facturas proveedores
        If Not ExisteReferencia Then
            Sql = DevuelveDesdeBD("codmacta", "factpro_lineas", "codmacta", Rs.Fields(0), "T")
            If Sql <> "" Then ExisteReferencia = True
        End If
        
        
        'AMortizacion
        '-------------
        'Proveedor
        If Not ExisteReferencia Then
            Sql = DevuelveDesdeBD("codprove", "inmovele", "codprove", Rs.Fields(0), "T")
            If Sql <> "" Then ExisteReferencia = True
        End If
        If Not ExisteReferencia Then
            Sql = DevuelveDesdeBD("codmact1", "inmovele", "codmact1", Rs.Fields(0), "T")
            If Sql <> "" Then ExisteReferencia = True
        End If
        If Not ExisteReferencia Then
            Sql = DevuelveDesdeBD("codmact2", "inmovele", "codmact2", Rs.Fields(0), "T")
            If Sql <> "" Then ExisteReferencia = True
        End If
        If Not ExisteReferencia Then
            Sql = DevuelveDesdeBD("codmact3", "inmovele", "codmact3", Rs.Fields(0), "T")
            If Sql <> "" Then ExisteReferencia = True
        End If
        
        
        
        If Not ExisteReferencia Then
           Set ItmX = ListView1.ListItems.Add(, , Rs.Fields(0))
           If Opcion = 0 Then
              
           Else
               ItmX.SmallIcon = 1
           End If
           ItmX.SubItems(1) = Rs.Fields(1)
           ItmX.EnsureVisible
        Else
            Conn.Execute "Delete from tmpbussinmov where codmacta='" & Rs.Fields(0) & "'"
        End If
        
        'Siguiente
        Rs.MoveNext
        'Miramos si hay algo por hacer
        DoEvents
        
        'Si han pulsado parar
        If HanPulsadoCancelar Then
            PonerCampos 2
            
            If Opcion = 0 Then VisualizarSeleccionar True
            
            Exit Sub
        End If
    Wend
    Rs.Close
    
    If ListView1.ListItems.Count > 0 Then
        PonerCampos 3
        
        If Opcion = 0 Then
            Me.cmdBus.Enabled = False
            Me.cmdCancel.Enabled = False
            Me.cmdBus.Visible = False
            Me.cmdCancel.Visible = False
            Me.CmdEliminar.Visible = True
            VisualizarSeleccionar True
        End If
    Else
        Label2.Caption = ""
        Label2.Refresh
        MsgBox "Ninguna cuenta sin movimientos", vbExclamation
        PonerCampos 0
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SePuedeCErrar Then
        Cancel = 1
    Else
        If Not PrimeraVez Then
            If Opcion = 0 Then BloqueoManual False, "Busquedas", ""
            IntentaCErrar
            If Opcion = 1 Then
                RaiseEvent DatoSeleccionado(CadenaDesdeOtroForm)

            End If
        End If
    End If
            
End Sub


Private Sub Eliminar()
Dim cad As String
    Sql = "DELETE FROM cuentas where codmacta = '"
    For I = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(I).Checked Then
            cad = BorrarCuenta(ListView1.ListItems(I).Text, Me.Label2)
            If cad = "" Then
                If EliminaCuenta(ListView1.ListItems(I).Text) Then ListView1.ListItems.Remove I
            Else
                cad = ListView1.ListItems(I).Text & " - " & ListView1.ListItems(I).SubItems(1) & vbCrLf & cad & vbCrLf
                MsgBox cad, vbExclamation
            End If
        End If
    Next I
End Sub


Private Function EliminaCuenta(ByRef Cuenta As String) As Boolean
    On Error Resume Next
    Conn.Execute Sql & Cuenta & "'"
    If Err.Number <> 0 Then
        MuestraError Err.Number, Cuenta
        EliminaCuenta = False
    Else
        EliminaCuenta = True
    End If
End Function

Private Sub MontaSQLBuscaAsien()
    Set Rs = New ADODB.Recordset
    
    Sql = ""
    'Fecha inicio
    If Text1(0).Text <> "" Then Sql = " fechaent >= '" & Format(Text1(0).Text, FormatoFecha) & "'"
    'Fecha fin
    If Text1(1).Text <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & " fechaent <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    End If
    If Sql <> "" Then Sql = " WHERE " & Sql
    
    Sql = "Select numasien,numdiari,fechaent from hlinapu " & Sql
    Sql = Sql & " group by numasien,numdiari,fechaent"
    
    If Opcion = 1 Then Sql = Sql & " order by fechaent, numasien "
    
    Rs.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    NumCuentas = 0
    I = 0
    While Not Rs.EOF
        I = I + 1
        Rs.MoveNext
    Wend
    NumCuentas = I
    I = 0
    If NumCuentas = 0 Then
        Rs.Close
        Exit Sub
    End If
    Rs.MoveFirst
    pb1.Visible = True
    Label2.Caption = ""
    pb1.Value = 0
    Me.Refresh

End Sub

Private Sub RecorriendoRecordsetDescuadres()


    If NumCuentas = 0 Then Exit Sub
    While Not Rs.EOF
        
        Label2.Caption = Rs.Fields(0) & " - " & Rs.Fields(2)
        Label2.Refresh
        
        pb1.Value = Int(((I / NumCuentas)) * 1000)
        I = I + 1
        
        ObtenerSumas
        
        'Siguiente
        Rs.MoveNext
        'Miramos si hay algo por hacer
        DoEvents
        
        'Si han pulsado parar
        If HanPulsadoCancelar Then
            PonerCampos 2
            Exit Sub
        End If
    Wend
    Rs.Close
    
    If ListView1.ListItems.Count > 0 Then
        PonerCampos 3
    Else
        MsgBox "Ningun asiento descuadrado.", vbExclamation
        PonerCampos 0
        If Opcion = 1 Then PonerCampos 3
    End If

End Sub


Private Function ObtenerSumas() As Boolean
    Dim Deb As Currency
    Dim hab As Currency
    Dim RsA As ADODB.Recordset

    Set RsA = New ADODB.Recordset
    Sql = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    Sql = Sql & " From hlinapu "
    Sql = Sql & " WHERE (((numdiari)=" & Rs!NumDiari
    Sql = Sql & ") AND ((fechaent)='" & Format(Rs!FechaEnt, FormatoFecha)
    Sql = Sql & "') AND ((numasien)=" & Rs!NumAsien
    Sql = Sql & "));"
    RsA.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RsA.EOF Then
        If IsNull(RsA.Fields(0)) Then
            Deb = 0
        Else
            Deb = RsA.Fields(0)
        End If
        
        'Deb = Round(Deb, 2)
        If IsNull(RsA.Fields(1)) Then
            hab = 0
        Else
            hab = RsA.Fields(1)
        End If
        
        
        
    Else
        Deb = 0
        hab = 0
    End If
    RsA.Close
    
    'Metemos en DEB el total
    Deb = Deb - hab
    If Deb <> 0 Then
            Sql = Format(Rs!NumAsien, "0000000")
            Sql = "    " & Sql
            Set ItmX = ListView1.ListItems.Add(, , Sql)
            If Opcion <> 1 Then
                ItmX.SmallIcon = 2
                ItmX.Icon = 2
            Else
                ItmX.SmallIcon = 3
                ItmX.Icon = 3
            End If
            ItmX.SubItems(1) = Format(Rs!FechaEnt, "dd/mm/yyyy")
            ItmX.SubItems(2) = Rs!NumDiari
            ItmX.SubItems(3) = Format(Deb, FormatoImporte)
    End If
End Function


Private Sub cargaAgrupacion(tabla As String)
    On Error GoTo E1
    Set Rs = New ADODB.Recordset
    Sql = "Select " & tabla & ".codmacta, nommacta from " & tabla & ",cuentas where "
    Sql = Sql & tabla & ".codmacta=cuentas.codmacta order by " & tabla & ".codmacta"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
            Sql = Rs!codmacta
            Set ItmX = ListView1.ListItems.Add(, , Sql)
            ItmX.SmallIcon = 2
            ItmX.Icon = 2
            ItmX.SubItems(1) = Rs!Nommacta
            Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Exit Sub
E1:
    MuestraError Err.Number, tabla
    Set Rs = Nothing
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
On Error Resume Next
    If Opcion = 3 Then
        Sql = "ctaexclusion"
    Else
        Sql = "ctaagrupadas"
    End If
    Sql = "INSERT INTO " & Sql & "(codmacta) VALUES ('" & RecuperaValor(CadenaSeleccion, 1) & "')"
    Conn.Execute Sql
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando la cuenta"
    Else
        ListView1.ListItems.Clear
        If Opcion = 3 Then
            cargaAgrupacion "ctaexclusion"
        Else
            cargaAgrupacion "ctaagrupadas"
        End If
    End If
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCLI.Text = RecuperaValor(CadenaSeleccion, 1)
        txtCLI2.Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(I).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imfech_Click(Index As Integer)
    I = Index
    Set frmF = New frmCal
    Sql = Now
    If Text1(I).Text <> "" Then
        If IsDate(Text1(I).Text) Then Sql = Text1(I).Text
    End If
    frmF.Fecha = CDate(Sql)
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim B As Boolean
Dim I As Long
    B = Index = 1
    For I = 1 To ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = B
        If (I Mod 50) = 0 Then DoEvents
    Next I
End Sub



Private Sub imgppal_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 1 ' contadores
        Set frmConta = New frmBasico
        If Opcion = 5 Then
            AyudaContadores frmConta, txtCLI.Text, "tiporegi REGEXP '^[0-9]+$' = 0"
        Else
            AyudaContadores frmConta, txtCLI.Text, "tiporegi REGEXP '^[0-9]+$' <> 0 and tiporegi > 0"
        End If
        Set frmConta = Nothing
        PonFoco txtCLI
    End Select
End Sub

Private Sub ListView1_DblClick()
    If Opcion = 1 Then
        CadenaDesdeOtroForm = ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|" & ListView1.SelectedItem.SubItems(2) & "|"
        Unload Me
    End If
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
            Case 1:  KEYFecha KeyAscii, 1
            Case 2:  KEYFecha KeyAscii, 2
            Case 3:  KEYFecha KeyAscii, 3
            Case 6:  KEYFecha KeyAscii, 6
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imfech_Click (Indice)
End Sub
'++

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
        Text1(Index).Text = ""
    End If
End Sub

Private Sub PonerDigitosPenultimoNivel()
    'Veremos cual es el ultimo nivel
    I = vEmpresa.numnivel
    If I < 2 Then
        MsgBox "Empresa mal configurada", vbExclamation
        Exit Sub
    End If
    NumCuentas = I - 1
    I = DigitosNivel(CInt(NumCuentas))
    lblHuecoCta.Caption = "Dígitos del nivel " & NumCuentas & ":    " & I
    lblHuecoCta.Tag = I
    Me.txtHuecoCta.MaxLength = I
End Sub






Private Sub Text2_GotFocus(Index As Integer)
    PonFoco Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    Sql = Text2(Index).Text
    If CuentaCorrectaUltimoNivelSIN(Sql, CadenaDesdeOtroForm) < 1 Then
        
        MsgBox CadenaDesdeOtroForm, vbExclamation
        Text2(Index).Text = ""
        PonFocus Text2(Index)
    Else
        Text2(Index).Text = Sql
    End If
    CadenaDesdeOtroForm = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1 'eliminar

            Sql = ""
            For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).Checked Then
                    Sql = "SI"
                    Exit For
                End If
            Next I
            If Sql = "" Then
                MsgBox "Seleccione alguna cuenta a eliminar", vbExclamation
                Exit Sub
            End If
            Sql = "Va a eliminar las cuentas seleccionadas. ¿ Esta seguro ?"
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
                Screen.MousePointer = vbHourglass
                Eliminar
                Screen.MousePointer = vbDefault
            End If
            
        Case 2 ' Nueva búsqueda
            Me.cmdBus.Enabled = True
            Me.cmdCancel.Enabled = True
            Me.cmdBus.Visible = True
            Me.cmdCancel.Visible = True
            
            cmdBus_Click

        
    End Select
    
End Sub

Private Sub txtCLI_KeyPress(KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYBusqueda KeyAscii, 1 ' serie
     Else
        KEYpress KeyAscii
     End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgppal_Click (Indice)
End Sub


Private Sub txtCLI_LostFocus()
    
    If Not PerderFocoGnral(txtCLI, 3) Then Exit Sub
    
    If txtCLI.Text = "" Then Exit Sub
    
    txtCLI.Text = UCase(txtCLI.Text)

    If Me.Opcion = 5 Then ' clientes
        txtCLI2.Text = DevuelveValor("select nomregis from contadores where tiporegi = " & DBSet(txtCLI.Text, "T") & " and tiporegi REGEXP '^[0-9]+$' = 0")
    Else 'proveedores
        txtCLI2.Text = DevuelveValor("select nomregis from contadores where tiporegi = " & DBSet(txtCLI.Text, "T") & " and tiporegi REGEXP '^[0-9]+$' <> 0 and tiporegi > 0")
    End If
    If txtCLI2.Text = "0" Then
        MsgBox "Letra de serie no existe.", vbExclamation
        txtCLI2.Text = ""
        txtCLI.Text = ""
        PonFoco txtCLI2
    End If

End Sub

Private Sub txtHuecoCta_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtHuecoCta_LostFocus()
    txtHuecoCta.Text = Trim(txtHuecoCta.Text)
    If txtHuecoCta.Text <> "" Then
        If Not IsNumeric(txtHuecoCta.Text) Then
            MsgBox "Campo numérico", vbExclamation
            Exit Sub
        End If
        txtHuecoCta.Text = Val(txtHuecoCta.Text)
    End If
End Sub


Private Sub CargarRecordSetCtasLibres()
Dim cad As String
Dim J As Long
Dim Multiplicador As Long
Dim vFormato As String

    I = vEmpresa.DigitosUltimoNivel - lblHuecoCta.Tag
    vFormato = Mid("00000000000", 1, I)
    Multiplicador = I
    cad = Me.txtHuecoCta.Text & Mid("0000000000", 1, I)
    I = 1   'Primer Numero de cuenta
    
    Set Rs = New ADODB.Recordset
    Sql = "DELETE FROM tmpbussinmov"
    Conn.Execute Sql
    
    
    
    Sql = "Select codmacta from cuentas where codmacta like '" & Me.txtHuecoCta.Text & "%' AND Apudirec='S'"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = "INSERT INTO tmpbussinmov VALUES ('"
    If Rs.EOF Then
        'Estan todas libres
        
        InsertaCtasLibres Format(I, vFormato), "TODAS LIBRES"
        Rs.Close
    Else
        
        While Not Rs.EOF
            NumCuentas = CLng(Right(CStr(Rs.Fields(0)), Multiplicador))
            If NumCuentas > I Then
                For J = I To NumCuentas - 1
                    InsertaCtasLibres Format(J, vFormato), "SALTO"
                Next J
            End If
            I = NumCuentas + 1
            Rs.MoveNext
        Wend
        Rs.Close
        'Cojemos desde la ultima
        I = vEmpresa.DigitosUltimoNivel - lblHuecoCta.Tag
        cad = Mid("999999999", 1, I)
        I = Val(cad) 'Utlima cta del subgrupo
        
        If NumCuentas < I Then
            NumCuentas = NumCuentas + 1
            InsertaCtasLibres Format(NumCuentas, vFormato), "Desde aqui LIBRES"
        End If
        
        
    End If
    
        Sql = "Select * from tmpbussinmov ORDER BY codmacta"
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add(, , Rs.Fields(0))
           ItmX.SmallIcon = 1
           ItmX.SubItems(1) = Rs.Fields(1)
           ItmX.EnsureVisible
            Rs.MoveNext
        Wend
        Rs.Close
End Sub


Private Sub InsertaCtasLibres(Cta As String, Descripcion As String)
Dim cad As String
        cad = Me.txtHuecoCta.Text & Cta
        cad = cad & "','" & Descripcion & "')"
        Conn.Execute Sql & cad
End Sub



Private Sub BuscarContabilizacionFacturas()
    
    cmdCancel.Enabled = False
    cmdBus.Enabled = False
    
    ListView1.ListItems.Clear
    NumCuentas = 0
    Set Rs = New ADODB.Recordset
    
    Label7(0).Caption = ""
    Label7(1).Caption = ""
    Me.FrameProgresoFac.Visible = True
    
    'Comprobamos facturas que estando contabilizadas no tienen apuntes
'    FacturasContabilizadas
    
    
    '[Monica]  Todas las Facturas deberian estar contabilizadas
    FacturasNoContabilizadas
    
    'Apuntes que siendo de factura, no esta la factura
    ApuntesSinFacturaNew
    
    If NumCuentas = 0 Then MsgBox "Proceso finalizado", vbInformation
    
    CadenaDesdeOtroForm = ""
    
EBuscarContabilizacionFacturas:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
    Set miRsAux = Nothing
    Me.FrameProgresoFac.Visible = False
    cmdCancel.Enabled = True
    cmdBus.Enabled = True
    
End Sub


Private Sub BuscarFacturasSaltos()
Dim cad As String
Dim Aux As String
Dim Anyo As Integer
Dim J As Integer


    On Error GoTo EBuscarFacturas
        
    
    If Opcion = 5 Then
        Sql = "numserie,anofactu as ano,numfactu as codigo"
        cad = "fecfactu"
        Sql = Sql & " FROM factcli"
    Else
        Sql = "numserie, anofactu as ano,numregis as codigo"
        cad = "fecharec"
        Sql = Sql & " FROM factpro"
    End If
    
    
    'Si hay fecha inicio
    If Text1(3).Text = "" Or Text1(2).Text = "" Then
        MsgBox "Debe escribir las fechas de inicio y fin", vbExclamation
        Exit Sub
    End If
    Aux = ""
    Aux = cad & " >= '" & Format(Text1(3).Text, FormatoFecha) & "'"
    
    
    Aux = Aux & " AND "
    Aux = Aux & cad & " <= '" & Format(Text1(2).Text, FormatoFecha) & "'"
    
    
    
    If txtCLI.Text <> "" Then
        If Aux <> "" Then Aux = Aux & " AND "
        Aux = Aux & " numserie = '" & txtCLI.Text & "'"
    End If
    If Aux <> "" Then Sql = Sql & " WHERE " & Aux
    Sql = Sql & " ORDER BY "
    If Opcion = 5 Then
        Sql = Sql & "numserie,anofactu,numfactu "
    Else
        Sql = Sql & "numserie,anofactu,numregis"
    End If
    Set Rs = New ADODB.Recordset
    
    
    '#FALTA revisar esto
    
    'Obtenego el minimo
    
    Set miRsAux = New ADODB.Recordset
    
    'Fale. Ya tenemos montado el SQL
    
    Rs.Open "SELECT " & Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Serie
    Aux = ""
    Anyo = 0
    While Not Rs.EOF
        If Rs!NUmSerie <> Aux Then
            'Nueva SERIE
            Aux = Rs!NUmSerie
            Anyo = Rs!Ano
            I = FacturaMinimo(Aux, CDate(Text1(3).Text), CDate(Text1(2).Text), Anyo)
        End If
        If Anyo <> Rs!Ano Then
            'AÑO DISTINTO
            Anyo = Rs!Ano
            I = FacturaMinimo(Aux, CDate(Text1(3).Text), CDate(Text1(2).Text), Anyo)
        End If
        
        'Para cada numero de factura
        If I = Rs!Codigo Then
            I = I + 1
            'no hacemos nada mas
        Else
            'Si si que es mayor. Hay salto o hueco
            
            
            If Rs!Codigo - I >= 2 Then
                'SALTO
                cad = Format(Rs!Codigo - 1, "000000000")
'                If opcion = 5 Then
                    Set ItmX = ListView1.ListItems.Add(, , Rs!NUmSerie)
                    ItmX.SubItems(1) = cad
                    J = 2
'                Else
'                    Set ItmX = ListView1.ListItems.Add(, , Cad)
'                    J = 1
'                End If
                ItmX.SubItems(J) = Anyo
                ItmX.SubItems(J + 1) = "Salto desde codigo: " & Format(I, "00000000")
                
                    
                
            Else
                'HUECO
                cad = Format(I, "000000000")
'                If opcion = 5 Then
                    Set ItmX = ListView1.ListItems.Add(, , Rs!NUmSerie)
                    ItmX.SubItems(1) = cad
                    J = 2
'                Else
'                    Set ItmX = ListView1.ListItems.Add(, , Cad)
'                    J = 1
'                End If
                ItmX.SubItems(J) = Anyo
                ItmX.SubItems(J + 1) = "Falta"
                'i = RS!Codigo + 1
            End If
            ItmX.SmallIcon = 1
             I = Rs!Codigo + 1
        End If
        'Movemos siguiente
        Rs.MoveNext
        
    Wend
    Rs.Close
    
    If ListView1.ListItems.Count = 0 Then MsgBox "Proceso finalizado", vbInformation
    
    
EBuscarFacturas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    Set miRsAux = Nothing
End Sub


Private Function FacturaMinimo(Serie As String, FIni As Date, fFin As Date, Anyo As Integer) As Long
Dim C As String
Dim Campo As String
Dim F1 As Date

    If Opcion = 5 Then
        C = "Select min(numfactu) FROM factcli WHERE "
        Campo = "fecfactu"
    Else
        C = "Select min(numregis) FROM factpro WHERE "
        Campo = "fecharec"
    End If
    
    'FEHAS   INICO
    If Anyo = Year(FIni) Then
        F1 = FIni
    Else
        F1 = CDate("01/01/" & Anyo)
    End If
    C = C & Campo & " >= '" & Format(F1, FormatoFecha) & "'"
    
    If Anyo = Year(fFin) Then
        F1 = fFin
    Else
        F1 = CDate("31/12/" & Anyo)
    End If
    C = C & " AND " & Campo & " <= '" & Format(F1, FormatoFecha) & "'"
    
    If Opcion = 5 Then C = C & " AND numserie = '" & Serie & "'"
    'Debug.Print C
    FacturaMinimo = 0

    miRsAux.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then FacturaMinimo = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
End Function

Private Sub PonFocus(ByRef Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub FacturasContabilizadas()
    Label7(0).Caption = "Facturas"
    Label7(1).Caption = "Obteniendo registros"
    Me.Refresh
    'Cogemos las facturas en RS
    Sql = "Select numasien,numdiari,fechaent, "
    
    If Opcion = 5 Then
        Sql = Sql & " numserie,numfactu c, fecfactu f"
    Else
        Sql = Sql & " numregis c, fecrecpr f"
    End If
    Sql = Sql & " FROM factcli"
    If Opcion = 6 Then Sql = Sql & "prov"
    Sql = Sql & " WHERE numasien>0 "
    If Opcion = 5 Then
        CadenaDesdeOtroForm = "fecfactu"
        If txtCLI.Text <> "" Then Sql = Sql & " AND numserie ='" & txtCLI.Text & "'"
    Else
        CadenaDesdeOtroForm = "fecrecpr"
    End If
    
    If Text1(3).Text <> "" Then Sql = Sql & " AND " & CadenaDesdeOtroForm & " >='" & Format(Text1(3).Text, FormatoFecha) & "'"
    If Text1(2).Text <> "" Then Sql = Sql & " AND " & CadenaDesdeOtroForm & " <='" & Format(Text1(2).Text, FormatoFecha) & "'"
    
    Sql = Sql & " ORDER BY numdiari,numasien,fechaent"
    
    
    'Cuento el recordset
    NumRegElim = 0
    Rs.Open "SELECT count(*) " & Mid(Sql, InStr(1, Sql, " FROM ")), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then NumRegElim = DBLet(Rs.Fields(0), "N")
    Rs.Close
    espera 0.2
    If NumRegElim = 0 Then Exit Sub

    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Hay facturas. Ahora en rsaux cargare los apuntes
        Sql = "Select numasien,numdiari,fechaent from hlinapu WHERE idcontab = 'FRA"
        If Opcion = 5 Then
            Sql = Sql & "CLI'"
        Else
            Sql = Sql & "PRO'"
        End If

        If Text1(3).Text <> "" Then Sql = Sql & " AND fechaent >='" & Format(Text1(3).Text, FormatoFecha) & "'"
        If Text1(2).Text <> "" Then Sql = Sql & " AND fechaent <='" & Format(Text1(2).Text, FormatoFecha) & "'"
        Sql = Sql & " GROUP BY numasien,numdiari,fechaent"

        miRsAux.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText


    
        'Recorremos las facturas
        I = 0
        While Not Rs.EOF
            'Label7(1).Caption = RS!C & " - " & RS!F
            Label7(1).Caption = I & " de " & NumRegElim
            Label7(1).Refresh
            If Not EstaEnMirsaux Then
                InsertaItemsFacturasContabilizadas True
            End If
            Rs.MoveNext
            I = I + 1
            If (I Mod 50) = 0 Then
                Me.Refresh
                DoEvents
            End If
        Wend
'        miRsAux.Close
    End If 'Rs.eof
    Rs.Close
    
End Sub


Private Sub FacturasNoContabilizadas()
' incluimos las facturas :que tienen numero de asiento y este no existe
'                         que no tienen numero de asiento


    Label7(0).Caption = "Facturas"
    Label7(1).Caption = "Obteniendo registros"
    Me.Refresh
    'Cogemos las facturas en RS
    Sql = "Select numasien,numdiari,fechaent, "
    
    If Opcion = 5 Then
        Sql = Sql & " numserie,numfactu c, fecfactu f"
        Sql = Sql & " FROM factcli"
    Else
        Sql = Sql & " numregis c, fecharec f"
        Sql = Sql & " FROM factpro"
    End If
    Sql = Sql & " WHERE (1=1) "

    If Opcion = 5 Then
        CadenaDesdeOtroForm = "fecfactu"
    Else
        CadenaDesdeOtroForm = "fecharec"
    End If
    
    If txtCLI.Text <> "" Then Sql = Sql & " AND numserie ='" & txtCLI.Text & "'"
    
    If Text1(3).Text <> "" Then Sql = Sql & " AND " & CadenaDesdeOtroForm & " >='" & Format(Text1(3).Text, FormatoFecha) & "'"
    If Text1(2).Text <> "" Then Sql = Sql & " AND " & CadenaDesdeOtroForm & " <='" & Format(Text1(2).Text, FormatoFecha) & "'"
    
    Sql = Sql & " ORDER BY numdiari,numasien,fechaent"
    
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Recorremos las facturas
    I = 0
    While Not Rs.EOF
        'Label7(1).Caption = RS!C & " - " & RS!F
        Label7(1).Caption = I & " de " & NumRegElim
        Label7(1).Refresh
        
        
        If DBLet(Rs!NumAsien) = 0 Then
            InsertaItemsFacturasContabilizadas True
        Else
            Sql = "select * from hlinapu where numasien = " & DBSet(Rs!NumAsien, "N") & " and fechaent = " & DBSet(Rs!F, "F")
            Sql = Sql & " and numdiari = " & DBSet(Rs!NumDiari, "N")
            Sql = Sql & " and idcontab='FRA"
    
            If Opcion = 5 Then
                Sql = Sql & "CLI'"
            Else
                Sql = Sql & "PRO'"
            End If
            
            
            If TotalRegistros(Sql) = 0 Then InsertaItemsFacturasContabilizadas True
            
        End If
        
        
        Rs.MoveNext
        I = I + 1
        If (I Mod 50) = 0 Then
            Me.Refresh
            DoEvents
        End If
    Wend
    
    Rs.Close
    
End Sub






Private Function EstaEnMirsaux() As Boolean
Dim Fin As Boolean
    EstaEnMirsaux = False
    If Not miRsAux.EOF Then miRsAux.MoveFirst
    If miRsAux.EOF Then Exit Function
    Fin = False
    While Not Fin
        If miRsAux!NumDiari = Rs!NumDiari Then
              If miRsAux!NumAsien = Rs!NumAsien Then
                If miRsAux!FechaEnt = Rs!FechaEnt Then
                    Fin = True
                    EstaEnMirsaux = True
                End If
            End If
        End If
        miRsAux.MoveNext
        If miRsAux.EOF Then Fin = True
    Wend
    
End Function


Private Sub InsertaItemsFacturasContabilizadas(RegistroFactura As Boolean)
    NumCuentas = NumCuentas + 1
    Set ItmX = ListView1.ListItems.Add(, "C" & NumCuentas)
    If RegistroFactura Then
        Sql = "  " & Format(Rs!C, "00000000") & "   " & Format(Rs!F, "dd/mm/yyyy")
        If Opcion = 5 Then Sql = Rs!NUmSerie & Sql
        ItmX.Text = Sql
        ItmX.SubItems(1) = " **** "
        
        If DBLet(Rs!NumAsien, "N") <> 0 Then
            ItmX.SubItems(2) = "   No existe asiento " & Format(DBLet(Rs!NumAsien), "0000000")
        Else
            ItmX.SubItems(2) = "   No tiene asiento para factura. "
        End If
    Else
        ItmX.Text = " **** "
        ItmX.SubItems(1) = Rs!NumDiari & "  " & Format(Rs!NumAsien, "0000000") & " " & Format(Rs!FechaEnt, "dd/mm/yyyy")
        ItmX.SubItems(2) = "   No existe factura para asiento."
    End If
End Sub



Private Sub ApuntesSinFactura()
    Label7(0).Caption = "Asientos"
    Label7(1).Caption = "Obteniendo registros"
    Me.Refresh


    Sql = "Select numasien,numdiari,fechaent FROM hlinapu WHERE idcontab='FRA"
    
    If Opcion = 5 Then
        Sql = Sql & "CLI'"
    Else
        Sql = Sql & "PRO'"
    End If
    If Text1(3).Text <> "" Then Sql = Sql & " AND fechaent >='" & Format(Text1(3).Text, FormatoFecha) & "'"
    If Text1(2).Text <> "" Then Sql = Sql & " AND fechaent <='" & Format(Text1(2).Text, FormatoFecha) & "'"
        
    Sql = Sql & " GROUP BY numasien,numdiari,fechaent"
    Sql = Sql & " ORDER BY numdiari,numasien,fechaent"
    
    
    'Cuento el recordset
    NumRegElim = 0
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        NumRegElim = NumRegElim + 1
        Rs.MoveNext
    Wend
    Rs.Close
    espera 0.2
    If NumRegElim = 0 Then Exit Sub



    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Hay apuntes. Busco sus facturaas
        Sql = "Select numasien,numdiari,fechaent "
        Sql = Sql & " FROM factcli"
        If Opcion = 6 Then Sql = Sql & "prov"
        Sql = Sql & " WHERE numasien>0 "
        If Opcion = 5 Then
            CadenaDesdeOtroForm = "fecfactu"
            If txtCLI.Text <> "" Then Sql = Sql & " AND numserie ='" & txtCLI.Text & "'"
        Else
            CadenaDesdeOtroForm = "fecrecpr"
        End If

        If Text1(3).Text <> "" Then Sql = Sql & " AND " & CadenaDesdeOtroForm & " >='" & Format(Text1(3).Text, FormatoFecha) & "'"
        If Text1(2).Text <> "" Then Sql = Sql & " AND " & CadenaDesdeOtroForm & " <='" & Format(Text1(2).Text, FormatoFecha) & "'"
        Sql = Sql & " GROUP BY numasien,numdiari,fechaent"

        Sql = Sql & " ORDER BY numdiari,numasien,fechaent"




        miRsAux.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText


        'Recorremos las facturas
        I = 0
        While Not Rs.EOF
            'Label7(1).Caption = RS!Numasien & " " & RS!fechaent
            Label7(1).Caption = I & " de " & NumRegElim
            Label7(1).Refresh
            If Not EstaEnMirsaux Then

                InsertaItemsFacturasContabilizadas False
            End If
            Rs.MoveNext
            I = I + 1
            If (I Mod 50) = 0 Then
                Me.Refresh
                DoEvents
            End If
        Wend
        miRsAux.Close
    End If 'Rs.eof
    Rs.Close
    
End Sub


Private Sub ApuntesSinFacturaNew()
    Label7(0).Caption = "Asientos"
    Label7(1).Caption = "Obteniendo registros"
    Me.Refresh


    Sql = "Select numasien,numdiari,fechaent FROM hlinapu WHERE idcontab='FRA"
    
    If Opcion = 5 Then
        Sql = Sql & "CLI'"
    Else
        Sql = Sql & "PRO'"
    End If
    If Text1(3).Text <> "" Then Sql = Sql & " AND fechaent >='" & Format(Text1(3).Text, FormatoFecha) & "'"
    If Text1(2).Text <> "" Then Sql = Sql & " AND fechaent <='" & Format(Text1(2).Text, FormatoFecha) & "'"
        
    Sql = Sql & " GROUP BY numasien,numdiari,fechaent"
    Sql = Sql & " ORDER BY numdiari,numasien,fechaent"
    
    
    'Cuento el recordset
    NumRegElim = TotalRegistrosConsulta(Sql)
    espera 0.2
        
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        Label7(1).Caption = I & " de " & NumRegElim
        Label7(1).Refresh
    
    
        If Opcion = 5 Then
            Sql = "Select numasien,numdiari,fechaent "
            Sql = Sql & " FROM factcli"
            If Opcion = 6 Then Sql = Sql & "prov"
            Sql = Sql & " WHERE numasien= " & DBSet(Rs!NumAsien, "N")
            Sql = Sql & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
            Sql = Sql & " and numdiari = " & DBSet(Rs!NumDiari, "N")
        Else
            Sql = "Select numasien,numdiari,fechaent "
            Sql = Sql & " FROM factpro"
            Sql = Sql & " WHERE numasien= " & DBSet(Rs!NumAsien, "N")
            Sql = Sql & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
            Sql = Sql & " and numdiari = " & DBSet(Rs!NumDiari, "N")
        End If
        If TotalRegistrosConsulta(Sql) = 0 Then
             InsertaItemsFacturasContabilizadas False
        End If
        Rs.MoveNext
        I = I + 1
        If (I Mod 50) = 0 Then
            Me.Refresh
            DoEvents
        End If
    Wend
        
    
End Sub


Private Sub QuitarOtrasCuentas()
Dim I As Integer
    Set Rs = New ADODB.Recordset
    
    'pRIMERO DE LAS CUENTAS BANCARIAS
    'codmacta ctagastos ctaingreso ctagastostarj    bancos
    Label4.Caption = "Cta bancaria "
    Label4.Refresh
    Sql = "Select codmacta , ctagastos , ctaingreso , ctagastostarj   FROM bancos"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = "|"
    While Not Rs.EOF
        For I = 0 To 3
            If Not IsNull(Rs.Fields(I)) Then
                If Rs.Fields(I) <> "" Then
                    If InStr(1, Sql, "|" & Rs.Fields(I) & "|") = 0 Then Sql = Sql & Rs.Fields(I) & "|"
                End If
            End If
        Next
        Rs.MoveNext
    Wend
    Rs.Close
    
    Sql = Mid(Sql, 2)
    While Sql <> ""
        I = InStr(1, Sql, "|")
        Conn.Execute "Delete from tmpbussinmov where codmacta ='" & RecuperaValor(Sql, 1) & "';"
        Sql = Mid(Sql, I + 1)
    Wend
    
    'PARAMETROS
    Sql = "Delete from tmpbussinmov where codmacta ='" & vParam.ctaperga & "';"
    Conn.Execute Sql
    espera 0.2
    
    If vEmpresa.TieneTesoreria Then
        Sql = "SELECT ctabenbanc  ,par_pen_apli,RemesaCancelacion,RemesaConfirmacion,taloncta,pagarectaPRO,talonctaPRO,ctaefectcomerciales from paramtesor"
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            For I = 0 To Rs.Fields.Count - 1
                Sql = DBLet(Rs.Fields(I), "T")
                If Sql <> "" Then
                    If Len(Sql) = vEmpresa.DigitosUltimoNivel Then
                        Sql = "Delete from tmpbussinmov where codmacta ='" & Sql & "';"
                        Conn.Execute Sql
                    End If
                End If
            Next I
        End If
        Rs.Close
    End If
    'IVAS
    Label4.Caption = "IVAS "
    Label4.Refresh
    Sql = "SELECT cuentare ,cuentarr ,cuentaso ,cuentasr ,cuentasn from tiposiva "
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = "|"
    While Not Rs.EOF
        For I = 0 To 4
            If Not IsNull(Rs.Fields(I)) Then
                If Rs.Fields(I) <> "" Then
                    If InStr(1, Sql, "|" & Rs.Fields(I) & "|") = 0 Then Sql = Sql & Rs.Fields(I) & "|"
                End If
            End If
        Next
        Rs.MoveNext
    Wend
    Rs.Close
    
    Sql = Mid(Sql, 2)
    While Sql <> ""
        I = InStr(1, Sql, "|")
        Conn.Execute "Delete from tmpbussinmov where codmacta ='" & RecuperaValor(Sql, 1) & "';"
        Sql = Mid(Sql, I + 1)
    Wend
    Set Rs = Nothing
    
    
    
    
End Sub
    

