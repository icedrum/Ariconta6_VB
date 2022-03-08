VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESRealizarPagos 
   Caption         =   "Realizar Pago"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   17520
   Icon            =   "frmTESRealizarPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   17520
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   750
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESRealizarPagos.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESRealizarPagos.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESRealizarPagos.frx":6B88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frame 
      Height          =   2325
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   17235
      Begin VB.Frame FrameRemesar 
         BorderStyle     =   0  'None
         Height          =   1905
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   16935
         Begin VB.CheckBox chkSoloBancoPrev 
            Caption         =   "Sólo Banco previsto"
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
            Left            =   12360
            TabIndex        =   34
            Top             =   1560
            Width           =   2325
         End
         Begin VB.TextBox Text3 
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
            Height          =   370
            Index           =   6
            Left            =   10680
            TabIndex        =   30
            Text            =   "0000000000"
            Top             =   960
            Width           =   1065
         End
         Begin VB.CheckBox chkVerPendiente 
            Caption         =   "Ver lo Pdte del proveedor"
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
            Left            =   8520
            TabIndex        =   29
            Top             =   1560
            Width           =   2865
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmTESRealizarPagos.frx":6EA2
            Left            =   8520
            List            =   "frmTESRealizarPagos.frx":6EA4
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "Tipo de pago|N|N|||formapago|tipforpa|||"
            Top             =   330
            Width           =   3255
         End
         Begin VB.TextBox txtCta 
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
            Height          =   370
            Index           =   5
            Left            =   1350
            TabIndex        =   5
            Text            =   "0000000000"
            Top             =   1470
            Width           =   1305
         End
         Begin VB.TextBox txtDCta 
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
            Height          =   370
            Index           =   4
            Left            =   2670
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "Text2"
            Top             =   930
            Width           =   5715
         End
         Begin VB.TextBox txtCta 
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
            Height          =   370
            Index           =   4
            Left            =   1350
            TabIndex        =   4
            Text            =   "0000000000"
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox Text3 
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
            Height          =   370
            Index           =   1
            Left            =   4230
            TabIndex        =   1
            Text            =   "0000000000"
            Top             =   330
            Width           =   1305
         End
         Begin VB.TextBox Text3 
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
            Height          =   370
            Index           =   2
            Left            =   7050
            TabIndex        =   2
            Text            =   "0000000000"
            Top             =   330
            Width           =   1305
         End
         Begin VB.CheckBox chkGenerico 
            Caption         =   "Cuenta genérica"
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
            Left            =   12360
            TabIndex        =   22
            Top             =   1230
            Width           =   2175
         End
         Begin VB.CheckBox chkPorFechaVenci 
            Caption         =   "Contab. por fecha vto."
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
            Left            =   12360
            TabIndex        =   21
            Top             =   900
            Width           =   2865
         End
         Begin VB.CheckBox chkContrapar 
            Caption         =   "Agrupar por Banco"
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
            Left            =   12360
            TabIndex        =   20
            Top             =   570
            Width           =   2745
         End
         Begin VB.CheckBox chkAsiento 
            Caption         =   "Un Asiento por recibo"
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
            Left            =   12360
            TabIndex        =   19
            Top             =   240
            Width           =   2385
         End
         Begin VB.TextBox txtDCta 
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
            Height          =   370
            Index           =   5
            Left            =   2670
            TabIndex        =   14
            Text            =   "Text3"
            Top             =   1470
            Width           =   5745
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   370
            Index           =   0
            Left            =   1380
            TabIndex        =   0
            Text            =   "Text3"
            Top             =   330
            Width           =   1365
         End
         Begin MSComctlLib.Toolbar ToolbarAyuda 
            Height          =   390
            Left            =   16590
            TabIndex        =   32
            Top             =   300
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
         Begin VB.CheckBox chkVtoCuenta 
            Caption         =   "Agrupar por Proveedor"
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
            Left            =   5400
            TabIndex        =   23
            Top             =   1440
            Width           =   2745
         End
         Begin VB.Label Label2 
            Caption         =   "Gastos banco"
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
            Left            =   8520
            TabIndex        =   31
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Pago"
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
            Index           =   8
            Left            =   8580
            TabIndex        =   28
            Top             =   60
            Width           =   2025
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   1
            Left            =   1080
            Top             =   1530
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   0
            Left            =   1080
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Banco"
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
            Left            =   60
            TabIndex        =   27
            Top             =   930
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Desde Vto"
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
            Index           =   5
            Left            =   2850
            TabIndex        =   25
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta Vto"
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
            Index           =   4
            Left            =   5640
            TabIndex        =   24
            Top             =   330
            Width           =   1125
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   3930
            Picture         =   "frmTESRealizarPagos.frx":6EA6
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   6780
            Picture         =   "frmTESRealizarPagos.frx":6F31
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1110
            Picture         =   "frmTESRealizarPagos.frx":6FBC
            ToolTipText     =   "Cambiar fecha contabilizacion"
            Top             =   330
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   16530
            Picture         =   "frmTESRealizarPagos.frx":72FE
            ToolTipText     =   "Seleccionar todos"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   16170
            Picture         =   "frmTESRealizarPagos.frx":7448
            ToolTipText     =   "Quitar seleccion"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
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
            TabIndex        =   16
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedor"
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
            Left            =   60
            TabIndex        =   15
            Top             =   1500
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   7740
      Width           =   16935
      Begin VB.CommandButton cmdImprimir 
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
         Left            =   1800
         TabIndex        =   35
         Top             =   120
         Width           =   1305
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Contabilizar"
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
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   8580
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   12060
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Seleccionado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   3780
         TabIndex        =   18
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " PENDIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   1
         Left            =   10650
         TabIndex        =   11
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Vencido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   7530
         TabIndex        =   9
         Top             =   180
         Width           =   990
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5025
      Left            =   60
      TabIndex        =   7
      Top             =   2460
      Width           =   17205
      _ExtentX        =   30348
      _ExtentY        =   8864
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin VB.Menu mnContextual 
      Caption         =   "Contextual"
      Visible         =   0   'False
      Begin VB.Menu mnNumero 
         Caption         =   "Poner numero Talón/Pagaré"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSelectAll 
         Caption         =   "Seleccionar todos"
      End
      Begin VB.Menu mnQUitarSel 
         Caption         =   "Quitar selección"
      End
   End
End
Attribute VB_Name = "frmTESRealizarPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 804



Public vSql As String

Public OrdenarEfecto As Boolean
Public Regresar As Boolean
Public vTextos As String  'Dependera de donde venga
Public SegundoParametro As String



Dim ContabTransfer2 As Boolean  'falta quitar esta variable. La he pasado de public a DIM


    'Diversas utilidades
    '-------------------------------------------------------------------------------
    'Para las transferencias me dice que transferencia esta siendo creada/modificada
    '
    'Para mostrar un check con los efectos k se van a generar en remesa y/o pagar
 
 
 ' 13 Mayo 08
    ' Cuando contabilice el los cobros por tarjeta entonces
    ' si lleva gastos los añadire
Public ImporteGastosTarjeta_ As Currency   'Para cuando viene de recepciondocumentos pondre el importe que le falta
                                          ' y asi ofertarlo al divisonvencimiento
     '-ABRIL 2014.  Navarres. Llevara el % interes
 
 
 
 
'Agosto 2009
'Desde recepcion de talones.
'Tendra la posibilidad de desdoblar un vencimiento
Public DesdeRecepcionTalones As Boolean
 
'Febrero 2010
'Para el pago de talones y pagareses ;)
'Enviara el nº de talon/pagare
Public NumeroTalonPagere As String


'Marzo 2013
'Cuando cobro/pago un mismo clie/prov aparecera un icono para poder añadir
'cualquier cobro /pago del mismo. Se contabilizaran con los datos pendientes
Public CodmactaUnica As String

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1
Private frmTESTalPag As frmTESRefTalonPagos

Dim Cad As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Fecha As Date
Dim Importe As Currency
Dim Vencido As Currency
Dim impo As Currency
Dim riesgo As Currency

Dim ImpSeleccionado As Currency
Dim i As Integer
Private PrimeraVez As Boolean
Dim RiesTalPag As Currency
Private FechaAsiento As Date
Private vp As Ctipoformapago
Private SubItemVto As Integer

Private DescripcionTransferencia As String
Private GastosTransferencia As Currency

Dim ObservacionApunteBanco As String

Dim CampoOrden As String
Dim Orden As Boolean
Dim Campo2 As Integer
Dim CtaAnt As String
Dim CtaAntBan As String



Dim FrasMarcadas As Collection



Private Sub chkAsiento_Click(Index As Integer)
   'Es incompatible asiento por pago y agrupar apunte bancario
   If chkAsiento(Index).Value = 1 Then
        If chkContrapar(Index).Value = 1 Then
            Incompatibilidad
            chkContrapar(Index).Value = 0
        End If
   End If
       
End Sub

Private Sub Incompatibilidad()
    If Not PrimeraVez Then MsgBox "Es incompatible agrupar apunte bancario y asiento por pago", vbExclamation
End Sub


Private Sub chkContrapar_Click(Index As Integer)
   'Es incompatible asiento por pago y agrupar apunte bancario
   If chkContrapar(Index).Value = 1 Then
        If chkAsiento(Index).Value = 1 Then
            Incompatibilidad
            chkAsiento(Index).Value = 0
        End If
   End If
End Sub

Private Sub chkGenerico_Click(Index As Integer)
    If chkGenerico(Index).Value = 0 Then
        chkGenerico(Index).FontBold = False
        chkGenerico(Index).Tag = ""
    End If
End Sub



Private Sub Generar2()
Dim Contador2 As Integer
Dim F2 As Date
Dim TipoAnt As Integer
Dim Aux As String
Dim SQL As String
Dim MasDeUnVto As Boolean
Dim MasDeUnProveedor As Boolean
Dim HacerPreguntaEspecial As Boolean

     'If vp.tipoformapago = vbTransferencia Then
    Aux = ""
    Cad = ""
    SQL = ""
    MasDeUnProveedor = False
    Importe = 0
    
    For i = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            Cad = Cad & "1"
            Importe = Importe + ImporteFormateado(ListView1.ListItems(i).SubItems(9))
            If Not MasDeUnProveedor Then
                If Aux = "" Then
                    Aux = ListView1.ListItems(i).Tag
                    SQL = ListView1.ListItems(i).SubItems(5)
                Else
                    If Aux <> ListView1.ListItems(i).Tag Then MasDeUnProveedor = True
                End If
            End If
        End If
    Next i
    If Cad = "" Then
        MsgBox "Deberias selecionar algún vencimiento", vbExclamation
        Exit Sub
    End If
    
    If Combo1.ListIndex = -1 Then
        MsgBox "Debe selecionar el tipo de pago", vbExclamation
        Exit Sub
    End If
         
    If txtCta(4).Text = "" Then
        MsgBox "Deberias introducir la cuenta de banco", vbExclamation
        Exit Sub
    End If
    MasDeUnVto = Len(Cad) > 1
    Aux = ""
    If Not MasDeUnProveedor Then
        Aux = SQL
    Else
        Aux = "Vtos: " & Len(Cad)
    End If
    CtaAntBan = Len(Cad) & "|" & Format(Importe, FormatoImporte) & "|" & Aux & "|"
    
    
    
    
    
    
    
    Importe = 0
    If Combo1.ItemData(Combo1.ListIndex) = 6 Then
        If Text3(6).Text <> "" Then
            If InStr(1, Text3(6).Text, ",") > 0 Then
                Importe = ImporteFormateado(Text3(6).Text)
            Else
                Importe = CCur(TransformaPuntosComas(Text3(6).Text))
            End If
        End If
    End If
    If vParamT.IntereseCobrosTarjeta2 > 0 Then
        If Importe < 0 Or Importe >= 100 Then
            MsgBox "Intereses cobro tarjeta. Valor entre 0..100", vbExclamation
            PonFoco Me.Text3(6)
            Exit Sub
            
        End If
        
        'Solo dejaremos IR cliente a cliente
        If Me.txtCta(5).Text = "" And Importe > 0 Then
            MsgBox "Seleccione una cuenta cliente", vbExclamation
            PonFoco Me.txtCta(5)
            Exit Sub
        End If
    End If
    
    
    
    
    'Alguna comprobacion
    'Si es un cobro, por tarjeta y tiene gastos
    'entonces tendra que ir todo en un unico apunte
    If False And (Combo1.ItemData(Combo1.ListIndex) = 6 Or Combo1.ItemData(Combo1.ListIndex) = 0) And ImporteGastosTarjeta_ > 0 Then
        Cad = ""
        '-----------------------------------------------------
        If Me.chkAsiento(0).Value Then
            Cad = "No debe marcar la opcion de varios asientos"
        Else
            If Me.chkPorFechaVenci.Value Then
                riesgo = 0
                For i = 1 To Me.ListView1.ListItems.Count
                    If ListView1.ListItems(i).Checked Then
                        
                        Fecha = ListView1.ListItems(i).SubItems(3)
                        If riesgo = 0 Then
                            F2 = Fecha
                            riesgo = 1
                        Else
                            'Si las fechas son distintas NO dejo seguir
                            If F2 <> Fecha Then
                                Cad = "Debe contabilizarlo todo en un único apunte"
                                Exit For
                            End If
                        End If
                    End If
        
                Next i
            End If
        End If
            
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
        
        
        'Compruebo que tiene configurada la cuenta de gastos de tarjeta
        If Combo1.ItemData(Combo1.ListIndex) = 6 Then   'SOLO TARJETA
            Cad = DevuelveDesdeBD("ctagastostarj", "bancos", "codmacta", Text3(1).Tag, "T")
            If Cad = "" Then
                MsgBox "Falta configurar la cuenta de gastos de tarjeta", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    
    
    'Fecha dentro de ejercicios
    Cad = ""
    If CDate(Text3(0).Text) < vParam.fechaini Then
       Cad = "Fuera de ejercios."
    Else
        Fecha = DateAdd("yyyy", 1, vParam.fechafin)
        If CDate(Text3(0).Text) > Fecha Then Cad = "Fecha de ejercicio aun no abierto. "
    End If
    
    If Cad <> "" Then
         MsgBox Cad, vbExclamation
         PonFoco Text3(0)
        Exit Sub
    End If
    

    If Combo1.ItemData(Combo1.ListIndex) = 3 Then
        Cad = ""
        For i = 1 To Me.ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
                If Me.ListView1.ListItems(i).ForeColor = vbRed Then
                    Cad = Cad & "1"
                    Exit For
                End If
            End If
        Next i
    
        If Cad <> "" Then
            'Significa que ha marcado alguno de los vencimientos que emitiero documento. Veremos si estan todos marcados
            Cad = ""
            For i = 1 To Me.ListView1.ListItems.Count
                If Not ListView1.ListItems(i).Checked Then
                    If Me.ListView1.ListItems(i).ForeColor = vbRed Then
                        Cad = Cad & "1"
                        Exit For
                    End If
                End If
            Next i
            
            If Cad <> "" Then
                Cad = "Ha seleccionado vencimientos que emitió documento, pero no estan todos seleccionados." & vbCrLf
                Cad = Cad & vbCrLf & "¿Es correcto?"
                If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
        End If
    End If

    
    'Embargado
    Cad = ""
    For i = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            If InStr(1, Cad, ListView1.ListItems(i).Tag) = 0 Then
                Cad = Cad & ", " & DBSet(ListView1.ListItems(i).Tag, "T")
            End If
        End If
    Next i

    If Cad <> "" Then
        'No deberia ser ""
        Cad = Mid(Cad, 2)
        Set miRsAux = New ADODB.Recordset
        Cad = "Select codmacta,nommacta from cuentas where embargo=1 AND codmacta in (" & Cad & ")"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not miRsAux.EOF
            Cad = Cad & miRsAux!codmacta & " " & miRsAux!Nommacta & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        If Cad <> "" Then
            MsgBox "Cuentas en situacion de embargo:" & vbCrLf & Cad, vbExclamation
            Exit Sub
        End If
    End If
    
    
    HacerPreguntaEspecial = True
    Cad = "Desea contabilizar los vencimientos seleccionados?"
    If Combo1.ItemData(Combo1.ListIndex) = 1 Then
        
        HacerPreguntaEspecial = False
        i = 0
        If Not ContabTransfer2 Then i = 1
        If i = 1 Then
            'Estamos creando la transferencia o el pago domiciliado
            Cad = RecuperaValor(Me.vTextos, 5)
            If Cad = "" Then
                Cad = "Desea generar la transferencia?"
            Else
                Cad = "Desea generar el " & Cad & "?"
            End If
            
        End If
    Else
        i = 1
        If chkContrapar(0).Value = 0 Then i = 0
        If chkAsiento(0).Value = 1 Then i = 0
        
        If i = 1 Then
            If Not MasDeUnVto Then HacerPreguntaEspecial = False
        Else
            HacerPreguntaEspecial = False
        End If
    End If
    
    
    
    If Combo1.ItemData(Combo1.ListIndex) = vbConfirming Then
        If Me.chkGenerico(0).Value = 1 Then
            MsgBox "No puede marcar cuenta generica", vbExclamation
            Exit Sub
        End If
    End If
    
    ObservacionApunteBanco = ""
    SubItemVto = 0
    CadenaDesdeOtroForm = CStr(CtaAntBan)
    CtaAntBan = ""
    If HacerPreguntaEspecial Then
        'Abriremos un FORM listado
        frmMensajes.Opcion = 71
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            SubItemVto = 1
        Else
            ObservacionApunteBanco = Mid(CadenaDesdeOtroForm, 2)
            CadenaDesdeOtroForm = ""
        End If
        
        
    Else
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then SubItemVto = 1
    End If
    If SubItemVto = 1 Then Exit Sub

    
    
    SQL = "delete from tmppendientes where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    
    Screen.MousePointer = vbHourglass
    
    'Una cosa mas.
    'Si la forma de pago es talon/pagere, y me ha escrito el numero de talon pagare...
    'Se lo tengo que pasar a la contabilizacion, con lo cual tendre que grabar
    'el nº de talon pagare en reftalonpag
    
    'Si el parametro dice k van todos en el mismo asiento, pues eso, todos en el mismo asiento
    'Primero leemos la forma de pago, el tipo perdon
    Set vp = New Ctipoformapago

    'en vtextos, en el 3 tenemos la forpa
    Cad = ""
    Cad = Combo1.ItemData(Combo1.ListIndex) 'RecuperaValor(vTextos, 3)
    If Cad = "" Then
        i = -1
    Else
        i = Val(Cad)
    End If
    If vp.Leer(i) = 1 Then
        'ERROR GRAVE LEYENDO LA FORMA DE PAGO
        Screen.MousePointer = vbDefault
        Set vp = Nothing
        End
    End If
    
    
    '--------------------------------------------------------
    'Si es realizar transferencia, crearemos la transferencia
    '--------------------------------------------------------
           
    
    '-----------------------------------------------------
    If Me.chkPorFechaVenci.Value Then
        'Contabilizaremos por fecha de vencimiento
        'Haremos una comrpobacion. Miraremos que todos los recibos marcados para
        'contabilizar , si la fecha no pertenece a actual y siguiente lo contabilizaremos con fecha
        'de cobro, es decir, la fecha con la que viene del otro form

        F2 = DateAdd("yyyy", 1, vParam.fechafin)
        Importe = 0
        riesgo = 0
        Cad = ""
        
        SubItemVto = 2
        
        For i = 1 To Me.ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
                Fecha = ListView1.ListItems(i).SubItems(SubItemVto)
                riesgo = 0
                If Fecha < vParam.fechaini Or Fecha > F2 Then
                    riesgo = 1
                Else
                    If Fecha < vParamT.fechaAmbito Then riesgo = 1
                End If
                If riesgo = 1 Then
                    If InStr(1, Cad, Format(Fecha, "dd/mm/yyyy")) = 0 Then
                        Cad = Cad & "    " & Format(Fecha)
                        Importe = Importe + 1
                        If Importe > 5 Then
                            Cad = Cad & vbCrLf
                            Importe = 0
                        End If
                    End If
                End If
            End If
        Next i
    
        If Cad <> "" Then
            Cad = "Las siguientes fechas están fuera de ejercicio (actual y siguiente):" & vbCrLf & vbCrLf & Cad
            Cad = Cad & vbCrLf & vbCrLf & "Se contabilizarán con fecha: " & Text3(0).Text & vbCrLf
            Cad = Cad & "¿Desea continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then Cad = ""
                
        End If
        Importe = 0
        riesgo = 0
    End If
    
    
    DescripcionTransferencia = ""
    If ContabTransfer2 Then
        'Estamos contabilizando la transferencia
        Cad = "stransfer"
        DescripcionTransferencia = DevNombreSQL(DevuelveDesdeBD("descripcion", Cad, "codigo", SegundoParametro, "N"))
        
    End If

    
    Cad = "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    Conn.Execute Cad


    Conn.BeginTrans
    
    'Si hay que generar la
    
    
    If HacerNuevaContabilizacion Then
        
        
        Conn.CommitTrans
              
        'Tenemos k borrar los listview
        For i = (ListView1.ListItems.Count) To 1 Step -1
            If ListView1.ListItems(i).Checked Then
'--
               EliminarCobroPago i
              
               ListView1.ListItems.Remove i
                
            End If
        Next i
        '-----------------------------------------------------------
    Else
        TirarAtrasTransaccion
    End If


    ImpSeleccionado = 0
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)

    
        
    Set vp = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

Private Function HacerNuevaContabilizacion() As Boolean
    On Error GoTo EHacer
    HacerNuevaContabilizacion = False
    
    'Paso1. Meto todos los seleccionados en una tabla
    If Not InsertarPagosEnTemporal2 Then Exit Function
    
    
    
    'Paso 2
    'Compruebo que los vtos a cobrar no tienen ni la cuenta bloqueada, ni,
    'si contabilizo por fecha de bloqueo, alguna de los vencimienotos
    'esta fuera del de fechas
    If Not ComprobarCuentasBloquedasYFechasVencimientos Then Exit Function
    
    
    
    'Contabilizo desde la tabla. Asi puedo agrupar mejor
    ContablizaDesdeTmp
    
    HacerNuevaContabilizacion = True
    
    
    Exit Function
EHacer:
    MuestraError Err.Number, "Contabilizando" & Err.Description
End Function



Private Sub Imprimir()

    frmTESImpRecibo.documentoDePago = CadenaDesdeOtroForm
    frmTESImpRecibo.pFecFactu = Text3(0).Text
    frmTESImpRecibo.Show vbModal
                                                                         
End Sub







Private Sub Refrescar()
    Screen.MousePointer = vbHourglass
    CargaList
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkVerPendiente_Click()
    CargaList
End Sub

Private Sub cmdAceptar_Click()
    Generar2
End Sub


Private Sub Combo1_Click()
    Combo1_Validate False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex = -1 Then
       ' MsgBox "Debe introducir un tipo de Pago. Revise.", vbExclamation
       ' Combo1.SetFocus
    
    Else
         'If Tipo = 1 And OrdenarEfecto And Not Cobros Then cmdGenerar.Caption = "Transferencia"
         i = 0
         If Combo1.ItemData(Combo1.ListIndex) = 1 And Me.SegundoParametro <> "" Then
             If Not ContabTransfer2 Then
                 i = 1
                 Cad = RecuperaValor(vTextos, 5) 'Dira si es PAGO DOMICILIADO
                 If Cad <> "" Then
                     If vParamT.PagosConfirmingTipo2 = 0 Then
'                         Me.Toolbar1.Buttons(2).ToolTipText = "Confirming"
                     Else
'                         Me.Toolbar1.Buttons(2).ToolTipText = "PAGO DOM."
                     End If
                 Else
'                     Me.Toolbar1.Buttons(2).ToolTipText = "Transferencia"
                 End If
             Else
                 'Caption = "ORDENAR pagos"
                 'Es una transferencia o, si es PAGO, puede ser un pago domiciliado
                 'If Not Cobros Then
                     Cad = RecuperaValor(vTextos, 5)
                     If Cad <> "" Then Caption = "Realizar pago DOMICILIADO "
                 'End If
             End If
         End If
         
         Me.chkPorFechaVenci.visible = i = 0
         chkGenerico(0).visible = i = 0
'         Me.chkVtoCuenta(0).Visible = I = 0
         
        '++
        If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
            If vParamT.IntereseCobrosTarjeta2 > 0 Then
                ImporteGastosTarjeta_ = vParamT.IntereseCobrosTarjeta2
                Text3(6).Text = Format(ImporteGastosTarjeta_, "##0.00")
            End If
        End If
        
        i = 0
        'If Cobros And (Combo1.ItemData(Combo1.ListIndex) = 2 Or Combo1.ItemData(Combo1.ListIndex) = 3) Then i = 1
        Me.mnbarra1.visible = i = 1
        Me.mnNumero.visible = i = 1
        
        
    
    
        CargaList
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        espera 0.1
        'Cargamos el LIST
        
'        CargaList
'''''
'''''        'OCTUBRE 2014
'''''        'PAgos por ventanilla.
'''''        'Pondre como fecha de vencimiento la fecha que el
'''''        'banco, en el fichero, me indica que realio el pago
'''''        If Cobros And Combo1.ListIndex <> -1 Then
'''''            If Combo1.ItemData(Combo1.ListIndex) = 0 Then
'''''                If InStr(1, vSql, "from tmpconext  WHERE codusu") > 0 Then AjustarFechaVencimientoDesdeFicheroBancario
'''''            End If
'''''        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub
 

Private Sub DevuelveCadenaPorTipo(Impresion As Boolean, ByRef Cad1 As String)

    
    Cad1 = ""
    
    '++
    If Combo1.ListIndex = -1 Then Exit Sub
    
    Select Case Combo1.ItemData(Combo1.ListIndex)
    Case 0
        If Impresion Then
            Cad1 = "He recibido mediante efectivo de"
        Else
            Cad1 = "[EFECTIVO]"
        End If
        
    Case 1
        Cad1 = "[TRANSFERENCIA]"
    Case 2
        If Impresion Then
            Cad1 = "He recibido mediante TALON de"
        Else
            Cad1 = "[TALON]"
        End If
    Case 3
        If Impresion Then
            Cad1 = "He recibido mediante PAGARE de"
        Else
            Cad1 = "[PAGARE]"
        End If
    
    Case 4
        Cad1 = "[RECIBO BANCARIO]"
    
    Case 5
        Cad1 = "[CONFIRMING]"
    
    Case 6
        If Impresion Then
            Cad1 = "He recibido mediante TARJETA DE CREDITO de"
        Else
            Cad1 = "[TARJETA CREDITO]"
        End If
    
    Case Else
        
        
    End Select
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    Limpiar Me
    Me.Icon = frmppal.Icon
    For i = 0 To imgFecha.Count - 1
        Me.imgFecha(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
    For i = 0 To imgCuentas.Count - 1
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    ContabTransfer2 = False
    CargaCombo
    
    CargaIconoListview Me.ListView1
    ListView1.Checkboxes = True
    imgCheck(0).visible = True
    imgCheck(1).visible = True
    chkPorFechaVenci.Value = 0
'    Me.cmdDividrVto.Visible = Me.DesdeRecepcionTalones  'Para poder dividir vto
    
'20170113: tarea de poner la ayuda en la fecha
'    imgFecha(2).Visible = False 'Para cambiar la fecha de contabilizacion de los pagos
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
     
    Text3(0).Text = Format(Now, "dd/mm/yyyy")
    Fecha = Now
    Me.ImporteGastosTarjeta_ = 0
    Me.CodmactaUnica = ""
     
    DevuelveCadenaPorTipo False, Cad
    LeerparametrosContabilizacion

    'Efectuar cobros
    FrameRemesar.visible = True
    ListView1.SmallIcons = Me.ImageList1
    CargaColumnas
    
    
''''    'Octubre 2014
''''    'Norma 57 pagos ventanilla
''''    'Si en el select , en el SQL, viene un
''''    If Cobros And Combo1.ListIndex <> -1 Then
''''        If Combo1.ItemData(Combo1.ListIndex) = 0 Then
''''            If InStr(1, vSql, "from tmpconext  WHERE codusu") > 0 Then chkPorFechaVenci.Value = 1
''''        End If
''''    End If
End Sub

Private Sub Form_Resize()
Dim i As Integer
Dim H As Integer

    If Me.WindowState = 1 Then Exit Sub  'Minimizar
    If Me.Height < 2700 Then Me.Height = 2700
    If Me.Width < 2700 Then Me.Width = 2700

    'Situamos el frame y demas
    Me.frame.Width = Me.Width - 120
    Me.Frame1.Left = Me.Width - 120 - Me.Frame1.Width
    Me.Frame1.top = Me.Height - Frame1.Height - 540 '360
    FrameRemesar.Width = Me.frame.Width - 320

    Me.ListView1.top = Me.frame.Height + 60
    Me.ListView1.Height = Me.Frame1.top - Me.ListView1.top - 60
    Me.ListView1.Width = Me.frame.Width

    'Las columnas
    H = ListView1.Tag
    ListView1.Tag = ListView1.Width - ListView1.Tag - 320 'Del margen
    For i = 1 To Me.ListView1.ColumnHeaders.Count
        If InStr(1, ListView1.ColumnHeaders(i).Tag, "%") Then
            Cad = (Val(ListView1.ColumnHeaders(i).Tag) * (Val(ListView1.Tag)) / 100)
        Else
            'Si no es de % es valor fijo
            Cad = Val(ListView1.ColumnHeaders(i).Tag)
        End If
        Me.ListView1.ColumnHeaders(i).Width = Val(Cad)
    Next i
    ListView1.Tag = H
End Sub


Private Sub CargaColumnas()
Dim ColX As ColumnHeader
Dim Columnas As String
Dim Ancho As String
Dim ALIGN As String
Dim NCols As Integer
Dim i As Integer

    ListView1.ColumnHeaders.Clear
 '  If Cobros Then
 '       NCols = 13 '11
 '       Columnas = "Serie|Factura|F.Factura|F. VTO|Nº|CLIENTE|Tipo|Importe|Gasto|Cobrado|Pendiente|"
 '       Ancho = "800|10%|12%|12%|520|23%|840|12%|8%|11%|12%|0%|0%|"
 '       ALIGN = "LLLLLLLDDDDDD"
 '
 '
 '       ListView1.Tag = 2200  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
 '
 '       If Combo1.ListIndex <> -1 Then
 '           If Combo1.ItemData(Combo1.ListIndex) = 2 Or Combo1.ItemData(Combo1.ListIndex) = 3 Then
 '               ''Si es un talon o pagare entonces añadire un campo mas
 '               NCols = NCols + 1
 '               Columnas = Columnas & "Nº Documento|"
 '               Ancho = Ancho & "2500|"
 '               ALIGN = ALIGN & "L"
 '           End If
 '       End If
 '  Else
        NCols = 12
        Columnas = "Serie|Factura|F. Fact|F. VTO|Nº|PROVEEDOR|Tipo|Importe|Pagado|Pendiente|Referencia|"
        Ancho = "800|15%|12%|12%|400|26%|800|12%|12%|12%|0%|0%|0%|"
        ALIGN = "LLLLLLLDDDDDD"
        ListView1.Tag = 2200  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
  ' End If
        
   For i = 1 To NCols
        Cad = RecuperaValor(Columnas, i)
        If Cad <> "" Then
            Set ColX = ListView1.ColumnHeaders.Add()
            ColX.Text = Cad
            'ANCHO
            Cad = RecuperaValor(Ancho, i)
            ColX.Tag = Cad
            'align
            Cad = Mid(ALIGN, i, 1)
            If Cad = "L" Then
                'NADA. Es valor x defecto
            Else
                If Cad = "D" Then
                    ColX.Alignment = lvwColumnRight
                Else
                    'CENTER
                    ColX.Alignment = lvwColumnCenter
                End If
            End If
        End If
    Next i

End Sub


Private Sub GuardarMarcados()
Dim i As Long
    
    Set FrasMarcadas = New Collection

    For i = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            FrasMarcadas.Add ListView1.ListItems(i).Text & "|" & ListView1.ListItems(i).SubItems(1) & "|" & ListView1.ListItems(i).SubItems(2) & "|" & ListView1.ListItems(i).SubItems(4) & "|" & ListView1.ListItems(i).Tag & "|"
        End If
    Next i

End Sub
 

Private Sub CargaList()
On Error GoTo ECargando

    Me.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    
    
    Set Rs = New ADODB.Recordset
'    Fecha = CDate(Text1.Text)

    GuardarMarcados

    ListView1.ListItems.Clear
    Importe = 0
    Vencido = 0
    riesgo = 0
    ImpSeleccionado = 0
    
    CargaPagos
    riesgo = ImpSeleccionado
    Set FrasMarcadas = Nothing
    
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
    Label2(2).Caption = "Selecionado"
    Label2(2).visible = True
    Text2(2).visible = True
    Label2(3).visible = True
    
ECargando:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Text2(0).Text = Format(Importe, FormatoImporte)
    Text2(1).Text = Format(Vencido, FormatoImporte)
    
    Text2(2).Text = Format(riesgo, FormatoImporte)
    riesgo = 0
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set Rs = Nothing
End Sub



Private Function DevSQL() As String
Dim Cad As String

    vSql = ""
    'Llegados a este punto montaremos el sql
    
    If Text3(1).Text <> "" Then
        If vSql <> "" Then vSql = vSql & " AND "
        vSql = vSql & " pagos.fecefect >= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    End If
        
        
    If Text3(2).Text <> "" Then
        If vSql <> "" Then vSql = vSql & " AND "
        vSql = vSql & " pagos.fecefect <= '" & Format(Text3(2).Text, FormatoFecha) & "'"
    End If

    
    'Forma de pago
    If Me.txtCta(5).Text <> "" Then
        'Los de un cliente solamente
        If vSql <> "" Then vSql = vSql & " AND "
        vSql = vSql & " pagos.codmacta = '" & txtCta(5).Text & "'"
        
        '++
        If Me.chkVerPendiente.Value = 0 Then
            If Combo1.ListIndex <> -1 Then
                If Combo1.ItemData(Combo1.ListIndex) >= 0 Then
                    If vSql <> "" Then vSql = vSql & " and "
                    vSql = vSql & " formapago.tipforpa = " & Combo1.ItemData(Combo1.ListIndex)
                End If
            End If
        End If
    Else
        If Combo1.ListIndex >= 0 Then
            If Combo1.ItemData(Combo1.ListIndex) >= 0 Then
                If vSql <> "" Then vSql = vSql & " AND "
                vSql = vSql & " formapago.tipforpa = " & Combo1.ItemData(Combo1.ListIndex)    'SubTipo
            
            End If
        End If
    End If

    If vSql <> "" Then vSql = vSql & " AND "
    vSql = vSql & " ((formapago.tipforpa in (" & vbTalon & "," & vbPagare & ") and pagos.nrodocum is null) or not formapago.tipforpa in (" & vbTalon & "," & vbPagare & "))"

    
    ' si está marcado los de solo el banco previsto
    If chkSoloBancoPrev.Value = 1 Then
        If txtCta(4).Text <> "" Then
            If vSql <> "" Then vSql = vSql & " AND "
            vSql = vSql & " pagos.ctabanc1 = " & DBSet(txtCta(4).Text, "T")
        End If
    End If
    
    
    
    ' no entran a jugar los recibos
    If vSql <> "" Then vSql = vSql & " and "
    vSql = vSql & " formapago.tipforpa <> " & vbTransferencia
    
    ' solo los pendientes de cobro
    If vSql <> "" Then vSql = vSql & " and "
    vSql = vSql & " (coalesce(impefect,0) - coalesce(imppagad,0)) <> 0 "
    
    
    ' pagos
    Cad = "SELECT pagos.*, tipofpago.siglas, "
    Cad = Cad & " coalesce(impefect,0) - coalesce(imppagad,0) imppdte "
    Cad = Cad & " FROM pagos , formapago, tipofpago"
    Cad = Cad & " Where  formapago.tipforpa = tipofpago.tipoformapago"
    Cad = Cad & " AND pagos.codforpa = formapago.codforpa"
    If vSql <> "" Then Cad = Cad & " AND " & vSql
    
    'SQL pedido
    DevSQL = Cad
End Function


Private Sub CargaPagos()

    Cad = DevSQL
    
    'ORDENACION
    If CampoOrden = "" Then CampoOrden = "pagos.fecefect"
    
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    If CampoOrden <> "pagos.fecefect" Then Cad = Cad & ", pagos.fecefect"
    

    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        InsertaItemPago
        Rs.MoveNext
    Wend
    Rs.Close

End Sub


Private Sub InsertaItemPago()
Dim J As Byte


     
     Set ItmX = ListView1.ListItems.Add()
     
     
     ItmX.Text = Rs!NUmSerie
     ItmX.SubItems(1) = Rs!numfactu
     ItmX.SubItems(2) = Format(Rs!FecFactu, "dd/mm/yyyy")
     ItmX.SubItems(3) = Format(Rs!fecefect, "dd/mm/yyyy")
     ItmX.SubItems(4) = Rs!numorden
     ItmX.SubItems(5) = DBLet(Rs!nomprove, "T")
     ItmX.SubItems(6) = DBLet(Rs!siglas, "T")
     ItmX.SubItems(7) = Format(Rs!ImpEfect, FormatoImporte)
     If Not IsNull(Rs!imppagad) Then
         ItmX.SubItems(8) = Format(Rs!imppagad, FormatoImporte)
         impo = Rs!ImpEfect - Rs!imppagad
         ItmX.SubItems(9) = Format(impo, FormatoImporte)
     Else
         impo = Rs!ImpEfect
         ItmX.SubItems(8) = "0.00"
         ItmX.SubItems(9) = ItmX.SubItems(7)
     End If
     If Rs!fecefect < Fecha Then
         'LO DEBE
         ItmX.SmallIcon = 1
         Vencido = Vencido + impo
     Else
'            ItmX.SmallIcon = 2
     End If
     If Combo1.ListIndex <> -1 Then
        If Combo1.ItemData(Combo1.ListIndex) = 1 Then
            If Not IsNull(Rs!nrodocum) Then
                ItmX.Checked = True
                ImpSeleccionado = ImpSeleccionado + impo
            End If
        End If
     End If
     ItmX.SubItems(10) = DBLet(Rs!Referencia, "T")
     
     
     'El tag lo utilizo para la cta proveedor
     ItmX.Tag = Rs!codmacta
     
     Importe = Importe + impo
     
     'Si el documento estaba emitido ya
     If Val(Rs!emitdocum) = 1 Then
         'Tiene marcado DOCUMENTO EMITIDO
         ItmX.ForeColor = vbRed
         For J = 1 To ListView1.ColumnHeaders.Count - 1
             ItmX.ListSubItems(J).ForeColor = vbRed
         Next J
        ' If DBLet(Rs!Referencia, "T") = "" Then ItmX.ListSubItems(4).ForeColor = vbMagenta
     End If
     
     ' nuevo si está marcada lo miramos
     
     For i = 1 To FrasMarcadas.Count
        Cad = FrasMarcadas.Item(i)
        If RecuperaValor(Cad, 1) = Rs!NUmSerie And CStr(RecuperaValor(Cad, 2)) = CStr(Rs!numfactu) And CStr(RecuperaValor(Cad, 3)) = CStr(Rs!FecFactu) And RecuperaValor(Cad, 4) = Rs!numorden And RecuperaValor(Cad, 5) = Rs!codmacta Then
            ItmX.Checked = True
    
            ImpSeleccionado = ImpSeleccionado + impo
            Exit For

        End If
     Next i
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Para dejar las variables bien
    ContabTransfer2 = False
    DesdeRecepcionTalones = False
    'Por si acaso
    NumeroTalonPagere = ""
    CodmactaUnica = ""
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCta(4).Text = RecuperaValor(CadenaSeleccion, 1)
        txtDCta(4).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCta(5) = RecuperaValor(CadenaSeleccion, 1)
        txtDCta(5).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub imgCheck_Click(Index As Integer)
    SeleccionarTodos Index = 1 Or Index = 2
End Sub

Private Sub imgCuentas_Click(Index As Integer)
Dim DevfrmCCtas As String
Dim Cad As String

    Select Case Index
        Case 0 ' cuenta de banco
            Set frmBan = New frmBasico2
            AyudaBanco frmBan
            Set frmBan = Nothing
            txtCta_LostFocus 4
            PonFoco txtCta(4)
       Case 1 ' cuenta de proveedor
            Set frmCCtas = New frmColCtas
            DevfrmCCtas = ""
            Cad = txtCta(5).Text
            frmCCtas.DatosADevolverBusqueda = "0"
            frmCCtas.FILTRO = "2" ' proveedores
            frmCCtas.Show vbModal
            Set frmCCtas = Nothing
            If Cad <> txtCta(5).Text Then txtCta_LostFocus 5
            PonFoco txtCta(Index + 4)
    End Select
    
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Fecha = Now
    If Text3(i).Text <> "" Then
        If IsDate(Text3(i).Text) Then Fecha = CDate(Text3(i).Text)
    End If
    Cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Fecha
    frmC.Show vbModal
    Set frmC = Nothing
    If Cad <> "" Then
        Text3(Index).Text = Cad
            
        If Index = 0 Then
            'Antes de poder cambiar la fecha hay que comprobar si la fecha devuelta es OK
            '                                                'Fecha OK
            If FechaCorrecta2(CDate(Cad), True) < 2 Then
                Text3(0).Text = Cad
                Fecha = CDate(Cad)
            End If
        End If
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
'        Columnas = "Serie|Factura|F. Fact|F. VTO|Nº|PROVEEDOR|Tipo|Importe|Pagado|Pendiente|"
    
    Select Case ColumnHeader
        Case "Serie"
            CampoOrden = "pagos.numserie"
        Case "Factura"
            CampoOrden = "pagos.numfactu"
        Case "F. Fact"
            CampoOrden = "pagos.fecfactu"
        Case "F. VTO"
            CampoOrden = "pagos.fecefect"
        Case "Nº"
            CampoOrden = "pagos.numorden"
        Case "PROVEEDOR"
            CampoOrden = "nomprove"
        Case "Tipo"
            CampoOrden = "siglas"
        Case "Importe"
            CampoOrden = "pagos.impefect"
        Case "Pagado"
            CampoOrden = "pagos.imppagad"
        Case "Pendiente"
            CampoOrden = "imppdte"
    End Select
    CargaList

End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    i = ColD(0)
    impo = ImporteFormateado(Item.SubItems(i))
    
    If Item.Checked Then
        Set ListView1.SelectedItem = Item
        i = 1
    Else
        i = -1
    End If
    ImpSeleccionado = ImpSeleccionado + (i * impo)
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Me.mnContextual
    End If
End Sub

Private Sub SeleccionarTodos(Seleccionar As Boolean)
Dim J As Integer
    J = ColD(0)
    ImpSeleccionado = 0
    For i = 1 To Me.ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = Seleccionar
        impo = ImporteFormateado(ListView1.ListItems(i).SubItems(J))
        ImpSeleccionado = ImpSeleccionado + impo
    Next i
    If Not Seleccionar Then ImpSeleccionado = 0
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
End Sub


Private Sub mnNumero_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
        
End Sub

Private Sub mnQUitarSel_Click()
    SeleccionarTodos False
End Sub

Private Sub mnSelectAll_Click()
    SeleccionarTodos True
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub





Private Function InsertarPagosEnTemporal2() As Boolean
Dim C As String
Dim Aux As String
Dim J As Long
Dim FechaContab As Date
Dim FechaFinEjercicios As Date
Dim vGasto As Currency

    
    InsertarPagosEnTemporal2 = False
    
    C = " WHERE codusu =" & vUsu.Codigo
    Conn.Execute "DELETE FROM tmpfaclin" & C
    riesgo = 0

    'Fechas fin ejercicios
    FechaFinEjercicios = DateAdd("yyyy", 1, vParam.fechafin)


    '++
    GastosTransferencia = CCur(ComprobarCero(Text3(6).Text))



    Aux = "INSERT INTO tmpfaclin (codusu, codigo, Fecha,Numfactura, cta, Cliente, NIF, Imponible,  Total) "
    Aux = Aux & "VALUES (" & vUsu.Codigo & ","
    For J = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(J).Checked Then
            C = J & ",'"

                'FechaContab = CDate(ListView1.ListItems(J).SubItems(3))
                FechaContab = CDate(ListView1.ListItems(J).SubItems(2))

            

            C = C & Format(FechaContab, FormatoFecha) & "','"
            
            '-----------------------------------------------------
            'Fecha de contabilizacion
            If Me.chkPorFechaVenci.Value Then
                
                i = 0
                
                'Meto la fecha VTO
                If FechaContab < vParam.fechaini Then
                    i = 1
                Else
                    If FechaContab > FechaFinEjercicios Then
                        i = 1
                    Else
                        If FechaContab < vParamT.fechaAmbito Then i = 1
                    End If
                End If
                
                
                
                If i = 1 Then FechaContab = CDate(Text3(0).Text)
                
                
            Else
                'La fecha de contabilizacion es la del text
                FechaContab = CDate(Text3(0).Text)
            End If
            'MEto la fecha de contabilizaccion
            C = C & Format(FechaContab, FormatoFecha) & "','"
            'Cuenta contable
            C = C & ListView1.ListItems(J).Tag & "','"
            C = C & DevNombreSQL(ListView1.ListItems(J).Text) & "|" & ListView1.ListItems(J).SubItems(1) & "|" & ListView1.ListItems(J).SubItems(2) & "|" & ListView1.ListItems(J).SubItems(4) & "|" & ListView1.ListItems(J).Tag
            C = C & "|','"
            
            'Cuenta agrupacion cobros
            If Combo1.ItemData(Combo1.ListIndex) = 1 And ContabTransfer2 Then
                C = C & Me.chkGenerico(1).Tag & "',"
            Else
                C = C & Me.chkGenerico(0).Tag & "',"
            End If
            'Dinerito
            'riesgo es GASTO
            i = ColD(0)
            impo = ImporteFormateado(ListView1.ListItems(J).SubItems(i))

            impo = impo - riesgo
            C = C & TransformaComasPuntos(CStr(impo)) & "," & TransformaComasPuntos(CStr(riesgo)) & ")"


            'Lo meto en la BD
            C = Aux & C
            Conn.Execute C
        End If
    Next J

'    If Combo1.ItemData(Combo1.ListIndex) = 1 And GastosTransferencia <> 0 Then
    If GastosTransferencia <> 0 Then
            'aqui ira los gastos asociados a la transferencia
            'Hay que ver los lados
            
           
            Cad = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", Text3(1).Tag, "T")
            If Cad = "" Then Err.Raise 513, , " Cuenta gastos sin configurar"
            
            FechaContab = CDate(Text3(0).Text)
            C = "'" & Format(FechaContab, FormatoFecha) & "'"
            C = C & "," & C
            C = J & "," & C & ",'" & Cad & "','"
            'Serie factura |FECHAfactura| ----> pondre: "gastos" | fecha contab
            C = C & "TRA" & Format(SegundoParametro, "0000000") & "|" & FechaContab & "|','" & Cad & "',"
            'Dinerito
            'riesgo es GASTO
            impo = GastosTransferencia
            C = C & TransformaComasPuntos(CStr(impo)) & ",0)"
            C = Aux & C
            Conn.Execute C
        
    End If
    
    InsertarPagosEnTemporal2 = True
    
    

End Function

'TENGO en la tabla tmpfaclin los vtos.
'Ahora en funcion de los check haremos la contabilizacion
'agrupando de un modo o de otro
Private Sub ContablizaDesdeTmp()
Dim SQL As String
Dim ContraPartidaPorLinea As Boolean
Dim UnAsientoPorCuenta As Boolean
Dim PonerCuentaGenerica As Boolean
Dim AgrupaCuenta As Boolean
Dim Rs As ADODB.Recordset
Dim MiCon As Contadores
Dim CampoCuenta As String
Dim CampoFecha As String
Dim GeneraAsiento As Boolean
Dim CierraAsiento As Boolean
Dim NumLinea As Integer
Dim ImpBanco As Currency
Dim NumVtos As Integer
Dim GastosTransDescontados As Boolean
Dim LineaUltima As Integer

    'Valores por defecto
    ContraPartidaPorLinea = True
    UnAsientoPorCuenta = False
    PonerCuentaGenerica = False
    AgrupaCuenta = False
    CampoFecha = "numfactura"
    GastosTransDescontados = False 'por lo que pueda pasar
    
    '++
    GastosTransferencia = Val(Text3(6).Text)
    
    'Si va agrupado por cta
    AgrupaCuenta = False
    If Combo1.ItemData(Combo1.ListIndex) = 1 And ContabTransfer2 Then
        If Me.chkContrapar(1).Value Then ContraPartidaPorLinea = False
        If Me.chkAsiento(1).Value Then UnAsientoPorCuenta = True
        If chkGenerico(1).Value Then PonerCuentaGenerica = True
        
        'Si lleva GastosTransferencia entonce AGRUPAMOS banco
        If GastosTransferencia <> 0 Then
            
            'gastos tramtiaacion transferenca descontados importe
            SQL = DevuelveDesdeBD("GastTransDescontad", "bancos", "codmacta", Text3(1).Tag, "T")
            GastosTransDescontados = SQL = "1"
            
            AgrupaCuenta = False
        Else
'            If Me.chkVtoCuenta(0).Value Then AgrupaCuenta = True
        End If
    Else
        'Si no es transferencia
        If Me.chkContrapar(0).Value Then ContraPartidaPorLinea = False
        If Me.chkAsiento(0).Value Then UnAsientoPorCuenta = True
        If chkGenerico(0).Value Then PonerCuentaGenerica = True
'        If Me.chkVtoCuenta(0).Value Then AgrupaCuenta = True
        'La contabiliacion es por fecha vencimiento , no por fecha solicitada
        'YA cuando inserto en temporal miro esto
        'If chkPorFechaVenci.Value Then CampoFecha = "fecha"
        
        '++
        'Si lleva GastosTransferencia entonce AGRUPAMOS banco
        If GastosTransferencia <> 0 Then
            
            'gastos tramtiaacion transferenca descontados importe
            SQL = DevuelveDesdeBD("GastTransDescontad", "bancos", "codmacta", Text3(1).Tag, "T")
            GastosTransDescontados = SQL = "1"
            
            AgrupaCuenta = False
        Else
'            If Me.chkVtoCuenta(0).Value Then AgrupaCuenta = True
        End If
        
        '++
        
    End If
    
    If PonerCuentaGenerica Then
        CampoCuenta = "NIF"
    Else
        CampoCuenta = "cta"
    End If
    'EL SQL lo empezamos aquin
    SQL = CampoCuenta & " AS cliprov,"
    'Selecciona
    SQL = "select count(*) as numvtos,codigo,numfactura,fecha,cliente," & SQL & "sum(imponible) as importe,sum(total) as gastos from tmpfaclin"
    SQL = SQL & " where codusu =" & vUsu.Codigo & " GROUP BY "
    Cad = ""
    If AgrupaCuenta Then
       If PonerCuentaGenerica Then
            Cad = "nif" 'La columna NIF lleva los datos de la cuenta generica
        Else
            Cad = "cta"
        End If
        'Como estamos agrupando por cuenta, marcaremos tb la fecha
        'Ya que si tienen fechas distintas son apuntes distintos
        Cad = Cad & "," & CampoFecha
    End If
    
    'Si no agrupo por nada agrupare por codigo(es decir como si no agrupara)
    If Cad = "" Then Cad = "codigo"
    
    'La ordenacion
    Cad = Cad & " ORDER BY " & CampoFecha
    If Not PonerCuentaGenerica Then Cad = Cad & ",cta"
        
    
    'Tanto si agrupamos por cuenta (Generica o no)
    'el recodset tendra las lineas que habra que insertar en/los apuntes(s)
    '
    'Es decir. Que si agrupo no tengo que ir moviendome por el recodset mirando a ver si
    'las cuentas son iguales.
    'Ya que al hacer group by ya lo estaran
    Cad = SQL & Cad
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Inicializamos variables
    Fecha = CDate("01/01/1900")
    GeneraAsiento = False
    While Not Rs.EOF
        'Comprobaciones iniciales
        If UnAsientoPorCuenta Then
            'Para cada linea ira su asiento
            GeneraAsiento = True
            CierraAsiento = True
            If Fecha < CDate("01/01/1950") Then CierraAsiento = False
            Fecha = CDate(Rs.Fields(CampoFecha))
        Else
            'Veremos en funcion de la fecha
            GeneraAsiento = False
            If CDate(Rs.Fields(CampoFecha)) = Fecha Then
                'Estamos en la misma fecha. Luego sera el mismo asiento
                'Excepto que asi no lo digan las variables
                If Not PonerCuentaGenerica Then
                    If UnAsientoPorCuenta Then
                        GeneraAsiento = True
                        If Fecha < CDate("01/01/1950") Then CierraAsiento = True
                    End If
                End If
                        
            Else
                'Fechas distintas.
                GeneraAsiento = True
                CierraAsiento = True
                If Fecha < CDate("01/01/1950") Then CierraAsiento = False
        
                Fecha = CDate(Rs.Fields(CampoFecha))
            End If
        End If 'de aseinto por cuenta
        
        
        
        
        
        'Si tengo que cerrar el asiento anterior
        If CierraAsiento Then
            'Tirar atras el RS
            If Not ContraPartidaPorLinea Then
                Rs.MovePrevious
                Fecha = CDate(Rs.Fields(CampoFecha))  'Para la fecha de asiento
                impo = ImpBanco
                'Generamos las lineas de apunte que faltan
                InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
                
                'Inserto para que actalice             3: Opcion para INSERT INTO tmpactualizar
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
                
                'Reestauramos variables
                NumVtos = 0
                'Ponemos la variable
                CierraAsiento = False
                'Volvemos el RS al sitio
                Rs.MoveNext
                Fecha = CDate(Rs.Fields(CampoFecha))
            Else
                'Inserto para que actalice             3: Opcion para INSERT INTO tmpactualizar
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
            End If
        End If
 
        
        'Si genero asiento
        If GeneraAsiento Then
            If MiCon Is Nothing Then Set MiCon = New Contadores
            MiCon.ConseguirContador "0", Fecha <= vParam.fechafin, True
                        
            'Genero la cabecera
            InsertarEnAsientosDesdeTemp Rs, MiCon, 0, NumLinea, NumVtos
            
            NumLinea = 1
            ImpBanco = 0
            'Reservo la primera linea para el banco
            If GastosTransferencia <> 0 Then
                NumLinea = 2
                If Not GastosTransDescontados Then
                        ImpBanco = -GastosTransferencia
                End If
            End If
            
            riesgo = 0
        End If
    
    
        
    
        'Para el cobro /pago  que tendremos en la fila actual del recordset
        '++
        Dim VienedeGastos As Boolean
        Dim CadRes As String
        CadRes = RecuperaValor(DBLet(Rs!Cliente, "T"), 1)
        VienedeGastos = (CadRes = "GASTOS" Or CadRes = "TRA")
        
        impo = Rs!Importe
        InsertarEnAsientosDesdeTemp Rs, MiCon, 1, NumLinea, Rs!NumVtos, VienedeGastos
    
        'If Cobros Then
        '    riesgo = riesgo + Rs!Gastos
        'Else
            riesgo = 0
        'End If
        ImpBanco = ImpBanco + Rs!Importe
        NumLinea = NumLinea + 1
        
        'Si tengo que generar la contrapartida
        If ContraPartidaPorLinea Then
            NumVtos = Rs!NumVtos
            InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
            NumLinea = NumLinea + 1
            ImpBanco = 0
            riesgo = 0
        Else
            NumVtos = NumVtos + Rs!NumVtos
        End If
        
        'Nos movemos
        Rs.MoveNext
        
        
        If Rs.EOF Then
            
            If Not ContraPartidaPorLinea Then
                
                'Era la ultima linea.
                Rs.MovePrevious
                
                LineaUltima = NumLinea
                
                'Cierro el apunte, del banco
                'Si fuera una transferenicia con gastos descontados, me he dejado el numlinea=1
                'si no, no hago nada
                If GastosTransferencia <> 0 Then
                    If Not GastosTransDescontados Then NumLinea = 1
                End If
                impo = ImpBanco
                InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
    
                If GastosTransferencia <> 0 Then
                    If Not GastosTransDescontados Then
                        NumLinea = LineaUltima + 1
                
                        impo = GastosTransferencia
                        
                        InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
                    End If
                End If
    
    
                'CIERRO EL APUNTE
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
                
                'Y vuelvo a ponerlo ande tocaba. Para que se salga del bucle
                Rs.MoveNext
                
            Else
                'Cada linea de asiento tiene su banco
                'Faltara insertarlo en tmpactualizar
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
            End If
        End If
    Wend
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Function ColD(Colu As Integer) As Integer
    Select Case Colu
    Case 0
            'IMporte pendiente
            ColD = 10
    Case 1
    
    End Select
    'If Not Cobros Then ColD = ColD - 1
    ColD = ColD - 1
End Function




Private Sub Text3_GotFocus(Index As Integer)
    ConseguirFoco Text3(Index), 3
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFec KeyAscii, 0
            Case 1: KEYFec KeyAscii, 1
            Case 2: KEYFec KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFec(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text = "" Then Exit Sub

    Select Case Index
        Case 6
            PonerFormatoDecimal Text3(Index), 3
            Me.ImporteGastosTarjeta_ = ComprobarCero(Text3(Index).Text)
            
        Case 0, 1, 2
        
            If Not EsFechaOK(Text3(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                Text3(Index).Text = ""
                PonFoco Text3(Index)
            
            End If
            If Index = 0 Then Fecha = IIf(Text3(Index).Text = "", Now, CDate(Text3(Index).Text))
    
            If Index = 1 Or Index = 2 Then CargaList
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
    If Index = 5 Then CtaAnt = txtCta(Index)
    If Index = 4 Then CtaAntBan = txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 4: KEYCuentas KeyAscii, 0
            Case 5: KEYCuentas KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYCuentas(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCuentas_Click (Indice)
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim DevfrmCCtas As String
Dim SQL As String

    Select Case Index
        Case 4 ' cuenta de banco
            txtCta(Index).Text = Trim(txtCta(Index).Text)
            DevfrmCCtas = txtCta(Index).Text
            i = 0
            If DevfrmCCtas <> "" Then
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                    DevfrmCCtas = DevuelveDesdeBD("codmacta", "bancos", "codmacta", DevfrmCCtas, "T")
                    If DevfrmCCtas = "" Then
                        SQL = ""
                        MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                    End If
                Else
                    MsgBox SQL, vbExclamation
                    DevfrmCCtas = ""
                    SQL = ""
                End If
                i = 1
            Else
                SQL = ""
            End If
            
            
            txtCta(Index).Text = DevfrmCCtas
            txtDCta(Index).Text = SQL
            If DevfrmCCtas = "" And i = 1 Then
                PonFoco txtCta(Index)
            Else
                Text3(1).Tag = txtCta(Index).Text
            End If
        
        
            If CtaAntBan <> txtCta(4).Text Then CargaList
        
        
        
        Case 5 ' cuenta cliente
            DevfrmCCtas = Trim(txtCta(Index).Text)
            i = 0
            If DevfrmCCtas <> "" Then
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                    
                Else
                    MsgBox SQL, vbExclamation
                    If Index < 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
                        DevfrmCCtas = ""
                        SQL = ""
                    End If
                End If
                i = 1
            Else
                SQL = ""
            End If
            
            txtCta(Index).Text = DevfrmCCtas
            txtDCta(Index).Text = SQL
            If DevfrmCCtas = "" And i = 1 Then
                PonFoco txtCta(Index)
            Else
                Me.CodmactaUnica = txtCta(Index)
            End If
        
            If CtaAnt <> txtCta(5).Text Then CargaList

    End Select
End Sub


Private Sub LeerparametrosContabilizacion()
Dim B As Boolean

    Me.chkAsiento(0).Value = Abs(vParamT.contapag)
    
   
    If vParamT.contapag Then
        B = False
    Else
        B = vParamT.AgrupaBancario
    End If
    Me.chkContrapar(0).Value = Abs(B)
End Sub

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'3 Opciones
'   0.- CABECERA
'   1.- LINEAS  de clientes o proeveedores
'   2.- Cierre del asiento con el BANCO, y o caja
'   3.- Para poner Boqueoactu a 0
'       Si vParam.contapag entonces cuando hago el de banco/caja lo updateo
'       pero si NO es uno por pagina entonces, la utlima vez k hago el apunte por banco/caja
'       ejecuto el update
'
'   La contrB sera la contrpartida , si lueog resulta k si, para la linea de banco o caja
'
'   FechaAsiento:  Antes estaba a "piñon" text3(0).text
'
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'ByRef m As Contadores, NumLine As Integer, Marcador As Integer, Cabecera As Byte, ByRef ContraB As String, ByRef LaUltimaAmpliacion As String, ContraParEnBanco As Boolean, CuentaDeCobroGenerica As Boolean, CodigoCtaCoborGenerica As String)
Private Function InsertarEnAsientosDesdeTemp(ByRef RS1 As ADODB.Recordset, ByRef m As Contadores, Cabecera As Byte, ByRef NumLine As Integer, NumVtos As Integer, Optional VienedeGastos As Boolean)
Dim SQL As String
Dim Ampliacion As String
Dim Debe As Boolean
Dim Conce As Integer
Dim TipoAmpliacion As Integer
Dim PonerContrPartida As Boolean
Dim Aux As String
Dim ImporteInterno As Currency

Dim ImporteInterno2 As Currency
    
    
    ImporteInterno = impo
    ImporteInterno2 = impo
    
    'LaUltimaAmpliacion  --> Servira pq si en parametros esta marcado un apunte por movimiento, o solo metemos
    '                        un unico pagao/cobro, repetiremos numdocum, textoampliacion
    
    'El diario

    FechaAsiento = Fecha

    Ampliacion = vp.diaripro

    
    If Cabecera = 0 Then
        'La cabecera
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
        SQL = SQL & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador
        SQL = SQL & ",  '"
        SQL = SQL & "Generado desde Tesorería el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre
        If Combo1.ItemData(Combo1.ListIndex) = 1 Then
            'TRANSFERENCIA
            Ampliacion = DevuelveDesdeBD("descripcion", "stransfer", "codigo", SegundoParametro, "N")
            If Ampliacion <> "" Then
                Ampliacion = "Concepto: " & Ampliacion
                Ampliacion = DevNombreSQL(Ampliacion)
                Ampliacion = vbCrLf & Ampliacion
                SQL = SQL & Ampliacion
            End If
        End If
        
        SQL = SQL & "',"
        SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Pagos'"

        
        SQL = SQL & ")"
        NumLine = 0
     
    Else
        If Cabecera < 3 Then
            'Lineas de apuntes o cabecera.
            'Comparten el principio
             SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
             SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfacpr, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
             SQL = SQL & "VALUES (" & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador & "," & NumLine & ",'"
             
             '1:  Asiento para el VTO
             If Cabecera = 1 Then
                 'codmacta
                 'Si agrupa la cuenta entonces
                 SQL = SQL & RS1!cliprov & "','"
                 
                 
                 'numdocum: la factura
                 If NumVtos > 1 Then
                    Ampliacion = "Vtos: " & NumVtos
                 Else
                 
                    If vParam.CodiNume = 1 Then
                        'Numero de registro SI existe
                        Aux = "  numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "T")
                        Aux = Aux & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F")
                        Aux = Aux & " and codmacta = " & DBSet(RecuperaValor(RS1!Cliente, 5), "T") & " AND 1"
                        Aux = DevuelveDesdeBD("numregis", "factpro", Aux, "1")
                        If Aux = "" Then
                            Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 2))
                        Else
                            Ampliacion = Right("0000000000" & Aux, 10)
                        End If
                        Aux = ""
                            
                    Else
                        'NUmero factura
                        Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 2))
                    End If
                    
                    
                    
                 End If
                 SQL = SQL & Ampliacion & "',"
                
               
                    'PROVEEDORES
                 If ImporteInterno < 0 Then
                    If vParam.abononeg Then
                        Debe = True
                    Else
                        'Va al debe pero cambiado de signo
                        Debe = False
                        ImporteInterno = Abs(ImporteInterno)
                    End If
                 Else
                    Debe = True
                 End If
                 If Debe Then
                     Conce = vp.condepro
                     TipoAmpliacion = vp.ampdepro
                     PonerContrPartida = vp.ctrdepro = 1
                 Else
                     Conce = vp.conhapro
                     TipoAmpliacion = vp.amphapro
                     PonerContrPartida = vp.ctrhapro = 1
                                          
                 End If
                  
            
                
                
                 SQL = SQL & Conce & ","
                 
                 'AMPLIACION
                 Ampliacion = ""
                


                Select Case TipoAmpliacion
                Case 0, 1
                   If TipoAmpliacion = 1 Then Ampliacion = Ampliacion & vp.siglas & " "
                   Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 2)
                
                Case 2
                
                   Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 3)
                
                Case 3
                    'NUEVA AMPLIC
                    Ampliacion = DescripcionTransferencia
                Case 4
                    Ampliacion = DevuelveDesdeBD("descripcion", "bancos", "codmacta", Me.txtCta(4).Text, "T")
                    If Ampliacion = "" Then Ampliacion = Me.txtDCta(4).Text
                   
                Case 5
                       
                        NumeroTalonPagere = "  numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T")
                        NumeroTalonPagere = NumeroTalonPagere & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "T")
                        NumeroTalonPagere = NumeroTalonPagere & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F")
                        NumeroTalonPagere = NumeroTalonPagere & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
                        NumeroTalonPagere = NumeroTalonPagere & " and codmacta = " & DBSet(RS1!cliprov, "T") & " AND 1"
                        NumeroTalonPagere = DevuelveDesdeBD("referencia", "pagos", NumeroTalonPagere, "1")
                        Ampliacion = ""
                        If NumeroTalonPagere <> "" Then
                            Ampliacion = DevuelveDesdeBD("descripcion", "bancos", "codmacta", Me.txtCta(4).Text, "T")
                            If Ampliacion = "" Then Ampliacion = Me.txtDCta(4).Text
                        
                          
                            Ampliacion = NumeroTalonPagere & " " & Ampliacion
                        End If
                        
                        If Ampliacion = "" Then
                            Ampliacion = RecuperaValor(RS1!Cliente, 2)
                        Else
                            Ampliacion = "NºDoc: " & Ampliacion
                        End If
                    
                Case 6
                          ' Como quiere ver toda la ampliacion en la ventana, vamos a suponore un maximo de 23 carcateres
                          
                        MiVariableAuxiliar = RecuperaValor(RS1!Cliente, 2)
                        If Len(MiVariableAuxiliar) > 23 Then MiVariableAuxiliar = Mid(MiVariableAuxiliar, 1, 20)
                        Aux = RecuperaValor(RS1!Cliente, 5)
                        Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
                        If Aux = "" Then Aux = "***   " & RecuperaValor(RS1!Cliente, 5) & " ****"
                       Ampliacion = Mid(Aux, 1, 39 - Len(MiVariableAuxiliar))
                        Ampliacion = Ampliacion & " " & MiVariableAuxiliar
                       
                End Select
                   
                If NumVtos > 1 Then
                    'TIENE MAS DE UN VTO. No puedo ponerlo en la ampliacion
                    Ampliacion = "Vtos: " & NumVtos
                End If
                
                 'Le concatenamos el texto del concepto para el asiento -ampliacion
                 Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce)) & " "
                 'Para la ampliacion de nºtal + ctrapar NO pongo la ampliacion del concepto
                 If TipoAmpliacion = 5 Then Aux = ""
                 If TipoAmpliacion = 6 Then Aux = ""
                 Ampliacion = Aux & Ampliacion
                 If Len(Ampliacion) > 50 Then Ampliacion = Mid(Ampliacion, 1, 50)
                
                 SQL = SQL & "'" & DevNombreSQL(Ampliacion) & "',"
                 
                 
                 If Debe Then
                    SQL = SQL & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                 Else
                    SQL = SQL & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                 End If
             
                'CENTRO DE COSTE
                SQL = SQL & "NULL,"
                
                'SI pone contrapardida
                If PonerContrPartida Then
                   SQL = SQL & "'" & Text3(1).Tag & "',"
                Else
                   SQL = SQL & "NULL,"
                End If
            
             
            Else
                    '----------------------------------------------------
                    'Cierre del asiento con el total contra banco o caja
                    '----------------------------------------------------
                    'codmacta
                    SQL = SQL & Text3(1).Tag & "','"
                     
  
                    PonerContrPartida = False
                    If NumVtos = 1 Then
                        PonerContrPartida = True
                    Else
                        PonerContrPartida = False
                        If Me.txtCta(5).Text <> "" Then PonerContrPartida = True
                        
                    End If
                       
                    If PonerContrPartida Then
                       Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 2))
                    Else
                       
                       Ampliacion = "Vtos: " & NumVtos
                    End If
                     
                    SQL = SQL & Ampliacion & "',"
                   
                    
                
                    'PROVEEDORES
                    If ImporteInterno < 0 Then
                       If vParam.abononeg Then
                           Debe = False
                       Else
                           'Va al debe pero cambiado de signo
                           Debe = True
                           ImporteInterno = Abs(ImporteInterno)
                       End If
                    Else
                       Debe = False
                    End If
                    
                    If Debe Then
                        Conce = vp.condepro
                        TipoAmpliacion = vp.ampdepro
                    Else
                        Conce = vp.conhapro
                        TipoAmpliacion = vp.amphapro
                    End If

                     
                     SQL = SQL & Conce & ","
                     'AMPLIACION
                     'AMPLIACION
                     Ampliacion = ""
                     
                     
                     
                     'Si estoy contabilizando pag de UN unico proveedor entonces NumeroTalonPageretendra valor
                     If NumVtos > 1 And NumeroTalonPagere <> "" Then NumVtos = 1
                        
                     
                     If NumVtos = 1 Then
                    
                        Select Case TipoAmpliacion
                        Case 0, 1
                           If TipoAmpliacion = 1 Then Ampliacion = Ampliacion & vp.siglas & " "
                           Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 2)
                        
                        Case 2
                        
                           Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 3)
                        
                        Case 3
                            'NUEVA AMPLIC
                             Ampliacion = DescripcionTransferencia
                        Case 4, 5
                            'Nombre ctrpartida
                            Ampliacion = CStr(DBLet(RS1!cliprov, "T"))
                            Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Ampliacion, "T")
                            DescripcionTransferencia = Ampliacion
                            
                                
                                Ampliacion = NumeroTalonPagere
                                If Ampliacion = "" Then
                                    Ampliacion = RecuperaValor(RS1!Cliente, 2)
                                Else
                                    Ampliacion = "NºDoc: " & Ampliacion
                                End If
                            
                          
                            Ampliacion = Ampliacion & " " & DescripcionTransferencia
                            DescripcionTransferencia = ""
                          
                        Case 6
                            Ampliacion = CStr(DBLet(RS1!cliprov, "T"))
                            Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Ampliacion, "T")
                            
                            
                            
                          
                            MiVariableAuxiliar = RecuperaValor(RS1!Cliente, 2)
                            If Len(MiVariableAuxiliar) > 23 Then MiVariableAuxiliar = Mid(MiVariableAuxiliar, 1, 20)
                            Aux = RecuperaValor(RS1!Cliente, 5)
                            Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
                            If Aux = "" Then Aux = "***   " & RecuperaValor(RS1!Cliente, 5) & " ****"
                            Ampliacion = Mid(Aux, 1, 39 - Len(MiVariableAuxiliar))
                            Ampliacion = Ampliacion & " " & MiVariableAuxiliar
                                
                            
                            
                            DescripcionTransferencia = ""
                        End Select
                    Else
                        'Ma de un VTO.  Si no
                        If vp.tipoformapago = vbTransferencia Then
                            'SI es transferencia
                            'If TipoAmpliacion = 3 Then Ampliacion = DescripcionTransferencia
                            Ampliacion = DescripcionTransferencia
                        
                        Else
                            
                            Ampliacion = ObservacionApunteBanco
                        End If
                    End If
                    
                     Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce))
                     Aux = Aux & " "
                     'Para la ampliacion de nºtal + ctrapar NO pongo la ampliacion del concepto
                     If TipoAmpliacion = 5 Then Aux = ""
                     Ampliacion = Trim(Aux & Ampliacion)
                     If Len(Ampliacion) > 50 Then Ampliacion = Mid(Ampliacion, 1, 50)
                    
                     SQL = SQL & "'" & DevNombreSQL(Ampliacion) & "',"
        
                         
                     If Debe Then
                        SQL = SQL & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                     Else
                        SQL = SQL & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                     End If
                 
                     'CENTRO DE COSTE
                     SQL = SQL & "NULL,"
                    
                     'SI pone contrapardida
                     If PonerContrPartida Then
                        SQL = SQL & "'" & RS1!cliprov & "',"
                     Else
                        SQL = SQL & "NULL,"
                     End If
                
                        
                 
            End If
            
            'Trozo comun
            '------------------------
            'IdContab
            SQL = SQL & "'PAGOS',"
            
            'Punteado
            SQL = SQL & "0,"
            
            If Cabecera = 1 And Not VienedeGastos Then
            
                ' nuevos campos de la factura
                'numSerie , numfacpr, FecFactu, numorden, TipForpa, reftalonpag, bancotalonpag
                If NumVtos > 1 Then
                    SQL = SQL & "null,null,null,null,"
                Else
                    SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 2), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & ","
                    SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & ","
                End If
                
                SQL = SQL & DBSet(Combo1.ItemData(Combo1.ListIndex), "N") & ","
                
                    
                    
                Dim SqlBanco As String
                Dim RsBanco As ADODB.Recordset
                
                'ANTES
                'SqlBanco = "select reftalonpag, bancotalonpag from tmppagos2 where codusu = " & vUsu.Codigo
                
                
                SqlBanco = "select Referencia , ctabanc1 from pagos where "
                SqlBanco = SqlBanco & " numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T")
                SqlBanco = SqlBanco & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "T")
                SqlBanco = SqlBanco & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F")
                SqlBanco = SqlBanco & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
                SqlBanco = SqlBanco & " and codmacta = " & DBSet(RS1!cliprov, "T")
        
                Set RsBanco = New ADODB.Recordset
                RsBanco.Open SqlBanco, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsBanco.EOF Then
                    SqlBanco = DevuelveDesdeBD("descripcion", "bancos", "codmacta", txtCta(4).Text, "T")
                    If SqlBanco = "" Then SqlBanco = Mid(Me.txtDCta(4).Text, 1, 30)
                
                    SQL = SQL & DBSet(RsBanco.Fields(0), "T") & "," & DBSet(SqlBanco, "T") & ")"
                Else
                    SQL = SQL & ValorNulo & "," & ValorNulo & ")"
                End If
                Set RsBanco = Nothing
                
            Else
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            End If
             
        End If 'De cabecera menor que 3, es decir : 1y 2
    
    End If
    
    'Ejecutamos si:
    '   Cabecera=0 o 1
    '   Cabecera=2 y impo=0.  Esto sginifica que estamos desbloqueando el apunte e insertandolo para pasarlo a hco
    Debe = True
    If Cabecera = 3 Then Debe = False
    If Debe Then Conn.Execute SQL
    
    If Debe Then
        '++monica
        If Cabecera = 1 And Not VienedeGastos Then
        
            Dim Situacion As Byte
            
            Situacion = 1

            Select Case Combo1.ItemData(Combo1.ListIndex)
                Case vbTalon, vbPagare
                    Situacion = 1
            End Select

'            SQL = "update pagos set imppagad = (select sum(imppago) from pagos_realizados where numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & " AND numfactu=" & DBSet(RecuperaValor(RS1!Cliente, 2), "N") & " and fecfactu=" & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & " AND numorden =" & RecuperaValor(RS1!Cliente, 4) & " and codmacta = " & DBSet(RecuperaValor(RS1!Cliente, 5), "T") & ") "
            SQL = "update pagos set imppagad = coalesce(imppagad,0) + " & DBSet(ImporteInterno2, "N") & " "
            SQL = SQL & ", fecultpa = " & DBSet(FechaAsiento, "F")
            SQL = SQL & ", situacion = " & DBSet(Situacion, "N")
            
            SQL = SQL & ", ctabanc1 = " & DBSet(txtCta(4).Text, "T")
            
            SQL = SQL & " where numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "T")
            SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
            SQL = SQL & " and codmacta = " & DBSet(RecuperaValor(RS1!Cliente, 5), "T")

            Conn.Execute SQL

        ' en tmppendientes metemos la clave primaria de pagos_realizados y el importe en letra
            SQL = "insert into tmppendientes (codusu,serie_cta,factura,fecha,numorden,codmacta,observa) values ("
            SQL = SQL & vUsu.Codigo & "," & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," 'numserie
            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 2), "T") & "," 'numfactu
            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & "," 'fecfactu
            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & "," 'numorden
            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 5), "T") & "," 'codmacta
            If ImporteInterno2 < 0 Then
                SQL = SQL & DBSet("menos " & EscribeImporteLetra(ImporteFormateado(CStr(ImporteInterno))), "T") & ") "
            Else
                SQL = SQL & DBSet(EscribeImporteLetra(ImporteFormateado(CStr(ImporteInterno))), "T") & ") "
            End If
            
            Conn.Execute SQL

        End If
    
    End If
    
    
  
        
End Function


'----------------------------------------------------------
'   A partir de la tabla tmp
'   Se que cuentas hay y los vencimientos.Por lo tanto, comprobare
'   que si la fechas estan fuera de ejercicios o de ambito
'   y si hay cuentas bloquedas
Private Function ComprobarCuentasBloquedasYFechasVencimientos() As Boolean
    ComprobarCuentasBloquedasYFechasVencimientos = False
    On Error GoTo EComprobarCuentasBloquedasYFechasVencimientos
    Set Rs = New ADODB.Recordset
    

    Cad = "select codmacta,nommacta,numfactura,fecha,fecbloq,cliente from tmpfaclin,cuentas where codusu=" & vUsu.Codigo & " and cta=codmacta and not (fecbloq is null )"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
        If CDate(Rs!NumFactura) > Rs!FecBloq Then Cad = Cad & Rs!codmacta & "    " & Rs!FecBloq & "     " & Format(Rs!NumFactura, "dd/mm/yyyy") & Space(15) & RecuperaValor(Rs!Cliente, 1) & RecuperaValor(Rs!Cliente, 2) & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close


    If Cad <> "" Then
        Cad = vbCrLf & String(90, "-") & vbCrLf & Cad
        Cad = "Cta           Fec. Bloq            Fecha contab         Factura" & Cad
        Cad = "Cuentas bloqueadas: " & vbCrLf & vbCrLf & vbCrLf & Cad
        MsgBox Cad, vbExclamation
    Else
        ComprobarCuentasBloquedasYFechasVencimientos = True
    End If
EComprobarCuentasBloquedasYFechasVencimientos:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function



'Busca VTO
' Para no hacer muchos seect WHERE, hacemos un unico SELECT (mirsaux)
' ahora en esta funcion buscaremos el registro correspondiente
'
Private Function BuscarVtoPago(ByRef IT As ListItem) As Boolean
Dim Fin As Boolean
    BuscarVtoPago = False
    Fin = False
    miRsAux.MoveFirst
    While Not Fin
        'numfactu fecfactu numorden
        If miRsAux!codmacta = IT.Tag Then
            If miRsAux!NUmSerie = IT.Text Then
              If miRsAux!numfactu = IT.SubItems(1) Then
                If miRsAux!FecFactu = IT.SubItems(2) Then
                    If miRsAux!numorden = IT.SubItems(4) Then
                        'ESTE ES
                        BuscarVtoPago = True
                        Fin = True
                    End If
                End If
            End If
            End If
        End If
        If Not Fin Then
            miRsAux.MoveNext
            Fin = miRsAux.EOF
        End If
    Wend
End Function




'***********************************************************************************
'***********************************************************************************
'
'   NORMA 57  Pagos por ventanilla
'
'***********************************************************************************
'***********************************************************************************
Private Sub AjustarFechaVencimientoDesdeFicheroBancario()
Dim Fin As Boolean
    'Para cada item buscare en la tabla from tmpconext  WHERE codusu
    Set Rs = New ADODB.Recordset
    '(numserie ,codfaccl,fecfaccl,numorden )
    Cad = "select ccost,pos,nomdocum,numdiari,fechaent from tmpconext  WHERE codusu =" & vUsu.Codigo & " and numasien=0 "
    Cad = Cad & " ORDER BY 1,2,3,4"
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    If Rs.EOF Then
        Cad = "NINGUN VENCIMIENTO"
    Else
    For i = 1 To Me.ListView1.ListItems.Count
        Fin = False
        Rs.MoveFirst
        With ListView1.ListItems(i)
            
            While Not Fin
                'Buscamos el registro... DEBERIA ESTAR
                If Rs!CCost = .Text Then
                    If Rs!Pos = .SubItems(1) Then
                        If Format(Rs!nomdocum, "dd/mm/yyyy") = .SubItems(2) Then
                            If Rs!NumDiari = .SubItems(4) Then
                                'Le pongo como fecha de vto la fecha del cobro del fichero
                                Fin = True
                                .SubItems(3) = Format(Rs!FechaEnt)
                                .Checked = True
                            End If
                        End If
                    End If
                End If
                If Not Fin Then
                    Rs.MoveNext
                    If Rs.EOF Then
                        'Ha llegado al final, y no lo ha encotrado
                        Cad = Cad & "     " & .Text & .SubItems(1) & "  -  " & .SubItems(2) & vbCrLf
                        'Para que vuelva al ppio
                        Fin = True
                    End If
                End If
            Wend
        End With
    Next
    End If
    Rs.Close
    
    If Cad <> "" Then
        Cad = Cad & vbCrLf & "El programa continuara con la fecha de vencimiento"
        MsgBox "No se ha encotrado la fecha de cobro para los siguientes vencimientos:" & vbCrLf & Cad, vbExclamation
    End If
    Set Rs = Nothing
End Sub


Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from tipofpago where tipoformapago <> 1 order by descformapago", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!descformapago
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!tipoformapago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(ButtonIndex As Integer)
    Select Case ButtonIndex
        Case 1
            Generar2
    End Select
End Sub


Private Function EsTalonOPagare(NumSer As String, NumFact As String, FecFact As String, NumOrd As String) As Boolean
Dim SQL As String
Dim Tipo As Byte

    SQL = "select tipforpa from formapago, pagos  where pagos.codforpa = formapago.codforpa and pagos.numserie = " & DBSet(NumSer, "T")
    SQL = SQL & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F") & " and numorden = " & DBSet(NumOrd, "N")
    SQL = SQL & " and codmacta = " & DBSet(txtCta(5).Text, "T")
    
    Tipo = DevuelveValor(SQL)
    EsTalonOPagare = (CByte(Tipo) = 2 Or CByte(Tipo) = 3)

End Function








Private Sub cmdImprimir_Click()
Dim Tipo As Integer
Dim NomFile As String
Dim FechaImprDoc As Date
        
    Cad = ""
    For i = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then Cad = Cad & "1"
    Next i
    If Cad = "" Then
        MsgBox "Deberias selecionar algun vencimiento.", vbExclamation
        Exit Sub
    End If
    
    If txtCta(4).Text = "" Then
        MsgBox "Seleccione el banco", vbExclamation
        Exit Sub
    End If
    If Combo1.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de pago", vbExclamation
        Exit Sub
    End If

    'Pagos
    Tipo = Combo1.ItemData(Combo1.ListIndex)
    Select Case Tipo
        
        Case vbTalon, vbPagare, vbConfirming
        
            'Para los pagares. Vere si alguno de los VTOs esta ya
            If Tipo = vbPagare Then
                Cad = ""
                For i = 1 To Me.ListView1.ListItems.Count
                    If ListView1.ListItems(i).Checked Then
                        If Me.ListView1.ListItems(i).ForeColor = vbRed Then
                            'Ese vto YA esta en otra "documentos de pagares"
                            Cad = Cad & "    - " & Me.ListView1.ListItems(i).SubItems(1) & " " & Me.ListView1.ListItems(i).SubItems(10) & vbCrLf
                        End If
                    End If
                Next i
                
                If Cad <> "" Then
                    Cad = "Los siguientes vencimientos fueron pagados en un documento anterior" & vbCrLf & vbCrLf & Cad
                    Cad = Cad & vbCrLf & " NO deberia seguir con el proceso. ¿Continuar?"
                    If MsgBox(Cad, vbExclamation + vbYesNoCancel) <> vbYes Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            
            
    
    
                  ' llamamos a un formulario para que me introduzca la referencia de los talones o pagarés
                  Dim CadInsert As String
                  Dim CadValues As String
                  Dim SQL As String
                  
              SQL = "delete from tmppagos2 where codusu = " & vUsu.Codigo
              Conn.Execute SQL
    
              CadInsert = "insert into tmppagos2 (codusu,numserie,numfactu,fecfactu,numorden,fecefect,reftalonpag,bancotalonpag,codmacta) values "
              CadValues = ""
    
              For i = 1 To Me.ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(ListView1.ListItems(i).Text, "T") & "," & DBSet(ListView1.ListItems(i).SubItems(1), "T") & ","
                    CadValues = CadValues & DBSet(ListView1.ListItems(i).SubItems(2), "F") & "," & DBSet(ListView1.ListItems(i).SubItems(4), "N") & ","
                    CadValues = CadValues & DBSet(ListView1.ListItems(i).SubItems(3), "F") & ","
                    CadValues = CadValues & DBSet(ListView1.ListItems(i).SubItems(10), "T", "S") & "," & ValorNulo & "," & DBSet(ListView1.ListItems(i).Tag, "T") & "),"
                End If
              Next i
              
              If CadValues <> "" Then
                  Conn.Execute CadInsert & Mid(CadValues, 1, Len(CadValues) - 1)
    
    
                  Set frmTESTalPag = New frmTESRefTalonPagos
    
                  frmTESTalPag.Show vbModal
                  
                  Set frmTESTalPag = Nothing
                  
              End If

    
            
            
            
            
            
            
            
            
            
            'Veo que documento es
            ' Primero en el banco
            If Tipo <> vbConfirming Then
                
                If Tipo = vbPagare Then
                    Cad = "DocPagare"
                Else
                    Cad = "DocTalon"
                End If
                
                Cad = DevuelveDesdeBD(Cad, "bancos", "codmacta", Me.txtCta(4).Text, "T")
            Else
                Cad = ""
            End If
            
            If Cad = "" Then
            
                NomFile = Format(IdPrograma, "0000") & "-03"
                If Not PonerParamRPT(NomFile, NomFile) Then Exit Sub
            Else
                NomFile = Cad
            End If
            
            FechaImprDoc = Now
            
            If vParamT.PideFechaImpresionTalonPagare Then
                    Fecha = Now
                    Set frmC = New frmCal
                    frmC.Fecha = Fecha
                    frmC.Show vbModal
                    Set frmC = Nothing
                    If Cad <> "" Then FechaImprDoc = CDate(Cad)
            End If
            
            
            If GenerarDocumentos(FechaImprDoc) Then
            
                'Imrpimimos
                Screen.MousePointer = vbHourglass
                CadenaDesdeOtroForm = NomFile
                
                Imprimir
                
                
                
                    
                        If MsgBox("Ha sido correcta la impresión?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        
                                                   
                        
                            'Marzo 2013
                            'Updateare los vtos al nuevo valor
                            'Juluio 2013. Ponia FECHA2 que era la maxima fecha de los vtos seleccionados
                            ' ahora pone Fecha1  que es la fecha seleccionada como fecha de pago
                            NomFile = DevuelveDesdeBD("fecha1", "tmptesoreriacomun", "codusu", CStr(vUsu.Codigo), "N")
                            If Trim(NomFile) = "" Then NomFile = Text3(0).Text
                            FechaAsiento = CDate(NomFile)
                        
                        
                            'Pagares.
                            'Marcamos los documentos como doc.recibido
                            NomFile = DevuelveDesdeBD("distinct(reftalonpag)", "tmppagos2", "codusu", CStr(vUsu.Codigo), "N")
                         
                            
                            
                            For i = 1 To Me.ListView1.ListItems.Count
                                If ListView1.ListItems(i).Checked Then
                                    Cad = "UPDATE pagos SET emitdocum=1"
                                    Cad = Cad & ",ctabanc1 = '" & txtCta(4).Text & "'"
                                    If NomFile <> "" Then Cad = Cad & ", referencia = '" & NomFile & "' "
                                    'Marzo 2013. Fecha vto
                                    Cad = Cad & ",fecefect = '" & Format(FechaAsiento, FormatoFecha) & "'"
                                    
                                    With ListView1.ListItems(i)
                                        Cad = Cad & " WHERE numserie = '" & .Text
                                        Cad = Cad & "' AND numfactu = '" & .SubItems(1)
                                        Cad = Cad & "' and fecfactu = '" & Format(.SubItems(2), FormatoFecha)
                                        Cad = Cad & "' and numorden = " & .SubItems(4)
                                        Cad = Cad & " and codmacta = '" & .Tag & "'"
                                        
                                        ListView1.ListItems(i).SubItems(10) = NomFile
                                    End With
                                    Conn.Execute Cad
                                End If
                            
                            Next i
                            
                        
                            'Ha sido correcta la impresion
                            'Contabilizamos
                            Generar2
                        
                        Else
                            'TEngo que tirar atras los contadores en los PAGARES
                            If Combo1.ItemData(Combo1.ListIndex) = vbConfirming Then
                                ' NumRegElim = NumRegElim - 1
                                 NomFile = "UPDATE contadores set "
                                 If CDate(Text3(0).Text) > vParam.fechafin Then
                                     NomFile = NomFile & "contado2=contado2-"
                                 Else
                                     NomFile = NomFile & "contado1=contado1-"
                                 End If
                                 NomFile = NomFile & NumRegElim & " WHERE tiporegi = 'ZZ0'"
                                 Ejecuta NomFile
                            End If
                        End If
                    
                
                    
                
                Screen.MousePointer = vbDefault
                
            End If
        Case vbEfectivo
            'Tipo=0 efectivo
            
            
            NomFile = Format(IdPrograma, "0000") & "-01"
            If Not PonerParamRPT(NomFile, NomFile) Then Exit Sub
            
            If GenerarRecibos2 Then
                'textoherecibido
                'Imprimimos
                CadenaDesdeOtroForm = NomFile
                Imprimir
            End If
            
            
        Case vbTipoPagoRemesa
            'Recibo bancario
            '
            NomFile = Format(IdPrograma, "0000") & "-00"
            If Not PonerParamRPT(NomFile, NomFile) Then Exit Sub
            'If NomFile = "" Then Exit Sub  'El msgbox ya lo da la funcion
            CadenaDesdeOtroForm = NomFile
            If ListadoOrdenPago Then
           
                Imprimir
            End If
           
            
        End Select
        


    



End Sub





'******************************************************************************************************************************
'******************************************************************************************************************************
'
'   Generacion de documentos para la impresion
'
'******************************************************************************************************************************
'******************************************************************************************************************************



Private Function GenerarDocumentos(FechaImpre As Date) As Boolean
Dim ListaProveedores As Collection
Dim Mc As Contadores
Dim SQL As String
Dim J As Integer

    
    On Error GoTo EGenDoc
    GenerarDocumentos = False
    
    'Preparo datos
    'Eliminamos temporales
    Cad = "Delete from tmpTesoreriaComun where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    'tmp347carta --> tmpTesoreria2    Ya no existe tabla tmp347carta. La nueva lleva los mismos campos
    Cad = "Delete from tmpTesoreria2 where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    
    'Enero2013. Ver abajo
    Cad = "DELETE from tmp340 where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    'Junio 2014
    'usuarios.z340
    ' -Grabara  datos del  banco propio(direccion y CCC).
    ' -Para herbelca dara mensaje de error si no existe
    
    
    'Datos carta
    'Datos basicos de la empresa para la carta
    Cad = "INSERT INTO tmpTesoreria2 (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, "
    Cad = Cad & "parrafo1, parrafo2, contacto, despedida,saludos,parrafo3, parrafo4, parrafo5, Asunto, Referencia)"
    Cad = Cad & " VALUES (" & vUsu.Codigo & ", "
    
    'Estos datos ya veremos com, y cuadno los relleno
    Set miRsAux = New ADODB.Recordset
    SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia,contacto from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Paarafo1 Parrafo2 contacto
    SQL = "'" & Format(Text3(0).Text, "dd mmmm yyyy") & "','',''"
    'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
    SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'," & SQL
    If Not miRsAux.EOF Then
        SQL = ""
        For i = 1 To 6
            SQL = SQL & DBLet(miRsAux.Fields(i), "T") & " "
        Next i
        SQL = Trim(SQL)
        SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
        SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"
        'Parrafo1, parrafo2
        SQL = SQL & ",'" & DevNombreSQL(DBLet(miRsAux!Poblacion)) & " " & Format(Text3(0).Text, "dd mmmm yyyy") & "','"
        SQL = SQL & DevNombreSQL(DBLet(miRsAux!Poblacion)) & "(" & DBLet(miRsAux!provincia) & ")'"
        'Contaccto
        SQL = SQL & ",'" & DevNombreSQL(DBLet(miRsAux!contacto)) & "' "
    End If
    miRsAux.Close
  
    Cad = Cad & SQL

   
    SQL = DevNombreSQL(txtDCta(4).Text)

    '
    Cad = Cad & ",'" & SQL & "',"
    
    NumRegElim = 0
    
    
    If Combo1.ItemData(Combo1.ListIndex) = vbTalon Or Combo1.ItemData(Combo1.ListIndex) = vbPagare Then
    
        SQL = "Select reftalonpag from tmppagos2 where codusu = " & vUsu.Codigo & " GROUP BY reftalonpag"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            SQL = ""
            While Not miRsAux.EOF
                NumeroTalonPagere = DBLet(miRsAux.Fields(0), "T")
                SQL = SQL & "X"
                miRsAux.MoveNext
            Wend
        End If
        miRsAux.Close
        If Len(SQL) <> 1 Then NumeroTalonPagere = ""
    End If
    
    If NumeroTalonPagere = "" Then
        NumeroTalonPagere = txtDCta(4).Text
       
    End If
    
    
    
    
    
    If NumeroTalonPagere = "" Then
        Cad = Cad & "NULL"
    Else
        Cad = Cad & "'" & DevNombreSQL(NumeroTalonPagere) & "'"
    End If

        
    
    
    'Pongo tb la fecha vto en parrafo 3
    Cad = Cad & ",'" & Text3(0).Text & "'"
    
    'Si tiene numerodetalonpagare entonces
    
    SQL = "NULL"
    If NumeroTalonPagere <> "" Then
        SQL = "codusu = " & vUsu.Codigo & " AND Pasivo = 'Z' AND codigo "
        SQL = DevuelveDesdeBD("QueCuentas", "tmpimpbalance", SQL, "1", "N")
        If SQL = "" Then
            SQL = "NULL"
        Else
            SQL = "'" & DevNombreSQL(SQL) & "'"
        End If
    End If
    Cad = Cad & "," & SQL
    'Parrafo 5 Updateare el importe total
    Cad = Cad & ", NULL,  NULL,  NULL)"
    Conn.Execute Cad
    SQL = ""
    
    
    'Contador de inserciones
    NumRegElim = 1
    
    
    
    DescripcionTransferencia = "|"
    'Veremos cuantos proveedores distintos hay y cuales son
    Set ListaProveedores = New Collection
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            Cad = "|" & ListView1.ListItems(i).Tag & "|"
            If InStr(1, DescripcionTransferencia, Cad) = 0 Then
                DescripcionTransferencia = DescripcionTransferencia & ListView1.ListItems(i).Tag & "|"
                ListaProveedores.Add ListView1.ListItems(i).Tag
            End If
        End If
    Next i
   
   
   Set Mc = New Contadores
   Fecha = CDate(Text3(0).Text)
   
   For J = 1 To ListaProveedores.Count
        '                     EL DOS es contadores pagare confirming
        
        
        If Combo1.ItemData(Combo1.ListIndex) = vbConfirming Then
            'ZZ0','CONFIRMING','0','0','0',NULL);
            If Mc.ConseguirContador("ZZ0", Fecha <= vParam.fechafin, True) = 0 Then
                GenerarDocumentos2 ListaProveedores.Item(J), Mc.Contador, FechaImpre
                NumRegElim = ""
            Else
                Exit Function
            End If
        Else
            GenerarDocumentos2 ListaProveedores.Item(J), 1, FechaImpre
        End If
    Next J
    
    
    'Enero 2013
    'Banco para los confirming que lo requieran
    
    'Julio 2014
    'Los graba para todo, solo que da mensaje su es operaciones aseguradas
    
    
   
    SQL = "select bancos.descripcion,cuentas.dirdatos,bancos.iban,nommacta  from bancos,cuentas "
    SQL = SQL & " where bancos.codmacta=cuentas.codmacta AND bancos.codmacta = '" & txtCta(4).Text & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        'ERROR obteniendo cuentas
        
        If vParamT.TieneOperacionesAseguradas Then MsgBox "Error obteniendo datos cta. contable banco", vbExclamation
    Else
        'ok
        'z340(codusu,codigo,razosoci,dom_intracom,nifdeclarado,nifrepresante,codpais,cp_intracom)
        SQL = DBLet(miRsAux!Descripcion, "T")
        If SQL = "" Then SQL = miRsAux!Nommacta
        SQL = ",1,'" & DevNombreSQL(SQL) & "','" & DevNombreSQL(DBLet(miRsAux!dirdatos, "T")) & "','"
        SQL = SQL & Mid(miRsAux!IBAN, 1, 4) & "','" & Mid(miRsAux!IBAN, 9, 4) & "','" & Mid(miRsAux!IBAN, 13, 2) & "','"
        SQL = SQL & Right(miRsAux!IBAN, 10) & "'," & Val(Mid(miRsAux!IBAN, 5, 4)) & ")"
        SQL = "INSERT INTO tmp340(codusu,codigo,razosoci,dom_intracom,nifdeclarado,nifrepresante,codpais,cp_intracom,numreg) VALUES (" & vUsu.Codigo & SQL
        Conn.Execute SQL
    End If
    miRsAux.Close

    
    
    DescripcionTransferencia = ""
    Set miRsAux = Nothing
    Set ListaProveedores = Nothing
    Set Mc = Nothing
    GenerarDocumentos = True
    Exit Function
EGenDoc:
    MuestraError Err.Number, , Err.Description
End Function


Private Function GenerarDocumentos2(Cta As String, numContador As Long, FechaImpresion As Date) As Boolean
Dim Aux As String
Dim SQL As String
Dim ColVtosQuePago As Collection
Dim FVto As Date
    
Dim QuitarFechaVto_ As Byte  '0 No quito vto.  1: Natural Horto  2: Morales -AVAB
Dim SqlBanco As String

    
    'Si me ha puesto numero de documento.
      'Sio ha escrito NUMERO de talong / pagare
    'De momento no utlizo nada de esto
    If False Then
        'St op
        SqlBanco = " codusu = " & vUsu.Codigo & " And codmacta = " & DBSet(Cta, "T") & " AND 1 "
        SQL = " 1  order by reftalonpag desc, bancotalonpag desc"
        SQL = DevuelveDesdeBD("concat(reftalonpag,'|', bancotalonpag,'|')", "tmppagos2", SqlBanco, SQL)
        If SQL <> "" Then
            'ACtualziamos la tabla donde estar los datos del proveedor, que llevara el numero de talon/pagare
        
        End If
    End If
    
    
    
    
    
    
    
    
    
    ' De momento, y a falta de poner parametro, buscare en nombre de la empresa
    'En hortonature / natural de montaña quitamos la fecha de veto para que quepan en dos clolumnas
    
    i = InStr(1, UCase(vEmpresa.nomempre), "MONTAÑA")
    If i = 0 Then i = InStr(1, UCase(vEmpresa.nomempre), "HORTONATURE")
    If i > 0 Then i = 1
        
    'AVAB y morales tampopo lo quiere
    If i = 0 Then
        i = InStr(1, UCase(vEmpresa.nomempre), "MORALES")
        If i = 0 Then i = InStr(1, UCase(vEmpresa.nomempre), "BLANQUETA")
        If i > 0 Then i = 2
    End If
    QuitarFechaVto_ = CByte(i)
    
    
    'La fecha de vencimiento debe coger la MAYOR de todas
    FVto = "01/01/1900"
    For i = 1 To ListView1.ListItems.Count
        With ListView1.ListItems(i)
            If .Checked Then
                If .Tag = Cta Then
                    If CDate(.SubItems(2)) > FVto Then FVto = CDate(.SubItems(2))
                End If
            End If
        End With
    Next
    
    impo = 0
    SubItemVto = 0 'Si vale uno es que ya hemos cojido los datos del proveedor
    SQL = ""
    Set ColVtosQuePago = New Collection
    For i = 1 To ListView1.ListItems.Count
        With ListView1.ListItems(i)
            If .Checked Then
                    If .Tag = Cta Then
                    Importe = ImporteFormateado(.SubItems(9))
                    impo = impo + Importe
                    
                    'Febrero 2010.   Llevara encolumnados los vtos que pago
                    'Llevara el listado de los pagos que efectuamos
                    'Antes: SQL = SQL & ".- " & Mid(.Text + Space(10), 1, 10)
                    '                 fra             fecfac              vto                  fecvenci
                    SQL = .Text & ":" & .SubItems(1) & "|" & .SubItems(2) & "|" & .SubItems(3) & "|" & .SubItems(9) & "|"
                    ColVtosQuePago.Add SQL
                    
                    'SaltoLinea
                    If SubItemVto = 0 Then
                        SubItemVto = 1 'Para que no vuelva a entrar
                        
                        
                        ', texto3, texto4, texto5,texto6
                        Cad = "Select nommacta,razosoci,dirdatos,codposta,despobla,desprovi,obsdatos from cuentas where codmacta ='" & Cta & "'"
                        
                        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If miRsAux.EOF Then Err.Raise 513, , "Error obteniendo datos cuenta: " & Cta
                        'NO PUEDE SER EOF
                        Cad = miRsAux!Nommacta
                        If Not IsNull(miRsAux!razosoci) Then Cad = miRsAux!razosoci
                        Cad = "'" & DevNombreSQL(Cad) & "'"
                        'Direccion
                        Cad = Cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!dirdatos))) & "'"
                        'Poblacion
                        Aux = DBLet(miRsAux!codposta)
                        If Aux <> "" Then Aux = Aux & " - "
                        Aux = Aux & DevNombreSQL(CStr(DBLet(miRsAux!desPobla)))
                        Cad = Cad & ",'" & Aux & "'"
                        'Provincia
                        Cad = Cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!desProvi))) & "'"
                        
                        
                        'Textos
                        '---------
                        '1.- Recibo nª    texto1,texto2 y en cad texto3,4,5,6
                        Cad = "'" & Format(numContador, "0000000") & "',''," & Cad
                        
                        'Marzo 2015
                        ' Herbelca Observaciones de la cuentas. Si las quiere sacar . Bajo de la direccion
                        '-------------------
                        Cad = Cad & ",'" & DevNombreSQL(Memo_Leer(miRsAux!obsdatos)) & "'"
                        
                        miRsAux.Close
                        
                        
                        'FECFAS
                        '--------------
                        'Libramiento o pago
                        Cad = Cad & ",'" & Format(Text3(0).Text, FormatoFecha) & "'"
                        'Cad = Cad & ",'" & Format(.SubItems(2), FormatoFecha) & "'"  antes Ene 2013
                        Cad = Cad & ",'" & Format(FVto, FormatoFecha) & "'"  '        AHORA Ene 2013
                        
                        '3era fecha  NULL
                        Cad = Cad & "," & DBSet(FechaImpresion, "F")
                        
                    
                    End If
                End If
            End If
        End With
    Next i
                
    'OBSERVACIONES1, observaciones 2 e importe en aux
    '------------------
    Importe = impo
    Aux = EscribeImporteLetra(impo)
    Aux = "       ** " & Aux
    Cad = Cad & ",'" & Aux & "**'"
    
    'Los vencimientos
    SQL = ""
    For i = 1 To ColVtosQuePago.Count
        'Codigo fra. Reservamos 10 espacios
        
        If Mid(ColVtosQuePago.Item(i), 1, 2) <> "1:" Then
            'Series 'creadas a mano
            Aux = RecuperaValor(CStr(ColVtosQuePago.Item(i)), 1)
            If Len(Aux) > 10 Then
                Aux = RecuperaValor(CStr(ColVtosQuePago.Item(i)), 1)
                Aux = Mid(Aux, 3)
            End If
            Aux = Mid(Aux & Space(20), 1, 20) & " "
            
        Else
            Aux = RecuperaValor(CStr(ColVtosQuePago.Item(i)), 1)
            Aux = Mid(Aux, 3)
            Aux = Mid(Aux & Space(20), 1, 20) & " "
        End If
        'Feb2020. No quito el espacio If QuitarFechaVto_ = 1 Then Aux = "   " & Aux
        
        'Separo la fecha para algunos
        If QuitarFechaVto_ = 2 Then Aux = Aux & "    "
        Aux = Aux & Mid(Format(RecuperaValor(CStr(ColVtosQuePago.Item(i)), 2), "dd/mm/yyyy") & Space(10), 1, 10) & "   "
        
        
        If QuitarFechaVto_ = 0 Then 'para todos menos para natural de montaña-morales
            Aux = Aux & Format(RecuperaValor(CStr(ColVtosQuePago.Item(i)), 3), "dd/mm/yyyy") & "   "
             'Solo reservo pocos espacios, muy justos
            Aux = Aux & Right(Space(13) & RecuperaValor(CStr(ColVtosQuePago.Item(i)), 4), 13) & " "
        Else
        '    'Solo reservo pocos espacios, muy justos
            Aux = Aux & Right(Space(14) & RecuperaValor(CStr(ColVtosQuePago.Item(i)), 4), 14) & " "
        End If
       
       
        If SQL <> "" Then SQL = SQL & vbCrLf
        SQL = SQL & Aux
    Next i
    
    Cad = Cad & ",'" & DevNombreSQL(SQL) & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
        
        
    SQL = "INSERT INTO tmptesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, "
    SQL = SQL & "texto6, observa2, fecha1, fecha2, fecha3, observa1, texto,importe1)"
    SQL = SQL & " VALUES (" & vUsu.Codigo & ","


    Conn.Execute SQL & NumRegElim & "," & Cad
    NumRegElim = NumRegElim + 1

       
    SQL = "UPDATE tmpTesoreria2 SET parrafo5 = '" & Format(Importe, FormatoImporte) & "' WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL
End Function
















'*******************************************************************************************
'*******************************************************************************************
'
'*******************************************************************************************
'*******************************************************************************************
Private Function GenerarRecibos2() As Boolean
Dim SQL As String
Dim Contador As Integer
Dim J As Integer
Dim Poblacion As String



    On Error GoTo EGenerarRecibos
    GenerarRecibos2 = False
    

    
    
    'Limpiamos
    Cad = "Delete from tmpTesoreriaComun where codusu = " & vUsu.Codigo
    Conn.Execute Cad


    'Guardamos datos empresa
    Cad = "Delete from tmpTesoreria2 where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "INSERT INTO tmpTesoreria2 (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos, "
    Cad = Cad & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, contacto, Asunto, Referencia)"
    Cad = Cad & " VALUES (" & vUsu.Codigo & ", "
    
    'Estos datos ya veremos com, y cuadno los relleno
    Set miRsAux = New ADODB.Recordset
    SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
    SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'"
    If Not miRsAux.EOF Then
        SQL = ""
        For J = 1 To 6
            SQL = SQL & DBLet(miRsAux.Fields(J), "T") & " "
        Next J
        SQL = Trim(SQL)
        SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
        SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"
        Poblacion = DevNombreSQL(DBLet(miRsAux!Poblacion, "T"))
        
    End If
    miRsAux.Close
 
    Cad = Cad & SQL
    'otralinea,saludos
    Cad = Cad & ",NULL"
    'parrafo1
    SQL = ""
    If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
        If vParamT.IntereseCobrosTarjeta2 > 0 And ImporteGastosTarjeta_ > 0 Then
            'FALTA###
        
            SQL = "1"
            If Fecha <= vParam.fechafin Then SQL = "2"
            
        
            'ZZ1','VTAS CREDITO(NAV)','0','0','0',NULL);

            SQL = DevuelveDesdeBD("contado" & SQL, "contadores", "tiporegi", "'ZZ1'")
            If SQL = "" Then SQL = "1"
            J = Val(SQL) + 1
            SQL = Format(J, "00000")
        End If
    End If
    
    Cad = Cad & ",'" & SQL & "'"
    

    
    
    '------------------------------------------------------------------------
    Cad = Cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
    Conn.Execute Cad

    'Empezamos
    SQL = "INSERT INTO tmptesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, "
    SQL = SQL & "texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion)"
    SQL = SQL & " VALUES (" & vUsu.Codigo & ","


    Contador = 0
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            'Lo insertamos tres veces
            
             RellenarCadenaSQLReciboPagos i
            
            'Lo rellenamos por triplicado    'VER ESTO
            'For J = 1 To 3
                Contador = Contador + 1
                Conn.Execute SQL & Contador & "," & Cad
            'Next J
        End If
    Next i
    GenerarRecibos2 = True
EGenerarRecibos:
    If Err.Number <> 0 Then
        MuestraError Err.Number
    End If
   
End Function






'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'
' Listado de efectos a pagar por el banco
Private Function ListadoOrdenPago() As Boolean
Dim SQL As String

    On Error GoTo EListadoOrdenPago
    ListadoOrdenPago = False

    'Borramos
    Cad = "DELETE from tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute Cad
    Set miRsAux = New ADODB.Recordset
    
    'insert into tmptesoreriacomun(texto1,texto2,observa1,texto3,fecha1,importe1,fecha2,opcion,texto4,codusu

    
   
    SQL = "select * from bancos where codmacta ='" & txtCta(4).Text & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        '---------------------------------------------------------
        SQL = DBLet(miRsAux!Descripcion, "T")
        If SQL = "" Then SQL = txtDCta(4).Text
        Cad = "'" & DevNombreSQL(SQL) & "','"
        
        'cad = cad & Format(DBLet(miRsAux!Entidad, "N"), "0000") & " "
        'cad = cad & Format(DBLet(miRsAux!oficina, "N"), "0000") & " "
        'cad = cad & DBLet(miRsAux!Control, "T") & " "
        'cad = cad & Format(DBLet(miRsAux!CtaBanco, "N"), "0000000000")
        Cad = Cad & DevNombreSQL(miRsAux!IBAN) & "',"
        Cad = ", (" & vUsu.Codigo & "," & Cad
     Else
        Cad = ""
    End If
    miRsAux.Close
    If Cad = "" Then
        MsgBox "Error leyendo el banco: " & "", vbExclamation
        Exit Function
    End If
    NumRegElim = 0
    
    SQL = DevSQL
    'Cargo el rs
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    SQL = ""
    For i = 1 To Me.ListView1.ListItems.Count
        NumRegElim = NumRegElim + 1
        If ListView1.ListItems(i).Checked Then
 
            impo = ImporteFormateado(ListView1.ListItems(i).SubItems(9))
            If impo > 0 Then
                
                
                If BuscarVtoPago(ListView1.ListItems(i)) Then
                    SQL = SQL & Cad
                    '`codusu`,`nombanco`,`cuentabanco`"  estan en cad
                    
                    
                   
                    SQL = SQL & "'" & DevNombreSQL(ListView1.ListItems(i).Tag) & "','" & DevNombreSQL(ListView1.ListItems(i).SubItems(1)) & "',"
                    SQL = SQL & "'" & Format(ListView1.ListItems(i).SubItems(2), FormatoFecha) & "'," & DevNombreSQL(ListView1.ListItems(i).SubItems(4)) & ","
                    SQL = SQL & "'" & Format(ListView1.ListItems(i).SubItems(3), FormatoFecha) & "',"
                    
                    'cad = cad & " `impefect`,`ctabanc1`,
                    SQL = SQL & TransformaComasPuntos(CStr(impo)) & ",'"
                    SQL = SQL & txtCta(4).Text & "'"
                    SQL = SQL & "," & miRsAux!emitdocum & ","
                                
                    '`entidad`,`oficina`,`CC`,`cuentaba`
                    'If Not IsNull(miRsAux!IBAN) Then
                    '    Sql = Sql & "'" & Format(miRsAux!IBAN, "0000") & "'"
                    'Else
                    '    Sql = Sql & "NULL"
                    'End If
                    SQL = SQL & "'" & ListView1.ListItems(i).SubItems(6) & "'"
                    
                    
                    'cad = cad & " `nomprove`"
                    SQL = SQL & ",'" & DevNombreSQL(ListView1.ListItems(i).SubItems(5)) & "'," & NumRegElim & ") "
                    NumRegElim = NumRegElim + 1
                    
                    
                Else
                    'NO HA ENCONTRADO EL VTO
                    MsgBox "Vto no encontrado: " & i, vbExclamation
                End If

                
            End If
        End If
        
    Next i
    
    
    'Cadena insercion
    If SQL <> "" Then
        SQL = Mid(SQL, 3)  'QUITO la primera coma
        
        'cad = "INSERT INTO usuarios.zlistadopagos (`codusu`,`nombanco`,`cuentabanco`,"
        'cad = cad & "`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,"
        'cad = cad & " `impefect`,`ctabanc1`,`ctabanc2`,`contdocu`,`entidad`,`oficina`,`CC`,`cuentaba`,"
        'cad = cad & " `nomprove`) VALUES "
        
        Cad = "insert into tmptesoreriacomun(codusu,observa1,texto1,texto5,texto3,"
        Cad = Cad & " fecha1,texto4,fecha2,importe1,texto2,opcion,texto6,observa2,codigo) VALUES "
        '(2000,'CUENTA FONDO OPERATIVO','ES4230582268522710000112','4000200758','3442',
        '2016-07-29',1,'2016-09-09',826.38,NULL,0,'UNITED PARCEL SERVICE')
        
        
        
        Cad = Cad & SQL
        Conn.Execute Cad
    End If
    
    If NumRegElim > 0 Then
        ListadoOrdenPago = True
    Else
        MsgBox "Ningun datos se ha generado", vbExclamation
    End If
    Set miRsAux = Nothing
    SegundoParametro = ""
    Exit Function
EListadoOrdenPago:
    MuestraError Err.Number, "ListadoOrdenPago"
    Set miRsAux = Nothing
End Function




Private Sub RellenarCadenaSQLReciboPagos(NumeroItem As Integer)
Dim Aux As String
Dim RT As ADODB.Recordset


    Set RT = New ADODB.Recordset
    With ListView1.ListItems(NumeroItem)
        'texto1 , texto2, texto3, texto4, texto5,
        'texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        
            
        'Textos
        '---------
        '1.- Recibo nª
        Cad = "'" & DevNombreSQL(.Text) & "',"
        
        Aux = "nommacta,razosoci,dirdatos,codposta,despobla,desprovi"
        Aux = "Select " & Aux & " from cuentas where codmacta = '" & .Tag & "'"
        RT.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            Aux = DBLet(RT.Fields(1), "T")
            If Aux = "" Then Aux = RT!Nommacta
            Aux = "'" & DevNombreSQL(Aux) & "'"
            For SubItemVto = 2 To 5
                Aux = Aux & ",'" & DevNombreSQL(DBLet(RT.Fields(SubItemVto), "T")) & "'"
            Next
        Else
            'VACIO. Error leyendo cuenta
            MsgBox "Error leyendo cuenta:" & .Tag, vbExclamation
            Aux = "'" & DevNombreSQL(.SubItems(4)) & "'"
            For SubItemVto = 2 To 5
                Aux = Aux & ",NULL"
            Next
        End If
        RT.Close
        Cad = Cad & Aux
        
        
        
        
        
        
        Importe = ImporteFormateado(.SubItems(9))
        
        'IMPORTES
        '--------------------
        Cad = Cad & "," & TransformaComasPuntos(CStr(Importe))
        
        'El segundo importe NULL
        Cad = Cad & ",NULL"
        
        'FECFAS
        '--------------
        'Libramiento o pago
        Cad = Cad & ",'" & Format(Text3(0).Text, FormatoFecha) & "'"
        Cad = Cad & ",'" & Format(.SubItems(2), FormatoFecha) & "'"
        
        '3era fecha  NULL
        Cad = Cad & ",NULL"
        
        'OBSERVACIONES
        '------------------
        Aux = EscribeImporteLetra(Importe)
        
        Aux = "       ** " & Aux
        Cad = Cad & ",'" & Aux & "**'"
        Cad = Cad & ",NULL"
        
        
        'OPCION
        '--------------
        Cad = Cad & ",NULL)"
        
        
    End With
    Set RT = Nothing
End Sub






Private Sub EliminarCobroPago(Indice As Integer)
Dim CadWhere As String
Dim C2 As String
    With ListView1.ListItems(Indice)
            C2 = ""
            CadWhere = " numserie = " & .Text
            CadWhere = CadWhere & " AND  numfactu= " & DBSet(.SubItems(1), "T")
            CadWhere = CadWhere & " and fecfactu = '" & Format(.SubItems(2), FormatoFecha)
            CadWhere = CadWhere & "' and numorden = " & .SubItems(4)
            CadWhere = CadWhere & " and codmacta = '" & .Tag & "'"
            If Combo1.ItemData(Combo1.ListIndex) = vbConfirming Then
                C2 = DevuelveDesdeBD("ctaconfirm", "pagos", CadWhere & " AND 1", "1")
                
            End If
    
            Cad = "UPDATE pagos set "
            Cad = Cad & " imppagad= " & DBSet(.SubItems(9), "N")
            Cad = Cad & " ,situacion = 1"
            
            If C2 <> "" Then   'es confirming que gestionamos por pagos por confirming
                Cad = Cad & " ,codmacta = " & DBSet(C2, "T")
            
            End If
            Cad = Cad & " WHERE " & CadWhere
    End With
    Ejecuta Cad
End Sub

