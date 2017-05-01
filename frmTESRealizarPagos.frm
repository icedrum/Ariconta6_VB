VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESRealizarPagos 
   Caption         =   "Realizar Pago"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   17520
   Icon            =   "frmTESRealizarPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
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
            Left            =   11430
            TabIndex        =   35
            Top             =   630
            Width           =   2325
         End
         Begin VB.CheckBox chkImprimir 
            Caption         =   "Imprimir Recibos"
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
            Height          =   255
            Left            =   11430
            TabIndex        =   32
            Top             =   330
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
            Left            =   9990
            TabIndex        =   30
            Text            =   "0000000000"
            Top             =   930
            Width           =   1305
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
            Width           =   2775
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
            Left            =   13770
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
            Left            =   13770
            TabIndex        =   21
            Top             =   930
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
            Left            =   13770
            TabIndex        =   20
            Top             =   630
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
            Left            =   13770
            TabIndex        =   19
            Top             =   330
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
            TabIndex        =   33
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
         Left            =   180
         TabIndex        =   34
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
Public Cobros As Boolean
Public OrdenarEfecto As Boolean
Public Regresar As Boolean
Public vTextos As String  'Dependera de donde venga
Public SegundoParametro As String
Public ContabTransfer As Boolean


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

    Cad = ""
    For i = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            Cad = Cad & "1"
            Exit For
        End If
    Next i
    If Cad = "" Then
        MsgBox "Deberias selecionar algún vencimiento", vbExclamation
        Exit Sub
    End If
    
    If Combo1.ListIndex = -1 Then
        MsgBox "Deberias selecionar el tipo de pago", vbExclamation
        Exit Sub
    End If
         
    If txtCta(4).Text = "" Then
        MsgBox "Deberias introducir la cuenta de banco", vbExclamation
        Exit Sub
    End If
    
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
    If vParamT.IntereseCobrosTarjeta > 0 Then
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
    If Cobros And (Combo1.ItemData(Combo1.ListIndex) = 6 Or Combo1.ItemData(Combo1.ListIndex) = 0) And ImporteGastosTarjeta_ > 0 Then
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
    If CDate(Text3(0).Text) < vParam.fechaini Then
        MsgBox "Fuera de ejercios.", vbExclamation
        Exit Sub
    Else
        Fecha = DateAdd("yyyy", 1, vParam.fechafin)
        If CDate(Text3(0).Text) > Fecha Then
            If MsgBox("Fecha de ejercicio aun no abierto. ¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
            
    End If
    
    
    
    If Not Cobros Then
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
    
    Cad = "Desea contabilizar los vencimientos seleccionados?"
    If Combo1.ItemData(Combo1.ListIndex) = 1 Then
        i = 0
        If Not Cobros Then
            If Not ContabTransfer Then i = 1
        Else
            If Not ContabTransfer And SegundoParametro <> "" Then i = 1
        End If
        If i = 1 Then
            'Estamos creando la transferencia o el pago domiciliado
            Cad = RecuperaValor(Me.vTextos, 5)
            If Cad = "" Then
                Cad = "Desea generar la transferencia?"
            Else
                Cad = "Desea generar el " & Cad & "?"
            End If
        End If
    End If
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
'???
    ' para la impresion
    Dim Sql As String
    
    Sql = "delete from tmppendientes where codusu = " & vUsu.Codigo
    Conn.Execute Sql
'???
    
    
    
    Screen.MousePointer = vbHourglass
    
    'Una cosa mas.
    'Si la forma de pago es talon/pagere, y me ha escrito el numero de talon pagare...
    'Se lo tengo que pasar a la contabilizacion, con lo cual tendre que grabar
    'el nº de talon pagare en reftalonpag
    
    

    If Combo1.ItemData(Combo1.ListIndex) = vbTalon Or Combo1.ItemData(Combo1.ListIndex) = vbPagare Then
    
    
          ' llamamos a un formulario para que me introduzca la referencia de los talones o pagarés
          Dim CadInsert As String
          Dim CadValues As String
          
          Sql = "delete from tmppagos2 where codusu = " & vUsu.Codigo
          Conn.Execute Sql

          CadInsert = "insert into tmppagos2 (codusu,numserie,numfactu,fecfactu,numorden,fecefect,reftalonpag,bancotalonpag,codmacta) values "
          CadValues = ""

          For i = 1 To Me.ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(ListView1.ListItems(i).Text, "T") & "," & DBSet(Val(ListView1.ListItems(i).SubItems(1)), "T") & ","
                CadValues = CadValues & DBSet(ListView1.ListItems(i).SubItems(2), "F") & "," & DBSet(ListView1.ListItems(i).SubItems(4), "N") & ","
                CadValues = CadValues & DBSet(ListView1.ListItems(i).SubItems(3), "F") & ","
                CadValues = CadValues & ValorNulo & "," & ValorNulo & "," & DBSet(ListView1.ListItems(i).Tag, "T") & "),"
            End If
          Next i
          
          If CadValues <> "" Then
              Conn.Execute CadInsert & Mid(CadValues, 1, Len(CadValues) - 1)


              Set frmTESTalPag = New frmTESRefTalonPagos

              frmTESTalPag.Show vbModal
              
              Set frmTESTalPag = Nothing
              
          End If
    
    End If
    
    
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
        If Cobros Then
            SubItemVto = 3
        Else
            SubItemVto = 2
        End If
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
    If ContabTransfer Then
        'Estamos contabilizando la transferencia
        Cad = "stransfer"
        If Cobros Then Cad = Cad & "cob"
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
'               EliminarCobroPago I
              
               ListView1.ListItems.Remove i
                
            End If
        Next i
        '-----------------------------------------------------------
    Else
        TirarAtrasTransaccion
    End If


    ImpSeleccionado = 0
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
'    If cad = "" Then GeneraLaContabilizacionFecha
'
'    Else
'        'La normal, lo que habia
'        GeneraLaContabilizacion
'
'    End If
        
    If chkImprimir.Value Then Imprimir2
    
        
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
    MuestraError Err.Number, "Contabilizando"
End Function



Private Sub Imprimir2()
Dim Sql As String
Dim EsCobroTarjetaNavarres As Boolean

    EsCobroTarjetaNavarres = False
    If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
        'Si tiene el parametro y le ha puesto valor
        If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then EsCobroTarjetaNavarres = True
        
        If EsCobroTarjetaNavarres Then
        
            'parrafo1
            Sql = ""
            If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
                If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then
                    Sql = "1"
                    If Fecha <= vParam.fechafin Then Sql = "2"
                    Sql = DevuelveDesdeBD("contado" & Sql, "contadores", "tiporegi", ContCreditoNav) 'tarjeta credito tipo NAVARRES
                    If Sql = "" Then Sql = "1"
                    J = Val(Sql) + 1
                    Sql = Format(J, "00000")
                End If
            End If
                
        End If
        
    End If


    frmTESImpRecibo.VienedeRealizarCobro = EsCobroTarjetaNavarres
    frmTESImpRecibo.pNumFactu = Sql
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
             If Not ContabTransfer Then
                 i = 1
                 Cad = RecuperaValor(vTextos, 5) 'Dira si es PAGO DOMICILIADO
                 If Cad <> "" Then
                     If vParamT.PagosConfirmingCaixa Then
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
                 If Not Cobros Then
                     Cad = RecuperaValor(vTextos, 5)
                     If Cad <> "" Then Caption = "Realizar pago DOMICILIADO "
                 End If
             End If
         End If
         
         Me.chkPorFechaVenci.Visible = i = 0
         chkGenerico(0).Visible = i = 0
'         Me.chkVtoCuenta(0).Visible = I = 0
         
        '++
        If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
            If vParamT.IntereseCobrosTarjeta > 0 Then
                ImporteGastosTarjeta_ = vParamT.IntereseCobrosTarjeta
                Text3(6).Text = Format(ImporteGastosTarjeta_, "##0.00")
            End If
        End If
        
        i = 0
        If Cobros And (Combo1.ItemData(Combo1.ListIndex) = 2 Or Combo1.ItemData(Combo1.ListIndex) = 3) Then i = 1
        Me.mnBarra1.Visible = i = 1
        Me.mnNumero.Visible = i = 1
        
        'check de impresion de documento
        '++
        Me.chkImprimir.Visible = True
        Me.chkImprimir.Enabled = True
        If Combo1.ItemData(Combo1.ListIndex) = 5 Then
            Me.chkImprimir.Visible = False
            Me.chkImprimir.Enabled = False
        End If
        Me.chkImprimir.Caption = "Imprimir Recibo"
        If Combo1.ItemData(Combo1.ListIndex) = vbTalon Or Combo1.ItemData(Combo1.ListIndex) = vbPagare Then
            Me.chkImprimir.Caption = "Imprimir Documento"
        End If
        '++
    
    
        CargaList
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        espera 0.1
        'Cargamos el LIST
        
'        CargaList
        
        'OCTUBRE 2014
        'PAgos por ventanilla.
        'Pondre como fecha de vencimiento la fecha que el
        'banco, en el fichero, me indica que realio el pago
        If Cobros And Combo1.ListIndex <> -1 Then
            If Combo1.ItemData(Combo1.ListIndex) = 0 Then
                If InStr(1, vSql, "from tmpconext  WHERE codusu") > 0 Then AjustarFechaVencimientoDesdeFicheroBancario
            End If
        End If
        
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
    
    
    CargaCombo
    
    CargaIconoListview Me.ListView1
    ListView1.Checkboxes = True
    imgCheck(0).Visible = True
    imgCheck(1).Visible = True
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
    Me.ImporteGastosTarjeta_ = 0
    Me.CodmactaUnica = ""
     
    DevuelveCadenaPorTipo False, Cad
    LeerparametrosContabilizacion

    'Efectuar cobros
    FrameRemesar.Visible = True
    ListView1.SmallIcons = Me.ImageList1
    CargaColumnas
    
    
    'Octubre 2014
    'Norma 57 pagos ventanilla
    'Si en el select , en el SQL, viene un
    If Cobros And Combo1.ListIndex <> -1 Then
        If Combo1.ItemData(Combo1.ListIndex) = 0 Then
            If InStr(1, vSql, "from tmpconext  WHERE codusu") > 0 Then chkPorFechaVenci.Value = 1
        End If
    End If
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
   If Cobros Then
        NCols = 13 '11
        Columnas = "Serie|Factura|F.Factura|F. VTO|Nº|CLIENTE|Tipo|Importe|Gasto|Cobrado|Pendiente|"
        Ancho = "800|10%|12%|12%|520|23%|840|12%|8%|11%|12%|0%|0%|"
        ALIGN = "LLLLLLLDDDDDD"
        
        
        ListView1.Tag = 2200  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
        
        If Combo1.ListIndex <> -1 Then
            If Combo1.ItemData(Combo1.ListIndex) = 2 Or Combo1.ItemData(Combo1.ListIndex) = 3 Then
                ''Si es un talon o pagare entonces añadire un campo mas
                NCols = NCols + 1
                Columnas = Columnas & "Nº Documento|"
                Ancho = Ancho & "2500|"
                ALIGN = ALIGN & "L"
            End If
        End If
   Else
        NCols = 12
        Columnas = "Serie|Factura|F. Fact|F. VTO|Nº|PROVEEDOR|Tipo|Importe|Pagado|Pendiente|"
        Ancho = "800|15%|12%|12%|400|26%|800|12%|12%|12%|0%|0%|"
        ALIGN = "LLLLLLLDDDDD"
        ListView1.Tag = 2200  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
   End If
        
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
    
    Set FrasMarcadas = Nothing
    
    Text2(2).Text = "0,00"
    Label2(2).Caption = "Selecionado"
    Label2(2).Visible = True
    Text2(2).Visible = True
    Label2(3).Visible = True
    
ECargando:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Text2(0).Text = Format(Importe, FormatoImporte)
    Text2(1).Text = Format(Vencido, FormatoImporte)
    
    Text2(2).Text = Format(riesgo, FormatoImporte)
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set Rs = Nothing
End Sub

Private Sub CargaCobros()
Dim Inserta As Boolean

    RiesTalPag = 0
    Cad = DevSQL
    
    'ORDENACION
    If CampoOrden = "" Then CampoOrden = "cobros.fecvenci"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    If CampoOrden <> "cobros.fecvenci" Then Cad = Cad & ", cobros.fecvenci"
    
    
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Inserta = True
        If Rs!tipoformapago = vbTipoPagoRemesa Then
            
        ElseIf Rs!tipoformapago = vbTalon Or Rs!tipoformapago = vbPagare Then
        
        End If
        
        If Inserta Then
            InsertaItemCobro
        End If  'de insertar
        
        Rs.MoveNext
    Wend
    Rs.Close
End Sub


Private Sub InsertaItemCobro()
Dim vImporte As Currency
Dim DiasDif As Long
Dim ImpAux As Currency

    Set ItmX = ListView1.ListItems.Add()
    
    ItmX.Text = Rs!NUmSerie
    ItmX.SubItems(1) = Rs!NumFactu
    ItmX.SubItems(2) = Format(Rs!FecFactu, "dd/mm/yyyy")
    ItmX.SubItems(3) = Format(Rs!FecVenci, "dd/mm/yyyy")
    ItmX.SubItems(4) = Rs!numorden
    ItmX.SubItems(5) = Rs!Nommacta
    ItmX.SubItems(6) = Rs!siglas
    
    ItmX.SubItems(7) = Format(Rs!ImpVenci, FormatoImporte)
    vImporte = DBLet(Rs!Gastos, "N")
    
    'Gastos
    ItmX.SubItems(8) = Format(vImporte, FormatoImporte)
    vImporte = vImporte + Rs!ImpVenci
    
    If Not IsNull(Rs!impcobro) Then
        ItmX.SubItems(9) = Format(Rs!impcobro, FormatoImporte)
        impo = vImporte - Rs!impcobro
        ItmX.SubItems(10) = Format(impo, FormatoImporte)
    Else
        impo = vImporte
        ItmX.SubItems(9) = "0.00"
        ItmX.SubItems(10) = Format(vImporte, FormatoImporte)
    End If
    If Rs!tipoformapago = vbTipoPagoRemesa Then
        '81--->
        'asc("Q") =81
        If Asc(Right(" " & DBLet(Rs!siturem, "T"), 1)) = 81 Then
            riesgo = riesgo + vImporte
        Else
      
        End If
    
    ElseIf Rs!tipoformapago = vbTalon Or Rs!tipoformapago = vbPagare Then
'            If RS!ImpVenci > 0 Then ItmX.SubItems(11) = DBLet(RS!reftalonpag, "T")
            If Rs!recedocu = 1 Then RiesTalPag = RiesTalPag + DBLet(Rs!impcobro, "N")
    End If
    
    If Rs!tipoformapago = vbTarjeta Then
        'Si tiene el parametro y le ha puesto valor
        If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then
            DiasDif = 0
            Fecha = CDate(Text3(0).Text)
            If Rs!FecVenci < Fecha Then DiasDif = DateDiff("d", Rs!FecVenci, Fecha)
            If DiasDif > 0 Then
                'Si ya tenia gastos.
                If DBLet(Rs!Gastos, "N") > 0 Then
                    MsgBox "Ya tenia gastos", vbExclamation
                    ItmX.ListSubItems(8).Bold = True
                    ItmX.ListSubItems(8).ForeColor = vbRed
                End If
                
                ImpAux = ((ImporteGastosTarjeta_ / 365) * DiasDif) / 100
                ImpAux = Round(ImpAux * impo, 2)
                
                impo = impo + ImpAux
                ItmX.SubItems(10) = Format(impo, FormatoImporte)
                'La de gastos
                ImpAux = DBLet(Rs!Gastos, "N") + ImpAux
                ItmX.SubItems(8) = Format(ImpAux, FormatoImporte)
            End If
            
        End If
    End If
    
    If Rs!FecVenci < Fecha Then
        'LO DEBE
        ItmX.SmallIcon = 1
        Vencido = Vencido + impo
    Else
'        ItmX.SmallIcon = 2
    End If
    Importe = Importe + impo
    
    ItmX.Tag = Rs!codmacta
    
    If Combo1.ListIndex <> -1 Then
        If Combo1.ItemData(Combo1.ListIndex) = 1 And SegundoParametro <> "" Then
            If Not IsNull(Rs!transfer) Then
                ItmX.Checked = True
                ImpSeleccionado = ImpSeleccionado + impo
            End If
        End If
    End If
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
        If Combo1.ListIndex > 0 Then
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
     ItmX.SubItems(1) = Rs!NumFactu
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
         If DBLet(Rs!Referencia, "T") = "" Then ItmX.ListSubItems(4).ForeColor = vbMagenta
     End If
     
     ' nuevo si está marcada lo miramos
     For i = 1 To FrasMarcadas.Count
        Cad = FrasMarcadas.Item(i)
        If RecuperaValor(Cad, 1) = Rs!NUmSerie And CStr(RecuperaValor(Cad, 2)) = CStr(Rs!NumFactu) And CStr(RecuperaValor(Cad, 3)) = CStr(Rs!FecFactu) And RecuperaValor(Cad, 4) = Rs!numorden And RecuperaValor(Cad, 5) = Rs!codmacta Then
            ItmX.Checked = True
'            FrasMarcadas.Remove I
        End If
     Next i
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Para dejar las variables bien
    ContabTransfer = False
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
            
       Case 1 ' cuenta de proveedor
            Set frmCCtas = New frmColCtas
            DevfrmCCtas = ""
            Cad = txtCta(5).Text
            frmCCtas.DatosADevolverBusqueda = "0"
            frmCCtas.FILTRO = "2" ' proveedores
            frmCCtas.Show vbModal
            Set frmCCtas = Nothing
            If Cad <> txtCta(5).Text Then txtCta_LostFocus 5
    End Select
    PonFoco txtCta(Index + 4)
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
            If FechaCorrecta2(CDate(Cad), True) < 2 Then Text3(0).Text = Cad
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


    'Fechas fin ejercicios
    FechaFinEjercicios = DateAdd("yyyy", 1, vParam.fechafin)


    '++
    GastosTransferencia = CCur(ComprobarCero(Text3(6).Text))



    Aux = "INSERT INTO tmpfaclin (codusu, codigo, Fecha,Numfactura, cta, Cliente, NIF, Imponible,  Total) "
    Aux = Aux & "VALUES (" & vUsu.Codigo & ","
    For J = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(J).Checked Then
            C = J & ",'"
            'Si la fecha de contabilizacion esta fuera de ejercicios
            If Cobros Then
                FechaContab = CDate(ListView1.ListItems(J).SubItems(3))
            Else
                FechaContab = CDate(ListView1.ListItems(J).SubItems(2))
            End If
            

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
            'Serie factura |FECHAfactura|
            'Neuvo febrero 2008 Serie |factura |FECHAfactura|numvto|
            If Cobros Then
                C = C & ListView1.ListItems(J).Text & "|" & ListView1.ListItems(J).SubItems(1) & "|" & ListView1.ListItems(J).SubItems(2) & "|" & ListView1.ListItems(J).SubItems(4)
            Else
                C = C & DevNombreSQL(ListView1.ListItems(J).Text) & "|" & ListView1.ListItems(J).SubItems(1) & "|" & ListView1.ListItems(J).SubItems(2) & "|" & ListView1.ListItems(J).SubItems(4) & "|" & ListView1.ListItems(J).Tag
            End If
            C = C & "|','"
            
            'Cuenta agrupacion cobros
            If Combo1.ItemData(Combo1.ListIndex) = 1 And ContabTransfer Then
                C = C & Me.chkGenerico(1).Tag & "',"
            Else
                C = C & Me.chkGenerico(0).Tag & "',"
            End If
            'Dinerito
            'riesgo es GASTO
            i = ColD(0)
            impo = ImporteFormateado(ListView1.ListItems(J).SubItems(i))
            If Cobros Then
                riesgo = ImporteFormateado(ListView1.ListItems(J).SubItems(i - 2))
            Else
                riesgo = 0
            End If
            impo = impo - riesgo
            C = C & TransformaComasPuntos(CStr(impo)) & "," & TransformaComasPuntos(CStr(riesgo)) & ")"


            'Lo meto en la BD
            C = Aux & C
            Conn.Execute C
        End If
    Next J

    'Si es por tarjeta hay una opcion para meter el gasto total
    'que a partir de la cuenta de banco gasto tarjeta crear una linea mas
    If Cobros And Combo1.ItemData(Combo1.ListIndex) = 6 And ImporteGastosTarjeta_ > 0 Then
            
            'Agosto 2014
            'Pagos credito NAVARRES  --> NO llevan esta linea
            
            If vParamT.IntereseCobrosTarjeta = 0 Then
                Cad = DevuelveDesdeBD("ctagastostarj", "bancos", "codmacta", Text3(1).Tag, "T")
                
                FechaContab = CDate(Text3(0).Text)
                C = "'" & Format(FechaContab, FormatoFecha) & "'"
                C = C & "," & C
                C = J & "," & C & ",'" & Cad & "','"
                'Serie factura |FECHAfactura| ----> pondre: "gastos" | fecha contab
                C = C & "GASTOS|" & FechaContab & "|','" & Cad & "',"
                'Dinerito
                'riesgo es GASTO
                impo = -ImporteGastosTarjeta_
                C = C & TransformaComasPuntos(CStr(impo)) & ",0)"
                C = Aux & C
                Conn.Execute C
            End If
    End If
    
    'Gastos contabilizacion transferencia
'    If Combo1.ItemData(Combo1.ListIndex) = 1 And GastosTransferencia <> 0 Then
    If GastosTransferencia <> 0 Then
            'aqui ira los gastos asociados a la transferencia
            'Hay que ver los lados
            
            'Cad = DevuelveDesdeBD("ctagastostarj", "ctabancaria", "codmacta", Text3(1).Tag, "T")
            Cad = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", Text3(1).Tag, "T")
            
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
Dim Sql As String
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
    If Combo1.ItemData(Combo1.ListIndex) = 1 And ContabTransfer Then
        If Me.chkContrapar(1).Value Then ContraPartidaPorLinea = False
        If Me.chkAsiento(1).Value Then UnAsientoPorCuenta = True
        If chkGenerico(1).Value Then PonerCuentaGenerica = True
        
        'Si lleva GastosTransferencia entonce AGRUPAMOS banco
        If GastosTransferencia <> 0 Then
            
            'gastos tramtiaacion transferenca descontados importe
            Sql = DevuelveDesdeBD("GastTransDescontad", "bancos", "codmacta", Text3(1).Tag, "T")
            GastosTransDescontados = Sql = "1"
            
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
            Sql = DevuelveDesdeBD("GastTransDescontad", "bancos", "codmacta", Text3(1).Tag, "T")
            GastosTransDescontados = Sql = "1"
            
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
    Sql = CampoCuenta & " AS cliprov,"
    'Selecciona
    Sql = "select count(*) as numvtos,codigo,numfactura,fecha,cliente," & Sql & "sum(imponible) as importe,sum(total) as gastos from tmpfaclin"
    Sql = Sql & " where codusu =" & vUsu.Codigo & " GROUP BY "
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
    Cad = Sql & Cad
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
                    If Cobros Then
                        ImpBanco = -GastosTransferencia
                    Else
                        ImpBanco = -GastosTransferencia
                    End If
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
    
        If Cobros Then
            riesgo = riesgo + Rs!Gastos
        Else
            riesgo = 0
        End If
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
    If Not Cobros Then ColD = ColD - 1
End Function

Private Sub EliminarCobroPago(Indice As Integer)
    
    With ListView1.ListItems(Indice)
        If Cobros Then
            
            Cad = "DELETE FROM  cobros WHERE "
            Cad = Cad & " numserie  = '" & .Text
            Cad = Cad & "' and numfactu = " & .SubItems(1)
            Cad = Cad & " and numorden = " & .SubItems(4)
            Cad = Cad & " and fecfactu = '" & Format(.SubItems(2), FormatoFecha) & "'"
            
        Else
            Cad = "DELETE FROM  pagos WHERE "
            Cad = Cad & " numfactu = '" & .Text
            Cad = Cad & "' and fecfactu = '" & Format(.SubItems(1), FormatoFecha)
            Cad = Cad & "' and numorden = " & .SubItems(3)
            Cad = Cad & " and ctaprove = '" & .Tag & "'"
        End If
    End With
    Ejecuta Cad
End Sub



Private Sub Text3_GotFocus(Index As Integer)
    ConseguirFoco Text3(Index), 3
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
    KEYpress KeyAscii
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim DevfrmCCtas As String
Dim Sql As String

    Select Case Index
        Case 4 ' cuenta de banco
            txtCta(Index).Text = Trim(txtCta(Index).Text)
            DevfrmCCtas = txtCta(Index).Text
            i = 0
            If DevfrmCCtas <> "" Then
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, Sql) Then
                    DevfrmCCtas = DevuelveDesdeBD("codmacta", "bancos", "codmacta", DevfrmCCtas, "T")
                    If DevfrmCCtas = "" Then
                        Sql = ""
                        MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                    End If
                Else
                    MsgBox Sql, vbExclamation
                    DevfrmCCtas = ""
                    Sql = ""
                End If
                i = 1
            Else
                Sql = ""
            End If
            
            
            txtCta(Index).Text = DevfrmCCtas
            txtDCta(Index).Text = Sql
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
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, Sql) Then
                    
                Else
                    MsgBox Sql, vbExclamation
                    If Index < 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
                        DevfrmCCtas = ""
                        Sql = ""
                    End If
                End If
                i = 1
            Else
                Sql = ""
            End If
            
            txtCta(Index).Text = DevfrmCCtas
            txtDCta(Index).Text = Sql
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
Dim Sql As String
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
    If Cobros Then
        Ampliacion = vp.diaricli
    Else
        Ampliacion = vp.diaripro
    End If
    
    If Cabecera = 0 Then
        'La cabecera
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
        Sql = Sql & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador
        Sql = Sql & ",  '"
        Sql = Sql & "Generado desde Tesorería el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre
        If Combo1.ItemData(Combo1.ListIndex) = 1 And Not Cobros Then
            'TRANSFERENCIA
            Ampliacion = DevuelveDesdeBD("descripcion", "stransfer", "codigo", SegundoParametro, "N")
            If Ampliacion <> "" Then
                Ampliacion = "Concepto: " & Ampliacion
                Ampliacion = DevNombreSQL(Ampliacion)
                Ampliacion = vbCrLf & Ampliacion
                Sql = Sql & Ampliacion
            End If
        End If
        
        Sql = Sql & "',"
        Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Pagos'"

        
        Sql = Sql & ")"
        NumLine = 0
     
    Else
        If Cabecera < 3 Then
            'Lineas de apuntes o cabecera.
            'Comparten el principio
             Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
             Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfacpr, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
             Sql = Sql & "VALUES (" & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador & "," & NumLine & ",'"
             
             '1:  Asiento para el VTO
             If Cabecera = 1 Then
                 'codmacta
                 'Si agrupa la cuenta entonces
                 Sql = Sql & RS1!cliprov & "','"
                 
                 
                 'numdocum: la factura
                 If NumVtos > 1 Then
                    Ampliacion = "Vtos: " & NumVtos
                 Else
                    Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 2))
                 End If
                 Sql = Sql & Ampliacion & "',"
                
                
                 'Veamos si va al debe, al haber, si ponemos concepto debe / haber etc eyc
                 If Cobros Then
                    'CLIENTES
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
                        Conce = vp.condecli
                        TipoAmpliacion = vp.ampdecli
                        PonerContrPartida = vp.ctrdecli = 1
                    Else
                        Conce = vp.conhacli
                        TipoAmpliacion = vp.amphacli
                        PonerContrPartida = vp.ctrhacli = 1
                    End If
                 
                 
                 Else
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
                     
                 End If
                
                
                 Sql = Sql & Conce & ","
                 
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
                    'Estamos en la amplicacion del cliente. Es una tonteria tener esta opcion marcada, pero bien
                    Ampliacion = RecuperaValor(vTextos, 3)
                    Ampliacion = Mid(Ampliacion, InStr(1, Ampliacion, "-") + 1)
                Case 5
                    'Si hubiera que especificar mas el documento
'                    If Tipo = vbTalon Then
'                        AUX = "TAL Nº"
'                    Else
'                        AUX = "PAG Nº"
'                    End If
'
                
                    If Cobros Then
                        'Veo la el camporefencia de ese talon
                        'Antes cogiamos numero fra
                        'ahora contrapar
                        Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)  'Num tal pag
                        If False Then
                            
                            Ampliacion = "numserie = '" & RecuperaValor(RS1!Cliente, 1) & "' AND RecuperaValor(RS1!Cliente, 2)"
                            Ampliacion = Ampliacion & " AND numorden = " & RecuperaValor(RS1!Cliente, 4) & " AND fecfactu "
                            Ampliacion = DevuelveDesdeBD("reftalonpag", "cobros", Ampliacion, Format(RecuperaValor(RS1!Cliente, 3), FormatoFecha), "F")
                            
                        Else
                            'Es numero tal pag + ctrpar
                            DescripcionTransferencia = RecuperaValor(vTextos, 2)
                            DescripcionTransferencia = Mid(DescripcionTransferencia, InStr(1, DescripcionTransferencia, "-") + 1)
                            Ampliacion = Ampliacion & " " & DescripcionTransferencia
                            DescripcionTransferencia = ""
                        End If
                        If Ampliacion = "" Then
                            Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                        Else
                            Ampliacion = " NºDoc: " & Ampliacion
                        End If
                    Else
                        If NumeroTalonPagere = "" Then
                            Ampliacion = ""
                        Else
                            'Cta banco
                            Ampliacion = RecuperaValor(vTextos, 2)
                            Ampliacion = Mid(Ampliacion, InStr(1, Ampliacion, "-") + 1)
                            'Numero tal/pag
                        
                            Ampliacion = NumeroTalonPagere & " " & Ampliacion
                        
                        End If
                        
                        If Ampliacion = "" Then
                            Ampliacion = RecuperaValor(RS1!Cliente, 2)
                        Else
                            Ampliacion = "NºDoc: " & Ampliacion
                        End If
                    End If
                    
                End Select
                   
                If NumVtos > 1 Then
                    'TIENE MAS DE UN VTO. No puedo ponerlo en la ampliacion
                    Ampliacion = "Vtos: " & NumVtos
                End If
                
                 'Le concatenamos el texto del concepto para el asiento -ampliacion
                 Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce)) & " "
                 'Para la ampliacion de nºtal + ctrapar NO pongo la ampliacion del concepto
                 If TipoAmpliacion = 5 Then Aux = ""
                 Ampliacion = Aux & Ampliacion
                 If Len(Ampliacion) > 30 Then Ampliacion = Mid(Ampliacion, 1, 30)
                
                 Sql = Sql & "'" & DevNombreSQL(Ampliacion) & "',"
                 
                 
                 If Debe Then
                    Sql = Sql & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                 Else
                    Sql = Sql & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                 End If
             
                'CENTRO DE COSTE
                Sql = Sql & "NULL,"
                
                'SI pone contrapardida
                If PonerContrPartida Then
                   Sql = Sql & "'" & Text3(1).Tag & "',"
                Else
                   Sql = Sql & "NULL,"
                End If
            
             
            Else
                    '----------------------------------------------------
                    'Cierre del asiento con el total contra banco o caja
                    '----------------------------------------------------
                    'codmacta
                    Sql = Sql & Text3(1).Tag & "','"
                     
  
                    PonerContrPartida = False
                    If NumVtos = 1 Then
                        PonerContrPartida = True
                    Else
                        PonerContrPartida = False
                    End If
                       
                    If PonerContrPartida Then
                       Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 2))
                    Else
                       
                       Ampliacion = ""
                    End If
                     
                    Sql = Sql & Ampliacion & "',"
                   
                    
                    If Cobros Then
                        '----------------------------------------------------------------------
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
                                   
                        
                        'COmo el banco o caja, siempre van al reves (Su abono es nuetro pago..)
                        If Not Debe Then
                            Conce = vp.condecli
                            TipoAmpliacion = vp.ampdecli
                        Else
                            Conce = vp.conhacli
                            TipoAmpliacion = vp.amphacli
                        End If
                        
                     Else
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
                        
                        If Not Debe Then
                            Conce = vp.condepro
                            TipoAmpliacion = vp.ampdepro
                        Else
                            Conce = vp.conhapro
                            TipoAmpliacion = vp.amphapro
                        End If
                     End If
                     
                     Sql = Sql & Conce & ","
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
                            If Cobros Then
                                
                            
                          
                                'Veo la el camporefencia de ese talon
                                Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Ampliacion = "numserie = '" & RecuperaValor(RS1!Cliente, 1) & "' AND numfactu = " & RecuperaValor(RS1!Cliente, 2)
                                Ampliacion = Ampliacion & " AND numorden = " & RecuperaValor(RS1!Cliente, 4) & " AND fecfactu "
                                Ampliacion = DevuelveDesdeBD("reftalonpag", "cobros", Ampliacion, Format(RecuperaValor(RS1!Cliente, 3), FormatoFecha), "F")
                                
                                If Ampliacion = "" Then
                                    Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Else
                                    Ampliacion = " NºDoc: " & Ampliacion
                                End If
                                Ampliacion = Ampliacion & " " & DescripcionTransferencia
     
                            Else
                                
                                Ampliacion = NumeroTalonPagere
                                If Ampliacion = "" Then
                                    Ampliacion = RecuperaValor(RS1!Cliente, 2)
                                Else
                                    Ampliacion = "NºDoc: " & Ampliacion
                                End If
                            End If
                          
                            Ampliacion = Ampliacion & " " & DescripcionTransferencia
                            DescripcionTransferencia = ""
                          
                          
                        End Select
                    Else
                        'Ma de un VTO.  Si no
                        If vp.tipoformapago = vbTransferencia Then
                            'SI es transferencia
                            'If TipoAmpliacion = 3 Then Ampliacion = DescripcionTransferencia
                            Ampliacion = DescripcionTransferencia
                        
                        End If
                    End If
                    
                     Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce))
                     Aux = Aux & " "
                     'Para la ampliacion de nºtal + ctrapar NO pongo la ampliacion del concepto
                     If TipoAmpliacion = 5 Then Aux = ""
                     Ampliacion = Trim(Aux & Ampliacion)
                     If Len(Ampliacion) > 30 Then Ampliacion = Mid(Ampliacion, 1, 30)
                    
                     Sql = Sql & "'" & DevNombreSQL(Ampliacion) & "',"
        
                         
                     If Debe Then
                        Sql = Sql & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                     Else
                        Sql = Sql & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                     End If
                 
                     'CENTRO DE COSTE
                     Sql = Sql & "NULL,"
                    
                     'SI pone contrapardida
                     If PonerContrPartida Then
                        Sql = Sql & "'" & RS1!cliprov & "',"
                     Else
                        Sql = Sql & "NULL,"
                     End If
                
                        
                 
            End If
            
            'Trozo comun
            '------------------------
            'IdContab
            Sql = Sql & "'PAGOS',"
            
            'Punteado
            Sql = Sql & "0,"
            
            If Cabecera = 1 And Not VienedeGastos Then
            
                ' nuevos campos de la factura
                'numSerie , numfacpr, FecFactu, numorden, TipForpa, reftalonpag, bancotalonpag
                Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 2), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & ","
                Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & "," & DBSet(Combo1.ItemData(Combo1.ListIndex), "N") & ","
                
                Dim SqlBanco As String
                Dim RsBanco As ADODB.Recordset
                
                SqlBanco = "select reftalonpag, bancotalonpag from tmppagos2 where codusu = " & vUsu.Codigo
                SqlBanco = SqlBanco & " and numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T")
                SqlBanco = SqlBanco & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "T")
                SqlBanco = SqlBanco & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F")
                SqlBanco = SqlBanco & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
                SqlBanco = SqlBanco & " and codmacta = " & DBSet(RS1!cliprov, "T")
        
                Set RsBanco = New ADODB.Recordset
                RsBanco.Open SqlBanco, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsBanco.EOF Then
                    Sql = Sql & DBSet(RsBanco.Fields(0), "T") & "," & DBSet(RsBanco.Fields(1), "T") & ")"
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & ")"
                End If
                Set RsBanco = Nothing
                
            Else
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            End If
             
        End If 'De cabecera menor que 3, es decir : 1y 2
    
    End If
    
    'Ejecutamos si:
    '   Cabecera=0 o 1
    '   Cabecera=2 y impo=0.  Esto sginifica que estamos desbloqueando el apunte e insertandolo para pasarlo a hco
    Debe = True
    If Cabecera = 3 Then Debe = False
    If Debe Then Conn.Execute Sql
    
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
            Sql = "update pagos set imppagad = coalesce(imppagad,0) + " & DBSet(ImporteInterno2, "N") & " "
            Sql = Sql & ", fecultpa = " & DBSet(FechaAsiento, "F")
            Sql = Sql & ", situacion = " & DBSet(Situacion, "N")
            Sql = Sql & " where numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "T")
            Sql = Sql & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
            Sql = Sql & " and codmacta = " & DBSet(RecuperaValor(RS1!Cliente, 5), "T")

            Conn.Execute Sql

        ' en tmppendientes metemos la clave primaria de pagos_realizados y el importe en letra
            Sql = "insert into tmppendientes (codusu,serie_cta,factura,fecha,numorden,codmacta,observa) values ("
            Sql = Sql & vUsu.Codigo & "," & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," 'numserie
            Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 2), "T") & "," 'numfactu
            Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & "," 'fecfactu
            Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & "," 'numorden
            Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 5), "T") & "," 'codmacta
            If ImporteInterno2 < 0 Then
                Sql = Sql & DBSet("menos " & EscribeImporteLetra(ImporteFormateado(CStr(ImporteInterno))), "T") & ") "
            Else
                Sql = Sql & DBSet(EscribeImporteLetra(ImporteFormateado(CStr(ImporteInterno))), "T") & ") "
            End If
            
            Conn.Execute Sql

        End If
    
    End If
    
    
    
    
    '-------------------------------------------------------------------
    'Si es apunte de banco, y hay gastos
    If Cabecera = 2 Then
        'SOOOOLO COBROS
        If Cobros And riesgo > 0 Then
                     
             Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
             Sql = Sql & " timporteH,  ctacontr,codccost, idcontab, punteada) "
             Sql = Sql & "VALUES (" & vp.diaricli & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador & ","
             
             Ampliacion = DevuelveDesdeBD("ctaingreso", "bancos", "codmacta", Text3(1).Tag, "T")
             If Ampliacion = "" Then
                MsgBox "Cta ingreso bancario MAL configurada. Se utilizara la misma del banco", vbExclamation
                Ampliacion = Text3(1).Tag
            End If
            'linea,numdocum,codconce  amconce
            For Conce = 1 To 2
                NumLine = NumLine + 1
                Aux = NumLine & ",'"
                If Conce = 1 Then
                    Aux = Aux & Text3(1).Tag
                Else
                    Aux = Aux & Ampliacion
                End If
                Aux = Aux & "',''," & vp.condecli & ",'" & DevNombreSQL(DevuelveDesdeBD("nomconce", "conceptos", "codconce", vp.condecli)) & "',"
                If Conce = 1 Then
                    Aux = Aux & TransformaComasPuntos(CStr(riesgo)) & ",NULL"
                Else
                    Aux = Aux & "NULL," & TransformaComasPuntos(CStr(riesgo))
                End If
                If Conce = 2 Then
                    Aux = Aux & ",'" & Text3(1).Tag
                Else
                    Aux = Aux & ",'" & Ampliacion
                End If
                Aux = Aux & "',"
                'CC
                If Conce = 1 Then
                    Aux = Aux & "NULL"
                Else
                    If vParam.autocoste Then
                        Ampliacion = DevuelveDesdeBD("codccost", "bancos", "codmacta", Text3(1).Tag, "T")
                        If Ampliacion = "" Then
                            Ampliacion = "NULL"
                        Else
                            Ampliacion = "'" & Ampliacion & "'"
                        End If
                    Else
                        'NO LLEVA ANALITICA
                        Ampliacion = "NULL"
                    End If
                    Aux = Aux & Ampliacion
                End If
                Aux = Aux & ",'PAGOS',0)"
                Aux = Sql & Aux
                Ejecuta Aux
            Next Conce
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
        If miRsAux!ctaprove = IT.Tag Then
            If miRsAux!NumFactu = IT.Text Then
                If miRsAux!FecFactu = IT.SubItems(1) Then
                    If miRsAux!numorden = IT.SubItems(3) Then
                        'ESTE ES
                        BuscarVtoPago = True
                        Fin = True
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
Dim Sql As String
Dim Tipo As Byte

    Sql = "select tipforpa from formapago, pagos  where pagos.codforpa = formapago.codforpa and pagos.numserie = " & DBSet(NumSer, "T")
    Sql = Sql & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F") & " and numorden = " & DBSet(NumOrd, "N")
    Sql = Sql & " and codmacta = " & DBSet(txtCta(5).Text, "T")
    
    Tipo = DevuelveValor(Sql)
    EsTalonOPagare = (CByte(Tipo) = 2 Or CByte(Tipo) = 3)

End Function
