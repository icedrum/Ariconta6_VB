VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInmoVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   Icon            =   "frmInmoVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   8160
      Left            =   0
      TabIndex        =   13
      Top             =   30
      Width           =   8685
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   435
         Left            =   390
         TabIndex        =   38
         Top             =   7530
         Visible         =   0   'False
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame FrameTesor 
         Height          =   2175
         Left            =   240
         TabIndex        =   23
         Top             =   5220
         Width           =   8175
         Begin VB.TextBox Text8 
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
            Index           =   3
            Left            =   1530
            TabIndex        =   29
            Top             =   1710
            Width           =   5220
         End
         Begin VB.TextBox Text8 
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
            Left            =   120
            TabIndex        =   10
            Top             =   1710
            Width           =   1335
         End
         Begin VB.TextBox Text8 
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
            Index           =   1
            Left            =   1530
            TabIndex        =   27
            Top             =   510
            Width           =   5220
         End
         Begin VB.TextBox Text8 
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
            Left            =   120
            TabIndex        =   8
            Top             =   510
            Width           =   1335
         End
         Begin VB.TextBox txtDescta 
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
            Index           =   3
            Left            =   1530
            TabIndex        =   24
            Top             =   1110
            Width           =   5220
         End
         Begin VB.TextBox txtcta 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Left            =   150
            TabIndex        =   28
            Top             =   1440
            Width           =   765
         End
         Begin VB.Image imgTesoreria 
            Height          =   240
            Index           =   2
            Left            =   1590
            Top             =   1470
            Width           =   240
         End
         Begin VB.Image imgTesoreria 
            Height          =   240
            Index           =   1
            Left            =   1590
            Top             =   870
            Width           =   240
         End
         Begin VB.Image imgTesoreria 
            Height          =   240
            Index           =   0
            Left            =   1590
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
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
            Left            =   150
            TabIndex        =   26
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cta cobro"
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
            Left            =   150
            TabIndex        =   25
            Top             =   840
            Width           =   1065
         End
      End
      Begin VB.TextBox Text7 
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
         Index           =   0
         Left            =   1770
         TabIndex        =   22
         Text            =   "Text7"
         Top             =   960
         Width           =   6495
      End
      Begin VB.TextBox Text6 
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
         Left            =   360
         TabIndex        =   0
         Text            =   "Text4"
         Top             =   960
         Width           =   1305
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Aceptar"
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
         Left            =   6090
         TabIndex        =   11
         Top             =   7500
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Cancelar"
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
         Left            =   7290
         TabIndex        =   12
         Top             =   7500
         Width           =   1095
      End
      Begin VB.Frame FrameVenta 
         Height          =   3075
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   8175
         Begin VB.TextBox txtCon 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   0
            Left            =   1470
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   2190
            Width           =   6555
         End
         Begin VB.TextBox txtdpto 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
            Top             =   1080
            Width           =   1275
         End
         Begin VB.TextBox txtDesdpto 
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
            Index           =   0
            Left            =   1470
            TabIndex        =   35
            Top             =   1080
            Width           =   5280
         End
         Begin VB.Frame Frame8 
            Caption         =   "Frame8"
            Height          =   555
            Left            =   240
            TabIndex        =   32
            Top             =   2430
            Visible         =   0   'False
            Width           =   495
            Begin VB.TextBox txtDescta 
               BackColor       =   &H80000018&
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   1440
               TabIndex        =   33
               Top             =   360
               Width           =   3975
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Base Fac."
               Height          =   195
               Index           =   3
               Left            =   600
               TabIndex        =   34
               Top             =   120
               Width           =   360
            End
            Begin VB.Image imgCta 
               Height          =   240
               Index           =   2
               Left            =   1080
               Picture         =   "frmInmoVenta.frx":000C
               Top             =   120
               Width           =   240
            End
         End
         Begin VB.TextBox txtCodCCost 
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
            Left            =   120
            TabIndex        =   6
            Top             =   1710
            Width           =   1275
         End
         Begin VB.TextBox txtNomCcost 
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
            Index           =   2
            Left            =   1470
            TabIndex        =   31
            Top             =   1710
            Width           =   5280
         End
         Begin VB.TextBox txtDescta 
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
            Index           =   1
            Left            =   1470
            TabIndex        =   18
            Top             =   420
            Width           =   5280
         End
         Begin VB.TextBox txtcta 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox Text5 
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
            Left            =   6780
            TabIndex        =   5
            Text            =   "Text5"
            Top             =   420
            Width           =   1245
         End
         Begin VB.Image imgCon 
            Height          =   240
            Index           =   0
            Left            =   1140
            Top             =   2190
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            Left            =   150
            TabIndex        =   37
            Top             =   2160
            Width           =   945
         End
         Begin VB.Image imgDpto 
            Height          =   240
            Index           =   0
            Left            =   1560
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label14 
            Caption         =   "Departamento"
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
            Left            =   120
            TabIndex        =   36
            Top             =   810
            Width           =   1410
         End
         Begin VB.Image imgCCost 
            Height          =   240
            Index           =   2
            Left            =   1560
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Centro coste"
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
            Left            =   150
            TabIndex        =   30
            Top             =   1440
            Width           =   1290
         End
         Begin VB.Label Label13 
            Caption         =   "Importe"
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
            Left            =   6780
            TabIndex        =   20
            Top             =   150
            Width           =   1065
         End
         Begin VB.Label Label14 
            Caption         =   "Cliente"
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
            Left            =   120
            TabIndex        =   19
            Top             =   150
            Width           =   810
         End
         Begin VB.Image imgCta 
            Height          =   240
            Index           =   1
            Left            =   1140
            Top             =   150
            Width           =   240
         End
      End
      Begin VB.TextBox txtDescta 
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
         Index           =   0
         Left            =   3240
         TabIndex        =   15
         Top             =   1680
         Width           =   5040
      End
      Begin VB.TextBox txtcta 
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
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
         Top             =   1680
         Width           =   1395
      End
      Begin VB.TextBox Text4 
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
         Left            =   360
         TabIndex        =   1
         Text            =   "Text4"
         Top             =   1680
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   40
         Top             =   300
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   39
         Top             =   300
         Value           =   -1  'True
         Width           =   1365
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   7980
         TabIndex        =   41
         Top             =   240
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
      Begin VB.Image imgElto 
         Height          =   240
         Index           =   0
         Left            =   1380
         Top             =   690
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Elemento"
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
         Left            =   360
         TabIndex        =   21
         Top             =   690
         Width           =   990
      End
      Begin VB.Label Label14 
         Caption         =   "Cuenta perdidas / beneficios"
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
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Top             =   1410
         Width           =   2970
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   0
         Left            =   4770
         Top             =   1410
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   1410
         Picture         =   "frmInmoVenta.frx":0A0E
         Top             =   1410
         Width           =   240
      End
      Begin VB.Label Label13 
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
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   1410
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmInmoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 511


Public Opcion As Byte
    '0.- Parametros
    '1.- Simular
    '2.- Cálculo amort.
    '3.- Venta/Baja inmovilizado
    '---------------------------
    'los siguiente utilizan el mismo frame, con opciones
    '4.- Listado estadisticas
    '5.- Ficha elementos
    '6.- Entre fechas


    '10.- Deshacer ultima amortizacion

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCt As frmColCtas
Attribute frmCt.VB_VarHelpID = -1
Private WithEvents frmE As frmInmoElto
Attribute frmE.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmAge As frmBasico
Attribute frmAge.VB_VarHelpID = -1
Private WithEvents frmFpa As frmBasico2
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmBa As frmBasico2
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmDpto As frmBasico
Attribute frmDpto.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private frmFCli As frmFacturasCli
Attribute frmFCli.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim Rs As Recordset
Dim cad As String
Dim i As Byte
Dim B As Boolean
Dim Importe As Currency
'
'Desde parametros
Dim Contabiliza As Boolean
Dim UltAmor As Date
Dim DivMes As Integer
Dim ParametrosContabiliza As String
Dim Mc As Contadores

'Tipo de IVA
Dim TipoIva As String
Dim aux2 As String


'Contador para las lineas de apuntes
Dim CONT As Integer

Dim Indice As Integer

Private Sub Command4_Click()
    Unload Me
End Sub

'BAJA / VENTA
Private Sub Command5_Click()
Dim Adelante As Boolean
Dim ContaLinASi As Long
Dim F As Date
'Comprobamos k esta el elemento
    If Text6(0).Text = "" Then
        MsgBox "El elemento no puede estar vacio", vbExclamation
        Exit Sub
    End If
    
    'Comprobamos k la fecha esta puesta
    If Text4(0).Text = "" Then
        MsgBox "Ponga la fecha de baja/venta", vbExclamation
        Text4(0).SetFocus
        Exit Sub
    End If
    If txtCta(0).Text = "" Then
        MsgBox "Introduzca la cuenta de pérdidas/beneficios.", vbExclamation
        txtCta(0).SetFocus
        Exit Sub
    End If
    'Si esta bloqueada
    If EstaLaCuentaBloqueada2(txtCta(0).Text, CDate(Text4(0).Text)) Then
        MsgBox "Cuenta bloqueada: " & txtCta(0).Text, vbExclamation
        Exit Sub
    End If
    
    'Si es venta tenemos k comprobar tb el importe y la cta de cliente
    If Option1(0).Value Then
        'Es venta
        'Comprobamos importe
        If Text5.Text = "" Then
            MsgBox "Introduzca el importe de la venta.", vbExclamation
            Text5.SetFocus
            Exit Sub
        End If
            
        If txtCta(1).Text = "" Then
            MsgBox "Introduzca la cuenta de venta.", vbExclamation
            txtCta(1).SetFocus
            Exit Sub
        End If
        '
        If EstaLaCuentaBloqueada2(txtCta(1).Text, CDate(Text4(0).Text)) Then
            MsgBox "Cuenta bloqueada: " & txtCta(1).Text, vbExclamation
            Exit Sub
        End If
        
        'Si la cuenta necesita CC
        If vParam.autocoste Then
            If HayKHabilitarCentroCoste(txtCta(0).Text) Then
                If Me.txtCodCCost(2).Text = "" Then
                    MsgBox "Debe poner el Centro de coste para la cuenta base", vbExclamation
                    Exit Sub
                End If
            End If
        End If
        'Si tiene tesoreria y es una venta. Entonce introducimos el vencimiento
        If vEmpresa.TieneTesoreria Then
            cad = ""
            If Text8(0).Text = "" Then cad = "Falta forma pago"
            If txtCta(3).Text = "" Then cad = "Falta cta prevista de pago"
            If Text8(2).Text = "" Then cad = "Falta el agente"
            If cad <> "" Then
                cad = "Campos requeridos." & cad
                MsgBox cad, vbExclamation
                Exit Sub
            End If
            
            If EstaLaCuentaBloqueada2(txtCta(3).Text, CDate(Text4(0).Text)) Then
                MsgBox "Cuenta bloqueada: " & txtCta(3).Text, vbExclamation
                Exit Sub
            End If
        End If
    End If


    i = FechaCorrecta2(CDate(Text4(0).Text))
    If i > 1 Then
        If i = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            If i = 3 Then
                MsgBox "Fecha  pertenece a un ejercicio cerrado.", vbExclamation
            Else
                MsgBox "Fecha  pertenece a un ejercicio todavia no abierto", vbExclamation
            End If
        End If
        Exit Sub
    End If



    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then
        MsgBox "Error obteniendo datos parametros amortización", vbExclamation
        Exit Sub
    End If
    
    
    If CDate(Text4(0).Text) <= UltAmor Then
        MsgBox "La fecha es menor que la ultima fecha de amortizacion", vbExclamation
        Exit Sub
    End If
        
        
    'Tenemos que comprobar si la fecha es mayor que la proxima fecha amortizacion
    F = CDate(SugerirFechaNuevo)
    'Debug.Print F
    If CDate(Text4(0).Text) > F Then
        cad = DevuelveDesdeBD("situacio", "inmovele", "codinmov", Me.Text6(0).Text)
        If cad = "1" Then
           'MsgBox "La fecha venta/baja es mayor que la próxima fecha de amortizacion(" & Format(F, "dd/mm/yyyy") & ")" & vbCrLf & vbCrLf & "¿Continuar?", vbInformation
           MsgBox "La fecha venta/baja es mayor que la próxima fecha de amortizacion(" & Format(F, "dd/mm/yyyy") & ")" & vbCrLf & vbCrLf, vbInformation
           Exit Sub
        End If
    End If
    
    If Not ComprobaDatosVentaBajaElemento Then Exit Sub
        
        
    
    aux2 = ""


'Llegados aqui todo bien, con lo cual hacemos ya lo siguiente
'------------------------------------------------------------
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then Exit Sub
    Contabiliza = RecuperaValor(ParametrosContabiliza, 1) = "1"
    'Si contabilizamos hay k conseguir el numero de asiento
    If Contabiliza Then
        Set Mc = New Contadores
        B = (Mc.ConseguirContador("0", CDate(Text4(0).Text) <= vParam.fechafin, True) = 0)
    Else
        B = True
    End If
    
    If B Then
        Screen.MousePointer = vbHourglass
        PreparaBloquear
        Conn.BeginTrans
        Adelante = False
        If Option1(0).Value Then
            i = 0
        Else
            i = 1
        End If
  
  
  
        'Intentamos cargar los datos
        If CargarDatosInmov Then
            'Veremos si ya esta totalmente amortizado o no.
            'Si  lo esta entonces generaremos la cabecera del apunte desde aqui, si no, al realizzar la amortizacion la crea
            If HayQueAmortizar Then
               'Cad = "Select * from sinmov where codinmov=" & Text6.Text & " for update "
                'cont=1  -> Lo inicaliza en el modulo
               B = GeneraCalculoInmovilizado(cad, CByte(i))
               
               'Volvemos a cargar los datos despues de la amortizacion
               If B Then
                    Rs.Close   'Cierro el RS. para volverlo abrir con los datos actualizados de amortiz
                    B = CargarDatosInmov
               End If
               
            Else
                CONT = 1 'Contador para las lineas de asiento
                B = GeneracabeceraApunte(CByte(i))
            End If
            'Contador de asiento
            ContaLinASi = CONT
            If B Then
            
                'Modificacion del 26 de Abril. Si hay venta se vende, pero
                'la cancelacion del elemento se produce siempre
            
                If Option1(0).Value Then
                    'VENTA ---------------------------------------
                    CadenaDesdeOtroForm = ""  'para guardar datos y despues pasarlos a la factura impresa
                    Adelante = VentaElemento
                    
                    'Aqui tb habra que recargar el elemento
                    
                Else
                    Adelante = True
                End If
                
                If Adelante Then
                    'BAJA
                    CONT = ContaLinASi
                    Adelante = CancelarCuentaElemento
                End If
            End If
            Set Rs = Nothing
        End If
        If Adelante Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
            B = False
        End If
        TerminaBloquear
        Pb1.visible = False
        Screen.MousePointer = vbDefault
        If B Then
            If Option1(0).Value Then

                Set frmFCli = New frmFacturasCli
                frmFCli.FACTURA = RecuperaValor(CadenaDesdeOtroForm, 2) & "|" & RecuperaValor(CadenaDesdeOtroForm, 3) & "|" & Year(CDate(RecuperaValor(CadenaDesdeOtroForm, 1))) & "|"
                frmFCli.Show vbModal
                Set frmFCli = Nothing


            Else
                'ha ido bien
                MsgBox "Venta / Baja realizada.", vbInformation
            End If
            Limpiar Me
            Unload Me
        Else
            If Contabiliza Then Mc.DevolverContador "0", Option1(0).Value, Mc.Contador
        End If
    End If
    Set Mc = Nothing
End Sub




Private Function ComprobaDatosVentaBajaElemento() As Boolean

    ComprobaDatosVentaBajaElemento = False
    
    
    'Por si acaso. Si estamos vendiendo, el elmento no puede estar de baja
    If Me.Option1(0).Value Then
        cad = DevuelveDesdeBD("situacio", "inmovele", "codinmov", Me.Text6(0).Text)
        'If Cad = "3" Or Cad = "2" Or Cad = "4" Then  MAYO 2019
        If cad = "3" Or cad = "2" Then
            If cad = "3" Then
                cad = "de baja"
            ElseIf cad = "2" Then
                cad = "vendido"
            Else
                cad = "totalmente amortizado"
            End If
            MsgBox "El elmento esta " & cad, vbExclamation
            Exit Function
        End If
    Else
        cad = DevuelveDesdeBD("situacio", "inmovele", "codinmov", Me.Text6(0).Text)
        
        If cad = "3" Then
            cad = "de baja"
        ElseIf cad = "2" Then
            cad = "vendido"
        Else
            cad = ""
        End If
        If cad <> "" Then
            MsgBox "El elmento esta " & cad, vbExclamation
            Exit Function
        End If
    End If
    
    
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then Exit Function

    
    If Not HazSimulacion("codinmov =" & Text6(0).Text, CDate(Text4(0).Text), 1, Nothing) Then Exit Function
    
    'Ahora, en ztmpsimula tengo los datos del elmento
    Set Rs = New ADODB.Recordset
    
    'NOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
    'NOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
    'NOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
    'NOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
    'cad = "Select valoradq,amortacu,totalamor from Usuarios.zsimulainm where codusu = " & vUsu.Codigo
    cad = "Select valoradq,amortacu,totalamor from tmpsimulainmo where codusu = " & vUsu.Codigo
    
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "      "
    Importe = 0
    aux2 = ""
    B = False
    
    
    
    
    If Not Rs.EOF Then
        B = True
        aux2 = "Importe adq : " & cad & Format(Rs!valoradq, FormatoImporte) & vbCrLf
        aux2 = aux2 & "Amort. acum. : " & cad & Format(Rs!amortacu, FormatoImporte) & vbCrLf
        Importe = Rs!valoradq - Rs!amortacu
        aux2 = aux2 & "Pendiente:     " & cad & Format(Importe, FormatoImporte) & vbCrLf & vbCrLf
        
        
        aux2 = aux2 & "Amort. periodo : " & cad & Format(Rs!totalamor, FormatoImporte) & vbCrLf
        Importe = Importe - Rs!totalamor
        'Si es venta.
        If Option1(0).Value Then
            aux2 = aux2 & "Importe venta : " & cad & Format(CCur(Text5.Text), FormatoImporte) & vbCrLf
            Importe = Importe - CCur(Text5.Text)
        End If

        
    Else
        aux2 = "- Totalmente amortizado" & vbCrLf & vbCrLf
        
        'Si es venta, todo sera ganancias
        If Option1(0).Value Then
            B = True
            Importe = -1 * CCur(Text5.Text)
        End If
    End If
    Rs.Close
    
    
    If B Then
        TipoIva = String(45, "*") & vbCrLf
        If Importe > 0 Then
            'Significa que a la baja o a la venta, falta por amortizar
            'Con lo cual vamos a una cuenta de perdidas
            aux2 = aux2 & "Pérdidas inm.: " & cad & Format(CCur(Importe), FormatoImporte) & vbCrLf & vbCrLf
            If Mid(txtCta(0).Text, 1, 1) <> "6" Then aux2 = aux2 & TipoIva & "Deberia poner una cuenta de PERDIDAS" & vbCrLf & TipoIva
        Else
            aux2 = aux2 & "Ganancias inm.: " & cad & Format(CCur(Abs(Importe)), FormatoImporte) & vbCrLf & vbCrLf
            If Mid(txtCta(0).Text, 1, 1) <> "7" Then aux2 = aux2 & TipoIva & "Deberia poner una cuenta de GANANCIAS" & vbCrLf & TipoIva
        End If
    End If
    
    
    
    TipoIva = ""
    'En importe tengo lo que me faltaria amortizar. Con lo cual. Lo que venda, o de de baja, ira a perdidas
    'o ganancias del grupo 6 o del 7

    If Option1(0).Value Then
        cad = "venta"
    Else
        cad = "baja"
    End If
    cad = "Va a realizar la " & cad & " del "
    
    
    
    aux2 = cad & "elemento:" & vbCrLf & vbCrLf & Text6(0).Text & " - " & Text7(0).Text & vbCrLf & vbCrLf & aux2
    
    

    aux2 = aux2 & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(aux2, vbQuestion + vbYesNo) = vbNo Then Exit Function
    ComprobaDatosVentaBajaElemento = True
End Function




'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False

    Select Case Opcion
    Case 3
        Me.Command4.Cancel = True
    End Select
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmppal.Icon

    Set miTag = New CTag
    Limpiar Me
    PrimeraVez = True
    
    Me.Caption = "Venta / Baja de inmovilizado"
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With

    
    imgElto(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgCta(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgCta(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgCCost(2).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgCon(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    For i = 0 To Me.imgTesoreria.Count - 1
        imgTesoreria(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    Frame3.visible = False
    Select Case Opcion
    Case 3
        Frame3.visible = True
        Me.Width = Frame3.Width + 150
        If vEmpresa.TieneTesoreria Then
            Frame3.Height = 8160 '7080
            
        Else
            Frame3.Height = 5985 '4800
        End If
        Me.Command4.top = Frame3.Height - 480
        Me.Command5.top = Command4.top
        Me.FrameTesor.Enabled = vEmpresa.TieneTesoreria
        Me.FrameTesor.visible = vEmpresa.TieneTesoreria
        Me.Height = Frame3.Height + 500
        
        
        txtCodCCost(2).visible = vParam.autocoste
        Label14(7).visible = vParam.autocoste
        imgCCost(2).visible = vParam.autocoste
        txtNomCcost(2).visible = vParam.autocoste
        'Caption = "Venta / Baja"
    
        Text4(0).Text = Format(Now, "dd/mm/yyyy")
    
    End Select


        
    '0.- Parametros
    '1.- Simular
    '2.- Cálculo amort.
    '3.- Venta/Baja inmovilizado
    '--- los siguiente utilizan el mismo frame, con opciones
    '4.- Listado estadisticas
    '5.- Ficha elementos
'    I = 0
'    If opcion = 0 Or opcion = 2 Then I = 1
'    If I = 0 Then Caption = "Informes"

End Sub

Private Function SugerirFechaNuevo() As String
Dim RC As String
    RC = "tipoamor"
    cad = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", "1", "N", RC)

    If cad <> "" Then
        Me.Tag = cad   'Ultima actualizacion
        Select Case Val(RC)
        Case 2
            'Semestral
            i = 6
            'Siempre es la ultima fecha de mes
        Case 3
            'Trimestral
            i = 3
        Case 4
            'Mensual
            i = 1
        Case Else
            'Anual
            i = 12
        End Select
        RC = PonFecha
    Else
        cad = "01/01/1991"
        RC = Format(Now, "dd/mm/yyyy")
    End If
    'If Simulacion Then
    '     txtFecha.Text = Format(RC, "dd/mm/yyyy")
    'Else
    '     txtFecAmo.Text = Format(RC, "dd/mm/yyyy")
    '     'Dejamos cambiar la fecha, si , y solo si, es administrador
    '     txtFecAmo.Enabled = vUsu.Nivel < 2
        
    'End If
    SugerirFechaNuevo = Format(RC, "dd/mm/yyyy")
    
End Function



Private Function PonFecha() As Date
Dim d As Date
'Dada la fecha en Cad y los meses k tengo k sumar
'Pongo la fecha
d = DateAdd("m", i, CDate(cad))
Select Case Month(d)
Case 2
    If ((Year(d) - 2000) Mod 4) = 0 Then
        i = 29
    Else
        i = 28
    End If
Case 1, 3, 5, 7, 8, 10, 12
    '31
        i = 31
Case Else
    '30
        i = 30
End Select
cad = i & "/" & Month(d) & "/" & Year(d)
PonFecha = CDate(cad)
End Function



Private Sub Form_Unload(Cancel As Integer)
    Set miTag = Nothing
End Sub



Private Sub frmAge_DatoSeleccionado(CadenaSeleccion As String)
    Text8(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text8(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    If i = 1 Then
        'Cuenta bancaria
        txtCta(3).Text = RecuperaValor(CadenaSeleccion, 1)
        txtDesCta(3).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    txtCodCCost(i).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNomCcost(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCt_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(i).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDesCta(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmE_DatoSeleccionado(CadenaSeleccion As String)
    cad = RecuperaValor(CadenaSeleccion, 3)
    If cad = "" Or cad = "1" Or cad = "2" Then
        MsgBox "El elemento esta dado de baja o vendido", vbExclamation
        Exit Sub
    End If
    i = CInt(Me.imgElto(0).Tag)
    Text6(i).Text = RecuperaValor(CadenaSeleccion, 1)
    Text7(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
    Select Case i
    Case 0
    Case 1
    Case 2
    Case 3
        Text4(0).Text = cad
    Case 4, 5
    End Select
End Sub


Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
    If i = 0 Then
        Text8(0).Text = RecuperaValor(CadenaSeleccion, 1)
        Text8(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        Text8(2).Text = RecuperaValor(CadenaSeleccion, 1)
        Text8(3).Text = RecuperaValor(CadenaSeleccion, 2)
    End If

End Sub

Private Sub frmDpto_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtDpto(0).Text = RecuperaValor(CadenaSeleccion, 1)
        txtDesdpto(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     txtCon(Indice).Text = vCampo
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    i = Index
    Select Case Index
    Case 0
    Case 1
    Case 2
    Case 3
        If Text4(0).Text <> "" Then frmF.Fecha = CDate(Text4(0).Text)
        
    Case 4, 5
        'Fechas inofrmes entre fechas
    End Select
    frmF.Show vbModal
    Set frmF = Nothing
End Sub



Private Sub imgCCost_Click(Index As Integer)
    i = Index
    Set frmCC = New frmBasico
    AyudaCC frmCC
    Set frmCC = Nothing
End Sub

Private Sub imgCon_Click(Index As Integer)
    ' observaciones
    Screen.MousePointer = vbDefault
    
    Indice = 0
    
    Set frmZ = New frmZoom
    frmZ.pValor = txtCon(Indice).Text
    frmZ.pModo = 4
    frmZ.Show vbModal
    Set frmZ = Nothing
End Sub

Private Sub imgcta_Click(Index As Integer)
    i = Index
    Set frmCt = New frmColCtas
    frmCt.DatosADevolverBusqueda = "0|1|"
    frmCt.Show vbModal
    Set frmCt = Nothing
End Sub

Private Sub imgDpto_Click(Index As Integer)
    ' departamento
    Indice = 1
    
    Set frmDpto = New frmBasico
    AyudaDepartamentos frmDpto, txtCta(Indice).Text, "codmacta = " & DBSet(txtCta(1).Text, "T")
    Set frmDpto = Nothing
    PonFoco txtDpto(Indice)

End Sub

Private Sub imgElto_Click(Index As Integer)
    Set frmE = New frmInmoElto
    imgElto(0).Tag = Index
    frmE.DatosADevolverBusqueda = "0|1|"
    frmE.Show vbModal
    Set frmE = Nothing
End Sub


Private Sub imgTesoreria_Click(Index As Integer)
    i = Index
    Select Case Index
    Case 0
        'FORMA PAGO
        Set frmFpa = New frmBasico2
        AyudaFPago frmFpa
        Set frmFpa = Nothing
        
    Case 1
        'Cuenta prevista pago
        Set frmBa = New frmBasico2
        AyudaCuentasBancarias frmBa
        Set frmBa = Nothing
    
    
    Case 2
        'Agente
        Set frmAge = New frmBasico
        AyudaAgentes frmAge
        Set frmAge = Nothing

    End Select
End Sub



Private Sub Option1_Click(Index As Integer)
    Me.Option1(0).FontBold = Index = 0
    Me.Option1(1).FontBold = Index = 1
    
    
    Me.FrameVenta.Enabled = (Option1(0).Value)
    Me.FrameTesor.Enabled = vEmpresa.TieneTesoreria And Me.FrameVenta.Enabled
    
    Me.FrameVenta.visible = (Option1(0).Value)
    
    If vEmpresa.TieneTesoreria Then
        FrameTesor.visible = (Option1(0).Value)
    Else
        FrameTesor.visible = False
    End If
    
    
    If Not FrameVenta.Enabled Then
        Text8(0).Text = ""
        Text8(1).Text = ""
        Text8(2).Text = ""
        Text8(3).Text = ""
        txtCta(3).Text = ""
        txtDesCta(3).Text = ""
        txtCta(1).Text = ""
        txtDesCta(1).Text = ""
        txtDpto(0).Text = ""
        txtDesdpto(0).Text = ""
        txtCodCCost(2).Text = ""
        txtNomCcost(2).Text = ""
        txtCon(0).Text = ""
        Text5.Text = ""
    End If
End Sub

Private Sub Text4_LostFocus(Index As Integer)
    If Text4(Index).Text <> "" Then
        If Not EsFechaOK(Text4(Index)) Then
            MsgBox "Fecha incorrecta: " & Text4(Index).Text, vbExclamation
            Text4(Index).Text = ""
        Else
            Text4(Index).Text = Format(Text4(Index).Text, "dd/mm/yyyy")
        End If
    End If
End Sub

Private Sub Text5_GotFocus()
    With Text5
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text5_LostFocus()
    Text5.Text = Trim(Text5.Text)
    If Text5.Text = "" Then Exit Sub
    If Not IsNumeric(Text5.Text) Then
        MsgBox "Importe debe ser numérico: " & Text5.Text, vbExclamation
        Text5.SetFocus
    End If
    DivMes = InStr(1, Text5.Text, ",")
    If DivMes > 0 Then
        'Esta formateado
        Importe = ImporteFormateado(Text5.Text)
    Else
        cad = TransformaPuntosComas(Text5.Text)
        Importe = CCur(cad)
    End If
    Text5.Text = Format(Importe, FormatoImporte)
    
End Sub

Private Sub Text6_LostFocus(Index As Integer)
Dim Sql As String
Dim Rs As ADODB.Recordset

    With Text6(Index)
        .Text = Trim(.Text)
        If .Text = "" Then
            Text7(Index).Text = ""
            Exit Sub
        End If
        If Not IsNumeric(.Text) Then
            MsgBox "Elemento de inmovilizado debe ser numérico: " & .Text, vbExclamation
            .Text = ""
            .SetFocus
            Exit Sub
        End If
        Text6(0).Text = Format(Text6(0).Text, "000000")
        ParametrosContabiliza = "situacio"
        cad = DevuelveDesdeBD("nominmov", "inmovele", "codinmov", .Text, "N", ParametrosContabiliza)
        If cad = "" Then
            MsgBox "elemento de inmovlizado NO encontrado: " & .Text, vbExclamation
        Else
            'Esta comprobacion solo es para la venta/baja
            If Index = 0 Then
                If ParametrosContabiliza = "2" Or ParametrosContabiliza = "3" Or ParametrosContabiliza = "4" Then
                    If ParametrosContabiliza = "4" Then
                       'If Option1(0).Value Then
                       '     MsgBox "Elemento totalmente amortizado", vbExclamation
                       '     Cad = ""
                       ' End If
                    Else
                        MsgBox "El elemento : " & cad & " ya ha sido vendido o dado de baja", vbExclamation
                        cad = ""
                    End If
                    
                Else
                    If vParam.autocoste Then
                        ' traemos si lo tiene el centro de coste del elemento inmovilizado
                        Sql = "select inmovele.codccost,nomccost from inmovele, ccoste where inmovele.codinmov = " & DBSet(Text6(0).Text, "N")
                        Sql = Sql & " and inmovele.codccost = ccoste.codccost"
                        Set Rs = New ADODB.Recordset
                        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If Not Rs.EOF And vParam.autocoste Then
                            txtCodCCost(2).Text = DBLet(Rs.Fields(0).Value, "T")
                            txtNomCcost(2).Text = DBLet(Rs.Fields(1).Value, "T")
                        End If
                        Rs.Close
                    End If
                End If
            End If
        End If
        Text7(Index).Text = cad
        If cad = "" Then
            .Text = ""
            .SetFocus
        End If
    End With
End Sub



Private Sub Text8_GotFocus(Index As Integer)
    PonFoco Text8(Index)
End Sub


Private Sub Text8_LostFocus(Index As Integer)
    Text8(Index).Text = Trim(Text8(Index).Text)
    
    If Index = 0 Then
        If Text8(0).Text = "" Then
            Text8(1).Text = ""
            Exit Sub
        End If
        If Not IsNumeric(Text8(0).Text) Then
            cad = ""
            i = 1
        Else
            cad = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", Text8(0).Text, "N")
            i = 2
        End If
        If cad = "" Then
            cad = "Error en forma pago."
            If i = 1 Then
                cad = cad & " Campo debe ser numérico"
            Else
                cad = cad & " No existe forma pago:" & Text8(0).Text
            End If
            MsgBox cad, vbExclamation
            Text8(0).Text = ""
            Text8(1).Text = ""
        Else
            Text8(1).Text = cad
        End If
    Else
        If Index = 2 Then
            If Text8(2).Text = "" Then
                Text8(3).Text = ""
                Exit Sub
            End If
        
            If Not IsNumeric(Text8(2).Text) Then
                cad = ""
                i = 1
            Else
                cad = DevuelveDesdeBD("nombre", "agentes", "codigo", Text8(2).Text, "N")
                i = 2
            End If
            If cad = "" Then
                cad = "Error en el agente."
                If i = 1 Then
                    cad = cad & " Campo debe ser numérico"
                Else
                    cad = cad & " No existe agente:" & Text8(2).Text
                End If
                MsgBox cad, vbExclamation
                Text8(2).Text = ""
                Text8(3).Text = ""
            Else
                Text8(3).Text = cad
            End If
        
        
        End If
        
    End If
End Sub



Private Function ParaBD(ByRef T As TextBox) As String
If T.Text = "" Then
    ParaBD = "NULL"
Else
    ParaBD = T.Text
End If
End Function

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtCodCCost_GotFocus(Index As Integer)
    With txtCodCCost(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCodCCost_LostFocus(Index As Integer)
    txtCodCCost(Index).Text = Trim(txtCodCCost(Index).Text)
    If txtCodCCost(Index).Text = "" Then
        txtNomCcost(Index).Text = ""
        Exit Sub
    End If
    cad = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtCodCCost(Index).Text, "T")
    If cad = "" Then
        MsgBox "C. coste NO encontrado: " & txtCodCCost(Index).Text, vbExclamation
        txtCodCCost(Index).Text = ""
        txtCodCCost(Index).SetFocus
    End If
    txtNomCcost(Index).Text = cad
End Sub




Private Sub txtCta_GotFocus(Index As Integer)
    With txtCta(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Rs As ADODB.Recordset
Dim Sql As String

With txtCta(Index)
    .Text = Trim(.Text)
    If .Text = "" Then
        txtDesCta(Index).Text = ""
        Exit Sub
    End If
    ParametrosContabiliza = .Text
    If CuentaCorrectaUltimoNivel(ParametrosContabiliza, cad) Then
        .Text = ParametrosContabiliza
        txtDesCta(Index).Text = cad
        If Index = 3 Then
            cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", ParametrosContabiliza, "T")
            If cad = "" Then
                MsgBox "Cuenta no asociada a ningun banco", vbExclamation
                .Text = ""
                txtDesCta(Index).Text = ""
            End If
        Else
            If Index = 1 Then
                txtDpto(0).Text = ""
                txtDesdpto(0).Text = ""
                
                If Me.Option1(0).Value Then
                    ' traemos todos los datos de la cuenta para la venta
                    Sql = "select ctabanco, cuentas.forpa, formapago.nomforpa from cuentas left join formapago on cuentas.forpa = formapago.codforpa "
                    Sql = Sql & " where codmacta = " & DBSet(txtCta(1).Text, "T")
                    
                    Set Rs = New ADODB.Recordset
                    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not Rs.EOF Then
                        If DBLet(Rs!forpa, "N") <> 0 Then Text8(0).Text = Format(DBLet(Rs!forpa, "N"), "000")
                        Text8(1).Text = DBLet(Rs!nomforpa, "T")
                        txtCta(3).Text = DBLet(Rs!CtaBanco, "T")
                        txtDesCta(3).Text = ""
                        If txtCta(3).Text <> "" Then txtDesCta(3).Text = DevuelveValor("select nommacta from cuentas where codmacta = " & DBSet(txtCta(3).Text, "T"))
                    End If
                    Set Rs = Nothing
                Else
                    ' limpiamos los datos
                    Text8(0).Text = ""
                    Text8(1).Text = ""
                    txtCta(3).Text = ""
                    txtDesCta(3).Text = ""
                End If
            End If
        End If
    Else
        MsgBox cad, vbExclamation
        .Text = ""
        txtDesCta(Index).Text = ""
        .SetFocus
    End If

End With
End Sub


'++
Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYCta KeyAscii, 0
            Case 1: KEYCta KeyAscii, 1
            Case 2: KEYCta KeyAscii, 2
            Case 3: KEYTesoreria KeyAscii, 1
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



Private Sub txtCodCCost_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYCCoste KeyAscii, 0
            Case 1: KEYCCoste KeyAscii, 1
            Case 2: KEYCCoste KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFec KeyAscii, 3
            Case 1: KEYFec KeyAscii, 6
            
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYElto KeyAscii, 0
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYTesoreria KeyAscii, 0
            Case 2: KEYTesoreria KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYTesoreria(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgTesoreria_Click (Indice)
End Sub

Private Sub KEYElto(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgElto_Click (Indice)
End Sub

Private Sub KEYCta(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgcta_Click (Indice)
End Sub

Private Sub KEYCCoste(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCCost_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub

Private Sub KEYFec(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub





'++


'TIPO:
'       0.- Venta
'       1.- Baja
'       2.- Calculo de amortizacion
Private Function GeneraCalculoInmovilizado(ByRef SeleccionInmovilizado As String, Tipo As Byte) As Boolean
Dim Codinmov As Long
Dim B As Boolean
On Error GoTo EGen

    GeneraCalculoInmovilizado = False
    If Tipo = 2 Then
        'Para el calculo del amortizado
        Set Rs = New ADODB.Recordset
        Rs.Open SeleccionInmovilizado, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        If Rs.EOF Then
            MsgBox "Ningun registro", vbExclamation
            Rs.Close
            Exit Function
        End If
    End If
    'Vemos cuantos hay
    CONT = 0
    While Not Rs.EOF
        CONT = CONT + 1
        Rs.MoveNext
    Wend
    Rs.MoveFirst
    If CONT > 3 Then Pb1.visible = True
    Pb1.Max = CONT + 1
    Pb1.Value = 0
    
    
    
    'Vemos si contabilizamos
    'Insertamos cabecera del asiento
    If Contabiliza Then GeneracabeceraApunte (Tipo)
    CONT = 1
    While Not Rs.EOF
        Codinmov = Rs!Codinmov
       
        'La fecha depende si estamos calculando normal o estamos vendiendo
       
        cad = Text4(0).Text

      
        B = CalculaAmortizacion(Codinmov, CDate(cad), DivMes, UltAmor, ParametrosContabiliza, Mc.Contador, CONT, Tipo < 2)
        If Not B Then
            Rs.Close
            Exit Function
        End If
        
        'Siguiente
        Pb1.Value = Pb1.Value + 1
        CONT = CONT + 1
        Rs.MoveNext
    Wend
    'Actualizamos la fecha de ultima amortizacion en paraemtros
    If Opcion <> 3 Then
        cad = "UPDATE paramamort SET ultfecha= '" & Format(cad, FormatoFecha)
        cad = cad & "' WHERE codigo=1"
        Conn.Execute cad
        Rs.Close
    Else
        'Estamos dando de baja o vendiendo un inmovilizado. Solo hay uno y hay k situarlo
        'en el primero
        Rs.Requery
        Rs.MoveFirst
    End If
    GeneraCalculoInmovilizado = True
    Exit Function
EGen:
    MuestraError Err.Number
End Function


'Para cancelar elto
Private Sub PonerCadenaLinea()
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce"
    cad = cad & ", ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) VALUES ("
    cad = cad & RecuperaValor(ParametrosContabiliza, 4) & ",'"
    cad = cad & Format(Text4(0).Text, FormatoFecha)
    cad = cad & "'," & Mc.Contador & ","
End Sub


Private Function CargarDatosInmov() As Boolean
On Error GoTo ECar
    CargarDatosInmov = False
    cad = "Select * from inmovele where codinmov =" & Text6(0).Text & " for update"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Error leyendo datos inmovilizado: " & Text6(0).Text, vbExclamation
        Rs.Close
    Else
        CargarDatosInmov = True
    End If
    Exit Function
ECar:
    MuestraError Err.Number, "Cargar datos inmovilizado"
End Function


Private Function CancelarCuentaElemento() As Boolean
Dim Aux As String
Dim NomConce As String

    On Error GoTo ECancelarCuentaElemento
    CancelarCuentaElemento = False
    If Not CargarDatosInmov Then Exit Function
    
    
    If Rs!repartos = 1 Then
     
        MsgBox "Error: REPARTOS incorrectos", vbExclamation
        Exit Function
    Else
        '---------------------------
        'NO tiene reparto de gastos
        '---------------------------
        PonerCadenaLinea
        cad = cad & CONT & ",'" & Rs!codmact3 & "','" & Format(Rs!Codinmov, "000000") & "',"
        cad = cad & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
        '[Monica]15/09/2015: en ampliacion tambien llevamos el nombre de concepto delante
        NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & RecuperaValor(ParametrosContabiliza, 2))
        cad = cad & ",'" & DevNombreSQL(NomConce) & " " & DevNombreSQL(Rs!nominmov)
        Aux = TransformaComasPuntos(CStr(Rs!amortacu))
        cad = cad & "'," & Aux & ",NULL" & ","     'AUX tiene el importe del inmovilizado
        Aux = "NULL"
        If Not IsNull(Rs!CodCCost) Then
            If HayKHabilitarCentroCoste(Rs!codmact3) Then Aux = "'" & Rs!CodCCost & "'"
        End If
        cad = cad & Aux
        cad = cad & ",'" & Rs!codmact1 & "','CONTAI',0)"
        Conn.Execute cad
        CONT = CONT + 1
        
        
        '------------------------------------------------------------------------
        'Cancelacion del elemento
        PonerCadenaLinea
        

        

        Importe = Rs!valoradq - Rs!amortacu

        'La diferencia, si la hubiere se va a las perd/ganan de inmobilizado
        
        If Importe > 0 Then
            PonerCadenaLinea
        
            cad = cad & CONT & ",'" & txtCta(0).Text & "','" & Format(Rs!Codinmov, "000000") & "',"
            cad = cad & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
            '[Monica]15/09/2015: en ampliacion tambien llevamos el nombre de concepto delante
            NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & RecuperaValor(ParametrosContabiliza, 2))
            cad = cad & ",'" & DevNombreSQL(NomConce) & " " & DevNombreSQL(Rs!nominmov)
            Aux = TransformaComasPuntos(CStr(Importe))
            cad = cad & "'," & Aux & ",NULL" & ","
            
            Aux = "NULL"
            If Not IsNull(Rs!CodCCost) Then
                If HayKHabilitarCentroCoste(txtCta(0).Text) Then Aux = "'" & Rs!CodCCost & "'"
            End If
            
            cad = cad & Aux
            cad = cad & ",'" & Rs!codmact3 & "','CONTAI',0)"
            Conn.Execute cad
            CONT = CONT + 1
        End If





        
        PonerCadenaLinea
        cad = cad & CONT & ",'" & Rs!codmact1 & "','" & Format(Rs!Codinmov, "000000") & "',"
        cad = cad & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
        '[Monica]15/09/2015: en ampliacion tambien llevamos el nombre de concepto delante
        NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & RecuperaValor(ParametrosContabiliza, 2))
        cad = cad & ",'" & DevNombreSQL(NomConce) & " " & DevNombreSQL(Rs!nominmov)
        Aux = TransformaComasPuntos(CStr(Rs!valoradq))
        cad = cad & "',NULL," & Aux & ","    'AUX tiene el importe del inmovilizado
        
        Aux = "NULL"
        If Not IsNull(Rs!CodCCost) Then
            If HayKHabilitarCentroCoste(Rs!codmact1) Then Aux = "'" & Rs!CodCCost & "'"
        End If
        
        cad = cad & Aux
        If Importe = 0 Then
            'SI QUE CANCELO en la ctapartida la cuenta
            cad = cad & ",'" & Rs!codmact3 & "'"
        Else
            cad = cad & ",NULL"
        End If
        cad = cad & ",'CONTAI',0)"
        Conn.Execute cad
        CONT = CONT + 1
        
        
        
        
    End If
    
    Rs.Close
    
    
'   Es para baja
    If Option1(1).Value Then
        cad = "UPDATE inmovele SET fecventa = '" & Format(Text4(0).Text, FormatoFecha)
        cad = cad & "', situacio =3  "
        cad = cad & " Where Codinmov = " & Text6(0).Text
        Conn.Execute cad
    End If
    CancelarCuentaElemento = True
ECancelarCuentaElemento:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cancelar cuenta elto."
    Set Rs = Nothing
End Function


Private Function GeneracabeceraApunte(vTipo As Byte) As Boolean
Dim Fecha As Date
On Error GoTo EGeneracabeceraApunte
        GeneracabeceraApunte = False
        cad = "INSERT INTO hcabapu (numdiari, fechaent, numasien,feccreacion,usucreacion,desdeaplicacion, obsdiari) VALUES ("
        cad = cad & RecuperaValor(ParametrosContabiliza, 4) & ",'"
        If Opcion = 3 Then
            Fecha = CDate(Text4(0).Text)
        End If
        cad = cad & Format(Fecha, FormatoFecha)
        cad = cad & "'," & Mc.Contador
        
        cad = cad & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Inmovilizado Amortización " & Fecha & "'"
        
        cad = cad & ",'"
        'Segun sea VENTA, BAJA, o calculo de inmovilizado pondremos una cosa u otra
        Select Case vTipo
        Case 0, 1
            'VENTA
            If txtCon(0).Text = "" Then
                cad = cad & DevNombreSQL(IIf(vTipo = 0, "Venta", "Baja") & " elemento " & Me.Text6(0).Text & " " & Text7(0).Text)
            Else
                cad = cad & txtCon(0).Text
            End If
        Case Else
            cad = cad & "Amortización: " & Fecha
        End Select
        cad = cad & "')"
        Conn.Execute cad
        GeneracabeceraApunte = True
        Exit Function
EGeneracabeceraApunte:
     MuestraError Err.Number, "Genera cabecera Apunte"
     Set Rs = Nothing
End Function


Private Function VentaElemento() As Boolean
Dim Aux As String
Dim RI As Recordset
Dim ImporteTotal As Currency
Dim mZ As Contadores
Dim TotalFactura As Currency
Dim Mc2 As Contadores

Dim ImpIva As Currency
Dim ImpRec As Currency
Dim SqlCtas As String
Dim RsCta As ADODB.Recordset
Dim Numasien2 As Long
Dim NumDiario As Integer
Dim CC As String

        

On Error GoTo EVentaElemento
    VentaElemento = False
        
        TipoIva = DevuelveDesdeBD("codiva", "paramamort", "codigo", "1", "N")
        If TipoIva = "" Then
            MsgBox "Error en el tipo de iva.", vbExclamation
            Exit Function
        End If
        Set RI = New ADODB.Recordset
        RI.Open "Select * from tiposiva WHERE codigiva=" & TipoIva, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If RI.EOF Then
            MsgBox "Error leyendo valores del IVA: " & TipoIva, vbExclamation
            RI.Close
            Exit Function
        End If
        
        'Conseguimos el contador para la factura
        Set mZ = New Contadores
        CONT = FechaCorrecta2(Text4(0).Text)
        If mZ.ConseguirContador("Z", (CONT = 0), True) = 1 Then
            RI.Close
            Rs.Close
            MsgBox "Error contador facuras inmovilizado", vbExclamation
            Set mZ = Nothing
            Exit Function
        End If
        
        'CABECEREA DE FACTURA ---
        'Genereamos la cabecera de factura
        cad = "INSERT INTO factcli (numserie, numfactu, fecfactu, codmacta, anofactu, codforpa, observa,"
        cad = cad & " totbases,totivas,totfaccl, fecliqcl, nommacta, dirdatos, codpobla,"
        cad = cad & " despobla,desprovi,nifdatos,codpais,dpto,codagente) VALUES ("
        'Ejemplo:  'A', 111111112, '2022-02-02', '1', 2002, 'VENTA elto 1',
     
        Aux = DBSet(mZ.TipoContador, "T") & "," & mZ.Contador & "," & DBSet(Text4(0).Text, "F") & ","
        Aux = Aux & DBSet(txtCta(1).Text, "T") & "," & Year(Text4(0).Text) & "," & DBSet(Text8(0).Text, "N") & "," & DBSet(txtCon(0).Text, "T") & ","

        'Pocentaje iva, imponible tal  ytal
        Importe = RI!porceiva
        
        'Para la facutra impresa   Fecha,Numfac,desc, %IVA, total IVA
        CadenaDesdeOtroForm = Text4(0).Text & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & mZ.TipoContador & "|" & Format(mZ.Contador, "0000000") & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & DevNombreSQL(Rs!nominmov) & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(Importe, FormatoImporte) & "|"
        
        '-----------
        ImporteTotal = ImporteFormateado(Text5.Text)
'        AUx = AUx & TransformaComasPuntos(CStr(Importe)) & "," & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & ","
        'Total iva
        Importe = (Importe * ImporteTotal) / 100
        Importe = Round2(Importe, 2)
        Aux = Aux & TransformaComasPuntos(CStr(ImporteTotal)) & "," & TransformaComasPuntos(CStr(Importe)) & ","
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(Importe, FormatoImporte) & "|" 'TOTALIVA
        
        'Total factura
        Importe = Importe + ImporteTotal
        Aux = Aux & TransformaComasPuntos(CStr(Importe))
        TotalFactura = Importe   'Para la tesoreria
        
        
        'Fecha liquidacion
        Aux = Aux & "," & DBSet(Text4(0).Text, "F")
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(Importe, FormatoImporte) & "|" 'TOTAL FAC
        
        ' datos fiscales de la factura
        SqlCtas = "select * from cuentas where codmacta = " & DBSet(txtCta(1).Text, "T")
        Set RsCta = New ADODB.Recordset
        RsCta.Open SqlCtas, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsCta.EOF Then
            Aux = Aux & "," & DBSet(RsCta!Nommacta, "T")
            Aux = Aux & "," & DBSet(RsCta!dirdatos, "T")
            Aux = Aux & "," & DBSet(RsCta!codposta, "T")
            Aux = Aux & "," & DBSet(RsCta!desPobla, "T")
            Aux = Aux & "," & DBSet(RsCta!desProvi, "T")
            Aux = Aux & "," & DBSet(RsCta!nifdatos, "T")
            Aux = Aux & "," & DBSet(RsCta!codpais, "T")
        Else
            Aux = Aux & "," & ValorNulo
            Aux = Aux & "," & ValorNulo
            Aux = Aux & "," & ValorNulo
            Aux = Aux & "," & ValorNulo
            Aux = Aux & "," & ValorNulo
            Aux = Aux & "," & ValorNulo
            Aux = Aux & "," & ValorNulo
        End If
        'departamento
        If txtDpto(0).Text <> "" Then
            Aux = Aux & "," & DBSet(txtDpto(0).Text, "N")
        Else
            Aux = Aux & "," & ValorNulo
        End If
        
        'agente
        If Text8(2).Text <> "" Then
            Aux = Aux & "," & DBSet(Text8(2).Text, "N")
        Else
            Aux = Aux & "," & ValorNulo
        End If
        
        Aux = Aux & ")"
        Conn.Execute cad & Aux
        
        ImporteTotal = ImporteFormateado(Text5.Text)
        CONT = 1
        
        '-------------------------------------------------------
        '-------------------------------------------------------
        'lINEAS . comun
        cad = "INSERT INTO factcli_lineas (numserie, numfactu, fecfactu, anofactu, numlinea, codmacta, baseimpo, codigiva, porciva, porcrec, impoiva, imporec, "
        cad = cad & "aplicret, codccost) VALUES ('" & mZ.TipoContador & "'," & mZ.Contador & "," & DBSet(Text4(0).Text, "F") & "," & Year(Text4(0).Text) & ","
        
        
        'Modificacion de 26 Abril 2004
        '--------------------------------
        'Estos apuntes van en la cancelacion del elemento
        'Amortizacion acumulada DEBE
        
        'Generara 2 lineas de factura
        Importe = Rs!valoradq - Rs!amortacu    'Ahora importe tiene: pendiente de anmortizar
        If Importe > 0 Then
            'Cancelamos la amortizacion del elmento
        
        
            'Elemento
            Aux = TransformaComasPuntos(CStr(Importe))
            Aux = CONT & ",'" & Rs!codmact3 & "'," & Aux & ","
            
            ' iva y porcentaje de iva
            Aux = Aux & TipoIva & "," & DBSet(RI!porceiva, "N") & "," & DBSet(RI!porcerec, "N") & ","
            
            ImpIva = Round2(Importe * DBLet(RI!porceiva) / 100, 2)
            ImpRec = Round2(Importe * DBLet(RI!porcerec) / 100, 2)
            Aux = Aux & DBSet(ImpIva, "N") & "," & DBSet(ImpRec, "N", "S") & ",0,"
            
            ' centro de coste
            CC = "NULL"
            If vParam.autocoste Then
                If HayKHabilitarCentroCoste(Rs!codmact3) Then CC = "'" & Me.txtCodCCost(2).Text & "'"
            End If
            CONT = CONT + 1
            Conn.Execute cad & Aux & CC & ")"
            
            
            'Ahora como he generado este apunte...
            'EL elto queda totalmente amortizado
            Aux = "UPDATE inmovele set amortacu=valoradq Where codinmov =" & Text6(0).Text
            Conn.Execute Aux
            
        Else
            Importe = 0 'Por si acaso estuviera mal. El importe acumulado NO puede ser mayor que el valoradq
        End If
            
        'Ganancias / perdidas de la venta
        Importe = ImporteTotal - Importe
        
        'Elemento
        Aux = TransformaComasPuntos(CStr(Importe))
        Aux = CONT & "," & DBSet(txtCta(0).Text, "T") & "," & Aux & ","
        
        ' iva y porcentaje de iva
        Aux = Aux & TipoIva & "," & DBSet(RI!porceiva, "N") & "," & DBSet(RI!porcerec, "N") & ","
        
        ImpIva = Round2(Importe * DBLet(RI!porceiva) / 100, 2)
        ImpRec = Round2(Importe * DBLet(RI!porcerec) / 100, 2)
        Aux = Aux & DBSet(ImpIva, "N") & "," & DBSet(ImpRec, "N", "S") & ",0,"
        
        
        'centro de coste
        CC = "NULL"
        If vParam.autocoste Then
            If HayKHabilitarCentroCoste(txtCta(0).Text) Then CC = "'" & Me.txtCodCCost(2).Text & "'"

        End If
        CONT = CONT + 1
        Conn.Execute cad & Aux & CC & ")"
        
        
        
        '-------------------------------------------------------
        '-------------------------------------------------------
        ' insertamos en la tabla factcli_totales
        cad = "INSERT INTO factcli_totales (numserie, numfactu, fecfactu, anofactu, numlinea, baseimpo, codigiva, porciva, porcrec, impoiva, imporec) "
        cad = cad & " VALUES ('" & mZ.TipoContador & "'," & mZ.Contador & "," & DBSet(Text4(0).Text, "F") & "," & Year(Text4(0).Text) & ", 1,"
        
        Importe = ImporteTotal
        
        'Elemento
        Aux = TransformaComasPuntos(CStr(Importe))
        Aux = Aux & ","
        
        ' iva y porcentaje de iva
        Aux = Aux & TipoIva & "," & DBSet(RI!porceiva, "N") & "," & DBSet(RI!porcerec, "N") & ","
        
        ImpIva = Round2(Importe * DBLet(RI!porceiva) / 100, 2)
        ImpRec = Round2(Importe * DBLet(RI!porcerec) / 100, 2)
        Aux = Aux & DBSet(ImpIva, "N") & "," & DBSet(ImpRec, "N", "S")
        
        Conn.Execute cad & Aux & ")"
        
        
        Rs.Close
        RI.Close
        
        
        '-------------------------------------------------------
        '-------------------------------------------------------
        ' contabilizar el asiento de la factura
        Set Mc2 = New Contadores
        
        i = FechaCorrecta2(CDate(Text4(0).Text))
        If Mc2.ConseguirContador("0", (i = 0), False) = 0 Then
            Numasien2 = Mc2.Contador
            
            With frmActualizar
                .OpcionActualizar = 6
                'NumAsiento     --> CODIGO FACTURA
                'NumDiari       --> AÑO FACTURA
                'NUmSerie       --> SERIE DE LA FACTURA
                'FechaAsiento   --> Fecha factura
                'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
                .NumFac = mZ.Contador
                .NumDiari = Year(CDate(Text4(0).Text))
                .NUmSerie = mZ.TipoContador
                
                .FechaAsiento = Text4(0).Text
                .FechaAnterior = Text4(0).Text
                
                If NumDiario <= 0 Then NumDiario = vParam.numdiacl
                .DiarioFacturas = NumDiario
                .NumAsiento = Numasien2
                .DentroBeginTrans = True
                
                .Show vbModal
                
                Screen.MousePointer = vbHourglass
                Me.Refresh
            End With
        Else
            Mc2.DevolverContador "0", (i = 0), Numasien2
        End If
        
        
        
        'Si tiene tesoreria genero el cobro
        If vEmpresa.TieneTesoreria Then
            If Not InsertarCobro(mZ) Then Exit Function
        
            'Tambien , para la factura meteremos en la tabla tesoreria comun
            'los datos del vto
             cad = "DELETE from tmptesoreriacomun where codusu =" & vUsu.Codigo
             EjecutaSQL cad
             If aux2 = "" Then aux2 = Text6(0).Text             'NO se que hace con AUX2
             cad = "INSERT INTO tmptesoreriacomun (codusu, codigo, texto1, texto2"
             cad = cad & " ,importe1,  fecha1) values (" & vUsu.Codigo & ",1,'"
             cad = cad & Text8(1).Text & "'," & aux2 & "," & TransformaComasPuntos(CStr(Importe)) & ",'"
             cad = cad & Format(Text4(0).Text, FormatoFecha) & "')"
             EjecutaSQL cad
        End If
        
        'Ahora hay k poner el elemento a vendido, con el importe de venta y la fecha de venta
        cad = "UPDATE inmovele SET fecventa = '" & Format(Text4(0).Text, FormatoFecha)
        cad = cad & "', situacio =2 , impventa="
        cad = cad & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & " WHERE codinmov =" & Text6(0).Text
        Conn.Execute cad
        
        VentaElemento = True
EVentaElemento:
    If Err.Number <> 0 Then MuestraError Err.Number, "Venta Elemento"
    Set Rs = Nothing
    Set RI = Nothing
End Function


Private Function InsertarCobro(ByRef mZ As Contadores) As Boolean
Dim Sql As String
Dim RsCta As ADODB.Recordset
Dim CadInsert As String
Dim CadValues As String
Dim textCSB As String

    On Error GoTo eInsertarCobro
    
    InsertarCobro = False
    
    
    Sql = "delete from tmpcobros where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute Sql
    
    B = CargarCobrosTemporal(Text8(0).Text, Text4(0).Text, CCur(Text5.Text))
    
    If B Then
        
        CadInsert = "insert into cobros (numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci," & _
                    "ctabanc1,fecultco,impcobro,emitdocum,recedocu,contdocu," & _
                    "text33csb,text41csb,ultimareclamacion,agente,departamento,transfer," & _
                    "nomclien,domclien,pobclien,cpclien,proclien,iban,codusu) values "
        CadValues = ""
        
        
        Sql = "select * from cuentas where codmacta = " & DBSet(txtCta(1).Text, "T")
        Set RsCta = New ADODB.Recordset
        RsCta.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsCta.EOF Then
        
            Sql = "select * from tmpcobros where codusu = " & DBSet(vUsu.Codigo, "N")
        
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            i = 0
            
            While Not Rs.EOF
                i = i + 1
                
                Sql = DBSet(mZ.TipoContador, "T") & "," & DBSet(mZ.Contador, "N") & "," & DBSet(Text4(0).Text, "F") & "," & DBSet(i, "N") & ","
                Sql = Sql & DBSet(txtCta(1).Text, "T") & "," & DBSet(Text8(0).Text, "N") & "," & DBSet(Rs!FecVenci, "F") & "," & DBSet(Rs!ImpVenci, "N") & ","
                Sql = Sql & DBSet(txtCta(3).Text, "T", "S") & ","
                
                Sql = Sql & ValorNulo & "," & ValorNulo & ","
                
                Sql = Sql & "0,0,0,"
                
                textCSB = "Factura " & Trim(mZ.TipoContador) & "-" & Trim(mZ.Contador) & " de Fecha " & Text4(0).Text
                
                Sql = Sql & DBSet(textCSB, "T") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Text8(2).Text, "N", "S") & "," & DBSet(txtDpto(0).Text, "N", "S") & "," & ValorNulo & ","
                Sql = Sql & DBSet(RsCta!Nommacta, "T", "S") & "," & DBSet(RsCta!dirdatos, "T", "S") & "," & DBSet(RsCta!desPobla, "T", "S") & "," & DBSet(RsCta!codposta, "T", "S") & ","
                Sql = Sql & DBSet(RsCta!desProvi, "T", "S") & "," & DBSet(RsCta!IBAN, "T", "S")
                
                ' el codusu
                Sql = Sql & "," & DBSet(vUsu.Id, "N")
                
                
                CadValues = CadValues & "(" & Sql & "),"
            
                Rs.MoveNext
            Wend
        
            If CadValues <> "" Then
                CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
                Conn.Execute CadInsert & CadValues
            End If
        
            Set Rs = Nothing
        
        End If
        Set RsCta = Nothing
    End If
    
    InsertarCobro = B
    Exit Function
eInsertarCobro:
    MuestraError Err.Number, "Insertar Cobro", Err.Description
End Function


'Devuelve TRUE si esta activo
Private Function HayQueAmortizar() As Boolean
HayQueAmortizar = False
cad = DevuelveDesdeBD("situacio", "inmovele", "codinmov", Text6(0).Text, "N")
If cad <> "" Then
    If cad = "1" Then HayQueAmortizar = True
End If
End Function




'///////////////////////////////////////////////////////////
'
'   Este procedimento utilizad dos tablas ya creadas que son
'   en USUARIOS  z347 y z347carta
Private Sub EmiteFacturaVentaInmmovilizado()

On Error GoTo EEmiteFacturaVentaInmmovilizado
    cad = "DELETE FROM usuarios.z347  WHERE codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'Los datos del encabezado
    CargaEncabezadoCarta 1
    DivMes = 0
    
    
    Set Rs = New ADODB.Recordset
    cad = "Select * from Cuentas where codmacta='" & txtCta(1).Text & "'"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "INSERT INTO usuarios.z347 (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla) VALUES (" & vUsu.Codigo
    If Rs.EOF Then
        cad = cad & "," & DivMes & ",'nif'," & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & ",'" & DevNombreSQL(txtDesCta(1).Text) & "','Direccion','codpos','Poblacion')"
    Else
        cad = cad & "," & DivMes & ",'"
        cad = cad & DBLet(Rs!nifdatos) & "'," & TransformaComasPuntos(ImporteSinFormato(Text5.Text)) & ",'" & DevNombreSQL(txtDesCta(1).Text) & "','"
        cad = cad & DevNombreSQL(DBLet(Rs!dirdatos)) & "','" & DBLet(Rs!codposta) & "','"
        cad = cad & DevNombreSQL(DBLet(Rs!desPobla)) & "')"
    End If
    Conn.Execute cad
    
    Exit Sub
EEmiteFacturaVentaInmmovilizado:
    MuestraError Err.Number, "Generando factura"
End Sub


Private Sub txtDpto_GotFocus(Index As Integer)
Dim Sql As String
Dim i As Integer
    With txtDpto(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

    If txtCta(1).Text <> "" Then
        Sql = "select dpto, descripcion from departamentos where codmacta = " & DBSet(txtCta(1).Text, "T")
        i = TotalRegistrosConsulta(Sql)
        Select Case i
            Case 0
                PonFoco Text5
            Case 1
            Case Else
                PonFoco txtDpto(0)
        End Select
    End If

End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    If txtDpto(Index).Text = "" Then
        txtDesdpto(Index).Text = ""
        Exit Sub
    End If
    cad = DevuelveValor("select descripcion from departamentos where codmacta = " & DBSet(txtCta(1).Text, "T") & " and dpto = " & DBSet(txtDpto(Index).Text, "N"))
    If cad = "" Then
        MsgBox "Departamento NO encontrado: " & txtDpto(Index).Text, vbExclamation
        txtDpto(Index).Text = ""
        txtDpto(Index).SetFocus
    End If
    txtDesdpto(Index).Text = cad

End Sub
