VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFVARReimpresion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpresión de Facturas Varias"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12855
   Icon            =   "frmFVARReimpresion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8190
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   7950
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   12405
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
         Height          =   6855
         Left            =   8820
         TabIndex        =   30
         Top             =   270
         Width           =   3285
         Begin VB.CheckBox Check1 
            Caption         =   "Duplicado"
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
            Left            =   270
            TabIndex        =   8
            Top             =   540
            Width           =   2850
         End
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
         Left            =   225
         TabIndex        =   10
         Top             =   7290
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
         Left            =   225
         TabIndex        =   19
         Top             =   4500
         Width           =   8490
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
            TabIndex        =   29
            Top             =   720
            Value           =   -1  'True
            Width           =   1335
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
            TabIndex        =   28
            Top             =   1200
            Width           =   1515
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
            TabIndex        =   26
            Top             =   2160
            Width           =   975
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
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   720
            Width           =   4920
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
            TabIndex        =   24
            Top             =   1200
            Width           =   6240
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
            TabIndex        =   23
            Top             =   1680
            Width           =   6240
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   8025
            TabIndex        =   22
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   8025
            TabIndex        =   21
            Top             =   1680
            Width           =   255
         End
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
            Left            =   6765
            TabIndex        =   20
            Top             =   720
            Width           =   1515
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
         Height          =   4110
         Left            =   225
         TabIndex        =   13
         Top             =   270
         Width           =   8445
         Begin VB.TextBox txtCodigo 
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   3
            Top             =   2430
            Width           =   1350
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   1215
            MaxLength       =   10
            TabIndex        =   2
            Top             =   2025
            Width           =   1350
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "Text5"
            Top             =   2430
            Width           =   5655
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "Text5"
            Top             =   2040
            Width           =   5655
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   6
            Left            =   4530
            MaxLength       =   7
            TabIndex        =   6
            Tag             =   "Nº factura|N|S|0||factcli|numfactu|0000000|S|"
            Top             =   3195
            Width           =   1350
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   7
            Left            =   4530
            MaxLength       =   7
            TabIndex        =   7
            Tag             =   "Nº factura|N|S|0||factcli|numfactu|0000000|S|"
            Top             =   3570
            Width           =   1350
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   4
            Top             =   3150
            Width           =   1350
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   5
            Left            =   1215
            MaxLength       =   10
            TabIndex        =   5
            Top             =   3525
            Width           =   1350
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "Text5"
            Top             =   1230
            Width           =   6150
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "Text5"
            Top             =   855
            Width           =   6150
         End
         Begin VB.TextBox txtCodigo 
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
            MaxLength       =   6
            TabIndex        =   1
            Top             =   1230
            Width           =   830
         End
         Begin VB.TextBox txtCodigo 
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
            MaxLength       =   6
            TabIndex        =   0
            Top             =   855
            Width           =   830
         End
         Begin VB.Image ImgFec 
            Height          =   240
            Index           =   0
            Left            =   900
            Picture         =   "frmFVARReimpresion.frx":000C
            Top             =   3150
            Width           =   240
         End
         Begin VB.Image ImgFec 
            Height          =   240
            Index           =   1
            Left            =   900
            Picture         =   "frmFVARReimpresion.frx":0097
            Top             =   3540
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Cuenta Contable"
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
            Height          =   330
            Index           =   2
            Left            =   180
            TabIndex        =   41
            Top             =   1665
            Width           =   3120
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   5
            Left            =   225
            TabIndex        =   40
            Top             =   2025
            Width           =   600
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   39
            Top             =   2400
            Width           =   645
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   915
            MouseIcon       =   "frmFVARReimpresion.frx":0122
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cuenta"
            Top             =   2430
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   900
            MouseIcon       =   "frmFVARReimpresion.frx":0274
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cuenta"
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Factura"
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
            Height          =   375
            Index           =   1
            Left            =   3510
            TabIndex        =   36
            Top             =   2880
            Width           =   2490
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   3
            Left            =   3555
            TabIndex        =   35
            Top             =   3195
            Width           =   690
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   2
            Left            =   3555
            TabIndex        =   34
            Top             =   3570
            Width           =   645
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Factura"
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
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   33
            Top             =   2835
            Width           =   3120
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   32
            Top             =   3150
            Width           =   690
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   31
            Top             =   3525
            Width           =   645
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            MouseIcon       =   "frmFVARReimpresion.frx":03C6
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar contador"
            Top             =   1230
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   915
            MouseIcon       =   "frmFVARReimpresion.frx":0518
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar contador"
            Top             =   855
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   12
            Left            =   225
            TabIndex        =   18
            Top             =   1230
            Width           =   645
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   13
            Left            =   225
            TabIndex        =   17
            Top             =   855
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Serie"
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
            Height          =   375
            Index           =   8
            Left            =   180
            TabIndex        =   14
            Top             =   540
            Width           =   3120
         End
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   10800
         TabIndex        =   11
         Top             =   7290
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
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
         Index           =   1
         Left            =   9315
         TabIndex        =   9
         Top             =   7290
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmFVARReimpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Integer
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmFVARCon As frmFVARConceptos 'Conceptos
Attribute frmFVARCon.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas 'Cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico 'Cuentas contables
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Dim IndCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim Sql As String

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

Dim cDesde As String
Dim cHasta As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("fvarfactura.numserie", "SER", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHSerie=""") Then Exit Function
    If Not PonerDesdeHasta("fvarfactura.codmacta", "CTA", Me.txtCodigo(2), Me.txtNombre(2), Me.txtCodigo(3), Me.txtNombre(3), "pDHCta=""") Then Exit Function
    If Not PonerDesdeHasta("fvarfactura.fecfactu", "F", Me.txtCodigo(4), Me.txtCodigo(4), Me.txtCodigo(5), Me.txtCodigo(5), "pDHFecha=""") Then Exit Function
    If Not PonerDesdeHasta("fvarfactura.numfactu", "FRA", Me.txtCodigo(6), Me.txtCodigo(6), Me.txtCodigo(7), Me.txtCodigo(7), "pDHNumfactu=""") Then Exit Function
    
    cadParam = cadParam & "pDuplicado=" & Check1(0).Value & "|"
    numParam = numParam + 1
    
    
    MontaSQL = True
    
End Function






Private Sub cmdAccion_Click(Index As Integer)

    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("fvarfactura", cadselect) Then Exit Sub
    
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

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
        
    vMostrarTree = False
    conSubRPT = False
        
        
    indRPT = "421-00" '"facturas varias"

    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 1
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Sub AccionesCSV()
Dim Sql As String

    'Monto el SQL
    Sql = "Select codconce AS Código, nomconce as Descripcion, fvarconceptos.codmacta as CtaContable, nommacta as Descripcion,  "
    Sql = Sql & " tipoiva TipoIva, porceiva PorIva, codccost CentroCoste"
    Sql = Sql & " from (fvarconceptos inner join cuentas on fvarconceptos.codmacta = cuentas.codmacta) inner join tiposiva on fvarconceptos.tipoiva = tiposiva.codigiva "
    
    If cadselect <> "" Then Sql = Sql & " WHERE " & cadselect
    
    
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCanCtasContables_Click()
    RaiseEvent DatoSeleccionado("")
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 0
                PonFoco txtCodigo(0)
                
                txtcodigo_LostFocus (0)
                txtcodigo_LostFocus (1)
                txtcodigo_LostFocus (2)
                txtcodigo_LostFocus (3)
            Case 1
            
        
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
    For H = 0 To 3
        Me.imgBuscar(H).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next H
     
     
    FrameCobros.visible = False
    
    '###Descomentar
'    CommitConexion
    Select Case OpcionListado
        Case 0 ' reimpresion de facturas
            FrameCobrosVisible True, H, W
            indFrame = 5
            tabla = "fvarfactura"
                
            PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
            ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
        
            ' la salida a archivo csv la deshabilitamos
            Me.optTipoSal(1).Enabled = False
            Me.txtTipoSalida(1).Enabled = False
            Me.PushButton2(0).Enabled = False
        
    End Select
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub


Private Sub imgFec_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
        IndCodigo = Index + 4
    
        'FECHA
        Set frmC = New frmCal
        frmC.Fecha = Now
        If txtCodigo(IndCodigo).Text <> "" Then frmC.Fecha = CDate(txtCodigo(IndCodigo).Text)
        frmC.Show vbModal
        Set frmC = Nothing
        PonFoco txtCodigo(Index)
        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub frmFVARCon_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Contadores
            IndCodigo = Index
        
            Set frmConta = New frmBasico
            AyudaContadores frmConta, txtCodigo(Index), "tiporegi REGEXP '^[0-9]+$' = 0"
            Set frmConta = Nothing
    
            PonFoco Me.txtCodigo(Index)
        
        Case 2, 3 ' cuenta contable
            IndCodigo = Index
            Sql = ""
            AbiertoOtroFormEnListado = True
            Set frmCtas = New frmColCtas
            frmCtas.DatosADevolverBusqueda = True
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            If Sql <> "" Then
                Me.txtCodigo(Index).Text = RecuperaValor(Sql, 1)
                Me.txtNombre(Index).Text = RecuperaValor(Sql, 2)
            Else
                QuitarPulsacionMas Me.txtCodigo(Index)
            End If
            
            PonFoco Me.txtCodigo(Index)
            AbiertoOtroFormEnListado = False
        
    End Select
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
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

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtcodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtcodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'concepto desde
            Case 1: KEYBusqueda KeyAscii, 1 'concepto hasta
            Case 2: KEYBusqueda KeyAscii, 2 'cuenta desde
            Case 3: KEYBusqueda KeyAscii, 3 'cuenta hasta
            
            Case 4: KEYFecha KeyAscii, 0 'fecha desde
            Case 5: KEYFecha KeyAscii, 1 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub txtcodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim b As Boolean
Dim Hasta As Integer

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If txtCodigo(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 0, 1 ' tipos de movimiento
            txtCodigo(Index).Text = UCase(Trim(txtCodigo(Index).Text))
            
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "contadores", "nomregis", "tiporegi", "T")
            
        Case 2, 3 ' cuentas
            If Not IsNumeric(txtCodigo(Index).Text) Then
                If InStr(1, txtCodigo(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCodigo(Index).Text, vbExclamation
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
                Exit Sub
            End If
            Cta = (txtCodigo(Index).Text)
                                    '********
            b = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
            If b = 0 Then
                MsgBox "NO existe la cuenta: " & txtCodigo(Index).Text, vbExclamation
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
            Else
                txtCodigo(Index).Text = Cta
                txtNombre(Index).Text = Sql
                If b = 1 Then
                    txtNombre(Index).Tag = ""
                Else
                    txtNombre(Index).Tag = Sql
                End If
                Hasta = -1
                If Index = 2 Then
                    Hasta = 3
                End If
                    
                If Hasta >= 0 Then
                    txtCodigo(Hasta).Text = txtCodigo(Index).Text
                    txtNombre(Hasta).Text = txtNombre(Index).Text
                End If
            End If
            
        Case 4, 5 ' fechas
            PonerFormatoFecha txtCodigo(Index)
        
    End Select
  
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.top = -90
        Me.FrameCobros.Left = 0
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub


Private Sub AbrirFrmFVARConceptos(Indice As Integer)
    IndCodigo = Indice
    Set frmFVARCon = New frmFVARConceptos
    frmFVARCon.DatosADevolverBusqueda = "0|1|"
    frmFVARCon.DeConsulta = True
    frmFVARCon.CodigoActual = txtCodigo(IndCodigo)
    frmFVARCon.Show vbModal
    Set frmFVARCon = Nothing
End Sub
 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
    Unload Me
End Sub

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




