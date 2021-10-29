VERSION 5.00
Begin VB.Form frmTESGastosFijos2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTESGastosFijos2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameContabilizarGasto 
      Height          =   4485
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   10395
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   19
         Left            =   120
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   1080
         Width           =   1305
      End
      Begin VB.TextBox txtDiario 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   555
      End
      Begin VB.TextBox txtNDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   720
         TabIndex        =   45
         Top             =   1800
         Width           =   3075
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   3
         Left            =   1680
         TabIndex        =   48
         Text            =   "Text4"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCuentas 
         Height          =   360
         Index           =   5
         Left            =   120
         TabIndex        =   56
         Text            =   "Text9"
         Top             =   2520
         Width           =   1350
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   5
         Left            =   1560
         TabIndex        =   44
         Text            =   "Text9"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   19
         Left            =   8550
         TabIndex        =   63
         Top             =   3750
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   9
         Left            =   4800
         MaxLength       =   35
         TabIndex        =   58
         Text            =   "Text9"
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox txtConcepto 
         Height          =   360
         Index           =   0
         Left            =   3840
         TabIndex        =   54
         Top             =   1800
         Width           =   885
      End
      Begin VB.TextBox txtNConcepto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   4800
         TabIndex        =   43
         Top             =   1800
         Width           =   5175
      End
      Begin VB.TextBox txtCuentas 
         Height          =   360
         Index           =   4
         Left            =   3360
         TabIndex        =   50
         Text            =   "Text2"
         Top             =   1080
         Width           =   1350
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   4
         Left            =   4830
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   1080
         Width           =   5145
      End
      Begin VB.TextBox txtCC 
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   60
         Text            =   "Text10"
         Top             =   3330
         Width           =   795
      End
      Begin VB.TextBox txtNCC 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   990
         TabIndex        =   41
         Text            =   "Text9"
         Top             =   3330
         Width           =   4695
      End
      Begin VB.CommandButton cmdContabiliGasto 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   7080
         TabIndex        =   62
         Top             =   3750
         Width           =   1335
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   5
         Left            =   150
         TabIndex        =   64
         Top             =   390
         Width           =   2280
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   61
         Top             =   810
         Width           =   750
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   19
         Left            =   1140
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Diario"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   59
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   57
         Top             =   2250
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación"
         Height          =   285
         Index           =   13
         Left            =   4800
         TabIndex        =   55
         Top             =   2250
         Width           =   1095
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   750
         Top             =   1530
         Width           =   270
      End
      Begin VB.Label Label7 
         Caption         =   "Concepto"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   53
         Top             =   1530
         Width           =   990
      End
      Begin VB.Label Label7 
         Caption         =   "Importe"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   51
         Top             =   810
         Width           =   960
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   4830
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Cuenta banco"
         Height          =   255
         Left            =   3360
         TabIndex        =   49
         Top             =   810
         Width           =   1425
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   4890
         Top             =   1530
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Centro de coste"
         Height          =   225
         Index           =   14
         Left            =   150
         TabIndex        =   47
         Top             =   3060
         Width           =   1725
      End
      Begin VB.Image imgCC 
         Height          =   240
         Index           =   0
         Left            =   1890
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   1230
         Top             =   2250
         Width           =   240
      End
   End
   Begin VB.Frame FrameModGastoFijo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10395
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   7530
         TabIndex        =   27
         Top             =   3060
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   8820
         TabIndex        =   26
         Top             =   3060
         Width           =   1095
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   3
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2190
         Width           =   6015
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   2
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1650
         Width           =   6015
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   3
         Left            =   2490
         TabIndex        =   23
         Top             =   2190
         Width           =   1275
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   2
         Left            =   2460
         TabIndex        =   22
         Top             =   1650
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   8
         Left            =   2460
         MaxLength       =   30
         TabIndex        =   21
         Tag             =   "Descripción|T|N|||remesas|descripción|||"
         Top             =   1020
         Width           =   7365
      End
      Begin VB.Label Label3 
         Caption         =   "Modificacion Gasto Fijo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   270
         TabIndex        =   31
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label Label3 
         Caption         =   "Contrapartida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   30
         Top             =   2250
         Width           =   1770
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   3
         Left            =   2100
         Top             =   2220
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   2
         Left            =   2100
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   29
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Prevista"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   28
         Top             =   1650
         Width           =   1860
      End
   End
   Begin VB.Frame FrameAltaModLineaGasto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   0
         Left            =   2550
         TabIndex        =   2
         Text            =   "99/99/9999"
         Top             =   1140
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   4
         Left            =   2580
         TabIndex        =   3
         Text            =   "99/99/9999"
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5280
         TabIndex        =   1
         Top             =   2910
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   5
         Top             =   2910
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Linea a modificar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   5460
      End
      Begin VB.Label lblFecha1 
         Height          =   255
         Index           =   1
         Left            =   2580
         TabIndex        =   7
         Top             =   3990
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de gasto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2190
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1770
         Width           =   1935
      End
   End
   Begin VB.Frame FrameAltaGastoFijo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10395
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "imgConcepto"
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   5
         Left            =   2670
         TabIndex        =   39
         Top             =   3300
         Width           =   1365
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   5
         Left            =   2670
         TabIndex        =   37
         Top             =   2130
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   1
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   34
         Tag             =   "Descripción|T|N|||remesas|descripción|||"
         Top             =   450
         Width           =   6045
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   2670
         TabIndex        =   35
         Top             =   1020
         Width           =   1275
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   1
         Left            =   2700
         TabIndex        =   36
         Top             =   1560
         Width           =   1275
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1020
         Width           =   6015
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   1
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1560
         Width           =   5985
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8820
         TabIndex        =   11
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptarAltaCab 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7530
         TabIndex        =   10
         Top             =   3840
         Width           =   1155
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         Index           =   0
         ItemData        =   "frmTESGastosFijos2.frx":000C
         Left            =   2670
         List            =   "frmTESGastosFijos2.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2730
         Width           =   2100
      End
      Begin VB.Label Label3 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   4
         Left            =   210
         TabIndex        =   33
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   19
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Periodicidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   2310
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de gasto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   2190
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Prevista"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   6
         Left            =   210
         TabIndex        =   16
         Top             =   1020
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   8
         Left            =   2730
         TabIndex        =   15
         Top             =   450
         Width           =   1380
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   2310
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   2310
         Top             =   1590
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Contrapartida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   11
         Left            =   210
         TabIndex        =   14
         Top             =   1620
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmTESGastosFijos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '1.- Alta cabecera gasto fijo
    '2.- Modificacion cabecera gasto fijo
    '3.- Alta linea gasto fijo
    '4.- Modificacion linea gasto fijo
    '5.- Contabilizacion del gasto
    
    
Public Parametros As String
    '1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

'recepcion de talon/pagare
Public Importe As Currency
Public Codigo As String
Public Tipo As String
Public FecCobro As String
Public FecVenci As String
Public Banco As String
Public Referencia As String

Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico 'frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmDi As frmTiposDiario
Attribute frmDi.VB_VarHelpID = -1

Private PrimeraVez As Boolean

Dim i As Integer
Dim SQL As String
Dim Rs As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim Ok As Integer
Dim vWhere As String

Dim CampoOrden As String
Dim Orden As Boolean
Dim Indice As Integer


Private Function DatosOK() As Boolean
Dim B1 As Byte
Dim Cta As String

    DatosOK = False
    
    Select Case Opcion
        Case 1 ' insertar cabecera
            If Text1(1).Text = "" Then
                MsgBox "Debe introducir el concepto.", vbExclamation
                PonFoco Text1(1)
                Exit Function
            End If
            If txtCuentas(0).Text = "" Then
                MsgBox "Debe introducir una cuenta prevista. Reintroduzca.", vbExclamation
                PonFoco txtCuentas(0)
                Exit Function
            Else
                Cta = (txtCuentas(0).Text)
                                    '********
                B1 = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
                
                If B1 = 2 Then
                    'Si que existe la cuenta
                    SQL = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Cta, "T")
                    If SQL = "" Then
                        SQL = "No pertence a un banco "
                        B1 = 0
                    End If
                End If
                
                If B1 <> 2 Then
                    MsgBox SQL & txtCuentas(0).Text, vbExclamation
                    PonFoco txtCuentas(0)
                    Exit Function
                End If
                
                
                
            End If
            
            
            If txtCuentas(1).Text <> "" Then
                Cta = (txtCuentas(1).Text)
                                    '********
                B1 = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
            Else
                B1 = 0
                SQL = "Debe poner una cuenta"
            End If
            If B1 <> 2 Then
                    MsgBox SQL & txtCuentas(1).Text, vbExclamation
                    PonFoco txtCuentas(1)
                    Exit Function
            End If
            
            
            If txtFecha(5).Text = "" Then
                MsgBox "Debe introducir una fecha de gasto.", vbExclamation
                PonFoco txtFecha(5)
                Exit Function
            End If
            If Combo1(0).ListIndex = -1 Then
                MsgBox "Debe introducir una periodicidad", vbExclamation
                PonerFocoCmb Combo1(0)
                Exit Function
            End If
    End Select
    
    DatosOK = True
    
End Function






Private Sub cmdAceptar_Click(Index As Integer)
Dim B As Boolean
    If Not DatosOK Then Exit Sub
    
    B = False
    Select Case Index
        Case 0 ' modificar cabecera
            B = ModificarGasto
        Case 2
            B = InsertarModificarLinea
    End Select
    
    If B Then
'        MsgBox "Proceso realizado correctamente", vbExclamation
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If

End Sub

Private Sub cmdAceptarAltaCab_Click()


    If Not DatosOK Then Exit Sub
    
    If GenerarGasto Then
        MsgBox "Proceso realizado correctamente", vbInformation
        CadenaDesdeOtroForm = "OK"
        Unload Me
    End If


End Sub

Private Function InsertarModificarLinea() As Boolean
Dim SQL As String

    On Error GoTo eModificarGasto
    
    InsertarModificarLinea = False
    If Text1(4).Text = "" Then
        MsgBox "Introduzaca el importe", vbExclamation
        
    Else
    If Opcion = 3 Then
    
        SQL = "insert into gastosfijos_recibos(codigo,fecha,importe,contabilizado) values (" & DBSet(RecuperaValor(Parametros, 1), "N") & ","
        SQL = SQL & DBSet(txtFecha(0).Text, "F") & "," & DBSet(Text1(4).Text, "N") & ",0)"
    
    Else
        SQL = ""
        If txtFecha(0).Text <> txtFecha(0).Tag Then SQL = " , fecha =" & DBSet(txtFecha(0).Text, "F")
        
        SQL = "update gastosfijos_recibos set importe = " & DBSet(Text1(4).Text, "N") & SQL
        SQL = SQL & " where codigo = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and fecha = " & DBSet(txtFecha(0).Tag, "F")
    
    End If
    
    Conn.Execute SQL
    InsertarModificarLinea = True
    End If
    
    Exit Function
    
eModificarGasto:
    MuestraError Err.Number, "Modificar Gasto", Err.Description
End Function




Private Function ModificarGasto() As Boolean
Dim SQL As String

    On Error GoTo eModificarGasto
    
    ModificarGasto = False
    
    SQL = "update gastosfijos set descripcion = " & DBSet(Text1(8).Text, "T") & ", ctaprevista = " & DBSet(txtCuentas(2).Text, "T")
    SQL = SQL & ", contrapar = " & DBSet(txtCuentas(3), "T")
    SQL = SQL & " where codigo = " & DBSet(RecuperaValor(Parametros, 1), "N")
    
    Conn.Execute SQL
    
    ModificarGasto = True
    Exit Function
    
eModificarGasto:
    MuestraError Err.Number, "Modificar Gasto", Err.Description
End Function


Private Function GenerarGasto() As Boolean
Dim Perio As Integer
Dim nVeces As Integer
Dim SqlValues As String
Dim SqlInsert As String
Dim Fecha As Date
Dim NumGasto As Long

    On Error GoTo eGenerarGasto

    GenerarGasto = False
    
    Conn.BeginTrans
    
    NumGasto = SugerirCodigoSiguiente
    
    SQL = "insert into gastosfijos (codigo, descripcion ,ctaprevista,contrapar) values ( " & DBSet(NumGasto, "N") & ","
    SQL = SQL & DBSet(Text1(1), "T") & "," & DBSet(txtCuentas(0).Text, "T") & "," & DBSet(txtCuentas(1).Text, "T") & ")"
    
    Conn.Execute SQL
    
    Select Case Combo1(0).ListIndex
        Case 0 ' mensual
            Perio = 1
            nVeces = 12
        Case 1 ' bimensual
            Perio = 2
            nVeces = 6
        Case 2 ' trimestral
            Perio = 3
            nVeces = 4
        Case 3 ' semestral
            Perio = 6
            nVeces = 2
        Case 4 ' anual
            Perio = 12
            nVeces = 1
    End Select
    
    SqlInsert = "insert into gastosfijos_recibos(codigo, fecha, importe, contabilizado) values "
    SqlValues = ""
    
    Fecha = CDate(txtFecha(5).Text)
    SqlValues = SqlValues & "(" & DBSet(NumGasto, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Text1(5).Text, "N") & ",0),"
    
    
    J = 0
    For i = 1 To nVeces - 1
        J = J + Perio
        Fecha = DateAdd("m", J, CDate(txtFecha(5).Text))
        
        SqlValues = SqlValues & "(" & DBSet(NumGasto, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Text1(5).Text, "N") & ",0),"
    Next i
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert & SqlValues
    End If
    
    Conn.CommitTrans
    GenerarGasto = True
    
    
    Exit Function

eGenerarGasto:
    Conn.RollbackTrans
    MuestraError Err.Number, "Generar Gasto", Err.Description
End Function

Private Sub cmdCancelar_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdContabiliGasto_Click()
    If txtFecha(19).Text = "" Or txtCuentas(4).Text = "" Or Text1(3).Text = "" Or _
        txtDiario(0).Text = "" Or txtCuentas(5).Text = "" Or txtConcepto(0).Text = "" Then
            MsgBox "Campos vacios. Todos los campos son obligatorios", vbExclamation
            Exit Sub
    End If
    
    If txtCC(0).visible Then
        If txtCC(0).Text = "" Then
            MsgBox "Centro de coste obligatorio", vbExclamation
            Exit Sub
        End If
    End If
    
     
    'OK. Contabilizamos
    '---------------------------------------------
    
    
    Conn.BeginTrans
    
    If ContabilizarGastoFijo Then
        Conn.CommitTrans
        
        MsgBox "Proceso realizado correctamente.", vbExclamation
        CadenaDesdeOtroForm = "OK"
        
        Unload Me
    Else
        TirarAtrasTransaccion
    End If
    
    

End Sub


Private Function ContabilizarGastoFijo() As Boolean
Dim Mc As Contadores
Dim FechaAbono As Date
Dim Importe As Currency
    On Error GoTo EContabilizarGastoFijo
    ContabilizarGastoFijo = False
    Set Mc = New Contadores
    
    FechaAbono = CDate(txtFecha(19).Text)
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
   
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & txtDiario(0).Text & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Gasto fijo : " & RecuperaValor(Parametros, 1) & " - " & DevNombreSQL(RecuperaValor(Parametros, 2)) & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & DevNombreSQL(vUsu.Nombre) & "',"
    
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Gastos Fijos');"
    
    If Not Ejecuta(SQL) Then Exit Function
    
    If InStr(1, Text1(3).Text, ",") > 0 Then
        'Texto formateado
        Importe = ImporteFormateado(Text1(3).Text)
    Else
        Importe = CCur(TransformaPuntosComas(Text1(3).Text))
    End If
    i = 1
    Do
        'Lineas de apuntes .
         SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
         SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
         SQL = SQL & " timporteH, ctacontr, codccost,idcontab, punteada) "
         SQL = SQL & "VALUES (" & txtDiario(0).Text & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & i & ",'"
         
         'Cuenta
         If i = 1 Then
            SQL = SQL & txtCuentas(5).Text
         Else
            SQL = SQL & txtCuentas(4).Text
        End If
        SQL = SQL & "','" & Format(Val(RecuperaValor(Parametros, 1)), "000000000") & "'," & txtConcepto(0).Text & ",'"
        
        'Ampliacion
        SQL = SQL & DevNombreSQL(Mid(txtNConcepto(0).Text & " " & Text1(9).Text, 1, 30)) & "',"
                        
        If i = 1 Then
            SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ",NULL,'"
            'Contrapar
            SQL = SQL & txtCuentas(4).Text
        Else
            SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",'"
            'Contrpar
            SQL = SQL & txtCuentas(5).Text
        End If
        
        'Solo para la line NO banco
        If i = 1 And txtCC(0).visible Then
            SQL = SQL & "','" & txtCC(0).Text & "'"
        Else
            SQL = SQL & "',NULL"
        End If
        SQL = SQL & ",'CONTAB',0)"
        
        If Not Ejecuta(SQL) Then Exit Function
        i = i + 1
    Loop Until i > 2  'Una para el banoc, otra para la cuenta
   
    

    'AHora actualizamos el gasto
    FechaAbono = RecuperaValor(vWhere, 2)
    SQL = "UPDATE gastosfijos_recibos SET"
    SQL = SQL & " contabilizado=1"
    SQL = SQL & " WHERE codigo=" & RecuperaValor(vWhere, 1)
    SQL = SQL & " and fecha='" & Format(FechaAbono, FormatoFecha) & "'"
    Conn.Execute SQL


    
    
    ContabilizarGastoFijo = True
    Exit Function
EContabilizarGastoFijo:
    MuestraError Err.Number, "Contabilizar Gasto Fijo"
End Function




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
            Case 1 ' alta cabecera
                Text1(0).Text = SugerirCodigoSiguiente
                txtFecha(5).Text = Format(Now, "dd/mm/yyyy")
                
                PonFoco Text1(1)
                
            Case 2  ' modificacion cabecera
                Label3(3).Caption = "Gasto Fijo : " & RecuperaValor(Parametros, 1)
                Text1(8).Text = RecuperaValor(Parametros, 2)
                txtCuentas(2).Text = RecuperaValor(Parametros, 3)
                txtCuentas(3).Text = RecuperaValor(Parametros, 4)
                txtNCuentas(2).Text = RecuperaValor(Parametros, 5)
                txtNCuentas(3).Text = RecuperaValor(Parametros, 6)
                
                PonFoco Text1(8)
                
            Case 3 ' alta linea
                Label3(9).Caption = "Gasto Fijo : " & RecuperaValor(Parametros, 1) & " " & RecuperaValor(Parametros, 2)
                Text1(4).Text = ""
                txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
                
                PonFoco txtFecha(0)
            
            Case 4 ' modificacion linea
                Label3(9).Caption = "Gasto Fijo : " & RecuperaValor(Parametros, 1) & " " & RecuperaValor(Parametros, 2)
                Text1(4).Text = RecuperaValor(Parametros, 4)
                txtFecha(0).Text = RecuperaValor(Parametros, 3)
                txtFecha(0).Tag = txtFecha(0).Text
                
                'txtFecha(0).Enabled = False
                'ImgFec(0).Enabled = False
                'ImgFec(0).visible = False
                
                
                PonFoco Text1(4)
                
            Case 5 ' contabilizacion del gasto
                Label3(3).Caption = "Gasto Fijo : " & RecuperaValor(Parametros, 1)
            
            
                txtCuentas(4).Text = RecuperaValor(Parametros, 3)
                txtNCuentas(4).Text = RecuperaValor(Parametros, 4)
                txtCuentas(5).Text = RecuperaValor(Parametros, 5)
                txtNCuentas(5).Text = RecuperaValor(Parametros, 6)
                Text1(9).Text = RecuperaValor(Parametros, 2)
                'Fecha e Importe
                txtFecha(19).Text = RecuperaValor(Parametros, 7)
                Text1(3).Text = RecuperaValor(Parametros, 8)
                'ASignaremos cadenadesdeotroform el valor para hacer el UPDATE del registro SI se contabiliza
                vWhere = RecuperaValor(Parametros, 1) & "|"
                vWhere = vWhere & txtFecha(19).Text & "|" & Text1(9).Text & "|"
                
                VisibleCC
            
        
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub VisibleCC()
Dim B As Boolean

    B = False
    If vParam.autocoste Then
        If txtCuentas(5).Text <> "" Then
                SQL = "|" & Mid(txtNCuentas(5).Text, 1, 1) & "|"
                
                '###AQUI
                'If InStr(1, CuentasCC, Sql) > 0 Then B = True
        End If
    End If
    Label1(14).visible = B
    txtCC(0).visible = B
    txtNCC(0).visible = B
    imgCC(0).visible = B
End Sub





Private Function SugerirCodigoSiguiente() As String
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    
    SQL = "Select Max(codigo) from gastosfijos"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            SQL = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = SQL
End Function


'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And Opcion = 23 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim W, H
    PrimeraVez = True
    
    Me.imgCuentas(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(2).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(3).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(4).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(5).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgDiario(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgConcepto(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    Me.ImgFec(0).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Me.ImgFec(5).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Me.ImgFec(19).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    
    
    
    Me.FrameAltaGastoFijo.visible = False
    Me.FrameModGastoFijo.visible = False
    Me.FrameAltaModLineaGasto.visible = False
    Me.FrameContabilizarGasto.visible = False
    
    
    Select Case Opcion
    Case 1
        Me.Caption = "Nuevo Gasto Fijo"
        W = Me.FrameAltaGastoFijo.Width
        H = Me.FrameAltaGastoFijo.Height
        Me.FrameAltaGastoFijo.visible = True
        
        CargarCombo

    Case 2
        Me.Caption = "Modificación Gasto Fijo"
        W = Me.FrameModGastoFijo.Width
        H = Me.FrameModGastoFijo.Height + 150
        Me.FrameModGastoFijo.visible = True
    
    Case 3, 4
        If Opcion = 3 Then
            Me.Caption = "Nueva Linea de Gasto"
        Else
            Me.Caption = "Modificación Linea de Gasto"
        End If
        W = Me.FrameAltaModLineaGasto.Width
        H = Me.FrameAltaModLineaGasto.Height + 200
        Me.FrameAltaModLineaGasto.visible = True
        
        
    Case 5
        Me.Caption = "Contabilización Gasto Fijo"
        W = Me.FrameContabilizarGasto.Width
        H = Me.FrameContabilizarGasto.Height + 150
        Me.FrameContabilizarGasto.visible = True
    
    
    End Select
    
    Me.Width = W + 320
    Me.Height = H + 320
End Sub









Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)

End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
Dim RC As Byte
    'Concepto
    txtConcepto(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNConcepto(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmDi_DatoSeleccionado(CadenaSeleccion As String)
    'Diario
    txtDiario(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNDiario(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgCC_Click(Index As Integer)
    If txtCC(0).Enabled Then
        Set frmCC = New frmBasico
        AyudaCC frmCC
        Set frmCC = Nothing
    End If

End Sub


Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    'Centro de coste
    txtCC(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCC(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgConcepto_Click(Index As Integer)
    
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|"
    frmCon.Show vbModal
    Set frmCon = Nothing

End Sub

Private Sub imgCuentas_Click(Index As Integer)
    SQL = ""
    AbiertoOtroFormEnListado = True
    Set frmCtas = New frmColCtas
    frmCtas.DatosADevolverBusqueda = True
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    If SQL <> "" Then
        Me.txtCuentas(Index).Text = RecuperaValor(SQL, 1)
        Me.txtNCuentas(Index).Text = RecuperaValor(SQL, 2)
    Else
        QuitarPulsacionMas Me.txtCuentas(Index)
    End If
    
    PonFoco Me.txtCuentas(Index)
    AbiertoOtroFormEnListado = False

End Sub

Private Sub imgDiario_Click(Index As Integer)
    'Tipos diario
    Set frmDi = New frmTiposDiario
    frmDi.DatosADevolverBusqueda = "0"
    frmDi.Show vbModal
    Set frmDi = Nothing
    PonFoco txtDiario(0)

End Sub

Private Sub imgFec_Click(Index As Integer)
    'FECHA FACTURA
    Indice = Index
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtFecha(Indice).Text <> "" Then frmF.Fecha = CDate(txtFecha(Indice).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtFecha(Indice)

End Sub



Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    If Text1(Index).Text = "" Then
        Exit Sub
    End If
    
    Select Case Index
        Case 4, 5 ' importe
            PonerFormatoDecimal Text1(Index), 1
            
            
    End Select

End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub



Private Sub txtCC_GotFocus(Index As Integer)
    ConseguirFoco txtCC(Index), 0
End Sub

Private Sub txtCC_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCC_LostFocus(Index As Integer)
    txtCC(Index).Text = Trim(txtCC(Index).Text)
    SQL = ""
    i = 0
    If txtCC(Index).Text <> "" Then
            
        SQL = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtCC(Index).Text, "T")
        If SQL = "" Then
            MsgBox "Concepto no existe", vbExclamation
            i = 1
        End If

    End If
    Me.txtNCC(Index).Text = SQL
    If i = 1 Then
        txtCC(Index).Text = ""
        PonFoco txtCC(Index)
    End If

End Sub

Private Sub txtDiario_GotFocus(Index As Integer)
    ConseguirFoco txtDiario(Index), 3
End Sub

Private Sub txtDiario_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtDiario(Index).Tag, Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
        Case "imgDiario"
            imgDiario_Click Indice
        Case "imgConcepto"
            imgConcepto_Click Indice
    End Select
    
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String

    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'TiposDiarioS
            txtNDiario(Index).Text = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index), "N")
            If txtDiario(Index).Text <> "" Then txtDiario(Index).Text = Format(txtDiario(Index).Text, "000")
    End Select

End Sub


Private Sub txtConcepto_GotFocus(Index As Integer)
    ConseguirFoco txtConcepto(Index), 3
End Sub

Private Sub txtConcepto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtConcepto(Index).Tag, Index
    End If
End Sub



Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
   ' KEYpressGnral KeyAscii
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'CONCEPTOS
            txtNConcepto(Index).Text = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index), "N")
            If txtConcepto(Index).Text <> "" Then txtConcepto(Index).Text = Format(txtConcepto(Index).Text, "000")
    End Select

End Sub




Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub

Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
'        LanzaFormAyuda txtCuentas(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCuentas_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

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
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 2, 3 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
            Else
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = SQL
                If B = 1 Then
                    txtNCuentas(Index).Tag = ""
                Else
                    txtNCuentas(Index).Tag = SQL
                End If
'                Hasta = -1
'                If Index = 6 Then
'                    Hasta = 7
'                Else
'                    If Index = 0 Then
'                        Hasta = 1
'                    Else
'                        If Index = 5 Then
'                            Hasta = 4
'                        Else
'                            If Index = 23 Then Hasta = 24
'                        End If
'                    End If
'
'                End If
'
'                If Hasta >= 0 Then
'                    txtCuentas(Hasta).Text = txtCuentas(Index).Text
'                    txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
'                End If
            End If
    
    End Select
    
End Sub





Private Sub txtFecha_GotFocus(Index As Integer)
    txtFecha(Index).SelStart = 0
    txtFecha(Index).SelLength = Len(txtFecha(Index).Text)
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index))
    If txtFecha(Index) = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index), vbExclamation
        txtFecha(Index).Text = ""
        txtFecha(Index).SetFocus
    End If
End Sub



Private Sub EjecutarSQL()
    On Error Resume Next
    
    Conn.Execute SQL
    If Err.Number <> 0 Then
        If Conn.Errors(0).Number = 1062 Then
            Err.Clear
        Else
            'MuestraError Err.Number, Err.Description
        End If
        Err.Clear
    End If
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim J As Long
    
    Combo1(0).Clear

    Combo1(0).AddItem "Mensual "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Bimensual "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Trimestral "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Semestral "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "Anual "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5



End Sub


