VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfEvolSal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
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
      Height          =   5565
      Left            =   7110
      TabIndex        =   14
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   150
         TabIndex        =   42
         Top             =   2760
         Width           =   4155
         Begin VB.OptionButton Option3 
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
            Height          =   225
            Left            =   2760
            TabIndex        =   45
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton Option2 
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
            Height          =   240
            Left            =   1530
            TabIndex        =   44
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton Option1 
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
            Height          =   240
            Left            =   240
            TabIndex        =   43
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.CheckBox chkEvolSalMeses 
         Caption         =   "Mostrar meses sin movimientos"
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
         Left            =   270
         TabIndex        =   40
         Top             =   2970
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         Height          =   1905
         Left            =   150
         TabIndex        =   25
         Top             =   750
         Width           =   4185
         Begin VB.CheckBox ChkEvolSaldo 
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
            TabIndex        =   35
            Top             =   1290
            Width           =   1335
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            Left            =   2850
            TabIndex        =   34
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            Left            =   1470
            TabIndex        =   33
            Top             =   960
            Width           =   1305
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            TabIndex        =   32
            Top             =   930
            Width           =   1305
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            Left            =   2850
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            Left            =   1470
            TabIndex        =   30
            Top             =   600
            Width           =   1305
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            TabIndex        =   29
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            Left            =   2850
            TabIndex        =   28
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            Left            =   1470
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox ChkEvolSaldo 
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
            TabIndex        =   26
            Top             =   240
            Value           =   1  'Checked
            Width           =   1155
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   24
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
      Height          =   2835
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNCta 
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
         Index           =   6
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1050
         Width           =   4185
      End
      Begin VB.TextBox txtNCta 
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
         Index           =   7
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1470
         Width           =   4185
      End
      Begin VB.ComboBox cmbEjercicios 
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
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2220
         Width           =   4095
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
         Height          =   360
         Index           =   6
         Left            =   1230
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1050
         Width           =   1275
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
         Height          =   360
         Index           =   7
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Ejercicio"
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
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label lblFecha 
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
         Left            =   2520
         TabIndex        =   23
         Top             =   2310
         Width           =   4095
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   6
         Left            =   930
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   7
         Left            =   930
         Top             =   1500
         Width           =   255
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
         Index           =   0
         Left            =   2550
         TabIndex        =   22
         Top             =   990
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
         TabIndex        =   21
         Top             =   1440
         Width           =   4095
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
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1440
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
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta"
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
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Top             =   690
         Width           =   960
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
      TabIndex        =   4
      Top             =   5850
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
      TabIndex        =   2
      Top             =   5850
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
      TabIndex        =   3
      Top             =   5790
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
      TabIndex        =   5
      Top             =   2910
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
         TabIndex        =   17
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   15
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1830
      TabIndex        =   36
      Top             =   5850
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
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
      Top             =   5850
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
      Index           =   29
      Left            =   1830
      TabIndex        =   41
      Top             =   5850
      Width           =   5175
   End
End
Attribute VB_Name = "frmInfEvolSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 311

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

Public Legalizacion As String   'Datos para la legalizacion

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

Private Sub ChkEvolSaldo_Click(Index As Integer)
    If ChkEvolSaldo(Index).Value = 1 Then
        For I = 1 To 10
            If I <> Index Then ChkEvolSaldo(I).Value = 0
        Next I
    End If

End Sub

Private Sub ChkEvolSaldo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
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
    

    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True

    
    If Not MontaSQL Then Exit Sub
    
    
    
    
    Screen.MousePointer = vbHourglass
    SQL = "DELETE FROM tmpconextcab where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "DELETE FROM tmpconext where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "DELETE FROM tmpevolsal where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    
    DoEvents 'Para que no bloquee la pantalla
    Label2(29).Caption = "Leyendo datos BD"
    Label2(29).Refresh
    If ListadoEvolucionMensual Then
    
        If Not HayRegParaInforme("tmpevolsal", "codusu=" & vUsu.Codigo) Then Exit Sub
        
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
    
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    If Me.cmdCancelarAccion.Visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Evoluci�n de Saldos"

    For I = 6 To 7
        Me.imgCuentas(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    PrimeraVez = True
     
    PonerNiveles
    CargaComboEjercicios 0

    Me.Option3.Value = True
   
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
    PonerNiveles
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub Image1_Click(Index As Integer)

    Select Case Index
        Case 0 'cuentas agrupadas
            Set frmCtas = New frmCtasAgrupadas
            frmCtas.Show vbModal
            Set frmCtas = Nothing
    End Select

End Sub

Private Sub imgCuentas_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing

    PonFoco txtCta(Index)

End Sub



Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
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

Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
End Sub


Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgCuentas"
        imgCuentas_Click Indice
    End Select
    
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        txtNCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser num�rica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        txtNCta(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 6, 7 'Cuentas
            
            RC = txtCta(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, SQL) Then
                txtCta(Index) = RC
                txtNCta(Index).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtCta(Index).Text = ""
                txtNCta(Index).Text = ""
                PonFoco txtCta(Index)
            End If
            
            If Index = 0 Then Hasta = 1
            If Hasta >= 1 Then
                txtCta(Hasta).Text = txtCta(Index).Text
                txtNCta(Hasta).Text = txtNCta(Index).Text
            End If
    End Select

End Sub



Private Sub AccionesCSV()
Dim SQL2 As String
Dim Tipo As Byte
            

    Select Case Tipo
        Case 1
            SQL = "select cta Cuenta , nomcta Titulo, totald Saldo_deudor, totalh Saldo_acreedor from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
        Case 2
            SQL = "select cta Cuenta , nomcta Titulo, acumantd AcumAnt_deudor, acumanth AcumAnt_acreedor, acumperd AcumPer_deudor, acumperh AcumPer_acreedor, totald Saldo_deudor, totalh Saldo_acreedor from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
        Case 3
            SQL = "select cta Cuenta , nomcta Titulo, aperturad Apertura_deudor, aperturah Apertura_acreedor,  totald Saldo_deudor, totalh Saldo_acreedor from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
        Case 4
            SQL = "select cta Cuenta , nomcta Titulo, aperturad, aperturah, case when coalesce(aperturad,0) - coalesce(aperturah,0) > 0 then concat(coalesce(aperturad,0) - coalesce(aperturah,0),'D') when coalesce(aperturad,0) - coalesce(aperturah,0) < 0 then concat(coalesce(aperturah,0) - coalesce(aperturad,0),'H') when coalesce(aperturad,0) - coalesce(aperturah,0) = 0 then 0 end Apertura, "
            SQL = SQL & " acumantd AcumAnt_deudor, acumanth AcumAnt_acreedor, acumperd AcumPer_deudor, acumperh AcumPer_acreedor, "
            SQL = SQL & " totald Saldo_deudor, totalh Saldo_acreedor, case when coalesce(totald,0) - coalesce(totalh,0) > 0 then concat(coalesce(totald,0) - coalesce(totalh,0),'D') when coalesce(totald,0) - coalesce(totalh,0) < 0 then concat(coalesce(totalh,0) - coalesce(totald,0),'H') when coalesce(totald,0) - coalesce(totalh,0) = 0 then 0 end Saldo"
            SQL = SQL & " from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
    End Select
    


        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String
Dim k As Integer
Dim K1 As Integer

    If Option1.Value Then Tipo = 0
    If Option2.Value Then Tipo = 1
    If Option3.Value Then Tipo = 2
            
    cadParam = cadParam & "pTipo=" & Tipo & "|"
    numParam = numParam + 1
        
    '------------------------------
    'Cuentas
    RC = cmbEjercicios(0).List(cmbEjercicios(0).ListIndex)
    RC = "Ejercicio " & Mid(RC, 1, 23) & "  "
    If txtCta(6).Text <> "" Then RC = RC & " desde " & txtCta(6).Text & " -" & txtNCta(6).Text
    If txtCta(7).Text <> "" Then RC = RC & " hasta " & txtCta(7).Text & " -" & txtNCta(7).Text
    If RC <> "" Then RC = "Cuentas: " & RC

    
    
    cadParam = cadParam & "Cuenta= """ & RC & """|"
    numParam = numParam + 2

    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'A�o natural
        cadParam = cadParam & "pMes1=""ENERO""|"
        cadParam = cadParam & "pMes2=""FEBRERO""|"
        cadParam = cadParam & "pMes3=""MARZO""|"
        cadParam = cadParam & "pMes4=""ABRIL""|"
        cadParam = cadParam & "pMes5=""MAYO""|"
        cadParam = cadParam & "pMes6=""JUNIO""|"
        cadParam = cadParam & "pMes7=""JULIO""|"
        cadParam = cadParam & "pMes8=""AGOSTO""|"
        cadParam = cadParam & "pMes9=""SEPTIEMBRE""|"
        cadParam = cadParam & "pMes10=""OCTUBRE""|"
        cadParam = cadParam & "pMes11=""NOVIEMBRE""|"
        cadParam = cadParam & "pMes12=""DICIEMBRE""|"
    Else
        'A�os fiscales partidos . Coooperativas agricolas
        For k = 0 To 11
            K1 = Month(FechaIncioEjercicio) + k
            If K1 > 12 Then K1 = K1 - 12
            Select Case k
                Case 0
                    cadParam = cadParam & "pMes1=""" & UCase(MonthName(K1)) & """|"
                Case 1
                    cadParam = cadParam & "pMes2=""" & UCase(MonthName(K1)) & """|"
                Case 2
                    cadParam = cadParam & "pMes3=""" & UCase(MonthName(K1)) & """|"
                Case 3
                    cadParam = cadParam & "pMes4=""" & UCase(MonthName(K1)) & """|"
                Case 4
                    cadParam = cadParam & "pMes5=""" & UCase(MonthName(K1)) & """|"
                Case 5
                    cadParam = cadParam & "pMes6=""" & UCase(MonthName(K1)) & """|"
                Case 6
                    cadParam = cadParam & "pMes7=""" & UCase(MonthName(K1)) & """|"
                Case 7
                    cadParam = cadParam & "pMes8=""" & UCase(MonthName(K1)) & """|"
                Case 8
                    cadParam = cadParam & "pMes9=""" & UCase(MonthName(K1)) & """|"
                Case 9
                    cadParam = cadParam & "pMes10=""" & UCase(MonthName(K1)) & """|"
                Case 10
                    cadParam = cadParam & "pMes11=""" & UCase(MonthName(K1)) & """|"
                Case 11
                    cadParam = cadParam & "pMes12=""" & UCase(MonthName(K1)) & """|"
             End Select
        Next k
    End If
    
    numParam = numParam + 12




    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "0311-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"evolsald1.rpt"

    cadFormula = "{tmpevolsal.codusu}=" & vUsu.Codigo

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

    MontaSQL = False
    
    
    MontaSQL = True
           
End Function


Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    SQL = ""
    For I = 1 To 10
        If Me.ChkEvolSaldo(I).Visible Then
            If Me.ChkEvolSaldo(I).Value = 1 Then SQL = SQL & "1"
        End If
    Next I
    
    If Len(SQL) <> 1 Then
        MsgBox "Eliga un nivel (y solo uno) para el listado de  evoluci�n mesual de saldos", vbExclamation
        Exit Function
    End If



    DatosOK = True

End Function




Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
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
    ChkEvolSaldo(10).Visible = True
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        cad = "Digitos: " & J
        ChkEvolSaldo(I).Visible = True
        Me.ChkEvolSaldo(I).Caption = cad
        
    Next I
    For I = vEmpresa.numnivel To 9
        ChkEvolSaldo(I).Visible = False
    Next I
    
    
End Sub



'Cargo en el combo los ejercicios para que los seleccione
Private Sub CargaComboEjercicios(Indice As Integer)
Dim RS As Recordset
Dim PrimeraVez As Boolean
Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim cad As String
        On Error GoTo ECargaComboEjericios
        
        Set RS = New ADODB.Recordset
        cad = "Select min(fechaent) from hcabapu"  'FECHA MINIMA
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        FechaIncioEjercicio = vParam.fechaini
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then FechaIncioEjercicio = RS.Fields(0)
        End If
        RS.Close
        Set RS = Nothing
        
        'Cargo el combo
        '--------------------------------------------------------------------------------
        'Ajusto la primera fecha que devuelve a la que seria Inicio de ese ejercicio
        FechaFinEjercicio = CDate(Format(vParam.fechaini, "dd/mm/" & Year(FechaIncioEjercicio)))
        
        If FechaFinEjercicio > FechaIncioEjercicio Then
            'El ejercicio empieza un a�o antes
            FechaIncioEjercicio = DateAdd("yyyy", -1, FechaFinEjercicio)
        Else
            FechaIncioEjercicio = FechaFinEjercicio
        End If
        
        
        
        FechaFinEjercicio = DateAdd("yyyy", 1, vParam.fechafin)  'Final de a�o siguiente
        cmbEjercicios(Indice).Clear
        CONT = 0
        While FechaIncioEjercicio <= FechaFinEjercicio
                cad = Format(FechaIncioEjercicio, "dd/mm/yyyy")
                FechaIncioEjercicio = DateAdd("yyyy", 1, FechaIncioEjercicio)
                FechaIncioEjercicio = DateAdd("d", -1, FechaIncioEjercicio)
                cad = cad & " - " & Format(FechaIncioEjercicio, "dd/mm/yyyy")
                'Le pongo una marca de actual o ssiguiente
                I = 0 'pAra memorizar cual es el que apunta
                If FechaIncioEjercicio > vParam.fechaini Then
                    If FechaIncioEjercicio = vParam.fechafin Then
                        cad = cad & "     Actual"
                        I = 1
                    Else
                        cad = cad & "     Siguiente"
                    End If
                End If
                'Meto en el combo
                cmbEjercicios(Indice).AddItem cad
                If I = 1 Then CONT = cmbEjercicios(Indice).NewIndex
                'Paso a inicio del ejercicio siguiente sumandole un dia
                'al fin del anterior
                FechaIncioEjercicio = DateAdd("d", 1, FechaIncioEjercicio)
        Wend
             
        'En cont tengo actual
        Me.cmbEjercicios(Indice).ListIndex = CONT
        
        Exit Sub
ECargaComboEjericios:
    MuestraError Err.Number, "CargaComboEjericios"
    
End Sub

Private Function ListadoEvolucionMensual() As Boolean
Dim QuitarTambienElCierre As Boolean
Dim Tipo As Integer

    On Error GoTo EListadoEvolucionMensual
    ListadoEvolucionMensual = False


    'En cmbejercicios(0) tenemos las fechas
    '
    '   Con simple mid obtenemos inicio / fin
    
    
    SQL = cmbEjercicios(0).List(cmbEjercicios(0).ListIndex)
    RC = Mid(SQL, 1, 10)
    FechaIncioEjercicio = CDate(RC)
    RC = Mid(SQL, 14, 10)
    FechaFinEjercicio = CDate(RC)
    
    SQL = "Select hlinapu.codmacta,nommacta from hlinapu,cuentas where hlinapu.codmacta=cuentas.codmacta "

    'Si tienen desde /hasta
    If txtCta(6).Text <> "" Then SQL = SQL & " AND hlinapu.codmacta >= '" & txtCta(6).Text & "'"
    If txtCta(7).Text <> "" Then SQL = SQL & " AND hlinapu.codmacta <= '" & txtCta(7).Text & "'"

    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'A�o natural
        SQL = SQL & " AND year(fechaent) = " & Year(FechaIncioEjercicio)
    Else
        'A�os fiscales partidos . Coooperativas agricolas
        SQL = SQL & " AND ( (year(fechaent) = " & Year(FechaIncioEjercicio) & " and month(fechaent) >=" & Month(FechaIncioEjercicio) & ") OR "
        SQL = SQL & " (year(fechaent) =" & Year(FechaIncioEjercicio) + 1 & " AND month(fechaent) < " & Month(FechaIncioEjercicio) & "))"
    End If
    
    SQL = SQL & " GROUP BY hlinapu.codmacta"
    
    
    If Option1.Value Then Tipo = 0
    If Option2.Value Then Tipo = 1
    If Option3.Value Then Tipo = 2
    
    
    
    'stop
    QuitarTambienElCierre = FechaIncioEjercicio < vParam.fechaini
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        FijarValoresEvolucionMensualSaldos FechaIncioEjercicio, FechaFinEjercicio
        
        'PAra el SQL
        SQL = ""
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            'A�o natural
            SQL = SQL & " AND year(fechaent) = " & Year(FechaIncioEjercicio)
        Else
            'A�os fiscales partidos . Coooperativas agricolas
            SQL = SQL & " AND ( (year(fechaent) = " & Year(FechaIncioEjercicio) & " and month(fechaent) >=" & Month(FechaIncioEjercicio) & ") OR "
            SQL = SQL & " (year(fechaent) =" & Year(FechaIncioEjercicio) + 1 & " AND month(fechaent) < " & Month(FechaIncioEjercicio) & "))"
        End If
        CONT = 0
        While Not RS.EOF
            Label2(29).Caption = RS!codmacta & " " & Mid(RS!Nommacta, 1, 20) & " ..."
            Me.Refresh
            DatosEvolucionMensualSaldos2 RS!codmacta, RS!Nommacta, SQL, True, False, QuitarTambienElCierre, FechaIncioEjercicio, Tipo
            RS.MoveNext
            If CONT > 150 Then
                CONT = 0
                DoEvents
                Screen.MousePointer = vbHourglass
            End If
            CONT = CONT + 1
        Wend
    End If
    RS.Close
    
    
    
    'Hacemos el conteo para ver si tiene o no movimientos
    Label2(29).Caption = "Comprobando valores"
    Label2(29).Refresh
    SQL = "Select count(*) from tmpconextcab"
    SQL = SQL & " WHERE codusu =" & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not RS.EOF Then CONT = DBLet(RS.Fields(0), "N")
    RS.Close
    Set RS = Nothing
    If CONT > 0 Then
        ListadoEvolucionMensual = True
    Else
        MsgBox "No hay datos con estos valores", vbExclamation
    End If
    
    
    Label2(29).Caption = ""
    Exit Function
EListadoEvolucionMensual:
    MuestraError Err.Number
    Set RS = Nothing
    Label2(29).Caption = ""
End Function


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
