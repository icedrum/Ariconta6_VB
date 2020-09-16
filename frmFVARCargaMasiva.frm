VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFVARCargaMasiva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga Masiva de Facturas Varias"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   9015
   Icon            =   "frmFVARCargaMasiva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   7635
      Left            =   45
      TabIndex        =   13
      Top             =   45
      Width           =   8895
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
         Index           =   6
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   450
         Width           =   5070
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2265
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   450
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   10
         Left            =   2190
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|000|S|"
         Top             =   4695
         Width           =   6195
      End
      Begin VB.TextBox txtcodigo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   9
         Left            =   360
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Observaciones|T|S|||cabfact|observac|||"
         Top             =   3240
         Width           =   8025
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   3885
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   5430
         Width           =   1860
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   5430
         Width           =   1230
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   360
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   5430
         Width           =   1230
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
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   4200
         Width           =   5115
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2190
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|000|S|"
         Top             =   4200
         Width           =   1050
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
         Index           =   5
         Left            =   3585
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   2640
         Width           =   4845
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
         Index           =   4
         Left            =   3585
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   2235
         Width           =   4845
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   8
         Left            =   2265
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   1440
         Width           =   1050
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
         Index           =   8
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   1440
         Width           =   5070
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2250
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   945
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2625
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2235
         Width           =   1350
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
         Left            =   7365
         TabIndex        =   12
         Top             =   6855
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptar 
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
         Left            =   6180
         TabIndex        =   11
         Top             =   6855
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   285
         Left            =   315
         TabIndex        =   17
         Top             =   5895
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1935
         MouseIcon       =   "frmFVARCargaMasiva.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar contador"
         Top             =   495
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   33
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ampliación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   10
         Left            =   360
         TabIndex        =   31
         Top             =   4695
         Width           =   1035
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   2970
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   9
         Left            =   3915
         TabIndex        =   29
         Top             =   5190
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   8
         Left            =   2190
         TabIndex        =   28
         Top             =   5190
         Width           =   780
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1920
         MouseIcon       =   "frmFVARCargaMasiva.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1920
         MouseIcon       =   "frmFVARCargaMasiva.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cuenta"
         Top             =   2235
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   7
         Left            =   360
         TabIndex        =   27
         Top             =   5190
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1860
         MouseIcon       =   "frmFVARCargaMasiva.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar concepto"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Concepto "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   26
         Top             =   4200
         Width           =   1020
      End
      Begin VB.Label lblProgres 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   22
         Top             =   6525
         Width           =   7995
      End
      Begin VB.Label lblProgres 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   21
         Top             =   6165
         Width           =   8010
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1935
         MouseIcon       =   "frmFVARCargaMasiva.frx":0554
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar f.pago"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1935
         Picture         =   "frmFVARCargaMasiva.frx":06A6
         ToolTipText     =   "Buscar fecha"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   945
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   1845
         Width           =   1650
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
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Top             =   2655
         Width           =   690
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
         Index           =   0
         Left            =   1080
         TabIndex        =   14
         Top             =   2265
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFVARCargaMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmCtas As frmColCtas 'Cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFpa As frmBasico2 'formas de pago
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmCon As frmFVARConceptos ' conceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCont As frmBasico
Attribute frmCont.VB_VarHelpID = -1
Private WithEvents frmMens As frmFVARInformes ' ayuda de cuentas contables
Attribute frmMens.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim IndCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim BdConta As Integer
Dim Indice As Integer

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim Sql As String
Dim Tipo As Byte
Dim Nregs As Long
Dim NumError As Long
Dim cWhere As String
Dim MensError As String

    If Not DatosOK Then Exit Sub
          
        
    cWhere = " codmacta >= '" & Trim(txtCodigo(4).Text) & "' and codmacta <= '" & Trim(txtCodigo(5).Text) & "'"

    Set frmMens = New frmFVARInformes

    frmMens.OpcionListado = 1
    frmMens.CadTag = " and  " & cWhere
    frmMens.Label5 = "Cuentas Contables"
    frmMens.Caption = ""
    frmMens.Show vbModal

    Set frmMens = Nothing
    
    If cadselect <> "" Then
        GenerarFacturas cadTABLA, cadselect, NumError, MensError
        'Eliminar la tabla TMP
        BorrarTMP
    End If
    'Desbloqueamos ya no estamos contabilizando facturas
    BloqueoManual False, "", "VENCON"  'VENtas CONtabilizar
    
    
    If cadselect = "" Then
        MsgBox "No se ha realizado el proceso, no se han seleccionado cuentas.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        Exit Sub
    End If
    
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de generación. Llame a soporte." & vbCrLf & vbCrLf & MensError
    Else
        MsgBoxA "Proceso realizado correctamente.", vbInformation
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonFoco txtCodigo(6)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(2).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(3).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "fvarfacturas"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    

End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(Indice).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
' concepto de factura
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de concepto
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de concepto
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
' cta de banco
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
' forma de pago
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmCont_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'nroserie
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " cuentas.codmacta in (" & CadenaSeleccion & ")"
        Sql2 = " cuentas.codmacta in [" & CadenaSeleccion & "]"
    Else
        Sql = ""
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Secciones
Dim cad As String

    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    
    cad = RecuperaValor(CadenaSeleccion, 5)  'numconta
    If cad <> "" Then BdConta = CInt(cad)  'numero de conta

End Sub

Private Sub imgFec_Click(Index As Integer)
    Indice = 7
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtCodigo(7).Text <> "" Then frmC.Fecha = CDate(txtCodigo(7).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    PonFoco txtCodigo(7)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 ' Concepto
            IndCodigo = 0
            Set frmCon = New frmFVARConceptos
            frmCon.DatosADevolverBusqueda = "0|1|2|4|"
            frmCon.CodigoActual = txtCodigo(0).Text
            frmCon.Show vbModal
            Set frmCon = Nothing
        
        Case 1, 2 ' Ctas Contables de Socio
            IndCodigo = Index + 3
            Set frmCtas = New frmColCtas
            frmCtas.DatosADevolverBusqueda = "0|1|2|"
            frmCtas.ConfigurarBalances = 3  'NUEVO
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonFoco txtCodigo(IndCodigo)
            
        Case 3  ' letra de serie
            IndCodigo = 6
            Set frmCont = New frmBasico
            AyudaContadores frmCont, txtCodigo(IndCodigo).Text, "tiporegi REGEXP '^[0-9]+$' = 0"
            Set frmCont = Nothing
            PonFoco txtCodigo(IndCodigo)
        
        Case 8 ' Forma de Pago
            IndCodigo = Index
            
            Set frmFpa = New frmBasico2
        
            AyudaFPago frmFpa, txtCodigo(8)

            Set frmFpa = Nothing
    
    End Select
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.Caption = "Facturas por Cliente"
        Case 1
            Me.Caption = "Facturas por Tarjeta"
        Case 2
            Me.Caption = "Facturas por Cliente y por Tarjeta"
    End Select
    
End Sub

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub


Private Sub txtcodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtcodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 6: KEYBusqueda KeyAscii, 3 'letra de serie
            Case 8: KEYBusqueda KeyAscii, 8 'forma de pago
            Case 4: KEYBusqueda KeyAscii, 1 'cta contable desde
            Case 5: KEYBusqueda KeyAscii, 2 'cta contable hasta
            Case 0: KEYBusqueda KeyAscii, 0 'concepto
            Case 7: KEYFecha KeyAscii, 7 'fecha factura
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
Dim cadMen As String
Dim B As Integer
Dim RC As String
Dim Cta As String
Dim Sql As String
Dim Hasta As Integer

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 6 ' letra de serie
            If IsNumeric(txtCodigo(Index).Text) Then
                MsgBoxA "Debe ser una letra: " & txtCodigo(Index).Text, vbExclamation
                txtCodigo(Index).Text = ""
                PonFoco txtCodigo(Index)
            End If
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)

            txtNombre(Index).Text = DevuelveValor("select nomregis from contadores where tiporegi = " & DBSet(txtCodigo(Index), "T") & " and tiporegi REGEXP '^[0-9]+$' = 0")
            If txtNombre(Index).Text = "0" Then
                MsgBoxA "Letra de serie no existe o no es de facturas de cliente. Reintroduzca.", vbExclamation
                txtNombre(Index).Text = ""
                txtCodigo(Index).Text = ""
                PonFoco txtCodigo(Index)
            End If
        
        Case 4, 5 ' Cuenta de contables
            If Not IsNumeric(txtCodigo(Index).Text) Then
                If InStr(1, txtCodigo(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCodigo(Index).Text, vbExclamation
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
                Exit Sub
            End If
        
            Cta = (txtCodigo(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCodigo(Index).Text, vbExclamation
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
            Else
                txtCodigo(Index).Text = Cta
                txtNombre(Index).Text = Sql
                If B = 1 Then
                    txtNombre(Index).Tag = ""
                Else
                    txtNombre(Index).Tag = Sql
                End If
                Hasta = -1
                If Index = 4 Then
                    Hasta = 5
                End If
                    
                If Hasta >= 0 Then
                    txtCodigo(Hasta).Text = txtCodigo(Index).Text
                    txtNombre(Hasta).Text = txtNombre(Index).Text
                End If
            End If
    
            
        Case 8 ' Forma de pago
            txtNombre(8).Text = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", txtCodigo(8).Text, "N")
        
        Case 7  'FECHA FACTURA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 0 ' Concepto
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(0).Text = PonerNombreDeCod(txtCodigo(Index), "fvarconceptos", "nomconce", "codconce", "N")
                If txtNombre(0).Text = "" Then
                    cadMen = "No existe el Concepto: " & txtCodigo(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCon = New frmFVARConceptos
                        frmCon.DatosADevolverBusqueda = "0|1|"
                        frmCon.NuevoCodigo = txtCodigo(Index).Text
                        txtCodigo(Index).Text = ""
                        TerminaBloquear
                        frmCon.Show vbModal
                        Set frmCon = Nothing
                    Else
                        txtCodigo(Index).Text = ""
                    End If
                    PonFoco txtCodigo(Index)
                End If
            Else
                txtNombre(0).Text = ""
            End If
        
        Case 1 ' cantidad
            PonerFormatoDecimal txtCodigo(Index), 3
            txtCodigo(3).Text = Round2(CCur(ComprobarCero(txtCodigo(1).Text)) * CCur(ComprobarCero(txtCodigo(2).Text)), 2)
            PonerFormatoDecimal txtCodigo(3), 3
        
        Case 2 ' precio
            PonerFormatoDecimal txtCodigo(Index), 8
            txtCodigo(3).Text = Round2(CCur(ComprobarCero(txtCodigo(1).Text)) * CCur(ComprobarCero(txtCodigo(2).Text)), 2)
            PonerFormatoDecimal txtCodigo(3), 3
            
        Case 3 ' importe
            PonerFormatoDecimal txtCodigo(Index), 3
        
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 7365
        Me.FrameCobros.Width = 8895
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtCodigo(7).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim cad As String
Dim UltNiv As Integer
Dim PorcIva As String

    DatosOK = False

    If txtCodigo(7).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Factura.", vbExclamation
        PonFoco txtCodigo(7)
        Exit Function
    End If
    
    If txtCodigo(4).Text = "" Or txtCodigo(5).Text = "" Then
        MsgBox "Debe introducir un valor en las cuentas contables", vbExclamation
        PonFoco txtCodigo(4)
        Exit Function
    Else
        If Len(txtCodigo(4).Text) = Len(txtCodigo(5).Text) Then
            If CDbl(txtCodigo(4).Text) > CDbl(txtCodigo(5).Text) Then
                MsgBox "La cuenta desde ha de ser inferior a la cuenta hasta. Revise.", vbExclamation
                PonFoco txtCodigo(4)
                Exit Function
            End If
        Else
            MsgBox "Las cuentas contables deben de tener el mismo nivel. Revise.", vbExclamation
            PonFoco txtCodigo(4)
            Exit Function
        End If
    End If
    
    If txtCodigo(8).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una forma de pago.", vbExclamation
        PonFoco txtCodigo(8)
        Exit Function
    Else
        txtNombre(8).Text = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", txtCodigo(8).Text, "N")
        If txtNombre(8).Text = "" Then
            MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
            PonFoco txtCodigo(8)
            B = False
        End If
    End If
        
    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un concepto.", vbExclamation
        PonFoco txtCodigo(0)
        Exit Function
    Else
        cad = ""
        cad = DevuelveDesdeBD("tipoiva", "fvarconceptos", "codconce", txtCodigo(0).Text, "N")
        If cad = "" Then
            MsgBox "El concepto no tiene asociado un tipo de iva. Revise.", vbExclamation
            PonFoco txtCodigo(0)
            Exit Function
        Else
            ' comprobamos que existe el tipo de iva en contabilidad
            PorcIva = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", cad, "N")
            If PorcIva = "" Then
                MsgBox "No existe el tipo de Iva del concepto. Revise.", vbExclamation
                PonFoco txtCodigo(0)
                B = False
            End If
        End If
    End If
    
    DatosOK = True
End Function

Private Sub GenerarFacturas(cadTABLA As String, CadWhere As String, NumError As Long, MensError As String)
Dim Sql As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cad As String
Dim NumF As Long
Dim CabSql As String
Dim LinSql As String

Dim NumFact As Long

Dim TipoIva As String
Dim PorIva As String
Dim Impoiva As Currency
Dim TotalFact As Currency

Dim Rs As ADODB.Recordset
Dim NomCuenta As String
Dim Existe As Boolean

Dim Mc As Contadores
Dim i As Integer

    On Error GoTo EContab


    Sql = "GENFAC" 'generar facturas de venta

    'Bloquear para que nadie mas pueda contabilizar
    BloqueoManual False, "", Sql
    If Not BloqueoManual(True, Sql, "1") Then
        MsgBox "No se pueden Generar Facturas. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Conn.BeginTrans

    BorrarTMP
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    B = CrearTMP("cuentas", CadWhere, True)
    If Not B Then Exit Sub
            
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    
    NumF = DevuelveValor("select count(*) from tmpfactu")
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, CInt(NumF)
        
    Sql = "select ctaclien, cuentas.* from tmpfactu inner join cuentas on tmpfactu.ctaclien = cuentas.codmacta order by ctaclien"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
    While Not Rs.EOF
    
        IncrementarProgres Me.Pb1, 1
        
        Me.lblProgres(1).Caption = "Procesando Cuenta Contable ..."
        Me.Refresh
        
        Set Mc = New Contadores
        If txtCodigo(7).Text <> "" Then i = FechaCorrecta2(CDate(txtCodigo(7).Text))
        If Mc.ConseguirContador(txtCodigo(6).Text, (i = 0), True) = 0 Then
            NumFact = Mc.Contador
            Existe = False
            Do
                Sql = "select count(*) from fvarfactura where "
                Sql = Sql & " numserie = " & DBSet(Mc.TipoContador, "T")
                Sql = Sql & " and numfactu = " & DBSet(NumFact, "N")
                Sql = Sql & " and fecfactu = " & DBSet(txtCodigo(7).Text, "F")
                
                If TotalRegistros(Sql) > 0 Then
                    NumFact = NumFact + 1
                    Existe = True
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            TipoIva = ""
            PorIva = ""
            Impoiva = 0
            TotalFact = 0
            
            TipoIva = DevuelveDesdeBD("tipoiva", "fvarconceptos", "codconce", txtCodigo(0).Text, "N")
            PorIva = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", TipoIva, "N")
            Impoiva = Round2(CCur(ImporteSinFormato(txtCodigo(3).Text)) * ComprobarCero(PorIva) / 100, 2)
            TotalFact = CCur(ImporteSinFormato(txtCodigo(3).Text)) + Impoiva
            
            ' Insertamos en la cabecera de factura
            CabSql = "insert into fvarfactura ("
            CabSql = CabSql & "numserie,numfactu,fecfactu,codmacta,observac,intconta,baseiva1,baseiva2,baseiva3,"
            CabSql = CabSql & "impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,totalfac,tipoiva1,tipoiva2,tipoiva3,"
            CabSql = CabSql & "porciva1 , porciva2, porciva3, codforpa, porcrec1, porcrec2, porcrec3, retfaccl, trefaccl, cuereten,"
            CabSql = CabSql & "nommacta, dirdatos, codposta, despobla, desprovi, nifdatos, codpais)  values  "
            
            CabSql = CabSql & "(" & DBSet(txtCodigo(6), "T")
            CabSql = CabSql & "," & DBSet(NumFact, "N")
            CabSql = CabSql & "," & DBSet(txtCodigo(7).Text, "F")
            CabSql = CabSql & "," & DBSet(Rs!ctaclien, "T")
            CabSql = CabSql & "," & DBSet(txtCodigo(9).Text, "T", "S")
            CabSql = CabSql & ",0"
            CabSql = CabSql & "," & DBSet(txtCodigo(3).Text, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(Impoiva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            
            CabSql = CabSql & "," & DBSet(TotalFact, "N")
            CabSql = CabSql & "," & DBSet(TipoIva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(PorIva, "N")
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & DBSet(txtCodigo(8).Text, "N") ' forma de pago
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            CabSql = CabSql & "," & ValorNulo
            
            ' datos fiscales
            CabSql = CabSql & "," & DBSet(Rs!Nommacta, "T")
            CabSql = CabSql & "," & DBSet(Rs!dirdatos, "T")
            CabSql = CabSql & "," & DBSet(Rs!codposta, "T")
            CabSql = CabSql & "," & DBSet(Rs!desPobla, "T")
            CabSql = CabSql & "," & DBSet(Rs!desProvi, "T")
            CabSql = CabSql & "," & DBSet(Rs!nifdatos, "T")
            CabSql = CabSql & "," & DBSet(Rs!codpais, "T")
            
            
            CabSql = CabSql & ")"
            
            Conn.Execute CabSql
            
            
            ' insertamos en la linea de factura
            LinSql = "insert into fvarfactura_lineas (numserie, numfactu, fecfactu, NumLinea, codConce, ampliaci, precio, cantidad, Importe, TipoIva) values "
            LinSql = LinSql & "(" & DBSet(txtCodigo(6), "T")
            LinSql = LinSql & "," & DBSet(NumFact, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(7).Text, "F")
            LinSql = LinSql & ",1"
            LinSql = LinSql & "," & DBSet(txtCodigo(0).Text, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(10).Text, "T")
            LinSql = LinSql & "," & DBSet(txtCodigo(2).Text, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(1).Text, "N")
            LinSql = LinSql & "," & DBSet(txtCodigo(3).Text, "N")
            LinSql = LinSql & "," & DBSet(TipoIva, "N")
            LinSql = LinSql & ")"
            
            Conn.Execute LinSql
        End If
                    
        
        Rs.MoveNext
    Wend
    
EContab:
    If Err.Number <> 0 Then
        NumError = Err.Number
        MensError = "Generar Facturas " '& Err.Description
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        
    End If
End Sub




Private Sub BorrarTMP()
On Error Resume Next

    Conn.Execute " DROP TABLE IF EXISTS tmpfactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CrearTMP(cadTABLA As String, CadWhere As String, Optional facturas As Boolean, Optional Telefono As Boolean) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    Sql = "CREATE TABLE tmpfactu ( "
    Sql = Sql & "ctaclien varchar(10) NOT NULL default '')"
    Conn.Execute Sql
     
    Sql = "SELECT codmacta "
    Sql = Sql & " FROM " & cadTABLA
    Sql = Sql & " WHERE " & CadWhere
    Sql = " INSERT INTO tmpfactu " & Sql
    Conn.Execute Sql

    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpfactu;"
        Conn.Execute Sql
    End If
End Function


