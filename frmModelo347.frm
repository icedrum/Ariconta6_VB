VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModelo347 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3285
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtFecha 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   810
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1230
         Width           =   1305
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
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   27
         Top             =   510
         Width           =   960
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
         Left            =   480
         TabIndex        =   26
         Top             =   870
         Width           =   690
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
         Left            =   480
         TabIndex        =   25
         Top             =   1230
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1170
         Picture         =   "frmModelo347.frx":0000
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1170
         Picture         =   "frmModelo347.frx":008B
         Top             =   1230
         Width           =   240
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
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         Top             =   5190
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
         Index           =   0
         Left            =   2520
         TabIndex        =   20
         Top             =   4800
         Width           =   4095
      End
   End
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
      Height          =   6015
      Left            =   7050
      TabIndex        =   22
      Top             =   0
      Width           =   4455
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   2865
         Left            =   180
         TabIndex        =   35
         Top             =   1170
         Width           =   4215
         Begin MSComctlLib.ListView ListView1 
            Height          =   2250
            Index           =   1
            Left            =   60
            TabIndex        =   36
            Top             =   510
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   3969
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   3390
            Picture         =   "frmModelo347.frx":0116
            ToolTipText     =   "Quitar al Debe"
            Top             =   120
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   3750
            Picture         =   "frmModelo347.frx":0260
            ToolTipText     =   "Puntear al Debe"
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Empresas"
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
            Index           =   15
            Left            =   30
            TabIndex        =   37
            Top             =   180
            Width           =   1110
         End
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "frmModelo347.frx":03AA
         Left            =   1920
         List            =   "frmModelo347.frx":03B4
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   4170
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   90
         TabIndex        =   29
         Top             =   5070
         Width           =   4245
         Begin VB.OptionButton OptProv 
            Caption         =   "Fecha recepci�n"
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
            Left            =   180
            TabIndex        =   32
            Top             =   330
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton OptProv 
            Caption         =   "Fecha factura"
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
            Left            =   2370
            TabIndex        =   31
            Top             =   330
            Width           =   1755
         End
      End
      Begin VB.TextBox Text347 
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
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   4650
         Width           =   1275
      End
      Begin VB.TextBox Text347 
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
         Left            =   210
         TabIndex        =   2
         Top             =   750
         Width           =   4065
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   23
         Top             =   150
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
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   3840
         TabIndex        =   38
         Top             =   4110
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda carta"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Informe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   270
         TabIndex        =   30
         Top             =   4200
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Importe L�mite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   270
         TabIndex        =   28
         Top             =   4680
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Responsable"
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
         Index           =   13
         Left            =   210
         TabIndex        =   24
         Top             =   450
         Width           =   1260
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
      TabIndex        =   7
      Top             =   6210
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
      Left            =   8760
      TabIndex        =   5
      Top             =   6210
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
      Left            =   90
      TabIndex        =   6
      Top             =   6210
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
      Left            =   60
      TabIndex        =   8
      Top             =   3360
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
         TabIndex        =   19
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   18
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "A.E.A.T."
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   60
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Label24"
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
      Index           =   30
      Left            =   1710
      TabIndex        =   34
      Top             =   6120
      Width           =   6585
   End
   Begin VB.Label Label2 
      Caption         =   "Label24"
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
      Index           =   31
      Left            =   1710
      TabIndex        =   33
      Top             =   6420
      Width           =   6585
   End
End
Attribute VB_Name = "frmModelo347"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 410


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


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmCar As frmCartas
Attribute frmCar.VB_VarHelpID = -1

Private SQL As String
Dim cad As String
Dim RC As String
Dim RS As Recordset

Dim I As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim Tablas As String

Dim Importe As Currency

Dim UltimoPeriodoLiquidacion As Boolean
Dim C2 As String



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


Private Sub cmdAccion_Click(Index As Integer)
Dim B As Boolean
Dim B2 As Boolean

Dim RS As ADODB.Recordset

Dim indRPT As String
Dim nomDocu As String


    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    Screen.MousePointer = vbHourglass
    
    'Modificacion de 26 Marzo 2007
    '------------------------------------
    'Hay una tabla auxiliar donde se guardan datos externos de 347.
    'Cuando voy a imprimir los datos pedire si de una y/o de la otra
    
    SQL = "DELETE FROM tmp347tot where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    SQL = "DELETE FROM tmp347trimestral where codusu = " & vUsu.Codigo
    Conn.Execute SQL
        
    
    Set miRsAux = New ADODB.Recordset
    
    'El de siempre
    B = ComprobarCuentas347_
    If Not B Then Exit Sub
    
    
    'Cobros efectivo
    'Updatearemos a cero los metalicos que no llegen al minimo
    SQL = "Select ImporteMaxEfec340 from parametros "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = DBLet(miRsAux!ImporteMaxEfec340, "N")
    miRsAux.Close
    If Val(SQL) > 0 Then
        SQL = "UPDATE tmp347trimestral set metalico=0  WHERE codusu = " & vUsu.Codigo & " AND metalico < " & TransformaComasPuntos(SQL)
         Conn.Execute SQL
    End If
     
     
    'Ahora borramos todas las entrdas k no superan el importe limite
    Label2(31).Caption = "Comprobar importes"
    Label2(31).Refresh
    Importe = ImporteFormateado(Text347(1).Text)
    SQL = "Delete from tmp347 where codusu = " & vUsu.Codigo & " AND Importe  <" & TransformaComasPuntos(CStr(Importe))
    Conn.Execute SQL
    
    
    'Comprobare si hay datos
    'Comprobamos si hay datos
    SQL = "Select count(*) FROM tmp347 where codusu = " & vUsu.Codigo
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            CONT = miRsAux.Fields(0)
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If CONT = 0 Then
        MsgBox "No se ha devuelto ningun dato", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Precomprobacion de NIFs
    If Not ComprobarNifs347 Then Exit Sub
    
    
    Label2(31).Caption = ""
    Label2(30).Caption = ""
    DoEvents
    Screen.MousePointer = vbDefault
    

    
    If B Then
        If optTipoSal(1).Value Then
            'Si es impresion y el numero de registros es superior a 25 no
            'puede imprimirse
            CONT = 0
            SQL = ""
                
            Set RS = New ADODB.Recordset
            SQL = "Select count(*) from tmp347tot where codusu =" & vUsu.Codigo
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then CONT = DBLet(RS.Fields(0), "N")
            RS.Close
            Set RS = Nothing
            If CONT > 25 Then
                SQL = "El numero de registros supera los 25." & vbCrLf & _
                    "Se debe presentar por diskette o via internet."
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
            
            'Modelo de hacienda
            B2 = Modelo347(False, Year(CDate(txtFecha(1).Text)))
            
            If B2 Then CopiarFicheroASalida False, txtTipoSalida(1).Text
        
        Else
            If optTipoSal(2).Value Or optTipoSal(3).Value Then
                ExportarPDF = True 'generaremos el pdf
            Else
                ExportarPDF = False
            End If
            SoloImprimir = False
            If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
            Select Case Combo5.ListIndex
            Case 0
                'La carta
                cad = "� Desea imprimir tambi�n los proveedores ?"
                If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then
                    cad = " AND {tmp347tot.cliprov} = " & Asc(0)
                    cadFormula = cadFormula & cad
                Else
                    cad = ""
                End If

                cadFormula = cadFormula & "{tmp347tot.codusu} = " & vUsu.Codigo
                cadFormula = cadFormula & " and {cartas.codcarta} = 999 "

                indRPT = "0410-01"
                
                If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
                
                cadNomRPT = nomDocu '"Carta.rpt"
                
                
                ImprimeGeneral
                
            Case Else
            
                'LISTADO
                '-----------------------------------------------------------------
                cadParam = cadParam & "Fechas= """ & "Desde " & txtFecha(0).Text & "      hasta  " & txtFecha(1).Text & """|"
                numParam = numParam + 1
                    
            
                cadFormula = "{tmp347tot.codusu} = " & vUsu.Codigo

                indRPT = "0410-00"
                
                If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
                
                cadNomRPT = nomDocu '"Carta.rpt"
                
                ImprimeGeneral
            
                
            End Select
            
            If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
            If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 21
                
            If SoloImprimir Or ExportarPDF Then Unload Me
            Screen.MousePointer = vbDefault
            
            
            
        End If
    End If
    
    
    
    
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub Combo5_Validate(Cancel As Boolean)
    optTipoSal(3).Enabled = (Combo5.ListIndex = 1)
    If Not optTipoSal(3).Enabled Then optTipoSal(0).Value = True
    
    Me.Toolbar1.Buttons(1).Enabled = (Combo5.ListIndex = 0)
    
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
    Me.Icon = frmPpal.Icon
        
    'Otras opciones
    Me.Caption = "Modelo 347"

    ' boton al mto de cartas
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4
    End With
     
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    
     
    txtFecha(0).Text = "01/01/" & Year(vParam.fechaini)
    txtFecha(1).Text = "31/12/" & Year(vParam.fechaini)
    Text347(1).Text = Format(vParam.limimpcl, FormatoImporte)
    Label2(30).Caption = ""
    Label2(31).Caption = ""
    
    Combo5.ListIndex = 1
    Toolbar1.Buttons(1).Enabled = False
     
    CargarListView 1
    
    FrameSeccion.Enabled = vParam.EsMultiseccion
    
    optTipoSal(3).Enabled = (Combo5.ListIndex = 1)
    
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    PonerDatosFicheroSalida
End Sub

Private Sub PonerDatosFicheroSalida()
    
    txtTipoSalida(1).Text = App.Path & "\Exportar\Mod347.txt"

End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim I As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' tabla de codigos de iva
        Case 0
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = False
            Next I
        Case 1
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = True
            Next I
    End Select
    
    Screen.MousePointer = vbDefault


End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Set frmCar = New frmCartas
            
            frmCar.CodigoActual = 999
            frmCar.Desde347 = True
            frmCar.Show vbModal
    
            Set frmCar = Nothing
    
    End Select

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtPag2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
End Sub


Private Sub AccionesCSV()
Dim SQL2 As String

    'Monto el SQL
    SQL = ""
    
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0410-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "FacturasCliFecha.rpt"

    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "Empresas= """
    For I = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            cadParam = cadParam & Me.ListView1(1).ListItems(I).SubItems(1) & "  "
        End If
    Next I
    cadParam = Trim(cadParam)
    
    cadParam = cadParam & """|"
    
    
    cadFormula = "{tmp347.codusu}=" & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 21
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Function CargarTemporal() As Boolean
Dim SQL As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    SQL = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "insert into tmpfaclin (codusu, codigo, numserie, nomserie, numfac, fecha, cta, cliente, nif, imponible, impiva, total, retencion,"
    SQL = SQL & " recargo, tipoopera, tipoformapago) "
    SQL = SQL & " select distinct " & vUsu.Codigo & ",0, factcli.numserie, contadores.nomregis, factcli.numfactu, factcli.fecfactu, factcli.codmacta, "
    SQL = SQL & " factcli.nommacta, factcli.nifdatos, factcli.totbases, factcli.totivas, factcli.totfaccl, factcli.trefaccl, "
    SQL = SQL & " factcli.totrecargo, tipofpago.descformapago , aa.denominacion"
    SQL = SQL & " from " & tabla
    SQL = SQL & " where " & cadselect
    
    Conn.Execute SQL
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
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
        
        LanzaFormAyuda txtFecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        MsgBox "Introduce las fechas de consulta.", vbExclamation
        Exit Function
    End If

    If Not ComprobarFechas(0, 1) Then Exit Function
    
    
    If Year(CDate(txtFecha(0).Text)) <> Year(CDate(txtFecha(1).Text)) Then
        MsgBox "Esta abarcando dos a�os. Se considera el a�o: " & Year(CDate(txtFecha(1).Text)), vbExclamation
    End If
    If Combo5.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de informe.", vbExclamation
        Exit Function
    End If
    
    
    If Combo5.ListIndex = 0 And Text347(0).Text = "" Then
        MsgBox "Escriba el nombre del responsable.", vbExclamation
        Exit Function
    End If
            
    
    If Combo5.ListIndex = 2 Then 'antes 3
        'Enero 2012
        'Tiene que ser una a�o exacto
        If Month(CDate(txtFecha(0).Text)) <> 1 Or Month(CDate(txtFecha(0).Text)) <> 1 Then
            MsgBox "A�o natural. Enero diciembre", vbExclamation
            Exit Function
        End If
        If Month(CDate(txtFecha(1).Text)) <> 12 Or Day(CDate(txtFecha(1).Text)) <> 31 Then
            MsgBox "A�o natural. Hasta 31 diciembre", vbExclamation
            Exit Function
        End If
        
    End If
    
    If Text347(1).Text = "" Then
        MsgBox "Importe limite en blanco", vbExclamation
        Exit Function
    End If
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
    End If
    
    
    '++ comprobamos que todas las facturas tienen nif asignado
    DatosOK = ComprobarNifFacturas
    
    If DatosOK Then DatosOK = ComprobarCPostalFacturas
    
   
End Function


Private Function ComprobarNifFacturas() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim CadResul As String

    ComprobarNifFacturas = False

    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            
            'facturas de clientes
            SQL = "select distinct factcli.codmacta from ariconta" & Me.ListView1(1).ListItems(I).Text & ".factcli, ariconta" & Me.ListView1(1).ListItems(I).Text & ".cuentas where "
            SQL = SQL & " cuentas.codmacta=factcli.codmacta and model347=1 "
            SQL = SQL & " AND fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            SQL = SQL & " AND fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            SQL = SQL & " and (factcli.nifdatos is null or factcli.nifdatos = '')"
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            CadResul = ""
            
            While Not RS.EOF
                CadResul = CadResul & DBLet(RS!codmacta, "T") & ","
                RS.MoveNext
            Wend
            
            If CadResul <> "" Then
                CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
                CadResul = Me.ListView1(1).ListItems(I).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de cliente sin nif de las cuentas: " & vbCrLf & vbCrLf & CadResul
                                
                MsgBox CadResul, vbExclamation
                
                Set RS = Nothing
                Exit Function
            End If
            Set RS = Nothing
        
            If OptProv(0).Value Then
                cad = "fecharec"
            Else
                If OptProv(1).Value Then
                    cad = "fecfactu"
                End If
            End If
            
            ' facturas de proveedores
            SQL = "SELECT distinct factpro.codmacta from ariconta" & Me.ListView1(1).ListItems(I).Text & ".factpro, ariconta" & Me.ListView1(1).ListItems(I).Text & ".cuentas  where "
            SQL = SQL & " cuentas.codmacta=factpro.codmacta and model347=1 "
            SQL = SQL & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            SQL = SQL & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            SQL = SQL & " and (factpro.nifdatos is null or factpro.nifdatos = '')"
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS.EOF
                CadResul = CadResul & DBLet(RS!codmacta, "T") & ","
                RS.MoveNext
            Wend
            If CadResul <> "" Then
                CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
                CadResul = Me.ListView1(1).ListItems(I).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de proveedor sin nif de las cuentas: " & vbCrLf & vbCrLf & CadResul
                                
                MsgBox CadResul, vbExclamation
                
                Set RS = Nothing
                Exit Function
            End If
            Set RS = Nothing
        
        
        End If
    Next I
    
    ComprobarNifFacturas = True
  
    

End Function


Private Function ComprobarCPostalFacturas() As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim CadResul As String

    ComprobarCPostalFacturas = False

    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            
            'facturas de clientes
            SQL = "select distinct factcli.codmacta from ariconta" & Me.ListView1(1).ListItems(I).Text & ".factcli, ariconta" & Me.ListView1(1).ListItems(I).Text & ".cuentas where "
            SQL = SQL & " cuentas.codmacta=factcli.codmacta and model347=1 "
            SQL = SQL & " AND fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            SQL = SQL & " AND fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            SQL = SQL & " and (factcli.codpobla is null or factcli.codpobla = '')"
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            CadResul = ""
            
            While Not RS.EOF
                CadResul = CadResul & DBLet(RS!codmacta, "T") & ","
                RS.MoveNext
            Wend
            
            If CadResul <> "" Then
                CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
                CadResul = Me.ListView1(1).ListItems(I).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de cliente sin c�digo postal de las cuentas: " & vbCrLf & vbCrLf & CadResul
                                
                MsgBox CadResul, vbExclamation
                
                Set RS = Nothing
                If MsgBox("� Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
            End If
            Set RS = Nothing
        
            If OptProv(0).Value Then
                cad = "fecharec"
            Else
                If OptProv(1).Value Then
                    cad = "fecfactu"
                End If
            End If
            
            ' facturas de proveedores
            SQL = "SELECT distinct factpro.codmacta from ariconta" & Me.ListView1(1).ListItems(I).Text & ".factpro, ariconta" & Me.ListView1(1).ListItems(I).Text & ".cuentas  where "
            SQL = SQL & " cuentas.codmacta=factpro.codmacta and model347=1 "
            SQL = SQL & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            SQL = SQL & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            SQL = SQL & " and (factpro.codpobla is null or factpro.codpobla = '')"
            
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS.EOF
                CadResul = CadResul & DBLet(RS!codmacta, "T") & ","
                RS.MoveNext
            Wend
            If CadResul <> "" Then
                CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
                CadResul = Me.ListView1(1).ListItems(I).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de proveedor sin c�digo postal de las cuentas: " & vbCrLf & vbCrLf & CadResul
                                
                MsgBox CadResul, vbExclamation
                
                Set RS = Nothing
                If MsgBox("� Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
            End If
            Set RS = Nothing
        
        
        End If
    Next I
    
    ComprobarCPostalFacturas = True
  
    

End Function





Private Function EmpresasSeleccionadas() As Integer
Dim SQL As String
Dim I As Integer
Dim NSel As Integer

    NSel = 0
    For I = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then NSel = NSel + 1
    Next I
    EmpresasSeleccionadas = NSel

End Function


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "C�digo", 600
    ListView1(Index).ColumnHeaders.Add , , "Descripci�n", 3200
    
    SQL = "SELECT codempre, nomempre, conta "
    SQL = SQL & " FROM usuarios.empresasariconta "
    
    If Not vParam.EsMultiseccion Then
        SQL = SQL & " where conta = " & DBSet(Conn.DefaultDatabase, "T")
    Else
        SQL = SQL & " where mid(conta,1,8) = 'ariconta'"
    End If
    SQL = SQL & " ORDER BY codempre "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        
        If vParam.EsMultiseccion Then
            If EsMultiseccion(DBLet(RS!CONTA)) Then
                Set ItmX = ListView1(Index).ListItems.Add
                
                If DBLet(RS!CONTA) = Conn.DefaultDatabase Then ItmX.Checked = True
                ItmX.Text = RS.Fields(0).Value
                ItmX.SubItems(1) = RS.Fields(1).Value
            End If
        Else
            Set ItmX = ListView1(Index).ListItems.Add
            
            ItmX.Checked = True
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = RS.Fields(1).Value
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Empresas.", Err.Description
    End If
End Sub

Private Function ComprobarCuentas347_() As Boolean
Dim I As Integer
Dim I1 As Currency
Dim I2 As Currency
Dim i3 As Currency
Dim i4 As Currency
Dim I5 As Currency
Dim PAIS As String
    ComprobarCuentas347_ = False
    
    'Esto sera para las inserciones de despues
    Tablas = "INSERT INTO tmp347tot (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla,Provincia,pais) "
    Tablas = Tablas & " VALUES (" & vUsu.Codigo
         

    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            Label2(30).Caption = Me.ListView1(1).ListItems(I).SubItems(1)
            Label2(31).Caption = "Comprobar Cuentas"
            Me.Refresh
            If Not ComprobarCuentas347_DOS("ariconta" & Me.ListView1(1).ListItems(I).Text, Me.ListView1(1).ListItems(I).SubItems(1)) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
        
           'Iremos NIF POR NIF
           
              Label2(31).Caption = "Insertando datos tmp(I)"
              Label2(31).Refresh
              SQL = "SELECT  cliprov,nif, sum(importe) as suma, razosoci,dirdatos,codposta,"
              SQL = SQL & "despobla,provincia,pais from ariconta" & Me.ListView1(1).ListItems(I).Text & ".tmp347 where codusu=" & vUsu.Codigo
              SQL = SQL & " group by cliprov, nif"
              
              Set RS = New ADODB.Recordset
              RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
              
              While Not RS.EOF
                   Label2(31).Caption = RS!NIF
                   Label2(31).Refresh
                   If ExisteEntrada Then
                        Importe = Importe + RS!Suma
                        SQL = "UPDATE tmp347tot SET importe=" & TransformaComasPuntos(CStr(Importe))
                        SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & RS!cliprov
                        SQL = SQL & " AND nif = '" & RS!NIF & "';"
                   Else
                        'Nuevo para lo de las agencias de viajes
                        SQL = "," & RS!cliprov & ",'" & RS!NIF & "'," & TransformaComasPuntos(CStr(RS!Suma))
                        SQL = SQL & ",'" & DevNombreSQL(DBLet(RS!razosoci, "T")) & "','" & DevNombreSQL(DBLet(RS!dirdatos)) & "','" & DBLet(RS!codposta, "T") & "','"
                        SQL = SQL & DevNombreSQL(DBLet(RS!desPobla, "T")) & "','" & DevNombreSQL(DBLet(RS!provincia, "T"))
                        If DBLet(RS!PAIS, "T") = "" Then
                            PAIS = "ES"
                        Else
                            PAIS = RS!PAIS
                        End If
                        SQL = SQL & "','" & DevNombreSQL(DBLet(PAIS, "T")) & "')"
                        SQL = Tablas & SQL
                   End If
                   Conn.Execute SQL
                   RS.MoveNext
              Wend
              RS.Close
              
              
              'trimestral
              Label2(31).Caption = "Insertando datos tmp(II)"
              Label2(31).Refresh
              SQL = "SELECT  tmp347trimestre.cliprov,tmp347.nif,tmp347.codposta, sum(trim1) as t1, sum(trim2) as t2,"
              SQL = SQL & " sum(trim3) as t3, sum(trim4) as t4,sum(metalico) as metalico"
              SQL = SQL & " from ariconta" & Me.ListView1(1).ListItems(I).Text & ".tmp347,ariconta" & Me.ListView1(1).ListItems(I).Text & ".tmp347trimestre where tmp347.codusu=" & vUsu.Codigo
              SQL = SQL & " and tmp347.codusu=tmp347trimestre.codusu"
              SQL = SQL & " and tmp347.cliprov=tmp347trimestre.cliprov"
              SQL = SQL & " and tmp347.cta=tmp347trimestre.cta "
              SQL = SQL & " group by tmp347.cliprov,tmp347.nif"
              
              Set RS = New ADODB.Recordset
              RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
              
              While Not RS.EOF
                   Label2(31).Caption = RS!NIF
                   Label2(31).Refresh
                   If ExisteEntradaTrimestral(I1, I2, i3, i4, I5) Then
                        I1 = I1 + RS!T1
                        I2 = I2 + RS!t2
                        i3 = i3 + RS!T3
                        i4 = i4 + RS!T4
                        I5 = I5 + RS!metalico
                        SQL = "UPDATE tmp347trimestral SET "
                        SQL = SQL & " trim1=" & TransformaComasPuntos(CStr(I1))
                        SQL = SQL & ", trim2=" & TransformaComasPuntos(CStr(I2))
                        SQL = SQL & ", trim3=" & TransformaComasPuntos(CStr(i3))
                        SQL = SQL & ", trim4=" & TransformaComasPuntos(CStr(i4))
                        SQL = SQL & ", metalico=" & TransformaComasPuntos(CStr(I5))
                        SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & RS!cliprov
                        SQL = SQL & " AND nif = '" & RS!NIF & "';"
                   Else
                        
                        SQL = "insert into tmp347trimestral (`codusu`,`cliprov`,`nif`,`trim1`,`trim2`"
                        SQL = SQL & ",`trim3`,`trim4`,`codposta`,metalico) values ( " & vUsu.Codigo
                        SQL = SQL & "," & RS!cliprov & ",'" & RS!NIF & "',"
                        SQL = SQL & TransformaComasPuntos(CStr(RS!T1)) & "," & TransformaComasPuntos(CStr(RS!t2)) & ","
                        SQL = SQL & TransformaComasPuntos(CStr(RS!T3)) & "," & TransformaComasPuntos(CStr(RS!T4)) & "," & DBSet(RS!codposta, "T") & ","
                        SQL = SQL & TransformaComasPuntos(CStr(RS!metalico)) & ")"
    
                   End If
                   Conn.Execute SQL
                   RS.MoveNext
              Wend
              RS.Close
              
              
              
              
              
              espera 0.5
         End If
    Next I
    ComprobarCuentas347_ = True
    
End Function



Private Sub CopiarFicheroHacienda(Modelo347 As Boolean)
    On Error GoTo ECopiarFichero347
    MsgBox "El archivo se ha generado con exito.", vbInformation
    SQL = ""
    cd1.CancelError = True
    cd1.ShowSave
    If Modelo347 Then
        SQL = App.Path & "\Hacienda\mod347\mod347.txt"
    Else
        SQL = App.Path & "\Hacienda\mod349\mod349.txt"
    End If
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El fichero ya existe. �Reemplazar?", vbQuestion + vbYesNo) = vbNo Then SQL = ""
        End If
        If SQL <> "" Then
            FileCopy SQL, cd1.FileName
            MsgBox Space(20) & "Copia efectuada correctamente" & Space(20), vbInformation
        End If
    End If
    Exit Sub
ECopiarFichero347:
    If Err.Number <> 32755 Then MuestraError Err.Number, "Copiar fichero 347"
    
End Sub

Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If txtFecha(Indice1).Text <> "" And txtFecha(Indice2).Text <> "" Then
        If CDate(txtFecha(Indice1).Text) > CDate(txtFecha(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function

Private Function ComprobarCuentas347_DOS(Contabilidad As String, Empresa As String) As Boolean
Dim SQL2 As String
Dim SqlTot As String
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim RsTot As ADODB.Recordset
Dim I1 As Currency
Dim I2 As Currency
Dim i3 As Currency
Dim Trimestre(3) As Currency
Dim Impor As Currency
Dim Tri As Byte

On Error GoTo EComprobarCuentas347
    ComprobarCuentas347_DOS = False
    'Utilizaremos la tabla tmpcierre1, prestada
    SQL = "DELETE FROM " & Contabilidad & ".tmp347 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    'Enero 2012
    'Calcular por trimiestre
    SQL = "DELETE FROM " & Contabilidad & ".tmp347trimestre where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    Set RS = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    
    'Para lo nuevo. Iremos codmacta a codmacta

   
    SQL = " Select factcli.codmacta,factcli.nifdatos,factcli.dirdatos,factcli.codpobla,factcli.nommacta,factcli.despobla,factcli.desprovi,factcli.codpais from "
    SQL = SQL & Contabilidad & ".factcli, " & Contabilidad & ".cuentas  where "
    SQL = SQL & " cuentas.codmacta=factcli.codmacta and model347=1 "
    SQL = SQL & " AND fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    SQL = SQL & " group by codmacta, factcli.nifdatos "
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
    
        Label2(31).Caption = RS!codmacta
        Label2(31).Refresh
        
        Trimestre(0) = 0: Trimestre(1) = 0: Trimestre(2) = 0: Trimestre(3) = 0
        
        
        SQL = "Select * from " & Contabilidad & ".factcli where codmacta = '" & RS.Fields(0) & "'"
        SQL = SQL & " and factcli.nifdatos = " & DBSet(RS.Fields(1).Value, "T")
        SQL = SQL & " AND fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        SQL = SQL & " AND fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I1 = 0
        I2 = 0
        While Not RT.EOF
            Set RsTot = New ADODB.Recordset
            
            SqlTot = "select sum(coalesce(baseimpo,0)) base, sum(coalesce(impoiva,0)) iva, sum(coalesce(imporec,0)) recargo "
            SqlTot = SqlTot & " from " & Contabilidad & ".factcli_totales "
            SqlTot = SqlTot & " where numserie = " & DBSet(RT!NUmSerie, "T")
            SqlTot = SqlTot & " and numfactu = " & DBSet(RT!NumFactu, "N")
            SqlTot = SqlTot & " and fecfactu = " & DBSet(RT!FecFactu, "F")
            Set RsTot = New ADODB.Recordset
            RsTot.Open SqlTot, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsTot.EOF Then
                I1 = I1 + DBLet(RsTot!Base, "N")
                I2 = I2 + DBLet(RsTot!iva, "N") + DBLet(RsTot!recargo, "N")
                Impor = DBLet(RsTot!Base, "N") + DBLet(RsTot!iva, "N") + DBLet(RsTot!recargo, "N")
            End If
            
            Set RsTot = Nothing
            
            
            'El trimestre
            Tri = QueTrimestre(RT!fecliqcl)
            Tri = Tri - 1
            Trimestre(Tri) = Trimestre(Tri) + Impor
            
            RT.MoveNext
        Wend
        RT.Close
        
        'El importe final es la suma de las bases mas los ivas
        I1 = I1 + I2
        SQL = "INSERT INTO " & Contabilidad & ".tmp347 (codusu, cliprov, cta, nif, codposta, importe, razosoci, dirdatos, despobla, provincia, pais )  "
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("0") & ",'" & RS!codmacta & "','"
        SQL = SQL & DBLet(RS!nifdatos) & "'," & DBSet(RS!CodPobla, "T") & "," & TransformaComasPuntos(CStr(I1))
        SQL = SQL & "," & DBSet(RS!Nommacta, "T") & "," & DBSet(RS!dirdatos, "T") & "," & DBSet(RS!desPobla, "T") & "," & DBSet(RS!desProvi, "T") & "," & DBSet(RS!codPAIS, "T") & ")"
        Conn.Execute SQL
        
       
        'El del trimestre
        SQL = "insert into " & Contabilidad & ".`tmp347trimestre` (`codusu`,`cliprov`,`cta`,`nif`,`codposta`,`trim1`,`trim2`,`trim3`,`trim4`)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("0") & ",'" & RS!codmacta & "'," & DBSet(RS!nifdatos, "T") & "," & DBSet(RS!CodPobla, "T")
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(0))) & "," & TransformaComasPuntos(CStr(Trimestre(1)))
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(2))) & "," & TransformaComasPuntos(CStr(Trimestre(3))) & ")"
        Conn.Execute SQL

     
        
        RS.MoveNext
    Wend
    RS.Close
    If OptProv(0).Value Then
        cad = "fecharec"
    Else
        If OptProv(1).Value Then
            cad = "fecfactu"
'        Else
'            Cad = "fecliqpr"
        End If
    End If
    
    Label2(31).Caption = "Comprobando datos facturas proveedor"
    DoEvents
    espera 0.2
    
    
    SQL = "SELECT factpro.codmacta,factpro.nifdatos, factpro.codpobla, factpro.dirdatos, factpro.nommacta,factpro.despobla,factpro.desprovi,factpro.codpais from " & Contabilidad & ".factpro," & Contabilidad & ".cuentas  where "
    SQL = SQL & Contabilidad & ".cuentas.codmacta=" & Contabilidad & ".factpro.codmacta and model347=1 "
    SQL = SQL & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    SQL = SQL & " group by factpro.codmacta,factpro.nifdatos "
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Label2(31).Caption = RS.Fields(0)
        Label2(31).Refresh
        DoEvents
        SQL = "Select factpro.*," & cad & " fecha from " & Contabilidad & ".factpro factpro where codmacta = '" & RS.Fields(0) & "'"
        SQL = SQL & " AND nifdatos = " & DBSet(RS!nifdatos, "T")
        SQL = SQL & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        SQL = SQL & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I1 = 0
        I2 = 0
        Trimestre(0) = 0: Trimestre(1) = 0: Trimestre(2) = 0: Trimestre(3) = 0
        While Not RT.EOF
                
            Set RsTot = New ADODB.Recordset
            
            Impor = 0
                
            SqlTot = "select sum(coalesce(baseimpo,0)) base, sum(coalesce(impoiva,0)) iva, sum(coalesce(imporec,0)) recargo "
            SqlTot = SqlTot & " from " & Contabilidad & ".factpro_totales "
            SqlTot = SqlTot & " where numserie = " & DBSet(RT!NUmSerie, "T")
            SqlTot = SqlTot & " and numregis = " & DBSet(RT!NumRegis, "N")
            SqlTot = SqlTot & " and anofactu = " & DBSet(RT!anofactu, "N")
            Set RsTot = New ADODB.Recordset
            RsTot.Open SqlTot, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RsTot.EOF Then
                I1 = I1 + DBLet(RsTot!Base, "N")
                I2 = I2 + DBLet(RsTot!iva, "N") + DBLet(RsTot!recargo, "N")
                Impor = DBLet(RsTot!Base, "N") + DBLet(RsTot!iva, "N") + DBLet(RsTot!recargo, "N")
            End If
            
            Set RsTot = Nothing
                
            'El trimestre
            Tri = QueTrimestre(RT!Fecha)
            Tri = Tri - 1
            Trimestre(Tri) = Trimestre(Tri) + Impor
            
            
            RT.MoveNext
        Wend
        RT.Close
        
        'El importe final es la suma de las bases mas los ivas
        I1 = I1 + I2
        SQL = "INSERT INTO " & Contabilidad & ".tmp347 (codusu, cliprov, cta, nif, codposta, importe, razosoci, dirdatos, despobla, provincia, pais)  "
        'SQL = SQL & " VALUES (" & vUsu.Codigo & ",1,'" & RS!Codmacta & "','"
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("1") & ",'" & RS!codmacta & "','"
        SQL = SQL & DBLet(RS!nifdatos) & "'," & DBSet(RS!CodPobla, "T") & "," & TransformaComasPuntos(CStr(I1))
        SQL = SQL & "," & DBSet(RS!Nommacta, "T") & "," & DBSet(RS!dirdatos, "T") & "," & DBSet(RS!desPobla, "T") & "," & DBSet(RS!desProvi, "T") & "," & DBSet(RS!codPAIS, "T") & ")"
        Conn.Execute SQL
        
        
        'El del trimestre
        SQL = "insert into " & Contabilidad & ".`tmp347trimestre` (`codusu`,`cliprov`,`cta`,`nif`,`codposta`,`trim1`,`trim2`,`trim3`,`trim4`)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & "," & Asc("1") & ",'" & RS!codmacta & "'," & DBSet(RS!nifdatos, "T") & "," & DBSet(RS!CodPobla, "T")
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(0))) & "," & TransformaComasPuntos(CStr(Trimestre(1)))
        SQL = SQL & "," & TransformaComasPuntos(CStr(Trimestre(2))) & "," & TransformaComasPuntos(CStr(Trimestre(3))) & ")"
        Conn.Execute SQL
        
        
        RS.MoveNext
        
    Wend
    RS.Close
    
    
    
    
    'DICIEMBRE 2012
    ' CObros en metalico superiores a 6000
    Label2(31).Caption = "Cobros metalico"
    Label2(31).Refresh
    DoEvents
    SQL = "Select ImporteMaxEfec340 from " & Contabilidad & ".parametros "
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO pues ser eof
    I1 = DBLet(RS!ImporteMaxEfec340, "N")
    RS.Close
    If I1 > 0 Then
        'SI que lleva control de cobros en efectivo
        'Veremos si hay conceptos de efectivo
        SQL = "Select codconce from " & Contabilidad & ".conceptos where EsEfectivo340 = 1"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not RS.EOF
            SQL = SQL & ", " & RS!codconce
            RS.MoveNext
        Wend
        RS.Close
        SQL2 = "" 'Errores en Datos en efectivo sin ventas
        If SQL <> "" Then
            SQL = Mid(SQL, 2) 'quit la coma
            
            cad = "Select * from " & Contabilidad & ".tmp347trimestre WHERE codusu = " & vUsu.Codigo & " ORDER BY cta"
            RT.Open cad, Conn, adOpenKeyset, adCmdText
            
            'HABER -DEBE"
            cad = "Select hlinapu.codmacta,sum(if(timporteh is null,0,timporteh))-sum(if(timported is null,0,timported)) importe, cuentas.nifdatos, cuentas.codposta"
            cad = cad & " from " & Contabilidad & ".hlinapu,cuentas WHERE hlinapu.codmacta =cuentas.codmacta "
            cad = cad & " AND model347=1 AND fechaent >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            cad = cad & " AND fechaent <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            cad = cad & " AND codconce IN (" & SQL & ")"
            cad = cad & " group by 1 order by 1"

            RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not RS.EOF
                Label2(31).Caption = RS!codmacta
                Label2(31).Refresh
        
                If RS!Importe <> 0 Then
                    SQL = "cta  = '" & RS!codmacta & "'"
                    RT.Find SQL, , adSearchForward, 1
                    
                    If RT.EOF Then
                        SQL2 = SQL2 & RS!codmacta & " (" & RS!Importe & ") " & vbCrLf
                    Else
                        SQL = "UPDATE " & Contabilidad & ".tmp347trimestre SET metalico = " & TransformaComasPuntos(CStr(RS!Importe))
                        SQL = SQL & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & RT!Cta & "'"
                        '++
                        SQL = SQL & " and nifdatos = " & DBSet(RS!nifdatos, "T")
                        Conn.Execute SQL
                    End If
                End If
                RS.MoveNext
            Wend
            RS.Close
            RT.Close
            
            If SQL2 <> "" Then
                SQL2 = "Cobros en efectivo sin asociar a ventas" & vbCrLf & SQL2
                MsgBox SQL2, vbExclamation
            End If
        End If
    End If
    
    Set RT = Nothing
    RC = ""
    cad = ""
    SQL2 = ""
    'Comprobaremos k el nif no es nulo, ni el codppos de las cuentas a tratar
    SQL = "Select cta from " & Contabilidad & ".tmp347 where (nif is null or nif = '') and codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        cad = cad & RS.Fields(0) & "       "
        If I = 3 Then
            cad = cad & vbCrLf
            I = 0
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    If cad <> "" Then
        RC = "Cuentas con NIF sin valor: " & vbCrLf & vbCrLf & cad
        cad = ""
    End If
    
    'Comprobamos el codpos
    SQL = "Select cta,razosoci,codposta from " & Contabilidad & ".tmp347 where codusu = " & vUsu.Codigo
    SQL = SQL & " AND (codposta is null or codposta='')"

    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        cad = cad & RS.Fields(0) & "       "
        If I = 3 Then
            cad = cad & vbCrLf
            I = 0
        End If
        RS.MoveNext
    Wend
    RS.Close
    
    If cad <> "" Then
        If RC <> "" Then RC = RC & vbCrLf & vbCrLf & vbCrLf
        RC = RC & "Cuentas con codigo postal sin valor: " & vbCrLf & vbCrLf & cad
    End If
    
    If RC <> "" Then
        RC = "Empresa: " & Empresa & vbCrLf & vbCrLf & RC & vbCrLf & " Desea continuar igualmente?"
        If MsgBox(RC, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    
    Set RS = Nothing
    
    ComprobarCuentas347_DOS = True
    Exit Function
EComprobarCuentas347:
    MuestraError Err.Number, "Comprobar Cuentas 347" & vbCrLf & vbCrLf & SQL & vbCrLf
End Function

Private Function ExisteEntrada() As Boolean
    SQL = "Select importe from tmp347tot  where codusu = " & vUsu.Codigo & " and cliprov =" & RS!cliprov & " AND nif ='" & RS!NIF & "'"
    SQL = SQL & " and codposta = " & DBSet(RS!codposta, "T") & ";"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        ExisteEntrada = True
        Importe = miRsAux!Importe
    Else
        ExisteEntrada = False
    End If
    miRsAux.Close
End Function

Private Function ExisteEntradaTrimestral(ByRef I1 As Currency, ByRef I2 As Currency, ByRef i3 As Currency, ByRef i4 As Currency, ByRef I5 As Currency) As Boolean
    SQL = "Select trim1,trim2,trim3,trim4,metalico from tmp347trimestral  where codusu = " & vUsu.Codigo & " and cliprov =" & RS!cliprov & " AND nif ='" & RS!NIF & "' and codposta = " & DBSet(RS!NIF, "T") & ";"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ExisteEntradaTrimestral = True
        I1 = miRsAux!trim1
        I2 = miRsAux!trim2
        i3 = miRsAux!trim3
        i4 = miRsAux!trim4
        I5 = DBLet(miRsAux!metalico, "N")
    Else
        ExisteEntradaTrimestral = False
        I1 = 0: I2 = 0: i3 = 0: i4 = 0: I5 = 0
    End If
    miRsAux.Close
End Function

'Dada una fecha me da el trimestre
Private Function QueTrimestre(Fecha As Date) As Byte
Dim C As Byte
    
        C = Month(Fecha)
        If C < 4 Then
            QueTrimestre = 1
        ElseIf C < 7 Then
            QueTrimestre = 2
        ElseIf C < 10 Then
            QueTrimestre = 3
        Else
            QueTrimestre = 4
        End If
    
End Function

