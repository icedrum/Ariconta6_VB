VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModelo303 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11685
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
      Height          =   2715
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   6915
      Begin VB.Frame FramePeriodo 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   90
         TabIndex        =   34
         Top             =   1290
         Width           =   3675
         Begin VB.TextBox txtperiodo 
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
            Left            =   960
            TabIndex        =   1
            Top             =   150
            Width           =   675
         End
         Begin VB.TextBox txtperiodo 
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
            Left            =   2670
            TabIndex        =   2
            Top             =   150
            Width           =   645
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
            Index           =   26
            Left            =   270
            TabIndex        =   36
            Top             =   150
            Width           =   870
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
            Index           =   27
            Left            =   2220
            TabIndex        =   35
            Top             =   165
            Width           =   390
         End
      End
      Begin VB.ComboBox cmbPeriodo 
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
         ItemData        =   "frmModelo303.frx":0000
         Left            =   330
         List            =   "frmModelo303.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   930
         Width           =   3330
      End
      Begin VB.TextBox txtAno 
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
         Index           =   0
         Left            =   3960
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   930
         Width           =   765
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   4800
         Width           =   4095
      End
      Begin VB.Label lblFecha1 
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
         Left            =   2580
         TabIndex        =   24
         Top             =   3990
         Width           =   4095
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
         Left            =   2580
         TabIndex        =   23
         Top             =   3630
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Per�odo"
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
         Left            =   360
         TabIndex        =   22
         Top             =   570
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "A�o"
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
         Index           =   6
         Left            =   3990
         TabIndex        =   21
         Top             =   570
         Width           =   960
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
      Height          =   5415
      Left            =   7110
      TabIndex        =   27
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtCuota 
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
         Left            =   1650
         TabIndex        =   5
         Top             =   4890
         Width           =   2595
      End
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   180
         TabIndex        =   31
         Top             =   1020
         Width           =   4185
         Begin MSComctlLib.ListView ListView1 
            Height          =   2880
            Index           =   1
            Left            =   60
            TabIndex        =   32
            Top             =   510
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   5080
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
            Picture         =   "frmModelo303.frx":0004
            ToolTipText     =   "Quitar al Debe"
            Top             =   120
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   3750
            Picture         =   "frmModelo303.frx":014E
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
            TabIndex        =   33
            Top             =   180
            Width           =   1110
         End
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
         Index           =   2
         Left            =   1350
         TabIndex        =   4
         Top             =   570
         Width           =   1485
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   28
         Top             =   270
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
      Begin VB.Label Label3 
         Caption         =   "Cuotas a compensar per�odos anteriores"
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
         Left            =   270
         TabIndex        =   37
         Top             =   4560
         Width           =   4125
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   2
         Left            =   1020
         Picture         =   "frmModelo303.frx":0298
         Top             =   570
         Width           =   240
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
         Index           =   13
         Left            =   210
         TabIndex        =   29
         Top             =   570
         Width           =   690
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
      Left            =   10350
      TabIndex        =   8
      Top             =   5490
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
      Left            =   8790
      TabIndex        =   6
      Top             =   5490
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
      TabIndex        =   7
      Top             =   5490
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
      TabIndex        =   9
      Top             =   2760
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
         TabIndex        =   20
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   19
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   60
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
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
      Left            =   1680
      TabIndex        =   30
      Top             =   5550
      Visible         =   0   'False
      Width           =   5475
   End
End
Attribute VB_Name = "frmModelo303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 408

Public OpcionListado As Byte
    ' 0 modelo 303
    ' 1 modelo 390



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

    ' en tmpliquidaiva la columna cliente indica
    '                   0- Facturas clientes
    '                   1- RECARGO EQUIVALENCIA
    '                   10- Intracomunitarias
    '                   12- Sujeto pasivo
    '                   14- Entregas intracomunitarias (no deducibles)
    '                   16- Exportaciones y operaciones asimiladas
    '                   2- Facturas proveedores
    '                   30- Proveedores bien de inversion
    '                   32- iva de importacion de bienes corrientes
    '                   36- iva intracomunitario de bienes corrientes
    '                   38- iva intracomunitario de bien de inversion
    '                   42- iva regimen especial agrario




Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String
Dim ImpTotal As Currency
Dim ImpCompensa As Currency
Dim Periodo As String
Dim Consolidado As String


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



Private Sub cmbPeriodo_Validate(Index As Integer, Cancel As Boolean)
    PonerDatosFicheroSalida
    
    If cmbPeriodo(0).ListIndex > 0 Then
        txtperiodo(0).Text = cmbPeriodo(0).ListIndex
        txtperiodo(1).Text = cmbPeriodo(0).ListIndex
    End If
    FramePeriodo.Enabled = (cmbPeriodo(0).ListIndex = 0)
    FramePeriodo.Visible = (cmbPeriodo(0).ListIndex = 0)
End Sub


Private Sub cmdAccion_Click(Index As Integer)
Dim B As Boolean
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    InicializarVbles True
    
    
    'Si tiene compensaciones de peridoso anteriores
    'CompensacionAnterior
    ImpTotal = 0
    If txtCuota(0).Text <> "" Then
        ImpTotal = ImporteFormateado(txtCuota(0).Text)
    End If
    cadParam = cadParam & "CompensacionAnterior=" & ImpTotal & "|"
    numParam = numParam + 1
    


    
    'Guardamos el valor del chk del IVA
'--
'    ModeloIva False
    Label13.Caption = "Elimina datos anteriores"
    Label13.Visible = True
    Label13.Refresh
    B = GeneraLasLiquidaciones
    If B Then
        
    
        Label13.Caption = ""
        Label13.Refresh
        espera 0.5
        'Periodos
        Sql = ""
        For i = 0 To 1
            Sql = Sql & txtperiodo(i).Text & "|"
        Next i
        Sql = Sql & txtAno(0).Text & "|"
        i = 1
        
        Periodo = Sql & i & "|"
        
        'Empresas para consolidado
        Sql = ""
        If EmpresasSeleccionadas = 1 Then
            For i = 1 To Me.ListView1(1).ListItems.Count
                If ListView1(1).ListItems(i).Checked Then
                    If Me.ListView1(1).ListItems(i).Text <> vEmpresa.nomempre Then Sql = Me.ListView1(1).ListItems(i).Text
                End If
            Next i
        Else
            'Mas de una empresa
            Sql = "'Empresas seleccionadas:' + Chr(13) "
            For i = 1 To Me.ListView1(1).ListItems.Count - 1
                Sql = Sql & " + ""        " & Me.ListView1(1).ListItems(i).Text & """ + Chr(13)"
            Next i
        End If

        Consolidado = Sql

        
    End If
    Label13.Visible = False
    Me.Refresh
    Screen.MousePointer = vbDefault

    If Not B Then Exit Sub
'++
    
    
    If Not HayRegParaInforme("tmpliquidaiva", "codusu = " & vUsu.Codigo) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        
        ModeloHaciend2
        
    
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

Private Sub cmdCancelar_Click()
    Unload Me
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
    Me.Icon = frmppal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    CargarListView 1
    
    PonerPeriodoPresentacion303
    
    'Otras opciones
    If OpcionListado = 0 Then
        Me.Caption = "Modelo 303"
    Else
        Me.Caption = "Modelo 390"
        Label3(6).Left = 360
        txtAno(0).Left = 330
        Label3(7).Visible = False
        cmbPeriodo(0).Visible = False
        cmbPeriodo(0).Enabled = False
        Label3(0).Visible = False
        txtCuota(0).Enabled = False
        txtCuota(0).Visible = False
        Me.ToolbarAyuda.Visible = False
        Me.ToolbarAyuda.Enabled = False
        
        txtperiodo(0).Text = 1
        If vParam.periodos = 0 Then
            txtperiodo(1).Text = 4
        Else
            txtperiodo(1).Text = 12
        End If
        
    End If
     
    FrameSeccion.Enabled = vParam.EsMultiseccion
    
    FramePeriodo.Enabled = (Me.cmbPeriodo(0).ListIndex = 0)
    FramePeriodo.Visible = (Me.cmbPeriodo(0).ListIndex = 0)
    
    txtFecha(2).Text = Format(Now, "dd/mm/yyyy")
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    
    PonerDatosFicheroSalida
    
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub

Private Sub PonerDatosFicheroSalida()
    
    txtTipoSalida(1).Text = App.Path & "\Exportar\Mod303_" & Format(Mid(Me.txtAno(0), 3, 2), "00") & Format(Me.cmbPeriodo(0).ListIndex, "00") & ".txt"

End Sub


Private Sub PonerPeriodoPresentacion303()

    cmbPeriodo(0).Clear
    Me.cmbPeriodo(0).AddItem "Manual"
    If vParam.periodos = 0 Then
        'Liquidacion TRIMESTRAL
        
        For i = 1 To 4
            If i = 1 Or i = 3 Then
                CadenaDesdeOtroForm = "er"
            Else
                CadenaDesdeOtroForm = "�"
            End If
            CadenaDesdeOtroForm = i & CadenaDesdeOtroForm & " "
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm & " trimestre"
        Next i
    Else
        'Liquidacion MENSUAL
        For i = 1 To 12
            CadenaDesdeOtroForm = MonthName(i)
            CadenaDesdeOtroForm = UCase(Mid(CadenaDesdeOtroForm, 1, 1)) & LCase(Mid(CadenaDesdeOtroForm, 2))
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm
        Next
    End If
    
    
    'Leeremos ultimo valor liquidado
    
    txtAno(0).Text = vParam.anofactu
    i = vParam.perfactu + 1
    If vParam.periodos = 0 Then
        NumRegElim = 4
    Else
        NumRegElim = 12
    End If
        
    If i > NumRegElim Then
            i = 1
            txtAno(0).Text = vParam.anofactu + 1
    End If
    Me.cmbPeriodo(0).ListIndex = i '- 1
     
     
    txtperiodo(0).Text = Me.cmbPeriodo(0).ListIndex
    txtperiodo(1).Text = Me.cmbPeriodo(0).ListIndex
    
     
    
    CadenaDesdeOtroForm = ""
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' tabla de codigos de iva
        Case 0
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = False
            Next i
        Case 1
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = True
            Next i
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


Private Sub txtPag2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub





Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
End Sub


Private Sub ModeloHaciend2()
Dim Sql2 As String
Dim i As String
Dim Es_A_Compensar As Byte
Dim CadenaImportes As String

    'Generamos la cadena con los importes a mostrar
    cad = ""
    GeneraCadenaImportes

    CadenaImportes = CStr(cad)


    'Si el importe es negativo tendriamos que preguntar si quiere realizar
    'compensacion o ingreso/devolucion
    If CCur(ImpTotal) < 0 Then
        'NEGATIVO
        cad = "Importe a devolver / compensar." & vbCrLf & vbCrLf & _
            "� Desea que sea a compensar ?"
        i = MsgBox(cad, vbQuestion + vbYesNoCancel)
        If i = vbCancel Then Exit Sub
        Es_A_Compensar = 0
        If i = vbYes Then Es_A_Compensar = 1
        
    Else
        cad = "Ingreso por cta banco?" & vbCrLf & vbCrLf
        '
        i = MsgBox(cad, vbQuestion + vbYesNoCancel)
        If i = vbCancel Then Exit Sub
        Es_A_Compensar = 2
        If i = vbYes Then Es_A_Compensar = 3
    End If


    'Generamos la cadena para el ultimo registro de la presentacion
    'Registro <T30303>
    cad = ""
    CadenaAdicional303_Nuevo


    'Ahora enviamos a generar fichero IVA
    If GenerarFicheroIVA_303_2014(CadenaImportes, ImpTotal, CDate(txtFecha(2).Text), Periodo, Es_A_Compensar, cad) Then
    
    GuardarComo
    End If
    
End Sub

Private Sub CadenaAdicional303_Nuevo()
Dim Sql As String
Dim Rs As ADODB.Recordset
    'REGISTRO T30303>
    
    
    'Entregas intracomunitarias
'    DevuelveImporte 35, 0  'base
    Sql = "select  sum(bases) bases from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 14 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
    Else
        DevuelveImporte 0, 0
    End If
    
    
    'Exportaciones y asimiladas todas las facturas que sean de
'    DevuelveImporte 37, 0  'base
    Sql = "select  sum(bases) bases from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 16 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
    Else
        DevuelveImporte 0, 0
    End If

    
    'DE estos dos NO hay text
    '---------------------
    'Op no sujetas o con conversion del sujeto pasivo
    Sql = "select  sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 12 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
    Else
        DevuelveImporte 0, 0
    End If
    
    'Diferencia antes de aplicar las
    DevuelveImporte ImpTotal, 0
     
    
    'Atribuible a la admon del estado
    DevuelveImporte 31, 0  '%
    DevuelveImporte ImpTotal * (-1), 0

    'A compensar de otros periodos
    ImpCompensa = ImporteSinFormato(ComprobarCero(txtCuota(0).Text))
    DevuelveImporte ImpCompensa, 0  'base
    
    'DE estos dos NO hay text
    'Diputacion foral
    cad = cad & String(17, "0")
    
    'Campo13. Resultado
    DevuelveImporte ImpTotal - ImpCompensa, 0

    'Campo14. A deducor
    DevuelveImporte ImpTotal - ImpCompensa, 0

    'Campo15. Resultado de la liquidacion
    DevuelveImporte ImpTotal - ImpCompensa, 0

End Sub



'Cojera los importes y los formateara para los programitas de hacineda
Private Sub GeneraCadenaImportes()
Dim TotalClien As Currency
Dim TotalProve As Currency
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset

    TotalClien = 0

    'En devuelveimporte
    ' Tipo 0:   11 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales

    
    Sql = "select iva,  bases, ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 0 "
    Sql = Sql & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        i = i + 1
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!IVA, "N"), 3
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    
    'por si hay menos de 3 porcentajes de iva hay que rellenarlos a ceros
    For J = i + 1 To 3
        DevuelveImporte 0, 0
        DevuelveImporte 0, 3
        DevuelveImporte 0, 0
    Next J
    
    Set Rs = Nothing
    
    'Adquisiciones intra
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 10 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    HayReg = False
    
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    ' Inversion de sujeto pasivo
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 12 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'modificacion bases y cuotas (no tenemos)
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0
    
    
    'Los recargos
    Sql = "select iva,  bases, ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 "
    Sql = Sql & " order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        i = i + 1
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!IVA, "N"), 3
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    
    'por si hay menos de 3 porcentajes de iva hay que rellenarlos a ceros
    For J = i + 1 To 3
        DevuelveImporte 0, 0
        DevuelveImporte 0, 3
        DevuelveImporte 0, 0
    Next J
    
    Set Rs = Nothing
    
    'modificacion bases y cuotas del recargo de equivalencia (no tenemos)
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0
    

    'total
'--
'    DevuelveImporte 33, 0
    DevuelveImporte TotalClien, 0
    
    '------------------------------------------------------------------------
    '------------------------------------------------------------------------
    'DEDUCIBLE
    TotalProve = 0
    
'    'operaciones interiores

    '[Monica]24/06/2016: en las facturas de proveedores faltaba a�adir las fras de ISP, he a�adido el 12
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente in ( 2, 12 )  "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'operaciones interiores BIENES INVERSION
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 30 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 32 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones BIEN INVERSION
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 34 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    
    
    'adqisiciones intracom
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 36 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'adqisiciones intracom BIEN INVERSION
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 38 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing

    ' rectificacion de deducciones tampoco tenemos
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0

'--
'    DevuelveImporte 28, 0  'Regimen especial
    Sql = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 42 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
    End If
    
    Set Rs = Nothing
    
    DevuelveImporte 0, 0  'Regularizacion inversiones
    DevuelveImporte 0, 0  'Regularizacion por aplicacion del porcentaje def de prorrata

    
    'total a deducir
    DevuelveImporte TotalProve, 0
    
    
    'Diferencia
'--
'    DevuelveImporte 29, 0  'base
    DevuelveImporte TotalClien - TotalProve, 0  'Regularizacion inversiones
    
    ImpTotal = TotalClien - TotalProve
    
     
End Sub


'Ahora desde un importe, antes Desde un text box
Private Sub DevuelveImporte(Importe As Currency, Tipo As Byte)
Dim J As Integer
Dim Aux As String
Dim Resul As String

    Dim modelo As Integer
    modelo = 4

    Resul = ""
    If Importe < 0 Then
        Aux = ""
        Resul = "N"
    Else
        Aux = "0"
    End If
    Importe = Importe * 100
'++ hasta aqui

    
    'Tipo sera la mascara.
    ' si Modelo<>303
        ' Tipo 0:   11 enteros y 2 decimales
        'Else
        ' Tipo 0:   15 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales
    Select Case Tipo
    Case 1
        Aux = Aux & "000"
    Case 2
        Aux = Aux & "00"
    Case 3
        Aux = Aux & "0000"
    Case Else
        If modelo = 4 Then
            Aux = Aux & String(16, "0")  '15 enteros 2 decima  17-1
        Else
            Aux = Aux & String(10, "0")   '11 enteros 2 decimales  13-1
        End If
    End Select
    
    cad = cad & Resul & Format(Importe, Aux)
        
End Sub



Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = True
        
    
    indRPT = "0408-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "FacturasCliFecha.rpt"

    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pTitulo=""" & Me.Caption & """|"
    numParam = numParam + 1
    
    If OpcionListado = 0 Then
        cadParam = cadParam & "pTipo=0|"
    Else
        cadParam = cadParam & "pTipo=1|"
    End If
    numParam = numParam + 1
    
    
    Sql = ""
    If EmpresasSeleccionadas = 1 Then
        For i = 1 To Me.ListView1(1).ListItems.Count
            If ListView1(1).ListItems(i).Checked Then
                If Me.ListView1(1).ListItems(i).Text <> vEmpresa.CodEMpre Then Sql = Me.ListView1(1).ListItems(i).SubItems(1)
            End If
        Next i
    Else
        'Mas de una empresa
        Sql = "Empresas seleccionadas: "" + Chr(13) "
        For i = 1 To Me.ListView1(1).ListItems.Count
            Sql = Sql & " + """ & Me.ListView1(1).ListItems(i).SubItems(1) & """ + Chr(13) "
        Next i
        Sql = Sql & " + """
    End If
    
    cadParam = cadParam & "empresas= """ & Sql & """|"
    numParam = numParam + 1
    

    cadParam = cadParam & "pPeriodo1=""" & txtperiodo(0).Text
    If vParam.periodos = 0 Then
        cadParam = cadParam & "T""|"
    Else
        cadParam = cadParam & """|"
    End If
    
    cadParam = cadParam & "pPeriodo2=""" & txtperiodo(1).Text
    If vParam.periodos = 0 Then
        cadParam = cadParam & "T""|"
    Else
        cadParam = cadParam & """|"
    End If
    
    cadParam = cadParam & "pAno=" & txtAno(0).Text & "|"
    numParam = numParam + 3
    
    
    cadFormula = "{tmpliquidaiva.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 19
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Function CargarTemporal() As Boolean
Dim Sql As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    Sql = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "insert into tmpfaclin (codusu, codigo, numserie, nomserie, numfac, fecha, cta, cliente, nif, imponible, impiva, total, retencion,"
    Sql = Sql & " recargo, tipoopera, tipoformapago) "
    Sql = Sql & " select distinct " & vUsu.Codigo & ",0, factcli.numserie, contadores.nomregis, factcli.numfactu, factcli.fecfactu, factcli.codmacta, "
    Sql = Sql & " factcli.nommacta, factcli.nifdatos, factcli.totbases, factcli.totivas, factcli.totfaccl, factcli.trefaccl, "
    Sql = Sql & " factcli.totrecargo, tipofpago.descformapago , aa.denominacion"
    Sql = Sql & " from " & tabla
    Sql = Sql & " where " & cadselect
    
    Conn.Execute Sql
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
End Function

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim i As Integer


    MontaSQL = False
    
            
    Sql = ""
    For i = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then
            Sql = Sql & Me.ListView1(1).ListItems(i).Text & ","
        End If
    Next i
    
    If Sql <> "" Then
        ' quitamos la ultima coma
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        
        If Not AnyadirAFormula(cadselect, "factcli_totales.codigiva in (" & Sql & ")") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{factcli_totales.codigiva} in [" & Sql & "]") Then Exit Function
    Else
        If Not AnyadirAFormula(cadselect, "factcli_totales.codigiva is null") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "isnull({factcli_totales.codigiva})") Then Exit Function
    End If
    
    
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
    If Not CargarTemporal Then Exit Function
    
    cadFormula = "{tmpfaclin.codusu} = " & vUsu.Codigo
    
            
    MontaSQL = True
End Function

Private Sub txtAno_GotFocus(Index As Integer)
    ConseguirFoco txtAno(Index), 3
End Sub

Private Sub txtAno_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAno_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtAno(Index).Text = Trim(txtAno(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'A�o
            txtAno(Index).Text = Format(txtAno(Index).Text, "0000")
            
            PonerDatosFicheroSalida
    End Select

End Sub


Private Sub txtCuota_GotFocus(Index As Integer)
    ConseguirFoco txtCuota(Index), 3
End Sub

Private Sub txtCuota_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCuota_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtCuota(Index).Text = Trim(txtCuota(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'Cuota
            If Not PonerFormatoDecimal(txtCuota(0), 1) Then
                txtCuota(0).Text = ""
            Else
                If ImporteFormateado(txtCuota(0).Text) < 0 Then
                    MsgBox "Importe positivo", vbExclamation
                    txtCuota(0).Text = ""
                    PonFoco txtCuota(0)
                End If
            End If
    End Select

End Sub






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
    
    If OpcionListado = 0 Then
        If cmbPeriodo(0).ListIndex = -1 Or txtperiodo(0).Text = "" Then
            MsgBox "Campos per�odo no pueden estar vacios", vbExclamation
            Exit Function
        End If
        
        If cmbPeriodo(0).ListIndex = 0 Then
            For i = 0 To 1
                If Me.txtperiodo(i).Text = "" Then
                    MsgBox "Campos per�odo no pueden estar vacios", vbExclamation
                    Exit Function
                End If
            Next i
            
            If Val(txtperiodo(0).Text) > Val(txtperiodo(1).Text) Then
                MsgBox "Per�odo desde mayor que per�odo hasta.", vbExclamation
                Exit Function
            End If
            
            
            If vParam.periodos = 1 Then
                If Val(txtperiodo(0).Text) > 12 Or Val(txtperiodo(1).Text) > 12 Then
                    MsgBox "Per�odo no puede ser superior a 12.", vbExclamation
                    Exit Function
                End If
            Else
                'TRIMESTRAL
                If Val(txtperiodo(0).Text) > 4 Or Val(txtperiodo(1).Text) > 4 Then
                    MsgBox "Per�odo no puede ser superior a 4.", vbExclamation
                    Exit Function
                End If
            End If
        End If
    Else
        If txtAno(0).Text = "" Then
            MsgBox "El a�o no puede estar vacio.", vbExclamation
            Exit Function
        End If
    End If
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
    End If

    For i = 1 To Me.ListView1(1).ListItems.Count
        If ListView1(1).ListItems(i).Checked Then
            If Not ComprobarContabilizacionFrasCliProv(True, ListView1(1).ListItems(i).Text) Then Exit Function
            If Not ComprobarContabilizacionFrasCliProv(False, ListView1(1).ListItems(i).Text) Then Exit Function
        End If
    Next i

    If txtCuota(0).Text <> "" Then
        If ImporteFormateado(txtCuota(0).Text) < 0 Then
            MsgBox "Importe a compensar debe ser positivo", vbExclamation
            Exit Function
        End If
    End If
            
        


    DatosOK = True

End Function


Private Function EmpresasSeleccionadas() As Integer
Dim Sql As String
Dim i As Integer
Dim NSel As Integer

    NSel = 0
    For i = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then NSel = NSel + 1
    Next i
    EmpresasSeleccionadas = NSel

End Function

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "C�digo", 600
    ListView1(Index).ColumnHeaders.Add , , "Descripci�n", 3200
    
    Sql = "SELECT codempre, nomempre, conta "
    Sql = Sql & " FROM usuarios.empresasariconta "
    
    If Not vParam.EsMultiseccion Then
        Sql = Sql & " where conta = " & DBSet(Conn.DefaultDatabase, "T")
    Else
        Sql = Sql & " where mid(conta,1,8) = 'ariconta'"
    End If
    Sql = Sql & " ORDER BY codempre "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        If vParam.EsMultiseccion Then
            If EsMultiseccion(DBLet(Rs!CONTA)) Then
                Set ItmX = ListView1(Index).ListItems.Add
                
                If DBLet(Rs!CONTA) = Conn.DefaultDatabase Then ItmX.Checked = True
                ItmX.Text = Rs.Fields(0).Value
                ItmX.SubItems(1) = Rs.Fields(1).Value
            End If
        Else
            Set ItmX = ListView1(Index).ListItems.Add
            
            ItmX.Checked = True
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Rs.Fields(1).Value
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Empresas.", Err.Description
    End If
End Sub


Private Function GeneraLasLiquidaciones() As Boolean
    
    '       cliprov     0- Facturas clientes
    '                   1- RECARGO EQUIVALENCIA
    '                   2- Facturas proveedores
    '                   3- libre
    '                   4- IVAS no deducible
    '                   5- Facturas NO DEDUCIBLES
    '                   6- IVA BIEN INVERSION
    '                   7- Compras extranjero
    '                   8- Inversion sujeto pasivo (Abril 2015)
    
    'Borramos los datos temporales
    Sql = "DELETE FROM tmpliquidaiva WHERE codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    'Modificacion para desglosaar los IVAS que sean:
    '   Intracom
    '   Regimen especial agrario
    '    inversion sujeto pasivo
    '...
    '  Para ello en tmpcierre1 pondremos para el usuario
    '  en nommacta: adqintra   ,  ventintra, campo
    '  para cada empresa
    Sql = "DELETE FROM tmpctaexplotacioncierre where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    
    'Si quiere ver el IVA detallado
'--
'    If Me.chkIVAdetallado.Value = 1 Then
        Sql = "DELETE FROM tmpimpbalan WHERE codusu =" & vUsu.Codigo
        Conn.Execute Sql
        Sql = "DELETE FROM tmpimpbalance WHERE codusu =" & vUsu.Codigo
        Conn.Execute Sql
'    End If
    
    
    Sql = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    NumRegElim = 0
    'Para cada empresa
    'Para cada periodo
    For i = 1 To Me.ListView1(1).ListItems.Count  'List2.ListCount - 1
        If Me.ListView1(1).ListItems(i).Checked Then
            For CONT = CInt(txtperiodo(0).Text) To CInt(txtperiodo(1).Text)
                Label13.Caption = Mid(ListView1(1).ListItems(i).SubItems(1), 1, 20) & ".  " & CONT
                Label13.Refresh
                If Not LiquidacionIVANew(CByte(CONT), CInt(txtAno(0).Text), Me.ListView1(1).ListItems(i).Text, True) Then      '(chkIVAdetallado.Value = 1)
                    GeneraLasLiquidaciones = False
                    Exit Function
                End If
            Next CONT
        End If
    Next i
    'Borraremos todos aquellos IVAS de Base imponible=0
    Sql = "DELETE From tmpliquidaiva WHERE codusu = " & vUsu.Codigo
    Sql = Sql & " AND bases = 0"
    Conn.Execute Sql
    
   'Insertamos en Usuarios para el posible informe
    Sql = "INSERT INTO tmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita) "
    Sql = Sql & " SELECT codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita FROM tmpimpbalance "
    Sql = Sql & " WHERE codusu=" & vUsu.Codigo
    Conn.Execute Sql
        
    
    GeneraLasLiquidaciones = True
End Function

Private Sub ModeloIva(Leer As Boolean)

On Error GoTo EModeloIva
EModeloIva:
    Err.Clear
End Sub


Private Sub GuardarComo()

    On Error GoTo EGuardarComo



    cd1.ShowSave
    cad = cd1.FileName
    If cad <> "" Then
        FileCopy App.Path & "\miIVA.txt", cad
    End If
    Exit Sub
EGuardarComo:
    MuestraError Err.Number
End Sub

