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
      Height          =   2715
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   6915
      Begin VB.CheckBox chk1 
         Caption         =   "Presentacion ultimo periodo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   38
         Top             =   1920
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   4335
      End
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
         Caption         =   "Período"
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
         Caption         =   "Año"
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
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
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
         Caption         =   "Cuotas a compensar períodos anteriores"
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

Private SQL As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String
Dim ImpTotal As Currency
Dim ImpCompensa As Currency
Dim Periodo As String
Dim Consolidado As String


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

Private Sub El_ultimoA_Presentar()
Dim B As Boolean
    B = False
    If vParam.SIITiene Then
        'If optTipoSal(1).Value Then
            If Me.cmbPeriodo(0).ListIndex = cmbPeriodo(0).ListCount - 1 Then B = True
        'End If
    End If
    chk1.visible = B
    If chk1.visible Then chk1.Value = IIf(B, 1, 0)
End Sub

Private Sub cmbPeriodo_Click(Index As Integer)
    El_ultimoA_Presentar
End Sub

Private Sub cmbPeriodo_Validate(Index As Integer, Cancel As Boolean)
    PonerDatosFicheroSalida
    
    If cmbPeriodo(0).ListIndex > 0 Then
        txtperiodo(0).Text = cmbPeriodo(0).ListIndex
        txtperiodo(1).Text = cmbPeriodo(0).ListIndex
    End If
    FramePeriodo.Enabled = (cmbPeriodo(0).ListIndex = 0)
    FramePeriodo.visible = (cmbPeriodo(0).ListIndex = 0)
    
End Sub


Private Sub cmdAccion_Click(Index As Integer)
Dim B As Boolean
    
    Label13.Caption = ""
    Label13.visible = True
    Screen.MousePointer = vbHourglass
    B = DatosOK
    Screen.MousePointer = vbDefault
    Label13.visible = False
    If Not B Then Exit Sub
    
    
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
    cadParam = cadParam & "CompensacionAnterior=" & Replace(CStr(ImpTotal), ",", ".") & "|"
    numParam = numParam + 1
    


    
    'Guardamos el valor del chk del IVA
'--
'    ModeloIva False
    Label13.Caption = "Elimina datos anteriores"
    Label13.visible = True
    Label13.Refresh
    B = GeneraLasLiquidaciones
    If B Then
        
    
        Label13.Caption = ""
        Label13.Refresh
        espera 0.5
        'Periodos
        SQL = ""
        For i = 0 To 1
            SQL = SQL & txtperiodo(i).Text & "|"
        Next i
        SQL = SQL & txtAno(0).Text & "|"
        i = 1
        
        Periodo = SQL & i & "|"
        
        'Empresas para consolidado
        SQL = ""
        If EmpresasSeleccionadas = 1 Then
            For i = 1 To Me.ListView1(1).ListItems.Count
                If ListView1(1).ListItems(i).Checked Then
                    If Me.ListView1(1).ListItems(i).Text <> vEmpresa.nomempre Then SQL = Me.ListView1(1).ListItems(i).Text
                End If
            Next i
        Else
            'Mas de una empresa
            SQL = "'Empresas seleccionadas:' + Chr(13) "
            For i = 1 To Me.ListView1(1).ListItems.Count - 1
                If ListView1(1).ListItems(i).Checked Then
                    SQL = SQL & " + ""        " & Me.ListView1(1).ListItems(i).Text & """ + Chr(13)"
                End If
            Next i
        End If

        Consolidado = SQL

        
    End If
    Label13.visible = False
    Me.Refresh
    Screen.MousePointer = vbDefault

    If Not B Then Exit Sub
'++
    
    B = HayRegParaInforme("tmpliquidaiva", "codusu = " & vUsu.Codigo)
    
    If Not B Then
        If Me.chk1.visible Then
            If Me.chk1.Value Then
                B = True
                
                'Meto una entrada a cero para linkar report
                SQL = "insert into `tmpliquidaiva` (`codusu`,`iva`,`bases`,`ivas`,`codempre`) values ( " & vUsu.Codigo & ",0,0,0,0)"
                Conn.Execute SQL
            End If
        End If
    End If
    
    If Not B Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        If Me.OpcionListado = 1 Then
            ModeloHaciend390
        Else
            ModeloHaciend303
        End If
    
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
        Label3(7).visible = False
        cmbPeriodo(0).visible = False
        cmbPeriodo(0).Enabled = False
        Label3(0).visible = False
        txtCuota(0).Enabled = False
        txtCuota(0).visible = False
        Me.ToolbarAyuda.visible = False
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
    FramePeriodo.visible = (Me.cmbPeriodo(0).ListIndex = 0)
    
    txtFecha(2).Text = Format(Now, "dd/mm/yyyy")
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    
    PonerDatosFicheroSalida
    
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub

Private Sub PonerDatosFicheroSalida()
    If OpcionListado = 1 Then
        txtTipoSalida(1).Text = App.Path & "\Exportar\Mod390.txt"     '_" & Format(Mid(Me.txtAno(0), 3, 2), "00") & ".txt"
    Else
        txtTipoSalida(1).Text = App.Path & "\Exportar\Mod303_" & Format(Mid(Me.txtAno(0), 3, 2), "00") & Format(Me.cmbPeriodo(0).ListIndex, "00") & ".txt"
    End If

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
                CadenaDesdeOtroForm = "º"
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
   ' El_ultimoA_Presentar
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


Private Sub ModeloHaciend303()
Dim Sql2 As String
Dim i As String
Dim Es_A_Compensar As Byte
Dim CadenaImportes As String
Dim B As Boolean

    'Generamos la cadena con los importes a mostrar
    cad = ""
    GeneraCadenaImportes303

    CadenaImportes = CStr(cad)


    'Si el importe es negativo tendriamos que preguntar si quiere realizar
    'compensacion o ingreso/devolucion
    If CCur(ImpTotal) < 0 Then
        'NEGATIVO
        cad = "Importe a devolver / compensar." & vbCrLf & vbCrLf & _
            "¿ Desea que sea a compensar ?"
        i = MsgBox(cad, vbQuestion + vbYesNoCancel)
        If i = vbCancel Then Exit Sub
        
        Es_A_Compensar = 0
        If i = vbYes Then Es_A_Compensar = 1
        
        
        If Es_A_Compensar = 0 Then
            cad = DevuelveDesdeBD("iban1", "empresa2", "1", "1")
            If cad = "" Then
                MsgBox "Falta configurar IBAN para la devolucion", vbExclamation
                Exit Sub
            End If
        End If
        
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
    i = 0
    If Me.chk1.visible Then
        If Me.chk1.Value = 1 Then i = 1
    End If
    
    
    '    B = GenerarFicheroIVA_390_2020(CadenaImportes, ImpTotal, CDate(txtFecha(2).Text), Periodo, Es_A_Compensar, cad, True)
    
    
    B = GenerarFicheroIVA_303_2017(CadenaImportes, ImpTotal, CDate(txtFecha(2).Text), Periodo, Es_A_Compensar, cad, i = 1)
    
    If B Then GuardarComo
    
    
    
End Sub

Private Sub CadenaAdicional303_Nuevo()
Dim SQL As String
Dim Rs As ADODB.Recordset
    'REGISTRO T30303>
    
    
    'Entregas intracomunitarias

    SQL = "select  sum(bases) bases from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 14 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
    Else
        DevuelveImporte 0, 0
    End If
    Rs.Close
    
    'Exportaciones y asimiladas todas las facturas que sean de

    SQL = "select  sum(bases) bases from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 16 "
    
   
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
    Else
        DevuelveImporte 0, 0
    End If
    Rs.Close
    
    'Op no sujetas o con conversion del sujeto pasivo
    'Segun MC, el punto en la liquidacion de esto solo afecta a aquellas ventas extentas de IVA.
    ' Sep 2019. Añadimos el tipo 61
    SQL = "select  sum(bases) base from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 61"
    
   
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveImporte DBLet(Rs!Base, "N"), 0
    Else
        DevuelveImporte 0, 0
    End If
    Rs.Close
    
    'Adiconal criterio de caja.
    cad = cad & String(17, "0")
    cad = cad & String(17, "0")
    cad = cad & String(17, "0")
    cad = cad & String(17, "0")
    
    'Reegularizacion cuotas
    cad = cad & String(17, "0")
    
    'Diferencia antes de aplicar las
    DevuelveImporte ImpTotal * 1, 0
    
    
    'Atribuible a la admon del estado
    'DevuelveImporte 31, 0   '%  PONIA 31 antes de ene 18
    cad = cad & "10000" '100%
    cad = cad & "0000"
    DevuelveImporte ImpTotal * 1, 0


    'IVA a la importación liquidado por la Aduana pendiente de ingreso  [77]
    cad = cad & String(17, "0")

    'A compensar de otros periodos
    ImpCompensa = ImporteSinFormato(ComprobarCero(txtCuota(0).Text))
    DevuelveImporte ImpCompensa * 1, 0  'base
    
    'DE estos dos NO hay text
    'Diputacion foral
    cad = cad & String(17, "0")
    
    'Campo19. Resultado
    DevuelveImporte ImpTotal - ImpCompensa, 0

    'Campo20. A deducor
    'DevuelveImporte ImpTotal - ImpCompensa, 0
    DevuelveImporte 0, 0

    'Campo21. Resultado de la liquidacion
    DevuelveImporte ImpTotal - ImpCompensa, 0

End Sub



'Cojera los importes y los formateara para los programitas de hacineda
Private Sub GeneraCadenaImportes303()
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
    
    'Enero 2020     El 390, tiene que hacer sumatoriso de iva
    
    SQL = "select iva ,  sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 0 "
    If vParam.RectificativasSeparadas303 Then SQL = SQL & " AND iva<100"
    SQL = SQL & " group by iva order by iva "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Msg = ""
        J = 0 'IVAS que si tengo
        K = 0 'Ivas que proceso
        While Not Rs.EOF
            Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
            J = J + 1
            Rs.MoveNext
        Wend
        Rs.MoveFirst
        Msg = "IVAs en contabilidad:  " & Msg & vbCrLf & vbCrLf & "Procesados: "
        For i = 1 To 3
           
            
            'primero el 4  despues el 10 despues el 21
            SQL = RecuperaValor("4|10|21|", i)
            Rs.Find "IVA = " & DBSet(SQL, "N"), , adSearchForward, 1
            
            If Rs.EOF Then
                DevuelveImporte 0, 0
                DevuelveImporte 0, 3
                DevuelveImporte 0, 0
            Else
                Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
                K = K + 1
                DevuelveImporte DBLet(Rs!Bases, "N"), 0
                DevuelveImporte DBLet(Rs!IVA, "N"), 3
                DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
                TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
            End If
            
        Next
        If K <> J Then
            SQL = "Error en IVAS regimen general. " & vbCrLf & " Existen " & Msg
            MsgBox SQL, vbQuestion
        End If
        
    Else
        'No hay IVA normal
        For J = 1 To 3
            DevuelveImporte 0, 0
            DevuelveImporte 0, 3
            DevuelveImporte 0, 0
        Next J
    End If
    Rs.Close
    
    
    Set Rs = Nothing
    
    
    
    
    'Adquisiciones intra
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 10 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    HayReg = False
    
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    ' Inversion de sujeto pasivo
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 12 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    
    
    'JUNIO 2019
    'modificacion bases y cuotas
    HayReg = False
    If vParam.RectificativasSeparadas303 Then
        Set Rs = New ADODB.Recordset
        SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 0 AND iva=100"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not Rs.EOF
            SQL = SQL & "X"
            HayReg = True
            DevuelveImporte DBLet(Rs!Bases, "N"), 0
            DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
            TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Len(SQL) > 1 Then
            MsgBox "Error en facturas rectificativas sin R.Equiv.  Mas de una linea devuelta", vbExclamation
            
        End If
        
        
    End If
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    
    'Los recargos
    Set Rs = New ADODB.Recordset
    
    
    SQL = "select iva,  bases, ivas,porcrec from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 "
    If vParam.RectificativasSeparadas303 Then SQL = SQL & " AND iva<100"
    SQL = SQL & " order by 1 "

    
    Rs.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Msg = ""
        J = 0 'IVAS que si tengo
        K = 0 'Ivas que proceso
        While Not Rs.EOF
            Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
            J = J + 1
            Rs.MoveNext
        Wend
        Rs.MoveFirst
        Msg = "IVAs en contabilidad:  " & Msg & vbCrLf & vbCrLf & "Procesados: "
    
    
        For i = 1 To 3
            
            
            'primero el 4  despues el 10 despues el 21
            SQL = RecuperaValor(vParam.OrdenIvas303Aeat, i)
            Rs.Find "IVA = " & DBSet(SQL, "N"), , adSearchForward, 1
            
            If Rs.EOF Then
                DevuelveImporte 0, 0
                DevuelveImporte 0, 3
                DevuelveImporte 0, 0
            Else
                Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
                K = K + 1
                DevuelveImporte DBLet(Rs!Bases, "N"), 0
                DevuelveImporte DBLet(Rs!porcrec, "N"), 3   'DevuelveImporte DBLet(Rs!IVA, "N"), 3
                DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
                TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
            End If
            
        Next
        If K <> J Then
            SQL = "Error en IVAS recargo equivalencia. Existen " & Msg
            MsgBox SQL, vbExclamation
        End If
    
    Else
    
        For J = 1 To 3
            DevuelveImporte 0, 0
            DevuelveImporte 0, 3
            DevuelveImporte 0, 0
        Next J
        
    End If
    Rs.Close
    
    'JUNIO 2019
    'modificacion bases y cuotas del recargo de equivalencia
    HayReg = False
    If vParam.RectificativasSeparadas303 Then
        Set Rs = New ADODB.Recordset
        SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 AND iva=101"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not Rs.EOF
            SQL = SQL & "X"
            HayReg = True
            DevuelveImporte DBLet(Rs!Bases, "N"), 0
            DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
            TotalClien = TotalClien + DBLet(Rs!Ivas, "N")
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Len(SQL) > 1 Then
            MsgBox "Error en facturas rectificativas sin R.Equiv.  Mas de una linea devuelta", vbExclamation
            
        End If
        
        
    End If
    
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    
    
    
    

    'total
'--
'    DevuelveImporte 33, 0
    DevuelveImporte 1 * TotalClien, 0
    
    '------------------------------------------------------------------------
    '------------------------------------------------------------------------
    'DEDUCIBLE
    TotalProve = 0
    
'    'operaciones interiores

    '[Monica]24/06/2016: en las facturas de proveedores faltaba añadir las fras de ISP, he añadido el 12
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente in ( 2, 12 )  "

    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'operaciones interiores BIENES INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 30 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 32 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones BIEN INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 34 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    
    
    'adqisiciones intracom
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 36 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'adqisiciones intracom BIEN INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 38 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing



    'JUNIO 2019
    ' rectificacion de deducciones
    HayReg = False
    If vParam.RectificativasSeparadas303 Then
        SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 40 "
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not Rs.EOF
            HayReg = True
            SQL = SQL & "X"
            DevuelveImporte DBLet(Rs!Bases, "N"), 0
            DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
            TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
            
            Rs.MoveNext
        Wend
        Rs.Close
        If Len(SQL) > 1 Then MsgBox "Error en facturas rectificativas DECUCIBLE .  Mas de una linea devuelta", vbExclamation

    End If

    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    
    
'--
'    DevuelveImporte 28, 0  'Regimen especial
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 42 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
    End If
    
    Set Rs = Nothing
    
    DevuelveImporte 0, 0  'Regularizacion inversiones
    DevuelveImporte 0, 0  'Regularizacion por aplicacion del porcentaje def de prorrata

    
    'total a deducir
    DevuelveImporte 1 * TotalProve, 0
    
    
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
        Importe = Importe * -1
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
    
    
    SQL = ""
    If EmpresasSeleccionadas = 1 Then
        For i = 1 To Me.ListView1(1).ListItems.Count
            If ListView1(1).ListItems(i).Checked Then
                If Me.ListView1(1).ListItems(i).Text <> vEmpresa.codempre Then SQL = Me.ListView1(1).ListItems(i).SubItems(1)
            End If
        Next i
    Else
        'Mas de una empresa
        SQL = "Empresas seleccionadas: "" + Chr(13) "
        For i = 1 To Me.ListView1(1).ListItems.Count
            If ListView1(1).ListItems(i).Checked Then
                SQL = SQL & " + """ & Me.ListView1(1).ListItems(i).SubItems(1) & """ + Chr(13) "
            End If
        Next i
        SQL = SQL & " + """
    End If
    
    cadParam = cadParam & "empresas= """ & SQL & """|"
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
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 19
        
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

Private Function MontaSQL() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim i As Integer


    MontaSQL = False
    
            
    SQL = ""
    For i = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then
            SQL = SQL & Me.ListView1(1).ListItems(i).Text & ","
        End If
    Next i
    
    If SQL <> "" Then
        ' quitamos la ultima coma
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        
        If Not AnyadirAFormula(cadselect, "factcli_totales.codigiva in (" & SQL & ")") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{factcli_totales.codigiva} in [" & SQL & "]") Then Exit Function
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
        Case 0 'Año
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
            MsgBox "Campos período no pueden estar vacios", vbExclamation
            Exit Function
        End If
        
        If cmbPeriodo(0).ListIndex = 0 Then
            For i = 0 To 1
                If Me.txtperiodo(i).Text = "" Then
                    MsgBox "Campos período no pueden estar vacios", vbExclamation
                    Exit Function
                End If
            Next i
            
            If Val(txtperiodo(0).Text) > Val(txtperiodo(1).Text) Then
                MsgBox "Período desde mayor que período hasta.", vbExclamation
                Exit Function
            End If
            
            
            If vParam.periodos = 1 Then
                If Val(txtperiodo(0).Text) > 12 Or Val(txtperiodo(1).Text) > 12 Then
                    MsgBox "Período no puede ser superior a 12.", vbExclamation
                    Exit Function
                End If
            Else
                'TRIMESTRAL
                If Val(txtperiodo(0).Text) > 4 Or Val(txtperiodo(1).Text) > 4 Then
                    MsgBox "Período no puede ser superior a 4.", vbExclamation
                    Exit Function
                End If
            End If
        End If
    Else
        If txtAno(0).Text = "" Then
            MsgBox "El año no puede estar vacio.", vbExclamation
            Exit Function
        End If
    End If
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
    End If

    For i = 1 To Me.ListView1(1).ListItems.Count
        If ListView1(1).ListItems(i).Checked Then
            Label13.Caption = "Comprobar fra: " & ListView1(1).ListItems(i).SubItems(1)
            Label13.Refresh
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
Dim SQL As String
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
Dim SQL As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "Código", 600
    ListView1(Index).ColumnHeaders.Add , , "Descripción", 3200
    
    SQL = "SELECT codempre, nomempre, conta "
    SQL = SQL & " FROM usuarios.empresasariconta "
    
    If Not vParam.EsMultiseccion Then
        SQL = SQL & " where conta = " & DBSet(Conn.DefaultDatabase, "T")
    Else
        SQL = SQL & " where mid(conta,1,8) = 'ariconta'"
    End If
    SQL = SQL & " ORDER BY codempre "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
    SQL = "DELETE FROM tmpliquidaiva WHERE codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    'Modificacion para desglosaar los IVAS que sean:
    '   Intracom
    '   Regimen especial agrario
    '    inversion sujeto pasivo
    '...
    '  Para ello en tmpcierre1 pondremos para el usuario
    '  en nommacta: adqintra   ,  ventintra, campo
    '  para cada empresa
    SQL = "DELETE FROM tmpctaexplotacioncierre where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    Conn.Execute "Delete from tmptesoreriacomun where codusu =" & vUsu.Codigo
    
    'Si quiere ver el IVA detallado
'--
'    If Me.chkIVAdetallado.Value = 1 Then
        SQL = "DELETE FROM tmpimpbalan WHERE codusu =" & vUsu.Codigo
        Conn.Execute SQL
        SQL = "DELETE FROM tmpimpbalance WHERE codusu =" & vUsu.Codigo
        Conn.Execute SQL
'    End If
    
    
    SQL = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    NumRegElim = 0
    'Para cada empresa
    'Para cada periodo
    For i = 1 To Me.ListView1(1).ListItems.Count  'List2.ListCount - 1
        If Me.ListView1(1).ListItems(i).Checked Then
            For CONT = CInt(txtperiodo(0).Text) To CInt(txtperiodo(1).Text)
                Label13.Caption = Mid(ListView1(1).ListItems(i).SubItems(1), 1, 20) & ".  " & CONT
                Label13.Refresh
                If Not LiquidacionIVANew(CByte(CONT), CInt(txtAno(0).Text), Me.ListView1(1).ListItems(i).Text, True) Then       '(chkIVAdetallado.Value = 1)
                    GeneraLasLiquidaciones = False
                    Exit Function
                End If
            Next CONT
        End If
    Next i
    'Borraremos todos aquellos IVAS de Base imponible=0
    SQL = "DELETE From tmpliquidaiva WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " AND bases = 0"
    Conn.Execute SQL
    
    SQL = ""
    If Me.chk1.visible Then
        If Me.chk1.Value = 1 Then SQL = "S"
    End If
    If SQL <> "" Then
        'Presentacion ULTIMO peridod
        Label13.Caption = "Ultimo periodo presentacion (I)"
        Label13.Refresh

        
        espera 0.5
        For i = 1 To Me.ListView1(1).ListItems.Count  'List2.ListCount - 1
            If Me.ListView1(1).ListItems(i).Checked Then
              
                    Label13.Caption = Mid(ListView1(1).ListItems(i).SubItems(1), 1, 20) & "   Ultimo peridodo"
                    Label13.Refresh
                    If Not LiquidacionIVAFinAnyo(CInt(txtAno(0).Text), Me.ListView1(1).ListItems(i).Text) Then
                        GeneraLasLiquidaciones = False
                        Exit Function
                    End If
              
            End If
        Next i
            
        'Se han generado los datos anuales junto a todos. Los sacmos sobre la tabla tmptesoreriacomun
        Label13.Caption = "Ultimo periodo presentacion (II)"
        Label13.Refresh
        SQL = "INSERT INTO tmptesoreriacomun (codusu,opcion,texto1,codigo,texto2,texto3,importe1,importe2)"
        SQL = SQL & " select codusu,cliente,codempre,@rownum:=@rownum+1 AS rownum  ,"
        SQL = SQL & " '','',sum(bases),sum(ivas) from tmpliquidaiva, (SELECT @rownum:=0) r "
        SQL = SQL & " where codusu=" & vUsu.Codigo & " and cliente<>199 and  periodo=100 group by 1,2,3"  'El 199 es l NO deducible
        Conn.Execute SQL
        
        SQL = " DELETE FROM tmpliquidaiva where codusu=" & vUsu.Codigo & " AND periodo=100"
        Conn.Execute SQL
        
        
        
        
        'Las opciones, 0,10,12  Regeimen genera, adquisino intra com, otras ope con ISP van juntas en la casilla regemine genera
        Label13.Caption = "Ultimo periodo presentacion (III)"
        Label13.Refresh
        SQL = "insert into tmptesoreriacomun (codusu ,codigo ,texto1 ,importe1 ,importe2 ,opcion)"
        SQL = SQL & " select codusu,codigo + 1000,texto1,sum(importe1),sum(importe2), 1 opcion from tmptesoreriacomun where"
        SQL = SQL & " codusu =" & vUsu.Codigo & " AND  opcion in (0,10,12)"
        SQL = SQL & " group by codusu,texto1"
        Conn.Execute SQL
        
        SQL = " DELETE FROM tmptesoreriacomun WHERE codusu =" & vUsu.Codigo & " AND  opcion in (0,10,12)"
        Conn.Execute SQL
        
        Label13.Caption = "Ultimo periodo presentacion (IV)"
        Label13.Refresh
        Set miRsAux = New ADODB.Recordset
        SQL = "Select distinct texto1 from tmptesoreriacomun where codusu=" & vUsu.Codigo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            SQL = "ariconta" & miRsAux!texto1 & ".empresa"
            SQL = DevuelveDesdeBD("nomresum", SQL, "1", "1")
            If SQL <> "" Then
                SQL = "UPDATE tmptesoreriacomun set texto1=" & DBSet(SQL, "T") & " WHERE codusu =" & vUsu.Codigo & " AND texto1 =" & DBSet(miRsAux!texto1, "T")
                Conn.Execute SQL
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If i = 1 Then
            'SOlo hay una empresa. NO lo detallo
            SQL = "UPDATE tmptesoreriacomun set texto1='' WHERE codusu =" & vUsu.Codigo
            Conn.Execute SQL
        End If
        
        
        Set miRsAux = New ADODB.Recordset
        SQL = "Select distinct opcion from tmptesoreriacomun where codusu=" & vUsu.Codigo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic
        cad = ""
        cad = cad & "@001@" & Mid("Operaciones en régimen general [80]" & Space(60), 1, 60)
        cad = cad & "@014@" & Mid("Entregas intracomunitarias exentas [93]" & Space(60), 1, 60)
        cad = cad & "@016@" & Mid("Exentas sin derecho a deducción [83]" & Space(60), 1, 60)
        cad = cad & "@061@" & Mid("No sujetas por reglas de loc. o con ISP[84]" & Space(60), 1, 60)
       ' Cad = Cad & "@001@" & Mid("Operaciones en régimen simplificado [86]" & Space(60), 1, 60)
       ' Cad = Cad & "@001@" & Mid("Exportaciones y operaciones exentas derecho a deducción [94]" & Space(60), 1, 60)
       ' Cad = Cad & "@001@" & Mid("Entregas de bienes de inversión [99]" & Space(60), 1, 60)
        
        While Not miRsAux.EOF
            
            SQL = "@" & Right("000" & miRsAux!Opcion, 3) & "@"
            i = InStr(1, cad, SQL)
            If i = 0 Then
                MsgBox "Opcin NO tratadas, Avise soporte técnico:  " & miRsAux!Opcion, vbExclamation
            Else
                SQL = Trim(Mid(cad, i + 5, 60))
                SQL = "UPDATE tmptesoreriacomun set texto2=" & DBSet(SQL, "T") & " WHERE codusu =" & vUsu.Codigo & " AND opcion =" & miRsAux!Opcion
                Conn.Execute SQL
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
        
        
        
        
        
        
        
        
        
        Set miRsAux = Nothing
        
    End If
    
    
    
    
    
    
   'Insertamos en Usuarios para el posible informe
    SQL = "INSERT INTO tmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita) "
    SQL = SQL & " SELECT codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita FROM tmpimpbalance "
    SQL = SQL & " WHERE codusu=" & vUsu.Codigo
    Conn.Execute SQL
        
    
    GeneraLasLiquidaciones = True
End Function

Private Sub ModeloIva(Leer As Boolean)

On Error GoTo EModeloIva
EModeloIva:
    Err.Clear
End Sub


Private Sub GuardarComo()

    On Error GoTo EGuardarComo


cd1.FileName = ""
    cd1.ShowSave
    cad = cd1.FileName
    If cad <> "" Then
        FileCopy App.Path & "\miIVA.txt", cad
        MsgBox "Fichero creado con éxito" & vbCrLf & vbCrLf & cad, vbInformation
    End If
    Exit Sub
EGuardarComo:
    MuestraError Err.Number
End Sub




'*********************************************************************************************************************
'*********************************************************************************************************************
'*********************************************************************************************************************
'
'Enero 2020
'Para el 390
'
Private Sub ModeloHaciend390()
Dim Sql2 As String
Dim i As String
Dim Es_A_Compensar As Byte
Dim CadenaImportes As String
Dim B As Boolean
    Dim Pagina2 As String
    Dim Pagina3 As String

    'Generamos la cadena con los importes a mostrar
    ImpTotal = 0
    cad = ""
    GeneraCadenaImportes390_Pagina2
    Pagina2 = CStr(cad)

     cad = ""
    GeneraCadenaImportes390_Pagina3
    Pagina3 = CStr(cad)

    If False Then
                            'Si el importe es negativo tendriamos que preguntar si quiere realizar
                            'compensacion o ingreso/devolucion
                            If CCur(ImpTotal) < 0 Then
                                'NEGATIVO
                                cad = "Importe a devolver / compensar." & vbCrLf & vbCrLf & _
                                    "¿ Desea que sea a compensar ?"
                                i = MsgBox(cad, vbQuestion + vbYesNoCancel)
                                If i = vbCancel Then Exit Sub
                                
                                Es_A_Compensar = 0
                                If i = vbYes Then Es_A_Compensar = 1
                                
                                
                                If Es_A_Compensar = 0 Then
                                    cad = DevuelveDesdeBD("iban1", "empresa2", "1", "1")
                                    If cad = "" Then
                                        MsgBox "Falta configurar IBAN para la devolucion", vbExclamation
                                        Exit Sub
                                    End If
                                End If
                                
                            Else
                                cad = "Ingreso por cta banco?" & vbCrLf & vbCrLf
                                '
                                i = MsgBox(cad, vbQuestion + vbYesNoCancel)
                                If i = vbCancel Then Exit Sub
                                Es_A_Compensar = 2
                                If i = vbYes Then Es_A_Compensar = 3
                            End If
    End If

    'Generamos la cadena para el ultimo registro de la presentacion
    'Registro <T30303>
    cad = ""
    CadenaAdicional303_Nuevo


    'Ahora enviamos a generar fichero IVA
    i = 0
    If Me.chk1.visible Then
        If Me.chk1.Value = 1 Then i = 1
    End If
    
    

    
    
    B = GenerarFicheroIVA_390_2020(CDate(txtFecha(2).Text), Periodo, Es_A_Compensar, Pagina2, Pagina3)
        
    If B Then GuardarComo
    
    
    
End Sub








Private Sub GeneraCadenaImportes390_Pagina2()
Dim TotCuotasYRecargo As Currency
Dim TotBases As Currency
Dim TotCuotas  As Currency
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset
       
    
    TotCuotasYRecargo = 0
    
    'En devuelveimporte
    ' Tipo 0:   11 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales
    
    
    TotBases = 0
    TotCuotas = 0
    SQL = "select iva ,  sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 0 "
    If vParam.RectificativasSeparadas303 Then SQL = SQL & " AND iva<100"
    SQL = SQL & " group by iva order by iva "


    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Msg = ""
        J = 0 'IVAS que si tengo
        K = 0 'Ivas que proceso
        While Not Rs.EOF
            Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
            J = J + 1
            Rs.MoveNext
        Wend
        Rs.MoveFirst
        Msg = "IVAs en contabilidad:  " & Msg & vbCrLf & vbCrLf & "Procesados: "
        For i = 1 To 3
           
            
            'primero el 4  despues el 10 despues el 21
            SQL = RecuperaValor("4|10|21|", i)
            Rs.Find "IVA = " & DBSet(SQL, "N"), , adSearchForward, 1
            
            If Rs.EOF Then
                DevuelveImporte 0, 0
                DevuelveImporte 0, 3
                DevuelveImporte 0, 0
            Else
                Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
                K = K + 1
                DevuelveImporte DBLet(Rs!Bases, "N"), 0
                DevuelveImporte DBLet(Rs!IVA, "N"), 3
                DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
                TotBases = TotBases + DBLet(Rs!Bases, "N")
                TotCuotas = TotCuotas + DBLet(Rs!Ivas, "N")
    
            End If
            
        Next
        If K <> J Then
            SQL = "Error en IVAS regimen general. " & vbCrLf & " Existen " & Msg
            MsgBox SQL, vbQuestion
        End If
        
    Else
        'No hay IVA normal
        For J = 1 To 3
            DevuelveImporte 0, 0
            DevuelveImporte 0, 3
            DevuelveImporte 0, 0
        Next J
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    '390
    '-----------------------------------------------------------------------------------------------------------------------
    'pos    Descripcion
    '115    5. Operaciones Reg. Gral. - Base imponible y cuota - operaciones intragrupo - Base imponible [500]
    '   6 campos de
    For J = 1 To 6
        DevuelveImporte 0, 0
    Next
    '217    5. Operaciones Reg. Gral. - Base imponible y cuota - regimen especial criterio caja - Base imponible [643]
    For J = 1 To 6
        DevuelveImporte 0, 0
    Next
    '319    5. Operaciones Reg. Gral. - Base Imponible y cuota - Reg. espec. bienes usados - Base imponible [07]
    For J = 1 To 5
        DevuelveImporte 0, 0
    Next
    '421   5. Operaciones Reg. Gral. - Base Imponible y cuota - Reg. espec. agencias viajes - Base imponible [13]
    For J = 1 To 2
        DevuelveImporte 0, 0
    Next
    
    
    'Adquisiciones intra
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 10 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    HayReg = False
    
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotBases = TotBases + DBLet(Rs!Bases, "N")
        TotCuotas = TotCuotas + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    '557    5. Operaciones Reg. Gral. - Base Imponible y cuota - Adquis. intracomunit. servicios - Base Imponible [545]
    For J = 1 To 6
        DevuelveImporte 0, 0
    Next
    
    
    ' Inversion de sujeto pasivo
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 12 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotBases = TotBases + DBLet(Rs!Bases, "N")
        TotCuotas = TotCuotas + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    

    HayReg = False
    If vParam.RectificativasSeparadas303 Then
        Set Rs = New ADODB.Recordset
        SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 0 AND iva=100"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not Rs.EOF
            SQL = SQL & "X"
            HayReg = True
            DevuelveImporte DBLet(Rs!Bases, "N"), 0
            DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
            TotBases = TotBases + DBLet(Rs!Bases, "N")
            TotCuotas = TotCuotas + DBLet(Rs!Ivas, "N")
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Len(SQL) > 1 Then
            MsgBox "Error en facturas rectificativas sin R.Equiv.  Mas de una linea devuelta", vbExclamation
            
        End If
        
        
    End If
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    
    
    '727    5. Operaciones Reg. Gral. - Base Imponible y cuota - Modificac. bases y cuotas intragrupo - Base [649]
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0
    '761    5. Operaciones Reg. Gral. - Base Imponible y cuota - Modificac. bases/cuotas concurso acreedores - Base imponible [31]
    DevuelveImporte 0, 0
    DevuelveImporte 0, 0
    
    '795    5. Operaciones Reg. Gral. - Base Imponible y cuota - Total bases y cuotas IVA - Base imponible [33]
    
    DevuelveImporte TotBases, 0
    DevuelveImporte TotCuotas, 0
    TotCuotasYRecargo = TotCuotas
    
    TotBases = 0
    TotCuotas = 0
    
    
    '829    Los recargos
    Set Rs = New ADODB.Recordset
    If OpcionListado = 1 Then
        SQL = "select iva,  sum(bases) bases , sum(ivas),porcrec from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 "
        If vParam.RectificativasSeparadas303 Then SQL = SQL & " AND iva<100"
        SQL = SQL & " group by iva ,porcrec order by 1 "
    Else
        SQL = "select iva,  bases, ivas,porcrec from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 "
        If vParam.RectificativasSeparadas303 Then SQL = SQL & " AND iva<100"
        SQL = SQL & " order by 1 "
    End If
    
    Rs.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Msg = ""
        J = 0 'IVAS que si tengo
        K = 0 'Ivas que proceso
        While Not Rs.EOF
            Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
            J = J + 1
            Rs.MoveNext
        Wend
        Rs.MoveFirst
        Msg = "IVAs en contabilidad:  " & Msg & vbCrLf & vbCrLf & "Procesados: "
    
    
        For i = 1 To 3
            
            
            'primero el 4  despues el 10 despues el 21
            SQL = RecuperaValor(vParam.OrdenIvas303Aeat, i)
            Rs.Find "IVA = " & DBSet(SQL, "N"), , adSearchForward, 1
            
            If Rs.EOF Then
                DevuelveImporte 0, 0
                DevuelveImporte 0, 0
            Else
                Msg = Msg & "  " & Format(Rs!IVA, FormatoImporte) & "%"
                K = K + 1
                DevuelveImporte DBLet(Rs!Bases, "N"), 0
                DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
                TotBases = TotBases + DBLet(Rs!Bases, "N")
                TotCuotas = TotCuotas + DBLet(Rs!Ivas, "N")
            End If
            
        Next
        If K <> J Then
            SQL = "Error en IVAS recargo equivalencia. Existen " & Msg
            MsgBox SQL, vbExclamation
        End If
    
    Else
    
        For J = 1 To 3
            DevuelveImporte 0, 0
            DevuelveImporte 0, 0
        Next J
        
    End If
    Rs.Close
    
    
    
    
    'modificacion bases y cuotas del recargo de equivalencia
    HayReg = False
    If vParam.RectificativasSeparadas303 Then
        Set Rs = New ADODB.Recordset
        SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 1 AND iva=101"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not Rs.EOF
            SQL = SQL & "X"
            HayReg = True
            DevuelveImporte DBLet(Rs!Bases, "N"), 0
            DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
            TotBases = TotBases + DBLet(Rs!Bases, "N")
            TotCuotas = TotCuotas + DBLet(Rs!Ivas, "N")
            
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Len(SQL) > 1 Then
            MsgBox "Error en facturas rectificativas sin R.Equiv.  Mas de una linea devuelta", vbExclamation
            
        End If
        
        
    End If
    
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    
    
    '1033 5. Operaciones Reg. Gral. - Base Imponible y cuota - Total cuotas IVA y recargo equivalencia [47]
    TotCuotasYRecargo = TotCuotasYRecargo + TotCuotas
    
    DevuelveImporte 1 * TotCuotasYRecargo, 0
    
    ImpTotal = TotCuotasYRecargo
    
    
End Sub


Private Sub GeneraCadenaImportes390_Pagina3()
Dim TotalProve  As Currency
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset
    
    
    '------------------------------------------------------------------------
    '------------------------------------------------------------------------
    'DEDUCIBLE
    TotalProve = 0
    
'    'operaciones interiores

    '[Monica]24/06/2016: en las facturas de proveedores faltaba añadir las fras de ISP, he añadido el 12
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente in ( 2, 12 )  "

    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'operaciones interiores BIENES INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 30 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 32 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'importaciones BIEN INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 34 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    
    
    'adqisiciones intracom
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 36 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing
    
    'adqisiciones intracom BIEN INVERSION
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 38 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Bases, "N"), 0
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    Set Rs = Nothing


    ' rectificacion de deducciones
    HayReg = False
    If vParam.RectificativasSeparadas303 Then
        SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 40 "
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not Rs.EOF
            HayReg = True
            SQL = SQL & "X"
            DevuelveImporte DBLet(Rs!Bases, "N"), 0
            DevuelveImporte DBLet(Rs!Ivas, "N"), 0
            
            TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
            
            Rs.MoveNext
        Wend
        Rs.Close
        If Len(SQL) > 1 Then MsgBox "Error en facturas rectificativas DECUCIBLE .  Mas de una linea devuelta", vbExclamation

    End If

    If Not HayReg Then
        DevuelveImporte 0, 0
        DevuelveImporte 0, 0
    End If
    
    

'    DevuelveImporte 28, 0  'Regimen especial
    SQL = "select sum(bases) bases, sum(ivas) ivas from tmpliquidaiva where codusu = " & DBSet(vUsu.Codigo, "N") & " and cliente = 42 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    HayReg = False
    While Not Rs.EOF
        HayReg = True
        DevuelveImporte DBLet(Rs!Ivas, "N"), 0
        
        TotalProve = TotalProve + DBLet(Rs!Ivas, "N")
        
        Rs.MoveNext
    Wend
    If Not HayReg Then
        DevuelveImporte 0, 0
    End If
    
    Set Rs = Nothing
    
    DevuelveImporte 0, 0  'Regularizacion inversiones
    DevuelveImporte 0, 0  'Regularizacion por aplicacion del porcentaje def de prorrata

    
    'total a deducir
    DevuelveImporte 1 * TotalProve, 0
    
    
    
    
    ImpTotal = ImpTotal - TotalProve
    
     
End Sub


