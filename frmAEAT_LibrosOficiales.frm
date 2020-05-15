VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAEAT_LibrosOficiales 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
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
      Height          =   4635
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   6915
      Begin VB.CheckBox chkSoloRea 
         Caption         =   "Solo R.E.A."
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
         Left            =   4200
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optTipoFac 
         Caption         =   "Recibidas"
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
         Left            =   2520
         TabIndex        =   32
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton optTipoFac 
         Caption         =   "Emitidas"
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
         Left            =   360
         TabIndex        =   31
         Top             =   1920
         Value           =   -1  'True
         Width           =   1575
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
         ItemData        =   "frmAEAT_LibrosOficiales.frx":0000
         Left            =   330
         List            =   "frmAEAT_LibrosOficiales.frx":0002
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
         TabIndex        =   1
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
      Height          =   4575
      Left            =   7110
      TabIndex        =   24
      Top             =   0
      Width           =   4455
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   180
         TabIndex        =   28
         Top             =   1020
         Width           =   4185
         Begin MSComctlLib.ListView ListView1 
            Height          =   2985
            Index           =   1
            Left            =   60
            TabIndex        =   29
            Top             =   510
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   5265
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
            Picture         =   "frmAEAT_LibrosOficiales.frx":0004
            ToolTipText     =   "Quitar al Debe"
            Top             =   120
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   3750
            Picture         =   "frmAEAT_LibrosOficiales.frx":014E
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
            TabIndex        =   30
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
         TabIndex        =   2
         Top             =   570
         Width           =   1485
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   25
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
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   2
         Left            =   1020
         Picture         =   "frmAEAT_LibrosOficiales.frx":0298
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
         TabIndex        =   26
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
      TabIndex        =   5
      Top             =   4890
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
      TabIndex        =   3
      Top             =   4890
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
      TabIndex        =   4
      Top             =   4890
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
      Left            =   1680
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
      TabIndex        =   27
      Top             =   4950
      Visible         =   0   'False
      Width           =   5475
   End
End
Attribute VB_Name = "frmAEAT_LibrosOficiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 419




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
Dim I As Integer
Dim IndCodigo As Integer

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





Private Sub cmdAccion_Click(Index As Integer)
Dim B As Boolean
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    InicializarVbles True
    
'
'    Si tiene compensaciones de peridoso anteriores
'    CompensacionAnterior
'    ImpTotal = 0
'    If txtCuota(0).Text <> "" Then
'        ImpTotal = ImporteFormateado(txtCuota(0).Text)
'    End If
'    cadParam = cadParam & "CompensacionAnterior=" & Replace(CStr(ImpTotal), ",", ".") & "|"
'    numParam = numParam + 1
'

    If Dir(App.Path & "\FraExpor.txt", vbArchive) <> "" Then Kill App.Path & "\FraExpor.txt"
    
    'Guardamos el valor del chk del IVA
'--
'    ModeloIva False
    Label13.Caption = "Elimina datos anteriores"
    Label13.visible = True
    Label13.Refresh
    NumRegElim = 0
    B = InsertaTmpFacturas
    Set miRsAux = Nothing
    Label13.Caption = ""
    Label13.Refresh
    
    
    Me.Refresh
    Screen.MousePointer = vbDefault

    If Not B Then Exit Sub
'++
    If NumRegElim = 0 Then
        MsgBox "No hya datos para guardar", vbExclamation
        Exit Sub
    End If
    espera 0.5
    
    'If optTipoSal(1).Value Then
    
    If True Then
        'EXPORTAR A CSV
        Label13.Caption = "Generando CSV"
        Label13.visible = True
        Me.Refresh
        ExportarCSV
        
        
    
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
    Label13.Caption = ""
    
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
  
        Me.Caption = "Libro oficial facturas AEAT "
        Me.ToolbarAyuda.visible = False
        Me.ToolbarAyuda.Enabled = False
  
     
    FrameSeccion.Enabled = vParam.EsMultiseccion
'
'    FramePeriodo.Enabled = (Me.cmbPeriodo(0).ListIndex = 0)
'    FramePeriodo.visible = (Me.cmbPeriodo(0).ListIndex = 0)
'
    txtFecha(2).Text = Format(Now, "dd/mm/yyyy")
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    
    PonerDatosFicheroSalida
    
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 1
    
End Sub

Private Sub PonerDatosFicheroSalida()
    
    txtTipoSalida(1).Text = App.Path & "\Exportar\Mod303_" & Format(Mid(Me.txtAno(0), 3, 2), "00") & Format(Me.cmbPeriodo(0).ListIndex, "00") & ".txt"

End Sub


Private Sub PonerPeriodoPresentacion303()

    cmbPeriodo(0).Clear
    Me.cmbPeriodo(0).AddItem "Anual"
    If vParam.periodos = 0 Then
        'Liquidacion TRIMESTRAL
        
        For I = 1 To 4
            If I = 1 Or I = 3 Then
                CadenaDesdeOtroForm = "er"
            Else
                CadenaDesdeOtroForm = "º"
            End If
            CadenaDesdeOtroForm = I & CadenaDesdeOtroForm & " "
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm & " trimestre"
        Next I
    Else
        'Liquidacion MENSUAL
        For I = 1 To 12
            CadenaDesdeOtroForm = MonthName(I)
            CadenaDesdeOtroForm = UCase(Mid(CadenaDesdeOtroForm, 1, 1)) & LCase(Mid(CadenaDesdeOtroForm, 2))
            Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm
        Next
    End If
    
    
    'Leeremos ultimo valor liquidado
    
    txtAno(0).Text = vParam.anofactu
    I = vParam.perfactu + 1
    If vParam.periodos = 0 Then
        NumRegElim = 4
    Else
        NumRegElim = 12
    End If
        
    If I > NumRegElim Then
            I = 1
            txtAno(0).Text = vParam.anofactu + 1
    End If
    Me.cmbPeriodo(0).ListIndex = I '- 1
     
     
'    txtperiodo(0).Text = Me.cmbPeriodo(0).ListIndex
'    txtperiodo(1).Text = Me.cmbPeriodo(0).ListIndex
'
     
    
    CadenaDesdeOtroForm = ""
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


Private Sub optTipoFac_Click(Index As Integer)
chkSoloRea.visible = Index = 1
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








Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
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
    
    cadParam = cadParam & "pTipo=1|"
    numParam = numParam + 1
    
    
    Sql = ""
    If EmpresasSeleccionadas = 1 Then
        For I = 1 To Me.ListView1(1).ListItems.Count
            If ListView1(1).ListItems(I).Checked Then
                If Me.ListView1(1).ListItems(I).Text <> vEmpresa.codempre Then Sql = Me.ListView1(1).ListItems(I).SubItems(1)
            End If
        Next I
    Else
        'Mas de una empresa
        Sql = "Empresas seleccionadas: "" + Chr(13) "
        For I = 1 To Me.ListView1(1).ListItems.Count
            Sql = Sql & " + """ & Me.ListView1(1).ListItems(I).SubItems(1) & """ + Chr(13) "
        Next I
        Sql = Sql & " + """
    End If
    
    cadParam = cadParam & "empresas= """ & Sql & """|"
    numParam = numParam + 1
    

   ' cadParam = cadParam & "pPeriodo1=""" & txtperiodo(0).Text
   ' If vParam.periodos = 0 Then
   '     cadParam = cadParam & "T""|"
   ' Else
   '     cadParam = cadParam & """|"
   ' End If
    
   ' cadParam = cadParam & "pPeriodo2=""" & txtperiodo(1).Text
   ' If vParam.periodos = 0 Then
   '     cadParam = cadParam & "T""|"
   ' Else
   '     cadParam = cadParam & """|"
   ' End If
    
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

Private Function CargarTemporal(NumeroConta As Integer) As Boolean
Dim Emitidas As Boolean
Dim TolLiqIva As Currency

    On Error GoTo eCargarTemporal

    CargarTemporal = False

    Emitidas = Me.optTipoFac(0).Value
    Sql = " select  " & vUsu.Codigo & ", f.numserie, f.numfactu, f.fecfactu, codpais,"
    Sql = Sql & " nifdatos,nommacta,codconce340,baseimpo,porciva,impoiva,porcrec,imporec,codopera"
    Sql = Sql & " , @rownum:=@rownum+1 AS rownum  , #porret# as porreten, #impret#  as imprete"
    If Emitidas Then
        cad = "insert into tmpfaclin (codusu, codigo, numserie,  numfac, fecha, cta, cliente, nif, imponible,IVA, impiva, porcrec , recargo,retencion,ImponibleAnt) "
      
        Sql = Replace(Sql, "#porret#", "F.retfaccl")
        Sql = Replace(Sql, "#impret#", "f.trefaccl")
       
     
        Sql = Sql & " FROM ariconta" & NumeroConta & ".factcli f,ariconta" & NumeroConta & ".factcli_totales ,  "
        Sql = Sql & " (SELECT @rownum:=0) r "
        Sql = Sql & " where f.numserie=factcli_totales.numserie and f.numfactu=factcli_totales.numfactu and"
        Sql = Sql & " f.anofactu=factcli_totales.anofactu"
    Else
        '## Dic 2019. En nodeducible ira la calave de operacion. De momento Si es INVERSION SUJETO PASIVO, ira clave S2 en el libro
                                                                                               'En total ira (deducible o no   tipoiva; numserie  suplidos:numregis
        cad = "insert into tmpfaclinprov (codusu, codigo ,Numfac ,FechaFac ,cta ,Cliente ,NIF ,Imponible ,IVA ,ImpIVA,FechaCon,Total,tipoiva,suplidos,nodeducible) "
        
        Sql = Replace(Sql, "#porret#", "F.retfacpr")
        Sql = Replace(Sql, "#impret#", "f.trefacpr")
        
        
        Sql = Sql & ", f.fecharec,f.numregis " 'para proveedores pondremos fecha reepcion -->FECHA OPERACION
        Sql = Sql & " FROM ariconta" & NumeroConta & ".factpro f,ariconta" & NumeroConta & ".factpro_totales , "
        Sql = Sql & " (SELECT @rownum:=" & NumRegElim & ") r "
        Sql = Sql & " where f.numserie=factpro_totales.numserie and f.numregis=factpro_totales.numregis and"
        Sql = Sql & " f.anofactu = factpro_totales.anofactu"
    End If
    
    
    Sql = Sql & " AND " & cadselect
    
    
    If Not Emitidas Then
        If chkSoloRea.Value = 1 Then Sql = Sql & " AND codconce340='X'"
    End If
    
    'ORDEN
    Sql = Sql & " ORDER BY "
    If Emitidas Then
        Sql = Sql & "f.fecfactu,f.numserie,f.numfactu "
    Else
        Sql = Sql & "f.fecharec,f.numserie,f.numregis "
    End If
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DoEvents
    Sql = ""
    While Not miRsAux.EOF
        '
            
        
        'Primer trozo comun codusu, codigo ,Numfac ,FechaFac ,cta ,Cliente ,NIF ,Imponible ,IVA ,ImpIVA
        NumRegElim = miRsAux!rownum
        Sql = Sql & ", (" & vUsu.Codigo & "," & miRsAux!rownum & ","
        If Emitidas Then
            RC = "'" & miRsAux!NUmSerie & "'," & Format(miRsAux!NumFactu, "000000")
        Else
            RC = DBSet(miRsAux!NumFactu, "T")
        End If
        
        Label13.Caption = RC
        Label13.Refresh
        
        'If RC = "337943" Then S top
        Sql = Sql & RC & "," & DBSet(miRsAux!FecFactu, "F") & ","
        RC = DBLet(miRsAux!codpais, "T")
        If RC = "" Then RC = "ES"
        Sql = Sql & DBSet(RC, "T") & "," & DBSet(miRsAux!Nommacta, "T") & "," & DBSet(miRsAux!nifdatos, "T") & ","
        
        'Impinible porceiva impoiva
        TolLiqIva = DBLet(miRsAux!Impoiva, "N")
        Sql = Sql & DBSet(miRsAux!Baseimpo, "N") & "," & DBSet(miRsAux!porciva, "N") & "," & DBSet(TolLiqIva, "N", "N") & ","
         
        If Emitidas Then
            RC = "null,0"
            If Not IsNull(miRsAux!porcrec) Then
                If miRsAux!porcrec > 0 Then RC = DBSet(miRsAux!porcrec, "N") & "," & DBSet(miRsAux!ImpoRec, "N", "N")
            End If
            Sql = Sql & RC & ","
            'Abril 2020
            Sql = Sql & DBSet(miRsAux!porreten, "N") & ","
            Sql = Sql & DBSet(miRsAux!imprete, "F")
            
            Sql = Sql & ")"
        Else
            ',FechaCon,Total
            Sql = Sql & DBSet(miRsAux!fecharec, "F") & ","
            
            'If miRsAux!nodeducible Then
            '
            'End If
            TolLiqIva = DBLet(miRsAux!Impoiva, "N")
            
            Sql = Sql & DBSet(TolLiqIva, "N") & "," & miRsAux!NUmSerie & "," & miRsAux!Numregis & ","
            If miRsAux!CodOpera = 4 Then
                Sql = Sql & "'S2'"
            Else
                Sql = Sql & "''"
            End If
            
            'Abril 2020
            'NOVA EN EL SQL. Lo tengo aqui por si acaso tuviera que alñadirlo
            'SQL = SQL & "," & DBSet(miRsAux!porreten, "N") & ","
            'SQL = SQL & DBSet(miRsAux!imprete, "F")

            
            
            Sql = Sql & ")"
            
        End If
        
        If Len(Sql) > 1000 Then HazInsertTmp
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    HazInsertTmp
    
 
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal ", Err.Description
End Function

Private Sub HazInsertTmp()
    If Sql <> "" Then
        Sql = Mid(Sql, 2)
        Sql = " VALUES " & Sql
        Sql = cad & Sql
        Conn.Execute Sql
        'Ejecuta SQL
        Sql = ""
    End If
End Sub

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
    
    If Val(Me.txtAno(0).Text) < 2000 Then
        MsgBox "Año incorrecto", vbExclamation
        Exit Function
    End If

    DatosOK = True

End Function


Private Function EmpresasSeleccionadas() As Integer
Dim Sql As String
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
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "Código", 600
    ListView1(Index).ColumnHeaders.Add , , "Descripción", 3200
    
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
                
                ItmX.Checked = True
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


Private Function InsertaTmpFacturas() As Boolean
Dim F1 As Date
Dim F2 As Date


    If Me.cmbPeriodo(0).ListIndex = 0 Then
        RC = " between '" & Me.txtAno(0).Text & "-01-01'  AND '" & txtAno(0).Text & "-12-31'"
    Else
        NumRegElim = Me.cmbPeriodo(0).ListIndex
        
        If vParam.periodos = 0 Then
            'Liquidacion TRIMESTRAL
            NumRegElim = ((NumRegElim - 1) * 3) + 1
            F1 = CDate("01/" & Format(NumRegElim, "00") & "/" & txtAno(0).Text)
            RC = " between " & DBSet(F1, "F")
            F1 = DateAdd("m", 3, F1)
            F1 = DateAdd("d", -1, F1)
            RC = RC & " AND " & DBSet(F1, "F")
        Else
             'Liquidacion TRIMESTRAL
            NumRegElim = ((NumRegElim - 1) * 3) + 1
            F1 = CDate("01/" & Format(NumRegElim, "00") & "/" & txtAno(0).Text)
            RC = " between " & DBSet(F1, "F")
            F1 = DateAdd("m", 1, F1)
            F1 = DateAdd("d", -1, F1)
            RC = RC & " AND " & DBSet(F1, "F")
        
        End If
    End If
    
    If Me.optTipoFac(0).Value Then
        'CLientes
        cadselect = "f.fecfactu "
        Sql = "delete from tmpfaclin where codusu = " & vUsu.Codigo
        Conn.Execute Sql
    Else
        'PRoveedores
        cadselect = "f.fecharec "
        Sql = "delete from tmpfaclinprov where codusu = " & vUsu.Codigo
        Conn.Execute Sql
    End If
    
    cadselect = cadselect & RC
    
    
    
    NumRegElim = 0
    'Para cada empresa
    'Para cada periodo
    For I = 1 To Me.ListView1(1).ListItems.Count  'List2.ListCount - 1
        If Me.ListView1(1).ListItems(I).Checked Then
                Set miRsAux = New ADODB.Recordset
                    
            Label13.Caption = "Leyendo facturas " & ListView1(1).ListItems(I).SubItems(1)
            Label13.Refresh

              If Not CargarTemporal(CInt(ListView1(1).ListItems(I).Text)) Then Exit Function
              
              Set miRsAux = Nothing
        End If
    Next I
    
    
    InsertaTmpFacturas = True
End Function



Private Sub GuardarComo()

    On Error GoTo EGuardarComo

    cd1.CancelError = True
    cd1.FileName = cad
    cd1.ShowSave
    cad = cd1.FileName
    If cad <> "" Then
        FileCopy App.Path & "\FraExpor.txt", cad
        MsgBox "Fichero creado correctamente: " & cd1.FileName, vbInformation
    End If
    Exit Sub
EGuardarComo:
    'MuestraError Err.Number
    Err.Clear
End Sub

Private Sub ExportarCSV()
Dim EpigrafeIAE As String


    cad = DevuelveDesdeBD("Epigrafe", "empresaactiv", "1", "1 ORDER By ppal DESC,id")
    cad = "    " & cad
    cad = Trim(Right(cad, 4))
    If Len(cad) <> 4 Then
        MsgBox "Error epigrafe actividad. Añada en epigrafes en configuracion->empresa", vbExclamation
        Exit Sub
    End If
    EpigrafeIAE = cad
    
    'Nombre fichero
    'Ejercicio
    '2) NIF
    '3) Tipo del Libro Registro del IVA que contiene el fichero, mediante uno de los siguientes valores:
    '   E: facturas Emitidas
    '   R: facturas Recibidas
    '4) Nombre o Razon social
    cad = IIf(optTipoFac(0).Value, "E", "R")
    
    Sql = vEmpresa.NombreEmpresaOficial
    Sql = Replace(Sql, ".", "")
    Sql = Replace(Sql, ",", "")
    cad = Me.txtAno(0).Text & vEmpresa.NIF & cad & Sql
    If Me.cmbPeriodo(0).ListIndex > 0 Then
        'HA pedido un periodo
        
        If vParam.periodos = 0 Then
            Sql = cmbPeriodo(0).ListIndex & "T"
        Else
            Sql = Format(cmbPeriodo(0).ListIndex, "00")
        End If
        Sql = "_" & Sql
    Else
        Sql = ""
    End If
    cad = cad & Sql & ".csv"
    CadenaDesdeOtroForm = cad
    If Me.optTipoFac(0).Value Then
        cad = ""
        'Abril 2020
        If vParam.periodos = 0 Then
            Sql = "concat(((month(Fecha)-1) div 3)+1,'T')"
        Else
            Sql = "lpad(month(Fecha),2,'0')"
        End If
        Sql = "SELECT year(fecha) ejercicio, " & Sql & " periodo , 1 tipoActiviad , '" & EpigrafeIAE & "' as epigrafe, "
        Sql = Sql & "'F1' as ""Tipo factura"", 'I01' as ""Concepto del ingreso"", '' as ""Concepto computable"" "
        'lo que estaba
        Sql = Sql & ", date_format(Fecha ,'%d/%m/%Y')  ""Fecha Expedición"",'' as ""Fecha Operación"",numserie ""Serie(Identificación de la Factura)"",numfac as ""Número(Identificación de la Factura)"""
        Sql = Sql & ",'' as ""Número-Final(Identificación de la Factura)"""
        Sql = Sql & ",  if(cta='ES','',if(coalesce(intracom,0)=1,'02','06'))   ""Tipo(NIF Destinatario)"""
        Sql = Sql & ",cta as ""Código País(NIF Destinatario)"""
        Sql = Sql & ",substring(nif,1,20)  ""Identificación(NIF Destinatario)"""
        Sql = Sql & ",substring(cliente,1,40) ""Nombre Destinatario"""
        'SQL = SQL & ",'' as ""Fa ctura Sustitutiva"""  marzo 2020 YA no esta
        Sql = Sql & ",'' as ""Clave de Operación"""
        Sql = Sql & ", imponible + impiva + recargo ""Total Factura"""
        Sql = Sql & ",imponible ""Base Imponible"""
        Sql = Sql & ",iva ""Tipo de IVA"""
        Sql = Sql & ",impiva ""Cuota IVA Repercutida"""
        Sql = Sql & ",coalesce(porcrec,'') as ""Tipo de Recargo Eq."""
        Sql = Sql & ",if (coalesce(recargo,0)=0,'',recargo) as ""Cuota Recargo Eq."""
        Sql = Sql & ",'' as ""Fecha(Cobro)"""
        Sql = Sql & ",'' as ""Importe(Cobro)"""
        Sql = Sql & ",'' as ""Medio Utilizado(Cobro)"""
        Sql = Sql & ", '' as ""Identificación Medio Utilizado(Cobro)""  "
        'abril 2020. De momento vacios
        Sql = Sql & ",coalesce(retencion,'') as ""Tipo de retencion"""
        Sql = Sql & ",if (coalesce(ImponibleAnt,0)=0,'',ImponibleAnt) as ""Importe retenido"""
        
        
        
        
        Sql = Sql & " from tmpfaclin left join paises on cta=codpais where codusu=" & vUsu.Codigo



    Else
        cad = ""
        'Abril 2020
        If vParam.periodos = 0 Then
            Sql = "concat(((month(Fechacon)-1) div 3)+1,'T')"
        Else
            Sql = "lpad(month(Fechacon),2,'0')"
        End If
        Sql = "SELECT year(Fechacon) ejercicio, " & Sql & " periodo , 1 tipoActiviad , '" & EpigrafeIAE & "' as epigrafe, "
        Sql = Sql & "'F1' as ""Tipo factura"", 'G01' as ""Concepto del gasto"", '' as ""Gasto deducible"" "
    
    
        Sql = Sql & ", date_format(FechaFac ,'%d/%m/%Y') as ""Fecha Expedición"""
        Sql = Sql & ",date_format(Fechacon ,'%d/%m/%Y')  as ""Fecha Operación"""
        Sql = Sql & ",numfac as ""Serie-Número(Identificación Factura del Expedidor)"""
        Sql = Sql & ", '' as ""Número-Final(Identificación Factura del Expedidor)"""
        Sql = Sql & ",concat(if(tipoiva=1,'',tipoiva),replace(format(suplidos,0),',','')) as  ""Número Recepción"""
        Sql = Sql & ",'' as ""Número Recepción Final"""
        Sql = Sql & ",  if(cta='ES','',if(coalesce(intracom,0)=1,'02','06'))  as ""Tipo(NIF Expedidor)"""
        Sql = Sql & ",cta as ""Código País(NIF Expedidor)"""
        Sql = Sql & ",substring(nif,1,20)  as ""Identificación(NIF Expedidor)"""
        Sql = Sql & ",substring(cliente,1,40) as ""Nombre Expedidor"""
        'SQL = SQL & ",'' as ""Factura Sustitutiva"""
        
        'Diciembre19. En clave de operacion hay que poner S2 para las Inv. sujeto pasivo. Esta grabado en el campo NoDeducible
        'SQL = SQL & ",'' as ""Clave de Operación"""
        Sql = Sql & ",NoDeducible as ""Clave de Operación"""
        Sql = Sql & ",imponible + impiva as ""Total Factura"""
        Sql = Sql & ",imponible  as ""Base Imponible"""
        Sql = Sql & ",iva as ""Tipo de IVA"""
        Sql = Sql & ",impiva as ""Cuota IVA Soportado"""
        Sql = Sql & ",impiva as ""Cuota Deducible"""
        Sql = Sql & ",'' as ""Tipo de Recargo Eq."""
        Sql = Sql & ",'' as ""Cuota Recargo Eq."""
        Sql = Sql & ",'' as ""Fecha(Pago)"""
        Sql = Sql & ",'' as ""Importe(Pago)"""
        Sql = Sql & ",'' as ""Medio Utilizado(Pago)"""
        Sql = Sql & ",'' as ""Identificación Medio Utilizado(Pago)"""
        'abril 2020
        'de momento vacio
        'SQL = SQL & ",coalesce(retencion,'') as ""Tipo de retencion"""
        'SQL = SQL & ",if (coalesce(ImponibleAnt,0)=0,'',ImponibleAnt) as ""Importe retenido"""
        Sql = Sql & ",'' as ""Tipo de retencion"""
        Sql = Sql & ",'' as ""Importe retenido"""
        
        Sql = Sql & "from tmpfaclinprov left join paises on cta=codpais where codusu=" & vUsu.Codigo
        
    End If
    Sql = Sql & " ORDER BY CODIGO"
    
    'LLamos a la funcion
        
    GeneraFicheroCSV Sql, App.Path & "\FraExpor.txt", True
    cad = CadenaDesdeOtroForm
    GuardarComo
    
    CadenaDesdeOtroForm = ""
End Sub


