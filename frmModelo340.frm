VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModelo340 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   11655
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
      Height          =   2025
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtSustitutiva 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         MaxLength       =   13
         TabIndex        =   3
         Top             =   1470
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkSutitutiva 
         Caption         =   "Declaracion SUSTITUTIVA. "
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
         Left            =   420
         TabIndex        =   2
         Top             =   1200
         Width           =   3135
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
         ItemData        =   "frmModelo340.frx":0000
         Left            =   420
         List            =   "frmModelo340.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   660
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
         Left            =   3930
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   660
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Número identificativo / justificante"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   3210
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   4800
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
         Left            =   450
         TabIndex        =   22
         Top             =   300
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
         Left            =   3960
         TabIndex        =   21
         Top             =   420
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
      Height          =   4845
      Left            =   7050
      TabIndex        =   25
      Top             =   0
      Width           =   4455
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   180
         TabIndex        =   29
         Top             =   990
         Width           =   4185
         Begin MSComctlLib.ListView ListView1 
            Height          =   2985
            Index           =   1
            Left            =   60
            TabIndex        =   5
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
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   3750
            Picture         =   "frmModelo340.frx":0004
            ToolTipText     =   "Puntear al Debe"
            Top             =   120
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   3390
            Picture         =   "frmModelo340.frx":014E
            ToolTipText     =   "Quitar al Debe"
            Top             =   120
            Width           =   240
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
         TabIndex        =   26
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
         Picture         =   "frmModelo340.frx":0298
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
         TabIndex        =   27
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
      Left            =   10290
      TabIndex        =   8
      Top             =   4950
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
      TabIndex        =   6
      Top             =   4950
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
      Left            =   60
      TabIndex        =   7
      Top             =   4950
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
      Left            =   0
      TabIndex        =   9
      Top             =   2160
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
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl340 
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
      Left            =   1650
      TabIndex        =   28
      Top             =   5010
      Width           =   5535
   End
End
Attribute VB_Name = "frmModelo340"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 409


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

Private Sql As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String


Dim UltimoPeriodoLiquidacion As Boolean
Dim C2 As String




Dim V340()   'Llevara un str
             'indicara si cada empresa a declarr tiene
             'los tickets como letra de serie o como cuenta
             'en los campos 2 y 3 llevara si es serie la serie
             ' y si es cta las cuentas 1 y dos


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


Private Sub Check1_Click()
    
End Sub

Private Sub chkSutitutiva_Click()
    Label1.Visible = chkSutitutiva.Value = 1
    txtSustitutiva.Visible = chkSutitutiva.Value = 1
End Sub

Private Sub cmbPeriodo_Change(Index As Integer)
    PonerDatosFicheroSalida
End Sub


Private Sub cmdAccion_Click(Index As Integer)
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    
    Screen.MousePointer = vbHourglass
    If Modelo340(Me.ListView1(1), CInt(txtAno(0).Text), cmbPeriodo(0).ListIndex + 1, cad, lbl340, False, Me.optTipoSal(1).Value, V340(), UltimoPeriodoLiquidacion) Then
        lbl340.Caption = ""
        
        If Not optTipoSal(1).Value Then
            
            'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
        
            If optTipoSal(2).Value Or optTipoSal(3).Value Then
                ExportarPDF = True 'generaremos el pdf
            Else
                ExportarPDF = False
            End If
            SoloImprimir = False
            If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
            
            AccionesCrystal
        
        Else
               'Adelante
    
            '¡Anño periodo. Variable que se le pasa al mod340
            '
            cad = Format(Me.txtAno(0).Text, "0000") & Format(Me.cmbPeriodo(0).ListIndex + 1, "00")
            If vParam.periodos = 0 Then
                'TRIMESTRAL
                cad = cad & Me.cmbPeriodo(0).ListIndex + 1 & "T"
            Else
                'MENSUAL
                cad = cad & Format(Me.cmbPeriodo(0).ListIndex + 1, "00")
            End If
            RC = Me.txtSustitutiva.Text
                                                    'Guardar como
            If GeneraFichero340(True, cad, False, RC) Then
                'INSERTO EL LOG
                If CuardarComo340 Then InsertaLog340
                    
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    lbl340.Caption = ""
    
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
        
    'Otras opciones
    Me.Caption = "Modelo 340"

     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
    txtFecha(2).Text = Format(Now, "dd/mm/yyyy")
     
    CargarListView 1
    
    
    PonerPeriodoPresentacion340
     
     
    FrameSeccion.Enabled = vParam.EsMultiseccion
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    
    PonerDatosFicheroSalida
    
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub


Private Sub PonerDatosFicheroSalida()
    
    txtTipoSalida(1).Text = App.Path & "\Exportar\Mod340_" & Format(Mid(Me.txtAno(0), 3, 2), "00") & Format(Me.cmbPeriodo(0).ListIndex + 1, "00") & ".txt"

End Sub

Private Sub PonerPeriodoPresentacion340()

    cmbPeriodo(0).Clear
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
    
    
    'Leeremos ultimo valor liquidaco
    
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
    Me.cmbPeriodo(0).ListIndex = i - 1
     
    
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


Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    Sql = "Select factcli.numserie Serie, tmpfaclin.nomserie Descripcion, factcli.numfactu Factura, factcli.fecfactu Fecha, factcli.codmacta Cuenta, factcli.nommacta Titulo, tmpfaclin.tipoformapago TipoPago, "
    Sql = Sql & " tmpfaclin.tipoopera TOperacion, factcli.codconce340 TFra, factcli.trefaccl Retencion, "
    Sql = Sql & " factcli_totales.baseimpo BaseImp,factcli_totales.codigiva IVA,factcli_totales.porciva PorcIva,factcli_totales.porcrec PorcRec,factcli_totales.impoiva ImpIva,factcli_totales.imporec ImpRec "
    Sql = Sql & " FROM (factcli inner join factcli_totales on factcli.numserie = factcli_totales.numserie and factcli.numfactu = factcli_totales.numfactu and factcli.fecfactu = factcli_totales.fecfactu) "
    Sql = Sql & " inner join tmpfaclin ON factcli.numserie=tmpfaclin.numserie AND factcli.numfactu=tmpfaclin.Numfac and factcli.fecfactu=tmpfaclin.Fecha "
    Sql = Sql & " WHERE  tmpfaclin.codusu = 22000 "
    Sql = Sql & " ORDER BY factcli.codmacta, factcli.nommacta, factcli_totales.numlinea "
            
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0409-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "FacturasCliFecha.rpt"

    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "Empresas= """
    For i = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then
            cadParam = cadParam & Me.ListView1(1).ListItems(i).SubItems(1) & "  "
        End If
    Next i
    cadParam = Trim(cadParam)
    
    cadParam = cadParam & """|"
    
    'Diciembre 2012. Pongo el peridodo en el rpt
    cadParam = cadParam & "Periodo= ""Periodo: " & cmbPeriodo(0).ListIndex + 1 & "/" & CInt(txtAno(0).Text) & """|"
    numParam = numParam + 2
    
    cadFormula = "{tmp340.codusu}=" & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 20
        
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
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
    End If
    If Me.cmbPeriodo(0).ListIndex < 0 Or txtAno(0).Text = "" Then
        MsgBox "Seleccione un periodo/año", vbExclamation
        Exit Function
    End If
    
    UltimoPeriodoLiquidacion = False
    If cmbPeriodo(0).ListIndex = cmbPeriodo(0).ListCount - 1 Then UltimoPeriodoLiquidacion = True
    
   
    'Si es a fichero, y esta marcada la opcion de sustitutiva el campo txtsusito tiene que ser length =13 , si o si
    If optTipoSal(1).Value Then
        'Fichero AEAT
        If chkSutitutiva.Value = 1 Then
            i = 0
            If Len(Me.txtSustitutiva.Text) <> 13 Then
                i = 1
            Else
                If Not IsNumeric(Me.txtSustitutiva.Text) Then i = 1
            End If
            If i = 1 Then
                MsgBox Label1.Caption & " debe ser de 13 caracteres numericos.", vbExclamation
                Exit Function
            End If
    
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

Private Sub txtAno_GotFocus(Index As Integer)
    ConseguirFoco txtAno(Index), 3
End Sub

Private Sub txtAnyo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAnyo_LostFocus(Index As Integer)
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

Private Function CuardarComo340() As Boolean
    On Error GoTo ECopiarFichero347
    
    CuardarComo340 = False
    cd1.CancelError = True
    cd1.InitDir = Mid(App.Path, 1, 3)
    cd1.ShowSave
        
    cad = App.Path & "\tmp340.dat"
    
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("Ya existe: " & cd1.FileName & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        FileCopy cad, cd1.FileName
        MsgBox Space(20) & "Copia efectuada correctamente" & Space(20), vbInformation
        CuardarComo340 = True
    End If
    Exit Function
ECopiarFichero347:
    If Err.Number <> 32755 Then MuestraError Err.Number, "Copiar fichero"
    
End Function


Private Sub InsertaLog340()
Dim C2 As String
    
    
    cad = cad & " "
    For i = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then
            cad = cad & Me.ListView1(1).ListItems(i).SubItems(1) & "  "
        End If
    Next i
    cad = Trim(cad)
    
    
    
    'DICIMEBRE 2012
    'Diciembre 2012
    'Pagos en efectivo
    'Para guardarme un LOG de pagos declardaos
    'Ya que si luego modifican un apunte ...  perderiamos datos realmente.
    'ASi, con este log me que declaramos de efectivo
    cad = Format(Me.txtAno(0).Text, "0000") & "-"
    If vParam.periodos = 0 Then
        'TRIMESTRAL
        cad = cad & Me.cmbPeriodo(0).ListIndex + 1 & "T"
    Else
        'MENSUAL
        cad = cad & Format(Me.cmbPeriodo(0).ListIndex + 1, "00")
    End If
                       
    cad = " SELECT  now() fecha, codusu,'" & cad & "',nifdeclarado,razosoci,fechaexp,base,totiva  "
    cad = cad & " FROM tmp340 where codusu =" & vUsu.Codigo & " and clavelibro='Z'"
    
    cad = "INSERT INTO slog340 " & cad
    If Not EjecutaSQL(cad) Then MsgBox "Error insertando LOG. Consulte soporte técnico", vbExclamation
    
    
    
End Sub


