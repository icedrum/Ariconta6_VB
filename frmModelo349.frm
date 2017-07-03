VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModelo349 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
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
      Height          =   1935
      Left            =   60
      TabIndex        =   13
      Top             =   0
      Width           =   6915
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
         ItemData        =   "frmModelo349.frx":0000
         Left            =   270
         List            =   "frmModelo349.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   900
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
         Left            =   3780
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   900
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         Left            =   300
         TabIndex        =   18
         Top             =   540
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
         Left            =   3810
         TabIndex        =   17
         Top             =   540
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
      Height          =   4605
      Left            =   7050
      TabIndex        =   21
      Top             =   0
      Width           =   4455
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   150
         TabIndex        =   24
         Top             =   720
         Width           =   4185
         Begin MSComctlLib.ListView ListView1 
            Height          =   2880
            Index           =   1
            Left            =   60
            TabIndex        =   25
            Top             =   450
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
            TabIndex        =   26
            Top             =   120
            Width           =   1110
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   3750
            Picture         =   "frmModelo349.frx":0004
            ToolTipText     =   "Puntear al Debe"
            Top             =   60
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   3390
            Picture         =   "frmModelo349.frx":014E
            ToolTipText     =   "Quitar al Debe"
            Top             =   60
            Width           =   240
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   22
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
      TabIndex        =   4
      Top             =   4860
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
      Top             =   4860
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
      TabIndex        =   3
      Top             =   4770
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
      TabIndex        =   5
      Top             =   1950
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
         TabIndex        =   16
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   15
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   14
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
      Left            =   1680
      TabIndex        =   23
      Top             =   4830
      Width           =   5535
   End
End
Attribute VB_Name = "frmModelo349"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 411


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
Dim Tablas As String

Dim UltimoPeriodoLiquidacion As Boolean
Dim C2 As String

Dim FechaI As String
Dim FechaF As String
Dim Rs As ADODB.Recordset
Dim Importe As Currency

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


Private Sub cmbPeriodo_Change(Index As Integer)
    PonerDatosFicheroSalida
End Sub


Private Sub cmbPeriodo_Validate(Index As Integer, Cancel As Boolean)
    PonerDatosFicheroSalida
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim B As Boolean
Dim ConCli As Integer 'Clientes
Dim ConPro As Integer  'proveedores

Dim indRPT As String
Dim nomDocu As String

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    
    Screen.MousePointer = vbHourglass
    
    Screen.MousePointer = vbHourglass
    B = ComprobarCuentas349(ConCli, ConPro)
    Screen.MousePointer = vbDefault
    
    If Me.cmbPeriodo(0).ListIndex = 0 Then
        RC = "0A"
    Else
        If cmbPeriodo(0).ListIndex >= 1 And cmbPeriodo(0).ListIndex <= 12 Then
            RC = "0" & cmbPeriodo(0).ListIndex
        Else
            RC = cmbPeriodo(0).ListIndex - 12 & "T"
        End If
    End If
    


    If B Then
        If Me.optTipoSal(0).Value Then
            'Comprobamos si va mas de una empresa
            cad = vEmpresa.nomempre
            If EmpresasSeleccionadas Then vEmpresa.nomempre = "CONSOLIDADO"
                
            
            
            'Desde hastas Abril 2012
            RC = "Fechas: " & FechaI & " - " & FechaF
            RC = RC & "       Periodo: " & cmbPeriodo(0).Text
            RC = "pdh1= """ & RC & """|"
            cadParam = cadParam & RC
            numParam = numParam + 1
            'Las que habian
            RC = "ContadorLinCli= " & ConCli & "|ContadorLinPRO= " & ConPro & "|" & RC
            cadParam = cadParam & RC
            numParam = numParam + 1
            

            cadFormula = "{tmp347tot.codusu} = " & vUsu.Codigo

            indRPT = "0411-00"
            
            If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
            
            cadNomRPT = nomDocu '"Carta.rpt"
            
            ImprimeGeneral


            vEmpresa.nomempre = cad
        Else
        
        
        
            'Impresion del modelo oficial
            If MODELO349(RC, CInt(txtAno(0).Text)) Then CopiarFicheroHaciend3 (False)                 'Modelo de haciend a

        End If
    End If
    
    
    Screen.MousePointer = vbDefault
    
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
    Me.Caption = "Modelo 349"

     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
     
    CargarListView 1
    
    
    CargarCombo
    cmbPeriodo_Change (0)
     
    FrameSeccion.Enabled = vParam.EsMultiseccion
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    
    PonerDatosFicheroSalida
    
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub


Private Sub PonerDatosFicheroSalida()
Dim CADENA As String

    txtTipoSalida(1).Text = App.Path & "\Exportar\Mod349_" & Format(Mid(Me.txtAno(0), 3, 2), "00") & cmbPeriodo(0).Text & ".txt"

    Select Case cmbPeriodo(0).ListIndex
        Case 0
            FechaI = "01/01/" & Format(txtAno(0), "0000")
            FechaF = "31/12/" & Format(txtAno(0), "0000")
        Case 1 To 12
            FechaI = "01/" & Format(cmbPeriodo(0).ListIndex, "00") & "/" & Format(txtAno(0), "0000")
            FechaF = DateAdd("d", -1, DateAdd("m", 1, CDate(FechaI)))
        Case 13
            FechaI = "01/01/" & Format(txtAno(0), "0000")
            FechaF = "31/03/" & Format(txtAno(0), "0000")
        Case 14
            FechaI = "01/04/" & Format(txtAno(0), "0000")
            FechaF = "30/06/" & Format(txtAno(0), "0000")
        Case 15
            FechaI = "01/07/" & Format(txtAno(0), "0000")
            FechaF = "30/09/" & Format(txtAno(0), "0000")
        Case 16
            FechaI = "01/10/" & Format(txtAno(0), "0000")
            FechaF = "31/12/" & Format(txtAno(0), "0000")
    End Select
        


End Sub

Private Sub CargarCombo()

    cmbPeriodo(0).Clear
    
    Me.cmbPeriodo(0).AddItem "Anual"
    
    For i = 1 To 12
        CadenaDesdeOtroForm = MonthName(i)
        CadenaDesdeOtroForm = UCase(Mid(CadenaDesdeOtroForm, 1, 1)) & LCase(Mid(CadenaDesdeOtroForm, 2))
        Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm
    Next i
    
    For i = 1 To 4
        If i = 1 Or i = 3 Then
            CadenaDesdeOtroForm = "er"
        Else
            CadenaDesdeOtroForm = "º"
        End If
        CadenaDesdeOtroForm = i & CadenaDesdeOtroForm & " "
        Me.cmbPeriodo(0).AddItem CadenaDesdeOtroForm & " trimestre"
    Next i
    
    
    
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
    If vParam.periodos = 0 Then
        Me.cmbPeriodo(0).ListIndex = i + 12 - 1
    Else
        Me.cmbPeriodo(0).ListIndex = i - 1
    End If
    
    CadenaDesdeOtroForm = ""
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
        
    
    indRPT = "0411-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "FacturasCliFecha.rpt"

    cadParam = cadParam & "pFecha=""" & FechaI & FechaF & """|"
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

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
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


Private Function ComprobarCuentas349(ByRef C1 As Integer, ByRef C2 As Integer) As Boolean
Dim i As Integer
Dim Trim(3) As Currency
'Contadores para facturas de abono

    ComprobarCuentas349 = False
    
    Sql = "DELETE FROM tmp347tot where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    

    
    'Para el listado de facturas utilizaremos los datos
    Sql = "DELETE FROM tmpfaclin WHERE codusu =" & vUsu.Codigo
    Conn.Execute Sql
    Sql = "DELETE FROM tmpfaclinprov WHERE codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    C1 = 0
    C2 = 0
    
    'Esto sera para las inserciones de despues
    Tablas = "INSERT INTO tmp347tot (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla) "
    Tablas = Tablas & " VALUES (" & vUsu.Codigo
         
    Set miRsAux = New ADODB.Recordset
    For i = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(i).Checked Then
            
            Sql = "DELETE FROM tmp347 where codusu = " & vUsu.Codigo
            Conn.Execute Sql
            
            
            If Not ComprobarCuentas349_DOS("ariconta" & ListView1(1).ListItems(i).Text, C1, C2) Then
                Set miRsAux = Nothing
                Exit Function
            End If
        
           'Iremos NIF POR NIF
           
          Sql = "SELECT  cliprov,nif, sum(importe) as suma, tmp347.razosoci,tmp347.dirdatos,tmp347.codposta,"
          Sql = Sql & "tmp347.pais despobla from tmp347 where codusu=" & vUsu.Codigo
          Sql = Sql & " group by cliprov,nif "
          
          Set Rs = New ADODB.Recordset
          Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
          While Not Rs.EOF
               If ExisteEntrada Then
                    Importe = Importe + Rs!Suma
                    Sql = "UPDATE tmp347tot SET importe=" & TransformaComasPuntos(CStr(Importe))
                    Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & Rs!cliprov
                    Sql = Sql & " AND nif = '" & Rs!NIF & "';"
               Else
                    
                    Sql = "," & Rs!cliprov & ",'" & Rs!NIF & "'," & TransformaComasPuntos(CStr(Rs!Suma))
                    Sql = Sql & ",'" & DevNombreSQL(DBLet(Rs!razosoci)) & "','" & DevNombreSQL(DBLet(Rs!dirdatos)) & "','" & Rs!codposta & "','" & DevNombreSQL(DBLet(Rs!desPobla)) & "')"
                    Sql = Tablas & Sql
               End If
               Conn.Execute Sql
               Rs.MoveNext
          Wend
          Rs.Close
       End If
    Next i
    
    
        
    'Comprobamos si hay datos
    Sql = "Select count(*) FROM tmp347tot where codusu = " & vUsu.Codigo
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            CONT = miRsAux.Fields(0)
        End If
    End If
    miRsAux.Close
    
    If CONT = 0 Then
        If optTipoSal(0).Value Then
            'Listado
            MsgBox "Ningún dato se ha generado con esos valores", vbExclamation
        Else
            'DEjo continuar
            ComprobarCuentas349 = True
        End If
    Else
        ComprobarCuentas349 = True
    End If
    Set miRsAux = Nothing
    
End Function

Private Function ComprobarCuentas349_DOS(Contabilidad As String, ByRef ContadorCli As Integer, ByRef ContadorPro As Integer) As Boolean
Dim Rs As ADODB.Recordset
Dim Importe As Currency

On Error GoTo EComprobarCuentas349
    ComprobarCuentas349_DOS = False
    
    'Cargamos la tabla con los valores
    Sql = "SELECT "
    Sql = Sql & " factcli.codmacta,factcli.nifdatos,factcli.nommacta,factcli.dirdatos,factcli.codpobla,sum(baseimpo)as s1"
    Sql = Sql & " from " & Contabilidad & ".factcli," & Contabilidad & ".factcli_totales  where "
    Sql = Sql & " factcli.numserie = factcli_totales.numserie and factcli.numfactu = factcli_totales.numfactu and factcli.fecfactu = factcli_totales.fecfactu "
    Sql = Sql & " AND factcli.fecfactu >='" & Format(FechaI, FormatoFecha) & "'"
    Sql = Sql & " AND factcli.fecfactu <='" & Format(FechaF, FormatoFecha) & "'"
    'Factura extranjero
    Sql = Sql & " AND factcli.codopera=1"
    
'    'Pero si tiene serie de AUTOFACTURAS, la quitamos
    Sql = Sql & " group by factcli.codmacta,factcli.nifdatos "

    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, razosoci, dirdatos, codposta, importe)  VALUES (" & vUsu.Codigo & ",0,'"
    
    While Not Rs.EOF
        'Antes 20 Abril
        'Importe = RS!s1 + RS!s2 + RS!s3
        Importe = Rs!s1
        cad = Rs!codmacta & "','" & Rs!nifdatos & "'," & DBSet(Rs!Nommacta, "T") & "," & DBSet(Rs!dirdatos, "T") & "," & DBSet(Rs!CodPobla, "T") & "," & TransformaComasPuntos(CStr(Importe))
        cad = Sql & cad & ")"
        Conn.Execute cad
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    cad = "factpro.fecharec"
    Sql = "SELECT factpro.codmacta,factpro.nifdatos,factpro.nommacta,factpro.dirdatos,factpro.codpobla,sum(baseimpo)as s1 from " & Contabilidad & ".factpro," & Contabilidad & ".factpro_totales  where "
    Sql = Sql & " factpro.numserie = factpro_totales.numserie and factpro.numregis = factpro_totales.numregis and factpro.anofactu=factpro_totales.anofactu "
    Sql = Sql & " AND " & cad & " >='" & Format(FechaI, FormatoFecha) & "'"
    Sql = Sql & " AND " & cad & " <='" & Format(FechaF, FormatoFecha) & "'"
    'Extranjero
    Sql = Sql & " AND factpro.codopera = 1"
    Sql = Sql & " group by factpro.codmacta,factpro.nifdatos "
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, razosoci, dirdatos, codposta, importe)  VALUES (" & vUsu.Codigo & ",1,'"
    While Not Rs.EOF
        'Antes 20 Abril
        'Importe = RS!s1 + RS!s2 + RS!s3
        Importe = Rs!s1
        cad = Rs!codmacta & "','" & Rs!nifdatos & "'," & DBSet(Rs!Nommacta, "T") & "," & DBSet(Rs!dirdatos, "T") & "," & DBSet(Rs!CodPobla, "T") & "," & TransformaComasPuntos(CStr(Importe))
        cad = Sql & cad & ")"
        Conn.Execute cad
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    
    
    
    
    RC = ""
    cad = ""
    'Comprobaremos k el nif no es nulo, ni el codppos de las cuentas a tratar
    Sql = "Select cta from tmp347 where (nif is null or nif = '') and codusu = " & vUsu.Codigo
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        i = i + 1
        cad = cad & Rs.Fields(0) & "       "
        If i = 3 Then
            cad = cad & vbCrLf
            i = 0
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    If cad <> "" Then
        RC = "Cuentas con NIF sin valor: " & vbCrLf & vbCrLf & cad
    End If
    
    
    If RC <> "" Then
       MsgBox RC, vbExclamation
       Exit Function
    End If
    
   
    '----------------------------------------------------------
    'Listado detallado de las facturas en negativo
    '----------------------------------------------
    'CLIENTES
    
    'Para insertar
    RC = "INSERT INTO tmpfaclin (codusu, codigo, Numfac, Fecha, cta,  NIF, "
    RC = RC & " IVA,  Total,cliente) VALUES (" & vUsu.Codigo & ","
    
    
    Sql = "SELECT  numserie,numfactu,fecfactu,totfaccl,nif,factcli.codmacta,nommacta,totbases baseimpo  from " & Contabilidad & ".factcli,"
    Sql = Sql & "tmp347  where "
    Sql = Sql & " tmp347.cta= factcli.codmacta"
    Sql = Sql & " AND fecfactu >='" & Format(FechaI, FormatoFecha) & "'"
    Sql = Sql & " AND fecfactu <='" & Format(FechaF, FormatoFecha) & "'"
    'Factura extranjero
    Sql = Sql & " AND codopera=1"
    'De compras / vetnas cojemos compras
    Sql = Sql & " AND cliprov = 0"
    
    
    'Importes negativos
    Sql = Sql & " AND totfaccl <0"
    
    
    'Modificacion del 27 Febrero 2006
    Sql = Sql & " AND tmp347.codusu = " & vUsu.Codigo
    
    
    'Nº Empresa
    i = Val(Mid(Contabilidad, 6))

    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        ContadorCli = ContadorCli + 1
        
        Sql = ContadorCli & ",'" & Rs!NUmSerie & Format(Rs!NumFactu, "0000000000") & "','" & Format(Rs!FecFactu, FormatoFecha) & "','"
        ', cta,  NIF, IVA,  Total   .- IVA= numero empresa
        Importe = Rs!Baseimpo
        Sql = Sql & Rs!codmacta & "','" & Rs!NIF & "'," & i & "," & TransformaComasPuntos(CStr(Importe))
        
        
        Sql = Sql & ",'" & DevNombreSQL(Rs!Nommacta)
        Sql = RC & Sql & "')"
    
        
        Conn.Execute Sql
    
        Rs.MoveNext
    Wend
    Rs.Close
    
 
    'PROVEEDORES
    
    RC = "INSERT INTO tmpfaclinprov (codusu, codigo, Numfac, FechaCon, cta,  NIF, "
    RC = RC & " IVA,  Total,Fechafac,cliente) VALUES (" & vUsu.Codigo & ","
    Sql = "SELECT  numregis,fecharec,fecfactu,totfacpr,numfactu,nif,nommacta,totbases baseimpo  from " & Contabilidad & ".factpro,"
    Sql = Sql & "tmp347  where "
    Sql = Sql & " tmp347.cta=factpro.codmacta "
    
    'Solo usuario 1
    Sql = Sql & " AND tmp347.codusu = " & vUsu.Codigo
    
    Sql = Sql & " AND fecharec >='" & Format(FechaI, FormatoFecha) & "'"
    Sql = Sql & " AND fecharec <='" & Format(FechaF, FormatoFecha) & "'"
    'Factura extranjero
    Sql = Sql & " AND codopera=1"
    
    'De compras / vetnas cojemos compras
    Sql = Sql & " AND cliprov = 1"
    
    'Importes negativos
    Sql = Sql & " AND totfacpr <0"

    
    'Modificacion del 27 Febrero 2006
    Sql = Sql & " AND tmp347.codusu = " & vUsu.Codigo
    
    
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        ContadorPro = ContadorPro + 1
        
        Sql = ContadorPro & ",'" & Format(Rs!numregis, "0000000000") & "','" & Format(Rs!fecharec, FormatoFecha) & "','"
        ', cta,  NIF, IVA,  Total   .- IVA= numero empresa    cta=cod factura
        
        'Abril 2006. Busco la base, no el total factura
        Importe = Rs!Baseimpo
        Sql = Sql & Mid(Rs!NumFactu, 1, 10) & "','" & Rs!NIF & "'," & i & "," & TransformaComasPuntos(CStr(Importe))
        
        
        
        Sql = Sql & ",'" & Format(Rs!FecFactu, FormatoFecha) & "','" & DevNombreSQL(Rs!Nommacta)
        Sql = RC & Sql & "')"
    
        
        Conn.Execute Sql
    
        Rs.MoveNext
    Wend
    Rs.Close
    
    Set Rs = Nothing
    ComprobarCuentas349_DOS = True
    Exit Function
EComprobarCuentas349:
    MuestraError Err.Number, "Comprobar Cuentas 349"
End Function

Private Sub CopiarFicheroHaciend3(Modelo347 As Boolean)
    On Error GoTo ECopiarFichero347
    MsgBox "El archivo se ha generado con exito.", vbInformation
    Sql = ""
    cd1.CancelError = True
    cd1.ShowSave
    If Modelo347 Then
        Sql = App.Path & "\mod347.txt"
    Else
        Sql = App.Path & "\mod349.txt"
    End If
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El fichero ya existe. ¿Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Sql = ""
        End If
        If Sql <> "" Then
            FileCopy Sql, cd1.FileName
            MsgBox Space(20) & "Copia efectuada correctamente" & Space(20), vbInformation
        End If
    End If
    Exit Sub
ECopiarFichero347:
    If Err.Number <> 32755 Then MuestraError Err.Number, "Copiar fichero 347"
    
End Sub

Private Function ExisteEntrada() As Boolean
    Sql = "Select importe from tmp347tot  where codusu = " & vUsu.Codigo & " and cliprov =" & Rs!cliprov & " AND nif ='" & Rs!NIF & "';"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ExisteEntrada = True
        Importe = miRsAux!Importe
    Else
        ExisteEntrada = False
    End If
    miRsAux.Close
End Function

