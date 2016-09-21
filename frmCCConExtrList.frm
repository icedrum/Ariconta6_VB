VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCCConExtrList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7305
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
      Height          =   3705
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNCCoste 
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
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1050
         Width           =   4605
      End
      Begin VB.TextBox txtNCCoste 
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
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1470
         Width           =   4605
      End
      Begin VB.TextBox txtAno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3270
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2700
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmCCConExtrList.frx":0000
         Left            =   1230
         List            =   "frmCCConExtrList.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2670
         Width           =   1935
      End
      Begin VB.TextBox txtAno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3270
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2220
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmCCConExtrList.frx":0004
         Left            =   1230
         List            =   "frmCCConExtrList.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2220
         Width           =   1935
      End
      Begin VB.TextBox txtCCoste 
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
         Top             =   1050
         Width           =   795
      End
      Begin VB.TextBox txtCCoste 
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
         Top             =   1470
         Width           =   795
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   6330
         TabIndex        =   30
         Top             =   300
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
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   6
         Left            =   930
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   7
         Left            =   930
         Top             =   1500
         Width           =   255
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   1080
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
         Left            =   240
         TabIndex        =   20
         Top             =   2640
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
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Centro de Coste"
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
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Mes / Año"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1410
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
      Left            =   5880
      TabIndex        =   4
      Top             =   6630
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
      Left            =   4290
      TabIndex        =   2
      Top             =   6630
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
      Top             =   6570
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
      Top             =   3720
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
      Left            =   5850
      TabIndex        =   27
      Top             =   6630
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1530
      TabIndex        =   31
      Top             =   6630
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
End
Attribute VB_Name = "frmCCConExtrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1002

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


Public Coste As String
Public Descripcion As String
Public FecDesde As String
Public FecHasta As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmCCoste  As frmBasico
Attribute frmCCoste.VB_VarHelpID = -1

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
Dim FechaInicio As String
Dim fechafin As String

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
    
    
    If Not PonerDesdeHasta("hlinapu.codccost", "CCO", Me.txtCCoste(6), Me.txtNCCoste(6), Me.txtCCoste(7), Me.txtNCCoste(7), "pDHCoste=""") Then Exit Sub
    
    If Not MontaSQL Then Exit Sub

    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True
    
    
    If Not HayRegParaInforme("tmpsaldoscc", "codusu=" & vUsu.Codigo) Then Exit Sub
    
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
    
    If Legalizacion <> "" Then
        CadenaDesdeOtroForm = "OK"
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

Private Sub Form_Activate()
Dim CONT As Integer

    If PrimeraVez Then
        PrimeraVez = False
    End If
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
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
    'Otras opciones
    Me.Caption = "Extractos por Centro de Coste"

    For I = 6 To 7
        Me.imgCCoste(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    PrimeraVez = True
     
    If Coste <> "" Then
        txtCCoste(6).Text = Coste
        txtCCoste(7).Text = Coste
        txtNCCoste(6).Text = Descripcion
        txtNCCoste(7).Text = Descripcion
    End If
     
     
    CargarComboFecha
     
    cmbFecha(0).ListIndex = Month(vParam.fechaini) - 1
    cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1

    txtAno(0).Text = Year(vParam.fechaini)
    txtAno(1).Text = Year(vParam.fechafin)
   
    If FecDesde <> "" Then
        txtAno(0).Text = Year(CDate(FecDesde))
        cmbFecha(0).ListIndex = Month(CDate(FecDesde)) - 1
    End If
    
    If FecHasta <> "" Then
        txtAno(1).Text = Year(CDate(FecHasta))
        cmbFecha(1).ListIndex = Month(CDate(FecHasta)) - 1
    End If
   
 
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
    
End Sub

Private Sub frmCCoste_DatoSeleccionado(CadenaSeleccion As String)
    txtCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub ImgCCoste_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmCCoste = New frmBasico
    AyudaCC frmCCoste
    Set frmCCoste = Nothing
    
    PonFoco txtCCoste(Index)

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


Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCCoste_GotFocus(Index As Integer)
    ConseguirFoco txtCCoste(Index), 3
End Sub


Private Sub txtCCoste_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCCoste", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgCCoste"
        ImgCCoste_Click Indice
    End Select
    
End Sub

Private Sub txtCCoste_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCCoste(Index).Text = Trim(txtCCoste(Index).Text)
    
    Select Case Index
        Case 6, 7 'Centros de Coste
            If txtCCoste(Index).Text <> "" Then txtCCoste(Index).Text = UCase(txtCCoste(Index).Text)
            txtNCCoste(Index) = PonerNombreDeCod(txtCCoste(Index), "ccoste", "nomccost", "codccost", "T")
            
    End Select

End Sub



Private Sub AccionesCSV()
Dim SQL2 As String
Dim Tipo As Byte

       
    SQL = "SELECT tmpsaldoscc.codccost CCoste, tmpsaldoscc.nomccost Descripcion, tmpsaldoscc.ano Año,tmpsaldoscc.mes Mes, tmpsaldoscc.impmesde Debe, tmpsaldoscc.impmesha Haber, coalesce(tmpsaldoscc.impmesde,0) - coalesce(tmpsaldoscc.impmesha) Saldo "
    SQL = SQL & " FROM  tmpsaldoscc tmpsaldoscc "
    SQL = SQL & " WHERE tmpsaldoscc.codusu = " & vUsu.Codigo
    SQL = SQL & " ORDER BY tmpsaldoscc.codccost, tmpsaldoscc.ano, tmpsaldoscc.mes"
        
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    cadParam = cadParam & "pTipo=" & Tipo & "|"
    numParam = numParam + 1
    
    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & txtAno(0).Text & " a " & cmbFecha(1).Text & " " & txtAno(1).Text & """|"
    numParam = numParam + 1
    
    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "1002-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"SumasySaldos.rpt"

    cadFormula = "{tmpsaldoscc.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String
Dim B As Boolean

    MontaSQL = False
    
    SQL = "delete from tmplinccexplo where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute SQL
    
    SQL = "insert into tmplinccexplo (codusu,codccost,codmacta,linapu,docum,fechaent,ampconce,ctactra,desctra,perD,perH) "
    SQL = SQL & " Select " & vUsu.Codigo & ", hlinapu.codccost CCoste,  hlinapu.codmacta as Cuenta, hlinapu.linliapu, hlinapu.numdocum, hlinapu.fechaent,  "
    SQL = SQL & " hlinapu.ampconce, ctacontr Contrapartida, cuentas1.nommacta Descripción , coalesce(timported,0) Debe, coalesce(timporteh,0) Haber "
    SQL = SQL & " FROM (hlinapu LEFT JOIN cuentas cuentas1 ON hlinapu.ctacontr = cuentas1.codmacta) "
    SQL = SQL & " where (1=1) "
    SQL = SQL & " and hlinapu.fechaent >= " & DBSet(FechaInicio, "F") & " and hlinapu.fechaent <= " & DBSet(fechafin, "F")
    If cadselect <> "" Then SQL = SQL & " and " & cadselect
    
    SQL = SQL & " union "
    SQL = SQL & " Select " & vUsu.Codigo & ", hlinapu.codccost CCoste,  hlinapu.codmacta as Cuenta, hlinapu.linliapu, hlinapu.numdocum, hlinapu.fechaent,  "
    SQL = SQL & " hlinapu.ampconce, ctacontr Contrapartida, cuentas1.nommacta Descripción , coalesce(timported,0) Debe, coalesce(timporteh,0) Haber "
    SQL = SQL & " FROM (hlinapu LEFT JOIN cuentas cuentas1 ON hlinapu.ctacontr = cuentas1.codmacta) INNER JOIN ccoste_lineas ON hlinapu.codccost = ccoste_lineas.codccost "
    SQL = SQL & " where (1=1) "
    SQL = SQL & " and hlinapu.fechaent >= " & DBSet(FechaInicio, "F") & " and hlinapu.fechaent <= " & DBSet(fechafin, "F")
    If cadselect <> "" Then SQL = SQL & " and " & Replace(cadselect, "hlinapu.codccost", "ccoste_lineas.subccost")
    
    
    SQL = SQL & " ORDER BY 1,2,3,4,5 "
    
    Conn.Execute SQL
    
    B = HacerRepartoSubcentrosCoste
    
    
    If B Then
        
        Conn.Execute "delete from tmpsaldoscc where codusu = " & vUsu.Codigo
        
        
        SQL = "insert into tmpsaldoscc (codusu,codccost,nomccost,ano,mes,impmesde,impmesha) "
        SQL = SQL & " select " & vUsu.Codigo & ", tmplinccexplo.codccost, ccoste.nomccost, year(fechaent), month(fechaent), sum(coalesce(perd,0)), sum(coalesce(perh,0)) "
        SQL = SQL & " from tmplinccexplo, ccoste "
        SQL = SQL & " where tmplinccexplo.codccost = ccoste.codccost "
        SQL = SQL & " and codusu = " & DBSet(vUsu.Codigo, "N")
        SQL = SQL & " group by 1,2,3"
        SQL = SQL & " order by 1,2,3"
        
        Conn.Execute SQL
        
        MontaSQL = True
    
    End If
    
End Function

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtAno(0).Text = "" Or txtAno(1).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Function
    End If
    
    If Not ComparaFechasCombos(0, 1, 0, 1) Then Exit Function
    
    If txtAno(0).Text = "" Then
        FechaInicio = "01/01/1900"
    Else
        FechaInicio = "01/" & Format((Me.cmbFecha(0).ListIndex + 1), "00") & "/" & txtAno(0).Text
    End If
    If txtAno(1).Text = "" Then
        fechafin = "31/12/2200"
    Else
        fechafin = "01/" & Format(Me.cmbFecha(1).ListIndex + 1) & "/" & txtAno(1).Text
    End If
    ' a la fecha hasta del listado le sumamos 1 mes y le restamos 1 dia ( para que me de el ultimo dia del mes )
    fechafin = DateAdd("d", -1, DateAdd("m", 1, CDate(fechafin)))
    
    If FechaInicio > fechafin Then
        MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
        Exit Function
    End If
    
    

    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer

QueCombosFechaCargar "0|1|"

End Sub




Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    cad = RecuperaValor(Lista, L)
    If cad <> "" Then
        I = Val(cad)
        With cmbFecha(I)
            .Clear
            For CONT = 1 To 12
                RC = "25/" & CONT & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next CONT
        End With
    End If
    L = L + 1
Loop Until cad = ""
End Sub


Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCCoste(Indice1).Text <> "" And txtCCoste(Indice2).Text <> "" Then
        L1 = Len(txtCCoste(Indice1).Text)
        L2 = Len(txtCCoste(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCCoste(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCCoste(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function

Private Function ComparaFechasCombos(Indice1 As Integer, Indice2 As Integer, InCombo1 As Integer, InCombo2 As Integer) As Boolean
    ComparaFechasCombos = False
    If txtAno(Indice1).Text <> "" And txtAno(Indice2).Text <> "" Then
        If Val(txtAno(Indice1).Text) > Val(txtAno(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        Else
            If Val(txtAno(Indice1).Text) = Val(txtAno(Indice2).Text) Then
                If Me.cmbFecha(InCombo1).ListIndex > Me.cmbFecha(InCombo2).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    ComparaFechasCombos = True
End Function


'Siempre k la fecha no este en fecha siguiente
Private Function HayAsientoCierre(Mes As Byte, Anyo As Integer, Optional Contabilidad As String) As Boolean
Dim C As String
    HayAsientoCierre = False
    'C = "01/" & CStr(Me.cmbFecha(1).ListIndex + 1) & "/" & txtAno(1).Text
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



Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtTitulo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Function HacerRepartoSubcentrosCoste() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim ImporteTot As Currency
Dim ImporteLinea As Currency
Dim UltSubCC As String
Dim Nregs As Long

    On Error GoTo eHacerRepartoSubcentrosCoste

    HacerRepartoSubcentrosCoste = False
    
    ' hacemos el desdoble
    SQL = "select * from tmplinccexplo where codusu = " & DBSet(vUsu.Codigo, "N") & " and codccost in (select ccoste.codccost from ccoste inner join ccoste_lineas on ccoste.codccost = ccoste_lineas.codccost) "

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Nregs = TotalRegistrosConsulta(SQL)

    If Nregs <> 0 Then
        pb2.Visible = True
        CargarProgres pb2, Nregs
    End If


    While Not RS.EOF
        IncrementarProgres pb2, 1
        
        SQL2 = "select ccoste.codccost, subccost, porccost from ccoste inner join ccoste_lineas on ccoste.codccost = ccoste_lineas.codccost where ccoste.codccost =  " & DBSet(RS!codccost, "T")

        ImporteTot = 0
        UltSubCC = ""

        Set Rs2 = New ADODB.Recordset

        Rs2.Open SQL2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            SQL = "insert into tmplinccexplo (codusu,codccost,codmacta,linapu,docum,fechaent,ampconce,ctactra,desctra,perD,perH,desdoblado) values ("
            SQL = SQL & vUsu.Codigo & "," & DBSet(Rs2!subccost, "T") & "," & DBSet(RS!codmacta, "T") & "," & DBSet(RS!LINAPU, "T") & ","
            SQL = SQL & DBSet(RS!DOCUM, "T") & "," & DBSet(RS!FechaEnt, "F") & "," & DBSet(RS!Ampconce, "T") & "," & DBSet(RS!ctactra, "T") & ","
            SQL = SQL & DBSet(RS!desctra, "T") & ","

            If DBLet(RS!perd, "N") <> 0 Then
                ImporteLinea = Round(DBLet(RS!perd, "N") * DBLet(Rs2!porccost, "N") / 100, 2)
                SQL = SQL & DBSet(ImporteLinea, "N") & ",0,1)"
            Else
                ImporteLinea = Round(DBLet(RS!perh, "N") * DBLet(Rs2!porccost, "N") / 100, 2)
                SQL = SQL & "0," & DBSet(ImporteLinea, "N") & ",1)"
            End If

            Conn.Execute SQL

            ImporteTot = ImporteTot + ImporteLinea

            UltSubCC = Rs2!subccost

            Rs2.MoveNext
        Wend

        If DBLet(RS!perd, "N") <> 0 Then
            If ImporteTot <> DBLet(RS!perd, "N") Then
                SQL = "update tmplinccexplo set perd = perd + (" & DBSet(Round(DBLet(RS!perd, "N") - ImporteTot, 2), "N") & ")"
                SQL = SQL & " where codusu = " & vUsu.Codigo
                SQL = SQL & " and codccost = " & DBSet(UltSubCC, "T")
                SQL = SQL & " and codmacta = " & DBSet(RS!codmacta, "T")
                SQL = SQL & " and fechaent = " & DBSet(RS!FechaEnt, "F")
                SQL = SQL & " and linapu = " & DBSet(RS!LINAPU, "N")
                SQL = SQL & " and docum = " & DBSet(RS!DOCUM, "T")
                SQL = SQL & " and ampconce = " & DBSet(RS!Ampconce, "T")
                SQL = SQL & " and desdoblado = 1"

                Conn.Execute SQL
            End If
        Else
            If ImporteTot <> DBLet(RS!perh, "N") Then
                SQL = "update tmplinccexplo set perh = perh + (" & DBSet(Round(DBLet(RS!perh, "N") - ImporteTot, 2), "N") & ")"
                SQL = SQL & " where codusu = " & vUsu.Codigo
                SQL = SQL & " and codccost = " & DBSet(UltSubCC, "T")
                SQL = SQL & " and codmacta = " & DBSet(RS!codmacta, "T")
                SQL = SQL & " and fechaent = " & DBSet(RS!FechaEnt, "F")
                SQL = SQL & " and linapu = " & DBSet(RS!LINAPU, "N")
                SQL = SQL & " and docum = " & DBSet(RS!DOCUM, "T")
                SQL = SQL & " and ampconce = " & DBSet(RS!Ampconce, "T")
                SQL = SQL & " and desdoblado = 1"

                Conn.Execute SQL
            End If
        End If

        SQL = "delete from tmplinccexplo where codusu = " & vUsu.Codigo
        SQL = SQL & " and codccost = " & DBSet(RS!codccost, "T")
        SQL = SQL & " and codmacta = " & DBSet(RS!codmacta, "T")
        SQL = SQL & " and fechaent = " & DBSet(RS!FechaEnt, "F")
        SQL = SQL & " and linapu = " & DBSet(RS!LINAPU, "N")
        SQL = SQL & " and docum = " & DBSet(RS!DOCUM, "T")
        SQL = SQL & " and ampconce = " & DBSet(RS!Ampconce, "T")
        SQL = SQL & " and desdoblado = 0"

        Conn.Execute SQL

        Set Rs2 = Nothing


        RS.MoveNext
    Wend

    Set RS = Nothing
    
    'falta el borrado de los que no tocan
    If txtCCoste(6).Text <> "" Or txtCCoste(7).Text <> "" Then
        SQL = "delete from tmplinccexplo where codusu = " & vUsu.Codigo
        SQL = SQL & " and not codccost in (select codccost from ccoste where (1=1) "
        If txtCCoste(6).Text <> "" Then SQL = SQL & " and codccost >= " & DBSet(txtCCoste(6).Text, "T")
        If txtCCoste(7).Text <> "" Then SQL = SQL & " and codccost <= " & DBSet(txtCCoste(7).Text, "T")
        SQL = SQL & ")"
        
        Conn.Execute SQL
    End If
    HacerRepartoSubcentrosCoste = True
    pb2.Visible = False
    Exit Function
    
eHacerRepartoSubcentrosCoste:
    MuestraError Err.Number, "Reparto Subcentros de Coste", Err.Description
    pb2.Visible = False
End Function


