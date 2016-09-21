VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresuList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7185
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
      TabIndex        =   17
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
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   30
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
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1470
         Width           =   4185
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
         Height          =   360
         Index           =   1
         Left            =   3270
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2670
         Width           =   885
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
         ItemData        =   "frmPresuList.frx":0000
         Left            =   1230
         List            =   "frmPresuList.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
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
         Height          =   360
         Index           =   0
         Left            =   3270
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2190
         Width           =   885
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
         ItemData        =   "frmPresuList.frx":0004
         Left            =   1230
         List            =   "frmPresuList.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2190
         Width           =   1935
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
         Index           =   0
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
         Index           =   1
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   6300
         TabIndex        =   31
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
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   990
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   990
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   2280
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
         TabIndex        =   22
         Top             =   690
         Width           =   960
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
         TabIndex        =   21
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
      Left            =   5820
      TabIndex        =   8
      Top             =   6510
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
      Left            =   4200
      TabIndex        =   6
      Top             =   6510
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
      TabIndex        =   9
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
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1830
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
      Left            =   5820
      TabIndex        =   28
      Top             =   6510
      Width           =   1215
   End
End
Attribute VB_Name = "frmPresuList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1101

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
            

    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True

    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmpevolsal", "tmpevolsal.codusu=" & vUsu.Codigo) Then Exit Sub
    
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

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Load()
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Presupuestos sobre Cuentas"

    For I = 0 To 1
        Me.imgCuentas(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    PrimeraVez = True
     
    CargarComboFecha
    
    'Fecha inicial
    cmbFecha(0).ListIndex = Month(vParam.fechaini) - 1
    cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1
    
    txtAno(0).Text = Year(vParam.fechaini)
    txtAno(1).Text = Year(vParam.fechafin)
   
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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


Private Sub txtAno_GotFocus(Index As Integer)
    ConseguirFoco txtAno(Index), 3
End Sub

Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
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


Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        LanzaFormAyuda "imgCuentas", Index
    Else
        KEYpress KeyAscii
    End If
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
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        txtNCta(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 0, 1 'Cuentas
            
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
    
    SQL = "select cta Cuenta , nomcta Titulo, aperturad, aperturah, case when coalesce(aperturad,0) - coalesce(aperturah,0) > 0 then concat(coalesce(aperturad,0) - coalesce(aperturah,0),'D') when coalesce(aperturad,0) - coalesce(aperturah,0) < 0 then concat(coalesce(aperturah,0) - coalesce(aperturad,0),'H') when coalesce(aperturad,0) - coalesce(aperturah,0) = 0 then 0 end Apertura, "
    SQL = SQL & " acumantd AcumAnt_deudor, acumanth AcumAnt_acreedor, acumperd AcumPer_deudor, acumperh AcumPer_acreedor, "
    SQL = SQL & " totald Saldo_deudor, totalh Saldo_acreedor, case when coalesce(totald,0) - coalesce(totalh,0) > 0 then concat(coalesce(totald,0) - coalesce(totalh,0),'D') when coalesce(totald,0) - coalesce(totalh,0) < 0 then concat(coalesce(totalh,0) - coalesce(totald,0),'H') when coalesce(totald,0) - coalesce(totalh,0) = 0 then 0 end Saldo"
    SQL = SQL & " from tmpbalancesumas where codusu = " & vUsu.Codigo
    SQL = SQL & " order by 1 "
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & txtAno(0).Text & " a " & cmbFecha(1).Text & " " & txtAno(1).Text & """|"
    numParam = numParam + 1
    
    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "1101-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"Presupuestos.rpt"

    cadFormula = "{tmpevolsal.codusu} = " & vUsu.Codigo

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
    
    If Not PonerDesdeHasta("presupuestos.codmacta", "CTA", Me.txtCta(0), Me.txtNCta(0), Me.txtCta(1), Me.txtNCta(1), "pDHCuentas=""") Then Exit Function
    
    CargarTemporal
    
    MontaSQL = True
           
End Function




Private Function DatosOK() As Boolean
    
    DatosOK = False
    

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



Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Sub CargarTemporal()
Dim SQL As String
Dim SQL2 As String
Dim FechaI As String
Dim FechaF As String
Dim CtaAnt As String
Dim imp1 As Currency
Dim Imp2 As Currency
Dim imp3 As Currency
Dim imp4 As Currency
Dim imp5 As Currency
Dim imp6 As Currency
Dim imp7 As Currency
Dim imp8 As Currency
Dim imp9 As Currency
Dim imp10 As Currency
Dim imp11 As Currency
Dim imp12 As Currency

Dim NuevaCampana As Boolean
Dim Campana As Integer
Dim CampanaAnt As Integer
Dim HayReg As Boolean
Dim NomAnt As String
Dim EjerciciosPartidos As Boolean

    On Error GoTo eCargarTemporal

    SQL = "delete from tmpevolsal where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    FechaI = "1900-01-01"
    FechaF = "2100-12-31'"
    
    If cmbFecha(0).ListIndex <> -1 And txtAno(0).Text <> "" Then FechaI = Format(txtAno(0), "0000") & "-" & Format(cmbFecha(0).ListIndex + 1, "00") & "-01"
    If cmbFecha(1).ListIndex <> -1 And txtAno(1).Text <> "" Then FechaF = Format(txtAno(1), "0000") & "-" & Format(cmbFecha(1).ListIndex + 1, "00") & "-01"
    
    
    SQL2 = "select presupuestos.codmacta, cuentas.nommacta, anopresu, mespresu, imppresu from presupuestos, cuentas where presupuestos.codmacta = cuentas.codmacta and "
    SQL2 = SQL2 & " date(concat(anopresu,'-',right(concat('00',mespresu),2),'-01')) between " & DBSet(FechaI, "F") & " and " & DBSet(FechaF, "F")
    If cadselect <> "" Then SQL2 = SQL2 & " and " & cadselect
    SQL2 = SQL2 & " order by 1,2,3,4"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        CtaAnt = DBLet(RS!codmacta)
        NomAnt = DBLet(RS!Nommacta)
        Campana = DBLet(RS!anopresu)
    
    
        If DBLet(RS!mespresu) < Month(vParam.fechaini) Then
            Campana = Campana - 1
        End If
        CampanaAnt = Campana
        
        imp1 = 0
        Imp2 = 0
        imp3 = 0
        imp4 = 0
        imp5 = 0
        imp6 = 0
        imp7 = 0
        imp8 = 0
        imp9 = 0
        imp10 = 0
        imp11 = 0
        imp12 = 0
        
        EjerciciosPartidos = (Year(vParam.fechaini) <> Year(vParam.fechafin))
        
        HayReg = False
        
        While Not RS.EOF
            HayReg = True
            
            If CtaAnt <> DBLet(RS!codmacta) Or CampanaAnt <> Campana Then
                SQL = "insert into tmpevolsal (codusu, codmacta, nommacta, apertura, "
                SQL = SQL & "importemes1, importemes2, importemes3, importemes4, importemes5, importemes6, importemes7, importemes8, importemes9, importemes10, importemes11, importemes12) values ("
                SQL = SQL & vUsu.Codigo & "," & DBSet(CtaAnt, "T") & "," & DBSet(NomAnt, "T") & "," & DBSet(CampanaAnt, "N") & ","
        
                SQL2 = SQL & DBSet(imp1, "N") & "," & DBSet(Imp2, "N") & "," & DBSet(imp3, "N") & "," & DBSet(imp4, "N") & "," & DBSet(imp5, "N") & "," & DBSet(imp6, "N") & "," & DBSet(imp7, "N") & "," & DBSet(imp8, "N") & "," & DBSet(imp9, "N") & "," & DBSet(imp10, "N") & "," & DBSet(imp11, "N") & "," & DBSet(imp12, "N") & ")"
        
                Conn.Execute SQL2
                
                imp1 = 0
                Imp2 = 0
                imp3 = 0
                imp4 = 0
                imp5 = 0
                imp6 = 0
                imp7 = 0
                imp8 = 0
                imp9 = 0
                imp10 = 0
                imp11 = 0
                imp12 = 0
                
                If CtaAnt = DBLet(RS!codmacta) Then
                    CampanaAnt = Campana
                Else
                    CtaAnt = DBLet(RS!codmacta)
                    NomAnt = DBLet(RS!Nommacta)
                    Campana = DBLet(RS!anopresu)
                    
                    If DBLet(RS!mespresu) < Month(vParam.fechaini) Then
                        Campana = Campana - 1
                    End If
                End If
                CampanaAnt = Campana
            End If
                
            If EjerciciosPartidos Then
                If DBLet(RS!mespresu) < Month(vParam.fechaini) And DBLet(RS!anopresu) - 1 = Campana Then
                    I = RS!mespresu - Month(vParam.fechaini) + 12 + 1
                Else
                    If DBLet(RS!anopresu) = Campana And DBLet(RS!mespresu) >= Month(vParam.fechaini) Then
                        I = RS!mespresu - Month(vParam.fechaini) + 1
                    Else
                        Campana = DBLet(RS!anopresu)
                        If DBLet(RS!mespresu) < Month(vParam.fechaini) Then
                            I = RS!mespresu - Month(vParam.fechaini) + 12 + 1
                        Else
                            I = RS!mespresu - Month(vParam.fechaini) + 1
                        End If
                        SQL = "insert into tmpevolsal (codusu, codmacta, nommacta, apertura, "
                        SQL = SQL & "importemes1, importemes2, importemes3, importemes4, importemes5, importemes6, importemes7, importemes8, importemes9, importemes10, importemes11, importemes12) values ("
                        SQL = SQL & vUsu.Codigo & "," & DBSet(CtaAnt, "T") & "," & DBSet(NomAnt, "T") & "," & DBSet(CampanaAnt, "N") & ","
                
                        SQL2 = SQL & DBSet(imp1, "N") & "," & DBSet(Imp2, "N") & "," & DBSet(imp3, "N") & "," & DBSet(imp4, "N") & "," & DBSet(imp5, "N") & "," & DBSet(imp6, "N") & "," & DBSet(imp7, "N") & "," & DBSet(imp8, "N") & "," & DBSet(imp9, "N") & "," & DBSet(imp10, "N") & "," & DBSet(imp11, "N") & "," & DBSet(imp12, "N") & ")"
                
                        Conn.Execute SQL2
                        
                        imp1 = 0
                        Imp2 = 0
                        imp3 = 0
                        imp4 = 0
                        imp5 = 0
                        imp6 = 0
                        imp7 = 0
                        imp8 = 0
                        imp9 = 0
                        imp10 = 0
                        imp11 = 0
                        imp12 = 0
                        
                        CtaAnt = DBLet(RS!codmacta)
                        NomAnt = DBLet(RS!Nommacta)
                        Campana = DBLet(RS!anopresu)
                        
                        If DBLet(RS!mespresu) < Month(vParam.fechaini) Then
                            Campana = Campana - 1
                        End If
                        CampanaAnt = Campana
                        
                   End If
                End If
            Else
                Campana = DBLet(RS!anopresu)
                I = RS!mespresu
                
                If CtaAnt <> DBLet(RS!codmacta) Or CampanaAnt <> Campana Then
                    SQL = "insert into tmpevolsal (codusu, codmacta, nommacta, apertura, "
                    SQL = SQL & "importemes1, importemes2, importemes3, importemes4, importemes5, importemes6, importemes7, importemes8, importemes9, importemes10, importemes11, importemes12) values ("
                    SQL = SQL & vUsu.Codigo & "," & DBSet(CtaAnt, "T") & "," & DBSet(NomAnt, "T") & "," & DBSet(CampanaAnt, "N") & ","
            
                    SQL2 = SQL & DBSet(imp1, "N") & "," & DBSet(Imp2, "N") & "," & DBSet(imp3, "N") & "," & DBSet(imp4, "N") & "," & DBSet(imp5, "N") & "," & DBSet(imp6, "N") & "," & DBSet(imp7, "N") & "," & DBSet(imp8, "N") & "," & DBSet(imp9, "N") & "," & DBSet(imp10, "N") & "," & DBSet(imp11, "N") & "," & DBSet(imp12, "N") & ")"
            
                    Conn.Execute SQL2
                    
                    imp1 = 0
                    Imp2 = 0
                    imp3 = 0
                    imp4 = 0
                    imp5 = 0
                    imp6 = 0
                    imp7 = 0
                    imp8 = 0
                    imp9 = 0
                    imp10 = 0
                    imp11 = 0
                    imp12 = 0
                    
                    If CtaAnt = DBLet(RS!codmacta) Then
                        CampanaAnt = Campana
                    Else
                        CtaAnt = DBLet(RS!codmacta)
                        NomAnt = DBLet(RS!Nommacta)
                        Campana = DBLet(RS!anopresu)
                        
                        If DBLet(RS!mespresu) < Month(vParam.fechaini) Then
                            Campana = Campana - 1
                        End If
                    End If
                    CampanaAnt = Campana
                    
                End If
                
                
            End If
            
            Select Case I
                Case 1
                    imp1 = DBLet(RS!imppresu, "N")
                Case 2
                    Imp2 = DBLet(RS!imppresu, "N")
                Case 3
                    imp3 = DBLet(RS!imppresu, "N")
                Case 4
                    imp4 = DBLet(RS!imppresu, "N")
                Case 5
                    imp5 = DBLet(RS!imppresu, "N")
                Case 6
                    imp6 = DBLet(RS!imppresu, "N")
                Case 7
                    imp7 = DBLet(RS!imppresu, "N")
                Case 8
                    imp8 = DBLet(RS!imppresu, "N")
                Case 9
                    imp9 = DBLet(RS!imppresu, "N")
                Case 10
                    imp10 = DBLet(RS!imppresu, "N")
                Case 11
                    imp11 = DBLet(RS!imppresu, "N")
                Case 12
                    imp12 = DBLet(RS!imppresu, "N")
            End Select
            
            RS.MoveNext
    
        Wend
        If HayReg Then
            SQL = "insert into tmpevolsal (codusu, codmacta, nommacta, apertura, "
            SQL = SQL & "importemes1, importemes2, importemes3, importemes4, importemes5, importemes6, importemes7, importemes8, importemes9, importemes10, importemes11, importemes12) values ("
            SQL = SQL & vUsu.Codigo & "," & DBSet(CtaAnt, "T") & "," & DBSet(NomAnt, "T") & "," & DBSet(CampanaAnt, "N") & ","
    
            
            SQL2 = SQL & DBSet(imp1, "N") & "," & DBSet(Imp2, "N") & "," & DBSet(imp3, "N") & "," & DBSet(imp4, "N") & "," & DBSet(imp5, "N") & "," & DBSet(imp6, "N") & "," & DBSet(imp7, "N") & "," & DBSet(imp8, "N") & "," & DBSet(imp9, "N") & "," & DBSet(imp10, "N") & "," & DBSet(imp11, "N") & "," & DBSet(imp12, "N") & ")"
    
            Conn.Execute SQL2
        End If
    
        Set RS = Nothing
    End If


eCargarTemporal:

End Sub


