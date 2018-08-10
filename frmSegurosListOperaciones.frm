VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSegurosListComunicacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11865
   Icon            =   "frmSegurosListOperaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   11865
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
      Height          =   2415
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   6945
      Begin VB.Frame FrameTipoListado 
         Height          =   735
         Left            =   600
         TabIndex        =   31
         Top             =   1440
         Width           =   6135
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Siniestro"
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
            Index           =   2
            Left            =   4440
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Prórroga"
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
            Left            =   2520
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optAsegAvisos 
            Caption         =   "Falta de pago"
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
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
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
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   900
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
         Index           =   0
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   900
         Width           =   1305
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
         TabIndex        =   27
         Top             =   3630
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
         TabIndex        =   26
         Top             =   3990
         Width           =   4095
      End
      Begin VB.Label lblNumFactu 
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
         Left            =   2580
         TabIndex        =   25
         Top             =   2700
         Width           =   4035
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   4680
         Top             =   960
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   1560
         Top             =   960
         Width           =   240
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
         Index           =   16
         Left            =   3960
         TabIndex        =   24
         Top             =   990
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
         Index           =   17
         Left            =   720
         TabIndex        =   23
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Height          =   240
         Index           =   18
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   690
      End
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
      TabIndex        =   10
      Top             =   2520
      Width           =   6915
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
         TabIndex        =   20
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
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
         TabIndex        =   19
         Top             =   1200
         Width           =   1515
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
         TabIndex        =   18
         Top             =   1680
         Width           =   975
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
         TabIndex        =   17
         Top             =   2160
         Width           =   975
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
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   720
         Width           =   3345
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
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   13
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   12
         Top             =   1680
         Width           =   255
      End
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
         TabIndex        =   11
         Top             =   720
         Width           =   1515
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
      Height          =   5145
      Left            =   7110
      TabIndex        =   9
      Top             =   0
      Width           =   4665
      Begin MSComctlLib.ListView ListView3 
         Height          =   3540
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   6244
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
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3960
         TabIndex        =   30
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3480
         Picture         =   "frmSegurosListOperaciones.frx":000C
         ToolTipText     =   "Quitar al Debe"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4080
         Picture         =   "frmSegurosListOperaciones.frx":0156
         ToolTipText     =   "Puntear al Debe"
         Top             =   1080
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
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1110
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
      Left            =   10440
      TabIndex        =   7
      Top             =   5340
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
      Left            =   8880
      TabIndex        =   6
      Top             =   5340
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
      Left            =   150
      TabIndex        =   8
      Top             =   5310
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Left            =   1980
      TabIndex        =   29
      Top             =   5400
      Width           =   6270
   End
End
Attribute VB_Name = "frmSegurosListComunicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 415

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
Public numero As String
    ' 0.- Comunicacion seguro
    ' 1.- Avisos aseguradora

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private SQL As String
Dim Cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim PrimeraVez As Boolean
Dim Cancelado As Boolean

'Dim vNIF As String


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
    cmdAccion(0).Enabled = False
    cmdAccionClick Index
    cmdAccion(0).Enabled = True
End Sub

Private Sub cmdAccionClick(Index As Integer)
Dim B As Boolean
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    tabla = " "
 
    Screen.MousePointer = vbHourglass
    B = ComunicaDatosSeguro_ 'CargarTemporales
    Screen.MousePointer = vbDefault
    If Not B Then Exit Sub
    
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
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        'PonFoco txtCuentas(0)
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
    Me.Icon = frmppal.Icon
        
    'Otras opciones
    FrameTipoListado.visible = False
    If numero = 0 Then
        Me.Caption = "Listado operaciones aseguradas"
        Label3(18).Caption = "Fecha factura"
    Else
        FrameTipoListado.BorderStyle = 0
        FrameTipoListado.visible = True
        Me.Caption = "Listado avisos seguro"
        Label3(18).Caption = "Fecha aviso"
    End If
    
    For i = 0 To 1
        Me.ImgFec(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
   
    
     
    CargarListViewEmpresas 1
    If numero = 1 Then Me.txtfecha(0).Text = "01" & Format(Now, "/mm/yyyy")
    Me.txtfecha(1).Text = Format(DateAdd("d", 0, Now), "dd/mm/yyyy")
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    

    
End Sub

Private Sub CargarListViewEmpresas(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim Prohibidas As String
Dim IT
Dim Aux As String
    
    On Error GoTo ECargarList

    'Los encabezados
    ListView3.ColumnHeaders.Clear

    ListView3.ColumnHeaders.Add , , "Empresa", 3800
    


    Set Rs = New ADODB.Recordset

    Prohibidas = DevuelveProhibidas
    
    ListView3.ListItems.Clear
    Aux = "Select * from usuarios.empresasariconta where tesor>0"
    
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        Aux = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, Aux) = 0 Then
            Set IT = ListView3.ListItems.Add
            IT.Key = "C" & Rs!codempre
            If vEmpresa.codempre = Rs!codempre Then IT.Checked = True
            IT.Text = Rs!nomempre
            IT.Tag = Rs!codempre
            IT.ToolTipText = Rs!CONTA
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

Private Function DevuelveProhibidas() As String
Dim i As Integer


    On Error GoTo EDevuelveProhibidas
    
    DevuelveProhibidas = ""

    Set miRsAux = New ADODB.Recordset

    i = vUsu.Codigo Mod 100
    miRsAux.Open "Select * from usuarios.usuarioempresasariconta WHERE codusu =" & i, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    DevuelveProhibidas = ""
    While Not miRsAux.EOF
        DevuelveProhibidas = DevuelveProhibidas & miRsAux.Fields(1) & "|"
        miRsAux.MoveNext
    Wend
    If DevuelveProhibidas <> "" Then DevuelveProhibidas = "|" & DevuelveProhibidas
    miRsAux.Close
    Exit Function
EDevuelveProhibidas:
    MuestraError Err.Number, "Cargando empresas prohibidas"
    Err.Clear
End Function






Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub



Private Sub frmF_Selec(vFecha As Date)
    txtfecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgCheck_Click(Index As Integer)
Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    For i = 1 To ListView3.ListItems.Count
        ListView3.ListItems(i).Checked = False
    Next i

    
    Screen.MousePointer = vbDefault

End Sub


'Private Sub imgCuentas_Click(Index As Integer)
'    Sql = ""
'    AbiertoOtroFormEnListado = True
'    Set frmCCtas = New frmColCtas
'    frmCCtas.DatosADevolverBusqueda = True
'    frmCCtas.Show vbModal
'    Set frmCCtas = Nothing
'    If Sql <> "" Then
'    '    Me.txtCuentas(Index).Text = RecuperaValor(Sql, 1)
'    '    Me.txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
'    '    If txtCuentas(Index).Text <> "" Then
'            Me.txtNIF(0).Text = ""
'            Me.txtNIF(0).Text = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", RecuperaValor(Sql, 1), "T")
'    '    End If
'
'    Else
'      '  QuitarPulsacionMas Me.txtCuentas(Index)
'    End If
'
'    'PonFoco Me.txtCuentas(Index)
'    AbiertoOtroFormEnListado = False
'End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
        IndCodigo = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtfecha(Index).Text <> "" Then frmF.Fecha = CDate(txtfecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtfecha(Index)
    End Select
    
    Screen.MousePointer = vbDefault

End Sub






Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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




Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    Case "imgCuentas"
      '  imgCuentas_Click Indice
    End Select
End Sub


Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    'FALTA###
'
'    Sql = "SELECT `tmptesoreriacomun`.`texto1` Nif , `tmptesoreriacomun`.`texto2` Conta, `tmptesoreriacomun`.`opcion` BD, `tmptesoreriacomun`.`texto5` Nombre, `tmptesoreriacomun`.`texto3` NroFra, `tmptesoreriacomun`.`fecha1` FecFra, `tmptesoreriacomun`.`fecha2` FecVto, `tmptesoreriacomun`.`importe1` Gasto, `tmptesoreriacomun`.`importe2` Recibo"
'    Sql = Sql & " FROM   `tmptesoreriacomun` `tmptesoreriacomun`"
'    Sql = Sql & " WHERE `tmptesoreriacomun`.codusu = " & vUsu.Codigo
'    Sql = Sql & " ORDER BY `tmptesoreriacomun`.`texto1`, `tmptesoreriacomun`.`texto2`, `tmptesoreriacomun`.`opcion`, `tmptesoreriacomun`.`fecha1`"
'
'    'LLamos a la funcion
'    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
'
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = Format(IdPrograma, "0000") & "-" & Format(numero, "00")
    
  
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    
    cadFormula = "{tmptesoreriacomun.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 50
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub





Private Function DameEmpresa(ByVal S As String) As String
    DameEmpresa = "NO ENCONTRADA"
    For i = 1 To ListView3.ListItems.Count
        If ListView3.ListItems(i).Tag = S Then
            DameEmpresa = DevNombreSQL(ListView3.ListItems(i).Text)
            Exit For
        End If
    Next i
  
End Function




Private Sub txtfecha_LostFocus(Index As Integer)
    txtfecha(Index).Text = Trim(txtfecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtfecha(Index)
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtfecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtfecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    

    
    If txtfecha(1).Text = "" Then
       
            SQL = "Fecha 'hasta' es campo obligado para considerar la fecha de baja de los asegurados." & vbCrLf
            SQL = SQL & "En el listado saldrán aquellos que si tienen fecha de baja , es superior al hasta solicitado "
         
            MsgBox SQL, vbExclamation
            Exit Function
    End If
    
    SQL = ""
    For i = 1 To ListView3.ListItems.Count
        If ListView3.ListItems(i).Checked Then
            SQL = "O"
            Exit For
        End If
    Next i
    If SQL = "" Then
        MsgBox "Seleccione al menos una empresa", vbExclamation
        Exit Function
    End If

    DatosOK = True


End Function

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub






'******************************************************************************
'******************************************************************************
'
'******************************************************************************
'******************************************************************************



Private Function ComunicaDatosSeguro_() As Boolean
Dim K As Integer

    ComunicaDatosSeguro_ = False
    
    CONT = 0
    Conn.Execute "DELETE from tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
    NumRegElim = 0
    
    For K = 1 To Me.ListView3.ListItems.Count
        If Me.ListView3.ListItems(K).Checked Then
            DatosSeguroUnaEmpresa CInt(ListView3.ListItems(K).Tag)
      
            SQL = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
            If SQL <> "" Then NumRegElim = Val(SQL)
        End If
    Next
    
    
    
    If NumRegElim > 0 Then
        SQL = "DELETE from tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " AND importe1<=0"
        
        
        
    
    
        '   Conn.Execute SQL
        SQL = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
        If SQL <> "" Then
            NumRegElim = Val(SQL)
        Else
            NumRegElim = 0
        End If
        
        
        ComunicaDatosSeguro_ = NumRegElim > 0
        If numero = 0 And NumRegElim > 0 Then
            SQL = "Select texto5 from tmptesoreriacomun WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not miRsAux.EOF
                SQL = miRsAux!texto5
                If SQL = "" Then
                    SQL = "ESPAÑA"
                Else
                    If InStr(1, SQL, " ") > 0 Then
                        SQL = Mid(SQL, 3)
                    Else
                        SQL = "" 'no updateamos
                    End If
                End If
                If SQL <> "" Then
                    'FALTA######
                    SQL = "UPDATE usuarios.ztesoreriacomun set texto5='" & DevNombreSQL(SQL) & "' WHERE codusu ="
                    SQL = SQL & vUsu.Codigo & " AND texto5='" & DevNombreSQL(miRsAux!texto5) & "'"
                    Conn.Execute SQL
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
    Else
        MsgBox "No existen datos para mostrar", vbExclamation
    End If
End Function

Private Sub DatosSeguroUnaEmpresa(NumConta As Integer)
Dim Importe As Currency
Dim B As Boolean
   
    
    If numero = 0 Then
        'Listado facturas
        SQL = "INSERT INTO tmptesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,"
        SQL = SQL & " importe1,  importe2,texto5) "
        RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,factcli.nifdatos,concat(numserie,right(concat('0000000',numfactu),8)),"
        RC = RC & " factcli.nommacta, totfaccl ,credicon,if(cuentas.codpais is null,'',cuentas.codpais) from"
        RC = RC & " ariconta" & NumConta & ".factcli,ariconta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
        RC = RC & " WHERE factcli.codmacta=cuentas.codmacta  and numpoliz<>''  and  (fecbajcre  is null or fecbajcre>'') and fecfactu>= fecconce"
        
        
        RC = RC & " AND (fecbajcre  is null or fecbajcre>'" & Format(txtfecha(1).Text, FormatoFecha) & "')"
        
        'Contemplamos facturas desde la fecha de concesion
        RC = RC & " and fecfactu>= fecconce"
         
        'D/H fecha
        If Me.txtfecha(0).Text <> "" Then RC = RC & " AND fecfactu >='" & Format(txtfecha(0).Text, FormatoFecha) & "'"
        If Me.txtfecha(1).Text <> "" Then RC = RC & " AND fecfactu <='" & Format(txtfecha(1).Text, FormatoFecha) & "'"
    
        
        
        
        
        
        SQL = SQL & RC & " ORDER BY numserie,fecfactu,numfactu"
        
        Conn.Execute SQL
        
        
    Else
    
        'feccomunica,fecprorroga,fecsiniestro
        SQL = ""
        If Me.optAsegAvisos(0).Value Then
            Cad = "feccomunica"
        ElseIf Me.optAsegAvisos(1).Value Then
            Cad = "fecprorroga"
        Else
            Cad = "fecsiniestro"
        End If
        RC = ""
        Msg = ""
        If Me.txtfecha(0).Text <> "" Then
            RC = RC & " AND " & Cad & ">='" & Format(txtfecha(0).Text, FormatoFecha) & "'"
            Msg = Msg & " desde " & txtfecha(0).Text
        End If
        If Me.txtfecha(1).Text <> "" Then
            RC = RC & " AND " & Cad & "<='" & Format(txtfecha(1).Text, FormatoFecha) & "'"
            Msg = Msg & " hasta " & txtfecha(1).Text
        End If
        If RC = "" Then RC = " AND " & Cad & ">='1900-01-01'"
        SQL = RC
        
        If Me.optAsegAvisos(0).Value Then
            Cad = "falta pago"
      
        Else
            If Me.optAsegAvisos(1).Value Then
            Cad = "prorroga"
            Else
                Cad = "siniestro"
            End If
        End If
        If Msg <> "" Then
            Msg = "Fecha " & Cad & " " & Msg
        Else
            Msg = ""
        End If
        cadParam = cadParam & "Titulo= """ & UCase(Cad) & """|cuenta= """ & Msg & """|"
        numParam = numParam + 2
        
        'RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
        'If RC <> "" Then SQL = SQL & " AND " & RC
        'RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
        'If RC <> "" Then SQL = SQL & " AND " & RC
        
        
        'ORDENACION
        If Me.optAsegAvisos(0).Value Then
            RC = "feccomunica"
        ElseIf Me.optAsegAvisos(1).Value Then
            RC = "fecprorroga"
        Else
            RC = "fecsiniestro"
        End If
        
        Cad = "Select numserie,numfactu,numorden,fecvenci,impvenci,impcobro,gastos,fecfactu,devuelto,cobros.codmacta,nommacta,numpoliz"
        Cad = Cad & ",credicon," & RC & " LaFecha" 'alias
        Cad = Cad & "  FROM cobros,cuentas,formapago where cobros.codmacta= cuentas.codmacta AND numpoliz<>"""""
        Cad = Cad & " and cobros.codforpa=formapago.codforpa "
        If SQL <> "" Then Cad = Cad & SQL
    
        Cad = Cad & " ORDER BY " & RC & ",cobros.codmacta"
      
    
      
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
       'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,faviso,fvto,impvto,disponible,vencida
        Cad = "INSERT INTO tmptesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,fecha3,importe1,importe2,opcion) VALUES "
        RC = ""
        
        While Not miRsAux.EOF
            
            CONT = CONT + 1
            SQL = ", (" & vUsu.Codigo & "," & CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            SQL = SQL & DevNombreSQL(miRsAux!numpoliz) & "'"
            SQL = SQL & ",'" & miRsAux!NUmSerie & Format(miRsAux!NumFactu, "00000000") & "',"  'texto4
            'Fecha fac
            SQL = SQL & DBSet(miRsAux!FecFactu, "F") & ","
            'Fecha aviso
            SQL = SQL & DBSet(miRsAux!lafecha, "F") & ","
            'Fecha vto
            SQL = SQL & DBSet(miRsAux!FecVenci, "F")
            
            SQL = SQL & "," & TransformaComasPuntos(CStr(miRsAux!ImpVenci))
            SQL = SQL & "," & TransformaComasPuntos(CStr(DBLet(miRsAux!Gastos, "N")))
            'Devuelto
            SQL = SQL & "," & DBLet(miRsAux!Devuelto, "N") & ")"
        
            RC = RC & SQL
            miRsAux.MoveNext
            
            If miRsAux.EOF Then
                B = True
            Else
                B = Len(RC) > 2000
            End If
            If B Then
                RC = Mid(RC, 2)
                Conn.Execute Cad & RC
                RC = ""
            End If
            
            
        Wend
        miRsAux.Close
    
        
      
      
    End If
End Sub

