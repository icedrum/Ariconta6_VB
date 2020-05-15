VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsolidadoFras 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11865
   Icon            =   "frmConsolidadoFras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
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
      Height          =   3015
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txtNIF 
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
         Left            =   1320
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   480
         Width           =   1485
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
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   2340
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
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1920
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
         Index           =   0
         Left            =   2610
         TabIndex        =   25
         Top             =   2340
         Width           =   4035
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
         TabIndex        =   24
         Top             =   2700
         Width           =   4035
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   1080
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "N.I.F."
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
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1020
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   960
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   960
         Top             =   1920
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
         Left            =   270
         TabIndex        =   22
         Top             =   2340
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
         Left            =   270
         TabIndex        =   21
         Top             =   1950
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha factura"
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
         Index           =   18
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   2280
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
      Left            =   150
      TabIndex        =   8
      Top             =   3090
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   11
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   10
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
         TabIndex        =   9
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
      Height          =   5745
      Left            =   7110
      TabIndex        =   7
      Top             =   0
      Width           =   4665
      Begin VB.CheckBox Check1 
         Caption         =   "Clientes"
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
         Index           =   0
         Left            =   720
         TabIndex        =   32
         Top             =   960
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Proveedores"
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
         Index           =   1
         Left            =   2640
         TabIndex        =   31
         Top             =   960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2580
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   3060
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   4551
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
         Left            =   3660
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
      Begin VB.Label Label3 
         Caption         =   "Factura"
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
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   2280
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3360
         Picture         =   "frmConsolidadoFras.frx":000C
         ToolTipText     =   "Quitar al Debe"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   3720
         Picture         =   "frmConsolidadoFras.frx":0156
         ToolTipText     =   "Puntear al Debe"
         Top             =   2760
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
         Top             =   2760
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
      TabIndex        =   5
      Top             =   5940
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
      TabIndex        =   4
      Top             =   5940
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
      TabIndex        =   6
      Top             =   5910
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
      Top             =   6000
      Width           =   6270
   End
End
Attribute VB_Name = "frmConsolidadoFras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 413

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

Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmGas As frmBasico
Attribute frmGas.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim I As Integer
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

Private Function MontaSQL() As Boolean
'Dim Sql As String
'Dim RC2 As String
'Dim RC As String
'Dim i As Integer

    MontaSQL = False
    
    
    cadselect = "nifdatos = " & DBSet(txtNIF(0).Text, "T")
    If Not PonerDesdeHasta("#@#", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    

    MontaSQL = True
End Function


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
    
    If Not MontaSQL Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    B = CargarTemporales
    If B Then
        If Not HayRegParaInforme("tmpfaclin", "codusu = " & vUsu.Codigo) Then B = False
    End If
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
    Me.Caption = "Informe de facturas CONSOLIDADO"

    
    For I = 0 To 1
        Me.ImgFec(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next I
     
    For I = 0 To 0
        Me.imgCuentas(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    Me.txtNIF(0).Text = ""
    
     
    CargarListViewEmpresas 1

    
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
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "Código", 600
    ListView1(Index).ColumnHeaders.Add , , "Empresa", 3200
    


    Set Rs = New ADODB.Recordset

    Prohibidas = DevuelveProhibidas
    
    ListView1(Index).ListItems.Clear
    Aux = "Select * from usuarios.empresasariconta where tesor>0 order by codempre"
    
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        
        'En cosolidados DEJO seleccionar cualquier empresa
        'Aux = "ariconta" & Rs!codempre & ".parametros"
        'Aux = DevuelveDesdeBD("esmultiseccion", Aux, "1", "1")
        'If Aux = "0" Then
        '    Aux = "N"
        'Else
            Aux = "|" & Rs!codempre & "|"
            If InStr(1, Prohibidas, Aux) = 0 Then Aux = ""
        'End If
        If Aux = "" Then
            Set IT = ListView1(Index).ListItems.Add
            IT.Key = "C" & Rs!codempre
            If vEmpresa.codempre = Rs!codempre Then IT.Checked = True
            IT.Text = Rs!codempre
            IT.SubItems(1) = Rs!nomempre
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
Dim I As Integer


    On Error GoTo EDevuelveProhibidas
    
    DevuelveProhibidas = ""

    Set miRsAux = New ADODB.Recordset

    I = vUsu.Codigo Mod 100
    miRsAux.Open "Select * from usuarios.usuarioempresasariconta WHERE codusu =" & I, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
    Sql = CadenaSeleccion
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
        ' empresas de usuarios
        Case 0
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = False
            Next I
        Case 1
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = True
            Next I
    
        ' tipos de forma de pago
        Case 2
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = False
            Next I
        Case 3
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = True
            Next I
    
    End Select
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub imgCuentas_Click(Index As Integer)
    Sql = ""
    AbiertoOtroFormEnListado = True
    Set frmCCtas = New frmColCtas
    frmCCtas.DatosADevolverBusqueda = True
    frmCCtas.Show vbModal
    Set frmCCtas = Nothing
    If Sql <> "" Then
    '    Me.txtCuentas(Index).Text = RecuperaValor(Sql, 1)
    '    Me.txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
    '    If txtCuentas(Index).Text <> "" Then
            Me.txtNIF(0).Text = ""
            Me.txtNIF(0).Text = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", RecuperaValor(Sql, 1), "T")
    '    End If

    Else
      '  QuitarPulsacionMas Me.txtCuentas(Index)
    End If
    
    'PonFoco Me.txtCuentas(Index)
    AbiertoOtroFormEnListado = False
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
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
        imgCuentas_Click Indice
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
        
    
    'indRPT = "0901-00"
    '
    '
   '
   ' If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    nomDocu = "ConsolidadoFra.rpt"
    cadNomRPT = nomDocu
    
    cadFormula = "{tmpfaclin.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 50
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub


Private Function CargarTemporales() As Boolean
Dim Sql As String
Dim RC As String
Dim RC2 As String
Dim I As Integer
Dim B As Boolean
Dim LaCuenta As Boolean

    CargarTemporales = False
    
    Label9.Caption = "Preparando tablas"
    Label9.Refresh
    'tmpfaclin(codusu,codigo,numserie,nomserie,Numfac,Fecha,cta,Cliente,NIF,Imponible,IVA,ImpIVA,Total,retencion,numfactura)
    Sql = "Delete from tmpfaclin where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "Delete from tmpcuentas where codusu =" & vUsu.Codigo
    Conn.Execute Sql
                
                
    Sql = ""
    LaCuenta = False
    
    RC = "INSERT INTO tmpfaclin(codusu,codigo,numserie,nomserie,numfactura,Fecha,cta,Cliente,NIF,Imponible,IVA,ImpIVA,retencion,Total,Numfac,tipoformapago) "
        
    '------------------------------------------
    'UNO SOLO
    For I = 1 To ListView1(1).ListItems.Count
        
        
        
        
        If ListView1(1).ListItems(I).Checked Then
            Screen.MousePointer = vbHourglass
            If Cancelado Then Exit For
            Label9.Caption = "Obteniendo fra. cli.: " & ListView1(1).ListItems(I).Text
            Label9.Refresh
            
            If Not LaCuenta Then
                RC2 = DevuelveDesdeBD("nommacta", "ariconta" & ListView1(1).ListItems(I).Tag & ".cuentas", "nifdatos", Me.txtNIF(0).Text, "T")
                If RC2 <> "" Then
                    RC2 = vUsu.Codigo & ",'n'," & DBSet(RC2, "T") & "," & DBSet(txtNIF(0).Text, "T")
                    RC2 = "INSERT INTO tmpcuentas(codusu,codmacta,nommacta,nifdatos) VALUES (" & RC2 & ")"
                    Conn.Execute RC2
                    LaCuenta = True
                End If
            End If
                
            'Clientes
            If Me.Check1(0).Value = 1 Then
                RC2 = cadselect
                RC2 = Replace(RC2, "#@#", "factcli.fecfactu")
                
                Sql = "SELECT " & vUsu.Codigo & " codusu," & Me.ListView1(1).ListItems(I).Tag & " empresa,"
                Sql = Sql & " numserie,nomregis, numfactu,fecfactu,codmacta,nommacta,nifdatos,totbases,0,"
                Sql = Sql & " totivas+coalesce(totrecargo,0),coalesce(trefaccl,0),totfaccl, 0 proveed ,nomforpa from "
                Sql = Sql & " ariconta" & ListView1(1).ListItems(I).Tag & ".factcli,ariconta" & ListView1(1).ListItems(I).Tag & ".contadores"
                Sql = Sql & " ,ariconta" & ListView1(1).ListItems(I).Tag & ".formapago"
                Sql = Sql & " WHERE factcli.numserie=contadores.tiporegi AND factcli.codforpa=formapago.codforpa  AND " & RC2
                Sql = RC & Sql
                If Not Ejecuta(Sql) Then Exit Function
                
                
            End If
            'Proveedores
            Label9.Caption = "Obteniendo fra pro.: " & ListView1(1).ListItems(I).Text
            Label9.Refresh
            If Me.Check1(1).Value = 1 Then
                RC2 = cadselect
                RC2 = Replace(RC2, "#@#", "factpro.fecharec")
                Sql = "SELECT " & vUsu.Codigo & " codusu," & Me.ListView1(1).ListItems(I).Tag & " empresa,"
                Sql = Sql & " numserie,nomregis, numfactu,fecharec,codmacta,nommacta,nifdatos,totbases,0,"
                Sql = Sql & " totivas+coalesce(totrecargo,0),coalesce(trefacpr,0),totfacpr, 1 proveed,nomforpa FROM "
                Sql = Sql & " ariconta" & ListView1(1).ListItems(I).Tag & ".factpro,ariconta" & ListView1(1).ListItems(I).Tag & ".contadores"
                Sql = Sql & " ,ariconta" & ListView1(1).ListItems(I).Tag & ".formapago"
                Sql = Sql & " WHERE  factpro.numserie=contadores.tiporegi AND factpro.codforpa=formapago.codforpa AND " & RC2
                Sql = RC & Sql
                If Not Ejecuta(Sql) Then Exit Function
                
                
            End If
            
            
            
        End If
    Next I
            
    Label9.Caption = "Obteniendo valores empresa "
    Label9.Refresh

    Set miRsAux = New ADODB.Recordset
    Sql = "Select codigo from tmpfaclin where codusu = " & vUsu.Codigo & " GROUP BY 1"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic
    While Not miRsAux.EOF
            Sql = miRsAux.Fields(0)
            Sql = DevuelveDesdeBD("nomempre", "ariconta" & Sql & ".empresa", "1", "1")
            
            Sql = "UPDATE tmpfaclin set tipoopera =" & DBSet(Sql, "T") & " WHERE codusu = " & vUsu.Codigo & " AND codigo =" & miRsAux!Codigo
            Conn.Execute Sql
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Label9.Caption = ""
    Label9.Refresh
            
    CargarTemporales = True
    Screen.MousePointer = vbDefault
End Function



Private Function DameEmpresa(ByVal S As String) As String
    DameEmpresa = "NO ENCONTRADA"
    For I = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(I).Tag = S Then
            DameEmpresa = DevNombreSQL(ListView1(1).ListItems(I).Text)
            Exit For
        End If
    Next I
  
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
    
    If Me.txtNIF(0).Text = "" Then
        MsgBox "Introduzca la Cuenta para obtener el NIF", vbExclamation
        Exit Function
    End If
    
    Sql = ""
    For I = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(I).Checked Then
            Sql = "O"
            Exit For
        End If
    Next I
    If Sql = "" Then
        MsgBox "Seleccione al menos una empresa", vbExclamation
        Exit Function
    End If

    If Me.Check1(0).Value = 0 And Me.Check1(1).Value = 0 Then
        MsgBox "Seleccione clientes y/o proveedores.", vbExclamation
        Exit Function
    End If
    DatosOK = True


End Function


Private Sub txtNIF_LostFocus(Index As Integer)
    If txtNIF(0).Text <> "" Then
        Sql = "nommacta"
        RC = DevuelveDesdeBD("codmacta", "cuentas", "nifdatos", txtNIF(0).Text, "T", Sql)
        If RC = "" Then
           ' txtNCuentas(0).Text = "NIF no pertenece a ninguan cuenta"
        Else
           ' txtNCuentas(0).Text = Sql
           ' Me.txtCuentas(0).Text = RC
            PonleFoco Me.cmdAccion
        End If
    Else
        
    End If
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

