VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESInfSituacionNIF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11865
   Icon            =   "frmTESInfSituacionNIF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
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
      TabIndex        =   18
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txtCuentas 
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
         Left            =   1260
         TabIndex        =   4
         Tag             =   "imgConcepto"
         Top             =   690
         Width           =   1485
      End
      Begin VB.TextBox txtNCuentas 
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
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   690
         Width           =   3735
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
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   2100
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
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1680
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
         Left            =   930
         Top             =   720
         Width           =   255
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
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   23
         Top             =   420
         Width           =   1020
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   960
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   960
         Top             =   1680
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
         Top             =   2100
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
         Top             =   1710
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Vencimiento"
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
         Left            =   270
         TabIndex        =   20
         Top             =   1410
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
      TabIndex        =   7
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   10
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   9
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
         TabIndex        =   8
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
      TabIndex        =   6
      Top             =   0
      Width           =   4305
      Begin MSComctlLib.ListView ListView1 
         Height          =   2100
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   3540
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3704
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2130
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3757
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
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3660
         TabIndex        =   33
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
         Index           =   2
         Left            =   3420
         Picture         =   "frmTESInfSituacionNIF.frx":000C
         ToolTipText     =   "Quitar al Debe"
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   3780
         Picture         =   "frmTESInfSituacionNIF.frx":0156
         ToolTipText     =   "Puntear al Debe"
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Pago"
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
         Left            =   210
         TabIndex        =   31
         Top             =   660
         Width           =   1920
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3360
         Picture         =   "frmTESInfSituacionNIF.frx":02A0
         ToolTipText     =   "Quitar al Debe"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   3720
         Picture         =   "frmTESInfSituacionNIF.frx":03EA
         ToolTipText     =   "Puntear al Debe"
         Top             =   3210
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
         TabIndex        =   29
         Top             =   3240
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
      Left            =   10200
      TabIndex        =   3
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
      Left            =   8640
      TabIndex        =   2
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
      TabIndex        =   5
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
      TabIndex        =   32
      Top             =   6000
      Width           =   6270
   End
End
Attribute VB_Name = "frmTESInfSituacionNIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 901

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
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim PrimeraVez As Boolean
Dim Cancelado As Boolean

Dim vNIF As String


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
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("cobros.FecFactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
            
    MontaSQL = True
End Function


Private Sub cmdAccion_Click(Index As Integer)

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    tabla = " "
    
    If Not MontaSQL Then Exit Sub
    
    If Not CargarTemporales Then Exit Sub
    
    If Not HayRegParaInforme("tmptesoreriacomun", "codusu = " & vUsu.Codigo) Then Exit Sub
    
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
        PonFoco txtCuentas(0)
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
    Me.Caption = "Informe de Situación por NIF"

    
    For i = 0 To 1
        Me.ImgFec(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
     
    For i = 0 To 0
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    vNIF = ""
    
     
    CargarListViewEmpresas 1
    CargarListViewTipoFPago 0
    
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

    ListView1(Index).ColumnHeaders.Add , , "Empresa", 3800
    


    Set Rs = New ADODB.Recordset

    Prohibidas = DevuelveProhibidas
    
    ListView1(Index).ListItems.Clear
    Aux = "Select * from Usuarios.empresasariconta where tesor>0"
    
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        Aux = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, Aux) = 0 Then
            Set IT = ListView1(Index).ListItems.Add
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

Private Sub CargarListViewTipoFPago(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , "Descripción", 3200
    ListView1(Index).ColumnHeaders.Add , , "Código", 0
    
    Sql = "SELECT descformapago, tipoformapago  "
    Sql = Sql & " FROM tipofpago "
    Sql = Sql & " order by 2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        
        ItmX.Text = Rs.Fields(0).Value
        ItmX.SubItems(1) = Rs.Fields(1).Value
        
        ItmX.Checked = True
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tipos de Forma de Pago.", Err.Description
    End If
End Sub





Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
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
        ' empresas de usuarios
        Case 0
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = False
            Next i
        Case 1
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = True
            Next i
    
        ' tipos de forma de pago
        Case 2
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = False
            Next i
        Case 3
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = True
            Next i
    
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
        Me.txtCuentas(Index).Text = RecuperaValor(Sql, 1)
        Me.txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
        If txtCuentas(Index).Text <> "" Then
            vNIF = ""
            vNIF = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", txtCuentas(Index), "T")
        End If

    Else
        QuitarPulsacionMas Me.txtCuentas(Index)
    End If
    
    PonFoco Me.txtCuentas(Index)
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


Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub

Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtCuentas(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
    If Index = 1 And KeyAscii = 13 Then
        PonFoco txtFecha(0)
    End If
End Sub

Private Sub txtCuentas_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    
    
    If txtCuentas(Index).Text = "" Then
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCuentas(Index).Text) Then
        If InStr(1, txtCuentas(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCuentas(Index).Text, vbExclamation
        txtCuentas(Index).Text = ""
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
            Else
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = Sql
                If B = 1 Then
                    txtNCuentas(Index).Tag = ""
                Else
                    txtNCuentas(Index).Tag = Sql
                End If
                Hasta = -1
                If Index = 6 Then
                    Hasta = 7
                Else
                    If Index = 0 Then
                        Hasta = 1
                    Else
                        If Index = 5 Then
                            Hasta = 4
                        Else
                            If Index = 23 Then Hasta = 24
                        End If
                    End If
                    
                End If
                vNIF = ""
                vNIF = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", txtCuentas(Index), "T")
            End If
    
    
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
    
    Sql = "SELECT `tmptesoreriacomun`.`texto1` Nif , `tmptesoreriacomun`.`texto2` Conta, `tmptesoreriacomun`.`opcion` BD, `tmptesoreriacomun`.`texto5` Nombre, `tmptesoreriacomun`.`texto3` NroFra, `tmptesoreriacomun`.`fecha1` FecFra, `tmptesoreriacomun`.`fecha2` FecVto, `tmptesoreriacomun`.`importe1` Gasto, `tmptesoreriacomun`.`importe2` Recibo"
    Sql = Sql & " FROM   `tmptesoreriacomun` `tmptesoreriacomun`"
    Sql = Sql & " WHERE `tmptesoreriacomun`.codusu = " & vUsu.Codigo
    Sql = Sql & " ORDER BY `tmptesoreriacomun`.`texto1`, `tmptesoreriacomun`.`texto2`, `tmptesoreriacomun`.`opcion`, `tmptesoreriacomun`.`fecha1`"
    
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0901-00"
    
        
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "GastosFijos.rpt"
    
    cadFormula = "{tmptesoreriacomun.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 50
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub


Private Function CargarTemporales() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim i As Integer
Dim B As Boolean

    CargarTemporales = False
    
    Label9.Caption = "Preparando tablas"
    Label9.Refresh
    Sql = "Delete from tmp347 where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    Sql = "Delete from tmptesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    'tmpfaclin  ... sera para cuando es mas de uno
    Sql = "Delete from tmpfaclin where codusu =" & vUsu.Codigo
    Conn.Execute Sql
                
    Sql = ""
    Screen.MousePointer = vbHourglass
    
    '------------------------------------------
    'UNO SOLO
    For i = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(i).Checked Then
            If Cancelado Then Exit For
            Label9.Caption = "Obteniendo tabla1: " & ListView1(1).ListItems(i).Text
            Label9.Refresh
            DoEvents
        
            Sql = "INSERT INTO tmp347 (codusu, cliprov, cta, nif) "
            
            Sql = Sql & " select " & vUsu.Codigo & "," & Mid(ListView1(1).ListItems(i).Key, 2) & ", codmacta, nifdatos from "
            Sql = Sql & "( "
        
            Sql = Sql & "select cobros.codmacta,nifclien nifdatos from ariconta" & ListView1(1).ListItems(i).Tag & ".cobros where not nifclien is null  "
            If vNIF <> "" Then Sql = Sql & " and cobros.nifclien = " & DBSet(vNIF, "T")
            Sql = Sql & " group by  codmacta,nifclien"
            Sql = Sql & " union "
            Sql = Sql & "select pagos.codmacta,nifprove nifdatos from ariconta" & ListView1(1).ListItems(i).Tag & ".pagos where not nifprove is null"
            If vNIF <> "" Then Sql = Sql & " and pagos.nifprove = " & DBSet(vNIF, "T")
            Sql = Sql & " group by  codmacta,nifprove"
            
            Sql = Sql & ") aaaaa "
            Sql = Sql & " group by 3, 4 "
            
            If Not Ejecuta(Sql) Then Exit Function
        End If
    Next i
            
    If Sql <> "" Then
        If GeneraCobrosPagosNIF Then
        
            Sql = ""
            For i = 1 To Me.ListView1(1).ListItems.Count
                If Me.ListView1(1).ListItems(i).Checked Then Sql = Sql & "1"
            Next
            If Len(Sql) > 1 Then
                Sql = "0"
            Else
                Sql = "1"
            End If
            Sql = "SoloUnaEmpresa= " & Sql & "|"

        
        End If
    End If
            
    Label9.Caption = ""
    Label9.Refresh
            
    CargarTemporales = True
    Screen.MousePointer = vbDefault
End Function


Public Function GeneraCobrosPagosNIF() As Boolean
Dim cad As String
Dim L As Long
Dim Empre As String
Dim Importe  As Currency
Dim Rs As ADODB.Recordset
Dim QueTipoPago As String


    On Error GoTo eGeneraCobrosPagosNIF

    'Guardaremos en la variable QueTipoPago que tipos de pago ha seleccionado
    'Si selecciona todos los tipos de pago NO pondremos el IN en el select
    QueTipoPago = ""
    cad = "" 'para saber si ha selccionado todos
    For L = 1 To Me.ListView1(0).ListItems.Count
        If ListView1(0).ListItems(L).Checked Then
            QueTipoPago = QueTipoPago & ", " & Me.ListView1(0).ListItems(L).Tag
        Else
            cad = "NO" 'No estan todos seleccionados
        End If
    Next
    If cad = "" Then
        'Estan todos. No tiene sentido hacer el Select in
        QueTipoPago = ""
    Else
        QueTipoPago = Mid(QueTipoPago, 2)
    End If
    
    
    
'En la tabla  INSERT INTO tmp347 (codusu, cliprov, cta, nif) VALUES ((
' Tendremos codccos: la empresa
'                  : cta, cada uno de los valores
'INSERT INTO ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,
'texto5, texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1,
'observa2, opcion) VALUES
    GeneraCobrosPagosNIF = False
    L = 1
    Sql = "Select * from tmp347 where codusu =" & vUsu.Codigo & " ORDER BY cliprov,nif"
    Set Rs = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        If Cancelado Then
            Rs.Close
            Exit Function
        End If
        'Los labels
        Label9.Caption = "Nif: " & Rs!NIF & " - " & Rs!Cta
        Label9.Refresh
        
        'SQL insert
        Sql = "INSERT INTO tmptesoreriacomun (codusu,texto1, codigo,texto2,  texto3,texto4, texto5,fecha1,fecha2,"   'texto5, texto6,
        Sql = Sql & " importe1, importe2,opcion"
        Sql = Sql & ") VALUES ("
        'NIF      Nombre
        Sql = Sql & vUsu.Codigo & ",'" & Rs!NIF & "',"
        
        
        '-------
        Empre = DameEmpresa(CStr(Rs!cliprov))
        
        'COBROS
        cad = "Select fecfactu,numserie,numfactu, numorden,impvenci,impcobro,gastos,fecvenci,nomclien nommacta from ariconta" & Rs!cliprov & ".cobros as c1 "
        If QueTipoPago <> "" Then cad = cad & ", ariconta" & Rs!cliprov & ".formapago as sforpa"
        cad = cad & " where c1.nifclien='" & Rs!NIF & "'"
        If QueTipoPago <> "" Then cad = cad & " AND c1.codforpa=sforpa.codforpa AND sforpa.tipforpa in (" & QueTipoPago & ")"
        'Fechas
        If txtFecha(0).Text <> "" Then cad = cad & " AND fecvenci >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then cad = cad & " AND fecvenci <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Los label
            If Cancelado Then
                miRsAux.Close
                Exit Function
            End If
            
            'Insetamos codigo,  texto3
            '                    empresa
            cad = L & ",'" & Empre & "','"
            cad = cad & miRsAux!NUmSerie & "/" & Format(miRsAux!NumFactu, "0000000000") & " : " & miRsAux!numorden & "','"
            cad = cad & Rs!Cta & "','"
            cad = cad & DevNombreSQL(miRsAux!Nommacta) & "','"
            'texto4: fecha
            cad = cad & Format(miRsAux!FecFactu, FormatoFecha) & "','"
            cad = cad & Format(miRsAux!FecVenci, FormatoFecha) & "',"
            
            
            'En importe1 estara el importe del cobro. En el 2 tb
'            Importe = DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
'            Importe = Importe + miRsAux!impvenci
'             Cad = Cad & TransformaComasPuntos(CStr(Importe)) & "," & TransformaComasPuntos(CStr(Importe))


            Importe = DBLet(miRsAux!Gastos, "N")
            cad = cad & TransformaComasPuntos(CStr(Importe))
            Importe = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N")
            cad = cad & "," & TransformaComasPuntos(CStr(Importe))
           
            
            
            'un cero para importe 2  y un cero para la opcion
            cad = cad & ",0)"
            
            'Ejecutamos
            cad = Sql & cad
            Ejecuta cad
            
            L = L + 1
            miRsAux.MoveNext
            DoEvents
        Wend
        miRsAux.Close
        
        'PAGOS
        cad = "Select numfactu,numorden,fecfactu,imppagad,fecefect,impefect,nomprove from ariconta" & Rs!cliprov & ".pagos "
        If QueTipoPago <> "" Then cad = cad & ", ariconta" & Rs!cliprov & ".formapago as sforpa"
        cad = cad & " where nifprove='" & Rs!NIF & "'"
        If QueTipoPago <> "" Then cad = cad & " AND pagos.codforpa=sforpa.codforpa AND sforpa.tipforpa in (" & QueTipoPago & ")"
        
        
        'Fechas
        If txtFecha(0).Text <> "" Then cad = cad & " AND fecefect >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then cad = cad & " AND fecefect <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Los label
            If Cancelado Then
                miRsAux.Close
                Exit Function
            End If
            
            'Insetamos codigo,  texto3,t5
            '                    empresa
            cad = L & ",'" & Empre & "','"
            cad = cad & DevNombreSQL(miRsAux!NumFactu) & " : " & miRsAux!numorden & "','"
            cad = cad & Rs!Cta & "','"
            cad = cad & DevNombreSQL(miRsAux!Nommacta) & "','"
            ' fecha1 y 2
            cad = cad & Format(miRsAux!FecFactu, FormatoFecha) & "','"
            cad = cad & Format(miRsAux!fecefect, FormatoFecha) & "',"
            
            
            'En importe1 estara el importe del cobro
            Importe = DBLet(miRsAux!imppagad, "N")

            Importe = miRsAux!ImpEfect - Importe
            cad = cad & TransformaComasPuntos(CStr(0)) & "," & TransformaComasPuntos(CStr(-1 * Importe))
            
            cad = cad & ",1)" '1: pago
            
            
            
            
            'Ejecutamos
            cad = Sql & cad
            Ejecuta cad
            
            L = L + 1
            miRsAux.MoveNext
            
            DoEvents
        Wend
        miRsAux.Close
        
        
        'SIGUIENTE CUENTA
        Rs.MoveNext
    Wend
    Rs.Close
    
    cad = "DELETE FROM tmptesoreriacomun where codusu = " & vUsu.Codigo & " AND importe1+importe2=0"
    Conn.Execute cad
    
    cad = "select count(*) from tmptesoreriacomun where codusu = " & vUsu.Codigo
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not Rs.EOF Then
        L = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    
    Set Rs = Nothing
    Set miRsAux = Nothing
    
    If L = 0 Then
        'ERROR. MO HAY DATOS
        MsgBox "Sin datos.", vbExclamation
    Else
        GeneraCobrosPagosNIF = True
    End If
    Exit Function
    
eGeneraCobrosPagosNIF:
    MuestraError Err.Number, "Genera Cobros/Pagos NIF", Err.Description
End Function

Private Function DameEmpresa(ByVal S As String) As String
    DameEmpresa = "NO ENCONTRADA"
    For i = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(i).Tag = S Then
            DameEmpresa = DevNombreSQL(ListView1(1).ListItems(i).Text)
            Exit For
        End If
    Next i
  
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
    
    If vNIF = "" Then
        MsgBox "Introduzca la Cuenta para obtener el NIF", vbExclamation
        Exit Function
    End If
    
    Sql = ""
    For i = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(i).Checked Then
            Sql = "O"
            Exit For
        End If
    Next i
    If Sql = "" Then
        MsgBox "Seleccione al menos una empresa", vbExclamation
        Exit Function
    End If
    
    'Tipos de pago
    Sql = ""
    For i = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(i).Checked Then
            Sql = "O"
            Exit For
        End If
    Next i
    If Sql = "" Then
        MsgBox "Seleccione al menos un tipo de pago", vbExclamation
        Exit Function
    End If
    
    DatosOK = True


End Function


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

