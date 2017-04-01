VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInfRatios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11745
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
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6915
      Begin TabDlg.SSTab SSTab1 
         Height          =   3765
         Left            =   150
         TabIndex        =   16
         Top             =   630
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6641
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Ratios"
         TabPicture(0)   =   "frmInfRatios.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Image2(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label6(14)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text3(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkRatio(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkRatio(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkRatio(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Gráficas"
         TabPicture(1)   =   "frmInfRatios.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "List1"
         Tab(1).Control(1)=   "cboMes"
         Tab(1).Control(2)=   "Label3(1)"
         Tab(1).Control(3)=   "Label3(0)"
         Tab(1).ControlCount=   4
         Begin VB.CheckBox chkRatio 
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   510
            TabIndex        =   22
            Top             =   840
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.CheckBox chkRatio 
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   510
            TabIndex        =   21
            Top             =   1380
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.CheckBox chkRatio 
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   510
            TabIndex        =   20
            Top             =   1890
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            Left            =   -74790
            Style           =   1  'Checkbox
            TabIndex        =   19
            Top             =   840
            Width           =   1815
         End
         Begin VB.ComboBox cboMes 
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
            Left            =   -71670
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox Text3 
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
            Left            =   2250
            TabIndex        =   17
            Top             =   2790
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Comparativa mes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   1
            Left            =   -71640
            TabIndex        =   25
            Top             =   510
            Width           =   1890
         End
         Begin VB.Label Label3 
            Caption         =   "Años"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   -74820
            TabIndex        =   24
            Top             =   480
            Width           =   960
         End
         Begin VB.Label Label6 
            Caption         =   "Hasta Fecha"
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
            Index           =   14
            Left            =   540
            TabIndex        =   23
            Top             =   2790
            Width           =   1245
         End
         Begin VB.Image Image2 
            Height          =   240
            Index           =   0
            Left            =   1890
            Picture         =   "frmInfRatios.frx":0038
            Top             =   2805
            Width           =   240
         End
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
      Height          =   7335
      Left            =   7110
      TabIndex        =   12
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox chkGraf1 
         Caption         =   "Resumen"
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
         Left            =   420
         TabIndex        =   26
         Top             =   930
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3810
         TabIndex        =   28
         Top             =   240
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
      TabIndex        =   2
      Top             =   7500
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
      TabIndex        =   0
      Top             =   7500
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
      TabIndex        =   1
      Top             =   7440
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
      TabIndex        =   3
      Top             =   4680
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
         TabIndex        =   15
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   14
         Top             =   1680
         Width           =   255
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox txtDescrip 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   2
      Left            =   270
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Text            =   "frmInfRatios.frx":00C3
      Top             =   4050
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox txtDescrip 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Text            =   "frmInfRatios.frx":00C9
      Top             =   2130
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.TextBox txtDescrip 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Text            =   "frmInfRatios.frx":00CF
      Top             =   210
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.Label lblInd 
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
      Left            =   2190
      TabIndex        =   27
      Top             =   7440
      Width           =   4785
   End
End
Attribute VB_Name = "frmInfRatios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 312

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

Public NumAsien As String
Public NumDiari As String
Public FechaEnt As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim Comparativo As Boolean


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
Dim tabla As String

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Me.SSTab1.Tab = 0 Then
        'Ratios
        B = HacerRatios
    Else
        'Graficos
        B = HacerGraficas
    End If
    
    Select Case SSTab1.Tab
        Case 0
            cadselect = "tmptesoreriacomun.codusu=" & vUsu.Codigo
            tabla = "tmptesoreriacomun"
            
        Case 1
            If chkGraf1(0).Value = 0 Then
                cadselect = "tmpbalancesumas.codusu=" & vUsu.Codigo
                tabla = "tmpbalancesumas"
            Else
                cadselect = "tmpsaldoscc.codusu=" & vUsu.Codigo
                tabla = "tmpsaldoscc"
            End If
    End Select
    
    Me.lblInd.Caption = ""
    
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
    
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
    
    'Otras opciones
    Me.Caption = "Informe de Ratios y Gráficas"

    SSTab1.Tab = 0
    Me.chkGraf1(0).Enabled = False
    
    
    optTipoSal(1).Enabled = False
    txtTipoSalida(1).Enabled = False
    PushButton2(0).Enabled = False

    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
   
   Sql = "01/" & Month(Now) & "/" & Year(Now)
   Sql = DateAdd("d", -1, CDate(Sql))
   Text3(0).Text = Sql
   CargaDatosRatios
   CargaDatosGraficas
    
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



Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    Sql = "Select  hcabapu.numdiari Diario, hcabapu.numasien Asiento, hcabapu.fechaent Fecha, hlinapu.linliapu Linea, hlinapu.codmacta Cuenta, nommacta Descripcion, numdocum Documento, ampconce Ampliacion, timporteD Debe, timporteH Haber"
    Sql = Sql & " FROM (hcabapu inner join hlinapu on hcabapu.numdiari = hlinapu.numdiari and hcabapu.numasien = hlinapu.numasien and hcabapu.fechaent = hlinapu.fechaent)"
    Sql = Sql & " inner join cuentas on hlinapu.codmacta = cuentas.codmacta "
    
    If cadselect <> "" Then Sql = Sql & " WHERE " & cadselect
    
    Sql = Sql & " ORDER BY 1,2,3,4"
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
Dim Sql As String
Dim Aux As String
    
    vMostrarTree = False
    conSubRPT = False
        
    If Me.SSTab1.Tab = 0 Then
        indRPT = "0312-01"
        
        cadFormula = "{tmptesoreriacomun.codusu}=" & vUsu.Codigo
    
        Sql = " hasta " & Text3(0).Text
        cadParam = cadParam & "pDesde=""" & " hasta " & Text3(0).Text & """|"
        numParam = numParam + 1
    Else
        If chkGraf1(0).Value = 0 Then
            indRPT = "0312-02"
        
            cadFormula = "{tmpbalancesumas.codusu}=" & vUsu.Codigo
        
            Sql = ""
            Aux = ""
            For NumRegElim = List1.ListCount - 1 To 0 Step -1
                If List1.Selected(NumRegElim) Then
                    Sql = Sql & "1"
                    If Aux = "" Then
                        'Primer ejercicio
                        Aux = "TextoEjer1=""" & List1.List(NumRegElim) & """|"
                        cadParam = cadParam & Aux
                        numParam = numParam + 1
                        
                    Else
                        'Segundo
                        Aux = Aux & "TextoEjer2=""" & List1.List(NumRegElim) & """|"
                        cadParam = cadParam & "TextoEjer2= """ & List1.List(NumRegElim) & """|"
                        numParam = numParam + 1
                    End If
                End If
            Next
            i = 0
            If Len(Sql) > 1 Then i = 1
            
            Sql = "Comparativo=" & i & "|" & Aux
            cadParam = cadParam & "Comparativo=" & i & "|"
            numParam = numParam + 1
        
        
        Else 'informe de graficas resumido
            indRPT = "0312-03"
            
            cadFormula = "{tmpsaldoscc.codusu}=" & vUsu.Codigo
            
            Sql = ""
            Aux = ""
            For NumRegElim = 0 To List1.ListCount - 1
                If List1.Selected(NumRegElim) Then
                    If Aux = "" Then Aux = "UltAno= " & Mid(List1.List(NumRegElim), 1, 4) & "|"
                End If
            Next
            Sql = Aux
            cadParam = cadParam & Aux
            numParam = numParam + 1
            
            NumRegElim = 1
            If cboMes.ListIndex > 0 Then
                'ha seleccionado mes
                Sql = Sql & "Desde=""Hasta " & cboMes.Text & """|"
                NumRegElim = 2
                
                cadParam = cadParam & "Desde=""Hasta " & cboMes.Text & """|"
                numParam = numParam + 1
            End If
            
        End If
    End If
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "ratios.rpt"

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 56
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
            
    MontaSQL = True

End Function

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    Select Case Me.SSTab1.Tab
        Case 0 ' hacer ratios
            If Text3(0).Text = "" Then
                MsgBox "Ponga la fecha", vbExclamation
                Exit Function
            End If
            
            If FechaCorrecta2(CDate(Text3(0).Text)) > 2 Then
                MsgBox "Fecha mayor que ejercicios abiertos.", vbExclamation
                Exit Function
            End If
            
            If chkRatio(0).Value = 0 And chkRatio(1).Value = 0 And chkRatio(2).Value = 0 Then
                MsgBox "Debe de seleccionar un tipo de ratio. Revise.", vbExclamation
                Exit Function
            End If
            
        Case 1 ' caso de graficas
            Sql = ""
            For i = 0 To Me.List1.ListCount - 1
                If List1.Selected(i) Then Sql = Sql & "1"
            Next
            If Len(Sql) < 1 Then
                MsgBox "Seleccione un año", vbExclamation
                Exit Function
            End If
            Comparativo = False
            If Len(Sql) = 2 Then
                Comparativo = True
                If cboMes.ListIndex <= 0 Then
                    MsgBox "Seleccione el mes para el comparativo", vbExclamation
                    Exit Function
                End If
            
            End If
            If Me.chkGraf1(0).Value = 0 And Len(Sql) > 2 Then
                MsgBox "Seleccione un año(dos para el comparativo)", vbExclamation
                Exit Function
            End If
         
    End Select
    
    DatosOK = True


End Function


Private Sub CargaDatosRatios()

    'NO puede dar error

    'En balances, del 51 al 53 tiene que existir  CUANDO ESTEN TODOS sera hasta el 55
    Sql = "Select * from balances where numbalan>=51 and numbalan<=54 order by numbalan"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        If i < 2 Then
    
            i = miRsAux!NumBalan - 51
            
            Me.chkRatio(i).Caption = miRsAux!NomBalan
'            Me.txtDescrip(i).Text = miRsAux!Descripcion
            
        End If
        
        miRsAux.MoveNext
        
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    Me.cboMes.Clear
    Me.cboMes.AddItem " " 'todos
    For i = 1 To 12
        Me.cboMes.AddItem Format("23/" & i & "/2000", "mmmm")
    Next
    
End Sub


Private Sub CargaDatosGraficas()

    Sql = "select year(fechaent) anopsald from hlinapu group by 1 order by 1 desc"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    Sql = ""
    While Not miRsAux.EOF
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            'Año natural
            Sql = miRsAux!anopsald
        
        Else
            'Sera yyyy - yyyy  . Posiciones fijas.  4 prim año 1  desde la 8 año 2
            If Sql = "" Then
                
                    If miRsAux!anopsald > Year(vParam.fechaini) Then
                        List1.AddItem Format(miRsAux!anopsald, "0000") & " - " & Format(miRsAux!anopsald + 1, "0000")
                    End If

            End If
        
            Sql = Format(miRsAux!anopsald - 1, "0000") & " - " & Format(miRsAux!anopsald, "0000")
                    
        End If
        List1.AddItem Sql
        i = i + 1
        miRsAux.MoveNext
    Wend
    If i > 0 Then List1.Selected(0) = True
    miRsAux.Close
        
    
End Sub

Private Function HacerGraficas() As Boolean
Dim Veces As Byte  'para años partidos SON dos
Dim Ingresos As Currency
Dim Gastos As Currency
Dim Aux As Currency
Dim AnyoMes As Long

    
    HacerGraficas = False
    
    
    Me.lblInd.Caption = "Prepara datos"
    Me.lblInd.Refresh
    
    
    Sql = "DELETE FROM tmpgraficas where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    Conn.Execute "DELETE FROM tmpbalancesumas where codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM tmpsaldoscc where codusu = " & vUsu.Codigo
    
    'la de los informes tb
    For i = 0 To List1.ListCount - 1
        Veces = 1
        If Year(vParam.fechafin) <> Year(vParam.fechaini) Then Veces = 2
        If List1.Selected(i) Then
            Me.lblInd.Caption = List1.List(i)
            Me.lblInd.Refresh
            'Este esta selecionado
            While Veces <> 0
                Sql = "select year(fechaent) anopsald, month(fechaent) mespsald,codmacta,sum(coalesce(timported,0)) impmesde,sum(coalesce(timporteh,0)) impmesha"
                Sql = Sql & "  from hlinapu where"
                Sql = Sql & " (codmacta like '6%' or codmacta like '7%') "
            
                If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                    'AÑO NATURAL
                    Sql = Sql & " AND year(fechaent)= " & List1.List(i)
                    
                                    'Quiere hasta un mes
                    If Me.cboMes.ListIndex > 0 Then Sql = Sql & " AND month(fechaent)<= " & cboMes.ListIndex
                    
                    
                Else
                    'Años aprtidos
                    'Si veces=1 entonces el primer trozo de año partido
                    If Veces = 2 Then
                        'Segundo trozo
                        Sql = Sql & " AND year(fechaent)= " & Mid(List1.List(i), 8)
                        Sql = Sql & " AND month(fechaent)<=  " & Month(vParam.fechafin)
                        'Quiere hasta un mes
                        If Me.cboMes.ListIndex > 0 Then
                            If cboMes.ListIndex < Month(vParam.fechaini) Then Sql = Sql & " AND month(fechaent)<= " & cboMes.ListIndex
                        End If
                        
                    Else
                        Sql = Sql & " AND year(fechaent)= " & Mid(List1.List(i), 1, 4)
                        Sql = Sql & " AND month(fechaent) >=  " & Month(vParam.fechaini)
                        If Me.cboMes.ListIndex > 0 Then
                            If cboMes.ListIndex >= Month(vParam.fechaini) Then Sql = Sql & " AND month(fechaent)<= " & cboMes.ListIndex
                        End If
                        
                    End If
                End If
                Sql = Sql & " group by 1,2,3"
                Sql = Sql & " ORDER BY 1,2,3"
                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                AnyoMes = 0
                While Not miRsAux.EOF
                    
                    NumRegElim = miRsAux!anopsald * 100 + miRsAux!mespsald
                    If NumRegElim <> AnyoMes Then
                        'Nuevo ano,mes
                        If AnyoMes > 0 Then
                            'Ya tienen valor
                            InsertaEnTmpGraf AnyoMes, Ingresos, Gastos
                            
                        End If
                       
                        Ingresos = 0: Gastos = 0
                        AnyoMes = NumRegElim
                    End If
                
                    Aux = miRsAux!impmesde - miRsAux!impmesha
                    If Mid(miRsAux!codmacta, 1, 1) = "6" Then
                        Gastos = Gastos + Aux
                    Else
                        Ingresos = Ingresos - Aux 'va saldo
                    End If
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                'El ultimo
                If AnyoMes > 0 Then InsertaEnTmpGraf AnyoMes, Ingresos, Gastos
                
                Veces = Veces - 1
                
             
            Wend
            
        End If
    Next
    
    
    'Si no el el de RESUMEN
    If chkGraf1(0).Value = 0 Then
            'Ya tengo en tmpgrafiacs los valores de los meses
            'Insertare los 12 meses a ceros
            Me.lblInd.Caption = "Carga meses"
            Me.lblInd.Refresh
            Sql = ""
            If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                For Veces = 1 To 12
                    Sql = Sql & ", (" & vUsu.Codigo & ",'" & Format(Veces, "00") & "','" & Format("20/" & Veces & "/2000", "mmmm") & "',0,0,0,0,0,0,0,0)"
                Next Veces
                Sql = Mid(Sql, 2) 'quito la primera cma
                Sql = "INSERT INTO tmpbalancesumas (`codusu`,`cta`,`nomcta`,`aperturaD`,`aperturaH`,`acumAntD`,`acumAntH`,`acumPerD`," & _
                    "`acumPerH`,`TotalD`,`TotalH`) values " & Sql
                Conn.Execute Sql
            
            Else
                Sql = ""
                For Veces = Month(vParam.fechaini) To 12
                    Sql = Sql & ", (" & vUsu.Codigo & ",'00" & Format(Veces, "00") & "','" & Format("20/" & Veces & "/2000", "mmmm") & "',0,0,0,0,0,0,0,0)"
                Next Veces
                For Veces = 1 To Month(vParam.fechafin)
                    Sql = Sql & ", (" & vUsu.Codigo & ",'10" & Format(Veces, "00") & "','" & Format("20/" & Veces & "/2000", "mmmm") & "',0,0,0,0,0,0,0,0)"
                Next Veces
                Sql = Mid(Sql, 2) 'quito la primera cma
                Sql = "INSERT INTO tmpbalancesumas (`codusu`,`cta`,`nomcta`,`aperturaD`,`aperturaH`,`acumAntD`,`acumAntH`,`acumPerD`," & _
                    "`acumPerH`,`TotalD`,`TotalH`) values " & Sql
                Conn.Execute Sql
                    
            End If
            
            
            Sql = "select * from tmpgraficas where codusu = " & vUsu.Codigo & " order by anyo,mes"
            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            i = 0 'tendre el primer año
            While Not miRsAux.EOF
                Me.lblInd.Caption = miRsAux!Anyo & " " & miRsAux!Mes
                Me.lblInd.Refresh
                If i = 0 Then i = miRsAux!Anyo
                
                Sql = "UPDATE tmpbalancesumas SET "
                
                If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                    'años normales
                    If miRsAux!Anyo = i Then
                        'Año 1
                        'aperturaD aperturaH TotalD
                        Sql = Sql & "aperturaD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        Sql = Sql & ",aperturaH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        Sql = Sql & ",TotalD = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    Else
                        '`acumAntD`,`acumAntH` TotalH
                        Sql = Sql & "acumAntD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        Sql = Sql & ",acumAntH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        Sql = Sql & ",TotalH = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    End If
                    Sql = Sql & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & Format(miRsAux!Mes, "00") & "'"
                    
                Else
                    'años partidos
                    Veces = 0
                    If miRsAux!Anyo <> i Then
                        'Es año siguiente. Pero si el mes es anterior a mesini entonces todavia es ejercicio anterior
                        If miRsAux!Mes < Month(vParam.fechaini) Then
                            Veces = 0
                        Else
                            Veces = 1
                        End If
                    End If
                    
                    If Veces = 0 Then
                        'Año 1
                        'aperturaD aperturaH TotalD
                        Sql = Sql & "aperturaD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        Sql = Sql & ",aperturaH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        Sql = Sql & ",TotalD = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    Else
                        '`acumAntD`,`acumAntH` TotalH
                        Sql = Sql & "acumAntD = " & TransformaComasPuntos(CStr(miRsAux!Ingresos))
                        Sql = Sql & ",acumAntH = " & TransformaComasPuntos(CStr(miRsAux!Gastos))
                        Sql = Sql & ",TotalH = " & TransformaComasPuntos(CStr(miRsAux!beneficio))
                    End If
                    Sql = Sql & " WHERE codusu = " & vUsu.Codigo & " AND cta like '%" & Format(miRsAux!Mes, "00") & "'"
                
                End If
                Conn.Execute Sql
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            
            
            'Debemos borrar los datos de los meses
            If cboMes.ListIndex > 0 Then
                If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                    Sql = "DELETE FROM tmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta > '" & Format(cboMes.ListIndex, "00") & "'"
                    Conn.Execute Sql
                Else
                    If Month(vParam.fechaini) <= cboMes.ListIndex Then
        
                        
                        Sql = "DELETE FROM tmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta > '00" & Format(cboMes.ListIndex, "00") & "'"
                        Conn.Execute Sql
                    Else
                        'Quiere  hasta parte del años siguiente
                        Sql = "DELETE FROM tmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta > '10" & Format(cboMes.ListIndex, "00") & "'"
                        Conn.Execute Sql
                    
                    End If
                End If
            End If
            
            'Si NO es comparativo ponogo los importes a NULL
            If Not Comparativo Then
                Sql = "update tmpbalancesumas set `acumAntD`=NULL,`acumAntH`=NULL,`acumPerD`=NULL,`acumPerH`=NULL,`TotalH`=NULL"
                Sql = Sql & " where `codusu`=" & vUsu.Codigo
                Conn.Execute Sql
            End If
            
            'Renumeramos mes
            
            Sql = "Select * from tmpbalancesumas WHERE codusu = " & vUsu.Codigo & " ORDER BY cta"
            NumRegElim = 1
            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                Sql = "UPDATE tmpbalancesumas SET cta = '" & Format(NumRegElim, "00") & "' WHERE codusu = " & vUsu.Codigo & " AND cta ='" & miRsAux!Cta & "'"
                NumRegElim = NumRegElim + 1
                miRsAux.MoveNext
                Conn.Execute Sql
            Wend
            miRsAux.Close
            
    
    Else
        'RESUMEN
        
        
        Sql = "INSERT INTO  tmpsaldoscc( codusu,codccost,nomccost,ano,mes,impmesde,impmesha)"
        Sql = Sql & " SELECT codusu,'','',anyo,mes,ingresos,gastos from tmpgraficas where codusu = " & vUsu.Codigo
        Conn.Execute Sql
        
        
        'Debemos borrar los datos de los meses
        If cboMes.ListIndex > 0 Then
            If Year(vParam.fechafin) = Year(vParam.fechaini) Then
                Sql = "DELETE FROM tmpsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes > " & Format(cboMes.ListIndex, "00")
                Conn.Execute Sql
            Else
                If Month(vParam.fechaini) <= cboMes.ListIndex Then
                    Sql = "DELETE FROM tmpsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes < " & Month(vParam.fechaini)
                    Conn.Execute Sql
                    
                    Sql = "DELETE FROM tmpsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes > " & cboMes.ListIndex
                    Conn.Execute Sql
                Else
                    'Quiere  hasta parte del años siguiente
                    Sql = "DELETE FROM tmpsaldoscc WHERE codusu = " & vUsu.Codigo & " AND mes < " & Month(vParam.fechaini) & " AND mes > " & cboMes.ListIndex
                    Conn.Execute Sql
                    
                    
                
                End If
            End If
        End If
        
            
        'El ejercicio va en NOMCOST
       If Year(vParam.fechafin) = Year(vParam.fechaini) Then
            Sql = "UPDATE tmpsaldoscc SET nomccost=ano WHERE codusu=" & vUsu.Codigo
            
        Else
            Sql = "UPDATE tmpsaldoscc set nomccost=if(mes<" & Month(vParam.fechaini) & ",ano-1,ano)  WHERE codusu=" & vUsu.Codigo
        
        End If
        Conn.Execute Sql
    End If
        
    
    HacerGraficas = True
End Function

Private Function HacerRatios() As Boolean
    HacerRatios = False

    NumRegElim = DiasMes(Month(Text3(0).Text), Year(Text3(0).Text))
    If Day(Text3(0).Text) <> NumRegElim Then
        MsgBox "Saldos mensuales", vbExclamation
        Sql = NumRegElim & "/" & Format(Month(Text3(0).Text), "00") & "/" & Year(Text3(0).Text)
        Text3(0).Text = Sql
    End If


    Conn.Execute "DELETE FROM tmpimpbalance where codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM tmpimpbalan where codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM tmptesoreriacomun where codusu = " & vUsu.Codigo
    
    If Me.chkRatio(0).Value = 1 Then CargarDatosRatio 51
    If Me.chkRatio(1).Value = 1 Then CargarDatosRatio 52
    If Me.chkRatio(2).Value = 1 Then CargarDatosRatio 53
    
    
    
    Sql = "Select count(*) from tmpimpbalance where codusu=" & vUsu.Codigo
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then NumRegElim = miRsAux.Fields(0)
    End If
    miRsAux.Close
    If NumRegElim = 0 Then
        MsgBox "No existen datos"
        Exit Function
    End If
    
    'Insertaremos en la usuarios.z
    Sql = "insert into tmpimpbalan (`codusu`,`Pasivo`,`codigo`,`descripcion`,`linea`,`importe1`)"
    Sql = Sql & " select codusu,pasivo,codigo,descripcion,linea,importe1 from tmpimpbalance where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    
    
    HacerRatios = True
    
    
    
    
End Function



Private Sub InsertaEnTmpGraf(id As Long, Ingr As Currency, Gast As Currency)
Dim Aux As Currency
    If Month(vParam.fechafin) = Val(Mid(CStr(id), 5, 2)) Then
        'MEs del cierre. Hay que quitar PyG
        If CDate("01/" & Mid(CStr(id), 5, 2) & "/" & Mid(CStr(id), 1, 4)) < vParam.fechaini Then
            'Hay que quitar Cierre y Pyg
            Sql = "fechaent='" & Mid(CStr(id), 1, 4) & "-" & Mid(CStr(id), 5, 2) & "-" & Day(vParam.fechafin) & "'  AND codmacta like '7%' AND codconce"
            Sql = DevuelveDesdeBD("sum(if(isnull(timported),0,timported))-sum(if(isnull(timporteh),0,timporteh))", "hlinapu", Sql, "960")
            If Sql = "" Then Sql = "0"
            Aux = CCur(Sql)
            Ingr = Ingr + Aux
            
            Sql = "fechaent='" & Mid(CStr(id), 1, 4) & "-" & Mid(CStr(id), 5, 2) & "-" & Day(vParam.fechafin) & "'  AND codmacta like '6%' AND codconce"
            Sql = DevuelveDesdeBD("sum(if(isnull(timporteh),0,timporteh))-sum(if(isnull(timported),0,timported))", "hlinapu", Sql, "960")
            If Sql = "" Then Sql = "0"
            Aux = CCur(Sql)
            Gast = Gast + Aux
        End If
            
    End If
    Sql = "insert into `tmpgraficas` (`codusu`,`anyo`,`mes`,`ingresos`,`gastos`,`beneficio`) "
    Sql = Sql & " VALUES (" & vUsu.Codigo & "," & Mid(CStr(id), 1, 4) & "," & Mid(CStr(id), 5, 2) & ","
    Sql = Sql & TransformaComasPuntos(CStr(Ingr)) & "," & TransformaComasPuntos(CStr(Gast)) & ","
    Ingr = Ingr - Gast
    Sql = Sql & TransformaComasPuntos(CStr(Ingr)) & ")"
    Conn.Execute Sql
End Sub


Private Sub CargarDatosRatio(Cual As Integer)
Dim Lin As Collection
Dim Col As Collection

Dim J As Integer
Dim Importe As Currency
Dim ImpLin As Currency
Dim EsPasivo As Boolean

Dim Sql1 As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Sql5 As String
Dim Sql6 As String
Dim Sql7 As String
Dim Sql8 As String
Dim Sql9 As String
Dim Sql10 As String
    
    Set Lin = New Collection

    Sql = "Select * from balances_texto where numbalan=" & Cual
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Lin.Add CStr(miRsAux!Codigo)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Lin.Count = 0 Then
        Set Lin = Nothing
        Exit Sub
    End If
    
    
    For i = 1 To Lin.Count
        Me.lblInd.Caption = "Lineas " & Lin.Item(i)
        Me.lblInd.Refresh
    
        Sql1 = ""
        Sql2 = ""
        Sql3 = ""
        Sql4 = ""
        Sql5 = ""
        Sql6 = ""
        Sql7 = ""
        Sql8 = ""
        Sql9 = ""
        Sql10 = ""
    
    
    
        Set Col = New Collection
        Sql = "Select length(codmacta) longitud, balances_ctas.* from balances_ctas where numbalan=" & Cual & " AND codigo=" & Lin.Item(i) & " order by 1 "
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        While Not miRsAux.EOF
            Select Case miRsAux!Longitud
                Case 1
                    Sql1 = Sql1 & DBSet(miRsAux!codmacta, "T") & ","
                Case 2
                    Sql2 = Sql2 & DBSet(miRsAux!codmacta, "T") & ","
                Case 3
                    Sql3 = Sql3 & DBSet(miRsAux!codmacta, "T") & ","
                Case 4
                    Sql4 = Sql4 & DBSet(miRsAux!codmacta, "T") & ","
                Case 5
                    Sql5 = Sql5 & DBSet(miRsAux!codmacta, "T") & ","
                Case 6
                    Sql6 = Sql6 & DBSet(miRsAux!codmacta, "T") & ","
                Case 7
                    Sql7 = Sql7 & DBSet(miRsAux!codmacta, "T") & ","
                Case 8
                    Sql8 = Sql8 & DBSet(miRsAux!codmacta, "T") & ","
                Case 9
                    Sql9 = Sql9 & DBSet(miRsAux!codmacta, "T") & ","
                Case 10
                    Sql10 = Sql10 & DBSet(miRsAux!codmacta, "T") & ","
            End Select
            
            
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        ' quitamos la ultima coma
        If Sql1 <> "" Then Sql1 = Mid(Sql1, 1, Len(Sql1) - 1)
        If Sql2 <> "" Then Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
        If Sql3 <> "" Then Sql3 = Mid(Sql3, 1, Len(Sql3) - 1)
        If Sql4 <> "" Then Sql4 = Mid(Sql4, 1, Len(Sql4) - 1)
        If Sql5 <> "" Then Sql5 = Mid(Sql5, 1, Len(Sql5) - 1)
        If Sql6 <> "" Then Sql6 = Mid(Sql6, 1, Len(Sql6) - 1)
        If Sql7 <> "" Then Sql7 = Mid(Sql7, 1, Len(Sql7) - 1)
        If Sql8 <> "" Then Sql8 = Mid(Sql8, 1, Len(Sql8) - 1)
        If Sql9 <> "" Then Sql9 = Mid(Sql9, 1, Len(Sql9) - 1)
        If Sql10 <> "" Then Sql10 = Mid(Sql10, 1, Len(Sql10) - 1)
        
        
        '-------------------------------------------------------
        '
        '
        
        
        Importe = 0
        
        'Cuentas de pasivo. Van con Haber-Debe
        EsPasivo = False
        Select Case Cual
        Case 51
            'Ratio tesoreria
            If Lin.Item(i) = 3 Then EsPasivo = True
        Case 52
            'Liquidez
            If Lin.Item(i) = 2 Then EsPasivo = True
        Case 53
            If Lin.Item(i) >= 3 Then EsPasivo = True  '3 y 4
        
        End Select
        
       
        
        For J = 1 To 10 'Col.Count
            Me.lblInd.Caption = "Saldos " & Lin.Item(i) & ": " & J '& " de " & Col.Count
            Me.lblInd.Refresh
                
            Sql = "SELECT sum(coalesce(timported,0)-coalesce(timporteh,0)) FROM hlinapu WHERE "
            Sql = Sql & " fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(Text3(0).Text, "F")
            
            Select Case J
                Case 1
                    If Sql1 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,1) in (" & Sql1 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 2
                    If Sql2 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,2) in (" & Sql2 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 3
                    If Sql3 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,3) in (" & Sql3 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 4
                    If Sql4 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,4) in (" & Sql4 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 5
                    If Sql5 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,5) in (" & Sql5 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 6
                    If Sql6 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,6) in (" & Sql6 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 7
                    If Sql7 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,7) in (" & Sql7 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 8
                    If Sql8 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,8) in (" & Sql8 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 9
                    If Sql9 <> "" Then
                        Sql = Sql & " and mid(codmacta,1,9) in (" & Sql9 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
                Case 10
                    If Sql10 <> "" Then
                        Sql = Sql & " and codmacta in (" & Sql10 & ")"
                    Else
                        Sql = Sql & " and codmacta is null"
                    End If
            End Select
                
            
           
            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux.Fields(0)) Then
                    If EsPasivo Then
                        ImpLin = -miRsAux.Fields(0)
                    Else
                        ImpLin = miRsAux.Fields(0)
                    End If
                    
                    Importe = Importe + ImpLin
                End If
            End If
            miRsAux.Close
            
        Next J
        Set Col = Nothing
        
        NumRegElim = Cual * 100
        NumRegElim = NumRegElim + Val(Lin.Item(i))
        
        Sql = "insert into `tmpimpbalance` (`codusu`,`Pasivo`,`codigo`,`importe1`,`descripcion`,`linea`,"
        Sql = Sql & "`importe2`,`negrita`,`orden`,`QueCuentas`) values ( " & vUsu.Codigo & ",'" & Chr(Cual + 14) & "',"
        Sql = Sql & NumRegElim & "," & TransformaComasPuntos(CStr(Importe))
        Sql = Sql & ",'',NULL,NULL,NULL,'0',NULL)"
        Conn.Execute Sql
        
        'Lo que seran los textos
        
    Next i
        
    Sql = "insert into tmptesoreriacomun (`codusu`,`codigo`,`texto1`,observa1,`Texto`)"
    Sql = Sql & " select " & vUsu.Codigo & ",balances.numbalan*100+codigo,nombalan,deslinea,descripcion from balances,balances_texto where balances.numbalan=balances_texto.numbalan and balances_texto.numbalan=" & Cual & " order by orden"
    Conn.Execute Sql
        
        
End Sub



Private Sub Image2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    Sql = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If Sql <> "" Then
        Text3(Index).Text = Sql
        Text3(Index).SetFocus
    End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    chkGraf1(0).Enabled = (PreviousTab = 0)
    If PreviousTab = 1 Then chkGraf1(0).Value = 0
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub


'++
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image2_Click (Indice)
End Sub
'++



Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    Else
        If FechaCorrecta2(CDate(Text3(Index).Text)) > 2 Then
            MsgBox "Fecha mayor que ejercicios abiertos.", vbExclamation
            Text3(Index).Text = ""
            Text3(Index).SetFocus
        End If
    End If
End Sub



Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
