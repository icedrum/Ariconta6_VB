VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCCDetalleExplota 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   7200
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
      Height          =   4605
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6915
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
         ItemData        =   "frmCCDetalleExplota.frx":0000
         Left            =   1230
         List            =   "frmCCDetalleExplota.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2220
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
         Index           =   1
         ItemData        =   "frmCCDetalleExplota.frx":0004
         Left            =   1230
         List            =   "frmCCDetalleExplota.frx":0006
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
         Index           =   1
         Left            =   3270
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2670
         Width           =   855
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
         Index           =   1
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3900
         Width           =   1305
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
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   3900
         Width           =   4185
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
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3480
         Width           =   4185
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1050
         Width           =   4185
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
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1470
         Width           =   4185
      End
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
         Left            =   1230
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1050
         Width           =   1275
      End
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
         Index           =   1
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
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
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3480
         Width           =   1305
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   6360
         TabIndex        =   32
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
         TabIndex        =   38
         Top             =   1920
         Width           =   1410
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
         TabIndex        =   37
         Top             =   2280
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
         TabIndex        =   36
         Top             =   2640
         Width           =   615
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
         Index           =   1
         Left            =   270
         TabIndex        =   35
         Top             =   3930
         Width           =   600
      End
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   1
         Left            =   900
         Top             =   3930
         Width           =   255
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
         Index           =   0
         Left            =   270
         TabIndex        =   33
         Top             =   3510
         Width           =   600
      End
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   0
         Left            =   900
         Top             =   3510
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   1500
         Width           =   255
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
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   3120
         Width           =   2130
      End
      Begin VB.Label lblAsiento 
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
         Left            =   2550
         TabIndex        =   27
         Top             =   990
         Width           =   4095
      End
      Begin VB.Label lblAsiento 
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
         Left            =   2550
         TabIndex        =   26
         Top             =   1440
         Width           =   4095
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   1080
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
         TabIndex        =   23
         Top             =   690
         Width           =   960
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
      TabIndex        =   10
      Top             =   7530
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
      Left            =   4260
      TabIndex        =   8
      Top             =   7530
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
      TabIndex        =   9
      Top             =   7470
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
      TabIndex        =   11
      Top             =   4710
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
         TabIndex        =   22
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   21
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   20
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1500
      TabIndex        =   39
      Top             =   7560
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
End
Attribute VB_Name = "frmCCDetalleExplota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1005



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


Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCCo As frmBasico
Attribute frmCCo.VB_VarHelpID = -1

Private SQL As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String

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
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmplinccexplo", "tmplinccexplo.codusu= " & vUsu.Codigo) Then Exit Sub
    
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
    Me.Caption = "Detalle de Explotación"

    For i = 0 To 1
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    For i = 0 To 1
        Me.imgCCoste(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    PrimeraVez = True
     
     
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
End Sub

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
        i = Val(cad)
        With cmbFecha(i)
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




Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCuentas(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCuentas(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCCo_DatoSeleccionado(CadenaSeleccion As String)
    txtCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ImgCCoste_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmCCo = New frmBasico
    
    AyudaCC frmCCo, txtCCoste(Index)
    
    Set frmCCo = Nothing
    
    PonFoco txtCCoste(Index)

End Sub


Private Sub imgCuentas_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing

    PonFoco txtCuentas(Index)

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

Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub


Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgCCuentas"
        imgCuentas_Click Indice
    Case "imgCCoste"
        ImgCCoste_Click Indice
    End Select
    
End Sub


Private Sub txtCuentas_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    
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



    Select Case Index
        Case 0, 1 'Cuentas
            
            RC = txtCuentas(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, SQL) Then
                txtCuentas(Index) = RC
                txtNCuentas(Index).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
                PonFoco txtCuentas(Index)
            End If
            
            If Index = 0 Then Hasta = 1
            If Hasta >= 1 Then
                txtCuentas(Hasta).Text = txtCuentas(Index).Text
                txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
            End If
    End Select

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


Private Sub txtCCoste_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCCoste(Index).Text = Trim(txtCCoste(Index).Text)
    
    Select Case Index
        Case 0, 1 'Centros de Coste
            If txtCCoste(Index).Text <> "" Then txtCCoste(Index).Text = UCase(txtCCoste(Index).Text)
            txtNCCoste(Index) = PonerNombreDeCod(txtCCoste(Index), "ccoste", "nomccost", "codccost", "T")
    End Select

End Sub


Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL

    SQL = "Select tmplinccexplo.codccost CCoste, ccoste.nomccost Descripción, tmplinccexplo.codmacta as Cuenta, cuentas.nommacta Título, tmplinccexplo.fechaent Fecha,  "
    SQL = SQL & " docum Documento, ctactra Contrapartida, cuentas1.nommacta Descripción , tmplinccexplo.ampconce Concepto, coalesce(perd,0) Debe, coalesce(perh,0) Haber, "
    SQL = SQL & " if (@Cta <> tmplinccexplo.codmacta,@Saldo:= coalesce(perd,0) - coalesce(perh,0), @Saldo:= @Saldo + coalesce(perd,0) - coalesce(perh,0) ) Saldo, "
    SQL = SQL & " @Cta:= tmplinccexplo.codmacta Cta"

    SQL = SQL & " FROM ((tmplinccexplo INNER JOIN cuentas ON tmplinccexplo.codmacta = cuentas.codmacta)  "
    SQL = SQL & " INNER JOIN ccoste ON tmplinccexplo.codccost = ccoste.codccost) "
    SQL = SQL & " LEFT JOIN cuentas cuentas1 ON tmplinccexplo.ctactra = cuentas1.codmacta , (select @Saldo:= 0) aaa, (select @Cta:= '') bbb   "
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")

    If cadselect <> "" Then SQL = SQL & " and " & cadselect
    SQL = SQL & " ORDER BY 1,2,3,4,5 "


    
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "1005-00" '"CC_X_CTA.rpt"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    
    cadFormula = "{tmplinccexplo.codusu}=" & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 33
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("hlinapu.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
    If Not PonerDesdeHasta("hlinapu.codccost", "CCO", Me.txtCCoste(0), Me.txtNCCoste(0), Me.txtCCoste(1), Me.txtNCCoste(1), "pDHCCoste=""") Then Exit Function
    
    MontaSQL = CargarTemporal
           
End Function

Private Function CargarTemporal() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim ImporteTot As Currency
Dim ImporteLinea As Currency
Dim UltSubCC As String
Dim B As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    
    SQL = "delete from tmplinccexplo where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute SQL
    
    SQL = "insert into tmplinccexplo (codusu,codccost,codmacta,linapu,docum,fechaent,ampconce,ctactra,desctra,perD,perH) "
    SQL = SQL & " Select " & vUsu.Codigo & ", hlinapu.codccost CCoste,  hlinapu.codmacta as Cuenta, hlinapu.linliapu, hlinapu.numdocum, hlinapu.fechaent,  "
    SQL = SQL & " hlinapu.ampconce, ctacontr Contrapartida, cuentas1.nommacta Descripción , coalesce(timported,0) Debe, coalesce(timporteh,0) Haber "
    SQL = SQL & " FROM (hlinapu LEFT JOIN cuentas cuentas1 ON hlinapu.ctacontr = cuentas1.codmacta) INNER JOIN ccoste on hlinapu.codccost = ccoste.codccost "
    SQL = SQL & " where mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & ")"
    SQL = SQL & " and hlinapu.fechaent >= " & DBSet(FechaInicio, "F") & " and hlinapu.fechaent <= " & DBSet(fechafin, "F")
    If cadselect <> "" Then SQL = SQL & " and " & cadselect
    SQL = SQL & " union "
    SQL = SQL & " Select " & vUsu.Codigo & ", hlinapu.codccost CCoste,  hlinapu.codmacta as Cuenta, hlinapu.linliapu, hlinapu.numdocum, hlinapu.fechaent,  "
    SQL = SQL & " hlinapu.ampconce, ctacontr Contrapartida, cuentas1.nommacta Descripción , coalesce(timported,0) Debe, coalesce(timporteh,0) Haber "
    SQL = SQL & " FROM (hlinapu LEFT JOIN cuentas cuentas1 ON hlinapu.ctacontr = cuentas1.codmacta) INNER JOIN ccoste_lineas ON hlinapu.codccost = ccoste_lineas.codccost "
    SQL = SQL & " where mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & ")"
    SQL = SQL & " and hlinapu.fechaent >= " & DBSet(FechaInicio, "F") & " and hlinapu.fechaent <= " & DBSet(fechafin, "F")
    If cadselect <> "" Then SQL = SQL & " and " & Replace(cadselect, "hlinapu.codccost", "ccoste_lineas.subccost")
    
    
    SQL = SQL & " ORDER BY 1,2,3,4,5 "
    
    Conn.Execute SQL
    
    B = HacerRepartoSubcentrosCoste
    
    CargarTemporal = B
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function


Private Function HacerRepartoSubcentrosCoste() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim ImporteTot As Currency
Dim ImporteLinea As Currency
Dim UltSubCC As String
Dim Nregs As Long

    On Error GoTo eHacerRepartoSubcentrosCoste

    HacerRepartoSubcentrosCoste = False
    
    ' hacemos el desdoble
    SQL = "select * from tmplinccexplo where codusu = " & DBSet(vUsu.Codigo, "N") & " and codccost in (select ccoste.codccost from ccoste inner join ccoste_lineas on ccoste.codccost = ccoste_lineas.codccost) "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Nregs = TotalRegistrosConsulta(SQL)

    If Nregs <> 0 Then
        pb2.visible = True
        CargarProgres pb2, Nregs
    End If


    While Not Rs.EOF
        IncrementarProgres pb2, 1
        
        Sql2 = "select ccoste.codccost, subccost, porccost from ccoste inner join ccoste_lineas on ccoste.codccost = ccoste_lineas.codccost where ccoste.codccost =  " & DBSet(Rs!codccost, "T")

        ImporteTot = 0
        UltSubCC = ""

        Set Rs2 = New ADODB.Recordset

        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            SQL = "insert into tmplinccexplo (codusu,codccost,codmacta,linapu,docum,fechaent,ampconce,ctactra,desctra,perD,perH,desdoblado) values ("
            SQL = SQL & vUsu.Codigo & "," & DBSet(Rs2!subccost, "T") & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(Rs!LINAPU, "T") & ","
            SQL = SQL & DBSet(Rs!DOCUM, "T") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!Ampconce, "T") & "," & DBSet(Rs!ctactra, "T") & ","
            SQL = SQL & DBSet(Rs!desctra, "T") & ","

            If DBLet(Rs!perd, "N") <> 0 Then
                ImporteLinea = Round(DBLet(Rs!perd, "N") * DBLet(Rs2!porccost, "N") / 100, 2)
                SQL = SQL & DBSet(ImporteLinea, "N") & ",0,1)"
            Else
                ImporteLinea = Round(DBLet(Rs!perh, "N") * DBLet(Rs2!porccost, "N") / 100, 2)
                SQL = SQL & "0," & DBSet(ImporteLinea, "N") & ",1)"
            End If

            Conn.Execute SQL

            ImporteTot = ImporteTot + ImporteLinea

            UltSubCC = Rs2!subccost

            Rs2.MoveNext
        Wend

        If DBLet(Rs!perd, "N") <> 0 Then
            If ImporteTot <> DBLet(Rs!perd, "N") Then
                SQL = "update tmplinccexplo set perd = perd + (" & DBSet(Round(DBLet(Rs!perd, "N") - ImporteTot, 2), "N") & ")"
                SQL = SQL & " where codusu = " & vUsu.Codigo
                SQL = SQL & " and codccost = " & DBSet(UltSubCC, "T")
                SQL = SQL & " and codmacta = " & DBSet(Rs!codmacta, "T")
                SQL = SQL & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                SQL = SQL & " and linapu = " & DBSet(Rs!LINAPU, "N")
                SQL = SQL & " and docum = " & DBSet(Rs!DOCUM, "T")
                SQL = SQL & " and ampconce = " & DBSet(Rs!Ampconce, "T")
                SQL = SQL & " and desdoblado = 1"

                Conn.Execute SQL
            End If
        Else
            If ImporteTot <> DBLet(Rs!perh, "N") Then
                SQL = "update tmplinccexplo set perh = perh + (" & DBSet(Round(DBLet(Rs!perh, "N") - ImporteTot, 2), "N") & ")"
                SQL = SQL & " where codusu = " & vUsu.Codigo
                SQL = SQL & " and codccost = " & DBSet(UltSubCC, "T")
                SQL = SQL & " and codmacta = " & DBSet(Rs!codmacta, "T")
                SQL = SQL & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                SQL = SQL & " and linapu = " & DBSet(Rs!LINAPU, "N")
                SQL = SQL & " and docum = " & DBSet(Rs!DOCUM, "T")
                SQL = SQL & " and ampconce = " & DBSet(Rs!Ampconce, "T")
                SQL = SQL & " and desdoblado = 1"

                Conn.Execute SQL
            End If
        End If

        SQL = "delete from tmplinccexplo where codusu = " & vUsu.Codigo
        SQL = SQL & " and codccost = " & DBSet(Rs!codccost, "T")
        SQL = SQL & " and codmacta = " & DBSet(Rs!codmacta, "T")
        SQL = SQL & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
        SQL = SQL & " and linapu = " & DBSet(Rs!LINAPU, "N")
        SQL = SQL & " and docum = " & DBSet(Rs!DOCUM, "T")
        SQL = SQL & " and ampconce = " & DBSet(Rs!Ampconce, "T")
        SQL = SQL & " and desdoblado = 0"

        Conn.Execute SQL

        Set Rs2 = Nothing


        Rs.MoveNext
    Wend

    Set Rs = Nothing

    'falta el borrado de los que no tocan
    If txtCCoste(0).Text <> "" Or txtCCoste(1).Text <> "" Then
        SQL = "delete from tmplinccexplo where codusu = " & vUsu.Codigo
        SQL = SQL & " and not codccost in (select codccost from ccoste where (1=1) "
        If txtCCoste(0).Text <> "" Then SQL = SQL & " and codccost >= " & DBSet(txtCCoste(0).Text, "T")
        If txtCCoste(1).Text <> "" Then SQL = SQL & " and codccost <= " & DBSet(txtCCoste(1).Text, "T")
        SQL = SQL & ")"
        
        Conn.Execute SQL
    End If


    HacerRepartoSubcentrosCoste = True
    pb2.visible = False
    Exit Function
    
eHacerRepartoSubcentrosCoste:
    MuestraError Err.Number, "Reparto Subcentros de Coste", Err.Description
    pb2.visible = False
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
