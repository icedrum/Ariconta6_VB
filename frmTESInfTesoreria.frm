VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESInfTesoreria 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   Icon            =   "frmTESInfTesoreria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ordenación"
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
      Left            =   7110
      TabIndex        =   34
      Top             =   3090
      Width           =   4365
      Begin VB.OptionButton optVarios 
         Caption         =   "Fecha vencimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   36
         Top             =   840
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   35
         Top             =   1320
         Width           =   1455
      End
   End
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
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   37
         Tag             =   "imgFecha"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1305
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1200
         Width           =   4155
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   780
         Width           =   4155
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
         Left            =   1260
         TabIndex        =   1
         Tag             =   "imgCuentas"
         Top             =   1200
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
         Index           =   0
         Left            =   1260
         TabIndex        =   0
         Tag             =   "imgCuentas"
         Top             =   780
         Width           =   1275
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
         Tag             =   "imgFecha"
         Top             =   2400
         Width           =   1305
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
         Left            =   4560
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   5250
         Top             =   2040
         Visible         =   0   'False
         Width           =   240
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
         Index           =   1
         Left            =   270
         TabIndex        =   33
         Top             =   480
         Width           =   2370
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
         Index           =   10
         Left            =   270
         TabIndex        =   32
         Top             =   810
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
         Index           =   9
         Left            =   270
         TabIndex        =   31
         Top             =   1170
         Width           =   615
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   960
         Top             =   1230
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   960
         Top             =   780
         Width           =   255
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   3990
         Width           =   4095
      End
      Begin VB.Label lblPrevInd 
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
         Left            =   2640
         TabIndex        =   22
         Top             =   2520
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
         TabIndex        =   21
         Top             =   2700
         Width           =   4035
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   960
         Top             =   2400
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
         TabIndex        =   20
         Top             =   2400
         Width           =   615
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
         TabIndex        =   19
         Top             =   1920
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
      Height          =   2985
      Left            =   7110
      TabIndex        =   6
      Top             =   0
      Width           =   4305
      Begin MSComctlLib.ListView ListView1 
         Height          =   1890
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3334
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
         TabIndex        =   28
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
         Picture         =   "frmTESInfTesoreria.frx":000C
         ToolTipText     =   "Quitar al Debe"
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   3780
         Picture         =   "frmTESInfTesoreria.frx":0156
         ToolTipText     =   "Puntear al Debe"
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Detallar"
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
         TabIndex        =   26
         Top             =   660
         Width           =   1920
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   27
      Top             =   6000
      Width           =   6270
   End
End
Attribute VB_Name = "frmTESInfTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 902

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

Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
'Private WithEvents frmGas As frmBasico

Private Sql As String
Dim cad As String
Dim RC As String
Dim I As Integer


Dim PrimeraVez As Boolean
Dim Cancelado As Boolean

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


    MontaSQL = False
    
    If Not PonerDesdeHasta("cobros.FecFactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    If Not PonerDesdeHasta("cobros.ctabanc1", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
            
    MontaSQL = True
End Function


Private Sub cmdAccion_Click(Index As Integer)

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    
    
    If Not MontaSQL Then Exit Sub
    
    If Not CargaDatos Then Exit Sub
    
  
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
    Me.Caption = "Informe de Situación por Cuenta"

    For I = 1 To 1
        Me.ImgFec(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next I
     
    For I = 0 To 1
        Me.imgCuentas(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
     
    CargarListtipos
    
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    
End Sub

Private Sub CargarListtipos()
  Dim IT
    On Error GoTo ECargarList

    'Los encabezados
    NumRegElim = 1
    ListView1(NumRegElim).ColumnHeaders.Clear

    ListView1(NumRegElim).ColumnHeaders.Add , , "Tipo", 1800
    ListView1(NumRegElim).ListItems.Clear
    
    
    For I = 1 To 3
            Set IT = ListView1(NumRegElim).ListItems.Add
            IT.Key = "C" & I
            IT.Checked = True
            IT.Text = RecuperaValor("Cobros|Pagos|Gastos|", I)
    Next

ECargarList:
    
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


Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Sql = Format(vFecha, "dd/mm/yyyy")
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
    Set frmBan = New frmBasico2
    AyudaBanco frmBan
    Set frmBan = Nothing
    If Sql <> "" Then
        
        txtCuentas(Index).Text = RecuperaValor(Sql, 1)
        Me.txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
        PonFoco Me.txtCuentas(Index)
    End If
    AbiertoOtroFormEnListado = False
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
        
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        Sql = ""
        frmF.Show vbModal
        Set frmF = Nothing
        If Sql <> "" Then txtFecha(Index).Text = Sql
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
        Case 0, 1 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
            If Not B Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
                PonFoco txtCuentas(Index)
            Else
            
                cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Cta, "T")
                If cad = "" Then
                    MsgBox "No existe banco: " & Cta & " - " & Sql, vbExclamation
                    
                    Sql = ""
                End If
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = Sql
            
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


    'Monto el SQL
    
    Sql = "SELECT `tmptesoreriacomun`.`texto1` cuenta , `tmptesoreriacomun`.`texto2` Conta, `tmptesoreriacomun`.`opcion` BD, `tmptesoreriacomun`.`texto5` Nombre, `tmptesoreriacomun`.`texto3` NroFra, `tmptesoreriacomun`.`fecha1` FecFra, `tmptesoreriacomun`.`fecha2` FecVto, `tmptesoreriacomun`.`importe1` Gasto, `tmptesoreriacomun`.`importe2` Recibo"
    Sql = Sql & " FROM   `tmptesoreriacomun` `tmptesoreriacomun`"
    Sql = Sql & " WHERE `tmptesoreriacomun`.codusu = " & vUsu.Codigo
    Sql = Sql & " ORDER BY `tmptesoreriacomun`.`texto1`, `tmptesoreriacomun`.`texto2`, `tmptesoreriacomun`.`opcion`, `tmptesoreriacomun`.`fecha1`"
    
    
    Sql = " select tmpconextcab.cta Cta, tmpconextcab.cuenta descripcion,"
    If Me.optVarios(0).Value Then
        Sql = Sql & " FechaEnt fechaVto  , contra tipo"
        RC = " fechaent,tipo"
    Else
        Sql = Sql & "  contra tipo,FechaEnt fechaVto  "
        RC = " tipo,fechaent"
    End If
    Sql = Sql & " ,ampconce,nomdocum,acumtotD SumDebe,acumtotH SumHaber, acumtotT SaldoInicial, timporteD debe ,timporteH  haber"
    Sql = Sql & " from tmpconextcab, tmpconext where tmpconextcab.codusu= tmpconext.codusu and  tmpconextcab.cta= tmpconext.cta"
    Sql = Sql & " and tmpconextcab.codusu=" & vUsu.Codigo & " ORDER BY tmpconextcab.cta," & RC & ",ampconce"



    
    
    
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0903-00"
    
        
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "GastosFijos.rpt"
    
    cadFormula = "{tmpconextcab.codusu} = " & vUsu.Codigo
    
    
    Sql = IIf(Me.optVarios(0).Value, "Fecha vto.", "tipo")
    Sql = "Titulo= ""Informe tesorería (" & Sql & ")""|"
    'Fechas intervalor
    Sql = Sql & "Fechas= ""Fecha hasta " & txtFecha(1).Text & """|"
    
    Sql = Sql & "Cuenta= """ & RC & """|"
    Sql = Sql & "pFecha= """ & Format(Now, "dd/mm/yyyy") & """|"
    Sql = Sql & "NumPag= 0|"
    Sql = Sql & "Salto= 2|"
    cadParam = cadParam & Sql
    numParam = 8
    
    
    If Me.optVarios(0).Value Then
        Sql = "orden1= {tmpconext.fechaent} |" & "orden2= {tmpconext.contra}|"
    Else
         Sql = "orden1= {tmpconext.contra} |" & "orden2= {tmpconext.fechaent}|"
    End If
    cadParam = cadParam & Sql
    numParam = numParam + 2
    
    
    
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 51
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
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
    
    If Me.txtFecha(1).Text = "" Then
        MsgBox "Seleccione fecha", vbExclamation
        PonFoco txtFecha(1)
        
    Else
        DatosOK = True
    End If

End Function


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub








Private Function CargaDatos() As Boolean
Dim Rs As ADODB.Recordset
    'Borramos las lineas en usuarios
    lblPrevInd.Caption = "Preparando ..."
    lblPrevInd.Refresh
    
    Conn.Execute "DELETE FROM tmpconext WHERE codusu =" & vUsu.Codigo
    Conn.Execute "DELETE FROM tmpconextcab WHERE codusu =" & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset



    'Hacemos el selecet
    Sql = "select cuentas.codmacta,nommacta from bancos,cuentas where cuentas.codmacta=bancos.codmacta"
    If txtCuentas(0).Text <> "" Then Sql = Sql & " AND bancos.codmacta >= " & DBSet(txtCuentas(0), "T")
    If txtCuentas(1).Text <> "" Then Sql = Sql & " AND bancos.codmacta <= " & DBSet(txtCuentas(1), "T")
    
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    NumRegElim = 0
    While Not Rs.EOF
        '---
        If Not HacerPrevisionCuenta(Rs!codmacta, Rs!Nommacta) Then
        '---
            Sql = "DELETE FROM tmpconextcab WHERE codusu =" & vUsu.Codigo
            Sql = Sql & " AND cta ='" & Rs!codmacta & "'"
            Conn.Execute Sql
        Else
            NumRegElim = NumRegElim + 1
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    lblPrevInd.Caption = ""
    Screen.MousePointer = vbDefault
    

    If NumRegElim = 0 Then
        MsgBox "Ningun dato generado", vbExclamation
    Else
        CargaDatos = True
    End If

   
    
    
 
End Function






Private Function HacerPrevisionCuenta(Cta As String, Nommacta As String) As Boolean
Dim SaldoArrastrado As Currency
Dim Id As Currency
Dim IH As Currency
Dim Importe As Currency


    HacerPrevisionCuenta = False
    
    lblPrevInd.Caption = Cta & " - " & Nommacta
    lblPrevInd.Refresh
    ' Las fechas son del periodo, luego me importa una mierda las fechas desde hasta
    '
    '
    cad = "INSERT INTO tmpconextcab (codusu,cta,fechini,fechfin,cuenta,acumtotD,acumtotH,acumtotT) VALUES (" & vUsu.Codigo & ", '" & Cta & "','" & Format(vParam.fechaini, "dd/mm/yyyy") & "','" & Format(Me.txtFecha(1).Text, "dd/mm/yyyy") & "','" & Nommacta & "',"
    
    RC = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH from hlinapu where codmacta=" & DBSet(Cta, "T")
    RC = RC & " AND fechaent >=  " & DBSet(vParam.fechaini, "F")
    RC = RC & " AND fechaent <=  " & DBSet(Now, "F")  'Hasta el dia de hoy
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        RC = "0,0,0"
    Else
        RC = DBSet(DBLet(miRsAux!SumaDetimporteD, "N"), "N") & "," & DBSet(DBLet(miRsAux!SumaDetimporteH), "N", "N") & ","
        RC = RC & DBSet(DBLet(miRsAux!SumaDetimporteD, "N") - DBLet(miRsAux!SumaDetimporteH, "N"), "N")
    End If
    miRsAux.Close
    
    RC = cad & RC & ")"
    Conn.Execute RC
    
    
    
    'RC = "INSERT INTO tmpfaclin (codusu, IVA,codigo, Fecha, Cliente, cta,"
    'RC = RC & " ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
    
    RC = "INSERT INTO tmpconext(codusu,cta,contra,Pos,fechaent,ampconce,nomdocum,timporteD,timporteH)  VALUES  "
            
            
            
    'PARA CADA CUENTA
    'mETEREMOS TODOS LOS REGISTROS EN LA TABLA  tmpconext
    
    
    'TANTO COBROS COMO PAGOS I GASTOS
    '
    'Luego, en funcion del orden(TIPO o fecha) los iremos insertando en la tabla, para que
    'el saldo que va arrastrando sea el correcto
    
    
       
        
    CONT = 0
    
    
    '--------------------
    'DETALLAR COBROS
    lblPrevInd.Caption = Cta & " - Cobros"
    lblPrevInd.Refresh
    Sql = " WHERE fecvenci<='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    Sql = Sql & " AND ctabanc1 ='" & Cta & "'"
    Sql = Sql & "  and ((impvenci +coalesce(gastos,0)) - coalesce(impcobro,0))<>0 " 'pendiente
    cad = ""
    If Not ListView1(1).ListItems(1).Checked Then
        'NO DETALLA COBROS
        Sql = "select sum(impvenci),sum(impcobro),fecvenci from cobros " & Sql
        Sql = Sql & " GROUP BY fecvenci"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
        
            Id = DBLet(miRsAux.Fields(0), "N")
            IH = DBLet(miRsAux.Fields(1), "N")
            Importe = Id - IH

            
            CONT = CONT + 1
            cad = cad & ", (" & vUsu.Codigo & ",'" & Cta & "','COBRO'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "',NULL,'COBROS PENDIENTES',"
            
            'HAY COBROS
            If Importe < 0 Then
                cad = cad & "0," & TransformaComasPuntos(CStr(Abs(Importe)))
            Else
                cad = cad & TransformaComasPuntos(CStr(Importe)) & ",0"
            End If
            cad = cad & ")"
                
                
            miRsAux.MoveNext
        Wend
        
                
    Else
         'DETALLAR PAGOS COBROS
            '(codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
            'timporteD,timporteH, saldo
            
        'SQL = "select scobro.*,nommacta from scobro,cuentas where scobro.codmacta=cuentas.codmacta"
        'SQL = SQL & " AND fecvenci<='2006-01-01'"
         
        Sql = "select cobros.* from cobros " & Sql
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not miRsAux.EOF
            CONT = CONT + 1
            cad = cad & ", (" & vUsu.Codigo & ",'" & Cta & "','COBRO'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "','"
            'NUmero factura
            cad = cad & miRsAux!NUmSerie & miRsAux!numfactu & IIf(miRsAux!numorden > 1, "/" & miRsAux!numorden, "") & "',"
            
            cad = cad & DBSet(Trim(miRsAux!codmacta & " " & DBLet(miRsAux!nomclien, "T")), "T") & ","
            Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
            If Importe <> 0 Then
                If Importe < 0 Then
                    cad = cad & "0," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Importe)) & ",0"
                End If
                cad = cad & ")"
                
                HacerInsertTmp RC, cad, False
            End If
            miRsAux.MoveNext
            
        Wend
        
        
    End If
    miRsAux.Close
    If Len(cad) > 0 Then HacerInsertTmp RC, cad, True
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR PAGOS
    '--------------------
    '--------------------
    lblPrevInd.Caption = Cta & " - pagos"
    lblPrevInd.Refresh
    Sql = " WHERE fecefect<='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    Sql = Sql & " AND ctabanc1 ='" & Cta & "' and (ImpEfect - coalesce(imppagad,0)) <>0"
    
    cad = ""
    If Not ListView1(1).ListItems(2).Checked Then
        Sql = "select sum(impefect),sum(imppagad),fecefect from pagos " & Sql & " GROUP BY fecefect"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Importe = 0
        While Not miRsAux.EOF

            Id = DBLet(miRsAux.Fields(0), "N")
            IH = DBLet(miRsAux.Fields(1), "N")
            Importe = Id - IH
            
            
            CONT = CONT + 1
            cad = cad & ", (" & vUsu.Codigo & ",'" & Cta & "','PAGO'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "',NULL,'PAGOS PENDIENTES',"
            
            'HAY COBROS
            If Importe > 0 Then
              cad = cad & "0," & TransformaComasPuntos(CStr(Importe))
            Else
              cad = cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",0"
            End If
            cad = cad & ")"
            miRsAux.MoveNext
        Wend
        
    Else
         'DETALLAR PAGOS
        
        
        Sql = "select pagos.* from pagos " & Sql
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            cad = cad & ", (" & vUsu.Codigo & ",'" & Cta & "','PAGO'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','"
            'NUmero factura
            cad = cad & DevNombreSQL(miRsAux!numfactu) & IIf(miRsAux!numorden = 1, "", "/" & miRsAux!numorden) & "',"
            
            cad = cad & DBSet(Trim(miRsAux!codmacta & " " & DBLet(miRsAux!nomprove, "T")), "T") & ","
            Importe = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
            If Importe <> 0 Then
                If Importe > 0 Then
                    cad = cad & "0," & TransformaComasPuntos(CStr(Importe))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",0"
                End If
                cad = cad & ")"
                
                HacerInsertTmp RC, cad, False
            End If
            miRsAux.MoveNext
        Wend
        
        
    End If
    miRsAux.Close
    
    If Len(cad) > 0 Then HacerInsertTmp RC, cad, True
    
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR GASTOS GASTOS
    '--------------------
    '--------------------
    
    Sql = " from gastosfijos,gastosfijos_recibos where gastosfijos.codigo= gastosfijos_recibos.codigo"
    Sql = Sql & " and fecha >='" & Format(Now, FormatoFecha)
    Sql = Sql & "' AND fecha <='" & Format(Format(txtFecha(1).Text, FormatoFecha), FormatoFecha) & "'"
    Sql = Sql & " and ctaprevista='" & Cta & "'"
    Sql = Sql & " and contabilizado=0"
    
    'recorro el recodset
    If Not ListView1(1).ListItems(3).Checked Then
        cad = "SELECT gastosfijos.codigo,concat('Nº: ',count(*), '-', descripcion) descripcion, max(fecha) fecha ,sum(importe) importe "
    Else
        cad = " select gastosfijos.codigo,descripcion,fecha,importe "
    End If
    Sql = cad & Sql
    If Not ListView1(1).ListItems(3).Checked Then Sql = Sql & " GROUP BY gastosfijos.codigo"
    
    cad = ""
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CONT = CONT + 1
        cad = cad & ", (" & vUsu.Codigo & ",'" & Cta & "','GASTO'," & CONT & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "','"
        'NUmero factura
        cad = cad & "ID " & Format(miRsAux!Codigo, "0000") & "'," & DBSet(miRsAux!Descripcion, "T") & ","
        
        'cad = cad & DBSet(Trim(miRsAux!codmacta & " " & DBLet(miRsAux!nomprove, "T")), "T") & ","
        
        Importe = miRsAux!Importe
        If Importe <> 0 Then
            If Importe > 0 Then
                cad = cad & "0," & TransformaComasPuntos(CStr(Importe))
            Else
                cad = cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",0"
            End If
            cad = cad & ")"
            
            HacerInsertTmp RC, cad, False
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Len(cad) > 0 Then HacerInsertTmp RC, cad, True
    
    If CONT > 0 Then HacerPrevisionCuenta = True
    
End Function



Private Sub HacerInsertTmp(ByRef InsertInto As String, ByRef LosValues As String, Forzar As Boolean)
    
    If Not Forzar Then
        If Len(LosValues) > 3000 Then Forzar = True
    End If
    
    If Forzar Then
        LosValues = Mid(LosValues, 2)
        LosValues = InsertInto & LosValues
        Conn.Execute LosValues
        LosValues = ""
    End If
    
        
End Sub
