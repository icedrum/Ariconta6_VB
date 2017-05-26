VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturasCliCtaVtas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11685
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
      Height          =   4395
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   6915
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
         Index           =   3
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2550
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
         Index           =   2
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2130
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1320
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   900
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
         Index           =   3
         Left            =   1260
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   2580
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
         Index           =   2
         Left            =   1260
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   2130
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
         Left            =   1260
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1320
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
         Tag             =   "imgConcepto"
         Top             =   900
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
         TabIndex        =   5
         Tag             =   "imgConcepto"
         Top             =   3810
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
         TabIndex        =   4
         Tag             =   "imgConcepto"
         Top             =   3390
         Width           =   1305
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   3
         Left            =   960
         Top             =   2610
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   2
         Left            =   960
         Top             =   2160
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
         Left            =   270
         TabIndex        =   34
         Top             =   2550
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
         Index           =   1
         Left            =   270
         TabIndex        =   33
         Top             =   2190
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Clientes"
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
         Left            =   270
         TabIndex        =   32
         Top             =   1800
         Width           =   2040
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Ventas"
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
         Index           =   11
         Left            =   270
         TabIndex        =   29
         Top             =   540
         Width           =   1650
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
         TabIndex        =   28
         Top             =   930
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
         TabIndex        =   27
         Top             =   1290
         Width           =   615
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   960
         Top             =   1350
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   960
         Top             =   900
         Width           =   255
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmFacturasCliCtaVtas.frx":0000
         Top             =   3810
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmFacturasCliCtaVtas.frx":008B
         Top             =   3420
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
         Index           =   4
         Left            =   270
         TabIndex        =   26
         Top             =   3810
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
         Left            =   270
         TabIndex        =   25
         Top             =   3450
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Factura"
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
         Left            =   270
         TabIndex        =   24
         Top             =   3090
         Width           =   1830
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
      Height          =   7095
      Left            =   7140
      TabIndex        =   30
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "Clasificar por importe"
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
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   780
         Width           =   3075
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Comparativo Año Anterior"
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
         Left            =   300
         TabIndex        =   8
         Top             =   1890
         Width           =   3075
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Detallar facturas"
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
         Left            =   300
         TabIndex        =   7
         Top             =   1320
         Width           =   3075
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   31
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
      Left            =   10380
      TabIndex        =   11
      Top             =   7230
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
      Left            =   8820
      TabIndex        =   9
      Top             =   7230
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
      TabIndex        =   10
      Top             =   7200
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
      TabIndex        =   12
      Top             =   4440
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
         TabIndex        =   23
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   22
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
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
      Index           =   24
      Left            =   1890
      TabIndex        =   35
      Top             =   7290
      Width           =   5145
   End
End
Attribute VB_Name = "frmFacturasCliCtaVtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 403


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
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim CadSelect1 As String


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



Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 0
            Check1(1).Enabled = (Check1(Index).Value = 0)
            If Check1(Index).Value = 1 Then Check1(1).Value = 0
        Case 1
            Check1(0).Enabled = (Check1(Index).Value = 0)
            If Check1(Index).Value = 1 Then Check1(0).Value = 0
    End Select
End Sub

Private Sub cmdAccion_Click(Index As Integer)

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    tabla = "factcli inner join factcli_lineas on factcli.numserie = factcli_lineas.numserie and factcli.numfactu = factcli_lineas.numfactu and factcli.anofactu = factcli_lineas.anofactu "
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmpfaclin", "tmpfaclin.codusu=" & vUsu.Codigo) Then Exit Sub
    
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
        
    'Otras opciones
    Me.Caption = "Relación de Clientes por Cta Ventas"
     
    For i = 0 To 3
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
    txtFecha(0).Text = vParam.fechaini
    txtFecha(1).Text = vParam.fechafin
    If Not vParam.FecEjerAct Then
        txtFecha(1).Text = Format(DateAdd("yyyy", 1, vParam.fechafin), "dd/mm/yyyy")
    End If
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
End Sub



Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
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

Private Sub imgCuentas_Click(Index As Integer)
    Sql = ""
    AbiertoOtroFormEnListado = True
    Set frmCtas = New frmColCtas
    frmCtas.DatosADevolverBusqueda = True
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    If Sql <> "" Then
        Me.txtCuentas(Index).Text = RecuperaValor(Sql, 1)
        Me.txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
    Else
        QuitarPulsacionMas Me.txtCuentas(Index)
    End If
    
    PonFoco Me.txtCuentas(Index)
    AbiertoOtroFormEnListado = False
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
   ' KEYpressGnral KeyAscii
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
        Case 0, 1, 2, 3 'cuentas
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
                If Index = 0 Then
                    Hasta = 1
                Else
                    If Index = 2 Then
                        Hasta = 3
                    End If
                End If
                    
                If Hasta >= 0 Then
                    txtCuentas(Hasta).Text = txtCuentas(Index).Text
                    txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
                End If
            End If
    
    
    End Select

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
Dim Rs As ADODB.Recordset
Dim Total As Currency
Dim TotalAnt As Currency

    'Monto el SQL
    If Check1(0).Value = 1 Then
        Sql = "Select  tmpfaclin.ctabase CtaBase, cuentas.nommacta Titulo, tmpfaclin.cta Cliente, tmpfaclin.cliente Titulo, tmpfaclin.numserie Serie, tmpfaclin.numfac Factura, tmpfaclin.Fecha, tmpfaclin.iva Iva, tmpfaclin.Imponible "
        Sql = Sql & "FROM  tmpfaclin inner join cuentas on tmpfaclin.ctabase = cuentas.codmacta "
        Sql = Sql & " WHERE  tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
        If Check1(2).Value = 1 Then
            Sql = Sql & " ORDER BY tmpfaclin.ctabase, tmpfaclin.codigo"
        Else
            Sql = Sql & " order by tmpfaclin.ctabase, tmpfaclin.cta"
        End If
    Else
        If Check1(1).Value = 1 Then
            Sql = "Select  tmpfaclin.ctabase,  tmpfaclin.cta,  sum(tmpfaclin.imponibleant) Anterior,  sum(tmpfaclin.imponible)  importe "
            Sql = Sql & "FROM  tmpfaclin inner join cuentas on tmpfaclin.ctabase = cuentas.codmacta "
            Sql = Sql & " WHERE  tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
            Sql = Sql & " group by 1,2"
            Sql = Sql & " order by 1,2"
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Total = DevuelveValor("select sum(imponible) from tmpfaclin where codusu = " & vUsu.Codigo & " and ctabase = " & DBSet(Rs!CtaBase, "T"))
                TotalAnt = DevuelveValor("select sum(imponibleant) from tmpfaclin where codusu = " & vUsu.Codigo & " and ctabase = " & DBSet(Rs!CtaBase, "T"))
            
            
                If Total <> 0 Then
                    Sql2 = "update tmpfaclin set iva = " & DBSet(Round(DBLet(Rs!Importe) * 100 / Total, 2), "N")
                    If TotalAnt <> 0 Then
                        Sql2 = Sql2 & ", porcrec = " & DBSet(Round(DBLet(Rs!Anterior) * 100 / TotalAnt, 2), "N")
                    Else
                        Sql2 = Sql2 & ", porcrec = 0"
                    End If
                Else
                    Sql2 = "update tmpfaclin set iva = 0 "
                    If TotalAnt <> 0 Then
                        Sql2 = Sql2 & ", porcrec = " & DBSet(Round(DBLet(Rs!Anterior) * 100 / TotalAnt, 2), "N")
                    Else
                        Sql2 = Sql2 & ", porcrec = 0"
                    End If
                End If
                
                Sql2 = Sql2 & " where codusu = " & DBSet(vUsu.Codigo, "N") & " and ctabase = " & DBSet(Rs!CtaBase, "T") & " and cta = " & DBSet(Rs!Cta, "T")
                
                Conn.Execute Sql2
            
                Rs.MoveNext
            Wend
            Set Rs = Nothing
            
            Sql = "Select  tmpfaclin.ctabase CtaBase, cuentas.nommacta Titulo, tmpfaclin.cta Cliente, ccc.nommacta TituloCli, sum(tmpfaclin.imponibleant) ImporteAnt, porcrec PorcAnt, sum(tmpfaclin.imponible) Importe, iva Porc"
            Sql = Sql & " FROM  (tmpfaclin inner join cuentas on tmpfaclin.ctabase = cuentas.codmacta) inner join cuentas ccc on tmpfaclin.cta = ccc.codmacta "
            Sql = Sql & " WHERE  tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
            Sql = Sql & " group by 1,2,3,4 "
            If Check1(2).Value = 1 Then
                Sql = Sql & " ORDER BY tmpfaclin.ctabase, tmpfaclin.codigo"
            Else
                Sql = Sql & " order by tmpfaclin.ctabase, tmpfaclin.cta"
            End If
            
            
        Else
            Sql = "Select  tmpfaclin.ctabase,  tmpfaclin.cta,  sum(tmpfaclin.imponible) Importe "
            Sql = Sql & " FROM  tmpfaclin inner join cuentas on tmpfaclin.ctabase = cuentas.codmacta "
            Sql = Sql & " WHERE  tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
            Sql = Sql & " group by 1,2"
            Sql = Sql & " ORDER BY 1,2"
        
        
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Total = DevuelveValor("select sum(imponible) from tmpfaclin where codusu = " & vUsu.Codigo & " and ctabase = " & DBSet(Rs!CtaBase, "T"))
                
                If Total <> 0 Then
                    Sql2 = "update tmpfaclin set iva = " & DBSet(Round(DBLet(Rs!Importe) * 100 / Total, 2), "N")
                Else
                    Sql2 = "update tmpfaclin set iva = 0 "
                End If
                
                Sql2 = Sql2 & " where codusu = " & DBSet(vUsu.Codigo, "N") & " and ctabase = " & DBSet(Rs!CtaBase, "T") & " and cta = " & DBSet(Rs!Cta, "T")
                
                Conn.Execute Sql2
            
                Rs.MoveNext
            Wend
            Set Rs = Nothing
            
            
            Sql = "Select  tmpfaclin.ctabase CtaBase, cuentas.nommacta Titulo, tmpfaclin.cta Cliente, ccc.nommacta TituloCli, sum(tmpfaclin.imponible) Importe, iva Porc "
            Sql = Sql & " FROM  (tmpfaclin inner join cuentas on tmpfaclin.ctabase = cuentas.codmacta) inner join cuentas ccc on tmpfaclin.cta = ccc.codmacta "
            Sql = Sql & " WHERE  tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
            Sql = Sql & " group by 1,2,3,4,6 "
            If Check1(2).Value = 1 Then
                Sql = Sql & " ORDER BY tmpfaclin.ctabase, tmpfaclin.codigo"
            Else
                Sql = Sql & " order by tmpfaclin.ctabase, tmpfaclin.cta"
            End If
        
        
        End If
    End If
    
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0403-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "FacturasCliBase.rpt"

    cadParam = cadParam & "pDesglose=" & Check1(0).Value & "|"
    numParam = numParam + 1
    cadParam = cadParam & "pComparativo=" & Check1(1).Value & "|"
    numParam = numParam + 1
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 16
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Function CargarTemporal() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    Set Rs = New ADODB.Recordset

    'Preparando tablas informe
    Sql = "DELETE from tmpfaclin where codusu =" & vUsu.Codigo
    Conn.Execute Sql
    
    If Check1(1).Value = 0 Then ' si no es comparativo

    
        If Check1(0).Value = 1 Then ' desglosar por cuenta
            Sql = "insert into tmpfaclin (codusu,ctabase,cta,cliente,numserie,Numfac,Fecha,iva,imponible) "
            Sql = Sql & "select " & vUsu.Codigo & ", factcli_lineas.codmacta, factcli.codmacta, factcli.nommacta, factcli.numserie, factcli.numfactu, factcli.fecfactu, factcli_lineas.porciva, factcli_lineas.baseimpo "
            Sql = Sql & " from " & tabla
            If cadselect <> "" Then Sql = Sql & " where " & cadselect
        Else
            Sql = "insert into tmpfaclin (codusu,ctabase,cta,cliente,imponible) "
            Sql = Sql & "select " & vUsu.Codigo & ", factcli_lineas.codmacta,factcli.codmacta, factcli.nommacta, sum(coalesce(factcli_lineas.baseimpo,0))  "
            Sql = Sql & " from " & tabla
            If cadselect <> "" Then Sql = Sql & " where " & cadselect
            Sql = Sql & " group by 1,2,3,4 "
        End If
    Else ' comparativo
        Sql = "insert into tmpfaclin (codusu,ctabase,cta,cliente,imponible, imponibleant) "
        Sql = Sql & " select c1, c2, c3, c4, sum(importe1), sum(importe2) "
        Sql = Sql & " from ("
        Sql = Sql & "select " & vUsu.Codigo & " c1, factcli_lineas.codmacta c2, factcli.codmacta c3, factcli.nommacta c4, sum(coalesce(factcli_lineas.baseimpo,0)) importe1, 0 importe2 "
        Sql = Sql & " from " & tabla
        If cadselect <> "" Then Sql = Sql & " where " & cadselect
        Sql = Sql & " group by 1,2,3,4 "
        Sql = Sql & " union "
        Sql = Sql & "select " & vUsu.Codigo & " c1, factcli_lineas.codmacta c2, factcli.codmacta c3, factcli.nommacta c4, 0 importe1, sum(coalesce(factcli_lineas.baseimpo,0)) importe2 "
        Sql = Sql & " from " & tabla
        Sql = Sql & " where factcli.fecfactu between date_sub(" & DBSet(txtFecha(0).Text, "F") & ", interval 1 year) and date_sub(" & DBSet(txtFecha(1).Text, "F") & ",  interval 1 year)"
        If CadSelect1 <> "" Then Sql = Sql & " and " & CadSelect1
        Sql = Sql & " group by 1,2,3,4 "
        Sql = Sql & ") aaaaaa "
        Sql = Sql & " group by 1,2,3,4 "
    End If
    
    Conn.Execute Sql
    
    
    If Check1(2).Value = 1 Then
            If Check1(0).Value = 1 Then ' desglosar por cuenta
                ' actualizamos el tmplinfac.codigo, que es el orden por importes
                Sql = " update tmpfaclin ddd, " & _
                        "( " & _
                        "select ctabase, cta, imponible, numserie, fecha, numfac, @rownum:=@rownum + 1 AS rownum " & _
                        "    from tmpfaclin, (SELECT @rownum:=0) r " & _
                        "   where codusu = " & DBSet(vUsu.Codigo, "N") & _
                        "  order by 1,2,3 " & _
                        ") fff " & _
                        " set ddd.Codigo = fff.rownum " & _
                        " where ddd.codusu = " & DBSet(vUsu.Codigo, "N") & "  and ddd.ctabase = fff.ctabase and ddd.Cta = fff.Cta and ddd.NumSerie = fff.NumSerie and ddd.Fecha = fff.Fecha and ddd.NumFac = fff.NumFac"
            
                Conn.Execute Sql
            
            Else
                ' actualizamos el tmplinfac.codigo, que indica que el orden es por cantidades
                Sql = "update tmpfaclin, " & _
                      " (" & _
                      "     select ctabase, cta, @rownum:=@rownum+1 AS rownum " & _
                      "     From " & _
                      "     ( " & _
                      "     select ctabase, cta, sum(imponible) " & _
                      "      From tmpfaclin " & _
                      "     Where codusu = " & DBSet(vUsu.Codigo, "N") & _
                      "     group by ctabase, cta " & _
                      "     order by 1, 3 desc " & _
                      "     ) aaaa, (SELECT @rownum:=0) r " & _
                      " ) ZZZZ " & _
                      " Set tmpfaclin.Codigo = zzzz.rownum " & _
                      " Where codusu = " & DBSet(vUsu.Codigo, "N") & " and tmpfaclin.ctabase = zzzz.ctabase And tmpfaclin.Cta = zzzz.Cta "
                Conn.Execute Sql
            End If
    End If
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
End Function


Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim i As Integer


    MontaSQL = False
    
    If Not PonerDesdeHasta("factcli.codmacta", "CTA", Me.txtCuentas(2), Me.txtNCuentas(2), Me.txtCuentas(3), Me.txtNCuentas(3), "pDHCuentas=""") Then Exit Function
    If Not PonerDesdeHasta("factcli_lineas.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCtaBase=""") Then Exit Function
    
    CadSelect1 = cadselect
    
    If Not PonerDesdeHasta("factcli.FecFactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
            
            
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
    If Not CargarTemporal Then Exit Function
    
    cadFormula = "{tmpfaclin.codusu} = " & vUsu.Codigo
    
            
    MontaSQL = True
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
    
    If Check1(1).Value = 1 Then
      If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
          MsgBox "Si marca la opcion de comparativo  debe indicar las fechas", vbExclamation
          Exit Function
      End If
      
      If DateDiff("d", CDate(txtFecha(0).Text), CDate(txtFecha(1).Text)) > 366 Then
          MsgBox "Si marca la opcion de comparativo , el periodo no puede ser superior a un año", vbExclamation
          Exit Function
      End If
        
    End If
    
  
    DatosOK = True


End Function

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


