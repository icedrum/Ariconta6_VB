VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInmoSimu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   5925
      Left            =   7110
      TabIndex        =   19
      Top             =   0
      Width           =   4455
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
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "imgConcepto"
         Top             =   750
         Width           =   1305
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3870
         TabIndex        =   24
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
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2610
         Picture         =   "frmInmoSimu.frx":0000
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Amortización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   25
         Top             =   780
         Width           =   2325
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
      Height          =   2985
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNConcepto 
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
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox txtConcepto 
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
         Top             =   1920
         Width           =   1305
      End
      Begin VB.TextBox txtNConcepto 
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
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txtConcepto 
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
         Top             =   2400
         Width           =   1305
      End
      Begin VB.TextBox txtNConcepto 
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
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1140
         Width           =   4125
      End
      Begin VB.TextBox txtNConcepto 
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
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   4125
      End
      Begin VB.TextBox txtConcepto 
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
         Top             =   720
         Width           =   1305
      End
      Begin VB.TextBox txtConcepto 
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
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Seccion"
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
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   3
         Left            =   960
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   2
         Left            =   960
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto"
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
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1260
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   1
         Left            =   960
         Top             =   1140
         Width           =   255
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   0
         Left            =   960
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   18
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   17
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   30
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   31
         Top             =   2400
         Width           =   735
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
      Left            =   10350
      TabIndex        =   6
      Top             =   6120
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
      Left            =   8790
      TabIndex        =   5
      Top             =   6120
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
      TabIndex        =   7
      Top             =   6120
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
      Height          =   2805
      Left            =   120
      TabIndex        =   8
      Top             =   3120
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label lblInd 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   33
      Top             =   6240
      Width           =   5460
   End
End
Attribute VB_Name = "frmInmoSimu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 508

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


Private WithEvents frmCon As frmInmoConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmSec As frmInmoSeccion
Attribute frmSec.VB_VarHelpID = -1
'Private WithEvents frmCC As frmCCCentroCoste
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer


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
Dim tabla As String
Dim B As Boolean
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    lblInd.Caption = "Iniciando"
    lblInd.Refresh
    B = HazSimulacion(cadselect, txtFecha(0).Text, 0, lblInd)
    lblInd.Caption = ""
    Screen.MousePointer = vbDefault
    If Not B Then Exit Sub
    
    cadselect = "tmpsimulainmo.codusu = " & vUsu.Codigo
    cadFormula = "{tmpsimulainmo.codusu}=" & vUsu.Codigo
    
    If Not HayRegParaInforme("tmpsimulainmo", cadselect) Then Exit Sub
    
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
    lblInd.Caption = ""
    'Otras opciones
    Me.Caption = "Simulación de Amortización"

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With


    For I = 0 To 3
        Me.imgConcepto(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
     
    txtFecha(0).Text = SugerirFechaNuevo
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub imgConcepto_Click(Index As Integer)
    
    Sql = ""
    AbiertoOtroFormEnListado = True
    If Index < 2 Then
        Set frmCon = New frmInmoConceptos
        frmCon.DatosADevolverBusqueda = True
        frmCon.Show vbModal
        Set frmCon = Nothing
    Else
        Set frmSec = frmInmoSeccion
        frmSec.DatosADevolverBusqueda = True
        frmSec.Show vbModal
        Set frmSec = Nothing
    
        
    End If
    If Sql <> "" Then
        Me.txtConcepto(Index).Text = RecuperaValor(Sql, 1)
        Me.txtNConcepto(Index).Text = RecuperaValor(Sql, 2)
    Else
        QuitarPulsacionMas Me.txtConcepto(Index)
    End If
    
    PonFoco Me.txtConcepto(Index)
    AbiertoOtroFormEnListado = False

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
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

Private Sub txtConcepto_GotFocus(Index As Integer)
    ConseguirFoco txtConcepto(Index), 3
End Sub


Private Sub txtConcepto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtConcepto(Index).Tag, Index
    End If
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgConcepto"
        imgConcepto_Click Indice
    End Select
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada

    Select Case Index
        Case 0, 1 'Tipos de concepto de inmovilizado
            txtNConcepto(Index).Text = DevuelveDesdeBD("nomconam", "inmovcon", "codconam", txtConcepto(Index), "N")
            If txtConcepto(Index).Text <> "" Then txtConcepto(Index).Text = Format(txtConcepto(Index).Text, "0000")
        Case 2, 3
            txtNConcepto(Index).Text = DevuelveDesdeBD("nomsecin", "inmovseccion", "codsecin", txtConcepto(Index), "N")
            If txtConcepto(Index).Text <> "" Then txtConcepto(Index).Text = Format(txtConcepto(Index).Text, "0000")
    End Select

End Sub



Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
   ' KEYpressGnral KeyAscii
End Sub



Private Sub AccionesCSV()

    Sql = " Select codsecin,nomsecin,tmpsimulainmo.conconam Concepto,nomconam Descripción,tmpsimulainmo.codinmov Elemento,tmpsimulainmo.nominmov Descripción"
    Sql = Sql & " ,tmpsimulainmo.fechaadq FechaAdquisicion,tmpsimulainmo.valoradq ValorAdquisicion,tmpsimulainmo.amortacu AmortAcumulada"
    Sql = Sql & " ,totalamor TotalAmortizacion, tmpsimulainmo.valoradq - tmpsimulainmo.amortacu - totalamor Pendiente"
    Sql = Sql & " FROM tmpsimulainmo  left join inmovele on tmpsimulainmo.codinmov=inmovele.codinmov"
    Sql = Sql & " left join inmovseccion on codsecin =seccion"
    
    If cadselect <> "" Then Sql = Sql & " WHERE " & cadselect
    
    Sql = Sql & " ORDER BY 1,3,5"
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
Dim CADENA As String

    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "0508-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    If UCase(Right(nomDocu, 4)) = "S.RPT" Then vMostrarTree = True
    If vParam.autocoste Then
        cadParam = cadParam & "pAnalitica=1|"
        numParam = numParam + 1
    End If

    cadParam = cadParam & "pFecha=""" & txtFecha(0).Text & """|"
    numParam = numParam + 1


    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 63
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim Situacion As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("inmovele.conconam", "COI", Me.txtConcepto(0), txtNConcepto(0), Me.txtConcepto(1), txtNConcepto(1), "pDHConcepto=""") Then Exit Function
    If Not PonerDesdeHasta("inmovele.seccion", "CCI", Me.txtConcepto(2), Me.txtConcepto(2), Me.txtConcepto(3), Me.txtConcepto(3), "pDHSecci=""") Then Exit Function
    
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
            
    MontaSQL = True
End Function



Private Function DatosOK() As Boolean
Dim I As Integer
Dim CADENA As String

    DatosOK = False
    
    If Me.txtFecha(0).Text = "" Then
        MsgBox "Inserte la fecha de la simulación.", vbExclamation
        Exit Function
    End If
    If Me.Tag <> "" Then
        If CDate(Me.txtFecha(0).Text) <= CDate(Me.Tag) Then
            MsgBox "Fecha no puede ser menor que la última fecha de amortización: " & Me.Tag, vbExclamation
            Exit Function
        End If
    End If
    
    DatosOK = True

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
    End If
End Sub


Private Function SugerirFechaNuevo() As String
Dim RC As String
    RC = "tipoamor"
    cad = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", "1", "N", RC)

    If cad <> "" Then
        Me.Tag = cad   'Ultima actualizacion
        Select Case Val(RC)
        Case 2
            'Semestral
            I = 6
            'Siempre es la ultima fecha de mes
        Case 3
            'Trimestral
            I = 3
        Case 4
            'Mensual
            I = 1
        Case Else
            'Anual
            I = 12
        End Select
        RC = PonFecha
    Else
        cad = "01/01/1991"
        RC = Format(Now, "dd/mm/yyyy")
    End If
    SugerirFechaNuevo = Format(RC, "dd/mm/yyyy")
    
End Function


Private Function PonFecha() As Date
Dim d As Date
'Dada la fecha en Cad y los meses k tengo k sumar
'Pongo la fecha
d = DateAdd("m", I, CDate(cad))
Select Case Month(d)
Case 2
    If ((Year(d) - 2000) Mod 4) = 0 Then
        I = 29
    Else
        I = 28
    End If
Case 1, 3, 5, 7, 8, 10, 12
    '31
        I = 31
Case Else
    '30
        I = 30
End Select
cad = I & "/" & Month(d) & "/" & Year(d)
PonFecha = CDate(cad)
End Function

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
