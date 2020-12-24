VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfConsBalan 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   7200
      TabIndex        =   33
      Top             =   2040
      Width           =   4695
      Begin MSComctlLib.ListView ListView1 
         Height          =   2580
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3840
         Picture         =   "frmInfConsBalan.frx":0000
         ToolTipText     =   "Quitar al Debe"
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4200
         Picture         =   "frmInfConsBalan.frx":014A
         ToolTipText     =   "Puntear al Debe"
         Top             =   240
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
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1110
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
      Height          =   1995
      Left            =   7200
      TabIndex        =   20
      Top             =   0
      Width           =   4695
      Begin VB.CheckBox chk1 
         Caption         =   "Incluir saldo de la 473 en la 470"
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
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   4155
      End
      Begin VB.CheckBox chk2 
         Caption         =   "Incluir saldo de grupos 6 y 7 en 129"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   4035
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   27
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
      Height          =   2925
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
         Index           =   2
         ItemData        =   "frmInfConsBalan.frx":0294
         Left            =   2910
         List            =   "frmInfConsBalan.frx":0296
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1890
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   120
         TabIndex        =   31
         Top             =   2250
         Width           =   4425
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
            Index           =   3
            ItemData        =   "frmInfConsBalan.frx":0298
            Left            =   2790
            List            =   "frmInfConsBalan.frx":029A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   150
            Width           =   1215
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
            ItemData        =   "frmInfConsBalan.frx":029C
            Left            =   1110
            List            =   "frmInfConsBalan.frx":029E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   150
            Width           =   1575
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
            Index           =   0
            Left            =   210
            TabIndex        =   32
            Top             =   210
            Width           =   615
         End
      End
      Begin VB.CheckBox chkBalPerCompa 
         Caption         =   "Comparativo"
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
         Left            =   5040
         TabIndex        =   3
         Top             =   1920
         Width           =   1545
      End
      Begin VB.TextBox txtNBalan 
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
         ItemData        =   "frmInfConsBalan.frx":02A0
         Left            =   1230
         List            =   "frmInfConsBalan.frx":02A2
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1890
         Width           =   1575
      End
      Begin VB.TextBox txtBalan 
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
      Begin VB.Label lblInd 
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
         Left            =   5160
         TabIndex        =   36
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Image imgBalan 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   1050
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
         Index           =   4
         Left            =   330
         TabIndex        =   26
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "NºInforme"
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
         TabIndex        =   25
         Top             =   690
         Width           =   1200
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
         TabIndex        =   24
         Top             =   1620
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
      Left            =   10560
      TabIndex        =   10
      Top             =   5790
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
      Left            =   8970
      TabIndex        =   8
      Top             =   5790
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
      Top             =   5730
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
      Top             =   2940
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
      Left            =   3960
      TabIndex        =   28
      Top             =   5760
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
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
      Left            =   10560
      TabIndex        =   29
      Top             =   5790
      Width           =   1215
   End
End
Attribute VB_Name = "frmInfConsBalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 308

Public Opcion As Byte

    'pcion = 0         Me.Caption = "Balance de Situación"
    '1:        Me.Caption = "Cuenta de Pérdidas y Ganancias"



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



'Private WithEvents frmDia As frmTiposDiario
Private WithEvents frmC As frmBasico
Attribute frmC.VB_VarHelpID = -1
'Private WithEvents frmCon  As frmConceptos
'Private frmCtas As frmCtasAgrupadas

Private Sql As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim Rs As ADODB.Recordset

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean



Dim HanPulsadoSalir As Boolean
Dim vIdPrograma As Integer

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




Private Sub chk1_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub chkBalPerCompa_Click()
    Frame2.visible = Me.chkBalPerCompa.Value = 1
    Frame2.Enabled = Me.chkBalPerCompa.Value = 1
End Sub

Private Sub chkBalPerCompa_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim Contabilidades As String

    If Not DatosOK Then Exit Sub
    
    PulsadoCancelar = False
    Me.cmdCancelarAccion.visible = True
    Me.cmdCancelarAccion.Enabled = True
    
    Me.cmdCancelar.visible = False
    Me.cmdCancelar.Enabled = False
        
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    Saldo473en470 = False
    Saldo6y7en129 = False
    If Opcion = 0 Then
        
        Saldo473en470 = (chk1.Value = 1)
        Saldo6y7en129 = (chk2.Value = 1)
    
    
        If Saldo473en470 Then
            'Deberiamos indicar si esta configurado para leer de la 470
            cad = "codmacta = '4' or codmacta='47' or codmacta like '473%' AND numbalan"
            RC = DevuelveDesdeBD("concat(pasivo,' ',codigo,': ',codmacta)", "balances_ctas", cad, txtBalan(0).Text & " ORDER BY codmacta")
            If RC <> "" Then MsgBox "La cuenta 470 ha sido configurada en el balance: " & RC, vbExclamation
                
        End If
    End If

    



    Screen.MousePointer = vbHourglass
    
    Contabilidades = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then Contabilidades = Contabilidades & ListView1.ListItems(I) & "|"
    Next I
    
    
    I = -1
    If chkBalPerCompa.Value = 1 Then
        I = Val(cmbFecha(1).ListIndex)
        I = I + 1
        If I = 0 Then I = -1
    End If
    
    
    
    
    GeneraDatosBalanceConfigurable_ CInt(txtBalan(0).Text), Me.cmbFecha(0).ListIndex + 1, CInt(cmbFecha(2).Text), I, Val(cmbFecha(3).Text), False, Contabilidades, pb2, False, lblInd

'



    Me.cmdCancelarAccion.visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.visible = True
    Me.cmdCancelar.Enabled = True

    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmpimpbalan", "codusu=" & vUsu.Codigo) Then Exit Sub
    
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
    lblInd.Caption = ""
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    If Me.cmdCancelarAccion.visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub Form_Activate()


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

    Me.Icon = frmppal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
        
        
    If Opcion = 0 Then
        Me.Caption = "Balance de Situación"
        vIdPrograma = 308
    Else
        Me.Caption = "Cuenta de Pérdidas y Ganancias"
        vIdPrograma = 309
    End If
    Me.Caption = Me.Caption & " CONSOLIDADO"
    
    ' solo se muestran si es balance de situacion
    chk1.visible = (Opcion = 0)
    chk1.Enabled = (Opcion = 0)
    chk2.visible = (Opcion = 0)
    chk2.Enabled = (Opcion = 0)
    

    Me.imgBalan(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    PrimeraVez = True
     
   
    CargarListViewEmpresas
    CargarComboFecha
    
    
    'Fecha inicial
    cmbFecha(0).ListIndex = Month(vParam.fechafin) - 1
    cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1

'    txtAno(0).Text = Year(vParam.fechafin)
'    txtAno(1).Text = Year(vParam.fechafin) - 1
    cmbFecha(2).Text = Year(vParam.fechafin)
    cmbFecha(3).Text = CInt(cmbFecha(2).Text) - 1
   
    PosicionarCombo cmbFecha(2), Year(vParam.fechafin)
    PosicionarCombo cmbFecha(3), CInt(cmbFecha(2).Text) - 1
    
   
    PonerBalancePredeterminado
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    Frame2.visible = False
    Frame2.Enabled = False
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.visible = False
    
    lblInd.Caption = ""
    
End Sub

Private Sub PonerBalancePredeterminado()

    'El balance de P y G tiene el campo Perdidas=1
    Sql = "Select * from balances where predeterminado = 1 AND perdidas =" & Opcion
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Me.txtBalan(0).Text = Rs.Fields(0)
        txtNBalan(0).Text = Rs.Fields(1)
    End If
    Rs.Close
    Set Rs = Nothing

End Sub





Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtBalan(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNBalan(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub




Private Sub ImgBalan_Click(Index As Integer)
Dim cWhere As String

    If Opcion = 0 Then
        cWhere = "numbalan < 50 and perdidas = 0"
    Else
        cWhere = "numbalan < 50 and perdidas = 1"
    End If
    
    Set frmC = New frmBasico
    AyudaBalances frmC, 0, , cWhere
    Set frmC = Nothing

    PonFoco txtBalan(Index)

End Sub




Private Sub imgCheck_Click(Index As Integer)
    For I = 1 To Me.ListView1.ListItems.Count
        Me.ListView1.ListItems(I).Checked = Index = 1
    Next I
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
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & vIdPrograma & ".html"
    End Select
End Sub

Private Sub txtBalan_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgBalan", Index
    End If
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgBalan"
        ImgBalan_Click Indice
    End Select
    
End Sub


Private Sub txtBalan_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtBalan(Index).Text = Trim(txtBalan(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtBalan(Index).Text = Trim(txtBalan(Index).Text)
    If txtBalan(Index).Text = "" Then
        txtNBalan(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtBalan(Index).Text) Then
        If InStr(1, txtBalan(Index).Text, "+") = 0 Then MsgBox "El Balance debe ser numérico: " & txtBalan(Index).Text, vbExclamation
        txtBalan(Index).Text = ""
        txtNBalan(Index).Text = ""
        Exit Sub
    End If

    If Opcion = 0 Then
        If EsPyG(txtBalan(Index)) Then
            MsgBox "Este código corresponde a un balance de Pérdidas y Ganancias. Reintroduzca.", vbExclamation
            txtBalan(Index).Text = ""
            txtNBalan(Index).Text = ""
            PonFoco txtBalan(Index)
            Exit Sub
        End If
    Else
        If Not EsPyG(txtBalan(Index)) Then
            MsgBox "Este código no corresponde a un balance de Pérdidas y Ganancias. Reintroduzca.", vbExclamation
            txtBalan(Index).Text = ""
            txtNBalan(Index).Text = ""
            PonFoco txtBalan(Index)
            Exit Sub
        End If
    
    End If
    
    txtNBalan(Index).Text = DevuelveDesdeBD("nombalan", "balances", "numbalan", txtBalan(Index), "N")


End Sub

Private Function EsPyG(Balance As Integer) As Boolean
Dim Sql As String

    EsPyG = DevuelveValor("select perdidas from balances where numbalan = " & DBSet(Balance, "N")) = 1


End Function

Private Sub AccionesCSV()
       
    
    Sql = "select descripcion 'Nº Cuentas',linea 'Debe (Haber)',"
    Sql = Sql & " importe1 '" & cmbFecha(2).Text & "'"
    If Me.chkBalPerCompa.Value = 1 Then Sql = Sql & " , importe2 '" & cmbFecha(3).Text & "'"
    
    Sql = Sql & " from tmpimpbalance where codusu = " & vUsu.Codigo & " order by pasivo,codigo"
    
      
            

        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String
Dim ConTexto As Byte
Dim optExportar As Integer

    cadParam = cadParam & "pTipo=" & Tipo & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pFecha=""" & Format(Now, "dd/mm/yyyy") & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & cmbFecha(2).Text & """|"
    numParam = numParam + 1
    
    
    cadNomRPT = ""
    ConTexto = 0
    For I = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(I).Checked Then
            UltimoNivel = UltimoNivel + 1
            
            cadNomRPT = cadNomRPT & "  -  " & ListView1.ListItems(I).SubItems(1)
            If Me.ListView1.ListItems(I).Tag = vEmpresa.codempre Then ConTexto = 1
    
        End If
    Next
    
    If UltimoNivel = 1 And ConTexto = 1 Then
        cadNomRPT = "" 'La empresa seleccionada(solo una) es la que estoy
    Else
        cadNomRPT = Trim(Mid(cadNomRPT, 4))
    End If
    cadParam = cadParam & "pdh=""" & cadNomRPT & """|"
    numParam = numParam + 1
    
        
    vMostrarTree = False
    conSubRPT = False
        

    
    ConTexto = DevuelveValor("select aparece from balances where numbalan = " & DBSet(txtBalan(0).Text, "N"))
        
            
    

    'If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    cadNomRPT = "balanceConso.rpt"
    If Me.chkBalPerCompa.Value = 1 Then cadNomRPT = "balanceConsoComp.rpt"
    cadFormula = "{tmpimpbalan.codusu}=" & vUsu.Codigo

'+++
    'Para saber k informe abriresmos
    CONT = 1
    RC = 1 'Perdidas y ganancias
    Set Rs = New ADODB.Recordset
    Sql = "Select * from balances where numbalan=" & Me.txtBalan(0).Text
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then

            If DBLet(Rs!Aparece, "N") = 0 Then
                CONT = 3
            Else
                CONT = 1
            End If

        RC = Rs!PERDIDAS
    End If
    Rs.Close
    Set Rs = Nothing
        
        
    'Si es comarativo o no
    If Me.chkBalPerCompa.Value = 1 Then CONT = CONT + 1
        
    'Textos
    If RC = 1 Then
        optExportar = 54
    Else
        optExportar = 25
    End If
    RC = "perdidasyganancias= " & RC & "|"
          
    Sql = RC & "FechaImp= """ & Format(Now, "dd/mm/yyyy") & """|"
    Sql = Sql & "Titulo= """ & Me.txtNBalan(0).Text & """|"
    'PGC 2008 SOlo pone el año, NO el mes
    If vParam.NuevoPlanContable Then
        RC = ""
    Else
        RC = cmbFecha(0).List(cmbFecha(0).ListIndex)
    End If
    
    'Agosto 2020
    'Si es años aprtidos, pintaresmos como año el de INICIO de ejercicio
    I = 0
    If Month(vParam.fechaini) > 1 Then
        If Month(vParam.fechaini) > (cmbFecha(0).ListIndex + 1) Then I = 1
    End If
    
    
    
    
    RC = RC & " " & Val(cmbFecha(2).Text) - I 'txtAno(0).Text
    RC = "fec1= """ & Trim(RC) & """|"
    Sql = Sql & RC
    
    
    If Me.chkBalPerCompa.Value = 1 Then
            'PGC 2008 SOlo pone el año, NO el mes
            If vParam.NuevoPlanContable Then
                RC = ""
                
                If Month(vParam.fechafin) <> Val(cmbFecha(0).ListIndex + 1) Then RC = Mid(cmbFecha(0).Text, 1, 3)
                Sql = Sql & "vMes= """ & RC & """|"
                RC = ""
            Else
                RC = cmbFecha(1).List(cmbFecha(1).ListIndex)
            End If
            RC = RC & " " & Val(cmbFecha(3).Text) - I 'txtAno(1).Text
            RC = "Fec2= """ & RC & """|"
            Sql = Sql & RC
            
            
            RC = ""
            If Month(vParam.fechafin) <> Val(cmbFecha(1).ListIndex + 1) Then RC = UCase(Mid(cmbFecha(1).Text, 1, 1)) & Mid(cmbFecha(1).Text, 2, 2)
            RC = "vMes2= """ & RC & """|"
            Sql = Sql & RC
            
            

    Else
        'Pong el nombre del mes
        RC = ""
        If Month(vParam.fechafin) <> Val(cmbFecha(0).ListIndex + 1) Then RC = UCase(Mid(cmbFecha(0).Text, 1, 1)) & Mid(cmbFecha(0).Text, 2, 2)
        RC = "vMes= """ & RC & """|"
        Sql = Sql & RC
    End If
    Sql = Sql & "Titulo= """ & Me.txtNBalan(0).Text & """|"


    cadParam = cadParam & Sql
    numParam = numParam + 4






    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text, True) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook optExportar
        
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
    
    If Me.txtBalan(0).Text = "" Then
        MsgBox "Número de balance incorrecto", vbExclamation
        Exit Function
    End If
    
    'Año 1
'    If txtAno(0).Text = "" Then
'        MsgBox "Año no puede estar en blanco", vbExclamation
'        Exit Function
'    End If
'
'    If Val(txtAno(0).Text) < 1900 Then
'        MsgBox "No se permiten años anteriores a 1900", vbExclamation
'        Exit Function
'    End If
    If cmbFecha(2).ListIndex < 0 Then
        MsgBox "Introduce la fecha(año) de consulta", vbExclamation
        Exit Function
    End If

    If chkBalPerCompa.Value = 1 Then
        If cmbFecha(3).ListIndex < 0 Then
            MsgBox "Introduce la fecha(año) de consulta", vbExclamation
            Exit Function
        End If
    End If
    
    cad = ""
    For I = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(I).Checked Then cad = cad & "X"
    Next I

    If cad = "" Then
        MsgBoxA "Seleccione almenos una empresa", vbExclamation
        Exit Function
    End If
    
    DatosOK = True

End Function

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

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub CargarComboFecha()
Dim J As Integer

    QueCombosFechaCargar "0|1|"
    
    cmbFecha(2).Clear
    cmbFecha(3).Clear
    
    J = Year(vParam.fechafin) + 1 - 2000
    For I = 1 To J
        cmbFecha(2).AddItem "20" & Format(I, "00")
        cmbFecha(3).AddItem "20" & Format(I, "00")
    Next I

End Sub


Private Sub CargarListViewEmpresas()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim Prohibidas As String
Dim IT
Dim Aux As String
    
    On Error GoTo ECargarList

    'Los encabezados
    ListView1.ColumnHeaders.Clear

    ListView1.ColumnHeaders.Add , , "Código", 600
    ListView1.ColumnHeaders.Add , , "Empresa", 3200
    
    
    


    Set Rs = New ADODB.Recordset

    Prohibidas = DevuelveProhibidas
    
    ListView1.ListItems.Clear
    Aux = "Select * from usuarios.empresasariconta order by codempre"
    
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
       '  Aux = "ariconta" & Rs!codempre & ".parametros"
       ' Aux = DevuelveDesdeBD("esmultiseccion", Aux, "1", "1")
       ' If Aux = "0" Then
       '     Aux = "N"
       ' Else
            Aux = "|" & Rs!codempre & "|"
            If InStr(1, Prohibidas, Aux) = 0 Then Aux = ""
       ' End If
        If Aux = "" Then
            Set IT = ListView1.ListItems.Add
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



