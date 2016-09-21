VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfBalances 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11685
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
      Height          =   5595
      Left            =   7110
      TabIndex        =   13
      Top             =   0
      Width           =   4485
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   1920
         TabIndex        =   25
         Top             =   750
         Width           =   1485
      End
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
         Left            =   150
         TabIndex        =   22
         Top             =   1950
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
         Left            =   150
         TabIndex        =   21
         Top             =   2490
         Value           =   1  'Checked
         Width           =   4035
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   20
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
         Caption         =   "Fecha"
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
         TabIndex        =   26
         Top             =   780
         Width           =   690
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   7
         Left            =   1560
         Picture         =   "frmInfBalances.frx":0000
         Top             =   780
         Width           =   240
      End
   End
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selecci�n"
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
      TabIndex        =   12
      Top             =   0
      Width           =   6915
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   120
         TabIndex        =   31
         Top             =   2250
         Width           =   4665
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
            Left            =   3150
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   150
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
            ItemData        =   "frmInfBalances.frx":008B
            Left            =   1110
            List            =   "frmInfBalances.frx":008D
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   150
            Width           =   1935
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
            TabIndex        =   34
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
         Left            =   5130
         TabIndex        =   30
         Top             =   1950
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
         TabIndex        =   29
         Top             =   1050
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
         Height          =   345
         Index           =   0
         Left            =   3270
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1890
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
         Index           =   0
         ItemData        =   "frmInfBalances.frx":008F
         Left            =   1230
         List            =   "frmInfBalances.frx":0091
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1890
         Width           =   1935
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
         TabIndex        =   19
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "N�Informe"
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
         TabIndex        =   18
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Mes / A�o"
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
         TabIndex        =   17
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
      Left            =   10320
      TabIndex        =   3
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
      Left            =   8730
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   4
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
         TabIndex        =   16
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   15
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1830
      TabIndex        =   27
      Top             =   5760
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
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
      Left            =   10320
      TabIndex        =   28
      Top             =   5790
      Width           =   1215
   End
End
Attribute VB_Name = "frmInfBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 308

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
Private WithEvents frmC As frmBasico
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon  As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private frmCtas As frmCtasAgrupadas

Private SQL As String
Dim Cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim Rs As ADODB.Recordset

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean

Public Legalizacion As String   'Datos para la legalizacion

Dim HanPulsadoSalir As Boolean
Dim vIdPrograma As Integer

Public Sub InicializarVbles(A�adireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If A�adireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub




Private Sub chkBalPerCompa_Click()
    Frame2.Visible = Me.chkBalPerCompa.Value = 1
    Frame2.Enabled = Me.chkBalPerCompa.Value = 1
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
    
    Saldo473en470 = False
    Saldo6y7en129 = False
    If Opcion = 0 Then
        
        Saldo473en470 = (chk1.Value = 1)
        Saldo6y7en129 = (chk2.Value = 1)
    
    
        If Saldo473en470 Then
            'Deberiamos indicar si esta configurado para leer de la 470
            Cad = "codmacta = '4' or codmacta='47' or codmacta like '473%' AND numbalan"
            RC = DevuelveDesdeBD("concat(pasivo,' ',codigo,': ',codmacta)", "balances_ctas", Cad, txtBalan(0).Text & " ORDER BY codmacta")
            If RC <> "" Then MsgBox "La cuenta 470 ha sido configurada en el balance: " & RC, vbExclamation
                
        End If
    End If

    



    Screen.MousePointer = vbHourglass
    i = -1
    If chkBalPerCompa.Value = 1 Then
        i = Val(cmbFecha(1).ListIndex)
        i = i + 1
        If i = 0 Then i = -1
    End If
    GeneraDatosBalanceConfigurable CInt(txtBalan(0).Text), Me.cmbFecha(0).ListIndex + 1, CInt(txtAno(0).Text), i, Val(txtAno(1).Text), False, -1, pb2

'



    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
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
    
    If Legalizacion <> "" Then
        CadenaDesdeOtroForm = "OK"
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

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        If Legalizacion <> "" Then
            optTipoSal(2).Value = True
                
            Cad = RecuperaValor(Legalizacion, 4)
            If Val(Cad) = 0 Then
                chkBalPerCompa.Value = 0
            Else
                txtAno(1).Text = Val(txtAno(0).Text) - 1
                cmbFecha(1).ListIndex = cmbFecha(0).ListIndex
                chkBalPerCompa.Value = 1
            End If
            
            cmdAccion_Click (1)
        End If
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

    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    If Opcion = 0 Then
        Me.Caption = "Balance de Situaci�n"
        vIdPrograma = 308
    Else
        Me.Caption = "Cuenta de P�rdidas y Ganancias"
        vIdPrograma = 309
    End If
        
    ' solo se muestran si es balance de situacion
    chk1.Visible = (Opcion = 0)
    chk1.Enabled = (Opcion = 0)
    chk2.Visible = (Opcion = 0)
    chk2.Enabled = (Opcion = 0)
    

    Me.imgBalan(0).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    PrimeraVez = True
     
    'Fecha informe
    txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
    
    CargarComboFecha
    
    
    'Fecha inicial
    cmbFecha(0).ListIndex = Month(vParam.fechafin) - 1
    cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1
    txtAno(0).Text = Year(vParam.fechafin)
    txtAno(1).Text = Year(vParam.fechafin) - 1
   
    PonerBalancePredeterminado
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    Frame2.Visible = False
    Frame2.Enabled = False
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    If Legalizacion <> "" Then
        PonerBalancePredeterminado
        
        txtFecha(7).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            
        txtAno(0).Text = Year(CDate(RecuperaValor(Legalizacion, 3)))     'Fin
        
        cmbFecha(0).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 3))) - 1
    End If
    
End Sub

Private Sub PonerBalancePredeterminado()

    'El balance de P y G tiene el campo Perdidas=1
    SQL = "Select * from balances where predeterminado = 1 AND perdidas =" & Opcion
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub ImgBalan_Click(Index As Integer)
Dim cWhere As String

    If Opcion = 0 Then
        cWhere = "numbalan < 50 and perdidas = 0"
    Else
        cWhere = "numbalan < 50 and perdidas = 1"
    End If
    
    Set frmC = New frmBasico
    AyudaBalances frmC, , cWhere
    Set frmC = Nothing

    PonFoco txtBalan(Index)

End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 7
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
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & vIdPrograma & ".html"
    End Select
End Sub


Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Sub txtBalan_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgBalan", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
    Case "imgBalan"
        ImgBalan_Click indice
    End Select
    
End Sub


Private Sub txtBalan_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
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
        If InStr(1, txtBalan(Index).Text, "+") = 0 Then MsgBox "El Balance debe ser num�rico: " & txtBalan(Index).Text, vbExclamation
        txtBalan(Index).Text = ""
        txtNBalan(Index).Text = ""
        Exit Sub
    End If

    If Opcion = 0 Then
        If EsPyG(txtBalan(Index)) Then
            MsgBox "Este c�digo corresponde a un balance de P�rdidas y Ganancias. Reintroduzca.", vbExclamation
            txtBalan(Index).Text = ""
            txtNBalan(Index).Text = ""
            PonFoco txtBalan(Index)
            Exit Sub
        End If
    Else
        If Not EsPyG(txtBalan(Index)) Then
            MsgBox "Este c�digo no corresponde a un balance de P�rdidas y Ganancias. Reintroduzca.", vbExclamation
            txtBalan(Index).Text = ""
            txtNBalan(Index).Text = ""
            PonFoco txtBalan(Index)
            Exit Sub
        End If
    
    End If
    
    txtNBalan(Index).Text = DevuelveDesdeBD("nombalan", "balances", "numbalan", txtBalan(Index), "N")


End Sub

Private Function EsPyG(Balance As Integer) As Boolean
Dim SQL As String

    EsPyG = DevuelveValor("select perdidas from balances where numbalan = " & DBSet(Balance, "N")) = 1


End Function

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
Dim ConTexto As Byte
            
    cadParam = cadParam & "pTipo=" & Tipo & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pFecha=""" & txtFecha(7).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & txtAno(0).Text & """|"
    numParam = numParam + 1
    
        
    vMostrarTree = False
    conSubRPT = False
        

    
    ConTexto = DevuelveValor("select aparece from balances where numbalan = " & DBSet(txtBalan(0).Text, "N"))
        
            
    indRPT = "0308-"
            
    If ConTexto Then
        If chkBalPerCompa.Value = 0 Then
            indRPT = indRPT & "00"
        Else
            indRPT = indRPT & "01"
        End If
    
    Else
        If chkBalPerCompa.Value = 0 Then
            indRPT = indRPT & "02"
        Else
            indRPT = indRPT & "03"
        End If
    
    End If
        
    
'    Stop
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    cadNomRPT = nomDocu '"balance1a.rpt"

    cadFormula = "{tmpimpbalan.codusu}=" & vUsu.Codigo

'+++
    'Para saber k informe abriresmos
    Cont = 1
    RC = 1 'Perdidas y ganancias
    Set Rs = New ADODB.Recordset
    SQL = "Select * from balances where numbalan=" & Me.txtBalan(0).Text
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then

            If DBLet(Rs!Aparece, "N") = 0 Then
                Cont = 3
            Else
                Cont = 1
            End If

        RC = Rs!perdidas
    End If
    Rs.Close
    Set Rs = Nothing
        
        
    'Si es comarativo o no
    If Me.chkBalPerCompa.Value = 1 Then Cont = Cont + 1
        
    'Textos
    RC = "perdidasyganancias= " & RC & "|"
          
    SQL = RC & "FechaImp= """ & txtFecha(7).Text & """|"
    SQL = SQL & "Titulo= """ & Me.txtNBalan(0).Text & """|"
    'PGC 2008 SOlo pone el a�o, NO el mes
    If vParam.NuevoPlanContable Then
        RC = ""
    Else
        RC = cmbFecha(0).List(cmbFecha(0).ListIndex)
    End If
    RC = RC & " " & txtAno(0).Text
    RC = "fec1= """ & RC & """|"
    SQL = SQL & RC
    
    
    If Me.chkBalPerCompa.Value = 1 Then
            'PGC 2008 SOlo pone el a�o, NO el mes
            If vParam.NuevoPlanContable Then
                RC = ""
            Else
                RC = cmbFecha(1).List(cmbFecha(1).ListIndex)
            End If
            RC = RC & " " & txtAno(1).Text
            RC = "Fec2= """ & RC & """|"
            SQL = SQL & RC
            

    Else
        'Pong el nombre del mes
        RC = UCase(Mid(cmbFecha(0).Text, 1, 1)) & Mid(cmbFecha(0).Text, 2, 2)
        RC = "vMes= """ & RC & """|"
        SQL = SQL & RC
    End If
    SQL = SQL & "Titulo= """ & Me.txtNBalan(0).Text & """|"


    cadParam = cadParam & SQL
    numParam = numParam + 4






    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
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
        
        LanzaFormAyuda "imgFecha", Index
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If Me.txtBalan(0).Text = "" Then
        MsgBox "N�mero de balance incorrecto", vbExclamation
        Exit Function
    End If
    
    'A�o 1
    If txtAno(0).Text = "" Then
        MsgBox "A�o no puede estar en blanco", vbExclamation
        Exit Function
    End If
    
    If Val(txtAno(0).Text) < 1900 Then
        MsgBox "No se permiten a�os anteriores a 1900", vbExclamation
        Exit Function
    End If
    
    If chkBalPerCompa.Value = 1 Then
        If txtAno(1).Text = "" Then
            MsgBox "A�o no puede estar en blanco", vbExclamation
            Exit Function
        End If
        If Val(txtAno(1).Text) < 1900 Then
            MsgBox "No se permiten a�os anteriores a 1900", vbExclamation
            Exit Function
        End If
    End If

    'Fecha informe
    If txtFecha(7).Text = "" Then
        MsgBox "Fecha informe incorrecta.", vbExclamation
        Exit Function
    End If
    

    DatosOK = True

End Function

Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    Cad = RecuperaValor(Lista, L)
    If Cad <> "" Then
        i = Val(Cad)
        With cmbFecha(i)
            .Clear
            For Cont = 1 To 12
                RC = "25/" & Cont & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next Cont
        End With
    End If
    L = L + 1
Loop Until Cad = ""
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub CargarComboFecha()
Dim J As Integer

    QueCombosFechaCargar "0|1|"

End Sub
