VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfBalances 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
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
      Height          =   5715
      Left            =   7110
      TabIndex        =   22
      Top             =   0
      Width           =   4485
      Begin VB.CheckBox chkSoloMes 
         Caption         =   "Saldos sólo del mes seleccionado"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   4035
      End
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
         TabIndex        =   6
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
         Left            =   120
         TabIndex        =   8
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
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Value           =   1  'Checked
         Width           =   4035
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   29
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
      Begin VB.Label lblInd 
         Caption         =   "indicador"
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
         Left            =   120
         TabIndex        =   36
         Top             =   5400
         Width           =   4095
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
         TabIndex        =   30
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
      Height          =   3045
      Left            =   120
      TabIndex        =   21
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
         ItemData        =   "frmInfBalances.frx":008B
         Left            =   2910
         List            =   "frmInfBalances.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1890
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   120
         TabIndex        =   34
         Top             =   2250
         Width           =   4665
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
            ItemData        =   "frmInfBalances.frx":008F
            Left            =   2760
            List            =   "frmInfBalances.frx":0091
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
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
            ItemData        =   "frmInfBalances.frx":0093
            Left            =   1110
            List            =   "frmInfBalances.frx":0095
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
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
            TabIndex        =   35
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
         Left            =   4440
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
         TabIndex        =   33
         Top             =   930
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
         ItemData        =   "frmInfBalances.frx":0097
         Left            =   1230
         List            =   "frmInfBalances.frx":0099
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
         Top             =   930
         Width           =   1275
      End
      Begin VB.Image imgBalan 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   930
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   570
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
         TabIndex        =   26
         Top             =   1560
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
      TabIndex        =   12
      Top             =   5910
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
      TabIndex        =   10
      Top             =   5910
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
      TabIndex        =   11
      Top             =   5850
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
      TabIndex        =   13
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
         TabIndex        =   25
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   24
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   23
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   405
      Left            =   1830
      TabIndex        =   31
      Top             =   5880
      Visible         =   0   'False
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   714
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
      TabIndex        =   32
      Top             =   5910
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

Public Legalizacion As String   'Datos para la legalizacion

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

Private Sub chk2_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub chkBalPerCompa_Click()
    Frame2.visible = Me.chkBalPerCompa.Value = 1
    Frame2.Enabled = Me.chkBalPerCompa.Value = 1
End Sub

Private Sub chkBalPerCompa_KeyPress(KeyAscii As Integer)
     KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub chkSoloMes_KeyPress(KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim F As Date

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
        
        
        
        'Si estasmos en jereccio actual o POSTERIOR, si la 129 (ctaperga ) tiene saldo AVISAMOS
        If Saldo6y7en129 Then
            I = Val(cmbFecha(2).Text) 'año
            If Me.chkBalPerCompa.Value Then
                'Comparativo
                If I < Val(cmbFecha(3).Text) Then I = Val(cmbFecha(3).Text)
                
            End If
            FechaIncioEjercicio = Format(vParam.fechaini, "dd/mm/") & CStr(I)
            If FechaIncioEjercicio >= vParam.fechaini Then
                'Esta en ejerccio actual y siguiente"
                cad = "fechaent>=" & DBSet(vParam.fechaini, "F") & " AND codmacta = '" & vParam.ctaperga & "' AND 1 "
                RC = DevuelveDesdeBD("sum(coalesce(timported,0))-sum(coalesce(timporteh,0))", "hlinapu", cad, "1")
                If RC <> "" Then
                    If RC <> "0" Then
                        cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", vParam.ctaperga, "T")
                        cad = vbCrLf & vParam.ctaperga & "  " & cad
                        cad = "La cuenta de perdidas y ganancias  tiene un saldo de : " & RC & vbCrLf & cad
                        
                        cad = cad & vbCrLf & vbCrLf & "Los saldos se solaparan"
                        MsgBox cad, vbInformation
                    End If
                End If
            End If
                
        End If
        
    End If

    
    
    
    



    Screen.MousePointer = vbHourglass
    I = -1
    If chkBalPerCompa.Value = 1 Then
        I = Val(cmbFecha(1).ListIndex)
        I = I + 1
        If I = 0 Then I = -1
    End If
    
    
    
    
    
    
    
    GeneraDatosBalanceConfigurable_ CInt(txtBalan(0).Text), Me.cmbFecha(0).ListIndex + 1, CInt(cmbFecha(2).Text), I, Val(cmbFecha(3).Text), False, -1, pb2, chkSoloMes.Value = 1, lblInd

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
    
    If Legalizacion <> "" Then
        CadenaDesdeOtroForm = "OK"
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
Dim F As Date

    If PrimeraVez Then
        PrimeraVez = False
        
        If Legalizacion <> "" Then
            optTipoSal(2).Value = True
                
            cad = RecuperaValor(Legalizacion, 3)
            F = CDate(cad)
            
            cmbFecha(0).ListIndex = Month(F) - 1
            cmbFecha(1).ListIndex = Month(F) - 1
            cmbFecha(2).Text = Year(F)
            cmbFecha(3).Text = CInt(cmbFecha(2).Text) - 1
            
            PosicionarCombo cmbFecha(2), Year(F)
            PosicionarCombo cmbFecha(3), CInt(cmbFecha(2).Text) - 1
                 
                
                
            cad = RecuperaValor(Legalizacion, 4)
                        
            If Val(cad) = 0 Then
                chkBalPerCompa.Value = 0
            Else
                'txtAno(1).Text = Val(txtAno(0).Text) - 1
                'cmbFecha(3).ListIndex = cmbFecha(2).ListIndex
                'cmbFecha(1).ListIndex = cmbFecha(0).ListIndex
                chkBalPerCompa.Value = 1
                Frame2.visible = True
                Frame2.Enabled = True
                DoEvent2
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
        
    ' solo se muestran si es balance de situacion
    chk1.visible = (Opcion = 0)
    chk1.Enabled = (Opcion = 0)
    chk2.visible = (Opcion = 0)
    chk2.Enabled = (Opcion = 0)
    

    Me.imgBalan(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    PrimeraVez = True
     
    'Fecha informe
    txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
    
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
    
    If Legalizacion <> "" Then
        PonerBalancePredeterminado
        
        txtFecha(7).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            
'        txtAno(0).Text = Year(CDate(RecuperaValor(Legalizacion, 3)))     'Fin
        PosicionarCombo cmbFecha(2), Year(CDate(RecuperaValor(Legalizacion, 3)))
        
        cmbFecha(0).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 3))) - 1
    End If
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
    AyudaBalances frmC, 0, , cWhere
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
Dim mesFinEjercicio As Boolean

    cadParam = cadParam & "pTipo=" & Tipo & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pFecha=""" & txtFecha(7).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & cmbFecha(2).Text & """|"
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
        
    

    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    cadNomRPT = nomDocu '"balance1a.rpt"

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
          
    Sql = RC & "FechaImp= """ & txtFecha(7).Text & """|"
    Sql = Sql & "Titulo= """ & Me.txtNBalan(0).Text & """|"
    
    
    
    
    If vParam.NuevoPlanContable Then
        RC = ""
        'Si es comparativo entonces idncaremos el mes y solo el mes, lo idnicaremos
                
        If chkBalPerCompa.Value = 1 Then
            If Me.chkSoloMes.Value = 1 Then
                RC = Mid(cmbFecha(0).List(cmbFecha(0).ListIndex), 1, 3) & " "
            Else
                If Month(vParam.fechafin) <> (cmbFecha(0).ListIndex + 1) Then RC = Mid(cmbFecha(0).List(cmbFecha(0).ListIndex), 1, 3) & " "
            End If
        End If
    Else
        RC = cmbFecha(0).List(cmbFecha(0).ListIndex)
    End If
    'Julio 2020
    'Si es años aprtidos, pintaresmos como año el de INICIO de ejercicio
    I = 0
    If Month(vParam.fechaini) > 1 Then
        If Month(vParam.fechaini) > (cmbFecha(0).ListIndex + 1) Then I = 1
    End If
    
    If Me.chkBalPerCompa.Value = 1 Then
        RC = RC & " " & Val(cmbFecha(2).Text) - I   ''Julio 2020  NO estaba
    Else
        RC = RC & " " & cmbFecha(2).Text - I 'txtAno(0).Text
    End If
    RC = "fec1= """ & RC & """|"
    Sql = Sql & RC
    
    
    If Me.chkBalPerCompa.Value = 1 Then
            'PGC 2008 SOlo pone el año, NO el mes
            If vParam.NuevoPlanContable Then
                RC = ""
                If Me.chkSoloMes.Value = 1 Then
                    RC = Mid(cmbFecha(1).List(cmbFecha(1).ListIndex), 1, 3) & " "
                Else
                
                    'NO ha pedido el mes de fin
                    If Month(vParam.fechafin) <> (cmbFecha(1).ListIndex + 1) Then RC = Mid(cmbFecha(1).List(cmbFecha(1).ListIndex), 1, 3) & " "
                End If
                
                
            Else
                RC = cmbFecha(1).List(cmbFecha(1).ListIndex)
            End If
            
            
            
            
            RC = RC & " " & Val(cmbFecha(3).Text) - I 'JULIO2020
            
            RC = "Fec2= """ & RC & """|"
            Sql = Sql & RC
            

    Else
        'Pong el nombre del mes
        RC = ""
        If Month(vParam.fechafin) <> (cmbFecha(0).ListIndex + 1) Then RC = UCase(Mid(cmbFecha(0).Text, 1, 1)) & Mid(cmbFecha(0).Text, 2, 2)
        RC = "vMes= """ & RC & """|"
        Sql = Sql & RC
    End If
    Sql = Sql & "Titulo= """ & Me.txtNBalan(0).Text & """|"
    Sql = Sql & "SoloMes= " & Me.chkSoloMes.Value & "|"


    cadParam = cadParam & Sql
    numParam = numParam + 4






    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text, (Legalizacion <> "")) Then ExportarPDF = False
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
