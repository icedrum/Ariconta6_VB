VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresuBal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11670
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
      Height          =   4995
      Left            =   7080
      TabIndex        =   14
      Top             =   0
      Width           =   4455
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
         ItemData        =   "frmPresuBal.frx":0000
         Left            =   1230
         List            =   "frmPresuBal.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   690
         Width           =   1935
      End
      Begin VB.CheckBox chkPreAct 
         Caption         =   "Ejercicio siguiente"
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
         Left            =   540
         TabIndex        =   37
         Top             =   1260
         Width           =   2505
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   150
         TabIndex        =   22
         Top             =   2370
         Width           =   4185
         Begin VB.CheckBox Check1 
            Caption         =   "9º nivel"
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
            Index           =   9
            Left            =   120
            TabIndex        =   32
            Top             =   1290
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "8º nivel"
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
            Index           =   8
            Left            =   2850
            TabIndex        =   31
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "7º nivel"
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
            Index           =   7
            Left            =   1470
            TabIndex        =   30
            Top             =   960
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "6º nivel"
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
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   930
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "5º nivel"
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
            Index           =   5
            Left            =   2850
            TabIndex        =   28
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "4º nivel"
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
            Index           =   4
            Left            =   1470
            TabIndex        =   27
            Top             =   600
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "3º nivel"
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
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2º nivel"
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
            Left            =   2850
            TabIndex        =   25
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1er nivel"
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
            Left            =   1470
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Último:  "
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
            Index           =   10
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   1  'Checked
            Width           =   1155
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   21
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
         Caption         =   "Mes"
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
         Left            =   540
         TabIndex        =   36
         Top             =   750
         Width           =   690
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
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNCta 
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
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1050
         Width           =   4185
      End
      Begin VB.TextBox txtNCta 
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
         Index           =   7
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1470
         Width           =   4185
      End
      Begin VB.TextBox txtCta 
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
         Index           =   6
         Left            =   1230
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1050
         Width           =   1275
      End
      Begin VB.TextBox txtCta 
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
         Index           =   7
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   6
         Left            =   990
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   7
         Left            =   990
         Top             =   1500
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
         Left            =   240
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
      Left            =   10320
      TabIndex        =   4
      Top             =   5160
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
      TabIndex        =   2
      Top             =   5160
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
      TabIndex        =   3
      Top             =   5130
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
      TabIndex        =   5
      Top             =   2340
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
         TabIndex        =   17
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   15
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
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
      TabIndex        =   33
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb4 
      Height          =   285
      Left            =   1560
      TabIndex        =   38
      Top             =   5160
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPresuBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1103

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
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon  As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private frmCtas As frmCtasAgrupadas

Private SQL As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim RS As ADODB.Recordset

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean


Dim HanPulsadoSalir As Boolean

Dim vFecIni As Date
Dim vFecFin As Date

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
    
    SQL = ""
    RC = ""
    If txtCta(6).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        RC = "Desde " & txtCta(6).Text & " - " & txtNCta(6).Text
        SQL = SQL & "presupuestos.codmacta >= '" & txtCta(6).Text & "'"
    End If
    
    
    If txtCta(7).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        If RC <> "" Then
            RC = RC & "       h"
        Else
            RC = "H"
        End If
        RC = RC & "asta " & txtCta(7).Text & " - " & txtNCta(7).Text
        SQL = SQL & "presupuestos.codmacta <= '" & txtCta(7).Text & "'"
    End If
    If SQL <> "" Then SQL = SQL & " AND"
    If chkPreAct.Value Then
        vFecIni = DateAdd("yyyy", 1, vParam.fechaini)
        vFecFin = DateAdd("yyyy", 1, vParam.fechafin)
        SQL = SQL & " date(concat(right(concat('0000',anopresu),4), right(concat('00',mespresu),2),'01')) between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
    Else
        vFecIni = vParam.fechaini
        vFecFin = vParam.fechafin
        SQL = SQL & " date(concat(right(concat('0000',anopresu),4), right(concat('00',mespresu),2),'01')) between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    End If
    
    If RC <> "" Then RC = """ + chr(13) +""" & RC
    If cmbFecha(0).ListIndex > 0 Then
        RC = "** " & Format("01/" & Format(cmbFecha(0).ListIndex, "00") & "/1999", "mmmm") & " ** " & RC
        RC = "  MENSUAL " & RC
    End If
    
    RC = "Ejercicio: " & vFecIni & " " & vFecFin & RC
    CadenaDesdeOtroForm = ""
    
    For CONT = 1 To 10
        If Check1(CONT).Value = 1 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "- " & CONT
    Next

    RC = RC & " Digitos: " & Mid(CadenaDesdeOtroForm, 2)
    RC = RC & "     Sin apertura"
    CadenaDesdeOtroForm = "CampoSeleccion= """ & RC & """|"

    cadParam = cadParam & CadenaDesdeOtroForm
    numParam = numParam + 1


    RC = ""
    For CONT = 1 To 9
        If Check1(CONT).Value = 1 Then
            If RC = "" Then RC = CONT
        End If
    Next
    If RC = "" Then RC = "11"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Remarcar= " & RC & "|"
    

    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True

    
    If Not HayRegParaInforme("tmppresu2", "codusu=" & vUsu.Codigo) Then Exit Sub
    
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
    
    Me.pb4.Visible = False
    
    
End Sub


Public Function GeneraBalancePresupuestario() As Boolean
Dim Aux As String
Dim Importe As Currency
Dim aux2 As String
Dim vMes  As Integer
Dim Cta As String

On Error GoTo EGeneraBalancePresupuestario

    GeneraBalancePresupuestario = False
    
    If Me.cmbFecha(0).ListIndex = 0 Then
        Aux = "select codmacta,sum(imppresu)  from presupuestos "
        If SQL <> "" Then Aux = Aux & " where " & SQL
        Aux = Aux & " group by codmacta"
        
        cad = "select sum(coalesce(timported,0)),sum(coalesce(timporteh,0)) from hlinapu where fechaent between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
        cad = cad & " and codmacta = '"

    Else
        Aux = "select codmacta,imppresu,mespresu, anopresu from presupuestos where " & SQL
        If cmbFecha(0).ListIndex <> 0 Then Aux = Aux & " and mespresu = " & cmbFecha(0).ListIndex 'txtMes(2).Text
        Aux = Aux & " ORDER BY codmacta,mespresu"
        cad = "select sum(coalesce(timported,0)),sum(coalesce(timporteh,0)) from hlinapu where fechaent between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
        If cmbFecha(0).ListIndex <> 0 Then cad = cad & " and month(fechaent)= " & cmbFecha(0).ListIndex  'txtMes(2).Text
        cad = cad & " and codmacta = '"
       
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ningún registro a mostrar.", vbExclamation
        RS.Close
        Exit Function
    End If
    
    'Borramos tmp de presu 2
    Aux = "DELETE FROM tmppresu2 where codusu =" & vUsu.Codigo
    Conn.Execute Aux
    
    SQL = "INSERT INTO tmppresu2 (codusu, codigo, cta, titulo,  mes, Presupuesto, realizado, anyo) VALUES ("
    SQL = SQL & vUsu.Codigo & ","
    
    CONT = 0
    Do
        CONT = CONT + 1
        RS.MoveNext
    Loop Until RS.EOF
    RS.MoveFirst
    
    'Ponemos el PB4
    pb4.Max = CONT + 1
    pb4.Value = 0
    If CONT > 3 Then pb4.Visible = True
    Cta = ""
    CONT = 1   'Contador
    While Not RS.EOF
        If Me.cmbFecha(0).ListIndex > 0 Then
            If Cta <> RS!codmacta Then
                vMes = 1
                Cta = RS!codmacta
            End If
            
            If RS!mespresu > vMes Then
                For I = vMes To RS!mespresu - 1
                
                    Aux = RS!codmacta  'Aqui pondremos el nombre
                    Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
                    Aux = CONT & ",'" & RS!codmacta & "','" & DevNombreSQL(Aux) & "',"
                    Aux = Aux & I
             
                    Aux = Aux & ",0,"
                    
                    aux2 = cad & RS!codmacta & "'"
                    aux2 = aux2 & " AND month(fechaent) =" & I
                    
                    Importe = ImporteBalancePresupuestario(aux2)
                    
                    Aux = Aux & TransformaComasPuntos(CStr(Importe)) & ","
                    Aux = Aux & DBSet(RS!anopresu, "N") & ")"
                    
                    If Importe <> 0 Then
                        Conn.Execute SQL & Aux
                        CONT = CONT + 1
                    End If
                Next I
            End If
            
        End If
                
        
    
    
        Aux = RS!codmacta  'Aqui pondremos el nombre
        Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
        Aux = CONT & ",'" & RS!codmacta & "','" & DevNombreSQL(Aux) & "',"
        If Me.cmbFecha(0).ListIndex = 0 Then
            Aux = Aux & "0"
        Else
            Aux = Aux & RS!mespresu
        End If
        Aux = Aux & "," & TransformaComasPuntos(CStr(RS.Fields(1))) & ","
        
        'SQL
        aux2 = cad & RS!codmacta & "'"
        If Me.cmbFecha(0).ListIndex > 0 Then
            aux2 = aux2 & " AND month(fechaent) =" & RS!mespresu
            'AUmento el mes
            vMes = RS!mespresu + 1
        End If
        
        
        Importe = ImporteBalancePresupuestario(aux2)
        'Debug.Print Importe
        Aux = Aux & TransformaComasPuntos(CStr(Importe)) & ","
'        If Me.chkPreMensual.Value = 0 Then
        If Me.cmbFecha(0).ListIndex = 0 Then
            Aux = Aux & "0)"
        Else
            Aux = Aux & DBSet(RS!anopresu, "N") & ")"
        End If
        Conn.Execute SQL & Aux
        
        'Sig
        pb4.Value = pb4.Value + 1
        CONT = CONT + 1
        RS.MoveNext
    Wend
    RS.Close
    
    
        '2013  Junio
    ' QUitaremos si asi lo pide, el saldo de la apertura
    RC = "" 'Por si quitamos el apunte de apertura. Guardare las cuentas para buscarlas despues en la apertura
        Aux = "SELECT cta from tmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
        RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            RC = RC & ", '" & RS!Cta & "'"
            RS.MoveNext
        Wend
        RS.Close
        
        
        
        'Subo qui lo de quitar apertura
        If RC <> "" Then
            RC = Mid(RC, 2)
            Aux = " AND codmacta IN (" & RC & ")"
            
            cad = "SELECT codmacta cta,sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) as importe"
            cad = cad & " from hlinapu where codconce=970 and fechaent='" & Format(vParam.fechaini, FormatoFecha) & "'"
            cad = cad & Aux
            cad = cad & " GROUP BY 1"
            RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                cad = "UPDATE tmppresu2 SET realizado=realizado-" & TransformaComasPuntos(CStr(RS!Importe))
                
                cad = cad & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & RS!Cta & "' AND mes = "
'                If Me.chkPreMensual.Value = 1 Then
                If Me.cmbFecha(0).ListIndex > 0 Then
                    cad = cad & " 1"
                Else
                    cad = cad & " 0"
                End If
                Conn.Execute cad
                RS.MoveNext
            Wend
            RS.Close
                
            
            
        End If
        
        
        
'    End If
    
    
    'Si pide a 3 DIGITOS este es el momemto
    'Sera facil.
    'Hacemos un insert into con substring
 
        'SUBNIVEL
        Aux = ""
        For I = 1 To 9
            If Check1(I).Value = 1 Then
                
                Aux = DevuelveDesdeBD("count(*)", "tmppresu2", "codusu", CStr(vUsu.Codigo))
                CONT = Val(Aux)
                
                '@rownum:=@rownum+1 AS rownum      (SELECT @rownum:=0) r
                Aux = "Select " & vUsu.Codigo & " us,@rownum:=@rownum+1 AS rownum,substring(cta,1," & I & ") as cta2,mes,sum(presupuesto),sum(realizado)"
                Aux = Aux & " FROM tmppresu2,(SELECT @rownum:=" & CONT & ") r WHERE codusu = " & vUsu.Codigo
                
                Aux = Aux & " AND length(cta)=" & vEmpresa.DigitosUltimoNivel
                
                Aux = Aux & " group by cta2,us,mes"
                Aux = "insert into tmppresu2 (codusu, codigo, cta,   mes, Presupuesto, realizado) " & Aux
                'Insertamos
                Conn.Execute Aux
                
                'Quito los de ultimo nivel

                
                Aux = "SELECT cta from tmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
                RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                    'Actualizo el nommacta
                    Aux = RS!Cta  'Aqui pondremos el nombre
                    Aux = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Aux, "T")
                    Aux = "UPDATE tmppresu2  SET titulo = '" & DevNombreSQL(Aux) & "' WHERE codusu = " & vUsu.Codigo & " AND Cta = '" & RS!Cta & "'"
                    Conn.Execute Aux
                    RS.MoveNext
                Wend
                RS.Close
                
                
                
            End If
        Next
        
        
        If Check1(10).Value = 0 Then
            Aux = "DELETE FROM tmppresu2 WHERE codusu = " & vUsu.Codigo & " AND cta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
            Conn.Execute Aux
        End If
        
    
    Set RS = Nothing
    GeneraBalancePresupuestario = True
    Exit Function
EGeneraBalancePresupuestario:
    MuestraError Err.Number, "Generar balance presupuestario"
    Set RS = Nothing
End Function




Private Sub cmdCancelar_Click()
    If Me.cmdCancelarAccion.Visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub Form_Activate()
Dim CONT As Integer

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
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Balance Presupuestario"

    For I = 6 To 7
        Me.imgCuentas(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    PrimeraVez = True
     
    CargarComboFecha
    
    cmbFecha(0).ListIndex = 0

    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    PonerNiveles
    
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub





Private Sub imgCuentas_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing

    PonFoco txtCta(Index)

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
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
End Sub


Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgCuentas"
        imgCuentas_Click Indice
    End Select
    
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        txtNCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        txtNCta(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 6, 7 'Cuentas
            
            RC = txtCta(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, SQL) Then
                txtCta(Index) = RC
                txtNCta(Index).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtCta(Index).Text = ""
                txtNCta(Index).Text = ""
                PonFoco txtCta(Index)
            End If
            
            If Index = 0 Then Hasta = 1
            If Hasta >= 1 Then
                txtCta(Hasta).Text = txtCta(Index).Text
                txtNCta(Hasta).Text = txtNCta(Index).Text
            End If
    End Select

End Sub



Private Sub AccionesCSV()
Dim SQL2 As String
Dim Tipo As Byte
        
    SQL = "SELECT `tmppresu2`.`cta` Cuenta, `tmppresu2`.`titulo` Nombre, `tmppresu2`.`anyo` Anyo, `tmppresu2`.`mes` Mes, `tmppresu2`.`Presupuesto` Presupuesto, `tmppresu2`.`realizado` Reslizado "
    SQL = SQL & " FROM  `tmppresu2` `tmppresu2`"
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")
    SQL = SQL & " ORDER BY `tmppresu2`.`cta`"
        
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    '------------------------------
    'Numero de niveles
    'Para cada nivel marcado veremos si tiene cuentas en la tmp
    CONT = 0
    UltimoNivel = 0
    For I = 1 To 10
        If Check1(I).Visible Then
            If Check1(I).Value = 1 Then
                If I = 10 Then
                    cad = vEmpresa.DigitosUltimoNivel
                Else
                    cad = CStr(DigitosNivel(I))
                End If
            End If
        End If
    Next I
    cad = "numeroniveles= " & CONT & "|"
    SQL = SQL & cad
    'Otro parametro mas
    cad = "vUltimoNivel= " & UltimoNivel & "|"
    
    cadParam = cadParam & cad
    numParam = numParam + 2

    
    vMostrarTree = False
    conSubRPT = False
        
    If cmbFecha(0).ListIndex > 0 Then
        indRPT = "1103-00"
    Else
        indRPT = "1103-01"
    End If
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"SumasySaldos.rpt"

    cadFormula = "{tmppresu2.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub



Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If Not ComprobarCuentas(6, 7) Then Exit Function
    
    SQL = ""
    For I = 1 To Me.Check1.Count
        If Me.Check1(I).Value Then SQL = SQL & "&"
    Next I
    If Len(SQL) <> 1 Then
        If cmbFecha(0).ListIndex > 0 Then
            MsgBox "Seleccione uno, y solo uno, de los niveles contables.", vbExclamation
            Exit Function
        End If
    End If
    
    
    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer

    QueCombosFechaCargar "0|"
    
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        Check1(I).Visible = True
        Check1(I).Caption = "Digitos: " & J
    Next I

End Sub


Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function


'Siempre k la fecha no este en fecha siguiente
Private Function HayAsientoCierre(Mes As Byte, Anyo As Integer, Optional Contabilidad As String) As Boolean
Dim C As String
    HayAsientoCierre = False
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierre = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            C = "Select count(*) from " & Contabilidad
            C = C & " hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then
                    If RS.Fields(0) > 0 Then HayAsientoCierre = True
                End If
            End If
            RS.Close
        End If
    End If
End Function



Private Function TieneCuentasEnTmpBalance(DigitosNivel As String) As Boolean
Dim RS As ADODB.Recordset
Dim C As String

    Set RS = New ADODB.Recordset
    TieneCuentasEnTmpBalance = False
    C = Mid("__________", 1, CInt(DigitosNivel))
    C = "Select count(*) from tmpbalancesumas  where cta like '" & C & "'"
    C = C & " AND codusu = " & vUsu.Codigo
    RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then TieneCuentasEnTmpBalance = True
        End If
    End If
    RS.Close
End Function

Private Sub PonerNiveles()
Dim I As Integer
Dim J As Integer


    Frame2.Visible = True
    Check1(10).Visible = True
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        cad = "Digitos: " & J
        Check1(I).Visible = True
        Me.Check1(I).Caption = cad
    Next I
    
    For I = vEmpresa.numnivel To 9
        Check1(I).Visible = False
    Next I
    
    
End Sub


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    cad = RecuperaValor(Lista, L)
    If cad <> "" Then
        I = Val(cad)
        With cmbFecha(I)
            .Clear
            RC = ""
            .AddItem RC
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
