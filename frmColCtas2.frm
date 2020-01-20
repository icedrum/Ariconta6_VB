VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmColCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11865
   Icon            =   "frmColCtas2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   6000
      TabIndex        =   21
      Top             =   150
      Width           =   3135
      Begin VB.ComboBox cboNiveles 
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
         ItemData        =   "frmColCtas2.frx":000C
         Left            =   90
         List            =   "frmColCtas2.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   210
         Width           =   1575
      End
      Begin VB.ComboBox cboOrden 
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
         ItemData        =   "frmColCtas2.frx":0050
         Left            =   1710
         List            =   "frmColCtas2.frx":005D
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   9180
      TabIndex        =   17
      Top             =   150
      Width           =   1785
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmColCtas2.frx":0094
         Left            =   120
         List            =   "frmColCtas2.frx":00A1
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   210
         Width           =   1605
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3600
      TabIndex        =   16
      Top             =   150
      Width           =   2295
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Búsqueda Avanzada"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cuentas Libres"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Observaciones Plan Contable"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Comprobar Cuentas"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cuentas sin Movimientos"
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   1410
         TabIndex        =   18
         Top             =   300
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   12
      Top             =   150
      Width           =   3405
      Begin VB.CheckBox Check2 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   150
         TabIndex        =   20
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   6840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAccion 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   10530
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Aceptar"
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
      Left            =   9300
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   5160
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   5160
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColCtas2.frx":00D8
      Height          =   5645
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11455
      _ExtentX        =   20214
      _ExtentY        =   9948
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   10530
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6030
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   11160
      TabIndex        =   19
      Top             =   300
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
   Begin VB.Label lblComprobar 
      Caption         =   "Pulse F2 para Consulta de Extractos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   2
      Left            =   3180
      TabIndex        =   24
      Top             =   6900
      Visible         =   0   'False
      Width           =   5265
   End
   Begin VB.Label lblComprobar 
      Caption         =   "Label1"
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
      Left            =   6090
      TabIndex        =   10
      Top             =   6900
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblComprobar 
      Caption         =   "Label1"
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
      Left            =   3180
      TabIndex        =   9
      Top             =   6900
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmColCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public ConfigurarBalances As Byte
    '0.- Normal
    '1.- Busqueda
    '2.- Agrupacion de cuentas
    '3.- BUSQUEDA NUEVA
    '4.- Nueva cuenta
    '5.- Busquedas de envio de e-mail
    '6.- Exclusion de cuentas en consolidado. Como la agrupacion pero acepta niveles inferiores al penultimo

Public Event DatoSeleccionado(CadenaSeleccion As String)
Public FILTRO As String   ' filtro desde otro formulario


Private Const IdPrograma = 201

Private frmCExt As frmConExtr

Private CadenaConsulta As String
Dim CadAncho As Boolean 'Para cuando llamemos al al form de lineas
Dim Rs As Recordset
Dim NF As Integer
Dim Errores As Long
Dim PrimeraVez As Boolean
Dim Aux As String

Dim vectorCboFiltroSQL() As String
 
Dim ordenColumna As Integer
Dim ordenAscendente As Boolean

Dim Modo As Byte

Dim CadB As String



'Dim Clik1 As Boolean

Private Sub BotonAnyadir(Cuenta As String)
    ParaBusqueda False
    frmCuentas.vModo = 1
    frmCuentas.CodCta = Cuenta
    CadenaDesdeOtroForm = ""
    frmCuentas.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        CargaGrid
        'Intentamos situar el grid
        SituaGrid CadenaDesdeOtroForm
        If Me.ConfigurarBalances = 4 Then cmdRegresar_Click
    Else
        If Me.ConfigurarBalances = 4 Then CargaGrid
    End If
End Sub

Private Sub BotonBuscar()
    CadenaConsulta = GeneraSQL(" false ")  'esto es para que no cargue ningun registro  Ponia codmacta= 'David'
    CargaGrid
    ParaBusqueda True
    txtaux(0).Text = "": txtaux(1).Text = "": txtaux(2).Text = ""
    '++
    Modo = 1
    PonerModoUsuarioGnral Modo, "ariconta"
    ConseguirFoco txtaux(1), Modo
    PonFoco txtaux(1)
End Sub

Private Sub BotonVerTodos()
    ParaBusqueda False
    CadenaConsulta = GeneraSQL("")
    CargaGrid
    PonerModoUsuarioGnral 2, "ariconta"
End Sub



Private Sub BotonModificar()
    
    ParaBusqueda False
    CadenaDesdeOtroForm = ""
    frmCuentas.vModo = 2
    frmCuentas.CodCta = Adodc1.Recordset!codmacta
    frmCuentas.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CargaGrid
        SituaGrid CadenaDesdeOtroForm
    End If
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    ParaBusqueda False
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar la cuenta:"
    SQL = SQL & vbCrLf & "Código: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & Adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = Adodc1.Recordset.Fields(0)
        Screen.MousePointer = vbHourglass
        If SepuedeEliminarCuenta(SQL) Then
            'Hay que eliminar
            Screen.MousePointer = vbHourglass
            
            SQL = "Delete from departamentos where codmacta = " & DBSet(Adodc1.Recordset!codmacta, "T")
            Conn.Execute SQL
            
            SQL = "Delete from cuentas where codmacta='" & Adodc1.Recordset!codmacta & "'"
            Conn.Execute SQL
            Screen.MousePointer = vbHourglass
            espera 0.5
            'Cancelamos el adodc1
            DataGrid1.Enabled = False
            NumRegElim = Adodc1.Recordset.AbsolutePosition - 1
            CargaGrid
            DataGrid1.Enabled = True
            If NumRegElim > 0 Then
                If NumRegElim >= Adodc1.Recordset.RecordCount Then
                    Adodc1.Recordset.MoveLast
                Else
                    Adodc1.Recordset.Move NumRegElim
                    'DataGrid1.Bookmark = Adodc1.Recordset.AbsolutePosition
                End If
            End If
            
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            Errores = Err.Number
            SQL = Err.Description
            FijarError SQL
            MsgBox "Error eliminando cuenta. " & vbCrLf & SQL, vbExclamation
        End If
End Sub

Private Sub FijarError(ByRef cad As String)
    On Error Resume Next
    cad = Conn.Errors(0).Description
    If Err.Number <> 0 Then
        Err.Clear
        cad = ""
    End If
End Sub


Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adReason = adRsnMove And adStatus = adStatusOK Then PonLblIndicador Me.lblIndicador, Adodc1
        
End Sub




Private Sub OpcionesCambiadas()
    If txtaux(0).visible Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    CadenaConsulta = GeneraSQL("")
    CargaGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    OpcionesCambiadas
End Sub


Private Sub cboNiveles_Click()
    If PrimeraVez Then Exit Sub
    OpcionesCambiadas
End Sub

Private Sub cboOrden_Click()
    If PrimeraVez Then Exit Sub
    OpcionesCambiadas
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim SQL As String


    If Index = 0 Then
        'Ha pulsado aceptar
        txtaux(0).Text = Trim(txtaux(0).Text)
        txtaux(1).Text = Trim(txtaux(1).Text)
        'Si estan vacios no hacemos nada
        SQL = ""
        Aux = ""
        If txtaux(0).Text <> "" Then
            If SeparaCampoBusqueda("T", "codmacta", txtaux(0).Text, Aux, False) = 0 Then SQL = Aux
        End If
        If txtaux(1).Text <> "" Then
            Aux = ""
            
            'VEO si ha puesto un *
            If InStr(1, txtaux(1).Text, "*") = 0 Then txtaux(1).Text = "*" & txtaux(1).Text & "*"
            If SeparaCampoBusqueda("T", "nommacta", txtaux(1).Text, Aux, False) = 0 Then
                If SQL <> "" Then SQL = SQL & " AND "
                SQL = SQL & Aux
            End If
        End If
        
        If txtaux(2).Text <> "" Then
            Aux = ""
            Dim CADENA As String
            
            If InStr(1, txtaux(2).Text, ">>") <> 0 Then
                CADENA = "select max(fecbloq) from cuentas "
            Else
                If InStr(1, txtaux(2).Text, "<<") <> 0 Then
                    CADENA = "select min(fecbloq) from cuentas "
                Else
                    If InStr(1, UCase(txtaux(2).Text), "=NULL") <> 0 Then
                        Aux = "(fecbloq is null)"
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & Aux
                    Else
                        If InStr(1, UCase(txtaux(2).Text), "<>NULL") <> 0 Then
                            Aux = "(not fecbloq is null)"
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & Aux
                        End If
                    End If
                End If
            End If
            If DevuelveValor(CADENA) <> "0" Then
                txtaux(2).Text = DevuelveValor(CADENA)
            End If
            
            
            If SeparaCampoBusqueda("F", "fecbloq", txtaux(2).Text, Aux, False) = 0 Then
                If SQL <> "" Then SQL = SQL & " AND "
                SQL = SQL & Aux
            End If
                
            
            
'                If SQL <> "" Then SQL = SQL & " AND "
'                SQL = SQL & AUx
'            End If
        End If
        
        
        'Si sql<>"" entonces hay puestos valores
        If SQL = "" Then Exit Sub
        
        'Llamamos a carga grid
        Screen.MousePointer = vbHourglass
        CadenaConsulta = GeneraSQL(SQL)
        CargaGrid
        Screen.MousePointer = vbDefault
        If Adodc1.Recordset.EOF Then
            MsgBox "Ningún resultado para la búsqueda.", vbExclamation
            Exit Sub
        Else
            
        End If
        PonerFoco DataGrid1
    End If
    ParaBusqueda False
    PonerModoUsuarioGnral Modo, "ariconta"
    'lblIndicador.Caption = ""
End Sub

Private Sub cmdRegresar_Click()
    If Adodc1.Recordset Is Nothing Then
        BotonBuscar
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    RaiseEvent DatoSeleccionado(Adodc1.Recordset!codmacta & "|" & Adodc1.Recordset!Nommacta & "|" & Adodc1.Recordset!bloqueada & "|")
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then
        cmdRegresar_Click
    Else
    
        If Adodc1.Recordset Is Nothing Then Exit Sub
        If Adodc1.Recordset.EOF Then Exit Sub
    
        'Vemos todos los valores de la cuenta
        frmCuentas.vModo = 0
        frmCuentas.CodCta = Adodc1.Recordset!codmacta
        frmCuentas.Show vbModal
    End If
End Sub


Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Modo = 0 Or Modo = 2 Then
        Form_KeyDown KeyCode, Shift
    End If
    
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then  ' si pulsa F2 vamos a consulta de extractos
    
        If Me.Adodc1.Recordset.EOF Then Exit Sub
        
        Aux = DBSet(vParam.fechaini, "F") & " AND "
        Aux = Aux & DBSet(DateAdd("yyyy", 1, vParam.fechafin), "F")
        Aux = Aux & " AND codmacta  "
        Aux = DevuelveDesdeBD("count(*)", "hlinapu", Aux, Adodc1.Recordset!codmacta, "T")
        If Val(Aux) = 0 Then
             MsgBox "La cuenta " & Adodc1.Recordset!codmacta & " NO tiene movimientos en las fechas", vbExclamation
        
        Else
            Set frmCExt = New frmConExtr
            
            frmCExt.Cuenta = Me.Adodc1.Recordset.Fields(0).Value
            frmCExt.Show vbModal
            
            Set frmCExt = Nothing
        End If
        KeyCode = 0
        Exit Sub
    End If

    If Shift = 0 And KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    End If
End Sub
'++
Private Sub Form_Activate()
    If PrimeraVez Then
        cboOrden.ListIndex = 0
        
        CargarNiveles 1
        PrimeraVez = False
        DoEvents
        'Vamos a ver si funciona 30 Sept 2003
        Select Case ConfigurarBalances
            Case 0, 1, 2, 5, 6
                If ConfigurarBalances = 2 Or ConfigurarBalances = 6 Then
                    cboNiveles.ListIndex = 0
                    CadenaConsulta = GeneraSQL("")
                End If
                If ConfigurarBalances = 5 Then
                    'Estoy buscando los que tienen e-mail
                    CadenaConsulta = CadenaConsulta & " AND maidatos <> ''"
                End If

                CargaGrid
            
            Case 3
                BotonBuscar
            Case 4
                
                BotonAnyadir CadenaDesdeOtroForm
            
        End Select
        CadenaDesdeOtroForm = ""
        
        If FILTRO <> "" Then
        
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    
    Me.Icon = frmppal.Icon

    Screen.MousePointer = vbHourglass
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
     
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 24
        .Buttons(2).Image = 47
        .Buttons(4).Image = 19
        .Buttons(6).Image = 37
        .Buttons(7).Image = 32
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With

    PrimeraVez = True
    
    CargarNiveles 0
    
    CargarOrden
    
    CargaFiltros
     
    pb1.visible = False
    'Poner niveles
    PonerOptionsVisibles 0
    
    'Opciones segun sea su nivel
    PonerOpcionesMenu
    
    Modo = 0
    PonerModoUsuarioGnral Modo, "ariconta"
    
    'Ocultamos busqueda
    ParaBusqueda False
    
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    CadAncho = False
    CadenaConsulta = GeneraSQL("")
    
    lblIndicador.Caption = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    ConfigurarBalances = 0
End Sub

Private Sub frmSelec_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion = "CANCEL" Then
        CadB = ""
    Else
        CadB = CadenaSeleccion
    End If
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir ""
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'----------------------------------------------------------------


Private Sub Option2_Click(Index As Integer)
    If PrimeraVez Then Exit Sub
    OpcionesCambiadas
End Sub

Private Sub Option2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String
    WheelUnHook
    Select Case Button.Index
        Case 1
                BotonAnyadir ""
        Case 2
                BotonModificar
        Case 3
                BotonEliminar
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
        Case 8
                'Imprimimos el listado
                frmColCtasList.Show vbModal
        
        Case Else
        
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim C As String
    WheelUnHook
    Select Case Button.Index
        Case 1
            'Busqueda avanzada
                CadenaDesdeOtroForm = ""
                frmCuentas.vModo = 3
                frmCuentas.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    Me.Refresh
                    Screen.MousePointer = vbHourglass
                    ParaBusqueda False
                    PonerResultadosBusquedaAvanzada
                    Screen.MousePointer = vbDefault
                End If

        Case 2 ' cuentas libres
            'Busqueda ctas libres
            Screen.MousePointer = vbHourglass
            CadenaDesdeOtroForm = ""
            frmUtilidades.Opcion = 7
            frmUtilidades.Show vbModal
            If CadenaDesdeOtroForm <> "" Then
                C = CadenaDesdeOtroForm
                CadenaDesdeOtroForm = ""
                BotonAnyadir C
                If CadenaDesdeOtroForm <> "" Then
                    Me.Refresh
                    DoEvents
                    Screen.MousePointer = vbHourglass
                    CadenaDesdeOtroForm = " codmacta = '" & CadenaDesdeOtroForm & "'"
                    PonerResultadosBusquedaAvanzada
                    Screen.MousePointer = vbDefault
                End If
            End If
    
        
        Case 6 '12
                'Comprobar cuentas
                ComprobarCuentas
                
        Case 7 ' cuentas sin movimientos
                Screen.MousePointer = vbHourglass
                frmUtilidades.Caption = ""
                frmUtilidades.Label1 = "Cuentas sin Movimientos"
                frmUtilidades.Opcion = 0
                frmUtilidades.Show vbModal
                
        Case Else
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim i
    For i = 16 To 19
        Toolbar1.Buttons(i).Enabled = bol
        Toolbar1.Buttons(i).visible = bol
    Next i
End Sub


Private Sub CargaGrid()
Dim B As Boolean

    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    espera 0.07
    lblComprobar(0).visible = True
    lblComprobar(0).Caption = "Leyendo BD"
    lblComprobar(0).Refresh
    CargaGrid2
'--
'    PonerFoco Check1(0)
    lblComprobar(0).visible = False
    DataGrid1.Enabled = B
    
    
    lblComprobar(2).visible = Not Me.Adodc1.Recordset.EOF
    
End Sub

Private Sub CargaGrid2()
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    Dim SQL As String
    Dim B As Boolean
    Adodc1.ConnectionString = Conn
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    SQL = CadenaConsulta
    SQL = SQL & " ORDER BY"
    If cboOrden.ListIndex = 0 Then
    'If Option2(0).Value Then
        SQL = SQL & " codmacta"
    Else
        SQL = SQL & " nommacta"
    End If
    Adodc1.RecordSource = SQL
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 350
    
    
   
        DataGrid1.Columns(0).Caption = "Cuenta"
        DataGrid1.Columns(0).Width = 1400 '1200
    
   
        DataGrid1.Columns(1).Caption = "Denominación"
        DataGrid1.Columns(1).Width = 8280 '5300
        TotalAncho = TotalAncho + DataGrid1.Columns(1).Width
    
   
        DataGrid1.Columns(2).Caption = "Direc."
        DataGrid1.Columns(2).Width = 600 '500
        DataGrid1.Columns(2).visible = False
        TotalAncho = TotalAncho + DataGrid1.Columns(2).Width
               
        DataGrid1.Columns(3).Caption = "F.Bloqueo"
        DataGrid1.Columns(3).Width = 1200 '600 '400
        TotalAncho = TotalAncho + DataGrid1.Columns(3).Width
               
               
               
        If Not CadAncho Then
            txtaux(0).Left = DataGrid1.Columns(0).Left + 150
            txtaux(0).Width = DataGrid1.Columns(0).Width - 30
            txtaux(0).top = DataGrid1.top + 260 ' 235
            txtaux(1).Left = DataGrid1.Columns(1).Left + 150
            txtaux(1).Width = DataGrid1.Columns(1).Width - 30
            txtaux(2).Left = DataGrid1.Columns(3).Left + 150
            txtaux(2).Width = DataGrid1.Columns(3).Width - 30
            txtaux(1).top = txtaux(0).top
            txtaux(2).top = txtaux(0).top
            txtaux(0).Height = DataGrid1.RowHeight - 15
            txtaux(1).Height = txtaux(0).Height
            txtaux(2).Height = txtaux(0).Height
            CadAncho = True
        End If
               
    'Habilitamos modificar y eliminar
    Toolbar1.Buttons(4).Enabled = vUsu.Nivel < 3
    Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    Else
        Toolbar1.Buttons(8).Enabled = False
    End If
   
    'Para k la barra de desplazamiento sea mas alta
    If Not Adodc1.Recordset.EOF Then
            DataGrid1.ScrollBars = dbgVertical
    End If
    DataGrid1.Enabled = B
End Sub


' 0 solo textos
'1 Solo enables
'2 todo
Private Sub PonerOptionsVisibles(Opcion As Byte)
Dim i As Integer
Dim J As Integer
Dim cad As String

End Sub



Private Function GeneraSQL(Busqueda As String) As String
Dim i As Integer
Dim SQL As String
Dim nexo As String
Dim J As Integer
Dim wildcar As String
Dim Digitos As Integer
Dim cad As String

    SQL = ""
    nexo = ""
    
    'If Check1(0).Value Then
    If cboNiveles.ListIndex = vEmpresa.numnivel - 1 Then
        SQL = "( apudirec = 'S')"
        nexo = " OR "
    End If

    If cboNiveles.ListIndex <> -1 Then
        Digitos = cboNiveles.ItemData(cboNiveles.ListIndex)
        wildcar = ""
        For J = 1 To Digitos
            wildcar = wildcar & "_"
        Next J
        SQL = SQL & nexo & " ( codmacta like '" & wildcar & "')"
        nexo = " OR "
    End If
    
    wildcar = "SELECT codmacta, nommacta, apudirec, fecbloq as bloqueada" 'apudirec,if(fecbloq is null,"""",""*"") as bloqueada"
    wildcar = wildcar & " FROM cuentas "
    
    
    'Nexo
    nexo = " WHERE "
    If Busqueda <> "" Then
        wildcar = wildcar & " WHERE (" & Busqueda & ")"
        nexo = " AND "
    End If
    If SQL <> "" Then wildcar = wildcar & nexo & "(" & SQL & ")"
    
    '++ añado el filtro
    'filtro
    cad = DevuelveSQLFiltro(True)
    If cad <> "" Then wildcar = wildcar & cad
    
    GeneraSQL = wildcar
End Function



Private Function SepuedeEliminarCuenta(Cuenta As String) As Boolean
Dim NivelCta As Integer
Dim i, J As Integer
Dim cad As String

    SepuedeEliminarCuenta = False
    If EsCuentaUltimoNivel(Cuenta) Then
        'ATENCION###
        ' Habra que ver casos particulares de eliminacion de una subcuenta de ultimo nivel
        'Si esta en apuntes, en ....
        'NO se puede borrar
        lblComprobar(0).Caption = "Comprobando"
        lblComprobar(0).visible = True
        cad = BorrarCuenta(Cuenta, lblComprobar(0))
        lblComprobar(0).visible = False
        If cad <> "" Then
            cad = Cuenta & vbCrLf & cad
            MsgBox cad, vbExclamation
            Exit Function
        End If
        
    Else
        'No
        'No
        'no es una cuenta de ultimo nivel
        NivelCta = NivelCuenta(Cuenta)
        If NivelCta < 1 Then
            MsgBox "Error obteniendo nivel de la subcuenta", vbExclamation
            Exit Function
        End If
        
        'Ctas agrupadas
        i = DigitosNivel(NivelCta)
        If i = 3 Then
            cad = DevuelveDesdeBD("codmacta", "ctaagrupadas", "codmacta", Cuenta, "T")
            If cad <> "" Then
                MsgBox "El subnivel pertenece a agrupacion de cuentas en balance"
                Exit Function
            End If
        End If
        For J = NivelCta + 1 To vEmpresa.numnivel
            cad = Cuenta & "__________"
            i = DigitosNivel(J)
            cad = Mid(cad, 1, i)
            If TieneEnBD(cad) Then
                MsgBox "Tiene cuentas en niveles superiores (" & J & ")", vbExclamation
                Exit Function
            End If
        Next J
    End If
    SepuedeEliminarCuenta = True
End Function

Private Function TieneEnBD(cad As String) As String
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select codmacta from cuentas where codmacta like '" & cad & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TieneEnBD = Not Rs.EOF
    Rs.Close
    Set Rs = Nothing
End Function


Private Sub SituaGrid(CADENA As String)
On Error GoTo ESituaGrid
If Adodc1.Recordset.EOF Then Exit Sub

Adodc1.Recordset.Find " codmacta =  " & CADENA & ""
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst

Exit Sub
ESituaGrid:
    MuestraError Err.Number, "Situando registro activo"
End Sub


Private Sub ParaBusqueda(Ver As Boolean)
txtaux(0).visible = Ver
txtaux(1).visible = Ver
txtaux(2).visible = Ver
cmdAccion(0).visible = Ver
cmdAccion(1).visible = Ver
If Ver Then
    lblIndicador.Caption = "BUSQUEDA"
Else
    lblIndicador.Caption = ""
    Modo = 0
End If

End Sub



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub




Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), 1
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub ComprobarCuentas()
Dim cad As String
Dim N As Integer
Dim i As Integer
Dim Col As Collection
Dim C1 As String

On Error GoTo EComprobarCuentas
    
    
    'NO hay cuentas
    If Me.Adodc1.Recordset.EOF Then Exit Sub
    'Buscando datos
    If txtaux(0).visible Then Exit Sub
    'Para cada nivel n comprobaremos si existe la cuenta en un
    'nivel n-1
    'La comprobacion se hara para cada cta de n sabiendo k
    ' para las cuentas de nivel 4 digitos  4300 ..4309 tienen
    'el mismo subnivel n-1 430
    lblComprobar(0).Caption = ""
    lblComprobar(1).Caption = ""
    lblComprobar(0).visible = True
    lblComprobar(1).visible = True
    lblComprobar(2).visible = False
    Me.lblIndicador.Caption = "Comprobar cuentas"
    Me.Refresh
    Errores = 0
    NF = FreeFile
    Open App.Path & "\Errorcta.txt" For Output As #NF
    
    
    'Primero comprobamos las cuentas de mayor longitud que la permitida
    CuentasDeMasNivel
    lblComprobar(0).Caption = "Cuentas > Ultimo nivel"
    lblComprobar(0).Refresh
    Set Col = New Collection
    'Hasta 2 pq el uno no tiene subniveles
    For i = vEmpresa.numnivel To 2 Step -1
        N = DigitosNivel(i)
        lblComprobar(0).Caption = "Nivel: " & i
        lblComprobar(0).Refresh
        Do
            If ObtenerCuenta(cad, i, N) Then
                lblComprobar(1).Caption = cad
                lblComprobar(1).Refresh
                ComprobarCuenta cad, i, Col
            End If
        Loop Until cad = ""
    Next i
    
    
    'Otras comprobaciones de las cuentas
    Me.lblComprobar(0).Caption = "Comp. cta numerica o con ' '"
    Me.lblComprobar(1).Caption = "Leyendo BD"
    Set Rs = New ADODB.Recordset
    OtrasComprobacionesCuentas
    Set Rs = Nothing
    
    Close #NF
    Me.lblComprobar(0).Caption = "Proceso"
    Me.lblComprobar(1).Caption = "Finalizado"
    
    If Errores = 0 Then
        Kill App.Path & "\Errorcta.txt"
        MsgBox "Comprobación finalizada", vbInformation
        
        Else
            cad = Dir("C:\WINDOWS\NOTEPAD.exe")
            If cad = "" Then
                cad = Dir("C:\WINNT\NOTEPAD.exe")
            End If
            If cad = "" Then
                MsgBox "Se ha producido errores. Vea el archivo Errorcta.txt"
                Else
                Shell cad & " " & App.Path & "\Errorcta.txt", vbMaximizedFocus
            End If
            espera 2
    
            If vUsu.Nivel < 2 Then
                If MsgBox("Desea crear los subniveles?", vbQuestion + vbYesNo) = vbYes Then
                        
                        cad = "insert into `cuentas` (`codmacta`,`nommacta`,`apudirec`,dirdatos) VALUES ('"
                        For NF = 1 To Col.Count
                            N = DigitosNivelAnterior(Col.Item(NF))
                            If N > 0 Then
                                C1 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Mid(Col.Item(NF), 1, N), "T")
                            Else
                                C1 = ""
                            End If
                            If C1 = "" Then C1 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Col.Item(NF), "T")
                            
                            If C1 = "" Then C1 = "AUTOM: " & Col.Item(NF)
    
                            EjecutaSQL cad & Col.Item(NF) & "','" & DevNombreSQL(C1) & "','N','AUTOMATICA en la comprobacion')"
                        Next NF
                End If
            End If
    End If

EComprobarCuentas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar cuentas: ", Err.Description
        Close #NF
    End If
    Me.lblComprobar(0).visible = False
    Me.lblComprobar(1).visible = False
    lblComprobar(2).visible = True
    Me.lblIndicador.Caption = ""
    Me.Refresh
    Set Col = Nothing
End Sub


Private Function ObtenerCuenta(ByRef CADENA As String, Nivel As Integer, ByRef Digitos As Integer) As Boolean
Dim RT As Recordset
Dim SQL As String


If CADENA = "" Then
    SQL = ""
Else
    SQL = DevuelveUltimaCuentaGrupo(CADENA, Nivel, Digitos)
    SQL = " codmacta > '" & SQL & "' AND "
End If
SQL = "Select codmacta from Cuentas WHERE " & SQL
SQL = SQL & " codmacta like '" & Mid("__________", 1, Digitos) & "'"

Set RT = New ADODB.Recordset
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RT.EOF Then
    ObtenerCuenta = False
    CADENA = ""
Else
    ObtenerCuenta = True
    CADENA = RT!codmacta
End If
RT.Close
Set RT = Nothing
End Function


Private Function DigitosNivelAnterior(Cuenta As String) As Integer
Dim i As Integer
    i = Len(Cuenta)
    Select Case i
    Case vEmpresa.numdigi7
        DigitosNivelAnterior = vEmpresa.numdigi6
        
    Case vEmpresa.numdigi5
        DigitosNivelAnterior = vEmpresa.numdigi4
        
    Case vEmpresa.numdigi4
        DigitosNivelAnterior = vEmpresa.numdigi3
        
    
    Case vEmpresa.numdigi3
        DigitosNivelAnterior = vEmpresa.numdigi2
    
    Case vEmpresa.numdigi2
        DigitosNivelAnterior = vEmpresa.numdigi1
    

    Case Else
        DigitosNivelAnterior = 0
    End Select
    
        
End Function


Private Sub ComprobarCuenta(Cuenta As String, Nivel As Integer, ByRef Cole As Collection)
Dim N As Integer
Dim aux2 As String

N = DigitosNivel(Nivel - 1)
Aux = Mid(Cuenta, 1, N)
aux2 = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Aux, "T")
If aux2 = "" Then
    'Error
   Errores = Errores + 1
   Print #NF, "Nivel: " & Nivel
   Print #NF, "Cuenta: " & Cuenta & "  -> " & Aux & " NO encontrada "
   Print #NF, ""
   Print #NF, ""
   Cole.Add Aux
End If

End Sub



Private Function DevuelveUltimaCuentaGrupo(Cta As String, Nivel As Integer, ByRef Digitos As Integer) As String
Dim cad As String
Dim N As Integer
N = DigitosNivel(Nivel - 1)
cad = Mid(Cta, 1, N)
cad = cad & "9999999999"
DevuelveUltimaCuentaGrupo = Mid(cad, 1, Digitos)
End Function


Private Sub CuentasDeMasNivel()
'###MYSQL
Set Rs = New ADODB.Recordset
Rs.Open "SELECT codmacta FROM cuentas WHERE ((Length(cuentas.codmacta)>" & vEmpresa.DigitosUltimoNivel & "))", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    Print #NF, "Cuentas de longitud mayor a la permitida"
    Print #NF, "Digitos ultimo nivel: " & vEmpresa.DigitosUltimoNivel
    While Not Rs.EOF
        Errores = Errores + 1
        Print #NF, "      .- " & Rs!codmacta
        Rs.MoveNext
    Wend
    Print #NF, ""
    Print #NF, ""
End If
Rs.Close
Set Rs = Nothing
End Sub



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub PonerFoco(ByRef Obje As Object)
On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerResultadosBusquedaAvanzada()

    On Error GoTo EC
        CadenaConsulta = GeneraSQL(CadenaDesdeOtroForm)
        CargaGrid
    Exit Sub
EC:
    MuestraError Err.Number, "Poner resultados busqueda avanzada"
End Sub

' ### [DavidV] 23/12/2016: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub



Private Sub OtrasComprobacionesCuentas()

    'Busco cuentas que no sean numericas
    Me.Refresh
    Aux = "Select codmacta from cuentas"
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = "Cuentas NO numericas o con espacios en blanco"
    While Not Rs.EOF
        Me.lblComprobar(1).Caption = Rs!codmacta
        Me.lblComprobar(1).Refresh
         If Not IsNumeric(Rs!codmacta) Then
            EscribeEnErrores Aux, Rs!codmacta
         Else
            If InStr(1, Rs!codmacta, " ") > 0 Then EscribeEnErrores Aux, Rs!codmacta
        End If
        Rs.MoveNext
    Wend
    Rs.Close
End Sub

Private Sub EscribeEnErrores(Titulito As String, Cuenta As String)
    'Error
   Errores = Errores + 1
   If Titulito <> "" Then
        Print #NF, " *****  " & Titulito
        Print #NF,: Print #NF,: Print #NF,
        Titulito = ""
   End If
   Print #NF, " - " & Cuenta


End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub

    If Index = 0 Then
        Aux = txtaux(0).Text
        If CuentaCorrectaUltimoNivel(Aux, "") Then txtaux(0).Text = Aux
    End If
End Sub


Private Sub CargarNiveles(Opcion As Byte)
Dim i As Integer

    
    If Opcion <> 1 Then
        cboNiveles.Clear
        For i = 1 To vEmpresa.numnivel - 1
            J = DigitosNivel(i)
            If J > 0 Then
                cboNiveles.AddItem "Digitos " & i
                cboNiveles.ItemData(cboNiveles.NewIndex) = i
            Else
            End If
        Next i
        'Ultimo nivel
        cboNiveles.AddItem "Ultimo nivel"
        cboNiveles.ItemData(cboNiveles.NewIndex) = 9
        cboNiveles.ListIndex = cboNiveles.ListCount - 1
    End If
    
    If Opcion <> 0 Then
        Select Case ConfigurarBalances
        Case 1
            cboNiveles.Clear
            For i = 1 To vEmpresa.numnivel - 1
                J = DigitosNivel(i)
                If J < 5 Then  'A balances van ctas de 4 digitos
                    cboNiveles.AddItem "Digitos " & i
                    cboNiveles.ItemData(cboNiveles.NewIndex) = i
                Else
                End If
            Next i
        Case 2
            cboNiveles.Clear

            'Agrupar ctas digitos . Realmete agrupamos al nivel de cuentas -1
            J = DevuelveDigitosNivelAnterior
            For i = 0 To 9
                If i = J Then
                    cboNiveles.AddItem "Digitos " & i
                    cboNiveles.ItemData(cboNiveles.NewIndex) = i
                End If
            Next i
    
        Case 6
            cboNiveles.Clear

            'Todos los niveles menos el ultimo
            'Agrupar ctas digitos . Realmete agrupamos al nivel de cuentas -1
            J = DevuelveDigitosNivelAnterior
            For i = 0 To 9
                If i < J And i > 0 Then
                    cboNiveles.AddItem "Digitos " & i
                    cboNiveles.ItemData(cboNiveles.NewIndex) = i
                End If
            Next i
        End Select
    End If
    

End Sub

Private Sub CargarOrden()
Dim i As Integer

    cboOrden.Clear
    
    cboOrden.AddItem "Código "
    cboOrden.ItemData(cboOrden.NewIndex) = 0
    cboOrden.AddItem "Alfabético "
    cboOrden.ItemData(cboOrden.NewIndex) = 1

End Sub



Private Sub CargaFiltros()
Dim Aux As String
    
    'Leo los filtros por defecto
    If Aux = "" Then Aux = "0"
    
    'Aqui, si no quisieramos filtro, pondriamos visible a false y a correr
    FrameFiltro.visible = True
    
    '-------------------------------------
    CargaVectoresFiltro 6, "Clientes|Proveedores|Acreedores|Bancos|Gastos|Ingresos|", Me.cboFiltro
    ReDim vectorCboFiltroSQL(6)   'El 0 es sin filtro
    vectorCboFiltroSQL(1) = "1cuentas.codmacta like '43%'"
    vectorCboFiltroSQL(2) = "1cuentas.codmacta like '40%'"
    vectorCboFiltroSQL(3) = "1cuentas.codmacta like '41%'"
    vectorCboFiltroSQL(4) = "1cuentas.codmacta like '57%'"
    vectorCboFiltroSQL(5) = "1cuentas.codmacta like '6%'"
    vectorCboFiltroSQL(6) = "1cuentas.codmacta like '7%'"
    
    
    'Situamos el filtro en un valor "guardado"
    If FILTRO <> "" Then
        If vParam.AplicarFiltrosEnCuentas Then Aux = FILTRO
    End If
    Me.cboFiltro.ListIndex = CInt(Aux)
    
    
End Sub

Private Function DevuelveSQLFiltro(AlWHERE As Boolean) As String
    
    If Me.cboFiltro.ListIndex <= 0 Then Exit Function
    
    If AlWHERE Then
        If Mid(vectorCboFiltroSQL(cboFiltro.ListIndex), 1, 1) = "0" Then
            DevuelveSQLFiltro = ""
        Else
            DevuelveSQLFiltro = Mid(vectorCboFiltroSQL(cboFiltro.ListIndex), 2)
            'EN ESTE FORM si que ponemos el AND
            DevuelveSQLFiltro = " AND " & DevuelveSQLFiltro
        End If
    Else
        If Mid(vectorCboFiltroSQL(cboFiltro.ListIndex), 1, 1) = "0" Then
            DevuelveSQLFiltro = Mid(vectorCboFiltroSQL(cboFiltro.ListIndex), 2)
        Else
            DevuelveSQLFiltro = ""
        End If
    End If
    
    
End Function

'**************************************************************************
'**************************************************************************
'**************************************************************************

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!especial, "N") And ConfigurarBalances <> 2 And ConfigurarBalances <> 6 'And (DatosADevolverBusqueda = "")
        Me.Toolbar2.Buttons(6).Enabled = DBLet(Rs!especial, "N") 'And (DatosADevolverBusqueda = "")
        Me.Toolbar2.Buttons(7).Enabled = DBLet(Rs!especial, "N") 'And (DatosADevolverBusqueda = "")
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



