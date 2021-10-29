VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmColCtasLW 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColCtasLW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   7095
      Left            =   120
      TabIndex        =   22
      Top             =   960
      Width           =   15615
      _Version        =   1114114
      _ExtentX        =   27543
      _ExtentY        =   12515
      _StockProps     =   64
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
      AllowColumnResize=   0   'False
      AllowColumnSort =   0   'False
      MultipleSelection=   0   'False
      FreezeColumnsAbs=   0   'False
      MultiSelectionMode=   -1  'True
      HeaderRowsAllowAccess=   0   'False
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
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
      Left            =   7800
      TabIndex        =   2
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
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
      Left            =   5400
      TabIndex        =   1
      Top             =   7920
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
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
      Left            =   4200
      TabIndex        =   0
      Top             =   7920
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6700
      TabIndex        =   18
      Top             =   120
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
         ItemData        =   "frmColCtasLW.frx":000C
         Left            =   90
         List            =   "frmColCtasLW.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   20
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
         ItemData        =   "frmColCtasLW.frx":0050
         Left            =   1710
         List            =   "frmColCtasLW.frx":005D
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame FrameFiltro 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   9870
      TabIndex        =   14
      Top             =   120
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
         ItemData        =   "frmColCtasLW.frx":0094
         Left            =   120
         List            =   "frmColCtasLW.frx":00A1
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   210
         Width           =   1605
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   2895
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
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
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Activar / Desactivar rejilla"
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   15
         Top             =   300
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.Frame FrameBotonGnral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3405
      Begin VB.CheckBox Check2 
         Caption         =   "Vista previa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   150
         TabIndex        =   17
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
      Left            =   14760
      TabIndex        =   4
      Top             =   8280
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
      Left            =   13560
      TabIndex        =   3
      Top             =   8280
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   8160
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
         TabIndex        =   6
         Top             =   210
         Width           =   2550
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   15360
      TabIndex        =   16
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
      Left            =   3120
      TabIndex        =   21
      Top             =   8400
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
      Left            =   9000
      TabIndex        =   8
      Top             =   8160
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
      Left            =   9000
      TabIndex        =   7
      Top             =   8520
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
Attribute VB_Name = "frmColCtasLW"
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

Dim fntStrike
Dim fntBold

Dim CargandoDatos As Boolean
Dim CuentasParaSaldos As String
Dim UltimaCuenta As String
Dim HayMasDatos As Boolean

Dim Sql_BusquedaAvanzada As String

'ME.tag llevará el SQL


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
       
    Else
        If Me.ConfigurarBalances = 4 Then CargaGrid
    End If
End Sub

Private Sub BotonBuscar()
    CadenaConsulta = GeneraSQL(" false ")  'esto es para que no cargue ningun registro  Ponia codmacta= 'David'
    CargaGrid
    ParaBusqueda True
    txtAux(0).Text = "": txtAux(1).Text = "": txtAux(2).Text = ""
    '++
    Modo = 1
    PonerModoUsuarioGnral Modo, "ariconta"
    ConseguirFoco txtAux(1), Modo
    PonFoco txtAux(1)
End Sub

Private Sub BotonVerTodos()
    ParaBusqueda False
    CadenaConsulta = GeneraSQL("")
    CargaGrid
    PonerModoUsuarioGnral 2, "ariconta"
    PonleFoco Me.wndReportControl
End Sub



Private Sub BotonModificar()
    If Not LWSeleccionado Then Exit Sub
    
    ParaBusqueda False
    CadenaDesdeOtroForm = ""
    frmCuentas.vModo = 2
    frmCuentas.CodCta = wndReportControl.SelectedRows(0).Record.Item(1).Caption
    frmCuentas.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
              
        
        CuentasParaSaldos = "concat(nommacta,'|',coalesce(fecbloq,''),'|')"
        CuentasParaSaldos = DevuelveDesdeBD(CuentasParaSaldos, "cuentas", "codmacta", CadenaDesdeOtroForm, "T")
        
        wndReportControl.SelectedRows(0).Record.Item(2).Caption = RecuperaValor(CuentasParaSaldos, 1)
        CuentasParaSaldos = RecuperaValor(CuentasParaSaldos, 2)
        If CuentasParaSaldos <> "" Then CuentasParaSaldos = Format(CuentasParaSaldos, "dd/mm/yyyy")
        wndReportControl.SelectedRows(0).Record.Item(3).Caption = CuentasParaSaldos
        CuentasParaSaldos = ""
        CadenaDesdeOtroForm = ""
        'SituaGrid CadenaDesdeOtroForm
    End If
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    ParaBusqueda False
    'Ciertas comprobaciones
    If Not LWSeleccionado Then Exit Sub
     
     
    
     
    
    
    '### a mano
    SQL = "Seguro que desea eliminar la cuenta:"
    SQL = SQL & vbCrLf & "Código: " & wndReportControl.SelectedRows(0).Record.Item(1).Caption
    SQL = SQL & vbCrLf & "Denominación: " & wndReportControl.SelectedRows(0).Record.Item(2).Caption
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        
        SQL = wndReportControl.SelectedRows(0).Record.Item(1).Caption
        Screen.MousePointer = vbHourglass
        If SepuedeEliminarCuenta(SQL) Then
        
        
            'Hay que eliminar
            Screen.MousePointer = vbHourglass
        
              
            Conn.Execute "Delete from departamentos where codmacta = " & DBSet(SQL, "T")
            Conn.Execute "Delete from cuentas where codmacta= " & DBSet(SQL, "T")
            
            Screen.MousePointer = vbHourglass
            espera 0.5
            
            SQL = ""
            NumRegElim = wndReportControl.SelectedRows(0).Record.Index
            If NumRegElim > 0 Then
                
                If NumRegElim < wndReportControl.Records.Count - 1 Then
                    'Esta a mitad
                    SQL = wndReportControl.Rows(NumRegElim + 1).Record.Item(1).Caption
                Else
                    
                    SQL = wndReportControl.Rows(NumRegElim - 1).Record.Item(1).Caption
                End If
            End If
            wndReportControl.Records.RemoveAt wndReportControl.SelectedRows(0).Record.Index
            
          wndReportControl.Populate
            
            If SQL <> "" Then SituaGrid SQL
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

Private Sub FijarError(ByRef Cad As String)
    On Error Resume Next
    Cad = Conn.Errors(0).Description
    If Err.Number <> 0 Then
        Err.Clear
        Cad = ""
    End If
End Sub



Private Sub OpcionesCambiadas()
    If txtAux(0).visible Then Exit Sub

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
        txtAux(0).Text = Trim(txtAux(0).Text)
        txtAux(1).Text = Trim(txtAux(1).Text)
        'Si estan vacios no hacemos nada
        SQL = ""
        Aux = ""
        If txtAux(0).Text <> "" Then
            If SeparaCampoBusqueda("T", "codmacta", txtAux(0).Text, Aux, False) = 0 Then SQL = Aux
        End If
        If txtAux(1).Text <> "" Then
            Aux = ""
            
            'VEO si ha puesto un *
            If InStr(1, txtAux(1).Text, "*") = 0 Then txtAux(1).Text = "*" & txtAux(1).Text & "*"
            If SeparaCampoBusqueda("T", "nommacta", txtAux(1).Text, Aux, False) = 0 Then
                If SQL <> "" Then SQL = SQL & " AND "
                SQL = SQL & Aux
            End If
        End If
        
        If txtAux(2).Text <> "" Then
            Aux = ""
            Dim Cadena As String
            
            If InStr(1, txtAux(2).Text, ">>") <> 0 Then
                Cadena = "select max(fecbloq) from cuentas "
            Else
                If InStr(1, txtAux(2).Text, "<<") <> 0 Then
                    Cadena = "select min(fecbloq) from cuentas "
                Else
                    If InStr(1, UCase(txtAux(2).Text), "=NULL") <> 0 Then
                        Aux = "(fecbloq is null)"
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & Aux
                    Else
                        If InStr(1, UCase(txtAux(2).Text), "<>NULL") <> 0 Then
                            Aux = "(not fecbloq is null)"
                            If SQL <> "" Then SQL = SQL & " AND "
                            SQL = SQL & Aux
                        End If
                    End If
                End If
            End If
            If DevuelveValor(Cadena) <> "0" Then
                txtAux(2).Text = DevuelveValor(Cadena)
            End If
            
            
            If SeparaCampoBusqueda("F", "fecbloq", txtAux(2).Text, Aux, False) = 0 Then
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
        ParaBusqueda False
        CadenaConsulta = GeneraSQL(SQL)
        CargaGrid
        Screen.MousePointer = vbDefault
        If False Then
            ParaBusqueda True
            MsgBox "Ningún resultado para la búsqueda.", vbExclamation
            Exit Sub
        Else
            
        End If
        
    End If
    ParaBusqueda False
    PonerModoUsuarioGnral Modo, "ariconta"
    'lblIndicador.Caption = ""
End Sub





'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then  ' si pulsa F2 vamos a consulta de extractos
    
        If Not LWSeleccionado Then Exit Sub
        
        Aux = DBSet(vParam.fechaini, "F") & " AND "
        Aux = Aux & DBSet(DateAdd("yyyy", 1, vParam.fechafin), "F")
        Aux = Aux & " AND codmacta  "
        'Aux = DevuelveDesdeBD("count(*)", "hlinapu", Aux, Adodc1.Recordset!codmacta, "T")
        If Val(Aux) = 0 Then
         '    MsgBox "La cuenta " & Adodc1.Recordset!codmacta & " NO tiene movimientos en las fechas", vbExclamation
        
        Else
            Set frmCExt = New frmConExtr
            
          '  frmCExt.Cuenta = Me.Adodc1.Recordset.Fields(0).Value
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
                PonleFoco Me.wndReportControl
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
        .Buttons(8).Image = 30
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
     
    
    'Poner niveles
    PonerOptionsVisibles 0
    
    'Opciones segun sea su nivel
    PonerOpcionesMenu
    
    Modo = 0
    PonerModoUsuarioGnral Modo, "ariconta"
    
    'Ocultamos busqueda
    ParaBusqueda False
    
    wndReportControl.ScrollMode = xtpReportScrollModeBlock
    
    CadAncho = False
    CadenaConsulta = GeneraSQL("")
    EstablecerFuente
    CreateReportControlPendientes
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
                Sql_BusquedaAvanzada = ""
                BotonBuscar
        Case 6
                Sql_BusquedaAvanzada = ""
                BotonVerTodos
        Case 8
                'Imprimimos el listado
                frmColCtasList.BusquedaAvanzada = Sql_BusquedaAvanzada
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
                Sql_BusquedaAvanzada = ""
                frmCuentas.vModo = 3
                frmCuentas.Show vbModal
                If CadenaDesdeOtroForm <> "" Then
                    Me.Refresh
                    Screen.MousePointer = vbHourglass
                    Sql_BusquedaAvanzada = CadenaDesdeOtroForm
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
        Case 8
            
                If wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines Then
                        wndReportControl.PaintManager.HorizontalGridStyle = xtpGridSolid
                        wndReportControl.PaintManager.VerticalGridStyle = xtpGridSolid
                Else
                    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                    wndReportControl.PaintManager.VerticalGridStyle = xtpGridNoLines
                End If
                
                For i = 1 To 10
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    espera 0.1
                Next
                wndReportControl.Populate
            Screen.MousePointer = vbDefault
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

   ' B = DataGrid1.Enabled
   ' DataGrid1.Enabled = False
    espera 0.07
    lblComprobar(0).visible = True
    lblComprobar(0).Caption = "Leyendo BD"
    lblComprobar(0).Refresh
    lblComprobar(2).visible = False
    
    HayMasDatos = True
    wndReportControl.Records.DeleteAll
    wndReportControl.EnsureFocusedRowVisible = False
    CargaGrid2 True

    

    lblComprobar(0).visible = False
        
      
    If Not CadAncho Then
        txtAux(0).Left = wndReportControl.Left + 15
        txtAux(0).Width = 1560
        txtAux(0).top = wndReportControl.top + 330 ' 235
        txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 9
        txtAux(1).Width = 6789
        txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 9
        txtAux(2).Width = 1575
        txtAux(1).top = txtAux(0).top
        txtAux(2).top = txtAux(0).top
        txtAux(0).Height = 330
        txtAux(1).Height = txtAux(0).Height
        txtAux(2).Height = txtAux(0).Height
        
         
        
        
        
        
        
        CadAncho = True
    End If
'

    'Habilitamos modificar y eliminar
    Toolbar1.Buttons(4).Enabled = vUsu.Nivel < 3
'    Toolbar1.Buttons(7).Enabled = Not Adodc1.Recordset.EOF
    If vUsu.Nivel < 2 Then
'        Toolbar1.Buttons(8).Enabled = Not Adodc1.Recordset.EOF
    Else
        Toolbar1.Buttons(8).Enabled = False
    End If
   
   
   
  
    If cboNiveles.ListIndex = vEmpresa.numnivel - 1 Then
        If Me.wndReportControl.Records.Count > 0 Then lblComprobar(2).visible = True
    End If
    
    
    
    
 
    
End Sub

Private Sub CargaGrid2(DesdeGargaGrid As Boolean)
    Screen.MousePointer = vbHourglass
    lblComprobar(0).visible = True
    lblComprobar(0).Caption = "Leyendo datos"
    lblComprobar(0).Refresh
    
    CargaGrid3 DesdeGargaGrid
    
    Screen.MousePointer = vbDefault
    lblComprobar(0).Caption = ""
End Sub


Private Sub CargaGrid3(DesdeGargaGrid As Boolean)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    Dim SQL As String
    Dim B As Boolean
    Dim Aux As String
    
   CargandoDatos = True
   
    If DesdeGargaGrid Then
        SQL = CadenaConsulta
    Else
        SQL = Me.Tag
        i = InStrRev(SQL, " ORDER BY ")
        SQL = Mid(SQL, 1, i)
        SQL = SQL & " AND codmacta >'" & UltimaCuenta & "'"
    End If
    SQL = SQL & " ORDER BY"
    If cboOrden.ListIndex = 0 Then
    'If Option2(0).Value Then
        SQL = SQL & " codmacta"
    Else
        SQL = SQL & " nommacta"
    End If
    
    SQL = SQL & " LIMIT 0," & IIf(DesdeGargaGrid, 30, 10)
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Me.Tag = SQL
    i = 0
    CuentasParaSaldos = ""
    If miRsAux.EOF Then
        HayMasDatos = False
    Else
        
        While Not miRsAux.EOF
                
                UltimaCuenta = miRsAux!codmacta
                CuentasParaSaldos = CuentasParaSaldos & ", " & DBSet(miRsAux!codmacta, "T")
                AddRecordCli
                miRsAux.MoveNext
                i = i + 1
        Wend
        J = IIf(DesdeGargaGrid, 28, 6)
        If i < J Then HayMasDatos = False
   End If
   miRsAux.Close
   
   wndReportControl.Populate
   If i > 0 Then
        
        lblComprobar(0).Caption = "Saldos"
        lblComprobar(0).Refresh
        If Len(UltimaCuenta) = vEmpresa.DigitosUltimoNivel Then
             Aux = " codmacta IN (" & Mid(CuentasParaSaldos, 2) & ")"
             SQL = "codmacta"
        Else
             Aux = " mid(codmacta,1," & Len(UltimaCuenta) & ") IN (" & Mid(CuentasParaSaldos, 2) & ")"
             SQL = " mid(codmacta,1," & Len(UltimaCuenta) & ") "
        End If
        
        SQL = "Select " & SQL & " as codmacta,sum(coalesce(timported,0)) debe, sum(coalesce(timporteh,0)) haber from hlinapu where "
        SQL = SQL & " fechaent>=  " & DBSet(vParam.fechaini, "F") & " AND " & Aux & " GROUP BY 1"
        miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        J = Me.wndReportControl.Records.Count - i
        While Not miRsAux.EOF
            For i = J To wndReportControl.Records.Count - 1
                If wndReportControl.Rows(i).Record.Item(1).Caption = miRsAux!codmacta Then
                    wndReportControl.Rows(i).Record.Item(4).Caption = Format(miRsAux!Debe, FormatoImporte)
                    wndReportControl.Rows(i).Record.Item(5).Caption = Format(miRsAux!Haber, FormatoImporte)
                    wndReportControl.Rows(i).Record.Item(6).Caption = Format(miRsAux!Debe - miRsAux!Haber, FormatoImporte)
                    Exit For
                End If
            Next
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        wndReportControl.Populate
        DoEvent2
    End If
   'utlimoIndiceTop = Me.wndReportControl.Records.Count - 24
   'If utlimoIndiceTop < 1 Then utlimoIndiceTop = 1
   
   
   Set miRsAux = Nothing
  '  utlimoIndice = Me.wndReportControl.Records.Count
   
    
    If Not DesdeGargaGrid Then
        
        If wndReportControl.Records.Count > 24 Then
            wndReportControl.Rows(wndReportControl.Records.Count - 4).EnsureVisible
            
            'If wndReportControl.Records.Count > 50 Then Debug.Print wndReportControl.Records.Count
            
        End If
        
    End If
    CargandoDatos = False
End Sub


' 0 solo textos
'1 Solo enables
'2 todo
Private Sub PonerOptionsVisibles(Opcion As Byte)
Dim i As Integer
Dim J As Integer
Dim Cad As String

End Sub



Private Function GeneraSQL(Busqueda As String) As String
Dim i As Integer
Dim SQL As String
Dim nexo As String
Dim J As Integer
Dim wildcar As String
Dim Digitos As Integer
Dim Cad As String

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
    Cad = DevuelveSQLFiltro(True)
    If Cad <> "" Then wildcar = wildcar & Cad
    
    GeneraSQL = wildcar
End Function



Private Function SepuedeEliminarCuenta(Cuenta As String) As Boolean
Dim NivelCta As Integer
Dim i, J As Integer
Dim Cad As String

    SepuedeEliminarCuenta = False
    If EsCuentaUltimoNivel(Cuenta) Then
        'ATENCION###
        ' Habra que ver casos particulares de eliminacion de una subcuenta de ultimo nivel
        'Si esta en apuntes, en ....
        'NO se puede borrar
        lblComprobar(0).Caption = "Comprobando"
        lblComprobar(0).visible = True
        Cad = BorrarCuenta(Cuenta, lblComprobar(0))
        lblComprobar(0).visible = False
        If Cad <> "" Then
            Cad = Cuenta & vbCrLf & Cad
            MsgBox Cad, vbExclamation
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
            Cad = DevuelveDesdeBD("codmacta", "ctaagrupadas", "codmacta", Cuenta, "T")
            If Cad <> "" Then
                MsgBox "El subnivel pertenece a agrupacion de cuentas en balance"
                Exit Function
            End If
        End If
        For J = NivelCta + 1 To vEmpresa.numnivel
            Cad = Cuenta & "__________"
            i = DigitosNivel(J)
            Cad = Mid(Cad, 1, i)
            If TieneEnBD(Cad) Then
                MsgBox "Tiene cuentas en niveles superiores (" & J & ")", vbExclamation
                Exit Function
            End If
        Next J
    End If
    SepuedeEliminarCuenta = True
End Function

Private Function TieneEnBD(Cad As String) As String
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select codmacta from cuentas where codmacta like '" & Cad & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TieneEnBD = Not Rs.EOF
    Rs.Close
    Set Rs = Nothing
End Function


Private Sub SituaGrid(Cadena As String)
On Error GoTo ESituaGrid

   For i = 0 To Me.wndReportControl.Rows.Count - 1
        If wndReportControl.Rows(i).Record.Item(1).Caption = Cadena Then
            
             wndReportControl.Rows(i).Selected = True
             wndReportControl.Rows(i).EnsureVisible
             Exit For
        End If
    Next
Exit Sub
ESituaGrid:
    MuestraError Err.Number, "Situando registro activo"
End Sub


Private Sub ParaBusqueda(Ver As Boolean)
txtAux(0).visible = Ver
txtAux(1).visible = Ver
txtAux(2).visible = Ver
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
    ConseguirFoco txtAux(Index), 1
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub ComprobarCuentas()
Dim Cad As String
Dim N As Integer
Dim i As Integer
Dim Col As Collection
Dim C1 As String

On Error GoTo EComprobarCuentas
    
    
    'NO hay cuentas
    If Not LWSeleccionado Then Exit Sub
    
    'Buscando datos
    If txtAux(0).visible Then Exit Sub
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
            If ObtenerCuenta(Cad, i, N) Then
                lblComprobar(1).Caption = Cad
                lblComprobar(1).Refresh
                ComprobarCuenta Cad, i, Col
            End If
        Loop Until Cad = ""
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
            Cad = Dir("C:\WINDOWS\NOTEPAD.exe")
            If Cad = "" Then
                Cad = Dir("C:\WINNT\NOTEPAD.exe")
            End If
            If Cad = "" Then
                MsgBox "Se ha producido errores. Vea el archivo Errorcta.txt"
                Else
                Shell Cad & " " & App.Path & "\Errorcta.txt", vbMaximizedFocus
            End If
            espera 2
    
            If vUsu.Nivel < 2 Then
                If MsgBox("Desea crear los subniveles?", vbQuestion + vbYesNo) = vbYes Then
                        
                        Cad = "insert into `cuentas` (`codmacta`,`nommacta`,`apudirec`,dirdatos) VALUES ('"
                        For NF = 1 To Col.Count
                            N = DigitosNivelAnterior(Col.Item(NF))
                            If N > 0 Then
                                C1 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Mid(Col.Item(NF), 1, N), "T")
                            Else
                                C1 = ""
                            End If
                            If C1 = "" Then C1 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Col.Item(NF), "T")
                            
                            If C1 = "" Then C1 = "AUTOM: " & Col.Item(NF)
    
                            EjecutaSQL Cad & Col.Item(NF) & "','" & DevNombreSQL(C1) & "','N','AUTOMATICA en la comprobacion')"
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


Private Function ObtenerCuenta(ByRef Cadena As String, Nivel As Integer, ByRef Digitos As Integer) As Boolean
Dim RT As Recordset
Dim SQL As String


If Cadena = "" Then
    SQL = ""
Else
    SQL = DevuelveUltimaCuentaGrupo(Cadena, Nivel, Digitos)
    SQL = " codmacta > '" & SQL & "' AND "
End If
SQL = "Select codmacta from Cuentas WHERE " & SQL
SQL = SQL & " codmacta like '" & Mid("__________", 1, Digitos) & "'"

Set RT = New ADODB.Recordset
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RT.EOF Then
    ObtenerCuenta = False
    Cadena = ""
Else
    ObtenerCuenta = True
    Cadena = RT!codmacta
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
Dim Cad As String
Dim N As Integer
N = DigitosNivel(Nivel - 1)
Cad = Mid(Cta, 1, N)
Cad = Cad & "9999999999"
DevuelveUltimaCuentaGrupo = Mid(Cad, 1, Digitos)
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

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub

    If Index = 0 Then
        Aux = txtAux(0).Text
        If CuentaCorrectaUltimoNivel(Aux, "") Then txtAux(0).Text = Aux
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
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) 'And (DatosADevolverBusqueda = "")
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!Especial, "N") And ConfigurarBalances <> 2 And ConfigurarBalances <> 6 'And (DatosADevolverBusqueda = "")
        Me.Toolbar2.Buttons(6).Enabled = DBLet(Rs!Especial, "N") 'And (DatosADevolverBusqueda = "")
        Me.Toolbar2.Buttons(7).Enabled = DBLet(Rs!Especial, "N") 'And (DatosADevolverBusqueda = "")
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Function LWSeleccionado() As Boolean
    LWSeleccionado = False
    If Me.wndReportControl.Records.Count = 0 Then Exit Function
    If Me.wndReportControl.SelectedRows.Count = 0 Then Exit Function
    LWSeleccionado = True
End Function



Public Sub CreateReportControlPendientes()
    'Start adding columns
    Dim Column As ReportColumn
    wndReportControl.Columns.DeleteAll
    
    'Chekc
     
    
    
        
        
    
        
        Set Column = wndReportControl.Columns.Add(1, "Cuenta", 10, True)
        Set Column = wndReportControl.Columns.Add(2, "Nombre", 43, True)
        Set Column = wndReportControl.Columns.Add(3, "Bloqueada", 10, True)
        
        Set Column = wndReportControl.Columns.Add(4, "Debe", 12, True)
        Column.Icon = 12
        Column.Alignment = xtpAlignmentRight
        
        Set Column = wndReportControl.Columns.Add(5, "Haber", 12, True)
        Column.Icon = 12
        Column.Alignment = xtpAlignmentRight
        
        Set Column = wndReportControl.Columns.Add(6, "Saldo", 12, True)
        Column.Icon = 12
        Column.Alignment = xtpAlignmentRight
        
    
     
    

    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
    wndReportControl.AllowSort = False
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    Set fntStrike = wndReportControl.PaintManager.TextFont
    fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    Set fntBold = wndReportControl.PaintManager.TextFont
    fntBold.Bold = True
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
    wndReportControl.ZOrder 1
End Sub




Private Sub AddRecordCli()
Dim SQL As String
Dim Record As ReportRecord
Dim ItemToolTip As String
Dim ItemIcon As Integer
Dim Color As Long
    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
   
    Dim Item As ReportRecordItem
    
          
    ItemIcon = RECORD_REPLIED_ICON
    ItemToolTip = ""

        'Check
    Set Item = Record.AddItem("")
    
        Item.Checked = False
        Item.HasCheckbox = False
    
    
    
    Set Item = Record.AddItem("")
    Item.Icon = ItemIcon
    Item.ToolTip = ItemToolTip
    Item.Caption = CStr(miRsAux!codmacta)
    
'    Sql = miRsAux!NUmSerie & Format(miRsAux!NumFactu, "0000000")
    Set Item = Record.AddItem(CStr(miRsAux!Nommacta))
    
    
    
    SQL = " "
    If Not IsNull(miRsAux!bloqueada) Then SQL = Format(miRsAux!bloqueada, "dd/mm/yyyy")
    Set Item = Record.AddItem(SQL)

    
    
    Set Item = Record.AddItem("")
    Item.Caption = ""
    Set Item = Record.AddItem("")
    Item.Caption = ""
    Set Item = Record.AddItem("")
    Item.Caption = ""
    
    Record.Tag = ""
                
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    'Record.PreviewText = "ID: "
    
        
    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
End Sub

Private Sub EstablecerFuente()

    On Error GoTo eEstablecerFuente
    'The following illustrate how to change the different fonts used in the ReportControl
    Dim TextFont As StdFont
    Set TextFont = Me.Font
    TextFont.SIZE = 9
    Set wndReportControl.PaintManager.TextFont = TextFont
    Set wndReportControl.PaintManager.CaptionFont = TextFont
    Set wndReportControl.PaintManager.PreviewTextFont = TextFont
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    'Set fntStrike = wndReportControl.PaintManager.TextFont
    'fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    'Set fntBold = wndReportControl.PaintManager.TextFont
    'fntBold.Bold = True


    Exit Sub
eEstablecerFuente:
    MuestraError Err.Number, Err.Description

End Sub

Private Sub wndReportControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then  ' si pulsa F2 vamos a consulta de extractos
    
        If Not LWSeleccionado Then Exit Sub
        
        If Len(wndReportControl.SelectedRows(0).Record.Item(1).Caption) <> vEmpresa.DigitosUltimoNivel Then
            MsgBoxA "Extracto solo cuentas de ultimo nivel", vbExclamation
            Exit Sub
        End If
            
        Aux = DBSet(vParam.fechaini, "F") & " AND "
        Aux = Aux & DBSet(DateAdd("yyyy", 1, vParam.fechafin), "F")
        Aux = Aux & " AND codmacta  "
        Aux = DevuelveDesdeBD("count(*)", "hlinapu", Aux, wndReportControl.SelectedRows(0).Record.Item(1).Caption, "T")
        If Val(Aux) = 0 Then
             MsgBox "La cuenta " & wndReportControl.SelectedRows(0).Record.Item(1).Caption & " NO tiene movimientos en las fechas", vbExclamation
        
        Else
            Set frmCExt = New frmConExtr
            
            frmCExt.Cuenta = wndReportControl.SelectedRows(0).Record.Item(1).Caption
            frmCExt.Show vbModal
            
            Set frmCExt = Nothing
        End If
        KeyCode = 0
        Exit Sub
    End If
End Sub

Private Sub wndReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   
    
    CadenaDesdeOtroForm = ""
    frmCuentas.vModo = 0
    frmCuentas.CodCta = Row.Record.Item(1).Caption
    frmCuentas.Show vbModal
    
    
End Sub

Private Sub wndReportControl_VScroll(ByVal Section As Long, ByVal Position As Long)
Dim N As Long

    'Debug.Print Now()

    If CargandoDatos Then Exit Sub
    N = Me.wndReportControl.Records.Count - (Position + 24)
    If N < 4 Then
        If HayMasDatos Then
            'Debug.Print Position
            CargaGrid2 False
        End If
    Else
        
    End If
   
End Sub

