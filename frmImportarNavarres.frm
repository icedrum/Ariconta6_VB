VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImportarNavarres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar facturas CONSUM"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   12195
   Icon            =   "frmImportarNavarres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   12015
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   22
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver todos"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Conceptos"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   " "
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Importar fichero"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
         Begin VB.CheckBox chkVistaPrevia 
            Caption         =   "Vista previa"
            Height          =   195
            Left            =   9840
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1109
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3440
      EndProperty
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   320
      Index           =   2
      Left            =   5280
      TabIndex        =   14
      Top             =   7560
      Width           =   135
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   320
      Index           =   1
      Left            =   3360
      TabIndex        =   11
      Top             =   7560
      Width           =   135
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      Height          =   320
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   5640
      Width           =   135
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Centro|N|N|0||importnavconcepcentro|codconcepto|0|S|"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9780
      TabIndex        =   4
      Top             =   8820
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10980
      TabIndex        =   5
      Top             =   8820
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "Centro|N|N|0||importnavconcepcentro|codcentro|0|S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10980
      TabIndex        =   8
      Top             =   8820
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   8640
      Width           =   4665
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4335
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   5970
      Top             =   0
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
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   3
      Left            =   3420
      MaxLength       =   30
      TabIndex        =   12
      Text            =   "Dato2"
      Top             =   7560
      Width           =   1395
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   5
      Left            =   5520
      MaxLength       =   30
      TabIndex        =   13
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   4320
      TabIndex        =   3
      Tag             =   "Centro|T|N|||importnavconcepcentro|codmacta|||"
      Text            =   "Dat"
      Top             =   7560
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmImportarNavarres.frx":000C
      Height          =   5505
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   9710
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Index           =   1
      Left            =   3120
      TabIndex        =   18
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod."
         Object.Width           =   1138
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Concepto"
         Object.Width           =   4789
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Index           =   2
      Left            =   7080
      TabIndex        =   19
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2963
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Centro"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fichero"
         Object.Width           =   4066
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Histórico ficheros importados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   195
      Index           =   3
      Left            =   7200
      TabIndex        =   22
      Top             =   600
      Width           =   2490
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cuentas por concepto y centro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   2595
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Conceptos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   20
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Centros CONSUM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cuentas por concepto y centro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   23
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   2595
   End
End
Attribute VB_Name = "frmImportarNavarres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents frmB As frmBuscaGrid
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBasico2
Attribute frmB.VB_VarHelpID = -1

Private CadenaConsulta As String
Private TextoBusqueda As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim jj As Long
Dim Sql As String

Dim IT As ListItem

Dim ImporteFacturaImportada As Currency
Dim FechaFraImportada As Date
Dim NumFacturaFichero As String
Dim TipoFichero  As Byte  '0. FC   1.- FA  2.-FV    Ver ImportarFactura

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
    Modo = vModo
    
    B = (Modo = 0)
    
    For jj = 0 To 5
        txtAux2(jj).visible = Not B
        If jj < 3 Then cmdAux(jj).visible = Not B
    Next jj
    
    Toolbar1.Buttons(1).Enabled = B
    Toolbar1.Buttons(2).Enabled = B
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    'DataGrid1.Enabled = b
    
    'Si es regresar
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.Visible = b
'    End If
    'Si estamo mod or insert
    If Modo = 2 Then
       txtAux2(0).BackColor = &H80000018
       txtAux2(2).BackColor = &H80000018
       Else
        txtAux2(0).BackColor = &H80000005
        txtAux2(2).BackColor = &H80000005
    End If
    txtAux2(0).Enabled = (Modo <> 2)
    txtAux2(0).Enabled = (Modo <> 2)
    txtAux2(2).Enabled = txtAux2(0).Enabled
    txtAux2(2).BackColor = txtAux2(0).BackColor
    cmdAux(0).Enabled = txtAux2(0).Enabled
    cmdAux(1).Enabled = txtAux2(0).Enabled
    
End Sub

Private Sub BotonAnyadir()
    Dim anc As Single

    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Not adodc1.Recordset.EOF Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.top + 210
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top
    End If
    txtAux2(0).Text = ""
    For jj = 1 To 5
        txtAux2(jj).Text = ""
    Next jj
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtAux2(0).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub



Private Sub BotonVerTodos()
    DataGrid1.Enabled = False
    espera 0.1
    TextoBusqueda = ""
    CargaGrid ""
    DataGrid1.Enabled = True
End Sub

Private Sub BotonBuscar()
    DataGrid1.Enabled = False
    CargaGrid " importnavconcepcentro.codcentro = -1"
    DataGrid1.Enabled = True
    'Buscar
    For jj = 0 To 5
        txtAux2(jj).Text = ""
    Next jj
    LLamaLineas DataGrid1.top + 206, 2
    txtAux2(0).SetFocus
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim cad As String
    Dim anc As Single
    Dim I As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    'If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.top  '320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top
    End If

    'Llamamos al form
    For jj = 0 To 5
        txtAux2(jj).Text = DataGrid1.Columns(jj).Text
    Next jj

    
    LLamaLineas anc, 1

   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid
    PonerModo xModo + 1
    'Fijamos el ancho
    For jj = 0 To 5
        txtAux2(jj).top = alto
        If jj < 3 Then cmdAux(jj).top = alto
    Next jj
    
End Sub




Private Sub BotonEliminar()
Dim Sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    '### a mano
    Sql = "Seguro que desea eliminar la linea de histórico:" & vbCrLf
    Sql = Sql & vbCrLf & "Codigo: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Centro: " & adodc1.Recordset.Fields(1)
    Sql = Sql & vbCrLf & "Seccion       : " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Nombre  : " & adodc1.Recordset.Fields(3)
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from importnavconcepcentro where codcentro=" & adodc1.Recordset!codcentro
        Sql = Sql & " AND codconcepto =" & adodc1.Recordset!codconcepto & ";"
        Conn.Execute Sql
        CargaGrid ""
        adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                Conn.Execute "commit"
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    Conn.Execute "commit"
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        
        'Para el texto
        TextoBusqueda = ""
        If txtAux2(0).Text <> "" Then TextoBusqueda = TextoBusqueda & "Cod. Inmov " & txtAux2(0).Text
        If txtAux2(2).Text <> "" Then TextoBusqueda = TextoBusqueda & "Fecha " & txtAux2(2).Text
        If txtAux2(3).Text <> "" Then TextoBusqueda = TextoBusqueda & "Porcentaje " & txtAux2(3).Text
        
            
        
        If CadB <> "" Then
            PonerModo 0
            DataGrid1.Enabled = False
            CargaGrid CadB
            DataGrid1.Enabled = True
        End If
    End Select


End Sub

Private Sub cmdAux_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    If Index = 2 Then
        'Cuentas
        Set frmC = New frmColCtas
        frmC.DatosADevolverBusqueda = "0|1"
        frmC.ConfigurarBalances = 3  'NUEVO
        Sql = ""
        frmC.Show vbModal
        Set frmC = Nothing
        jj = 4
    Else
        Sql = ""
        Set frmB = New frmBasico2
        
        
        If Index = 0 Then
            jj = 0
            AyudaImporNavarresCentro frmB
        Else
            AyudaImporNavarresSeccion frmB
            jj = 2
        End If
        Set frmB = Nothing
    End If
    If Sql <> "" Then
        txtAux2(jj).Text = RecuperaValor(Sql, 1)
        txtAux2(jj + 1).Text = RecuperaValor(Sql, 2)
    End If
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
TextoBusqueda = ""
DataGrid1.SetFocus
End Sub

'Private Sub cmdRegresar_Click()
'Dim cad As String
'
'If adodc1.Recordset.EOF Then
'    MsgBox "Ningún registro a devolver.", vbExclamation
'    Exit Sub
'End If
'
'cad = adodc1.Recordset.Fields(0) & "|"
'cad = cad & adodc1.Recordset.Fields(1) & "|"
'cad = cad & adodc1.Recordset.Fields(2) & "|"
'RaiseEvent DatoSeleccionado(cad)
'Unload Me
'End Sub

Private Sub DataGrid1_DblClick()
'If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If ListView1(0).ListItems.Count = 0 Then
        CargaCentros
        CargaConceptos
        CargaHistorico
    End If
End Sub


Private Sub Form_Load()
    Me.Icon = frmppal.Icon
          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.ImgListComun
        '.Buttons(1).Image = 1
        '.Buttons(2).Image = 2
        '.Buttons(6).Image = 3
        '.Buttons(7).Image = 4
        '.Buttons(8).Image = 5
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(6).Image = 1
        .Buttons(7).Image = 2
        '.Buttons(8).Image = 2
        
        
        .Buttons(10).Image = 12
        .Buttons(12).Image = 13   '11
        
        .Buttons(17).Image = 15
        
        '.Buttons(14).Image = 6
        '.Buttons(15).Image = 7
        '.Buttons(16).Image = 8
        '.Buttons(17).Image = 9
    End With

    Set miTag = New CTag
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
'    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    
    'cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    CadenaConsulta = "select  importnavconcepcentro.codcentro,importnavcentros.Descripcion ElCentro, codconcepto,"
    CadenaConsulta = CadenaConsulta & " importnavconceptos.descripcion,importnavconcepcentro.codmacta,nommacta"
    CadenaConsulta = CadenaConsulta & " from importnavconcepcentro inner join cuentas on importnavconcepcentro.codmacta=cuentas.codmacta"
    CadenaConsulta = CadenaConsulta & " left join importnavcentros on importnavconcepcentro.codcentro=importnavcentros.codcentro"
    CadenaConsulta = CadenaConsulta & " left join importnavconceptos on importnavconceptos.concepto= importnavconcepcentro.codconcepto WHERE 1=1 "
    
    CargaGrid "importnavconcepcentro.codcentro >0 "
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Set miTag = Nothing
End Sub




Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Sql = CadenaDevuelta
End Sub

Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
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
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    BotonAnyadir
        
Case 2
      BotonModificar
        
Case 3
    BotonEliminar
Case 6
        BotonBuscar
Case 7
        BotonVerTodos
Case 8
        
        
    
Case 10
        CadenaDesdeOtroForm = ""
        frmImportarNavConcep.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            Screen.MousePointer = vbHourglass
            CadenaDesdeOtroForm = ""
            CargaConceptos
            BotonVerTodos
            Screen.MousePointer = vbDefault
        End If
Case 12

        ImportarFactura
Case 17
        Unload Me
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim I
    For I = 19 To 22
        Toolbar1.Buttons(I).visible = bol
    Next I
End Sub

Private Sub CargaGrid(Optional vSql As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    
    adodc1.ConnectionString = Conn
    If vSql <> "" Then
        Sql = CadenaConsulta & " AND " & vSql
        Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " ORDER BY  importnavconcepcentro.codcentro,codconcepto"
    adodc1.RecordSource = Sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
        
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Centro"
        DataGrid1.Columns(I).Width = 700

    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Nom. centro"
        DataGrid1.Columns(I).Width = 2100
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    'El importe es campo calculado
    
    I = 2
        DataGrid1.Columns(I).Caption = "Conce."
        DataGrid1.Columns(I).Width = 800

    
    'Leemos del vector en 2
    I = 3
        DataGrid1.Columns(I).Caption = "Descripcion"
        DataGrid1.Columns(I).Width = 2800
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    
    I = 4
        DataGrid1.Columns(I).Caption = "Cuenta"
        DataGrid1.Columns(I).Width = 1150

    
    'Leemos del vector en 2
    I = 5
        DataGrid1.Columns(I).Caption = "Nom. cuenta"
        DataGrid1.Columns(I).Width = 3450
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    For I = 0 To 5
        DataGrid1.Columns(I).AllowSizing = False
    Next I
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux2(0).Left = DataGrid1.Left + 340
        txtAux2(0).Width = DataGrid1.Columns(0).Width - 60
        cmdAux(0).Left = DataGrid1.Columns(1).Left - 15
        txtAux2(1).Left = cmdAux(0).Left + cmdAux(0).Width
        txtAux2(1).Width = DataGrid1.Columns(1).Width - cmdAux(0).Width + 30
        txtAux2(2).Left = DataGrid1.Columns(2).Left + 90
        txtAux2(2).Width = DataGrid1.Columns(2).Width - 15
        cmdAux(1).Left = DataGrid1.Columns(3).Left
        txtAux2(3).Left = DataGrid1.Columns(3).Left + 90
        txtAux2(3).Width = DataGrid1.Columns(3).Width
        txtAux2(4).Left = DataGrid1.Columns(4).Left + 120
        txtAux2(4).Width = DataGrid1.Columns(4).Width - 15
        cmdAux(2).Left = DataGrid1.Columns(5).Left - 30
        txtAux2(5).Left = DataGrid1.Columns(5).Left + 60
        txtAux2(5).Width = DataGrid1.Columns(5).Width - 15
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
   
End Sub

Private Sub txtaux2_GotFocus(Index As Integer)
    With txtAux2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtaux2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtAux2_LostFocus(Index As Integer)

    txtAux2(Index).Text = Trim(txtAux2(Index).Text)
    If Modo = 3 Then Exit Sub 'Busquedas
    
    If txtAux2(Index).Text = "" Then
        txtAux2(Index + 1).Text = ""
        Exit Sub
    End If
    
    Select Case Index
    Case 0, 2
        If Modo >= 2 And Index < 4 Then Exit Sub
        Sql = ""
        If Not IsNumeric(txtAux2(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            
        Else
            If Index = 0 Then
                'importnavcentros CodCentro Descripcion
                Sql = DevuelveDesdeBD("Descripcion", "importnavcentros", "CodCentro", txtAux2(Index).Text, "N")
            Else
                'importnavconceptos concepto  descripcion
                Sql = DevuelveDesdeBD("Descripcion", "importnavconceptos", "concepto", txtAux2(Index).Text, "N")
            End If
        
            If Sql = "" Then MsgBox "Ningun valore en la BD para el campo : " & txtAux2(Index).Text, vbExclamation
            
            
        End If
        txtAux2(Index + 1).Text = Sql
        If Sql = "" Then
            txtAux2(Index).Text = ""
            txtAux2(Index).SetFocus
        End If
        
    Case 4
        TextoBusqueda = txtAux2(Index).Text
        If CuentaCorrectaUltimoNivel(TextoBusqueda, Sql) Then
            txtAux2(Index).Text = TextoBusqueda
            txtAux2(Index + 1).Text = Sql

        Else
            MsgBox Sql, vbExclamation
            txtAux2(Index).Text = ""
            txtAux2(Index + 1).Text = ""
            txtAux2(Index).SetFocus
        End If
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 1 Then
        'Estamos insertando
        
    End If
    DatosOK = B
End Function

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function




Private Sub CargaCentros()
    Set miRsAux = New ADODB.Recordset
    Sql = "Select * FROM importnavcentros order by  CodCentro  "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1(0).ListItems.Add
        IT.Text = miRsAux.Fields(0)
        IT.SubItems(1) = miRsAux.Fields(1)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


Private Sub CargaConceptos()
    Set miRsAux = New ADODB.Recordset
    ListView1(1).ListItems.Clear
    Sql = "select * from importnavconceptos order by concepto  "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1(1).ListItems.Add
        IT.Text = miRsAux.Fields(0)
        IT.SubItems(1) = miRsAux.Fields(1)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub CargaHistorico()
    Set miRsAux = New ADODB.Recordset
    Sql = "select fecha ,centro,Fichero  from importnavhcofich WHERE fecha >= " & DBSet(DateAdd("m", -6, Now), "F") & " order by fecha desc  "
    ListView1(2).ListItems.Clear
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1(2).ListItems.Add
        IT.Text = miRsAux.Fields(0)
        IT.SubItems(1) = miRsAux.Fields(1)
        IT.SubItems(2) = miRsAux.Fields(2)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub


























'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'
'
'   proceso de importacion
'
'
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************


Private Function ImportarFactura_(Fichero As String, ByRef L As Label) As Boolean
Dim B As Boolean
    
    

    
    Screen.MousePointer = vbHourglass
    ImportarFactura_ = False
    
    
    L.Caption = "Procesando"
    L.Refresh
    
    B = ProcesarFichero(Fichero)
    
    'Si el proceso ha sido correcto, veremos si todos los datos que vienen en el fichero
    'son corectos. Existen IVAs, existen centros, existen conceptos.....
    L.Caption = "Comprobar"
    L.Refresh
    If B Then B = ComprobarDatosFichero

    
    
    'Calculamos las abses por seccion
    L.Caption = "Calculos"
    L.Refresh
    If B Then CalculoDeBases
    
    If B Then
        jj = InStrRev(Fichero, "\")
        Fichero = Mid(Fichero, jj + 1)
        
        Sql = "UPDATE importanatmptotal SET Fichero  = '" & Fichero & "'"
        Conn.Execute Sql
    End If
    
    
    ImportarFactura_ = B
    
    Screen.MousePointer = vbDefault
End Function

Private Function ProcesarFichero(Ficher As String) As Boolean
Dim NF As Integer
Dim Linea As String
Dim TrozoComun As String  'TODAS LAS LINEAS LO LLEVAN
Dim LenTrozoComun As Integer
Dim InicioLineas As String
Dim ColLineas As Collection
'Dim K As Long
Dim TextoAProcesar As String
Dim OtraCadena2 As String
Dim LineaVacia As Boolean
Dim Depuracion As Boolean
Dim NF2 As Integer
Dim LongitudLinea As Integer

    On Error GoTo EProcesarFichero
    ProcesarFichero = False
    
    'Borramos la temporal
    Conn.Execute "DELETE FROM  importnavtmp"
    Conn.Execute "DELETE FROM  importanatmptotal"
    
    NF = FreeFile
    
    Open Ficher For Input As #NF
    
    'Primer registro CABECERA
    If False Then Line Input #NF, Linea
    
    
    Depuracion = False
    NF2 = -1
    If Dir(App.Path & "\depuraF.txt", vbArchive) <> "" Then Depuracion = True
    If Depuracion Then
        NF2 = NF + 1
        Open App.Path & "\navarr.dat" For Output As #NF2
    End If
    
    'Separamos en lineas
    'En los .dat vienen una unica linea
    Set ColLineas = New Collection
    TextoAProcesar = ""
    
    LongitudLinea = 220
    If TipoFichero > 0 Then LongitudLinea = IIf(TipoFichero = 1, 290, 200)  '0. FC   1.- FA  2.-FV
    
    If UCase(Right(Ficher, 3)) = "DAT" Then
        
        
        
        
         While Not EOF(NF)
                Line Input #NF, Sql
                If Len(Sql) > LongitudLinea Then Err.Raise 513, , "Longitud linea cabecera incorrecta " & Sql
                Sql = Mid(Sql & Space(LongitudLinea), 1, LongitudLinea)
                Linea = Linea & Sql
         Wend
        'FICHERO .dat
        Sql = ""
        While Linea <> ""
            If Len(Linea) < LongitudLinea Then
                Err.Raise 513, , "Longitud linea cabecera incorrecta: " & ColLineas.Count
            Else
                Sql = Mid(Linea, 1, LongitudLinea)
                ColLineas.Add Sql
                If Depuracion Then
                    'If Mid(Linea, 66, 2) = "02" Then Print #NF2, SQL
                     Print #NF2, Sql
                End If
                Linea = Mid(Linea, LongitudLinea + 1)
            End If
                
        Wend
        Sql = ""
    Else
        'If UCase(Right(Ficher, 3)) = "RTF" Then
            
            Err.Raise 513, , "Solo formato DAT"
                        
                        
         'Lo que habia para los rtf lo he puesto bajo comentado
        'End If
    End If
    
    Close #NF
    NF = -1
    If Depuracion Then Close #NF2: NF2 = -1
    If ColLineas.Count = 0 Then Exit Function
    
    
    'Trozo comun a todas las lineas
    'Sera los primeros "J" caracteres
    LenTrozoComun = 14  'para fc
    If TipoFichero > 0 Then LenTrozoComun = IIf(TipoFichero = 1, 22, 23)   '0. FC   1.- FA  2.-FV
    
    '38220500052453  EN FCs
    Linea = ColLineas(1)
    TrozoComun = Mid(Linea, 1, LenTrozoComun)
    
    'Para los insertes de las lineas
    If TipoFichero = 0 Then
        Sql = Mid(Linea, 17, 8)
        InicioLineas = ", ('" & Mid(Linea, 2, 3) & "','" & Mid(Linea, 5, 10) & "','"
    
    Else
        '  1.- FA  2.-FV    Ver ImportarFactura
        If TipoFichero = 1 Then
            InicioLineas = ", ('" & Mid(Linea, 2, 3) & "','" & Mid(Linea, 5, 10) & "',"
            Sql = Mid(Linea, 15, 8)
        Else
            Sql = Mid(Linea, 16, 8)
            InicioLineas = ", ('" & Mid(Linea, 2, 3) & "','" & Mid(Linea, 5, 10) & "',"
        End If
    End If
    InicioLineas = InicioLineas & Mid(Sql, 1, 4) & "-" & Mid(Sql, 5, 2) & "-" & Mid(Sql, 8, 2) & "',"
    
    
    'importnavtmp(tienda,numfac,fechafac,secuencial,articulo,area,seccion,subseccion,grupo,subgrupo,
    ' Longitud      3       10     10                   8      2     2        2         2       2
    
    'precioventa,importeventa,precosteud,imporcoste,porceniva,porcenrecequiv, precosteiva1   impcosteiva signo )
    '    8           10          9           11          4           4           11                 11     1
    
    'Lineas
   
    TextoBusqueda = ""
    For jj = 2 To ColLineas.Count
        
        Linea = ColLineas(jj)
        If Len(Linea) <> LongitudLinea Then Err.Raise 513, , "Longitud linea incorrecta. Linea " & jj & " - " & Len(Linea)
        
        
        'Trozo comun
        If Mid(Linea, 1, LenTrozoComun) <> TrozoComun Then Err.Raise 513, , "Inicio linea erroneo " & Linea & " [" & TrozoComun & "]"
        
        
        LineaVacia = False
        Sql = Mid(Linea, 29)
        If Trim(Sql) = "" Then
           ' If Depuracion Then Print #NF2, "Linea vacia (" & jj & ")" & Linea
            LineaVacia = True
        End If
        
        
        
        'Montamos el INSERT
        'tienda,numfac,fechafac,secuencial
        Sql = InicioLineas & jj - 1 & ","
        OtraCadena2
                
                
        Select Case TipoFichero
        Case 0
            '                       _________
            'LO que habia. Facturas  CENTRAL
            'articulo,area,seccion,subseccion,grupo,subgrupo
            Sql = Sql & "'" & Mid(Linea, 56, 8) & "','" & Mid(Linea, 64, 2) & "','" & Mid(Linea, 66, 2) & "','"
            Sql = Sql & Mid(Linea, 68, 2) & "','" & Mid(Linea, 70, 2) & "','" & Mid(Linea, 72, 2) & "','"
            
            
            'cajaformser,cajaformserDec,udformato  -> de ahi extraeremos las unidades
            'If Mid(Linea, 80, 3) <> "000" Then S top
            Sql = Sql & Mid(Linea, 74, 6) & "','" & Mid(Linea, 80, 3) & "','" & Mid(Linea, 82, 6) & "','"
            
            'precioventa,importeventa,precosteud,
            Sql = Sql & Mid(Linea, 88, 8) & "','" & Mid(Linea, 96, 10) & "','" & Mid(Linea, 106, 9) & "','"
            
            
            'imporcoste,porceniva,porcenrecequiv
            Sql = Sql & Mid(Linea, 115, 11) & "','" & Mid(Linea, 126, 4) & "','" & Mid(Linea, 130, 4) & "','"
            
            'precosteiva1   impcosteiva ,signo
            Sql = Sql & Mid(Linea, 134, 11) & "','" & Mid(Linea, 145, 11) & "','" & Mid(Linea, 205, 1)
            
            
            
            Sql = Sql & "','" & oTRAcADENA & "',null)"
        
        Case 1
            '          ______________
            'Facturas  ADMINISTRACION
            'articulo,area,seccion,subseccion,grupo,subgrupo
            Sql = Sql & "null,null,100,null,null,null,null"
            
            
            'cajaformser,cajaformserDec,udformato  -> de ahi extraeremos las unidades
            'If Mid(Linea, 80, 3) <> "000" Then S top
            Sql = Sql & Mid(Linea, 74, 6) & "','" & Mid(Linea, 80, 3) & "','" & Mid(Linea, 82, 6) & "','"
            
            'precioventa,importeventa,precosteud,
            Sql = Sql & Mid(Linea, 88, 8) & "','" & Mid(Linea, 96, 10) & "','" & Mid(Linea, 106, 9) & "','"
            
            
            'imporcoste,porceniva,porcenrecequiv
            Sql = Sql & Mid(Linea, 115, 11) & "','" & Mid(Linea, 126, 4) & "','" & Mid(Linea, 130, 4) & "','"
            
            'precosteiva1   impcosteiva ,signo
            Sql = Sql & Mid(Linea, 134, 11) & "','" & Mid(Linea, 145, 11) & "','" & Mid(Linea, 205, 1)
            
            
            
            Sql = Sql & "','" & oTRAcADENA & "')"
        
            
        
        
        
        Case 2
            '         ______________
            'Facturas   AUTOVENTAS
        
        
        
        End Select
        
        
        If Not LineaVacia Then TextoBusqueda = TextoBusqueda & Sql
        
        If jj = ColLineas.Count Then
            Sql = ""  'el ultimo
        Else
            If Len(TextoBusqueda) > 2000 Then Sql = ""
        End If
        If Sql = "" Then
            TextoBusqueda = Mid(TextoBusqueda, 2) 'quitamos la primera cma
            Sql = "INSERT INTO importnavtmp(tienda,numfac,fechafac,secuencial,articulo,area,seccion,subseccion,grupo,subgrupo,cajaformser,cajaformserDec,udformato,"
            Sql = Sql & "precioventa,importeventa,precosteud,imporcoste,porceniva,porcenrecequiv,precosteiva1,impcosteiva,signo,unidades,observaciones) VALUES "
            Sql = Sql & TextoBusqueda
            
            Conn.Execute Sql
            TextoBusqueda = ""
        End If
        
        
    Next jj


    If jj > 0 Then
        ProcesarFichero = True
    Else
        MsgBox "Ningun dato en el fichero", vbExclamation
    End If



EProcesarFichero:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    If NF > 0 Then Close #NF
    If NF2 > 0 Then Close #NF2
End Function



'Comprobamos los datos
'Centros, conceptos, cuentas , ivas
Private Function ComprobarDatosFichero() As Boolean
Dim Porcen As Currency
Dim Fin As Boolean
    
    On Error GoTo eComprobarDatosFichero
    ComprobarDatosFichero = False
    
    Set miRsAux = New ADODB.Recordset
    
    'Centros / Tiendas
    Sql = "select distinct(tienda) from importnavtmp where  not tienda in (select CodCentro from importnavcentros)"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not miRsAux.EOF
        Sql = Sql & miRsAux.Fields(0) & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    If Sql <> "" Then Err.Raise 513, , "Existen centros sin dar de alta en el sistema" & vbCrLf & Sql
        
    
    
    'Secciones
    Sql = "select distinct(seccion) from importnavtmp where not seccion in (select concepto from importnavconceptos)"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not miRsAux.EOF
        Sql = Sql & miRsAux.Fields(0) & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close

    If Sql <> "" Then Err.Raise 513, , "Existen secciones sin dar de alta en el sistema" & vbCrLf & Sql
    
    
    'Seccion - centro de coste - codmacta
    Sql = "select distinct tienda,seccion from importnavtmp where not (tienda,seccion) in (select codcentro,codconcepto from importnavconcepcentro)"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not miRsAux.EOF
        Sql = Sql & miRsAux.Fields(0) & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Sql <> "" Then Err.Raise 513, , "Existen secciones-tiendas-cuentas sin dar de alta en el sistema" & vbCrLf & Sql
    
    
    
    
    'Tipos de iva
    Sql = "select distinct(porceniva) from importnavtmp order by 1 desc"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TextoBusqueda = ""
    jj = 0
    While Not miRsAux.EOF
        TextoBusqueda = TextoBusqueda & miRsAux.Fields(0) & "|"
        jj = jj + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If TextoBusqueda = "" Then Err.Raise 513, , "No hay IVAS(1)" & vbCrLf & Sql
    If jj > 3 Then Err.Raise 513, , "Mas de tres IVAS(2)" & vbCrLf & Sql

    'Veo los IVAS de parametros
    Sql = "select IVANormal ,IVAReducid ,IVASuperReducido from importnavparam "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    Fin = False
    For jj = 1 To 3
        If IsNull(miRsAux.Fields(jj - 1)) Then
            Fin = True
        Else
            Sql = Sql & ", " & miRsAux.Fields(jj - 1)
        End If
    Next
    miRsAux.Close
    
    'Ya tengo los IVAS
    If Fin Then Err.Raise 513, , "Error en parametros IVAS(3)" & vbCrLf & Sql
    Sql = Mid(Sql, 2)
    
    'Cargo los tipos de iVA en el rs
    Sql = "Select * from tiposiva where codigiva in (" & Sql & ")"
    miRsAux.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    
    While TextoBusqueda <> ""
        jj = InStr(1, TextoBusqueda, "|")
        If jj = 0 Then
            TextoBusqueda = ""
        Else
            Sql = Mid(TextoBusqueda, 1, jj - 1)
            TextoBusqueda = Mid(TextoBusqueda, jj + 1)
        
            Porcen = Val(Sql) / 100
            
            miRsAux.MoveFirst
            Fin = False
            While Not Fin
                If miRsAux!porceiva = Porcen Then
                
                    Sql = " WHERE porceniva = '" & Sql & "'"
                    Sql = "UPDATE importnavtmp SET tipoiva  =" & miRsAux!codigiva & Sql
                    Conn.Execute Sql
                
                    Sql = ""
                    Fin = True
                Else
                    miRsAux.MoveNext
                    If miRsAux.EOF Then Fin = True
                End If
            Wend
            If Sql <> "" Then Err.Raise 513, , "No se encuentra iva: " & CStr(Porcen) & vbCrLf & Sql
            
        End If
    Wend
    miRsAux.Close
    
    
    ComprobarDatosFichero = True
    
eComprobarDatosFichero:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function

Private Sub ImprimeLineaDebug(LineaDebug As String, Valor, Posiciones As Integer)
    LineaDebug = LineaDebug & Right(String(Posiciones, " ") & Valor, Posiciones) & "|"
End Sub

Private Sub CalculoDeBases()
Dim Base As Currency
Dim PorcenFranquicia As Currency
Dim Unidades As Currency
Dim IVA As Currency
Dim Intermedio As Currency
Dim PreCoste As Currency
Dim Margen As Currency
Dim Aux As Currency
Dim BaseImponible As Currency
Dim CosteSinPorcen As Currency

Dim ImprimeLinea As Boolean   'Para el debug
Dim LinDebug As String

    Sql = "select * from importnavtmp "
    
    PorcenFranquicia = 0.05
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
       ImprimeLinea = False
       LinDebug = ""
        
       If miRsAux!seccion = 16 Then
             
            ImprimeLinea = True
        End If
        
       'If miRsAux!seccion = 21 Then
       '     If miRsAux!porceniva = "1000" Then
       '
       '         ImprimeLinea = True
       '     End If
       ' End If
        
        '0718647605270102220000010000100000000005000000500000000004000000040000210000000000000004800000048000                    0000000000000000100000000000C+
        
        'Ejemplo debug
        '   Unid.   Coste   Venta   Imp. Coste  IVA Vta s/iva   Imp. V.s/iva    Margen  Cuota 5%    Base Imp.   IVA Total
        '   1000    0,04    0,05    40,00   21,00   0,041   41,00   1,00    0,05    40,05   8,41    48,46

               
        'UNIDADES
        Base = miRsAux!cajaformser + (miRsAux!cajaformserDec / 1000)
       
        
        Intermedio = Val(miRsAux!udformato)
        Unidades = CCur(Base * Intermedio)
        If miRsAux!signo = "-" Then Unidades = -Unidades
            
        ImprimeLineaDebug LinDebug, Unidades, 8
            
        'IVA
        '-------------------
        IVA = CCur(miRsAux!porceniva) / 100    'ahora tengo el %,
        IVA = IVA / 100     'calculo el tanto por uno
                
        'Coste con IVA
        PreCoste = miRsAux!precosteud / 1000
        ImprimeLineaDebug LinDebug, PreCoste, 8
        
        'Coste * uDs
        PreCoste = Round2(PreCoste * Unidades, 2)
        ImprimeLineaDebug LinDebug, PreCoste, 8
        
        'DEBGU: Venta   Imp. Coste  IVA Vta s/iva   Imp. V.s/iva    Margen  Cuota 5%    Base Imp.   IVA Total
        'PVP sin IVA.   Octubre 2016 UNITARIO
        Base = miRsAux!precioventa / 100
        ImprimeLineaDebug LinDebug, Base, 8
        Base = Round(((Base / Unidades) / (1 + IVA)) * Unidades, 3)
        
        'Debug. Im`porte coste:
        ImprimeLineaDebug LinDebug, Base, 8 'PVP sin iva
        ImprimeLineaDebug LinDebug, IVA * 100, 10
        
        'BaseImponible = Base * Unidades
        BaseImponible = Round2(Base * Unidades, 2)
        ImprimeLineaDebug LinDebug, BaseImponible, 8
    
        'Intermedio = Base - PreCoste
        'Intermedio = BaseImponible - PreCoste
        Intermedio = Round2(BaseImponible - PreCoste, 2)
        
        ImprimeLineaDebug LinDebug, Intermedio, 8
        
        
        Margen = (Intermedio * PorcenFranquicia)
        ImprimeLineaDebug LinDebug, Margen, 8
        
        Margen = Intermedio - Margen
        ImprimeLineaDebug LinDebug, Margen, 8  'margen burot
       
        'Base = Round(PreCoste + (Intermedio * PorcenFranquicia), 2)
        
        ''ImprimeLineaDebug LinDebug, Base, 8
        
       ' Intermedio = Base * Iva
       ' Iva = Intermedio
        ''ImprimeLineaDebug LinDebug, Iva, 8
        
       ' Intermedio = Base + Iva
        ''ImprimeLineaDebug LinDebug, Intermedio, 8

        If ImprimeLinea Then Debug.Print LinDebug
        Sql = "UPDATE importnavtmp SET coste =" & TransformaComasPuntos(CStr(PreCoste))
        Sql = Sql & ", margenNeto=" & TransformaComasPuntos(CStr(Intermedio))
        Sql = Sql & ", margenBruto=" & TransformaComasPuntos(CStr(Margen))  'estaba CosteSinPorcen
        Sql = Sql & " WHERE secuencial = " & miRsAux!secuencial
        Conn.Execute Sql
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    

    
    espera 0.5
    'Insertamos en tmptotales
    Sql = "INSERT INTO importanatmptotal(seccion,porceniva,base,iva,Modificad,tipoiva) "
    
    
    'ANTES
    'SQL = SQL & "select seccion,(porceniva + 0)/100,truncate(sum(basecalculada+coste),2),round(sum(ivacalculado),2),0,tipoiva"
    Sql = Sql & "select seccion,(porceniva + 0)/100,truncate(sum(basecalculada+coste),2),round(sum(ivacalculado),2),0,tipoiva"
    Sql = Sql & " from importnavtmp group by 1,2"
    'Conn.Execute SQL
    
    
    Sql = "select seccion,(porceniva + 0)/100 porceiva,sum(coste) coste,sum(margenneto) margenN,round(sum(margenbruto),2) margenB,0,tipoiva"
    Sql = Sql & " from importnavtmp group by 1,2"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not miRsAux.EOF
        'SQL = "INSERT INTO importanatmptotal(seccion,porceniva,base,iva,Modificad,tipoiva) "
        IVA = CCur(miRsAux!porceiva)
        PreCoste = miRsAux!Coste
        Margen = miRsAux!margenN
        Intermedio = miRsAux!margenB
        PreCoste = PreCoste + Margen - Intermedio
        Intermedio = (IVA * PreCoste) / 100
        
        Sql = Sql & ", (" & miRsAux!seccion & "," & TransformaComasPuntos(CStr(miRsAux!porceiva)) & ","
        Sql = Sql & TransformaComasPuntos(CStr(PreCoste)) & "," & TransformaComasPuntos(CStr(Intermedio))
        Sql = Sql & ",0," & miRsAux!TipoIva & ")"
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Sql = Mid(Sql, 2)
    Sql = "INSERT INTO importanatmptotal(seccion,porceniva,base,iva,Modificad,tipoiva) VALUES " & Sql
    Conn.Execute Sql
    
    Set miRsAux = Nothing
    
End Sub

Private Sub ElLbl(ByRef TEXTO As String)
    lblIndicador.Caption = TEXTO
    lblIndicador.Refresh
End Sub
Private Sub ImportarFactura()
Dim B As Boolean
Dim Fichero As String
Dim NumeroReg As Long
Dim rCta As ADODB.Recordset
Dim Fecha As Date

    'Seleccionamos el fichero
    
    
    CadenaDesdeOtroForm = ""
    ElLbl "Leyendo fichero"
    frmImportarNavProceso.Show vbModal
    ElLbl ""
    If CadenaDesdeOtroForm = "" Then Exit Sub

    
    'Mayo 2020
    'Tratamos 3 tipos de factura FC (el de antes)    FA , administracion   y FV -autoventas
    Fichero = Dir(CadenaDesdeOtroForm)
    If Fichero = "" Then
        Fichero = "Error obteniedo nombre fichero"
    Else
        Fichero = Mid(Fichero, 1, 2)
        If InStr(1, "FC-FA-FV", Fichero) = 0 Then
            Fichero = "Tipo de fichero no tratado (FC,FA,FV)"
        Else
            TipoFichero = 0
            If Fichero = "FA" Then TipoFichero = 1
            If Fichero = "FV" Then TipoFichero = 2
            Fichero = ""
        End If
    End If
        
    If Fichero <> "" Then
        MsgBoxA Fichero, vbExclamation
        Exit Sub
    End If
    
    'Vere si el fichero YA ha sido procesado
    Fichero = Dir(CadenaDesdeOtroForm)
    Fichero = DevuelveDesdeBD("fecha", "importnavhcofich", "fichero", Fichero, "T")
    If Fichero <> "" Then
        MsgBox "El fichero ya ha sido procesado con anterioridad. " & Fichero, vbExclamation
        Exit Sub
    End If
    
    
    If Not BloqueoManual(True, "Consum", "1") Then
        MsgBox "Proceso bloqueado por otro usuario", vbExclamation
        Exit Sub
    End If
    
    
    ElLbl "Abrir fichero"
    'Fichero = "C:\Users\David\Desktop\FC3822093514311052453000000.rtf"
    Fichero = CStr(CadenaDesdeOtroForm)
    B = ImportarFactura_(Fichero, Me.lblIndicador)
    
    If B Then
        CadenaDesdeOtroForm = ""
        ElLbl "Ajustar importes"
        frmImportacionNavFra.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            Set rCta = New ADODB.Recordset
            
            'OK: Generamos la factura
            
            Conn.BeginTrans
            ElLbl "Genera factura"
            NumeroReg = GenerarLaFactura1(rCta)
            If NumeroReg >= 0 Then
                
                Conn.CommitTrans
                
            Else
                Conn.RollbackTrans
            End If
    
    
            If NumeroReg >= 0 Then
                ElLbl "Finalizar"
                GenerarApunteyPago NumeroReg
                
                'Insertamos en la tablas de hoc de procesados
                InsertamosEnHcos
                
                'OK. METEMOS EL nuevo fichero
                CargaHistorico
                
                
                'Finalmente metemos el pago
                ProcesoCrearPagos rCta
                rCta.Close
                
                MsgBox "Proceso finalizado correctamente", vbInformation
            End If
            Set rCta = Nothing
        End If
    End If
    
    ElLbl ""
    'Cerramos el bloqueo del proceso
    BloqueoManual False, "Consum", 1

End Sub

Private Function GenerarLaFactura1(ByRef rCta As ADODB.Recordset) As Long
Dim Mc As Contadores
Dim Aux As String
Dim Bases As Currency
Dim Ivas As Currency

    
    GenerarLaFactura1 = -1

    Set Mc = New Contadores
    
    On Error GoTo eGenerarLaFactura
    FechaFraImportada = CDate(RecuperaValor(CadenaDesdeOtroForm, 1))
    If Mc.ConseguirContador("1", FechaFraImportada <= vParam.fechafin, True) = 1 Then
        Set Mc = Nothing
        Exit Function
    End If




    'Insertamos la cabecera
    'cabfactprov(numregis,fecfacpr,anofacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,fecliqpr
    Aux = "numfac"
    Sql = DevuelveDesdeBD("fechafac", "importnavtmp", "1", "1", "N", Aux)
    NumFacturaFichero = Aux
    Sql = "1," & Mc.Contador & ",'" & Sql & "'," & Year(FechaFraImportada) & ",'" & Format(FechaFraImportada, FormatoFecha) & "','" & Aux & "'"
    

    Aux = DevuelveDesdeBD("Ctaconsum", "importnavparam ", "1", "1", "N")
    
    Aux = "SELECT codmacta,nommacta,dirdatos,codposta,despobla,desprovi,nifdatos,codpais,forpa,ctabanco,iban FROM cuentas where codmacta= '" & Aux & "'"
    rCta.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Cosas importantes
    
    If rCta.EOF Then Err.Raise 513, , "Error leyendo cuenta contable"
    If IsNull(rCta!Forpa) Then Err.Raise 513, , "Forma pago NULA"
    If IsNull(rCta!CtaBanco) Then Err.Raise 513, , "Cuenta banco asociada cuenta"

    Sql = Sql & ",'" & rCta!codmacta & "','" & DevNombreSQL(RecuperaValor(CadenaDesdeOtroForm, 2)) & "',"
    
    
    
    'ba1facpr,  pi1facpr,  pr1facpr, ti1facpr,  tr1facpr, tp1facpr
    Aux = "select  porceniva '%iva',sum(base) base,sum(iva) IVA,sum(base+iva) Total,tipoiva "
    Aux = Aux & " from importanatmptotal  group by 1 order by 1 desc"
    Set miRsAux = New ADODB.Recordset
    jj = 0
    ImporteFacturaImportada = 0
    Bases = 0
    Ivas = 0
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If DBLet(miRsAux.Fields(1), "N") <> 0 Then
            jj = jj + 1
            If jj > 3 Then Err.Raise 513, , "Mas de tres tipos de iva"
            Bases = Bases + DBLet(miRsAux!Base, "N")
            Ivas = Ivas + DBLet(miRsAux!IVA, "N")
            ImporteFacturaImportada = ImporteFacturaImportada + miRsAux.Fields(3)
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
   
    
    'codconce340, CodOpera,   codforpa,  totbases, totbasesret, totivas"
    Sql = Sql & "0,0," & rCta!Forpa & "," & TransformaComasPuntos(CStr(Bases)) & ",0," & TransformaComasPuntos(CStr(Ivas)) & ","
    
    ',nommacta,dirdatos,codpobla,despobla,desprovi,,codpais"
    Sql = Sql & DBSet(rCta!Nommacta, "T") & "," & DBSet(rCta!dirdatos, "T") & "," & DBSet(rCta!codposta, "T") & ","
    Sql = Sql & DBSet(rCta!desPobla, "T") & "," & DBSet(rCta!desProvi, "T") & "," & DBSet(rCta!nifdatos, "T") & ","
    Sql = Sql & DBSet(rCta!codpais, "T") & ","
    
    ' totfacpr, extranje, fecliqpr
    Sql = Sql & TransformaComasPuntos(CStr(ImporteFacturaImportada)) & ",'" & Format(FechaFraImportada, FormatoFecha) & "'"
    
    
   
    Aux = "INSERT INTO factpro(NUmSerie , numregis,FecFactu, anofactu,fecharec, NumFactu, codmacta,observa,"
    
    Aux = Aux & "codconce340, CodOpera, codforpa,  totbases, totbasesret, totivas"
    Aux = Aux & ",nommacta,dirdatos,codpobla,despobla,desprovi,nifdatos,codpais"
    
    
    
    Aux = Aux & " ,totfacpr,  fecliqpr) VALUES (" & Sql & ")"
    
    Conn.Execute Aux
    
    

    'Los IVAS
    Aux = "select  porceniva '%iva',sum(base) base,sum(iva) IVA,sum(base+iva) Total,tipoiva "
    Aux = Aux & " from importanatmptotal  group by 1 order by 1 desc"
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = ""
    jj = 0
    While Not miRsAux.EOF
        If DBLet(miRsAux.Fields(1), "N") <> 0 Then
            jj = jj + 1
            'factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
            Sql = ", (1," & Mc.Contador & ",'" & Format(FechaFraImportada, FormatoFecha) & "'," & Year(FechaFraImportada) & "," & jj & ","
            Sql = Sql & DBSet(miRsAux!Base, "N") & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux.Fields(0), "N") & ","
            Sql = Sql & "NULL," & DBSet(miRsAux!IVA, "N") & ",NULL)"
            Aux = Aux & Sql
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Sql = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec) VALUES "
    Aux = Mid(Aux, 2)
    Aux = Sql & Aux
    Conn.Execute Aux
    
    
    
    'Las bases
    Aux = "select  base,porceniva,iva,tipoiva,codmacta from importanatmptotal left join importnavconcepcentro  "
    Aux = Aux & " on seccion=codconcepto And codcentro = " & RecuperaValor(CadenaDesdeOtroForm, 3)
    Aux = Aux & " ORDER BY seccion,porceniva"
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    jj = 0
    While Not miRsAux.EOF
        'factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
        jj = jj + 1
        Sql = Sql & ", (1," & Mc.Contador & ",'" & Format(FechaFraImportada, FormatoFecha) & "'," & Year(FechaFraImportada) & "," & jj & ","
        Sql = Sql & DBSet(miRsAux!codmacta, "T") & "," & DBSet(miRsAux!Base, "N") & "," & miRsAux!TipoIva & ","
        Sql = Sql & DBSet(miRsAux!porceniva, "N") & ",NULL," & DBSet(miRsAux!IVA, "N") & ",NULL)"
        
        miRsAux.MoveNext
        jj = jj + 1
    Wend
    miRsAux.Close
    
    Sql = Mid(Sql, 2)
    Aux = "INSERT INTO factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec) VALUES "
    Sql = Aux & Sql
    Conn.Execute Sql
    
    
    'Log de inserciones
    vLog.Insertar 7, vUsu, "Importar consum: " & Mc.Contador & " - "
    
    GenerarLaFactura1 = Mc.Contador
    
eGenerarLaFactura:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set Mc = Nothing
End Function


Private Sub GenerarApunteyPago(NumeroDeRegistro As Long)

    espera 0.5
    If vParam.ContabilizaFactura Then
                   

        With frmActualizar
            .OpcionActualizar = 8
            'NumAsiento     --> CODIGO FACTURA
            'NumDiari       --> AÑO FACTURA
            'NUmSerie       --> SERIE DE LA FACTURA
            'FechaAsiento   --> Fecha factura
            .NumFac = NumeroDeRegistro
            .NumDiari = Year(FechaFraImportada)
            .NUmSerie = 1
            .FechaAsiento = FechaFraImportada
            .DiarioFacturas = vParam.numdiapr
            .NumAsiento = 0
            .Show vbModal
    
        End With
    End If
    
    If vEmpresa.TieneTesoreria Then
        '|0500052453||07/11/2014|40000010|CONSUM puesto por David|
        
        CadenaDesdeOtroForm = DevuelveDesdeBD("Ctaconsum", "importnavparam ", "1", "1")
        Sql = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CadenaDesdeOtroForm, "T")
        CadenaDesdeOtroForm = "|" & CadenaDesdeOtroForm & "|" & Sql & "|"
        
        Sql = "concat(numfac,'||',DATE_FORMAT(fechafac,'%d/%m/%Y'))"
        Sql = DevuelveDesdeBD(Sql, "importnavtmp", "1", "1")
        Sql = "|" & Sql & CadenaDesdeOtroForm
        
     '   frmVto.Opcion = 1
     '   frmVto.Importe = ImporteFacturaImportada
     '   frmVto.Datos = SQL
     '   frmVto.Show vbModal

    End If
End Sub



Private Sub InsertamosEnHcos()
    Sql = DevuelveDesdeBD("max(id)", "importnavhcofich", "1", "1")
    If Sql = "" Then Sql = "0"
    jj = Val(Sql) + 1
    
    
    Sql = DevuelveDesdeBD("Fichero", "importanatmptotal", "1", "1", "N")
    TextoBusqueda = DevuelveDesdeBD("tienda", "importnavtmp", "1", "1", "N")
    'importnavhcofich id fecha fichero centro
    Sql = jj & ",now(),'" & Sql & "'," & TextoBusqueda & ")"
    Sql = "INSERT INTO importnavhcofich(id ,fecha ,fichero ,centro) VALUES (" & Sql
    Conn.Execute Sql
    
    Sql = "seccion,porceniva,base,iva,Modificad,Fichero,tipoiva"
    Sql = "INSERT INTO importnavtotalhco(id," & Sql & ") SELECT " & jj & "," & Sql & " FROM importanatmptotal"
    Conn.Execute Sql
     

End Sub






Private Function Truncar(numero As Single, Decimales) As Single
Dim CADENA As String

    CADENA = Format(numero, "#0.00000")
    CADENA = Mid(CADENA, 1, Len(CADENA) - (5 - Decimales))
    Truncar = CSng(CADENA)
End Function






'--------------------------------------------------------------
'Para crear los pagos

Private Sub ProcesoCrearPagos(ByRef rCta As ADODB.Recordset)
Dim TipForpa  As Integer
Dim CadValues As String


On Error GoTo eInsertaPagos

    ElLbl "Calcula pagos"
    
    CadValues = "DELETE from tmppagos where codusu =" & vUsu.Codigo
    Conn.Execute CadValues
    espera 0.1
    
    'NomBanco = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", rCta!CtaBanco, "T")
    TipForpa = DevuelveValor("select formapago.tipforpa from formapago where codforpa = " & DBSet(rCta!Forpa, "N"))
    
    
    CargarPagosTemporal CStr(rCta!Forpa), CStr(FechaFraImportada), ImporteFacturaImportada
    

        
'    CadInsert = "insert into pagos (numserie,codmacta,numfactu,fecfactu,numorden,codforpa,fecefect,impefect," & _
                "ctabanc1,fecultpa,imppagad,emitdocum," & _
                "text1csb,text2csb,nrodocum,referencia, iban,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais,situacion,codusu) values "
    


    Dim Borra
    I = 0
    
    Set miRsAux = New ADODB.Recordset
    CadValues = "Select * from tmppagos where codusu =" & vUsu.Codigo
    'numserie='1' and numregis = " &  & " AND anofactu =" & Year(FechaFraImportada)
    
    miRsAux.Open CadValues, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    While Not miRsAux.EOF
        I = I + 1
        'numserie,codmacta,numfactu,fecfactu,numorden
        Sql = "1," & DBSet(rCta!codmacta, "T") & "," & DBSet(NumFacturaFichero, "T") & "," & DBSet(FechaFraImportada, "F") & "," & DBSet(I, "N") & ","
        'codforpa,fecefect,impefec
        Sql = Sql & DBSet(rCta!Forpa, "N") & "," & DBSet(miRsAux!FecVenci, "F") & "," & DBSet(miRsAux!ImpVenci, "N") & ","
        'ctabanc1,fecultpa,imppagad,emitdocum,"
        Sql = Sql & DBSet(rCta!CtaBanco, "T", "S") & ","
        Sql = Sql & ValorNulo & "," & ValorNulo & ",0,"
        'ext1csb,text2csb,nrodocum,referencia,iban
        Sql = Sql & DBSet("Factura " & NumFacturaFichero & " de Fecha " & FechaFraImportada, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(rCta!IBAN, "T", "S") & ","
        
        'nomprove , domprove, pobprove, cpprove, proprove, nifprove, codpais
        Sql = Sql & DBSet(rCta!Nommacta, "T") & "," & DBSet(rCta!dirdatos, "T") & "," & DBSet(rCta!desPobla, "T") & "," & DBSet(rCta!codposta, "T") & ","
        Sql = Sql & DBSet(rCta!desProvi, "T") & "," & DBSet(rCta!nifdatos, "T") & "," & DBSet(rCta!codpais, "T") & ","
        
        'situacion,codusu
        Sql = Sql & "0," & DBSet(vUsu.Id, "N")
        
        
        CadValues = CadValues & ", (" & Sql & ")"
    
        miRsAux.MoveNext
    Wend

    If CadValues <> "" Then
        
        Sql = "insert into pagos (numserie,codmacta,numfactu,fecfactu,numorden,codforpa,fecefect,impefect,"
        Sql = Sql & "ctabanc1,fecultpa,imppagad,emitdocum,"
        Sql = Sql & "text1csb,text2csb,nrodocum,referencia, iban,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais,situacion,codusu) values "
        CadValues = Mid(CadValues, 2, Len(CadValues) - 1)
        Conn.Execute Sql & CadValues
    End If

    
    

eInsertaPagos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Sub






Private Function CargarPagosTemporal(Forpa As String, FecFactu As String, TotalFac As Currency) As Boolean
Dim Sql As String
Dim CadValues As String
Dim Rsvenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarPagos

    CargarPagosTemporal = False

    Sql = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set Rsvenci = New ADODB.Recordset
    Rsvenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    If Not Rsvenci.EOF Then
        If Rsvenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            I = 1
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(Rsvenci!primerve, "N"), FecVenci)
            '===
            
            'IMPORTE del Vencimiento
            If Rsvenci!numerove = 1 Then
                ImpVenci = TotalFac
            Else
                ImpVenci = Round(TotalFac / Rsvenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * Rsvenci!numerove <> TotalFac Then
                    ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * Rsvenci.Fields(0).Value), 2)
                End If
            End If
            CadValues = "(" & vUsu.Codigo & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For I = 2 To Rsvenci!numerove
                FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round(TotalFac / Rsvenci.Fields(0).Value, 2)
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next I
        End If
    End If
    
    Set Rsvenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO tmppagos (codusu, numorden, fecvenci, impvenci)"
        Sql = Sql & " VALUES " & Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute Sql
    End If
    
    CargarPagosTemporal = True
    Exit Function

eCargarPagos:

End Function


'***********************************************************************************************************************
'****************    LO que habia para los rtfs
'********
'            'La primera linea NO vale ej: {\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Courier New;}}
'            InicioLineas = ""
'            While Not EOF(NF)
'                Line Input #NF, Linea
'
'
'
'                    'Buscamos el {
'                    jj = InStrRev(Linea, "}")
'                    If jj > 0 Then Linea = Mid(Linea, jj + 1)
'
'                    'La \
'                    Do
'                        jj = InStr(Linea, "\")
'                        If jj > 0 Then
'                            Sql = Mid(Linea, 1, jj - 1)
'
'                            K = InStr(jj, Linea, " ")
'
'                            If K > 0 Then
'                                Linea = Mid(Linea, K + 1) 'quito el espacio
'                            Else
'                                Linea = ""
'                            End If
'                            Linea = Sql & Linea
'                        End If
'                    Loop Until jj = 0
'
'                    TextoAProcesar = TextoAProcesar & Linea
'
'
'            Wend
'
'
'            'El incio de linea es el mismoa para todos
'            '382205000524530
'            If Len(TextoAProcesar) < 15 Then
'                MsgBox "Fichero incorrecto(I)", vbExclamation
'            Else
'                InicioLineas = Mid(TextoAProcesar, 1, 15)
'                jj = 0
'                While TextoAProcesar <> ""
'                    jj = jj + 1
'                    If Len(TextoAProcesar) <= 220 Then
'                        Linea = Mid(TextoAProcesar & Space(220), 1, 220)
'                        ColLineas.Add Linea
'                        TextoAProcesar = ""
'
'                    Else
'                         K = InStr(3, TextoAProcesar, InicioLineas)
'                         If K = 0 Then
'                            Sql = TextoAProcesar
'                            TextoAProcesar = ""
'                         Else
'
'                            Sql = Mid(TextoAProcesar, 1, K - 1)
'                            If Len(Sql) < 220 Then Sql = Left(Sql & Space(220), 220)
'                            TextoAProcesar = Mid(TextoAProcesar, K)
'                         End If
'
'
'                         If Len(Sql) > 220 Then
'                            MsgBox "Fichero incorrecto(II). Avise Ariadna", vbExclamation
'
'                         Else
'
'                            ColLineas.Add Sql
'
'                         End If
'                         'Debug.Print SQL
'
'                    End If
'
'                Wend
'
'            End If 'Len<16
'

