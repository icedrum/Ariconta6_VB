VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmColCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmColCtas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   5685
      Left            =   6450
      TabIndex        =   5
      Top             =   450
      Width           =   1425
      Begin VB.CheckBox Check1 
         Caption         =   "9º nivel"
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   4245
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "8º nivel"
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   3805
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "7º nivel"
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   3365
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "6º nivel"
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   2925
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "5º nivel"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2485
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4º nivel"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   2045
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "3º nivel"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "2º nivel"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1165
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1er nivel"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   725
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Último nivel"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Value           =   1  'Checked
         Width           =   1125
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   7
         Top             =   5205
         Width           =   1035
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cod. Cta"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   4890
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   75
         X2              =   1290
         Y1              =   4755
         Y2              =   4755
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColCtas.frx":000C
      Height          =   5565
      Left            =   90
      TabIndex        =   4
      Top             =   570
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   9816
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5550
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1035
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":0021
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":0133
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":0245
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":0357
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":0469
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":057B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":0E55
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":172F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":2009
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":28E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":31BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":360F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":3721
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":3833
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":3945
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColCtas.frx":3FBF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private CadenaConsulta As String
Dim CadAncho As String 'Para cuando llamemos al al form de lineas

Private Sub BotonAnyadir()
 
End Sub

Private Sub BotonBuscar()
    CadenaConsulta = GeneraSQL("codmacta= 'David'")  'esto es para que no cargue ningun registro
    CargaGrid
    'Buscar
  
'    'CadenaDevueltaFormHijo
'    If FormularioHijoModificado Then
'        CadenaConsulta = GeneraSQL(CadenaDevueltaFormHijo)
'        CargaGrid
'    End If
End Sub

Private Sub BotonVerTodos()
    CargaGrid
End Sub



Private Sub BotonModificar()
  
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar el concepto:"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from conceptos where codconce=" & adodc1.Recordset!codconce
        Conn.Execute SQL
    End If
    CargaGrid
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
        End If
End Sub




Private Sub Check1_Click(Index As Integer)
    OpcionesCambiadas
End Sub

Private Sub OpcionesCambiadas()
    Screen.MousePointer = vbHourglass
    CadenaConsulta = GeneraSQL("")
    CargaGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRegresar_Click()

If adodc1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If
RaiseEvent DatoSeleccionado(adodc1.Recordset!codmacta & "|" & adodc1.Recordset!nommacta & "|")
Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

'
Private Sub Form_Load()
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
     
    'Poner niveles
    PonerOptionsVisibles
    
    'Bloqueo de tabla, cursor type
    adodc1.Password = vUsu.Passwd
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    CadAncho = ""
    'Cadena consulta
    CadenaConsulta = GeneraSQL("")
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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



'----------------------------------------------------------------


Private Sub Option2_Click(Index As Integer)
OpcionesCambiadas
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
        BotonBuscar
Case 2
        BotonVerTodos
Case 6
        BotonAnyadir
Case 7
        BotonModificar
Case 8
        BotonEliminar
Case 11
        'Imprimimos el listado
            frmImprimir.Opcion = 1
            frmImprimir.Show vbModal

Case 12
        Unload Me
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = bol
    Next i
End Sub

Private Sub CargaGrid()
    Dim j As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    Dim SQL As String
    
    adodc1.ConnectionString = Conn
    SQL = CadenaConsulta
    SQL = SQL & " ORDER BY"
    If Option2(0).Value Then
        SQL = SQL & " codmacta"
    Else
        SQL = SQL & " nommacta"
    End If
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    
    'Nombre producto
        DataGrid1.Columns(0).Caption = "Cod"
        DataGrid1.Columns(0).Width = 1200
    
    'Leemos del vector en 2
        DataGrid1.Columns(1).Caption = "Denominación"
        DataGrid1.Columns(1).Width = 4000
        TotalAncho = TotalAncho + DataGrid1.Columns(1).Width
    
    'El importe es campo calculado
        DataGrid1.Columns(2).Caption = "Direc."
        DataGrid1.Columns(2).Width = 500
        TotalAncho = TotalAncho + DataGrid1.Columns(2).Width
            
    
        'Fiajamos el cadancho
    If CadAncho = "" Then
        For i = 0 To DataGrid1.Columns.Count - 1
            If DataGrid1.Columns(i).Visible Then
                CadAncho = CadAncho & DataGrid1.Columns(i).Width & "|"
            End If
        Next i
    End If
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
End Sub



Private Sub PonerOptionsVisibles()
Dim i As Integer

For i = vEmpresa.numnivel To 9
    Check1(i).Visible = False
Next i

End Sub



Private Function GeneraSQL(Busqueda As String) As String
Dim i As Integer
Dim SQL As String
Dim nexo As String
Dim j As Integer
Dim wildcar As String

SQL = ""
nexo = ""
If Check1(0).Value Then
    SQL = "( apudirec = 'S')"
    nexo = " OR "
End If
For i = 1 To vEmpresa.numnivel - 1
    If Check1(i).Value = 1 Then
        wildcar = ""
        For j = 1 To i
            wildcar = wildcar & "_"
        Next j
        SQL = SQL & nexo & " ( codmacta like '" & wildcar & "')"
        nexo = " OR "
    End If
Next i
wildcar = "SELECT codmacta, nommacta, apudirec"
wildcar = wildcar & " FROM cuentas "


'Nexo
nexo = " WHERE "
If Busqueda <> "" Then
    wildcar = wildcar & " WHERE (" & Busqueda & ")"
    nexo = " AND "
End If
If SQL <> "" Then wildcar = wildcar & nexo & "(" & SQL & ")"

GeneraSQL = wildcar
End Function
