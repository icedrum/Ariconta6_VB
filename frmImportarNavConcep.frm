VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImportarNavConcep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   5325
   Icon            =   "frmImportarNavConcep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4140
      TabIndex        =   3
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   6000
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   310
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Numero de diario|N|N|0|100|importnavconceptos|concepto|00|S|"
      Text            =   "Dat"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   310
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Tag             =   "Denominación|T|N|||importnavconceptos|descripcion|||"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4140
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   5880
      Width           =   1755
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
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
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3840
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   495
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmImportarNavConcep.frx":000C
      Height          =   5055
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
Attribute VB_Name = "frmImportarNavConcep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos
Dim Modo As Byte

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

txtaux(0).visible = Not B
txtaux(1).visible = Not B
mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(6).Enabled = B And vUsu.Nivel < 2
Me.mnNuevo.Enabled = Toolbar1.Buttons(6).Enabled
Toolbar1.Buttons(7).Enabled = B And vUsu.Nivel < 2
Me.mnModificar.Enabled = Toolbar1.Buttons(7).Enabled
Toolbar1.Buttons(8).Enabled = B And vUsu.Nivel < 2
Me.mnEliminar.Enabled = Toolbar1.Buttons(8).Enabled



cmdAceptar.visible = Not B
cmdCancelar.visible = Not B
DataGrid1.Enabled = B

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.visible = B
End If


'Si estamo mod or insert
If Modo = 2 Then
   txtaux(0).BackColor = &H80000018
   Else
    txtaux(0).BackColor = &H80000005
End If
txtaux(0).Enabled = (Modo <> 2)
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.top + 240
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top + 60
        
    End If
    txtaux(0).Text = NumF
    txtaux(1).Text = ""
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtaux(0).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub

Private Sub BotonBuscar()
    CargaGrid "concepto= -1"
    'Buscar
    txtaux(0).Text = ""
    txtaux(1).Text = ""
    LLamaLineas 770, 2
    txtaux(0).SetFocus
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
End Sub



Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim cad As String
    Dim anc As Single
    Dim I As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.top + 220
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top + 15
    End If
    cad = ""
    For I = 0 To 1
        cad = cad & DataGrid1.Columns(I).Text & "|"
    Next I
    'Llamamos al form
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtaux(1).Text = DataGrid1.Columns(1).Text
    LLamaLineas anc, 1
   
   'Como es modificar
   txtaux(1).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub


'Private Sub DeseleccionaGrid()
'    On Error GoTo EDeseleccionaGrid
'
'    While Datagrid1.SelBookmarks.Count > 0
'        Datagrid1.SelBookmarks.Remove 0
'    Wend
'    Exit Sub
'EDeseleccionaGrid:
'        Err.Clear
'End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid DataGrid1
    PonerModo xModo + 1
    'Fijamos el ancho
    txtaux(0).top = alto
    txtaux(1).top = alto
    txtaux(0).Left = DataGrid1.Left + 340
    txtaux(1).Left = txtaux(0).Left + txtaux(0).Width + 45
End Sub

Private Sub BotonEliminar()
Dim Sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    

    '### a mano
    Sql = "Seguro que desea eliminar el diario:"
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from importnavconceptos where concepto=" & adodc1.Recordset!Concepto
        Conn.Execute Sql
        CancelaADODC
        CargaGrid ""
        CancelaADODC
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar diario", Err.Description
End Sub


Private Sub CancelaADODC()
On Error Resume Next
adodc1.Recordset.Cancel
If Err.Number <> 0 Then Err.Clear
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
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
                CadenaDesdeOtroForm = "OK" 'para que recarge los datos en el formulario ppal
            End If
        End If
    Case 2
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CancelaADODC
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                    CadenaDesdeOtroForm = "OK" 'para que recarge los datos en el formulario ppal
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    End Select


End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

    cad = adodc1.Recordset.Fields(0) & "|"
    cad = cad & adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
If cmdRegresar.visible = True Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If Modo = 0 Then Unload Me
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    ' ICONITOS DE LA BARRA
    Me.Icon = frmppal.Icon
    With Me.Toolbar1
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 16
        .Buttons(11).Image = 15
    End With
    
    
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    CadAncho = False
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    PonerModo 0
    'Cadena consulta
    CadenaConsulta = "Select concepto,descripcion from importnavconceptos"
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
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Function SugerirCodigoSiguiente() As String
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    
    Sql = "Select Max(concepto) from importnavconceptos where concepto<99"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, , , adCmdText
    Sql = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            Sql = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = Sql
End Function

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
Case 10
        Screen.MousePointer = vbHourglass
        
Case 11
        Unload Me
Case Else

End Select
End Sub


Private Sub CargaGrid(Optional Sql As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    Dim B As Boolean
    
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    adodc1.ConnectionString = Conn
    If Sql <> "" Then
        Sql = CadenaConsulta & " WHERE " & Sql
        Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " ORDER BY concepto"
    adodc1.RecordSource = Sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = txtaux(0).Height
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Codigo"
        DataGrid1.Columns(I).Width = 700
        DataGrid1.Columns(I).NumberFormat = "000"
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Descripción"
        DataGrid1.Columns(I).Width = 3400
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtaux(0).Width = DataGrid1.Columns(0).Width - 60
        txtaux(0).Left = DataGrid1.Left + 340
        txtaux(1).Width = DataGrid1.Columns(1).Width - 60
        txtaux(1).Left = DataGrid1.Columns(1).Left + 90
        CadAncho = True
    End If
   
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
   DataGrid1.Enabled = B
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
'With txtaux(Index)
'    .SelStart = 0
'    .SelLength = Len(.Text)
'End With
    PonFoco txtaux(Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        SendKeys "{tab}"
'    End If
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

txtaux(Index).Text = Trim(txtaux(Index).Text)
If txtaux(Index).Text = "" Then Exit Sub
If Modo = 3 Then Exit Sub 'Busquedas
If Index = 0 Then
    If Not IsNumeric(txtaux(0).Text) Then
        MsgBox "Código diario tiene que ser numérico", vbExclamation
        Exit Sub
    End If
    txtaux(0).Text = Format(txtaux(0).Text, "00")
End If
End Sub

Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function




If Modo = 1 Then
    If Val(txtaux(0).Text) > 98 Then
        MsgBox "Valor maximo 99", vbExclamation
        B = False
    End If


    'Estamos insertando
     Datos = DevuelveDesdeBD("concepto", "importnavconceptos", "concepto", txtaux(0).Text, "T")
     If Datos <> "" Then
        MsgBox "Ya existe el diario : " & txtaux(0).Text, vbExclamation
        B = False
    End If
End If
DatosOK = B
End Function


Private Function SepuedeBorrar() As Boolean
Dim Sql As String
    SepuedeBorrar = False
    Sql = DevuelveDesdeBD("codcentro", "importnavconcepcentro", "codconcepto", adodc1.Recordset!Concepto, "N")
    If Sql <> "" Then
        MsgBox "Esta vinculada con centros", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True
End Function



















