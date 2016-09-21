VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmColPresu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmColPresu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   5880
      TabIndex        =   3
      Tag             =   "Importe|N|N|||presupuestos|imppresu||N|"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4440
      TabIndex        =   2
      Tag             =   "Denominación|N|N|1|12|presupuestos|mespresu||S|"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2880
      TabIndex        =   1
      Tag             =   "Año|N|N|1900||presupuestos|anopresu||S|"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7740
      TabIndex        =   5
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6540
      TabIndex        =   4
      Top             =   6000
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Cuenta|T|N|||presupuestos|codmacta||S|"
      Text            =   "Dat"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   1
      Left            =   1260
      TabIndex        =   6
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7740
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   8
      Top             =   5880
      Width           =   2115
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1440
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Object.ToolTipText     =   "Generar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7080
         TabIndex        =   10
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
      Bindings        =   "frmColPresu.frx":000C
      Height          =   5235
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9234
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
Attribute VB_Name = "frmColPresu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 1101


Private CadenaConsulta As String
Private CadAncho As Boolean  'Para saber si hemos fijado el ancho de los campos
Dim SQL As String
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
Dim i As Integer
Modo = vModo

B = (Modo = 0)


For i = 0 To txtAux.Count - 1
    txtAux(i).Visible = Not B
Next i

mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(6).Enabled = B
Toolbar1.Buttons(7).Enabled = B
Toolbar1.Buttons(8).Enabled = B
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid1.Enabled = B

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If


'Si estamo mod or insert
If Modo = 2 Then
    'Modificando
    For i = 0 To 3
        txtAux(i).BackColor = &H80000018
        txtAux(i).Enabled = False
    Next i
   Else
        For i = 0 To 3
            txtAux(i).BackColor = &H80000005
            txtAux(i).Enabled = True
        Next i
        txtAux(1).Enabled = False
        txtAux(1).BackColor = &H80000018
End If

End Sub

Private Sub BotonAnyadir()
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
 
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    anc = FijarVariableAnc(DataGrid1)

    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(3).Text = ""
    txtAux(4).Text = ""
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    PonFoco txtAux(0)
    
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
    DataGrid1.Enabled = False
    CargaGrid "anopresu= -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(3).Text = ""
    txtAux(4).Text = ""
    
    LLamaLineas FijarVariableAnc(DataGrid1), 2
    PonFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
End Sub



Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim i As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    
    anc = FijarVariableAnc(DataGrid1)
    
    
    Cad = ""
    For i = 0 To 1
        Cad = Cad & DataGrid1.Columns(i).Text & "|"
    Next i
    'Llamamos al form
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = DataGrid1.Columns(i).Text
    Next i
    LLamaLineas anc, 1
   
   'Como es modificar
   PonFoco txtAux(4)
   
    Screen.MousePointer = vbDefault
End Sub


Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Integer
DeseleccionaGrid
PonerModo xModo + 1
'Fijamos el ancho
For i = 0 To txtAux.Count - 1
    txtAux(i).Top = alto
Next i
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar el presupuesto para:"
    SQL = SQL & vbCrLf & "Cuenta: " & adodc1.Recordset!nommacta
    SQL = SQL & vbCrLf & "Mes / Anyo: " & adodc1.Recordset!mespresu & " - " & adodc1.Recordset!anopresu
    SQL = SQL & vbCrLf & "Importe: " & adodc1.Recordset!imppresu
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "DELETE FROM presupuestos WHERE codmacta='" & adodc1.Recordset!codmacta
        SQL = SQL & "' AND anopresu=" & adodc1.Recordset!anopresu
        SQL = SQL & " AND mespresu=" & adodc1.Recordset!mespresu
        Conn.Execute SQL
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
Dim i As Integer
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
            End If
        End If
    Case 2
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    PonerModo 0
                    CancelaADODC
                    CargaGrid
                    SituarData1
                End If
            End If
    Case 3
        
        'Algunas consideracion
        'Comprobamos que en cueta no hay un punto
        'y si lo hay es k quiere la cuenta
        If InStr(1, txtAux(0).Text, ".") > 0 Then
            CadB = txtAux(0).Text
            If CuentaCorrectaUltimoNivel(CadB, SQL) Then _
                txtAux(0).Text = CadB
        End If
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
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    SQL = "anopresu = -1"
    CargaGrid SQL
End Select
PonerModo 0
lblIndicador.Caption = ""
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

    Cad = adodc1.Recordset.Fields(0) & "|"
    Cad = Cad & adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible = True Then cmdRegresar_Click
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 21
        .Buttons(12).Image = 16
        .Buttons(13).Image = 15
    End With
    
    
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    CadAncho = False
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    PonerModo 0
    'Cadena consulta
    CadenaConsulta = "select presupuestos.codmacta, nommacta, anopresu,mespresu,imppresu from presupuestos,cuentas where presupuestos.codmacta=cuentas.codmacta"
    SQL = "anopresu = -1"
    CargaGrid SQL
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
        'Generacion de presupuestaria
        CadenaDesdeOtroForm = ""
        frmGenerar.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            CargaGrid CadenaDesdeOtroForm
            CadenaDesdeOtroForm = ""
        Else
            Me.Refresh
        End If
Case 12
        'Impresion
        Screen.MousePointer = vbHourglass
        frmListado.opcion = 9 'Listado de cuentas
        frmListado.Show vbModal
Case 13
        'Salir
        Unload Me
End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
Dim B As Boolean
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    CargaGrid2 vSQL
    DataGrid1.Enabled = B
End Sub



Private Sub CargaGrid2(Optional SQL As String)
    Dim i As Integer
    
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codmacta,anopresu,mespresu"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "Cuenta"
        DataGrid1.Columns(i).Width = 1300
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Título"
        DataGrid1.Columns(i).Width = 3600
            
    i = 2
        DataGrid1.Columns(i).Caption = "Año"
        DataGrid1.Columns(i).Width = 800

            
    i = 3
        DataGrid1.Columns(i).Caption = "Mes"
        DataGrid1.Columns(i).Width = 600
        DataGrid1.Columns(i).Alignment = dbgRight
    
    i = 4
        DataGrid1.Columns(i).Caption = "Importe"
        DataGrid1.Columns(i).Width = 1600
        DataGrid1.Columns(i).NumberFormat = "#,###,##0.00"
        DataGrid1.Columns(i).Alignment = dbgRight
        
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        For i = 0 To txtAux.Count - 1
            txtAux(i).Width = DataGrid1.Columns(i).Width - 30
        Next i
        txtAux(0).Left = DataGrid1.Left + 340
        For i = 1 To txtAux.Count - 1
            txtAux(i).Left = DataGrid1.Columns(i).Left + 150
        Next i
        CadAncho = True
    End If
   
    'Habilitamos modificar y eliminar
   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
With txtAux(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim RC As String
txtAux(Index).Text = Trim(txtAux(Index).Text)
If txtAux(Index).Text = "" Then Exit Sub
If Modo = 3 Then Exit Sub 'Busquedas

Select Case Index
Case 0
        RC = txtAux(0).Text
        If CuentaCorrectaUltimoNivel(RC, SQL) Then
            txtAux(0).Text = RC
            txtAux(1).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            txtAux(0).Text = ""
            txtAux(1).Text = ""
            PonFoco txtAux(0)
        End If
Case 2, 3, 4
        If Index = 2 Then
            RC = "Año"
        Else
            If Index = 3 Then
                RC = "mes"
            Else
                RC = "importe"
            End If
        End If
        
        If Not IsNumeric(txtAux(Index).Text) Then
            MsgBox "El valor del " & RC & " debe de sr numérico.", vbExclamation
            PonFoco txtAux(Index)
            Exit Sub
        End If
        
        'Particularidades
        If Index = 2 Then
            If Val(txtAux(2).Text) < 1000 Then
                MsgBox "Año incorrecto", vbExclamation
                PonFoco txtAux(2)
                Exit Sub
            End If
        Else
            If (Val(txtAux(3).Text) < 1) Or (Val(txtAux(3).Text) > 12) Then
                MsgBox "Mes incorrecto", vbExclamation
                PonFoco txtAux(3)
                Exit Sub
            End If
        End If
End Select
End Sub

Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function

If Modo = 1 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD("numdiari", "tiposdiario", "numdiari", txtAux(0).Text, "T")
     If Datos <> "" Then
        MsgBox "Ya existe el diario : " & txtAux(0).Text, vbExclamation
        B = False
    End If
End If
DatosOK = B
End Function


Private Function SituarData1() As Boolean
    Dim i As Integer
    On Error GoTo ESituarData1
        'Actualizamos el recordset
        Do
            If adodc1.Recordset!codmacta = txtAux(0).Text Then
                If adodc1.Recordset!anopresu = txtAux(2).Text Then
                    i = adodc1.Recordset!mespresu
                    If i = txtAux(3).Text Then
                        'Mover el scroll
                        Exit Function
                    End If
                End If
            End If
            adodc1.Recordset.MoveNext
        Loop Until adodc1.Recordset.EOF
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function


Private Sub PonFoco(Objeto As Object)
On Error Resume Next
Objeto.SetFocus
If Err.Number <> 0 Then Err.Clear
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub

