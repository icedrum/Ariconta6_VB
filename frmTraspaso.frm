VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTraspaso 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "frmTraspaso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   5760
      Width           =   3795
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   6120
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   5760
      Width           =   1395
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   6000
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
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
            Object.ToolTipText     =   "Generar datos traspaso ultimo nivel"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4200
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
      Bindings        =   "frmTraspaso.frx":000C
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   660
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
   Begin VB.Label Label1 
      Caption         =   "ULTIMO NIVEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   6120
      Width           =   3255
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
Attribute VB_Name = "frmTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'EL FORMULARIO ES EL MISMO PARA
'  traspaso
'  traspaso ultimo nivel

Public UltimoNivel As Boolean

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

txtAux(0).Visible = Not B
txtAux(1).Visible = Not B
txtAux(2).Visible = Not B
mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(11).Enabled = B

    'Generar solo es enabled en NO ultimo nivel y NO nuevo plan
    B = False
    If Not UltimoNivel And Not vParam.NuevoPlanContable Then B = True
    Toolbar1.Buttons(10).Enabled = B And mnOpciones.Enabled 'es como si fuera modo=0
    
    
    'Modificar y eso solo sera posible o bien es root, o bien
    'plan antiguo
    B = Not vParam.NuevoPlanContable And vUsu.Nivel < 2
    B = ((vUsu.Codigo Mod 1000) = 0 Or B) And mnOpciones.Enabled   'modo=0
    Toolbar1.Buttons(6).Enabled = B
    Toolbar1.Buttons(7).Enabled = B
    Toolbar1.Buttons(8).Enabled = B
    If Toolbar1.Buttons(10).Enabled Then Toolbar1.Buttons(10).Enabled = B
    


'
''Toolbar1.Buttons(6).Enabled = B And vUsu.Nivel < 2
'Me.mnNuevo.Enabled = Toolbar1.Buttons(6).Enabled
'Toolbar1.Buttons(7).Enabled = B And vUsu.Nivel < 2
'Me.mnModificar.Enabled = Toolbar1.Buttons(7).Enabled
'Toolbar1.Buttons(8).Enabled = B And vUsu.Nivel < 2
'Me.mnEliminar.Enabled = Toolbar1.Buttons(8).Enabled


B = (Modo = 0)
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid1.Enabled = B

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If


'Si estamo mod or insert
If Modo = 2 Then
   txtAux(0).BackColor = &H80000018
   Else
    txtAux(0).BackColor = &H80000005
End If
txtAux(0).Enabled = (Modo <> 2)
End Sub

Private Sub BotonAnyadir()

    Dim anc As Single
    

    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.Top + 220
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top + 20
    End If
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtAux(0).SetFocus

End Sub

Private Sub BotonBuscar()
    CargaGrid "codmactaold = -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    LLamaLineas 880, 2
    txtAux(0).SetFocus
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
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top + 20
    End If
'    Cad = ""
'    For I = 0 To 1
'        Cad = Cad & DataGrid1.Columns(I).Text & "|"
'    Next I
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(2).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text
    LLamaLineas anc, 1
   
   'Como es modificar
   PonleFoco txtAux(1)
   
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
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    'txtaux(0).Left = DataGrid1.Left + 340
    'txtaux(1).Left = txtaux(0).Left + txtaux(0).Width + 45
End Sub

Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    

    '### a mano
    SQL = "Seguro que desea eliminar la correspondencia de cuentas:"
    SQL = SQL & vbCrLf & "Antigua: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Nueva   : " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from straspaso"
        If Me.UltimoNivel Then SQL = SQL & "UltNivel"
        SQL = SQL & " Where codmactaold = " & adodc1.Recordset!codmactaold
        Conn.Execute SQL
        CancelaADODC
        CargaGrid ""
        SituarDespuesBorrar
        CancelaADODC
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Traspaso", Err.Description
End Sub

Private Sub SituarDespuesBorrar()
    On Error Resume Next
    NumRegElim = NumRegElim - 1
    If NumRegElim < 1 Then Exit Sub
    adodc1.Recordset.Move NumRegElim, 1
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub CancelaADODC()
On Error Resume Next
adodc1.Recordset.Cancel
If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdAceptar_Click()

Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
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
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    CadB = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CancelaADODC
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " = '" & CadB & "'")
                    'Esto es especial para este form
                    'Situaremos en el siguiente registro con modificar
                    If Not adodc1.Recordset.EOF Then
                        adodc1.Recordset.MoveNext
                        If Not adodc1.Recordset.EOF Then BotonModificar
                    End If
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
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""

DataGrid1.SetFocus
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

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If Modo = 0 Then Unload Me
    End If
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
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
    End With
    
    Caption = "Configurar traspaso P.G.C 2008"
    If UltimoNivel Then Caption = Caption & "     ULT. NIVEL"
    
           

    Label1.Visible = UltimoNivel
    
    PonerTags
    
    
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
    CadenaConsulta = "Select codmactaold,nommacta,codmactanew from straspaso"
    If Me.UltimoNivel Then CadenaConsulta = CadenaConsulta & "UltNivel"
    CadenaConsulta = CadenaConsulta & " left join cuentas on codmacta = codmactaold "
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
        'Esto no deberia de pasar
        If UltimoNivel Then Exit Sub
        
        'Generar ultimo nivel
        'Compruebo si ya lo han generado
        'Lo genero
        Set miRsAux = New ADODB.Recordset
        
        GenerarCambiosUltimoNivel
        
        CadenaDesdeOtroForm = ""
        DatosADevolverBusqueda = ""
        Set miRsAux = Nothing
        Screen.MousePointer = vbDefault
Case 11
        If TieneRegistrosTraspasos Then
            If Me.UltimoNivel Then
                CadenaDesdeOtroForm = "ULTIMO NIVEL"
            Else
                CadenaDesdeOtroForm = ""
            End If
            CadenaDesdeOtroForm = "Fechas= """ & CadenaDesdeOtroForm & """|"
            
            With frmImprimir
                .OtrosParametros = CadenaDesdeOtroForm
                .NumeroParametros = 1
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                '.Opcion = 34 + Cont
                .Opcion = 92
                .Show vbModal
            End With

        
        
        End If
Case 12
        Unload Me
Case Else

End Select
End Sub


Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer

    Dim I As Integer
    Dim B As Boolean
    
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codmactaold"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Cta antigua"
        DataGrid1.Columns(I).Width = 1150

    
    I = 1
        DataGrid1.Columns(I).Caption = "Título"
        DataGrid1.Columns(I).Width = 4700

    
    'Leemos del vector en 2
    I = 2
        DataGrid1.Columns(I).Caption = "Cta Nueva"
        DataGrid1.Columns(I).Width = 1150

            
    'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Left = DataGrid1.Columns(0).Left + 150
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(2).Left = DataGrid1.Columns(1).Left + 150
        txtAux(2).Width = DataGrid1.Columns(1).Width - 60
        txtAux(1).Left = DataGrid1.Columns(2).Left + 150
        txtAux(1).Width = DataGrid1.Columns(2).Width - 60
        CadAncho = True
    End If
   
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        If Toolbar1.Buttons(7).Enabled Then Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        If Toolbar1.Buttons(8).Enabled Then Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
   DataGrid1.Enabled = B
End Sub

Private Sub txtaux_GotFocus(Index As Integer)

    PonFoco txtAux(Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)

    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

txtAux(Index).Text = Trim(txtAux(Index).Text)
If txtAux(Index).Text = "" Then
    If Index = 0 Then txtAux(2).Text = ""
    Exit Sub
End If
If Modo = 3 Then Exit Sub 'Busquedas



    If Not UltimoNivel Then

        If Not IsNumeric(txtAux(Index).Text) Then
            MsgBox "Cuenta tiene que ser numérica: " & txtAux(Index).Text, vbExclamation
            txtAux(Index).Text = ""
            PonleFoco txtAux(Index)
            Exit Sub
        End If
        If Index = 0 Then
            CadenaDesdeOtroForm = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", txtAux(Index).Text, "T")
            If CadenaDesdeOtroForm = "" Then
                MsgBox "No existe la cuenta: " & txtAux(Index).Text, vbExclamation
                'txtAux(Index).Text = ""
                txtAux(2).Text = ""
            Else
                txtAux(2).Text = CadenaDesdeOtroForm
            End If
        End If
    Else
        
        If Index = 0 Or Index = 1 Then
            DatosADevolverBusqueda = txtAux(0).Text
            
            
            If Index = 0 Then
                'CODMACTA  OLD
                If CuentaCorrectaUltimoNivel(DatosADevolverBusqueda, CadenaDesdeOtroForm) Then
                    txtAux(0).Text = DatosADevolverBusqueda
                    
                Else
                    MsgBox CadenaDesdeOtroForm, vbExclamation
                    txtAux(0).Text = ""
                    CadenaDesdeOtroForm = ""
                    PonleFoco txtAux(0)
                End If
                
                txtAux(2).Text = CadenaDesdeOtroForm
            
            
            Else
            
                'CODMACTA NEW.
                CadenaDesdeOtroForm = txtAux(1).Text
                If Not IsNumeric(CadenaDesdeOtroForm) Then
                    MsgBox "Cuenta debe ser numérica: " & CadenaDesdeOtroForm, vbExclamation
                    DatosADevolverBusqueda = ""
                Else
                    DatosADevolverBusqueda = RellenaCodigoCuenta(CadenaDesdeOtroForm)
                    If InStr(1, DatosADevolverBusqueda, ".") > 0 Then
                        MsgBox "Error cuenta: " & DatosADevolverBusqueda, vbExclamation
                        DatosADevolverBusqueda = ""
                    End If
                    
                    If Len(DatosADevolverBusqueda) <> vEmpresa.DigitosUltimoNivel Then
                        MsgBox "No es cuenta ultimo nivel: " & DatosADevolverBusqueda, vbExclamation
                        DatosADevolverBusqueda = ""
                    End If
                End If
                txtAux(1).Text = DatosADevolverBusqueda
                If DatosADevolverBusqueda = "" Then PonleFoco txtAux(1)
                
            End If
            
            CadenaDesdeOtroForm = ""
            DatosADevolverBusqueda = ""
            
        End If
            
    End If
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function



    If Me.UltimoNivel Then
        B = False
        'ESOBLIGADA LA CUENTA
        If txtAux(1).Text = "" Then
            MsgBox "Cuentas obligadas", vbExclamation
            Exit Function
        End If
        
        If Len(txtAux(1).Text) <> vEmpresa.DigitosUltimoNivel Then
            MsgBox "No es cuenta ultimo nivel", vbExclamation
            Exit Function
        End If
        B = True
    Else
        If Len(txtAux(0).Text) > 5 Then
            MsgBox "5 digitos como máximo(antigua)", vbExclamation
            Exit Function
        End If
        If Len(txtAux(1).Text) > 5 Then
            MsgBox "5 digitos como máximo(nueva)", vbExclamation
            Exit Function
        End If
    End If
    
    If txtAux(1).Text <> "" Then
        Set miRsAux = New ADODB.Recordset
        Datos = "Select * from straspaso where codmactanew like '" & txtAux(1).Text & "%'"
        If Modo = 2 Then Datos = Datos & " AND codmactaold <> '" & txtAux(0).Text & "'"
        miRsAux.Open Datos, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Datos = ""
        If Not miRsAux.EOF Then
            Datos = miRsAux!codmactaold
            Datos = "La cuenta " & txtAux(1).Text & " ya esta relacionada con : " & Datos
            Datos = Datos & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(Datos, vbQuestion + vbYesNoCancel) = vbYes Then Datos = ""
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        If Datos <> "" Then Exit Function
    End If
DatosOk = B
End Function


Private Function SepuedeBorrar() As Boolean

    
    SepuedeBorrar = True
End Function




' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
 ' WheelHook DataGrid1

End Sub
Private Sub DataGrid1_LostFocus()
'  WheelUnHook
End Sub




Private Sub PonerTags()
    
    '
    'Cta nueva|T|S|||straspaso||||
    CadenaConsulta = ""
    If Me.UltimoNivel Then CadenaConsulta = "UltNivel"
    txtAux(0).Tag = "Cta antigua|T|N|||straspaso" & CadenaConsulta & "|codmactaold||S|"
    txtAux(1).Tag = "Cta nueva|T|S|||straspaso" & CadenaConsulta & "|codmactanew|||"
End Sub



Private Sub GenerarCambiosUltimoNivel()

    On Error GoTo EGenerarCambiosUltimoNivel
    
    
    'Compruebo que no se han efectuado cambios
    CadenaDesdeOtroForm = "Select count(*) from straspasoUltNivel"
    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = "Quiere generar la asignacion de cuentas a ultimo nivel?"
    NumRegElim = 0
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then
            CadenaDesdeOtroForm = "Ya se han generado la asignacion de cuentas a ultimo nivel. Borrar los datos y repetir el proceso? "
            NumRegElim = 1
        End If
    End If
    miRsAux.Close
    
    If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
  
    Screen.MousePointer = vbHourglass
  
    'Compruebo que no se van a solapar los de tres y cuatro digitos.
    'Esto es. Si he puesto  valores para la 120, y para la 1200 o la 1203, se solparian
    
    CadenaDesdeOtroForm = "DELETE from tmpcierre1 where codusu = " & vUsu.Codigo
    Conn.Execute CadenaDesdeOtroForm
    
    'Cargo las entradas a 4 digitos , pero cogiendo los 3 primeros
    CadenaDesdeOtroForm = "insert into `tmpcierre1` (`codusu`,`cta`) SELECT " & vUsu.Codigo
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ",substring(codmactaold,1,3)  from straspaso where codmactaold like '____' group by 1 "
    Conn.Execute CadenaDesdeOtroForm
    
    
    CadenaDesdeOtroForm = "select cta from straspaso left join tmpcierre1 on codusu= " & vUsu.Codigo
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " and cta = codmactaold where not (cta  is null)"
    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    CadenaDesdeOtroForm = ""
    While Not miRsAux.EOF
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux!Cta & "      "
        NumRegElim = NumRegElim + 1
        If (NumRegElim Mod 5) = 0 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Ha indicado cuentas a cuatro digitos  para las  cuentas : " & vbCrLf & CadenaDesdeOtroForm, vbExclamation
        Exit Sub
    End If
    'Primero elimino todos los registros que hubieran
    Conn.Execute "DELETE FROM straspasoUltNivel"
    
    
        
    CadenaDesdeOtroForm = "select codmactaold,codmactanew from straspaso where codmactanew like '___%'"
    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DatosADevolverBusqueda = ""
    While Not miRsAux.EOF
        'Esto va a ser sencillo. Raiz
        CadenaDesdeOtroForm = miRsAux.Fields(1)
        'Genero un UPDATE del estilo siguiente para cada cuenta
        NumRegElim = Len(CadenaDesdeOtroForm) + 1
        DatosADevolverBusqueda = "Select codmacta, concat('" & CadenaDesdeOtroForm & "',substring(codmacta," & NumRegElim & ","
        NumRegElim = vEmpresa.DigitosUltimoNivel - Len(CadenaDesdeOtroForm)
        DatosADevolverBusqueda = DatosADevolverBusqueda & NumRegElim & ")) from cuentas where codmacta like '"
        'Aqui va la cuenta OLD
        DatosADevolverBusqueda = DatosADevolverBusqueda & miRsAux.Fields(0) & "%' and apudirec ='S'"
        
        'INSERTAMOS EN straspasoNEW
        CadenaDesdeOtroForm = "INSERT INTO straspasoUltNivel(codmactaold,codmactanew) " & DatosADevolverBusqueda
        EjecutaSQL CadenaDesdeOtroForm
        
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If DatosADevolverBusqueda <> "" Then MsgBox "Se generaron las entradas a ultimo nivel para el traspaso de plan contable", vbExclamation
        
    Exit Sub
EGenerarCambiosUltimoNivel:
    
    MuestraError Err.Number, "Generar Cambios UltimoNivel"
End Sub



Private Function TieneRegistrosTraspasos() As Boolean
    On Error GoTo ETien
    NumRegElim = 0
    Set miRsAux = New ADODB.Recordset
    CadenaDesdeOtroForm = "Select count(*) from straspaso"
    If UltimoNivel Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "ultnivel"
    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        NumRegElim = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    
    'Cargo la tabla temporal donde mostrare las cuentas
    '-------------------------------------------------------
    CadenaDesdeOtroForm = "Delete from usuarios.zcuentas where codusu = " & vUsu.Codigo
    Conn.Execute CadenaDesdeOtroForm
    
    CadenaDesdeOtroForm = "INSERT INTO usuarios.zcuentas (`codusu`,`codmacta`,`nommacta`,`razosoci`) "
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " SELECT " & vUsu.Codigo & ",codmactaold,nommacta,codmactanew from straspaso"
    If Me.UltimoNivel Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "UltNivel"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " left join cuentas on codmacta = codmactaold "
    Conn.Execute CadenaDesdeOtroForm
    
ETien:
    If Err.Number <> 0 Then MuestraError Err.Number, "Calculando numero registros"
    Set miRsAux = Nothing
    TieneRegistrosTraspasos = NumRegElim > 0
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
End Function

