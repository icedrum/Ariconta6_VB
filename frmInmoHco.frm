VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInmoHco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico inmovilizado"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInmoHco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   420
      Left            =   90
      TabIndex        =   17
      Top             =   6270
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   5790
      Width           =   195
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
      Left            =   90
      TabIndex        =   12
      Top             =   60
      Width           =   1485
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
         Left            =   3750
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   14
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   810
      TabIndex        =   11
      Top             =   5760
      Width           =   195
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   350
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6300
      Width           =   1575
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   350
      Index           =   4
      Left            =   4650
      TabIndex        =   3
      Tag             =   "Importe|N|N|0||shisin|imporinm|#,###,##0.00||"
      Text            =   "Dat"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   350
      Index           =   3
      Left            =   3810
      MaxLength       =   5
      TabIndex        =   2
      Tag             =   "Porcentaje|N|N|0||shisin|porcinm|##0.00||"
      Text            =   "Dat"
      Top             =   5760
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   350
      Index           =   2
      Left            =   2370
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||shisin|fechainm|dd/mm/yyyy|S|"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6945
      TabIndex        =   4
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8145
      TabIndex        =   6
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   870
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   5
      Text            =   "Dato2"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   350
      Index           =   0
      Left            =   30
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Código|N|N|0||shisin|codinmov|000000|S|"
      Text            =   "Dat"
      Top             =   5760
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8145
      TabIndex        =   7
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmInmoHco.frx":000C
      Height          =   5265
      Left            =   90
      TabIndex        =   8
      Top             =   930
      Width           =   9080
      _ExtentX        =   16007
      _ExtentY        =   9287
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   8760
      TabIndex        =   15
      Top             =   150
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total €"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   6330
      Width           =   1065
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
Attribute VB_Name = "frmInmoHco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 507


'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmEI As frmInmoElto
Attribute frmEI.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1


Private CadenaConsulta As String
Private TextoBusqueda As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim jj As Integer
Dim SQL As String
Dim PrimeraVez As Boolean

Private CadB As String

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas BUSCAR
'   Modo 2 -> Recorrer registros
'   Modo 3 -> Lineas  INSERTAR
'   Modo 4 -> Lineas MODIFICAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
    Modo = vModo
    
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    B = (Modo = 0 Or Modo = 2)
    
    For jj = 0 To 4
        txtaux(jj).Visible = Not B
    Next jj
    
    cmdAux(0).Visible = Not B
    cmdAux(1).Visible = Not B
    
    For I = 0 To txtaux.Count - 1
        If I <> 1 Then txtaux(I).BackColor = vbWhite
    Next I
    
    Toolbar1.Buttons(1).Enabled = B
    Toolbar1.Buttons(2).Enabled = B
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    
    txtaux(0).Enabled = (Modo <> 2)
    txtaux(2).Enabled = txtaux(0).Enabled
    txtaux(2).BackColor = txtaux(0).BackColor
    cmdAux(0).Enabled = txtaux(0).Enabled
    cmdAux(1).Enabled = txtaux(2).Enabled
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub BotonAnyadir()
    Dim anc As Single

    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Not Adodc1.Recordset.EOF Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
   
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.Top + 210
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If
    txtaux(0).Text = ""
    For jj = 1 To 4
        txtaux(jj).Text = ""
    Next jj
    LLamaLineas anc, 3
    
    'Ponemos el foco
    PonFoco txtaux(0)
    
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
    CargaGrid " inmovele_his.codinmov = -1"
    DataGrid1.Enabled = True
    'Buscar
    For jj = 0 To 4
        txtaux(jj).Text = ""
    Next jj
    LLamaLineas DataGrid1.Top + 250, 1
    PonFoco txtaux(0)
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim I As Integer
    If Adodc1.Recordset.EOF Then Exit Sub


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
        anc = DataGrid1.RowTop(DataGrid1.Row) + 600
    End If

    'Llamamos al form
    For jj = 0 To 2
        txtaux(jj).Text = DataGrid1.Columns(jj).Text
    Next jj
    'El porcentaje
    SQL = Adodc1.Recordset!porcinm
    txtaux(3).Text = TransformaComasPuntos(SQL)
        'El porcentaje
    SQL = Adodc1.Recordset!imporinm
    txtaux(4).Text = TransformaComasPuntos(SQL)
    
    LLamaLineas anc, 4
   
    'Como es modificar
    PonFoco txtaux(3)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid
    PonerModo xModo
    'Fijamos el ancho
    For jj = 0 To 4
        txtaux(jj).Top = alto
    Next jj
    cmdAux(0).Top = alto
    cmdAux(1).Top = alto
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    '### a mano
    SQL = "Seguro que desea eliminar la linea de histórico:" & vbCrLf
    SQL = SQL & vbCrLf & "Inmovilizado: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & Adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Fecha       : " & Adodc1.Recordset.Fields(2)
    SQL = SQL & vbCrLf & "Importe(€)  : " & Adodc1.Recordset.Fields(3)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from inmovele_his where codinmov=" & Adodc1.Recordset!Codinmov
        SQL = SQL & " AND fechainm ='" & Format(Adodc1.Recordset!fechainm, FormatoFecha) & "';"
        Conn.Execute SQL
        CargaGrid ""
        Adodc1.Recordset.Cancel
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
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        
        'Para el texto
        TextoBusqueda = ""
        If txtaux(0).Text <> "" Then TextoBusqueda = TextoBusqueda & "Cod. Inmov " & txtaux(0).Text
        If txtaux(2).Text <> "" Then TextoBusqueda = TextoBusqueda & "Fecha " & txtaux(2).Text
        If txtaux(3).Text <> "" Then TextoBusqueda = TextoBusqueda & "Porcentaje " & txtaux(3).Text
        If txtaux(4).Text <> "" Then TextoBusqueda = TextoBusqueda & "Importe " & txtaux(4).Text
        
        If CadB <> "" Then
            PonerModo 0
            DataGrid1.Enabled = False
            CargaGrid CadB
            DataGrid1.Enabled = True
        End If

    Case 3
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
    
    Case 4
        'Modificar
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If ModificaDesdeFormulario(Me) Then
                Conn.Execute "commit"
                I = Adodc1.Recordset.Fields(0)
                PonerModo 0
                CargaGrid
                Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & I)
            End If
        End If
    End Select
End Sub


Private Sub cmdAux_Click(Index As Integer)
Dim F As Date

    Select Case Index
        Case 0 ' elemento del inmovilizado
            Screen.MousePointer = vbHourglass
            Set frmEI = New frmInmoElto
            frmEI.DatosADevolverBusqueda = "0|1|"
            frmEI.Show vbModal
            Set frmEI = Nothing
            
            PonFoco txtaux(0)
        
        Case 1 ' fecha
            Set frmF = New frmCal
            F = Now
            If txtaux(2).Text <> "" Then F = CDate(txtaux(2).Text)
            
            frmF.Fecha = F
            frmF.Show vbModal
            Set frmF = Nothing
            
            PonFoco txtaux(2)

    End Select
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1
            CargaGrid
        Case 3
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
    End Select
    
    PonerModo 0
    lblIndicador.Caption = ""
    TextoBusqueda = ""
    DataGrid1.SetFocus

End Sub


Private Sub DataGrid1_DblClick()
'If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub


Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon

    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With

    PrimeraVez = True
    
    Set miTag = New CTag
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    
    CadenaConsulta = "SELECT inmovele_his.codinmov,inmovele.nominmov,fechainm,porcinm,imporinm "
    CadenaConsulta = CadenaConsulta & " FROM inmovele_his,inmovele WHERE "
    CadenaConsulta = CadenaConsulta & " inmovele_his.codinmov=inmovele.codinmov"
    
    CargaGrid "inmovele_his.codinmov = -1 "  'Para k lo carge vacio
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Set miTag = Nothing
End Sub


Private Sub frmEI_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtaux(2).Text = Format(vFecha, "dd/mm/yyyy")
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
                BotonBuscar
        Case 2
                BotonVerTodos
        Case Else
    End Select

End Sub



Private Sub CargaGrid(Optional vSQL As String)
Dim J As Integer
Dim TotalAncho As Integer
Dim I As Integer
Dim tots As String
    
    
    Text1.Text = ""
    Adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codinmov,fechainm"
    
    
    CargaGridGnral Me.DataGrid1, Me.Adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Código|1000|;S|cmdAux(0)|B|||;S|txtAux(1)|T|Descripción|3800|;S|txtAux(2)|T|Fecha|1455|;S|cmdAux(1)|B|||;"
    tots = tots & "S|txtAux(3)|T|Porcentaje|800|;S|txtAux(4)|T|Importe|1450|;"

    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft

    'Habilitamos modificar y eliminar
    CargarSumas vSQL
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'elemento
            Case 2: KEYBusqueda KeyAscii, 1 'fecha
            
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim Valor As Currency

    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub

    Select Case Index
    Case 0
        txtaux(1).Text = PonerNombreDeCod(txtaux(Index), "inmovele", "nominmov", "codinmov", "N")
        
        
    Case 2
        PonerFormatoFecha txtaux(Index)
    Case Else
    
        miTag.Cargar txtaux(Index)
        If miTag.Comprobar(txtaux(Index)) Then
            If Index > 2 Then
                'Son los numeros
                If InStr(1, txtaux(Index).Text, ",") > 0 Then
                    Valor = ImporteFormateado(txtaux(Index).Text)
                    SQL = CStr(Valor)
                Else
                     SQL = TransformaPuntosComas(txtaux(Index).Text)
                End If
                txtaux(Index).Text = Format(SQL, FormatoImporte)
            End If
            
        Else
            'Error con los datos
            txtaux(Index).Text = ""
            If Modo <> 0 Then txtaux(Index).SetFocus
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



Private Sub CargarSumas(ByRef vS As String)
On Error GoTo ECargarSumas

    Set miRsAux = New ADODB.Recordset
    SQL = "Select sum(imporinm) from inmovele_his"
    If vS <> "" Then SQL = SQL & " WHERE  " & vS
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then _
            Text1.Text = Format(miRsAux.Fields(0), FormatoImporte)
    End If
    miRsAux.Close
ECargarSumas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar sumas"
    Set miRsAux = Nothing
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim RS As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub



