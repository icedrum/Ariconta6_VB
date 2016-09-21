VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmContadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contadores"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   12735
   Icon            =   "frmContadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmContadores.frx":000C
      Left            =   6450
      List            =   "frmContadores.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5640
      Width           =   5190
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3810
      TabIndex        =   14
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   180
         TabIndex        =   15
         Top             =   150
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   60
      TabIndex        =   11
      Top             =   30
      Width           =   3585
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   12
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   13
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
               Object.ToolTipText     =   "Comprobar Contadores"
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
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      ItemData        =   "frmContadores.frx":0010
      Left            =   5430
      List            =   "frmContadores.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Fra cta ajena|N|N|||contadores|FacliAjena|||"
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
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
      Height          =   350
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Tag             =   "Siguiente|N|N|0||contadores|contado2|||"
      Text            =   "Dato4"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
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
      Height          =   350
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Tag             =   "Cont. actual|N|N|0||contadores|contado1|||"
      Text            =   "Dato3"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
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
      Left            =   10230
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   11370
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
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
      Height          =   350
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Denominación|T|N|||contadores|nomregis||N|"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
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
      Height          =   350
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Tipo registro|T|N|||contadores|tiporegi||S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmContadores.frx":0026
      Height          =   5295
      Left            =   90
      TabIndex        =   10
      Top             =   840
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   9340
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
      Left            =   11370
      TabIndex        =   9
      Top             =   6390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   7
      Top             =   6285
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
         Left            =   210
         TabIndex        =   8
         Top             =   210
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6000
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
      Left            =   11910
      TabIndex        =   16
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
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   6390
      TabIndex        =   18
      Tag             =   "Tipo Fra|T|S|||contadores|codconce340|||"
      Text            =   "Dato4"
      Top             =   5670
      Width           =   1395
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
Attribute VB_Name = "frmContadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 104

Private CadenaConsulta As String
Private CadB As String


Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte

Dim Cont1Ant As String
Dim Cont2Ant As String
Dim SqlLog As String

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

    txtaux(0).Visible = Not B
    txtaux(1).Visible = Not B
    txtaux(2).Visible = Not B
    txtaux(3).Visible = Not B
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    DataGrid1.Enabled = B
    Combo1.Visible = Not B
    Combo2.Visible = Not B
    
    For I = 0 To txtaux.Count - 1
        txtaux(I).BackColor = vbWhite
    Next I
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    End If
    'Si estamo mod or insert
    txtaux(0).Enabled = (Modo <> 2)
    If Modo = 3 Then
        Combo2.ListIndex = -1
    End If
    
    ' solo dejamos elegir el tipo de factura en el caso de que sean clientes
    If Modo = 3 Or Modo = 4 Then
        Combo2.Enabled = (TotalRegistros("select count(*) from contadores where " & DBSet(txtaux(0).Text, "T") & " REGEXP '^[0-9]+$' = 0 ") <> 0)
    End If
    
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
    

    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    txtaux(1).Enabled = True
    
   
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.Top + 250
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If
    txtaux(0).Text = ""
    txtaux(1).Text = ""
    txtaux(2).Text = ""
    txtaux(3).Text = ""
    
    LLamaLineas anc, 3
    Combo1.ListIndex = 0
    Combo2.ListIndex = -1
    
    'Ponemos el foco
    PonFoco txtaux(0)

End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    cmdCancelar.Visible = True
    cmdCancelar.SetFocus
    CargaGrid "tiporegi = ' '"
    'Buscar
    txtaux(0).Text = "":    txtaux(1).Text = "": txtaux(2).Text = "": txtaux(3).Text = ""
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    LLamaLineas DataGrid1.Top + 250, 1
    PonFoco txtaux(0)
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim cad As String
    Dim anc As Single
    Dim I As Integer
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    'Peculiar
    

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
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If

    'Llamamos al form
    For I = 0 To 3
        txtaux(I).Text = DataGrid1.Columns(I).Text
    Next I
    If DataGrid1.Columns(4).Text = "" Then
        Combo1.ListIndex = 0
    Else
        Combo1.ListIndex = 1
    End If
    
    If DataGrid1.Columns(5).Text = "" Then
        Combo2.ListIndex = -1
    Else
        PosicionarCombo2 Combo2, DataGrid1.Columns(5).Text
    End If
    
    Cont1Ant = txtaux(2).Text
    Cont2Ant = txtaux(3).Text
    
    LLamaLineas anc, 4
   
   'a mano###
    If Adodc1.Recordset!tiporegi = "0" Or Adodc1.Recordset!tiporegi = "1" Then
        txtaux(1).Enabled = False
    Else
        'Como es modificar
        txtaux(1).Enabled = True
        PonFoco txtaux(1)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    PonerModo xModo
    'Fijamos el ancho
    txtaux(0).Top = alto
    txtaux(1).Top = alto
    txtaux(2).Top = alto
    txtaux(3).Top = alto
    Combo1.Top = alto
    Combo2.Top = alto
End Sub




Private Sub BotonEliminar()
Dim Sql As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset!tiporegi = "0" Or Adodc1.Recordset!tiporegi = "1" Or Adodc1.Recordset!tiporegi = "Z" Then
        MsgBox "Este contador no se puede eliminar", vbExclamation
        Exit Sub
    End If
    '### a mano
    Sql = "Seguro que desea eliminar el contador:"
    Sql = Sql & vbCrLf & "Código: " & Adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Denominación: " & Adodc1.Recordset.Fields(1)
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from contadores where tiporegi='" & Adodc1.Recordset!tiporegi & "'"
        Conn.Execute Sql
        CargaGrid ""
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
        End If
End Sub






Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If adReason = adRsnMove And adStatus = adStatusOK Then PonLblIndicador Me.lblIndicador, Adodc1
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
On Error GoTo EAceptar
Select Case Modo
    Case 1
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
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
                If Cont1Ant <> txtaux(2).Text Or Cont2Ant <> txtaux(3).Text Then
                    SqlLog = "Contador : " & txtaux(0).Text & " " & txtaux(1).Text & vbCrLf & vbCrLf
                    SqlLog = SqlLog & "Ejercicio " & RellenaABlancos("Valor Anterior", False, 15) & " " & RellenaABlancos("Valor Actual", False, 15) & vbCrLf
                    SqlLog = SqlLog & String(42, "-") & vbCrLf
                    If Cont1Ant <> txtaux(2).Text Then
                        SqlLog = SqlLog & "Actual    " & RellenaABlancos(Cont1Ant, False, 15) & " " & RellenaABlancos(txtaux(2), False, 15) & vbCrLf
                    End If
                    If Cont2Ant <> txtaux(3).Text Then
                        SqlLog = SqlLog & "Siguiente " & RellenaABlancos(Cont2Ant, False, 15) & " " & RellenaABlancos(txtaux(3), False, 15) & vbCrLf
                    End If
                    
                    vLog.Insertar 25, vUsu, SqlLog
                End If
                        
                CadB = Adodc1.Recordset.Fields(0)
                DataGrid1.Enabled = False
                CargaGrid
                Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " = '" & CadB & "'")
                PonerModo 0
            End If
        End If
    End Select

Exit Sub
EAceptar:
    Err.Clear
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
    DataGrid1.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String

    If Adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation
        Exit Sub
    End If
    
    If Asc(Adodc1.Recordset.Fields(0)) <= 57 Then
        MsgBox "No es contador de tipo factura.", vbExclamation
        Exit Sub
    End If
    
    
    cad = Adodc1.Recordset.Fields(0) & "|"
    cad = cad & Adodc1.Recordset.Fields(1) & "|"
    cad = cad & Adodc1.Recordset.Fields(2) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Combo2_Validate(Cancel As Boolean)
    If (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        txtaux(4).Text = Chr(Combo2.ItemData(Combo2.ListIndex))
        If Combo2.ListIndex = -1 Or Chr(Combo2.ItemData(Combo2.ListIndex)) = "-" Then txtaux(4).Text = ""
    End If
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    Me.Icon = frmPpal.Icon
    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 20
    End With

    ' desplazamiento
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 26
    End With

    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    
    CargarCombo
    
    'Cadena consulta
    CadenaConsulta = "Select tiporegi,nomregis,contado1,contado2,if(FacliAjena=0,"""",""Si""), cc.descripcion from contadores left join usuarios.wconce340 cc on contadores.codconce340 = cc.codigo "
    CargaGrid
    PonerModo 2
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
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                BotonAnyadir
        Case 2
                BotonModificar
        Case 3
                BotonEliminar
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
        Case 8
                Screen.MousePointer = vbHourglass
                ComprobarContadores
                Screen.MousePointer = vbDefault
        Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub CargaGrid(Optional Sql As String)
Dim B As Boolean
    B = DataGrid1.Enabled
    CargaGrid2 Sql
    DataGrid1.Enabled = B
End Sub



Private Sub CargaGrid2(Optional Sql As String)
    Dim I As Integer
    Dim anc As Integer
    Adodc1.ConnectionString = Conn
    If Sql <> "" Then
        Sql = CadenaConsulta & " WHERE " & Sql
        Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " ORDER BY tiporegi"
    Adodc1.RecordSource = Sql
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockOptimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 350
    
    
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Tipo"
        DataGrid1.Columns(I).Width = 500
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Denominación"
        DataGrid1.Columns(I).Width = 4040
    
    'El importe es campo calculado
    I = 2
        DataGrid1.Columns(I).Caption = "Actual"
        DataGrid1.Columns(I).Width = 1100
        DataGrid1.Columns(I).Alignment = dbgRight
        
    I = 3
        DataGrid1.Columns(I).Caption = "Siguiente"
        DataGrid1.Columns(I).Width = 1100
        DataGrid1.Columns(I).Alignment = dbgRight
    
    I = 4
        DataGrid1.Columns(I).Caption = "Fac.ajena"
        DataGrid1.Columns(I).Width = 1100
        DataGrid1.Columns(I).Alignment = dbgRight
     
    I = 5
        DataGrid1.Columns(I).Caption = "Tipo Factura para Modelo 340"
        DataGrid1.Columns(I).Width = 4100
        DataGrid1.Columns(I).Alignment = dbgLeft
        
     
    
    
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        anc = 73 '60
        For I = 0 To 3
            txtaux(I).Left = DataGrid1.Columns(I).Left + anc
            txtaux(I).Width = DataGrid1.Columns(I).Width - 15
        Next I
        I = 4
        Combo1.Left = DataGrid1.Columns(I).Left + anc
        Combo1.Width = DataGrid1.Columns(I).Width - 15
        I = 5
        Combo2.Left = DataGrid1.Columns(I).Left + anc
        Combo2.Width = DataGrid1.Columns(I).Width - 15
        CadAncho = True
    End If
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
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub


    If Index = 4 Then Exit Sub

    txtaux(Index).Text = Trim(txtaux(Index).Text)
    If txtaux(Index).Text = "" Then Exit Sub
    If Modo = 1 Then Exit Sub 'Busquedas
    If Index > 1 Then
        If Not IsNumeric(txtaux(Index).Text) Then
            MsgBox "Los contadores tiene que ser numéricos", vbExclamation
            Exit Sub
        End If
        Else
            If Index = 0 Then txtaux(0).Text = UCase(txtaux(0).Text)
    End If
End Sub


Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function

If InStr(1, txtaux(0).Text, " ") > 0 Then
    MsgBox "No se permiten blancos", vbExclamation
    Exit Function
End If


'Del 0 al 9 nos los reservamos.
'   0-Asientos 1- Proveedores 2.- Contador confirming
'
If IsNumeric(txtaux(0).Text) Then
    If Combo1.ListIndex = 1 Then MsgBox "Facturación por cuenta ajena valido solo para SERIES DE FACTURAS.", vbExclamation
End If
If Modo = 3 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD("tiporegi", "contadores", "tiporegi", txtaux(0).Text, "T")
     If Datos <> "" Then
        MsgBox "Ya existe el contador : " & txtaux(0).Text, vbExclamation
        B = False
    End If
End If
DatosOK = B
End Function


Private Function ComprobarContadores()
Dim I As Long
Dim F As Date
Dim Sql As String
Dim aux2 As String
Dim Aux As String
Dim MaxA As Long
Dim CadenaError As String
    Set miRsAux = New ADODB.Recordset

    '-------------------------------------------------------
    '-------------------------------------------------------
    'Probamos los contadores
    '----------------------------------------------------------
    '-------------------------------------------------------
    '
    ' Asientos
    'actual
    CadenaError = ""
    NumRegElim = 1
    I = 2
    F = vParam.fechaini
    Sql = "Select max(numasien) from hlinapu where fechaent>='" & Format(F, FormatoFecha) & "'"
    F = vParam.fechafin
    Sql = Sql & " AND fechaent<='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        NumRegElim = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    
    '--------------------------------------
    'siguiente   --------------------------
    F = DateAdd("yyyy", 1, vParam.fechaini)
    Sql = "Select max(numasien) from hlinapu where fechaent>='" & Format(F, FormatoFecha) & "'"
    F = DateAdd("yyyy", 1, vParam.fechafin)
    Sql = Sql & " AND fechaent<='" & Format(F, FormatoFecha) & "'"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        I = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    'Ahora vemos si son correctos los contadores
    CadenaDesdeOtroForm = "contado2"
    Sql = DevuelveDesdeBD("contado1", "contadores", "tiporegi", "0", "T", CadenaDesdeOtroForm)
    If Sql <> "" Then
        If Val(Sql) <> NumRegElim Then
            Aux = "          Actual.     " & Aux & " : " & _
                 Format(NumRegElim, "000000") & "    --    Contadores: " & Format(Sql, "000000") & vbCrLf
            CadenaError = CadenaError & Aux
        End If
        If Val(CadenaDesdeOtroForm) <> I Then
            'Si es distinto de uno, pq el Nos reservamos el 1 para el asiento de apertura
            If I > 1 Or (I = 0 And Val(CadenaDesdeOtroForm) > 1) Then
                
                aux2 = "          Siguiente. " & aux2 & ": " & _
                     Format(I, "000000") & "    --    Contadores: " & Format(CadenaDesdeOtroForm, "000000") & vbCrLf
                CadenaError = CadenaError & aux2
            End If
        End If
        If CadenaError <> "" Then CadenaError = "ASIENTOS." & vbCrLf & CadenaError & vbCrLf & vbCrLf
    End If
    
    
    '-------------------------------------------------------
    '-------------------------------------------------------
    'Comprobaremos si querenos las facturas proveedores
    '-------------------------------------------------------
    '-------------------------------------------------------
    Dim Rs As ADODB.Recordset
    
    Me.Tag = ""
    
    Sql = "select * from contadores where tiporegi REGEXP '^[0-9]+$' <> 0 and cast(tiporegi as unsigned) > 0"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        NumRegElim = 1
        I = 2
        F = vParam.fechaini
        Sql = "Select max(numregis) from factpro where fecharec >='" & Format(F, FormatoFecha) & "'"
        F = vParam.fechafin
        Sql = Sql & " AND fecharec <='" & Format(F, FormatoFecha) & "'"
        Sql = Sql & " AND numserie = " & DBSet(Rs!tiporegi, "T")
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        'siguiente
        F = DateAdd("yyyy", 1, vParam.fechaini)
        Sql = "Select max(numregis) from factpro where fecharec >='" & Format(F, FormatoFecha) & "'"
        F = DateAdd("yyyy", 1, vParam.fechafin)
        Sql = Sql & " AND fecharec <='" & Format(F, FormatoFecha) & "'"
        Sql = Sql & " and numserie = " & DBSet(Rs!tiporegi, "T")
        
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            I = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        
        
        
        'Ahora vemos si son correcto los contadores
        CadenaDesdeOtroForm = "contado2"
        Sql = DevuelveDesdeBD("contado1", "contadores", "tiporegi", Rs!tiporegi, "T", CadenaDesdeOtroForm)
        If Sql <> "" Then
            aux2 = ""
            If Val(Sql) <> NumRegElim Then
                aux2 = aux2 & "          Actual:          " & Format(Sql, "000000") & "    --    Contadores: " & Format(NumRegElim, "000000") & vbCrLf
            End If
            If Val(CadenaDesdeOtroForm) <> I Then
                aux2 = aux2 & "          Siguiente:   " & Format(CadenaDesdeOtroForm, "000000") & "    --    Contadores: " & Format(I, "000000") & vbCrLf
            End If
            If aux2 <> "" Then Me.Tag = Me.Tag & "    -Registro: " & Rs!tiporegi & " - " & Rs!nomregis & vbCrLf & aux2 & vbCrLf
                
        End If
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    If Me.Tag <> "" Then CadenaError = CadenaError & vbCrLf & "FACTURAS PROVEEDORES" & vbCrLf & Me.Tag
    
    
    
    
    'Podemos comprobar las facturas tb
    'Por eso recorreremos el adodc1
    Me.Tag = ""
    
    Sql = "select * from contadores where tiporegi REGEXP '^[0-9]+$' = 0 "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        
        CadenaDesdeOtroForm = ""
        'Son FACTURAS DE CLIENTES
        'Actual
        
        I = DBLet(Rs!Contado1, "N")
        aux2 = MontaSQLFacCli(True)
        Aux = "Select max(numfactu) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
        Sql = Aux & " AND " & aux2
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        MaxA = 0
        If Not miRsAux.EOF Then MaxA = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        
        If I <> MaxA Then
            
            aux2 = "               Actual:      " & Format(MaxA, "000000") & "    -   " & " Contadores: " & I
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & aux2 & vbCrLf
            'MsgBox AUX2, vbExclamation
        End If


        'Actual
        I = DBLet(Rs!Contado2, "N")
        aux2 = MontaSQLFacCli(False)
        Aux = "Select max(numfactu) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
        Sql = Aux & " AND " & aux2
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        MaxA = 0
        If Not miRsAux.EOF Then MaxA = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        If I <> MaxA Then
            aux2 = "               Siguiente: " & Format(MaxA, "000000") & "    -   " & " Contadores: " & I
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & aux2 & vbCrLf
        End If


        If CadenaDesdeOtroForm <> "" Then Me.Tag = Me.Tag & "    -Registro: " & Rs!tiporegi & " - " & Rs!nomregis & vbCrLf & CadenaDesdeOtroForm & vbCrLf
        
        Rs.MoveNext
    Wend
    
    If Me.Tag <> "" Then CadenaError = CadenaError & vbCrLf & "FACTURAS CLIENTES" & vbCrLf & Me.Tag
    Me.Tag = ""
    
    'Fin comprobacion
    Set miRsAux = Nothing
    If CadenaError = "" Then
        MsgBox "Comprobacion finalizada", vbInformation
    Else
        MsgBox CadenaError, vbExclamation
    End If
End Function

Private Function MontaSQLFacCli(Actual As Boolean) As String
Dim I As Integer
Dim F As Date
    If Actual Then
        I = 0
    Else
        I = 1
    End If
    F = DateAdd("yyyy", I, vParam.fechaini)
    MontaSQLFacCli = "fecfactu >='" & Format(F, FormatoFecha) & "'"
    F = DateAdd("yyyy", I, vParam.fechafin)
    MontaSQLFacCli = MontaSQLFacCli & " AND fecfactu<='" & Format(F, FormatoFecha) & "'"
End Function


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String

    Combo2.Clear

    'Tipo de factura
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wconce340 ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ' simulamos el -1
    Combo2.AddItem "Ninguno"
    Combo2.ItemData(Combo2.NewIndex) = Asc("-")
    
    
    I = 0
    While Not Rs.EOF
        Combo2.AddItem Rs!Descripcion
        Combo2.ItemData(Combo2.NewIndex) = Asc(Rs!Codigo)
        I = I + 1
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing

End Sub
