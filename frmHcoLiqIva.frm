VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHcoLiqIVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación I.V.A."
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H00000000&
      Height          =   350
      Index           =   3
      Left            =   6360
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Compensado|N|N|||liqiva|Compensa|###,###,##0.00||"
      Text            =   "Dato2"
      Top             =   5760
      Width           =   825
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
      ForeColor       =   &H00000000&
      Height          =   350
      Index           =   2
      Left            =   5040
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Liquidacion|N|N|||liqiva|Resultado|###,###,##0.00||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1425
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3540
      TabIndex        =   16
      Top             =   30
      Width           =   915
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Liquidación"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   60
      TabIndex        =   12
      Top             =   30
      Width           =   3405
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
         TabIndex        =   14
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
   Begin VB.ComboBox Combo2 
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
      Left            =   2670
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Es complementaria|N|N|||liqiva|escomplem||S|"
      Top             =   5640
      Width           =   945
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
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Periodo|N|N|||liqiva|periodo||S|"
      Top             =   5640
      Width           =   1575
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
      Left            =   8790
      TabIndex        =   6
      Top             =   6510
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
      Left            =   9990
      TabIndex        =   7
      Top             =   6510
      Visible         =   0   'False
      Width           =   1035
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
      ForeColor       =   &H00000000&
      Height          =   350
      Index           =   1
      Left            =   3750
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Importe|N|N|||liqiva|importe|###,###,##0.00||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1425
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
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Año Liquidación|N|N|||liqiva|anoliqui|000|S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5265
      Left            =   60
      TabIndex        =   8
      Top             =   900
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   9287
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   23
      RowDividerStyle =   6
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
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
      Left            =   9990
      TabIndex        =   11
      Top             =   6510
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
      Height          =   420
      Left            =   120
      TabIndex        =   9
      Top             =   6480
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
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   4560
      Top             =   120
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
      Left            =   10320
      TabIndex        =   15
      Top             =   120
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
Attribute VB_Name = "frmHcoLiqIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 412

Private frmMens As frmMensajes

Private WithEvents frmLiq As frmModelo303Liq
Attribute frmLiq.VB_VarHelpID = -1
Private frmAsi As frmAsientosHco
Attribute frmAsi.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadB As String

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim PrimeraVez As Boolean

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
Dim I As Integer

    Modo = vModo

    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    B = (Modo = 0 Or Modo = 2)
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = Not B
    Next
    Combo1.visible = Not B
    Combo2.visible = Not B
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I
    Combo1.BackColor = vbWhite
    Combo2.BackColor = vbWhite
    
    'Prueba
    
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = B
    End If
    txtAux(0).Enabled = (Modo <> 4)
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
'    NumF = SugerirCodigoSiguiente
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.top + 260
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top
    End If
    
    txtAux(0).Text = NumF
    txtAux(1).Text = ""
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    LLamaLineas anc, 3
    
    
    'Ponemos el foco
    PonFoco txtAux(0)

End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "anoliqui = -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    Me.Combo1.ListIndex = -1
    Me.Combo2.ListIndex = -1
    LLamaLineas DataGrid1.top + 250, 1
    PonFoco txtAux(0)
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


    If adodc1.Recordset!CodConce > 899 Then
        MsgBox "La aplicación se reserva los 100 ultimos conceptos", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    DeseleccionaGrid
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    I = adodc1.Recordset!TipoConce
    Combo1.ListIndex = I - 1
    If DBLet(adodc1.Recordset!EsEfectivo340, "T") = "" Then
        Combo2.ListIndex = 0
    Else
        Combo2.ListIndex = 1
    End If
  
    LLamaLineas anc, 4
   
   'Como es modificar
   PonFoco txtAux(1)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
PonerModo xModo
'Fijamos el ancho
txtAux(0).top = alto
txtAux(1).top = alto
Combo1.top = alto - 15
Combo2.top = alto - 15

End Sub




Private Sub BotonEliminar()
Dim Sql As String
Dim Actualizar As Boolean
Dim Ano As Integer
Dim Periodo As Integer
Dim Rs As ADODB.Recordset
Dim Mc As Contadores
Dim SqlLog As String

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    '### a mano
    Sql = "Seguro que desea eliminar la liquidación:"
    Sql = Sql & vbCrLf & "Año: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Período: " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Tipo: " & adodc1.Recordset.Fields(4)
    
    Conn.BeginTrans
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        Sql = "Select numasien, numdiari, fechaent  from liqiva where anoliqui=" & adodc1.Recordset!anoliqui
        Sql = Sql & " and periodo = " & DBSet(adodc1.Recordset!Periodo, "N")
        Sql = Sql & " and escomplem = " & DBSet(adodc1.Recordset!escomplem, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not Rs.EOF Then
            If DBLet(Rs!NumAsien, "N") <> 0 Then
                Sql = "Se va a proceder a eliminar el asiento correspondiente." & vbCrLf & vbCrLf
                Sql = Sql & "Diario: " & Rs!NumDiari & vbCrLf
                Sql = Sql & "Asiento: " & Rs!NumAsien & vbCrLf
                Sql = Sql & "Fecha: " & Rs!FechaEnt & vbCrLf
                
                MsgBox Sql, vbExclamation
                
                
                Sql = "delete from hlinapu where numdiari = " & DBSet(Rs!NumDiari, "N")
                Sql = Sql & " and numasien = " & DBSet(Rs!NumAsien, "N")
                Sql = Sql & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                
                Conn.Execute Sql
                
                Sql = "delete from hcabapu where numdiari = " & DBSet(Rs!NumDiari, "N")
                Sql = Sql & " and numasien = " & DBSet(Rs!NumAsien, "N")
                Sql = Sql & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                
                Conn.Execute Sql
                
                ' faltaria ver si devolvemos el contador
                Set Mc = New Contadores
                Mc.DevolverContador 0, DBLet(Rs!FechaEnt, "F") <= vParam.fechafin, Rs!NumAsien, True
                Set Mc = Nothing
                
                SqlLog = "Asiento Eliminado: " & DBLet(Rs!NumAsien, "N") & " - " & DBLet(Rs!FechaEnt, "F") & " - " & DBLet(Rs!NumDiari, "N")
            End If
        End If
        
        
        
        Sql = "Delete from liqiva where anoliqui=" & adodc1.Recordset!anoliqui
        Sql = Sql & " and periodo = " & DBSet(adodc1.Recordset!Periodo, "N")
        Sql = Sql & " and escomplem = " & DBSet(adodc1.Recordset!escomplem, "N")
        
        Conn.Execute Sql
        
        Actualizar = False
        If vParam.periodos = "1" Then ' mensual
            ' es la ultima liquidacion
            If vParam.Anofactu = adodc1.Recordset!anoliqui And vParam.perfactu = adodc1.Recordset!Periodo Then
                Periodo = vParam.perfactu - 1
                Ano = vParam.Anofactu
                If Periodo = 0 Then
                    Periodo = 12
                    Ano = Ano - 1
                End If
                Actualizar = True
            End If
        Else ' trimestral
            ' es la ultima liquidacion
            If vParam.Anofactu = adodc1.Recordset!anoliqui And vParam.perfactu = (Me.adodc1.Recordset!Periodo - 12) Then
                Periodo = adodc1.Recordset!Periodo - 13
                Ano = vParam.Anofactu
                If Periodo = 0 Then
                    Periodo = 4
                    Ano = Ano - 1
                End If
                Actualizar = True
            End If
        End If
        
        If Actualizar Then
            Sql = "update parametros set anofactu = " & DBSet(Ano, "N") & ", perfactu= " & DBSet(Periodo, "N")
            Conn.Execute Sql
            
            SqlLog = SqlLog & vbCrLf & "Ultima liquidacion: " & vParam.perfactu & "/ " & vParam.Anofactu & " ---> " & Periodo & "/" & Ano
            vParam.Anofactu = Ano
            vParam.perfactu = Periodo
           
        End If
        'Insertamos log
        If Actualizar Then
        vLog.Insertar 16, vUsu, "Liquidacion : " & Ano & " / " & adodc1.Recordset.Fields(2) & vbCrLf & SqlLog
            
        End If
        
    End If
    
    Conn.CommitTrans
    CargaGrid ""
    adodc1.Recordset.Cancel
    
    Exit Sub
    
Error2:
    Conn.RollbackTrans
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub


Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If adReason = adRsnMove And adStatus = adStatusOK Then PonLblIndicador Me.lblIndicador, adodc1
End Sub


Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
    Select Case Modo
    Case 1
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
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
                I = adodc1.Recordset.Fields(0)
                PonerModo 0
                CargaGrid
                adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
            End If
        End If
    End Select


End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1
            CargaGrid
        Case 3
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
            
    End Select
    PonerModo 0
    lblIndicador.Caption = ""
    DataGrid1.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation
        Exit Sub
    End If
    
    cad = adodc1.Recordset.Fields(0) & "|"
    cad = cad & adodc1.Recordset.Fields(1) & "|"
    cad = cad & adodc1.Recordset.Fields(2) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_GotFocus()
    If Modo = 1 Then
        Combo1.BackColor = vbLightBlue
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
    Combo1.BackColor = vbWhite
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo2_GotFocus()
    If Modo = 1 Then
        Combo2.BackColor = vbLightBlue
    End If
End Sub

Private Sub Combo2_LostFocus()
    Combo2.BackColor = vbWhite
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
    
    If adodc1.Recordset.EOF Then Exit Sub
    
    If IsNull(adodc1.Recordset!NumAsien) Then Exit Sub
    
    
    If Val(adodc1.Recordset!NumAsien) <> 0 Then
        Set frmAsi = New frmAsientosHco
        
        frmAsi.Asiento = adodc1.Recordset!NumDiari & "|" & adodc1.Recordset!FechaEnt & "|" & adodc1.Recordset!NumAsien & "|"
        frmAsi.SoloImprimir = True
        frmAsi.Show vbModal
        
        Set frmAsi = Nothing
    End If
    
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++
Private Sub Form_Load()

    Me.Icon = frmppal.Icon

    PrimeraVez = True

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
        .Buttons(1).Image = 48
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With


    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargaCombo
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    
    
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    
    CadenaConsulta = "Select liqiva.anoliqui,periodo, CASE periodo WHEN 0 THEN 'Anual' WHEN 1 THEN 'Enero' WHEN 2 THEN 'Febrero' WHEN 3 THEN 'Marzo' WHEN 4 THEN 'Abril' WHEN 5 THEN 'Mayo' WHEN 6 THEN 'Junio' WHEN 7 THEN 'Julio' WHEN 8 THEN 'Agosto' WHEN 9 THEN 'Septiembre' WHEN 10 THEN 'Octubre' WHEN 11 THEN 'Noviembre' WHEN 12 THEN 'Diciembre' WHEN 13 THEN '1er trimestre' WHEN 14 THEN '2do trimestre' WHEN 15 THEN '3er trimestre' WHEN 16 THEN '4to trimestre' END, "
    CadenaConsulta = CadenaConsulta & " escomplem, CASE escomplem WHEN 0 THEN 'Normal' WHEN 1 THEN 'Complementaria' END, importe"
    CadenaConsulta = CadenaConsulta & " ,resultado,compensa,numdiari, numasien, fechaent "
    CadenaConsulta = CadenaConsulta & " FROM liqiva "
    CadenaConsulta = CadenaConsulta & " WHERE (1=1) "
    
    CargaGrid
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
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla

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
        Case 8 'impresion
            frmModelo303.OpcionListado = 1
            frmModelo303.Show vbModal
        
        Case Else
    End Select
End Sub


Private Sub CargaGrid(Optional vSql As String)
Dim Sql As String
Dim tots As String

    
    
'    adodc1.ConnectionString = Conn
    If vSql <> "" Then
        Sql = CadenaConsulta & " AND " & vSql
    Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " order by anoliqui desc, periodo desc "
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Año|1000|;N||||0|;S|Combo1|C|Periodo|1700|;N||||0|;S|Combo2|C|Tipo|2000|;" '
    tots = tots & "S|txtAux(1)|T|Liquidacion|1830|;S|txtAux(2)|T|Periodo|1830|;S|txtAux(3)|T|Compensa|1830|;"


    tots = tots & "N||||0|;N||||0|;N||||0|;"
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight
    DataGrid1.Columns(7).Alignment = dbgRight
    DataGrid1.RowHeight = 350
    
    
    
    
End Sub



Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String
    
    Select Case Button.Index
        Case 1
            'Liquidacion de iva
            Set frmLiq = New frmModelo303Liq
            
            frmLiq.Show vbModal
            
            Set frmLiq = Nothing

            CargaGrid

        Case Else
    End Select

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'año
            PonerFormatoEntero txtAux(Index)
        Case 1 'importe
            PonerFormatoDecimal txtAux(Index), 3
    End Select
    

End Sub


Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
txtAux(1).Text = UCase(txtAux(1).Text)
B = CompForm(Me)
If Not B Then Exit Function

If Modo = 3 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtAux(0).Text, "N")
     If Datos <> "" Then
        MsgBox "Ya existe el concepto : " & txtAux(0).Text, vbExclamation
        B = False
    End If
End If
DatosOK = B
End Function

Private Sub CargaCombo()
    Combo1.Clear
    'periodo liquidado
    Combo1.AddItem "Enero"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Febrero"
    Combo1.ItemData(Combo1.NewIndex) = 2
    Combo1.AddItem "Marzo"
    Combo1.ItemData(Combo1.NewIndex) = 3
    Combo1.AddItem "Abril"
    Combo1.ItemData(Combo1.NewIndex) = 4
    Combo1.AddItem "Mayo"
    Combo1.ItemData(Combo1.NewIndex) = 5
    Combo1.AddItem "Junio"
    Combo1.ItemData(Combo1.NewIndex) = 6
    Combo1.AddItem "Julio"
    Combo1.ItemData(Combo1.NewIndex) = 7
    Combo1.AddItem "Agosto"
    Combo1.ItemData(Combo1.NewIndex) = 8
    Combo1.AddItem "Septiembre"
    Combo1.ItemData(Combo1.NewIndex) = 9
    Combo1.AddItem "Octubre"
    Combo1.ItemData(Combo1.NewIndex) = 10
    Combo1.AddItem "Noviembre"
    Combo1.ItemData(Combo1.NewIndex) = 11
    Combo1.AddItem "Diciembre"
    Combo1.ItemData(Combo1.NewIndex) = 12
    Combo1.AddItem "1er trimestre"
    Combo1.ItemData(Combo1.NewIndex) = 13
    Combo1.AddItem "2do trimestre"
    Combo1.ItemData(Combo1.NewIndex) = 14
    Combo1.AddItem "3er trimestre"
    Combo1.ItemData(Combo1.NewIndex) = 15
    Combo1.AddItem "4to trimestre"
    Combo1.ItemData(Combo1.NewIndex) = 16
    
    Combo2.Clear
    Combo2.AddItem "Normal"
    Combo2.ItemData(Combo2.NewIndex) = 0
    Combo2.AddItem "Complementaria"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim Sql As String
    
    SepuedeBorrar = False
    
    ' si no es complementaria miramos si es la ultima liquidacion
   If Val(adodc1.Recordset!escomplem) = 0 Then
        If vParam.periodos = "1" Then ' mensual
            ' es la ultima liquidacion
            If vParam.Anofactu = adodc1.Recordset!anoliqui And vParam.perfactu = Me.adodc1.Recordset!Periodo Then
                SepuedeBorrar = True
            End If
        Else ' trimestral
            ' es la ultima liquidacion
            If vParam.Anofactu = adodc1.Recordset!anoliqui And vParam.perfactu = (Me.adodc1.Recordset!Periodo - 12) Then
                SepuedeBorrar = True
            End If
        End If
    Else
        SepuedeBorrar = True
    End If
    If Not SepuedeBorrar Then
        MsgBox "No es la última liquidación y no se permite eliminar.", vbExclamation
    
    End If

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



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = (vUsu.Nombre = "root") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!Especial, "N")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



