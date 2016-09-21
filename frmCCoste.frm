VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCCoste 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de centros de coste"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmCCoste.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Left            =   960
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Left            =   1080
      TabIndex        =   6
      Text            =   "Dato2"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Text            =   "Dato2"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Text            =   "Dat"
      Top             =   4680
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   675
      Left            =   4320
      TabIndex        =   5
      Top             =   540
      Width           =   2355
      Begin VB.CheckBox Check1 
         Caption         =   "Tiene subcentros reparto"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Tag             =   "Repartos|N|N|||cabccost|idsubcos|||"
         Top             =   270
         Width           =   2775
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCCoste.frx":000C
      Height          =   3705
      Left            =   180
      TabIndex        =   9
      Top             =   1320
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   6535
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4320
      Top             =   3360
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   582
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
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1380
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "FDenominación C.C.|T|N|||ccoste|nomccost|||"
      Text            =   "123456789012345678901234567890"
      Top             =   900
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   180
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Codigo C.C.|T|N|||ccoste|codccost||S|"
      Text            =   "Text1"
      Top             =   900
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   180
      TabIndex        =   11
      Top             =   5100
      Width           =   2775
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5340
      TabIndex        =   10
      Top             =   5220
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   5220
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   3960
      Top             =   2880
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
         Height          =   400
         Left            =   5880
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   5340
      TabIndex        =   8
      Top             =   5220
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   14
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   630
      Width           =   1215
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmCCoste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 1001

'-----A mano: hay que definirse a mano los formularios refernciados
Private WithEvents frmB As frmBasico
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmBB As frmBasico   'Este sera para buscar los datos
Attribute frmBB.VB_VarHelpID = -1
                                           ' del centro de coste.

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'   Nuevo modo --> Modificando lineas
'  5.- Modificando lineas

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private NumRegistro As Long
Private kCampo As Integer
Private TotalReg As Long
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Private SQL As String
Dim i As Integer
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim ModificaLineas As Byte  '1.- Añadir  2.- Modificar 0.-Pasar control a Lineas


Dim MaxValor As Single  'Valor maximo para
                        'Cuando % en centro de coste

Dim Rs As ADODB.Recordset
Dim AuxAnt As String


Private Sub Check1_Click()
    Dim Tiene As Boolean
    If Modo = 4 Then
        Tiene = False
        If Not adodc1.Recordset.EOF Then
            If adodc1.Recordset.RecordCount > 0 Then Tiene = True
        End If
        If Check1.Value = 0 And Tiene Then
            MsgBox "El centro de coste tiene asignados subcentros de coste. Eliminelos primero.", vbExclamation
            Check1.Value = 1
            Exit Sub
        End If
        If Check1.Value = 1 Then
            SQL = DevuelveDesdeBD("linscost", "ccoste_lineas", "subccost", Text1(0).Text, "T")
            If SQL <> "" Then
                MsgBox "El centro de coste ya esta incluido como subcentro en otros y no puede tener reparto", vbExclamation
                Check1.Value = 0
            End If
        End If
    End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                If SituarData1 Then
                    PonerModo 5
                    'Haremos como si pulsamo el boton de insertar nuevas lineas
                    cmdCancelar.Caption = "C. coste"
                    MaxValor = 100
                    AnyadirLinea 'CLng(Text1(0).Text)
                End If
            End If
        End If
    Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1 Then PonerModo 2
                    'lblIndicador.Caption = "Insertado"
                End If
            End If
    Case 5
            Cad = ""
            If txtaux(1).Text = "" Then Cad = "Escriba un % para el reparto"
            If txtaux2.Text = "" Then Cad = "No es un subcentro de coste correcto"
            If Not IsNumeric(txtaux(1).Text) Then Cad = "% de reparto debe de ser numérico"
            If Cad <> "" Then
                Screen.MousePointer = vbDefault
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
            If InsertarModificarLinea Then
                ObtenerMaxValor
                CargaGrid Data1.Recordset!codccost
                
                If MaxValor = 0 Then
                    cmdAceptar.Visible = False
                    DataGrid1.AllowAddNew = False
                    'Recalculo totales
                
                    txtaux(0).Visible = False
                    txtaux(1).Visible = False
                    txtaux2.Visible = False
                    cmdAux.Visible = False
                    ModificaLineas = 0
                    
                    cmdCancelar.Caption = "C. de coste"
                Else
                    AnyadirLinea
                End If
            End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click()
       Dim Cad As String
        'Llamamos a al form
        '##A mano
        HaDevueltoDatos = False
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 20, "C.Coste:")
        Cad = Cad & ParaGrid(Text1(1), 80, "Denominación:")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmBB = New frmBasico
'            frmBB.vCampos = Cad
'            frmBB.vTabla = NombreTabla
'            frmBB.vSQL = " idsubcos = 0"
'            HaDevueltoDatos = False
'            '###A mano
'            frmBB.vDevuelve = "0|1|"
'            frmBB.vTitulo = "Centro para % de reparto"
'            frmBB.vSelElem = 1
'            '#
'            frmBB.Show vbModal
'            Set frmBB = Nothing
'            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'            'tendremos que cerrar el form lanzando el evento
'            If HaDevueltoDatos Then
'                PonleFoco txtAux(1)
'                HaDevueltoDatos = False
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                PonleFoco txtAux(0)
'            End If
            
        End If
End Sub

Private Sub cmdCancelar_Click()

Select Case Modo
Case 1, 3
    LimpiarCampos
    PonerModo 0
Case 4
    PonerModo 2
    PonerCampos
Case 5
    If ModificaLineas = 0 Then
        'Estamos viendo las lineas, es decir nos vamos a cabeceras
        'Pero antes de irnos un detallito, que le % sea exactamente 100
        If PorcentajeCorrecto Then PonerModo 2
    
    Else
        cmdCancelar.Caption = "C. de coste"
        txtaux(0).Visible = False
        txtaux(1).Visible = False
        txtaux2.Visible = False
        cmdAux.Visible = False
        DataGrid1.AllowAddNew = False
        If ModificaLineas = 2 Then
            CargaGrid Data1.Recordset!codccost
            Else
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        End If
        ModificaLineas = 0
        cmdAceptar.Visible = False
        DataGrid1.Enabled = True
    End If
End Select
End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim SQL As String
    On Error GoTo ESituarData1
            'Actualizamos el recordset
            Data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = " codccost = '" & Text1(0).Text & "'"
            Data1.Recordset.Find SQL
            If Data1.Recordset.EOF Then GoTo ESituarData1
            SituarData1 = True
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid ""
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    '###A mano
    Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid ""
        
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
        Else
            HacerBusqueda
            If TotalReg = 0 Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Text1(kCampo).SetFocus
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid ""
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
        NumRegistro = 1
    Case 1
        Data1.Recordset.MovePrevious
        NumRegistro = NumRegistro - 1
        If Data1.Recordset.BOF Then
            Data1.Recordset.MoveFirst
            NumRegistro = 1
        End If
    Case 2
        Data1.Recordset.MoveNext
        NumRegistro = NumRegistro + 1
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveLast
            NumRegistro = TotalReg
        End If
    Case 3
        Data1.Recordset.MoveLast
        NumRegistro = TotalReg
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    
    '
    
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim i As Integer
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    If Not BorrarCC Then Exit Sub
    
    '### a mano
    Cad = "Seguro que desea eliminar de el centro de coste:"
    Cad = Cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Denominación: " & Data1.Recordset.Fields(1)
    i = MsgBox(Cad, vbQuestion + vbYesNo)
    'Borramos
    If i = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        'Borramos sus lineas de factura
        Cad = "Delete from ccoste where nomccost = '" & Data1.Recordset!codccost & "'"
        Conn.Execute Cad
        CargaGrid Data1.Recordset!codccost
        Data1.Recordset.Delete
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            PonerModo 0
            Else
                If NumRegistro = TotalReg Then
                        'He borrado el que era el ultimo
                        Data1.Recordset.MoveLast
                        NumRegistro = NumRegistro - 1
                        Else
                            For i = 1 To NumRegistro - 1
                                Data1.Recordset.MoveNext
                            Next i
                End If
                TotalReg = TotalReg - 1
                PonerCampos
        End If
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim AUx As String

If Data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

Cad = ""
i = 0
Do
    J = i + 1
    i = InStr(J, DatosADevolverBusqueda, "|")
    If i > 0 Then
        AUx = Mid(DatosADevolverBusqueda, J, i - J)
        J = Val(AUx)
        Cad = Cad & Text1(J).Text & "|"
    End If
Loop Until i = 0
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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

    LimpiarCampos
    'Asignamos toolbar
    With Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        
        .Buttons(10).Image = 11
        .Buttons(11).Image = 15
        
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With
    
    'Si hay algun combo los cargamos
    CargarCombo
    '## A mano
    NombreTabla = "ccoste"
    Ordenacion = " ORDER BY nomccost"
    CadAncho = False
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
'    Data1.UserName = vUsu.Login
'    Data1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    Data1.ConnectionString = Conn
    
    PonerOpcionesMenu
    
    'Bloqueo de tabla, cursor type
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
    'Data1.RecordSource = "Select * from ccoste ORDER BY nomccost"
    Data1.Refresh
    'Cargamos el grid
    CargaGrid "-11"
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        Text1(0).BackColor = vbYellow
    End If
    Me.Check1.Height = 520
    
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    'Combo1.ListIndex = -1
    Check1.Value = 0
    'Combo1.Text = ""
End Sub


Private Sub CargarCombo()
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'
'Ejemplo
''''''''Dim Rs As ADODB.Recordset
''''''''Set Rs = New ADODB.Recordset
''''''''
''''''''Rs.Open "TABLADONDESACARDATOS", Conn, adOpenForwardOnly, adLockOptimistic, adCmdTable
''''''''Combo1.Clear
''''''''While Not Rs.EOF
''''''''    Combo1.AddItem Rs!Nombre
''''''''    Combo1.ItemData(Combo1.newindex) = Rs!idSeccion
''''''''    'Siguiente
''''''''    Rs.MoveNext
''''''''Wend
''''''''Rs.Close
''''''''
''''''''ECargarCombo:
''''''''    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar combo"
''''''''    Set Rs = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Modo = 5 Then Cancel = 1
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim AUx As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        AUx = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = AUx
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        'If CadB <> "" Then CadB = CadB & " AND "
        'CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub frmBB_Selecionado(CadenaDevuelta As String)
    Dim C1 As String
    C1 = RecuperaValor(CadenaDevuelta, 1)
    If C1 <> AuxAnt Then
        If Not CentroYaAsignado(Text1(0).Text, RecuperaValor(CadenaDevuelta, 1)) Then
            txtaux(0).Text = RecuperaValor(CadenaDevuelta, 1)
            txtaux2.Text = RecuperaValor(CadenaDevuelta, 2)
            HaDevueltoDatos = True
        End If
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    If Modo <> 2 Then Exit Sub
    
    If CargaInforme(CadenaConsulta) Then
            frmImprimir.opcion = 94
            frmImprimir.FormulaSeleccion = "{ado.codusu} = " & vUsu.Codigo
            frmImprimir.NumeroParametros = 0
            frmImprimir.Show vbModal
    End If
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

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
        Else
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
    Dim i As Integer
    Dim SQL As String
    Dim mTag As CTag
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = &H80000018
    End If
    If Modo <> 1 Then _
        FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
    
    
    If Index = 0 Then Text1(0).Text = UCase(Text1(0).Text)
    
   '-------------------------------------------------------
    'Si queremos hacer algo ..
'    Select Case Index
'        Case 0
'
'        Case 1
'
'        '....
'    End Select
 
    '---
End Sub

Private Sub HacerBusqueda()
Dim Cad As String
    Dim CadB As String
    CadB = ObtenerBusqueda(Me)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
        Else
            'Se muestran en el mismo form
            If CadB <> "" Then
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " '& Ordenacion
                PonerCadenaBusqueda
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 20, "C.Coste:")
        Cad = Cad & ParaGrid(Text1(1), 80, "Denominación:")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBasico
'            frmB.vCampos = Cad
'            frmB.vTabla = NombreTabla
'            frmB.vSQL = CadB
'            HaDevueltoDatos = False
'            '###A mano
'            frmB.vDevuelve = "0|1"
'            frmB.vTitulo = "Centros de coste"
'            frmB.vSelElem = 1
'            '#
'            frmB.Show vbModal
'            Set frmB = Nothing
'            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'            'tendremos que cerrar el form lanzando el evento
'            If HaDevueltoDatos Then
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                    cmdRegresar_Click
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                Text1(kCampo).SetFocus
'            End If
        End If
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta & Ordenacion
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        TotalReg = 0
        Exit Sub
    
        Else
            PonerModo 2
            'Data1.Recordset.MoveLast
            Data1.Recordset.MoveFirst
            TotalReg = Data1.Recordset.RecordCount
            NumRegistro = 1
            PonerCampos
    End If
    
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
        MuestraError Err.Number, "PonerCadenaBusqueda"
        PonerModo 0
        Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    Dim i As Integer
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = NumRegistro & " de " & TotalReg
    
    '##A mano
    'Cargamos el LINEAS
    CargaGrid Data1.Recordset!codccost
    
    'Habilitamos o no el boton de lineas
    Toolbar1.Buttons(10).Enabled = (Check1.Value = 1)
    
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim i As Integer
    Dim B As Boolean
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For i = 0 To Text1.Count - 1
            'Text1(i).BackColor = vbWhite
            Text1(0).BackColor = &H80000018
        Next i
        'chkVistaPrevia.Visible = False
    End If
    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nuevo centro"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar centro"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar centro"
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea centro"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea centro"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea centro"
    End If
    
        

    chkVistaPrevia.Visible = (Modo < 5)
    
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B And vUsu.Nivel < 3 'Lineas factur
    
    
   
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 3
    cmdAceptar.Visible = B Or Modo = 1
    '
    B = B Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    mnOpciones.Enabled = Not B
   
   
        'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5
    'Modificar
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    Else
        cmdRegresar.Visible = False
    End If
    B = (B Or (Modo = 5)) And vUsu.Nivel < 2
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B

   
   
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    B = B Or Modo = 0   'En B tenemos modo=2 o a 5
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = B
        If B Then
            Text1(i).BackColor = &H80000018
        ElseIf Modo <> 1 Then
            Text1(i).BackColor = vbWhite
        End If
    Next i
    Frame2.Enabled = Not B
    B = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = B
    
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    
End Sub


Private Function DatosOK() As Boolean
    Dim B As Boolean
    B = CompForm(Me)
    DatosOK = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        BotonBuscar
    Case 2
        BotonVerTodos
    Case 6
        If Modo <> 5 Then
            BotonAnyadir
        Else
            'AÑADIR linea factura
            AnyadirLinea
        End If
    Case 7
        If Modo <> 5 Then
            BotonModificar
        Else
            'MODIFICAR linea factura
            ModificarLinea
        End If
    Case 8
        If Modo <> 5 Then
            BotonEliminar
        Else
            'ELIMINAR linea factura
            EliminarLineaFactura
        End If
    Case 10
        'Nuevo Modo
        PonerModo 5
        DataGrid1.Enabled = True
        'Obtenemos el máximo valor paa las lineas
        ObtenerMaxValor
        ModificaLineas = 0
        cmdCancelar.Caption = "C.Coste"
        Me.lblIndicador.Caption = "Lineas detalle"
        'CargaGrid Data1.Recordset!NumFac, True
    Case 11
        Unload Me
    Case 14 To 17
        Desplazamiento (Button.Index - 14)
    
    Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = bol
    Next i
End Sub


'--------------------- Controles para las líneas ----------------
Private Sub CargaGrid(NumFac As String)
    
On Error GoTo ECargaGRid
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = MontaSQLCarga(NumFac)
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockPessimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-MNumfac    3.- Nomprodu    5.- Importe
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    'NUMFAC
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Visible = False
    
    
    DataGrid1.Columns(2).Caption = "C. Coste"
    DataGrid1.Columns(2).Width = 1000
    
    
    DataGrid1.Columns(3).Caption = "Denominación sub centro"
    DataGrid1.Columns(3).Width = 3500
    
    DataGrid1.Columns(4).Caption = "% Reparto"
    DataGrid1.Columns(4).Width = 1000
    DataGrid1.Columns(4).NumberFormat = "0.00"
    DataGrid1.Columns(4).Alignment = dbgRight
    
    For i = 1 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).Locked = True
        DataGrid1.Columns(i).AllowSizing = False
    Next i
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        txtaux(0).Left = DataGrid1.Left + 340
        txtaux(0).Width = DataGrid1.Columns(2).Width - 60
        
        cmdAux.Left = txtaux(0).Left + txtaux(0).Width + 45
        
        txtaux2.Left = cmdAux.Left + cmdAux.Width
        txtaux2.Width = DataGrid1.Columns(3).Width - 30 - cmdAux.Width
        
        txtaux(1).Width = DataGrid1.Columns(4).Width - 30
        txtaux(1).Left = txtaux2.Left + txtaux2.Width + 45
        
        CadAncho = True
    End If
    Toolbar1.Buttons(10).Enabled = Not Data1.Recordset.EOF
    Exit Sub
ECargaGRid:
    MuestraError Err.Number, "Carga grid", Err.Description
End Sub

Private Function MontaSQLCarga(vNumC As String) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    '--------------------------------------------------------------------
    Dim SQL As String

    SQL = "SELECT ccoste_lineas.codccost,ccoste_lineas.linscost,ccoste_lineas.subccost, ccoste.nomccost,ccoste_lineas.porccost"
    SQL = SQL & " FROM ccoste_lineas INNER JOIN ccoste ON ccoste_lineas.subccost = ccoste.codccost"
    SQL = SQL & " WHERE (((ccoste_lineas.codccost)='" & vNumC & "'))"
    SQL = SQL & " ORDER BY ccoste_lineas.linscost;"
    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea()
    Dim NumF As Long
    Dim anc As Single

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub


    'Si el valor maximo del siguiente reparto es de 0 significa que no se puede añadir nada mas
    If MaxValor = 0 Then
        MsgBox "Ya se ha alcanzado el total de reparto.", vbExclamation
        Exit Sub
    End If

  'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    cmdAceptar.Tag = NumF
    cmdAceptar.Visible = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
    anc = DataGrid1.Top + 30
    If DataGrid1.Row >= 0 Then
        anc = DataGrid1.RowTop(DataGrid1.Row) + anc
    Else
        anc = anc + 200
    End If
    txtaux(0).Text = ""
    txtaux(1).Text = Format(MaxValor, "0.00")
    txtaux2.Text = ""
    LLamaLineas anc, 1

    'Ponemos el foco en la linea
    txtaux(0).SetFocus
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single

    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub
    
    Me.lblIndicador.Caption = "MODIFICAR"
   
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top + 30
    If DataGrid1.Row >= 0 Then
        anc = DataGrid1.RowTop(DataGrid1.Row) + anc
    End If
    Cad = ""
    'Llamamos al form
    cmdAceptar.Tag = adodc1.Recordset!linscost
    cmdAceptar.Visible = True
    txtaux(0).Text = DataGrid1.Columns(2).Text
    txtaux2.Text = DataGrid1.Columns(3).Text
    txtaux(1).Text = DataGrid1.Columns(4).Text
    
    LLamaLineas anc, 2
    txtaux(0).SetFocus
End Sub

Private Sub EliminarLineaFactura()
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas <> 0 Then Exit Sub
    
    SQL = "Seguro que desea eliminar els subcentro de coste: " & adodc1.Recordset!nomccost & " "
    SQL = SQL & "  Porcentaje : " & adodc1.Recordset!porccost & "?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = "Delete from ccoste_lineas WHERE codccost ='" & Data1.Recordset!codccost
        SQL = SQL & "' AND linscost=" & adodc1.Recordset!linscost
        Conn.Execute SQL
        'Recalculo totales
        ObtenerMaxValor
        CargaGrid Data1.Recordset!codccost
    End If
End Sub


Private Function SugerirCodigoSiguiente() As Integer


    SQL = "Select Max(linscost) from ccoste_lineas where codccost='" & Data1.Recordset!codccost & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SugerirCodigoSiguiente = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            SugerirCodigoSiguiente = Rs.Fields(0)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = SugerirCodigoSiguiente + 1
End Function


Private Sub LLamaLineas(alto As Single, xModo As Byte)
DataGrid1.Enabled = False
ModificaLineas = xModo
txtaux(0).Visible = True
txtaux(1).Visible = True
txtaux2.Visible = True
cmdAux.Visible = True
cmdCancelar.Caption = "Cancelar"
cmdAceptar.Caption = "Aceptar"
'Fijamos el ancho
cmdAux.Top = alto
txtaux(0).Top = alto
txtaux(1).Top = alto
txtaux2.Top = alto
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
With txtaux(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
    AuxAnt = .Text
End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim Cad As String
Dim Sng As Single

   txtaux(Index).Text = UCase(txtaux(Index).Text)
   If txtaux(Index).Text = AuxAnt Then
        
        Exit Sub
    End If
   If txtaux(Index).Text = "" Then Exit Sub
   If ModificaLineas = 0 Then Exit Sub
If Index = 0 Then
   txtaux(0).Text = UCase(txtaux(0).Text)
        'Reutilizacion de variables
    Cad = "idsubcos"
    SQL = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtaux(0).Text, "T", Cad)
    Sng = 0
    If SQL <> "" Then
        If Cad = "0" Then  'Tiene sub centro de coste
            txtaux2.Text = SQL
        Else
            SQL = "El subcentro de coste tiene reparto"
            Sng = 2
        End If
    Else
        SQL = "No existe el subcentro de coste para: " & Text1(0).Text
        Sng = 2
    End If
    If Sng > 1 Then
        MsgBox SQL, vbExclamation
        txtaux(0).Text = ""
        txtaux2.Text = ""
    End If
    
    'Ahora comprobamos que el centrod de coste no esta YA asignado
    If CentroYaAsignado(Text1(0).Text, txtaux(0).Text) Then
        MsgBox "El centro de coste ya ha sido asigenado en este reparto", vbExclamation
        txtaux(0).Text = ""
        txtaux2.Text = ""
        txtaux(0).SetFocus
    End If

    
Else
        If Not IsNumeric(txtaux(1).Text) Then
            MsgBox "El % de reparto debe de ser numérico", vbExclamation
            Exit Sub
        End If
        Sng = CSng(TransformaPuntosComas(txtaux(1).Text))
        
        If ModificaLineas = 2 Then
            Sng = Sng - adodc1.Recordset!porccost
        End If
                
        If Sng > MaxValor Then
            MsgBox "El porcentaje de reparto debe ser como máximo 100", vbExclamation
            If ModificaLineas Then
                txtaux(1).Text = adodc1.Recordset!porccost
            Else
                txtaux(1).Text = MaxValor
            End If
        End If
        
        txtaux(1).Text = TransformaPuntosComas(txtaux(1).Text)
        txtaux(1).Text = Format(txtaux(1).Text, "0.00")
        
End If
End Sub



Private Sub ObtenerMaxValor()
Dim Sng As Single
    Set Rs = New ADODB.Recordset
    SQL = "SELECT * FROM ccoste_lineas where codccost='" & Data1.Recordset!codccost & "'"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sng = 0
    While Not Rs.EOF
        Sng = Sng + Rs!porccost
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If Sng > 100 Then MsgBox "El valor total del % de reparto de los subcentros de coste excede de 100", vbCritical
    
    MaxValor = Round(100 - Sng, 2)

End Sub




Private Function InsertarModificarLinea() As Boolean

On Error GoTo EInsertarModificarCCoste
    InsertarModificarLinea = False
    If ModificaLineas = 1 Then
        SQL = "INSERT INTO ccoste_lineas VALUES ("
        SQL = SQL & "'" & Data1.Recordset!codccost & "',"
        SQL = SQL & cmdAceptar.Tag & ",'"
        SQL = SQL & txtaux(0).Text & "',"
        SQL = SQL & TransformaComasPuntos(txtaux(1).Text) & ")"
        Else
        SQL = "UPDATE ccoste_lineas Set subccost = '" & txtaux(0).Text
        SQL = SQL & "', porccost = " & TransformaComasPuntos(txtaux(1).Text)
        SQL = SQL & " WHERE codccost ='" & Data1.Recordset!codccost & "'"
        SQL = SQL & " AND linscost=" & cmdAceptar.Tag
    End If
    Conn.Execute SQL
    InsertarModificarLinea = True
    Exit Function
EInsertarModificarCCoste:
    MuestraError Err.Number, "Insertar/Modificar Cliente" & vbCrLf & Err.Description
End Function



Private Function PorcentajeCorrecto() As Boolean
Dim S As Single

On Error GoTo EPorcentajeCorrecto
    PorcentajeCorrecto = False
    SQL = "Select sum(porccost) from ccoste_lineas"
    SQL = SQL & " WHERE codccost  = '" & Text1(0).Text & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    S = 0
    If Not Rs.EOF Then
        S = DBLet(Rs.Fields(0), "N")
        S = Round(S, 3)
    End If
    Rs.Close
    Set Rs = Nothing
    If S > 0 And S <> 100# Then
        SQL = vbCrLf & "El porcentaje total del centro de coste  no suma 100. (" & Format(S, "0.00") & ")" & vbCrLf
        SQL = SQL & "Puede generar inconsistencia en los datos." & vbCrLf
        SQL = SQL & vbCrLf & vbCrLf & vbCrLf & "Desea salir de cualquier modo ?" & vbCrLf
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    PorcentajeCorrecto = True
Exit Function
EPorcentajeCorrecto:
    MuestraError Err.Number, "Obtener porcentaje.", Err.Description
End Function


Private Function CentroYaAsignado(Cod As String, SubCod As String) As Boolean
    On Error GoTo ECentroYaAsignado
    CentroYaAsignado = False
    SQL = "Select * from ccoste_lineas where codccost = '" & Cod & "'"
    SQL = SQL & " AND subccost ='" & SubCod & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then CentroYaAsignado = True
    Rs.Close
    Set Rs = Nothing
    Exit Function
ECentroYaAsignado:
        MuestraError Err.Number, "Comprobar asignacion centro.", Err.Description
End Function


Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub


Private Function BorrarCC() As Boolean
On Error GoTo EBorrarCC

    BorrarCC = False
    'En las lineas
    SQL = DevuelveDesdeBD("linscost", "ccoste_lineas", "subccost", Text1(0).Text, "T")
    If SQL <> "" Then
        MsgBox "El centro de coste esta asignado como subcentro a otros costes.", vbExclamation
        Exit Function
    End If

    'Referencias a eltos inmovilizado
    SQL = DevuelveDesdeBD("codinmov", "inmovele", "codccost", Text1(0).Text, "T")
    If SQL <> "" Then
        MsgBox "Tiene referencias en eltos. inmov", vbExclamation
        Exit Function
    End If

    'Referncias en reparto de gastos de inmovilizado
    SQL = DevuelveDesdeBD("codinmov", "inmovele_rep", "codccost", Text1(0).Text, "T")
    If SQL <> "" Then
        MsgBox "Tiene referencias a repartos en eltos. inmov", vbExclamation
        Exit Function
    End If


    'Cuenta bancaria
    SQL = DevuelveDesdeBD("codccost", "bancos", "codccost", Text1(0).Text, "T")
    If SQL <> "" Then
        MsgBox "Tiene referencias en cuentas bancarias", vbExclamation
        Exit Function
    End If



    BorrarCC = True
Exit Function
EBorrarCC:
    MuestraError Err.Number, "Comprobar borrar" & vbCrLf & Err.Description
End Function




Private Function CargaInforme(ByRef vSQL As String) As Boolean

On Error GoTo EGI_Conceptos
    CargaInforme = False
    'Borramos los anteriores
    Conn.Execute "Delete from Usuarios.zconceptos where codusu = " & vUsu.Codigo
    AuxAnt = "INSERT INTO Usuarios.zconceptos (codusu, codconce, nomconce,tipoconce) VALUES ("
    AuxAnt = AuxAnt & vUsu.Codigo & ",'"
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 0
    While Not Rs.EOF
        i = i + 1
        SQL = AuxAnt & DevNombreSQL(Rs.Fields(0))
        SQL = SQL & "','" & DevNombreSQL(Rs.Fields(1)) & "','')"
        Conn.Execute SQL
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    CargaInforme = True
Exit Function
EGI_Conceptos:
    MuestraError Err.Number
    Set Rs = Nothing
End Function

