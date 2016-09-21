VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCCoste 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de facturas"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frmMiFac2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmMiFac2.frx":000C
      Height          =   4845
      Left            =   225
      TabIndex        =   18
      Top             =   2100
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   8546
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7005
      Top             =   435
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
      Left            =   1440
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||CabFac|fecha|dd/mm/yyyy||"
      Text            =   "Text1"
      Top             =   810
      Width           =   1215
   End
   Begin VB.PictureBox picBus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   1080
      Picture         =   "frmMiFac2.frx":0021
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1290
      Width           =   240
   End
   Begin VB.TextBox txtInf 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1530
      Width           =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Tag             =   "Código de cliente|N|N|||CabFac|codcli|0000||"
      Text            =   "Text1"
      Top             =   1530
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegresar 
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   6300
      TabIndex        =   11
      Top             =   7410
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "Número de factura|N|N|||CabFac|numfac|0000|S|"
      Text            =   "Text1"
      Top             =   810
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   270
      TabIndex        =   7
      Top             =   7290
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6300
      TabIndex        =   5
      Top             =   7410
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7410
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4980
      TabIndex        =   3
      Top             =   7410
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
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
            Object.ToolTipText     =   "Modificar Lineas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":0123
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":0235
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":0347
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":0459
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":056B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":067D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":0F57
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":1831
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":210B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":29E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":32BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":3711
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":3823
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMiFac2.frx":3935
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   5610
      Top             =   420
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
   Begin VB.TextBox txtPseudo 
      Height          =   285
      Left            =   90
      TabIndex        =   17
      Top             =   2205
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Top             =   570
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del cliente:"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   14
      Top             =   1290
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Cod.Cli:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Num.Factura:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   570
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



'-----A mano: hay que definirse a mano los formularios refernciados
Private WithEvents FrmB1 As frmClientes
Attribute FrmB1.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmP As frmProdu
Attribute frmP.VB_VarHelpID = -1
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
'------------------------------------------------------------
'   Variables de referencia para búsquedas sobre otras tablas
Dim rCampo() As Integer '-- Campo Text1 que contiene la clave
Dim rTabla() As String '-- Tabla a la que referencia
Dim rClave() As String '-- Nombre en la tabla de referencia de la columna clave
Dim rColumna() As String '-- Valor de columna que queremos devolver
Private mRs As ADODB.Recordset


'------------- Definiciones para líneas ------------------
Private vCampos() As String ' Contiene los TAG de los campos que se desea mantener
Private vAncho() As Integer ' Contiene el ancho en porcentaje de cada columna
Private SQL As String
Dim mTag As CTag
Dim i As Integer
Dim ancho As Integer
Dim colMes As Integer

Dim CadAncho As String 'Para cuando llamemos al al form de lineas

'-------------------------------------------------------------


'Para pasar de lineas a cabeceras
Private ModificandoLineas As Boolean







Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    txtPseudo.Tag = ""
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                If SituarData1 Then
                    PonerModo 5
                    'Haremos como si pulsamo el boton de insertar nuevas lineas
                    cmdCancelar.Caption = "Factura"
                    AnyadirLinea 'CLng(Text1(0).Text)
                End If
            End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1 Then PonerModo 2
                    'lblIndicador.Caption = "Insertado"
                End If
            End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
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
    If Adodc1.Recordset.RecordCount > 0 Then
        'Modo=5
        If Adodc1.Recordset.EditMode = adEditAdd Or Adodc1.Recordset.EditMode = adEditInProgress Then
            Adodc1.Recordset.Cancel
            Data1.Recordset.Cancel
            CargaGrid Data1.Recordset!NumFac, False
        End If
        Data1.Recordset.CancelUpdate
    End If
    lblIndicador.Caption = NumRegistro & " de " & TotalReg
    PonerModo 2
End Select
cmdCancelar.Caption = "Cancelar"
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
            SQL = " numfac = " & Text1(0).Text
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
    CargaGrid -1, False
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    SugerirCodigoSiguiente
    '###A mano
    Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid -1, False
        
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        '### A mano
        '################################################
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
    CargaGrid -1, False
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
    '### a mano
    Cad = "Seguro que desea eliminar de la BD la factura:"
    Cad = Cad & vbCrLf & "Numero: " & Data1.Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields(2)
    i = MsgBox(Cad, vbQuestion + vbYesNo)
    'Borramos
    If i = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        'Borramos sus lineas de factura
        Cad = "Delete from linfac where numfac = " & Data1.Recordset!NumFac
        Conn.Execute Cad
        CargaGrid Data1.Recordset!NumFac, True
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
Dim j As Integer
Dim aux As String

If Data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

Cad = ""
i = 0
Do
    j = i + 1
    i = InStr(j, DatosADevolverBusqueda, "|")
    If i > 0 Then
        aux = Mid(DatosADevolverBusqueda, j, i - j)
        j = Val(aux)
        Cad = Cad & Text1(j).Text & "|"
    End If
Loop Until i = 0
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdSalir_Click()
    If Modo = 6 Then Adodc1.Recordset.Cancel
    Unload Me
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    LimpiarCampos
    'Si hay algun combo los cargamos
    CargarCombo
    '## A mano
    NombreTabla = "CabFac"
    Ordenacion = " ORDER BY numfac"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    
    
    'Bloqueo de tabla, cursor type
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    Data1.RecordSource = "Select * from " & NombreTabla
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
        Else
        PonerModo 1
        '### A mano
        Text1(0).BackColor = vbYellow
    End If
    '--- A mano // definir las relaciones externas
    ReDim rCampo(0)
    ReDim rTabla(0)
    ReDim rClave(0)
    ReDim rColumna(0)
    rCampo(0) = 2: rTabla(0) = "Clientes": rClave(0) = "codcli": rColumna(0) = "nomcli"
    '-------------
    
    CadAncho = ""
    'La lineas
    LineaInicio
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    'Combo1.ListIndex = -1
    'Check1.Value = 0
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
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        'If CadB <> "" Then CadB = CadB & " AND "
        'CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
DataGrid1.Columns(2).Text = RecuperaValor(CadenaSeleccion, 1)
DataGrid1.Columns(3).Text = RecuperaValor(CadenaSeleccion, 2)
DataGrid1.Columns(5).Text = RecuperaValor(CadenaSeleccion, 3)
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

Private Sub picBus_Click(Index As Integer)
    '-- A mano // llamadas a las búsquedas por prismático
    Select Case Index
        Case 0
            'No se puede si no es insertanod o modif..
            If Modo = 2 Or Modo = 0 Then Exit Sub
            '-- En este caso buscar proveedores
            Screen.MousePointer = vbHourglass
            Set FrmB1 = New frmClientes
            FrmB1.DatosADevolverBusqueda = "0|1|"
            FrmB1.Show vbModal
            Set FrmB1 = Nothing
    End Select
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
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
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
    
   '-------------------------------------------------------
    'Si queremos hacer algo ..
'    Select Case Index
'        Case 0
'
'        Case 1
'
'        '....
'    End Select
    '------------------------------------------------------------
    '-- A mano // control de los campos referenciales
    If Text1(Index) = "" Then Exit Sub
    For i = 0 To UBound(rCampo)
        If rCampo(i) = Index Then
            Set mTag = New CTag
            If mTag.Comprobar(Text1(rCampo(i))) Then
                SQL = "SELECT " & rColumna(i) & " FROM " & rTabla(i) & " WHERE " & rClave(i) & " = "
                Select Case mTag.TipoDato
                    Case "T"
                        SQL = SQL & "'" & Text1(rCampo(i)) & "'"
                    Case "F"
                        '-- Ojito (solo access)
                        SQL = SQL & "#" & Format(Text1(rCampo(i)), "yyyy/mm/dd") & "#"
                    Case Else
                        SQL = SQL & Text1(rCampo(i))
                End Select
                Set mRs = New ADODB.Recordset
                mRs.Open SQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
                If mRs.EOF Then
                    '-- No ha encontrado nada
                    MsgBox "No existe referencia en la tabla " & rTabla(i) & " para la clave " & Text1(rCampo(i)), vbInformation
                    Text1(Index).SetFocus
                Else
                    '-- Encontramos el registro
                    txtInf(i) = mRs.Fields(0)
                End If
            End If
            Exit For
        End If
    Next i
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
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 50, "Num, fac:")
        Cad = Cad & ParaGrid(Text1(1), 50, "Fecha")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSql = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1"
            frmB.vTitulo = "CabFac"
            frmB.vSelElem = 1
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                Text1(kCampo).SetFocus
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
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
    txtPseudo.Tag = ""
    PonerCamposForma Me, Data1
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = NumRegistro & " de " & TotalReg
    '-- A mano // Control de campos enlazados
    For i = 0 To UBound(rCampo)
        Set mTag = New CTag
        mTag.Cargar Text1(rCampo(i))
        SQL = "SELECT " & rColumna(i) & " FROM " & rTabla(i) & " WHERE " & rClave(i) & " = "
        Select Case mTag.TipoDato
            Case "T"
                SQL = SQL & "'" & Text1(rCampo(i)) & "'"
            Case "F"
                '-- Ojito (solo access)
                SQL = SQL & "#" & Format(Text1(rCampo(i)), "yyyy/mm/dd") & "#"
            Case Else
                SQL = SQL & Text1(rCampo(i))
        End Select
        Set mRs = New ADODB.Recordset
        mRs.Open SQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        If mRs.EOF Then
            '-- No ha encontrado nada
            MsgBox "No existe referencia en la tabla " & rTabla(i) & " para la clave " & Text1(rCampo(i)), vbInformation
        Else
            '-- Encontramos el registro
            txtInf(i) = mRs.Fields(0)
        End If
    Next i
    
    '##A mano
    'Cargamos el LINEAS
    CargaGrid Data1.Recordset!NumFac, False
    
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim i As Integer
    Dim b As Boolean
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
        Toolbar1.Buttons(6).ToolTipText = "Nueva factura"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar factura"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar factura"
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea fac"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea fac"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea fac"
    End If
    
    
    chkVistaPrevia.Visible = (Modo < 5)
    
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    DespalzamientoVisible b
    Toolbar1.Buttons(10).Enabled = b  'Lineas factur
    
    
   
    'Modo insertar o modificar
    b = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not b
    cmdAceptar.Visible = b Or Modo = 1
    '
    b = b Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not b
    Toolbar1.Buttons(2).Enabled = Not b
    mnOpciones.Enabled = Not b
   
   
        'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5
    'Modificar
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = b
    Else
        cmdRegresar.Visible = False
    End If
    b = b Or (Modo = 5)
    Toolbar1.Buttons(7).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(8).Enabled = b
    mnEliminar.Enabled = b

   
   
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    b = b Or Modo = 0   'En B tenemos modo=2 o a 5
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = b
        If b Then
            Text1(i).BackColor = &H80000018
        ElseIf Modo <> 1 Then
            Text1(i).BackColor = vbWhite
        End If
    Next i
    
    b = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = b
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
        cmdSalir.Cancel = True
    End If
    
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    
End Sub


Private Function DatosOk() As Boolean
Dim Rs As ADODB.Recordset
Dim b As Boolean
b = CompForm(Me)
DatosOk = b
End Function


'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    
    SQL = "Select Max(numfac) from " & NombreTabla
    Text1(0).Text = 1
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            Text1(0).Text = Rs.Fields(0) + 1
        End If
    End If
    Rs.Close
End Sub

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
    cmdCancelar.Caption = "Factura"
    Me.lblIndicador.Caption = "Lineas detalle"
    'CargaGrid Data1.Recordset!NumFac, True
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

'--- A mano // control de devoluciones de prismáticos
Private Sub FrmB1_DatoSeleccionado(CadenaSeleccion As String) '-- Proveedores
    If CadenaSeleccion <> "" Then
        Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
        txtInf(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


'--------------------- Controles para las líneas ----------------
Private Sub LineaInicio()


    'La SQL es la siguiente
    'SELECT LinFac.numfac, LinFac.numlin, LinFac.codprodu, Productos.nomprodu,
    'LinFac.cantidad, Productos.precio, [cantidad]*[precio] AS Exp
    'FROM Productos INNER JOIN LinFac ON Productos.codprodu = LinFac.codprodu;



    '-- ## A mano ----------------------
    ReDim vCampos(2)
    ReDim vAncho(2)   'En porcentaje
    vCampos(0) = "Linea:|N|N|||LinFac|numlin|000|S|": vAncho(0) = 10
    vCampos(1) = "Cod. produ:|N|N|||LinFac|nomcli||N|": vAncho(1) = 10
    vCampos(2) = "Cantidad:|N|N|||LinFac|cantidad||N|": vAncho(2) = 10
    colMes = 1
    '-------------------------------------
    CargaGrid -1, False  'Ponemos el -1 para k carge en blanco
End Sub



Private Sub CargaGrid(NumFac As Long, Bloqueado As Boolean)
    Dim j As Integer
    Dim TotalAncho As Integer
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = MontaSQLCarga(NumFac)
    Adodc1.CursorType = adOpenDynamic
    If Bloqueado Then
        'Adodc1.RecordSource = Adodc1.RecordSource & " FOR UPDATE"
        'Adodc1.RecordSource = Adodc1.RecordSource & " LOCK IN SHARE MODE;"
        Adodc1.LockType = adLockPessimistic
        Caption = "adLockPessimistic " & Adodc1.LockType
    Else
        Adodc1.LockType = adLockOptimistic
        Caption = "adLock Optimistic " & Adodc1.LockType
    End If
    Adodc1.Refresh
    
    ancho = DataGrid1.Width - 500
    Set mTag = New CTag
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-MNumfac    3.- Nomprodu    5.- Importe
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    'NUMFAC
    DataGrid1.Columns(0).Visible = False
    
    For j = 0 To 1
        i = j + 1
        txtPseudo.Tag = vCampos(j)
        mTag.Cargar txtPseudo
        DataGrid1.Columns(i).Caption = mTag.Nombre
        DataGrid1.Columns(i).NumberFormat = mTag.Formato
        If j = 0 Then
            DataGrid1.Columns(i).Width = 800
            Else
            DataGrid1.Columns(i).Width = 1200
        End If
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
    Next j
    'La columna 2 lleva un boton
    DataGrid1.Columns(2).Button = True
    
    'Nombre producto
    i = 3
        DataGrid1.Columns(i).Caption = "Producto"
        DataGrid1.Columns(i).Width = 2800
        DataGrid1.Columns(i).Locked = True
        
    'Leemos del vector en 2
    i = 2
        txtPseudo.Tag = vCampos(i)
        mTag.Cargar txtPseudo
        'Asignamos a la columna cuatro
        i = 4
        DataGrid1.Columns(i).Caption = mTag.Nombre
        DataGrid1.Columns(i).NumberFormat = mTag.Formato
        DataGrid1.Columns(i).Width = 700
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
    
    
    'El importe es campo calculado
    i = 5
        DataGrid1.Columns(i).Caption = "Precio UD."
        DataGrid1.Columns(i).NumberFormat = "#,##0.00"
        DataGrid1.Columns(i).Width = 1000
        'DataGrid1.Columns(i).Locked = True
        DataGrid1.Columns(i).Alignment = dbgRight
        
    i = 6
        DataGrid1.Columns(i).Caption = "Importe"
        DataGrid1.Columns(i).NumberFormat = "#,##0.00"
        DataGrid1.Columns(i).Width = 1200
        'DataGrid1.Columns(i).Locked = True
        DataGrid1.Columns(i).Alignment = dbgRight
        
    For i = 1 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).Locked = True
    Next i
    
    'Fiajamos el cadancho
    If CadAncho = "" Then
        For i = 0 To DataGrid1.Columns.Count - 1
            If DataGrid1.Columns(i).Visible Then
                CadAncho = CadAncho & DataGrid1.Columns(i).Width & "|"
            End If
        Next i
    End If
    
End Sub

Private Function MontaSQLCarga(vNumFac As Long) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    '--------------------------------------------------------------------
    Dim SQL As String

    SQL = "SELECT LinFac.numfac, LinFac.numlin, LinFac.codprodu, Productos.nomprodu,"
    SQL = SQL & " LinFac.cantidad, Productos.precio, cantidad*precio AS Exp"
    SQL = SQL & " FROM Productos ,LinFac "
    SQL = SQL & " WHERE Productos.codprodu = LinFac.codprodu"
    SQL = SQL & " AND LinFac.numFac = " & vNumFac
    SQL = SQL & " ORDER BY LinFac.NumLin"

    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea()
    Dim NumF As Long
    Dim anc As Long
    'Obtenemos la siguiente numero de factura
    NumF = ObtenerSigueinteNumeroFactura
    'Situamos el grid al final
    
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.MoveLast
    End If
    
    
    If DataGrid1.Row < 0 Then
        anc = 840
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 970
    End If
    With frm1LineaDe6
        .vFac = Data1.Recordset!NumFac
        .vCadena = NumF
        .vAnchoCampos = CadAncho
        .vBotonesVisibles = "N|N|S|N|N|"
        .vCamposHabilitados = "S|S|N|S|N|N|"
        .vModo = 0
        .vLeft = Me.Left + DataGrid1.Left + 360
        .vTop = Me.Top + DataGrid1.Top + anc
        FormularioHijoModificado = False
        .Show vbModal
    End With
    If FormularioHijoModificado Then
        CargaGrid Data1.Recordset!NumFac, False
        AnyadirLinea
        Else
            cmdCancelar.SetFocus
    End If
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 610
    End If
    Cad = ""
    For i = 1 To 6
        Cad = Cad & DataGrid1.Columns(i).Text & "|"
    Next i
    With frm1LineaDe6
        .vFac = Data1.Recordset!NumFac
        .vCadena = Cad
        .vAnchoCampos = CadAncho
        .vBotonesVisibles = "N|N|S|N|N|N|"
        .vCamposHabilitados = "S|S|N|S|N|N|"
        .vModo = 1
        .vLeft = Me.Left + DataGrid1.Left + 360
        .vTop = Me.Top + DataGrid1.Top + anc
        FormularioHijoModificado = False
        .Show vbModal
    End With

    If FormularioHijoModificado Then CargaGrid Data1.Recordset!NumFac, True
End Sub

Private Sub EliminarLineaFactura()
If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
SQL = "Seguro que desea eliminar la linea: " & Adodc1.Recordset!numlin & " "
SQL = SQL & Adodc1.Recordset!nomprodu & " - " & Adodc1.Recordset!cantidad & "?"
If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
    SQL = "Delete from linfac WHERE numfac =" & Data1.Recordset!NumFac
    SQL = SQL & " AND numlin=" & Adodc1.Recordset!numlin
    Conn.Execute SQL
    CargaGrid Data1.Recordset!NumFac, True
End If
End Sub



Private Function ObtenerSigueinteNumeroFactura() As Long
Dim Rs As ADODB.Recordset
Dim i As Long

Set Rs = New ADODB.Recordset
Rs.Open "SELECT Max(NumLin) FROM LinFac where numfac =" & Text1(0).Text, Conn, adOpenDynamic, adLockOptimistic, adCmdText
i = 0
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields(0)) Then i = Rs.Fields(0)
End If
Rs.Close
ObtenerSigueinteNumeroFactura = i + 1
End Function


