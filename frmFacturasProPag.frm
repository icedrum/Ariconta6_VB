VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacturasProPag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vencimientos de la Factura"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   7155
   Icon            =   "frmFacturasProPag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   30
      TabIndex        =   18
      Top             =   60
      Width           =   7035
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   2490
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1860
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizar Pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   150
         TabIndex        =   2
         Top             =   1650
         Width           =   2235
      End
      Begin VB.TextBox Text1 
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
         Index           =   13
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1140
         Width           =   5880
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   420
         Width           =   5265
      End
      Begin VB.TextBox Text1 
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
         Index           =   26
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   420
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Index           =   29
         Left            =   1980
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   420
         Width           =   3720
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Pago"
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
         Left            =   2490
         TabIndex        =   23
         Top             =   1590
         Width           =   1260
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   3750
         Picture         =   "frmFacturasProPag.frx":000C
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   1650
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta banco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN"
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
         Index           =   26
         Left            =   150
         TabIndex        =   20
         Top             =   810
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdRegresar 
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
      Left            =   4770
      TabIndex        =   5
      Top             =   7290
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   3150
      MaskColor       =   &H00000000&
      TabIndex        =   15
      ToolTipText     =   "Buscar fecha"
      Top             =   4410
      Visible         =   0   'False
      Width           =   195
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
      Index           =   2
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Fecha Vencimiento|F|S|||tmppagos|fecvenci|dd/mm/yyyy||"
      Text            =   "1234567890"
      Top             =   4410
      Width           =   1065
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
      Index           =   1
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   8
      Tag             =   "Nro Orden|N|N|1|999999|tmppagos|numorden|0000|S|"
      Top             =   4410
      Width           =   705
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
      Left            =   3390
      MaxLength       =   12
      TabIndex        =   9
      Tag             =   "Importe Vto. |N|N|||tmppagos|impvenci|###,###,##0.00||"
      Top             =   4410
      Width           =   1185
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
      Left            =   120
      MaxLength       =   250
      TabIndex        =   7
      Tag             =   "Usuario|N|N|||tmppagos|codusu||S|"
      Top             =   4410
      Width           =   1095
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
      Left            =   4770
      TabIndex        =   10
      Tag             =   "   "
      Top             =   7290
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
      Left            =   5940
      TabIndex        =   11
      Top             =   7290
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacturasProPag.frx":0097
      Height          =   3780
      Left            =   90
      TabIndex        =   14
      Top             =   3180
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   6668
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
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   12
      Top             =   7080
      Width           =   2385
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   13
         Top             =   210
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2205
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   60
      TabIndex        =   16
      Top             =   2430
      Width           =   975
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "&Continuar"
      Enabled         =   0   'False
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
      Left            =   5850
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   1125
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
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnCargaLecturas 
         Caption         =   "&Cargar Lecturas"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnActualizar 
         Caption         =   "&Actualizar Contadores"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "&Filtro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "Con fecha actual"
         Index           =   1
      End
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "Sin fecha actual"
         Index           =   2
      End
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "Sin filtro"
         Checked         =   -1  'True
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmFacturasProPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado2(CadenaSeleccion As String)

Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBan As frmBasico2 ' cuenta de banco
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Dim Ordenacion As String
Dim ImporteFactura As Currency
Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

Dim FechaAnt As String
Dim Ok As Boolean
Dim CadB1 As String
Dim FILTRO As Byte
Dim SQL As String
Dim EsReciboBancario As Boolean

Dim CadB2 As String

Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    Frame2.Enabled = (Modo = 2)
    
    For i = 0 To txtaux.Count - 1
        txtaux(i).visible = (Modo = 1)
        txtaux(i).Enabled = (Modo = 1)
    Next i
    
    For i = 2 To 3
        txtaux(i).visible = (Modo = 1 Or Modo = 4)
        txtaux(i).Enabled = (Modo = 1 Or Modo = 4)
    Next i
    
    
    btnBuscar(0).visible = (Modo <> 2)
    btnBuscar(0).Enabled = (Modo <> 2)
    
    
    cmdCancelar.visible = B Or Modo = 4
    cmdRegresar.visible = B Or Modo = 4
    
    DataGrid1.Enabled = B
    
'    'Si es regresar
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
     txtaux(1).Enabled = (Modo = 4)
     
     If cmdRegresar.visible And Modo = 2 Then cmdRegresar.SetFocus
     
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


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    
    B = (Adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    
    
End Sub


Private Sub BotonVerTodos()
Dim Sql2 As String
Dim SQL As String

    CargaGrid "" 'CadB
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "tmppagos.codusu is null"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtaux.Count - 1
        txtaux(i).Text = ""
    Next i
    
    LLamaLineas DataGrid1.top + 206, 1 'Pone el form en Modo=1, Buscar
    PonFoco txtaux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top + 5 '+ 670 '545
    End If

    'Llamamos al form
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtaux(1).Text = DataGrid1.Columns(1).Text 'codsocio
    txtaux(2).Text = DataGrid1.Columns(2).Text
    txtaux(3).Text = DataGrid1.Columns(3).Text
    

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonFoco txtaux(3)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtaux.Count - 1
        txtaux(i).top = alto
    Next i
    For i = 0 To Me.btnBuscar.Count - 1
        btnBuscar(i).top = alto
    Next i
    ' ### [Monica] 12/09/2006
    
End Sub




Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el Cobro?"
    SQL = SQL & vbCrLf & "Código: " & Adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Adodc1.Recordset.AbsolutePosition
        
        SQL = "Delete from tmppagos where codusu=" & DBSet(vUsu.Codigo, "N") & " and numorden = " & Adodc1.Recordset.Fields(1)
        Conn.Execute SQL
        CargaGrid CadB
        
        temp = SituarDataTrasEliminar(Adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        Adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    Select Case Index
        Case 0 ' fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim Obj As Object
        
            Set frmC = New frmCal
            
            Indice = Index
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).top
                
            Set Obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> Obj.Name
                    esq = esq + Obj.Left
                    dalt = dalt + Obj.top
                    Set Obj = Obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            If txtaux(2).Text <> "" Then frmC.Fecha = txtaux(2).Text
            
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.top = dalt + btnBuscar(Index).Parent.top + btnBuscar(Index).Height + menu - 40
        
        
            btnBuscar(Index).Tag = Index '<===
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonFoco txtaux(2) '<===
            ' ********************************************
            
    End Select
    
End Sub

Private Sub Check1_Click()
    Text1(1).Enabled = (Check1.Value = 1)
    imgppal(0).Enabled = (Check1.Value = 1)
End Sub

Private Sub cmdAceptar_Click()
    Dim i As String
    Dim NReg As Long
    Dim SQL As String
    Dim Sql2 As String
    
    
    
    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
            
                CargaGrid "" ' CadB & AnyadeCadenaFiltro(True)
                PonerModo 2
                
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 4 'MODIFICAR
            
            Ok = False
            If DatosOK Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    Ok = True
                
                    TerminaBloquear
                    i = Adodc1.Recordset.Fields(1)
                    PonerModo 2
                    CargaGrid "" 'CadB
                    Adodc1.Recordset.Find (Adodc1.Recordset.Fields(1).Name & " ='" & i & "'")
                    PonerFocoGrid Me.DataGrid1
                    
                    
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 2
        
            'modificar
            TerminaBloquear
            ContinuarPago = False
            Unload Me
        Case 4
            LLamaLineas 0, 2
    End Select
    
End Sub

Private Sub CmdContinuar_Click()
    
    If Not DatosOKContinuar Then Exit Sub

    mnModificar_Click
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    ContinuarPago = True
    
    cmdAceptar_Click
    
    If Not DatosOKContinuar Then Exit Sub
    
    If Text1(26).Text = "" Then
        MsgBox "Debe introducir los datos del banco.", vbExclamation
        PonFoco Text1(26)
    Else
        If Check1.Value Then
            If Text1(1).Text = "" Then
                MsgBox "Debe introducir la fecha de contabilización del pago.", vbExclamation
                PonFoco Text1(1)
                Exit Sub
            End If
        End If
        'Text1(26).Text & "|" & Text1(13).Text & "|" & Text1(14).Text & "|" & Text1(15).Text & "|" & Text1(16).Text & "|" & Text1(29).Text & "|" & Me.Check1.Value & "|" & Me.Text1(1).Text & "|"
        RaiseEvent DatoSeleccionado2(Text1(26).Text & "|" & Replace(Text1(13).Text, " ", "") & "|" & Me.Check1.Value & "|" & Me.Text1(1).Text & "|")
        
        Unload Me
    End If
End Sub


Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim Cad As String

    If Adodc1.Recordset Is Nothing Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    Ordenacion = "ORDER BY " & DataGrid1.Columns(0).DataField
    
    CargaGrid CadB

    Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then PonerContRegIndicador
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        
        Else
            PonerModo 2
            If Me.CodigoActual <> "" Then
                SituarData Me.Adodc1, "numorden=" & CodigoActual, "", True
                
            End If
        End If
        
        ' cargamos los datos del banco
        Text1(13).Text = ""
      
        
        Text1(26).Text = RecuperaValor(CodigoActual, 1)
        Text1(29).Text = RecuperaValor(CodigoActual, 6)
       
        Text1(13).Text = DevuelveIBANSeparado(Text1(29).Text)
            
        EsReciboBancario = (RecuperaValor(CodigoActual, 7) = 4)
        Me.Text2(1).Text = RecuperaValor(CodigoActual, 8)
        Check1.Value = 0
        ImporteFactura = RecuperaValor(CodigoActual, 9)
        Text1(1).Text = Format(Now, "dd/mm/yyyy")
        Text1(1).Enabled = (Check1.Value = 1)
        imgppal(0).Enabled = (Check1.Value = 1)
        If Text1(26).Text <> "" Then
            PonleFoco cmdRegresar
        Else
            PonFoco Text1(26)
        End If
    End If
End Sub

Private Sub Form_Load()
Dim Sql2 As String

    PrimeraVez = True

    Me.Icon = frmppal.Icon

    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With

    Image1(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT tmppagos.codusu, tmppagos.numorden, tmppagos.fecvenci, tmppagos.impvenci "
    CadenaConsulta = CadenaConsulta & " FROM tmppagos "
    CadenaConsulta = CadenaConsulta & " WHERE tmppagos.codusu = " & vUsu.Codigo
    '************************************************************************
    
    Ordenacion = " ORDER BY codusu, numorden "
    
    CadB = ""
    CargaGrid
    
    ' podemos marcar de si se da por cobrado solo en el caso de haya un solo efecto
    Check1.visible = (Adodc1.Recordset.RecordCount = 1)
    Check1.Enabled = (Adodc1.Recordset.RecordCount = 1)
    
    FechaAnt = ""
    
    Modo = 2
    PonerModo Modo
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub Frame2_Click()
    Frame2.Enabled = True
End Sub


Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frame2.Enabled = True
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtaux(2).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub Image1_Click(Index As Integer)
    Set frmBan = New frmBasico2
    AyudaBanco frmBan
    Set frmBan = Nothing
End Sub

Private Sub imgppal_Click(Index As Integer)
    If (Modo = 5 Or Modo = 0) And (Index <> 6) And (Index <> 8) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        'FECHA FACTURA
        Indice = 1
        
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco Text1(1)
        
    End Select
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If Adodc1.Recordset.EOF Then Exit Sub
    
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, Adodc1, 1) Then BotonModificar
End Sub


Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim i As Integer
    Dim Sql2 As String, Sql3 As String
    Dim mTag As CTag
    Dim Im As Currency
    Dim Result As Byte
    Dim RC As String
    
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 13
            If Text1(Index).Text = "" Then Exit Sub
            If Not EsIBAN_Correcto(Text1(Index).Text) Then
                MsgBox "IBAN incorrecto", vbExclamation
                
            End If
        Case 26
            If Text1(26).Text = "" Then Exit Sub
            
            SQL = Text1(26).Text
            If CuentaCorrectaUltimoNivel(SQL, Sql2) Then
                SQL = DevuelveDesdeBD("codmacta", "bancos", "codmacta", SQL, "T")
                If SQL = "" Then
                    MsgBox "La cuenta NO pertenece a ningúna cta. bancaria", vbExclamation
                    Sql2 = ""
                Else
                    'CORRECTO
                End If
            Else
                SQL = ""
                MsgBox Sql2, vbExclamation
                Sql2 = ""
            End If
            Text1(26).Text = SQL
            Text2(1).Text = Sql2
            If SQL = "" Then PonleFoco Text1(26)
             
             
        Case 1 '1 - fecha de cobro
            SQL = ""
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                PonFoco Text1(Index)
                Exit Sub
            End If
            Result = FechaCorrecta2(CDate(Text1(Index).Text))
            If Result > 1 Then
                If Result = 2 Then
                    RC = varTxtFec
                Else
                    If Result = 3 Then
                        RC = "ya esta cerrado"
                    Else
                        RC = " todavia no ha sido abierto"
                    End If
                    RC = "La fecha pertenece a un ejercicio que " & RC
                End If
                MsgBox RC, vbExclamation
                Text1(Index).Text = ""
                If Index = 1 Then Text1(14).Text = ""
                PonFoco Text1(Index)
                Exit Sub
            End If
            
            Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
            
    End Select
    

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 26: KEYImage KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    PonFoco Text1(Index)
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
        Case 2
                mnModificar_Click
        Case 3
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSql As String)
    Dim SQL As String
    Dim tots As String
    Dim Sql2 As String
    
'    adodc1.ConnectionString = Conn
    If vSql <> "" Then
        SQL = CadenaConsulta & " AND " & vSql
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " " & Ordenacion
    '**************************************************************++
    
    
    CargaGridGnral Me.DataGrid1, Me.Adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "N||||0|;S|txtAux(1)|T|Orden|1000|;S|txtAux(2)|T|Fecha Vto|2250|;S|btnBuscar(0)|B||195|;"
    tots = tots & "S|txtAux(3)|T|Importe|3000|;"
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtaux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 3 ' importe de vto
            If Not PonerFormatoDecimal(txtaux(Index), 1) Then PonFoco txtaux(Index)
            
             
        Case 2 ' fecha de lectura actual
            PonerFormatoFecha txtaux(Index)
    End Select
    
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Mens As String
Dim NroDig As Integer
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim Limite As Long

    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtaux(0)) Then B = False
    End If
    
    If B And Modo = 4 Then
    End If
    
    
    DatosOK = B
End Function


Private Function DatosOKContinuar() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Mens As String
Dim NroDig As Integer
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim Limite As Long
Dim Index As Integer

    DatosOKContinuar = False
    
    If Text1(26).Text = "" And EsReciboBancario Then
        MsgBox "Debe de introducir la cuenta del Banco. Reintroduzca.", vbExclamation
        B = False
        PonleFoco Text1(26)
    Else
        If Text1(26).Text <> "" Then
            SQL = Text1(26).Text
            If CuentaCorrectaUltimoNivel(SQL, Sql2) Then
                SQL = DevuelveDesdeBD("codmacta", "bancos", "codmacta", SQL, "T")
                If SQL = "" Then
                    MsgBox "La cuenta NO pertenece a ningúna cta. bancaria", vbExclamation
                    Sql2 = ""
                Else
                    'CORRECTO
                    B = True
                End If
            Else
                SQL = ""
                MsgBox Sql2, vbExclamation
                Sql2 = ""
            End If
            Text1(26).Text = SQL
            Text2(1).Text = Sql2
            If SQL = "" Then
                B = False
                PonleFoco Text1(26)
            End If
        Else
            MsgBox "Cuanta bancaria obligatoria", vbExclamation
        End If
    
        If B And Me.Text1(13).Text <> "" Then
            If Not EsIBAN_Correcto(Text1(13).Text) Then
                If Mid(Text1(13).Text, 1, 2) = "ES" Then
                    MsgBox "IBAN incorrecto", vbExclamation
                    PonleFoco Text1(13)
                    B = False
                Else
                    If MsgBox("IBAN paraece incorrecto. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then B = False
                End If
            End If
        End If
        
        If B Then
            SQL = DevuelveDesdeBD("sum(impvenci)", "tmppagos", "codusu", CStr(vUsu.Codigo))
            If SQL = "" Then SQL = "0"
            
            If CCur(SQL) <> ImporteFactura Then
                SQL = "Introducido; " & SQL & vbCrLf
                SQL = "Factura: " & ImporteFactura & vbCrLf & SQL
                SQL = "Importes diferentes:" & vbCrLf & SQL & "¿Continuar?"
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then B = False
            End If
        End If
    End If
        
    DatosOKContinuar = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 3 Then ' estoy introduciendo la lectura
       If KeyAscii = 13 Then 'ENTER
            PonerFormatoDecimal txtaux(Index), 1
            If Modo = 4 Then
                cmdAceptar_Click
                'ModificarLinea

                If Ok Then PasarSigReg
                    
            End If
            If Modo = 1 Or Modo = 3 Then
                cmdAceptar.SetFocus
            End If
            
       ElseIf KeyAscii = 27 Then
            cmdCancelar_Click 'ESC
       End If
    Else
        KEYpress KeyAscii
    End If

End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    
    
    ' si no estamos en muestra salimos
    If Index <> 3 Then Exit Sub
    
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            cmdAceptar_Click
            If Ok Then PasarAntReg
        Case 40 'Desplazamiento Flecha Hacia Abajo
            cmdAceptar_Click
            
            If Ok Then PasarSigReg
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGrid1.Bookmark < Me.Adodc1.Recordset.RecordCount Then
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        BotonModificar
        PonFoco txtaux(3)
    ElseIf DataGrid1.Bookmark = Adodc1.Recordset.RecordCount Then
        BotonModificar
        PonFoco txtaux(3)
    End If
End Sub


Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGrid1.Bookmark > 1 Then
        DataGrid1.Bookmark = DataGrid1.Bookmark - 1
        BotonModificar
        PonFoco txtaux(3)
    ElseIf DataGrid1.Bookmark = 1 Then
        BotonModificar
        PonFoco txtaux(3)
    End If
End Sub



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


