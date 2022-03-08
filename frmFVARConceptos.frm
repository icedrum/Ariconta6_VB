VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFVARConceptos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos de Facturas"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15150
   Icon            =   "frmFVARConceptos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   8
      Left            =   10305
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3870
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   5
      Left            =   7965
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3870
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   4
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3870
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   3
      Left            =   4905
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3870
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   17
      Top             =   90
      Width           =   3585
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   18
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   19
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3900
      TabIndex        =   15
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   180
         TabIndex        =   16
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
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Index           =   2
      Left            =   10020
      MaskColor       =   &H00000000&
      TabIndex        =   14
      ToolTipText     =   "Buscar Centro Coste"
      Top             =   3885
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
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
      Height          =   290
      Index           =   8
      Left            =   9420
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "Centro Coste|T|S|||fvarconceptos|codccost|||"
      Text            =   "Ccos"
      Top             =   3885
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Index           =   1
      Left            =   6615
      MaskColor       =   &H00000000&
      TabIndex        =   13
      ToolTipText     =   "Buscar Tipo Iva"
      Top             =   3915
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
      Height          =   290
      Index           =   3
      Left            =   3780
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Cuenta Contable|T|N|||fvarconceptos|codmacta|||"
      Text            =   "CtaContabl"
      Top             =   3915
      Width           =   870
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Index           =   0
      Left            =   4635
      MaskColor       =   &H00000000&
      TabIndex        =   12
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3915
      Visible         =   0   'False
      Width           =   195
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
      Height          =   290
      Index           =   0
      Left            =   135
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Codigo|N|N|1|999|fvarconceptos|codconce|000|S|"
      Text            =   "Cod"
      Top             =   3930
      Width           =   555
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
      Height          =   290
      Index           =   1
      Left            =   810
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||fvarconceptos|nomconce|||"
      Text            =   "Nombre"
      Top             =   3930
      Width           =   2295
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
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
      Height          =   290
      Index           =   4
      Left            =   6075
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "Tipo Iva|N|N|0|99|fvarconceptos|tipoiva|00||"
      Text            =   "Iv"
      Top             =   3915
      Width           =   555
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
      Left            =   12615
      TabIndex        =   5
      Top             =   8670
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   13770
      TabIndex        =   6
      Top             =   8670
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFVARConceptos.frx":000C
      Height          =   7425
      Left            =   135
      TabIndex        =   10
      Top             =   900
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   13097
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
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
      Left            =   13770
      TabIndex        =   11
      Top             =   8640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   135
      TabIndex        =   8
      Top             =   8505
      Width           =   2475
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
         Left            =   135
         TabIndex        =   9
         Top             =   180
         Width           =   2205
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4410
      Top             =   45
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   3150
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "C"
      Top             =   3915
      Width           =   555
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
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFVARConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO (Se lo copia)                          +-+-
' +-+- Menú: Bancos Propios (con un par)                    +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

Private Const IdPrograma = 420

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean
Public CodigoActual As String


' *** adrede: per a quan busque suplements/desconters des de frmViagrc ***
Public ExpedBusca As Long
Public TipoSuplem As Integer
' *********************************************************************

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid

Private CadenaConsulta As String
Private CadB As String

' ### [Monica] 08/09/2006
Private WithEvents frmCtas As frmColCtas 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTIva As frmBasico2  'tipos de via de contabilidad
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico 'centros de coste de contabilidad
Attribute frmCC.VB_VarHelpID = -1


Private SQL As String



Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos

Private kCampo As Integer

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim i As Integer
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    ' **** posar tots els controls (botons inclosos) que siguen del Grid
    For i = 0 To 4
        txtAux(i).visible = Not B
    Next i
    txtAux(8).visible = Not B ' centro de coste
    
    txtAux2(3).visible = Not B
    txtAux2(4).visible = Not B
    txtAux2(5).visible = Not B
    txtAux2(8).visible = Not B
    
    
    
    
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).visible = (Modo = 3 Or Modo = 4)
        btnBuscar(i).Enabled = (Modo = 3 Or Modo = 4)
    Next i
    ' **************************************************
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
'    For i = 5 To 7
'        BloquearTxt txtAux(i), True
'    Next i
'    BloquearTxt txtAux(9), True ' nombre de centro de coste


    
    ' ********************************************************

    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es retornar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari
    
    ' *** bloquejar tota la PK quan estem en modificar  ***
    BloqueaTXT txtAux(0), (Modo = 4)   'codconce
    
'    BloquearImgBuscar Me, Modo

    PonerModoUsuarioGnral Modo, "ariconta"

End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim B As Boolean
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim B As Boolean

    ' *** adrede: per a que no es puga fer res si estic cridant des de frmViagrc ***

    B = (Modo = 2) And ExpedBusca = 0
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta And ExpedBusca = 0
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B

    'Eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
    Me.mnImprimir.Enabled = B

    ' ******************************************************************************
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim i As Integer
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("fvarconceptos", "codconce")
        'NumF = SugerirCodigoSiguienteStr("sdexpgrp", "codsupdt", "codempre=" & vSesion.Empresa)
        'NumF = ""
    End If
    '***************************************************************
    'Situem el grid al final
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = NumF
'    For i = 1 To 9
'        txtaux(i).Text = ""
'    Next i
    txtAux(1).Text = ""
    txtAux(3).Text = ""
    txtAux(4).Text = ""
    txtAux(8).Text = ""
    
    txtAux2(3).Text = ""
    txtAux2(4).Text = ""
    txtAux2(5).Text = ""
    txtAux2(8).Text = ""


    ' **************************************************

    LLamaLineas anc, 3
       
    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonFoco txtAux(1) '**** 1r camp visible que NO siga PK ****
    Else
        PonFoco txtAux(0) '**** 1r camp visible que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    Dim i As Integer
    
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    CargaGrid "codconce = -1"
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(3).Text = ""
    txtAux(4).Text = ""
    txtAux(8).Text = ""
    
    txtAux2(3).Text = ""
    txtAux2(4).Text = ""
    txtAux2(5).Text = ""
    txtAux2(8).Text = ""

    LLamaLineas DataGrid1.top + 240, 1
    
    ' *** posar el foco al 1r camp visible que siga PK ***
    PonFoco txtAux(0)
    ' ***************************************************************
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.top 'DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(3).Text = DataGrid1.Columns(2).Text
    txtAux2(3).Text = DataGrid1.Columns(3).Text
    txtAux(4).Text = ComprobarCero(Trim(DataGrid1.Columns(4).Text))
    txtAux2(4).Text = DataGrid1.Columns(5).Text
    txtAux2(5).Text = DataGrid1.Columns(6).Text
    txtAux(8).Text = DataGrid1.Columns(7).Text
    txtAux2(8).Text = DataGrid1.Columns(8).Text
    
    ' ********************************************************

    LLamaLineas anc, 4 'modo 4
   
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonFoco txtAux(1)
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    txtAux(0).top = alto
    txtAux(1).top = alto
    txtAux(3).top = alto
    txtAux(4).top = alto
    txtAux(8).top = alto
    
    txtAux2(3).top = alto
    txtAux2(4).top = alto
    txtAux2(5).top = alto
    txtAux2(8).top = alto

    If Not vParam.autocoste Then
        btnBuscar(2).visible = False
        btnBuscar(2).Enabled = False
        txtAux(8).visible = False
        txtAux(8).Enabled = False
        txtAux2(8).visible = False
        txtAux2(8).Enabled = False
    End If

    btnBuscar(0).top = alto - 15
    btnBuscar(1).top = alto - 15
    btnBuscar(2).top = alto - 15
    ' ***************************************************
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Certes comprovacions
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ***************************************************************************
    
    '*** canviar la pregunta, els noms dels camps i el DELETE; repasar codEmpre ***
    SQL = "¿Seguro que desea eliminar el Concepto?"
    'SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBoxA(SQL, vbQuestion + vbYesNo) = vbYes Then
        'N'hi ha que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from fvarconceptos where codconce = " & adodc1.Recordset!CodConce
        
        Conn.Execute SQL
        CargaGrid CadB
        
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
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
    TerminaBloquear
     Select Case Index
        Case 0 'Cuentas Contables (de contabilidad)
        
            '¿?  !!!!!!    Por que moni? !!!!!!!
            'If txtAux(3).Text = "" Then Exit Sub
            'If CInt(txtAux(3).Text) = 0 Then Exit Sub
            
            Indice = 3
            Set frmCtas = New frmColCtas
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.ConfigurarBalances = 3  'NUEVO
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonFoco txtAux(Indice)
            
        Case 1 'Tipos de Iva (de contabilidad)
            
            '¿?  !!!!!!    Por que moni? !!!!!!!
            'If txtAux(2).Text = "" Then Exit Sub
            'If CInt(txtAux(2).Text) = 0 Then Exit Sub
            
            Indice = 4
            
            Set frmTIva = New frmBasico2
            AyudaTiposIva frmTIva
            Set frmTIva = Nothing
            
            PonFoco txtAux(Indice)
        
        Case 2 'Centros de coste de contabilidad
            Indice = 8
            If txtAux(Indice).Enabled Then
                Set frmCC = New frmBasico
                AyudaCC frmCC
                Set frmCC = Nothing
            End If
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1

End Sub

Private Sub cmdAceptar_Click()
Dim i As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOK Then
                'If InsertarDesdeForm(Me) Then
                If InsertarDesdeForm2(Me, 0) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            'adodc1.Recordset.Filter = "codempre = " & txtAux(0).Text & " AND codsupdt = " & txtAux(1).Text
                            adodc1.Recordset.Filter = "codconce = " & txtAux(0).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    i = adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Move i - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid CadB
    End Select
    
    If adodc1.Recordset.EOF Then
        LimpiarCampos
    End If
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    ' *** adrede: per a tornar el TipoSuplem ***
    ' cad = cad & TipoSuplem & "|"
    ' ******************************************
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonFoco txtAux(1)
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer

    '******* repasar si n'hi ha botó d'imprimir o no******
    '*****************************************************
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

    ' desplazamiento
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With

    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    DespalzamientoVisible False
    
    ' *** SI N'HI HAN COMBOS ***
    ' CargaCombo 0
    ' **************************
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT fvarconceptos.codconce, fvarconceptos.nomconce, fvarconceptos.codmacta, cuentas.nommacta, "
    CadenaConsulta = CadenaConsulta & " fvarconceptos.tipoiva, tiposiva.nombriva, tiposiva.porceiva, fvarconceptos.codccost, "
    CadenaConsulta = CadenaConsulta & " ccoste.nomccost "
    CadenaConsulta = CadenaConsulta & " FROM ((fvarconceptos inner join cuentas ON  fvarconceptos.codmacta = cuentas.codmacta) "
    CadenaConsulta = CadenaConsulta & " INNER JOIN tiposiva ON fvarconceptos.tipoiva = tiposiva.codigiva) "
    CadenaConsulta = CadenaConsulta & " LEFT JOIN ccoste on fvarconceptos.codccost = ccoste.codccost"
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    ' ****** Si n'hi han camps fora del grid ******
    ' *********************************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
    End If
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub



Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
'Centro de coste de la Contabilidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codccost
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcoste
End Sub

' ### [Monica] 08/09/2006
Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtAux2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de Iva de la Contabilidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'nombriva
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 3) 'Porcentaje iva
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
Dim frmFVARInf As frmFVARInformes

    Set frmFVARInf = New frmFVARInformes

    frmFVARInf.OpcionListado = 0
    frmFVARInf.Show vbModal
    
    Set frmFVARInf = Nothing
    
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    
    'Prepara para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
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
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 8 'Imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSql As String)
    Dim i As Integer
    Dim SQL As String
    Dim tots As String
    
    If vSql <> "" Then
        SQL = CadenaConsulta & " WHERE " & vSql  ' ### [Monica] 08/09/2006: antes habia AND
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    'SQL = SQL & " ORDER BY codempre, codsupdt"
    SQL = SQL & " ORDER BY codconce"
    '**************************************************************++
    
'    adodc1.RecordSource = SQL
'    adodc1.CursorType = adOpenDynamic
'    adodc1.LockType = adLockOptimistic
'    DataGrid1.ScrollBars = dbgNone
'    adodc1.Refresh
'    Set DataGrid1.DataSource = adodc1 ' per a que no ixca l'error de "la fila actual no está disponible"
       
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, False
       
       
    ' *** posar només els controls del grid ***
    tots = "S|txtAux(0)|T|Código|850|;S|txtAux(1)|T|Denominación|3700|;"
    tots = tots & "S|txtAux(3)|T|Cuenta|1500|;S|btnBuscar(0)|B|||;S|txtAux2(3)|T|Nombre Cuenta|3450|;S|txtAux(4)|T|Iva|600|;"
    tots = tots & "S|btnBuscar(1)|B|||;S|txtAux2(4)|T|Tipo Iva|1400|;S|txtAux2(5)|T|%Iva|800|;S|txtAux(8)|T|C.Coste|830|;S|btnBuscar(2)|B|||;S|txtAux2(8)|T|Nombre|1000|;"
'    For i = 1 To 11
'        tots = tots & "N||||0|;"
'    Next i
    arregla tots, DataGrid1, Me
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************
    
    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    DataGrid1.Columns(2).Alignment = dbgCenter
    DataGrid1.Columns(4).Alignment = dbgCenter
    DataGrid1.Columns(5).Alignment = dbgCenter
    DataGrid1.Columns(6).Alignment = dbgCenter
    ' *****************************
    
    
    ' *** Si n'hi han camps fora del grid ***
    If adodc1.Recordset.EOF Then
        LimpiarCampos
    End If
    ' **************************************
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Index = 3 And KeyAscii = 43 Then '+
'        KeyAscii = 0
'    Else
'        KEYpress KeyAscii
'    End If
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 12: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
    
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim RC As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case Index
        Case 0
            PonerFormatoEntero txtAux(Index)
        
        Case 1
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        
        Case 3 'cuenta contable
            RC = txtAux(3).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtAux(3).Text = RC
                If Modo = 1 Then Exit Sub
                
                txtAux2(3).Text = SQL
                
                RC = ""
            Else
                If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                    'NO EXISTE LA CUENTA, añado que debe de tener permiso de creacion de cuentas
                    If vUsu.PermiteOpcion("ariconta", 201, vbOpcionCrearEliminar) Then
                        SQL = SQL & " ¿Desea crearla?"
                        If MsgBoxA(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                            CadenaDesdeOtroForm = RC
                            Set frmCtas = New frmColCtas
                            frmCtas.DatosADevolverBusqueda = "0|1|"
                            frmCtas.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                            frmCtas.Show vbModal
                            Set frmCtas = Nothing
                            If txtAux(3).Text = RC Then
                                SQL = "" 'Para k no los borre
                                ' traemos el tipo de iva de la cuenta
                                txtAux(4).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtAux(3).Text, "N")
                                txtAux_LostFocus (4)
                            End If
                        End If
                    Else
                        MsgBoxA SQL, vbExclamation
                    End If
                Else
                    'MsgBoxA Sql, vbExclamation
                End If
                    
                If SQL <> "" Then
                  txtAux(3).Text = ""
                  txtAux2(3).Text = ""
                  RC = "NO"
                End If
            End If
            HabilitarCentroCoste
            If RC <> "" Then
                If txtAux(3).Text <> "" Then PonFoco txtAux(3)
            End If
            If Modo = 3 Or Modo = 4 Then MostrarObservaciones txtAux(Index)
            
       Case 4 'tipo de iva
            If txtAux(Index).Text = "" Then Exit Sub
       
            RC = "porceiva"
            txtAux2(4).Text = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", txtAux(4), "N", RC)
            If txtAux2(4).Text = "" Then
                MsgBoxA "No existe el Tipo de Iva. Reintroduzca.", vbExclamation
                If txtAux(4).Text <> "" Then PonFoco txtAux(4)
            Else
                If RC = 0 Then
                    txtAux2(5).Text = ""
                Else
                    txtAux2(5).Text = RC
                End If
            End If
       
       Case 8 ' centro de coste
            RC = ""
            If txtAux(Index).Text <> "" Then
                txtAux(Index).Text = UCase(txtAux(Index).Text)
                RC = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtAux(Index).Text, "T")
            End If
            txtAux2(Index).Text = RC
            If RC = "" And txtAux(Index).Text <> "" Then
                MsgBox "No existe centro de coste : " & txtAux(Index).Text, vbExclamation
                txtAux(Index).Text = ""
                PonFoco txtAux(Index)
            End If
       
    End Select
    '**************************************************************************
End Sub

Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtAux(3).Text <> "" Then
            hab = HayKHabilitarCentroCoste(txtAux(3).Text)
        Else
            txtAux(8).Text = ""
        End If
        If hab Then
            txtAux(8).BackColor = &H80000005
            Else
            txtAux(8).BackColor = &H80000018
            txtAux(8).Text = ""
        End If
    End If
    txtAux(8).Enabled = hab
End Sub

Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
' *** només per ad este manteniment ***
Dim Rs As Recordset
Dim Cad As String
'Dim exped As String
' *************************************

    B = CompForm(Me)
    If Not B Then Exit Function


    If B And (Modo = 3) Then
        'Estem insertant
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numèric
         
        ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
        Datos = DevuelveDesdeBD("codconce", "fvarconceptos", "codconce", txtAux(0).Text, "N")
'       Datos = DevuelveDesdeBDNew(1, "sdexpgrp", "codsupdt", "codsupdt", txtAux(1).Text, "N", "", "codempre", CStr(vSesion.Empresa), "N")
         
        If Datos <> "" Then
            MsgBox "Ya existe el Código de Concepto: " & txtAux(0).Text, vbExclamation
            B = False
            PonFoco txtAux(1) '*** posar el foco al 1r camp visible de la PK de la capçalera ***
            Exit Function
        End If
        '*************************************************************************************
    End If

    ' *** Si cal fer atres comprovacions ***
    'comprobamos que la cta contable sea de gastos
    If B And vParam.autocoste Then
        If txtAux(8).Text = "" Then
            MsgBox "La Contabilidad tiene Analítica debe introducir el Centro de Coste.", vbExclamation
            B = False
            PonFoco txtAux(8)
        Else
            
        End If
    End If
    ' *********************************************

    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
'    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 And Modo = 2 Then Unload Me  'ESC
    End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Modo <> 4 Then 'Modificar
    Else
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    
    PonerContRegIndicador
    
End Sub


Private Sub LimpiarCampos()
Dim i As Integer
On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    For i = 5 To 7
        txtAux(i).Text = ""
    Next i
    ' ****************************************************
'    text2(12).Text = "" ' el nombre de la cuenta contable la ponemos a cero

    If Err.Number <> 0 Then Err.Clear
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



