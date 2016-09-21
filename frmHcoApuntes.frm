VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHcoApuntes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico de apuntes"
   ClientHeight    =   8340
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12015
   Icon            =   "frmHcoApuntes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameResultdos 
      Height          =   7635
      Left            =   0
      TabIndex        =   39
      Top             =   450
      Width           =   12015
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   4440
         Top             =   7800
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Adodc2"
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
      Begin MSDataGridLib.DataGrid DataGridResult 
         Bindings        =   "frmHcoApuntes.frx":000C
         Height          =   6735
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   11880
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
      Begin VB.Label Label4 
         Caption         =   "Cargando datos desde BD"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   7320
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Resultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.TextBox txtAmp 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   7600
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSComctlLib.Toolbar toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Resultado búsqueda anterior"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver asiento"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtaux 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtaux 
      Height          =   885
      Index           =   12
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmHcoApuntes.frx":0021
      Top             =   1320
      Width           =   5775
   End
   Begin VB.TextBox txtaux 
      Height          =   315
      Index           =   10
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "commor"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text4"
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   25
      Top             =   6240
      Width           =   195
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   1
      Left            =   1080
      TabIndex        =   24
      Top             =   6240
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   5
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   6
      Top             =   6240
      Width           =   885
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   5400
      TabIndex        =   7
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   5
      Left            =   6480
      MaxLength       =   30
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   6
      Left            =   8340
      TabIndex        =   9
      Top             =   6240
      Width           =   1125
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   7
      Left            =   9480
      TabIndex        =   10
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   8
      Left            =   10620
      MaxLength       =   4
      TabIndex        =   11
      Top             =   6240
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   8400
      TabIndex        =   17
      Top             =   1200
      Width           =   3375
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "HABER"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   22
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "DEBE"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   300
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   2040
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   14
      Top             =   7650
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Height          =   315
      Index           =   11
      Left            =   3120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmHcoApuntes.frx":0028
      Height          =   4395
      Left            =   120
      TabIndex        =   16
      Top             =   2340
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7752
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
   Begin VB.Frame frameextras 
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   6720
      Width           =   10215
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nomctapar"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   420
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nombreconcepto"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   4
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "centrocoste"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   3
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Cta. Contrapartida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   32
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4320
         TabIndex        =   31
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "C. coste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7800
         TabIndex        =   30
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10440
      TabIndex        =   12
      Top             =   7650
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hco.  Apuntes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BCBCBC&
      Height          =   585
      Left            =   6600
      TabIndex        =   43
      Top             =   7500
      Width           =   3285
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   1
      Left            =   3960
      Picture         =   "frmHcoApuntes.frx":003D
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   0
      Left            =   2040
      Picture         =   "frmHcoApuntes.frx":0A3F
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Asiento"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   36
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   35
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   195
      Index           =   5
      Left            =   1500
      TabIndex        =   34
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Diario"
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   15
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Cod Diario"
      Height          =   195
      Index           =   0
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.Menu mnOpcionesAsiPre 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Busqueda anterior"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnVerAsiento 
         Caption         =   "Ver &Asiento"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
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
         Shortcut        =   ^P
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFIltro 
      Caption         =   "Filtro"
      Begin VB.Menu mnNinguno 
         Caption         =   "Ninguno"
      End
      Begin VB.Menu mnFechas 
         Caption         =   "-"
      End
      Begin VB.Menu mnAyer 
         Caption         =   "Ayer"
         HelpContextID   =   1
      End
      Begin VB.Menu mnActual 
         Caption         =   "Mes actual"
         HelpContextID   =   2
      End
      Begin VB.Menu mnHace30Dias 
         Caption         =   "Hace 30 dias"
         HelpContextID   =   3
      End
      Begin VB.Menu mnNum 
         Caption         =   "-"
      End
      Begin VB.Menu mnUltimo 
         Caption         =   "Ultimo"
         HelpContextID   =   4
      End
      Begin VB.Menu mnDiezultimos 
         Caption         =   "10 ultimos"
         HelpContextID   =   5
      End
   End
End
Attribute VB_Name = "frmHcoApuntes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 302

Public ASIENTO As String   'Vendra separado por pies los tres campos del asiento
Public LINASI As Integer 'Si es k keremos ver una  linea de asiento

Public EjerciciosCerrados As Boolean
Dim Tablas As String

Private Const NO = "No encontrado"
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
'Private WithEvents frmCon As frmConceptos
'Private WithEvents frmCC As frmCCoste
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDi As frmTiposDiario
Attribute frmDi.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos   . 2 MODOS
' 1- Buscar
' 2- Con resultado, mostramos el grid resultados
' 3- El asiento, directamente


'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
'Private kCampo As Integer
'-------------------------------------------------------------------------
'Private HaDevueltoDatos As Boolean
Private SQL As String
Dim i As Integer
Dim ancho As Integer
'Dim colMes As Integer

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

Dim lon As Integer
Private NOSoloVerAsiento As Boolean
Dim Rs As Recordset

Dim PrimeraVez As Boolean
Dim FILTRO As Byte


Private Sub Adodc2_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    PonerLabelRecordset
End Sub

Private Sub PonerLabelRecordset()
On Error Resume Next
    Label4.Caption = Adodc2.Recordset.AbsolutePosition & " de " & Adodc2.Recordset.RecordCount
    If Err.Number <> 0 Then
        Err.Clear
        Label4.Caption = ""
    End If
End Sub



Private Sub cmdAceptar_Click()

       Screen.MousePointer = vbHourglass
       ' HacerBusqueda
       If GeneraCadenaConuslta Then
            'Ya tenemos la subcadena de busqueda. Ahora vamos a poner
            'Vamos a generar la cadena consulta
            If EjerciciosCerrados Then
                Tablas = "1"
            Else
                Tablas = ""
            End If
            
            CadenaConsulta = "SELECT hcabapu" & Tablas & ".fechaent, hcabapu" & Tablas & ".numasien,hcabapu" & Tablas & ".numdiari, hlinapu" & Tablas & ".codmacta, cuentas.nommacta,"
            CadenaConsulta = CadenaConsulta & " hlinapu" & Tablas & ".numdocum, hlinapu" & Tablas & ".ampconce, hlinapu" & Tablas & ".timporteD, hlinapu" & Tablas & ".timporteH,hlinapu" & Tablas & ".linliapu"
            'CCOSTE
            If vParam.autocoste Then
                CadenaConsulta = CadenaConsulta & ",hlinapu" & Tablas & ".codccost"
            End If
            CadenaConsulta = CadenaConsulta & " FROM (hcabapu" & Tablas & " INNER JOIN hlinapu" & Tablas & " ON (hcabapu" & Tablas & ".numasien = hlinapu" & Tablas & ".numasien) AND (hcabapu" & Tablas & ".fechaent = hlinapu" & Tablas & ".fechaent)"
            CadenaConsulta = CadenaConsulta & " AND (hcabapu" & Tablas & ".numdiari = hlinapu" & Tablas & ".numdiari)) INNER JOIN cuentas ON hlinapu" & Tablas & ".codmacta = cuentas.codmacta"
            CadenaConsulta = CadenaConsulta & " WHERE " & SQL & " ORDER BY hcabapu" & Tablas & ".fechaent,hcabapu" & Tablas & ".numasien,hlinapu" & Tablas & ".linliapu;"
            PonerCadenaBusqueda True
       End If
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click(Index As Integer)
    cmdAux(0).Tag = 0
    LlamaContraPar
End Sub


'Private Sub cmdRegresar_Click()
'Dim Cad As String
'Dim i As Integer
'Dim j As Integer
'Dim AUX As String
'
'If Data1.Recordset.EOF Then
'    MsgBox "Ningún registro devuelto.", vbExclamation
'    Exit Sub
'End If
'
'Cad = ""
'i = 0
'Do
'    j = i + 1
'    i = InStr(j, DatosADevolverBusqueda, "|")
'    If i > 0 Then
'        AUX = Mid(DatosADevolverBusqueda, j, i - j)
'        j = Val(AUX)
'        Cad = Cad & Text1(j).Text & "|"
'    End If
'Loop Until i = 0
'RaiseEvent DatoSeleccionado(Cad)
'Unload Me
'End Sub



Private Sub DataGridResult_DblClick()
If Not Adodc2.Recordset.EOF Then
    Screen.MousePointer = vbHourglass
    Label3.Caption = "Leyendo BD"
    Label4.Caption = "Leyendo BD"
    Me.Refresh
    DoEvents
    Screen.MousePointer = vbHourglass
    ASIENTO = Adodc2.Recordset!NumDiari & "|" & Adodc2.Recordset!FechaEnt & "|" & Adodc2.Recordset!NumAsien & "|"
    LINASI = Adodc2.Recordset!Linliapu
    PonerAsiento
    PonerModo 3
    Screen.MousePointer = vbDefault

End If
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()
Dim Fec As Date
    If PrimeraVez Then
        PrimeraVez = False
        If EjerciciosCerrados Then
            Caption = Caption & "     -   EJERCICIOS CERRADOS"
            FILTRO = 0
        Else
'            FijarFiltro True
        End If
'        PonerOpcionMenuFiltro
        
        
        If ASIENTO = "" Then
            'En espera de que ponga buscar
            PonerModo 1
            
            'ADemas, segun sea el filtro puedo poner
            Select Case FILTRO
            Case 1
                'Ayer
                i = 1
            Case 2
                'Este mes
                i = Day(Now) - 1
            Case 3
                'Hace un mes
                i = 30
            Case 4, 5
                i = 0
                'Al leer el filtro ya pone el asiento
            Case Else
                i = -1
            End Select
            If i > 0 Then
                'Fecha
                Fec = DateAdd("d", -i, Now)
                If Fec < vParam.fechaini Then Fec = vParam.fechaini
                txtaux(10).Text = ">=" & Format(Fec, "dd/mm/yyyy")
            End If
            If i >= 0 Then
                'Ponemos el foco en el boton de aceptar
                Me.cmdAceptar.SetFocus
            End If
        Else
            'Realmente quiero ver este asiento: ASIENTO
            PonerAsiento
            'y el modo
            PonerModo 3
        End If
    End If
    Caption = "Histórico de apuntes."
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    LimpiarCampos
    NOSoloVerAsiento = (ASIENTO = "")
    'Ponemos los iconos
   ' ESTE; TROZO; SERVIRA; PARA; MUCHOS, casi; todos; los; Forms
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 19
        .Buttons(4).Image = 18
        .Buttons(5).Image = 4
        .Buttons(6).Image = 5
        .Buttons(7).Image = 16
        .Buttons(9).Image = 15
    End With
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
        Me.Width = 12000
        'Me.Height = Screen.Height.
    End If
    Me.Height = 8800
    
    CadAncho = False
    'ASignamos un SQL al DATA1
'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login
'    Adodc2.password = vUsu.Passwd
'    Adodc2.UserName = vUsu.Login
   
    'Maxima longitud cuentas
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    txtaux(3).MaxLength = vEmpresa.DigitosUltimoNivel
    CargaGrid False
    
    PrimeraVez = True
End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
End Sub

'Private Sub Form_Resize()
'If Me.WindowState <> 0 Then Exit Sub
'If Me.Width < 11610 Then Me.Width = 11610
'End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Me.DataGrid1.Enabled = False
    espera 0.2
    'Set Adodc2.Recordset = Nothing
    'FijarFiltro False
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas
If cmdAux(0).Tag = 0 Then
    'Cuenta normal
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
Else
    'contrapartida
    txtaux(3).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 2)
End If
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtaux(10).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgppal_Click(Index As Integer)
    If Modo <> 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        'If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
    Case 1
        'Tipos diario
        Set frmDi = New frmTiposDiario
        frmDi.DatosADevolverBusqueda = "0"
        frmDi.Show vbModal
        Set frmDi = Nothing
    End Select
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnActual_Click()
      ClickMenu 2
End Sub

Private Sub mnAyer_Click()
  ClickMenu 1
End Sub

Private Sub mnBuscar_Click()
    If Modo = 1 Then
        cmdAceptar_Click
    Else
        LimpiarCampos
        CargaGrid False
        PonerModo 1
    End If
End Sub

Private Sub mnDiezultimos_Click()
    ClickMenu 5
End Sub

Private Sub mnEliminar_Click()
    HacerToolBar 6
End Sub

Private Sub mnHace30Dias_Click()
    ClickMenu 3
End Sub

Private Sub mnImprimir_Click()
    HacerToolBar 7
End Sub

Private Sub mnModificar_Click()
    HacerToolBar 5
End Sub

Private Sub mnNinguno_Click()
    ClickMenu 0
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    DataGrid1.Enabled = False
    Unload Me
End Sub



Private Sub mnUltimo_Click()
    ClickMenu 4
End Sub

Private Sub mnVerAsiento_Click()
    HacerToolBar 4
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    HacerToolBar Button.Index
    
End Sub

Private Sub HacerToolBar(IndiceTool As Integer)
Dim ANtiguoActualiza As Boolean
Dim Mc As Contadores
Dim CadFac As String
Dim SoloModificar As Boolean

    Select Case IndiceTool
    Case 1
          mnBuscar_Click
    Case 2
          PonerModo 2
          
    Case 5, 6
        
        CadFac = ""
        
        If Adodc1.Recordset.EOF Then Exit Sub
        'Primero comprobamos si esta cerrado el ejercicio
        varFecOk = FechaCorrecta2(Adodc1.Recordset!FechaEnt)
        If varFecOk >= 2 Then
            If varFecOk = 2 Then
                MsgBox varTxtFec, vbExclamation
            Else
                MsgBox "El asiento pertenece a un ejercicio cerrado.", vbExclamation
            End If
            Exit Sub
        End If
        
        'Cojo prestado esta variabel un momento CadenaDesdeOtroForm
        If Not IsNull(Adodc1.Recordset!idcontab) Then
            If Adodc1.Recordset!idcontab = "FRACLI" Then
                CadFac = "FRACLI"
                CadenaDesdeOtroForm = " clientes "
            Else
                If Adodc1.Recordset!idcontab = "FRAPRO" Then
                    CadFac = "FRAPRO"
                    CadenaDesdeOtroForm = " proveedores "
                End If
            End If
        End If
        If CadFac <> "" Then
            'Esta intentando modificar o eliminar apuntes k vienen de facturas
            'Luego si en el parametro le dice k no puede, no le dejamos segir
            If vParam.modhcofa Then
                  i = MsgBox("Va a modificar el apunte de una factura de " & CadenaDesdeOtroForm & vbCrLf & _
                    " DEBERIA hacerlo desde registro de facturas de " & CadenaDesdeOtroForm & "." & vbCrLf & vbCrLf & _
                        "¿Desea continuar?", vbQuestion + vbYesNoCancel)
            Else
                MsgBox "Este apunte pertenece a una factura de " & CadenaDesdeOtroForm & " y solo se puede modificar en el registro" & _
                    " de facturas de " & CadenaDesdeOtroForm & ".", vbExclamation
                i = -1
            End If
            CadenaDesdeOtroForm = ""
            If i <> vbYes Then Exit Sub
        Else
            'No es de FACTURAS es asiento. Si es eliminar pregunto
            If IndiceTool = 6 Then
                SQL = String(40, "*") & vbCrLf
                SQL = SQL & SQL & SQL
                SQL = SQL & "Va a eliminar el asiento: " & vbCrLf & vbCrLf
                SQL = SQL & "Nº asiento: " & Adodc1.Recordset!NumAsien & vbCrLf
                SQL = SQL & "Fecha     : " & Adodc1.Recordset!FechaEnt & String(4, vbCrLf)
                
                'CUANTAS LINEAS TIENE
                PonerNumeroDeLineasParaEliminar
                
                
                SQL = SQL & "¿ Desea continuar de cualquier modo?" & vbCrLf & vbCrLf & vbCrLf
                'SQL = SQL & String(40, "*")
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        i = 0
        SoloModificar = False
        If IndiceTool = 5 Then
            'Si hay que modificar el asiento entonces, si es de factura o cliente entonces NO isnertamos en
            ' diario
            If CadFac = "" Then
                i = 3  'Desactualizar para modificar
                SoloModificar = True
            Else
                'Si pongo 2 lo borro y no lo dejo modificar
                'I = 2
                'Si pongo un 3 lo modifico
                i = 3
                SoloModificar = True
                CadFac = ""
            End If
        Else
            i = 2  'Desactualizar para eliminar
        End If
        
        frmActualizar.OpcionActualizar = i
        frmActualizar.NUmSerie = SQL
        frmActualizar.NumAsiento = CLng(txtaux(9).Text)
        frmActualizar.FechaAsiento = CDate(txtaux(10).Text)
        frmActualizar.NumDiari = CInt(txtaux(11).Text)
        AlgunAsientoActualizado = False
        frmActualizar.Show vbModal
        Screen.MousePointer = vbHourglass
        Me.Refresh
        If AlgunAsientoActualizado Then
            Conn.Execute "commit"
            
            'El LOG
            If IndiceTool = 5 Then
                i = 2
            Else
                i = 3
            End If
            vLog.Insertar CByte(i), vUsu, txtaux(9).Text & " : " & txtaux(10).Text
            
            
            'Devolvemos contador
            If Not SoloModificar Then
                i = FechaCorrecta2(CDate(txtaux(10).Text))
                Set Mc = New Contadores
                NumRegElim = CLng(txtaux(9).Text)
                Mc.DevolverContador "0", i = 0, NumRegElim
                Set Mc = Nothing
            End If
        
            If CadFac = "" Then
                If IndiceTool = 5 Then
                    ANtiguoActualiza = vParam.AsienActAuto
                    vParam.AsienActAuto = True
                    frmAsientos.ASIENTO = txtaux(11).Text & "|" & txtaux(10).Text & "|" & txtaux(9).Text & "|"
                    espera 1
                    AsientoConExtModificado = 1   'Nos garantizamos k vuelva a cargar los datos
                    'frmAsientos.vLinapu = 1
                    frmAsientos.vLinapu = LINASI
                    frmAsientos.Show vbModal
                    vParam.AsienActAuto = ANtiguoActualiza
                    If Not NOSoloVerAsiento Then
                        'Si viene del la consulta de extractos
                        Unload Me
                        Exit Sub
                    End If
                Else
                    'BORRAR
                    AsientoConExtModificado = 1
                End If
            End If
            LimpiarCampos
            CargaGrid False
            Me.Refresh
            PonerCadenaBusqueda False
            Screen.MousePointer = vbDefault
        End If
    Case 4
        DataGridResult_DblClick
    Case 7
        'Imprimir
        Imprimir
    Case 9
        Me.DataGrid1.Enabled = False
        Unload Me
    End Select
End Sub


Private Sub PonerNumeroDeLineasParaEliminar()
    On Error Resume Next
    
    
    SQL = SQL & "      Total lineas apunte: "
    SQL = SQL & Right("            " & Adodc1.Recordset.RecordCount, 12) & vbCrLf & "      "
    SQL = SQL & String(19, "=")
    
    SQL = SQL & String(5, vbCrLf)
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Imprimir()
Dim Cad As String
'Resto parametros
    If Modo = 1 Then
        MsgBox "Esta buscando asientos. Ponga algun criterio de busqueda e imprimalos despues", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    SQL = ""
    
    'Fechas intervalor
    SQL = "Fechas= """"|"
    
    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    
    
    If Modo = 3 Then
        Cad = "hcabapu" & Tablas & ".numasien =" & txtaux(9).Text & " AND hcabapu" & Tablas & ".fechaent= '" & Format(txtaux(10).Text, FormatoFecha) & "'"
    Else
        Cad = Me.Tag
    End If
    'Cuentas
    SQL = SQL & "Cuenta= """"|"
    
    'Fecha impresion
    SQL = SQL & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
    If IHcoApuntes(Cad, Tablas) Then
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            'Opcion dependera del combo
            .opcion = 12
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda(MostrarMSG As Boolean)
Dim bol As Boolean
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    If CadenaConsulta = "" Then
        bol = True  'EOF
    Else
        Adodc2.ConnectionString = Conn
        Adodc2.RecordSource = CadenaConsulta
        Adodc2.Refresh
        bol = Adodc2.Recordset.EOF
    End If
    If bol Then
        If MostrarMSG Then MsgBox "No hay ningún registro con los valores solicitados", vbInformation
        Screen.MousePointer = vbDefault
        txtaux(9).SetFocus
        PonerModo 1
        Exit Sub
        Else
            CargaGridResultados
            Label4.Caption = Adodc2.Recordset.AbsolutePosition & " de " & Adodc2.Recordset.RecordCount
            PonerModo 2
            PonFocoGrid
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
        MuestraError Err.Number, "PonerCadenaBusqueda"
        PonerModo 1
        Screen.MousePointer = vbDefault
End Sub

Private Sub PonFocoGrid()

    On Error Resume Next
        Me.DataGridResult.SetFocus
        If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean

    
  'ASIGNAR MODO
   Modo = Kmodo
   
   frameResultdos.Visible = (Modo = 2)
   DataGrid1.Enabled = Modo = 3
   
   Me.Toolbar1.Buttons(1).Enabled = NOSoloVerAsiento
   Me.mnBuscar.Enabled = NOSoloVerAsiento
    Select Case Modo
    Case 1
        LLamaLineas Me.DataGrid1.Top + 220, 1, True
         
    Case 2
        
        Label4.Caption = ""
        Label3.Caption = "Resultados"
    Case 3
        For i = 0 To 8
            txtaux(i).Visible = False
        Next i
        cmdAux(0).Visible = False
    End Select
    
    For i = 9 To 12
        txtaux(i).Locked = Modo <> 1
    Next i
    'Modificar eliminar imprimi solo si estamos en asiento
    cmdAceptar.Visible = (Modo = 1) And NOSoloVerAsiento
    B = (Modo = 3)

    
    'Y ademas modificar y eliminar si no estamos en EJERCICIOS CERRADOS
    B = B And Not (EjerciciosCerrados)
        Toolbar1.Buttons(5).Enabled = B
        Toolbar1.Buttons(6).Enabled = B
        'Antes
        'Toolbar1.Buttons(7).Enabled = B
        Toolbar1.Buttons(7).Enabled = True
     
    B = (Modo = 3) And NOSoloVerAsiento
        Me.mnModificar.Enabled = B
        Me.mnEliminar.Enabled = B
        Me.mnImprimir.Enabled = B
        
        
    'Ver asiento solo en modo 2
    B = (Modo = 2) And NOSoloVerAsiento
        Toolbar1.Buttons(4).Enabled = B
        Me.mnVerAsiento.Enabled = B
    
    'Busqueda anterior si no estamos en el grid resultdos
    B = TieneDatosBusqueda And (Modo <> 2) And NOSoloVerAsiento
        Toolbar1.Buttons(2).Enabled = B
        Me.mnVerTodos.Enabled = B
        
    PonerOpcionesMenuGeneral Me
End Sub


Private Function TieneDatosBusqueda() As Boolean
        TieneDatosBusqueda = False
        If Not Adodc2.Recordset Is Nothing Then
            If Not Adodc2.Recordset.EOF Then
                TieneDatosBusqueda = True
            End If
        End If
End Function

Private Function DatosOK() As Boolean
    Dim Rs As ADODB.Recordset
    Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function
    DatosOK = B
End Function




Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = MontaSQLCarga(Enlaza)
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockPessimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    DataGrid1.Tag = "Asignando"
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-numaspre  1.- Lin aspre
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    'Claves lineas asientos predefinidos
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Visible = False

    'Cuenta
    DataGrid1.Columns(2).Caption = "Cuenta"
    DataGrid1.Columns(2).Width = 1005
    
    DataGrid1.Columns(3).Caption = "Denominación"
    DataGrid1.Columns(3).Width = 2395


    DataGrid1.Columns(4).Caption = "Docu."
    DataGrid1.Columns(4).Width = 1005

    DataGrid1.Columns(5).Caption = "Contra."
    DataGrid1.Columns(5).Width = 1005
    
    DataGrid1.Columns(6).Caption = "Cto."
    DataGrid1.Columns(6).Width = 465
    
    DataGrid1.Columns(7).Visible = False
    

        
    DataGrid1.Columns(8).Caption = "Ampliación"
    DataGrid1.Columns(8).Width = 2400

    'Cuenta contrapartida
    DataGrid1.Columns(9).Visible = False
    
    If vParam.autocoste Then
        ancho = 0
    Else
        ancho = 255 'Es la columna del centro de coste divida entre dos
    End If
    
    DataGrid1.Columns(10).Caption = "Debe"
    DataGrid1.Columns(10).NumberFormat = "#,##0.00"
    DataGrid1.Columns(10).Width = 1075 + ancho
    DataGrid1.Columns(10).Alignment = dbgRight
            
    DataGrid1.Columns(11).Caption = "Haber"
    DataGrid1.Columns(11).NumberFormat = "#,##0.00"
    DataGrid1.Columns(11).Width = 1075 + ancho
    DataGrid1.Columns(11).Alignment = dbgRight
            
            
    If vParam.autocoste Then
        DataGrid1.Columns(12).Caption = "C.C."
        DataGrid1.Columns(12).Width = 510
    Else
        DataGrid1.Columns(12).Visible = False
    End If
    
    DataGrid1.Columns(13).Visible = False
    DataGrid1.Columns(14).Visible = False
    DataGrid1.Columns(15).Visible = False
    DataGrid1.Columns(16).Visible = False
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        anc = 323
        txtaux(0).Left = DataGrid1.Left + 330
        txtaux(0).Width = DataGrid1.Columns(2).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(3).Left + 90
                
        txtaux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 6
        txtaux(1).Width = DataGrid1.Columns(3).Width - 180
    
        txtaux(2).Left = DataGrid1.Columns(4).Left + 150
        txtaux(2).Width = DataGrid1.Columns(4).Width - 30
    
        txtaux(3).Left = DataGrid1.Columns(5).Left + 150
        txtaux(3).Width = DataGrid1.Columns(5).Width - 45

        
        'Concepto
        txtaux(4).Left = DataGrid1.Columns(6).Left + 150
        txtaux(4).Width = DataGrid1.Columns(6).Width - 45
        
        txtaux(5).Left = DataGrid1.Columns(8).Left + 150
        txtaux(5).Width = DataGrid1.Columns(8).Width - 45
        
        txtaux(6).Left = DataGrid1.Columns(10).Left + 150
        txtaux(6).Width = DataGrid1.Columns(10).Width - 30
        
       
        txtaux(7).Left = DataGrid1.Columns(11).Left + 150
        txtaux(7).Width = DataGrid1.Columns(11).Width - 30
       
        txtaux(8).Left = DataGrid1.Columns(12).Left + 150
        txtaux(8).Width = DataGrid1.Columns(12).Width - 45
       
        CadAncho = True
    End If
        
    For i = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(i).AllowSizing = False
    Next i
    
    DataGrid1.Tag = "Calculando"
    'Obtenemos las sumas
    ObtenerSumas
    
    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub ObtenerSumas()
    Dim Deb As Currency
    Dim hab As Currency

    
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""

    
    If Adodc1.Recordset.EOF Then Exit Sub
    
    
    
    Set Rs = New ADODB.Recordset
    SQL = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
    SQL = SQL & " ,numdiari,fechaent,numasien"
    SQL = SQL & " From hlinapu"
    If EjerciciosCerrados Then SQL = SQL & "1"
    SQL = SQL & " GROUP BY numdiari, fechaent, numasien "
    SQL = SQL & " HAVING (((numdiari)=" & txtaux(11).Text
    SQL = SQL & ") AND ((fechaent)='" & Format(txtaux(10).Text, FormatoFecha)
    SQL = SQL & "') AND ((numasien)=" & txtaux(9).Text
    SQL = SQL & "));"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Deb = 0
    hab = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Deb = Rs.Fields(0)
        If Not IsNull(Rs.Fields(1)) Then hab = Rs.Fields(1)
    End If
    Rs.Close
    Set Rs = Nothing
    Text2(0).Text = Format(Deb, FormatoImporte): Text2(1).Text = Format(hab, FormatoImporte)
    'Metemos en DEB el total
    Deb = Deb - hab
    If Deb < 0 Then
        Text2(2).ForeColor = vbRed
        Else
        Text2(2).ForeColor = vbBlack
    End If
    If Deb = 0 Then
        Text2(2).Text = ""
    Else
        Text2(2).Text = Format(Deb, FormatoImporte)
    End If
End Sub


Private Function MontaSQLCarga(Enlaza As Boolean) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    ' Si ENLAZA -> Enlaza con el data1
    '           -> Si no lo cargamos sin enlazar a nngun campo
    '--------------------------------------------------------------------
    Dim SQL As String
    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    SQL = "SELECT numasien, linliapu, cuentas.codmacta, cuentas.nommacta,"
    SQL = SQL & " numdocum, ctacontr, conceptos.codconce, conceptos.nomconce as nombreconcepto, ampconce, cuentas_1.nommacta as nomctapar,"
    SQL = SQL & " timporteD, timporteH, cabccost.codccost, cabccost.nomccost as centrocoste,"
    SQL = SQL & " numdiari, fechaent,idcontab"
    SQL = SQL & " FROM (((hlinapu" & Tablas
    SQL = SQL & " LEFT JOIN cuentas AS cuentas_1 ON hlinapu" & Tablas & ".ctacontr ="
    SQL = SQL & " cuentas_1.codmacta) LEFT JOIN cabccost ON hlinapu" & Tablas & ".codccost = cabccost.codccost)"
    SQL = SQL & " INNER JOIN cuentas ON hlinapu" & Tablas & ".codmacta = cuentas.codmacta) INNER JOIN"
    SQL = SQL & " conceptos ON hlinapu" & Tablas & ".codconce = conceptos.codconce"
    If Enlaza Then
        SQL = SQL & " WHERE numasien = " & txtaux(9).Text
        SQL = SQL & " AND numdiari =" & txtaux(11).Text
        SQL = SQL & " AND fechaent= '" & Format(txtaux(10).Text, FormatoFecha) & "'"
        Else
        SQL = SQL & " WHERE numasien = -1"
    End If
    SQL = SQL & " ORDER BY hlinapu" & Tablas & ".linliapu"
    MontaSQLCarga = SQL
End Function




Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid
    B = (xModo = 0)
    frameextras.Visible = Not B
    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim i As Integer
    Dim J As Integer
    
    If vParam.autocoste Then
        J = 8
        Else
        J = 7
        txtaux(8).Visible = False
    End If
    For i = 0 To J
        txtaux(i).Visible = Visible
        txtaux(i).Top = Altura
    Next i
        
        cmdAux(0).Visible = Visible
        cmdAux(0).Top = Altura
    If Limpiar Then
        For i = 0 To J
            txtaux(i).Text = ""
        Next i
    End If
    
End Sub



Private Sub txtaux_Change(Index As Integer)
If Modo = 2 Then Exit Sub
With txtaux(Index)
    lon = MaximaLongitudCampo(Index)
    If Len(.Text) > lon Then
        txtAmp.Visible = True
        txtAmp.Text = .Text
        Else
        txtAmp.Visible = False
    End If
End With
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
lon = MaximaLongitudCampo(Index)
With txtaux(Index)
        .SelStart = Len(.Text)
    If Len(.Text) > lon Then
        Me.txtAmp.Text = .Text
        Me.txtAmp.Visible = True
        Else
            txtAmp.Text = ""
            txtAmp.Visible = False
    End If
End With

End Sub


'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:   KEYBusqueda1 KeyAscii, 0 ' cuenta
            Case 10:  KEYBusqueda KeyAscii, 0 ' fecha
            Case 11:  KEYBusqueda KeyAscii, 1 ' tipo de diario
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgppal_Click (indice)
End Sub


Private Sub KEYBusqueda1(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    cmdAux_Click (indice)
End Sub

'++





Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Sng As Double
    Dim B As Boolean
        
        txtAmp.Visible = False
        
        'Ponemos visible la primera parte del texto
        txtaux(Index).SelStart = 0
        
        
        
        'Comprobaremos ciertos valores
        txtaux(Index).Text = Trim(txtaux(Index).Text)
    
        'Comun a todos
        If txtaux(Index).Text = "" Then
            Select Case Index
            Case 0
                txtaux(1).Text = ""
            Case 3
                'Text3(0).Text = ""
            Case 11
                Text4.Text = ""
            End Select
            Exit Sub
        End If
        
        Select Case Index
        Case 0
            'Cta
            
            RC = txtaux(0).Text
            B = False
            If InStr(1, RC, ".") Then
                B = True
            Else
                If Len(B) = vEmpresa.DigitosUltimoNivel Then B = True
            End If
         
            If B Then
                If CuentaCorrectaUltimoNivel(RC, SQL) Then
                    txtaux(0).Text = RC
                    txtaux(1).Text = SQL
                    RC = ""
                Else
                    'MsgBox SQL, vbExclamation
                    'txtaux(0).Text = ""
                    'txtaux(1).Text = ""
                    'NO existe la cuenta N.N. Si no ha puesto * se lo reemplzao yo
                    If InStr(1, txtaux(0).Text, "*") = 0 Then
                        SQL = Replace(txtaux(0).Text, ".", "*")
                        txtaux(0).Text = SQL
                        RC = ""
                    End If
                End If
            Else
                'No pinto la cadena de texto
                txtaux(1).Text = ""
                RC = ""
            End If
            If RC <> "" Then txtaux(0).SetFocus
            
        Case 3
            RC = txtaux(3).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(3).Text = RC
                'Text3(0).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtaux(3).Text = ""
                'Text3(0).Text = ""
                txtaux(3).SetFocus
            End If
            
        Case 4
             If Not IsNumeric(txtaux(4).Text) Then
                    'MsgBox "El concepto debe de ser numérico", vbExclamation
                    Exit Sub
                End If
                RC = "tipoconce"
                SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtaux(4).Text, "N", RC)
                If SQL = "" Then
                    'MsgBox "Concepto NO encontrado: " & txtaux(4).Text, vbExclamation
                    txtaux(4).Text = ""
                    RC = "0"
                End If
                'Text3(1).Text = SQL
        Case 6, 7
''''''                'Ponemos el otro campo a ""
''''''                If Index = 6 Then
''''''                    txtaux(7).Text = ""
''''''                Else
''''''                    txtaux(6).Text = ""
''''''                End If
        Case 8
                RC = "idsubcos"
                SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtaux(8).Text, "T", RC)
                If SQL = "" Then
                    MsgBox "Centro coste NO encontrado: " & txtaux(8).Text, vbExclamation
                    'txtAux(8).Text = ""
                End If
                'Text3(2).Text = SQL
                
                    
        Case 11
                Text4.Text = ""
                If IsNumeric(txtaux(11).Text) Then
                    SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtaux(11).Text, "N")
                    Text4.Text = SQL
                End If
        End Select
End Sub


Private Sub LlamaContraPar()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing
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



Private Sub CargaGrid(Enlaza As Boolean)
Dim B As Boolean
    B = DataGrid1.Enabled
    DataGrid1.Enabled = False
    CargaGrid2 Enlaza
    DataGrid1.Enabled = B
End Sub


'//////////////////////////////////////////////
'Esta funcion nos da cual es el tamño campo del
'campo INDICE
'A partir del maximo, mostremos el campo ampliacion
'e iremos copianod el texto en la ampliacion
Private Function MaximaLongitudCampo(indice As Integer) As Integer
    Select Case indice
    Case 0, 2, 3
        MaximaLongitudCampo = 10
    Case 4
        MaximaLongitudCampo = 3
    Case 5
        MaximaLongitudCampo = 25
    Case 6, 7
        MaximaLongitudCampo = 15
    Case 8
        MaximaLongitudCampo = 5
    Case Else
        'Los de "cabeceras"
        MaximaLongitudCampo = 7
    End Select
End Function



Private Function GeneraCadenaConuslta() As Boolean
Dim AUx As String
Dim txtImportes As String
Dim TxtMaxMin As String
Dim TipoDato As String
Dim Cad As String
Dim RC As Byte
Dim J As Integer
GeneraCadenaConuslta = False
'Comprobamos que tiene valores de criterio
'es decir que algun campo tiene valor
SQL = ""
For i = 0 To txtaux.Count - 1
    If Trim(txtaux(i).Text) <> "" Then
        SQL = "x"
        Exit For
    End If
Next i
If SQL = "" Then
    MsgBox "Ponga algún criterio de busqueda", vbExclamation
    Exit Function
End If

SQL = ""
txtImportes = ""
TxtMaxMin = ""
For i = 0 To txtaux.Count - 1
    If txtaux(i).Text <> "" Then
        'Valores para la BD
        ValoresCampo i, AUx, TipoDato
        'Si tiene >> o <<
        If txtaux(i).Text = "<<" Or txtaux(i).Text = ">>" Then
            Cad = "Select "
            If txtaux(i).Text = "<<" Then
                Cad = Cad & "MIN("
            Else
                Cad = Cad & "MAX("
            End If
            J = InStr(1, AUx, ".")
            If J > 0 Then
                Cad = Cad & Mid(AUx, J + 1) & ") FROM "
                Cad = Cad & Mid(AUx, 1, J - 1)
            Else
                Cad = Cad & AUx & ") FROM hlinapu"
                If EjerciciosCerrados Then Cad = Cad & "1"
            End If
            Cad = DevuelveMaxMin(Cad)
            If Cad <> "" Then
                Select Case TipoDato
                Case "N"
                    Cad = AUx & " = " & TransformaComasPuntos(Cad)
                Case "F"
                    Cad = AUx & " = '" & Format(Cad, FormatoFecha) & "'"
                Case "T"
                    Cad = AUx & " = '" & Cad & "'"
                End Select
                If TxtMaxMin <> "" Then TxtMaxMin = TxtMaxMin & " AND "
                TxtMaxMin = TxtMaxMin & Cad
            End If
        Else
            Select Case i
            Case 1
                'Nombre de cuenta, no hacemos nada
                
            Case 6, 7
                    RC = SeparaCampoBusqueda(TipoDato, AUx, txtaux(i).Text, Cad)
                    If RC = 0 Then
                        If txtImportes <> "" Then txtImportes = txtImportes & " OR "
                        txtImportes = txtImportes & "(" & Cad & ")"
                    End If
            Case Else
                    RC = SeparaCampoBusqueda(TipoDato, AUx, txtaux(i).Text, Cad)
                    If RC = 0 Then
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
             End Select
        End If
    End If
Next i

If txtImportes <> "" Then txtImportes = "(" & txtImportes & ")"
If txtImportes <> "" Then
    If SQL <> "" Then SQL = SQL & " AND "
    SQL = SQL & txtImportes
End If

If TxtMaxMin <> "" Then
    TxtMaxMin = "(" & TxtMaxMin & ")"
    If SQL <> "" Then SQL = SQL & " AND "
    SQL = SQL & TxtMaxMin
End If

If SQL = "" Then
    MsgBox "Ninguna cadena de busqueda se ha generado", vbExclamation
    Else
    GeneraCadenaConuslta = True
    Me.Tag = SQL
End If
End Function


Private Sub ValoresCampo(indice As Integer, ByRef CampoBD As String, ByRef TipoDeDatos As String)
    If EjerciciosCerrados Then
        Tablas = "1"
    Else
        Tablas = ""
    End If
    Select Case indice
    Case 0
        CampoBD = "hlinapu" & Tablas & ".codmacta"
        TipoDeDatos = "T"  'Realmente es texto, pero
    Case 2
        CampoBD = "numdocum"
        TipoDeDatos = "T"
    Case 3
        CampoBD = "ctacontr"
        TipoDeDatos = "T"  'Realmente, tb es texto
    Case 4
        CampoBD = "codconce"
        TipoDeDatos = "N"
    Case 5
        CampoBD = "ampconce"
        TipoDeDatos = "T"
    Case 6
        CampoBD = "timporteD"
        TipoDeDatos = "N"
    Case 7
        CampoBD = "timporteH"
        TipoDeDatos = "N"
    Case 8
        CampoBD = "codccost"
        TipoDeDatos = "T"
    'Esto es para cabeceras de apuntes
    Case 9
        CampoBD = "hcabapu" & Tablas & ".numasien"
        TipoDeDatos = "N"
    Case 10
        CampoBD = "hcabapu" & Tablas & ".fechaent"
        TipoDeDatos = "F"
    Case 11
        CampoBD = "hcabapu" & Tablas & ".numdiari"
        TipoDeDatos = "N"
    Case 12
        CampoBD = "obsdiari"
        TipoDeDatos = "T"
    End Select
End Sub




Private Sub CargaGridResultados()
    With DataGridResult
        .Tag = "Estableciendo"
        .AllowRowSizing = False
        .RowHeight = 270
        
        .Tag = "Asignando"
        '------------------------------------------
        'Sabemos que de la consulta los campos
        ' 0.-numaspre  1.- Lin aspre
        '   No se pueden modificar
        ' y ademas el 0 es NO visible
        .Columns(0).Caption = "Fecha"
        .Columns(0).NumberFormat = "dd/mm/yy"
        .Columns(0).Width = 900
        
        .Columns(1).Caption = "Asien."
        .Columns(1).Width = 700
    
        .Columns(2).Visible = False
        
        .Columns(3).Caption = "Cuenta"
        .Columns(3).Width = 1000
    
        .Columns(4).Caption = "Titulo"
        .Columns(4).Width = 2060
    
        .Columns(5).Caption = "Docum."
        .Columns(5).Width = 1030
        
        .Columns(6).Caption = "Ampliación"
        .Columns(6).Width = 2750
        
        .Columns(7).Caption = "Debe"
        .Columns(7).NumberFormat = "#,##0.00"
        .Columns(7).Width = 1000 + ancho
        .Columns(7).Alignment = dbgRight
                
        .Columns(8).Caption = "Haber"
        .Columns(8).NumberFormat = "#,##0.00"
        .Columns(8).Width = 1000 + ancho
        .Columns(8).Alignment = dbgRight
        
        .Columns(9).Visible = False
        
        
        'Centro de coste
        If vParam.autocoste Then
            .Columns(10).Caption = "C.C."
            .Columns(10).Width = 600
            .Columns(10).Alignment = dbgRight
            .Columns(10).Visible = True
        End If
    End With
End Sub



Private Sub PonerAsiento()
Dim B As Byte
    B = Screen.MousePointer
    
    DoEvents
    Me.Refresh
    Screen.MousePointer = vbHourglass
    txtaux(9).Text = RecuperaValor(ASIENTO, 3)
    txtaux(10).Text = Format(RecuperaValor(ASIENTO, 2), "dd/mm/yyyy")
    txtaux(11).Text = RecuperaValor(ASIENTO, 1)
    PonObservaciones
    'Poner diario
    txtAux_LostFocus 11
    
    CargaGrid True 'Cargamos las lineas
    If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.Find "linliapu = " & LINASI
    Screen.MousePointer = B
End Sub



Private Sub PonObservaciones()

        SQL = "SELECT obsdiari FROM hcabapu "
        SQL = SQL & " WHERE numasien = " & txtaux(9).Text
        SQL = SQL & " AND numdiari =" & txtaux(11).Text
        SQL = SQL & " AND fechaent= '" & Format(txtaux(10).Text, FormatoFecha) & "'"
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        txtaux(12).Text = ""
        If Not Rs.EOF Then txtaux(12).Text = DBMemo(Rs.Fields(0))
        Rs.Close
        Set Rs = Nothing
End Sub



Private Function DevuelveMaxMin(ByRef vSQL As String) As String
On Error GoTo EDevuelveMaxMin
Set Rs = New ADODB.Recordset
Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
DevuelveMaxMin = ""
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields(0)) Then
        DevuelveMaxMin = Rs.Fields(0)
    End If
End If
Rs.Close

EDevuelveMaxMin:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Obteniendo valores maximos/minimos"
        Err.Clear
    End If
    Set Rs = Nothing

End Function



'Private Sub FijarFiltro(LeerFiltro As Boolean)
'On Error GoTo EFij
'
'
'    SQL = App.path & "\hcofilt.dat"
'    If LeerFiltro Then
'        'LEER
'        FILTRO = 0
'        If Dir(SQL) <> "" Then
'            lon = FreeFile
'            Open SQL For Input As #lon
'            Line Input #lon, SQL
'            Close #lon
'            lon = 0
'            If SQL <> "" Then
'                If IsNumeric(SQL) Then lon = Val(SQL)
'            End If
'            lon = Abs(lon)
'            If lon > 5 Then lon = 0
'            FILTRO = CByte(lon)
'
'        End If
'        Me.Tag = FILTRO
'        Else
'            'ESCRIBIR
'            If FILTRO = 0 Then
'                'Nos cargamos el archivo
'                If Dir(SQL) <> "" Then Kill SQL
'            Else
'                If FILTRO <> CByte(Me.Tag) Then
'                    lon = FreeFile
'                    Open SQL For Output As #lon
'                    Print #lon, FILTRO
'                    Close #lon
'                End If
'            End If
'        End If
'
'    Exit Sub
'EFij:
'    Err.Clear
'End Sub



'Private Sub PonerOpcionMenuFiltro()
'    'En EjerciciosCerrados no se verá
'    mnFiltro.Visible = Not EjerciciosCerrados
'
'    ClickMenu FILTRO
'
'    If FILTRO >= 4 Then
'        If FILTRO = 4 Then
'            i = 0
'        Else
'            i = 10
'        End If
'        'Comprobaremos el ultimo numero de asiento
'        PonerNumeroASiento i
'    End If
'End Sub


Private Sub PonerNumeroASiento(Diferencia As Integer)
Dim L As Long
    SQL = "Select max(numasien) from hcabapu where fechaent >='" & Format(vParam.fechaini, FormatoFecha) & "'"
    SQL = SQL & " AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not Rs.EOF Then L = DBLet(Rs.Fields(0), "N")
    Rs.Close
    Set Rs = Nothing
    If L = 0 Then Exit Sub
    
    L = L - Diferencia
    If L <= 0 Then L = 1
    'Lo ponemos en el text
    txtaux(9).Text = " >= " & L
    txtaux(10).Text = Format(vParam.fechaini, "ddmmyy") & ":" & Format(vParam.fechafin, "ddmmyy")
        
End Sub


Private Sub ClickMenu(vFilt As Byte)
    FILTRO = vFilt
    mnNinguno.Checked = (FILTRO = 0)
    mnAyer.Checked = (FILTRO = 1)
    Me.mnActual.Checked = (FILTRO = 2)
    Me.mnHace30Dias.Checked = (FILTRO = 3)

    Me.mnUltimo.Checked = (FILTRO = 4)
    Me.mnDiezultimos.Checked = (FILTRO = 5)
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub


Private Sub DataGridResult_GotFocus()

  WheelHook DataGridResult
End Sub
Private Sub DataGridResult_LostFocus()

  WheelUnHook
End Sub

