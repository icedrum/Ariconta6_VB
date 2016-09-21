VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAsientosErr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apuntes con errores"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAsientosErr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSaldoHco 
      Height          =   495
      Index           =   1
      Left            =   11220
      Picture         =   "frmAsientosErr.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Extractos"
      Top             =   7020
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSaldoHco 
      Height          =   495
      Index           =   0
      Left            =   10620
      Picture         =   "frmAsientosErr.frx":685E
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Saldos en historico"
      Top             =   7020
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   7380
      TabIndex        =   53
      Top             =   540
      Width           =   4455
      Begin VB.Label Label3 
         Caption         =   "Asientos con errores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   54
         Top             =   180
         Width           =   3675
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FEF7E4&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Tag             =   "Nº asiento|N|S|0||cabapue|numasien||S|"
      Text            =   "Text1"
      Top             =   810
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Index           =   3
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "Obser|T|S|||cabapue|obsdiari|||"
      Text            =   "frmAsientosErr.frx":D0B0
      Top             =   1320
      Width           =   5775
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "Text4"
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   7380
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Nº asiento predefinido|N|S|0||cabapue|numaspre|||"
      Text            =   "commor"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Fecha entrada|F|N|||cabapue|fechaent|dd/mm/yyyy|S|"
      Text            =   "commor"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   46
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
      TabIndex        =   36
      Top             =   6240
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10680
      TabIndex        =   14
      Top             =   7800
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   60
      TabIndex        =   5
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
      TabIndex        =   35
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
      TabIndex        =   6
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
      TabIndex        =   7
      Top             =   6240
      Width           =   885
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   8
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
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   6
      Left            =   8340
      TabIndex        =   10
      Top             =   6240
      Width           =   1125
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   7
      Left            =   9480
      TabIndex        =   11
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   8
      Left            =   10620
      MaxLength       =   4
      TabIndex        =   12
      Top             =   6240
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   795
      Left            =   7440
      TabIndex        =   21
      Top             =   1440
      Width           =   4335
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2940
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   4
         Left            =   2940
         TabIndex        =   27
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "HABER"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   26
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "DEBE"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   25
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2520
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
      Left            =   10680
      TabIndex        =   18
      Top             =   7890
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Tag             =   "numero diario|N|N|0||cabapue|numdiari||S|"
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   7800
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAsientosErr.frx":D0B7
      Height          =   4455
      Left            =   120
      TabIndex        =   20
      Top             =   2340
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
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
            Object.ToolTipText     =   "Actualizar asiento"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8880
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   495
      Left            =   120
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Frame framelineas 
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   6840
      Width           =   10275
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   2
         Left            =   7800
         TabIndex        =   33
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   32
         Text            =   "Text3"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   420
         Width           =   3135
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   8580
         Picture         =   "frmAsientosErr.frx":D0CC
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   5160
         Picture         =   "frmAsientosErr.frx":DACE
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1980
         Picture         =   "frmAsientosErr.frx":E4D0
         Top             =   180
         Width           =   240
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
         Index           =   2
         Left            =   7800
         TabIndex        =   34
         Top             =   180
         Width           =   795
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
         Index           =   1
         Left            =   4320
         TabIndex        =   31
         Top             =   180
         Width           =   855
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
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame frameextras 
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   6840
      Width           =   10215
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nomctapar"
         DataSource      =   "Adodc1"
         Height          =   315
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   2
      Left            =   7920
      Picture         =   "frmAsientosErr.frx":EED2
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   1
      Left            =   3960
      Picture         =   "frmAsientosErr.frx":EFD4
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   0
      Left            =   2040
      Picture         =   "frmAsientosErr.frx":F9D6
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Asiento predefinido"
      Height          =   195
      Index           =   9
      Left            =   8520
      TabIndex        =   52
      Top             =   600
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Asiento"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   51
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   50
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   195
      Index           =   6
      Left            =   7380
      TabIndex        =   48
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   195
      Index           =   5
      Left            =   1500
      TabIndex        =   47
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Diario"
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   19
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Cod Diario"
      Height          =   195
      Index           =   0
      Left            =   3120
      TabIndex        =   17
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
Attribute VB_Name = "frmAsientosErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DesdeNorma43 As Byte  'La uno y la 2 son validas
Public Datos As String  'Tendra, empipado, numero asiento  y demas

Private Const NO = "No encontrado"
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDi As frmTiposDiario
Attribute frmDi.VB_VarHelpID = -1
Private WithEvents frmPre As frmAsiPre
Attribute frmPre.VB_VarHelpID = -1

'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busquedaa
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
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private SQL As String
Dim I As Integer
Dim ancho As Integer
'Dim colMes As Integer

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

'-------------------------------------------------------------


'Cuando la cuenta lleva contrapartida
Private LlevaContraPartida As Boolean
'Para pasar de lineas a cabeceras
Dim Linliapu As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean



Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    Dim Limp As Boolean
    Dim Mc As Contadores
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
            'No deberia darse el caso
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos modificar
                'PreparaBloquear
                Limp = Modificar
                'TerminaBloquear
                If Limp Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1 Then
                        lblIndicador.Caption = ""
                        PonerModo 2
                    Else
                        PonerModo 0
                    End If
                End If
            End If
            
    Case 5
        Cad = AuxOK
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
        Else
            'Insertaremos, o modificaremos
            If InsertarModificar Then
                'Reestablecemos los campos
                'y ponemos el grid
                cmdAceptar.Visible = False
                DataGrid1.AllowAddNew = False
                CargaGrid True
                Limp = True
                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    'Si ha puesto contrapartida borramos
                    If txtAux(3).Text <> "" Then
                        If LlevaContraPartida Then
                            'Ya lleva la contra partida, luego no hacemos na
                            LlevaContraPartida = False
                        Else
                            If DesdeNorma43 = 0 Then
                                Cad = "Generar asiento de la contrapartida?"
                                If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
                                    FijarContraPartida
                                    Limp = False
                                    LlevaContraPartida = True
                                End If
                            End If
                        End If
                    Else
                        LlevaContraPartida = False
                    End If
                    txtAux(8).Text = ""
                    Text3(2).Text = ""
                    If Limp Then
                        For I = 0 To 2
                            Text3(I).Text = ""
                        Next I
                        For I = 0 To 7
                            txtAux(I).Text = ""
                        Next I
                    End If
                    ModificandoLineas = 0
                    cmdAceptar.Visible = True
                    cmdCancelar.Caption = "Cabecera"
                    AnyadirLinea False
                    If Limp Then
                        txtAux(0).SetFocus
                    Else
                        txtAux(2).SetFocus
                    End If
                Else
                    ModificandoLineas = 0
                    CamposAux False, 0, False
                    cmdCancelar.Caption = "Cabecera"
                End If
            End If
        End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click(Index As Integer)
    cmdAux(0).Tag = 0
    LlamaContraPar
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3
        LimpiarCampos
        PonerModo 0
    Case 4
        lblIndicador.Caption = ""
        PonerModo 2
        PonerCampos
    Case 5
        CamposAux False, 0, False
        frameextras.Visible = True
        framelineas.Visible = False
        LlevaContraPartida = False
        If adodc1.Recordset.EOF Then
              cmdSaldoHco(1).Visible = False
              cmdSaldoHco(0).Visible = False
        End If
        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
            If adodc1.Recordset.EOF Then
                SQL = "El asiento no tiene lineas."
                If Me.DesdeNorma43 > 0 Then SQL = SQL & " Si elije salir cancelar la generacion del asiento. " & vbCrLf
                SQL = SQL & " ¿ Desea salir igualmente?"
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            Else
                'Si el asiento esta descuadrado hbar que dar una notificacion
                If Val(Text2(2).Text) <> 0 Then
                    SQL = "El asiento esta descuadrado."
                    If Me.DesdeNorma43 > 0 Then SQL = SQL & " Si elije SALIR cancelará la generación del asiento. " & vbCrLf
                    SQL = SQL & " Seguro que desea salir de la edición de lineas de asiento ?"
                    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                End If
            End If
            
            'Auqi, si viene de norma 43 saldremos
            If Me.DesdeNorma43 > 0 Then
                CadenaDesdeOtroForm = ""
                If Val(Text2(2).Text) = 0 Then
                    ToolbarActualizarAsiento
                    If CadenaDesdeOtroForm = "" Then
                        MsgBox "No se ha podido generar el asiento.", vbExclamation
                    End If
                End If
                PulsadoSalir = True
                Unload Me
                
            Else
                'Datos
                lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                PonerModo 2
            End If
            

        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
                 DataGrid1.Refresh
            End If
            frameextras.Visible = Not adodc1.Recordset.EOF
            cmdAceptar.Visible = False
            cmdCancelar.Caption = "Cabeceras"
            ModificandoLineas = 0
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
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!Numasien) = Text1(4).Text Then
                If CStr(.Fields!NumDiari) = Text1(0).Text Then
                    If Format(CStr(.Fields!fechaent), "dd/mm/yyyy") = Text1(1).Text Then
                        SituarData1 = True
                        Exit Function
                    End If
                End If
            End If
            .MoveNext
        Wend
    End With
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    cmdSaldoHco(0).Visible = False
    cmdSaldoHco(1).Visible = False
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    '###A mano
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(1).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(4).SetFocus
        Text1(4).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
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
    Me.chkVistaPrevia.SetFocus
    CargaGrid False
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
    Case 1
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 3
        Data1.Recordset.MoveLast
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
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano

    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
End Sub

Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not EliminarDesdeActualizar Then
        '### a mano
        SQL = "Cabecera de apuntes." & vbCrLf
        SQL = SQL & "-----------------------------" & vbCrLf & vbCrLf
        SQL = SQL & "Va a eliminar el asiento erróneo:"
        SQL = SQL & vbCrLf & "Nº Asiento   :   " & Data1.Recordset.Fields(2)
        SQL = SQL & vbCrLf & "Fecha ent    :   " & CStr(Data1.Recordset.Fields(1))
        SQL = SQL & vbCrLf & "Diario           :   " & Text1(0).Text & " - " & Text4.Text & vbCrLf & vbCrLf
        SQL = SQL & "      ¿Desea continuar ? "
        I = MsgBox(SQL, vbQuestion + vbYesNoCancel)
        'Borramos
        If I <> vbYes Then
            Exit Sub
        End If
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
        
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    DataGrid1.Enabled = False
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid False
        PonerModo 0
        Else
            If NumRegElim > Data1.Recordset.RecordCount Then
                Data1.Recordset.MoveLast
            Else
                Data1.Recordset.MoveFirst
                Data1.Recordset.Move NumRegElim
            End If
            PonerCampos
            DataGrid1.Enabled = True
    End If


Error2:
        Screen.MousePointer = vbDefault
        If Not EliminarDesdeActualizar Then

        End If
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

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
Unload Me
End Sub







Private Sub cmdSaldoHco_Click(Index As Integer)
Dim Cta As String
    If Modo = 5 And ModificandoLineas > 0 Then
        If txtAux(0).Text = "" Then
            MsgBox "Seleccione una cuenta", vbExclamation
            Exit Sub
        End If
        SQL = txtAux(0).Text
        Cta = txtAux(1).Text
    Else
        If adodc1.Recordset.EOF Then
            MsgBox "Ningún registro activo.", vbExclamation
            Exit Sub
        End If
        SQL = adodc1.Recordset!codmacta
        Cta = DBLet(adodc1.Recordset!nommacta)
    End If
    If Index = 0 Then
        SaldoHistorico SQL, "", txtAux(1).Text, False
    Else
        Screen.MousePointer = vbHourglass
        frmConExtr.EjerciciosCerrados = False
        frmConExtr.Cuenta = SQL
        frmConExtr.Show vbModal
    End If
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()
Dim B As Boolean
  
    If PrimeraVez Then
        B = False
        PrimeraVez = False
    
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE numasien = "
        If DesdeNorma43 = 0 Then
            CadenaConsulta = CadenaConsulta & "-1"
            Modo = 0
        Else
            CadenaConsulta = CadenaConsulta & RecuperaValor(Datos, 1)
            Modo = 5
        End If
        
        Data1.RecordSource = CadenaConsulta
        Data1.Refresh
        If Modo = 5 Then PonerCampos
        PonerModo CInt(Modo)
        espera 0.2
        CargaGrid (Modo = 5)
        
        
        If DesdeNorma43 > 0 Then
            Label1(8).Caption = "Temporal"
            ModificandoLineas = 0
            'Ponemos en marcha, la maquinaria. Desde variable DATOS extraemos
            If DesdeNorma43 = 1 Then
                AnyadirLinea True
            Else
                'Es TIPO 2. Es decir lo dejamos modificando lineas
                PonerOpcionModificar
                txtAux(5).SelStart = 0
                txtAux(5).SelLength = Len(txtAux(5).Text)
                txtAux(5).SetFocus
            End If
        Else
            Label1(8).Caption = "Nº Asiento"
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerOpcionModificar()

    On Error GoTo EPonerOpcionModificar
    adodc1.Recordset.MoveLast
    ModificarLinea
    Exit Sub
EPonerOpcionModificar:
    MuestraError Err.Number, Err.Description
End Sub



Private Sub Form_Load()

    Me.Icon = frmPpal.Icon


    If Me.DesdeNorma43 = 0 Then
        'Asientos con errores
        Label3.Caption = "Asientos con errores"
    Else
        Label3.Caption = "Generar desde norma 43"
    End If
    Caption = Label3.Caption
    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False
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
        .Buttons(10).Image = 10
        .Buttons(11).Image = 18
        .Buttons(13).Image = 16
        .Buttons(14).Image = 15
        .Buttons(16).Image = 6
        .Buttons(17).Image = 7
        .Buttons(18).Image = 8
        .Buttons(19).Image = 9
    End With
    
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
        Me.Width = 12000
        Me.Height = Screen.Height
    End If
    Me.Height = 9000
    'Los campos auxiliares
    CamposAux False, 0, True
    
    'Si no es analitica no mostramos el label
    Text3(2).Visible = vParam.autocoste
    Label2(2).Visible = vParam.autocoste
    Image1(2).Visible = vParam.autocoste
    Text3(3).Visible = vParam.autocoste
    Label2(3).Visible = vParam.autocoste
    
    
    '## A mano
    NombreTabla = "cabapue"
    Ordenacion = " ORDER BY numasien"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
'    Data1.UserName = vUsu.Login
'    Data1.Password = vUsu.Passwd
'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login

    
    'Maxima longitud cuentas
    txtAux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    txtAux(3).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
'    Data1.CursorType = adOpenDynamic
'    Data1.LockType = adLockPessimistic
    'CadAncho = False

End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
End Sub


Private Sub Form_Resize()
If Me.WindowState <> 0 Then Exit Sub
If Me.Width < 11610 Then Me.Width = 11610
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 3)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas
If cmdAux(0).Tag = 0 Then
    'Cuenta normal
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
Else
    'contrapartida
    txtAux(3).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 2)
End If
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
'Centro de coste
txtAux(8).Text = RecuperaValor(CadenaSeleccion, 1)
Text3(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
Dim RC As Byte
'Concepto
txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1)
Text3(1).Text = RecuperaValor(CadenaSeleccion, 2)
'Habilitamos importes
RC = CByte(Val(RecuperaValor(CadenaSeleccion, 3)))
HabilitarImportes RC
End Sub

Private Sub frmDi_DatoSeleccionado(CadenaSeleccion As String)
Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
Text4.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmPre_DatoSeleccionado(CadenaSeleccion As String)
Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
Text5.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    'Cta contrapartida
    cmdAux(0).Tag = 1
    LlamaContraPar
    txtAux(4).SetFocus
    
Case 1
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|"
    frmCon.Show vbModal
    Set frmCon = Nothing
Case 2
    If txtAux(8).Enabled Then
        Set frmCC = New frmCCoste
        frmCC.DatosADevolverBusqueda = "0|1|"
        frmCC.Show vbModal
        Set frmCC = Nothing
    End If
End Select
End Sub

Private Sub imgppal_Click(Index As Integer)
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
    Case 1
        'Tipos diario
        Set frmDi = New frmTiposDiario
        frmDi.DatosADevolverBusqueda = "0"
        frmDi.Show vbModal
        Set frmDi = Nothing
    Case 2
        'ASiento predefinido
        If Modo = 3 Then
            'Solo si es nuevo
            Set frmPre = New frmAsiPre
            frmPre.DatosADevolverBusqueda = "0"
            frmPre.Show vbModal
            Set frmPre = Nothing
        End If
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar False
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    'Condiciones para NO salir
    
    If Modo = 5 Then
        If DesdeNorma43 = 0 Then
            Exit Sub
        Else
            'Esta en el punteo BANCARIO. CON lo cual preguntamos si desea cancelar la edicioon del punteo
            ' y borramos de la tabla
            SQL = "Desea cancelar la generacion del asiento?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            If Not Eliminar Then Exit Sub
            CadenaDesdeOtroForm = "" 'para que no punteee el banco
        End If
    End If
    PulsadoSalir = True
    Screen.MousePointer = vbHourglass
    DataGrid1.Enabled = False
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


'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 1:  KEYBusqueda KeyAscii, 0
            Case 0:  KEYBusqueda KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgppal_Click (Indice)
End Sub
'++

'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim RC As Byte
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite  '&H80000018
    End If
    
    'Si estamos insertando o modificando o buscando
    If Modo = 3 Or Modo = 4 Then
        If Text1(Index).Text = "" Then
            If Index = 0 Then
                Text4.Text = ""
            Else
                If Index = 2 Then Text5.Text = ""
            End If
            Exit Sub
        End If
        Select Case Index
        Case 0
            'Tipo diario
            If Not IsNumeric(Text1(0).Text) Then
                MsgBox "Tipo de diario no es numérico: " & Text1(0).Text, vbExclamation
                Text1(0).Text = ""
                Text4.Text = ""
                Text1(0).SetFocus
                Exit Sub
            End If
             SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(0).Text, "N")
             If SQL = "" Then
                    SQL = "Diario no encontrado: " & Text1(0).Text
                    Text1(0).Text = ""
                    Text4.Text = ""
                    MsgBox SQL, vbExclamation
                    Text1(0).SetFocus
            End If
            Text4.Text = SQL
        Case 1
            SQL = ""
            If Not EsFechaOK(Text1(1)) Then
                MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
                SQL = "mal"
            Else
                RC = FechaCorrecta2(CDate(Text1(1).Text))
                'Text1(1).Text = Format(Text1(1).Text, "dd/mm/yyyy")
                SQL = ""
                If RC > 1 Then
                    If RC = 2 Then
                        SQL = varTxtFec
                    Else
                        If RC = 3 Then
                            SQL = "El ejercicio al que pertenece la fecha: " & Text1(Index).Text & " está cerrado."
                        Else
                            SQL = "Ejercicio para: " & Text1(Index).Text & " todavía no activo"
                        End If
                     End If
                     MsgBox SQL, vbExclamation
                 End If
            End If
            If SQL <> "" Then
                Text1(1).Text = ""
                Text1(1).SetFocus
            End If
            
        Case 2
            SQL = DevuelveDesdeBD("nomaspre", "cabasipre", "numaspre", Text1(2).Text, "N")
            If SQL = "" Then
                Text1(2).Text = "-1"
                SQL = NO
            End If
            Text5.Text = SQL
        End Select
    End If
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
        Cad = Cad & ParaGrid(Text1(4), 20, "Nº Asiento:")
        Cad = Cad & ParaGrid(Text1(1), 30, "Fecha Entrada")
        Cad = Cad & ParaGrid(Text1(0), 15, "Nº Diario")
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.VCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Asientos"
            frmB.vSelElem = 0
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                'If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
               ' Text1(kCampo).SetFocus
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla de Apuntes", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
        Else
            PonerModo 2
            'Data1.Recordset.MoveLast
            Data1.Recordset.MoveFirst
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
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    
    Text1(4).Text = Data1.Recordset!Numasien
    Text1(1).Text = Format(Data1.Recordset!fechaent, "dd/mm/yyyy")
    Text1(0).Text = Data1.Recordset!NumDiari
    Text1(3).Text = DBLet(Data1.Recordset!obsdiari)
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True
    'Cargamos datos extras
    SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(0).Text, "N")
    If SQL = "" Then SQL = ""
    Text4.Text = SQL
    
    If Text1(2).Text = "" Then
        SQL = ""
    Else
        SQL = DevuelveDesdeBD("nomaspre", "cabasipre", "numaspre", Text1(2).Text, "N")
        If SQL = "" Then SQL = "Error en nº de asiento predefinido"
    End If
    Text5.Text = SQL
    frameextras.Visible = Not adodc1.Recordset.EOF
    cmdSaldoHco(0).Visible = Not adodc1.Recordset.EOF
    cmdSaldoHco(1).Visible = Not adodc1.Recordset.EOF

    If Modo = 2 Then lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean
    

    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
'        For i = 0 To Text1.Count - 1
'            Text1(i).BackColor = vbWhite
'            'Text1(0).BackColor = &H80000018
'        Next i
'        'chkVistaPrevia.Visible = False
        'Reestablecemos el color del nuº asien
        Text1(4).BackColor = &HFEF7E4
    End If
    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nuevo apunte diario"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar apunte diario"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar apunte diario"
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea apunte diario"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea apunte diario"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea apunte diario"
    End If
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    frameextras.Visible = B
    If B Then framelineas.Visible = False
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B  'Lineas factur
    Toolbar1.Buttons(11).Enabled = B
    If Not B Then frameextras.Visible = False
        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = (Modo = 5)
    cmdAceptar.Visible = B Or Modo = 1
    'PRueba###
    
    If Modo < 2 Then
        Me.cmdSaldoHco(0).Visible = False
        Me.cmdSaldoHco(1).Visible = False
    End If

    '
    B = B Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    mnOpcionesAsiPre.Enabled = Not B
   
   
        'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5

    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.Visible = (Modo = 2)
'    Else
'        cmdRegresar.Visible = False
'    End If
    
    '
    Text1(4).Enabled = (Modo = 1)
    Text1(2).Enabled = (Modo = 3 Or Modo = 1) 'Solo insertar
    B = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    Text1(0).Enabled = B
    Text1(1).Enabled = B
    Text1(3).Enabled = B
    'El text
    B = (Modo = 2) Or (Modo = 5)
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
    
    B = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = B
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    PonerOpcionesMenuGeneral Me
    
End Sub


Private Function DatosOk() As Boolean
    Dim RS As ADODB.Recordset
    Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Campos encabezados correctos
    'Comprobaremos el resto de detalles
    DatosOk = False
    '       0 .- Año actual
    '       1 .- Siguiente
    '       2 .- Anterior al inicio
    '       3 .- Posterior al fin
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            MsgBox "La fecha no pertenece al ejercicio actual ni al siguiente", vbExclamation
        End If
        Exit Function
    End If
    
    DatosOk = True
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
            AnyadirLinea True
        End If
    Case 7
        If Modo <> 5 Then
            'Intentamos bloquear la cuenta
            BotonModificar
        Else
            'MODIFICAR linea factura
            ModificarLinea
        End If
    Case 8
        If Modo <> 5 Then
            BotonEliminar False
        Else
            'ELIMINAR linea factura
            EliminarLineaFactura
        End If
    Case 10
   
        'Nuevo Modo
        PonerModo 5
        'Fuerzo que se vean las lineas
        frameextras.Visible = True
        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
    Case 11
        'Acturalizarr asiento
        ToolbarActualizarAsiento
    Case 13
        'Imprimir asientos
        Screen.MousePointer = vbHourglass
        frmActualizar.NUmSerie = Text1(4).Text
        'frmActualizar.OpcionActualizar = 4
        
        frmActualizar.OpcionActualizar = 13
        frmActualizar.Show vbModal
    
    Case 14
        mnSalir_Click
    Case 16 To 19
        Desplazamiento (Button.Index - 16)
    Case Else
    
    End Select
End Sub



Private Sub ToolbarActualizarAsiento()
Dim NuAsi As Long
        'ACtualizar asiento
        If Data1.Recordset.EOF Then
            MsgBox "Ningún asiento para actualizar.", vbExclamation
            Exit Sub
        End If
        If adodc1 Is Nothing Then Exit Sub
        If adodc1.Recordset.EOF Then
            MsgBox "No hay lineas insertadas para este asiento", vbExclamation
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        If ComprobarLineas Then
            If ActualizarASiento(NuAsi) Then ActualizarASientoHCO (NuAsi)
            BotonEliminar True
            'Si hallegado hasta aqui es k ha ido bien
            If Me.DesdeNorma43 > 0 Then CadenaDesdeOtroForm = "OK"
        End If
        Screen.MousePointer = vbDefault

End Sub

'Private Function BloqAsien() As Boolean
'Dim b As Byte
''Lo bloqueamos
'        b = Screen.MousePointer
'        Screen.MousePointer = vbHourglass
'        BloqAsien = True
'        I = (BloquearAsiento(Text1(4).Text, Text1(0).Text, Format(Text1(1).Text, FormatoFecha)))
'        'Bloqueamos el registro
'        If I = 0 Then
'            MsgBox "Asiento bloqueado", vbExclamation
'        Else
'            BloqAsien = False
'        End If
'        Screen.MousePointer = b
'End Function
'
'
'Private Sub DesBloqAsien()
'Dim b As Byte
''Lo bloqueamos
'        b = Screen.MousePointer
'        Screen.MousePointer = vbHourglass
'        I = (DesbloquearAsiento(Text1(4).Text, Text1(0).Text, Format(Text1(1).Text, FormatoFecha)))
'        Screen.MousePointer = b
'End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    For I = 16 To 19
        Toolbar1.Buttons(I).Visible = bol
    Next I
End Sub


Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = MontaSQLCarga(Enlaza)
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockPessimistic
    adodc1.Refresh
    
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
    DataGrid1.Columns(10).Width = 1154 + ancho
    DataGrid1.Columns(10).Alignment = dbgRight
            
    DataGrid1.Columns(11).Caption = "Haber"
    DataGrid1.Columns(11).NumberFormat = "#,##0.00"
    DataGrid1.Columns(11).Width = 1154 + ancho
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
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        anc = 323
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(2).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(3).Left + 90
                
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 6
        txtAux(1).Width = DataGrid1.Columns(3).Width - 180
    
        txtAux(2).Left = DataGrid1.Columns(4).Left + 150
        txtAux(2).Width = DataGrid1.Columns(4).Width - 30
    
        txtAux(3).Left = DataGrid1.Columns(5).Left + 150
        txtAux(3).Width = DataGrid1.Columns(5).Width - 45

        
        'Concepto
        txtAux(4).Left = DataGrid1.Columns(6).Left + 150
        txtAux(4).Width = DataGrid1.Columns(6).Width - 45
        
        txtAux(5).Left = DataGrid1.Columns(8).Left + 150
        txtAux(5).Width = DataGrid1.Columns(8).Width - 45
        
        txtAux(6).Left = DataGrid1.Columns(10).Left + 150
        txtAux(6).Width = DataGrid1.Columns(10).Width - 30
        
       
        txtAux(7).Left = DataGrid1.Columns(11).Left + 150
        txtAux(7).Width = DataGrid1.Columns(11).Width - 30
       
        txtAux(8).Left = DataGrid1.Columns(12).Left + 150
        txtAux(8).Width = DataGrid1.Columns(12).Width - 45
       
        CadAncho = True
    End If
        
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
    
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
    Dim RS As ADODB.Recordset
    If Data1.Recordset.EOF Then
        Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
        Exit Sub
    End If
    
    If adodc1.Recordset.EOF Then
        Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
        Exit Sub
    End If
    
    
    
    Set RS = New ADODB.Recordset
    SQL = "SELECT Sum(linapue.timporteD) AS SumaDetimporteD, Sum(linapue.timporteH) AS SumaDetimporteH"
    SQL = SQL & " ,linapue.numdiari,linapue.fechaent,linapue.numasien"
    SQL = SQL & " From linapue GROUP BY linapue.numdiari, linapue.fechaent, linapue.numasien "
    SQL = SQL & " HAVING (((linapue.numdiari)=" & Data1.Recordset!NumDiari
    SQL = SQL & ") AND ((linapue.fechaent)='" & Format(Data1.Recordset!fechaent, FormatoFecha)
    SQL = SQL & "') AND ((linapue.numasien)=" & Data1.Recordset!Numasien
    SQL = SQL & "));"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Deb = 0
    hab = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Deb = RS.Fields(0)
        If Not IsNull(RS.Fields(1)) Then hab = RS.Fields(1)
    End If
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
    SQL = "SELECT linapue.numasien, linapue.linliapu, linapue.codmacta, cuentas.nommacta,"
    SQL = SQL & " linapue.numdocum, linapue.ctacontr, linapue.codconce, conceptos.nomconce as nombreconcepto, linapue.ampconce, cuentas_1.nommacta as nomctapar,"
    SQL = SQL & " linapue.timporteD,linapue.timporteH, linapue.codccost, cabccost.nomccost as centrocoste,"
    SQL = SQL & " linapue.numdiari, linapue.fechaent"
    SQL = SQL & " FROM (((linapue LEFT JOIN cuentas AS cuentas_1 ON linapue.ctacontr ="
    SQL = SQL & " cuentas_1.codmacta) LEFT JOIN cabccost ON linapue.codccost = cabccost.codccost)"
    SQL = SQL & " LEFT JOIN cuentas ON linapue.codmacta = cuentas.codmacta) LEFT JOIN"
    SQL = SQL & " conceptos ON linapue.codconce = conceptos.codconce"
    If Enlaza Then
        SQL = SQL & " WHERE numasien = " & Data1.Recordset!Numasien
        SQL = SQL & " AND numdiari =" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent= '" & Format(Data1.Recordset!fechaent, FormatoFecha) & "'"
        Else
        SQL = SQL & " WHERE numasien = -1"
    End If
    SQL = SQL & " ORDER BY linapue.linliapu"
    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea(Limpiar As Boolean)
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    Linliapu = ObtenerSigueinteNumeroLinea
    'Situamos el grid al final
    
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If
    LLamaLineas anc, 1, Limpiar
    HabilitarImportes 0
    'Ponemos el foco
    txtAux(0).SetFocus
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    Linliapu = adodc1.Recordset!Linliapu
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    txtAux(0).Text = adodc1.Recordset.Fields!codmacta
    txtAux(1).Text = DBLet(adodc1.Recordset.Fields!nommacta)
    txtAux(2).Text = DataGrid1.Columns(4).Text
    txtAux(3).Text = DataGrid1.Columns(5).Text
    txtAux(4).Text = DataGrid1.Columns(6).Text
    txtAux(5).Text = DataGrid1.Columns(8).Text
    Cad = DBLet(adodc1.Recordset.Fields!timported)
    If Cad <> "" Then
        txtAux(6).Text = Format(Cad, "0.00")
    Else
        txtAux(6).Text = Cad
    End If
    Cad = DBLet(adodc1.Recordset.Fields!timporteH)
    If Cad <> "" Then
        txtAux(7).Text = Format(Cad, "0.00")
    Else
        txtAux(7).Text = Cad
    End If
    txtAux(8).Text = DBLet(adodc1.Recordset.Fields!codccost)
    HabilitarImportes 3
    HabilitarCentroCoste
    Text3(0).Text = Text3(5).Text
    Text3(1).Text = Text3(4).Text
    Text3(2).Text = Text3(3).Text
    LLamaLineas anc, 2, False
    txtAux(0).SetFocus
End Sub

Private Sub EliminarLineaFactura()
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Lineas de apuntes con errores." & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar la linea: "
    SQL = SQL & adodc1.Recordset.Fields(3) & " - " & DataGrid1.Columns(10).Text & " " & DataGrid1.Columns(11).Text
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = "Delete from linapue"
        SQL = SQL & " WHERE linapue.linliapu = " & adodc1.Recordset!Linliapu
        SQL = SQL & " AND linapue.numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND linapue.fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND linapue.numasien=" & Data1.Recordset!Numasien & ";"
        DataGrid1.Enabled = False
        Conn.Execute SQL
        CargaGrid (Not Data1.Recordset.EOF)
        DataGrid1.Enabled = True
    End If
End Sub



Private Function ObtenerSigueinteNumeroLinea() As Long
    Dim RS As ADODB.Recordset
    Dim I As Long
    
    Set RS = New ADODB.Recordset
    SQL = "SELECT Max(linliapu) FROM linapue"
    SQL = SQL & " WHERE linapue.numdiari=" & Data1.Recordset!NumDiari
    SQL = SQL & " AND linapue.fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
    SQL = SQL & "' AND linapue.numasien=" & Data1.Recordset!Numasien & ";"
    RS.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then I = RS.Fields(0)
    End If
    RS.Close
    ObtenerSigueinteNumeroLinea = I + 1
End Function



'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------

Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtAux(0).Text <> "" Then
              hab = HayKHabilitarCentroCoste(txtAux(0).Text)
        Else
            txtAux(8).Text = ""
        End If
        If hab Then
            txtAux(8).BackColor = &H80000005
            Else
            txtAux(8).Text = ""
            Text3(2).Text = ""
            txtAux(8).BackColor = &H80000018
        End If
    End If
    txtAux(8).Enabled = hab
    Image1(2).Enabled = hab
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid DataGrid1
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)
    framelineas.Visible = Not B
    frameextras.Visible = B
    'Habilitamos los botones de cuenta
    'Habilitamos los botones de cuenta
    cmdSaldoHco(1).Visible = Not B
    cmdSaldoHco(0).Visible = Not B
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    frameextras.Visible = Not B
    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim I As Integer
    Dim J As Integer
    
    DataGrid1.Enabled = Not Visible
    If vParam.autocoste Then
        J = txtAux.Count - 1
        Else
        J = txtAux.Count - 2
        txtAux(8).Visible = False
    End If
    For I = 0 To J
        txtAux(I).Visible = Visible
        txtAux(I).Top = Altura
    Next I
        cmdAux(0).Visible = Visible
        cmdAux(0).Top = Altura
    If Limpiar Then
        For I = 0 To J
            txtAux(I).Text = ""
        Next I
        For I = 0 To 3
            Text3(I).Text = ""
        Next I
    End If
    
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
With txtAux(Index)
    If Index <> 5 Then
        .SelStart = 0
        .SelLength = Len(.Text)
    Else
        .SelStart = Len(.Text)
    End If
End With

End Sub

Private Sub txtaux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Ha pulsado F5. Ponemos linea anterior
     If KeyCode = 116 Then PonerLineaAnterior (Index)
End Sub


'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYImage KeyAscii, 3
            Case 3:  KEYImage1 KeyAscii, 0
            Case 4:  KEYImage1 KeyAscii, 1
            Case 8:  KEYImage1 KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub

Private Sub KEYImage1(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub

'++

'1.-Debe    2.-Haber   3.-Decide en asiento
Private Sub HabilitarImportes(tipoConcepto As Byte)
    Dim bDebe As Boolean
    Dim bHaber As Boolean
    
    'Vamos a hacer .LOCKED =
    Select Case tipoConcepto
    Case 1
        bDebe = False
        bHaber = True
    Case 2
        bDebe = True
        bHaber = False
    Case 3
        bDebe = False
        bHaber = False
    Case Else
        bDebe = True
        bHaber = True
    End Select
    
    txtAux(6).Enabled = Not bDebe
    txtAux(7).Enabled = Not bHaber
    
    If bDebe Then
        txtAux(6).BackColor = &H80000018
        Else
        txtAux(6).BackColor = &H80000005
    End If
    If bHaber Then
        txtAux(7).BackColor = &H80000018
        Else
        txtAux(7).BackColor = &H80000005
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Sng As Double
        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtAux(Index).Text = Trim(txtAux(Index).Text)
    
        'Comun a todos
        If txtAux(Index).Text = "" Then
            Select Case Index
            Case 0
                HabilitarCentroCoste
                txtAux(1).Text = ""
            Case 3
                Text3(0).Text = ""
            Case 4
                HabilitarImportes 0
            End Select
            Exit Sub
        End If
        
        Select Case Index
        Case 0
            'Cta
            
            RC = txtAux(0).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtAux(0).Text = RC
                txtAux(1).Text = SQL
                RC = ""
            Else
                MsgBox SQL, vbExclamation
                txtAux(0).Text = ""
                txtAux(1).Text = ""
                RC = "NO"
            End If
            HabilitarCentroCoste
            If RC <> "" Then txtAux(0).SetFocus
            
        Case 3
            RC = txtAux(3).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtAux(3).Text = RC
                Text3(0).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtAux(3).Text = ""
                Text3(0).Text = ""
                txtAux(3).SetFocus
            End If
            
        Case 4
             If Not IsNumeric(txtAux(4).Text) Then
                    MsgBox "El concepto debe de ser numérico", vbExclamation
                    Exit Sub
                End If
                
                If Val(txtAux(4).Text) >= 900 Then
                    MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
                    Text3(1).Text = ""
                    txtAux(4).Text = ""
                    txtAux(4).SetFocus
                    Exit Sub
                End If
                
               
                RC = "tipoconce"
                SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtAux(4).Text, "N", RC)
                If SQL = "" And RC = "tipoconce" Then
                    MsgBox "Concepto NO encontrado: " & txtAux(4).Text, vbExclamation
                    txtAux(4).Text = ""
                    RC = "0"
                End If
                HabilitarImportes CByte(Val(RC))
                Text3(1).Text = SQL
                txtAux(5).Text = SQL
                If txtAux(5).Text <> "" Then txtAux(5).Text = txtAux(5).Text & " "
                If RC = "0" Then txtAux(4).SetFocus
                
                
        Case 6, 7
                'LOS IMPORTES

                
                If Not IsNumeric(txtAux(Index).Text) Then
                    MsgBox "Importes deben ser numéricos.", vbExclamation
                    On Error Resume Next
                    txtAux(Index).Text = ""
                    txtAux(Index).SetFocus
                    Exit Sub
                End If
                
                
                'Es numerico
                SQL = TransformaPuntosComas(txtAux(Index).Text)
                Sng = Round(CDbl(SQL), 2)
                txtAux(Index).Text = Format(Sng, "0.00")
                
                'Ponemos el otro campo a ""
                If Index = 6 Then
                    txtAux(7).Text = ""
                Else
                    txtAux(6).Text = ""
                End If
        Case 8
                txtAux(8).Text = UCase(txtAux(8).Text)
                RC = "idsubcos"
                SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtAux(8).Text, "T", RC)
                If SQL = "" Then
                    MsgBox "Concepto NO encontrado: " & txtAux(8).Text, vbExclamation
                    txtAux(8).Text = ""
                End If
                Text3(2).Text = SQL
        End Select
End Sub


Private Function AuxOK() As String
    
    'Cuenta
    If txtAux(0).Text = "" Then
        AuxOK = "Cuenta no puede estar vacia."
        Exit Function
    End If
    
    If Not IsNumeric(txtAux(0).Text) Then
        AuxOK = "Cuenta debe ser numrica"
        Exit Function
    End If
    
    If txtAux(1).Text = NO Then
        AuxOK = "La cuenta debe estar dada de alta en el sistema"
        Exit Function
    End If
    
    If Not EsCuentaUltimoNivel(txtAux(0).Text) Then
        AuxOK = "La cuenta no es de último nivel"
        Exit Function
    End If
    
    
    'Contrapartida
    If txtAux(3).Text <> "" Then
        If Not IsNumeric(txtAux(3).Text) Then
            AuxOK = "Cuenta contrapartida debe ser numérica"
            Exit Function
        End If
        If Text3(0).Text = NO Then
            AuxOK = "La cta. contrapartida no esta dad de alta en el sistema."
            Exit Function
        End If
        If Not EsCuentaUltimoNivel(txtAux(3).Text) Then
            AuxOK = "La cuenta contrapartida no es de último nivel"
            Exit Function
        End If
    End If
        
    'Concepto
    If txtAux(4).Text = "" Then
        AuxOK = "El concepto no puede estar vacio"
        Exit Function
    End If
        
    If txtAux(4).Text <> "" Then
        If Not IsNumeric(txtAux(4).Text) Then
            AuxOK = "El concepto debe de ser numérico."
            Exit Function
        End If
    End If
    
    'Importe
    If txtAux(6).Text <> "" Then
        If Not IsNumeric(txtAux(6).Text) Then
            AuxOK = "El importe DEBE debe ser numérico"
            Exit Function
        End If
    End If
    
    If txtAux(7).Text <> "" Then
        If Not IsNumeric(txtAux(7).Text) Then
            AuxOK = "El importe HABER debe ser numérico"
            Exit Function
        End If
    End If
    
    If Not (txtAux(6).Text = "" Xor txtAux(7).Text = "") Then
        AuxOK = "Solo el debe, o solo el haber, tiene que tener valor"
        Exit Function
    End If
    
    
    'cENTRO DE COSTE
    If txtAux(8).Enabled Then
        If txtAux(8).Text = "" Then
            AuxOK = "Centro de coste no puede ser nulo"
            Exit Function
        End If
    End If
    
    AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS
        'INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab) VALUES (1, '2003-01-16', 1, 2, '5720001', 'doc', 1, NULL, 1600, NULL, NULL, NULL, NULL)
        SQL = "INSERT INTO linapue (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
        SQL = SQL & "codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada) VALUES ("
        'Nudiari, fechaentra y numasien es fijo
        SQL = SQL & Data1.Recordset!NumDiari & ",'"
        SQL = SQL & Format(Data1.Recordset!fechaent, FormatoFecha) & "'," & Data1.Recordset!Numasien & ","
        SQL = SQL & Linliapu & ",'"
        SQL = SQL & txtAux(0).Text & "','"
        SQL = SQL & txtAux(2).Text & "',"
        SQL = SQL & txtAux(4).Text & ",'"
        SQL = SQL & txtAux(5).Text & "',"
        If txtAux(6).Text = "" Then
          SQL = SQL & ValorNulo & "," & TransformaComasPuntos(txtAux(7).Text) & ","
          Else
          SQL = SQL & TransformaComasPuntos(txtAux(6).Text) & "," & ValorNulo & ","
        End If
        'Centro coste
        If txtAux(8).Text = "" Then
          SQL = SQL & ValorNulo & ","
          Else
          SQL = SQL & "'" & txtAux(8).Text & "',"
        End If
        
        If txtAux(3).Text = "" Then
          SQL = SQL & ValorNulo & ","
          Else
          SQL = SQL & "'" & txtAux(3).Text & "',"
        End If
        'Marca de entrada manual de datos
        SQL = SQL & "'CONTAB',"
        
        'Si viene desde norma43 entonces la marcamos como punteada
        SQL = SQL & "0)"
        
    Else
    
        'MODIFICAR
        'UPDATE linasipre SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
        '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
        SQL = "UPDATE linapue SET "
        
        SQL = SQL & " codmacta = '" & txtAux(0).Text & "',"
        SQL = SQL & " numdocum = '" & txtAux(2).Text & "',"
        SQL = SQL & " codconce = " & txtAux(4).Text & ","
        SQL = SQL & " ampconce = '" & txtAux(5).Text & "',"
        If txtAux(6).Text = "" Then
          SQL = SQL & " timporteD = " & ValorNulo & "," & " timporteH = " & TransformaComasPuntos(txtAux(7).Text) & ","
          Else
          SQL = SQL & " timporteD = " & TransformaComasPuntos(txtAux(6).Text) & "," & " timporteH = " & ValorNulo & ","
        End If
        'Centro coste
        If txtAux(8).Text = "" Then
          SQL = SQL & " codccost = " & ValorNulo & ","
          Else
          SQL = SQL & " codccost = '" & txtAux(8).Text & "',"
        End If
        
        If txtAux(3).Text = "" Then
          SQL = SQL & " ctacontr = " & ValorNulo
          Else
          SQL = SQL & " ctacontr = '" & txtAux(3).Text & "'"
        End If
    
        SQL = SQL & " WHERE linapue.linliapu = " & Linliapu
        SQL = SQL & " AND linapue.numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND linapue.fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND linapue.numasien=" & Data1.Recordset!Numasien & ";"
    
    End If
    Conn.Execute SQL
    InsertarModificar = True
    Exit Function
EInsertarModificar:
        MuestraError Err.Number, "InsertarModificar linea asiento.", Err.Description
End Function
 

Private Sub LlamaContraPar()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

'Private Sub DeseleccionaGrid()
'    On Error GoTo EDeseleccionaGrid
'
'    While DataGrid1.SelBookmarks.Count > 0
'        DataGrid1.SelBookmarks.Remove 0
'    Wend
'    Exit Sub
'EDeseleccionaGrid:
'        Err.Clear
'End Sub

Private Sub FijarContraPartida()
    Dim Cad As String
    'Hay contrapartida
    'Reasignamos campos de cuentas
    Cad = txtAux(0).Text
    txtAux(0).Text = txtAux(3).Text
    txtAux(3).Text = Cad
    HabilitarCentroCoste
    Cad = txtAux(1).Text
    txtAux(1).Text = Text3(0).Text
    Text3(0).Text = Cad
    
    'Los importes
    HabilitarImportes 3
    Cad = txtAux(6).Text
    txtAux(6).Text = txtAux(7).Text
    txtAux(7).Text = Cad
End Sub








Private Sub CargaGrid(Enlaza As Boolean)
Dim B As Boolean
    B = DataGrid1.Enabled
    CargaGrid2 Enlaza
    DataGrid1.Enabled = B
End Sub




Private Function Eliminar() As Boolean
On Error GoTo FinEliminar
        Conn.BeginTrans
        SQL = " WHERE  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien & ";"
        
        'Lineas
        Conn.Execute "Delete  from linapue " & SQL
        
        'Cabeceras1
        Conn.Execute "Delete  from cabapue " & SQL
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Function Modificar() As Boolean
      
    On Error GoTo EModificar
        Conn.BeginTrans
        'Comun
        
        SQL = " WHERE  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien

        'BLoqueamos
        Conn.Execute "Select * from cabapue " & SQL
        
        SQL = " numdiari= " & Text1(0).Text & " , fechaent ='" & Format(Text1(1).Text, FormatoFecha) & "'" & SQL
        
       'Las lineas de apuntes
        Conn.Execute "UPDATE linapue SET " & SQL
      
        
        'Modificamos la cabecera
        If Text1(3).Text = "" Then
            SQL = "obsdiari = NULL," & SQL
        Else
            SQL = "Obsdiari ='" & Text1(3).Text & "'," & SQL
        End If
        'Data1.Recordset.Close
        Conn.Execute "UPDATE cabapue SET " & SQL
        
                
  
EModificar:
        If Err.Number <> 0 Then
            MuestraError Err.Number
            Conn.RollbackTrans
            Modificar = False
        Else
            Conn.CommitTrans
            Modificar = True
        End If
End Function



Private Function ComprobarLineas() As Boolean
Dim Cad As String
Dim Errores As String

        ComprobarLineas = False
        SQL = " numasien = " & Data1.Recordset!Numasien
        SQL = SQL & " AND numdiari =" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent= '" & Format(Data1.Recordset!fechaent, FormatoFecha) & "'"
        Errores = ""
        Linliapu = adodc1.Recordset.RecordCount
        
        
        'Comprobamos la fecha del asiento
        I = FechaCorrecta2(CDate(Text1(1).Text))
        If I > 1 Then
            If I = 2 Then
                    Errores = Errores & "    .- " & varTxtFec & vbCrLf
            Else
                If I < 4 Then
                    Errores = Errores & "    .- La fecha pertenece a un ejercicio cerrado" & vbCrLf
                Else
                    Errores = Errores & "    .- La fecha pertenece a un ejercicio todavia no abierto" & vbCrLf
                End If
            End If
        End If
        
                
        'Si esta cuadrado o no el asiento
        If Text2(2).Text <> "" Then Errores = Errores & "    .- Asiento sin cuadrar" & vbCrLf
        
        
        'El diario
        If Text4.Text = "" Then Errores = Errores & "    .- Número de diario incorrecto" & vbCrLf
        
        '----------------------------------------------------------
        ' Cuentas
        Set miRsAux = New ADODB.Recordset
        Cad = "SELECT count(linapue.linliapu)"
        Cad = Cad & " FROM linapue LEFT JOIN cuentas ON linapue.codmacta = cuentas.codmacta"
        Cad = Cad & " WHERE cuentas.apudirec =""S"" AND " & SQL
        I = 0
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            I = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        If I <> Linliapu Then Errores = Errores & "    .- Cuentas incorrectas" & vbCrLf
            
        
        '----------------------------------------------------------
        ' Concepto
        Set miRsAux = New ADODB.Recordset
        Cad = "SELECT count(linapue.linliapu) FROM linapue INNER JOIN conceptos ON linapue.codconce = conceptos.codconce"
        Cad = Cad & " WHERE " & SQL
        I = 0
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            I = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        If I <> Linliapu Then Errores = Errores & "    .- Conceptos incorrectos" & vbCrLf
        
        
        
        'Contrapartidas
        Cad = "Select count(ctacontr) from linapue where  not (ctacontr is null) and " & SQL
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Linliapu = 0
        If Not miRsAux.EOF Then Linliapu = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        If Linliapu <> 0 Then
            'Hay contrapartidas, vamos a enlazar
            Cad = "SELECT count(ctacontr)"
            Cad = Cad & " FROM linapue INNER JOIN cuentas ON linapue.ctacontr = cuentas.codmacta"
            Cad = Cad & " WHERE cuentas.apudirec =""S"" AND " & SQL
            I = 0
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then I = DBLet(miRsAux.Fields(0), "N")
            miRsAux.Close
            If I <> Linliapu Then Errores = Errores & "    .- Contrapartidas incorrectas" & vbCrLf
        End If
        
        'Centros de coste
        Cad = "Select count(codccost) from linapue where  not (codccost is null) and " & SQL
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Linliapu = 0
        If Not miRsAux.EOF Then Linliapu = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        If Linliapu <> 0 Then
            'Hay contrapartidas, vamos a enlazar
            Cad = "SELECT count(linliapu)"
            Cad = Cad & " FROM linapue INNER JOIN cabccost ON linapue.codccost = cabccost.codccost WHERE " & SQL
            I = 0
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then I = DBLet(miRsAux.Fields(0), "N")
            miRsAux.Close
            If I <> Linliapu Then Errores = Errores & "    .- Centros de coste incorrectos" & vbCrLf
        End If
        
        
        
        'Bloqueo de cuentas a partir de una fecha
        If vParam.CuentasBloqueadas <> "" Then
            Cad = "select codmacta from linapue WHERE " & SQL & " GROUP BY codmacta"
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 0
            While Not miRsAux.EOF
                If EstaLaCuentaBloqueada(miRsAux!codmacta, CDate(Text1(1).Text)) Then
                    If I = 0 Then Errores = Errores & "    .- Cuentas bloqueadas: " & vbCrLf
                    Errores = Errores & "                " & miRsAux!codmacta & vbCrLf
                    I = I + 1
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        
        Set miRsAux = Nothing
        If Errores <> "" Then
            Cad = "No se puede contabilizar el asiento. Contiene datos erróneos." & vbCrLf & vbCrLf & Errores
            MsgBox Cad, vbExclamation
        Else
            ComprobarLineas = True
        End If
End Function



Private Function ActualizarASientoHCO(NumeroAsiento As Long) As Boolean


    Screen.MousePointer = vbHourglass
    frmActualizar.OpcionActualizar = 1
    frmActualizar.NumAsiento = NumeroAsiento
    frmActualizar.FechaAsiento = CDate(Text1(1).Text)
    frmActualizar.NumDiari = CInt(Text1(0).Text)
    AlgunAsientoActualizado = False
    frmActualizar.Show vbModal
    Me.Refresh
    If AlgunAsientoActualizado Then
        'Vamos a puntear las lineas
        'MODIFICACION MAYO 2005. NOOOO punteamos
        'SQL = "UPDATE hlinapu SET Punteada=1 WHERE"
        'SQL = SQL & " numasien = " & NumeroAsiento
        'SQL = SQL & " AND fechaent = '" & Format(CDate(Text1(1).Text), FormatoFecha)
        'SQL = SQL & "' AND NumDiari = " & CInt(Text1(0).Text)
        'Conn.Execute SQL
        ActualizarASientoHCO = True
    End If

    
End Function



'Si todo ha ido bien solo necsitamos asignar un nuevo numero de asiento e integrarlo
Private Function ActualizarASiento(ByRef NumeroAsiento As Long) As Boolean
Dim Mc As Contadores
    
    ActualizarASiento = False
    If CDate(Text1(1).Text) < vParam.fechaini Then Exit Function
    If CDate(Text1(1).Text) > DateAdd("yyyy", 1, vParam.fechafin) Then Exit Function
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", (CDate(Text1(1).Text) <= vParam.fechafin), False) = 0 Then
        SQL = " WHERE  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien & ";"
        
        'Guardamos numero asiento
        NumeroAsiento = Mc.Contador
        
        'Lineas
        Conn.Execute "UPDATE linapue set numasien= " & Mc.Contador & SQL
        
        'Cabeceras1
        Conn.Execute "UPDATE cabapue set numasien= " & Mc.Contador & SQL
        
        'Para el nuevo valor del SQL
        SQL = " WHERE  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Mc.Contador & ";"
        
        
        'Transaccionamos
        Conn.BeginTrans
        If InsertaEnDiario Then
            Conn.CommitTrans
            Conn.Execute "DELETE FROM linapue " & SQL
            Conn.Execute "DELETE FROM cabapue " & SQL
            ActualizarASiento = True
        Else
            Conn.RollbackTrans
            Mc.DevolverContador "0", (CDate(Text1(1).Text) <= vParam.fechafin), Mc.Contador
        End If
        
    End If
    Set Mc = Nothing
End Function



Private Function InsertaEnDiario() As Boolean
    On Error GoTo EInsertaEnDiario
        InsertaEnDiario = False
        
        Conn.Execute "INSERT INTO cabapu SELECT * from cabapue" & SQL
        
        Conn.Execute "INSERT INTO linapu SELECT * from linapue" & SQL
        InsertaEnDiario = True
        
        Exit Function
EInsertaEnDiario:
        MuestraError Err.Number, "InsertaEnDiario"
End Function



Private Sub PonerLineaAnterior(Indice As Integer)
Dim RT As ADODB.Recordset
Dim C As String
On Error GoTo EponerLineaAnterior

    'Si no estamos insertando,modificando lineas
    
    If Modo <> 5 Then Exit Sub
    

    If adodc1.Recordset.EOF Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    
    'Todos los casos menos la ampliacion del concepto
    If Indice <> 5 Then
        SQL = "SELECT "
        Select Case Indice
        Case 0
            C = "codmacta"
            I = 2
        Case 2
            C = "numdocum"
            I = 3
        Case 3
            C = "ctacontr"
            I = 4
        Case 4
            C = "codconce"
            I = 5
        Case 8
            C = "codccost"
            I = -1
        Case Else
            C = ""
        End Select
        If C <> "" Then
            SQL = SQL & C & "  FROM linapue"
            SQL = SQL & " WHERE numdiari=" & Data1.Recordset!NumDiari
            SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
            SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien
            If ModificandoLineas = 2 Then SQL = SQL & " AND linliapu <" & Linliapu
            SQL = SQL & " ORDER BY linliapu DESC"
            Set RT = New ADODB.Recordset
            RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            C = ""
            If Not RT.EOF Then C = DBLet(RT.Fields(0))
            
            'Lo ponemos en txtaux
            If C <> "" Then
                txtAux(Indice).Text = C
                If I >= 0 Then
                    PonerFoco txtAux(I)
                End If
            End If
            RT.Close
        End If





    Else
        SQL = "Select linliapu,ampconce,nomconce FROM linapue,conceptos"
        SQL = SQL & " WHERE conceptos.codconce=linapue.codconce AND  numdiari=" & Data1.Recordset!NumDiari
        SQL = SQL & " AND fechaent='" & Format(Data1.Recordset!fechaent, FormatoFecha)
        SQL = SQL & "' AND numasien=" & Data1.Recordset!Numasien
        If ModificandoLineas = 2 Then SQL = SQL & " AND linliapu <" & Linliapu
           
        SQL = SQL & " ORDER BY linliapu DESC"
        Set RT = New ADODB.Recordset
        RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        C = ""
        If Not RT.EOF Then
            SQL = DBLet(RT.Fields(1))
            C = DBLet(RT.Fields(2))
        End If
        
        'Lo ponemos en txtaux
        If SQL <> "" Then
            If C <> "" Then
                I = InStr(1, SQL, C)
                If I > 0 Then SQL = Trim(Mid(SQL, Len(C) + 1))
            End If
            txtAux(5).Text = txtAux(5).Text & SQL & " "
            txtAux(5).SelStart = Len(txtAux(5).Text)
            PonerFoco txtAux(6)
        End If
        RT.Close

    
    End If
    
EponerLineaAnterior:
    If Err.Number <> 0 Then Err.Clear
    Set RT = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub

