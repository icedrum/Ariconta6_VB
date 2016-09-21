VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAsiPre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Predefinidos"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   16605
   Icon            =   "frmAsiPre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   16605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   13470
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   50
      Top             =   180
      Width           =   3675
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   51
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
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   3
      Left            =   11220
      TabIndex        =   48
      Top             =   6240
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   2
      Left            =   7440
      TabIndex        =   47
      Top             =   6240
      Width           =   195
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   6030
      TabIndex        =   46
      Top             =   6240
      Width           =   195
   End
   Begin VB.Frame FrameToolAux 
      Height          =   555
      Left            =   120
      TabIndex        =   44
      Top             =   1680
      Width           =   1545
      Begin MSComctlLib.Toolbar ToolbarAux 
         Height          =   330
         Left            =   180
         TabIndex        =   45
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insertar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3930
      TabIndex        =   42
      Top             =   180
      Width           =   2505
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   43
         Top             =   180
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
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   34
      Top             =   6240
      Width           =   195
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
      Left            =   15330
      TabIndex        =   12
      Top             =   8670
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
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
      Height          =   320
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtaux 
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
      Height          =   320
      Index           =   1
      Left            =   1080
      TabIndex        =   33
      Top             =   6240
      Width           =   2235
   End
   Begin VB.TextBox txtaux 
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
      Height          =   320
      Index           =   2
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   4
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
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
      Height          =   320
      Index           =   3
      Left            =   4560
      TabIndex        =   5
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
      TabIndex        =   6
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txtaux 
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
      Height          =   320
      Index           =   5
      Left            =   6240
      MaxLength       =   30
      TabIndex        =   7
      Top             =   6240
      Width           =   2070
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
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
      Height          =   320
      Index           =   6
      Left            =   8340
      TabIndex        =   8
      Top             =   6240
      Width           =   1125
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
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
      Height          =   320
      Index           =   7
      Left            =   9480
      TabIndex        =   9
      Top             =   6240
      Width           =   945
   End
   Begin VB.TextBox txtaux 
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
      Height          =   320
      Index           =   8
      Left            =   10620
      MaxLength       =   4
      TabIndex        =   10
      Top             =   6240
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   915
      Left            =   10710
      TabIndex        =   19
      Top             =   1170
      Width           =   5625
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   3720
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   420
         Width           =   1665
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   1950
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   420
         Width           =   1665
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "SALDO"
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
         Index           =   4
         Left            =   3720
         TabIndex        =   25
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "HABER"
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
         Index           =   3
         Left            =   1950
         TabIndex        =   24
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "DEBE"
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
         Index           =   2
         Left            =   180
         TabIndex        =   23
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   3540
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
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Nombre asiento predefinido|T|N|||asipre|nomaspre|||"
      Text            =   "commor"
      Top             =   1260
      Width           =   6375
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
      Left            =   15330
      TabIndex        =   16
      Top             =   8670
      Visible         =   0   'False
      Width           =   1035
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
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Tag             =   "Nº asiento predefinido|N|N|||asipre|numaspre|0000|S|"
      Text            =   "Text1"
      Top             =   1260
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   13
      Top             =   8550
      Width           =   3495
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
         TabIndex        =   14
         Top             =   210
         Width           =   2955
      End
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
      Left            =   14130
      TabIndex        =   11
      Top             =   8670
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   0
      Top             =   3060
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAsiPre.frx":000C
      Height          =   5310
      Left            =   120
      TabIndex        =   18
      Top             =   2310
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   9366
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame framelineas 
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   7650
      Width           =   16245
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
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
         Index           =   2
         Left            =   9420
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   420
         Width           =   4965
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
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
         Left            =   4680
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   420
         Width           =   4635
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
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
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   420
         Width           =   4185
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   14880
         Top             =   300
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   10470
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   5850
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   2430
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "C. coste"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   9420
         TabIndex        =   32
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
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
         Index           =   1
         Left            =   4680
         TabIndex        =   29
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Cta. Contrapartida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   180
         Width           =   1875
      End
   End
   Begin VB.Frame frameextras 
      Height          =   855
      Left            =   120
      TabIndex        =   35
      Top             =   7650
      Width           =   16215
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nomctapar"
         DataSource      =   "Adodc1"
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
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text3"
         Top             =   420
         Width           =   4185
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "nombreconcepto"
         DataSource      =   "Adodc1"
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
         Index           =   4
         Left            =   4680
         TabIndex        =   37
         Text            =   "Text3"
         Top             =   420
         Width           =   4605
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         DataField       =   "centrocoste"
         DataSource      =   "Adodc1"
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
         Index           =   3
         Left            =   9420
         TabIndex        =   36
         Text            =   "Text3"
         Top             =   420
         Width           =   4965
      End
      Begin VB.Label Label2 
         Caption         =   "Cta. Contrapartida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   41
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
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
         Index           =   4
         Left            =   4680
         TabIndex        =   40
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "C. coste"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   9450
         TabIndex        =   39
         Top             =   180
         Width           =   1275
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   15930
      TabIndex        =   49
      Top             =   180
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
      Caption         =   "Nombre"
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
      Index           =   1
      Left            =   1380
      TabIndex        =   17
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Número"
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
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   990
      Width           =   1215
   End
   Begin VB.Menu mnOpcionesAsiPre 
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
Attribute VB_Name = "frmAsiPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 211

Private Const NO = "No encontrado"

Private WithEvents frmAsiP As frmBasico
Attribute frmAsiP.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCCentroCoste
Attribute frmCC.VB_VarHelpID = -1

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
Dim NumLin As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar



Dim CadB As String





Private Sub cmdAceptar_Click()
    Dim cad As String
    Dim I As Integer
    Dim Limp As Boolean
    
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
                    cmdCancelar.Caption = "Cabecera"
                    AnyadirLinea True
                End If
            End If
        End If
    Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    If SituarData1 Then PonerModo 2
                    lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
                End If
            End If
            
    Case 5
        cad = AuxOK
        If cad <> "" Then
            MsgBox cad, vbExclamation
        Else
            'Insertaremos, o modificaremos
            If InsertarModificar Then
                'Reestablecemos los campos
                'y ponemos el grid
                cmdAceptar.Visible = False
                DataGrid1.AllowAddNew = False
                CargaGrid data1.Recordset!numaspre
                Limp = True
                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    'Si ha puesto contrapartida borramos
                    If txtAux(3).Text <> "" Then
                        If LlevaContraPartida Then
                            'Ya lleva la contra partida, luego no hacemos na
                            LlevaContraPartida = False
                        Else
                            FijarContraPartida
                            Limp = False
                            LlevaContraPartida = True
                        End If
                    Else
                        LlevaContraPartida = False
                    End If
                    txtAux(8).Text = ""
                    Text3(2).Text = ""
                    If Limp Then
                        For I = 0 To 7
                            txtAux(I).Text = ""
                        Next I
                        Text3(0).Text = ""
                        Text3(1).Text = ""
                    End If
                    ModificandoLineas = 0
                    cmdAceptar.Visible = True
                    AnyadirLinea False
                    If Limp Then
                        txtAux(0).SetFocus
                    Else
                        txtAux(2).SetFocus
                    End If
                Else
                    ModificandoLineas = 0
                    CamposAux False, 0, False
                End If
                '++
                cmdCancelar_Click
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
    Select Case Index
        Case 0
            cmdAux(0).Tag = 0
            LlamaContraPar
        Case 1
            'Cta contrapartida
            cmdAux(0).Tag = 1
            LlamaContraPar
            txtAux(4).SetFocus
        Case 2
            Set frmCon = New frmConceptos
            frmCon.DatosADevolverBusqueda = "0|"
            frmCon.Show vbModal
            Set frmCon = Nothing
        Case 3
            If txtAux(8).Enabled Then
                Set frmCC = New frmCCCentroCoste
                frmCC.DatosADevolverBusqueda = "0|1|"
                frmCC.Show vbModal
                Set frmCC = Nothing
            End If
    End Select
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
    CamposAux False, 0, False
    LlevaContraPartida = False
    'Si esta insertando/modificando lineas haremos unas cosas u otras
    DataGrid1.Enabled = True
    If ModificandoLineas = 0 Then
        lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
        PonerModo 2
    Else
        If ModificandoLineas = 1 Then
             DataGrid1.AllowAddNew = False
             If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
             DataGrid1.Refresh
        End If
        frameextras.Visible = True
        framelineas.Visible = False
'--
'        cmdAceptar.Visible = False
'        cmdCancelar.Caption = "Cabeceras"
        ModificandoLineas = 0
        
    End If
'++
    cmdCancelar_Click

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
            CadenaConsulta = "Select * from " & NombreTabla
            data1.RecordSource = CadenaConsulta
            data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = " numaspre = " & Val(Text1(0).Text)
            data1.Recordset.Find SQL
            If data1.Recordset.EOF Then GoTo ESituarData1
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
    CargaGrid -1
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
        CargaGrid -1
        
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonFoco Text1(0)
        Else
            HacerBusqueda
            If data1.Recordset.EOF Then
                Text1(kCampo).Text = ""
                PonFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid -1
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 1
        data1.Recordset.MoveFirst
    Case 2
        data1.Recordset.MovePrevious
        If data1.Recordset.BOF Then data1.Recordset.MoveFirst

    Case 3
        data1.Recordset.MoveNext
        If data1.Recordset.EOF Then data1.Recordset.MoveLast

    Case 4
        data1.Recordset.MoveLast
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
    Text1(0).Locked = True
    Text1(0).BackColor = &H80000018
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
    Dim cad As String
    Dim I As Integer
    'Ciertas comprobaciones
    If data1.Recordset.EOF Then Exit Sub
    '### a mano
    cad = "Seguro que desea eliminar de la BD el asiento predefinido:"
    cad = cad & vbCrLf & "Nº Asiento: " & data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Descrpcion: " & data1.Recordset.Fields(1)
    I = MsgBox(cad, vbQuestion + vbYesNo)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        'Eliminar cabeceras
        cad = "Delete from asipre_lineas where numaspre = " & data1.Recordset!numaspre
        Conn.Execute cad
        
        'Borramos sus lineas
        cad = "Delete from asipre where numaspre = " & data1.Recordset!numaspre
        NumRegElim = data1.Recordset.AbsolutePosition
        Conn.Execute cad

        data1.Refresh
        If data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            PonerModo 0
            Else
                data1.Recordset.MoveFirst
                NumRegElim = NumRegElim - 1
                If NumRegElim > 1 Then
                    For I = 1 To NumRegElim - 1
                        data1.Recordset.MoveNext
                    Next I
                End If
                PonerCampos
        End If

    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdRegresar_Click()

If data1.Recordset.EOF Then
    MsgBox "Ningún registro devuelto.", vbExclamation
    Exit Sub
End If

RaiseEvent DatoSeleccionado(data1.Recordset.Fields(0) & "|" & data1.Recordset.Fields(1) & "|")
Unload Me
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
    
   
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
   
    With Me.ToolbarAux
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    For I = 0 To 2
        Image1(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
    End If
    
    'Los campos auxiliares
    CamposAux False, 0, True
    
    'Si no es analitica no mostramos el label
    Text3(2).Visible = vParam.autocoste
    Label2(2).Visible = vParam.autocoste
    
    '## A mano
    NombreTabla = "asipre"
    Ordenacion = " ORDER BY numaspre"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    data1.ConnectionString = Conn
    
    PonerOpcionesMenu
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    'Maxima longitud
    txtAux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    txtAux(3).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
    data1.CursorType = adOpenDynamic
    data1.LockType = adLockPessimistic
    data1.RecordSource = "Select * from " & NombreTabla & " WHERE numaspre = -1"
    data1.Refresh
    CargaGrid -1
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If

    
    CadAncho = False

End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmAsiP_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    CadB = "numaspre = " & RecuperaValor(CadenaSeleccion, 1)
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

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
'Concepto
txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1)
Text3(1).Text = RecuperaValor(CadenaSeleccion, 2) & " "
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
        Set frmCC = New frmCCCentroCoste
        frmCC.DatosADevolverBusqueda = "0|1|"
        frmCC.Show vbModal
        Set frmCC = Nothing
    End If
End Select
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


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
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
    ''Quitamos blancos por los lados
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    If Modo <> 1 Then _
        FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor

End Sub

Private Sub HacerBusqueda()
Dim cad As String
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
Dim cWhere As String
    
    Set frmAsiP = New frmBasico
    
    AyudaAsientosP frmAsiP
    
    Set frmAsiP = Nothing

End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    data1.RecordSource = CadenaConsulta
    data1.Refresh
    If data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Else
            PonerModo 2
            data1.Recordset.MoveFirst
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
    If data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma Me, data1
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid data1.Recordset!numaspre
    DataGrid1.Enabled = True
    
    frameextras.Visible = Not Adodc1.Recordset.EOF
    lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
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
        For I = 0 To Text1.Count - 1
            Text1(I).BackColor = vbWhite
        Next I
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    B = (Modo = 0 Or Modo = 2)
    
    
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    frameextras.Visible = B
    If B Then framelineas.Visible = False
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B And Me.data1.Recordset.RecordCount > 1
    If Not B Then frameextras.Visible = False
        
    
    DataGrid1.Enabled = B Or (Modo = 5)
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    cmdAceptar.Visible = B Or Modo = 1

    '
    B = B Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    mnOpcionesAsiPre.Enabled = Not B
   
   
    'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5
    'Modificar
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    Else
        cmdRegresar.Visible = False
    End If
    B = B Or (Modo = 5)
    
    ToolbarAux.Buttons(2).Enabled = B
    'eliminar
    ToolbarAux.Buttons(3).Enabled = B

   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    B = B Or Modo = 0   'En B tenemos modo=2 o a 5
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = B
        If Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    
    B = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = B
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    PonerOpcionesMenu
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.data1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Function DatosOK() As Boolean
    Dim RS As ADODB.Recordset
    Dim B As Boolean
    B = CompForm(Me)
    DatosOK = B
End Function


'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()
    Dim SQL As String
    Dim RS As ADODB.Recordset
    
    SQL = "Select Max(numaspre) from " & NombreTabla
    Text1(0).Text = 1
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            Text1(0).Text = RS.Fields(0) + 1
        End If
    End If
    RS.Close
End Sub

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

                frmAsiPreList.Show vbModal

        Case Else
        
    End Select


End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub CargaGrid(NumFac As Long)
Dim B As Boolean
B = DataGrid1.Enabled
CargaGrid2 NumFac
DataGrid1.Enabled = B
End Sub


Private Sub CargaGrid2(NumFac As Long)
    Dim anc As Single
    
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = MontaSQLCarga(NumFac)
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockPessimistic
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 350 '320
    
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
    DataGrid1.Columns(2).Width = 1405
    
    DataGrid1.Columns(3).Caption = "Denominación"
    DataGrid1.Columns(3).Width = 3995


    DataGrid1.Columns(4).Caption = "Documento"
    DataGrid1.Columns(4).Width = 1405

    DataGrid1.Columns(5).Caption = "Contrapart."
    DataGrid1.Columns(5).Width = 1405
    
    DataGrid1.Columns(6).Caption = "Cto."
    DataGrid1.Columns(6).Width = 465
    
    DataGrid1.Columns(7).Visible = False
    

        
    DataGrid1.Columns(8).Caption = "Ampliación"
    DataGrid1.Columns(8).Width = 3000

    'Cuenta contrapartida
    DataGrid1.Columns(9).Visible = False
    
    If vParam.autocoste Then
        ancho = 0
    Else
        ancho = 355 'Es la columna del centro de coste divida entre dos
    End If
    
    DataGrid1.Columns(10).Caption = "Debe"
    DataGrid1.Columns(10).NumberFormat = "#,##0.00"
    DataGrid1.Columns(10).Width = 1654 + ancho
    DataGrid1.Columns(10).Alignment = dbgRight
            
    DataGrid1.Columns(11).Caption = "Haber"
    DataGrid1.Columns(11).NumberFormat = "#,##0.00"
    DataGrid1.Columns(11).Width = 1654 + ancho
    DataGrid1.Columns(11).Alignment = dbgRight
            
            
    If vParam.autocoste Then
        DataGrid1.Columns(12).Caption = "C.C."
        DataGrid1.Columns(12).Width = 710
    Else
        DataGrid1.Columns(12).Visible = False
    End If
    DataGrid1.Columns(13).Visible = False
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        anc = 323
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(2).Width
        
        
        anc = 150
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(3).Left + 90
        
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 6
        txtAux(1).Width = DataGrid1.Columns(3).Width - 180
        
    
        txtAux(2).Left = DataGrid1.Columns(4).Left + anc
        txtAux(2).Width = DataGrid1.Columns(4).Width - 180
        
        
        txtAux(3).Left = DataGrid1.Columns(5).Left + anc
        txtAux(3).Width = DataGrid1.Columns(5).Width - 30
        
        
        
        'Concepto
        cmdAux(1).Left = DataGrid1.Columns(6).Left + 90
        
        txtAux(4).Left = cmdAux(1).Left + cmdAux(1).Width + 6
        txtAux(4).Width = DataGrid1.Columns(6).Width - 180
        
        cmdAux(2).Left = DataGrid1.Columns(8).Left + 90
        
        
        txtAux(5).Left = cmdAux(2).Left + cmdAux(2).Width + 6
        txtAux(5).Width = DataGrid1.Columns(8).Width - 180
        
        
        
        txtAux(6).Left = DataGrid1.Columns(10).Left + anc
        txtAux(6).Width = DataGrid1.Columns(10).Width - 30
        
        
        txtAux(7).Left = DataGrid1.Columns(11).Left + anc
        txtAux(7).Width = DataGrid1.Columns(11).Width - 30
        
        cmdAux(3).Left = DataGrid1.Columns(12).Left + 90
        
        txtAux(8).Left = cmdAux(3).Left + cmdAux(2).Width + 6
        txtAux(8).Width = DataGrid1.Columns(12).Width - 180
      
        CadAncho = True
    End If
    
    
    
    
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
    
    
    
    'Obtenemos las sumas
    ObtenerSumas
    
    PonerModoUsuarioGnral Modo, "ariconta"
     
    
End Sub

Private Sub ObtenerSumas()
Dim Deb As Currency
Dim hab As Currency
Dim RS As ADODB.Recordset
If data1.Recordset.EOF Then
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    Exit Sub
End If

If Adodc1.Recordset.EOF Then
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    Exit Sub
End If



Set RS = New ADODB.Recordset
SQL = "SELECT Sum(asipre_lineas.timporteD) AS SumaDetimporteD, Sum(asipre_lineas.timporteH) AS SumaDetimporteH,asipre_lineas.numaspre"
SQL = SQL & " From asipre_lineas"
SQL = SQL & " GROUP BY asipre_lineas.numaspre"
SQL = SQL & " HAVING (((asipre_lineas.numaspre)=" & data1.Recordset!numaspre & "));"
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
Text2(2).Text = Format(Deb, FormatoImporte)

End Sub


Private Function MontaSQLCarga(vNumFac As Long) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    '--------------------------------------------------------------------
    Dim SQL As String

    SQL = "SELECT asipre_lineas.numaspre,asipre_lineas.linlapre, asipre_lineas.codmacta, cuentas.nommacta, asipre_lineas.numdocum,"
    SQL = SQL & " asipre_lineas.ctacontr, asipre_lineas.codconce, conceptos.nomconce as nombreconcepto, asipre_lineas.ampconce,"
    SQL = SQL & " cuentas_1.nommacta as nomctapar, asipre_lineas.timporteD, asipre_lineas.timporteH, asipre_lineas.codccost, ccoste.nomccost as centrocoste"
    SQL = SQL & " FROM (((asipre_lineas INNER JOIN conceptos ON asipre_lineas.codconce = conceptos.codconce)"
    SQL = SQL & " INNER JOIN cuentas ON asipre_lineas.codmacta = cuentas.codmacta)"
    SQL = SQL & " LEFT JOIN cuentas AS cuentas_1 ON asipre_lineas.ctacontr = cuentas_1.codmacta)"
    SQL = SQL & " LEFT JOIN ccoste ON asipre_lineas.codccost = ccoste.codccost"
    SQL = SQL & " WHERE numaspre = " & vNumFac
    SQL = SQL & " ORDER BY asipre_lineas.linlapre"

    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea(Limpiar As Boolean)
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    NumLin = ObtenerSigueinteNumeroLinea
    'Situamos el grid al final
    
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 270 '220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If
    LLamaLineas anc, 1, Limpiar
    HabilitarImportes 0
    'Ponemos el foco
    txtAux(0).SetFocus
    
End Sub

Private Sub ModificarLinea()
Dim cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    NumLin = Adodc1.Recordset!linlapre
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 270 '220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    txtAux(0).Text = Adodc1.Recordset.Fields!codmacta
    txtAux(1).Text = Adodc1.Recordset.Fields!Nommacta
    txtAux(2).Text = DataGrid1.Columns(4).Text
    txtAux(3).Text = DataGrid1.Columns(5).Text
    txtAux(4).Text = DataGrid1.Columns(6).Text
    txtAux(5).Text = DataGrid1.Columns(8).Text
    cad = DBLet(Adodc1.Recordset.Fields!timported)
    If cad <> "" Then
        txtAux(6).Text = Format(cad, "0.00")
    Else
        txtAux(6).Text = cad
    End If
    cad = DBLet(Adodc1.Recordset.Fields!timporteH)
    If cad <> "" Then
        txtAux(7).Text = Format(cad, "0.00")
    Else
        txtAux(7).Text = cad
    End If
    txtAux(8).Text = DBLet(Adodc1.Recordset.Fields!codccost)
    HabilitarImportes 3
    HabilitarCentroCoste
    Text3(0).Text = Text3(5).Text
    Text3(1).Text = Text3(4).Text
    Text3(2).Text = Text3(3).Text
    LLamaLineas anc, 2, False
 
End Sub

Private Sub EliminarLineaFactura()
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Seguro que desea eliminar la linea: " & Adodc1.Recordset.Fields(3) & " "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = "Delete from asipre_lineas WHERE numaspre =" & data1.Recordset!numaspre
        SQL = SQL & " AND linlapre = " & Adodc1.Recordset!linlapre
        Conn.Execute SQL
        CargaGrid data1.Recordset!numaspre
    End If
    '++
    cmdCancelar_Click
End Sub



Private Function ObtenerSigueinteNumeroLinea() As Long
Dim RS As ADODB.Recordset
Dim I As Long

    Set RS = New ADODB.Recordset
    RS.Open "SELECT Max(linlapre) FROM asipre_lineas where numaspre =" & Text1(0).Text, Conn, adOpenDynamic, adLockOptimistic, adCmdText
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
Dim Ch As String
    If Not vParam.autocoste Then Exit Sub
    hab = False
    If txtAux(0).Text <> "" Then
            Ch = Mid(txtAux(0).Text, 1, 1)
            If Ch = vParam.grupogto Or Ch = vParam.grupovta Or Ch = vParam.grupoord Then hab = True
    Else
        txtAux(8).Text = ""
    End If
    If hab Then
        txtAux(8).BackColor = &H80000005
        Else
        txtAux(8).BackColor = &H80000018
    End If
    txtAux(8).Enabled = hab
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
Dim B As Boolean
    DeseleccionaGrid DataGrid1
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)
    framelineas.Visible = Not B
    'frameextras.Visible = b
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
    
        cmdAux(1).Visible = Visible
        cmdAux(1).Top = Altura
        cmdAux(2).Visible = Visible
        cmdAux(2).Top = Altura
        If vParam.autocoste Then
            cmdAux(3).Visible = Visible
            cmdAux(3).Top = Altura
        Else
            cmdAux(3).Visible = False
        End If
    
    If Limpiar Then
        For I = 0 To J
            txtAux(I).Text = ""
        Next I
    End If
End Sub


Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)

    PonerModo 5
    'Fuerzo que se vean las lineas
    frameextras.Visible = True
    
    Select Case Button.Index
        Case 1
            'AÑADIR linea factura
            AnyadirLinea True
        Case 2
            'MODIFICAR linea factura
            ModificarLinea
        Case 3
            'ELIMINAR linea factura
            EliminarLineaFactura
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYBusqueda KeyAscii, 0
            Case 3:  KEYBusqueda1 KeyAscii, 0
            Case 4:  KEYBusqueda1 KeyAscii, 1
            Case 8:  KEYBusqueda1 KeyAscii, 2
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

Private Sub KEYBusqueda1(KeyAscii As Integer, Indice As Integer)
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
Dim Sng As Single

    'Si no estamos modificando o insertando lineas no hacemos na de na
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
            RC = "tipoconce"
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtAux(4).Text, "N", RC)
            If SQL = "" Then
                MsgBox "Concepto NO encontrado: " & txtAux(4).Text, vbExclamation
                txtAux(4).Text = ""
                RC = "0"
            Else
                SQL = SQL & " "
            End If
            HabilitarImportes CByte(Val(RC))
            Text3(1).Text = SQL
            txtAux(5).Text = SQL
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
            Sng = Round(CSng(SQL), 2)
            txtAux(Index).Text = Format(Sng, "0.00")
            
            'Ponemos el otro campo a ""
            If Index = 6 Then
                txtAux(7).Text = ""
            Else
                txtAux(6).Text = ""
            End If
    Case 8
            RC = "idsubcos"
            SQL = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtAux(8).Text, "T", RC)
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
    SQL = "INSERT INTO asipre_lineas (numaspre, linlapre, codmacta, numdocum, codconce,"
    SQL = SQL & "ampconce, timporteD, timporteH, codccost, ctacontr, idcontab) VALUES ("
    SQL = SQL & data1.Recordset.Fields(0) & ","
    SQL = SQL & NumLin & ",'"
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
      SQL = SQL & txtAux(3).Text & ","
    End If
    'Marca de entrada manual de datos
    SQL = SQL & "'contab')"
    
    
Else

    SQL = "UPDATE asipre_lineas SET "
    
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
    SQL = SQL & " Where numaspre = " & data1.Recordset.Fields(0)
    SQL = SQL & " And linlapre = " & NumLin


End If
Conn.Execute SQL
InsertarModificar = True
Exit Function
EInsertarModificar:
    MuestraError Err.Number, "InsertarModificar linea asiento predefinido.", Err.Description
End Function
 

Private Sub LlamaContraPar()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.ConfigurarBalances = 3
    frmC.Show vbModal
    Set frmC = Nothing
End Sub


Private Sub FijarContraPartida()
Dim cad As String
'Hay contrapartida
'Reasignamos campos de cuentas
cad = txtAux(0).Text
txtAux(0).Text = txtAux(3).Text
txtAux(3).Text = cad
HabilitarCentroCoste
cad = txtAux(1).Text
txtAux(1).Text = Text3(0).Text
Text3(0).Text = cad

'Los importes
HabilitarImportes 3
cad = txtAux(6).Text
txtAux(6).Text = txtAux(7).Text
txtAux(7).Text = cad
End Sub


Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub


' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub


'**************************************************************************
'**************************************************************************
'**************************************************************************

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim RS As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2 And Me.data1.Recordset.RecordCount > 0)
        Toolbar1.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2 And Me.data1.Recordset.RecordCount > 0)
        
        Toolbar1.Buttons(5).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N")
        
       
        ToolbarAux.Buttons(1).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2)
        ToolbarAux.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2 And Me.Adodc1.Recordset.RecordCount > 0)
        ToolbarAux.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2 And Me.Adodc1.Recordset.RecordCount > 0)
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub





