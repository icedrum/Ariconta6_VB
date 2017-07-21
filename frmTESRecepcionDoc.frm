VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTESRecepcionDoc 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10035
   Icon            =   "frmTESRecepcionDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
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
      Left            =   7590
      TabIndex        =   42
      Top             =   330
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   210
      TabIndex        =   38
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   39
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
      Left            =   4920
      TabIndex        =   36
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   37
         Top             =   210
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3870
      TabIndex        =   34
      Top             =   90
      Width           =   975
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   35
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Contabilizar Recepción"
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtSuma 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEF7E4&
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
      Left            =   8130
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   2550
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Llevado banco"
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
      Left            =   7890
      TabIndex        =   31
      Tag             =   "C|N|N|||talones|LlevadoBanco|||"
      Top             =   1260
      Width           =   1785
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Contabilizado"
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
      Left            =   5820
      TabIndex        =   4
      Tag             =   "C|N|N|||talones|Contabilizada|||"
      Top             =   1260
      Width           =   1905
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
      Index           =   6
      Left            =   4290
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Fecha vencimiento|F|N|||talones|fechavto|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   1200
      Width           =   1485
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   6210
      TabIndex        =   8
      Tag             =   "Importe|N|N|0||talones|importe|#,##0.00||"
      Text            =   "Text1"
      Top             =   2550
      Width           =   1515
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
      Index           =   3
      Left            =   4290
      TabIndex        =   6
      Tag             =   "Banco|T|N|||talones|banco|||"
      Text            =   "Text1"
      Top             =   1860
      Width           =   5415
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmTESRecepcionDoc.frx":000C
      Left            =   1140
      List            =   "frmTESRecepcionDoc.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Talon|N|N|0||talones|talon|||"
      Top             =   1200
      Width           =   1545
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FEF7E4&
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
      Left            =   240
      TabIndex        =   0
      Tag             =   "Codigo|N|S|0||talones|codigo||S|"
      Text            =   "Text1"
      Top             =   1200
      Width           =   795
   End
   Begin VB.TextBox Text5 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text4"
      Top             =   2520
      Width           =   4185
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
      Index           =   2
      Left            =   240
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Cliente|T|N|||talones|codmacta|||"
      Text            =   "commor"
      Top             =   2520
      Width           =   1365
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
      Left            =   2730
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Fecha recepcion|F|N|||talones|fecharec|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   1200
      Width           =   1425
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   16
      Top             =   6930
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
      Left            =   8730
      TabIndex        =   15
      Top             =   8130
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
      Height          =   345
      Index           =   0
      Left            =   660
      MaxLength       =   3
      TabIndex        =   9
      Top             =   6930
      Width           =   975
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
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   6930
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
      Height          =   345
      Index           =   2
      Left            =   4020
      MaxLength       =   10
      TabIndex        =   11
      Top             =   6930
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
      Height          =   345
      Index           =   3
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   12
      Top             =   6930
      Width           =   885
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
      Height          =   345
      Index           =   4
      Left            =   6240
      TabIndex        =   13
      Top             =   6930
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6240
      Top             =   480
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
      Left            =   8730
      TabIndex        =   20
      Top             =   8130
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
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Tag             =   "Referencia|T|N|||talones|numeroref|||"
      Text            =   "Text1"
      Top             =   1860
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   7920
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
         Height          =   315
         Left            =   240
         TabIndex        =   18
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
      Left            =   7530
      TabIndex        =   14
      Top             =   8130
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTESRecepcionDoc.frx":0029
      Height          =   4455
      Left            =   240
      TabIndex        =   21
      Top             =   3390
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   7858
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   495
      Left            =   4680
      Top             =   0
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
   Begin MSComctlLib.Toolbar ToolbarAux 
      Height          =   330
      Left            =   240
      TabIndex        =   40
      Top             =   2970
      Width           =   1305
      _ExtentX        =   2302
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   9390
      TabIndex        =   41
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
   Begin VB.Label lblSuma 
      Caption         =   "Suma"
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
      Left            =   9090
      TabIndex        =   33
      Top             =   2310
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Vto."
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
      Left            =   4320
      TabIndex        =   30
      Top             =   930
      Width           =   1155
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   1
      Left            =   5520
      Picture         =   "frmTESRecepcionDoc.frx":003E
      Top             =   930
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
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
      Left            =   6240
      TabIndex        =   29
      Top             =   2310
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Banco"
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
      Index           =   2
      Left            =   4290
      TabIndex        =   28
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
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
      Left            =   1170
      TabIndex        =   27
      Top             =   900
      Width           =   975
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   2
      Left            =   1380
      Picture         =   "frmTESRecepcionDoc.frx":00C9
      Top             =   2280
      Width           =   240
   End
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   0
      Left            =   3960
      Picture         =   "frmTESRecepcionDoc.frx":0ACB
      Top             =   930
      Width           =   240
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
      Height          =   225
      Index           =   9
      Left            =   1710
      TabIndex        =   26
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
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
      Index           =   8
      Left            =   240
      TabIndex        =   25
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "F.Recepción"
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
      Index           =   5
      Left            =   2730
      TabIndex        =   22
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Referencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Menu mnOpcionesAsiPre 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^F
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
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "Lineas"
         Shortcut        =   ^L
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
Attribute VB_Name = "frmTESRecepcionDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const IdPrograma = 611



Private Const NO = "No encontrado"
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private frmMens As frmMensajes
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
Dim i As Integer
Dim Ancho As Integer
'Dim colMes As Integer

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

Dim PrimeraVez As Boolean
Dim ImporteVto As Currency
Dim PosicionGrid As Integer


Dim ModoLineas As Byte

Private BuscaChekc As String


Private Sub PonerLineaModificadaSeleccionada()
    On Error GoTo E1


    Exit Sub
E1:
    Err.Clear
End Sub



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Select Case Modo
    Case 1
        HacerBusqueda
    Case 3
        If Not DatosOK Then Exit Sub
        
        If InsertarRegistro Then
            NuevoTalonPagareDefecto False   'Para que memorize la ultima opcion
            If SituarData1(True) Then
                PonerModo 2
            End If
        End If
        
    Case 4
        If Not DatosOK Then Exit Sub
        
        If ModificaDesdeFormulario(Me) Then
            'Ha cambiado fecha vto
            CambiaFechaVto
            If SituarData1(False) Then
                lblIndicador.Caption = ""
                PonerModo 2
            Else
                PonerModo 0
            End If
        End If
      
    End Select
End Sub

Private Sub cmdAux_Click(Index As Integer)
Dim Im As Currency

        Im = 0
        If Me.txtSuma.Text <> "" Then Im = ImporteFormateado(txtSuma.Text)
        Im = ImporteFormateado(Text1(5).Text) - Im
        

        CadenaDesdeOtroForm = ""
        'Todos los cobros pendientes de este
        SQL = " cobros.codmacta = '" & Text1(2).Text & "' AND ( impcobro =0 or impcobro is null)"
        
        'MODIFICADO Agosto 2009
        SQL = " cobros.codmacta = '" & Text1(2).Text & "' AND  ( tiporem is null or tiporem>1)"
        
        'Docu recibido NO
        SQL = SQL & " AND recedocu = 0" 'por si acoaso
        
        SQL = SQL & " and formapago.tipforpa in ( " & vbTalon & "," & vbPagare & ")"
        
        '###FALTA
        
        Set frmMens = New frmMensajes
        
        frmMens.Parametros = SQL
        frmMens.Opcion = 54
        frmMens.Importe = Text1(5)
        frmMens.Banco = Text1(3).Text
        frmMens.Referencia = Text1(0).Text
        frmMens.FecCobro = Text1(1).Text
        frmMens.FecVenci = Text1(6).Text
        frmMens.Tipo = Combo1.ListIndex
        frmMens.Codigo = Text1(4).Text
        
        frmMens.Label33.Caption = Text5.Text
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        CargaGrid True
        
        
        If CadenaDesdeOtroForm <> "" Then
            SQL = " and (numserie, numfactu, fecfactu, numorden) in (" & CadenaDesdeOtroForm & ")"
        End If
        
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
        
    End Select
End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1(Insertar As Boolean) As Boolean
    Dim SQL As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        SQL = "Select * from talones WHERE codigo =" & Text1(4).Text
        data1.RecordSource = SQL
    'End If
    
    data1.Refresh
    With data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not data1.Recordset.EOF
            If CStr(data1.Recordset!Codigo) = Text1(4).Text Then
                lblIndicador.Caption = ""
                SituarData1 = True
                Exit Function
                
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
   ' CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
   ' PonerCadenaBusqueda True
    
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3

    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    

    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    NuevoTalonPagareDefecto True
    'Combo1.ListIndex = 1 'Talon
    
    
    PonerFoco Combo1
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
        PonerFoco Text1(4)
        Else
            HacerBusqueda
            If data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                PonerFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda False
    End If
End Sub


Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
 
    PonerModo 4


    'Si tienen NO tiene lineas dejaremos modificar la cuenta contable
    If adodc1.Recordset.EOF Then Text1(2).Enabled = True
    
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    PonerFoco Text1(0)
End Sub

Private Sub BotonEliminar()
Dim Importe As Currency
Dim Ok As Boolean
    If data1.Recordset.EOF Then Exit Sub
    
    
    SQL = DevuelveDesdeBD("Contabilizada", "talones", "codigo", Text1(4).Text)
    If SQL = "1" Then
        'Esta realizado el apunte. Hay que deshacer
        SQL = DevuelveDesdeBD("sum(importe)", "talones_facturas", "codigo", Text1(4).Text)
        If SQL = "" Then SQL = "0"
        i = 0
        If CCur(SQL) <> ImporteFormateado(Text1(5).Text) Then
            SQL = CStr(CCur(SQL) - ImporteFormateado(Text1(5).Text))
            If CCur(SQL) > 0 Then
                i = -1   'Mayor las lineas que el importe del talon
            Else
                i = 1   'Mayor el total que la suma de las lineas
            End If
        End If
        
        If Not HacerDES_Contabilizacion_(i) Then Exit Sub
    Else
        'NO ha hecho nada, se borra directamente
        If MsgBox("¿Desea eliminar el documento recibido?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    Conn.BeginTrans
    NumRegElim = data1.Recordset.AbsolutePosition
    
        i = 0
        If Not adodc1.Recordset Is Nothing Then
            If Not adodc1.Recordset.EOF Then
               adodc1.Recordset.MoveFirst
               While Not adodc1.Recordset.EOF
               
                   'Obtengo el importe del vto
                   SQL = MontaSQLDelVto(False)
                   SQL = SQL & " AND 1 " 'Para hacer un truqiot
                   SQL = DevuelveDesdeBD("impcobro", "cobros", SQL, "1")
                   If SQL = "" Then SQL = "0"
                   Importe = CCur(SQL)
                   If Importe <> adodc1.Recordset!Importe Then
                       'TODO EL IMPORTE estaba en la linea. Fecultco a NULL
                       i = 1
                       Importe = Importe - adodc1.Recordset!Importe
                   Else
                       i = 0
                   End If
               
                   SQL = "UPDATE cobros SET recedocu=0"
                   If i = 0 Then
                       SQL = SQL & ", impcobro = NULL, fecultco = NULL"
                   Else
                       SQL = SQL & ", impcobro = " & TransformaComasPuntos(CStr(Importe))  'NO somos capace sde ver cual fue la utlima fecha de amortizacion
                   End If
                   SQL = SQL & ", observa= NULL"
                   SQL = SQL & " WHERE " & MontaSQLDelVto(False)
                   
                   If Not EjecutarSQL(SQL) Then
                       MsgBox "Error actualizadno cobros", vbExclamation
                       i = 100
                       adodc1.Recordset.MoveLast
                   End If
                   
                   adodc1.Recordset.MoveNext
               Wend
            End If
        End If
        If i = 100 Then
            Ok = False
        Else
            Ok = Eliminar
        End If
    If Ok Then
        Conn.CommitTrans
        
        data1.Refresh
        If data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            CargaGrid False
            PonerModo 0
            
        Else
            data1.Recordset.MoveFirst
            NumRegElim = NumRegElim - 1
            If NumRegElim > 1 Then
                For i = 1 To NumRegElim - 1
                    data1.Recordset.MoveNext
                Next i
            End If
            PonerCampos
        End If
    Else
        'Conn.RollbackTrans
        TirarAtrasTransaccion
    End If

    
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
Dim B As Boolean
  
    If PrimeraVez Then
        B = False
        PrimeraVez = False
        PonerModo 0
        CargaGrid False
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    LimpiarCampos
    PrimeraVez = True
    CadAncho = False


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
        .Buttons(1).Image = 42
    End With

    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
   
    With Me.ToolbarAux
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
    
    Caption = "Recepcion de documentos TALON,PAGARE (" & vEmpresa.nomresum & ")"
    'Los campos auxiliares
    CamposAux False, 0, True
    


    '## A mano
    NombreTabla = "talones"
    Ordenacion = " ORDER BY codigo"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    data1.ConnectionString = Conn

    PonerModoUsuarioGnral 0, "ariconta"


End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    Me.Combo1.ListIndex = -1
    Me.Check1.Value = 0
    Me.Check2.Value = 0
    txtSuma.Text = ""
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
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

        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda False
        Screen.MousePointer = vbDefault
    End If

End Sub









Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text5.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(i).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgppal_Click(Index As Integer)
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub

    Select Case Index
    Case 0, 1
        'FECHA
        If Index = 0 Then
            i = 1
        Else
            i = 6
        End If
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(i).Text <> "" Then frmF.Fecha = CDate(Text1(i).Text)
        frmF.Show vbModal
        Set frmF = Nothing
    
    Case 2
    
       ' If Text1(2).Enabled Then   'Solo insertando
            Set frmCCtas = New frmColCtas
            SQL = ""
            frmCCtas.DatosADevolverBusqueda = "0"
            frmCCtas.Show vbModal
            Set frmCCtas = Nothing
            If SQL <> "" Then
                Text1(2) = RecuperaValor(SQL, 1)
                Text5.Text = RecuperaValor(SQL, 2)
            End If
       ' End If
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    'BotonEliminar False
    HacerToolBar 8
End Sub

Private Sub mnLineas_Click()
Dim B As Button
    Set B = Toolbar1.Buttons(10)
    Toolbar1_ButtonClick B
    Set B = Nothing
End Sub

Private Sub mnModificar_Click()
    'BotonModificar
    HacerToolBar 7
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    'Condiciones para NO salir
    If Modo = 5 Then Exit Sub
    

    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 1: KEYPPal KeyAscii, 0
            Case 6: KEYPPal KeyAscii, 1
            Case 2: KEYPPal KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYPPal(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgppal_Click (Indice)
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
Dim RC As Byte
Dim EntrarEnSelect As Boolean
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbYellow Then
        Text1(Index).BackColor = vbWhite  '&H80000018
    End If
    
    'Si estamos insertando o modificando o buscando
    EntrarEnSelect = False
    If Modo = 3 Or Modo = 4 Then
        EntrarEnSelect = True
    Else
        If Modo = 1 And Index = 2 Then EntrarEnSelect = True
    End If
    If EntrarEnSelect Then
        If Text1(Index).Text = "" Then
            If Index = 0 Then
               
            Else
                If Index = 2 Then Text5.Text = ""
            End If
            Exit Sub
        End If
        Select Case Index
        Case 0
        
            
        Case 1, 6

            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
        Case 2
        
            RC = CByte(CuentaCorrectaUltimoNivelTXT(Text1(2), Text5))
            If RC = 0 Then
                'Error. En busqueda dejamos pasar
                
                If Modo <> 1 Then
                    MsgBox Text5.Text, vbExclamation
                    Text1(2).Text = ""
                    PonerFoco Text1(2)
                End If
                Text5.Text = ""
            End If
        Case 5
            PonerFormatoDecimal Text1(Index), 3
            If Text1(Index).Text = "" Then PonerFoco Text1(Index)
            
        End Select
    End If
End Sub



Private Sub MandaBusquedaPrevia(CadB As String)

   frmTESRecepcionDocPrev.Show vbModal

End Sub

Private Sub PonerCadenaBusqueda(Insertando As Boolean)
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    data1.RecordSource = CadenaConsulta
    data1.Refresh
    If Insertando Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla de recepcion de documentos", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
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
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True
    'Cargamos datos extras

    If Text1(2).Text = "" Then
        SQL = ""
    Else
        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(2).Text, "N")
        If SQL = "" Then SQL = "Error en cuenta contable"
    End If
    Text5.Text = SQL
    PonerImporteLinea
    If Modo = 2 Then lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean


    If Modo = 1 Then
        Text1(4).BackColor = &HFEF7E4
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    
    
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    If Not data1.Recordset Is Nothing Then
        DespalzamientoVisible B And (data1.Recordset.RecordCount > 1)
    End If
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    
    
    
    
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    B = (Modo < 5)
    chkVistaPrevia.visible = B

    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    If Not data1.Recordset Is Nothing Then
        DespalzamientoVisible B And (data1.Recordset.RecordCount > 1)
    End If
        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    cmdAceptar.visible = B Or Modo = 1
    'PRueba###
    
    '
    B = B Or (Modo = 5)
   
    'MODIFICAR Y ELIMINAR DISPONIBLES TB CUANDO EL MODO ES 5

    Text1(4).Enabled = (Modo = 1)
    Text1(2).Enabled = (Modo = 3 Or Modo = 1) 'Solo insertar
    B = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    
    Me.Check1.Enabled = (Modo = 1) Or (B And vUsu.Nivel = 0)
    Me.Check2.Enabled = Me.Check1.Enabled
    Combo1.Enabled = B
    
    Text1(0).Enabled = B
    Text1(1).Enabled = B
    Text1(3).Enabled = B
    Text1(5).Enabled = B
    Text1(6).Enabled = B
    
    If Modo = 4 Then
        'Esta contabilizada
        If Me.Check1.Value = 1 Then
            ' no dejaremos cambiar el importe tampoc
            Text1(5).Enabled = False
            Combo1.Enabled = False
        End If
    End If
    'El text
    B = (Modo = 2) Or (Modo = 5)
   
    If Modo <= 2 Then
         Me.cmdAceptar.Caption = "Aceptar"
         Me.cmdCancelar.Caption = "Cancelar"
    End If
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    
    B = Modo > 2 Or Modo = 1
    cmdCancelar.visible = B
    'Detalles
    
End Sub


Private Function DatosOK() As Boolean
    
    Dim B As Boolean
    B = CompForm(Me)
    

    DatosOK = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub


Private Sub HacerToolBar(Boton As Integer)
Dim vCadena As String

    Select Case Boton
    
    Case 1
        BotonAnyadir
    Case 2
        'Intentamos bloquear la cuenta
        If PuedeRealizarAccion(False, True, False) Then BotonModificar
    Case 3
        If PuedeRealizarAccion(False, False, True) Then BotonEliminar
    
    Case 5
        BotonBuscar
    Case 6
        BotonVerTodos
        
    Case 8
        'Imprimir
        vCadena = ""
        If Text1(4).Text <> "" Then
            vCadena = vCadena & Text1(4).Text & "|" & Text1(1).Text & "|"
            If CCur(data1.Recordset!Talon) = 0 Then 'pagare
                vCadena = vCadena & "1|0|"
            Else
                vCadena = vCadena & "0|1|"
            End If
        End If
        frmTESRecepcionDocList.CadenaInicio = vCadena
        frmTESRecepcionDocList.Show vbModal
    
    End Select
    
End Sub
    
    
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 '        'Contabilizar

            i = 1
            If Combo1.ListIndex = 0 Then
                'PAGARE. Ver si tiene cta puente pagare
                If vParamT.PagaresCtaPuente Then i = 0
            Else
                If vParamT.TalonesCtaPuente Then i = 0
            End If
            If i = 1 Then
                MsgBox "Falta configurar en parametros", vbExclamation
                Exit Sub
            End If
    
            If Not PuedeRealizarAccion(True, False, False) Then Exit Sub
    
    
            SQL = DevuelveDesdeBD("count(*)", "talones_facturas", "codigo", Text1(4).Text)
            If SQL = "" Then SQL = "0"
            If Val(SQL) = 0 Then
                MsgBox "No tiene vencimientos asociados", vbExclamation
                Exit Sub
            End If
    
            'Los importes
            SQL = DevuelveDesdeBD("sum(importe)", "talones_facturas", "codigo", Text1(4).Text)
            If SQL = "" Then SQL = "0"
            i = 0
            If CCur(SQL) <> ImporteFormateado(Text1(5).Text) Then
                SQL = CStr(CCur(SQL) - ImporteFormateado(Text1(5).Text))
                If CCur(SQL) > 0 Then
                    i = -1   'Mayor las lineas que el importe del talon
                Else
                    i = 1   'Mayor el total que la suma de las lineas
                End If
    
                SQL = "Suma de importes distintos del importe del talon: " & SQL
                If vUsu.Nivel <= 1 Then
                    SQL = SQL & vbCrLf & "Seguro que desea continuar?"
                    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                Else
                    MsgBox SQL, vbExclamation
                    Exit Sub
                End If
            End If
    
    
            'Hacemos contabilizacion
            HacerContabilizacion i

    End Select

End Sub


Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LINASI As Long
Dim Ampliacion As String
    
    'Llegados aqui bloqueamos desde form
    '--If Not BloqueaRegistroForm(Me) Then Exit Sub
    If Not PuedeRealizarAccion(False, False, False) Then Exit Sub
    
    If Not BLOQUEADesdeFormulario2(Me, data1, 1) Then Exit Sub

    'Fuerzo que se vean las lineas
    
    Select Case Button.Index
        Case 1
            'AÑADIR linea factura
            AnyadirLinea
        Case 2
            'MODIFICAR linea factura
        Case 3
            'ELIMINAR linea factura
            EliminarLineaFactura
    End Select

End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
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
    DataGrid1.Columns(0).Caption = "Serie"
    DataGrid1.Columns(0).Width = 1000
    
    DataGrid1.Columns(1).Caption = "Factura"
    DataGrid1.Columns(1).Width = 2395
    

    DataGrid1.Columns(2).Caption = "Fecha"
    DataGrid1.Columns(2).Width = 1905
    DataGrid1.Columns(2).NumberFormat = "dd/mm/yyyy"
    
    DataGrid1.Columns(3).Caption = "Vto"
    DataGrid1.Columns(3).Width = 1000
    
    DataGrid1.Columns(4).Caption = "Importe"
    DataGrid1.Columns(4).Width = 2200
    DataGrid1.Columns(4).NumberFormat = FormatoImporte
    DataGrid1.Columns(4).Alignment = dbgRight
    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(0).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = DataGrid1.Columns(1).Left + DataGrid1.Left - cmdAux(0).Width - 15
                
        anc = DataGrid1.Left + 15
        txtAux(1).Left = DataGrid1.Columns(1).Left + anc
        txtAux(1).Width = DataGrid1.Columns(1).Width - 45
    
        txtAux(2).Left = DataGrid1.Columns(2).Left + anc
        txtAux(2).Width = DataGrid1.Columns(2).Width - 45
    
        txtAux(3).Left = DataGrid1.Columns(3).Left + anc
        txtAux(3).Width = DataGrid1.Columns(3).Width - 45

        
        'Concepto
        txtAux(4).Left = DataGrid1.Columns(4).Left + anc
        txtAux(4).Width = DataGrid1.Columns(4).Width - 45
        

       
        CadAncho = True
    End If
        
    For i = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(i).AllowSizing = False
    Next i
    
    DataGrid1.Tag = "Calculando"

    If Modo = 5 Then PonerImporteLinea
    

    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Sub PonerImporteLinea()
Dim C As String
        C = DevuelveDesdeBD("sum(importe)", "talones_facturas", "codigo", Text1(4).Text)
        If C = "" Then C = "0"
        txtSuma.Text = Format(C, FormatoImporte)

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
    SQL = "SELECT numserie,numfactu,fecfactu,numorden,importe From talones_facturas WHERE codigo = "
    If Enlaza Then
        SQL = SQL & Text1(4).Text ' Data1.Recordset!Codigo
    Else
        SQL = SQL & "-1"
    End If
    SQL = SQL & " ORDER BY numserie,numfactu,fecfactu,numorden"
    MontaSQLCarga = SQL
End Function


Private Sub AnyadirLinea()
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
   
    cmdAux_Click 0
    
End Sub

Private Sub EliminarLineaFactura()
Dim Importe As Currency

    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    
    cmdAux_Click (0)
    
  
    CargaGrid True
    DataGrid1.Enabled = True
    PosicionaLineas CInt(NumRegElim)
End Sub

Private Sub PosicionaLineas(Pos As Integer)
    On Error GoTo EPosicionaLineas
    If Pos > 1 Then
        If Pos >= adodc1.Recordset.RecordCount Then Pos = adodc1.Recordset.RecordCount - 1
        adodc1.Recordset.Move Pos
    End If
    
    Exit Sub
EPosicionaLineas:
    Err.Clear
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid DataGrid1
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)


    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B

    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(visible As Boolean, Altura As Single, Limpiar As Boolean)
    Dim i As Integer
    
    
    DataGrid1.Enabled = Not visible

    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = visible
        txtAux(i).top = Altura
    Next i
    
    cmdAux(0).visible = visible
    cmdAux(0).top = Altura
    If Limpiar Then
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    
End Sub



Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub Desplazamiento(Index As Integer)
    If data1.Recordset Is Nothing Then Exit Sub
    If data1.Recordset.EOF Then Exit Sub
    
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

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo

End Sub

Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        'Esto sera k hemos pulsado el ENTER
        txtAux_LostFocus Index
        'cmdAceptar_Click   FALTA###
    Else
        If KeyCode = 113 Then
            
            
        Else
            'Ha pulsado F5. Ponemos linea anterior
            Select Case KeyCode
            Case 116
               
                
            Case 117
                'F6

                
            Case Else
                If (Shift And vbCtrlMask) > 0 Then
                    If UCase(Chr(KeyCode)) = "B" Then
                        'OK. Ha pulsado Control + B
                        '----------------------------------------------------
                        '----------------------------------------------------
                        '
                        ' Dependiendo de index lanzaremos una opcion uotra
                        '
                        '----------------------------------------------------
                        
                        'De momento solo para el 5. Cliente
                        Select Case Index
                        Case 4
                            'txtaux(4).Text = ""
                           ' Image1_Click 1
                        Case 8
                            'txtaux(8).Text = ""
                            'Image1_Click 2
                        End Select
                     End If
                End If
            End Select
        End If
    End If
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 112 Then
        
        End If
    End If
End Sub

Private Sub txtaux_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo <> 1 Then
        If KeyCode = 107 Or KeyCode = 187 Then
                KeyCode = 0
                LanzaPantalla Index
        End If
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    

        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtAux(Index).Text = Trim(txtAux(Index).Text)
    
    
        'Comun a todos
        If txtAux(Index).Text = "" Then
           ' Select Case Index
           ' Case 0
           '
           '
           ' Case 3
           '
           ' Case 4
           '
           ' End Select
            Exit Sub
        End If
        

        
        Select Case Index
        Case 0
            txtAux(0).Text = UCase(txtAux(0).Text)

        Case 1, 3
            If Not EsNumerico(txtAux(Index).Text) Then
                
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
        
          
                
        Case 2
            If Not EsFechaOK(txtAux(Index)) Then
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
            
        Case 4
            PonerFormatoDecimal txtAux(Index), 3
            If txtAux(Index).Text = "" Then
                PonerFoco txtAux(Index)
            Else
                'El importe no puede ser mayor
                If ImporteFormateado(txtAux(Index).Text) > ImporteVto Then
                    MsgBox "El importe NO puede ser mayor al del vencimiento", vbExclamation
                    PonerFoco txtAux(Index)
                End If
            End If
        End Select
        
        
        If Index = 0 Or Index = 1 Then
            If txtAux(0).Text <> "" And txtAux(1).Text <> "" Then PonerCamposVencimiento False
        End If
End Sub



Private Sub PonerCamposVencimiento(DesdeElButon As Boolean)
Dim cad As String
Dim Importe As Currency

        'Veresmos si existe un unico vto para esta factura
        cad = "Select numserie,numfactu,fecfactu,numorden,impvenci,impcobro,tipforpa,Gastos from cobros,formapago"
        cad = cad & "  WHERE cobros.codforpa=formapa.codforpa"
        cad = cad & " AND codmacta ='" & Text1(2).Text & "'"
        
        'Numero de serie y numfac
        If Not DesdeElButon Then
            cad = cad & " AND numserie ='" & txtAux(0).Text & "' AND numfactu = " & txtAux(1).Text
        Else
            cad = cad & SQL  'SQL traera los datos del venciemietno
        End If
            
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        SQL = ""
        While Not miRsAux.EOF
            Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
            Importe = Importe - DBLet(miRsAux!impcobro, "N")
            
            If Importe = 0 Then
                
                    If IsNull(miRsAux!impcobro) Then
                        cad = "Importe es cero"
                    Else
                        cad = "Totalmente cobrado"
                    End If
                    MsgBox cad, vbExclamation
                
            Else
                NumRegElim = vbTalon
                If Me.Combo1.ListIndex = 0 Then NumRegElim = vbPagare
                
                
                SQL = SQL & miRsAux!NUmSerie & "|" & miRsAux!NumFactu & "|"
                SQL = SQL & miRsAux!numorden & "|" & Format(miRsAux!FecFactu, "dd/mm/yyyy") & "|"
                SQL = SQL & Importe & "|" 'Importe
                'si la forma de pago corresponde al documento que estamos procesando
                SQL = SQL & Abs((miRsAux!TipForpa = NumRegElim)) & "|:"
                i = i + 1
                
            End If
            
            miRsAux.MoveNext
        Wend
        
        miRsAux.Close
        Set miRsAux = Nothing
        
        
        If i > 0 Then
            'HAY DATOS
            If i = 1 Then
                'SOLO HAY UNO
                SQL = Mid(SQL, 1, Len(SQL) - 1) 'Le quito los dos puntos
                
            Else
                'Hay mas de uno. Mostraremos una windows
            
                SQL = ""
            End If
            'Pongo los datos
            If SQL <> "" Then PonerDatosVencimiento SQL, DesdeElButon
        End If
End Sub



Private Function AuxOK_() As Boolean
Dim Importe As Currency
Dim cad As String
    AuxOK_ = False
    For i = 0 To txtAux.Count - 1
        If txtAux(i).Text = "" Then
            MsgBox "Campo obligatorio", vbExclamation
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next
    
    If ModificandoLineas = 1 Then
        SQL = DevuelveDesdeBD("sum(importe)", "talones_facturas", "id", Text1(4).Text)
        If SQL = "" Then SQL = "0"
        Importe = ImporteFormateado(SQL)
        If Importe + ImporteFormateado(txtAux(4).Text) > ImporteFormateado(Text1(5).Text) Then
            SQL = CStr(Importe + ImporteFormateado(txtAux(4).Text) - ImporteFormateado(Text1(5).Text))
            SQL = "Suma de importes execede del importe del talon : " & SQL & vbCrLf & vbCrLf
            Importe = ImporteFormateado(Text1(5).Text) - Importe
            SQL = SQL & "Importe maximo del vto: " & Importe
            SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then
                'Pongo el foco en el campo
                PonerFoco txtAux(4)
                Exit Function
            End If
        End If

    
        'Ahora veremos si esta introduciendo un VTO sin el importe total....
        cad = "Select impvenci,impcobro,Gastos from cobros"
        cad = cad & "  WHERE numserie ='" & txtAux(0).Text & "' AND numfactu = " & txtAux(1).Text
        cad = cad & "  AND fecfactu='" & Format(txtAux(2).Text, FormatoFecha) & "' AND numorden = " & txtAux(3).Text
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
                Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
                Importe = Importe - ImporteFormateado(txtAux(4).Text)
                If Importe > 0 Then
                    cad = "Deberia dividir el vencimiento si no lo va a remesar por el total pendiente."
                    cad = cad & vbCrLf & vbCrLf & "¿Continuar?"
                    If MsgBox(cad, vbQuestion + vbYesNo + vbMsgBoxRight) = vbYes Then AuxOK_ = True
                Else
                    AuxOK_ = True
                End If
        Else
            MsgBox "Vencimiento NO encontrado.  Funcion: auxok", vbCritical
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If

    
End Function


Private Function InsertarModificar() As Boolean
Dim Importe As Currency

    On Error GoTo EInsertarModificar
    Set miRsAux = New ADODB.Recordset
    
    InsertarModificar = False
    
    'Cargaremos el VTO de la cobros
    SQL = MontaSQLDelVto(True)
    SQL = " WHERE cobros.codforpa=formapago.codforpa AND " & SQL
    SQL = "select cobros.*,tipforpa from cobros,formapago " & SQL
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "El vencimiento introducido no se corresponde con ningún cobro pendiente", vbExclamation
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Function
    End If
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS

        SQL = "insert into `talones_facturas` (`codigo`,`numserie`,`numfactu`,`fecfactu`, "
        SQL = SQL & "`numvenci`,`importe`,`contabilizado`) VALUES ("
        SQL = SQL & data1.Recordset!Codigo & ",'"
        SQL = SQL & txtAux(0).Text & "',"
        SQL = SQL & txtAux(1).Text & ",'"
        
        SQL = SQL & Format((txtAux(2).Text), FormatoFecha) & "'," & txtAux(3).Text & ","
        SQL = SQL & TransformaComasPuntos(ImporteFormateado(txtAux(4).Text)) & ",0)"
    Else
    End If
    Conn.Execute SQL
    
    
    'Segunda parte del meollo. En la cobros MARCAREMOS el vencimiento
    '
    '      Documento recibido
    '      Importe cobrado
    '      si no tiene forma de pago talon / pager se la pongo

    
    SQL = "UPDATE cobros SET recedocu=1,reftalonpag = '" & DevNombreSQL(Text1(0).Text) & "'"
    Importe = DBLet(miRsAux!impcobro, "N") + ImporteFormateado(txtAux(4).Text)
    SQL = SQL & ", impcobro = " & TransformaComasPuntos(CStr(Importe))
    SQL = SQL & ", fecultco = '" & Format(Text1(1).Text, FormatoFecha) & "'"
    'Febrero 2010
    'Fecha vencimiento tb le pongo la de la recpcion
    SQL = SQL & ", fecvenci = '" & Format(Text1(6).Text, FormatoFecha) & "'"
    'BANCO LO PONGO EN OBSERVACION
    SQL = SQL & ", obs = '" & DevNombreSQL(Text1(3).Text) & "'"
    'Si no era forma de pago talon/pagare la pongo
    If Me.Combo1.ListIndex = 0 Then
        i = vbPagare
    Else
        i = vbTalon
    End If
    If miRsAux!TipForpa <> i Then
        'AQUI BUSCARE una forma de pago
        i = Val(DevuelveDesdeBD("codforpa", "formapago", "tipforpa", CStr(i)))
        If i > 0 Then SQL = SQL & ", codforpa = " & i
        
    End If
    SQL = SQL & " WHERE " & MontaSQLDelVto(True)
    miRsAux.Close
    
    If Not EjecutarSQL(SQL) Then MsgBox "Actualizando Cobros. Avise soporte", vbExclamation
    
    InsertarModificar = True
    
EInsertarModificar:
    If Err.Number <> 0 Then MuestraError Err.Number, "InsertarModificar linea asiento.", Err.Description
    Set miRsAux = Nothing
End Function
 



Private Sub CargaGrid(Enlaza As Boolean)
Dim B As Boolean
    B = DataGrid1.Enabled
    
    DataGrid1.Enabled = False
    DoEvents
    CargaGrid2 Enlaza
    DoEvents
    DataGrid1.Enabled = B
    
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    PonerImporteLinea

End Sub



Private Function Eliminar() As Boolean
On Error GoTo FinEliminar
        'Alguna comprobacion
        
        
        
        'Lineas
        Conn.Execute "Delete  from talones_facturas WHERE codigo =" & Text1(4).Text
        
        'Cabeceras
        Conn.Execute "Delete  from talones WHERE codigo =" & Text1(4).Text
        
                

        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        
        
        Eliminar = False
    Else
       
        Eliminar = True
    End If
End Function






Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Function RecodsetVacio() As Boolean
    RecodsetVacio = True
    If Not adodc1.Recordset Is Nothing Then
        If Not adodc1.Recordset.EOF Then RecodsetVacio = False
    End If
End Function




Private Sub LanzaPantalla(Index As Integer)
Dim miI As Integer
        '----------------------------------------------------
        '----------------------------------------------------
        '
        ' Dependiendo de index lanzaremos una opcion uotra
        '
        '----------------------------------------------------
        
        'De momento solo para el 5. Cliente
        miI = -1
        Select Case Index
        Case 0
            txtAux(0).Text = ""
            miI = 3
        Case 3
            txtAux(3).Text = ""
            miI = 0
        Case 4
            txtAux(4).Text = ""
            miI = 1
            
        Case 8
            txtAux(8).Text = ""
            miI = 2
        End Select
       ' If miI >= 0 Then Image1_Click miI
End Sub





Private Function InsertarRegistro() As Boolean
On Error GoTo EInsertarLinea
    InsertarRegistro = False
    
    
    SQL = DevuelveDesdeBD("max(codigo)", "talones", "1", "1") 'Truco del almendruco par obtener el max
    If SQL = "" Then SQL = "0"
    NumRegElim = Val(SQL) + 1
    
    
    Text1(4).Text = NumRegElim
    
    InsertarRegistro = InsertarDesdeForm(Me)
    
    If InsertarRegistro Then cmdAux_Click 0
    
    
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, Err.Description
End Function



'A partir de un string empipado separaremos
Private Sub PonerDatosVencimiento(CADENA As String, Todo As Boolean)
    If RecuperaValor(CADENA, 6) = "0" Then
        If MsgBox("No tiene forma de pago correcta. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    If Todo Then
        txtAux(0).Text = RecuperaValor(CADENA, 1)
        txtAux(1).Text = RecuperaValor(CADENA, 2)
    End If
    txtAux(2).Text = Format(RecuperaValor(CADENA, 4), "dd/mm/yyyy")
    txtAux(3).Text = RecuperaValor(CADENA, 3)
    CADENA = RecuperaValor(CADENA, 5)
    ImporteVto = CCur(CADENA)
    txtAux(4).Text = Format(CADENA, FormatoImporte)
    
    PonerFoco txtAux(4)
End Sub

'ImporteCoincide
'       0:  IMporte del tal/pag igual que el de la suma de las lineas
'       1:  Importe del  "       MAYOR  "
'       -1: Importe    "         MENOR  "
Private Sub HacerContabilizacion(ImporteCoincide As Integer)

    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(1).Text), True) > 1 Then Exit Sub
    
    'Cuenta bloqueada
    If CuentaBloqeada(Text1(2).Text, CDate(Text1(1).Text), True) Then Exit Sub
    
        
    'Para llevarlos a hco
    Conn.Execute "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    
      
    
    'Abrireremos una ventana para seleccionar un par de cosillas
    If Combo1.ListIndex = 0 Then
        CadenaDesdeOtroForm = CStr(vbPagare)
    Else
        CadenaDesdeOtroForm = CStr(vbTalon)
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|" & CStr(ImporteCoincide) & "|"
    
    
    
    
    frmTESListado.Opcion = 23
    frmTESListado.Show vbModal



    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        CadAncho = RemesasCancelacionTALONPAGARE_(Combo1.ListIndex = 1, CInt(Text1(4).Text), CDate(Text1(1).Text), CadenaDesdeOtroForm)
        If CadAncho Then
            Conn.CommitTrans
            'Ahora actualizamos los registros que estan en tmpactualziar

            'Espera
            espera 0.2
            If SituarData1(True) Then PonerCampos
            
        Else
            TirarAtrasTransaccion
        End If
        CadAncho = True  'la vuelvo a poner como estaba
    End If
    Screen.MousePointer = vbDefault
End Sub


'Vamos a borrar el apunte generado anteriormente
'ImporteCoincide
'       0:  IMporte del tal/pag igual que el de la suma de las lineas
'       1:  Importe del  "       MAYOR  "
'       -1: Importe    "         MENOR  "
Private Function HacerDES_Contabilizacion_(ImporteCoincide As Integer) As Boolean

    
    HacerDES_Contabilizacion_ = False
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(1).Text), True) > 1 Then Exit Function
    
    'Cuenta bloqueada
    If CuentaBloqeada(Text1(2).Text, CDate(Text1(1).Text), True) Then Exit Function
    
        
    'Para llevarlos a hco
    Conn.Execute "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    
      
    
    'Abrireremos una ventana para seleccionar un par de cosillas
    If Combo1.ListIndex = 0 Then
        CadenaDesdeOtroForm = CStr(vbPagare)
    Else
        CadenaDesdeOtroForm = CStr(vbTalon)
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "|" & CStr(ImporteCoincide) & "|"
    
   
    frmTESListado.Opcion = 34
    frmTESListado.Show vbModal

    If CadenaDesdeOtroForm <> "" Then
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        CadAncho = EliminarCancelacionTALONPAGARE(Combo1.ListIndex = 1, CInt(Text1(4).Text), CDate(Text1(1).Text), CadenaDesdeOtroForm)
        If CadAncho Then
            Conn.CommitTrans
            HacerDES_Contabilizacion_ = True
        Else
            TirarAtrasTransaccion
        End If
        CadAncho = True  'la vuelvo a poner como estaba
    End If
    Screen.MousePointer = vbDefault
End Function

Private Function MontaSQLDelVto(EnLasLineas As Boolean) As String
    If EnLasLineas Then
        MontaSQLDelVto = " numserie = '" & txtAux(0).Text & "' AND numfactu = " & txtAux(1).Text
        MontaSQLDelVto = MontaSQLDelVto & " and fecfactu ='" & Format(txtAux(2).Text, FormatoFecha) & "' AND numorden = " & txtAux(3).Text
    Else
        With adodc1.Recordset
          MontaSQLDelVto = " numserie = '" & !NUmSerie & "' AND numfactu = " & !NumFactu
          MontaSQLDelVto = MontaSQLDelVto & " and fecfactu ='" & Format(!FecFactu, FormatoFecha) & "' AND numorden = " & !numorden
        End With
    End If
    
End Function

Private Function PuedeRealizarAccion(PermisoAdministrador As Boolean, ModificarCab As Boolean, Eliminar As Boolean) As Boolean
Dim TieneCtaPte As Boolean

    PuedeRealizarAccion = False
    If data1.Recordset.EOF Then Exit Function
    If Modo <> 2 Then Exit Function
    If PermisoAdministrador Then
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene permisos", vbExclamation
            Exit Function
        End If
    End If
    
    
    'AHora compruebo que no esta contabilizado
    SQL = DevuelveDesdeBD("LlevadoBanco", "talones", "codigo", Text1(4).Text)
    If SQL = "1" Then
        'ESTA LLEVADA A BANCO
        If Combo1.ListIndex = 1 Then
            TieneCtaPte = vParamT.TalonesCtaPuente
        Else
            TieneCtaPte = vParamT.PagaresCtaPuente
        End If
        If Check1.Value = 0 And TieneCtaPte Then
            'Hay un error y no esta marcada como contabilziada
            MsgBox "Falta actualizar datos", vbExclamation
            PonerModo 0
            Exit Function
        End If
        
        
        SQL = DevuelveDesdeBD("Contabilizada", "talones", "codigo", Text1(4).Text)
        If SQL = "0" Then
            If Not ModificarCab And TieneCtaPte Then
                MsgBox "Esta contabilizada pero no ha sido llevada a banco", vbExclamation
                Exit Function
            End If
        End If
        
        If ModificarCab Then
        
        Else
            If Not Eliminar Then
                MsgBox "Ya esta en banco", vbExclamation
                Exit Function
            End If
        
            'Si es eliminar
            SQL = "Select cobros.numserie,cobros.numfactu,cobros.fecfactu,cobros.numorden"
            SQL = SQL & " FROM talones_facturas left join cobros on cobros.numserie=talones_facturas.numserie AND cobros.numfactu=talones_facturas.numfactu and"
            SQL = SQL & " cobros.fecfactu = talones_facturas.fecfactu And cobros.numorden = talones_facturas.numorden"
            SQL = SQL & " WHERE id =" & data1.Recordset!Codigo
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            SQL = ""
            NumRegElim = 0
            While Not miRsAux.EOF
                If Not IsNull(miRsAux!codfaccl) Then
                    SQL = SQL & DBLet(miRsAux!NUmSerie, "T") & Format(miRsAux!NumFactu, "000000") & "  " & Format(miRsAux!FecFactu, "dd/mm/yyyy") & vbCrLf
                    NumRegElim = NumRegElim + 1
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            If NumRegElim > 0 Then
                'Hay vencimientos sin eliminar. No se pude eliminar el regisro
                If NumRegElim = 1 Then
                    SQL = "Existe un vencimiento pendiente de eliminar: " & vbCrLf & SQL
                Else
                    SQL = "Existen vencimientos(" & NumRegElim & ") pendientes de eliminar: " & vbCrLf & SQL
                End If
                MsgBox SQL, vbExclamation
                Exit Function
            End If
        End If


    Else
        'Si no esta llevada a banco
        'Si no es para modificar la cabecera si esta contabilizada TAMPOCO dejo continuar
        If Not ModificarCab Then
            'Para eliminar si que dejare pasar
            If Not Eliminar Then
                SQL = DevuelveDesdeBD("Contabilizada", "talones", "codigo", Text1(4).Text)
                If SQL = "1" Then
                    'ESTA CONTABILIZADO
                    MsgBox "Esta contabilizada", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    PuedeRealizarAccion = True
    
End Function


Private Sub NuevoTalonPagareDefecto(Leer As Boolean)
Dim i As Integer
    On Error GoTo ENuevoTalonPagareDefecto
    If Leer Then
        i = CheckValueLeer("talpag")
        Me.Combo1.ListIndex = i
        
    Else
        'Escribir
        i = Combo1.ListIndex
        CheckValueGuardar "talpag", CByte(i)
    End If
    Exit Sub
ENuevoTalonPagareDefecto:
    Err.Clear

End Sub


Private Sub CambiaFechaVto()

    If Me.data1.Recordset!fechavto <> CDate(Text1(6).Text) Then
        Set miRsAux = New ADODB.Recordset
        SQL = "SELECT numserie,numfactu,fecfactu,numorden,importe FROm talones_facturas WHERE id = " & data1.Recordset!Codigo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            SQL = "UPDATE cobros set fecultco='" & Format(Text1(6).Text, FormatoFecha) & "' WHERE"
            SQL = SQL & " numserie = '" & miRsAux!NUmSerie & "' AND fecfactu='" & Format(miRsAux!FecFactu, FormatoFecha)
            SQL = SQL & "' AND numorden= " & miRsAux!numorden & " AND numfactu = " & miRsAux!NumFactu
            Ejecuta SQL
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
    End If
End Sub


Private Sub HacerBusqueda()

Dim CadB As String
CadB = ObtenerBusqueda(Me)

If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda False
        End If
End If
End Sub

'Cuand esta añadiendo una nueva, veremos si coinciden los importes
Private Function ComprobarImportes() As Boolean

On Error GoTo eComprobarImportes
    'Si ha ha sido llevada NO deberia haber entrado
    ComprobarImportes = True 'dejare que salga de las lineas
    SQL = DevuelveDesdeBD("LlevadoBanco", "talones", "codigo", Text1(4).Text)
    
    If SQL = "1" Then
        MsgBox "No deberia haber entrado en edicion de lineas. Llevado a banco", vbExclamation
        Exit Function
    End If
    
    
    
    'Sumas lineas
    ImporteVto = 0
    SQL = DevuelveDesdeBD("sum(importe)", "talones_facturas", "codigo", Text1(4).Text)
    If SQL <> "" Then ImporteVto = CCur(SQL)
    SQL = Format(ImporteVto, FormatoImporte)
    
    
    If Me.Text1(5).Text <> SQL Then
        SQL = "Importes distintos: " & vbCrLf & "Talon/Pagaré: " & Text1(5).Text & vbCrLf & "Lineas vtos: " & SQL
        SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then ComprobarImportes = False
    End If
    
    Exit Function
eComprobarImportes:
    MuestraError Err.Number, Err.Description
End Function

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N")
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N") And (Modo <> 0 And Modo <> 5)
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!especial, "N") And Modo = 2 And vEmpresa.TieneTesoreria
        Me.Toolbar2.Buttons(3).Enabled = DBLet(Rs!especial, "N") And Modo = 2
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        If Not Me.adodc1.Recordset Is Nothing Then
            ToolbarAux.Buttons(2).Enabled = False
            ToolbarAux.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.adodc1.Recordset.RecordCount > 0)
        Else
            ToolbarAux.Buttons(2).Enabled = False
            ToolbarAux.Buttons(3).Enabled = False
        End If
        vUsu.LeerFiltros "ariconta", IdPrograma
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


