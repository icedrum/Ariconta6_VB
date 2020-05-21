VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmempresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresa"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9660
   Icon            =   "frmempresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame Frame1 
      Height          =   420
      Left            =   150
      TabIndex        =   79
      Top             =   5820
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
         TabIndex        =   80
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   75
      Top             =   90
      Width           =   1125
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   76
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   77
         Top             =   180
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   90
      TabIndex        =   39
      Top             =   1020
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8070
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmempresa.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(4)=   "Label1(8)"
      Tab(0).Control(5)=   "Label1(9)"
      Tab(0).Control(6)=   "Label1(10)"
      Tab(0).Control(7)=   "Label1(11)"
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(9)=   "Label1(13)"
      Tab(0).Control(10)=   "Label1(14)"
      Tab(0).Control(11)=   "Label1(15)"
      Tab(0).Control(12)=   "Label1(16)"
      Tab(0).Control(13)=   "Label1(17)"
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(15)=   "Text1(0)"
      Tab(0).Control(16)=   "Text1(1)"
      Tab(0).Control(17)=   "Text1(2)"
      Tab(0).Control(18)=   "Text1(7)"
      Tab(0).Control(19)=   "Text1(8)"
      Tab(0).Control(20)=   "Text1(9)"
      Tab(0).Control(21)=   "Text1(10)"
      Tab(0).Control(22)=   "Text1(11)"
      Tab(0).Control(23)=   "Text1(12)"
      Tab(0).Control(24)=   "Text1(13)"
      Tab(0).Control(25)=   "Text1(14)"
      Tab(0).Control(26)=   "Text1(15)"
      Tab(0).Control(27)=   "Text1(16)"
      Tab(0).Control(28)=   "Text1(17)"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Otros datos"
      TabPicture(1)   =   "frmempresa.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(18)"
      Tab(1).Control(1)=   "Label1(6)"
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).Control(4)=   "Label1(2)"
      Tab(1).Control(5)=   "Label1(19)"
      Tab(1).Control(6)=   "Label1(21)"
      Tab(1).Control(7)=   "Label1(22)"
      Tab(1).Control(8)=   "Label1(23)"
      Tab(1).Control(9)=   "Label1(24)"
      Tab(1).Control(10)=   "Label1(25)"
      Tab(1).Control(11)=   "Label1(26)"
      Tab(1).Control(12)=   "Label1(28)"
      Tab(1).Control(13)=   "Label1(30)"
      Tab(1).Control(14)=   "Label1(20)"
      Tab(1).Control(15)=   "Label1(31)"
      Tab(1).Control(16)=   "Text1(18)"
      Tab(1).Control(17)=   "Text1(6)"
      Tab(1).Control(18)=   "Text1(5)"
      Tab(1).Control(19)=   "Text1(4)"
      Tab(1).Control(20)=   "Text1(3)"
      Tab(1).Control(21)=   "Text1(19)"
      Tab(1).Control(22)=   "Text1(21)"
      Tab(1).Control(23)=   "Text1(22)"
      Tab(1).Control(24)=   "Text1(23)"
      Tab(1).Control(25)=   "Text1(24)"
      Tab(1).Control(26)=   "Text1(25)"
      Tab(1).Control(27)=   "Text1(26)"
      Tab(1).Control(28)=   "Text1(20)"
      Tab(1).Control(29)=   "Text1(31)"
      Tab(1).Control(30)=   "Text1(32)"
      Tab(1).Control(31)=   "Text1(33)"
      Tab(1).Control(32)=   "Text1(34)"
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "Presentación IVA"
      TabPicture(2)   =   "frmempresa.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(27)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label4(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(29)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "AdoAux(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text1(27)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text1(28)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text1(29)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text1(30)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame"
         Height          =   2535
         Left            =   120
         TabIndex        =   82
         Top             =   1920
         Width           =   9135
         Begin VB.TextBox txtaux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   1
            Left            =   5640
            MaxLength       =   10
            TabIndex        =   35
            Text            =   "factura"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            ItemData        =   "frmempresa.frx":0060
            Left            =   240
            List            =   "frmempresa.frx":0062
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1680
            Visible         =   0   'False
            Width           =   5460
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            ItemData        =   "frmempresa.frx":0064
            Left            =   7080
            List            =   "frmempresa.frx":006E
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1680
            Visible         =   0   'False
            Width           =   780
         End
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   2520
            TabIndex        =   84
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
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
         Begin MSDataGridLib.DataGrid DataGridAux 
            Height          =   2040
            Index           =   0
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Width           =   8850
            _ExtentX        =   15610
            _ExtentY        =   3598
            _Version        =   393216
            AllowUpdate     =   0   'False
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   16
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
               Size            =   9
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
         Begin VB.Label Label1 
            Caption         =   "Actividades empresa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   32
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   2295
         End
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
         Index           =   34
         Left            =   -72720
         MaxLength       =   40
         TabIndex        =   16
         Tag             =   "T|T|S|||empresa2|NomEmpresaOficial|||"
         Text            =   "Text1"
         Top             =   840
         Width           =   6300
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
         Index           =   30
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   33
         Tag             =   "IBAN Ingreso|T|S|||empresa2|iban2|||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   4245
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
         Left            =   240
         MaxLength       =   40
         TabIndex        =   32
         Tag             =   "IBAN Devolución|T|S|||empresa2|iban1|||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   4005
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
         Index           =   33
         Left            =   -74760
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "NIF|T|S|||empresa2|siglaempre|||"
         Text            =   "Text1"
         Top             =   840
         Width           =   645
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
         Index           =   17
         Left            =   -67860
         MaxLength       =   8
         TabIndex        =   13
         Tag             =   "Digitos 10º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3960
         Width           =   615
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
         Index           =   16
         Left            =   -67860
         MaxLength       =   8
         TabIndex        =   12
         Tag             =   "Digitos 9º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3510
         Width           =   615
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
         Index           =   15
         Left            =   -67860
         MaxLength       =   8
         TabIndex        =   11
         Tag             =   "Digitos 8º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3060
         Width           =   615
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
         Index           =   14
         Left            =   -67860
         MaxLength       =   8
         TabIndex        =   10
         Tag             =   "Digitos 7º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   2625
         Width           =   615
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
         Left            =   -67860
         MaxLength       =   8
         TabIndex        =   9
         Tag             =   "Digitos 6º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   2175
         Width           =   615
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
         Index           =   12
         Left            =   -70200
         MaxLength       =   8
         TabIndex        =   8
         Tag             =   "Digitos 5º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3960
         Width           =   615
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
         Index           =   11
         Left            =   -70200
         MaxLength       =   8
         TabIndex        =   7
         Tag             =   "Digitos 4º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3510
         Width           =   615
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
         Index           =   10
         Left            =   -70200
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "Digitos 3er nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   3060
         Width           =   615
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
         Index           =   9
         Left            =   -70200
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "Digitos 2º nivel|N|S|||||||"
         Text            =   "Text1"
         Top             =   2625
         Width           =   615
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
         Index           =   8
         Left            =   -70200
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "Digitos 1er nivel|N|N|||||||"
         Text            =   "Text1"
         Top             =   2175
         Width           =   615
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
         Index           =   7
         Left            =   -73200
         MaxLength       =   1
         TabIndex        =   3
         Tag             =   "Numero niveles|N|N|||||||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   480
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
         Left            =   -69045
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   2670
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
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   4485
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
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   840
         Width           =   975
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
         Index           =   32
         Left            =   -67200
         TabIndex        =   58
         Tag             =   "NIF|T|S|||empresa2|codigo||S|"
         Text            =   "CODIGO"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1110
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
         Index           =   28
         Left            =   8280
         MaxLength       =   1
         TabIndex        =   31
         Tag             =   "Letras|T|S|||empresa2|letraseti|||"
         Text            =   "Text1"
         Top             =   540
         Width           =   495
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
         Index           =   31
         Left            =   -68130
         MaxLength       =   9
         TabIndex        =   18
         Tag             =   "Código postal|T|S|||empresa2|telefono|||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   1695
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
         Index           =   20
         Left            =   -74040
         TabIndex        =   15
         Tag             =   "NIF|T|S|||empresa2|nifempre|||"
         Text            =   "Text1"
         Top             =   840
         Width           =   1290
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
         Index           =   27
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   30
         Tag             =   "Admon|T|S|||empresa2|administracion|||"
         Text            =   "Text1"
         Top             =   540
         Width           =   975
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
         Left            =   -68130
         MaxLength       =   9
         TabIndex        =   20
         Tag             =   "Tfno|T|S|||empresa2|tfnocontacto|||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   1725
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
         Index           =   25
         Left            =   -66960
         MaxLength       =   2
         TabIndex        =   26
         Tag             =   "pta|T|S|||empresa2|puerta|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   555
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
         Index           =   24
         Left            =   -67560
         MaxLength       =   2
         TabIndex        =   25
         Tag             =   "Piso|T|S|||empresa2|piso|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   555
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
         Index           =   23
         Left            =   -68160
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "Esca|T|S|||empresa2|escalera|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   555
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
         Index           =   22
         Left            =   -68880
         MaxLength       =   5
         TabIndex        =   23
         Tag             =   "Numero|T|S|||empresa2|numero|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   645
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
         Index           =   21
         Left            =   -74760
         MaxLength       =   2
         TabIndex        =   21
         Tag             =   "Via|T|S|||empresa2|siglasvia|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   585
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
         Index           =   19
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   19
         Tag             =   "Contacto|T|S|||empresa2|contacto|||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   6450
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
         Left            =   -74160
         MaxLength       =   30
         TabIndex        =   22
         Tag             =   "Dirección|T|S|||empresa2|direccion|||"
         Text            =   "Text1"
         Top             =   3150
         Width           =   5130
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
         Index           =   4
         Left            =   -70185
         TabIndex        =   28
         Tag             =   "Código postal|T|S|||empresa2|codpos|||"
         Text            =   "Text1"
         Top             =   3990
         Width           =   1155
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
         Index           =   5
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   27
         Tag             =   "Población|T|S|||empresa2|poblacion|||"
         Text            =   "Text1"
         Top             =   3990
         Width           =   4440
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
         Left            =   -68895
         MaxLength       =   30
         TabIndex        =   29
         Tag             =   "Provincia|T|S|||empresa2|provincia|||"
         Text            =   "Text1"
         Top             =   3990
         Width           =   2490
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
         Index           =   18
         Left            =   -74745
         MaxLength       =   30
         TabIndex        =   17
         Tag             =   "Apoderado|T|S|||empresa2|apoderado|||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   6420
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   6720
         Top             =   960
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
         Caption         =   "AdoAux(1)"
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
      Begin VB.Label Label1 
         Caption         =   "Nombre empresa Oficial"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   -72720
         TabIndex        =   81
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
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
         Index           =   20
         Left            =   -74040
         TabIndex        =   55
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Siglas"
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
         Index           =   30
         Left            =   -74760
         TabIndex        =   74
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "DIGITOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72465
         TabIndex        =   73
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "10º Nivel"
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
         Index           =   17
         Left            =   -68850
         TabIndex        =   72
         Top             =   3945
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "9º Nivel"
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
         Index           =   16
         Left            =   -68820
         TabIndex        =   71
         Top             =   3510
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "8º Nivel"
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
         Index           =   15
         Left            =   -68820
         TabIndex        =   70
         Top             =   3060
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "7º Nivel"
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
         Index           =   14
         Left            =   -68850
         TabIndex        =   69
         Top             =   2655
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "6º Nivel"
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
         Index           =   13
         Left            =   -68850
         TabIndex        =   68
         Top             =   2190
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "5º Nivel"
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
         Index           =   12
         Left            =   -71250
         TabIndex        =   67
         Top             =   3960
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "4º Nivel"
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
         Index           =   11
         Left            =   -71250
         TabIndex        =   66
         Top             =   3510
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "3er Nivel"
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
         Index           =   10
         Left            =   -71250
         TabIndex        =   65
         Top             =   3075
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "2º Nivel"
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
         Index           =   9
         Left            =   -71250
         TabIndex        =   64
         Top             =   2655
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "1er Nivel"
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
         Index           =   8
         Left            =   -71250
         TabIndex        =   63
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Nº de niveles"
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
         Index           =   7
         Left            =   -74595
         TabIndex        =   62
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Abreviado"
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
         Left            =   -69045
         TabIndex        =   61
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre empresa"
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
         Left            =   -73680
         TabIndex        =   60
         Top             =   600
         Width           =   2655
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   59
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reg.devolución mensual (Art. 30 RIVA)  (S-N)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   3600
         TabIndex        =   57
         Top             =   600
         Width           =   4545
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
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
         Index           =   28
         Left            =   -68130
         TabIndex        =   56
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "IBAN ingreso"
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
         Left            =   4440
         TabIndex        =   54
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label Label4 
         Caption         =   "IBAN devolución"
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
         Left            =   240
         TabIndex        =   53
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo admon. AEAT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   27
         Left            =   210
         TabIndex        =   52
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
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
         Index           =   26
         Left            =   -68160
         TabIndex        =   51
         Top             =   2040
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Pta"
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
         Index           =   25
         Left            =   -66930
         TabIndex        =   50
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Piso"
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
         Index           =   24
         Left            =   -67530
         TabIndex        =   49
         Top             =   2880
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Esca."
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
         Index           =   23
         Left            =   -68130
         TabIndex        =   48
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Num."
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
         Index           =   22
         Left            =   -68850
         TabIndex        =   47
         Top             =   2880
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Via"
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
         Index           =   21
         Left            =   -74760
         TabIndex        =   46
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Persona de contacto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   -74760
         TabIndex        =   45
         Top             =   2040
         Width           =   2445
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección"
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
         Left            =   -74160
         TabIndex        =   44
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "C.Postal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   -70185
         TabIndex        =   43
         Top             =   3750
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   -74760
         TabIndex        =   42
         Top             =   3750
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   -68895
         TabIndex        =   41
         Top             =   3750
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre apoderado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   -74745
         TabIndex        =   40
         Top             =   1320
         Width           =   2160
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5640
      Top             =   1050
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5850
      Width           =   1035
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
      Left            =   6885
      TabIndex        =   37
      Top             =   5850
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   8520
      TabIndex        =   78
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
End
Attribute VB_Name = "frmempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public PrimeraConfiguracion As Boolean

Private Const IdPrograma = 101

Dim Rs As ADODB.Recordset
Dim Modo As Byte
Dim ModoLineas As Byte

Private Sub cmdAceptar_Click()
    Dim cad As String
    Dim I As Integer
    Dim HayPpal As Boolean
    Dim Aux As String
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1


    Select Case Modo
    Case 0
    
        
    Case 3
        
        If DatosOK Then
            If vEmpresa Is Nothing Then InsertarModificarEmpresa True
            If InsertarDesdeForm(Me) Then PonerModo 0
        End If
    
    Case 4
            'Modificar
            If DatosOK Then
                InsertarModificarEmpresa False
                '-----------------------------------------
                'Hacemos insertar
                If adodc1.Recordset.EOF Then
                    I = InsertarDesdeForm(Me)
                Else
                    I = ModificaDesdeFormulario(Me)
                End If
                If I = -1 Then PonerModo 0
            End If

    Case 5
        'IAE
        
        
        cad = ""
        If Combo2(0).ListIndex < 0 Then cad = cad & "Tipo actividad" & vbCrLf
        If Combo2(1).ListIndex < 0 Then cad = cad & "Principal: SI - NO" & vbCrLf
        txtaux(1).Text = Trim(txtaux(1).Text)
        If txtaux(1).Text <> "" Then
            If Len(txtaux(1).Text) <> 4 Then cad = "Epigrafe, si se indica, de cuatro caracteres"
        End If
       ' If txtaux(1).Text = "" Then Cad = Cad & "Epigrafe "
        If cad <> "" Then
            MsgBox "Campos obligatorios: " & vbCrLf & cad, vbExclamation
            Exit Sub
        End If
        
        
        
        
        Aux = ""
        If ModoLineas = 0 Then
            'insertar
            
            cad = DevuelveDesdeBD("max(codigo)", "empresaactiv", "1", "1")
            NumRegElim = Val(cad) + 1
            cad = "INSERT INTO empresaactiv(codigo,id,Epigrafe,ppal)  VALUES   (" & NumRegElim & "," & Combo2(0).ItemData(Combo2(0).ListIndex) & ","
            cad = cad & DBSet(txtaux(1).Text, "T", "S") & "," & Combo2(1).ListIndex & ")"
            If Combo2(1).ListIndex = 1 Then Aux = "UPDATE empresaactiv SET ppal=0 "
        Else
            'modificad
            NumRegElim = Me.AdoAux(0).Recordset!Codigo
            cad = "REPLACE INTO empresaactiv (codigo,id,Epigrafe,ppal)  VALUES   (" & NumRegElim & "," & Combo2(0).ItemData(Combo2(0).ListIndex) & ","
            cad = cad & DBSet(txtaux(1).Text, "T", "S") & "," & Combo2(1).ListIndex & ")"
            
            If Me.AdoAux(0).Recordset!Ppal = "*" Then
                
                If Combo2(1).ListIndex = 0 Then  'era ppal y lo ha quitadp
                    Aux = DevuelveDesdeBD("min(codigo)", "empresaactiv", "1", "1")
                    If Aux <> "" Then Aux = "UPDATE empresaactiv SET ppal=1 WHERE codigo = " & Aux
                End If
            Else
                If Combo2(1).ListIndex = 1 Then Aux = "UPDATE empresaactiv SET ppal=0 "
            End If
        End If
        
        If Aux <> "" Then Conn.Execute Aux
        Conn.Execute cad
        LLamaLineas False, 0
        CargaGrid True
        PonerModo 0
    End Select

        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
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
End Sub


Private Sub cmdCancelar_Click()
Select Case Modo
Case 0
   
Case 3
    PonerModo 3
Case 4
    PonerCampos
    PonerModo 0
Case 5
    LLamaLineas False, 0
    CargaGrid True
    PonerModo 0
End Select
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
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

    Me.Icon = frmppal.Icon
    
    Me.top = 200
    Me.Left = 400
    Limpiar Me
    'Lista imagen
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(2).Image = 4
    End With
    
    Text1(0).Enabled = False
    Me.SSTab1.Tab = 0
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
    With Me.ToolbarAux
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    
    
    
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = "Select * from empresa2"
    adodc1.Refresh
    
    If vEmpresa Is Nothing Then
        'No hay datos
        CargaGrid False
        PonerModo 3 '1
                
        'SQl
        Me.Tag = "select * from usuarios.empresasariconta where ariconta='" & vUsu.CadenaConexion & "'"
        Set Rs = New ADODB.Recordset
        Rs.Open Me.Tag, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Rs.EOF Then
            MsgBox "Error fatal.  ---  NO HAY EMPRESA ---", vbCritical
            End
            Exit Sub
        End If
        Text1(0).Text = Rs!codempre
        Text1(1).Text = Rs!nomempre
        Text1(2).Text = Rs!nomresum
        Rs.Close
        
    Else
        PonerCampos
        PonerModo 0
    End If
    
    
    
    If adodc1.Recordset.EOF Then Text1(38).Text = "1"  'Codigo para la tabla 2 de empresa
    If Toolbar1.Buttons(1).Enabled Then _
        Toolbar1.Buttons(1).Enabled = (vUsu.Nivel <= 1)
    cmdAceptar.Enabled = (vUsu.Nivel <= 1)
End Sub


'

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
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
    Dim mTag As CTag
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    FormateaCampo Text1(Index)  'Formateamos el campo si tiene valor
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Modo = Kmodo
    
    For I = 1 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
    
    Select Case Kmodo
    Case 0
        'Preparamos para ver los datos
        For I = 1 To Text1.Count - 1
            Text1(I).Locked = True
        Next I
        lblIndicador.Caption = ""
    Case 3
        'Preparamos para que pueda insertar
        For I = 1 To Text1.Count - 1
            Text1(I).Text = ""
            Text1(I).Locked = False
        Next I
        lblIndicador.Caption = "INSERTAR"
    Case 4
        For I = 1 To Text1.Count - 1
            Text1(I).Locked = False
        Next I
        lblIndicador.Caption = "MODIFICAR"
        
        
    Case 5
        lblIndicador.Caption = "Lineas"
    End Select
    Me.Toolbar1.Buttons(1).Enabled = Modo <> 3 And Modo <> 5
    cmdAceptar.visible = Modo > 0
    cmdCancelar.visible = Modo > 0

    PonerModoUsuarioGnral Modo, "ariconta"
    
    
    Me.Frame2.Enabled = Modo < 3 Or Modo > 4
    
    Me.ToolbarAux.Buttons(1).Enabled = Modo = 0
    Me.ToolbarAux.Buttons(2).Enabled = Modo = 0
    Me.ToolbarAux.Buttons(3).Enabled = Modo = 0
    
    
End Sub

Private Sub PonerCampos()
    If Not vEmpresa Is Nothing Then
        With vEmpresa
            Text1(0).Text = .codempre
            Text1(1).Text = .nomempre
            Text1(2).Text = .nomresum
            Text1(7).Text = .numnivel
            Text1(8).Text = .numdigi1
            Text1(9).Text = PonTextoNivel(.numdigi2)
            Text1(10).Text = PonTextoNivel(.numdigi3)
            Text1(11).Text = PonTextoNivel(.numdigi4)
            Text1(12).Text = PonTextoNivel(.numdigi5)
            Text1(13).Text = PonTextoNivel(.numdigi6)
            Text1(14).Text = PonTextoNivel(.numdigi7)
            Text1(15).Text = PonTextoNivel(.numdigi8)
            Text1(16).Text = PonTextoNivel(.numdigi9)
            Text1(17).Text = PonTextoNivel(.numdigi10)
        End With
        CargaGrid True
    Else
        Limpiar Me
        CargaGrid False
    End If
    If adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, adodc1
End Sub

Private Function PonTextoNivel(Nivel As Integer) As String
If Nivel <> 0 Then
    PonTextoNivel = Nivel
Else
    PonTextoNivel = ""
End If
End Function


Private Function DatosOK() As Boolean
    Dim Rs As ADODB.Recordset
    Dim B As Boolean
    Dim I As Integer
    Dim J As Integer
    Dim K As Integer
    
    DatosOK = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Otras cosas importantes
    'Comprobamos que tienen n niveles, y solo n
    J = CInt(Text1(7).Text)
    K = 0
    For I = 8 To 17
        If Text1(I).Text <> "" Then K = K + 1
    Next I
    If K <> J Then
        MsgBox "Niveles contables: " & J & vbCrLf & "Niveles parametrizados: " & K, vbExclamation
        Exit Function
    End If
    
    'K los niveles sean consecitivos sin saltar ninguno y sin ser menor
    J = 1
    K = CInt(Text1(8).Text)
    For I = 9 To 17
        If Text1(I).Text = "" Then
            J = 1000 + I
        Else
            J = CInt(Text1(I).Text)
        End If
        If J <= K Then
            MsgBox "Error en la asignacion de niveles. ", vbExclamation
            Exit Function
        End If
        K = J
    Next I
    
    'letraseti   Es si esta inscrita en el regimend e dolucion
    If Text1(28).Text <> "" Then
        If Text1(28).Text <> "S" And Text1(28).Text <> "N" Then
            MsgBox "Registro de devolución mensual (Art. 30 RIVA)" & vbCrLf & vbCrLf & " Valores posibles: S ó N"
            Exit Function
        End If
    End If
    DatosOK = True
End Function




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 2
        'A modificar
        PonerModo 4 '2
        PonFoco Text1(2)
        PonFoco Text1(1)
    End Select
End Sub


Private Sub InsertarModificarEmpresa(Insertar As Boolean)

If Insertar Then Set vEmpresa = New Cempresa
With vEmpresa
    If Insertar Then .codempre = Val(Text1(0).Text)
    .nomempre = Text1(1).Text
    .nomresum = Text1(2).Text
    .numnivel = Val(Text1(7).Text)
    .numdigi1 = Val(Text1(8).Text)
    .numdigi2 = Val(Text1(9).Text)
    .numdigi3 = Val(Text1(10).Text)
    .numdigi4 = Val(Text1(11).Text)
    .numdigi5 = Val(Text1(12).Text)
    .numdigi6 = Val(Text1(13).Text)
    .numdigi7 = Val(Text1(14).Text)
    .numdigi8 = Val(Text1(15).Text)
    .numdigi9 = Val(Text1(16).Text)
    .numdigi10 = Val(Text1(17).Text)
    If Insertar Then
        If .Agregar = 1 Then
            MsgBox "Error fatal insertando datos empresa.", vbCritical
            End
        End If
    Else
        .Modificar
    End If
End With
End Sub

Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cad As String
    If Modo <> 0 Then Exit Sub
    
    
    If Button.Index < 3 Then
        If Combo2(0).ListCount = 0 Then
            CargarCombo_Tabla Combo2(0), "usuarios.wepigrafeiae ", "id", "concat(clave,' - ',substring(descripcion,if(instr(descripcion,':')>0,instr(descripcion,':')+2,1),255))", "", False, "id"
        End If
        
        
        'Si llevamos 6, no deja meter mas. Preentacion por fichero 303
        If Button.Index = 1 Then
            If Not Me.AdoAux(0).Recordset.EOF Then
              If Me.AdoAux(0).Recordset.RecordCount >= 6 Then
                    MsgBox "Solo se pueden configurar 6 actividades IAE (Modelo 303)", vbExclamation
                    Exit Sub
              End If
            End If
        End If
    End If
    
    
    
    If Button.Index < 3 Then
        PonerModo 5
    
        If Button.Index = 1 Then AnyadirLinea DataGridAux(0), AdoAux(0)
        LLamaLineas True, CByte(Button.Index)
    
    Else
        If Me.AdoAux(0).Recordset.EOF Then Exit Sub
        
        
        
        cad = "Desea eliminar la actividad: " & vbCrLf & AdoAux(0).Recordset!descr & vbCrLf
        cad = cad & "Epigrafe: " & DBLet(AdoAux(0).Recordset!epigrafe, "T")
        cad = cad & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbYesNoCancel + vbQuestion) <> vbYes Then Exit Sub
        
        
        cad = "Delete from empresaactiv WHERE codigo =" & AdoAux(0).Recordset!Codigo
        Conn.Execute cad
        
        If AdoAux(0).Recordset!Ppal = "*" Then
            cad = " codigo <> " & AdoAux(0).Recordset!Codigo & " AND 1"
            cad = DevuelveDesdeBD("min(codigo)", "empresaactiv", cad, "1")
            If cad <> "" Then cad = "UPDATE empresaactiv SET ppal=1 WHERE codigo = " & cad
        End If
        CargaGrid True
    End If
    
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
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
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2)
    
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub CargaGrid(Enlaza As Boolean)
Dim Sql
Dim Index As Integer
Dim B As Boolean
Dim tots As String

    Index = 0
    
    Sql = "select codigo,empresaactiv.id,"
    'concat(clave,' - ',substring(descripcion,1,60)) descr"
    Sql = Sql & "concat(clave,' - ',substring(descripcion,if(instr(descripcion,':')>0,instr(descripcion,':')+2,1),255)) descr"
    Sql = Sql & " ,epigrafe,if(ppal=1,'*','') Ppal"
    Sql = Sql & " from empresaactiv left join usuarios.wepigrafeiae on empresaactiv.id=wepigrafeiae.id WHERE "
    If Enlaza Then
        Sql = Sql & " true"
    Else
        Sql = Sql & " false"
    End If
    Sql = Sql & " ORDER BY ppal desc,codigo"
    
    
    B = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = Conn
    AdoAux(Index).RecordSource = Sql
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 350
    
    'If PrimeraVez Then
    If True Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For I = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(I).AllowSizing = False
    Next I
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), CStr(Sql), True
    
    'codigo,empresaactiv.id,concat(clave,' - ',substring(descripcion,1,50)) descr,epigrafe,if(ppal=1,'*','') Ppal"
            tots = "N||||0|;N||||0|;S|Combo2(0)|C|Clave|6405|;S|txtaux(1)|T|Epigrafe|995|;S|Combo2(1)|C|Ppal|805|;"
                
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(Index).Columns(4).Alignment = dbgLeft

            If (Enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For I = 0 To 4
                 '   txtaux(i).Text = ""
                Next I
                
            End If

    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
 '       DataGridAux_RowColChange Index, 1, 1
    Else
'        LimpiarCamposFrame Index
    End If
    
    
    
    




End Sub



Private Sub LLamaLineas(visible As Boolean, Insertar As Byte)
Dim anc As Single
Dim alto As Single
Dim jj As Integer
Dim B As Boolean

        DeseleccionaGrid DataGridAux(0)
        
        If visible Then
            anc = DataGridAux(0).top
            If DataGridAux(0).Row < 0 Then
                anc = anc + 230
            Else
                anc = anc + DataGridAux(0).RowTop(DataGridAux(0).Row) + 5
            End If
            
            Me.Combo2(0).top = anc
            Me.Combo2(1).top = anc
            txtaux(1).top = anc
                
                
            If Insertar = 1 Then
                
            
                ModoLineas = 0
                Combo2(0).ListIndex = -1
                Combo2(1).ListIndex = 0
                txtaux(1).Text = ""
            Else
                ModoLineas = 1
                txtaux(1).Text = AdoAux(0).Recordset!epigrafe
                For jj = 0 To Me.Combo2(0).ListCount - 1
                    If Combo2(0).ItemData(jj) = AdoAux(0).Recordset!Id Then
                        Combo2(0).ListIndex = jj
                        Exit For
                    End If
                Next jj
                
                Combo2(1).ListIndex = IIf(CStr(AdoAux(0).Recordset!Ppal) = "*", 1, 0)
                
            End If
                
            
        Else
             DataGridAux(0).AllowAddNew = False
             
        End If

        Me.Combo2(0).visible = visible
        Me.Combo2(1).visible = visible
        txtaux(1).visible = visible
        
        
        If visible Then PonleFoco Combo2(0)
       
           
           

End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
