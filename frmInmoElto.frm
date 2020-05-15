VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmInmoElto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Elementos de inmovilizado"
   ClientHeight    =   10410
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13500
   Icon            =   "frmInmoElto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Index           =   1
      Left            =   8070
      Locked          =   -1  'True
      TabIndex        =   101
      Text            =   "Text4"
      Top             =   5400
      Width           =   5055
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
      Left            =   6720
      MaxLength       =   30
      TabIndex        =   23
      Tag             =   "Ubicacion|N|N|0||inmovele|ubicacion|||"
      Text            =   "commor"
      Top             =   5400
      Width           =   1275
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
      Index           =   0
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   99
      Text            =   "Text4"
      Top             =   5400
      Width           =   4935
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
      Left            =   240
      MaxLength       =   30
      TabIndex        =   22
      Tag             =   "Ubicacion|N|N|0||inmovele|seccion|||"
      Text            =   "commor"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
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
      Left            =   11160
      TabIndex        =   97
      Tag             =   "Sub|N|N|0||inmovele|subvencionado|||"
      Top             =   3270
      Width           =   255
   End
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
      Height          =   300
      Left            =   8190
      TabIndex        =   84
      Top             =   330
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3630
      TabIndex        =   70
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   71
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   62
      Top             =   0
      Width           =   3405
      Begin VB.CheckBox Check2 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   63
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   150
         TabIndex        =   69
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   240
      TabIndex        =   59
      Top             =   5880
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   6694
      _Version        =   393216
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Histórico Inmovilizado"
      TabPicture(0)   =   "frmInmoElto.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAux0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameTotales"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cuentas de Reparto"
      TabPicture(1)   =   "frmInmoElto.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Observaciones"
      TabPicture(2)   =   "frmInmoElto.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(23)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
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
         Height          =   2760
         Index           =   23
         Left            =   -74520
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   103
         Tag             =   "O|T|S|||inmovele|observaciones|||"
         Text            =   "frmInmoElto.frx":0060
         Top             =   600
         Width           =   11835
      End
      Begin VB.Frame FrameTotales 
         BorderStyle     =   0  'None
         Height          =   2835
         Left            =   6960
         TabIndex        =   90
         Top             =   600
         Width           =   5595
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
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
            Index           =   4
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   93
            Text            =   "commor"
            Top             =   690
            Width           =   1785
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   92
            Text            =   "commor"
            Top             =   1320
            Width           =   1785
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   2970
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   91
            Text            =   "commor"
            Top             =   1950
            Width           =   1785
         End
         Begin VB.Label Label1 
            Caption         =   "Valor adquisición"
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
            Left            =   960
            TabIndex        =   96
            Top             =   750
            Width           =   2085
         End
         Begin VB.Label Label1 
            Caption         =   "Amort. Acumulada"
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
            Left            =   960
            TabIndex        =   95
            Top             =   1350
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Amort. Pendiente"
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
            Left            =   960
            TabIndex        =   94
            Top             =   1980
            Width           =   1785
         End
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   3405
         Left            =   150
         TabIndex        =   80
         Top             =   360
         Width           =   7035
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   1
            Left            =   2490
            TabIndex        =   85
            Top             =   2190
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux1 
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
            Left            =   3780
            TabIndex        =   88
            Tag             =   "Importe|N|N|0||inmovele_his|imporinm|#,###,##0.00||"
            Top             =   2190
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtaux1 
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
            Index           =   2
            Left            =   2730
            TabIndex        =   87
            Tag             =   "Porcentaje|N|N|0||inmovele_his|porcinm|##0.00||"
            Top             =   2190
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox txtaux1 
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
            Left            =   420
            TabIndex        =   89
            Tag             =   "Elto inmovilizado|N|N|0||inmovele_his|codinmov||S|"
            Top             =   2190
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox txtaux1 
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
            Left            =   1470
            TabIndex        =   86
            Tag             =   "Fecha|F|N|||inmovele_his|fechainm|dd/mm/yyyy|S|"
            Top             =   2190
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Frame FrameToolAux 
            Height          =   555
            Index           =   1
            Left            =   150
            TabIndex        =   81
            Top             =   0
            Width           =   1845
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Index           =   0
               Left            =   210
               TabIndex        =   82
               Top             =   150
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   4
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Insertar"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Modificar"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Eliminar"
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Búsqueda avanzada"
                  EndProperty
               EndProperty
            End
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmInmoElto.frx":0067
            Height          =   2655
            Index           =   0
            Left            =   150
            TabIndex        =   83
            Top             =   600
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   4683
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            TabAction       =   2
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   330
            Index           =   0
            Left            =   1890
            Top             =   60
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
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   -74760
         TabIndex        =   73
         Top             =   600
         Width           =   12165
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   2
            Left            =   6900
            TabIndex        =   61
            Top             =   2280
            Visible         =   0   'False
            Width           =   195
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
            Height          =   350
            Index           =   1
            Left            =   1470
            TabIndex        =   65
            Tag             =   "Lin|N|N|0||inmovele_rep|numlinea|000000|S|"
            Top             =   2310
            Visible         =   0   'False
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
            Height          =   350
            Index           =   0
            Left            =   450
            TabIndex        =   64
            Tag             =   "Cod|N|N|0||inmovele_rep|codinmov|000000|S|"
            Top             =   2310
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame FrameToolAux 
            Height          =   555
            Index           =   0
            Left            =   150
            TabIndex        =   78
            Top             =   0
            Width           =   1545
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Index           =   1
               Left            =   210
               TabIndex        =   79
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
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   0
            Left            =   3480
            TabIndex        =   60
            Top             =   2310
            Visible         =   0   'False
            Width           =   195
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
            Height          =   350
            Index           =   2
            Left            =   2580
            TabIndex        =   66
            Tag             =   "Cuenta|T|N|||inmovele_rep|codmacta2|||"
            Top             =   2310
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtaux2 
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
            Height          =   350
            Index           =   0
            Left            =   3600
            TabIndex        =   75
            Top             =   2310
            Visible         =   0   'False
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
            Height          =   350
            Index           =   3
            Left            =   5940
            MaxLength       =   4
            TabIndex        =   67
            Tag             =   "Centro Coste|T|S|||inmovele_rep|codccost|||"
            Top             =   2310
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtaux2 
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
            Height          =   350
            Index           =   1
            Left            =   7080
            MaxLength       =   10
            TabIndex        =   74
            Top             =   2310
            Visible         =   0   'False
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
            Height          =   350
            Index           =   4
            Left            =   8130
            MaxLength       =   5
            TabIndex        =   68
            Tag             =   "Porcentaje|N|N|0||inmovele_rep|porcenta|##0.00||"
            Top             =   2310
            Visible         =   0   'False
            Width           =   765
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmInmoElto.frx":007C
            Height          =   2265
            Index           =   1
            Left            =   150
            TabIndex        =   76
            Top             =   600
            Width           =   11625
            _ExtentX        =   20505
            _ExtentY        =   3995
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            TabAction       =   2
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   330
            Index           =   1
            Left            =   1890
            Top             =   60
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
      End
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Index           =   7
      Left            =   6570
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   57
      Text            =   "commor"
      Top             =   2520
      Width           =   1875
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "frmInmoElto.frx":0091
      Left            =   11070
      List            =   "frmInmoElto.frx":00A1
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "Situación|N|N|1||inmovele|situacio|||"
      Top             =   2520
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
      Index           =   19
      Left            =   9030
      MaxLength       =   30
      TabIndex        =   19
      Tag             =   "Fecha Venta/baja|F|S|||inmovele|fecventa|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   3990
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   16
      Tag             =   "Años vida|N|N|0||inmovele|anovidas|||"
      Text            =   "commor"
      Top             =   3240
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   6690
      MaxLength       =   30
      TabIndex        =   14
      Tag             =   "Años máximo|N|N|0||inmovele|anomaxim|||"
      Text            =   "commor"
      Top             =   3240
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   8160
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "Años minimos|N|N|0||inmovele|anominim|||"
      Text            =   "commor"
      Top             =   3240
      Width           =   1395
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
      Index           =   15
      Left            =   6690
      MaxLength       =   30
      TabIndex        =   18
      Tag             =   "Valor venta/baja|N|S|||inmovele|impventa|#,###,##0.00||"
      Text            =   "commor"
      Top             =   3990
      Width           =   1875
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
      Index           =   14
      Left            =   8910
      MaxLength       =   30
      TabIndex        =   11
      Tag             =   "Valor residual|N|N|0||inmovele|valorres|#,###,##0.00||"
      Text            =   "commor"
      Top             =   2520
      Width           =   1785
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
      Index           =   13
      Left            =   4350
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "Amortización acumulada|N|N|0||inmovele|amortacu|#,###,##0.00||"
      Text            =   "commor"
      Top             =   2520
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
      Index           =   12
      Left            =   2070
      MaxLength       =   30
      TabIndex        =   9
      Tag             =   "Valor adquisición|N|N|0||inmovele|valoradq|#,###,##0.00||"
      Text            =   "commor"
      Top             =   2520
      Width           =   1785
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
      Left            =   180
      MaxLength       =   30
      TabIndex        =   17
      Tag             =   "Cta de amortizacion|N|N|||inmovele|codmact3|||"
      Text            =   "commor"
      Top             =   3990
      Width           =   1665
   End
   Begin VB.TextBox Text4 
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
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "Text4"
      Top             =   3990
      Width           =   4665
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
      Left            =   240
      MaxLength       =   30
      TabIndex        =   20
      Tag             =   "Cta de gastos|N|N|||inmovele|codmact2|||"
      Text            =   "commor"
      Top             =   4710
      Width           =   1575
   End
   Begin VB.TextBox Text4 
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
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "Text4"
      Top             =   4710
      Width           =   4725
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
      Left            =   180
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "Cta inmovilizado|N|N|||inmovele|codmact1|||"
      Text            =   "commor"
      Top             =   3240
      Width           =   1545
   End
   Begin VB.TextBox Text4 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "Text4"
      Top             =   3240
      Width           =   4785
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
      Index           =   8
      Left            =   120
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "Porcentaje|N|N|0|100|inmovele|coeficie|0.00||"
      Text            =   "commor"
      Top             =   2520
      Width           =   1725
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
      ItemData        =   "frmInmoElto.frx":00CD
      Left            =   11070
      List            =   "frmInmoElto.frx":00DD
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Situación|N|N|1||inmovele|tipoamor|||"
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
      Index           =   7
      Left            =   6690
      MaxLength       =   30
      TabIndex        =   21
      Tag             =   "Centro de coste|T|S|||inmovele|codccost|||"
      Text            =   "commor"
      Top             =   4710
      Width           =   885
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
      Left            =   7650
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text3"
      Top             =   4710
      Width           =   5115
   End
   Begin VB.TextBox Text2 
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
      Left            =   6570
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Text4"
      Top             =   1650
      Width           =   4425
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
      Left            =   5400
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Concepto|N|N|0||inmovele|conconam|||"
      Text            =   "commor"
      Top             =   1650
      Width           =   1095
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
      Left            =   2070
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Num|T|S|||inmovele|numserie|||"
      Text            =   "commor"
      Top             =   1650
      Width           =   3285
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
      Left            =   120
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "Fecha adquisición|F|N|||inmovele|fechaadq|dd/mm/yyyy||"
      Text            =   "commor"
      Top             =   1650
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
      Index           =   3
      Left            =   11070
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Fact. proveedor|T|S|||inmovele|factupro|||"
      Text            =   "commor"
      Top             =   990
      Width           =   2205
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
      Left            =   5400
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Proveedor|N|S|||inmovele|codprove|||"
      Text            =   "commor"
      Top             =   990
      Width           =   1395
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5790
      Top             =   60
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   0
      Tag             =   "Cod|N|N|0||inmovele|codinmov|000000|S|"
      Text            =   "Text1"
      Top             =   990
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
      Index           =   1
      Left            =   1230
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||inmovele|nominmov|||"
      Text            =   "000000000000000000000000000000"
      Top             =   990
      Width           =   4125
   End
   Begin VB.TextBox Text4 
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text4"
      Top             =   990
      Width           =   4155
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
      Left            =   12240
      TabIndex        =   25
      Top             =   9960
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   150
      TabIndex        =   26
      Top             =   9810
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
         TabIndex        =   27
         Top             =   240
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
      Left            =   11010
      TabIndex        =   24
      Top             =   9960
      Width           =   1035
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
      Left            =   12240
      TabIndex        =   28
      Top             =   9960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   12450
      TabIndex        =   72
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
      Height          =   315
      Index           =   20
      Left            =   7440
      MaxLength       =   30
      TabIndex        =   77
      Tag             =   "repartos|N|N|0||inmovele|repartos|||"
      Text            =   "commor"
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación"
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
      Index           =   28
      Left            =   6750
      TabIndex        =   102
      Top             =   5160
      Width           =   945
   End
   Begin VB.Image imgOtro 
      Height          =   240
      Index           =   1
      Left            =   7800
      Picture         =   "frmInmoElto.frx":010A
      Top             =   5160
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sección"
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
      Left            =   270
      TabIndex        =   100
      Top             =   5160
      Width           =   780
   End
   Begin VB.Image imgOtro 
      Height          =   240
      Index           =   0
      Left            =   1320
      Picture         =   "frmInmoElto.frx":0B0C
      Top             =   5160
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "z"
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
      Left            =   11400
      TabIndex        =   98
      Top             =   3270
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Amort.Pendiente"
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
      Left            =   6600
      TabIndex        =   58
      Top             =   2250
      Width           =   1785
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   1
      Left            =   10440
      Picture         =   "frmInmoElto.frx":150E
      Top             =   3750
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   0
      Left            =   1680
      Picture         =   "frmInmoElto.frx":1599
      Top             =   1410
      Width           =   240
   End
   Begin VB.Image imgConcep 
      Height          =   240
      Left            =   6570
      Picture         =   "frmInmoElto.frx":1624
      Top             =   1410
      Width           =   240
   End
   Begin VB.Image imgCC 
      Height          =   240
      Left            =   8400
      Picture         =   "frmInmoElto.frx":2026
      Top             =   4470
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   3
      Left            =   3180
      Picture         =   "frmInmoElto.frx":2A28
      Top             =   3720
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   2
      Left            =   1800
      Picture         =   "frmInmoElto.frx":342A
      Top             =   4470
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   1
      Left            =   2280
      Picture         =   "frmInmoElto.frx":3E2C
      Top             =   3000
      Width           =   240
   End
   Begin VB.Image imgCta 
      Height          =   240
      Index           =   0
      Left            =   6510
      Picture         =   "frmInmoElto.frx":482E
      Top             =   780
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Situación"
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
      Left            =   11070
      TabIndex        =   56
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo amortización"
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
      Left            =   11100
      TabIndex        =   55
      Top             =   1410
      Width           =   1785
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   13320
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "F. Venta/Baja"
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
      Index           =   19
      Left            =   9030
      TabIndex        =   54
      Top             =   3720
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "Años vida"
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
      Index           =   18
      Left            =   9840
      TabIndex        =   53
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Años máximo"
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
      Index           =   17
      Left            =   6720
      TabIndex        =   52
      Top             =   2970
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Años mínimos"
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
      Index           =   16
      Left            =   8160
      TabIndex        =   51
      Top             =   2970
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Valor Venta/Baja"
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
      Index           =   15
      Left            =   6720
      TabIndex        =   50
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Valor residual"
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
      Index           =   14
      Left            =   8940
      TabIndex        =   49
      Top             =   2250
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Amort.Acumulada"
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
      Index           =   13
      Left            =   4380
      TabIndex        =   48
      Top             =   2250
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Valor adquisicion"
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
      Index           =   12
      Left            =   2100
      TabIndex        =   47
      Top             =   2250
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta Inmovilizado"
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
      Index           =   11
      Left            =   210
      TabIndex        =   46
      Top             =   3000
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta. Amortizacion Acumulada"
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
      Index           =   10
      Left            =   210
      TabIndex        =   44
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta gastos"
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
      Index           =   9
      Left            =   210
      TabIndex        =   42
      Top             =   4470
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Porcentaje"
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
      Index           =   7
      Left            =   120
      TabIndex        =   40
      Top             =   2280
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Centro de coste"
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
      Index           =   6
      Left            =   6720
      TabIndex        =   39
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Height          =   225
      Index           =   5
      Left            =   5400
      TabIndex        =   36
      Top             =   1410
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Serie"
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
      Left            =   2100
      TabIndex        =   35
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fec.Adquisición"
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
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   1410
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Factura"
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
      Left            =   11070
      TabIndex        =   33
      Top             =   750
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
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
      Index           =   1
      Left            =   5400
      TabIndex        =   32
      Top             =   750
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
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
      Left            =   1230
      TabIndex        =   31
      Top             =   750
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Elemento"
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
      Left            =   120
      TabIndex        =   30
      Top             =   750
      Width           =   1035
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
Attribute VB_Name = "frmInmoElto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 503

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Nuevo As String                      'Nuevo desde el form de facturas
Public Event DatoSeleccionado(CadenaSeleccion As String)


Private Const NO = "No encontrado"
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCCentroCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmCC2 As frmCCCentroCoste
Attribute frmCC2.VB_VarHelpID = -1
Private WithEvents frmCI As frmInmoConceptos
Attribute frmCI.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmEleInmo As frmBasico2
Attribute frmEleInmo.VB_VarHelpID = -1
Private WithEvents frmISec As frmInmoSeccion
Attribute frmISec.VB_VarHelpID = -1
Private WithEvents frmIUbi As frmInmoUbicacion
Attribute frmIUbi.VB_VarHelpID = -1


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
'6 Modo momentaneo para poder poner los campos


'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer
Private BuscaChekc As String
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private Sql As String
Dim I As Integer
Dim Ancho As Integer
'Dim colMes As Integer

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas

'-------------------------------------------------------------


'Para pasar de lineas a cabeceras
Dim Linliapu As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar


Dim TotalLin As Currency
Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean
Dim RC As String

Dim ModoLineas As Byte
Dim NumTabMto As Integer



Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
    Dim cad As String
    Dim I As Integer
    Dim Limp As Boolean

    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
    
        Case 3
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                    If InsertarDesdeForm2(Me, 1) Then
                       espera 0.5
                        
                        
                       'Si nuevo <>"" siginifac k veni,mos de facturas proveedores. Les vamos a devolver
                       'el valor de la cta de amortizacion
                       'y nos salimos
                       If Nuevo <> "" Then
                            CadenaDesdeOtroForm = Text1(9).Text & "|" & Text4(1).Text & "|"
                            'Para que pueda salir
                            Modo = 10
                            PulsadoSalir = True
                            Unload Me
                            Exit Sub
                       End If
                        
                       'Ponemos la cadena consulta
                        If SituarData1 Then
                            lblIndicador.Caption = ""
                            'Ahora preguntamos si tiene centros de desea agregar centros de reparto
                            Sql = "¿Desea agregar sub centros de reparto?"
                            If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
                                PonerModo 5
                                'Haremos como si pulsamo el boton de insertar nuevas lineas
                                ModificandoLineas = 0
                                AnyadirLinea True, 1
                            Else
                                PonerModo 2
                            End If
                        Else
                            Sql = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                            Sql = Sql & vbCrLf & " CLAVE: FRMiNMOV. cmdAceptar. SituarData1"
                            MsgBox Sql, vbCritical
                            Exit Sub
                        End If
                    End If
            End If
        Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario2(Me, 1) Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1 Then
                        lblIndicador.Caption = ""
                        PonerModo 2
                    End If
                End If
            End If
                
        Case 5
            cad = DatosOkLin("")
            If cad <> "" Then
                MsgBox cad, vbExclamation
            Else
                Select Case ModoLineas
                    Case 1
                        InsertarLinea
                    Case 2
                        ModificarLinea
'                        PosicionarData
                End Select
                
                cmdCancelar_Click
                
            End If
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click(Index As Integer)
Dim F As Date
    

    Select Case Index
        Case 0
            cmdAux(0).Tag = 100
            LLamaCuenta
        Case 1
            Set frmF = New frmCal
            F = Now
            If txtaux1(1).Text <> "" Then F = CDate(txtaux1(1).Text)
            I = 2
            frmF.Fecha = F
            frmF.Show vbModal
            Set frmF = Nothing
        
        Case 2 ' centro de coste
            Set frmCC2 = New frmCCCentroCoste
            frmCC2.DatosADevolverBusqueda = "0|1|"
            frmCC2.Show vbModal
            Set frmCC2 = Nothing
        
    End Select


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
        CamposAux False, 0, False, NumTabMto
        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGridAux(1).Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
        Else
            If ModificandoLineas = 1 Then
                 DataGridAux(NumTabMto).AllowAddNew = False
                 If Not AdoAux(NumTabMto).Recordset.EOF Then AdoAux(NumTabMto).Recordset.MoveFirst
                 DataGridAux(NumTabMto).Refresh
            End If
            ModificandoLineas = 0
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
        End If
    End Select
End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim Sql As String
    
    On Error GoTo ESituarData1
    Data1.Refresh
    espera 0.2
    Data1.Recordset.Find "codinmov = " & Text1(0).Text
    If Not Data1.Recordset.EOF Then
        SituarData1 = True
        Exit Function
    End If
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid 0, False
    CargaGrid 1, False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    Text1(0).Text = ObtenerSigueinteNumeroLinea(True)
    
    
    Text1(21).Text = "1"
    Text1(22).Text = "1"
    
    'Ponemos otros valores por defecto
    Text1(13).Text = "0"
    Text1(14).Text = "0"
    Text1(4).Text = Format(Now, "dd/mm/yyyy")
    Combo1.ListIndex = 3
    Combo2.ListIndex = 0
    Text1(1).SetFocus

End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid 0, False
        CargaGrid 1, False
        
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        Combo1.ListIndex = -1
        Combo2.ListIndex = -1
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
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
    CargaGrid 0, False
    CargaGrid 1, False
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
    
Select Case Index
    Case 1
        If Not Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 3
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 4
        Data1.Recordset.MoveLast
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
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
        Sql = "INMOVILIZADO." & vbCrLf
        Sql = Sql & "-----------------------------" & vbCrLf & vbCrLf
        Sql = Sql & "Va a eliminar el elemento de inmovilizado:"
        Sql = Sql & vbCrLf & "Cod.         :   " & Data1.Recordset.Fields(0)
        Sql = Sql & vbCrLf & "Descripción    :   " & CStr(Data1.Recordset.Fields(2))
        Sql = Sql & "      ¿Desea continuar ? "
        I = MsgBox(Sql, vbQuestion + vbYesNoCancel)
        'Borramos
        If I <> vbYes Then Exit Sub
        
        Sql = DevuelveDesdeBD("fechainm", "inmovele_his", "codinmov", Data1.Recordset.Fields(0), "N")
        If Sql <> "" Then
            Sql = "Los datos del histórico de inmovilizado del elemento se borrarán también. ¿Continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        
        'Borro, por si existieran, las lineas
        Sql = "Delete from inmovele_his  WHERE codinmov =" & Data1.Recordset!Codinmov
        Conn.Execute Sql
        
        'Borro el elemento
        Sql = "Delete from inmovele  WHERE codinmov =" & Data1.Recordset!Codinmov
        DataGridAux(1).Enabled = False
        NumRegElim = Data1.Recordset.AbsolutePosition
        Conn.Execute Sql
        
    End If
    
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid 0, False
        CargaGrid 1, False
        PonerModo 0
        Else
            Data1.Recordset.MoveFirst
            NumRegElim = NumRegElim - 1
            If NumRegElim > 1 Then
                For I = 1 To NumRegElim - 1
                    Data1.Recordset.MoveNext
                Next I
            End If
            PonerCampos
            DataGridAux(1).Enabled = True
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Elimina Elto."
            Data1.Recordset.CancelUpdate
        End If
End Sub




Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until I = 0
    
    '###a mano
    'Devuelvo tb el estado de elemento
    cad = cad & Combo2.ListIndex & "|"
    
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        
        Modo = 0
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codinmov = -1"
        Data1.RecordSource = CadenaConsulta
        Data1.Refresh
        
        
        If Nuevo <> "" Then
            BotonAnyadir
            
            'Ahora de la cadena NUEVO desglosamos los datos
            ' Codprove, nomprove, numfac, cta amort
            Text1(2).Text = RecuperaValor(Nuevo, 1)  'codprove
            Text4(0).Text = RecuperaValor(Nuevo, 2)  'nombre
            Text1(3).Text = RecuperaValor(Nuevo, 3)  'numfac
            Text1(4).Text = RecuperaValor(Nuevo, 4)  'fecha adq
            Text1(12).Text = RecuperaValor(Nuevo, 5)  'importe
            Text1(9).Text = RecuperaValor(Nuevo, 6)  'Cuenta
            Text4(1).Text = RecuperaValor(Nuevo, 7)  'Des. cuenta
            
        Else
        
            'Procedimiento normal
            PonerModo CInt(Modo)
            CargaGrid 0, (Modo = 2)
            CargaGrid 1, (Modo = 2)
            
            If Modo <> 2 Then
                CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
                Data1.RecordSource = CadenaConsulta
            End If
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    Me.Icon = frmppal.Icon

    Screen.MousePointer = vbHourglass
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
     
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
   
    For I = 0 To Me.ToolbarAux.Count - 1
        With Me.ToolbarAux(I)
            .HotImageList = frmppal.imgListComun_OM16
            .DisabledImageList = frmppal.imgListComun_BN16
            .ImageList = frmppal.imgListComun16
            .Buttons(1).Image = 3
            .Buttons(2).Image = 4
            .Buttons(3).Image = 5
            If I = 0 Then .Buttons(4).Image = 29
        End With
    Next I
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With


    Me.Icon = frmppal.Icon

    CargaCombo
    Label1(26).Caption = vParam.TextoInmoSubencionado
    Set miTag = New CTag
    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False
    
    'Los campos auxiliares
    CamposAux False, 0, True
    
    
    '## A mano
    NombreTabla = "inmovele"
    Ordenacion = " ORDER BY codinmov"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    
    imgCC.Enabled = vParam.autocoste
    Text1(7).Enabled = vParam.autocoste
        
    PonerModo 0
    
    PonerOpcionesMenu
    PonerModoUsuarioGnral 0, "ariconta"
    
    
    'Maxima longitud cuentas
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(2).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(9).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(10).MaxLength = vEmpresa.DigitosUltimoNivel
    Text1(11).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    'CadAncho = False
    cmdRegresar.visible = DatosADevolverBusqueda <> ""
    
    SSTab1.Tab = 0
    
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    Check1(0).Value = 0
    lblIndicador.Caption = ""
End Sub




Private Sub Form_Unload(Cancel As Integer)
    If Modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Nuevo = ""
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Me.DatosADevolverBusqueda = "" Then Set miTag = Nothing
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas
Select Case cmdAux(0).Tag
Case 100
    'Cuenta normal
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
Case 0, 1, 2, 3
    I = Val(cmdAux(0).Tag)
    Text4(I).Text = RecuperaValor(CadenaSeleccion, 2)
    If I = 0 Then
        I = 2
    Else
        I = I + 8
    End If
    Text1(I).Text = RecuperaValor(CadenaSeleccion, 1)
    
End Select
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCC2_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(3).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCI_DatoSeleccionado(CadenaSeleccion As String)
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
    Text1(8).Text = RecuperaValor(CadenaSeleccion, 3)
    Text1(18).Text = RecuperaValor(CadenaSeleccion, 4)
End Sub

Private Sub frmEleInmo_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "codinmov = " & RecuperaValor(CadenaSeleccion, 1)
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmF_Selec(vFecha As Date)
    Select Case I
        Case 0
            I = 4
            Text1(I).Text = Format(vFecha, "dd/mm/yyyy")
        Case 1
            I = 19
            Text1(I).Text = Format(vFecha, "dd/mm/yyyy")
        Case 2
            txtaux1(1).Text = Format(vFecha, "dd/mm/yyyy")
    End Select
End Sub


Private Sub frmISec_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmIUbi_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub imgCC_Click()
    Set frmCC = New frmCCCentroCoste
    frmCC.DatosADevolverBusqueda = "0|1|"
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub imgConcep_Click()
    Set frmCI = New frmInmoConceptos
    frmCI.DatosADevolverBusqueda = "0|1|2|3|"
    frmCI.Show vbModal
    Set frmCI = Nothing
End Sub

Private Sub imgcta_Click(Index As Integer)
   cmdAux(0).Tag = Index
   LLamaCuenta
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim F As Date
    
    Set frmF = New frmCal
    F = Now
    If Index = 0 Then
        If Text1(4).Text <> "" Then F = CDate(Text1(4).Text)
    Else
        If Text1(19).Text <> "" Then F = CDate(Text1(19).Text)
    End If
    I = Index
    frmF.Fecha = F
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub imgOtro_Click(Index As Integer)

    Sql = ""
    If Index = 0 Then
        Set frmISec = New frmInmoSeccion
        frmISec.DatosADevolverBusqueda = "0|1|"
        frmISec.Show vbModal
        Set frmISec = Nothing
    Else
        Set frmIUbi = New frmInmoUbicacion
        frmIUbi.DatosADevolverBusqueda = "0|1|"
        frmIUbi.Show vbModal
        Set frmIUbi = Nothing
        
        
    End If
    If Sql <> "" Then
        Text1(21 + Index).Text = RecuperaValor(Sql, 1)
        Text5(Index).Text = RecuperaValor(Sql, 2)
        PonFoco Text1(21 + Index)
    End If
End Sub

Private Sub Label1_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Check1(0).Value = IIf(Check1(0).Value = 0, 1, 0)
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
    If Modo = 5 Then Exit Sub
        
    PulsadoSalir = True
    Screen.MousePointer = vbHourglass
    DataGridAux(1).Enabled = False
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

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 0 'cuenta
            Case 9: KEYBusqueda KeyAscii, 1 'cuenta
            Case 10: KEYBusqueda KeyAscii, 2 'cuenta
            Case 11: KEYBusqueda KeyAscii, 3 'cuenta
            Case 6: KEYConcepto KeyAscii 'concepto
            Case 7: KEYCCoste KeyAscii, 0 'centro de coste
            Case 4: KEYFecha KeyAscii, 0 'fecha
            Case 19: KEYFecha KeyAscii, 1 'fecha
            Case 21, 22:   KEYOtro KeyAscii, Index - 21
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
    imgcta_Click (Indice)
End Sub

Private Sub KEYConcepto(KeyAscii As Integer)
    KeyAscii = 0
    imgConcep_Click
End Sub

Private Sub KEYCCoste(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCC_Click
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub



Private Sub KEYOtro(ByRef KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgOtro_Click (Indice)
End Sub

'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim Im As Currency

    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).BackColor = vbLightBlue Then
        Text1(Index).BackColor = vbWhite  '&H80000018
    End If
    
    
    If Modo = 0 Then
        Text1(Index).Text = ""
        Exit Sub
    End If
    
    If Text1(Index).Text = "" Then
        Select Case Index
            Case 2, 9, 10, 11
                    If Index = 2 Then
                        I = 0
                    Else
                        I = (Index - 8)
                    End If
                    Text4(I).Text = ""
                    
            Case 6
                Text2.Text = ""
            Case 7
                Text3.Text = ""
            Case 21, 22
                Text5(Index - 21).Text = ""
        End Select
        Exit Sub
    End If
    'Si estamos insertando o modificando o buscando
    If Modo >= 3 Then  'Or Modo = 4 Or Modo = 1
    
        Select Case Index
        Case 2, 9, 10, 11
                If Index = 2 Then
                    I = 0
                Else
                    I = (Index - 8)
                End If
                RC = Text1(Index).Text
                If CuentaCorrectaUltimoNivel(RC, Sql) Then
                    Text1(Index).Text = RC
                    Text4(I).Text = Sql
                    RC = ""
                Else
                    If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                        Sql = Sql & " ¿Desea crearla?"
                        If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                            CadenaDesdeOtroForm = RC
                            If Index = 2 Then
                                cmdAux(0).Tag = 0
                            Else
                                cmdAux(0).Tag = Index - 8
                            End If
                            Set frmC = New frmColCtas
                            frmC.DatosADevolverBusqueda = "0|1|"
                            frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                            frmC.Show vbModal
                            Set frmC = Nothing
                            If Text1(Index).Text = RC Then Sql = "" 'Para k no los borre
                        End If
                     Else
                        MsgBox Sql, vbExclamation
                     End If
                        
                        
                    If Sql <> "" Then
                        Text1(Index).Text = ""
                        Text4(I).Text = ""
                        RC = "NO"
                    End If
                End If
                If RC <> "" Then PonFoco Text1(Index)
                
        Case 6
                'Concepto de imovilizado
                If Not IsNumeric(Text1(6).Text) Then
                    MsgBox "Concepto debe ser numérico:", vbExclamation
                    Text1(6).SetFocus
                    Exit Sub
                End If
                Sql = DevuelveDesdeBD("nomconam", "inmovcon", "codconam", Text1(6).Text, "N")
                If Sql = "" Then
                    MsgBox "Concepto NO encontrado: " & Text1(6).Text, vbExclamation
                    Text1(6).Text = ""
                    Text1(6).SetFocus
                    Exit Sub
                End If
                Text2.Text = Sql
                Sql = "Select * from inmovcon where codconam =" & Text1(6).Text
                Set miRsAux = New ADODB.Recordset
                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                If miRsAux.EOF Then
                    MsgBox "Centro de coste NO encontrado: " & Text1(6).Text, vbExclamation
                    Text1(6).Text = ""
                    Text2.Text = ""
                    Text1(6).SetFocus
                Else
                    Text1(6).Text = miRsAux.Fields(0)
                    Text2.Text = miRsAux.Fields(1)
                    If Modo < 6 Then
                        'Solo insertando y modificando
                        Text1(8).Text = miRsAux.Fields(2)
                        Text1(18).Text = miRsAux.Fields(3)
                    End If
                End If
                miRsAux.Close
                Set miRsAux = Nothing
                
        Case 7
                'Centro de coste
                Text1(7).Text = UCase(Text1(7).Text)
                Sql = DevuelveDesdeBD("nomccost", "ccoste", "codccost", Text1(7).Text, "T")
                If Sql = "" Then
                    MsgBox "Centro de coste NO encontrado: " & Text1(7).Text, vbExclamation
                    Text1(7).Text = ""
                    Text1(7).SetFocus
                End If
                Text3.Text = Sql

        Case 4, 19
                If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    Text1(Index).SetFocus
                End If
        Case 21, 22
            'Concepto de imovilizado
            Sql = ""
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Campo debe ser numérico:", vbExclamation
            Else
                If Index = 21 Then
                    Sql = DevuelveDesdeBD("nomsecin", "inmovseccion", "codsecin", Text1(Index).Text, "N")
                Else
                    Sql = DevuelveDesdeBD("nomubiin", "inmovubicacion", "codubiin", Text1(Index).Text, "N")
                End If
            End If
            Text5(Index - 21).Text = Sql
            If Sql = "" Then
                If Text1(Index).Text <> "" Then
                    Text1(Index).Text = ""
                    PonFoco Text1(Index)
                End If
            End If
        Case Else
                If Index = 8 Or Index = 12 Or Index = 13 Or Index = 14 Or Index = 15 Then
                    If InStr(1, Text1(Index).Text, ",") Then
                        'El importe ya esta formateado
                        If CadenaCurrency(Text1(Index).Text, Im) Then
                            Text1(Index).Text = Format(Im, FormatoImporte)
                        Else
                            MsgBox "importe incorrecto", vbExclamation
                            Text1(Index).Text = ""
                        End If
                    Else
                        Text1(Index).Text = TransformaPuntosComas(Text1(Index).Text)
                    End If
                End If
                
                If Index = 12 Or Index = 13 Then
                    Dim Valor As Currency
                    Valor = CCur(ComprobarCero(Text1(12).Text)) - CCur(ComprobarCero(Text1(13).Text))
                    Text4(7).Text = ""
                    'If Valor <> 0 Then
                    Text4(7).Text = Format(Valor, "###,###,##0.00")
                End If
                
                
                If miTag.Cargar(Text1(Index)) Then
                    If miTag.Comprobar(Text1(Index)) Then
                            If miTag.Formato <> "" Then
                                miTag.DarFormato Text1(Index)
                            End If
                    Else
                        Text1(Index).Text = ""
                        Text1(Index).SetFocus
                    End If
                End If
        End Select
        
    End If
End Sub

Private Sub HacerBusqueda()
    Dim cad As String
    Dim CadB As String

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
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
    Set frmEleInmo = New frmBasico2
    
    AyudaEleInmo frmEleInmo, , CadB
    
    Set frmEleInmo = Nothing
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla de elementos de inmovilizado", vbInformation
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
Dim Antmodo As Byte

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1
     
    
    'Cargamos el LINEAS
    DataGridAux(1).Enabled = False
    CargaGrid 0, True
    CargaGrid 1, True
    
    If Modo = 2 Then DataGridAux(1).Enabled = True

    'Vamos a poner los valores de los campos referencales
    Antmodo = Modo
    Modo = 6
    Text1_LostFocus (6)
    Text1_LostFocus (2)
    Text1_LostFocus (7)
    Text1_LostFocus (9)
    Text1_LostFocus (10)
    Text1_LostFocus (11)
    
    Text1_LostFocus (12)
    Text1_LostFocus (21)
    Text1_LostFocus (22)

    Modo = Antmodo
    If Modo = 2 Then lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean
    
    'ASIGNAR MODO
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
        
    
    B = (Modo = 2) Or Modo = 0
    
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = B
        If Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
       
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    
    
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    
    B = B Or Modo = 1  'o buscar
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = Not B
        Text1(I).Enabled = B
    Next I
    Check1(0).Enabled = B
    For I = 0 To Text4.Count - 1
        Text4(I).Locked = Not B
        Text4(I).Enabled = B
    Next I
    For I = 0 To Text5.Count - 1
        Text5(I).Locked = Not B
        Text5(I).Enabled = B
    Next I
    For I = 0 To 3
        Me.imgCta(I).Enabled = B
    Next I
    Text2.Locked = Not B
    Text2.Enabled = B
    Text3.Locked = Not B
    Text3.Enabled = B
    
    Combo1.Enabled = B
    Combo2.Enabled = B
    Me.imgFecha(0).Enabled = B
    Me.imgFecha(1).Enabled = B
    Me.imgCC.Enabled = B And vParam.autocoste
    Me.imgConcep.Enabled = B
    
    
    B = B Or (Modo = 5)
   
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    B = Modo > 2 Or Modo = 1
    
    'Ahora , si es superusuario podra modificar el valor del combo de situacion
    'Si no se lo forazremos nosotros
    If Combo2.Enabled Then
        If Modo = 4 Then
            If vUsu.Nivel > 1 Then Combo2.Enabled = False
        End If
    End If
    
    
    
End Sub


Private Function DatosOK() As Boolean
    Dim Rs As ADODB.Recordset
    Dim B As Boolean
    Dim Im As Currency
    'Comprobamos si tiene lineas
    'Para guardarlo en la BD esta el campo oculto
    If AdoAux(1).Recordset.EOF Then
        Text1(20).Text = 0
    Else
        Text1(20).Text = 1
    End If
    
    
    'Si el usuario es administrador entonces el valor del objeto lo tiene
    'k poner a mano
    If vUsu.Nivel > 1 Then   'Usuarios generales
        If Text1(19).Text <> "" Then
            Combo2.ListIndex = 2 'De baja
        Else
            If CCur(Text1(12).Text) <= CCur(Text1(13).Text) Then
                Combo2.ListIndex = 3
            Else
                Combo2.ListIndex = 1
            End If
        End If
    End If
    
    'Error por temas de campo 0
    
    
    B = CompForm2(Me, 1)
    
    
    If B Then
        If Text1(15).Text <> "" Then
            Im = ImporteFormateado(Text1(15).Text)
            If Im = 0 Then Text1(15).Text = ""
        End If
        
        If Modo = 4 Then
            Sql = ""
            If IsNull(Data1.Recordset!fecventa) Then
                If Text1(19).Text <> "" Then Sql = "No deberia ponerle fecha de baja. El proceso de dar de baja esta en otro punto de la aplicación"
            Else
                If Text1(19).Text = "" Then Sql = "No deberia quitarle la fecha de baja al elemento."
            End If
            If Sql <> "" Then
                Sql = Sql & vbCrLf & "¿Continuar?"
                If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then B = False
                Sql = ""
            End If
        End If
    End If
    DatosOK = B
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String
    WheelUnHook
    Select Case Button.Index
        Case 1
                BotonAnyadir
        Case 2
                BotonModificar
        Case 3
                BotonEliminar False
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
        Case 8 ' Impresion
            frmInmoEltoList.Show vbModal
        
        Case Else
        
    End Select






End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGridAux(1).Tag = "Estableciendo"
    AdoAux(1).ConnectionString = Conn
    AdoAux(1).RecordSource = MontaSQLCarga(1, Enlaza)
    AdoAux(1).CursorType = adOpenDynamic
    AdoAux(1).LockType = adLockPessimistic
    AdoAux(1).Refresh
    
    DataGridAux(1).AllowRowSizing = False
    DataGridAux(1).RowHeight = 320
    
    DataGridAux(1).Tag = "Asignando"
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-numaspre  1.- Lin aspre
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    'Claves lineas asientos predefinidos
    DataGridAux(1).Columns(0).visible = False
    DataGridAux(1).Columns(1).visible = False

    'Cuenta
    DataGridAux(1).Columns(2).Caption = "Cuenta"
    DataGridAux(1).Columns(2).Width = 1005
    
    DataGridAux(1).Columns(3).Caption = "Denominación"
    DataGridAux(1).Columns(3).Width = 2900

    If vParam.autocoste Then
        DataGridAux(1).Columns(4).Caption = "C.C."
        DataGridAux(1).Columns(4).Width = 800

        DataGridAux(1).Columns(5).Caption = "Centro de coste"
        DataGridAux(1).Columns(5).Width = 2100
    Else
        DataGridAux(1).Columns(4).visible = False
        DataGridAux(1).Columns(5).visible = False
    End If
    
    DataGridAux(1).Columns(6).Caption = "% Porce."
    DataGridAux(1).Columns(6).Width = 900
    DataGridAux(1).Columns(6).NumberFormat = "0.00"
        

    
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGridAux(1).Tag = "Fijando ancho"
        anc = 323
        txtaux(0).Left = DataGridAux(1).Left + 330
        txtaux(0).Width = DataGridAux(1).Columns(2).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = DataGridAux(1).Columns(3).Left + 90
                
        txtaux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 6
        txtaux(1).Width = DataGridAux(1).Columns(3).Width - 180
    
        txtaux(2).Left = DataGridAux(1).Columns(4).Left + 150
        txtaux(2).Width = DataGridAux(1).Columns(4).Width - 30
    
        txtaux(3).Left = DataGridAux(1).Columns(5).Left + 150
        txtaux(3).Width = DataGridAux(1).Columns(5).Width - 45

        
        'Concepto
        txtaux(4).Left = DataGridAux(1).Columns(6).Left + 150
        txtaux(4).Width = DataGridAux(1).Columns(6).Width - 45

        
        
    
       
        CadAncho = True
    End If
        
    For I = 0 To DataGridAux(1).Columns.Count - 1
            DataGridAux(1).Columns(I).AllowSizing = False
    Next I
    
    DataGridAux(1).Tag = "Calculando"

    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(1).Tag, Err.Description
End Sub


Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    ' Si ENLAZA -> Enlaza con el data1
    '           -> Si no lo cargamos sin enlazar a nngun campo
    '--------------------------------------------------------------------
    Dim Sql As String
    
    Select Case Index
        Case 0 ' historico
            Sql = "SELECT inmovele_his.codinmov,fechainm,porcinm,imporinm  "
            Sql = Sql & " FROM inmovele_his "
            If Enlaza Then
                Sql = Sql & " WHERE codinmov = " & Data1.Recordset!Codinmov
                Else
                Sql = Sql & " WHERE codinmov = -1"
            End If
            Sql = Sql & " ORDER BY fechainm desc"
    
        Case 1 ' porcentaje de cuentas
            Sql = "SELECT inmovele_rep.codinmov, inmovele_rep.numlinea, inmovele_rep.codmacta2, cuentas.nommacta,"
            Sql = Sql & " ccoste.codccost, ccoste.nomccost, inmovele_rep.porcenta"
            Sql = Sql & " FROM (inmovele_rep INNER JOIN cuentas ON inmovele_rep.codmacta2 = cuentas.codmacta) LEFT"
            Sql = Sql & " JOIN ccoste ON inmovele_rep.codccost = ccoste.codccost"
            If Enlaza Then
                Sql = Sql & " WHERE codinmov = " & Data1.Recordset!Codinmov
                Else
                Sql = Sql & " WHERE codinmov = -1"
            End If
            Sql = Sql & " ORDER BY numlinea"
    End Select

    MontaSQLCarga = Sql


End Function


Private Sub AnyadirLinea(Limpiar As Boolean, Index As Integer)
    Dim anc As Single
    
    If ModificandoLineas <> 0 Then Exit Sub
        
    Select Case Index
        Case 0 ' hco de inmovilizado
        
    
        Case 1 ' cuentas
        
             CalculaTotalLineas
             If TotalLin = 0 Then
                 MsgBox "No se pueden insertar nuevas lineas." & TotalLin, vbExclamation
                 Exit Sub
             End If
             'Obtenemos la siguiente numero de factura
             Linliapu = ObtenerSigueinteNumeroLinea(False)
             'Situamos el grid al final
             
            'Situamos el grid al final
             DataGridAux(Index).AllowAddNew = True
             If AdoAux(Index).Recordset.RecordCount > 0 Then
                 DataGridAux(Index).HoldFields
                 AdoAux(Index).Recordset.MoveLast
                 DataGridAux(Index).Row = DataGridAux(1).Row + 1
             End If
             anc = DataGridAux(Index).top
             If DataGridAux(Index).Row < 0 Then
                 anc = anc + 220
                 Else
                 anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 15
             End If
             LLamaLineas anc, 1, Limpiar, Index
             txtaux(4).Text = TotalLin
             'Ponemos el foco
             PonFoco txtaux(0)
    End Select
             
End Sub


Private Function ObtenerSigueinteNumeroLinea(Cabecera As Boolean) As Long
    Dim Rs As ADODB.Recordset
    Dim I As Long
    
    Set Rs = New ADODB.Recordset
    If Cabecera Then
        Sql = "SELECT Max(codinmov) FROM inmovele"
    Else
        Sql = "Select max(numlinea) from inmovele_rep where codinmov=" & Data1.Recordset!Codinmov
    End If
    Rs.Open Sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then I = Rs.Fields(0)
    End If
    Rs.Close
    ObtenerSigueinteNumeroLinea = I + 1
End Function


'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------

Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean, Optional Index As Integer)
    Dim B As Boolean
    DeseleccionaGrid Index
    ModificandoLineas = xModo
    B = (xModo = 0)
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    CamposAux Not B, alto, Limpiar, Index
End Sub

Private Sub CamposAux(visible As Boolean, Altura As Single, Limpiar As Boolean, Optional Index As Integer)
    Dim I As Integer
    Dim B As Boolean
    
    DeseleccionaGrid Index

    DataGridAux(Index).Enabled = Not visible
    
    Select Case Index
        Case 0 ' hco
            For I = 1 To 3
                txtaux1(I).visible = visible
                txtaux1(I).top = Altura
            Next I
            cmdAux(1).visible = visible
            cmdAux(1).top = Altura
        
            If Limpiar Then
                For I = 1 To 4
                    txtaux(I).Text = ""
                Next I
            End If
        
        
        Case 1
            For I = 2 To 4
                If I = 2 Or I = 3 Then
                    B = vParam.autocoste
                Else
                    B = True
                End If
                txtaux(I).visible = visible And B
                txtaux(I).top = Altura
            Next I
            
            txtAux2(0).visible = visible
            txtAux2(0).top = Altura
            txtAux2(1).visible = visible And B
            txtAux2(1).top = Altura
        
            cmdAux(0).visible = visible
            cmdAux(0).top = Altura
            
            cmdAux(2).visible = visible And B
            cmdAux(2).top = Altura
            
            If Limpiar Then
                For I = 0 To 4
                    txtaux(I).Text = ""
                Next I
                txtAux2(0).Text = ""
                txtAux2(1).Text = ""
            End If
    
    End Select
    
End Sub



Private Sub ToolbarAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
     Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
            If Modo = 4 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
        Case 4 ' busqueda avanzada
            frmInmoHco.Show vbModal
    End Select

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), Modo
End Sub

'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda1 KeyAscii, 0 'cuenta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda1(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub
'++

Private Sub txtAux_LostFocus(Index As Integer)
    Dim Sng As Double
        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtaux(Index).Text = Trim(txtaux(Index).Text)
    
        'Comun a todos
        If txtaux(Index).Text = "" Then
            Select Case Index
                Case 0
                    txtaux(1).Text = ""
                Case 2
                    txtaux(3).Text = ""
            End Select
            Exit Sub
        End If
        
        Select Case Index
        Case 2 'Cuenta
            RC = txtaux(2).Text
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                txtaux(2).Text = RC
                txtAux2(0).Text = Sql
                RC = ""
            Else
                MsgBox Sql, vbExclamation
                txtaux(2).Text = ""
                txtAux2(0).Text = ""
                RC = "NO"
            End If
            If RC <> "" Then PonFoco txtaux(2)
            
        Case 3 'Centro de Coste
            txtaux(3).Text = UCase(txtaux(3).Text)
            Sql = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtaux(3).Text, "T")
            If Sql = "" Then
                MsgBox "Centro de coste NO encontrado: " & txtaux(3).Text, vbExclamation
                txtaux(3).Text = ""
                PonFoco txtaux(3)
            End If
            txtAux2(1).Text = Sql
        
        Case 4
            If Not IsNumeric(txtaux(4).Text) Then
                MsgBox "Porcentaje debe ser numérico", vbExclamation
                txtaux(4).Text = ""
                PonFoco txtaux(4)
            Else
                cmdAceptar.SetFocus
            End If
        End Select
End Sub


Private Function DatosOkLin(nomframe As String) As String
    
    Select Case nomframe
        Case "FrameAux0" ' historico
        
        
        Case "FrameAux1" ' centros de reparto
             'Cuenta
             If txtaux(2).Text = "" Then
                 DatosOkLin = "Cuenta no puede estar vacia."
                 Exit Function
             End If
             
             If Not IsNumeric(txtaux(2).Text) Then
                 DatosOkLin = "Cuenta debe ser numerica"
                 Exit Function
             End If
             
             If txtAux2(0).Text = NO Then
                 DatosOkLin = "La cuenta debe estar dada de alta en el sistema"
                 Exit Function
             End If
             
             If Not EsCuentaUltimoNivel(txtaux(2).Text) Then
                 DatosOkLin = "La cuenta no es de último nivel"
                 Exit Function
             End If
                     
             'Porcentaje
             If txtaux(4).Text = "" Then
                 DatosOkLin = "Porcentaje en blanco"
                 Exit Function
             End If
            
             If Not IsNumeric(txtaux(4).Text) Then
                 DatosOkLin = "El porcentaje DEBE debe ser numérico"
                 Exit Function
             End If
             I = Val(txtaux(4).Text)
             If I < 0 Or I > 100 Then
                 DatosOkLin = "Porcentajes incorrecto"
             End If
    End Select
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        'INSERTAR LINEAS
        Sql = "INSERT INTO inmovele_rep (codinmov, numlinea, codmacta2, codccost, porcenta) VALUES (" & Data1.Recordset!Codinmov & ","
        Sql = Sql & Linliapu & ",'"
        Sql = Sql & txtaux(0).Text & "',"
        If txtaux(2).Text = "" Then
            Sql = Sql & "NULL"
        Else
            Sql = Sql & "'" & txtaux(2).Text & "'"
        End If
        Sql = Sql & "," & TransformaComasPuntos(txtaux(4).Text) & ")"
        
        
    Else
    
        'MODIFICAR
        'UPDATE asipre_lineas SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
        '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
        Sql = "UPDATE inmovele_rep SET "
        Sql = Sql & " codmacta2 = '" & txtaux(0).Text & "',"
        Sql = Sql & " codccost = "
        If txtaux(2).Text = "" Then
            Sql = Sql & "NULL"
        Else
            Sql = Sql & "'" & txtaux(2).Text & "'"
        End If
        Sql = Sql & ", porcenta = " & TransformaComasPuntos(txtaux(4).Text)
        Sql = Sql & " WHERE inmovele_rep.numlinea = " & Linliapu
        Sql = Sql & " AND inmovele_rep.codinmov =" & Data1.Recordset!Codinmov
        
    End If
    Conn.Execute Sql
    
    'Ahora actualizamos la BD para ver si tiene centro de repartos
    ActualizaRepartos
    InsertarModificar = True
    Exit Function
EInsertarModificar:
        MuestraError Err.Number, "InsertarModificar linea asiento.", Err.Description
End Function
 
Private Sub ActualizaRepartos()
    Sql = "UPDATE inmovele SET Repartos="
    If ModificandoLineas = 1 Or ModificandoLineas = 2 Then
        RC = "1"
    Else
        If AdoAux(1).Recordset.EOF Then
            RC = "0"
        Else
            RC = "1"
        End If
    End If
    Text1(20).Text = RC
    Sql = Sql & RC & " WHERE codinmov =" & Data1.Recordset!Codinmov
    Conn.Execute Sql
End Sub

Private Sub DeseleccionaGrid(Index As Integer)
    On Error GoTo EDeseleccionaGrid
        
    While DataGridAux(Index).SelBookmarks.Count > 0
        DataGridAux(Index).SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub



Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim I As Byte
Dim tots As String

    tots = MontaSQLCarga(Index, Enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez

    Select Case Index
        Case 0 'historico
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtaux1(1)|T|Fecha|1855|;S|cmdAux(1)|B|||;S|txtaux1(2)|T|Porcentaje|1855|;"
            tots = tots & "S|txtaux1(3)|T|Importe|2105|;"

            arregla tots, DataGridAux(Index), Me

            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            CalculaTotales
        
        Case 1 'Cuentas de reparto
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codforfait
            tots = tots & "S|txtaux(2)|T|Cuenta|1405|;S|cmdAux(0)|B|||;"
            tots = tots & "S|txtaux2(0)|T|Denominación|3995|;"
            If vParam.autocoste Then
                tots = tots & "S|txtaux(3)|T|C.C|800|;S|cmdAux(2)|B|||;S|txtaux2(1)|T|Centro de Coste|3500|;"
            Else
                tots = tots & "N||||0|;N||||0|;"
            End If
            tots = tots & "S|txtaux(4)|T|%Reparto|1300|;"
            
            arregla tots, DataGridAux(Index), Me
        
        
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description


End Sub

Private Sub CalculaTotales()
Dim Sql As String
Dim VAdqui As Currency
Dim VAmor As Currency
Dim VPdte As Currency

    Text4(4).Text = ""
    Text4(5).Text = ""
    Text4(6).Text = ""
    FrameTotales.visible = False
    
    If Data1.Recordset.EOF Then Exit Sub
    
    FrameTotales.visible = (Combo2.ListIndex = 0)

    Sql = "select sum(imporinm) from inmovele_his where codinmov = " & DBSet(Text1(0).Text, "N")
    VAdqui = CCur(ComprobarCero(Text1(12).Text))
    VAmor = DevuelveValor(Sql)
    VPdte = VAdqui - VAmor
    
    Text4(4).Text = Format(VAdqui, "#,###,##0.00")
    Text4(5).Text = Format(VAmor, "#,###,##0.00")
    Text4(6).Text = Format(VPdte, "#,###,##0.00")
    
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub LLamaCuenta()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0"
    frmC.ConfigurarBalances = 3  'nuevo
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub CalculaTotalLineas()
    TotalLin = 0
    Sql = "Select Sum(porcenta) from inmovele_rep where codinmov=" & Data1.Recordset!Codinmov
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then _
            TotalLin = miRsAux.Fields(0)
    End If
    miRsAux.Close
    TotalLin = 100 - TotalLin
    Set miRsAux = Nothing
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'historico
        Case 1: nomframe = "FrameAux1" 'centros de reparto
    End Select
    ' ***************************************************************
    
    If DatosOkLin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If B Then AnyadirLinea True, NumTabMto
            End Select
           
'            SituarTab (NumTabMto)
            
            
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim v As Integer
Dim cad As String
Dim TablaAux As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'historico
        Case 1: nomframe = "FrameAux1" 'reparto de cuentas
    End Select
    ' **************************************************************

    If DatosOkLin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            
            Select Case NumTabMto
                Case 0: TablaAux = "inmovele_his" 'historico
                Case 1: TablaAux = "inmovele_rep" 'reparto de cuentas
            End Select
    
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModoLineas = 0

            v = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True
                
            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & v)
            ' ***********************************************************

            LLamaLineas 0, ModoLineas, True, NumTabMto
            
        End If
    End If
        
End Sub


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo2.Clear
    
    'tipo de situacion del elemento
    Sql = "select situacio, descsituacion from usuarios.wcontiposituinmo order by situacio"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not Rs.EOF
        Combo2.AddItem DBLet(Rs.Fields(1).Value, "T")
        Combo2.ItemData(Combo2.NewIndex) = DBLet(Rs.Fields(0).Value, "N")
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    

End Sub



Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
Dim SumLin As Currency

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    'If ModificaLineas = 2 Then Exit Sub
    ModoLineas = 1 'Ponemos Modo Añadir Linea
    
    ModificandoLineas = 0
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modifcar Cabecera
        cmdAceptar_Click
        'No se ha insertado la cabecera
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5

    'Obtener el numero de linea ha insertar
    Select Case Index
        Case 0: vTabla = "inmovele_his"
        Case 1: vTabla = "inmovele_rep"
    End Select


    DataGridAux(Index).AllowAddNew = True
    If AdoAux(Index).Recordset.RecordCount > 0 Then
        DataGridAux(Index).HoldFields
        AdoAux(Index).Recordset.MoveLast
        DataGridAux(Index).Row = DataGridAux(Index).Row + 1
    End If
    DataGridAux(Index).Enabled = False


    anc = DataGridAux(Index).top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 275 '248
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 8 '15
    End If


    LLamaLineas anc, ModoLineas, True, NumTabMto

    Select Case Index
        Case 0 'historico
            txtaux1(0).Text = Text1(0).Text 'elemento
            For I = 1 To 3
                txtaux1(I).Text = ""
            Next I
            PonFoco txtaux1(1)
            
        Case 1 'cuentas de reparto
            'Obtener el sig. nº de linea a insertar
            vWhere = Replace(ObtenerWhereCab(False), "inmovele", vTabla)
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
            
            txtaux(0).Text = Text1(0).Text
            txtaux(1).Text = NumF
            
            For I = 2 To 4
                txtaux(I).Text = ""
            Next I
            txtAux2(0).Text = ""
            txtAux2(1).Text = ""
            
            PonFoco txtaux(2)
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
    
    ModificandoLineas = 0
    
    If Modo = 4 Then 'Modificar Cabecera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
    
    NumTabMto = Index
    PonerModo 5
    
    If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
        I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
        DataGridAux(Index).Scroll 0, I
        DataGridAux(Index).Refresh
    End If
      
    anc = DataGridAux(Index).top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 245
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 10 '+ 15
    End If

    Select Case Index
        Case 0 'hco
            For J = 0 To 3
                txtaux1(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
        
        Case 1 'cuentas de reparto
            txtaux(0).Text = DataGridAux(Index).Columns(0).Text
            txtaux(1).Text = DataGridAux(Index).Columns(1).Text
            txtaux(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux2(0).Text = DataGridAux(Index).Columns(3).Text
            txtaux(3).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(1).Text = DataGridAux(Index).Columns(5).Text
            txtaux(4).Text = DataGridAux(Index).Columns(6).Text
    End Select
    
    LLamaLineas anc, ModoLineas, False, Index
   
    Select Case Index
        Case 0 'hco
            PonFoco txtaux1(1)
        Case 1 'cuentas de reparto
            PonFoco txtaux(2)
    End Select

End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5

    Eliminar = False

    Select Case Index
        Case 0 'lineas del hco
            Sql = "¿Seguro que desea eliminar la Amortización?"
            Sql = Sql & vbCrLf & "Fecha: " & Format(DBLet(AdoAux(Index).Recordset!fechainm), "dd/mm/yyyy")
            Sql = Sql & vbCrLf & "Importe: " & Format(DBLet(AdoAux(Index).Recordset!imporinm), "###,###,##0.00")
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                Eliminar = True
                Sql = "DELETE FROM inmovele_his"
                Sql = Sql & Replace(ObtenerWhereCab(True), "inmovele", "inmovele_his") & " AND fechainm= " & DBSet(AdoAux(Index).Recordset!fechainm, "F")
            End If
        Case 1 'cuentas de reparto
            Sql = "¿Seguro que desea eliminar la Cuenta?"
            Sql = Sql & vbCrLf & "Cuenta: " & DBLet(AdoAux(Index).Recordset!codmacta2)
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                Eliminar = True
                Sql = "DELETE FROM inmovele_rep"
                Sql = Sql & Replace(ObtenerWhereCab(True), "inmovele", "inmovele_rep") & " AND numlinea= " & DBSet(AdoAux(Index).Recordset!NumLinea, "N")
            End If
        
    End Select

    If Eliminar Then
        TerminaBloquear
        Conn.Execute Sql
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        End If
       
        
        
        'antes estaba debajo de situardata
        CargaGrid Index, True
        SituarDataTrasEliminar AdoAux(Index), NumRegElim, True
        
        
        
    End If

    ModoLineas = 0
    
    cmdCancelar_Click
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    vWhere = ""
    If conW Then vWhere = " WHERE "
    vWhere = vWhere & "codinmov=" & Trim(Text1(0).Text)
    ObtenerWhereCab = vWhere
End Function

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
Dim I As Integer
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N")
        
        For I = 0 To ToolbarAux.Count - 1
            ToolbarAux(I).Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
            ToolbarAux(I).Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
            ToolbarAux(I).Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
            
            If Modo = 2 Then
                If AdoAux(I).Recordset.EOF Then
                    ToolbarAux(I).Buttons(2).Enabled = False
                    ToolbarAux(I).Buttons(3).Enabled = False
                End If
            End If
        Next I
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub txtaux1_GotFocus(Index As Integer)
    ConseguirFoco txtaux1(Index), Modo
End Sub

'++
Private Sub txtaux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda2 KeyAscii, 0 'fecha
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda2(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub
'++

Private Sub txtAux1_LostFocus(Index As Integer)
    Dim Sng As Double
        
    If ModificandoLineas = 0 Then Exit Sub
    
    'Comprobaremos ciertos valores
    txtaux1(Index).Text = Trim(txtaux1(Index).Text)

    'Comun a todos
    If txtaux1(Index).Text = "" Then
        Select Case Index
            Case 0
                txtaux1(1).Text = ""
            Case 2
                txtaux1(3).Text = ""
        End Select
        Exit Sub
    End If
    
    Select Case Index
        Case 1 ' fecha de inmovilizado
            PonerFormatoFecha txtaux1(1)
        Case 2 ' porcentaje de inmovilizado
            PonerFormatoDecimal txtaux1(Index), 4
        Case 3 ' importe de amortizado
            PonerFormatoDecimal txtaux1(Index), 1
            
    End Select

End Sub

