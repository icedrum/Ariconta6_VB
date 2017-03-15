VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTESPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15840
   Icon            =   "frmTESPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   15840
   StartUpPosition =   2  'CenterScreen
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
      Left            =   14610
      TabIndex        =   98
      Top             =   9300
      Visible         =   0   'False
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
      Left            =   13470
      TabIndex        =   96
      Top             =   9300
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
      Left            =   14610
      TabIndex        =   97
      Top             =   9300
      Width           =   1035
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   8490
      TabIndex        =   61
      Top             =   30
      Width           =   3255
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmTESPagos.frx":000C
         Left            =   90
         List            =   "frmTESPagos.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   210
         Width           =   3075
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   58
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   59
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
      Left            =   5880
      TabIndex        =   56
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   57
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
      Left            =   12630
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3810
      TabIndex        =   54
      Top             =   30
      Width           =   1965
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   55
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Datos Fiscales"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Dividir Vencimiento"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Pago"
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4260
      Top             =   8340
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   90
      TabIndex        =   33
      Top             =   1740
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   12991
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Datos Pago"
      TabPicture(0)   =   "frmTESPagos.frx":0050
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgCuentas(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgCuentas(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgCuentas(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(33)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(34)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(35)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgFecha(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgFecha(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "imgppal(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(10)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(17)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(16)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(32)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(20)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Combo1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(19)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(26)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(21)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(30)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(29)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(28)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text2(0)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(4)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text2(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(5)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(6)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text2(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(9)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text2(3)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(10)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(17)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "frameContene"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtPendiente"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(12)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(11)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text1(7)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(8)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "FrameRemesa"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "SSTab2"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "FrameDatosFiscales"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cboSituRem"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text1(31)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).ControlCount=   52
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
         Index           =   31
         Left            =   13830
         MaxLength       =   4
         TabIndex        =   108
         Tag             =   "Año remesa|N|S|0||pagos|anyodocum|||"
         Text            =   "Text"
         Top             =   3420
         Width           =   1395
      End
      Begin VB.ComboBox cboSituRem 
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
         ItemData        =   "frmTESPagos.frx":006C
         Left            =   9780
         List            =   "frmTESPagos.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   4200
         Width           =   2805
      End
      Begin VB.Frame FrameDatosFiscales 
         Caption         =   "DATOS FISCALES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         Left            =   300
         TabIndex        =   80
         Top             =   1080
         Visible         =   0   'False
         Width           =   9375
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
            Index           =   25
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "Text4"
            Top             =   2250
            Width           =   3345
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
            Left            =   1470
            TabIndex        =   86
            Tag             =   "País|T|S|||pagos|codpais|||"
            Top             =   2250
            Width           =   465
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
            Left            =   1470
            TabIndex        =   87
            Tag             =   "Nif|T|S|||pagos|nifprove|||"
            Top             =   2820
            Width           =   1350
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
            Left            =   1470
            TabIndex        =   85
            Tag             =   "Provincia|T|S|||pagos|proprove|||"
            Top             =   1770
            Width           =   7800
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
            Left            =   4020
            TabIndex        =   84
            Tag             =   "Poblacion|T|S|||pagos|pobprove|||"
            Top             =   1320
            Width           =   5250
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
            Left            =   1470
            TabIndex        =   83
            Tag             =   "CP|T|S|||pagos|cpprove|||"
            Top             =   1290
            Width           =   1320
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
            Left            =   1470
            TabIndex        =   82
            Tag             =   "Dirección|T|S|||pagos|domprove|||"
            Top             =   840
            Width           =   7800
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
            Left            =   1470
            TabIndex        =   81
            Tag             =   "Nombre|T|S|||pagos|nomprove|||"
            Top             =   390
            Width           =   7800
         End
         Begin VB.Image imgppal 
            Height          =   240
            Index           =   1
            Left            =   1170
            Top             =   2280
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "País"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   28
            Left            =   300
            TabIndex        =   95
            Top             =   2310
            Width           =   555
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   29
            Left            =   300
            TabIndex        =   94
            Top             =   1830
            Width           =   1395
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   30
            Left            =   3000
            TabIndex        =   93
            Top             =   1350
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   31
            Left            =   300
            TabIndex        =   92
            Top             =   2880
            Width           =   1065
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   32
            Left            =   300
            TabIndex        =   91
            Top             =   1350
            Width           =   855
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   36
            Left            =   300
            TabIndex        =   90
            Top             =   900
            Width           =   1545
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   37
            Left            =   300
            TabIndex        =   89
            Top             =   450
            Width           =   1545
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2475
         Left            =   300
         TabIndex        =   63
         Top             =   4680
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4366
         _Version        =   393216
         Tabs            =   1
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
         TabCaption(0)   =   "Pagos Realizados"
         TabPicture(0)   =   "frmTESPagos.frx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FrameAux0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame FrameAux0 
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
            Height          =   2085
            Left            =   150
            TabIndex        =   64
            Top             =   330
            Width           =   14475
            Begin VB.TextBox txtaux 
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
               Index           =   14
               Left            =   12000
               TabIndex        =   73
               Tag             =   "Banco Talon/Pag|T|S|||pagos_realizados|bancotalonpag|||"
               Text            =   "Banco"
               Top             =   1050
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txtaux 
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
               Index           =   13
               Left            =   11370
               TabIndex        =   72
               Tag             =   "Ref. Talon/Pag|T|S|||pagos_realizados|reftalonpag|||"
               Text            =   "ref"
               Top             =   1050
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox txtaux 
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
               Index           =   12
               Left            =   6870
               TabIndex        =   68
               Tag             =   "Cta.Real Cobro|T|S|||pagos_realizados|ctabanc2|||"
               Text            =   "cta.cobro"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Frame FrameToolAux 
               Height          =   555
               Index           =   0
               Left            =   0
               TabIndex        =   101
               Top             =   0
               Width           =   975
               Begin MSComctlLib.Toolbar ToolbarAux 
                  Height          =   330
                  Left            =   180
                  TabIndex        =   102
                  Top             =   150
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   582
                  ButtonWidth     =   609
                  ButtonHeight    =   582
                  Style           =   1
                  _Version        =   393216
                  BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                     NumButtons      =   2
                     BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Object.ToolTipText     =   "Ver Asiento"
                     EndProperty
                     BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Enabled         =   0   'False
                        Object.Visible         =   0   'False
                        Object.ToolTipText     =   "Imprimir Recibo"
                     EndProperty
                  EndProperty
               End
            End
            Begin VB.TextBox txtaux 
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
               Height          =   315
               Index           =   11
               Left            =   9870
               TabIndex        =   70
               Tag             =   "F.Realizado|F|S|||pagos_realizados|fecrealizado|dd/mm/yyyy||"
               Text            =   "FRealizado"
               Top             =   1080
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox txtaux 
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
               Height          =   315
               Index           =   10
               Left            =   10620
               TabIndex        =   71
               Tag             =   "Tipo|T|S|||tipofpago|siglas|||"
               Text            =   "T FP"
               Top             =   1080
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.CommandButton cmdAux 
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
               Height          =   315
               Index           =   1
               Left            =   5400
               TabIndex        =   100
               ToolTipText     =   "Buscar cuenta"
               Top             =   1110
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.CommandButton cmdAux 
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
               Height          =   315
               Index           =   0
               Left            =   4020
               TabIndex        =   99
               ToolTipText     =   "Buscar cuenta"
               Top             =   1110
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.TextBox txtaux 
               Alignment       =   1  'Right Justify
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
               Index           =   9
               Left            =   12810
               TabIndex        =   74
               Tag             =   "Importe Cobro|T|S|||pagos_realizados|impcobro|###,##0.00||"
               Text            =   "importe"
               Top             =   1050
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.TextBox txtaux 
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
               Index           =   0
               Left            =   720
               MaxLength       =   3
               TabIndex        =   79
               Tag             =   "Serie|T|N|||pagos_devolucion|numserie||S|"
               Text            =   "ser"
               Top             =   1090
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtaux 
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
               Index           =   4
               Left            =   2340
               MaxLength       =   4
               TabIndex        =   78
               Tag             =   "Linea|N|N|0||pagos_devolucion|numlinea||S|"
               Text            =   "lin"
               Top             =   1090
               Visible         =   0   'False
               Width           =   390
            End
            Begin VB.TextBox txtaux 
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
               Left            =   1125
               MaxLength       =   10
               TabIndex        =   77
               Tag             =   "Nº Factura|N|N|||pagos_devolucion|numfactul|000000|S|"
               Text            =   "fac"
               Top             =   1090
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtaux 
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
               Left            =   1515
               MaxLength       =   10
               TabIndex        =   76
               Tag             =   "Fecha Factura|F|N|||pagos_devolucion|fecfactu|dd/mm/yyyy|S|"
               Text            =   "Fec"
               Top             =   1090
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtaux 
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
               Left            =   1905
               MaxLength       =   30
               TabIndex        =   75
               Tag             =   "Nº Vencimiento|N|N|0||pagos_devolucion|numorden||S|"
               Text            =   "vto"
               Top             =   1090
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtaux 
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
               Index           =   8
               Left            =   8250
               TabIndex        =   69
               Tag             =   "Usuario Cobro|T|S|||pagos_realizados|usuariocobro|||"
               Text            =   "usuario"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.TextBox txtaux 
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
               Index           =   7
               Left            =   5550
               TabIndex        =   67
               Tag             =   "Nº asiento|N|S|0||pagos_realizados|numasien|00000000||"
               Text            =   "asiento"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.TextBox txtaux 
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
               Index           =   6
               Left            =   4200
               TabIndex        =   66
               Tag             =   "Fecha entrada|F|S|||pagos_realizados|fechaent|dd/mm/yyyy||"
               Text            =   "fecha"
               Top             =   1090
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.TextBox txtaux 
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
               Index           =   5
               Left            =   2850
               TabIndex        =   65
               Tag             =   "Diario|N|S|0||pagos_realizados|numdiari|00||"
               Text            =   "diario"
               Top             =   1090
               Visible         =   0   'False
               Width           =   1275
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   0
               Left            =   1830
               Top             =   120
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
               Caption         =   "AdoAux(0)"
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
            Begin MSComctlLib.ListView lwpagos 
               Height          =   1425
               Left            =   -30
               TabIndex        =   105
               Top             =   570
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   2514
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Tipo"
                  Object.Width           =   1410
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Factura"
                  Object.Width           =   2116
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Fecha"
                  Object.Width           =   2381
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Vto"
                  Object.Width           =   1234
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Fecha Vto"
                  Object.Width           =   2381
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Forma pago"
                  Object.Width           =   6174
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   6
                  Text            =   "Importe"
                  Object.Width           =   3590
               EndProperty
            End
         End
      End
      Begin VB.Frame FrameRemesa 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   525
         Left            =   9690
         TabIndex        =   43
         Top             =   3360
         Width           =   3015
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
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   22
            Tag             =   "Nro Documento|N|S|0||pagos|nrodocum|0000000000||"
            Text            =   "Text1"
            Top             =   30
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   19
            Left            =   60
            TabIndex        =   103
            Top             =   90
            Width           =   1260
         End
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
         Left            =   13830
         MaxLength       =   30
         TabIndex        =   20
         Tag             =   "Importe|N|S|||pagos|imppagad|#,##0.00||"
         Text            =   "1.999.999.00"
         Top             =   2130
         Width           =   1455
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
         Index           =   7
         Left            =   13980
         TabIndex        =   18
         Tag             =   "Fecha ult. pago|F|S|||pagos|fecultpa|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1380
         Width           =   1305
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
         Left            =   360
         MaxLength       =   80
         TabIndex        =   13
         Tag             =   "CSB|T|S|||pagos|text1csb|||"
         Text            =   "WWW4567890WWW4567890WWW4567890WWW456789WWWW4567890WWW4567890WWW4567890WWW456789J"
         Top             =   2790
         Width           =   9225
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
         Left            =   360
         MaxLength       =   60
         TabIndex        =   14
         Tag             =   "T|T|S|||pagos|text2csb|||"
         Top             =   3420
         Width           =   9225
      End
      Begin VB.TextBox txtPendiente 
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
         Left            =   13830
         TabIndex        =   47
         Text            =   "Text4"
         Top             =   2850
         Width           =   1425
      End
      Begin VB.Frame frameContene 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   12900
         TabIndex        =   45
         Top             =   4200
         Width           =   2235
         Begin VB.CheckBox Check1 
            Caption         =   "Documento emitido"
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
            Left            =   60
            TabIndex        =   21
            Tag             =   "Emitido|N|S|||pagos|emitdocum|||"
            Top             =   90
            Width           =   2505
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
         Height          =   495
         Index           =   17
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Tag             =   "obs|T|S|||pagos|observa|||"
         Text            =   "frmTESPagos.frx":008C
         Top             =   4080
         Width           =   9225
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
         Index           =   10
         Left            =   360
         TabIndex        =   6
         Text            =   "0000"
         Top             =   1410
         Width           =   675
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
         Index           =   3
         Left            =   7410
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   720
         Width           =   2235
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
         Index           =   9
         Left            =   360
         TabIndex        =   12
         Tag             =   "Cta prevista|T|N|||pagos|ctabanc1|||"
         Text            =   "Text1"
         Top             =   2130
         Width           =   1350
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
         Index           =   2
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   2130
         Width           =   5925
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
         Index           =   6
         Left            =   10950
         MaxLength       =   30
         TabIndex        =   19
         Tag             =   "Importe|N|N|||pagos|impefect|#,###,##0.00||"
         Text            =   "1.999.999.00"
         Top             =   2130
         Width           =   1455
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
         Index           =   5
         Left            =   11160
         TabIndex        =   17
         Tag             =   "Fecha Efecto|F|N|||pagos|fecefect|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1380
         Width           =   1245
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
         Index           =   0
         Left            =   9720
         TabIndex        =   16
         Tag             =   "Forma Pago|N|N|0||pagos|codforpa|000||"
         Text            =   "Text1"
         Top             =   720
         Width           =   735
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
         Index           =   1
         Left            =   10500
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   720
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
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Tag             =   "Cta.proveedor|T|N|||pagos|codmacta||S|"
         Text            =   "Text1"
         Top             =   720
         Width           =   1350
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
         Index           =   0
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   720
         Width           =   5625
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
         Index           =   28
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "0000"
         Top             =   1410
         Width           =   675
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
         Index           =   29
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "0000"
         Top             =   1410
         Width           =   675
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
         Index           =   30
         Left            =   3420
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "0000"
         Top             =   1410
         Width           =   675
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
         Index           =   21
         Left            =   4170
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "0000"
         Top             =   1410
         Width           =   675
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
         Index           =   26
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "0000"
         Top             =   1410
         Width           =   675
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
         Height          =   285
         Index           =   19
         Left            =   2370
         MaxLength       =   40
         TabIndex        =   24
         Tag             =   "Iban|T|S|||pagos|iban|||"
         Text            =   "ES99"
         Top             =   780
         Width           =   3795
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmTESPagos.frx":0092
         Left            =   5190
         List            =   "frmTESPagos.frx":009F
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Tag             =   "Situación|N|N|||pagos|situacion|||"
         Top             =   720
         Width           =   2175
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
         Left            =   13860
         MaxLength       =   30
         TabIndex        =   106
         Tag             =   "Usuario|N|S|||pagos|codusu|####0||"
         Top             =   2850
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
         Height          =   285
         Index           =   32
         Left            =   11130
         MaxLength       =   4
         TabIndex        =   111
         Tag             =   "Situacion|T|S|||pagos|situdocum|||"
         Text            =   "Text1"
         Top             =   4200
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
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
         Left            =   13020
         TabIndex        =   110
         Top             =   3450
         Width           =   540
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
         Height          =   195
         Index           =   17
         Left            =   9720
         TabIndex        =   109
         Top             =   3900
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   10
         Left            =   7410
         TabIndex        =   104
         Top             =   420
         Width           =   915
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   1890
         Top             =   3810
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   13710
         Picture         =   "frmTESPagos.frx":00D6
         Top             =   1410
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   10890
         Picture         =   "frmTESPagos.frx":0161
         Top             =   1410
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pagado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   12510
         TabIndex        =   53
         Top             =   2130
         Width           =   1200
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   12510
         TabIndex        =   52
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Linea2 SEPA"
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
         Index           =   35
         Left            =   360
         TabIndex        =   51
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Linea1 SEPA"
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
         Index           =   34
         Left            =   360
         TabIndex        =   50
         Top             =   2550
         Width           =   1335
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
         Height          =   195
         Index           =   33
         Left            =   390
         TabIndex        =   49
         Top             =   1140
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "Pendiente"
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
         Left            =   12510
         TabIndex        =   48
         Top             =   2880
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Observaciones"
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
         Left            =   360
         TabIndex        =   44
         Top             =   3810
         Width           =   1455
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   2700
         Top             =   1860
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Prevista Pago"
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
         Left            =   360
         TabIndex        =   42
         Top             =   1860
         Width           =   2130
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
         Height          =   285
         Index           =   6
         Left            =   9720
         TabIndex        =   41
         Top             =   2130
         Width           =   780
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
         Height          =   195
         Index           =   5
         Left            =   9720
         TabIndex        =   40
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   11250
         Top             =   435
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
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
         Index           =   0
         Left            =   9750
         TabIndex        =   39
         Top             =   420
         Width           =   1470
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   2190
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Proveedor"
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
         Left            =   360
         TabIndex        =   38
         Top             =   420
         Width           =   1770
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   26
      Top             =   9150
      Width           =   4095
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
         Top             =   210
         Width           =   3675
      End
   End
   Begin VB.Frame FrameClaves 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   765
      Left            =   120
      TabIndex        =   28
      Top             =   870
      Width           =   15555
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
         Left            =   8460
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Referencia|T|S|0||pagos|referencia|||"
         Text            =   "Text1"
         Top             =   270
         Width           =   2145
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
         Index           =   13
         Left            =   360
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Serie|T|N|||pagos|numserie||S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   765
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
         Index           =   1
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Nº Factura|T|N|||pagos|numfactu||S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   1335
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
         Index           =   3
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Nº Vencimiento|N|N|0||pagos|numorden||S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   1125
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
         Index           =   2
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||pagos|fecfactu|dd/mm/yyyy|S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   1275
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3780
         Picture         =   "frmTESPagos.frx":01EC
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Left            =   900
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia"
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
         Index           =   20
         Left            =   8460
         TabIndex        =   46
         Top             =   0
         Width           =   1710
      End
      Begin VB.Label Label1 
         Caption         =   "Serie"
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
         Left            =   360
         TabIndex        =   32
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Nº  Factura"
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
         Left            =   1260
         TabIndex        =   31
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Vencimiento"
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
         Left            =   4170
         TabIndex        =   30
         Top             =   30
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "F.Factura"
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
         Index           =   4
         Left            =   2760
         TabIndex        =   29
         Top             =   0
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   15060
      TabIndex        =   60
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
   Begin VB.Menu mnOpciones 
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmTESPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 801


Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmDpto As frmBasico
Attribute frmDpto.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmF As frmFormaPago
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmS As frmBasico
Attribute frmS.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmPais As frmBasico2
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmDev As frmBasico
Attribute frmDev.VB_VarHelpID = -1

Private frmAsi As frmAsientosHco
Attribute frmAsi.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private DevfrmCCtas As String

'NUEVO: DICIEMBRE 2005. PARA BUSCAR POR CHECKS TB
'------------------------------------------------
Private BuscaChekc As String

Dim PrimeraVez As Boolean

Dim Indice As Byte

Dim CadB As String
Dim CadB1 As String
Dim cadFiltro As String
Dim ModoLineas As Byte
Dim NumTabMto As Integer
Dim PosicionGrid As Integer

Dim vTipForpa As String
Dim CtaAnt As String


Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar



Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    If Modo = 0 Then Exit Sub
    
    LimpiarCampos
    CargaList 0, False 'CargaGrid 0, False
    
    HacerBusqueda2
End Sub



Private Sub cboSituRem_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboSituRem_Validate(Cancel As Boolean)
    If (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        If cboSituRem.ListIndex = 0 Then
            Text1(32).Text = ""
        Else
            If cboSituRem.ListIndex <> -1 Then Text1(32).Text = Chr(cboSituRem.ItemData(cboSituRem.ListIndex))
        End If
    End If

End Sub
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
    Dim Cad As String
    Dim i As Integer
    Dim Clave As String
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 1
        HacerBusqueda
    
    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm2(Me, 1) Then
                Data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCab(True) & Ordenacion
                PosicionarData
                PonerCampos
            End If
        End If
    
    Case 4
        'Modificar
        If DatosOK Then
            Clave = "numserie = " & DBSet(Data1.Recordset!NUmSerie, "T") & " AND codmacta =" & DBSet(Data1.Recordset!codmacta, "T") ' codmacta numfactu
            Clave = Clave & " AND fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F") & " AND numorden =" & DBSet(Data1.Recordset!numorden, "N")  ' codmacta numfactu fecfactu numorden
            Clave = Clave & " AND numfactu = " & DBSet(Data1.Recordset!NumFactu, "T")
            
            '       If ModificaDesdeFormulario2(Me, 1) Then
            If ModificaDesdeFormularioClaves2(Me, 1, "", Clave) Then
                'TerminaBloquear
                DesBloqueaRegistroForm Me.Text1(0)
                lblIndicador.Caption = ""
                If SituarData Then
                
                    Text1_LostFocus 0
                    Cad = vTipForpa 'para que no pierda el valor
                    PonerModo 2
                    vTipForpa = Cad
                    Cad = ""
                    PonPendiente
                    '-- Esto permanece para saber donde estamos
                    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

                Else
                    LimpiarCampos
                    'PonerModo 0
                End If
            End If
        End If
    
    
    
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub


Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 ' cuenta base
            cmdAux(0).Tag = 1
            
            Set frmDia = New frmTiposDiario
            frmDia.Show vbModal
            Set frmDia = Nothing
            
            PonFoco txtAux(5)
            
            
    End Select
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        Case 4
            'Modificar
            lblIndicador.Caption = ""
            'TerminaBloquear
            DesBloqueaRegistroForm Me.Text1(0)
            PonerModo 2
            PonerCampos
        Case 5 'LLÍNIES
    End Select
End Sub

Private Function SituarData() As Boolean
    Dim posicion As Long
    Dim Sql As String
    On Error GoTo ESituarData1
        SituarData = False
                    
        With Data1
            'Vemos poscion
            posicion = .Recordset.AbsolutePosition - 1
            'Actualizamos el recordset
            .Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            .Recordset.MoveFirst
            
            If .Recordset.RecordCount <= posicion Then
                'Era el utlimo
                .Recordset.MoveLast
            Else
                If posicion > 0 Then .Recordset.Move posicion
            End If
            SituarData = True
        End With
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    
    CargaList 0, False 'CargaGrid 0, False
    
    
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    
    Combo1.ListIndex = 0
    Text2(3).Text = Combo1.Text
    
    'añadimos el codusu
    Text1(20).Text = vUsu.Id
    
    cboSituRem.ListIndex = -1
    '###A mano
    PonFoco Text1(13)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        CargaList 0, False
        
        lblIndicador.Caption = "Búsqueda"
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarillo
        PonFoco Text1(13)
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
    
    CadB = ""
    CadB1 = ""
    
    LimpiarCampos
    CargaList 0, False
    
    HacerBusqueda2
    
    
End Sub

Private Sub Desplazamiento(Index As Integer)
    If Data1.Recordset.EOF Then Exit Sub
    Select Case Index
        Case 1
            Data1.Recordset.MoveFirst
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
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
Dim N As Byte
Dim BloquearClave As Boolean

    N = SePuedeEliminar()
    If N = 0 Then Exit Sub
    
    If PertenceAlgunoDocumentoEmitido Then Exit Sub

    If Not BloqueaRegistroForm(Me) Then Exit Sub
    '---------
    'MODIFICAR
    '----------
    
    'Si se puede modificar entonces habilito todooos los campos
    PonerModo 4
    
    'Si tiene algun pago hehco, NO puede modiciar ningun campo de la clave ppal
    BloquearClave = False
    If DBLet(Data1.Recordset!imppagad, "N") <> 0 Then BloquearClave = True
        'Tiene importes pagados. NO dejo cambiar la clave
        'numserie codmacta numfactu fecfactu numorden
    
    '13 1 2 3 4
    BloqueaTXT Me.Text1(1), BloquearClave
    BloqueaTXT Me.Text1(2), BloquearClave
    BloqueaTXT Me.Text1(3), BloquearClave
    BloqueaTXT Me.Text1(4), BloquearClave
    BloqueaTXT Me.Text1(13), BloquearClave
        
        
        
        
        
    
    
    
    If N < 3 Then
        'Se puede modifcar la CC
        Dim T As TextBox
        For Each T In Text1
            If T.Index = 10 Or T.Index = 26 Or T.Index = 28 Or T.Index = 29 Or T.Index = 30 Or T.Index = 21 Then
                T.Locked = True
                T.BackColor = &H80000018
            End If
        Next T
        'Tabbien dejamos modificar el IBAN
        Text1(19).Locked = False
        Text1(19).BackColor = vbWhite
        'Pongo visible false los img
         For N = 0 To 6
            If N < 4 And N <> 3 Then imgCuentas(N).Visible = False
            Me.imgFecha(N).Visible = False
         Next N
        
        
        'Si es una remesa de talon/pagare tb dejare modificar el numero de talon pagare
        If Val(DBLet(Data1.Recordset!Tiporem)) > 1 Then
            Text1(27).Locked = False
            Text1(27).BackColor = vbWhite
        End If
            
        PonerFoco Text1(10)
    Else
        PonerFoco Text1(6)
    End If
    
    
    'Si no tienen permisos NO permito modificar
    If vParamT.TieneOperacionesAseguradas Then
        If vUsu.Nivel >= 1 Then Me.SSTab2.TabEnabled(2) = False  'FrameSeguro.Enabled = False
    End If
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    
End Sub

Private Sub BotonEliminar()
    Dim Cad As String
    Dim i As Integer
    Dim Sql As String
    Dim SqlLog As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'Comprobamos si se puede eliminar
    If Not SePuedeEliminar Then Exit Sub
    If PertenceAlgunoDocumentoEmitido Then Exit Sub
    
    
    '### a mano
    Cad = "Seguro que desea eliminar de la BD el registro actual:"
    Cad = Cad & vbCrLf & Data1.Recordset.Fields(0) & "  " & Data1.Recordset.Fields(1) & " "
    Cad = Cad & Data1.Recordset.Fields(2) & "  " & Data1.Recordset.Fields(3) & "  " & Data1.Recordset.Fields(4)
    i = MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton2)
    'Borramos
    If i = vbYes Then
        'Borro el elemento
        Sql = "Delete from pagos  WHERE numserie = '" & Data1.Recordset!NUmSerie & "' AND numfactu = " & DBSet(Data1.Recordset!NumFactu, "T")
        Sql = Sql & " AND fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F") & " AND numorden =" & Data1.Recordset!numorden
        Sql = Sql & " and codmacta = " & DBSet(Text1(4).Text, "T")
        NumRegElim = Data1.Recordset.AbsolutePosition
        Conn.Execute Sql


        SqlLog = "Serie      : " & Text1(13).Text
        SqlLog = SqlLog & vbCrLf & "Factura    : " & Text1(1).Text
        SqlLog = SqlLog & vbCrLf & "Fecha      : " & Text1(2).Text
        SqlLog = SqlLog & vbCrLf & "Vencimiento: " & Text1(3).Text & vbCrLf
        SqlLog = SqlLog & vbCrLf & "Proveedor  : " & Text1(4).Text & " " & Text2(0).Text
        SqlLog = SqlLog & vbCrLf & "Importe    : " & Text1(6) & vbCrLf
        
        vLog.Insertar 24, vUsu, SqlLog

        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
            CargaList 0, False
            PonerModo 0
            Else
                Data1.Recordset.MoveFirst
                NumRegElim = NumRegElim - 1
                If NumRegElim > 1 Then
                    For i = 1 To NumRegElim - 1
                        Data1.Recordset.MoveNext
                    Next i
                End If
                PonerCampos
'                DataGridAux(1).Enabled = True
        End If
        
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim impo As Currency
    
    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    
    If Not SePuedeEliminar Then Exit Sub
    If PertenceAlgunoDocumentoEmitido Then Exit Sub
    

    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    If impo < 0 Then
        MsgBox "Los abonos no se realizan por caja", vbExclamation
        Exit Sub
    End If

    'Menos ya pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)
    
    If impo <= 0 Then
        MsgBox "Totalmente cobrado", vbExclamation
        Exit Sub
    End If
    
    'Devolvera muuuuchas cosas
    'serie factura fecfac numvto
    Cad = Text1(13).Text & "|" & Format(Text1(1).Text, "0000000") & "|" & Text1(2).Text & "|" & Text1(3).Text & "|"
    'Codmacta nommacta codforpa   nomforpa   importe
    Cad = Cad & Text1(4).Text & "|" & Text2(0).Text & "|" & Text1(0).Text & "|" & Text2(1).Text & "|" & CStr(impo) & "|"
    'Lo que lleva cobrado
    Cad = Cad & Text1(8).Text & "|"
    
    
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub



Private Sub DataGridAux_DblClick(Index As Integer)
    If Index = 0 Then BotonVerAsiento
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
        cboFiltro.ListIndex = vUsu.FiltroPagos
    
        PrimeraVez = False
    
    End If
    
    CargarSqlFiltro
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim i As Integer


    PrimeraVez = True
    
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
        .Buttons(1).Image = 47
        .Buttons(2).Image = 44
        .Buttons(3).Image = 37
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
   
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    With Me.ToolbarAux
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 1
        .Buttons(2).Image = 16
    End With
    

    
    
    CargarColumnas
    
    CargaFiltros
    Me.cboFiltro.ListIndex = 0
    CargarCombo
    
    'Cargo los iconos
    For i = 0 To imgCuentas.Count - 1
        imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    For i = 0 To imgppal.Count - 1
        imgppal(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    imgSerie.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    Me.SSTab1.Tab = 0
    Me.Icon = frmppal.Icon
    LimpiarCampos
    
    'Recaudacion ejecutiva
    
    
    '## A mano
    NombreTabla = "pagos"
    Ordenacion = " ORDER BY numserie,numfactu,fecfactu,numorden"
        
    PonerOpcionesMenu
    
    CargaList 0, False
'    CargaGrid 1, False
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If

    Me.SSTab2.Tab = 0

End Sub

Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    txtPendiente.Text = ""
    check1(0).Value = 0
    Combo1.ListIndex = -1
    Text2(3).Text = ""
    lblIndicador.Caption = ""
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    vUsu.ActualizarFiltro "ariconta", IdPrograma, Me.cboFiltro.ListIndex
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim Cad As String

    If CadenaDevuelta <> "" Then
        If DevfrmCCtas <> "" Then
    
            HaDevueltoDatos = True
            DevfrmCCtas = CadenaDevuelta
            
        Else
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            DevfrmCCtas = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 1)
            Cad = DevfrmCCtas
            DevfrmCCtas = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            Cad = Cad & " AND " & DevfrmCCtas
            DevfrmCCtas = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            Cad = Cad & " AND " & DevfrmCCtas
            DevfrmCCtas = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
            Cad = Cad & " AND " & DevfrmCCtas
            DevfrmCCtas = Cad
            If DevfrmCCtas = "" Then Exit Sub
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & DevfrmCCtas & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    Else
        DevfrmCCtas = ""
    End If
End Sub

Private Sub PonerDatoDevuelto(CadenaDevuelta As String)
Dim Cad As String
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(13), CadenaDevuelta, 1)
    Cad = DevfrmCCtas
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
    Cad = Cad & " AND " & DevfrmCCtas
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
    Cad = Cad & " AND " & DevfrmCCtas
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
    Cad = Cad & " AND " & DevfrmCCtas
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 5)
    Cad = Cad & " AND " & DevfrmCCtas
    
    DevfrmCCtas = Cad
    If DevfrmCCtas = "" Then Exit Sub
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & DevfrmCCtas & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault
End Sub




Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
    txtAux(CInt(cmdAux(1).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    DevfrmCCtas = CadenaSeleccion
End Sub

Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(36).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(36).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
    End If
End Sub

Private Sub frmDpto_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(33).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(4).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmF_DatoSeleccionado(CadenaSeleccion As String)
       Text1(0) = RecuperaValor(CadenaSeleccion, 1)
       Text2(1) = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub ImgAgente_Click()
    Set frmA = New frmAgentes
    frmA.DatosADevolverBusqueda = "0|1|"
    frmA.Show vbModal
    Set frmA = Nothing
    
End Sub

Private Sub imgCuentas_Click(Index As Integer)
Dim Cad As String
Dim Z
    Screen.MousePointer = vbHourglass
    If Index = 1 Then
       
        Set frmF = New frmFormaPago
        frmF.DatosADevolverBusqueda = "0|"
        frmF.Show vbModal
        Set frmF = Nothing
    
        
    
    Else
        'Cuentas
        If Index = 0 And Me.Text1(4).Locked Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imgFecha(0).Tag = Index
        Set frmCCtas = New frmColCtas
        DevfrmCCtas = ""
        frmCCtas.DatosADevolverBusqueda = "0"
        frmCCtas.Show vbModal
        Set frmCCtas = Nothing
        If DevfrmCCtas <> "" Then
            If Index = 0 Then
                Text1(4 + Index) = RecuperaValor(DevfrmCCtas, 1)
            Else
                Text1(7 + Index) = RecuperaValor(DevfrmCCtas, 1)
            End If

            Text2(Index).Text = RecuperaValor(DevfrmCCtas, 2)
        End If
    End If
    
End Sub


Private Sub imgDepart_Click()

        ' departamento
        Indice = 33
        
        Set frmDpto = New frmBasico
        AyudaDepartamentos frmDpto, Text1(Indice).Text, "codmacta = " & DBSet(Text1(4).Text, "T")
        Set frmDpto = Nothing
        PonFoco Text1(Indice)
        


    
End Sub

Private Sub imgFecha_Click(Index As Integer)

    If Index = 0 And Me.Text1(2).Locked Then Exit Sub

    'En tag pongo el txtfecha asociado
    Select Case Index
    Case 0
        imgFecha(0).Tag = 2
    Case 1
        imgFecha(0).Tag = 5
    Case 2
        imgFecha(0).Tag = 7
    End Select
    DevfrmCCtas = Format(Now, "dd/mm/yyyy")
    If IsDate(Text1(CInt(imgFecha(0).Tag)).Text) Then _
        DevfrmCCtas = Format(Text1(CInt(imgFecha(0).Tag)).Text, "dd/mm/yyyy")
    Set frmC = New frmCal
    frmC.Fecha = CDate(DevfrmCCtas)
    DevfrmCCtas = ""
    frmC.Show vbModal
    Set frmC = Nothing
    
    
End Sub

Private Sub imgppal_Click(Index As Integer)
    If (Modo = 2 Or Modo = 5 Or Modo = 0) And (Index <> 0) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        ' observaciones
        Screen.MousePointer = vbDefault
        
        Indice = 17
        
        Set frmZ = New frmZoom
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
        frmZ.Caption = "Observaciones pagos"
        frmZ.Show vbModal
        Set frmZ = Nothing
        
    Case 1 ' pais
        Set frmPais = New frmBasico2
        AyudaPais frmPais
        Set frmPais = Nothing
    
    End Select

End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub



Private Sub imgSerie_Click()
        If Me.Text1(1).Locked Then Exit Sub
        Set frmConta = New frmBasico
        AyudaContadores frmConta, Text1(13).Text, "tiporegi REGEXP '^[0-9]+$' <> 0 and cast(tiporegi as unsigned) > 0"
        Set frmConta = Nothing
        PonFoco Text1(1)
    
    
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 4 Then CtaAnt = Text1(Index)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo = 1 Then
        'BUSQUEDA
        If KeyCode = 112 Then HacerF1
    ElseIf Modo = 0 Then
        If KeyCode = 27 Then Unload Me
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 27 Then
        PonleFoco Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
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
    Dim i As Integer
    Dim Sql As String
    Dim Valor
    
    If Text1(Index).Text = "" Then Exit Sub
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    
    If Not (Index = 4 Or Index = 10 Or Index = 9) Then
        If Modo < 2 Then Exit Sub
    End If
    'Campo con valor
    Select Case Index
    Case 4, 9
            'Cuentas          'Cuentas
            'Cuentas          'Cuentas
        i = DevuelveText2Relacionado(Index)
        DevfrmCCtas = Text1(Index).Text
        If CuentaCorrectaUltimoNivel(DevfrmCCtas, Sql) Then
            Text1(Index).Text = DevfrmCCtas
            If Modo >= 2 Then Text2(i).Text = Sql
        Else
            If Modo >= 2 Then
                MsgBox Sql, vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
            Text2(i).Text = ""
        End If
        
        'Poner la cuenta bancaria a partir de la cuenta
        If DevfrmCCtas <> "" Then
            If Modo > 2 And Index = 4 Then
                Sql = ""
                Valor = DevuelveLaCtaBanco(DevfrmCCtas)
                If Len(Valor) = 1 Then Valor = ""
                If CStr(Valor) <> "" Then
                    If Sql <> "" Then
                        If MsgBox("Poner Cuenta bancaria de la registro del cliente: " & Replace(CStr(Valor), "|", " - ") & "?", vbQuestion + vbYesNo) = vbYes Then Sql = ""
                    End If
                    If Sql = "" Then
                        Sql = DevuelveLaCtaBanco(DevfrmCCtas)

                        Text1(10).Text = Mid(RecuperaValor(Sql, 1), 1, 4)
                        Text1(26).Text = Mid(RecuperaValor(Sql, 1), 5, 4)
                        Text1(28).Text = Mid(RecuperaValor(Sql, 1), 9, 4)
                        Text1(29).Text = Mid(RecuperaValor(Sql, 1), 13, 4)
                        Text1(30).Text = Mid(RecuperaValor(Sql, 1), 17, 4)
                        Text1(21).Text = Mid(RecuperaValor(Sql, 1), 21, 4)

                        Text1(19).Text = RecuperaValor(Sql, 5)
                    End If
                End If

                Sql = DevuelveLaCtaBanco(DevfrmCCtas)
                


            End If
            If Index = 4 Then
                'Veremos si es asegurado
                If vParamT.TieneOperacionesAseguradas Then
                    Sql = DevuelveDesdeBD("numpoliz", "cuentas", "codmacta", DevfrmCCtas, "T")
                End If
                
                
                If Modo = 3 Then
                    Sql = "concat(if( isnull(forpa),'',forpa),'|',if(isnull(ctabanco),'',ctabanco),'|')"
                    Sql = DevuelveDesdeBD(Sql, "cuentas", "codmacta", DevfrmCCtas, "T")
                    If Sql <> "" Then
                        Text1(0).Text = RecuperaValor(Sql, 1)
                        Text1(9).Text = RecuperaValor(Sql, 2)
                        If Text1(9).Text <> "" Then Text2(2).Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(9).Text, "T", Text1(9).Text)
                        If Text1(0).Text <> "" Then Text1_LostFocus 0   'VOLVEMOS A LLAMR a la lostfocus, cuidado con las variables
                    End If
                End If
            End If
            If Index = 4 And Trim(Text1(Index).Text) = CtaAnt Then Exit Sub
            
            If Index = 4 And (Modo = 3 Or (Modo = 4 And Trim(Text1(Index).Text) <> CtaAnt)) Then
                CargarDatosCuenta Text1(Index)
            End If
            
            
        End If
     Case 0
        'FORMA DE PAGO
        vTipForpa = ""
        DevfrmCCtas = "tipforpa"
        If Not IsNumeric(Text1(Index).Text) Then
            Sql = "Campo Forma pago debe ser numérico: " & Text1(Index).Text
            MsgBox Sql, vbExclamation
            Sql = ""
        Else
            Sql = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", Text1(Index).Text, "N", DevfrmCCtas)
            If Sql = "" Then
                Sql = "Forma de pago inexistente: " & Text1(Index).Text
                MsgBox Sql, vbExclamation
                Sql = ""
            Else
                vTipForpa = DevfrmCCtas
            End If
        End If
        Text2(1).Text = Sql
        If vTipForpa = "" Then
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
        
        
    Case 2, 5, 7
        'FECHAS
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
        
    Case 6, 8
        PonerFormatoDecimal Text1(Index), 3

    Case 3
        'Vencimiento
        'Debe ser numerico
        If Not IsNumeric(Text1(3).Text) Then
            MsgBox "Campo debe ser numerico", vbExclamation
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
        
    Case 13
        If Not IsNumeric(Text1(13).Text) Then
            MsgBox "Serie es un numérico.", vbExclamation
            Text1(13).Text = ""
            PonerFoco Text1(13)
        Else
            Text1(13).Text = UCase(Text1(13).Text)
        End If
        

    Case 28 To 30, 10, 26, 21
        If Index <> 10 Then
            'Cuenta bancaria
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Cuenta banco debe ser numérico: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            Else
                'Formateamos
                If Text1(Index) <> "" Then
                    Text1(Index).Text = Format(Text1(Index).Text, "0000")
                End If
            End If
        Else
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
        End If
        
        Sql = Text1(26).Text & Text1(28).Text & Text1(29).Text & Text1(30).Text & Text1(21).Text
        
        If Len(Sql) = 20 And Index = 21 Then 'solo cuando pierde el foco la cuentaban
            'OK. Calculamos el IBAN
            If Text1(10).Text = "" Then
                'NO ha puesto IBAN
                If DevuelveIBAN2("ES", Sql, Sql) Then Text1(10).Text = "ES" & Sql
            Else
                Valor = CStr(Mid(Text1(10).Text, 1, 2))
                If DevuelveIBAN2(CStr(Valor), Sql, Sql) Then
                    If Mid(Text1(10).Text, 3) <> Sql Then
                        MsgBox "Codigo IBAN distinto del calculado [" & Valor & Sql & "]", vbExclamation
                    End If
                End If
            End If
        End If
        
        Text1(19).Text = Text1(10).Text & Text1(26).Text & Text1(28).Text & Text1(29).Text & Text1(30).Text & Text1(21).Text
        
        
    Case 25 ' codigo de pais
        If Text1(Index).Text <> "" Then
            Text2(Index).Text = PonerNombreDeCod(Text1(Index), "paises", "nompais", "codpais", "T")
            If Text2(Index) = "" Then
                MsgBox "No existe el País. Reintroduzca.", vbExclamation
                PonFoco Text1(Index)
            End If
        Else
            Text2(Index).Text = ""
        End If
        
        
    End Select
            
End Sub

Public Function DevuelveText2Relacionado(Index As Integer) As Integer
        DevuelveText2Relacionado = -1
        Select Case Index
        Case 0
            DevuelveText2Relacionado = 1
        Case 4
            DevuelveText2Relacionado = 0
        Case 9
            DevuelveText2Relacionado = 2
        Case 10
            DevuelveText2Relacionado = 3
        End Select
End Function


Private Sub HacerBusqueda()
Dim Cad As String

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    ' busqueda por las partes del iban
    
    If Text1(10).Text <> "" Then
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & "mid(iban,1,4) = " & DBSet(Text1(10).Text, "T")
    End If
    If Text1(26).Text <> "" Then
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & "mid(iban,5,4) = " & DBSet(Text1(26).Text, "T")
    End If
    If Text1(28).Text <> "" Then
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & "mid(iban,9,4) = " & DBSet(Text1(28).Text, "T")
    End If
    If Text1(29).Text <> "" Then
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & "mid(iban,13,4) = " & DBSet(Text1(29).Text, "T")
    End If
    If Text1(30).Text <> "" Then
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & "mid(iban,17,4) = " & DBSet(Text1(30).Text, "T")
    End If
    If Text1(21).Text <> "" Then
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & "mid(iban,21,4) = " & DBSet(Text1(21).Text, "T")
    End If
    
    
    CadB1 = ObtenerBusqueda2(Me, , 2, "FrameAux1")
    
    HacerBusqueda2


End Sub

Private Sub HacerBusqueda2()

    CargarSqlFiltro
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Or CadB1 <> "" Or cadFiltro <> "" Then
        CadenaConsulta = "select distinct pagos.* from "
        CadenaConsulta = CadenaConsulta & " (tipofpago  INNER JOIN hlinapu ON hlinapu.tipforpa = tipofpago.tipoformapago ) "
        CadenaConsulta = CadenaConsulta & " right join pagos on pagos.numserie = hlinapu.numserie and pagos.numfactu = hlinapu.numfacpr and pagos.fecfactu = hlinapu.fecfactu and pagos.numorden = hlinapu.numorden and pagos.codmacta = hlinapu.codmacta "

        
        CadenaConsulta = CadenaConsulta & " WHERE (1=1) "
        If CadB <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB & " "
        If CadB1 <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB1 & " "
        If cadFiltro <> "" Then CadenaConsulta = CadenaConsulta & " and " & cadFiltro & " "
        
        CadenaConsulta = CadenaConsulta & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonFoco Text1(0)
        ' **********************************************************************
    End If
    

End Sub

Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    cadFiltro = ""
    
    Select Case cboFiltro.ItemData(cboFiltro.ListIndex) 'cboFiltro.ListIndex
        Case 0 ' pendientes de cobro
            cadFiltro = "(coalesce(pagos.impefect,0)  - coalesce(pagos.imppagad,0) <> 0) "
        Case 1 ' cobrados
            cadFiltro = "(coalesce(pagos.impefect,0)  - coalesce(pagos.imppagad,0) = 0) and pagos.situacion = 1"
        Case 9 ' todos
            cadFiltro = "(1=1)"
    End Select
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    CadenaDesdeOtroForm = ""
    
    frmTESVerCobrosPagos.vSql = cadFiltro
    If CadB <> "" Then frmTESVerCobrosPagos.vSql = frmTESVerCobrosPagos.vSql & " and " & CadB
    frmTESVerCobrosPagos.OrdenarEfecto = False
    frmTESVerCobrosPagos.Regresar = True
    frmTESVerCobrosPagos.Cobros = False
    frmTESVerCobrosPagos.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        PonerDatoDevuelto CadenaDesdeOtroForm
        If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then cmdRegresar_Click
    Else   'de ha devuelto datos, es decir NO ha devuelto datos
        PonerFoco Text1(kCampo)
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
    Dim i As Integer
    Dim mTag As CTag
    Dim Sql As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1
'    PonerCtasIVA
    
    Text2(0).Text = PonerNombreDeCod(Text1(4), "cuentas", "nommacta", "codmacta", "T")
    Text2(2).Text = PonerNombreDeCod(Text1(9), "cuentas", "nommacta", "codmacta", "T")
    Text2(3).Text = Combo1.Text 'PonerNombreDeCod(Text1(10), "cuentas", "nommacta", "codmacta", "T")
    Text2(25).Text = PonerNombreDeCod(Text1(25), "paises", "nompais", "codpais", "T")
    Text2(1).Text = PonerNombreDeCod(Text1(0), "formapago", "nomforpa", "codforpa", "N")
    
    vTipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Text1(0).Text, "N")
    
    If Text1(32).Text = "" Then
        cboSituRem.ListIndex = -1
    Else
        PosicionarCombo cboSituRem, Asc(Text1(32).Text)
    End If
    
    cboSituRem_Validate False
    
    
    
    ' situamos los campos del iban
    Text1(10).Text = ""
    Text1(26).Text = ""
    Text1(28).Text = ""
    Text1(29).Text = ""
    Text1(30).Text = ""
    Text1(21).Text = ""
    
    Text1(10).ToolTipText = ""
    Text1(26).ToolTipText = ""
    Text1(28).ToolTipText = ""
    Text1(29).ToolTipText = ""
    Text1(30).ToolTipText = ""
    Text1(21).ToolTipText = ""
    
    If Text1(19).Text <> "" Then
        Text1(10) = Mid(Text1(19), 1, 4)
        Text1(26) = Mid(Text1(19), 5, 4)
        Text1(28) = Mid(Text1(19), 9, 4)
        Text1(29) = Mid(Text1(19), 13, 4)
        Text1(30) = Mid(Text1(19), 17, 4)
        Text1(21) = Mid(Text1(19), 21, 4)
        
        Dim CCC As String
        CCC = Text1(10).Text & " " & Text1(26).Text & " " & Text1(28).Text & " " & Mid(Text1(29).Text, 1, 2) & " " & Mid(Text1(29).Text, 3, 2) & Text1(30).Text & Text1(21).Text
        
        Text1(10).ToolTipText = CCC
        Text1(26).ToolTipText = CCC
        Text1(28).ToolTipText = CCC
        Text1(29).ToolTipText = CCC
        Text1(30).ToolTipText = CCC
        Text1(21).ToolTipText = CCC
    
    End If
    
    
    
    'Cargamos el LINEAS
    CargaList 0, True
    
    PonPendiente
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
End Sub


Private Sub PonPendiente()
Dim Importe As Currency

    On Error GoTo EPonPendiente
    'Pendiente
    Importe = Data1.Recordset!ImpEfect - DBLet(Data1.Recordset!imppagad, "N")
    txtPendiente.Text = Format(Importe, FormatoImporte)
    
EPonPendiente:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        Err.Clear
    End If
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer, Optional indFrame As Integer)
    Dim i As Integer
    Dim B As Boolean
    
    BuscaChekc = ""
    If Modo = 1 Then
        'Ponemos todos a fondo blanco
        '### a mano
        For i = 0 To Text1.Count - 1
            Text1(0).BackColor = &H80000018
        Next i
        Text1(28).MaxLength = 4
        Text1(29).MaxLength = 4
    ElseIf Modo = 4 Then
    End If
    
    'Modo buscar
    If Kmodo = 1 Then
        Text1(28).MaxLength = 0
        Text1(29).MaxLength = 0
    End If
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    If Not Data1.Recordset Is Nothing Then
        DespalzamientoVisible B And (Data1.Recordset.RecordCount > 1)
    End If
    
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    Else
        cmdRegresar.Visible = False
    End If
    
    FrameRemesa.Enabled = Kmodo = 1
    Text1(27).Enabled = Kmodo = 1
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
       
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    
    B = Modo = 2 Or Modo = 0 Or Modo = 5
    
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = B
        If Modo <> 1 Then
            Text1(i).BackColor = vbWhite
        End If
    Next i
    
    'Empieza siempre a false
    Toolbar2.Buttons(2).Enabled = False
    
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    B = (Modo = 2) Or Modo = 0
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = B
    Next i
    
    cboSituRem.Locked = B

    frameContene.Enabled = Not B
    
    For i = 0 To 2
        If i <> 3 Then imgCuentas(i).Visible = Not B
        Me.imgFecha(i).Visible = Not B
    Next i
    
    Me.imgSerie.Visible = Not B
        
        
    
    'lineas de pagos_realizados
    Dim anc As Single
    
    Combo1.Enabled = False '(Modo = 1) Or ((vUsu.Nombre = "root") And Modo = 4)
    
        
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Tipo As Integer

    DatosOK = False
    
    If cboSituRem.ListIndex = 0 Or cboSituRem.ListIndex = -1 Then
        Text1(32).Text = ""
    Else
        Text1(32).Text = Chr(cboSituRem.ItemData(cboSituRem.ListIndex))
    End If
    
    
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    'NUmero serie
    DevfrmCCtas = DevuelveDesdeBD("tiporegi", "contadores", "tiporegi", Text1(13).Text, "T")
    If DevfrmCCtas = "" Then
        B = False
        MsgBox "Serie no existe", vbExclamation
        Exit Function
    End If
    
    DevfrmCCtas = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Text1(0).Text, "N")
    Tipo = CInt(DevfrmCCtas)

    DevfrmCCtas = Trim(Text1(26).Text) & Trim(Text1(28).Text) & Text1(29).Text & Trim(Text1(30).Text) & Trim(Text1(21).Text)
    
    'AHora comprobare Si tiene cuenta bancaria, si es correcta
    B = False
    If DevfrmCCtas <> "" Then
        If Val(DevfrmCCtas) <> 0 Then B = True
    End If
        
    If B Then
        'Vale, hay campos y son numericos
        'La cuenta contable si digi control, si tiene valor, tiene que ser longitud 18
        If Len(DevfrmCCtas) < 20 Then
            MsgBox "Cuenta bancaria incorrecta", vbExclamation
            Exit Function
        End If
    End If
        
    If B Then
        'Compruebo EL IBAN
        'Meto el CC
        BuscaChekc = ""
        If Me.Text1(10).Text <> "" Then BuscaChekc = Mid(Text1(10).Text, 1, 2)
            
        If DevuelveIBAN2(BuscaChekc, DevfrmCCtas, DevfrmCCtas) Then
            If Me.Text1(10).Text = "" Then
                If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(10).Text = BuscaChekc & DevfrmCCtas
            Else
                If Mid(Text1(10).Text, 3) <> DevfrmCCtas Then
                    DevfrmCCtas = "Calculado : " & BuscaChekc & DevfrmCCtas
                    DevfrmCCtas = "Introducido: " & Me.Text1(10).Text & vbCrLf & DevfrmCCtas & vbCrLf
                    DevfrmCCtas = "Error en codigo IBAN" & vbCrLf & DevfrmCCtas & "Continuar?"
                    If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
            End If
        End If
    End If
    
    If Not B Then Text1(19).Text = ""
    
    If CuentaBloqeada(Me.Text1(4).Text, CDate(Text1(2).Text), True) Then Exit Function
        

    If Modo = 4 Then
        If DBLet(Me.Data1.Recordset!emitdocum, "N") = 1 Then
            'Tiene la marca de documento emitido
            'Veremos si se la ha quitado
            If Me.check1(0).Value = 0 Then
                DevfrmCCtas = "Seguro que desea quitarle la marca de documento emitido?"
                If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
            End If
        End If
    End If

    
    'Nuevo. 12 Mayo 2008
    B = CuentaBloqeada(Me.Text1(4).Text, CDate(Text1(2).Text), True)
    If B Then
        If (vUsu.Codigo Mod 100) > 0 Then Exit Function
    End If
    
    
    If Modo = 4 Then
        If Text1(8).Text = "" Then
            Text1(7).Text = ""
            Combo1.ListIndex = 0
        End If
    End If
    
    
    DatosOK = True
End Function




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub


Private Sub HacerToolBar(Boton As Integer)

    Select Case Boton
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
            'Imprimir factura
            
            frmTESPagosPdtesList.Show vbModal

    End Select
End Sub



Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerCtasIVA()
On Error GoTo EPonerCtasIVA

    Text1_LostFocus 4
    Text1_LostFocus 0
    Text1_LostFocus 9
    Text1_LostFocus 10
Exit Sub
EPonerCtasIVA:
    MuestraError Err.Number, "Poniendo valores ctas. IVA", Err.Description
End Sub

Private Sub PonerFoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

'Si no esta en transferencia o en una remesa
'entonces dejare que modifique algun dato basico
'Realmente solo la cta bancaria
Private Function SePuedeEliminar() As Byte
    
    SePuedeEliminar = False

    If Not IsNull(Me.Data1.Recordset!nrodocum) Then
        If Val(Me.Data1.Recordset!nrodocum) > 0 Then
            MsgBox "Pertenece a una transferencia.", vbExclamation
            Exit Function
        End If
    End If
    
    SePuedeEliminar = True

    Screen.MousePointer = vbDefault
End Function

Private Function PertenceAlgunoDocumentoEmitido() As Boolean

    PertenceAlgunoDocumentoEmitido = False
    If Val(Data1.Recordset!emitdocum) = 1 Then
        If MsgBox("Pertence a un documento emtitido.  No deberia seguir con el proceso." & vbCrLf & vbCrLf & "Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then PertenceAlgunoDocumentoEmitido = True
    End If

End Function


Private Sub RealizarPagoCuenta()
Dim impo As Currency
    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    'Pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)

    'Si impo>0 entonces TODAVIA puedn pagarme algo
    If impo = 0 Then
        'Cosa rara. Esta todo el importe pagado
        MsgBox "La factura está totalmente cobrada.", vbExclamation
        Exit Sub
    End If

    frmTESParciales.Cobro = False
    frmTESParciales.Vto = Text1(13).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & Text1(5).Text & "|"
    frmTESParciales.Importes = Text1(6).Text & "|" & Text1(8).Text & "|"
    frmTESParciales.Cta = Text1(4).Text & "|" & Text2(0).Text & "|" & Text1(9).Text & "|" & Text2(2).Text & "|"
    frmTESParciales.FormaPago = Val(vTipForpa)
    frmTESParciales.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        'Hay que refrescar los datos
        lblIndicador.Caption = ""
        If SituarData Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
End Sub

Private Sub HacerF1()
Dim C As String
    
    C = ObtenerBusqueda2(Me, BuscaChekc)
    If C = "" Then Text1(13).Text = "*"  'Para que busqu toooodo
    cmdAceptar_Click
End Sub




Private Sub DividirVencimiento()
Dim Im As Currency

    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    
    'Si esta totalmente cobrado pues no podemos desdoblar ekl vto
    
    
    
    If Val(DBLet(Data1.Recordset!nrodocum, "N")) = 1 Then
        MsgBox "Pertenece a una transferencia", vbExclamation
        Exit Sub
    End If
    
    
    Im = Data1.Recordset!ImpEfect
    Im = Im - DBLet(Data1.Recordset!imppagad, "N")
    If Im = 0 Then
        MsgBox "NO puede dividir el vencimiento. Importe totalmente pagado", vbExclamation
        Exit Sub
    End If
    
    
       'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
    
    CadenaDesdeOtroForm = "numserie = '" & Data1.Recordset!NUmSerie & "' AND numfactu = " & DBSet(Data1.Recordset!NumFactu, "T")
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND fecfactu = '" & Format(Data1.Recordset!FecFactu, FormatoFecha) & "' and codmacta = " & DBSet(Data1.Recordset!codmacta, "T") & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Data1.Recordset!numorden & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CStr(Im) & "|"
    
    
    'Ok, Ahora pongo los labels
    frmTESPagosDivVto.Opcion = 27
    frmTESPagosDivVto.Label4(56).Caption = Text2(0).Text
    frmTESPagosDivVto.txtCodigo(2).Text = Text1(5).Text
    
    frmTESPagosDivVto.Label4(57).Caption = Data1.Recordset!NUmSerie & Format(Data1.Recordset!NumFactu, "000000") & " / " & Data1.Recordset!numorden & "      de " & Format(Data1.Recordset!FecFactu, "dd/mm/yyyy")
    
    'Si ya ha cobrado algo...
    Im = DBLet(Data1.Recordset!imppagad, "N")
    If Im > 0 Then frmTESPagosDivVto.txtCodigo(1).Text = txtPendiente.Text
    
    If Text1(0).Text = "" Then
        MsgBox "El pago no tiene forma de pago. Revise.", vbExclamation
        Exit Sub
    End If
    
    frmTESPagosDivVto.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = "Select * from pagos WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1) 'CadenaConsulta
        Data1.RecordSource = CadenaConsulta
        Data1.Refresh
        If Data1.Recordset.RecordCount <= 0 Then
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Else
            DevfrmCCtas = ""
            PonerCampos
            PonerModo 2
        End If
    End If

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Datos Fiscales
            Me.FrameDatosFiscales.Visible = Not Me.FrameDatosFiscales.Visible
           
        Case 2
            'dividir vencimientos
            If Text1(13).Text = "" Then Exit Sub
            
            DividirVencimiento
        
        
        
        Case 3
            'Generar pagos
            If Me.Data1.Recordset.EOF Then Exit Sub
            If Modo <> 2 Then Exit Sub
            If vTipForpa <> "" Then
                If (Val(vTipForpa) <> vbTransferencia) Or (Val(vTipForpa) = vbTransferencia And Val(DBLet(Data1.Recordset!nrodocum)) = 0) Then
                    If Not SePuedeEliminar Then Exit Sub
                
                    If PertenceAlgunoDocumentoEmitido Then Exit Sub

                
                    'Bloqueamos
                    If BloqueaRegistroForm(Me) Then
                        RealizarPagoCuenta
                        DesBloqueaRegistroForm Text1(0)
                    End If
                Else
                    MsgBox "Lo pagos a cuenta no se realizan sobre RECIBOS BANCARIOS Remesados", vbExclamation
                End If
            End If
    End Select

End Sub

Private Sub CargaList(Index As Integer, Enlaza As Boolean)
Dim IT
Dim Cad As String

    lwpagos.ListItems.Clear
    Set Me.lwpagos.SmallIcons = frmppal.imgListComun16 'imgListComun 'ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "SELECT  hlinapu.numdiari, hlinapu.fechaent, "
    Cad = Cad & " hlinapu.numasien, hlinapu.ctacontr, "
    Cad = Cad & " hcabapu.usucreacion, hcabapu.feccreacion, tipofpago.siglas, "
    Cad = Cad & " hlinapu.reftalonpag, hlinapu.bancotalonpag, "
    Cad = Cad & " (coalesce(hlinapu.timported,0) - coalesce(hlinapu.timporteh,0)) imppago, hlinapu.numserie, hlinapu.numfacpr, hlinapu.fecfactu, hlinapu.numorden, hlinapu.codmacta "
    Cad = Cad & " FROM (hlinapu INNER JOIN tipofpago ON hlinapu.tipforpa = tipofpago.tipoformapago) INNER JOIN hcabapu ON hlinapu.numdiari = hcabapu.numdiari and hlinapu.fechaent = hcabapu.fechaent and hlinapu.numasien = hcabapu.numasien "
    If Enlaza Then
        Cad = Cad & Replace(Replace(ObtenerWhereCab(True), "pagos", "hlinapu"), "numfactu", "numfacpr")
    Else
        Cad = Cad & " WHERE hlinapu.codmacta is null"
    End If
    Cad = Cad & " ORDER BY hlinapu.numserie, hlinapu.numfacpr, hlinapu.fecfactu, hlinapu.numorden, hlinapu.fechaent, hlinapu.numasien"
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwpagos.ListItems.Add()
        IT.Text = DBLet(miRsAux!NumDiari, "N")
        IT.SubItems(1) = Format(miRsAux!FechaEnt, "dd/mm/yyyy")
        IT.SubItems(2) = DBLet(miRsAux!NumAsien, "N")
        IT.SubItems(3) = DBLet(miRsAux!ctacontr, "T")
        IT.SubItems(4) = DBLet(miRsAux!usucreacion, "T")
        IT.SubItems(5) = Format(DBLet(miRsAux!feccreacion, "F"), "dd/mm/yyyy")
        IT.SubItems(6) = DBLet(miRsAux!siglas, "T")
        IT.SubItems(7) = DBLet(miRsAux!reftalonpag, "T")
        IT.SubItems(8) = DBLet(miRsAux!bancotalonpag, "T")
        IT.SubItems(9) = Format(miRsAux!imppago, "###,###,##0.00")
        
        IT.SubItems(10) = DBLet(miRsAux!NUmSerie)
        IT.SubItems(11) = DBLet(miRsAux!NumFacpr)
        IT.SubItems(12) = DBLet(miRsAux!FecFactu, "F")
        IT.SubItems(13) = DBLet(miRsAux!numorden)
        IT.SubItems(14) = DBLet(miRsAux!codmacta)
        
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    PonerModoUsuarioGnral Modo, "ariconta"

End Sub




Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'hlinapu
            ' esta en un listview
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & "pagos.numserie=" & DBSet(Text1(13).Text, "T") & " and pagos.numfactu=" & DBSet(Text1(1).Text, "T") & " and pagos.fecfactu = " & DBSet(Text1(2).Text, "F")
    vWhere = vWhere & " and pagos.numorden = " & DBSet(Text1(3).Text, "N")
    vWhere = vWhere & " and pagos.codmacta = " & DBSet(Text1(4).Text, "T")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(numserie=" & DBSet(Text1(13).Text, "T") & " and numfactu = " & DBSet(Text1(1).Text, "T") & " and fecfactu = " & DBSet(Text1(2).Text, "F") & " and numorden = " & DBSet(Text1(3).Text, "N") & ") "
    
    If SituarDataMULTI(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N") And (Modo <> 0 And Modo <> 5)
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!especial, "N") And Modo = 2
        Me.Toolbar2.Buttons(3).Enabled = DBLet(Rs!especial, "N") And Modo = 2
        
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!especial, "N") And (Modo = 2) And Not Me.AdoAux(0).Recordset.EOF
        ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 2) And Not Me.AdoAux(0).Recordset.EOF
            
        
        vUsu.LeerFiltros "ariconta", IdPrograma
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub





Private Sub CargaFiltros()
Dim Aux As String
    

    cboFiltro.Clear
    
    cboFiltro.AddItem "Pendientes de Pago "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Pagado "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1

    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 9


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


Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim J As Long
    

    Combo1.Clear

    Combo1.AddItem "Pendientes de Pago"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Pagado"
    Combo1.ItemData(Combo1.NewIndex) = 1


    
    cboSituRem.Clear

    cboSituRem.AddItem ""
    cboSituRem.ItemData(cboSituRem.NewIndex) = Asc("NULL")


    'Tipo de situacion de la transferencia
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtiposituacionrem ORDER BY situacio"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        cboSituRem.AddItem Rs!descsituacion
        cboSituRem.ItemData(cboSituRem.NewIndex) = Asc(Rs!situacio)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

End Sub


'************* LLINIES: ****************************
Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LINASI As Long
Dim Ampliacion As String

    'Fuerzo que se vean las lineas
    Select Case Button.Index
        Case 1
            'Acceder a asiento del cobro
            BotonVerAsiento
        Case 2
            'Impresion del recibo
            BotonImprimirRecibo
    End Select

End Sub



Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'linea de asiento
            Sql = "¿Seguro que desea eliminar la línea de la factura?"
            Sql = Sql & vbCrLf & "Serie: " & AdoAux(Index).Recordset!NUmSerie & " - " & AdoAux(Index).Recordset!NumFactu & " - " & AdoAux(Index).Recordset!FecFactu & " - " & AdoAux(Index).Recordset!NumLinea
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM factcli_lineas "
                Sql = Sql & Replace(vWhere, "factcli", "factcli_lineas") & " and numlinea = " & DBLet(AdoAux(Index).Recordset!NumLinea, "N")
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute Sql
        
  '      RecalcularTotales
        
        '**** parte de contabilizacion de la factura
        TerminaBloquear
        
        
        'LOG
        vLog.Insertar 6, vUsu, Text1(2).Text & Text1(0).Text & " " & Text1(1).Text
        'Creo que no hace falta volver a situar el datagrid
        If True Then
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
            Data1.Refresh
            PonerModo 2
        Else
            PonerModo 0
        End If
        '**** hasta aqui
        
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then
        End If
        
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto)
        ' ************************
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub CargaFrame(Index As Integer, Enlaza As Boolean)
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Sub BotonVerAsiento()

    If lwpagos.SelectedItem Is Nothing Then Exit Sub


    Set frmAsi = New frmAsientosHco
    
    frmAsi.ASIENTO = lwpagos.SelectedItem.Text & "|" & lwpagos.SelectedItem.SubItems(1) & "|" & lwpagos.SelectedItem.SubItems(2) & "|" '& Me.AdoAux(0).Recordset.Fields(5) & "|" & Me.AdoAux(0).Recordset.Fields(6) & "|" & Me.AdoAux(0).Recordset.Fields(7) & "|"
    frmAsi.SoloImprimir = True
    frmAsi.Show vbModal
    
    Set frmAsi = Nothing

End Sub

Private Sub BotonImprimirRecibo()

    If lwpagos.SelectedItem Is Nothing Then Exit Sub


    CargarTemporal

    frmTESImpRecibo.pImporte = lwpagos.SelectedItem.SubItems(9) 'Me.AdoAux(0).Recordset.Fields(14)
    frmTESImpRecibo.pFechaRec = lwpagos.SelectedItem.SubItems(1) 'Me.AdoAux(0).Recordset.Fields(6)
    frmTESImpRecibo.pFecFactu = lwpagos.SelectedItem.SubItems(12) 'Me.AdoAux(0).Recordset.Fields(2)
    frmTESImpRecibo.pNumFactu = lwpagos.SelectedItem.SubItems(11) 'Me.AdoAux(0).Recordset.Fields(1)
    frmTESImpRecibo.pNumSerie = lwpagos.SelectedItem.SubItems(10) 'Me.AdoAux(0).Recordset.Fields(0)
    frmTESImpRecibo.pNumOrden = lwpagos.SelectedItem.SubItems(13) 'Me.AdoAux(0).Recordset.Fields(3)
    frmTESImpRecibo.pNumlinea = lwpagos.SelectedItem.SubItems(14) 'Me.AdoAux(0).Recordset.Fields(4)
    
    frmTESImpRecibo.Show vbModal


End Sub

Private Sub CargarTemporal()
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "delete from tmppendientes where codusu = " & vUsu.Codigo
    Conn.Execute Sql
                                                                              
    ' en tmppendientes metemos la clave primaria de pagos_recibidos y el importe en letra
                                                      'importe=nro factura,   codforpa=linea de pagos_realizados
    Sql = "insert into tmppendientes (codusu,serie_cta,importe,fecha,numorden,codforpa, observa) values ("
    Sql = Sql & vUsu.Codigo & "," & DBSet(lwpagos.SelectedItem.SubItems(10), "T") & "," 'numserie
    Sql = Sql & DBSet(lwpagos.SelectedItem.SubItems(11), "N") & "," 'numfactu
    Sql = Sql & DBSet(lwpagos.SelectedItem.SubItems(12), "F") & "," 'fecfactu
    Sql = Sql & DBSet(lwpagos.SelectedItem.SubItems(13), "N") & "," 'numorden
    Sql = Sql & DBSet(lwpagos.SelectedItem.SubItems(14), "N") & "," 'numlinea
    Sql = Sql & DBSet(EscribeImporteLetra(ImporteFormateado(CStr(lwpagos.SelectedItem.SubItems(9)))), "T") & ") "
    
    Conn.Execute Sql

End Sub





Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    SepuedeBorrar = True
End Function





Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim B As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte


    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And (Modo = 5 And ModoLineas = 1) Then  'insertar
    
    End If
    
    If B And Modo = 5 Then ' tanto si insertamos como si modificamos en lineas
        
        
    End If
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5:  KEYImage KeyAscii, 0 ' cta base
            Case 7:  KEYImage KeyAscii, 1 ' iva
            Case 12:  KEYImage KeyAscii, 2 ' Centro Coste
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub
'++


Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Importe As Currency
        
        
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    If txtAux(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 5 ' diario
            RC = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtAux(5), "N")
            If RC = "" Then
                MsgBox "No existe el tipo de diario. Reintroduzca.", vbExclamation
                PonFoco txtAux(5)
            End If
                
        Case 6, 11 ' fecha
            If Not EsFechaOK(txtAux(Index)) Then
                MsgBox "Fecha incorrecta: " & txtAux(Index).Text, vbExclamation
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
            
        Case 7 ' asiento
            PonerFormatoEntero txtAux(Index)
        
        Case 8 ' usuario
        
        Case 9
           ' IMPORTE
             txtAux(Index) = ImporteSinFormato(txtAux(Index))
            
        Case 10 'tipo
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 12 ' cuenta de cobro
            RC = txtAux(12).Text
            If CuentaCorrectaUltimoNivel(RC, "") Then
                txtAux(12).Text = RC
            End If
        
    End Select
End Sub



Private Sub CargarDatosCuenta(Cuenta As String)
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo eTraerDatosCuenta
    
    Sql = "select * from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(0).Text = ""
    Text2(1).Text = ""
    
    For i = 14 To 16
        Text1(i).Text = ""
    Next i
    For i = 22 To 25
        Text1(i).Text = ""
    Next i
    
    If Not Rs.EOF Then
        
        Text1(0).Text = DBLet(Rs!Forpa, "N")
        Text2(1).Text = PonerNombreDeCod(Text1(0), "formapago", "nomforpa", "codforpa", "N")
        
        Text1(16).Text = DBLet(Rs!Nommacta, "T")
        Text1(15).Text = DBLet(Rs!dirdatos, "T")
        Text1(14).Text = DBLet(Rs!codposta, "T")
        Text1(22).Text = DBLet(Rs!desPobla, "T")
        Text1(23).Text = DBLet(Rs!desProvi, "T")
        Text1(24).Text = DBLet(Rs!nifdatos, "T")
        Text1(25).Text = DBLet(Rs!codPAIS, "T")
        Text2(25).Text = PonerNombreDeCod(Text1(25), "paises", "nompais", "codpais", "T")
    End If
    Exit Sub
    
eTraerDatosCuenta:
    MuestraError Err.Number, "Cargar Datos de Cuenta", Err.Description

End Sub

Private Sub CargarColumnas()
    
    lwpagos.ColumnHeaders.Clear
    
    lwpagos.ColumnHeaders.Add , , "Diario", 830
    lwpagos.ColumnHeaders.Add , , "Fecha", 1320
    lwpagos.ColumnHeaders.Add , , "Asiento", 1200, 1
    lwpagos.ColumnHeaders.Add , , "Cta.Pago", 1405
    lwpagos.ColumnHeaders.Add , , "Usuario", 1405
    lwpagos.ColumnHeaders.Add , , "Realizado", 1455
    lwpagos.ColumnHeaders.Add , , "Tipo", 905
    lwpagos.ColumnHeaders.Add , , "Ref.Talon", 1755
    lwpagos.ColumnHeaders.Add , , "Banco Talon/Pag", 1855
    lwpagos.ColumnHeaders.Add , , "Importe", 1900, 1
    
    lwpagos.ColumnHeaders.Add , , "serie", 0
    lwpagos.ColumnHeaders.Add , , "Factura", 0
    lwpagos.ColumnHeaders.Add , , "Fecha", 0
    lwpagos.ColumnHeaders.Add , , "vto", 0
    lwpagos.ColumnHeaders.Add , , "Linea", 0
    lwpagos.ColumnHeaders.Add , , "Cuenta", 0

End Sub


