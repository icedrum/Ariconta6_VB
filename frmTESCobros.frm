VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTESCobros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15840
   Icon            =   "frmTESCobros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
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
      TabIndex        =   144
      Top             =   10200
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
      TabIndex        =   142
      Top             =   10200
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
      TabIndex        =   143
      Top             =   10200
      Width           =   1035
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   8490
      TabIndex        =   88
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
         ItemData        =   "frmTESCobros.frx":000C
         Left            =   90
         List            =   "frmTESCobros.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   210
         Width           =   3075
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   85
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   86
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
      TabIndex        =   83
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   84
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
      TabIndex        =   81
      Top             =   30
      Width           =   1965
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   82
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
               Object.ToolTipText     =   "Generar Cobro"
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4260
      Top             =   9660
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
      Height          =   8235
      Left            =   90
      TabIndex        =   48
      Top             =   1740
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   14526
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
      TabCaption(0)   =   "Datos Cobro"
      TabPicture(0)   =   "frmTESCobros.frx":0050
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
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgDepart"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(11)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(18)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "imgAgente"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label6"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(33)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(34)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(35)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "imgFecha(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "imgFecha(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "imgFecha(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "imgppal(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(10)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Combo1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(19)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(26)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(31)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(30)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(29)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(28)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text2(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(4)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text2(1)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(5)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(6)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text2(2)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(9)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text2(3)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(10)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(33)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text2(4)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text1(32)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(16)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text1(17)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "frameContene"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text2(5)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Text1(34)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtPendiente"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Text1(12)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Text1(11)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Text1(7)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Text1(8)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "FrameRemesa"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Check1(2)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "SSTab2"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Check1(4)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "FrameDatosFiscales"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).ControlCount=   61
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
         Height          =   3435
         Left            =   270
         TabIndex        =   125
         Top             =   1140
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
            Index           =   36
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   133
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
            Index           =   36
            Left            =   1470
            TabIndex        =   131
            Tag             =   "País|T|S|||cobros|codpais|||"
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
            Index           =   37
            Left            =   1470
            TabIndex        =   132
            Tag             =   "Nif|T|S|||cobros|nifclien|||"
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
            Index           =   38
            Left            =   1470
            TabIndex        =   130
            Tag             =   "Provincia|T|S|||cobros|proclien|||"
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
            Index           =   39
            Left            =   4020
            TabIndex        =   129
            Tag             =   "Poblacion|T|S|||cobros|pobclien|||"
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
            Index           =   40
            Left            =   1470
            TabIndex        =   128
            Tag             =   "CP|T|S|||cobros|cpclien|||"
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
            Index           =   41
            Left            =   1470
            TabIndex        =   127
            Tag             =   "Dirección|T|S|||cobros|domclien|||"
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
            Index           =   42
            Left            =   1470
            TabIndex        =   126
            Tag             =   "Nombre|T|S|||cobros|nomclien|||"
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
            TabIndex        =   140
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
            TabIndex        =   139
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
            TabIndex        =   138
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
            TabIndex        =   137
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
            TabIndex        =   136
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
            TabIndex        =   135
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
            TabIndex        =   134
            Top             =   450
            Width           =   1545
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "NO remesar"
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
         Left            =   13590
         TabIndex        =   33
         Tag             =   "s|N|S|||cobros|noremesar|||"
         Top             =   5040
         Width           =   1545
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2475
         Left            =   360
         TabIndex        =   90
         Top             =   5550
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4366
         _Version        =   393216
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
         TabCaption(0)   =   "Cobros Realizados"
         TabPicture(0)   =   "frmTESCobros.frx":006C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FrameAux0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Devoluciones"
         TabPicture(1)   =   "frmTESCobros.frx":0088
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrameAux1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Fechas Op.Asegurada"
         TabPicture(2)   =   "frmTESCobros.frx":00A4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrameSeguro"
         Tab(2).ControlCount=   1
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
            TabIndex        =   109
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
               TabIndex        =   118
               Tag             =   "Banco Talon/Pag|T|S|||cobros_realizados|bancotalonpag|||"
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
               TabIndex        =   117
               Tag             =   "Ref. Talon/Pag|T|S|||cobros_realizados|reftalonpag|||"
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
               TabIndex        =   113
               Tag             =   "Cta.Real Cobro|T|S|||cobros_realizados|ctabanc2|||"
               Text            =   "cta.cobro"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Frame FrameToolAux 
               Height          =   555
               Index           =   0
               Left            =   0
               TabIndex        =   147
               Top             =   0
               Width           =   1335
               Begin MSComctlLib.Toolbar ToolbarAux 
                  Height          =   330
                  Left            =   180
                  TabIndex        =   148
                  Top             =   150
                  Width           =   915
                  _ExtentX        =   1614
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
               TabIndex        =   115
               Tag             =   "F.Realizado|F|S|||cobros_realizados|fecrealizado|dd/mm/yyyy||"
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
               TabIndex        =   116
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
               TabIndex        =   146
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
               TabIndex        =   145
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
               TabIndex        =   119
               Tag             =   "Importe Cobro|T|S|||cobros_realizados|impcobro|###,##0.00||"
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
               TabIndex        =   124
               Tag             =   "Serie|T|N|||cobros_devolucion|numserie||S|"
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
               TabIndex        =   123
               Tag             =   "Linea|N|N|0||cobros_devolucion|numlinea||S|"
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
               TabIndex        =   122
               Tag             =   "Nº Factura|N|N|||cobros_devolucion|numfactul|000000|S|"
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
               TabIndex        =   121
               Tag             =   "Fecha Factura|F|N|||cobros_devolucion|fecfactu|dd/mm/yyyy|S|"
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
               TabIndex        =   120
               Tag             =   "Nº Vencimiento|N|N|0||cobros_devolucion|numorden||S|"
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
               TabIndex        =   114
               Tag             =   "Usuario Cobro|T|S|||cobros_realizados|usuariocobro|||"
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
               TabIndex        =   112
               Tag             =   "Nº asiento|N|S|0||cobros_realizados|numasien|00000000||"
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
               TabIndex        =   111
               Tag             =   "Fecha entrada|F|S|||cobros_realizados|fechaent|dd/mm/yyyy||"
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
               TabIndex        =   110
               Tag             =   "Diario|N|S|0||cobros_realizados|numdiari|00||"
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
            Begin MSComctlLib.ListView lwCobros 
               Height          =   1425
               Left            =   -30
               TabIndex        =   155
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
         Begin VB.Frame FrameSeguro 
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
            Height          =   945
            Left            =   -74910
            TabIndex        =   100
            Top             =   540
            Width           =   11415
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
               Index           =   23
               Left            =   2010
               TabIndex        =   104
               Tag             =   "Fecha Aviso falta pago|F|S|||cobros|feccomunica|dd/mm/yyyy||"
               Text            =   "Text1"
               Top             =   300
               Width           =   1275
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
               Index           =   22
               Left            =   4710
               TabIndex        =   103
               Tag             =   "Aviso prorroga|F|S|||cobros|fecprorroga|dd/mm/yyyy||"
               Text            =   "Text1"
               Top             =   300
               Width           =   1275
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
               Left            =   7110
               TabIndex        =   102
               Tag             =   "Aviso siniestro|F|S|||cobros|fecsiniestro|dd/mm/yyyy||"
               Text            =   "Text1"
               Top             =   300
               Width           =   1275
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
               Index           =   20
               Left            =   9900
               TabIndex        =   101
               Tag             =   "Fecha ult ejecucion|F|S|||cobros|fecejecutiva|dd/mm/yyyy||"
               Text            =   "Text1"
               Top             =   300
               Width           =   1275
            End
            Begin VB.Label Label1 
               Caption         =   "Comunicación"
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
               Index           =   24
               Left            =   210
               TabIndex        =   108
               Top             =   330
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Prorroga"
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
               Left            =   3480
               TabIndex        =   107
               Top             =   330
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "Aviso"
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
               Left            =   6210
               TabIndex        =   106
               Top             =   330
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "Ejecutiva"
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
               Index           =   27
               Left            =   8610
               TabIndex        =   105
               Top             =   300
               Width           =   945
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   4
               Left            =   1740
               Picture         =   "frmTESCobros.frx":00C0
               Top             =   300
               Width           =   240
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   6
               Left            =   6810
               Picture         =   "frmTESCobros.frx":014B
               Top             =   330
               Width           =   240
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   5
               Left            =   4410
               Picture         =   "frmTESCobros.frx":01D6
               Top             =   300
               Width           =   240
            End
            Begin VB.Image imgFecha 
               Height          =   240
               Index           =   7
               Left            =   9630
               Picture         =   "frmTESCobros.frx":0261
               Top             =   330
               Width           =   240
            End
         End
         Begin VB.Frame FrameAux1 
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
            Height          =   1995
            Left            =   -74850
            TabIndex        =   91
            Top             =   330
            Width           =   14295
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
               Height          =   290
               Index           =   4
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   159
               Top             =   1080
               Visible         =   0   'False
               Width           =   1740
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
               Height          =   290
               Index           =   12
               Left            =   8130
               MaxLength       =   255
               TabIndex        =   158
               Tag             =   "TipoRemesa|N|S|||hlinapu|tiporem|||"
               Text            =   "Tipo"
               Top             =   780
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Frame FrameToolAux 
               Height          =   555
               Index           =   1
               Left            =   0
               TabIndex        =   156
               Top             =   -30
               Width           =   855
               Begin MSComctlLib.Toolbar ToolbarAux1 
                  Height          =   330
                  Left            =   210
                  TabIndex        =   157
                  Top             =   180
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   582
                  ButtonWidth     =   609
                  ButtonHeight    =   582
                  Style           =   1
                  _Version        =   393216
                  BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                     NumButtons      =   1
                     BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Object.ToolTipText     =   "Modificar"
                     EndProperty
                  EndProperty
               End
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
               Index           =   2
               Left            =   4590
               TabIndex        =   154
               ToolTipText     =   "Buscar código"
               Top             =   780
               Visible         =   0   'False
               Width           =   195
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
               Height          =   290
               Index           =   11
               Left            =   11910
               MaxLength       =   255
               TabIndex        =   153
               Tag             =   "Gastos Dev|N|S|||hlinapu|gastodev|###,##0.00||"
               Text            =   "Gastos"
               Top             =   780
               Visible         =   0   'False
               Width           =   1845
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
               Height          =   290
               Index           =   10
               Left            =   11040
               MaxLength       =   255
               TabIndex        =   152
               Tag             =   "Año Remesa|N|S|||hlinapu|anyorem|###0||"
               Text            =   "Año"
               Top             =   780
               Visible         =   0   'False
               Width           =   795
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
               Height          =   290
               Index           =   9
               Left            =   9510
               MaxLength       =   255
               TabIndex        =   151
               Tag             =   "Remesa|N|S|||hlinapu|codrem|######0||"
               Text            =   "Remesa"
               Top             =   780
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.TextBox txtaux1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
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
               Left            =   4830
               MaxLength       =   255
               TabIndex        =   141
               Text            =   "nomconcedevol"
               Top             =   780
               Visible         =   0   'False
               Width           =   2625
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
               Height          =   290
               Index           =   6
               Left            =   2850
               MaxLength       =   10
               TabIndex        =   98
               Tag             =   "Código Devolucion|T|S|||hlinapu|coddevol|||"
               Text            =   "CodDev"
               Top             =   780
               Visible         =   0   'False
               Width           =   1740
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
               Height          =   290
               Index           =   0
               Left            =   45
               MaxLength       =   3
               TabIndex        =   97
               Tag             =   "Serie|T|N|||hlinapu|numserie||S|"
               Text            =   "ser"
               Top             =   765
               Visible         =   0   'False
               Width           =   375
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
               Height          =   290
               Index           =   7
               Left            =   7560
               MaxLength       =   255
               TabIndex        =   96
               Text            =   "Tipo"
               Top             =   780
               Visible         =   0   'False
               Width           =   495
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
               Height          =   290
               Index           =   1
               Left            =   450
               MaxLength       =   10
               TabIndex        =   95
               Tag             =   "Nº Factura|N|N|||hlinapu|numfaccl|000000|S|"
               Text            =   "fac"
               Top             =   750
               Visible         =   0   'False
               Width           =   375
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
               Height          =   290
               Index           =   2
               Left            =   840
               MaxLength       =   10
               TabIndex        =   94
               Tag             =   "Fecha Factura|F|N|||hlinapu|fecfactu|dd/mm/yyyy|S|"
               Text            =   "Fec"
               Top             =   750
               Visible         =   0   'False
               Width           =   375
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
               Height          =   290
               Index           =   3
               Left            =   1230
               MaxLength       =   30
               TabIndex        =   93
               Tag             =   "Nº Vencimiento|N|N|0||hlinapu|numorden||S|"
               Text            =   "vto"
               Top             =   750
               Visible         =   0   'False
               Width           =   375
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
               Height          =   290
               Index           =   5
               Left            =   2100
               MaxLength       =   10
               TabIndex        =   92
               Tag             =   "Fecha Devolucion|F|N|||hlinapu|fecdevol|||"
               Text            =   "fec"
               Top             =   750
               Visible         =   0   'False
               Width           =   390
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   1
               Left            =   3720
               Top             =   225
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
            Begin MSDataGridLib.DataGrid DataGridAux 
               Bindings        =   "frmTESCobros.frx":02EC
               Height          =   1335
               Index           =   1
               Left            =   0
               TabIndex        =   99
               Top             =   570
               Width           =   13545
               _ExtentX        =   23892
               _ExtentY        =   2355
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
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Devuelto"
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
         Left            =   13590
         TabIndex        =   32
         Tag             =   "Devuelto|N|S|||cobros|Devuelto|||"
         Top             =   4740
         Width           =   1215
      End
      Begin VB.Frame FrameRemesa 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   9660
         TabIndex        =   60
         Top             =   3270
         Width           =   5715
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
            Height          =   370
            Index           =   27
            Left            =   4170
            MaxLength       =   10
            TabIndex        =   29
            Tag             =   "Transferencia|N|S|0||cobros|transfer|0000000000||"
            Text            =   "Text1"
            Top             =   840
            Width           =   1425
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
            ItemData        =   "frmTESCobros.frx":0304
            Left            =   1290
            List            =   "frmTESCobros.frx":0306
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   840
            Width           =   2835
         End
         Begin VB.ComboBox cboTipoRem 
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
            ItemData        =   "frmTESCobros.frx":0308
            Left            =   1290
            List            =   "frmTESCobros.frx":0315
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Tag             =   "Remesa|N|S|0||cobros|tiporem|||"
            Top             =   240
            Width           =   1455
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
            Index           =   15
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   37
            Tag             =   "Situacion|T|S|||cobros|siturem|||"
            Text            =   "Text1"
            Top             =   870
            Width           =   885
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
            Height          =   370
            Index           =   14
            Left            =   4170
            MaxLength       =   4
            TabIndex        =   27
            Tag             =   "Año remesa|N|S|0||cobros|anyorem|||"
            Text            =   "Text"
            Top             =   240
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
            Height          =   370
            Index           =   35
            Left            =   2820
            MaxLength       =   30
            TabIndex        =   26
            Tag             =   "Remesa|N|S|0||cobros|codrem|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1275
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
            Index           =   43
            Left            =   1320
            MaxLength       =   7
            TabIndex        =   160
            Tag             =   "Usuario|N|S|||cobros|codusu|######0||"
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Transferencia"
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
            Index           =   19
            Left            =   4170
            TabIndex        =   149
            Top             =   600
            Width           =   1380
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
            Left            =   60
            TabIndex        =   64
            Top             =   870
            Width           =   1080
         End
         Begin VB.Label Label4 
            Caption         =   "REMESA"
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
            Left            =   60
            TabIndex        =   63
            Top             =   270
            Width           =   855
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
            Left            =   4170
            TabIndex        =   62
            Top             =   0
            Width           =   540
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
            Height          =   195
            Index           =   15
            Left            =   2850
            TabIndex        =   61
            Top             =   0
            Width           =   1170
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
         TabIndex        =   23
         Tag             =   "Importe|N|S|||cobros|impcobro|#,##0.00||"
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
         TabIndex        =   21
         Tag             =   "Fecha ult. pago|F|S|||cobros|fecultco|dd/mm/yyyy||"
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
         TabIndex        =   16
         Tag             =   "CSB|T|S|||cobros|text33csb|||"
         Text            =   "WWW4567890WWW4567890WWW4567890WWW456789WWWW4567890WWW4567890WWW4567890WWW456789J"
         Top             =   3510
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
         TabIndex        =   17
         Tag             =   "T|T|S|||cobros|text41csb|||"
         Top             =   4200
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
         TabIndex        =   72
         Text            =   "Text4"
         Top             =   2850
         Width           =   1425
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
         Index           =   34
         Left            =   360
         TabIndex        =   8
         Tag             =   "Agente|N|S|0||cobros|agente|||"
         Text            =   "Text1"
         Top             =   2130
         Width           =   795
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
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   2130
         Width           =   3735
      End
      Begin VB.Frame frameContene 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   9720
         TabIndex        =   68
         Top             =   4560
         Width           =   3675
         Begin VB.CheckBox Check1 
            Caption         =   "Documento recibido"
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
            TabIndex        =   31
            Tag             =   "Recibido|N|S|||cobros|recedocu|||"
            Top             =   450
            Width           =   2505
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Situacion jurídica"
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
            Left            =   60
            TabIndex        =   30
            Tag             =   "s|N|S|||cobros|situacionjuri|||"
            Top             =   150
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
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
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
            Width           =   2205
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
         Height          =   555
         Index           =   17
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "obs|T|S|||cobros|observa|||"
         Text            =   "frmTESCobros.frx":0334
         Top             =   4920
         Width           =   9225
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
         Index           =   16
         Left            =   10950
         MaxLength       =   30
         TabIndex        =   24
         Tag             =   "Gastos|N|S|||cobros|gastos|#,###,##0.00||"
         Text            =   "1.999.999.00"
         Top             =   2850
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
         Index           =   32
         Left            =   7710
         TabIndex        =   15
         Tag             =   "Ultima reclamacion|F|S|||cobros|ultimareclamacion|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   2850
         Width           =   1875
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
         Index           =   4
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text2"
         Top             =   1440
         Width           =   8025
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
         Left            =   360
         TabIndex        =   7
         Tag             =   "departamento|N|S|||cobros|departamento|||"
         Text            =   "Text1"
         Top             =   1440
         Width           =   1095
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
         Left            =   5100
         TabIndex        =   34
         Text            =   "0000"
         Top             =   2130
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
         TabIndex        =   52
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
         TabIndex        =   14
         Tag             =   "Cta prevista|T|N|||cobros|ctabanc1|||"
         Text            =   "Text1"
         Top             =   2850
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
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   2850
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
         TabIndex        =   22
         Tag             =   "Importe|N|N|||cobros|impvenci|#,###,##0.00||"
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
         TabIndex        =   20
         Tag             =   "Fecha vencimiento|F|N|||cobros|fecvenci|dd/mm/yyyy||"
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
         TabIndex        =   19
         Tag             =   "Forma Pago|N|N|0||cobros|codforpa|000||"
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
         TabIndex        =   50
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
         Tag             =   "Cta. cliente|T|N|||cobros|codmacta|||"
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
         TabIndex        =   49
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
         Left            =   6630
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "0000"
         Top             =   2130
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
         Left            =   7380
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "0000"
         Top             =   2130
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
         Left            =   8145
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "0000"
         Top             =   2130
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
         Index           =   31
         Left            =   8910
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "0000"
         Top             =   2130
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
         Left            =   5865
         MaxLength       =   20
         TabIndex        =   35
         Text            =   "0000"
         Top             =   2130
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
         Height          =   360
         Index           =   19
         Left            =   5790
         MaxLength       =   40
         TabIndex        =   9
         Tag             =   "Iban|T|S|||cobros|iban|||"
         Text            =   "ES99"
         Top             =   1440
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
         ItemData        =   "frmTESCobros.frx":033A
         Left            =   5190
         List            =   "frmTESCobros.frx":0347
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Situación|N|N|||cobros|situacion|||"
         Top             =   720
         Width           =   2175
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
         TabIndex        =   150
         Top             =   420
         Width           =   915
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   1890
         Top             =   4620
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   9270
         Picture         =   "frmTESCobros.frx":037E
         Top             =   2580
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   13710
         Picture         =   "frmTESCobros.frx":0409
         Top             =   1410
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   10890
         Picture         =   "frmTESCobros.frx":0494
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   3960
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
         TabIndex        =   77
         Top             =   3270
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
         Left            =   5160
         TabIndex        =   76
         Top             =   1860
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
         TabIndex        =   73
         Top             =   2880
         Width           =   1245
      End
      Begin VB.Image imgAgente 
         Height          =   255
         Left            =   1170
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
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
         Index           =   14
         Left            =   360
         TabIndex        =   70
         Top             =   1860
         Width           =   705
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
         TabIndex        =   67
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Gastos"
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
         Index           =   18
         Left            =   9720
         TabIndex        =   66
         Top             =   2850
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Ultima reclam."
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
         Left            =   7770
         TabIndex        =   65
         Top             =   2580
         Width           =   1485
      End
      Begin VB.Image imgDepart 
         Height          =   240
         Left            =   1860
         ToolTipText     =   "Buscar departamento"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
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
         Index           =   12
         Left            =   360
         TabIndex        =   59
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   2700
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Prevista Cobro"
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
         TabIndex        =   57
         Top             =   2580
         Width           =   2220
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   420
         Width           =   1470
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   2100
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Cliente"
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
         TabIndex        =   53
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   120
      TabIndex        =   41
      Top             =   10050
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
         TabIndex        =   42
         Top             =   210
         Width           =   3675
      End
   End
   Begin VB.Frame FrameClaves 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   765
      Left            =   120
      TabIndex        =   43
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
         Index           =   25
         Left            =   10500
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Referencia1|T|S|||cobros|referencia1|||"
         Text            =   "Text1"
         Top             =   270
         Width           =   2145
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
         Left            =   12840
         MaxLength       =   15
         TabIndex        =   40
         Tag             =   "Referencia2|T|S|||cobros|referencia2|||"
         Text            =   "Text1"
         Top             =   270
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
         Index           =   18
         Left            =   8100
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Referencia|T|S|0||cobros|referencia|||"
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
         Tag             =   "Serie|T|N|||cobros|numserie||S|"
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
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº Factura|N|N|||cobros|numfactu|0000000|S|"
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
         Tag             =   "Nº Vencimiento|N|N|0||cobros|numorden||S|"
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
         Tag             =   "Fecha Factura|F|N|||cobros|fecfactu|dd/mm/yyyy|S|"
         Text            =   "Text1"
         Top             =   270
         Width           =   1275
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3780
         Picture         =   "frmTESCobros.frx":051F
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia (I)"
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
         Index           =   22
         Left            =   10500
         TabIndex        =   75
         Top             =   0
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia (II)"
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
         Index           =   23
         Left            =   12840
         TabIndex        =   74
         Top             =   0
         Width           =   1770
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
         Left            =   8100
         TabIndex        =   71
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
         Top             =   0
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   15060
      TabIndex        =   87
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
Attribute VB_Name = "frmTESCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 601


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
    CargaGrid 1, False
    
    HacerBusqueda2
End Sub

Private Sub cboSituRem_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboSituRem_Validate(Cancel As Boolean)
    If (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        If cboSituRem.ListIndex = 0 Then
            Text1(15).Text = ""
        Else
            If cboSituRem.ListIndex <> -1 Then Text1(15).Text = Chr(cboSituRem.ItemData(cboSituRem.ListIndex))
        End If
    End If

End Sub


Private Sub cboTipoRem_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
    Dim cad As String
    Dim i As Integer
    
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
             If ModificaDesdeFormulario2(Me, 1) Then
                'TerminaBloquear
                DesBloqueaRegistroForm Me.Text1(0)
                lblIndicador.Caption = ""
                If SituarData Then
                
                    Text1_LostFocus 0
                    cad = vTipForpa 'para que no pierda el valor
                    PonerModo 2
                    vTipForpa = cad
                    cad = ""
                    PonPendiente
                    '-- Esto permanece para saber donde estamos
                    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

                Else
                    LimpiarCampos
                    'PonerModo 0
                End If
            End If
        End If
    
    Case 5
        
        Select Case ModoLineas
            Case 1 'afegir llínia
                InsertarLinea
            Case 2 'modificar llínies
                ModificarLinea
                                    
                '**** parte de contabilizacion de la factura
                TerminaBloquear
                
            
                'LOG
                vLog.Insertar 5, vUsu, Text1(2).Text & Text1(0).Text & " " & Text1(1).Text
                PosicionarData
                
        End Select
            
    
    
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
            
            PonFoco txtaux(5)
            
        Case 2 ' cocnepto de devolucion
            Set frmDev = New frmBasico
            
            AyudaDevolucion frmDev, txtaux1(6)
            
            Set frmDev = Nothing
            
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
            TerminaBloquear
        
            If ModoLineas = 1 Then 'INSERTAR
                ModoLineas = 0
                DataGridAux(0).AllowAddNew = False
            End If
            ModoLineas = 0
            LLamaLineas 0, 0, 0
            
            Modo = 2   'Para que el lostfocus NO haga nada
            
            PosicionarData
            PonerCampos
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
    
    Check1(1).Value = 0
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    
    CargaList 0, False 'CargaGrid 0, False
    CargaGrid 1, False
    
    
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    
    'metemos el codusu
    Text1(43).Text = vUsu.Id
    
    Combo1.ListIndex = 0
    Text2(3).Text = Combo1.Text
    cboSituRem.ListIndex = -1
    '###A mano
    PonFoco Text1(13)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        CargaList 0, False
        CargaGrid 1, False
        
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
    CargaGrid 1, False
    
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

    N = SePuedeEliminar2()
    If N = 0 Then Exit Sub


    If Not BloqueaRegistroForm(Me) Then Exit Sub
    '---------
    'MODIFICAR
    '----------
    'Añadiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    
    'Si se puede modificar entonces habilito todooos los campos
    
    PonerModo 4
    If N < 3 Then
        'Se puede modifcar la CC
        Dim T As TextBox
        For Each T In Text1
            If (T.Index < 28 Or T.Index > 31) And T.Index <> 10 And T.Index <> 26 Then
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
    Dim cad As String
    Dim i As Integer
    Dim Sql As String
    Dim SqlLog As String
    

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'Comprobamos si se puede eliminar
    i = SePuedeEliminar2
    If i < 3 Then Exit Sub
    '### a mano
    cad = "Seguro que desea eliminar de la BD el registro actual:"
    cad = cad & vbCrLf & Data1.Recordset.Fields(0) & "  " & Data1.Recordset.Fields(1) & " "
    cad = cad & Data1.Recordset.Fields(2) & "  " & Data1.Recordset.Fields(3)
    i = MsgBox(cad, vbQuestion + vbYesNoCancel + vbDefaultButton2)
    'Borramos
    If i = vbYes Then
        'Borro el elemento
        Sql = "Delete from cobros  WHERE numserie = '" & Data1.Recordset!NUmSerie & "' AND numfactu = " & Data1.Recordset!NumFactu
        Sql = Sql & " AND fecfactu = " & DBSet(Data1.Recordset!FecFactu, "F") & " AND numorden =" & Data1.Recordset!numorden
        
        DataGridAux(1).Enabled = False
        NumRegElim = Data1.Recordset.AbsolutePosition
        Conn.Execute Sql

        SqlLog = "Serie      : " & Text1(13).Text
        SqlLog = SqlLog & vbCrLf & "Factura    : " & Text1(1).Text
        SqlLog = SqlLog & vbCrLf & "Fecha      : " & Text1(2).Text
        SqlLog = SqlLog & vbCrLf & "Vencimiento: " & Text1(3).Text & vbCrLf
        SqlLog = SqlLog & vbCrLf & "Cliente    : " & Text1(4).Text & " " & Text2(0).Text
        SqlLog = SqlLog & vbCrLf & "Importe    : " & Text1(6) & vbCrLf
        
        vLog.Insertar 23, vUsu, SqlLog


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
                DataGridAux(1).Enabled = True
        End If
        
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim impo As Currency
    
    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    
    If SePuedeEliminar2 < 3 Then Exit Sub

    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    If impo < 0 Then
        MsgBox "Los abonos no se realizan por caja", vbExclamation
        Exit Sub
    End If


    'Mas gastos
    If Text1(38).Text <> "" Then impo = impo + ImporteFormateado(Text1(38).Text)
    'Menos ya pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)
    
    If impo <= 0 Then
        MsgBox "Totalmente cobrado", vbExclamation
        Exit Sub
    End If
    
    'Devolvera muuuuchas cosas
    'serie factura fecfac numvto
    cad = Text1(13).Text & "|" & Format(Text1(1).Text, "0000000") & "|" & Text1(2).Text & "|" & Text1(3).Text & "|"
    'Codmacta nommacta codforpa   nomforpa   importe
    cad = cad & Text1(4).Text & "|" & Text2(0).Text & "|" & Text1(0).Text & "|" & Text2(1).Text & "|" & CStr(impo) & "|"
    'Lo que lleva cobrado
    cad = cad & Text1(8).Text & "|"
    
    
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub



Private Sub DataGridAux_DblClick(Index As Integer)
    If Index = 0 Then BotonVerAsiento
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
        cboFiltro.ListIndex = vUsu.FiltroCobros
    
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
    
    With Me.ToolbarAux1
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 4
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
    imgDepart.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgAgente.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    Me.SSTab1.Tab = 0
    Me.Icon = frmppal.Icon
    LimpiarCampos
    Me.SSTab2.TabEnabled(2) = vParamT.TieneOperacionesAseguradas
    
    'Recaudacion ejecutiva
    Label1(27).Visible = vParamT.RecaudacionEjecutiva
    Text1(20).Visible = vParamT.RecaudacionEjecutiva
    imgFecha(7).Visible = vParamT.RecaudacionEjecutiva
    
    
    '## A mano
    NombreTabla = "cobros"
    Ordenacion = " ORDER BY numserie,numfactu,fecfactu,numorden"
        
    PonerOpcionesMenu
    
    CargaList 0, False
    CargaGrid 1, False
    
    
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
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    cboTipoRem.ListIndex = -1
    Combo1.ListIndex = -1
    Text2(3).Text = ""
    lblIndicador.Caption = ""
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    vUsu.ActualizarFiltro "ariconta", IdPrograma, Me.cboFiltro.ListIndex
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Text1(34).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cad As String

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
            cad = DevfrmCCtas
            DevfrmCCtas = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            cad = cad & " AND " & DevfrmCCtas
            DevfrmCCtas = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
            cad = cad & " AND " & DevfrmCCtas
            DevfrmCCtas = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
            cad = cad & " AND " & DevfrmCCtas
            DevfrmCCtas = cad
            If DevfrmCCtas = "" Then Exit Sub
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & DevfrmCCtas & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    Else
        DevfrmCCtas = ""
    End If
End Sub

Private Sub PonerDatoDevuelto(CadenaDevuelta As String)
Dim cad As String
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(13), CadenaDevuelta, 1)
    cad = DevfrmCCtas
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
    cad = cad & " AND " & DevfrmCCtas
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
    cad = cad & " AND " & DevfrmCCtas
    DevfrmCCtas = ValorDevueltoFormGrid(Text1(3), CadenaDevuelta, 4)
    cad = cad & " AND " & DevfrmCCtas
    DevfrmCCtas = cad
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
    txtaux(CInt(cmdAux(1).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    DevfrmCCtas = CadenaSeleccion
End Sub

Private Sub frmDev_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtaux1(6).Text = RecuperaValor(CadenaSeleccion, 1)
        txtaux1(8).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
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
Dim cad As String
Dim Z
    Screen.MousePointer = vbHourglass
    If Index = 1 Then
       
        Set frmF = New frmFormaPago
        frmF.DatosADevolverBusqueda = "0|"
        frmF.Show vbModal
        Set frmF = Nothing
    
        
    
    Else
        'Cuentas
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
    'En tag pongo el txtfecha asociado
    Select Case Index
    Case 0
        imgFecha(0).Tag = 2
    Case 1
        imgFecha(0).Tag = 5
    Case 2
        imgFecha(0).Tag = 7
    Case 3
        imgFecha(0).Tag = 32
    Case 4
        imgFecha(0).Tag = 23
    Case 5
        imgFecha(0).Tag = 22
    Case 6
        imgFecha(0).Tag = 21
    Case 7
        imgFecha(0).Tag = 20
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
        frmZ.Caption = "Observaciones Cobros"
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
    
        Set frmConta = New frmBasico
        AyudaContadores frmConta, Text1(13).Text, "tiporegi REGEXP '^[0-9]+$' = 0"
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
    If Index = 24 Then
        'Despues de la fecha prorroga va el btn
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
    
    
    If Not (Index = 4 Or Index = 9 Or Index = 10 Or Index = 26 Or Index = 28 Or Index = 29 Or Index = 30 Or Index = 31) Then
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
                        Text1(31).Text = Mid(RecuperaValor(Sql, 1), 21, 4)

                        Text1(19).Text = RecuperaValor(Sql, 1)
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
            
            
            
                If Modo = 3 Then
                    If Trim(Text1(Index).Text) <> CtaAnt Then CargarDatosCuenta Text1(Index), False
                Else
                    If Modo = 4 Then
                        'No ha puesto los datos fiscales
                        If Text1(42).Text = "" Then
                            If Trim(Text1(Index).Text) <> CtaAnt Then
                                i = 1 'Cargamos todo
                            Else
                                i = 0 'Solo datos fiscales, NO forma pago
                            End If
                        
                            CargarDatosCuenta Text1(Index), i = 0
                        End If
                    End If
                End If
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
        
        
    Case 2, 5, 7, 32, 23, 22, 20, 21
        'FECHAS,32
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
        
    Case 6, 8, 16
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
        If IsNumeric(Text1(13).Text) Then
            MsgBox "Serie es una letra.", vbExclamation
            Text1(13).Text = ""
            PonerFoco Text1(13)
        Else
            Text1(13).Text = UCase(Text1(13).Text)
        End If
        

    Case 28 To 31, 10, 26
        'Cuenta bancaria
        If Index <> 10 Then
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Cuenta banco debe ser numérico: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            Else
                'Formateamos
                If Text1(Index).Text <> "" Then
                    Text1(Index).Text = Format(Text1(Index).Text, "0000")
                End If
            End If
        Else
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
        End If
        
        Sql = Text1(26).Text & Text1(28).Text & Text1(29).Text & Text1(30).Text & Text1(31).Text
        
        If Len(Sql) = 20 And Index = 31 Then 'solo cuando pierde el foco la cuentaban
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
        
        Text1(19).Text = Text1(10).Text & Text1(26).Text & Text1(28).Text & Text1(29).Text & Text1(30).Text & Text1(31).Text
        
        
    Case 33
        
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Departamento debe ser numérico: " & Text1(Index).Text, vbExclamation
            i = 0
        Else
            i = 1
            PonerDepartamenteo
            If Text2(4).Text = "" Then i = 0
        End If
        If i = 0 Then
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
            Text2(4).Text = ""
        End If
        
    Case 34
        i = 0
        If Text1(34).Text <> "" Then
            Sql = DevuelveDesdeBD("nombre", "agentes", "codigo", Text1(Index).Text, "N")
            If Sql = "" Then
                MsgBox "No existe el agente: " & Text1(34).Text, vbExclamation
                i = 2
            Else
                i = 1
            End If
        Else
            Sql = ""
        End If
        Text2(5).Text = Sql
        If i = 2 Then PonerFoco Text1(34)
            
    Case 19
        Text1(Index).Text = UCase(Text1(Index).Text)
        
    Case 36 ' codigo de pais
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
Dim cad As String

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
    If Text1(31).Text <> "" Then
        If CadB <> "" Then CadB = CadB & " and "
        CadB = CadB & "mid(iban,21,4) = " & DBSet(Text1(31).Text, "T")
    End If

    
    CadB1 = ObtenerBusqueda2(Me, , 2, "FrameAux1")
    
    HacerBusqueda2


End Sub

Private Sub HacerBusqueda2()

    CargarSqlFiltro
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Or CadB1 <> "" Or cadFiltro <> "" Then
        
        
            
        If CadB1 <> "" Then
            CadenaConsulta = "select distinct cobros.* from "
            CadenaConsulta = CadenaConsulta & " (tipofpago  INNER JOIN hlinapu ON hlinapu.tipforpa = tipofpago.tipoformapago ) "
            CadenaConsulta = CadenaConsulta & " right join cobros on cobros.numserie = hlinapu.numserie and cobros.numfactu = hlinapu.numfaccl and cobros.fecfactu = hlinapu.fecfactu and cobros.numorden = hlinapu.numorden"
        Else
            CadenaConsulta = "select distinct cobros.* from cobros "
        End If
        
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
    
'    CargaDatosLW
    

End Sub

Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    cadFiltro = ""
    
    Select Case cboFiltro.ItemData(cboFiltro.ListIndex) 'cboFiltro.ListIndex
        Case 0 ' pendientes de cobro
            cadFiltro = "(coalesce(cobros.impvenci,0) + coalesce(cobros.gastos,0) - coalesce(cobros.impcobro,0) <> 0) "
        Case 1 ' cobrados
            cadFiltro = "(coalesce(cobros.impvenci,0) + coalesce(cobros.gastos,0) - coalesce(cobros.impcobro,0) = 0) and cobros.situacion = 1"
        Case 9 ' todos
            cadFiltro = "(1=1)"
    End Select
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim vCadena As String
Dim frmTESVerCobPag As frmTESVerCobrosPagos


    CadenaDesdeOtroForm = ""
    
    vCadena = cadFiltro  '"(coalesce(Cobros.ImpVenci, 0) + coalesce(Cobros.Gastos, 0) - coalesce(Cobros.impcobro, 0) <> 0)"
    
    Set frmTESVerCobPag = New frmTESVerCobrosPagos
    
    
    frmTESVerCobPag.Situacion = 1
    frmTESVerCobPag.vSql = vCadena
    If CadB <> "" Then frmTESVerCobPag.vSql = frmTESVerCobPag.vSql & " and " & CadB
    frmTESVerCobPag.OrdenarEfecto = False
    frmTESVerCobPag.Regresar = True
    frmTESVerCobPag.Cobros = True
    frmTESVerCobPag.Show vbModal
    
    Set frmTESVerCobPag = Nothing
    
    If CadenaDesdeOtroForm <> "" Then
        PonerDatoDevuelto CadenaDesdeOtroForm
        If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then cmdRegresar_Click
    Else   'de ha devuelto datos, es decir NO ha devuelto datos
        ' Text1(kCampo).SetFocus
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
    
    Text2(36).Text = PonerNombreDeCod(Text1(36), "paises", "nompais", "codpais", "T")
    
    Text2(5).Text = PonerNombreDeCod(Text1(34), "agentes", "nombre", "codigo", "N")
    Text2(1).Text = PonerNombreDeCod(Text1(0), "formapago", "nomforpa", "codforpa", "N")
    vTipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Text1(0).Text, "N")
    
    PonerDepartamenteo
    
    
    If Text1(15).Text = "" Then
        cboSituRem.ListIndex = -1
    Else
        PosicionarCombo cboSituRem, Asc(Text1(15).Text)
    End If
    
    cboSituRem_Validate False
    
    
    
    
    ' situamos los campos del iban
    Text1(10).Text = ""
    Text1(26).Text = ""
    Text1(28).Text = ""
    Text1(29).Text = ""
    Text1(30).Text = ""
    Text1(31).Text = ""
    
    Text1(10).ToolTipText = ""
    Text1(26).ToolTipText = ""
    Text1(28).ToolTipText = ""
    Text1(29).ToolTipText = ""
    Text1(30).ToolTipText = ""
    Text1(31).ToolTipText = ""
    
    If Text1(19).Text <> "" Then
        Text1(10) = Mid(Text1(19), 1, 4)
        Text1(26) = Mid(Text1(19), 5, 4)
        Text1(28) = Mid(Text1(19), 9, 4)
        Text1(29) = Mid(Text1(19), 13, 4)
        Text1(30) = Mid(Text1(19), 17, 4)
        Text1(31) = Mid(Text1(19), 21, 4)
        
        Dim CCC As String
        CCC = Text1(10).Text & " " & Text1(26).Text & " " & Text1(28).Text & " " & Mid(Text1(29).Text, 1, 2) & " " & Mid(Text1(29).Text, 3, 2) & Text1(30).Text & Text1(31).Text
        
        Text1(10).ToolTipText = CCC
        Text1(26).ToolTipText = CCC
        Text1(28).ToolTipText = CCC
        Text1(29).ToolTipText = CCC
        Text1(30).ToolTipText = CCC
        Text1(31).ToolTipText = CCC
    
    End If
    
    
    'Cargamos el LINEAS
    CargaList 0, True
    CargaGrid 1, True

    
    
    
    PonPendiente
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
End Sub


Private Sub PonPendiente()
Dim Importe As Currency

    On Error GoTo EPonPendiente
    'Pendiente
    Importe = Data1.Recordset!ImpVenci + DBLet(Data1.Recordset!Gastos, "N") - DBLet(Data1.Recordset!impcobro, "N")
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
        Me.SSTab2.TabEnabled(2) = True
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
    
    frameContene.Enabled = Not B
    
    For i = 0 To 6
        If i < 4 And i <> 3 Then imgCuentas(i).Visible = Not B
        Me.imgFecha(i).Visible = Not B
    Next i
    
    Me.imgSerie.Visible = Not B
    Me.imgDepart.Visible = Not B
    Me.imgAgente.Visible = Not B
        
    cboSituRem.Locked = B
    cboTipoRem.Locked = B
        
        
    Text2(4).Tag = ""
    
    'lineas de cobros_realizados
    Dim anc As Single
    anc = DataGridAux(1).top
    If DataGridAux(1).Row < 0 Then
        anc = anc + 230
    Else
        anc = anc + DataGridAux(1).RowTop(DataGridAux(1).Row) + 5
    End If
    If Modo = 1 Then
        LLamaLineas 1, Modo, anc
    Else
        LLamaLineas 1, 3, anc
    End If

    For i = 0 To txtaux1.Count - 1
        If i <> 8 Then txtaux1(i).BackColor = vbWhite
    Next i
    
    Check1(2).Enabled = (Modo = 1)
    Check1(4).Enabled = (Modo = 1)
    
    Combo1.Enabled = False '(Modo = 1) Or ((vUsu.Nombre = "root") And Modo = 4)
    
        
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Tipo As Integer

    DatosOK = False
    
    
    If cboSituRem.ListIndex = 0 Or cboSituRem.ListIndex = -1 Then
        Text1(15).Text = ""
    Else
        Text1(15).Text = Chr(cboSituRem.ItemData(cboSituRem.ListIndex))
    End If
    
    
    
    DevfrmCCtas = ""
    
    If Text1(34).Text = "" Then
        DevfrmCCtas = vbCrLf & "-  Agente "
        Tipo = 34
    End If
    
    If Text1(9).Text = "" Then
        DevfrmCCtas = DevfrmCCtas & vbCrLf & "-  Cuenta prevista cobro "
        Tipo = 9
    End If
    
    If Text1(4).Text = "" Then
        DevfrmCCtas = DevfrmCCtas & vbCrLf & "-  Cuenta cliente "
        Tipo = 4
    End If
    If DevfrmCCtas <> "" Then
        DevfrmCCtas = "Los siguientes campos son requeridos:" & vbCrLf & vbCrLf & DevfrmCCtas
        MsgBox DevfrmCCtas, vbExclamation
        PonerFoco Text1(Tipo)
        Exit Function
    End If
    
    vTipForpa = ""
    
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
    
    DevfrmCCtas = Trim(Text1(26).Text) & Trim(Text1(28).Text) & Trim(Text1(29).Text) & Trim(Text1(30).Text) & Trim(Text1(31).Text)
    
    'Para preguntar por el Banco
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
            '###FALTA
            'DevfrmCCtas = mid(DevfrmCCtas, 5, 20)
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
            
            
    Else
        If Tipo = vbTipoPagoRemesa Then
                DevfrmCCtas = "Debe poner cuenta bancaria. Desea continuar?"
                If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
    End If
    
   
        If Modo = 4 Then
            If DBLet(Me.Data1.Recordset!recedocu, "N") = 1 Then
                'Tiene la marca de documento recibido
                'Veremos si se la ha quitado
                If Me.Check1(0).Value = 0 Then
                    DevfrmCCtas = "Seguro que desea quitarle la marca de documento recibido?"
                    If MsgBox(DevfrmCCtas, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
            End If
        End If

    
    'Nuevo. 12 Mayo 2008
    B = CuentaBloqeada(Me.Text1(4).Text, CDate(Text1(2).Text), True)
    If B Then
        If (vUsu.Codigo Mod 100) > 0 Then Exit Function
    End If
    
    
    
    'Ultimas comprobaciones
    If vParamT.TieneOperacionesAseguradas Then
        B = Me.Text1(23).Text <> "" Or Me.Text1(22).Text <> "" Or Me.Text1(21).Text <> ""
        If B Then
             MsgBox "No debe indicar fechas de operaciones aseguradas" & vbCrLf & "-Falta pago/prorroga/aviso siniestro" & vbCrLf & " Si no esta asegurado", vbExclamation
             PonerFoco Me.Text1(23)
             Exit Function
        End If
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
            
            frmTESCobrosPdtesList.Show vbModal

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
Private Function SePuedeEliminar2() As Byte


    SePuedeEliminar2 = 0 'NO se puede eliminar

    SePuedeEliminar2 = 1
    If Val(DBLet(Data1.Recordset!CodRem)) > 0 Then
        MsgBox "Pertenece a una remesa", vbExclamation
        'Noviembre 2009
        If vUsu.Nivel < 2 Then
            If CStr(Data1.Recordset!siturem) = "Q" Or CStr(Data1.Recordset!siturem) = "Y" Then
                'DEJO ELIMINARLO
                If MsgBox("Efecto remesado. Situacion: " & Data1.Recordset!siturem & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Function
                espera 1
                If MsgBox("¿Seguro?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            Else
                'Tampoco dejamos continuar
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    'Si no esta en transferencia
    If Val(DBLet(Data1.Recordset!transfer)) > 0 Then
        MsgBox "Pertenece a una transferencia", vbExclamation
        Exit Function
    End If
    
    
    'Si  tiene documento recibido
    If Val(DBLet(Data1.Recordset!recedocu)) > 0 Then
        'Documento recibido
        '
        DevfrmCCtas = "numserie='" & Data1.Recordset!NUmSerie
        DevfrmCCtas = DevfrmCCtas & "' AND fecfactu='" & Format(Data1.Recordset!FecFactu, FormatoFecha)
        DevfrmCCtas = DevfrmCCtas & "' AND numfactu=" & Data1.Recordset!NumFactu
        DevfrmCCtas = DevfrmCCtas & " AND numorden"
        DevfrmCCtas = DevuelveDesdeBD("codigo", "talones_facturas", DevfrmCCtas, Data1.Recordset!numorden)
        If DevfrmCCtas <> "" Then
            DevfrmCCtas = "Esta en la recepcion de documentos. Numero: " & DevfrmCCtas
            MsgBox DevfrmCCtas, vbExclamation
            DevfrmCCtas = ""
            Exit Function
        End If
    End If
    
    
    
    
    
    SePuedeEliminar2 = 3  'SI SE PUEDE ELIMINAR

    Screen.MousePointer = vbDefault
End Function


Private Sub PonerDepartamenteo()
Dim C As String
Dim o As Boolean

    o = False
    
    If Text1(4).Text <> "" Then
        If Text1(33).Text <> "" Then
                    
            Set miRsAux = New ADODB.Recordset
            C = "Select Descripcion FROM Departamentos WHERE codmacta ='" & Text1(4).Text
            C = C & "' AND Dpto =" & Text1(33).Text
            miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                If Not IsNull(miRsAux.Fields(0)) Then
                    C = miRsAux.Fields(0)
                    o = True
                End If
            End If
            miRsAux.Close
            Set miRsAux = Nothing
        End If
    End If
    If o Then
        Text2(4).Text = C
    Else
        Text2(4).Text = ""
    End If
    
End Sub
    



Private Sub RealizarPagoCuenta()
Dim impo As Currency
    'Para realizar pago a cuenta... Varias cosas.
    'Primero. Hay por pagar
    impo = ImporteFormateado(Text1(6).Text)
    'Gastos
    If Text1(16).Text <> "" Then impo = impo + ImporteFormateado(Text1(16).Text)
    'Pagado
    If Text1(8).Text <> "" Then impo = impo - ImporteFormateado(Text1(8).Text)

    'Si impo>0 entonces TODAVIA puedn pagarme algo
    If impo = 0 Then
        'Cosa rara. Esta todo el importe pagado
        MsgBox "La factura está totalmente cobrada.", vbExclamation
        Exit Sub
    End If

    frmTESParciales.Cobro = True
    frmTESParciales.Vto = Text1(13).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & Text1(5).Text & "|"
    frmTESParciales.Importes = Text1(6).Text & "|" & Text1(16).Text & "|" & Text1(8).Text & "|"
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
    
    
    
    If Val(DBLet(Data1.Recordset!transfer, "N")) = 1 Then
        MsgBox "Pertenece a una transferencia", vbExclamation
        Exit Sub
    End If
    
    
    Im = Data1.Recordset!ImpVenci + DBLet(Data1.Recordset!Gastos, "N")
    Im = Im - DBLet(Data1.Recordset!impcobro, "N")
    If Im = 0 Then
        MsgBox "NO puede dividir el vencimiento. Importe totalmente cobrado", vbExclamation
        Exit Sub
    End If
    
    
       'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
    
    CadenaDesdeOtroForm = "numserie = '" & Data1.Recordset!NUmSerie & "' AND numfactu = " & Data1.Recordset!NumFactu
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND fecfactu = '" & Format(Data1.Recordset!FecFactu, FormatoFecha) & "'|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Data1.Recordset!numorden & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CStr(Im) & "|"
    
    
    'Ok, Ahora pongo los labels
    frmTESCobrosDivVto.Opcion = 27
    frmTESCobrosDivVto.Label4(56).Caption = Text2(0).Text
    frmTESCobrosDivVto.txtCodigo(2).Text = Text1(5).Text
    
    frmTESCobrosDivVto.Label4(57).Caption = Data1.Recordset!NUmSerie & Format(Data1.Recordset!NumFactu, "000000") & " / " & Data1.Recordset!numorden & "      de " & Format(Data1.Recordset!FecFactu, "dd/mm/yyyy")
    
    'Si ya ha cobrado algo...
    Im = DBLet(Data1.Recordset!impcobro, "N")
    If Im > 0 Then frmTESCobrosDivVto.txtCodigo(1).Text = txtPendiente.Text
    
    If Text1(0).Text = "" Then
        MsgBox "El cobro no tiene forma de pago. Revise.", vbExclamation
        Exit Sub
    End If
   
    
    frmTESCobrosDivVto.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        CadenaConsulta = "Select * from cobros WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1) 'CadenaConsulta
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
            'Generar Cobros
            If Me.Data1.Recordset.EOF Then Exit Sub
            If Modo <> 2 Then Exit Sub
            If vTipForpa <> "" Then
                If (Val(vTipForpa) <> vbTipoPagoRemesa) Or (Val(vTipForpa) = vbTipoPagoRemesa And Val(DBLet(Data1.Recordset!CodRem)) = 0) Then
                    If SePuedeEliminar2 < 3 Then Exit Sub
                
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
Dim cad As String

    lwCobros.ListItems.Clear
    Set Me.lwCobros.SmallIcons = frmppal.imgListComun16 'imgListComun 'ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    cad = "SELECT  hlinapu.numdiari, hlinapu.fechaent, "
    cad = cad & " hlinapu.numasien, hlinapu.ctacontr, "
    cad = cad & " hcabapu.usucreacion, hcabapu.feccreacion, tipofpago.siglas, "
    cad = cad & " hlinapu.reftalonpag, hlinapu.bancotalonpag, "
    cad = cad & " (coalesce(hlinapu.timporteh,0) - coalesce(hlinapu.timported,0)) impcobro, hlinapu.numserie, hlinapu.numfaccl, hlinapu.fecfactu, hlinapu.numorden,  hlinapu.codrem "
    cad = cad & " FROM (hlinapu INNER JOIN tipofpago ON hlinapu.tipforpa = tipofpago.tipoformapago)  INNER JOIN hcabapu ON hlinapu.numdiari = hcabapu.numdiari and hlinapu.fechaent = hcabapu.fechaent and hlinapu.numasien = hcabapu.numasien  "
    If Enlaza Then
        cad = cad & Replace(Replace(ObtenerWhereCab(True), "cobros", "hlinapu"), "numfactu", "numfaccl")
    Else
        cad = cad & " WHERE hlinapu.codmacta is null"
    End If
    cad = cad & " ORDER BY hlinapu.numserie, hlinapu.numfaccl, hlinapu.fecfactu, hlinapu.numorden, hlinapu.fechaent, hlinapu.numasien "
    
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = DBLet(miRsAux!NumDiari, "N")
        IT.SubItems(1) = Format(miRsAux!FechaEnt, "dd/mm/yyyy")
        IT.SubItems(2) = DBLet(miRsAux!NumAsien, "N")
        IT.SubItems(3) = DBLet(miRsAux!ctacontr, "T")
        IT.SubItems(4) = DBLet(miRsAux!usucreacion, "T")
        IT.SubItems(5) = Format(DBLet(miRsAux!feccreacion, "F"), "dd/mm/yyyy")
        IT.SubItems(6) = DBLet(miRsAux!siglas, "T")
        IT.SubItems(7) = DBLet(miRsAux!reftalonpag, "T")
        IT.SubItems(8) = DBLet(miRsAux!bancotalonpag, "T")
        IT.SubItems(9) = Format(miRsAux!impcobro, "###,###,##0.00")
        
        IT.SubItems(10) = DBLet(miRsAux!NUmSerie)
        IT.SubItems(11) = DBLet(miRsAux!numfaccl)
        IT.SubItems(12) = DBLet(miRsAux!FecFactu, "F")
        IT.SubItems(13) = DBLet(miRsAux!numorden)
        IT.SubItems(14) = DBLet(miRsAux!CodRem)
        
        
        If DBLet(miRsAux!CodRem, "N") <> 0 Then
            IT.SmallIcon = 42
        End If
         
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing


End Sub



Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim tots As String

    tots = MontaSQLCarga(Index, Enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez

    Select Case Index
        Case 0 ' hlinapu
            ' es un listview
        
        Case 1 'DEVOLUCIONES
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux1(5)|T|Fecha|1255|;S|txtaux1(6)|T|Código|1005|;S|cmdAux(2)|B|||;"
            tots = tots & "S|txtaux1(8)|T|Descripción|5055|;"
            tots = tots & "N||||0|;S|txtaux1(7)|T|Tipo |1500|;"
            tots = tots & "S|txtaux1(9)|T|Remesa|1000|;"
            tots = tots & "S|txtaux1(10)|T|Año|1000|;"
            tots = tots & "S|txtaux1(11)|T|Importe|2120|;"
            

            arregla tots, DataGridAux(Index), Me

            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
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
        
        Case 1 'hlinapu
            tabla = "hlinapu"
            Sql = "SELECT hlinapu.numserie, hlinapu.numfaccl, hlinapu.fecfactu, hlinapu.numorden, hlinapu.fecdevol, "
            Sql = Sql & "hlinapu.coddevol, usuarios.wdevolucion.descripcion, hlinapu.tiporem, "
            Sql = Sql & "CASE hlinapu.tiporem WHEN 1 THEN 'Efectos' WHEN 2 THEN 'Pagarés' WHEN 3 THEN 'Talones' END as TTipo, codrem, anyorem, (coalesce(timported,0) - coalesce(timporteh,0)) impcobro  "
            Sql = Sql & " FROM " & tabla & " LEFT JOIN usuarios.wdevolucion ON hlinapu.coddevol = usuarios.wdevolucion.codigo "
            If Enlaza Then
                Sql = Sql & Replace(Replace(ObtenerWhereCab(True), "cobros", "hlinapu"), "numfactu", "numfaccl")
            Else
                Sql = Sql & " WHERE hlinapu.codmacta is null"
            End If
            
            Sql = Sql & " and hlinapu.esdevolucion = 1 "
            Sql = Sql & " ORDER BY 1,2,3,4,5"
            
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & "cobros.numserie=" & DBSet(Text1(13).Text, "T") & " and cobros.numfactu=" & DBSet(Text1(1).Text, "N") & " and cobros.fecfactu = " & DBSet(Text1(2).Text, "F")
    vWhere = vWhere & " and cobros.numorden = " & DBSet(Text1(3).Text, "N")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(numserie=" & DBSet(Text1(13).Text, "T") & " and numfactu = " & DBSet(Text1(1).Text, "N") & " and fecfactu = " & DBSet(Text1(2).Text, "F") & " and numorden = " & DBSet(Text1(3).Text, "N") & ") "
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
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



Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'hlinapu
            For jj = 5 To txtaux.Count - 1
                txtaux(jj).Visible = B
                txtaux(jj).top = alto
            Next jj
        
        Case 1 'lineas de factura
            txtaux1(6).Visible = B
            txtaux1(6).top = alto
            txtaux1(8).Visible = B
            txtaux1(8).top = alto
            txtaux1(8).Locked = True
            Me.cmdAux(2).Visible = B
            Me.cmdAux(2).top = alto
            
            
            txtaux1(5).Visible = B And (Modo = 1)
            txtaux1(5).top = alto
            txtaux1(9).Visible = B And (Modo = 1)
            txtaux1(9).top = alto
            txtaux1(10).Visible = B And (Modo = 1)
            txtaux1(10).top = alto
            txtaux1(11).Visible = B And (Modo = 1)
            txtaux1(11).top = alto
            
    End Select

End Sub


Private Sub CargaFiltros()
Dim Aux As String
    

    cboFiltro.Clear
    
    cboFiltro.AddItem "Pendientes de Cobro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Cobrado "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1

    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 9


End Sub


Private Sub ToolbarAux1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LINASI As Long
Dim Ampliacion As String

    'Fuerzo que se vean las lineas
    Select Case Button.Index
        Case 1
            'Acceder a asiento del cobro
            BotonModificarLinea 1

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


Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim J As Long
    
    cboSituRem.Clear

    cboSituRem.AddItem ""
    cboSituRem.ItemData(cboSituRem.NewIndex) = Asc("NULL")


    'Tipo de situacion de remesa
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

    Combo1.Clear

    Combo1.AddItem "Pendientes de Cobro"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Cobrado"
    Combo1.ItemData(Combo1.NewIndex) = 1
    




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
        '--DesBloqueaRegistroForm Me.Text1(0)
        TerminaBloquear
        
        
        'LOG
        vLog.Insertar 6, vUsu, Text1(2).Text & Text1(0).Text & " " & Text1(1).Text
        'Creo que no hace falta volver a situar el datagrid
        'If SituarData1(0) Then
        If True Then
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
            Data1.Refresh
            PonerModo 2
        Else
            PonerModo 0
        End If
        '**** hasta aqui
        
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
        ' ***************************************
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
    
'    If numTab = 0 Then
'        SSTab1.Tab = 2
'    ElseIf numTab = 1 Then
'        SSTab1.Tab = 1
'    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Sub BotonVerAsiento()

    If lwCobros.SelectedItem Is Nothing Then Exit Sub


    Set frmAsi = New frmAsientosHco
    
    frmAsi.ASIENTO = lwCobros.SelectedItem.Text & "|" & lwCobros.SelectedItem.SubItems(1) & "|" & lwCobros.SelectedItem.SubItems(2) & "|" '& Me.AdoAux(0).Recordset.Fields(5) & "|" & Me.AdoAux(0).Recordset.Fields(6) & "|" & Me.AdoAux(0).Recordset.Fields(7) & "|"
    frmAsi.SoloImprimir = True
    frmAsi.Show vbModal
    
    Set frmAsi = Nothing

End Sub

Private Sub BotonImprimirRecibo()

    If lwCobros.SelectedItem Is Nothing Then Exit Sub


    CargarTemporal

    frmTESImpRecibo.pImporte = lwCobros.SelectedItem.SubItems(9) 'Me.AdoAux(0).Recordset.Fields(14)
    frmTESImpRecibo.pFechaRec = lwCobros.SelectedItem.SubItems(1) 'Me.AdoAux(0).Recordset.Fields(6)
    frmTESImpRecibo.pFecFactu = lwCobros.SelectedItem.SubItems(12) 'Me.AdoAux(0).Recordset.Fields(2)
    frmTESImpRecibo.pNumFactu = lwCobros.SelectedItem.SubItems(11) 'Me.AdoAux(0).Recordset.Fields(1)
    frmTESImpRecibo.pNumSerie = lwCobros.SelectedItem.SubItems(10) 'Me.AdoAux(0).Recordset.Fields(0)
    frmTESImpRecibo.pNumOrden = lwCobros.SelectedItem.SubItems(13) 'Me.AdoAux(0).Recordset.Fields(3)
    frmTESImpRecibo.pNumlinea = lwCobros.SelectedItem.SubItems(14) 'Me.AdoAux(0).Recordset.Fields(4)
    
    frmTESImpRecibo.Show vbModal


End Sub

Private Sub CargarTemporal()
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "delete from tmppendientes where codusu = " & vUsu.Codigo
    Conn.Execute Sql
                                                                              
    ' en tmppendientes metemos la clave primaria de cobros_recibidos y el importe en letra
                                                      'importe=nro factura,   codforpa=linea de cobros_realizados
    Sql = "insert into tmppendientes (codusu,serie_cta,importe,fecha,numorden,codforpa, observa) values ("
    Sql = Sql & vUsu.Codigo & "," & DBSet(lwCobros.SelectedItem.SubItems(10), "T") & "," 'numserie
    Sql = Sql & DBSet(lwCobros.SelectedItem.SubItems(11), "N") & "," 'numfactu
    Sql = Sql & DBSet(lwCobros.SelectedItem.SubItems(12), "F") & "," 'fecfactu
    Sql = Sql & DBSet(lwCobros.SelectedItem.SubItems(13), "N") & "," 'numorden
    Sql = Sql & DBSet(lwCobros.SelectedItem.SubItems(14), "N") & "," 'numlinea
    Sql = Sql & DBSet(EscribeImporteLetra(ImporteFormateado(CStr(lwCobros.SelectedItem.SubItems(9)))), "T") & ") "
    
    Conn.Execute Sql

End Sub


Private Sub BotonAnyadirLinea(Index As Integer, Limpia As Boolean)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer

    ModoLineas = 1 'Posem Modo Afegir Llínia

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    ' *** bloquejar la clau primaria de la capçalera ***
'    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vTabla = "cobros"
        
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 0   'cobros_realizados
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = ""
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", Replace(vWhere, "cobros", "cobros"))
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)

            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 230 '248
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'lineas de cobros realizados
                    If Limpia Then
                        For i = 0 To txtaux.Count - 1
                            txtaux(i).Text = ""
                        Next i
                    End If
                    txtaux(0).Text = Text1(13).Text 'serie
                    txtaux(1).Text = Text1(1).Text 'numfactu
                    txtaux(2).Text = Text1(2).Text 'fecha
                    txtaux(3).Text = Text1(3).Text 'nro vencimiento
                    
                    txtaux(4).Text = Format(NumF, "0000") 'linea contador
                    
                    PonFoco txtaux(5)
            
            End Select

    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub


    ModoLineas = 2 'Modificar llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    Select Case Index
        Case 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If

            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

    End Select

    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 'lineas de facturas
            txtaux1(0).Text = DataGridAux(Index).Columns(0).Text 'serie
            txtaux1(1).Text = DataGridAux(Index).Columns(1).Text 'factura
            txtaux1(2).Text = DataGridAux(Index).Columns(2).Text 'fecha
            txtaux1(3).Text = DataGridAux(Index).Columns(3).Text 'vencimiento
            txtaux1(4).Text = DataGridAux(Index).Columns(4).Text 'linea
            
            txtaux1(5).Text = DataGridAux(Index).Columns(5).Text ' concepto de devolucion
            txtaux1(6).Text = DataGridAux(Index).Columns(6).Text ' concepto
            txtaux1(8).Text = DataGridAux(Index).Columns(7).Text ' nombre del concepto
            
            txtaux1(9).Text = DataGridAux(Index).Columns(10).Text ' numero remesa
            txtaux1(10).Text = DataGridAux(Index).Columns(11).Text ' año
            txtaux1(11).Text = DataGridAux(Index).Columns(12).Text ' importe
            txtaux1(12).Text = DataGridAux(Index).Columns(8).Text ' tipo de remesa
            
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    
    PonFoco txtaux1(6)
    
    ' ***************************************************************************************
End Sub

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    SepuedeBorrar = True
End Function


Private Function RecalcularTotalesCobros() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim SqlValues As String
Dim i As Long
Dim Rs As ADODB.Recordset

Dim Baseimpo As Currency

    On Error GoTo eRecalcularTotalesCobros

    RecalcularTotalesCobros = False

    Baseimpo = 0
    
    Sql = "select sum(coalesce(timporteh,0) - coalesce(timported,0)) importe from hlinapu "
    Sql = Sql & " where numserie = " & DBSet(Text1(13).Text, "T") & " and numfaccl = " & DBSet(Text1(1).Text, "N") & " and fecfactu = " & DBSet(Text1(2).Text, "F")
    Sql = Sql & " and numorden = " & DBSet(Text1(3).Text, "N")
    
    Baseimpo = DevuelveValor(Sql)
    
    Text1(6).Text = Format(Baseimpo, FormatoImporte)
    
    Sql = "update cobros set "
    Sql = Sql & " impcobro = " & DBSet(Baseimpo, "N")
    Sql = Sql & " where numserie= " & DBSet(Text1(13).Text, "T") & " and numfactu= " & DBSet(Text1(1).Text, "N") & " and fecfactu = " & DBSet(Text1(2).Text, "F")
    Sql = Sql & " and numorden = " & DBSet(Text1(3).Text, "N")
    
    Conn.Execute Sql
    
    RecalcularTotalesCobros = True
    Exit Function
    
eRecalcularTotalesCobros:
    MuestraError Err.Number, "Recalcular Totales Cobros", Err.Description
End Function

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean
Dim Limp As Boolean
Dim cad As String



    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0"
        Case 1: nomframe = "FrameAux1"
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        Conn.BeginTrans
        
        B = True
    
        If B And InsertarDesdeForm2(Me, 2, nomframe) Then
        
            B = RecalcularTotalesCobros
            
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    
                    DataGridAux(NumTabMto).AllowAddNew = False
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then PosicionGrid = DataGridAux(NumTabMto).FirstRow
                    CargaGrid NumTabMto, True
                    Limp = True

                    If Limp Then
                        For i = 0 To txtaux.Count - 1
                            txtaux(i).Text = ""
                        Next i
                    End If
                    ModoLineas = 0
                    If B Then
                         BotonAnyadirLinea NumTabMto, True
                    End If
            End Select
           
        Else
           Conn.RollbackTrans
        End If
    End If
End Sub

Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim v As Integer
Dim cad As String
Dim B As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'apuntes
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        Conn.BeginTrans
        
        B = True
        
        If B And ModificaDesdeFormulario2(Me, 2, nomframe) Then
        
            B = RecalcularTotalesCobros
        
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            ' ******************************************************
            ModoLineas = 0

            If NumTabMto <> 3 Then
                v = AdoAux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(3).Name & " =" & v)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
            
        Else
            Conn.RollbackTrans
        End If
    End If
        
End Sub


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
    ConseguirFoco txtaux(Index), Modo
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
        
        
    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
    
    If txtaux(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 5 ' diario
            RC = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtaux(5), "N")
            If RC = "" Then
                MsgBox "No existe el tipo de diario. Reintroduzca.", vbExclamation
                PonFoco txtaux(5)
            End If
                
        Case 6, 11 ' fecha
            If Not EsFechaOK(txtaux(Index)) Then
                MsgBox "Fecha incorrecta: " & txtaux(Index).Text, vbExclamation
                txtaux(Index).Text = ""
                PonerFoco txtaux(Index)
            End If
            
        Case 7 ' asiento
            PonerFormatoEntero txtaux(Index)
        
        Case 8 ' usuario
        
        Case 9
           ' IMPORTE
             txtaux(Index) = ImporteSinFormato(txtaux(Index))
            
        Case 10 'tipo
            txtaux(Index).Text = UCase(txtaux(Index).Text)
        
        Case 12 ' cuenta de cobro
            RC = txtaux(12).Text
            If CuentaCorrectaUltimoNivel(RC, "") Then
                txtaux(12).Text = RC
            End If
        
    End Select
End Sub



Private Sub CargarDatosCuenta(Cuenta As String, SoloDatosFiscales As Boolean)
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo eTraerDatosCuenta
    
    Sql = "select * from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not SoloDatosFiscales Then
        Text1(0).Text = ""
        Text2(1).Text = ""
    End If
    For i = 36 To 42
        Text1(i).Text = ""
    Next i
    
    If Not Rs.EOF Then
        If Not SoloDatosFiscales Then
            Text1(0).Text = DBLet(Rs!Forpa, "N")
            Text2(1).Text = PonerNombreDeCod(Text1(0), "formapago", "nomforpa", "codforpa", "N")
        End If
        Text1(42).Text = DBLet(Rs!Nommacta, "T")
        Text1(41).Text = DBLet(Rs!dirdatos, "T")
        Text1(40).Text = DBLet(Rs!codposta, "T")
        Text1(39).Text = DBLet(Rs!desPobla, "T")
        Text1(38).Text = DBLet(Rs!desProvi, "T")
        Text1(37).Text = DBLet(Rs!nifdatos, "T")
        Text1(36).Text = DBLet(Rs!codPAIS, "T")
        Text2(36).Text = PonerNombreDeCod(Text1(36), "paises", "nompais", "codpais", "T")
    End If
    Rs.Close
    
eTraerDatosCuenta:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar Datos de Cuenta", Err.Description
    Set Rs = Nothing
End Sub

Private Sub CargarColumnas()
    
    lwCobros.ColumnHeaders.Clear
    
    lwCobros.ColumnHeaders.Add , , "Diario", 830
    lwCobros.ColumnHeaders.Add , , "Fecha", 1320
    lwCobros.ColumnHeaders.Add , , "Asiento", 1200, 1
    lwCobros.ColumnHeaders.Add , , "Cta.Cobro", 1405
    lwCobros.ColumnHeaders.Add , , "Usuario", 1405
    lwCobros.ColumnHeaders.Add , , "Realizado", 1455
    lwCobros.ColumnHeaders.Add , , "Tipo", 905
    lwCobros.ColumnHeaders.Add , , "Ref.Talon", 1755
    lwCobros.ColumnHeaders.Add , , "Banco Talon/Pag", 1855
    lwCobros.ColumnHeaders.Add , , "Importe", 1900, 1
    
    lwCobros.ColumnHeaders.Add , , "serie", 0
    lwCobros.ColumnHeaders.Add , , "Factura", 0
    lwCobros.ColumnHeaders.Add , , "Fecha", 0
    lwCobros.ColumnHeaders.Add , , "vto", 0
    lwCobros.ColumnHeaders.Add , , "Linea", 0
    lwCobros.ColumnHeaders.Add , , "Remesa", 0

End Sub

Private Sub txtaux1_GotFocus(Index As Integer)
    ConseguirFoco txtaux1(Index), Modo
End Sub


Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub txtaux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 6:  KEYImage KeyAscii, 2 ' concepto de devolucion
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

'++


Private Sub txtAux1_LostFocus(Index As Integer)
    Dim RC As String
    Dim Importe As Currency
        
    If Not PerderFocoGnral(txtaux1(Index), Modo) Then Exit Sub
    
    If txtaux1(Index).Text = "" Then Exit Sub
    
    Select Case Index
            
        Case 6 ' iva
            txtaux1(6).Text = UCase(txtaux1(6).Text)
            txtaux1(8).Text = DevuelveDesdeBD("descripcion", "usuarios.wdevolucion", "codigo", txtaux1(6), "T")
            If txtaux1(8).Text = "" Then
                MsgBox "No existe el Tipo de Devolución. Reintroduzca.", vbExclamation
                PonFoco txtaux1(6)
            End If
                
    End Select
End Sub

