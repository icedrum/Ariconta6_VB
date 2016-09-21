VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmempresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresa"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9150
   Icon            =   "frmempresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame Frame1 
      Height          =   420
      Left            =   150
      TabIndex        =   75
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
         TabIndex        =   76
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   71
      Top             =   90
      Width           =   1125
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   72
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   73
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
      TabIndex        =   35
      Top             =   1020
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   8070
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
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmempresa.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(10)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(11)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(13)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(14)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(15)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(17)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(8)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(10)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(12)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(13)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(14)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(15)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(16)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(17)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Otros datos"
      TabPicture(1)   =   "frmempresa.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(18)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(19)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(21)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(22)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(23)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(24)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(25)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(26)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(28)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(30)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(20)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text1(18)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text1(6)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text1(5)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text1(4)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text1(3)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text1(19)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text1(21)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text1(22)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text1(23)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text1(24)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Text1(25)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Text1(26)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Text1(20)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Text1(31)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Text1(32)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text1(33)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "Presentación IVA"
      TabPicture(2)   =   "frmempresa.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(27)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label4(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(29)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text1(27)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text1(28)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text1(29)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text1(30)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
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
         Left            =   -72300
         MaxLength       =   40
         TabIndex        =   32
         Tag             =   "IBAN Ingreso|T|S|||empresa2|iban2|||"
         Text            =   "Text1"
         Top             =   3210
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
         Index           =   29
         Left            =   -72300
         MaxLength       =   40
         TabIndex        =   31
         Tag             =   "IBAN Devolución|T|S|||empresa2|iban1|||"
         Text            =   "Text1"
         Top             =   2670
         Width           =   4485
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
         Left            =   7140
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
         Left            =   7140
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
         Left            =   7140
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
         Left            =   7140
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
         Left            =   7140
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
         Left            =   4800
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
         Left            =   4800
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
         Left            =   4800
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
         Left            =   4800
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
         Left            =   4800
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
         Left            =   1800
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
         Left            =   6195
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   1710
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
         Left            =   1695
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   4365
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
         Left            =   360
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
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
         Left            =   -68640
         TabIndex        =   54
         Tag             =   "NIF|T|S|||empresa2|codigo||S|"
         Text            =   "CODIGO"
         Top             =   600
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
         Left            =   -69150
         MaxLength       =   4
         TabIndex        =   30
         Tag             =   "Letras|T|S|||empresa2|letraseti|||"
         Text            =   "Text1"
         Top             =   1500
         Width           =   855
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
         TabIndex        =   17
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
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   29
         Tag             =   "Admon|T|S|||empresa2|administracion|||"
         Text            =   "Text1"
         Top             =   1500
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
         TabIndex        =   19
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   20
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
         TabIndex        =   18
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
         TabIndex        =   21
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   28
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
         TabIndex        =   16
         Tag             =   "Apoderado|T|S|||empresa2|apoderado|||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   6420
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
         TabIndex        =   51
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
         TabIndex        =   70
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
         Left            =   2535
         TabIndex        =   69
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
         Left            =   6150
         TabIndex        =   68
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
         Left            =   6180
         TabIndex        =   67
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
         Left            =   6180
         TabIndex        =   66
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
         Left            =   6150
         TabIndex        =   65
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
         Left            =   6150
         TabIndex        =   64
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
         Left            =   3750
         TabIndex        =   63
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
         Left            =   3750
         TabIndex        =   62
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
         Left            =   3750
         TabIndex        =   61
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
         Left            =   3750
         TabIndex        =   60
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
         Left            =   3750
         TabIndex        =   59
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
         Left            =   405
         TabIndex        =   58
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
         Left            =   6195
         TabIndex        =   57
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
         Left            =   1695
         TabIndex        =   56
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
         Left            =   375
         TabIndex        =   55
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Letras etiqueta"
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
         Index           =   29
         Left            =   -70770
         TabIndex        =   53
         Top             =   1560
         Width           =   1590
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
         TabIndex        =   52
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
         Left            =   -74160
         TabIndex        =   50
         Top             =   3240
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
         Left            =   -74160
         TabIndex        =   49
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Administración"
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
         Left            =   -74190
         TabIndex        =   48
         Top             =   1560
         Width           =   2160
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   2925
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
         Height          =   255
         Index           =   3
         Left            =   -70185
         TabIndex        =   39
         Top             =   3735
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
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   1320
         Width           =   1575
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
      TabIndex        =   34
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
      TabIndex        =   33
      Top             =   5850
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   8520
      TabIndex        =   74
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

Private Sub cmdAceptar_Click()
    Dim cad As String
    Dim i As Integer
    
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
                If Adodc1.Recordset.EOF Then
                    i = InsertarDesdeForm(Me)
                Else
                    i = ModificaDesdeFormulario(Me)
                End If
                If i = -1 Then PonerModo 0
            End If

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
End Select
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

    Me.Top = 200
    Me.Left = 400
    Limpiar Me
    'Lista imagen
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(2).Image = 4
    End With
    
    Text1(0).Enabled = False
    Me.SSTab1.Tab = 0
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    
    
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = "Select * from empresa2"
    Adodc1.Refresh
    
    If vEmpresa Is Nothing Then
        'No hay datos
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
    If Adodc1.Recordset.EOF Then Text1(38).Text = "1"  'Codigo para la tabla 2 de empresa
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
    Dim i As Integer
    Modo = Kmodo
    
    For i = 1 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i
    
    Select Case Kmodo
    Case 0
        'Preparamos para ver los datos
        For i = 1 To Text1.Count - 1
            Text1(i).Locked = True
        Next i
        lblIndicador.Caption = ""
    Case 3
        'Preparamos para que pueda insertar
        For i = 1 To Text1.Count - 1
            Text1(i).Text = ""
            Text1(i).Locked = False
        Next i
        lblIndicador.Caption = "INSERTAR"
    Case 4
        For i = 1 To Text1.Count - 1
            Text1(i).Locked = False
        Next i
        lblIndicador.Caption = "MODIFICAR"
    End Select
    Me.Toolbar1.Buttons(1).Enabled = Modo <> 3 '1
    cmdAceptar.Visible = Modo > 0
    cmdCancelar.Visible = Modo > 0

    PonerModoUsuarioGnral Modo, "ariconta"


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
            
    Else
        Limpiar Me
    End If
    If Adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Adodc1
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
    Dim i As Integer
    Dim J As Integer
    Dim k As Integer
    
    DatosOK = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Otras cosas importantes
    'Comprobamos que tienen n niveles, y solo n
    J = CInt(Text1(7).Text)
    k = 0
    For i = 8 To 17
        If Text1(i).Text <> "" Then k = k + 1
    Next i
    If k <> J Then
        MsgBox "Niveles contables: " & J & vbCrLf & "Niveles parametrizados: " & k, vbExclamation
        Exit Function
    End If
    
    'K los niveles sean consecitivos sin saltar ninguno y sin ser menor
    J = 1
    k = CInt(Text1(8).Text)
    For i = 9 To 17
        If Text1(i).Text = "" Then
            J = 1000 + i
        Else
            J = CInt(Text1(i).Text)
        End If
        If J <= k Then
            MsgBox "Error en la asignacion de niveles. ", vbExclamation
            Exit Function
        End If
        k = J
    Next i
    
    
    
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

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
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

