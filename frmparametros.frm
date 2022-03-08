VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmparametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Contables"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11880
   Icon            =   "frmparametros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Height          =   465
      Left            =   120
      TabIndex        =   107
      Top             =   7080
      Width           =   2565
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label3"
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
         Left            =   120
         TabIndex        =   108
         Top             =   120
         Width           =   2310
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   103
      Top             =   0
      Width           =   1185
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   104
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   105
         Top             =   180
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
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
      Left            =   10650
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   7230
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6165
      Left            =   90
      TabIndex        =   50
      Top             =   840
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   10874
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "frmparametros.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgFec(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgFec(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgFec(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(27)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(33)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgPathFich(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(8)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(12)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(16)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(17)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(31)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(32)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Clientes - Proveedores "
      TabPicture(1)   =   "frmparametros.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "I.V.A. - Norma 43"
      TabPicture(2)   =   "frmparametros.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(13)"
      Tab(2).Control(1)=   "Text1(15)"
      Tab(2).Control(2)=   "Text1(14)"
      Tab(2).Control(3)=   "Frame17"
      Tab(2).Control(4)=   "Frame8"
      Tab(2).Control(5)=   "Frame1"
      Tab(2).Control(6)=   "Label1(13)"
      Tab(2).Control(7)=   "Label1(14)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Inmovilizado"
      TabPicture(3)   =   "frmparametros.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).Control(2)=   "Frame16"
      Tab(3).Control(3)=   "Frame15"
      Tab(3).Control(4)=   "Frame14"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Tesorería I"
      TabPicture(4)   =   "frmparametros.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame66"
      Tab(4).Control(1)=   "FrameValDefecto"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Tesorería II"
      TabPicture(5)   =   "frmparametros.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameTalones(1)"
      Tab(5).Control(1)=   "FrameTalones(0)"
      Tab(5).Control(2)=   "FrameOpAseguradas"
      Tab(5).Control(3)=   "FrameTalones(2)"
      Tab(5).Control(4)=   "FrameTalones(3)"
      Tab(5).ControlCount=   5
      Begin VB.Frame Frame11 
         Caption         =   "Integración histórico apuntes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   210
         Top             =   4560
         Width           =   11175
         Begin VB.CheckBox Check1 
            Caption         =   "Abonos negativos"
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
            Index           =   7
            Left            =   240
            TabIndex        =   22
            Tag             =   "Abonos negativos|N|N|||parametros|abononeg|||"
            Top             =   480
            Width           =   2145
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Añadir nº fra proveedor en ampliacion"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   22
            Left            =   240
            TabIndex        =   24
            Tag             =   "Prov. numfra hlinapu|N|N|||parametros|anyadenumfacproContab|||"
            ToolTipText     =   "Añadir numero factura en la ampliacion del apunte"
            Top             =   840
            Width           =   4755
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Observaciones lineas factura proveedor"
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
            Index           =   23
            Left            =   5160
            TabIndex        =   25
            Tag             =   "Observalinfac|N|N|||parametros|obsrlinfact|||"
            Top             =   960
            Width           =   4905
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
            ItemData        =   "frmparametros.frx":00B4
            Left            =   7080
            List            =   "frmparametros.frx":00C1
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Tag             =   "Ampliación clientes|T|S|||parametros|nctafact|||"
            Top             =   427
            Width           =   3975
         End
         Begin VB.Label Label1 
            Caption         =   "Ampliación clientes / Proveedores"
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
            Left            =   3480
            TabIndex        =   211
            Top             =   480
            Width           =   3795
         End
      End
      Begin VB.Frame FrameTalones 
         Caption         =   "Confirming clientes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Index           =   3
         Left            =   -74880
         TabIndex        =   197
         Top             =   4920
         Width           =   6885
         Begin VB.CheckBox Check5 
            Caption         =   "Contabiliza contra cuentas puente"
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
            Left            =   2610
            TabIndex        =   199
            Top             =   210
            Width           =   3825
         End
         Begin VB.TextBox Text5 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   174
            Text            =   "0"
            Top             =   570
            Width           =   1515
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
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
            Index           =   1
            Left            =   1950
            TabIndex        =   198
            Text            =   "Text4"
            Top             =   570
            Width           =   4785
         End
         Begin VB.Image imgCta2 
            Height          =   240
            Index           =   1
            Left            =   1650
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Cancelacion cliente"
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
            Left            =   120
            TabIndex        =   200
            Top             =   270
            Width           =   1995
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
         Index           =   32
         Left            =   2760
         MaxLength       =   250
         TabIndex        =   195
         Tag             =   "Path|T|S|||parametros|PathFicherosInteg|||"
         Top             =   5400
         Width           =   8475
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
         Left            =   -65070
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "Importe limite 347|N|S|0||parametros|limimpcl|0.00||"
         Text            =   "3"
         Top             =   2010
         Width           =   1260
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
         Left            =   -64170
         MaxLength       =   2
         TabIndex        =   35
         Tag             =   "Ultimo periodo liquidación I.V.A.|N|S|0|100|parametros|perfactu|||"
         Text            =   "2"
         Top             =   1590
         Width           =   360
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
         Left            =   -65070
         MaxLength       =   8
         TabIndex        =   34
         Tag             =   "Ultimo año liquidación I.V.A.|N|S|0|9999|parametros|anofactu|||"
         Text            =   "1999"
         Top             =   1590
         Width           =   660
      End
      Begin VB.Frame Frame17 
         Caption         =   "Contadores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -67260
         TabIndex        =   185
         Top             =   900
         Width           =   3525
         Begin VB.OptionButton Option1 
            Caption         =   "Año natural"
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
            Index           =   1
            Left            =   1860
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Año fiscal"
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
            Index           =   0
            Left            =   180
            TabIndex        =   32
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.Frame FrameTalones 
         Caption         =   "Efectos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1815
         Index           =   2
         Left            =   -74820
         TabIndex        =   162
         Top             =   690
         Width           =   6885
         Begin VB.CheckBox Check5 
            Caption         =   "Contabiliza contra cuentas de efectos"
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
            Left            =   2670
            TabIndex        =   164
            Top             =   300
            Width           =   4125
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
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
            Index           =   9
            Left            =   1920
            TabIndex        =   168
            Text            =   "Text4"
            Top             =   600
            Width           =   4845
         End
         Begin VB.TextBox Text5 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   165
            Text            =   "0"
            Top             =   600
            Width           =   1485
         End
         Begin VB.TextBox Text5 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   166
            Text            =   "0"
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
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
            Index           =   10
            Left            =   1890
            TabIndex        =   163
            Text            =   "Text4"
            Top             =   1320
            Width           =   4875
         End
         Begin VB.Label Label4 
            Caption         =   "Efectos descontados"
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
            Left            =   120
            TabIndex        =   172
            Top             =   360
            Width           =   2085
         End
         Begin VB.Image imgCta2 
            Height          =   240
            Index           =   9
            Left            =   1620
            Top             =   630
            Width           =   240
         End
         Begin VB.Image imgCta2 
            Height          =   240
            Index           =   10
            Left            =   1620
            Stretch         =   -1  'True
            Top             =   1350
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Efectos descontados a cobrar"
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
            TabIndex        =   170
            Top             =   1080
            Width           =   3495
         End
      End
      Begin VB.Frame FrameOpAseguradas 
         Caption         =   "Operaciones aseguradas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3105
         Left            =   -67830
         TabIndex        =   154
         Top             =   690
         Width           =   4275
         Begin VB.TextBox Text5 
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   175
            Top             =   720
            Width           =   465
         End
         Begin VB.TextBox Text5 
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
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   176
            Top             =   720
            Width           =   465
         End
         Begin VB.TextBox Text5 
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   177
            Top             =   1440
            Width           =   465
         End
         Begin VB.TextBox Text5 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   178
            Top             =   1440
            Width           =   465
         End
         Begin VB.TextBox Text5 
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   179
            Top             =   2220
            Width           =   465
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Fecha factura para dias asegurados"
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
            Index           =   11
            Left            =   240
            TabIndex        =   180
            Top             =   2550
            Width           =   3945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Aviso falta de pago"
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
            Index           =   17
            Left            =   240
            TabIndex        =   161
            Top             =   360
            Width           =   1920
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Dias aviso siniestro"
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
            Left            =   270
            TabIndex        =   160
            Top             =   1140
            Width           =   1890
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
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
            Left            =   690
            TabIndex        =   159
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
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
            Left            =   1800
            TabIndex        =   158
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
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
            Left            =   720
            TabIndex        =   157
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
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
            Left            =   1800
            TabIndex        =   156
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Aviso siniestro desde prorroga"
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
            Index           =   15
            Left            =   240
            TabIndex        =   155
            Top             =   1920
            Width           =   2985
         End
      End
      Begin VB.Frame FrameTalones 
         Caption         =   "Talones cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Index           =   0
         Left            =   -74820
         TabIndex        =   151
         Top             =   2520
         Width           =   6885
         Begin VB.TextBox Text5 
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
            Left            =   90
            MaxLength       =   10
            TabIndex        =   169
            Text            =   "0"
            Top             =   540
            Width           =   1515
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
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
            Index           =   7
            Left            =   1920
            TabIndex        =   152
            Text            =   "Text4"
            Top             =   540
            Width           =   4815
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Contabiliza contra cuentas puente  "
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
            Left            =   2610
            TabIndex        =   167
            Top             =   240
            Width           =   3825
         End
         Begin VB.Label Label4 
            Caption         =   "Cancelacion cliente"
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
            Left            =   120
            TabIndex        =   153
            Top             =   270
            Width           =   1995
         End
         Begin VB.Image imgCta2 
            Height          =   240
            Index           =   7
            Left            =   1620
            Top             =   570
            Width           =   240
         End
      End
      Begin VB.Frame FrameTalones 
         Caption         =   "Pagarés clientes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Index           =   1
         Left            =   -74820
         TabIndex        =   148
         Top             =   3720
         Width           =   6885
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
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
            Index           =   8
            Left            =   1950
            TabIndex        =   149
            Text            =   "Text4"
            Top             =   570
            Width           =   4785
         End
         Begin VB.TextBox Text5 
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
            MaxLength       =   10
            TabIndex        =   173
            Text            =   "0"
            Top             =   570
            Width           =   1515
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Contabiliza contra cuentas puente"
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
            Index           =   4
            Left            =   2610
            TabIndex        =   171
            Top             =   210
            Width           =   3825
         End
         Begin VB.Label Label4 
            Caption         =   "Cancelacion cliente"
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
            Index           =   6
            Left            =   120
            TabIndex        =   150
            Top             =   270
            Width           =   1995
         End
         Begin VB.Image imgCta2 
            Height          =   240
            Index           =   8
            Left            =   1650
            Top             =   600
            Width           =   240
         End
      End
      Begin VB.Frame FrameValDefecto 
         Caption         =   "Valores por defecto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2835
         Left            =   -68070
         TabIndex        =   147
         Top             =   720
         Width           =   4515
         Begin VB.CheckBox Check5 
            Caption         =   "Un Asiento por recibo"
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
            Left            =   120
            TabIndex        =   130
            Top             =   330
            Width           =   2865
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Agrupar apunte banco"
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
            Index           =   1
            Left            =   120
            TabIndex        =   131
            Top             =   690
            Width           =   2805
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Abonos cambiados"
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
            Index           =   2
            Left            =   120
            TabIndex        =   132
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Comprobar riesgo al inicio"
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
            Index           =   8
            Left            =   120
            TabIndex        =   133
            Top             =   1470
            Width           =   2865
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Norma 19 por fecha vto"
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
            Left            =   120
            TabIndex        =   134
            Top             =   1860
            Width           =   2805
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Eliminar en recepcion de documentos al eliminar riesgo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   9
            Left            =   120
            TabIndex        =   135
            Top             =   2130
            Width           =   4215
         End
      End
      Begin VB.Frame Frame66 
         Height          =   2175
         Left            =   -74910
         TabIndex        =   140
         Top             =   750
         Width           =   6795
         Begin VB.TextBox Text5 
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
            Left            =   5490
            MaxLength       =   10
            TabIndex        =   129
            Top             =   1740
            Width           =   1125
         End
         Begin VB.TextBox Text5 
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
            Left            =   120
            TabIndex        =   128
            Top             =   1740
            Width           =   4755
         End
         Begin VB.TextBox Text5 
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
            MaxLength       =   10
            TabIndex        =   127
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
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
            Index           =   4
            Left            =   1920
            TabIndex        =   142
            Text            =   "Text4"
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H80000018&
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
            Index           =   0
            Left            =   1920
            TabIndex        =   141
            Text            =   "Text4"
            Top             =   480
            Width           =   4695
         End
         Begin VB.TextBox Text5 
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
            MaxLength       =   10
            TabIndex        =   126
            Top             =   480
            Width           =   1515
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "% Interes tarjeta"
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
            Index           =   16
            Left            =   4830
            TabIndex        =   146
            Top             =   1500
            Width           =   1725
         End
         Begin VB.Image ImageAyudaImpcta 
            Height          =   240
            Index           =   12
            Left            =   5160
            Top             =   1770
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Responsable"
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
            Left            =   120
            TabIndex        =   145
            Top             =   1500
            Width           =   1425
         End
         Begin VB.Label Label4 
            Caption         =   "Partidas pendientes aplicacion"
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
            Left            =   120
            TabIndex        =   144
            Top             =   840
            Width           =   2175
         End
         Begin VB.Image imgCta2 
            Height          =   240
            Index           =   4
            Left            =   1650
            Top             =   1110
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta beneficios bancarios"
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
            Left            =   120
            TabIndex        =   143
            Top             =   240
            Width           =   2925
         End
         Begin VB.Image imgCta2 
            Height          =   240
            Index           =   0
            Left            =   1650
            Top             =   510
            Width           =   240
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Datos Generales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   -74700
         TabIndex        =   119
         Top             =   810
         Width           =   3945
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
            Left            =   240
            MaxLength       =   40
            TabIndex        =   201
            Tag             =   "T|T|S|0||parametros|TextoInmoSubencionado|||"
            Top             =   2160
            Width           =   3240
         End
         Begin VB.TextBox Text2 
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
            Left            =   240
            TabIndex        =   121
            Text            =   "Text2"
            Top             =   1440
            Width           =   1435
         End
         Begin VB.ComboBox Combo3 
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
            ItemData        =   "frmparametros.frx":00EA
            Left            =   240
            List            =   "frmparametros.frx":00FA
            Style           =   2  'Dropdown List
            TabIndex        =   120
            Top             =   600
            Width           =   2355
         End
         Begin VB.Label Label2 
            Caption         =   "Etiqueta tipo inmovilizado"
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
            Left            =   240
            TabIndex        =   202
            Top             =   1920
            Width           =   2565
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   1590
            Picture         =   "frmparametros.frx":0125
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Ultima fecha amorti."
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
            TabIndex        =   138
            Top             =   1200
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de amortización"
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
            Left            =   270
            TabIndex        =   136
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Datos Contables"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   -70470
         TabIndex        =   112
         Top             =   840
         Width           =   6765
         Begin VB.TextBox Text2 
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
            Left            =   2100
            TabIndex        =   124
            Text            =   "Text2"
            Top             =   1710
            Width           =   735
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000018&
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
            Index           =   2
            Left            =   2910
            TabIndex        =   115
            Text            =   "Text3"
            Top             =   1710
            Width           =   3705
         End
         Begin VB.TextBox Text2 
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
            Left            =   2100
            TabIndex        =   123
            Text            =   "Text2"
            Top             =   1230
            Width           =   735
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000018&
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
            Index           =   1
            Left            =   2910
            TabIndex        =   114
            Text            =   "Text3"
            Top             =   1230
            Width           =   3705
         End
         Begin VB.TextBox Text2 
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
            Left            =   2100
            TabIndex        =   122
            Text            =   "Text2"
            Top             =   750
            Width           =   735
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000018&
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
            Index           =   0
            Left            =   2910
            TabIndex        =   113
            Text            =   "Text3"
            Top             =   750
            Width           =   3705
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Contabilizacion automática"
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
            Left            =   180
            TabIndex        =   137
            Top             =   330
            Width           =   2985
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   1
            Left            =   1800
            Top             =   1725
            Width           =   240
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   0
            Left            =   1800
            Top             =   1245
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   1800
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Concepto haber"
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
            Index           =   2
            Left            =   210
            TabIndex        =   118
            Top             =   1725
            Width           =   1590
         End
         Begin VB.Label Label3 
            Caption         =   "Concepto debe"
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
            Index           =   1
            Left            =   210
            TabIndex        =   117
            Top             =   1245
            Width           =   1545
         End
         Begin VB.Label Label3 
            Caption         =   "Nº Diario"
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
            Left            =   210
            TabIndex        =   116
            Top             =   765
            Width           =   1260
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Venta Inmovilizado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   -70440
         TabIndex        =   109
         Top             =   3240
         Width           =   6735
         Begin VB.TextBox txtIVA 
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
            Left            =   2010
            MaxLength       =   8
            TabIndex        =   125
            Top             =   390
            Width           =   765
         End
         Begin VB.TextBox txtIVA 
            BackColor       =   &H80000018&
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
            Index           =   1
            Left            =   2850
            TabIndex        =   110
            Top             =   390
            Width           =   3705
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Papel preimpreso"
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
            Left            =   210
            TabIndex        =   139
            Top             =   900
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.Image imgiva 
            Height          =   240
            Left            =   1740
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label14 
            Caption         =   "IVA"
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
            Index           =   2
            Left            =   210
            TabIndex        =   111
            Top             =   420
            Width           =   1245
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
         Index           =   31
         Left            =   4830
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha fin|F|S|||parametros|fechaActiva|dd/mm/yyyy||"
         Text            =   "1/2/3"
         Top             =   1110
         Width           =   1305
      End
      Begin VB.Frame Frame8 
         Caption         =   "Parámetros importación datos bancarios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74760
         TabIndex        =   92
         Top             =   900
         Width           =   7365
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   24
            Left            =   2520
            TabIndex        =   94
            Text            =   "Text2"
            Top             =   900
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
            Index           =   24
            Left            =   1860
            MaxLength       =   8
            TabIndex        =   27
            Tag             =   "Diario norma 43|N|N|0|100|parametros|diario43|000||"
            Text            =   "1"
            Top             =   900
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   23
            Left            =   2520
            TabIndex        =   93
            Text            =   "Text2"
            Top             =   420
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
            Index           =   23
            Left            =   1860
            MaxLength       =   8
            TabIndex        =   26
            Tag             =   "Concepto norma 43|N|N|0||parametros|conce43|||"
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
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
            Index           =   25
            Left            =   150
            TabIndex        =   96
            Top             =   960
            Width           =   615
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   2
            Left            =   1590
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   4
            Left            =   1590
            Top             =   480
            Width           =   240
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
            Height          =   255
            Index           =   24
            Left            =   150
            TabIndex        =   95
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Soporte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   -74700
         TabIndex        =   88
         Top             =   3360
         Visible         =   0   'False
         Width           =   10545
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
            Left            =   2340
            MaxLength       =   100
            TabIndex        =   47
            Tag             =   "Web|T|S|||parametros|webversion|||"
            Text            =   "3"
            Top             =   1320
            Width           =   6060
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
            Left            =   2340
            MaxLength       =   100
            TabIndex        =   46
            Tag             =   "M|T|S|||parametros|mailsoporte|||"
            Text            =   "3"
            Top             =   780
            Width           =   6060
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
            Left            =   2340
            MaxLength       =   100
            TabIndex        =   45
            Tag             =   "W|T|S|||parametros|websoporte|||"
            Text            =   "3"
            Top             =   300
            Width           =   6060
         End
         Begin VB.Label Label1 
            Caption         =   "Web check version"
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
            Left            =   300
            TabIndex        =   91
            Top             =   1380
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "Mail soporte"
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
            Left            =   300
            TabIndex        =   90
            Top             =   840
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Web de soporte"
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
            Left            =   300
            TabIndex        =   89
            Top             =   360
            Width           =   1680
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Envio E-Mail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   -74700
         TabIndex        =   83
         Top             =   1020
         Visible         =   0   'False
         Width           =   10545
         Begin VB.CheckBox Check1 
            Caption         =   "Enviar desde el Outlook"
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
            Index           =   14
            Left            =   6720
            TabIndex        =   102
            Tag             =   "Outlook|N|N|||parametros|EnvioDesdeOutlook|||"
            Top             =   960
            Width           =   3075
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   41
            Tag             =   "Direccion e-mail|T|S|||parametros|diremail|||"
            Text            =   "3"
            Top             =   420
            Width           =   4860
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
            Index           =   20
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   42
            Tag             =   "Servidor SMTP|T|S|||parametros|smtpHost|||"
            Text            =   "3"
            Top             =   900
            Width           =   4860
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   43
            Tag             =   "Usuario SMTP|T|S|||parametros|smtpUser|||"
            Text            =   "3"
            Top             =   1440
            Width           =   4860
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
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   7980
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   44
            Tag             =   "Password SMTP|T|S|||parametros|smtpPass|||"
            Text            =   "3"
            Top             =   1440
            Width           =   2220
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
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
            Left            =   300
            TabIndex        =   87
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
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
            Index           =   21
            Left            =   300
            TabIndex        =   86
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
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
            Index           =   22
            Left            =   300
            TabIndex        =   85
            Top             =   1500
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   23
            Left            =   6720
            TabIndex        =   84
            Top             =   1500
            Width           =   1020
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
         Index           =   17
         Left            =   2730
         MaxLength       =   8
         TabIndex        =   82
         Text            =   "1"
         Top             =   4230
         Visible         =   0   'False
         Width           =   660
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
         Left            =   2010
         MaxLength       =   8
         TabIndex        =   81
         Text            =   "1"
         Top             =   4230
         Visible         =   0   'False
         Width           =   660
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
         Left            =   1290
         MaxLength       =   8
         TabIndex        =   80
         Text            =   "1"
         Top             =   4230
         Visible         =   0   'False
         Width           =   660
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
         Left            =   570
         MaxLength       =   8
         TabIndex        =   79
         Text            =   "1"
         Top             =   4230
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Frame Frame3 
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74760
         TabIndex        =   51
         Top             =   2640
         Width           =   11235
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
            ItemData        =   "frmparametros.frx":01B0
            Left            =   4950
            List            =   "frmparametros.frx":01BA
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Tag             =   "Documento proveedores|T|S|||parametros|codinume|||"
            Top             =   1230
            Width           =   2445
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   6
            Left            =   5550
            TabIndex        =   54
            Text            =   "Text2"
            Top             =   540
            Width           =   4875
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   7
            Left            =   840
            TabIndex        =   53
            Text            =   "Text2"
            Top             =   1260
            Width           =   3915
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   5
            Left            =   840
            TabIndex        =   52
            Text            =   "Text2"
            Top             =   540
            Width           =   3915
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
            Left            =   240
            MaxLength       =   8
            TabIndex        =   18
            Tag             =   "Diario proveedores|N|N|0|100|parametros|numdiapr|000||"
            Text            =   "3"
            Top             =   540
            Width           =   540
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
            Left            =   4950
            MaxLength       =   8
            TabIndex        =   19
            Tag             =   "Conceptos facturas proveedores|N|S|0|1000|parametros|concefpr|000||"
            Text            =   "2"
            Top             =   540
            Width           =   540
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
            Left            =   240
            MaxLength       =   8
            TabIndex        =   20
            Tag             =   "Conceptos abonos proveedores|N|S|0|1000|parametros|conceapr|000||"
            Text            =   "1"
            Top             =   1260
            Width           =   525
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   1
            Left            =   2100
            Top             =   1020
            Width           =   240
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   0
            Left            =   6900
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   1
            Left            =   870
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Documento proveedores"
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
            Index           =   19
            Left            =   4950
            TabIndex        =   67
            Top             =   990
            Width           =   2595
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
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
            Left            =   240
            TabIndex        =   57
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto facturas"
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
            Left            =   4950
            TabIndex        =   56
            Top             =   300
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto abonos"
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
            Left            =   240
            TabIndex        =   55
            Top             =   990
            Width           =   2550
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "I.V.A."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   -74760
         TabIndex        =   76
         Top             =   2520
         Width           =   11115
         Begin VB.CheckBox Check1 
            Caption         =   "Mod. importe IVA lin. fra. cliente"
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
            Index           =   21
            Left            =   7410
            TabIndex        =   192
            Tag             =   "Cambia importes iva lineasl|N|N|||parametros|MoficaImporIVALineas|||"
            ToolTipText     =   "Permite modificar el importe de IVA en lineas de factura de cliente"
            Top             =   1830
            Width           =   3585
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Inscrito AEAT facturas DUA"
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
            Left            =   5640
            TabIndex        =   208
            Tag             =   "Inscrito AEAT DUA|N|N|||parametros|inscritoDeclarDUA|||"
            Top             =   3000
            Width           =   3975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Excluir IVA cero en soportado[28]"
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
            Index           =   17
            Left            =   1560
            TabIndex        =   205
            Tag             =   "Excluir IVA cero  repercutido[28]|N|N|||parametros|ExcluirBasesIvaCeroRecibidas303|||"
            Top             =   3000
            Width           =   3975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Facturas rectificativas separadas"
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
            Index           =   13
            Left            =   5640
            TabIndex        =   204
            Tag             =   "303 Fra.rect separadas|N|N|||parametros|RectificativasSeparadas303|||"
            Top             =   2640
            Width           =   3855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Se contabiliza apunte de iva 0"
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
            Index           =   4
            Left            =   7410
            TabIndex        =   191
            Tag             =   "Contabiliza Apunte iva 0l|N|N|||parametros|contabapteiva0|||"
            Top             =   1440
            Width           =   3585
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
            Left            =   1860
            MaxLength       =   8
            TabIndex        =   29
            Tag             =   "Diario modelo 303|N|S|0|100|parametros|diario303|000||"
            Text            =   "1"
            Top             =   810
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   36
            Left            =   2520
            TabIndex        =   189
            Text            =   "Text2"
            Top             =   810
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
            Index           =   35
            Left            =   1860
            MaxLength       =   8
            TabIndex        =   28
            Tag             =   "Concepto modelo 303|N|S|0||parametros|conce303|0||"
            Top             =   360
            Width           =   645
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   35
            Left            =   2520
            TabIndex        =   188
            Text            =   "Text2"
            Top             =   360
            Width           =   4725
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H80000018&
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
            Left            =   3210
            TabIndex        =   183
            Text            =   "Text4"
            Top             =   1680
            Width           =   4035
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
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   31
            Tag             =   "Cuenta HP Acreedora|T|S|0||parametros|ctahpacreedor|||"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H80000018&
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
            Left            =   3210
            TabIndex        =   181
            Text            =   "Text4"
            Top             =   1230
            Width           =   4035
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
            Index           =   33
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   30
            Tag             =   "Cuenta HP Deudora|T|S|0||parametros|ctahpdeudor|||"
            Top             =   1230
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Multisección"
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
            Left            =   7410
            TabIndex        =   40
            Tag             =   "es multisecciónl|N|N|||parametros|esmultiseccion|||"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Caption         =   "349. Presentacion mensual"
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
            Left            =   7410
            TabIndex        =   39
            Tag             =   "Periodo mensual|N|N|||parametros|Presentacion349Mensual|||"
            Top             =   360
            Width           =   3045
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Presentación mensual"
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
            Left            =   1560
            TabIndex        =   37
            Tag             =   "Periodo mensual|N|N|||parametros|periodos|||"
            Top             =   2640
            Width           =   3015
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Modificar apuntes factura"
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
            TabIndex        =   38
            Tag             =   "Mod apuntes factura|N|N|||parametros|modhcofa|||"
            Top             =   720
            Width           =   2910
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C00000&
            X1              =   1440
            X2              =   11040
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label1 
            Caption         =   "Modelo 303"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   34
            Left            =   120
            TabIndex        =   206
            Top             =   2280
            Width           =   1410
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   3
            Left            =   1590
            Top             =   870
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
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
            Index           =   31
            Left            =   150
            TabIndex        =   190
            Top             =   840
            Width           =   615
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   5
            Left            =   1590
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "H.P.Acreedora"
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
            Index           =   30
            Left            =   120
            TabIndex        =   184
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Image imgCta 
            Height          =   240
            Index           =   2
            Left            =   1560
            Top             =   1710
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "H.P Deudora"
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
            Index           =   29
            Left            =   150
            TabIndex        =   182
            Top             =   1260
            Width           =   1230
         End
         Begin VB.Image imgCta 
            Height          =   240
            Index           =   1
            Left            =   1590
            Top             =   1260
            Width           =   240
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
            Height          =   255
            Index           =   15
            Left            =   150
            TabIndex        =   78
            Top             =   420
            Width           =   1080
         End
         Begin VB.Label Label5 
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
            Left            =   3600
            TabIndex        =   77
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1035
         Left            =   240
         TabIndex        =   73
         Top             =   1560
         Width           =   6765
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
            Left            =   240
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cuentas pérdidas y ganancias|T|S|0||parametros|ctaperga|||"
            Top             =   360
            Width           =   1515
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H80000018&
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
            Left            =   2100
            TabIndex        =   75
            Text            =   "Text4"
            Top             =   360
            Width           =   4455
         End
         Begin VB.Image imgCta 
            Height          =   240
            Index           =   0
            Left            =   1860
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta pérdidas y ganancias"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   -30
            Width           =   2985
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4305
         Left            =   7050
         TabIndex        =   72
         Top             =   840
         Width           =   4455
         Begin VB.CheckBox Check1 
            Caption         =   "Mantenimiento cuentas con saldos"
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
            Index           =   20
            Left            =   240
            TabIndex        =   209
            Tag             =   "c|N|N|||parametros|colCtasConSaldos|||"
            Top             =   3240
            Width           =   4095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SII. Proveedor por fecha recepción"
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
            Left            =   240
            TabIndex        =   207
            Tag             =   "c|N|N|||parametros|SII_FraPro_Fecregis|||"
            Top             =   2760
            Width           =   4095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Comprobar asientos descuadrados"
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
            Left            =   240
            TabIndex        =   203
            Tag             =   "c|N|N|||parametros|ComprobarAsientosIncio|||"
            Top             =   2280
            Width           =   3855
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
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   193
            Tag             =   "Nro Ariges|N|N|||parametros|nroariges|||"
            Text            =   "2"
            Top             =   3750
            Width           =   405
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ofertar Fechas Ejercicio Actual"
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
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Tag             =   "c|N|N|||parametros|fecejeract|||"
            Top             =   1815
            Width           =   3855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Gran empresa (8 y 9)"
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
            Index           =   15
            Left            =   240
            TabIndex        =   6
            Tag             =   "c|N|N|||parametros|granempresa|||"
            Top             =   1365
            Width           =   3015
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Agencia de viajes"
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
            Left            =   240
            TabIndex        =   5
            Tag             =   "c|N|N|||parametros|agenciaviajes|||"
            Top             =   780
            Width           =   2235
         End
         Begin VB.CheckBox Check1 
            Caption         =   "I.V.A. por fecha de pago"
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
            Left            =   240
            TabIndex        =   4
            Tag             =   "Constructoras|N|N|||parametros|constructoras|||"
            Top             =   330
            Width           =   3075
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Contabilizar factura automáticamente(oculto)"
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
            Height          =   510
            Index           =   6
            Left            =   3840
            TabIndex        =   8
            Tag             =   "Asiento automatico|N|N|||parametros|ContabilizaFact|||"
            Top             =   3480
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Nro.Ariges asociado"
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
            Index           =   32
            Left            =   240
            TabIndex        =   194
            Top             =   3780
            Width           =   2445
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Analítica"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   240
         TabIndex        =   68
         Top             =   2640
         Width           =   6765
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   735
            Left            =   3810
            TabIndex        =   100
            Top             =   1650
            Width           =   2775
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
            Left            =   5910
            MaxLength       =   1
            TabIndex        =   98
            Tag             =   "Grupo 2|T|S|||parametros|Subgrupo2|||"
            Text            =   "Text1"
            Top             =   1830
            Width           =   600
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
            Left            =   2550
            MaxLength       =   3
            TabIndex        =   14
            Tag             =   "Grupo 1|T|S|||parametros|Subgrupo1|||"
            Text            =   "Text1"
            Top             =   1830
            Width           =   600
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Grabar C.C. en la contabilización de facturas"
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
            Index           =   2
            Left            =   300
            TabIndex        =   10
            Tag             =   "Autocoste|N|S|||paramertros|CCenFacturas|||"
            Top             =   780
            Width           =   4785
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
            Left            =   6150
            MaxLength       =   1
            TabIndex        =   13
            Tag             =   "Grupo|T|S|||parametros|grupoord|||"
            Text            =   "2"
            Top             =   1260
            Width           =   360
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
            Left            =   4500
            MaxLength       =   1
            TabIndex        =   12
            Tag             =   "Grupo|T|S|||parametros|grupovta|||"
            Text            =   "1"
            Top             =   1260
            Width           =   360
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
            Left            =   2040
            MaxLength       =   1
            TabIndex        =   11
            Tag             =   "Grupo|T|S|||parametros|grupogto|||"
            Text            =   "Text1"
            Top             =   1260
            Width           =   360
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Contabilidad analítica"
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
            Left            =   300
            TabIndex        =   9
            Tag             =   "Autocoste|N|N|||paramertros|autocoste|||"
            Top             =   360
            Width           =   2460
         End
         Begin VB.Label Label1 
            Caption         =   "Subgrupo a 3 dígitos"
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
            Left            =   3810
            TabIndex        =   99
            Top             =   1890
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Subgrupo a 3 dígitos"
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
            Left            =   300
            TabIndex        =   97
            Top             =   1890
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Otro Grupo Analitica"
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
            Left            =   4950
            TabIndex        =   71
            Top             =   1305
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Grupo de ventas"
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
            Left            =   2700
            TabIndex        =   70
            Top             =   1305
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Grupo de gastos"
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
            Left            =   300
            TabIndex        =   69
            Top             =   1305
            Width           =   1875
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
         Index           =   0
         Left            =   480
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Fecha inicio|F|N|||parametros|fechaini|dd/mm/yyyy|S|"
         Text            =   "1/2/3"
         Top             =   1110
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
         Index           =   1
         Left            =   2385
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha fin|F|N|||parametros|fechafin|dd/mm/yyyy||"
         Text            =   "1/2/3"
         Top             =   1110
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74760
         TabIndex        =   58
         Top             =   720
         Width           =   11235
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
            Left            =   240
            MaxLength       =   8
            TabIndex        =   17
            Tag             =   "Conceptos abonos clientes|N|S|0|1000|parametros|conceacl|000||"
            Text            =   "3"
            Top             =   1320
            Width           =   540
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
            Left            =   4950
            MaxLength       =   8
            TabIndex        =   16
            Tag             =   "Conceptos facturas clientes|N|S|0|1000|parametros|concefcl|000||"
            Text            =   "2"
            Top             =   540
            Width           =   540
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
            Left            =   240
            MaxLength       =   8
            TabIndex        =   15
            Tag             =   "Diario clientes|N|N|0|100|parametros|numdiacl|000||"
            Text            =   "1"
            Top             =   540
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   9
            Left            =   840
            TabIndex        =   61
            Text            =   "Text2"
            Top             =   540
            Width           =   3945
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   10
            Left            =   5550
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   540
            Width           =   4785
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
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
            Index           =   11
            Left            =   840
            TabIndex        =   59
            Text            =   "Text2"
            Top             =   1320
            Width           =   3945
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   3
            Left            =   2070
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgConcep 
            Height          =   240
            Index           =   2
            Left            =   6900
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   0
            Left            =   900
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto abonos"
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
            Index           =   11
            Left            =   240
            TabIndex        =   64
            Top             =   1080
            Width           =   2010
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto facturas"
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
            Index           =   10
            Left            =   4950
            TabIndex        =   63
            Top             =   270
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Diario"
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
            Left            =   240
            TabIndex        =   62
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Image imgPathFich 
         Height          =   240
         Index           =   0
         Left            =   2400
         Top             =   5400
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Path ficheros integ "
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
         Index           =   33
         Left            =   360
         TabIndex        =   196
         Top             =   5400
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Importe limite 347"
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
         Index           =   13
         Left            =   -67230
         TabIndex        =   187
         Top             =   2010
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Último período"
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
         Index           =   14
         Left            =   -67200
         TabIndex        =   186
         Top             =   1650
         Width           =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha activa"
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
         Index           =   27
         Left            =   4830
         TabIndex        =   101
         Top             =   870
         Width           =   1320
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   6180
         Picture         =   "frmparametros.frx":01DF
         Top             =   870
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   3390
         Picture         =   "frmparametros.frx":026A
         Top             =   870
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "frmparametros.frx":02F5
         Top             =   870
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inicio"
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
         Left            =   480
         TabIndex        =   66
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha fin"
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
         Left            =   2385
         TabIndex        =   65
         Top             =   870
         Width           =   1020
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1020
      Top             =   4620
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
      Left            =   9450
      TabIndex        =   48
      Top             =   7230
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   11310
      TabIndex        =   106
      Top             =   150
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
Attribute VB_Name = "frmparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 102

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmC1 As frmConceptos
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmD1 As frmTiposDiario
Attribute frmD1.VB_VarHelpID = -1
Private WithEvents frmI As frmIVA
Attribute frmI.VB_VarHelpID = -1


Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCo As frmConceptos
Attribute frmCo.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmCta2 As frmColCtas
Attribute frmCta2.VB_VarHelpID = -1


Dim Rs As ADODB.Recordset
Dim Modo As Byte
Dim i As Integer

Dim Cad As String
Dim CadB As String
Dim Indice As Integer

Private MaxLen As Integer 'Para los txt k son de ultimo nivel o de nivel anterior
                          'Ej:



Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim ModificaClaves As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1


    Select Case Modo
    Case 0
        'Preparao para modificar
        PonerModo 2
        
    Case 3
        If DatosOK Then
            'Cambiamos el path
            'CambiaPath True
            CambiaBarrasPath True
            If InsertarDesdeForm(Me) Then
                InsertarModificar True
                PonerModo 0
            End If
            CambiaBarrasPath False
        End If
    
    Case 4
        'Modificar
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            'CambiaPath True
            
            
            ModificaClaves = False
            If vUsu.Nivel = 0 Then
                If vParam.fechaini <> Text1(0).Text Then
                    ModificaClaves = True
                    Cad = " fechaini = '" & Format(vParam.fechaini, FormatoFecha) & "'"
                End If
            End If
            CambiaBarrasPath True
            If ModificaClaves Then
                If ModificaDesdeFormularioClaves(Me, Cad) Then
                    ReestableceVPARAM
                    PonerModo 0
                End If
            Else
                If ModificaDesdeFormulario(Me) Then
                    If InsertarModificar(False) Then
                        If InsertarModificarTesoreria(False) Then
                            PonerModo 0
                        End If
                    End If
                End If
            End If
            CambiaBarrasPath False
        End If

    End Select
    
    'Si el modo es 0 significa k han insertado o modificado cosas
    If Modo = 0 Then _
        MsgBox "Para que los cambios tengan efecto debe reiniciar la aplicación.", vbExclamation
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub CambiaBarrasPath(paraInsertar As Boolean)
    If Text1(32).Text = "" Then Exit Sub
    
    If paraInsertar Then
        Text1(32).Text = Replace(Text1(32).Text, "\", "#")
        Text1(32).Text = Replace(Text1(32).Text, "#", "\\")
    Else
        Text1(32).Text = Replace(Text1(32).Text, "\\", "#")
        Text1(32).Text = Replace(Text1(32).Text, "#", "\")
    End If

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
If Modo = 4 Then PonerCampos
PonerModo 0
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub Form_Load()

    Me.Icon = frmppal.Icon

    SSTab1.Tab = 0
    Me.top = 200
    Me.Left = 100
        ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 4
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    imgCta2(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image2.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgiva.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgCta2(4).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    imgCta2(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    For i = 7 To 10
        imgCta2(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    For i = 0 To 2
        imgCta(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    For i = 0 To 3
        imgDiario(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    For i = 0 To 5
        imgConcep(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    For i = 0 To 1
        Image3(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    imgPathFich(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    'Se muestra unicamente si hay tesorería
    Me.SSTab1.TabEnabled(4) = (vEmpresa.TieneTesoreria)
    Me.SSTab1.TabVisible(4) = (vEmpresa.TieneTesoreria)
    Me.SSTab1.TabEnabled(5) = (vEmpresa.TieneTesoreria)
    Me.SSTab1.TabVisible(5) = (vEmpresa.TieneTesoreria)
    
    PonerLongitudCampoNivelAnterior
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = "Select * from parametros "
    adodc1.Refresh
    Limpiar Me
    If adodc1.Recordset.EOF Then
        'No hay datos
        
        PonerModo 3
    Else
        If Not vParamT Is Nothing Then FrameOpAseguradas.visible = vParamT.TieneOperacionesAseguradas
    
        PonerCampos
        PonerModo 0
        'Campos que nos se tocaran los ponemos con colorcitos bonitos
        If vUsu.Nivel <> 0 Then
            Text1(0).BackColor = &H80000018
            Text1(1).BackColor = &H80000018
        End If
    End If
    Toolbar1.Buttons(1).Enabled = (vUsu.Nivel <= 1)
    cmdAceptar.Enabled = (vUsu.Nivel <= 1)
    

End Sub

Private Sub frmC_Selec(vFecha As Date)
    imgFec(1).Tag = vFecha
End Sub

Private Sub frmCo_DatoSeleccionado(CadenaSeleccion As String)
    imgConcep(1).Tag = CadenaSeleccion
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Select Case Indice
        Case 18
            Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1)
            Text4.Text = RecuperaValor(CadenaSeleccion, 2)
        Case 33
            Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1)
            Text7.Text = RecuperaValor(CadenaSeleccion, 2)
        Case 34
            Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1)
            Text8.Text = RecuperaValor(CadenaSeleccion, 2)
    End Select
End Sub

Private Sub frmCta2_DatoSeleccionado(CadenaSeleccion As String)
    Me.Tag = CadenaSeleccion
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    imgDiario(1).Tag = CadenaSeleccion
End Sub



Private Sub frmF_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
    Select Case i
    Case 0
        Text2(0).Text = Cad
    End Select
End Sub

Private Sub frmI_DatoSeleccionado(CadenaSeleccion As String)
    txtIVA(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtIVA(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    
    i = Index
    Select Case Index
        Case 0
            If Text2(0).Text <> "" Then
                If IsDate(Text2(0).Text) Then frmF.Fecha = CDate(Text2(0).Text)
            End If
    End Select
    frmF.Show vbModal
    Set frmF = Nothing

End Sub

Private Sub Image2_Click()
    Set frmD1 = New frmTiposDiario
    frmD1.DatosADevolverBusqueda = "0|1|"
    frmD1.Show vbModal
    Set frmD1 = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
    i = Index
    Set frmC1 = New frmConceptos
    frmC1.DatosADevolverBusqueda = "0|1|"
    frmC1.Show vbModal
    Set frmC1 = Nothing
End Sub


Private Sub imgConcep_Click(Index As Integer)
    
    imgConcep(1).Tag = ""
    Select Case Index
    Case 0, 1
        imgConcep(0).Tag = 6
    Case 2, 3
        imgConcep(0).Tag = 8
    Case 4
        imgConcep(0).Tag = 19
    Case 5
        imgConcep(0).Tag = 30
    End Select
    Set frmCo = New frmConceptos
    frmCo.DatosADevolverBusqueda = "0|1|"
    frmCo.Show vbModal
    Set frmCo = Nothing
    Index = CInt(imgConcep(0).Tag) + Index
    If imgConcep(1).Tag <> "" Then
        Text1(Index).Text = Format(RecuperaValor(imgConcep(1).Tag, 1), "000")
        Text2(Index).Text = RecuperaValor(imgConcep(1).Tag, 2)
    End If
End Sub

Private Sub imgcta_Click(Index As Integer)
    Select Case Index
        Case 0
            Indice = 18
        Case 1
            Indice = 33
        Case 2
            Indice = 34
    End Select
    
    
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    frmCta.DatosADevolverBusqueda = "0|1|"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
        
    PonFoco Text1(Indice)
End Sub

Private Sub imgCta2_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    Set frmCta2 = New frmColCtas
    Me.Tag = ""
    'Para el text de despues
    frmCta2.DatosADevolverBusqueda = "0|1"
    frmCta2.ConfigurarBalances = 0
    frmCta2.Show vbModal
    Set frmCta2 = Nothing
    If Me.Tag <> "" Then
        Text6(Index).Text = RecuperaValor(Me.Tag, 2)
        Text5(Index).Text = RecuperaValor(Me.Tag, 1)
    End If
    Me.Tag = ""
End Sub

Private Sub imgDiario_Click(Index As Integer)
    imgDiario(1).Tag = ""
    Select Case Index
        Case 0
            Index = 9
        Case 1
            Index = 5
        Case 2
            Index = 24
        Case 3
            Index = 36
    End Select
    
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
    If imgDiario(1).Tag <> "" Then
        Text1(Index).Text = Format(RecuperaValor(imgDiario(1).Tag, 1), "000")
        Text2(Index).Text = RecuperaValor(imgDiario(1).Tag, 2)
    End If
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim F As Date
    'En los tag
    'En el 0 tendremos quien lo ha llamado y en el 1 el valor que devuelve
    imgFec(0).Tag = Index
    F = Now
    imgFec(1).Tag = ""
    If Text1(Index).Text <> "" Then
        If IsDate(Text1(Index).Text) Then F = Text1(Index).Text
    End If
    Set frmC = New frmCal
    frmC.Fecha = F
    frmC.Show vbModal
    Set frmC = Nothing
    If imgFec(1).Tag <> "" Then
        If IsDate(imgFec(1).Tag) Then Text1(Index).Text = Format(CDate(imgFec(1).Tag), "dd/mm/yyyy")
    End If
End Sub

Private Sub imgiva_Click()
    Set frmI = New frmIVA
    frmI.DatosADevolverBusqueda = "0|1|"
    frmI.Show vbModal
    Set frmI = Nothing
End Sub


Private Sub imgPathFich_Click(Index As Integer)

    If Modo = 0 Then Exit Sub
    
    Cad = GetFolder("Carpeta integracion")
    If Cad = "" Then Exit Sub
    Text1(32).Text = Cad
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
            Case 1:  KEYFecha KeyAscii, 1
            Case 31:  KEYFecha KeyAscii, 2
            
            Case 18:  KEYCuentas KeyAscii, 0
            Case 33:  KEYCuentas KeyAscii, 1
            Case 34:  KEYCuentas KeyAscii, 2
            
            Case 9:  KEYDiario KeyAscii, 0
            Case 5:  KEYDiario KeyAscii, 1
            Case 24:  KEYDiario KeyAscii, 2
            Case 36:  KEYDiario KeyAscii, 3
            
            Case 10:  KEYConcepto KeyAscii, 2
            Case 11:  KEYConcepto KeyAscii, 3
            Case 6:  KEYConcepto KeyAscii, 0
            Case 7:  KEYConcepto KeyAscii, 1
            Case 23:  KEYConcepto KeyAscii, 4
            Case 35:  KEYConcepto KeyAscii, 5
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


Private Sub KEYCuentas(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgcta_Click Indice
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYDiario(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgDiario_Click (Indice)
End Sub

Private Sub KEYConcepto(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgConcep_Click (Indice)
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
    Dim Cad As String
    Dim SQL As String
    ''Quitamos blancos por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)

    If Index > 1 And Index <> 31 Then FormateaCampo Text1(Index)     'Formateamos el campo si tiene valor excepto las fechas
    
    'Si queremos hacer algo ..
    Select Case Index
    Case 0, 1, 31
        If Text1(Index).Text = "" Then Exit Sub
        If Not EsFechaOK(Text1(Index)) Then
            MsgBox "Fecha incorrecta : " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Text1(Index).SetFocus
            Exit Sub
        End If
                        
    Case 2, 3, 4
        If Text1(Index).Text = "" Then Exit Sub
        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(Index).Text, "T")
        If SQL = "" Then
            MsgBox "La cuenta no existe: " & Text1(Index).Text, vbExclamation
            Text1(Index).Text = ""
            Text1(Index).SetFocus
        End If
    Case 5, 9, 24, 36
     ' Diarios
       If Not IsNumeric(Text1(Index).Text) Then Exit Sub
       SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(Index).Text)
       If SQL = "" Then
            SQL = "Codigo incorrecto"
            Text1(Index).Text = "-1"
        End If
       Text2(Index).Text = SQL
    Case 6, 7, 10, 11, 23, 35
       'Conceptos
       Cad = "NO"
       If Not IsNumeric(Text1(Index).Text) Then
            SQL = ""
            MsgBox "Campo debe ser numerico", vbExclamation
            Text1(Index).Text = ""
        Else
            Cad = "codconce"
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(Index).Text, "N", Cad)
            If Cad = "codconce" Then
                'NO existe el concepto
                MsgBox "No existe el concepto", vbExclamation
                Text1(Index).Text = ""
            Else
                Cad = ""
            End If
        End If
        
        Text2(Index).Text = SQL
        If Cad <> "" Then PonFoco Text1(Index)
    Case 18
        Cad = Text1(18).Text
        If CuentaCorrectaUltimoNivel(Cad, SQL) Then
            Text1(18).Text = Cad
            Text4.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            If Modo > 2 And Text1(18).Text <> "" Then PonFoco Text1(18)
            Text1(18).Text = "" ' cad
            Text4.Text = "" 'SQL
            
        End If
    
    Case 28, 29
        Cad = Trim(Text1(Index).Text)
        If Cad <> "" Then
            If Not IsNumeric(Cad) Then
                MsgBox Cad & ": No es un campo numerico", vbExclamation
                Text1(Index).Text = ""
                Text1(Index).SetFocus
            End If
        End If
        
    Case 33
        Cad = Text1(33).Text
        If Cad = "" Then
            Text7.Text = ""
            Exit Sub
        End If
        If CuentaCorrectaUltimoNivel(Cad, SQL) Then
            Text1(33).Text = Cad
            Text7.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text1(33).Text = "" ' cad
            Text7.Text = "" 'SQL
            If Modo > 2 And Cad <> "" Then Text1(33).SetFocus
        End If
        
    Case 34
        Cad = Text1(34).Text
        If Cad = "" Then
            Text8.Text = ""
            Exit Sub
        End If
        If CuentaCorrectaUltimoNivel(Cad, SQL) Then
            Text1(34).Text = Cad
            Text8.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text1(34).Text = Cad
            Text8.Text = SQL
            If Modo > 2 Then
                If Cad <> "" Then PonFoco Text1(34)
            End If
        End If
        
    Case 30 ' nro de ariges asociado
        PonerFormatoEntero Text1(30)
    
    End Select
    '---
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim Valor As Boolean
    Dim B As Boolean
    
    
    
    Modo = Kmodo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    
    Select Case Kmodo
        Case 0
            'Preparamos para ver los datos
            Valor = True
    
        Case 3
            'Preparamos para que pueda insertar
            Valor = False
    
        Case 4
            Valor = False
    
    End Select
    
    cmdAceptar.visible = Modo > 0
    cmdCancelar.visible = Modo > 0
    
    For i = 0 To Text1.Count - 1
        If i <> 30 And i <> 32 Then Text1(i).BackColor = vbWhite
    Next i
    
    Combo1.BackColor = vbWhite
    Combo2.BackColor = vbWhite
    Combo3.BackColor = vbWhite
    
    
    'Ponemos los valores
    For i = 0 To Text1.Count - 1
        If i <> 30 And i <> 32 Then Text1(i).Locked = Valor
    Next i
    Frame1.Enabled = Not Valor
    Frame4.Enabled = Not Valor
    Frame6.Enabled = Not Valor
    Frame3.Enabled = Not Valor
    Frame5.Enabled = Not Valor
    Frame2.Enabled = Not Valor
    Frame7.Enabled = Not Valor
    Frame14.Enabled = Not Valor
    Frame15.Enabled = Not Valor
    Frame16.Enabled = Not Valor
    Frame11.Enabled = Not Valor
    
    For i = 0 To FrameTalones.Count - 1
        FrameTalones(i).Enabled = Not Valor
    Next i
    FrameValDefecto.Enabled = Not Valor
    FrameOpAseguradas.Enabled = Not Valor
    Frame66.Enabled = Not Valor

    For i = 0 To imgDiario.Count - 1
        imgDiario(i).Enabled = Not Valor
    Next i
    For i = 0 To imgConcep.Count - 1
        imgConcep(i).Enabled = Not Valor
    Next i
    
    BloqueaTXT Text1(32), Modo = 0
    If Modo > 0 Then Me.imgPathFich(0).Enabled = Not Valor
    
    
    'Campos que solo estan habilitados para insercion
    If Not Valor Then
        Text1(0).Locked = (vUsu.Nivel >= 1)
        Text1(1).Locked = (vUsu.Nivel >= 1)
    End If
    For i = 0 To imgFec.Count - 1
        imgFec(i).Enabled = Not Text1(0).Locked
    Next i
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
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



Private Sub PonerCampos()
    Dim Cam As String
    Dim tabla As String
    Dim Cod As String
    
        If adodc1.Recordset.EOF Then Exit Sub
        If PonerCamposForma(Me, adodc1) Then
           'Correcto, ponemos los datos auxiliares
           '----------------------------------------
           ' Diarios
           Cam = "desdiari"
           tabla = "tiposdiario"
           Cod = "numdiari"
           Text2(9).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(9).Text)
           Text2(5).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(5).Text)
           Text2(24).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(24).Text)
           Text2(36).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(36).Text)
           
           'Conceptos
           Cam = "nomconce"
           tabla = "conceptos"
           Cod = "codconce"
           Text2(10).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(10).Text)
           Text2(11).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(11).Text)
           Text2(6).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(6).Text)
           Text2(7).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(7).Text)
           Text2(23).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(23).Text)
           Text2(35).Text = DevuelveDesdeBD(Cam, tabla, Cod, Text1(35).Text)
           
           'Cuenta de pérdidas y ganancias
           Text4.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(18).Text, "T")
           Text7.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(33).Text, "T")
           Text8.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(34).Text, "T")
           
           CargarDatos
            
           CargarDatosTesoreria
           
           
        End If
End Sub
'
Private Function DatosOK() As Boolean
    Dim B As Boolean
    Dim J As Integer
    
    
    DatosOK = False
    
    'Si esta marcado la analitica, entonces debe tener valor grupo ventas y grupo gastos
    If Me.Check1(0).Value = 1 Then
        Text1(2).Text = Trim(Text1(2).Text)
        Text1(3).Text = Trim(Text1(3).Text)
        If Text1(2).Text = "" Or Text1(3).Text = "" Then
            MsgBox "Si selecciona la contabilidad analítica debe poner valor al grupo gastos y ventas.", vbExclamation
            Exit Function
        End If
    End If
    

    'NO puede marcar constructora y Agencia viajes a la vez
    If Check1(8).Value = 1 And Check1(9).Value Then
        MsgBox "No puede marcar Contructora y Agencia de viajes a la vez", vbExclamation
        Exit Function
    End If
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Si tiene puesta fecha Activa, esta no puede ser ni menor que fechaini ni mayor que fecha fin
    If Text1(31).Text <> "" Then
        If CDate(Text1(31).Text) < CDate(Text1(0).Text) Or CDate(Text1(31).Text) > DateAdd("yyyy", 1, CDate(Text1(1).Text)) Then
            MsgBox "Fecha activa debe estar entre incio ejercicio - fin ejercicio siguiente", vbExclamation
            B = False
            Exit Function
        End If
    End If
    
    
    
    
    
    'Año natural
    
    If Modo = 1 Then
        If CDate(Text1(0).Text) >= CDate(Text1(1).Text) Then
            MsgBox "Fecha inicio mayor o igual que la fecha final de ejercicio.", vbExclamation
            B = False
        End If
    End If
    
    i = DateDiff("d", CDate(Text1(0).Text), CDate(Text1(1).Text))
    If i > 366 Then
        MsgBox "Rango fecha mayor a un año", vbExclamation
        B = False
    Else
        If i < 364 Then If MsgBox("Año ejercicio inferior a 365 dias. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then B = False
    End If
    
    
    'Comprobamos que si el periodo de liquidacion de IVA es mensual
    'Comprobaremos que el valor del periodo no excede de
    If Check1(5).Value = 1 Then
        i = 12  'Mensual
    Else
        i = 4   'Trimestral
    End If
    If Text1(15).Text <> "" Then
        J = CInt(Text1(15).Text)
        If J = 0 Then
            MsgBox "El periodo de liquidación no puede ser 0", vbExclamation
            Exit Function
        End If
        If J > i Then
            MsgBox "Periodo de liquidacion incorrecto", vbExclamation
            Exit Function
        End If
    End If
    
    
    
    ' Parte correspondiente al arimoney
    If vEmpresa.TieneTesoreria Then
        Dim C As String
        
        If B Then
            J = 0
            C = ""
            
            If Me.Check5(3).Value = 0 Xor Text5(7).Text = "" Then C = C & "-  talones "
            If Me.Check5(4).Value = 0 Xor Text5(8).Text = "" Then C = C & "-  pagarés "
            If Me.Check5(6).Value = 0 Xor Text5(1).Text = "" Then C = C & "-  confirming "
            
            J = Len(Text5(9).Text) + Len(Text5(10).Text)
            If Me.Check5(5).Value = 0 Xor J = 0 Then C = C & "-  efectos "
    '        'Proveedores
            
            If C = "" Then
                'Todo bien. Compruebo esto tb
                'Si pone cta puente, la primera de las ctas es obligada
                If Me.Check5(5).Value = 1 And Text5(9).Text = "" Then C = "   La cuenta puente es campo obligatorio"
            End If
            
            If C <> "" Then
                C = Trim(Mid(C, 2)) 'quitamos el primer guion
                C = C & vbCrLf & vbCrLf & "Si marca que utiliza la cuenta puente entonces debe indicarla. En otro caso debe dejarla a blancos"
                MsgBox C, vbExclamation
                Exit Function
            End If
            
            C = ""
            'AHora es desde i=5
            For J = 7 To 10
                i = IIf(J = 10, 1, J)
                Text5(i).Text = Trim(Text5(i).Text)
                
                If Text5(i).Text <> "" Then
                    
                    If Len(Text5(i).Text) <> vEmpresa.DigitosUltimoNivel Then
                                                'Aqui tenemos los digitos a ultnivel-1
                        If Len(Text5(i).Text) <> MaxLen Then
                            If i = 7 Then
                                C = C & "talon" & vbCrLf
                            ElseIf i = 8 Then
                                C = C & "pagares" & vbCrLf
                            ElseIf i = 9 Then
                                C = C & "efectos" & vbCrLf
                            Else
                                C = C & "confirming" & vbCrLf
                            End If
                           
                        End If
                    End If
                End If
            Next J
            
            If C <> "" Then
                C = "Error en la longitud de las cuentas para: " & vbCrLf & C
                C = C & vbCrLf & "Ha de tener longitud "
                C = C & "a " & vEmpresa.DigitosUltimoNivel & " digitos o "
                C = C & "a " & MaxLen & " digitos. "
                C = C & vbCrLf & "¿Continuar?"
                If MsgBox(C, vbQuestion + vbYesNo) <> vbYes Then Exit Function
            End If
            
            DatosOK = False
            
            'Si es nuevo
            If Modo = 1 Then Text5(3).Text = 1
            
        End If
    End If
    
        
        
    If B Then
        If Combo2.ListIndex = 1 And Me.Check1(22).Value = 1 Then
                C = "Ha seleccionado numero factura en apunte/documento y " & vbCrLf
                C = C & " añadir nº factura en ampliacion apunte"
                C = C & vbCrLf & "¿Continuar?"
                If MsgBox(C, vbQuestion + vbYesNo) <> vbYes Then B = False
        End If
    End If
    
    If B Then
        If Text1(32).Text <> "" Then B = HacerDir
    End If
    DatosOK = B

End Function

Private Function HacerDir() As Boolean
    On Error Resume Next
    HacerDir = False
    Cad = Dir(Text1(32).Text, vbDirectory)
    If Err.Number <> 0 Then
        MuestraError Err.Number
    Else
        If Cad <> "" Then
            HacerDir = True
        Else
            MsgBox "No existe carpeta", vbExclamation
        End If
    End If
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Sub Text6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        'Modificar
         PonerModo 4
    End Select
End Sub



Private Sub ReestableceVPARAM()
    Set vParam = Nothing
    Set vParam = New Cparametros
    vParam.Leer
    vParam.FijarAplicarFiltrosEnCuentas vEmpresa.nomempre
    If vEmpresa.TieneTesoreria Then vParamT.Leer
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
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
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

'#####################################
'######### INMOVILIZADO


Private Function InsertarModificar(Insertar As Boolean) As Boolean
On Error GoTo EInsertarModificar
    InsertarModificar = False

    Cad = DevuelveDesdeBD("codigo", "paramamort", "codigo", "1")


    If Cad <> "" Then
        'Modificar
        Cad = "UPDATE paramamort SET tipoamor= " & Combo3.ListIndex + 1
        Cad = Cad & ", intcont= " & Check1(16).Value
        Cad = Cad & ", ultfecha= '" & Format(Text2(0).Text, FormatoFecha)
        Cad = Cad & "', condebes= " & DBSet(Text2(2), "N", "S")
        Cad = Cad & ", conhaber= " & DBSet(Text2(3), "N", "S")
        Cad = Cad & ", numdiari=  " & DBSet(Text2(1), "N", "S")
        
        Cad = Cad & ", codiva = " & DBSet(txtIVA(0), "N", "S")
        Cad = Cad & ", preimpreso =" & Check2.Value
        Cad = Cad & " WHERE codigo=1"
    Else
        'INSERTAR
        Cad = "INSERT INTO paramamort (codigo, tipoamor, intcont, ultfecha, condebes, conhaber, numdiari,codiva,preimpreso) VALUES (1,"
        Cad = Cad & Combo3.ListIndex + 1 & "," & Me.Check1(16).Value & ",'" & Format(Text2(0).Text, FormatoFecha)
        Cad = Cad & "'," & DBSet(Text2(2), "N", "S") & "," & DBSet(Text2(3), "N", "S") & "," & DBSet(Text2(1), "N", "S")
        Cad = Cad & ",'" & DBSet(txtIVA(0), "N", "S") & "'," & Check2.Value & ")"
    End If
    Conn.Execute Cad
    InsertarModificar = True
    Exit Function
EInsertarModificar:
    MuestraError Err.Number, "Insertar-Modificar"
End Function


Private Function CargarDatos() As Boolean
Dim Cad As String
Dim Rs As ADODB.Recordset

On Error GoTo ECargarDatos
    CargarDatos = False
    Set Rs = New ADODB.Recordset
    Cad = "Select * from paramamort where codigo=1"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        CargarDatos = True
        '------------------  Ponemos los datos
        Combo3.ListIndex = Rs!tipoamor - 1
        Check1(16).Value = Rs!intcont
        Text2(0).Text = Format(Rs!ultfecha, "dd/mm/yyyy")
        Text2(1).Text = DBLet(Rs!NumDiari)
        Text2_LostFocus 1
        Text2(2).Text = DBLet(Rs!condebes)
        Text2_LostFocus 2
        Text2(3).Text = DBLet(Rs!conhaber)
        Text2_LostFocus 3
        txtIVA(0).Text = DBLet(Rs!CodIVA)
        txtIVA_LostFocus 0
        Check2.Value = Rs!Preimpreso
    End If
    Rs.Close
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando parametros"
    Set Rs = Nothing
End Function




Private Sub txtiva_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYIva KeyAscii, Index
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtIVA_LostFocus(Index As Integer)
    If Index = 1 Then Exit Sub
    txtIVA(0).Text = Trim(txtIVA(0).Text)
    If txtIVA(0).Text = "" Then
        txtIVA(1).Text = ""
        Exit Sub
    End If
    
    i = 0
    Cad = ""
    If Not IsNumeric(txtIVA(0).Text) Then
        MsgBox "Tipo IVA debe ser numérico: " & txtIVA(0).Text, vbExclamation
        txtIVA(0).Text = ""
        i = 1
    End If
    If i = 0 Then
        Cad = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", txtIVA(0).Text, "N")
        If Cad = "" Then
            MsgBox "IVA no encontrado: " & txtIVA(0).Text, vbExclamation
            txtIVA(0).Text = ""
        End If
    End If
    txtIVA(1).Text = Cad
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text2_LostFocus(Index As Integer)
    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then
        If Index > 0 Then Text3(Index - 1).Text = ""
        Exit Sub
    End If
    If Index = 0 Then
        'Fecha
        If Not EsFechaOK(Text2(0)) Then
            MsgBox "Fecha incorrecta", vbExclamation
            Text2(0).Text = ""
            Text2(0).SetFocus
            Exit Sub
        End If
        Text2(0).Text = Format(Text2(0).Text)
    Else
        If Not IsNumeric(Text2(Index).Text) Then
            MsgBox "El campo tiene que ser numérico", vbExclamation
            Text2(Index).Text = ""
            Text2(Index).SetFocus
            Exit Sub
        End If
        Select Case Index
        Case 1
             Cad = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text2(1).Text, "N")
             If Cad = "" Then
                    MsgBox "Diario no encontrado: " & Text2(1).Text, vbExclamation
                    Text2(1).Text = ""
                    Text2(1).SetFocus
            End If
            Text3(0).Text = Cad
        Case 2, 3
                Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text2(Index).Text, "N")
                If Cad = "" Then
                    MsgBox "Concepto NO encontrado: " & Text2(Index).Text, vbExclamation
                    Text2(Index).Text = ""
                End If
                Text3(Index - 1).Text = Cad
                
        End Select
    End If
End Sub

Private Sub frmC1_DatoSeleccionado(CadenaSeleccion As String)
    Text2(i + 2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(i + 1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmD1_DatoSeleccionado(CadenaSeleccion As String)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub KEYIva(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgiva_Click
End Sub

'#####################################




'#####################################
'######### TESORERIA

Private Function InsertarModificarTesoreria(Insertar As Boolean) As Boolean
On Error GoTo EInsertarModificar
    InsertarModificarTesoreria = False

    If Not Insertar Then
        'Modificar
        Cad = "UPDATE paramtesor SET ctabenbanc= " & DBSet(Text5(0).Text, "T", "S")
        Cad = Cad & ", par_pen_apli= " & DBSet(Text5(4).Text, "T", "S")
        Cad = Cad & ", responsable= " & DBSet(Text5(2).Text, "T", "S")
        Cad = Cad & ", interesescobrostarjeta= " & DBSet(Text5(16), "N")
        Cad = Cad & ", remesacancelacion= " & DBSet(Text5(9), "T", "S")
        Cad = Cad & ", ctaefectcomerciales=  " & DBSet(Text5(10), "T", "S")
        Cad = Cad & ", taloncta = " & DBSet(Text5(7), "T", "S")
        Cad = Cad & ", pagarecta =" & DBSet(Text5(8).Text, "T", "S")
        Cad = Cad & ", contapag =" & DBSet(Check5(0).Value, "N")
        Cad = Cad & ", generactrpar =" & DBSet(Check5(1).Value, "N")
        Cad = Cad & ", abonocambiado =" & DBSet(Check5(2).Value, "N")
        Cad = Cad & ", comprobarinicio =" & DBSet(Check5(8).Value, "N")
        Cad = Cad & ", nor19xvto =" & DBSet(Check5(10).Value, "N")
        Cad = Cad & ", eliminarecibidosriesgo =" & DBSet(Check5(9).Value, "N")
        Cad = Cad & ", diasmaxavisodesde =" & DBSet(Text5(11).Text, "N", "S")
        Cad = Cad & ", diasmaxavisohasta =" & DBSet(Text5(12).Text, "N", "S")
        Cad = Cad & ", diasmaxsiniestrodesde =" & DBSet(Text5(13).Text, "N", "S")
        Cad = Cad & ", diasmaxsiniestrohasta =" & DBSet(Text5(14).Text, "N", "S")
        Cad = Cad & ", diasavisodesdeprorroga =" & DBSet(Text5(15).Text, "N", "S")
        Cad = Cad & ", fechaasegesfra =" & DBSet(Check5(11).Value, "N")
        Cad = Cad & ", contatalonpte =" & DBSet(Check5(3).Value, "N")
        Cad = Cad & ", contapagarepte =" & DBSet(Check5(4).Value, "N")
        Cad = Cad & ", contaefecpte =" & DBSet(Check5(5).Value, "N")
        Cad = Cad & ", confirmingcta =" & DBSet(Text5(1).Text, "T", "S")
        Cad = Cad & ", contaconfirmpte =" & DBSet(Check5(6).Value, "N")
        
        Cad = Cad & " WHERE codigo=1"
    Else
        'INSERTAR
        Cad = "INSERT INTO paramtesor (codigo, ctabenbanc, par_pen_apli, responsable, interesescobrostarjeta, remesacancelacion, ctaefectcomerciales, "
        Cad = Cad & "taloncta,pagarecta,contapag,generactrpar,abonocambiado,comprobarinicio,nor19xvto,eliminarecibidosriesgo, "
        Cad = Cad & "diasmaxavisodesde,diasmaxavisohasta,diasmaxsiniestrodesde,diasmaxsiniestrohasta,diasavisodesdeprorroga,fechaasegesfra,contatalonpte, "
        Cad = Cad & "contapagarepte,contaefecpte) values (1,"
        Cad = Cad & DBSet(Text5(0).Text, "T", "S") & "," & DBSet(Text5(4).Text, "T", "S") & "," & DBSet(Text5(2).Text, "T", "S") & ","
        Cad = Cad & DBSet(Text5(16), "N") & "," & DBSet(Text5(9), "T", "S") & "," & DBSet(Text5(10), "T", "S") & "," & DBSet(Text5(7), "T", "S") & ","
        Cad = Cad & DBSet(Text5(8).Text, "T", "S") & "," & DBSet(Check5(0).Value, "N") & ","
        Cad = Cad & DBSet(Check5(1).Value, "N") & "," & DBSet(Check5(2).Value, "N") & "," & DBSet(Check5(8).Value, "N") & "," & DBSet(Check5(10).Value, "N") & " ,"
        Cad = Cad & DBSet(Check5(9).Value, "N") & "," & DBSet(Text5(11).Text, "N", "S") & "," & DBSet(Text5(12).Text, "N", "S") & "," & DBSet(Text5(13).Text, "N", "S") & ","
        Cad = Cad & DBSet(Text5(14).Text, "N", "S") & "," & DBSet(Text5(15).Text, "N", "S") & "," & DBSet(Check5(11).Value, "N") & "," & DBSet(Check5(3).Value, "N") & ","
        Cad = Cad & DBSet(Check5(4).Value, "N") & "," & DBSet(Check5(5).Value, "N") & ")"
    
    End If
    Conn.Execute Cad
    InsertarModificarTesoreria = True
    Exit Function
EInsertarModificar:
    MuestraError Err.Number, "Insertar-Modificar"
End Function


Private Function CargarDatosTesoreria() As Boolean
Dim Cad As String
Dim Rs As ADODB.Recordset

On Error GoTo ECargarDatos
    CargarDatosTesoreria = False
    Set Rs = New ADODB.Recordset
    Cad = "Select * from paramtesor where codigo=1"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        CargarDatosTesoreria = True
        '------------------  Ponemos los datos
        Text5(0).Text = DBLet(Rs!ctabenbanc, "T")
        Text5_LostFocus 0
        Text5(4).Text = DBLet(Rs!Par_pen_apli, "T")
        Text5_LostFocus 4
        Text5(2).Text = DBLet(Rs!responsable, "T")
        Text5(16).Text = ""
        If Not IsNull(Rs!InteresesCobrosTarjeta) Then Text5(16).Text = Rs!InteresesCobrosTarjeta
        Text5(9).Text = DBLet(Rs!remesacancelacion, "T")
        Text5_LostFocus 9
        Text5(10).Text = DBLet(Rs!ctaefectcomerciales, "T")
        Text5_LostFocus 10
        Text5(7).Text = DBLet(Rs!taloncta, "T")
        Text5_LostFocus 7
        Text5(8).Text = DBLet(Rs!pagarecta, "T")
        Text5_LostFocus 8
        Text5(1).Text = DBLet(Rs!confirmingcta, "T")
        Text5_LostFocus 1
    
        Check5(0).Value = Rs!contapag
        Check5(1).Value = Rs!generactrpar
        Check5(2).Value = Rs!abonocambiado
        Check5(8).Value = Rs!comprobarinicio
        Check5(10).Value = Rs!nor19xvto
        Check5(9).Value = DBLet(Rs!EliminaRecibidosRiesgo, "N")
        Check5(11).Value = Rs!fechaasegesfra
        
        
        Text5(11).Text = DBLet(Rs!diasmaxavisodesde, "N")
        Text5(12).Text = DBLet(Rs!diasmaxavisohasta, "N")
        Text5(13).Text = DBLet(Rs!diasmaxsiniestrodesde, "N")
        Text5(14).Text = DBLet(Rs!diasmaxsiniestrohasta, "N")
        Text5(15).Text = DBLet(Rs!DiasAvisoDesdeProrroga, "N")
    
        Check5(3).Value = Rs!contatalonpte
        Check5(4).Value = Rs!contapagarepte
        Check5(5).Value = Rs!contaefecpte
        Check5(6).Value = DBLet(Rs!contaconfirmpte, "N")
    End If
    Rs.Close
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando parametros de Tesoreria"
    Set Rs = Nothing
End Function




'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text5_GotFocus(Index As Integer)
    Text5(Index).SelStart = 0
    Text5(Index).SelLength = Len(Text5(Index).Text)
End Sub


Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
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
Private Sub Text5_LostFocus(Index As Integer)
    Dim Cad As String
    Dim SQL As String
    Dim Valor As Currency
    
    ''Quitamos blancos por los lados
    Text5(Index).Text = Trim(Text5(Index).Text)

    'Si queremos hacer algo ..
    Select Case Index
'    Case 1
'        If Text5(Index).Text = "" Then Exit Sub
'        If Not EsFechaOK(Text5(Index)) Then
'            MsgBox "Fecha incorrecta : " & Text5(Index).Text, vbExclamation
'            Text5(Index).Text = ""
'            Text5(Index).SetFocus
'            Exit Sub
'        End If
                        
    Case 0, 4
        Cad = Text5(Index).Text
        If Cad = "" Then
            Text6(Index).Text = ""
            Exit Sub
        End If
        If CuentaCorrectaUltimoNivel(Cad, SQL) Then
            Text5(Index).Text = Cad
            Text6(Index).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text5(Index).Text = "" 'cad
            Text6(Index).Text = "" 'SQL
            If Modo > 2 Then Text5(Index).SetFocus
        End If

    Case 1, 5, 6, 7, 8, 9, 10
        Text6(Index).Text = ""
        If Text5(Index).Text = "" Then Exit Sub
        If Not IsNumeric(Text5(Index).Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            Text5(Index).Text = ""
            Exit Sub
        End If
        i = Len(Text5(Index).Text)
        NumRegElim = InStr(1, Text5(Index).Text, ".")
        If NumRegElim = 0 Then
            If i <> vEmpresa.DigitosUltimoNivel And i <> MaxLen Then
                If MsgBox("Longitud de campo incorrecta. Digitos: " & vEmpresa.DigitosUltimoNivel & " o " & MaxLen & "¿Continuar?", vbQuestion + vbYesNo) <> vbYes Then
                    Text5(Index).Text = ""
                    Exit Sub
                End If
            End If
        End If
        
        'Llegados aqui, si es de ultimo nivel pondre la cuenta
        If NumRegElim > 0 Or i = vEmpresa.DigitosUltimoNivel Then
            Cad = Text5(Index).Text
            If CuentaCorrectaUltimoNivel(Cad, SQL) Then
                Text5(Index).Text = Cad
            Else
                MsgBox SQL, vbExclamation
                Text5(Index).Text = ""
                SQL = ""
            End If
        Else
            SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text5(Index).Text, "T")
            If SQL = "" Then
                MsgBox "No existe la cuenta indicada", vbExclamation
                Text5(Index).Text = ""
            End If
        End If
        Text6(Index).Text = SQL
        
    Case 11, 12, 13, 14, 16
        If Text5(Index).Text = "" Then Exit Sub
        
        If Not IsNumeric(Text5(Index).Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            Text5(Index).Text = ""
            PonFoco Text5(Index)
        Else
            If Index = 16 Then
                If InStr(1, Text5(Index).Text, ",") > 0 Then
                    Valor = ImporteFormateado(Text5(Index).Text)
                Else
                    Valor = CCur(TransformaPuntosComas(Text5(Index).Text))
                End If
                Text5(Index).Text = Format(Valor, FormatoImporte)
            End If
        End If
        
    End Select
    '---
End Sub


Private Sub PonerLongitudCampoNivelAnterior()
    On Error GoTo EPonerLongitudCampoNivelAnterior
    
    
    i = DigitosNivel(vEmpresa.numnivel - 1)
    If i = 0 Then i = 4
    MaxLen = i
    
  
    Exit Sub
EPonerLongitudCampoNivelAnterior:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub ImageAyudaImpcta_Click(Index As Integer)
Dim C As String
Dim C2 As String

    Select Case Index
    Case 0, 2, 3, 5, 6
        'Cancelarcion cliente
        C2 = "cuando dentro del punto ""Recepción de documentos"" se realice la contabilización"
        Select Case Index
        Case 0, 5
            C = "talones"
        Case 2, 6
            C = "pagarés"
        Case Else
            C = "Efectos"
            C2 = "cuando dentro del punto ""Cancelación cliente"" del apartado Remesas  se realice el abono de la remesa,"
        End Select
        'If Index > 0 Then C = C & " de PROVEEDORES"
        C = "Para la cancelacion de los " & C & ":" & vbCrLf & vbCrLf
        C = C & "Si tiene marcada la opcion de 'Contabiliza contra cuentas puente', " & C2
        C = C & " tendremos dos opciones:" & vbCrLf
        If Index = 3 Then C = C & "  Efectos descontados" & vbCrLf
        C = C & "    -   Una única cuenta a último nivel (Ej: 4310000), con lo que todos los apuntes irán a esa cuenta genérica." & vbCrLf
        C = C & "    -   Introducir una cuenta raíz a 4 dígitos (Ej: 4310), con lo que el programa creará cuentas a último nivel haciéndolas coincidir con las terminaciones de las cuentas del cliente." & vbCrLf
        
        
        'Nuevo Nov 2009
        If Index = 3 Then
            C = C & vbCrLf & vbCrLf
            C = C & "    Efectos descontados a cobrar " & vbCrLf
            C = C & "    -   Una única cuenta a último nivel (Ej: 4310000), con lo que todos los apuntes irán a esa cuenta genérica." & vbCrLf
            C = C & "    -   Introducir una cuenta raíz a 4 dígitos (Ej: 4310), con lo que el programa creará cuentas a último nivel haciéndolas coincidir con las terminaciones de las cuentas del cliente." & vbCrLf
        End If
        
    Case 1
        C = "Cuenta beneficios bancarios." & vbCrLf & vbCrLf
        C = C & "Si no esta indicada en la configuración del banco  " & vbCrLf
        C = C & "con el que estemos trabajando, utilizará esta cuenta  " & vbCrLf
    Case 4
        C = "Valores que ofertará para la contabilización de cobros/pagos. " & vbCrLf
        C = C & "Luego podrá ser modificado para cada caso  " & vbCrLf
    Case 7
        C = "Responsable para poder firmar en documentos(recibos, cheques)"
        
    Case 8
        C = "Al entrar en la empresa que compruebe si hay posibilidad de eliminar "
        C = C & vbCrLf & "riesgo, tanto en efectos como en talones y pagarés"
    Case 9
        C = "Cuando eliminamos riesgo en talones y pagarés, eliminar tambien en la tabla de  "
        C = C & vbCrLf & "recepcion de documentos."
    Case 10
        C = "Operaciones aseguradas. "
        C = C & vbCrLf & "Dias maximo(desde/hasta) para mostrar avisos de falta de pago y/o de siniestro"

        C = C & vbCrLf & "Check  'Fecha factura...'"
        C = C & vbCrLf & "Para calcular los dias de aviso,riesgo, prorroga... puede coger"
        C = C & vbCrLf & "la fecha factura o la fecha de vencimiento."
    

    Case 11
        C = "Norma 19. "
        C = C & vbCrLf & "Se contabilizara la remesa por fecha vencimiento"
        C = C & vbCrLf & "Tantos apuntes como fechas distintas haya en la remesa"
    
    Case 12
        C = "% Intereses tarjeta"
        C = C & vbCrLf & "Porcentaje anual de interes para las ventas a credito(Forpa: Tarjeta) "
        C = C & vbCrLf & "Calculo:  (% / 365) * dias_desde_vto "
        C = C & vbCrLf & vbCrLf & "Una vez impresos los recibos, si la impresión es correcta, graba columna gastos "

        '
    
    End Select
    MsgBox C, vbInformation
End Sub


