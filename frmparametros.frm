VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmparametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Par�metros Contables"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11880
   Icon            =   "frmparametros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Height          =   465
      Left            =   120
      TabIndex        =   106
      Top             =   6210
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
         TabIndex        =   107
         Top             =   120
         Width           =   2310
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   102
      Top             =   0
      Width           =   1185
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   103
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   104
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
      TabIndex        =   47
      Top             =   6270
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5325
      Left            =   90
      TabIndex        =   48
      Top             =   840
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   9393
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
      Tab(0).Control(6)=   "Text1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(16)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text1(17)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(31)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Clientes - Proveedores "
      TabPicture(1)   =   "frmparametros.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "I.V.A. - Norma 43"
      TabPicture(2)   =   "frmparametros.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(14)"
      Tab(2).Control(1)=   "Label1(13)"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(4)=   "Frame17"
      Tab(2).Control(5)=   "Text1(14)"
      Tab(2).Control(6)=   "Text1(15)"
      Tab(2).Control(7)=   "Text1(13)"
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
      TabCaption(4)   =   "Tesorer�a I"
      TabPicture(4)   =   "frmparametros.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameValDefecto"
      Tab(4).Control(1)=   "Frame66"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Tesorer�a II"
      TabPicture(5)   =   "frmparametros.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameTalones(2)"
      Tab(5).Control(1)=   "FrameOpAseguradas"
      Tab(5).Control(2)=   "FrameTalones(0)"
      Tab(5).Control(3)=   "FrameTalones(1)"
      Tab(5).ControlCount=   4
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
         TabIndex        =   34
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
         TabIndex        =   33
         Tag             =   "Ultimo periodo liquidaci�n I.V.A.|N|S|0|100|parametros|perfactu|||"
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
         TabIndex        =   32
         Tag             =   "Ultimo a�o liquidaci�n I.V.A.|N|S|0|9999|parametros|anofactu|||"
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
         TabIndex        =   183
         Top             =   900
         Width           =   3525
         Begin VB.OptionButton Option1 
            Caption         =   "A�o natural"
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
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton Option1 
            Caption         =   "A�o fiscal"
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
            TabIndex        =   30
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
         Height          =   1935
         Index           =   2
         Left            =   -74820
         TabIndex        =   161
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
            TabIndex        =   163
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
            TabIndex        =   167
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
            TabIndex        =   164
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
            TabIndex        =   165
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
            TabIndex        =   162
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
            TabIndex        =   171
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
            TabIndex        =   169
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
         TabIndex        =   153
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
            TabIndex        =   173
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
            TabIndex        =   174
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
            TabIndex        =   175
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
            TabIndex        =   176
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
            TabIndex        =   177
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
            TabIndex        =   178
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
            TabIndex        =   160
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
            TabIndex        =   159
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
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
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
            TabIndex        =   155
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
            TabIndex        =   154
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
         TabIndex        =   150
         Top             =   2700
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
            TabIndex        =   168
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
            TabIndex        =   151
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
            TabIndex        =   166
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
            TabIndex        =   152
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
         Caption         =   "Pagar�s clientes"
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
         TabIndex        =   147
         Top             =   3960
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
            TabIndex        =   148
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
            TabIndex        =   172
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
            TabIndex        =   170
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
            TabIndex        =   149
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
         TabIndex        =   146
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
            TabIndex        =   129
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
            TabIndex        =   130
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
            TabIndex        =   131
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
            TabIndex        =   132
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
            TabIndex        =   133
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
            TabIndex        =   134
            Top             =   2130
            Width           =   4215
         End
      End
      Begin VB.Frame Frame66 
         Height          =   2175
         Left            =   -74910
         TabIndex        =   139
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   125
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
            TabIndex        =   145
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
            TabIndex        =   144
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
            TabIndex        =   143
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
            TabIndex        =   142
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
         Height          =   2265
         Left            =   -74700
         TabIndex        =   118
         Top             =   810
         Width           =   3945
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
            Left            =   270
            TabIndex        =   120
            Text            =   "Text2"
            Top             =   1620
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
            ItemData        =   "frmparametros.frx":00B4
            Left            =   240
            List            =   "frmparametros.frx":00C4
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   720
            Width           =   2355
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   1590
            Picture         =   "frmparametros.frx":00EF
            Top             =   1320
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
            Left            =   300
            TabIndex        =   137
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de amortizaci�n"
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
            TabIndex        =   135
            Top             =   450
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
         TabIndex        =   111
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
            TabIndex        =   123
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
            TabIndex        =   114
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
            TabIndex        =   122
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
            TabIndex        =   113
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
            TabIndex        =   121
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
            TabIndex        =   112
            Text            =   "Text3"
            Top             =   750
            Width           =   3705
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Contabilizacion autom�tica"
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
            TabIndex        =   136
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
            TabIndex        =   117
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
            TabIndex        =   116
            Top             =   1245
            Width           =   1545
         End
         Begin VB.Label Label3 
            Caption         =   "N� Diario"
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
            TabIndex        =   115
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
         TabIndex        =   108
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
            TabIndex        =   124
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
            Height          =   390
            Index           =   1
            Left            =   2850
            TabIndex        =   109
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
            TabIndex        =   138
            Top             =   900
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.Image imgiva 
            Height          =   240
            Left            =   1740
            Top             =   390
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
            TabIndex        =   110
            Top             =   390
            Width           =   405
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
         Caption         =   "Par�metros importaci�n datos bancarios"
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
         TabIndex        =   91
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
            TabIndex        =   93
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
            TabIndex        =   25
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
            TabIndex        =   92
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
            TabIndex        =   24
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
            TabIndex        =   95
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
            TabIndex        =   94
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
         TabIndex        =   87
         Top             =   3180
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   88
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
         TabIndex        =   82
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
            TabIndex        =   101
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
            TabIndex        =   39
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
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
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
         Index           =   16
         Left            =   2010
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
         Index           =   12
         Left            =   1290
         MaxLength       =   8
         TabIndex        =   79
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
         TabIndex        =   78
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
         Left            =   -74820
         TabIndex        =   49
         Top             =   3090
         Width           =   10995
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
            ItemData        =   "frmparametros.frx":017A
            Left            =   4950
            List            =   "frmparametros.frx":0184
            Style           =   2  'Dropdown List
            TabIndex        =   23
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   20
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
            TabIndex        =   21
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
            TabIndex        =   22
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
            TabIndex        =   66
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
         Height          =   2685
         Left            =   -74760
         TabIndex        =   75
         Top             =   2370
         Width           =   11115
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
            TabIndex        =   189
            Tag             =   "Contabiliza Apunte iva 0l|N|N|||parametros|contabapteiva0|||"
            Top             =   2040
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
            TabIndex        =   27
            Tag             =   "Diario modelo 303|N|S|0|100|parametros|diario303|000||"
            Text            =   "1"
            Top             =   930
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
            TabIndex        =   187
            Text            =   "Text2"
            Top             =   930
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
            TabIndex        =   26
            Tag             =   "Concepto modelo 303|N|S|0||parametros|conce303|||"
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
            TabIndex        =   186
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
            TabIndex        =   181
            Text            =   "Text4"
            Top             =   2040
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
            TabIndex        =   29
            Tag             =   "Cuenta HP Acreedora|T|S|0||parametros|ctahpacreedor|||"
            Top             =   2040
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
            TabIndex        =   179
            Text            =   "Text4"
            Top             =   1470
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
            TabIndex        =   28
            Tag             =   "Cuenta HP Deudora|T|S|0||parametros|ctahpdeudor|||"
            Top             =   1470
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Multisecci�n"
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
            TabIndex        =   38
            Tag             =   "es multisecci�nl|N|N|||parametros|esmultiseccion|||"
            Top             =   1650
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
            TabIndex        =   37
            Tag             =   "Periodo mensual|N|N|||parametros|Presentacion349Mensual|||"
            Top             =   780
            Width           =   3045
         End
         Begin VB.CheckBox Check1 
            Caption         =   "303. Presentaci�n mensual"
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
            Left            =   7410
            TabIndex        =   35
            Tag             =   "Periodo mensual|N|N|||parametros|periodos|||"
            Top             =   390
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
            TabIndex        =   36
            Tag             =   "Mod apuntes factura|N|N|||parametros|modhcofa|||"
            Top             =   1200
            Width           =   2910
         End
         Begin VB.Image imgDiario 
            Height          =   240
            Index           =   3
            Left            =   1590
            Top             =   990
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
            TabIndex        =   188
            Top             =   960
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
            Left            =   150
            TabIndex        =   182
            Top             =   2040
            Width           =   1410
         End
         Begin VB.Image imgCta 
            Height          =   240
            Index           =   2
            Left            =   1590
            Top             =   2070
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
            TabIndex        =   180
            Top             =   1500
            Width           =   1230
         End
         Begin VB.Image imgCta 
            Height          =   240
            Index           =   1
            Left            =   1590
            Top             =   1500
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
            TabIndex        =   77
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
            TabIndex        =   76
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1035
         Left            =   240
         TabIndex        =   72
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
            Tag             =   "Cuentas p�rdidas y ganancias|T|S|0||parametros|ctaperga|||"
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
            TabIndex        =   74
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
            Caption         =   "Cuenta p�rdidas y ganancias"
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
            TabIndex        =   73
            Top             =   -30
            Width           =   2985
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4305
         Left            =   7050
         TabIndex        =   71
         Top             =   840
         Width           =   4455
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
            TabIndex        =   190
            Tag             =   "Nro Ariges|N|N|||parametros|nroariges|||"
            Text            =   "2"
            Top             =   3030
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
            Top             =   2010
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
            Top             =   1470
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
            Top             =   900
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
            Top             =   450
            Width           =   3075
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Contabilizar factura autom�ticamente(oculto)"
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
            Left            =   120
            TabIndex        =   8
            Tag             =   "Asiento automatico|N|N|||parametros|ContabilizaFact|||"
            Top             =   3540
            Visible         =   0   'False
            Width           =   4140
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
            TabIndex        =   191
            Top             =   3060
            Width           =   2445
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Anal�tica"
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
         TabIndex        =   67
         Top             =   2640
         Width           =   6765
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   735
            Left            =   3810
            TabIndex        =   99
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
            TabIndex        =   97
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
            Caption         =   "Grabar C.C. en la contabilizaci�n de facturas"
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
            Caption         =   "Contabilidad anal�tica"
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
            Caption         =   "Subgrupo a 3 d�gitos"
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
            TabIndex        =   98
            Top             =   1890
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Subgrupo a 3 d�gitos"
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
            TabIndex        =   96
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
         Height          =   1935
         Left            =   -74820
         TabIndex        =   56
         Top             =   1110
         Width           =   10995
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
            Left            =   8580
            TabIndex        =   19
            Tag             =   "Abonos negativos|N|N|||parametros|abononeg|||"
            Top             =   1380
            Width           =   2145
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
            ItemData        =   "frmparametros.frx":01A9
            Left            =   4950
            List            =   "frmparametros.frx":01B6
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Tag             =   "Ampliaci�n clientes|T|S|||parametros|nctafact|||"
            Top             =   1320
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            Caption         =   "Ampliaci�n clientes / Proveedores"
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
            Left            =   4950
            TabIndex        =   65
            Top             =   1080
            Width           =   3795
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   60
            Top             =   300
            Width           =   555
         End
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
         TabIndex        =   185
         Top             =   2010
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "�ltimo per�odo"
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
         TabIndex        =   184
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
         TabIndex        =   100
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
         TabIndex        =   64
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
         TabIndex        =   63
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
      TabIndex        =   46
      Top             =   6270
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   11310
      TabIndex        =   105
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


Dim RS As ADODB.Recordset
Dim Modo As Byte
Dim I As Integer

Dim cad As String
Dim CadB As String
Dim Indice As Integer

Private MaxLen As Integer 'Para los txt k son de ultimo nivel o de nivel anterior
                          'Ej:



Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim cad As String
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
            If InsertarDesdeForm(Me) Then
                InsertarModificar True
                PonerModo 0
            End If
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
                    cad = " fechaini = '" & Format(vParam.fechaini, FormatoFecha) & "'"
                End If
            End If
            If ModificaClaves Then
                If ModificaDesdeFormularioClaves(Me, cad) Then
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
        End If

    End Select
    
    'Si el modo es 0 significa k han insertado o modificado cosas
    If Modo = 0 Then _
        MsgBox "Para que los cambios tengan efecto debe reiniciar la aplicaci�n.", vbExclamation
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub


Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'A�adiremos el boton de aceptar y demas objetos para insertar
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

    Me.Icon = frmPpal.Icon

    SSTab1.Tab = 0
    Me.Top = 200
    Me.Left = 100
        ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    imgCta2(0).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Image2.Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    imgiva.Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    imgCta2(4).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    For I = 7 To 10
        imgCta2(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    For I = 0 To 2
        imgCta(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    For I = 0 To 3
        imgDiario(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    For I = 0 To 5
        imgConcep(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    For I = 0 To 1
        Image3(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    'Se muestra unicamente si hay tesorer�a
    Me.SSTab1.TabEnabled(4) = (vEmpresa.TieneTesoreria)
    Me.SSTab1.TabVisible(4) = (vEmpresa.TieneTesoreria)
    Me.SSTab1.TabEnabled(5) = (vEmpresa.TieneTesoreria)
    Me.SSTab1.TabVisible(5) = (vEmpresa.TieneTesoreria)
    
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = "Select * from parametros "
    Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
        'No hay datos
        Limpiar Me
        PonerModo 3
    Else
        If Not vParamT Is Nothing Then FrameOpAseguradas.Visible = vParamT.TieneOperacionesAseguradas
    
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
    PonerLongitudCampoNivelAnterior

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
    cad = Format(vFecha, "dd/mm/yyyy")
    Select Case I
    Case 0
        Text2(0).Text = cad
    End Select
End Sub

Private Sub frmI_DatoSeleccionado(CadenaSeleccion As String)
    txtIVA(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtIVA(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    
    I = Index
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
    I = Index
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
    Dim cad As String
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
       If Not IsNumeric(Text1(Index).Text) Then Exit Sub
       SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text1(Index).Text)
       If SQL = "" Then
            SQL = "Codigo incorrecto"
            Text1(Index).Text = "-1"
        End If
       Text2(Index).Text = SQL
        '....
    Case 18
        cad = Text1(18).Text
        If CuentaCorrectaUltimoNivel(cad, SQL) Then
            Text1(18).Text = cad
            Text4.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text1(18).Text = cad
            Text4.Text = SQL
            If Modo > 2 Then Text1(18).SetFocus
        End If
    
    Case 28, 29
        cad = Trim(Text1(Index).Text)
        If cad <> "" Then
            If Not IsNumeric(cad) Then
                MsgBox cad & ": No es un campo numerico", vbExclamation
                Text1(Index).Text = ""
                Text1(Index).SetFocus
            End If
        End If
        
    Case 33
        cad = Text1(33).Text
        If CuentaCorrectaUltimoNivel(cad, SQL) Then
            Text1(33).Text = cad
            Text7.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text1(33).Text = cad
            Text7.Text = SQL
            If Modo > 2 Then Text1(33).SetFocus
        End If
        
    Case 34
        cad = Text1(34).Text
        If CuentaCorrectaUltimoNivel(cad, SQL) Then
            Text1(34).Text = cad
            Text8.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text1(34).Text = cad
            Text8.Text = SQL
            If Modo > 2 Then Text1(34).SetFocus
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
    
    cmdAceptar.Visible = Modo > 0
    cmdCancelar.Visible = Modo > 0
    
    For I = 0 To Text1.Count - 1
        If I <> 30 And I <> 32 Then Text1(I).BackColor = vbWhite
    Next I
    
    Combo1.BackColor = vbWhite
    Combo2.BackColor = vbWhite
    Combo3.BackColor = vbWhite
    
    
    'Ponemos los valores
    For I = 0 To Text1.Count - 1
        If I <> 30 And I <> 32 Then Text1(I).Locked = Valor
    Next I
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

    For I = 0 To FrameTalones.Count - 1
        FrameTalones(I).Enabled = Not Valor
    Next I
    FrameValDefecto.Enabled = Not Valor
    FrameOpAseguradas.Enabled = Not Valor
    Frame66.Enabled = Not Valor

    For I = 0 To imgDiario.Count - 1
        imgDiario(I).Enabled = Not Valor
    Next I
    For I = 0 To imgConcep.Count - 1
        imgConcep(I).Enabled = Not Valor
    Next I
    
    'Campos que solo estan habilitados para insercion
    If Not Valor Then
        Text1(0).Locked = (vUsu.Nivel >= 1)
        Text1(1).Locked = (vUsu.Nivel >= 1)
    End If
    For I = 0 To imgFec.Count - 1
        imgFec(I).Enabled = Not Text1(0).Locked
    Next I
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adodc1)
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
    
        If Adodc1.Recordset.EOF Then Exit Sub
        If PonerCamposForma(Me, Adodc1) Then
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
           
           'Cuenta de p�rdidas y ganancias
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
            MsgBox "Si selecciona la contabilidad anal�tica debe poner valor al grupo gastos y ventas.", vbExclamation
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
    
    
    'A�o natural
    
    If Modo = 1 Then
        If CDate(Text1(0).Text) >= CDate(Text1(1).Text) Then
            MsgBox "Fecha inicio mayor o igual que la fecha final de ejercicio.", vbExclamation
            B = False
        End If
    End If
    
    
    'Comprobamos que si el periodo de liquidacion de IVA es mensual
    'Comprobaremos que el valor del periodo no excede de
    If Check1(5).Value = 1 Then
        I = 12  'Mensual
    Else
        I = 4   'Trimestral
    End If
    If Text1(15).Text <> "" Then
        J = CInt(Text1(15).Text)
        If J = 0 Then
            MsgBox "El periodo de liquidaci�n no puede ser 0", vbExclamation
            Exit Function
        End If
        If J > I Then
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
            If Me.Check5(4).Value = 0 Xor Text5(8).Text = "" Then C = C & "-  pagar�s "
            
            
            J = Len(Text5(9).Text) + Len(Text5(10).Text)
            If Me.Check5(5).Value = 0 Xor J = 0 Then C = C & "-  efectos "
    '        'Proveedores
            
            If C = "" Then
                'Todo bien. Compruebo esto tb
                'Si pone cta puente, la primera de las ctas es obligada
                If Me.Check5(5).Value = 1 And Text5(9).Text = "" Then C = "   La cuenta puente es campo obligatorio"
            End If
            
            If C <> "" Then
                C = Mid(C, 2) 'quitamos el primer guion
                C = C & vbCrLf & "Si marca que utiliza la cuenta puente entonces debe indicarla. En otro caso debe dejarla a blancos"
                MsgBox C, vbExclamation
                Exit Function
            End If
            
            C = ""
            'AHora es desde i=5
            For I = 7 To 9
                Text5(I).Text = Trim(Text5(I).Text)
                
                If Text5(I).Text <> "" Then
                    
                    If Len(Text5(I).Text) <> vEmpresa.DigitosUltimoNivel Then
                                                'Aqui tenemos los digitos a ultnivel-1
                        If Len(Text5(I).Text) <> MaxLen Then
                            C = C & RecuperaValor(Text5(I).Tag, 1) & vbCrLf
                        End If
                    End If
                End If
            Next I
            
            If C <> "" Then
                C = "Error en la longitud de las cuentas para: " & vbCrLf & C
                C = C & vbCrLf & "Ha de tener longitud "
                C = C & "a " & vEmpresa.DigitosUltimoNivel & " digitos o "
                C = C & "a " & MaxLen & " digitos. "
                
                MsgBox C, vbExclamation
                Exit Function
            End If
            
            DatosOK = False
            
            'Si es nuevo
            If Modo = 1 Then Text5(3).Text = 1
            
        End If
    End If
    
    DatosOK = B

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
    If vEmpresa.TieneTesoreria Then vParamT.Leer
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim RS As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(RS!Modificar, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub

'#####################################
'######### INMOVILIZADO


Private Function InsertarModificar(Insertar As Boolean) As Boolean
On Error GoTo EInsertarModificar
    InsertarModificar = False

    If Not Insertar Then
        'Modificar
        cad = "UPDATE paramamort SET tipoamor= " & Combo3.ListIndex + 1
        cad = cad & ", intcont= " & Check1(16).Value
        cad = cad & ", ultfecha= '" & Format(Text2(0).Text, FormatoFecha)
        cad = cad & "', condebes= " & DBSet(Text2(2), "N", "S")
        cad = cad & ", conhaber= " & DBSet(Text2(3), "N", "S")
        cad = cad & ", numdiari=  " & DBSet(Text2(1), "N", "S")
        
        cad = cad & ", codiva = " & DBSet(txtIVA(0), "N", "S")
        cad = cad & ", preimpreso =" & Check2.Value
        cad = cad & " WHERE codigo=1"
    Else
        'INSERTAR
        cad = "INSERT INTO paramamort (codigo, tipoamor, intcont, ultfecha, condebes, conhaber, numdiari,codiva,preimpreso) VALUES (1,"
        cad = cad & Combo3.ListIndex + 1 & "," & Me.Check1(16).Value & ",'" & Format(Text2(0).Text, FormatoFecha)
        cad = cad & "'," & DBSet(Text2(2), "N", "S") & "," & DBSet(Text2(3), "N", "S") & "," & DBSet(Text2(1), "N", "S")
        cad = cad & ",'" & DBSet(txtIVA(0), "N", "S") & "'," & Check2.Value & ")"
    End If
    Conn.Execute cad
    InsertarModificar = True
    Exit Function
EInsertarModificar:
    MuestraError Err.Number, "Insertar-Modificar"
End Function


Private Function CargarDatos() As Boolean
Dim cad As String
Dim RS As ADODB.Recordset

On Error GoTo ECargarDatos
    CargarDatos = False
    Set RS = New ADODB.Recordset
    cad = "Select * from paramamort where codigo=1"
    RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        CargarDatos = True
        '------------------  Ponemos los datos
        Combo3.ListIndex = RS!tipoamor - 1
        Check1(16).Value = RS!intcont
        Text2(0).Text = Format(RS!ultfecha, "dd/mm/yyyy")
        Text2(1).Text = DBLet(RS!NumDiari)
        Text2_LostFocus 1
        Text2(2).Text = DBLet(RS!condebes)
        Text2_LostFocus 2
        Text2(3).Text = DBLet(RS!conhaber)
        Text2_LostFocus 3
        txtIVA(0).Text = DBLet(RS!codiva)
        txtIVA_LostFocus 0
        Check2.Value = RS!Preimpreso
    End If
    RS.Close
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando parametros"
    Set RS = Nothing
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
    
    I = 0
    cad = ""
    If Not IsNumeric(txtIVA(0).Text) Then
        MsgBox "Tipo IVA debe ser num�rico: " & txtIVA(0).Text, vbExclamation
        txtIVA(0).Text = ""
        I = 1
    End If
    If I = 0 Then
        cad = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", txtIVA(0).Text, "N")
        If cad = "" Then
            MsgBox "IVA no encontrado: " & txtIVA(0).Text, vbExclamation
            txtIVA(0).Text = ""
        End If
    End If
    txtIVA(1).Text = cad
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
            MsgBox "El campo tiene que ser num�rico", vbExclamation
            Text2(Index).Text = ""
            Text2(Index).SetFocus
            Exit Sub
        End If
        Select Case Index
        Case 1
             cad = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text2(1).Text, "N")
             If cad = "" Then
                    MsgBox "Diario no encontrado: " & Text2(1).Text, vbExclamation
                    Text2(1).Text = ""
                    Text2(1).SetFocus
            End If
            Text3(0).Text = cad
        Case 2, 3
                cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text2(Index).Text, "N")
                If cad = "" Then
                    MsgBox "Concepto NO encontrado: " & Text2(Index).Text, vbExclamation
                    Text2(Index).Text = ""
                End If
                Text3(Index - 1).Text = cad
                
        End Select
    End If
End Sub

Private Sub frmC1_DatoSeleccionado(CadenaSeleccion As String)
    Text2(I + 2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(I + 1).Text = RecuperaValor(CadenaSeleccion, 2)
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
        cad = "UPDATE paramtesor SET ctabenbanc= " & DBSet(Text5(0).Text, "T", "S")
        cad = cad & ", par_pen_apli= " & DBSet(Text5(4).Text, "T", "S")
        cad = cad & ", responsable= " & DBSet(Text5(2).Text, "T", "S")
        cad = cad & ", interesescobrostarjeta= " & DBSet(Text5(16), "N")
        cad = cad & ", remesacancelacion= " & DBSet(Text5(9), "T", "S")
        cad = cad & ", ctaefectcomerciales=  " & DBSet(Text5(10), "T", "S")
        cad = cad & ", taloncta = " & DBSet(Text5(7), "T", "S")
        cad = cad & ", pagarecta =" & DBSet(Text5(8).Text, "T", "S")
        cad = cad & ", contapag =" & DBSet(Check5(0).Value, "N")
        cad = cad & ", generactrpar =" & DBSet(Check5(1).Value, "N")
        cad = cad & ", abonocambiado =" & DBSet(Check5(2).Value, "N")
        cad = cad & ", comprobarinicio =" & DBSet(Check5(8).Value, "N")
        cad = cad & ", nor19xvto =" & DBSet(Check5(10).Value, "N")
        cad = cad & ", eliminarecibidosriesgo =" & DBSet(Check5(9).Value, "N")
        cad = cad & ", diasmaxavisodesde =" & DBSet(Text5(11).Text, "N", "S")
        cad = cad & ", diasmaxavisohasta =" & DBSet(Text5(12).Text, "N", "S")
        cad = cad & ", diasmaxsiniestrodesde =" & DBSet(Text5(13).Text, "N", "S")
        cad = cad & ", diasmaxsiniestrohasta =" & DBSet(Text5(14).Text, "N", "S")
        cad = cad & ", diasavisodesdeprorroga =" & DBSet(Text5(15).Text, "N", "S")
        cad = cad & ", fechaasegesfra =" & DBSet(Check5(11).Value, "N")
        cad = cad & ", contatalonpte =" & DBSet(Check5(3).Value, "N")
        cad = cad & ", contapagarepte =" & DBSet(Check5(4).Value, "N")
        cad = cad & ", contaefecpte =" & DBSet(Check5(5).Value, "N")
        
        cad = cad & " WHERE codigo=1"
    Else
        'INSERTAR
        cad = "INSERT INTO paramtesor (codigo, ctabenbanc, par_pen_apli, responsable, interesescobrostarjeta, remesacancelacion, ctaefectcomerciales, "
        cad = cad & "taloncta,pagarecta,contapag,generactrpar,abonocambiado,comprobarinicio,nor19xvto,eliminarecibidosriesgo, "
        cad = cad & "diasmaxavisodesde,diasmaxavisohasta,diasmaxsiniestrodesde,diasmaxsiniestrohasta,diasavisodesdeprorroga,fechaasegesfra,contatalonpte, "
        cad = cad & "contapagarepte,contaefecpte) values (1,"
        cad = cad & DBSet(Text5(0).Text, "T", "S") & "," & DBSet(Text5(4).Text, "T", "S") & "," & DBSet(Text5(2).Text, "T", "S") & ","
        cad = cad & DBSet(Text5(16), "N") & "," & DBSet(Text5(9), "T", "S") & "," & DBSet(Text5(10), "T", "S") & "," & DBSet(Text5(7), "T", "S") & ","
        cad = cad & DBSet(Text5(8).Text, "T", "S") & "," & DBSet(Check5(0).Value, "N") & ","
        cad = cad & DBSet(Check5(1).Value, "N") & "," & DBSet(Check5(2).Value, "N") & "," & DBSet(Check5(8).Value, "N") & "," & DBSet(Check5(10).Value, "N") & " ,"
        cad = cad & DBSet(Check5(9).Value, "N") & "," & DBSet(Text5(11).Text, "N", "S") & "," & DBSet(Text5(12).Text, "N", "S") & "," & DBSet(Text5(13).Text, "N", "S") & ","
        cad = cad & DBSet(Text5(14).Text, "N", "S") & "," & DBSet(Text5(15).Text, "N", "S") & "," & DBSet(Check5(11).Value, "N") & "," & DBSet(Check5(3).Value, "N") & ","
        cad = cad & DBSet(Check5(4).Value, "N") & "," & DBSet(Check5(5).Value, "N") & ")"
    
    End If
    Conn.Execute cad
    InsertarModificarTesoreria = True
    Exit Function
EInsertarModificar:
    MuestraError Err.Number, "Insertar-Modificar"
End Function


Private Function CargarDatosTesoreria() As Boolean
Dim cad As String
Dim RS As ADODB.Recordset

On Error GoTo ECargarDatos
    CargarDatosTesoreria = False
    Set RS = New ADODB.Recordset
    cad = "Select * from paramtesor where codigo=1"
    RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        CargarDatosTesoreria = True
        '------------------  Ponemos los datos
        Text5(0).Text = RS!ctabenbanc
        Text5_LostFocus 0
        Text5(4).Text = RS!par_pen_apli
        Text5_LostFocus 4
        Text5(2).Text = DBLet(RS!responsable, "T")
        Text5(16).Text = RS!InteresesCobrosTarjeta
        Text5(9).Text = DBLet(RS!remesacancelacion, "T")
        Text5_LostFocus 9
        Text5(10).Text = DBLet(RS!ctaefectcomerciales, "T")
        Text5_LostFocus 10
        Text5(7).Text = DBLet(RS!taloncta, "T")
        Text5_LostFocus 7
        Text5(8).Text = DBLet(RS!pagarecta, "T")
        Text5_LostFocus 8
    
        Check5(0).Value = RS!contapag
        Check5(1).Value = RS!generactrpar
        Check5(2).Value = RS!abonocambiado
        Check5(8).Value = RS!comprobarinicio
        Check5(10).Value = RS!nor19xvto
        Check5(9).Value = RS!EliminaRecibidosRiesgo
        Check5(11).Value = RS!fechaasegesfra
        
        
        Text5(11).Text = DBLet(RS!diasmaxavisodesde, "N")
        Text5(12).Text = DBLet(RS!diasmaxavisohasta, "N")
        Text5(13).Text = DBLet(RS!diasmaxsiniestrodesde, "N")
        Text5(14).Text = DBLet(RS!diasmaxsiniestrohasta, "N")
        Text5(15).Text = DBLet(RS!DiasAvisoDesdeProrroga, "N")
    
        Check5(3).Value = RS!contatalonpte
        Check5(4).Value = RS!contapagarepte
        Check5(5).Value = RS!contaefecpte
    
    End If
    RS.Close
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando parametros de Tesoreria"
    Set RS = Nothing
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
    Dim cad As String
    Dim SQL As String
    Dim Valor As Currency
    
    ''Quitamos blancos por los lados
    Text5(Index).Text = Trim(Text5(Index).Text)

    'Si queremos hacer algo ..
    Select Case Index
    Case 1
        If Text5(Index).Text = "" Then Exit Sub
        If Not EsFechaOK(Text5(Index)) Then
            MsgBox "Fecha incorrecta : " & Text5(Index).Text, vbExclamation
            Text5(Index).Text = ""
            Text5(Index).SetFocus
            Exit Sub
        End If
                        
    Case 0, 4
        cad = Text5(Index).Text
        If cad = "" Then
            Text6(Index).Text = ""
            Exit Sub
        End If
        If CuentaCorrectaUltimoNivel(cad, SQL) Then
            Text5(Index).Text = cad
            Text6(Index).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text5(Index).Text = cad
            Text6(Index).Text = SQL
            If Modo > 2 Then Text5(Index).SetFocus
        End If

    Case 5, 6, 7, 8, 9, 10
        Text6(Index).Text = ""
        If Text5(Index).Text = "" Then Exit Sub
        If Not IsNumeric(Text5(Index).Text) Then
            MsgBox "Campo debe ser num�rico", vbExclamation
            Text5(Index).Text = ""
            Exit Sub
        End If
        I = Len(Text5(Index).Text)
        NumRegElim = InStr(1, Text5(Index).Text, ".")
        If NumRegElim = 0 Then
            If I <> vEmpresa.DigitosUltimoNivel And I <> MaxLen Then
                MsgBox "Longitud de campo incorrecta. Digitos: " & vEmpresa.DigitosUltimoNivel & " o " & MaxLen, vbExclamation
                Text5(Index).Text = ""
                Exit Sub
            End If
        End If
        
        'Llegados aqui, si es de ultimo nivel pondre la cuenta
        If NumRegElim > 0 Or I = vEmpresa.DigitosUltimoNivel Then
            cad = Text5(Index).Text
            If CuentaCorrectaUltimoNivel(cad, SQL) Then
                Text5(Index).Text = cad
            Else
                MsgBox SQL, vbExclamation
                Text5(Index).Text = ""
                SQL = ""
            End If
        Else
            SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text5(Index).Text, "T")
        End If
        Text6(Index).Text = SQL
        
    Case 11, 12, 13, 14, 16
        If Text5(Index).Text = "" Then Exit Sub
        
        If Not IsNumeric(Text5(Index).Text) Then
            MsgBox "Campo debe ser num�rico", vbExclamation
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
    
    
    I = DigitosNivel(vEmpresa.numnivel - 1)
    If I = 0 Then I = 4
    MaxLen = I
    
  
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
        C2 = "cuando dentro del punto ""Recepci�n de documentos"" se realice la contabilizaci�n"
        Select Case Index
        Case 0, 5
            C = "talones"
        Case 2, 6
            C = "pagar�s"
        Case Else
            C = "Efectos"
            C2 = "cuando dentro del punto ""Cancelaci�n cliente"" del apartado Remesas  se realice el abono de la remesa,"
        End Select
        'If Index > 0 Then C = C & " de PROVEEDORES"
        C = "Para la cancelacion de los " & C & ":" & vbCrLf & vbCrLf
        C = C & "Si tiene marcada la opcion de 'Contabiliza contra cuentas puente', " & C2
        C = C & " tendremos dos opciones:" & vbCrLf
        If Index = 3 Then C = C & "  Efectos descontados" & vbCrLf
        C = C & "    -   Una �nica cuenta a �ltimo nivel (Ej: 4310000), con lo que todos los apuntes ir�n a esa cuenta gen�rica." & vbCrLf
        C = C & "    -   Introducir una cuenta ra�z a 4 d�gitos (Ej: 4310), con lo que el programa crear� cuentas a �ltimo nivel haci�ndolas coincidir con las terminaciones de las cuentas del cliente." & vbCrLf
        
        
        'Nuevo Nov 2009
        If Index = 3 Then
            C = C & vbCrLf & vbCrLf
            C = C & "    Efectos descontados a cobrar " & vbCrLf
            C = C & "    -   Una �nica cuenta a �ltimo nivel (Ej: 4310000), con lo que todos los apuntes ir�n a esa cuenta gen�rica." & vbCrLf
            C = C & "    -   Introducir una cuenta ra�z a 4 d�gitos (Ej: 4310), con lo que el programa crear� cuentas a �ltimo nivel haci�ndolas coincidir con las terminaciones de las cuentas del cliente." & vbCrLf
        End If
        
    Case 1
        C = "Cuenta beneficios bancarios." & vbCrLf & vbCrLf
        C = C & "Si no esta indicada en la configuraci�n del banco  " & vbCrLf
        C = C & "con el que estemos trabajando, utilizar� esta cuenta  " & vbCrLf
    Case 4
        C = "Valores que ofertar� para la contabilizaci�n de cobros/pagos. " & vbCrLf
        C = C & "Luego podr� ser modificado para cada caso  " & vbCrLf
    Case 7
        C = "Responsable para poder firmar en documentos(recibos, cheques)"
        
    Case 8
        C = "Al entrar en la empresa que compruebe si hay posibilidad de eliminar "
        C = C & vbCrLf & "riesgo, tanto en efectos como en talones y pagar�s"
    Case 9
        C = "Cuando eliminamos riesgo en talones y pagar�s, eliminar tambien en la tabla de  "
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
        C = C & vbCrLf & vbCrLf & "Una vez impresos los recibos, si la impresi�n es correcta, graba columna gastos "

        '
    
    End Select
    MsgBox C, vbInformation
End Sub


