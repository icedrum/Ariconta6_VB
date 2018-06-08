VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESTransferencias 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencias"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16035
   Icon            =   "frmTESTransferencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   16035
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6930
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCreacionRemesa 
      BorderStyle     =   0  'None
      Height          =   9045
      Left            =   30
      TabIndex        =   19
      Top             =   -60
      Visible         =   0   'False
      Width           =   15855
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
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
         Left            =   12060
         TabIndex        =   57
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   4515
         Left            =   150
         TabIndex        =   26
         Top             =   3840
         Width           =   11655
         Begin MSComctlLib.ListView lwCobros 
            Height          =   4095
            Left            =   0
            TabIndex        =   27
            Top             =   360
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   7223
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
            NumItems        =   8
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
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ENTIDAD"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   11310
            Picture         =   "frmTESTransferencias.frx":000C
            ToolTipText     =   "Puntear al Debe"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   10950
            Picture         =   "frmTESTransferencias.frx":0156
            ToolTipText     =   "Quitar al Debe"
            Top             =   30
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   180
         TabIndex        =   24
         Top             =   8340
         Width           =   1755
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
            TabIndex        =   25
            Top             =   210
            Width           =   1200
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
         Index           =   0
         Left            =   13170
         TabIndex        =   17
         Top             =   8460
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   0
         Left            =   14430
         TabIndex        =   18
         Top             =   8460
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   7950
         Top             =   150
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.Frame FrameCreaRem 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         TabIndex        =   29
         Top             =   60
         Width           =   15645
         Begin VB.CheckBox chkAbonosSoloTransferencia 
            Caption         =   "Sólo tipo pago transferencia"
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
            Left            =   12270
            TabIndex        =   77
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   2400
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Pago en Fecha introducida"
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
            Left            =   12270
            TabIndex        =   75
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   1980
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.ComboBox cboConcepto 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmTESTransferencias.frx":02A0
            Left            =   12210
            List            =   "frmTESTransferencias.frx":02AD
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   780
            Width           =   2265
         End
         Begin VB.TextBox txtCuentas 
            BeginProperty Font 
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
            TabIndex        =   14
            Text            =   "Text2"
            Top             =   3150
            Width           =   1335
         End
         Begin VB.TextBox txtNCuentas 
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
            Left            =   1620
            TabIndex        =   55
            Text            =   "Text2"
            Top             =   3150
            Width           =   3525
         End
         Begin VB.TextBox txtRemesa 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6630
            MaxLength       =   50
            TabIndex        =   16
            Tag             =   "Descripción|T|N|||remesas|descripción|||"
            Top             =   3150
            Width           =   5145
         End
         Begin VB.TextBox txtFecha 
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
            Index           =   4
            Left            =   5250
            TabIndex        =   15
            Tag             =   "Fecha Reclamación|F|N|||reclama|fecreclama|dd/mm/yyyy||"
            Text            =   "99/99/9999"
            Top             =   3150
            Width           =   1245
         End
         Begin VB.TextBox txtFecha 
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
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "imgConcepto"
            Top             =   810
            Width           =   1305
         End
         Begin VB.TextBox txtFecha 
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
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "imgConcepto"
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox txtNumFac 
            Alignment       =   1  'Right Justify
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
            Left            =   1230
            TabIndex        =   8
            Tag             =   "Nº factura|N|S|0||factcli|numfactu|0000000|S|"
            Top             =   1950
            Width           =   1275
         End
         Begin VB.TextBox txtNumFac 
            Alignment       =   1  'Right Justify
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
            Left            =   1230
            TabIndex        =   9
            Tag             =   "Nº factura|N|S|0||factcli|numfactu|0000000|S|"
            Top             =   2370
            Width           =   1275
         End
         Begin VB.TextBox txtSerie 
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
            Left            =   6300
            TabIndex        =   7
            Tag             =   "imgConcepto"
            Top             =   1200
            Width           =   765
         End
         Begin VB.TextBox txtSerie 
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
            Left            =   6300
            TabIndex        =   6
            Tag             =   "imgConcepto"
            Top             =   810
            Width           =   765
         End
         Begin VB.TextBox txtCuentas 
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
            Left            =   6240
            TabIndex        =   12
            Tag             =   "imgConcepto"
            Top             =   1950
            Width           =   1275
         End
         Begin VB.TextBox txtCuentas 
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
            Left            =   6240
            TabIndex        =   13
            Tag             =   "imgConcepto"
            Top             =   2370
            Width           =   1275
         End
         Begin VB.TextBox txtNSerie 
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
            Left            =   7110
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   810
            Width           =   4665
         End
         Begin VB.TextBox txtNSerie 
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
            Left            =   7110
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1200
            Width           =   4665
         End
         Begin VB.TextBox txtNCuentas 
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
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   1950
            Width           =   4185
         End
         Begin VB.TextBox txtNCuentas 
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
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   2370
            Width           =   4185
         End
         Begin VB.TextBox txtFecha 
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
            Index           =   3
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "imgConcepto"
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox txtFecha 
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
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "imgConcepto"
            Top             =   810
            Width           =   1305
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   3660
            TabIndex        =   11
            Tag             =   "imgConcepto"
            Top             =   2370
            Width           =   1275
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   3660
            TabIndex        =   10
            Tag             =   "imgConcepto"
            Top             =   1950
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Concepto"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Index           =   13
            Left            =   12180
            TabIndex        =   73
            Top             =   480
            Visible         =   0   'False
            Width           =   1170
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
            Index           =   1
            Left            =   240
            TabIndex        =   56
            Top             =   2880
            Width           =   975
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   2
            Left            =   1320
            Top             =   2880
            Width           =   240
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
            Left            =   6630
            TabIndex        =   54
            Top             =   2880
            Width           =   1245
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   6210
            Top             =   2880
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
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
            Left            =   5250
            TabIndex        =   53
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label Label3 
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
            Index           =   0
            Left            =   5280
            TabIndex        =   52
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label3 
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
            Index           =   1
            Left            =   5280
            TabIndex        =   51
            Top             =   1230
            Width           =   585
         End
         Begin VB.Image imgSerie 
            Height          =   255
            Index           =   0
            Left            =   5940
            Top             =   840
            Width           =   255
         End
         Begin VB.Image imgSerie 
            Height          =   255
            Index           =   1
            Left            =   5940
            Top             =   1230
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Serie"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   5250
            TabIndex        =   50
            Top             =   510
            Width           =   960
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Factura"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   2700
            TabIndex        =   49
            Top             =   480
            Width           =   2280
         End
         Begin VB.Label Label3 
            Caption         =   "Nro.Factura"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   48
            Top             =   1650
            Width           =   1590
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   2700
            TabIndex        =   47
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lblFecha1 
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
            Left            =   2580
            TabIndex        =   46
            Top             =   3990
            Width           =   4095
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   2700
            TabIndex        =   45
            Top             =   1260
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   44
            Top             =   2010
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   43
            Top             =   2430
            Width           =   615
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   3390
            Top             =   855
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   3390
            Top             =   1260
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   255
            Index           =   0
            Left            =   5940
            Top             =   1980
            Width           =   255
         End
         Begin VB.Image imgCuentas 
            Height          =   255
            Index           =   1
            Left            =   5940
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   5250
            TabIndex        =   42
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   5250
            TabIndex        =   41
            Top             =   2010
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Cuenta Cliente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   5250
            TabIndex        =   40
            Top             =   1650
            Width           =   2910
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   960
            Top             =   1230
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   960
            Top             =   855
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   270
            TabIndex        =   39
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   270
            TabIndex        =   38
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Vencimiento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   18
            Left            =   270
            TabIndex        =   37
            Top             =   480
            Width           =   2280
         End
         Begin VB.Label Label3 
            Caption         =   "Importe Vencimiento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   14
            Left            =   2670
            TabIndex        =   36
            Top             =   1650
            Width           =   2340
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   2670
            TabIndex        =   35
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   2670
            TabIndex        =   34
            Top             =   2010
            Width           =   690
         End
      End
      Begin VB.Frame FrameModRem 
         Caption         =   "Datos Transferencia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         TabIndex        =   59
         Top             =   60
         Width           =   15645
         Begin VB.CheckBox chkFecha2 
            Caption         =   "Pago en Fecha introducida"
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
            Left            =   11820
            TabIndex        =   76
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   2130
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.ComboBox cboConcepto2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmTESTransferencias.frx":02CD
            Left            =   11790
            List            =   "frmTESTransferencias.frx":02DA
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   2100
            Width           =   2265
         End
         Begin VB.TextBox txtFecha 
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
            Left            =   5370
            TabIndex        =   61
            Tag             =   "Fecha Reclamación|F|N|||reclama|fecreclama|dd/mm/yyyy||"
            Text            =   "99/99/9999"
            Top             =   2130
            Width           =   1245
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
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   62
            Tag             =   "Descripción|T|N|||remesas|descripción|||"
            Top             =   2130
            Width           =   5025
         End
         Begin VB.TextBox txtNCuentas 
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
            Index           =   3
            Left            =   1740
            TabIndex        =   64
            Text            =   "Text2"
            Top             =   2130
            Width           =   3525
         End
         Begin VB.TextBox txtCuentas 
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   60
            Text            =   "Text2"
            Top             =   2130
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Concepto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Index           =   15
            Left            =   11760
            TabIndex        =   74
            Top             =   1800
            Width           =   1170
         End
         Begin VB.Label Label3 
            Caption         =   "Transferencia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   12
            Left            =   360
            TabIndex        =   69
            Top             =   750
            Width           =   8940
         End
         Begin VB.Label lblFecha1 
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
            Left            =   2580
            TabIndex        =   68
            Top             =   3990
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
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
            Left            =   5370
            TabIndex        =   67
            Top             =   1860
            Width           =   795
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   6330
            Top             =   1860
            Width           =   240
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
            Index           =   3
            Left            =   6690
            TabIndex        =   66
            Top             =   1860
            Width           =   1245
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   3
            Left            =   2340
            Top             =   1860
            Width           =   240
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
            Left            =   360
            TabIndex        =   65
            Top             =   1860
            Width           =   1845
         End
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   2130
         TabIndex        =   70
         Top             =   8550
         Width           =   8400
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   72
         Left            =   12090
         TabIndex        =   58
         Top             =   3900
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   30
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   15915
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   4020
         TabIndex        =   28
         Top             =   180
         Width           =   1365
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   180
            TabIndex        =   71
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Grabación Fichero"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Cargo Transferencia"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   21
         Top             =   180
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   180
            TabIndex        =   1
            Top             =   210
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
      Begin MSComctlLib.ListView lw1 
         Height          =   7305
         Left            =   240
         TabIndex        =   22
         Top             =   990
         Width           =   15525
         _ExtentX        =   27384
         _ExtentY        =   12885
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
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   15270
         TabIndex        =   23
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
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   1
         Left            =   14670
         TabIndex        =   0
         Top             =   7860
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTESTransferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)



Private Const SaltoLinea = """ + chr(13) + """

' en este formulario el IdPrograma no es una constante pq sirve para transf. de abonos (cobros) y de pagos
Private IdPrograma As Integer  ' transferencias de abonos = 614
                               ' transferencias de pagos = 805


Public TipoTrans As Integer  '0 = transferencia desde abonos
                             '1 = transferencia desde pagos
                             '2 = transferencia de pagos domiciliados o bancarios
                             '3 = transferencia de confirming

Public vSql As String
Public Opcion As Byte      ' 0.- Nueva remesa    1.- Modifcar remesa
                           ' 2.- Devolucion remesa
Public vRemesa As String   ' nºremesa|fecha remesa
Public ImporteRemesa As Currency

Public ValoresDevolucionRemesa As String
        'NOV 2009
        'antes: 4 campos     AHORA 5 campos
        'Concepto|ampliacion|
        'Concepto banco|ampliacion banco|
        'ahora+ Agrupa vtos

Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1

Private frmMens3 As frmMensajes
Private frmMens2 As frmMensajes
Attribute frmMens2.VB_VarHelpID = -1
Private frmMens As frmMensajes

Dim SQL As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim i As Integer
Dim TotalReg As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Dim CampoOrden As String
Dim Orden As Boolean
Dim Modo As Byte

Dim Txt33Csb As String
Dim Txt41Csb As String

Dim Indice As Integer
Dim Codigo As Long

Dim SubTipo As Integer

Dim ModoInsertar As Boolean

Dim IndCodigo As Integer
Dim nTipo As String

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub


Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function

Private Sub cmdCancelar_Click(Index As Integer)
Dim i As Integer
Dim PantallaPpal As Boolean
    If Index = 0 Then
        
        If Modo = 3 Then
            If Not ModoInsertar Then
                PantallaPpal = True
            Else
                PantallaPpal = False
            End If
        Else
            PantallaPpal = True
        End If
        
        If PantallaPpal Then
            'Pidiendo datos
            If ModoInsertar Then
                cmdAceptar(0).Caption = "&Aceptar"
                ModoInsertar = False
            End If
        
            Frame1.visible = True
            Frame1.Enabled = True
            
            FrameCreacionRemesa.visible = False
            FrameCreacionRemesa.Enabled = False
            If i >= 0 Then lw1.SetFocus
    
        Else
            Me.lwCobros.ListItems.Clear
            cmdAceptar(0).Caption = "&Aceptar"
            ModoInsertar = False
            Me.Text1(4).Text = ""
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
    Select Case Index
        Case 0
            Select Case Modo
                Case 3  ' insertar
                    If Not DatosOK(0) Then Exit Sub
                
                    If Not ModoInsertar Then
                        ModoInsertar = True
                        cmdAceptar(0).Caption = "C&onfirmar"
                        If TipoTrans = 0 Then ' transferencia de abonos
                            NuevaTransf
                        Else
                            NuevaTransfPagos
                        End If
                    Else
                        If GenerarTransferencia(0) Then
                            Select Case TipoTrans
                                Case 0, 1
                                    MsgBox "Transferencia generada correctamente.", vbExclamation
                                Case 2
                                    MsgBox "Pago domiciliado generado correctamente.", vbExclamation
                                Case 3
                                    MsgBox "Confirming generado correctamente.", vbExclamation
                            End Select
                            ModoInsertar = False
                            cmdCancelar_Click (0)
                            CargaList
                        End If
                    End If
                    
                    Screen.MousePointer = vbDefault
                    
                    
                Case 4  ' modificar
                    If Not DatosOK(1) Then Exit Sub
                    
                    If Not ModoInsertar Then
                        ModoInsertar = True
                        
                        cmdAceptar(0).Caption = "C&onfirmar"
                    Else
                        If GenerarTransferencia(1) Then
                            'Refrescamos los datos en el lw de remesas
                            'MsgBox "Remesa modificada correctamente.", vbExclamation
                            cmdCancelar_Click (0)
                            CargaList
                        End If
                        
                    End If
            End Select
    End Select
End Sub


Private Function DatosOK(Opcion As Integer) As Boolean
Dim B As Boolean
Dim nTipo As String
Dim C As String
Dim J As Integer

    DatosOK = False

    If Opcion = 0 Then
        If txtCuentas(2).Text = "" Then
            MsgBox "Indique la cuenta bancaria", vbExclamation
            Exit Function
        Else
            If TipoTrans < 2 Then
                SQL = "select count(*) from bancos where codmacta = " & DBSet(txtCuentas(2).Text, "T") & " and not sufijoem is null and sufijoem <> ''"
                If TotalRegistros(SQL) = 0 Then
                    MsgBox "El banco no tiene Sufijo Transferencia. Reintroduzca.", vbExclamation
                    PonleFoco txtCuentas(2)
                    Exit Function
                End If
            End If
        End If
    
        'Fecha transferencia tiene k tener valor
        If txtfecha(4).Text = "" Then
            MsgBox "Fecha de " & nTipo & " debe tener valor", vbExclamation
            PonFoco txtfecha(4)
            Exit Function
        End If
        
        'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
        If FechaCorrecta2(CDate(txtfecha(4).Text), True) > 1 Then
            PonFoco txtfecha(4)
            Exit Function
        End If
        
        
        If TipoTrans < 2 Then
            If Me.cboConcepto.ListIndex = -1 Then
                MsgBox "Debe introducir un Concepto.", vbExclamation
                PonerFocoCmb cboConcepto
                Exit Function
            End If
        End If
    Else
        If txtCuentas(3).Text = "" Then
            MsgBox "Indique la cuenta bancaria", vbExclamation
            Exit Function
        End If
    
        'Fecha remesa tiene k tener valor
        If txtfecha(5).Text = "" Then
            MsgBox "Fecha de " & nTipo & " debe tener valor", vbExclamation
            PonFoco txtfecha(5)
            Exit Function
        Else
            If Year(CDate(txtfecha(5).Text)) <> lw1.SelectedItem.SubItems(1) Then
                MsgBox "La fecha de " & nTipo & " ha de ser del mismo año. Revise.", vbExclamation
                PonFoco txtfecha(5)
                Exit Function
            End If
        End If
        
        'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
        If FechaCorrecta2(CDate(txtfecha(5).Text), True) > 1 Then
            PonFoco txtfecha(5)
            Exit Function
        End If
    
        If TipoTrans < 2 Then
            If Me.cboConcepto2.ListIndex = -1 Then
                MsgBox "Debe introducir un Concepto.", vbExclamation
                PonerFocoCmb cboConcepto2
                Exit Function
            End If
        End If
    
    
        'Sept 2017
        'Si hay algun "NO" no sigue
        
    End If
    
        C = ""
        For J = 1 To Me.lwCobros.ListItems.Count
            If lwCobros.ListItems(J).ListSubItems(3).Tag = "NO" And lwCobros.ListItems(J).Checked Then
                C = C & lwCobros.ListItems(J).Text & " - " & lwCobros.ListItems(J).SubItems(1) & " "
                C = C & lwCobros.ListItems(J).SubItems(5) & " " & lwCobros.ListItems(J).ListSubItems(3).ToolTipText & vbCrLf
            End If
        Next
        If C <> "" Then
            C = "Vencimientos incorrectos. ¿Continuar de igualmente?" & vbCrLf & C
            
            If MsgBox(C, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
        End If
    
    DatosOK = True

End Function

Private Sub Insertar()
Dim NumF As Long
Dim B As Boolean

    On Error GoTo eInsertar
    
    Conn.BeginTrans
    
eInsertar:
    If Err.Number = 0 And B Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
End Sub

Private Function InsertarLineas() As Boolean
Dim Rs As ADODB.Recordset
Dim CadValues As String
Dim CadInsert As String

    On Error GoTo eInsertarLineas

    InsertarLineas = False

    InsertarLineas = True
    Exit Function
    
eInsertarLineas:
    MuestraError Err.Number, "Insertar Lineas", Err.Description
End Function

Private Sub cmdVtoDestino(Index As Integer)
    
    If Index = 0 Then
        TotalReg = 0
        If Not Me.lwCobros.SelectedItem Is Nothing Then TotalReg = Me.lwCobros.SelectedItem.Index
    
    
        For i = 1 To Me.lwCobros.ListItems.Count
            If Me.lwCobros.ListItems(i).Bold Then
                Me.lwCobros.ListItems(i).Bold = False
                Me.lwCobros.ListItems(i).ForeColor = vbBlack
                For CONT = 1 To Me.lwCobros.ColumnHeaders.Count - 1
                    Me.lwCobros.ListItems(i).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwCobros.ListItems(i).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalReg > 0 Then
            i = TotalReg
            Me.lwCobros.ListItems(i).Bold = True
            Me.lwCobros.ListItems(i).ForeColor = vbRed
            For CONT = 1 To Me.lwCobros.ColumnHeaders.Count - 1
                Me.lwCobros.ListItems(i).ListSubItems(CONT).ForeColor = vbRed
                Me.lwCobros.ListItems(i).ListSubItems(CONT).Bold = True
            Next
        End If
        lwCobros.Refresh
        
        PonerFocoLw Me.lwCobros

    Else
    

    End If
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not Frame1.visible Then
            If CadenaDesdeOtroForm <> "" Then
            Else
'                PonFoco Text1(2)
            End If
            CadenaDesdeOtroForm = ""
        End If
        CargaList
        PonerFocoBtn cmdCancelar(1)
        If lw1.ListItems.Count > 0 Then Set lw1.SelectedItem = Nothing

    End If
    Screen.MousePointer = vbDefault
End Sub
    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmppal.Icon
    
    For i = 0 To 1
        Me.imgSerie(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    Me.imgCuentas(2).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(3).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    For i = 0 To 5
        Me.ImgFec(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i

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
        .Buttons(2).Image = 37
    End With
    
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    Select Case TipoTrans
        Case 0
            Me.Caption = "Transferencias Abonos"
            IdPrograma = 614 ' transf. de abonos
            nTipo = "transferencia"
                        
            'concepto
            Label3(13).Enabled = True
            Label3(13).visible = True
            cboConcepto.Enabled = True
            cboConcepto.visible = True
            Label3(15).Enabled = True
            Label3(15).visible = True
            cboConcepto2.Enabled = True
            cboConcepto2.visible = True
            'fecha introducida
            chkFecha.Enabled = False
            chkFecha.visible = False
            chkFecha2.Enabled = False
            chkFecha2.visible = False
            
            Label3(11).Caption = "Cuenta Cliente"
            Label3(18).Caption = "Fecha Vencimiento"
            Label3(14).Caption = "Importe Vencimiento"
        Case 1
            Me.Caption = "Transferencias Pagos"
            IdPrograma = 805 ' transf. de pagos
            nTipo = "transferencia"
            
            'concepto
            Label3(13).Enabled = True
            Label3(13).visible = True
            cboConcepto.Enabled = True
            cboConcepto.visible = True
            Label3(15).Enabled = True
            Label3(15).visible = True
            cboConcepto2.Enabled = True
            cboConcepto2.visible = True
            'fecha introducida
            chkFecha.Enabled = False
            chkFecha.visible = False
            chkFecha2.Enabled = False
            chkFecha2.visible = False
            
            Label3(11).Caption = "Cuenta Proveedor"
            Label3(18).Caption = "Fecha Efecto"
            Label3(14).Caption = "Importe Efecto"
        
        Case 2
            Me.Caption = "Pagos Domiciliados"
            IdPrograma = 806 ' transf. de pagos domiciliados
            nTipo = "pago domiciliado"
            
            'concepto
            Label3(13).Enabled = False
            Label3(13).visible = False
            cboConcepto.Enabled = False
            cboConcepto.visible = False
            Label3(15).Enabled = False
            Label3(15).visible = False
            cboConcepto2.Enabled = False
            cboConcepto2.visible = False
            'fecha introducida
            chkFecha.Enabled = True
            chkFecha.visible = True
            chkFecha2.Enabled = True
            chkFecha2.visible = True
            
            Label3(11).Caption = "Cuenta Proveedor"
            Label3(18).Caption = "Fecha Efecto"
            Label3(14).Caption = "Importe Efecto"
        Case 3
            Me.Caption = "Confirming"
            IdPrograma = 810 ' transf. de pagos confirming caixa
            nTipo = "confirming"
            
            'concepto
            Label3(13).Enabled = False
            Label3(13).visible = False
            cboConcepto.Enabled = False
            cboConcepto.visible = False
            Label3(15).Enabled = False
            Label3(15).visible = False
            cboConcepto2.Enabled = False
            cboConcepto2.visible = False
            'fecha introducida
            chkFecha.Enabled = True
            chkFecha.visible = True
            chkFecha2.Enabled = True
            chkFecha2.visible = True
            
            Label3(11).Caption = "Cuenta Proveedor"
            Label3(18).Caption = "Fecha Efecto"
            Label3(14).Caption = "Importe Efecto"
    End Select
    Me.Toolbar2.Buttons(2).ToolTipText = "Cargo " & UCase(Me.Caption)
    If TipoTrans < 2 Then cboConcepto.ListIndex = 1
    chkAbonosSoloTransferencia.visible = IIf(TipoTrans = 0, True, False)
    'Limpiamos el tag
    PrimeraVez = True
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
        
    H = FrameCreacionRemesa.Height + 120
    W = FrameCreacionRemesa.Width
    
    FrameCreacionRemesa.visible = False
    Me.Frame1.visible = True
    
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    Me.cmdCancelar(0).Cancel = True
    
    
    Orden = True
    CampoOrden = "transferencias.fecha"
    
    
    lwCobros.ColumnHeaders.Clear
    
    lwCobros.ColumnHeaders.Add , , "Tipo", 800
    lwCobros.ColumnHeaders.Add , , "Factura", 1200
    lwCobros.ColumnHeaders.Add , , "Fecha", 1350
    
    If TipoTrans = 0 Then
        lwCobros.ColumnHeaders.Add , , "Vto", 699
        lwCobros.ColumnHeaders.Add , , "Fecha Vto", 1350
        'lwCobros.ColumnHeaders.Add , , "Forma pago", 3500
        lwCobros.ColumnHeaders.Add , , "Cliente", 3500
    Else
        lwCobros.ColumnHeaders.Add , , "Efec", 699
        lwCobros.ColumnHeaders.Add , , "Fecha Efec", 1350
        lwCobros.ColumnHeaders.Add , , "Proveedor", 3500
    End If
    
    lwCobros.ColumnHeaders.Add , , "Importe", 2035, 1
    lwCobros.ColumnHeaders.Add , , "ENTIDAD", 0, 1
    lwCobros.ColumnHeaders.Add , , "CtaProve", 0, 1

End Sub


Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCuentas(2).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNCuentas(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub


Private Sub Image3_Click(Index As Integer)

    Select Case Index
        Case 1 ' cuenta contable
            Screen.MousePointer = vbHourglass
            
            Set frmCtas = New frmColCtas
            RC = Index
            frmCtas.DatosADevolverBusqueda = "0|1"
            frmCtas.ConfigurarBalances = 3
            frmCtas.Show vbModal
            Set frmCtas = Nothing
    
    End Select
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim IT
Dim i As Integer
    For i = 1 To Me.lwCobros.ListItems.Count
        Set IT = lwCobros.ListItems(i)
        lwCobros.ListItems(i).Checked = (Index = 1)
        lwCobros_ItemCheck (IT)
        Set IT = Nothing
    Next i
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtfecha(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'FECHA FACTURA
    Indice = Index
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtfecha(Indice).Text <> "" Then frmF.Fecha = CDate(txtfecha(Indice).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtfecha(Indice)

End Sub

Private Sub imgCuentas_Click(Index As Integer)
    
    If Index = 2 Then
            Set frmBan = New frmBasico2
            AyudaBanco frmBan
            Set frmBan = Nothing
    
    
    Else
        SQL = ""
        AbiertoOtroFormEnListado = True
        Set frmCtas = New frmColCtas
        frmCtas.DatosADevolverBusqueda = True
        frmCtas.Show vbModal
        Set frmCtas = Nothing
        If SQL <> "" Then
            Me.txtCuentas(Index).Text = RecuperaValor(SQL, 1)
            Me.txtNCuentas(Index).Text = RecuperaValor(SQL, 2)
        Else
            QuitarPulsacionMas Me.txtCuentas(Index)
        End If
        
        PonFoco Me.txtCuentas(Index)
        AbiertoOtroFormEnListado = False
    End If
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2, 3, 4, 5
        Indice = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtfecha(Index).Text <> "" Then frmF.Fecha = CDate(txtfecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtfecha(Index)
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgSerie_Click(Index As Integer)
    IndCodigo = Index

    Set frmConta = New frmBasico
    AyudaContadores frmConta, txtSerie(Index), "tiporegi REGEXP '^[0-9]+$' = 0"
    Set frmConta = Nothing
    
    PonFoco Me.txtSerie(Index)
End Sub

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
    
    Select Case ColumnHeader
        Case "Código"
            CampoOrden = "transferencias.codigo"
        Case "Fecha"
            CampoOrden = "transferencias.fecha"
        Case "Cuenta"
            CampoOrden = "transferencias.codmacta"
        Case "Nombre"
            CampoOrden = "cuentas.nommacta"
        Case "Año"
            CampoOrden = "transferencias.anyo"
        Case "Importe"
            CampoOrden = "transferencias.importe"
        Case "Descripción"
            CampoOrden = "transferencias.descripcion"
        Case "Situación"
            CampoOrden = "descsituacion"
    End Select
    CargaList


End Sub

Private Sub lw1_DblClick()
    If lw1.SelectedItem Is Nothing Then Exit Sub
    'detalle de facturas
    Set frmMens = New frmMensajes
    
    If TipoTrans = 0 Then ' transferencias de abonos
        frmMens.Opcion = 55
    Else
        frmMens.Opcion = 56
    End If
    frmMens.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|"
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    
End Sub

Private Sub lw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'FALTA###  ¿Porque esta asi?
'    PonerModoUsuarioGnral 2, "ariconta"
End Sub

Private Sub lwCobros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
    C = Item.Tag
    
    Importe = 0
    For i = 1 To lwCobros.ListItems.Count
        If lwCobros.ListItems(i).Checked Then Importe = Importe + lwCobros.ListItems(i).SubItems(6)
    Next i
    Text1(4).Text = Format(Importe, "###,###,##0.00")
    
    If ComprobarCero(Text1(4).Text) = 0 Then Text1(4).Text = ""
            
End Sub

Private Sub HacerToolBar(Boton As Integer)
Dim frmTESTransList As frmTESTransferenciasList

    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
            BotonModificar
        Case 3
            BotonEliminar
        Case 5
'            BotonBuscar
        Case 6 ' ver todos
            CargaList
        Case 8
            Set frmTESTransList = New frmTESTransferenciasList
        
            'Imprimir factura
            If Not lw1.SelectedItem Is Nothing Then
                frmTESTransList.numero = lw1.SelectedItem.Text
                frmTESTransList.Anyo = lw1.SelectedItem.SubItems(1)
            Else
                frmTESTransList.numero = ""
                frmTESTransList.Anyo = ""
            End If
            
            frmTESTransList.EsTransfAbonos = (TipoTrans = 0)
            If TipoTrans > 0 Then
                frmTESTransList.TipoTransfPagos = (TipoTrans - 1)
            Else
                'si son abonos me da igual este parametro
                frmTESTransList.TipoTransfPagos = 0
            End If
                    
            frmTESTransList.Show vbModal
            
            Set frmTESTransList = Nothing

    End Select
    
End Sub

Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    
    SepuedeBorrar = False

    If lw1.SelectedItem.SubItems(9) = "Q" Then
        MsgBox "No se pueden modificar ni eliminar " & nTipo & "s en situación abonada.", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True

End Function


Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    
    If Me.lw1.SelectedItem Is Nothing Then Exit Sub
    If Me.lw1.SelectedItem = "" Then Exit Sub
        
    If Not SepuedeBorrar Then Exit Sub
        
        
    '*************** canviar els noms i el DELETE **********************************
    Select Case TipoTrans
        Case 0, 1
            SQL = "¿Seguro que desea eliminar la Transferencia?"
        Case 2
            SQL = "¿Seguro que desea eliminar el Pago Domiciliado?"
        Case 3
            SQL = "¿Seguro que desea eliminar el Confirming?"
    End Select
    SQL = SQL & vbCrLf & " Código: " & lw1.SelectedItem.Text
    SQL = SQL & vbCrLf & " Fecha: " & lw1.SelectedItem.SubItems(2)
    SQL = SQL & vbCrLf & " Banco: " & lw1.SelectedItem.SubItems(5)
    SQL = SQL & vbCrLf & " Importe: " & lw1.SelectedItem.SubItems(8)
    
    
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = lw1.SelectedItem.Text
        If TipoTrans = 0 Then
            If ModificarCobros Then
                lw1.ListItems.Remove (lw1.SelectedItem.Index)
                If lw1.ListItems.Count > 0 Then
                    lw1.SetFocus
                End If
            End If
        Else
            If ModificarPagos Then
                lw1.ListItems.Remove (lw1.SelectedItem.Index)
                If lw1.ListItems.Count > 0 Then
                    lw1.SetFocus
                End If
            End If
        End If
            
'        CargaList
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Function ModificarCobros() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim FecUltCob As String
Dim Importe As Currency
Dim NumLinea As Integer


    ModificarCobros = False
    
    Conn.BeginTrans
    

    SQL = "select * from cobros where transfer = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        ' antes lo sumaba de los cobros_realizados
        ' ahora lo dejo todo a nulo
    
    
        FecUltCob = ""
        Importe = 0
    
        SQL = "update cobros set fecultco = " & DBSet(FecUltCob, "F", "S")
        If Importe = 0 Then
            SQL = SQL & " , impcobro = " & ValorNulo
        Else
            SQL = SQL & " , impcobro = " & DBSet(Importe, "N", "S")
        End If
        SQL = SQL & ", tiporem = " & ValorNulo
        SQL = SQL & ", transfer = " & ValorNulo
        SQL = SQL & ", anyorem = " & ValorNulo
        SQL = SQL & ", siturem = " & ValorNulo
        SQL = SQL & ", situacion = 0 "
        SQL = SQL & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and "
        SQL = SQL & " numfactu = " & DBSet(Rs!NumFactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and "
        SQL = SQL & " numorden = " & DBSet(Rs!numorden, "N")
                    
        Conn.Execute SQL
    
        Rs.MoveNext
    Wend

    SQL = "delete from transferencias where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Conn.Execute SQL

    Set Rs = Nothing
    ModificarCobros = True
    Conn.CommitTrans
    Exit Function
    
eModificarCobros:
    Conn.RollbackTrans
    MuestraError Err.Number, "Modificar Cobros", Err.Description
End Function

Private Function ModificarPagos() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim FecUltCob As String
Dim Importe As Currency
Dim NumLinea As Integer


    ModificarPagos = False
    
    Conn.BeginTrans
    

    SQL = "select * from pagos where nrodocum = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyodocum = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        ' antes lo sumaba de los cobros_realizados
        ' ahora lo dejo todo a nulo
    
    
        FecUltCob = ""
        Importe = 0
    
        SQL = "update pagos set fecultpa = " & DBSet(FecUltCob, "F", "S")
        If Importe = 0 Then
            SQL = SQL & " , imppagad = " & ValorNulo
        Else
            SQL = SQL & " , imppagad = " & DBSet(Importe, "N", "S")
        End If
        SQL = SQL & ", nrodocum = " & ValorNulo
        SQL = SQL & ", anyodocum = " & ValorNulo
        SQL = SQL & ", situdocum = " & ValorNulo
        SQL = SQL & ", situacion = 0 "
        SQL = SQL & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and "
        SQL = SQL & " numfactu = " & DBSet(Rs!NumFactu, "T") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and "
        SQL = SQL & " numorden = " & DBSet(Rs!numorden, "N") & " AND codmacta = " & DBSet(Rs!codmacta, "T")
                    
        Conn.Execute SQL
    
        Rs.MoveNext
    Wend

    SQL = "delete from transferencias where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    SQL = SQL & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Conn.Execute SQL

    Set Rs = Nothing
    ModificarPagos = True
    Conn.CommitTrans
    Exit Function
    
eModificarCobros:
    Conn.RollbackTrans
    MuestraError Err.Number, "Modificar Pagos", Err.Description
End Function





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub


Private Sub BotonAnyadir()

    
    ModoInsertar = False
    
    LimpiarCampos
    Modo = 3
    PonerModo Modo

    txtfecha(4).Text = Format(Now, "dd/mm/yyyy")

    txtCuentas(2).Text = BancoPropio
    If txtCuentas(2).Text <> "" Then
        txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "bancos", "descripcion", "codmacta", txtCuentas(2), "T")
        If txtNCuentas(2).Text = "" Then txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtCuentas(2).Text, "T")
    End If

    PonleFoco txtfecha(2)
    
    Label2.Caption = ""

    Me.Label3(8).Caption = "Fecha Pago"
    Label1(1).Caption = "Banco"



End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    
    Me.lwCobros.ListItems.Clear
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub BotonModificar()
Dim SQL As String
    
 
    If lw1.SelectedItem Is Nothing Then Exit Sub
    If lw1.SelectedItem = 0 Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    ModoInsertar = True
    
    LimpiarCampos
    
    CadenaDesdeOtroForm = ""
    
    Modo = 4
    PonerModo Modo

    txtfecha(5).Text = Format(lw1.SelectedItem.SubItems(2), "dd/mm/yyyy")
    txtCuentas(3).Text = lw1.SelectedItem.SubItems(5)
    txtNCuentas(3).Text = lw1.SelectedItem.SubItems(6)
    Text2.Text = lw1.SelectedItem.SubItems(7)

    Select Case TipoTrans
        Case 0, 1
            Label3(12).Caption = "Transferencia: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""
        Case 2
            Label3(12).Caption = "Pago Domiciliado: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""
        Case 3
            Label3(12).Caption = "Confirming: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""
    End Select
        
        

    Me.Label3(8).Caption = "Fecha Pago"
    Label1(1).Caption = "Banco"
    
    If TipoTrans = 0 Then
        SQL = "from cobros, formapago where transfer = " & DBSet(lw1.SelectedItem.Text, "N") & " and anyorem = " & lw1.SelectedItem.SubItems(1)
        SQL = SQL & " and cobros.codforpa = formapago.codforpa"
    
        PonerVtosTransferencia SQL, True
    Else
        SQL = "from pagos where nrodocum = " & DBSet(lw1.SelectedItem.Text, "N") & " and anyodocum = " & lw1.SelectedItem.SubItems(1)
    
        PonerVtosTransferenciaPagos SQL, True
    End If
    
    If TipoTrans = 0 Or TipoTrans = 1 Then
    
        PosicionarCombo Me.cboConcepto2, lw1.SelectedItem.SubItems(11)
        
    Else
    
        chkFecha.Value = lw1.SelectedItem.SubItems(11)
        
    End If
    
    
    

    PonleFoco txtCuentas(3)

End Sub




Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    PonerIndicador lblIndicador, Modo
    
    If Modo = 3 Or Modo = 4 Then
        Frame1.visible = False
        Frame1.Enabled = False
    
        Me.FrameCreacionRemesa.visible = True
        Me.FrameCreacionRemesa.Enabled = True
    End If
    
    If Modo = 3 Then
        Me.FrameCreaRem.visible = True
        Me.FrameCreaRem.Enabled = True
        
        Me.FrameModRem.visible = False
        Me.FrameModRem.Enabled = False
        
        Me.lwCobros.Enabled = True
        Me.imgCheck(0).Enabled = True
        Me.imgCheck(1).Enabled = True
    Else
        Me.FrameCreaRem.visible = False
        Me.FrameCreaRem.Enabled = False
        
        Me.FrameModRem.visible = True
        Me.FrameModRem.Enabled = True
        
        Me.lwCobros.Enabled = False
        Me.imgCheck(0).Enabled = False
        Me.imgCheck(1).Enabled = False
    End If
    
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar2 Button.Index
End Sub

Private Sub HacerToolBar2(Boton As Integer)

Dim VtosAgrupados As Integer

    Select Case Boton
        Case 1
            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
        
        
            CadenaDesdeOtroForm = ""
            If Asc(UCase(lw1.SelectedItem.SubItems(9))) > Asc("B") Then
                Select Case TipoTrans
                    Case 0, 1
                        CadenaDesdeOtroForm = "No se puede modificar una transferencia " & lw1.SelectedItem.SubItems(3)
                    Case 2
                        CadenaDesdeOtroForm = "No se puede modificar un pago domiciliado " & lw1.SelectedItem.SubItems(3)
                    Case 3
                        CadenaDesdeOtroForm = "No se puede modificar un confirming " & lw1.SelectedItem.SubItems(3)
                End Select
            End If
            
            If CadenaDesdeOtroForm <> "" Then
                MsgBox CadenaDesdeOtroForm, vbExclamation
                Exit Sub
            End If
        
            
            'Ene 2018
            VtosAgrupados = 0
            If TipoTrans < 2 Then
                'Transferencias /abonos. Podemos agrupar vtos en fichero
                CadenaDesdeOtroForm = ""
                frmMensajes.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & TipoTrans & "|"
                frmMensajes.Opcion = 64
                frmMensajes.Show vbModal
                If CadenaDesdeOtroForm = "" Then Exit Sub
                Fecha = RecuperaValor(CadenaDesdeOtroForm, 1)
                
                If RecuperaValor(CadenaDesdeOtroForm, 2) = "1" Then
                    VtosAgrupados = ComprobacionesAgrupaFicheroTransfer(TipoTrans = 0)
                    If VtosAgrupados = -1 Then Exit Sub
                    
                End If
                
                
                
            End If
        
            If BloqueoManual(True, "ModTransfer", CStr(lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1))) Then
        
                CadenaDesdeOtroForm = ""
        
                GeneraNormaBancaria VtosAgrupados
                
                
                'Desbloqueamos
                BloqueoManual False, "ModTransfer", CStr(lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1))
            
                'Hay que poner en el formualrio de arriba valor a cadenadesdeotroform si ha modificado
                If CadenaDesdeOtroForm <> "" Then CargaList
            
            
            
            Else
                MsgBox "Registro bloqueado", vbExclamation
            End If
    
        Case 2 ' CONTABILIZACION TRANSFERENCIA
            
            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
            
            
            HaHabidoCambios = False
        
            SQL = "No se puede contabilizar "
            Select Case TipoTrans
                Case 0, 1
                    SQL = SQL & "una transferencia abierta. Sin llevar al banco."
                Case 2
                    SQL = SQL & "un pago domiciliado abierto. Sin llevar al banco."
                Case 3
                    SQL = SQL & "un confirming abierto. Sin llevar al banco."
            End Select
            
            CadenaDesdeOtroForm = ""
            If lw1.SelectedItem.SubItems(9) = "A" Then CadenaDesdeOtroForm = SQL
            
            If lw1.SelectedItem.SubItems(9) = "Q" Then
                SQL = "No se puede contabilizar "
                Select Case TipoTrans
                    Case 0, 1
                        CadenaDesdeOtroForm = SQL & "Transferencia abonada."
                    Case 2
                        CadenaDesdeOtroForm = SQL & "Pago domiciliado abonado."
                    Case 3
                        CadenaDesdeOtroForm = SQL & "Confirming abonado."
               End Select
            End If
            
            
            SQL = ""
            If CadenaDesdeOtroForm <> "" Then
                If TipoTrans < 2 Then
                    If vUsu.Nivel = 0 Then
                        SQL = "S"
                    Else
                        SQL = "N"
                    End If
                Else
                    SQL = "N"
                End If
            End If
            
            If SQL <> "" Then
                If SQL = "N" Then
                    MsgBox CadenaDesdeOtroForm, vbExclamation
                    CadenaDesdeOtroForm = ""
                    Exit Sub
                Else
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & vbCrLf & "¿Seguro que desea contabilizarla?" & vbCrLf
                    CadenaDesdeOtroForm = String(70, "*") & vbCrLf & CadenaDesdeOtroForm & vbCrLf & String(70, "*")
                    If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                End If
            End If
            CadenaDesdeOtroForm = ""
            
            frmTESTransferenciasCont.Opcion = 8
            frmTESTransferenciasCont.Cobros = (TipoTrans = 0)
            frmTESTransferenciasCont.NumeroDocumento = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(4) & "|" & lw1.SelectedItem.SubItems(5) & "|" & lw1.SelectedItem.SubItems(8) & "|"
            frmTESTransferenciasCont.TipoTrans = TipoTrans
            frmTESTransferenciasCont.Show vbModal
         
            'Hay que poner en el formualrio de arriba valor a cadenadesdeotroform si ha modificado
            If HaHabidoCambios Then CargaList
         
    End Select
End Sub


Private Function GeneraNormaBancaria(AgrupaVtos As Integer) As Boolean
Dim B As Boolean
Dim NumF As String
Dim NIF As String
Dim IdFich As String

    On Error GoTo EGeneraNormaBancaria
    GeneraNormaBancaria = False
    
    
    NumF = DevuelveDesdeBD("nifempre", "empresa2", "1", "1", "N")
    If NumF = "" Then
        MsgBox "La empresa no tiene NIF. No puede generar fichero norma bancaria", vbExclamation
        Exit Function
    Else
        NIF = NumF
    End If
    
   
    
    'Comprobamos las cuentas del banco de los recibos
    Set miRsAux = New ADODB.Recordset
    
    'Que el banco este bien
    
    ' tipotrans:  1 = pagos  (transferencias de pagos)
    '             0 = cobros (transferencias de abono)
    
   
    If TipoTrans < 2 Then
        If Not comprobarCuentasBancariasPagos(lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(1), TipoTrans = 1) Then
            Set miRsAux = Nothing
            Exit Function
        End If
    End If
    
    If Not ComprobarNifDatosProveedor Then Exit Function
        
   
    'Si es Norma34, transferenia, ofertaremos si queremos que el fichero sea
   
    Set miRsAux = Nothing
        'B = GeneraFicheroNorma34(NIF, Adodc1.Recordset!Fecha, Adodc1.Recordset!codmacta, "9", Adodc1.Recordset!Codigo, Adodc1.Recordset!descripcion, TipoDeFrm <> 0)
        
    If TipoTrans < 2 Then
            'Estaba como fecha: lw1.SelectedItem.SubItems(2)
        B = GeneraFicheroNorma34(NIF, Fecha, lw1.SelectedItem.SubItems(5), CStr(lw1.SelectedItem.SubItems(11)), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), TipoTrans = 1, lw1.SelectedItem.SubItems(1), IdFich, AgrupaVtos > 0)
        
    Else
         If TipoTrans = 3 Then ' si es caixa confirmning
            'Van por una "norma" de la caixa. De momento picassent
            B = GeneraFicheroCaixaConfirming(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
            
        Else  ' pagos domiciliados
            'Q68
            'Fontenas, herbelca....
            B = GeneraFicheroNorma68(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
        End If
    End If
   
    
    
    If B Then
        If cd1.FileName <> "" Then cd1.FileName = ""
        cd1.ShowSave
        If cd1.FileName <> "" Then
            If Dir(cd1.FileName, vbArchive) <> "" Then
                If MsgBox("El archivo " & cd1.FileName & " ya existe" & vbCrLf & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
                Kill cd1.FileName
            End If
        
            'CopiarFicheroNorma43 (TipoDeFrm < 2), cd1.FileName
            CopiarFicheroNormaBancaria CByte(TipoTrans), cd1.FileName
            
            CadenaDesdeOtroForm = "OK"
                        
            
            If TipoTrans < 2 Then
                Set miRsAux = New ADODB.Recordset
                SQL = "Select * from transferencias where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
                SQL = SQL & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    'Si a estaba generado el fichero, guardo en LOG
                    If Asc(DBLet(miRsAux!Situacion, "T")) >= Asc("B") Then
                        'YA se habia generado
                        SQL = "Transfer: " & Format(miRsAux!Codigo, "0000") & "/" & miRsAux!Anyo & "   Fecha:" & miRsAux!Fecha & vbCrLf
                        SQL = SQL & DBLet(miRsAux!Descripcion, "T") & vbCrLf
                        
                        SQL = SQL & "Banco: " & miRsAux!codmacta & " " & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux!codmacta, "T") & vbCrLf
                        SQL = SQL & "Situacion: " & miRsAux!Situacion & " Importe : " & miRsAux!Importe & "€ "
                        If TipoTrans = 0 Then
                            RC = "transfer= " & miRsAux!Codigo & " AND anyorem =" & miRsAux!Anyo & " AND 1"
                            RC = DevuelveDesdeBD("count(*)", "cobros", RC, "1")
                        Else
                            RC = "nrodocum= " & miRsAux!Codigo & " AND anyodocum =" & miRsAux!Anyo & " AND 1"
                            RC = DevuelveDesdeBD("count(*)", "pagos", RC, "1")
                        End If
                        SQL = SQL & " Vtos: " & RC
                        If DBLet(miRsAux!LlevaAgrupados, "N") > 0 Then SQL = SQL & " Cli/Prov agr: " & miRsAux!LlevaAgrupados
                        SQL = SQL & vbCrLf
                        RC = DBLet(miRsAux!IdFicheroSEPA, "T")
                        If RC <> "" Then
                            RC = "ID_Fich: " & RC & "       Usuario: " & DBLet(miRsAux!usurioFich, "T")
                            SQL = SQL & RC
                        End If
                        RC = ""
                        vLog.Insertar 31, vUsu, SQL
                    End If
                End If
                miRsAux.Close
                Set miRsAux = Nothing
                
            End If
            
            
            
            
            NIF = "update transferencias set situacion = 'B' "
            If TipoTrans < 2 Then
                '`IdFicheroSEPA` `LlevaAgrupados``usurioFich`
                NIF = NIF & ", IdFicheroSEPA = " & DBSet(IdFich, "T")
                NIF = NIF & ", LlevaAgrupados = " & DBSet(AgrupaVtos, "T")
                NIF = NIF & ", usurioFich = " & DBSet(vUsu.Login, "T")
                NIF = NIF & ", fecha = " & DBSet(Fecha, "F")
            End If
            NIF = NIF & " where codigo = " & DBSet(lw1.SelectedItem.Text, "N") & " and anyo = " & DBSet(lw1.SelectedItem.SubItems(1), "N")
 
            
            
            If Ejecuta(NIF) Then CadenaDesdeOtroForm = "OK"
            
            
            If CadenaDesdeOtroForm = "OK" Then
            
                Set miRsAux = New ADODB.Recordset
                If TipoTrans = 0 Then
                    If Not UpdatearCobrosTransferencia Then MsgBox "Error updateando cobros transferencia", vbExclamation
                Else
                    If Not UpdatearPagosTransferencia Then
                        Select Case TipoTrans
                            Case 1
                                MsgBox "Error updateando pagos transferencia", vbExclamation
                            Case 2
                                MsgBox "Error updateando pagos domiciliados", vbExclamation
                            Case 3
                                MsgBox "Error updateando pagos confirming", vbExclamation
                        End Select
                    End If
                End If
                
                Set miRsAux = Nothing
            End If
            
        End If
    Else
        MsgBox "Error generando fichero", vbExclamation
    End If
    
    
    Exit Function
EGeneraNormaBancaria:
    MuestraError Err.Number, "Genera Fichero Norma34"
End Function


Private Function UpdatearCobrosTransferencia() As Boolean
Dim Im As Currency
    On Error GoTo EUpdatearCobrosTransferencia
    UpdatearCobrosTransferencia = False
    
    SQL = "Select * from cobros WHERE transfer=" & DBSet(lw1.SelectedItem.Text, "N")
    SQL = SQL & " AND anyorem =" & DBSet(lw1.SelectedItem.SubItems(1), "N")
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            While Not miRsAux.EOF
                SQL = "UPDATE cobros SET fecultco = '" & Format(lw1.SelectedItem.SubItems(2), FormatoFecha) & "', impcobro = "
                Im = miRsAux!ImpVenci
                If Not IsNull(miRsAux!Gastos) Then Im = Im + miRsAux!Gastos
                SQL = SQL & TransformaComasPuntos(CStr(Im))
                
                SQL = SQL & " ,siturem = 'B'"
                
                
                'WHERE
                SQL = SQL & " WHERE numserie='" & miRsAux!NUmSerie
                SQL = SQL & "' AND  numfactu =  " & miRsAux!NumFactu
                SQL = SQL & "  AND  fecfactu =  '" & Format(miRsAux!FecFactu, FormatoFecha)
                SQL = SQL & "' AND  numorden =  " & miRsAux!numorden
                'Muevo siguiente
                miRsAux.MoveNext
                
                'Ejecuto SQL
                If Not Ejecuta(SQL) Then MsgBox "Error: " & SQL, vbExclamation
            Wend
    End If
    miRsAux.Close
                    
                    
                    
    UpdatearCobrosTransferencia = True
    Exit Function
EUpdatearCobrosTransferencia:
    
End Function

Private Function UpdatearPagosTransferencia() As Boolean
Dim Im As Currency
    On Error GoTo EUpdatearCobrosTransferencia
    UpdatearPagosTransferencia = False
    
    SQL = "Select * from pagos WHERE nrodocum=" & DBSet(lw1.SelectedItem.Text, "N")
    SQL = SQL & " AND anyodocum =" & DBSet(lw1.SelectedItem.SubItems(1), "N")
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            While Not miRsAux.EOF
                SQL = "UPDATE pagos SET fecultpa = '" & Format(lw1.SelectedItem.SubItems(2), FormatoFecha) & "', imppagad = "
                Im = miRsAux!ImpEfect
                SQL = SQL & TransformaComasPuntos(CStr(Im))
                
                SQL = SQL & " ,situdocum = 'B'"
                
                
                'WHERE
                SQL = SQL & " WHERE numserie='" & miRsAux!NUmSerie
                SQL = SQL & "' AND  numfactu =  " & DBSet(miRsAux!NumFactu, "T")
                SQL = SQL & "  AND  fecfactu =  '" & Format(miRsAux!FecFactu, FormatoFecha)
                SQL = SQL & "' AND  numorden =  " & miRsAux!numorden
                SQL = SQL & " and codmacta = " & DBSet(miRsAux!codmacta, "T")
                'Muevo siguiente
                miRsAux.MoveNext
                
                'Ejecuto SQL
                If Not Ejecuta(SQL) Then MsgBox "Error: " & SQL, vbExclamation
            Wend
    End If
    miRsAux.Close
                    
                    
                    
    UpdatearPagosTransferencia = True
    Exit Function
EUpdatearCobrosTransferencia:
    
End Function





Private Function ComprobarNifDatosProveedor() As Boolean
Dim SQL As String

    ComprobarNifDatosProveedor = False
    
    If TipoTrans = 0 Then
    
        SQL = "select nifclien, codmacta, nomclien from cobros where transfer = " & lw1.SelectedItem.Text & " and anyorem = " & DBSet(lw1.SelectedItem.SubItems(1), "N")
        SQL = SQL & " GROUP BY 1"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            
            If Trim(DBLet(miRsAux!nifclien, "T")) = "" Then SQL = SQL & "- " & miRsAux!codmacta & " " & miRsAux!nomclien & vbCrLf
            miRsAux.MoveNext
        
        Wend
        
    Else
        
        SQL = "select nifprove, codmacta, nomprove from pagos where nrodocum = " & lw1.SelectedItem.Text & " and anyodocum = " & DBSet(lw1.SelectedItem.SubItems(1), "N")
        SQL = SQL & " GROUP BY 1"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        While Not miRsAux.EOF
            
            If Trim(DBLet(miRsAux!nifprove, "T")) = "" Then SQL = SQL & "- " & miRsAux!codmacta & " " & miRsAux!nomprove & vbCrLf
            miRsAux.MoveNext
        
        Wend
    
    End If
        
    miRsAux.Close
    If SQL <> "" Then
        MsgBox "Error en NIFs: " & vbCrLf & SQL, vbExclamation
        Set miRsAux = Nothing
    Else
        ComprobarNifDatosProveedor = True
    End If
    

End Function






Private Function HacerEliminacionTransferenciaVtos() As Boolean

    On Error GoTo EHacerEliminacionTransferenciaVtos

    HacerEliminacionTransferenciaVtos = False

    'Eliminamos los vencimientos asociados
    Conn.Execute "DELETE FROM cobros where transfer=" & lw1.SelectedItem.Text & " AND anyorem =" & lw1.SelectedItem.SubItems(1)
    
    'Eliminamos la remesa
    Conn.Execute "DELETE FROM transferencias where codigo=" & lw1.SelectedItem.Text & " AND anyo =" & lw1.SelectedItem.SubItems(1)
    
    HacerEliminacionTransferenciaVtos = True
    Exit Function
    
EHacerEliminacionTransferenciaVtos:
    MuestraError Err.Number, "Function: HacerEliminacionTransferenciaVtos"
End Function


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    PonFoco Text1(Index)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    If Text1(Index).Text = "" Then
        Exit Sub
    End If
    
    Select Case Index
        Case 1 ' fecha
            PonerFormatoFecha Text1(Index)
        
        Case 2 ' cuenta
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "La cuenta debe ser numérica: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                Text1(3).Text = ""
                Text1(6).Tag = Text1(6).Text
                PonFoco Text1(Index)
                
                Exit Sub
            End If
            
            Select Case Index
            Case Else
                'DE ULTIMO NIVEL
                Cta = (Text1(Index).Text)
                If CuentaCorrectaUltimoNivel(Cta, SQL) Then
                    Text1(Index).Text = Cta
                    Text1(3).Text = SQL
                Else
                    MsgBox SQL, vbExclamation
                    Text1(Index).Text = ""
                    Text1(3).Text = ""
                    Text1(Index).SetFocus
                End If
                
            End Select
        Case 4
            PonerFormatoDecimal Text1(Index), 1
    End Select
End Sub

Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If Text1(Indice1).Text <> "" And Text1(Indice2).Text <> "" Then
        L1 = Len(Text1(Indice1).Text)
        L2 = Len(Text1(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(Text1(Indice1).Text & "000000000", 1, L1)) > Val(Mid(Text1(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function


Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


'Si tiene valor el campo fecha, entonces lo ponemos con el BD
Private Function CampoABD(ByRef T As TextBox, Tipo As String, CampoEnLaBD, Mayor_o_Igual As Boolean) As String

    CampoABD = ""
    If T.Text <> "" Then
        If Mayor_o_Igual Then
            CampoABD = " >= "
        Else
            CampoABD = " <= "
        End If
        Select Case Tipo
        Case "F"
            CampoABD = CampoEnLaBD & CampoABD & "'" & Format(T.Text, FormatoFecha) & "'"
        Case "T"
            CampoABD = CampoEnLaBD & CampoABD & "'" & T.Text & "'"
        Case "N"
            CampoABD = CampoEnLaBD & CampoABD & T.Text
        End Select
    End If
End Function



Private Function CampoBD_A_SQL(ByRef C As ADODB.Field, Tipo As String, Nulo As Boolean) As String

    If IsNull(C) Then
        If Nulo Then
            CampoBD_A_SQL = "NULL"
        Else
            If Tipo = "T" Then
                CampoBD_A_SQL = "''"
            Else
                CampoBD_A_SQL = "0"
            End If
        End If
    Else
    
        Select Case Tipo
        Case "F"
            CampoBD_A_SQL = "'" & Format(C.Value, FormatoFecha) & "'"
        Case "T"
            CampoBD_A_SQL = "'" & DevNombreSQL(C.Value) & "'"
        Case "N"
            CampoBD_A_SQL = TransformaComasPuntos(CStr(C.Value))
        End Select
    End If
End Function

Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
    If TEXTO = "" Then TEXTO = "Generando datos"
    Me.Refresh
End Sub


Private Sub PonerVtosTransferencia(vSql As String, Modificar As Boolean)
Dim IT
Dim ImporteTot As Currency

    lwCobros.ListItems.Clear
    If Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
    
'    Set Me.lwCobros.SmallIcons = frmPpal.ImgListviews
    Set lwCobros.SmallIcons = frmppal.imgListComun16
    
    
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select cobros.*,nomforpa " & vSql
    Cad = Cad & " ORDER BY fecvenci"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.ListSubItems(3).Tag = ""
        
        
        If Not ComprobarIBANCuentaBancaria(DBLet(miRsAux!IBAN, "T"), Cad) Then
            IT.ListSubItems(3).Tag = "NO"
            IT.ListSubItems(3).Bold = True
            IT.ListSubItems(3).ForeColor = vbRed
            IT.ListSubItems(3).ToolTipText = Cad
        End If
        
        IT.SubItems(4) = miRsAux!FecVenci
        If IsNull(miRsAux!nomclien) Then
            IT.SubItems(5) = miRsAux!codmacta
        Else
            IT.SubItems(5) = miRsAux!nomclien
        End If
    
        If Modificar Then IT.Checked = True
    
    
            
    
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not Modificar Then
            If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        End If
        
        IT.SubItems(6) = Format(Importe, FormatoImporte)
        
        ImporteTot = ImporteTot + Importe

        IT.Tag = Abs(Importe)  'siempre valor absoluto
            
        If DBLet(miRsAux!Devuelto, "N") = 1 Then
            IT.SmallIcon = 42
        End If
            
        If Modificar Then
            IT.SubItems(7) = txtCuentas(3).Text
        Else
            IT.SubItems(7) = txtCuentas(2).Text
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Not Modificar Then
        Text1(4).Text = Format(ImporteTot, "###,###,##0.00")
    Else
        Text1(4).Text = Format(lw1.SelectedItem.SubItems(8), "###,###,##0.00")
    End If
    

End Sub


Private Sub PonerVtosTransferenciaPagos(vSql As String, Modificar As Boolean)
Dim IT
Dim ImporteTot As Currency

    lwCobros.ListItems.Clear
    If Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
    
'    Set Me.lwCobros.SmallIcons = frmPpal.ImgListviews
    Set lwCobros.SmallIcons = frmppal.imgListComun16
    
    
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select pagos.* " & vSql
    Cad = Cad & " ORDER BY fecefect"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = miRsAux!NumFactu
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        
         IT.ListSubItems(3).Tag = ""
        
        
        If Not ComprobarIBANCuentaBancaria(DBLet(miRsAux!IBAN, "T"), Cad) Then
            IT.ListSubItems(3).Tag = "NO"
            IT.ListSubItems(3).Bold = True
            IT.ListSubItems(3).ForeColor = vbRed
            IT.ListSubItems(3).ToolTipText = Cad
        End If
       
        
        
        
        
        
        IT.SubItems(4) = miRsAux!fecefect
        IT.SubItems(5) = IIf(IsNull(miRsAux!nomprove), "N/D", miRsAux!nomprove)
    
        If Modificar Then IT.Checked = True
    
        Importe = 0
        Importe = Importe + miRsAux!ImpEfect
        
        'Si ya he cobrado algo
        If Not Modificar Then
            If Not IsNull(miRsAux!imppagad) Then Importe = Importe - miRsAux!imppagad
        End If
        
        IT.SubItems(6) = Format(Importe, FormatoImporte)
        
        ImporteTot = ImporteTot + Importe

        IT.Tag = Abs(Importe)  'siempre valor absoluto
            
'        If DBLet(miRsAux!Devuelto, "N") = 1 Then
'            IT.SmallIcon = 42
'        End If
            
        If Modificar Then
            IT.SubItems(7) = txtCuentas(3).Text
        Else
            IT.SubItems(7) = txtCuentas(2).Text
        End If
        
        IT.SubItems(8) = miRsAux!codmacta
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Not Modificar Then
        Text1(4).Text = Format(ImporteTot, "###,###,##0.00")
    Else
        '###revisar
        Text1(4).Text = Format(Me.lw1.SelectedItem.SubItems(8), "###,###,##0.00")
    End If
    

End Sub


Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    SQL = ""
    For i = 1 To Me.lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(i).Checked Then
        
            Insertar = True
            If Me.lwCobros.ListItems(i).Bold Then
                RegistroDestino = i
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwCobros.ListItems(i).Text & "'," & lwCobros.ListItems(i).SubItems(1)
                SQL = SQL & ",'" & Format(lwCobros.ListItems(i).SubItems(2), FormatoFecha) & "'," & lwCobros.ListItems(i).SubItems(3) & ")"
            End If
            
        End If
    Next
    SQL = Mid(SQL, 2)
            
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
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = False 'DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N")
    
        Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N") And Not (lw1.SelectedItem Is Nothing)
        Toolbar2.Buttons(2).Enabled = DBLet(Rs!especial, "N") And Not (lw1.SelectedItem Is Nothing)
        Toolbar2.Buttons(3).Enabled = DBLet(Rs!especial, "N") 'And Not (lw1.SelectedItem Is Nothing)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub CargaList()
Dim IT

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select transferencias.codigo,transferencias.anyo, transferencias.fecha, "
    'CASE situacion WHEN 0 THEN 'ABIERTA' WHEN 1 THEN 'GENERADO FICHERO' WHEN 2 THEN 'CONTABILIZADA' END as descsituacion,"
    Cad = Cad & " wtiposituacionrem.descsituacion, "
    If TipoTrans < 2 Then
        Cad = Cad & " CASE concepto WHEN 0 THEN 'PENSION' WHEN 1 THEN 'NOMINA' WHEN 9 THEN 'ORDINARIA' END as desconcepto, "
    Else
        ' pagos domiciliados como caixa confirming
        Cad = Cad & " if (concepto = 0,'Vencimiento','Fecha intro.') as desconcepto, "
    End If
    Cad = Cad & " transferencias.codmacta,cuentas.nommacta,"
    Cad = Cad & " transferencias.descripcion, Importe , transferencias.tipotrans, situacion, transferencias.concepto "
    Cad = Cad & " from cuentas,transferencias,usuarios.wtiposituacionrem where transferencias.codmacta=cuentas.codmacta"
    Cad = Cad & " and wtiposituacionrem.situacio = transferencias.situacion "
    
    Cad = Cad & PonerFiltro
    
    If CampoOrden = "" Then CampoOrden = "transferencias.anyo, transferencias.codigo "
    Cad = Cad & " ORDER BY " & CampoOrden ' transferencias.anyo desc,
    If Orden Then Cad = Cad & " DESC"
    
    lw1.ColumnHeaders.Clear
    
    lw1.ColumnHeaders.Add , , "Código", 950
    lw1.ColumnHeaders.Add , , "Año", 700
    lw1.ColumnHeaders.Add , , "Fecha", 1350
    lw1.ColumnHeaders.Add , , "Situación", 1540
    lw1.ColumnHeaders.Add , , "Concepto", 1500
    lw1.ColumnHeaders.Add , , "Cuenta", 1440
    lw1.ColumnHeaders.Add , , "Nombre", 2940
    lw1.ColumnHeaders.Add , , "Descripción", 2840
    lw1.ColumnHeaders.Add , , "Importe", 1940, 1
    lw1.ColumnHeaders.Add , , "S", 0, 1
    lw1.ColumnHeaders.Add , , "T", 0, 1
    lw1.ColumnHeaders.Add , , "C", 0, 1
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = DBLet(miRsAux!Codigo, "N")
        IT.SubItems(1) = DBLet(miRsAux!Anyo, "N")
        IT.SubItems(2) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        IT.SubItems(3) = DBLet(miRsAux!descsituacion, "T")
        IT.ListSubItems(3).ToolTipText = DBLet(miRsAux!descsituacion, "T")
        
        IT.SubItems(4) = DBLet(miRsAux!desconcepto, "T")
        IT.ListSubItems(4).ToolTipText = DBLet(miRsAux!desconcepto, "T")
        
        
        IT.SubItems(5) = miRsAux!codmacta
        IT.SubItems(6) = DBLet(miRsAux!Nommacta, "T")
        IT.ListSubItems(6).ToolTipText = DBLet(miRsAux!Nommacta, "T")
        IT.SubItems(7) = DBLet(miRsAux!Descripcion, "T")
        IT.ListSubItems(7).ToolTipText = DBLet(miRsAux!Descripcion, "T")
        IT.SubItems(8) = Format(miRsAux!Importe, "###,###,##0.00")
        IT.SubItems(9) = miRsAux!Situacion
        IT.SubItems(10) = miRsAux!TipoTrans
        IT.SubItems(11) = miRsAux!Concepto
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    If lw1.ListItems.Count > 0 Then
        Modo = 2
    Else
        Modo = 0
    End If
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub



Private Function PonerFiltro()
Dim C As String
    'Filtro
    Select Case TipoTrans
        Case 0 ' abonos
            C = " and transferencias.tipotrans = 1"
        Case 1 ' pagos
            C = " and transferencias.tipotrans = 0 and transferencias.subtipo = 0"
        Case 2 ' pagos domiciliados
            C = " and transferencias.tipotrans = 0 and transferencias.subtipo = 1"
        Case 3 ' confirming
            C = " and transferencias.tipotrans = 0 and transferencias.subtipo = 2"
    End Select
    
    PonerFiltro = C
End Function


Private Sub NuevaTransf()

Dim Forpa As String
Dim Cad As String
Dim Impor As Currency
Dim Sql2 As String

    SQL = " (1=1) "
    
    
    
    'Del vto
    If txtfecha(2).Text <> "" Then SQL = SQL & " AND cobros.fecvenci >= '" & Format(txtfecha(2).Text, FormatoFecha) & "'"
    If txtfecha(3).Text <> "" Then SQL = SQL & " AND cobros.fecvenci <= '" & Format(txtfecha(3).Text, FormatoFecha) & "'"
   
    
    'Si ha puesto importe desde Hasta
    If txtImporte(0).Text <> "" Then SQL = SQL & " AND impvenci >= " & TransformaComasPuntos(ImporteFormateado(txtImporte(0).Text))
    If txtImporte(1).Text <> "" Then SQL = SQL & " AND impvenci <= " & TransformaComasPuntos(ImporteFormateado(txtImporte(1).Text))
    
    
    'Desde hasta cuenta
    If Me.txtCuentas(0).Text <> "" Then SQL = SQL & " AND cobros.codmacta >= '" & txtCuentas(0).Text & "'"
    If Me.txtCuentas(1).Text <> "" Then SQL = SQL & " AND cobros.codmacta <= '" & txtCuentas(1).Text & "'"
    
    'El importe
    SQL = SQL & " AND (impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) < 0"
    

    'serie
    If txtSerie(0).Text <> "" Then _
        SQL = SQL & " AND cobros.numserie >= '" & txtSerie(0).Text & "'"
    If txtSerie(1).Text <> "" Then _
        SQL = SQL & " AND cobros.numserie <= '" & txtSerie(1).Text & "'"
    
    'Fecha factura
    If txtfecha(0).Text <> "" Then _
        SQL = SQL & " AND cobros.fecfactu >= '" & Format(txtfecha(0).Text, FormatoFecha) & "'"
    If txtfecha(1).Text <> "" Then _
        SQL = SQL & " AND cobros.fecfactu <= '" & Format(txtfecha(1).Text, FormatoFecha) & "'"
    
    'Codigo factura
    If txtNumFac(0).Text <> "" Then _
        SQL = SQL & " AND cobros.numfactu >= '" & txtNumFac(0).Text & "'"
    If txtNumFac(1).Text <> "" Then _
        SQL = SQL & " AND cobros.numfactu <= '" & txtNumFac(1).Text & "'"
    
    
    SQL = SQL & " and situacion = 0 "
     
    If Me.TipoTrans = 0 Then
        If Me.chkAbonosSoloTransferencia.Value = 1 Then SQL = SQL & " AND formapago.tipforpa=1"
    End If
    
    
    
    CadenaDesdeOtroForm = ""

    Sql2 = SQL & " and not cobros.impcobro is null and cobros.impcobro <> 0 and cobros.codmacta=cuentas.codmacta AND (transfer is null) AND cobros.codforpa = formapago.codforpa "
    
    Sql2 = "select cobros.* FROM cobros,cuentas,formapago  WHERE " & Sql2
    
    If TotalRegistrosConsulta(Sql2) <> 0 Then
    
        Set frmMens3 = New frmMensajes
        
        frmMens3.Opcion = 53
        frmMens3.Parametros = Sql2
        frmMens3.Show vbModal
        
        Set frmMens = Nothing
        
        If CadenaDesdeOtroForm <> "OK" Then
            cmdCancelar_Click (0)
        End If
    
    End If
    
    SQL = SQL & " and (cobros.impcobro is null or cobros.impcobro = 0)"
        
     
    Screen.MousePointer = vbHourglass
    Set Rs = New ADODB.Recordset
    
        
    
    'Que la cuenta NO este bloqueada
    i = 0
    
    Cad = " FROM cobros,formapago,cuentas WHERE cobros.codforpa = formapago.codforpa AND transfer is null AND situacion = 0 and "
    Cad = Cad & " cobros.codmacta=cuentas.codmacta AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtfecha(4).Text), FormatoFecha) & "') AND "
    Cad = "Select cobros.codmacta,nommacta,fecbloq" & Cad & SQL & " GROUP BY 1 ORDER BY 1"
        
    
    
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Cad = ""
        i = 1
        While Not Rs.EOF
            Cad = Cad & Rs!codmacta & " - " & Rs!Nommacta & " : " & Rs!FecBloq & vbCrLf
            Rs.MoveNext
        Wend
    End If

    Rs.Close
    
    If i > 0 Then
        Cad = "Las siguientes cuentas estan bloqueadas." & vbCrLf & String(60, "-") & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
        Exit Sub
    End If
    
    
    Cad = " FROM cobros,formapago,cuentas WHERE cobros.codforpa = formapago.codforpa AND transfer is null AND "
    Cad = Cad & " cobros.codmacta=cuentas.codmacta AND situacion = 0 and "
    
    'Hacemos un conteo
    Rs.Open "SELECT Count(*) " & Cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        i = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Cad = Cad & SQL
    
    
    
    If i > 0 Then
        i = 1  'Para que siga por abajo
    End If
    
    

    'La suma
    If i > 0 Then
        SQL = "select sum(impvenci),sum(impcobro),sum(gastos) " & Cad
        Impor = 0
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N") + DBLet(Rs.Fields(2), "N")
        Rs.Close
        If Impor = 0 Then i = 0
    End If
        

    Set Rs = Nothing
    
    If i = 0 Then
        MsgBox "Ningun dato a transferir con esos valores", vbExclamation
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
    Else
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        SQL = "Delete from tmpcierre1 where codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        CadenaDesdeOtroForm = ""
        
        SQL = Cad 'Le paso el SELECT
        
        PonerVtosTransferencia SQL, False
        
        Dim CadAux As String
        
        CadAux = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES (" & vUsu.Codigo
        CadAux = CadAux & ",'" & txtCuentas(2).Text & "','" & txtNCuentas(2).Text & "'," & DBSet(Text1(4).Text, "N") & ")"
        If Not Ejecuta(CadAux) Then Exit Sub
        
        CadenaDesdeOtroForm = "'" & Trim(txtCuentas(2).Text) & "'"
                
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub NuevaTransfPagos()

Dim Forpa As String
Dim Cad As String
Dim Impor As Currency
Dim Sql2 As String


    If TipoTrans = 1 Then
        ' transferencias
        SQL = "formapago.tipforpa = " & vbTransferencia
        
    Else
        ' pagos domiciliados o confirming
        If TipoTrans = 2 Then
            SQL = "formapago.tipforpa = " & vbPagoDomiciliado
        Else
            SQL = "formapago.tipforpa = " & vbConfirming
        End If
    End If
    
    'Del vto
    If txtfecha(2).Text <> "" Then SQL = SQL & " AND pagos.fecefect >= '" & Format(txtfecha(2).Text, FormatoFecha) & "'"
    If txtfecha(3).Text <> "" Then SQL = SQL & " AND pagos.fecefect <= '" & Format(txtfecha(3).Text, FormatoFecha) & "'"
   
    
    'Si ha puesto importe desde Hasta
    If txtImporte(0).Text <> "" Then SQL = SQL & " AND impefect >= " & TransformaComasPuntos(ImporteFormateado(txtImporte(0).Text))
    If txtImporte(1).Text <> "" Then SQL = SQL & " AND impefect <= " & TransformaComasPuntos(ImporteFormateado(txtImporte(1).Text))
    
    
    'Desde hasta cuenta
    If Me.txtCuentas(0).Text <> "" Then SQL = SQL & " AND pagos.codmacta >= '" & txtCuentas(0).Text & "'"
    If Me.txtCuentas(1).Text <> "" Then SQL = SQL & " AND pagos.codmacta <= '" & txtCuentas(1).Text & "'"
    
    'El importe
    SQL = SQL & " AND (impefect - coalesce(imppagad,0)) > 0"
    

    'serie
    If txtSerie(0).Text <> "" Then _
        SQL = SQL & " AND pagos.numserie >= '" & txtSerie(0).Text & "'"
    If txtSerie(1).Text <> "" Then _
        SQL = SQL & " AND pagos.numserie <= '" & txtSerie(1).Text & "'"
    
    'Fecha factura
    If txtfecha(0).Text <> "" Then _
        SQL = SQL & " AND pagos.fecfactu >= '" & Format(txtfecha(0).Text, FormatoFecha) & "'"
    If txtfecha(1).Text <> "" Then _
        SQL = SQL & " AND pagos.fecfactu <= '" & Format(txtfecha(1).Text, FormatoFecha) & "'"
    
    'Codigo factura
    If txtNumFac(0).Text <> "" Then _
        SQL = SQL & " AND pagos.numfactu >= '" & txtNumFac(0).Text & "'"
    If txtNumFac(1).Text <> "" Then _
        SQL = SQL & " AND pagos.numfactu <= '" & txtNumFac(1).Text & "'"
    
    
    SQL = SQL & " and situacion = 0 "
     
    ' si hay pagos con imppagad <> 0 damos aviso y no los incluimos
    
    
    
    CadenaDesdeOtroForm = ""

    Sql2 = SQL & " and not pagos.imppagad is null and pagos.imppagad <> 0 and pagos.codmacta=cuentas.codmacta AND (nrodocum is null) AND pagos.codforpa = formapago.codforpa "
   
    
    Sql2 = "select pagos.* FROM pagos,cuentas,formapago  WHERE " & Sql2
    
    If TotalRegistrosConsulta(Sql2) <> 0 Then
    
        Set frmMens3 = New frmMensajes
        
        frmMens3.Opcion = 53
        frmMens3.Parametros = Sql2
        frmMens3.Show vbModal
        
        Set frmMens = Nothing
        
        If CadenaDesdeOtroForm <> "OK" Then
            cmdCancelar_Click (0)
        End If
    
    End If
    
    SQL = SQL & " and (pagos.imppagad is null or pagos.imppagad = 0)"
        
     
    Screen.MousePointer = vbHourglass
    Set Rs = New ADODB.Recordset
    
        
    
    'Que la cuenta NO este bloqueada
    i = 0
    
    Cad = " FROM pagos,formapago,cuentas WHERE pagos.codforpa = formapago.codforpa AND nrodocum is null AND situacion = 0 and "
    Cad = Cad & " pagos.codmacta=cuentas.codmacta AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtfecha(4).Text), FormatoFecha) & "') AND "
    Cad = "Select pagos.codmacta,nommacta,fecbloq" & Cad & SQL & " GROUP BY 1 ORDER BY 1"
        
    
    
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Cad = ""
        i = 1
        While Not Rs.EOF
            Cad = Cad & Rs!codmacta & " - " & Rs!Nommacta & " : " & Rs!FecBloq & vbCrLf
            Rs.MoveNext
        Wend
    End If

    Rs.Close
    
    If i > 0 Then
        Cad = "Las siguientes cuentas estan bloqueadas." & vbCrLf & String(60, "-") & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
        Exit Sub
    End If
    
    
    Cad = " FROM pagos,formapago,cuentas WHERE pagos.codforpa = formapago.codforpa AND nrodocum is null AND "
    Cad = Cad & " pagos.codmacta=cuentas.codmacta AND situacion = 0 and "
    
    'Hacemos un conteo
    Rs.Open "SELECT Count(*) " & Cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        i = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Cad = Cad & SQL
    
    
    
    If i > 0 Then
        i = 1  'Para que siga por abajo
    End If
    
    

    'La suma
    If i > 0 Then
        SQL = "select sum(impefect),sum(imppagad) " & Cad
        Impor = 0
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N")
        Rs.Close
        If Impor = 0 Then i = 0
    End If
        

    Set Rs = Nothing
    
    If i = 0 Then
        MsgBox "Ningun dato  con esos valores", vbExclamation
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
    Else
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        SQL = "Delete from tmpcierre1 where codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        CadenaDesdeOtroForm = ""
        
        SQL = Cad 'Le paso el SELECT
        
        PonerVtosTransferenciaPagos SQL, False
        
        Dim CadAux As String
        
        CadAux = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES (" & vUsu.Codigo
        CadAux = CadAux & ",'" & txtCuentas(2).Text & "','" & txtNCuentas(2).Text & "'," & DBSet(Text1(4).Text, "N") & ")"
        If Not Ejecuta(CadAux) Then Exit Sub
        
        CadenaDesdeOtroForm = "'" & Trim(txtCuentas(2).Text) & "'"
                
    End If
    
    Screen.MousePointer = vbDefault
End Sub





Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgSerie"
        imgSerie_Click Indice
    Case "imgFecha"
        imgFec_Click Indice
    Case "imgCuentas"
        imgCuentas_Click Indice
    End Select
End Sub


Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub

Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtCuentas(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCuentas_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    
    
    If txtCuentas(Index).Text = "" Then
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCuentas(Index).Text) Then
        If InStr(1, txtCuentas(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCuentas(Index).Text, vbExclamation
        txtCuentas(Index).Text = ""
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1, 2, 3 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            'B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
            
            B = CuentaCorrectaUltimoNivel(Cta, SQL)
            
            If Not B Then
                MsgBox SQL, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
                PonleFoco txtCuentas(Index)
            Else
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = SQL
                If B = 1 Then
                    txtNCuentas(Index).Tag = ""
                Else
                    txtNCuentas(Index).Tag = SQL
                End If
                Hasta = -1
                If Index = 6 Then
                    Hasta = 7
                Else
                    If Index = 0 Then
                        Hasta = 1
                    Else
                        If Index = 5 Then
                            Hasta = 4
                        Else
                            If Index = 23 Then Hasta = 24
                        End If
                    End If
                    
                End If
                    
                If Hasta >= 0 Then
                    txtCuentas(Hasta).Text = txtCuentas(Index).Text
                    txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
                End If
            End If
    
    End Select
    
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtfecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtfecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtfecha(Index).Text = Trim(txtfecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtfecha(Index)
End Sub

Private Sub txtRemesa_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSerie_GotFocus(Index As Integer)
    ConseguirFoco txtSerie(Index), 3
End Sub

Private Sub txtSerie_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtSerie(Index).Text = UCase(Trim(txtSerie(Index).Text))
    
    If txtSerie(Index).Text = "" Then
        txtNSerie(Index).Text = ""
        Exit Sub
    End If
    
    Select Case Index
        Case 0, 1 'tipos de movimiento
            txtNSerie(Index).Text = DevuelveDesdeBD("nomregis", "contadores", "tiporegi", txtSerie(Index), "T")
    End Select
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

End Sub


Private Sub txtNumFac_GotFocus(Index As Integer)
    ConseguirFoco txtNumFac(Index), 3
End Sub

Private Sub txtNumFac_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtNumFac_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtNumFac_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtNumFac(Index).Text = UCase(Trim(txtNumFac(Index).Text))
    
    
    Select Case Index
        Case 0, 1 'numero de factura
            PonerFormatoEntero txtNumFac(Index)
    End Select
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

End Sub


Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtImporte(Index).Text = UCase(Trim(txtImporte(Index).Text))
    
    Select Case Index
        Case 0, 1 'importe de vencimiento
            PonerFormatoEntero txtImporte(Index)
    End Select
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

End Sub


Private Function VencimientosPorEntidadBancaria() As String
Dim SQL As String

    VencimientosPorEntidadBancaria = ""

    SQL = " and length(cobros.iban) <> 0 and mid(cobros.iban,5,4) = (select mid(iban,5,4) from bancos where codmacta = " & DBSet(txtCuentas(2).Text, "T") & ")"
    
    VencimientosPorEntidadBancaria = SQL

End Function


Private Function GenerarTransferencia(Opcion As Integer) As Boolean
Dim C As String
Dim NumeroRemesa As Long
Dim Rs As ADODB.Recordset
Dim J As Integer
Dim i As Integer
Dim ImporteQueda As Currency

    On Error GoTo eGenerarTransferencia
    
    GenerarTransferencia = False
    
    
    'EMBARGO
    Cad = ""
    For i = 1 To Me.lwCobros.ListItems.Count
        If lwCobros.ListItems(i).Checked Then
            If InStr(1, Cad, lwCobros.ListItems(i).SubItems(8)) = 0 Then
                Cad = Cad & ", " & DBSet(lwCobros.ListItems(i).SubItems(8), "T")
            End If
        End If
    Next i

    If Cad <> "" Then
        'No deberia ser ""
        Cad = Mid(Cad, 2)
        Set miRsAux = New ADODB.Recordset
        Cad = "Select codmacta,nommacta from cuentas where embargo=1 AND codmacta in (" & Cad & ")"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not miRsAux.EOF
            Cad = Cad & miRsAux!codmacta & " " & miRsAux!Nommacta & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        If Cad <> "" Then
            MsgBox "Cuentas en situacion de embargo:" & vbCrLf & Cad, vbExclamation
            Exit Function
        End If
    End If
    
    
    
    
    'Lo qu vamos a hacer es, primero bloquear la opcion de remesar
    If Opcion = 0 Then
        If Not BloqueoManual(True, "Transferencias", "Transferencias") Then
            MsgBox "Otro usuario esta generando transferencias", vbExclamation
            Exit Function
        End If
    End If
    
    i = 0
    For J = 1 To lwCobros.ListItems.Count
        If lwCobros.ListItems(J).Checked Then
            i = J
            Exit For
        End If
    Next J
    If i = 0 Then
        If TipoTrans = 0 Then
            MsgBox "No se ha seleccionado cobros. Revise.", vbExclamation
        Else
            MsgBox "No se ha seleccionado pagos. Revise.", vbExclamation
        End If
        If Opcion = 0 Then BloqueoManual False, "Transferencias", ""
        Exit Function
    End If
    
    
    'A partir de la fecha generemos leemos k remesa corresponde
    If Opcion = 0 Then
        SQL = "select max(codigo) from transferencias where anyo=" & Year(CDate(txtfecha(4).Text))
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        NumeroRemesa = 0
        If Not miRsAux.EOF Then
            NumeroRemesa = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        
        
        NumeroRemesa = NumeroRemesa + 1
    Else
        NumeroRemesa = lw1.SelectedItem.Text
        txtfecha(4).Text = lw1.SelectedItem.SubItems(2)
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    Conn.BeginTrans
    
    
    Set Rs = New ADODB.Recordset
    Cad = "Select * from tmpcierre1 where codusu =" & vUsu.Codigo
    If CadenaDesdeOtroForm <> "" Then Cad = Cad & " and cta in (" & CadenaDesdeOtroForm & ")"
    
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Error grave. Datos temporales vacios", vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Function
    End If
    
    
    Set miRsAux = New ADODB.Recordset
    
    'Para ver si existe la transferencia... pero esto no tendria k pasar
    '------------------------------------------------------------
    
    Label2.Caption = ""
    Label2.visible = True
    
    While Not Rs.EOF
    
        Select Case TipoTrans
            Case 0, 1
                Label2.Caption = "Generando Transferencia " & NumeroRemesa & " del banco " & Rs!Cta
            Case 2
                Label2.Caption = "Generando Pago Domiciliado " & NumeroRemesa & " del banco " & Rs!Cta
            Case 3
                Label2.Caption = "Generando Confirming " & NumeroRemesa & " del banco " & Rs!Cta
        End Select
        Me.Refresh
        DoEvents
    
    
        If Opcion = 0 Then
            'Ahora insertamos la remesa
            Cad = "INSERT INTO transferencias (tipotrans, codigo, anyo, fecha,situacion,codmacta,descripcion,subtipo,concepto) "
            If TipoTrans = 0 Then
                Cad = Cad & " VALUES (1,"
            Else
                Cad = Cad & " VALUES (0,"
            End If
            Cad = Cad & NumeroRemesa & "," & Year(CDate(txtfecha(4).Text)) & ",'" & Format(txtfecha(4).Text, FormatoFecha) & "','A','"
            Cad = Cad & Rs!Cta & "','" & DevNombreSQL(txtRemesa.Text) & "',"
            Select Case TipoTrans
                Case 0, 1 ' transf.abonos y transf.pagos
                    Cad = Cad & "0,"
                    Cad = Cad & DBSet(cboConcepto2.ItemData(cboConcepto.ListIndex), "N") & ")"
                Case 2 ' transf.pagos domiciliados
                    Cad = Cad & "1,"
                    If chkFecha.Value = 0 Then
                        Cad = Cad & "0)"
                    Else
                        Cad = Cad & "1)"
                    End If
                Case 3 ' pagos confirming
                    Cad = Cad & "2,"
                    If chkFecha.Value = 0 Then
                        Cad = Cad & "0)"
                    Else
                        Cad = Cad & "1)"
                    End If
            End Select
                
            Conn.Execute Cad
            
        Else
            'Paso la remesa a estado: A
            'Vuelvo a poner los vecnimientos a NULL para poder
            'meterlos luego
            
            '---remesa estado A
            
            Cad = "UPDATE transferencias SET Situacion = 'A'"
            Cad = Cad & ", descripcion ='" & DevNombreSQL(Text2.Text) & "'"
            Cad = Cad & ", fecha= " & DBSet(txtfecha(5).Text, "F")
            Cad = Cad & ", codmacta= " & DBSet(txtCuentas(3).Text, "T")
            Cad = Cad & ", concepto = " & DBSet(cboConcepto2.ItemData(cboConcepto2.ListIndex), "N")
            Cad = Cad & " WHERE codigo=" & NumeroRemesa
            Cad = Cad & " AND anyo =" & Year(CDate(txtfecha(4).Text))
            If Not Ejecuta(Cad) Then Exit Function
            
            
            If TipoTrans = 0 Then ' transferencias de abonos
                Cad = "UPDATE cobros SET siturem=NULL, transfer=NULL, anyorem=NULL ,tiporem =NULL "
                Cad = Cad & " ,fecultco=NULL, impcobro = NULL "
                Cad = Cad & " WHERE transfer = " & NumeroRemesa
                Cad = Cad & " AND anyorem=" & Year(CDate(txtfecha(4).Text)) '& " AND tiporem = 1"
                If Not Ejecuta(Cad) Then Exit Function
            Else
                ' transferencias de pagos
                Cad = "UPDATE pagos SET situdocum=NULL,nrodocum=NULL,anyodocum=NULL "
                Cad = Cad & " ,fecultpa=NULL, imppagad = NULL "
                Cad = Cad & " WHERE nrodocum = " & NumeroRemesa
                Cad = Cad & " AND anyodocum=" & Year(CDate(txtfecha(4).Text))
                If Not Ejecuta(Cad) Then Exit Function
            
            End If
        End If
        
        'Ahora cambiamos los cobros y les ponemos la remesa
        If TipoTrans = 0 Then ' transferencias de abonos
            If Opcion = 0 Then
                Cad = "UPDATE cobros SET siturem = 'A', transfer= " & NumeroRemesa & ", anyorem =" & Year(CDate(txtfecha(4).Text))
            Else
                Cad = "UPDATE cobros SET siturem = 'A', transfer= " & NumeroRemesa & ", anyorem =" & Year(CDate(txtfecha(5).Text))
            End If
        
            'Para cada cobro UPDATE
            For J = 1 To lwCobros.ListItems.Count
               With lwCobros.ListItems(J)
                    If .Checked And .SubItems(7) = Rs!Cta Then   ' si el subitem es del banco
                        C = " WHERE numserie = '" & .Text & "' and numfactu = "
                        C = C & Val(.SubItems(1)) & " and fecfactu ='" & Format(.SubItems(2), FormatoFecha)
                        C = C & "' AND numorden =" & .SubItems(3)
                        C = Cad & C
                        Conn.Execute C
                    Else
                   
                    End If
               End With
            Next J
            espera 0.5
            
            If Opcion = 0 Then
                'Hacemos un select sum para el importe
                Cad = "Select sum(impvenci),sum(coalesce(impcobro,0)),sum(coalesce(gastos,0)) from cobros "
                Cad = Cad & " WHERE transfer=" & NumeroRemesa
                Cad = Cad & " AND anyorem =" & Year(CDate(txtfecha(4).Text))
                
                miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                C = "0"
                If Not miRsAux.EOF Then
                    If Not IsNull(miRsAux.Fields(0)) Then
                                       'Impvenci                               impcobro                      gastos
                        ImporteQueda = DBLet(miRsAux.Fields(0), "N") - DBLet(miRsAux.Fields(1), "N") + DBLet(miRsAux.Fields(2), "N")
                        C = TransformaComasPuntos(CStr(ImporteQueda))
                    End If
                End If
                miRsAux.Close
            End If
        Else ' transferencias de pagos
        
            If Opcion = 0 Then
                Cad = "UPDATE pagos SET situdocum = 'A', nrodocum= " & NumeroRemesa & ", anyodocum =" & Year(CDate(txtfecha(4).Text))
                'Pagos domiciliados, pude pagar por otro banco. Ponemos el banco seleccionado aqui
                If TipoTrans = 2 Then Cad = Cad & ", ctabanc1 ='" & txtCuentas(2).Text & "', emitdocum = 1"
            Else
                Cad = "UPDATE pagos SET situdocum = 'A', nrodocum= " & NumeroRemesa & ", anyodocum =" & Year(CDate(txtfecha(5).Text))
            End If
        
            'Para cada cobro UPDATE
            For J = 1 To lwCobros.ListItems.Count
               With lwCobros.ListItems(J)
                    If .Checked And .SubItems(7) = Rs!Cta Then   ' si el subitem es del banco
                        C = " WHERE numserie = '" & .Text & "' and numfactu = "
                        C = C & DBSet(.SubItems(1), "T") & " and fecfactu ='" & Format(.SubItems(2), FormatoFecha)
                        C = C & "' AND numorden =" & .SubItems(3)
                        C = C & " and codmacta = " & DBSet(.SubItems(8), "T")
                    
                        C = Cad & C
                        Conn.Execute C
                    Else
     
                    End If
               End With
            Next J
            espera 0.5
            
            If Opcion = 0 Then
                'Hacemos un select sum para el importe
                Cad = "Select sum(impefect),sum(coalesce(imppagad,0)) from pagos "
                Cad = Cad & " WHERE nrodocum=" & NumeroRemesa
                Cad = Cad & " AND anyodocum =" & Year(CDate(txtfecha(4).Text))
                
                miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                C = "0"
                If Not miRsAux.EOF Then
                    If Not IsNull(miRsAux.Fields(0)) Then
                                       'Impvenci                               impcobro
                        ImporteQueda = DBLet(miRsAux.Fields(0), "N") - DBLet(miRsAux.Fields(1), "N")
                        C = TransformaComasPuntos(CStr(ImporteQueda))
                    End If
                End If
                miRsAux.Close
            
            
            
            
            
            End If
        End If
        
        If Opcion = 0 Then
            Cad = "UPDATE transferencias SET importe=" & C
            Cad = Cad & " WHERE codigo=" & NumeroRemesa
            Cad = Cad & " AND anyo =" & Year(CDate(txtfecha(4).Text))
            Conn.Execute Cad
        End If
        
        NumeroRemesa = NumeroRemesa + 1
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    Set miRsAux = Nothing
    
    GenerarTransferencia = True
    Conn.CommitTrans
    If Opcion = 0 Then BloqueoManual False, "Transferencias", "Transferencias"
    
    Label2.Caption = ""
    Label2.visible = False
    
    Exit Function
    
eGenerarTransferencia:
    Conn.RollbackTrans
    Select Case TipoTrans
        Case 0, 1
            MuestraError Err.Number, "Generar transferencia", Err.Description
        Case 2
            MuestraError Err.Number, "Generar Pago Domiciliado", Err.Description
        Case 3
            MuestraError Err.Number, "Generar Confirming", Err.Description
    End Select
    If Opcion = 0 Then BloqueoManual False, "Transferencias", "Transferencias"

    Label2.Caption = ""
    Label2.visible = False
End Function






' -1: ERROR     0: Ninguno     1 Si que hay agrupaciones
Private Function ComprobacionesAgrupaFicheroTransfer(Abonos As Boolean) As Integer
Dim C As String

    On Error GoTo eComrpobacionesAgrupacionFichero
    Set miRsAux = New ADODB.Recordset
    ComprobacionesAgrupaFicheroTransfer = -1
    
   
    
    If Abonos Then
        SQL = "select codmacta,count(*) from cobros where"
        SQL = SQL & " transfer = " & lw1.ListItems(lw1.SelectedItem.Index).Text
        SQL = SQL & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)

    Else
        SQL = "select codmacta,count(*) from pagos where nrodocum = " & lw1.ListItems(lw1.SelectedItem.Index).Text
        SQL = SQL & " and anyodocum = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    End If
    SQL = SQL & " group by codmacta having count(*) >1"
     miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Msg = ""
    RC = ""
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        Msg = Msg & miRsAux!codmacta & "|"
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    
    If i = 0 Then
        ComprobacionesAgrupaFicheroTransfer = 0
        Exit Function
    End If
    
    'Comprobare la cuenta IBAN
    MsgErr = ""
    K = 0
    NumRegElim = 0
    For J = 1 To i
        SQL = " AND codmacta = '" & RecuperaValor(Msg, CInt(J)) & "' GROUP BY iban "
        If Abonos Then
            SQL = " transfer = " & lw1.ListItems(lw1.SelectedItem.Index).Text & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1) & SQL
            SQL = "select  iban,count(*) ctos from cobros where  " & SQL
            
        Else
            SQL = " nrodocum = " & lw1.ListItems(lw1.SelectedItem.Index).Text & " and anyodocum = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1) & SQL
            SQL = "select iban,count(*) ctos from pagos where  " & SQL
        End If
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If miRsAux.EOF Then
            Err.Raise 513, , "no se encuentran vtos. en select: " & Mid(miRsAux.Source, 30)
        Else
            C = DevuelveDesdeBD("concat(codmacta,' ',nommacta)", "cuentas", "codmacta", RecuperaValor(Msg, CInt(J)), "T") & "    Nº:" & Format(miRsAux!ctos, "0000")
            SQL = miRsAux!IBAN
            miRsAux.MoveNext
            If miRsAux.EOF Then
                'perfeto. SOLO un IBAN
                RC = RC & "X"
                NumRegElim = NumRegElim + Val(Right(C, 4))
                SQL = ""
            Else
                SQL = C & vbCrLf & SQL & " // " & miRsAux!IBAN & vbCrLf
            End If
        End If
        If SQL <> "" Then MsgErr = MsgErr & SQL & vbCrLf
        miRsAux.Close
    Next J
    
    
    Set miRsAux = Nothing
    
    If MsgErr <> "" Then
        MsgErr = "IBAN distinto para agrupacion por cliente:" & vbCrLf & MsgErr
        MsgBox MsgErr, vbExclamation
        
    Else
        
        SQL = "Fecha de pago: " & RecuperaValor(CadenaDesdeOtroForm, 1) & vbCrLf
        RC = IIf(TipoTrans = 0, "Clientes", "Proveedores") & " agrupados: " & Len(RC) & vbCrLf
        RC = RC & "Total vencimientos agrupados: " & NumRegElim & vbCrLf
        SQL = SQL & RC & " ¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then ComprobacionesAgrupaFicheroTransfer = i
    End If
eComrpobacionesAgrupacionFichero:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    RC = ""
End Function

