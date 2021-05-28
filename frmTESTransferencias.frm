VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESTransferencias 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencias"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16035
   Icon            =   "frmTESTransferencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   16035
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   11640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCreacionRemesa 
      BorderStyle     =   0  'None
      Height          =   9285
      Left            =   30
      TabIndex        =   25
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
         Left            =   13680
         TabIndex        =   63
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   4755
         Left            =   120
         TabIndex        =   32
         Top             =   3840
         Width           =   15495
         Begin MSComctlLib.ListView lwCobros 
            Height          =   4095
            Left            =   0
            TabIndex        =   33
            Top             =   600
            Width           =   15375
            _ExtentX        =   27120
            _ExtentY        =   7223
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
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
            Left            =   11880
            TabIndex        =   83
            Top             =   180
            Width           =   1335
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   480
            Picture         =   "frmTESTransferencias.frx":000C
            ToolTipText     =   "Seleccionar"
            Top             =   180
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "frmTESTransferencias.frx":0156
            ToolTipText     =   "Quitar seleccion"
            Top             =   180
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   240
         TabIndex        =   30
         Top             =   8640
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
            TabIndex        =   31
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
         TabIndex        =   22
         Top             =   8760
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
         TabIndex        =   24
         Top             =   8760
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
         TabIndex        =   35
         Top             =   60
         Width           =   15645
         Begin VB.CheckBox chkIncluirAbonos 
            Caption         =   "Incluir abonos"
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
            TabIndex        =   18
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   1260
            Width           =   3195
         End
         Begin VB.CheckBox chkCompensarAbonos 
            Caption         =   "Compensar abonos"
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
            TabIndex        =   19
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   1680
            Width           =   3195
         End
         Begin VB.CheckBox chkConfirmingPP 
            Caption         =   "Confirming pronto pago"
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
            TabIndex        =   23
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   840
            Width           =   3195
         End
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
            TabIndex        =   21
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   2520
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
            TabIndex        =   20
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   2100
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
            TabIndex        =   17
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
            TabIndex        =   61
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
            BeginProperty Font 
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
            BeginProperty Font 
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   77
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
            TabIndex        =   62
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
            TabIndex        =   60
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
         TabIndex        =   64
         Top             =   60
         Width           =   15645
         Begin VB.CheckBox chkConfirmingPP2 
            Caption         =   "Confirming pronto pago"
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
            TabIndex        =   80
            Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
            Top             =   2160
            Width           =   3195
         End
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
            TabIndex        =   79
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
            TabIndex        =   68
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
            TabIndex        =   66
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
            TabIndex        =   67
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
            TabIndex        =   69
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
            TabIndex        =   65
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
            TabIndex        =   78
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
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
            Top             =   1860
            Width           =   1245
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   3
            Left            =   1200
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
            TabIndex        =   70
            Top             =   1860
            Width           =   1005
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
         TabIndex        =   75
         Top             =   8550
         Width           =   8400
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   30
      TabIndex        =   26
      Top             =   30
      Visible         =   0   'False
      Width           =   15915
      Begin VB.Frame FrameFiltro 
         Height          =   705
         Left            =   5940
         TabIndex        =   81
         Top             =   180
         Width           =   2445
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
            ItemData        =   "frmTESTransferencias.frx":02FA
            Left            =   120
            List            =   "frmTESTransferencias.frx":0307
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   210
            Width           =   2235
         End
      End
      Begin VB.Frame FrameBotonGnral2 
         Height          =   705
         Left            =   4020
         TabIndex        =   34
         Top             =   180
         Width           =   1725
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   180
            TabIndex        =   76
            Top             =   180
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Grabación Fichero"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Cargo Transferencia"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Devolucion"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   27
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
         Height          =   7905
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   15525
         _ExtentX        =   27384
         _ExtentY        =   13944
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
         TabIndex        =   29
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


Public TipoTrans2 As Integer  '0 = transferencia desde abonos
                             '1 = transferencia desde pagos
                             '2 = transferencia de pagos domiciliados o bancarios
                             '3 = transferencia de confirming
                             '4 = Anticipar recibos




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

Dim Sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim cad As String
Dim CONT As Long
Dim I As Integer
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

Private Sub cboConcepto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboConcepto2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    CargaList
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkAbonosSoloTransferencia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCompensarAbonos_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub chkConfirmingPP_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub chkConfirmingPP2_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub chkFecha_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub chkFecha2_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
Dim I As Integer
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
            If I >= 0 Then lw1.SetFocus
    
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
                        If TipoTrans2 = 0 Or TipoTrans2 = 4 Then ' transferencia de abonos
                            NuevaTransf
                        Else
                            NuevaTransfPagos
                        End If
                    Else
                        If GenerarTransferencia(0) Then
                            Select Case TipoTrans2
                                Case 0, 1
                                    MsgBox "Transferencia generada correctamente.", vbExclamation
                                Case 2
                                    MsgBox "Pago domiciliado generado correctamente.", vbExclamation
                                Case 3
                                    MsgBox "Confirming generado correctamente.", vbExclamation
                                Case 4
                                    MsgBox "Proceso generado correctamente.", vbExclamation
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
            If TipoTrans2 < 2 Then
                Sql = "select count(*) from bancos where codmacta = " & DBSet(txtCuentas(2).Text, "T") & " and not sufijoem is null and sufijoem <> ''"
                If TotalRegistros(Sql) = 0 Then
                    MsgBox "El banco no tiene Sufijo Transferencia. Reintroduzca.", vbExclamation
                    PonleFoco txtCuentas(2)
                    Exit Function
                End If
            End If
        End If
    
        'Fecha transferencia tiene k tener valor
        If txtFecha(4).Text = "" Then
            MsgBox "Fecha de " & nTipo & " debe tener valor", vbExclamation
            PonFoco txtFecha(4)
            Exit Function
        End If
        
        'Confirming
        If TipoTrans2 = 3 Then
        
            'NO HACE FALTA configurar la cuenta confirmnng.
            'Si no la tiene el preoceso de generacion NO updatea  la codmacta
        
            'El banco tienen que tener confuigurada la cuenta confirming
            'SQL = DevuelveDesdeBD("ctaConfirming", "bancos", "codmacta", txtCuentas(2).Text, "T")
            'If SQL = "" Then
            '    MsgBox "No esta bien configurado el banco. Falta cuenta confirming", vbExclamation
            '    Exit Function
            'End If
            
            
            If Me.chkCompensarAbonos.Value = 1 And Me.chkIncluirAbonos.Value = 1 Then
                MsgBox "Si elige incluir abonos no puede realizar la compensacion", vbExclamation
                Exit Function
            End If
            
        End If
        
        'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
        If FechaCorrecta2(CDate(txtFecha(4).Text), True) > 1 Then
            PonFoco txtFecha(4)
            Exit Function
        End If
        
        
        If TipoTrans2 < 2 Then
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
        If txtFecha(5).Text = "" Then
            MsgBox "Fecha de " & nTipo & " debe tener valor", vbExclamation
            PonFoco txtFecha(5)
            Exit Function
        Else
            If Year(CDate(txtFecha(5).Text)) <> lw1.SelectedItem.SubItems(1) Then
                MsgBox "La fecha de " & nTipo & " ha de ser del mismo año. Revise.", vbExclamation
                PonFoco txtFecha(5)
                Exit Function
            End If
        End If
        
        'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
        If FechaCorrecta2(CDate(txtFecha(5).Text), True) > 1 Then
            PonFoco txtFecha(5)
            Exit Function
        End If
    
        If TipoTrans2 < 2 Then
            If Me.cboConcepto2.ListIndex = -1 Then
                MsgBox "Debe introducir un Concepto.", vbExclamation
                PonerFocoCmb cboConcepto2
                Exit Function
            End If
        End If
    
    
        'Sept 2017
        'Si hay algun "NO" no sigue
        
    End If
    
        'Comprobar si permite NEGATIVOS en confirming, que el resultado de cada proveedor es POSITIVO
        If TipoTrans2 = 3 And chkIncluirAbonos.Value = 1 Then
                    
            'Cuentas que van al confirming
            C = ""
            B = False
            For J = 1 To Me.lwCobros.ListItems.Count
                If lwCobros.ListItems(J).Checked Then
                    If InStr(1, lwCobros.ListItems(J).SubItems(7), "-") > 0 Then
                        B = True
                        Msg = lwCobros.ListItems(J).SubItems(14) & "|"
                        If InStr(1, C, Msg) = 0 Then C = C & lwCobros.ListItems(J).SubItems(14) & "|"
                    End If
                End If
            Next

            'Para cada cuenta , el importe final.
            If B Then 'ha habido algun abonno
                RC = ""
                Do
                    I = InStr(1, C, "|")
                    If I = 0 Then
                        C = 0
                    Else
                        Msg = Mid(C, 1, I - 1)
                        C = Mid(C, I + 1)
                        Importe = 0
                        For J = 1 To Me.lwCobros.ListItems.Count
                            If lwCobros.ListItems(J).Checked Then
                                If lwCobros.ListItems(J).SubItems(14) = Msg Then Importe = Importe + ImporteFormateado(lwCobros.ListItems(J).SubItems(7))
                            End If
                        Next
                        
                        If Importe <= 0 Then
                            Msg = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Msg, "T")
                            RC = RC & "     - " & Msg & " (" & Format(Importe, FormatoImporte) & ")" & vbCrLf
                        End If
                    End If
                Loop Until C = ""
                If RC <> "" Then
                    RC = "Proveedores con importe final menor o igual a cero." & vbCrLf & vbCrLf & RC
                    MsgBoxA RC, vbExclamation
                    Exit Function
                End If
            End If
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
            If TipoTrans2 <> 4 Then
                If MsgBox(C, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            End If
        End If
    
    DatosOK = True

End Function


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
    
    For I = 0 To 1
        Me.imgSerie(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgCuentas(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
    Me.imgCuentas(2).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(3).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    For I = 0 To 5
        Me.ImgFec(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next I

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
        .Buttons(3).Image = 45
    End With
    
    'En pagos domiciliados, NO contabiliza
    If TipoTrans2 = 2 Then Toolbar2.Buttons(2).Enabled = False
    Toolbar2.Buttons(3).visible = TipoTrans2 <= 1
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
        
     CargaFiltrosEjer Me.cboFiltro
    
    chkConfirmingPP.visible = False
    chkConfirmingPP.Enabled = False
    chkConfirmingPP2.visible = False
    chkConfirmingPP2.Enabled = False
    chkCompensarAbonos.visible = True
    chkCompensarAbonos.Enabled = True
    chkIncluirAbonos.visible = False
    chkIncluirAbonos.Enabled = False
    Select Case TipoTrans2
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
            chkConfirmingPP.visible = False
            
            
            Label3(11).Caption = "Cuenta Cliente"
            Label3(18).Caption = "Fecha Vencimiento"
            Label3(14).Caption = "Importe Vencimiento"
            Me.chkCompensarAbonos.visible = False
            chkCompensarAbonos.Enabled = False
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
            chkConfirmingPP.visible = False
            
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
            chkFecha.Enabled = False
            chkFecha.visible = False
            chkFecha2.Enabled = False
            chkFecha2.visible = False
            chkConfirmingPP.visible = True
            chkConfirmingPP.Enabled = True
            chkConfirmingPP2.visible = True
            chkConfirmingPP2.Enabled = True
            chkIncluirAbonos.visible = True
            chkIncluirAbonos.Enabled = True
            
            
            Label3(11).Caption = "Cuenta Proveedor"
            Label3(18).Caption = "Fecha Efecto"
            Label3(14).Caption = "Importe Efecto"
            
        Case 4
            Me.Caption = "Anticipo recibos"
            IdPrograma = 614 ' transf. de abonos
            nTipo = "transferencia"
                        
            'concepto
            Label3(13).Enabled = True
            Label3(13).visible = False
            cboConcepto.Enabled = False
            cboConcepto.visible = False
            Label3(15).Enabled = True
            Label3(15).visible = True
            cboConcepto2.Enabled = True
            cboConcepto2.visible = True
            'fecha introducida
            chkFecha.Enabled = False
            chkFecha.visible = False
            chkFecha2.Enabled = False
            chkFecha2.visible = False
            chkConfirmingPP.visible = False
            
            
            Label3(11).Caption = "Cuenta Cliente"
            Label3(18).Caption = "Fecha Vencimiento"
            Label3(14).Caption = "Importe Vencimiento"
            Me.chkCompensarAbonos.visible = False
            chkCompensarAbonos.Enabled = False
            
            
    End Select
    
   
    
    Me.Toolbar2.Buttons(2).ToolTipText = "Cargo " & UCase(Me.Caption)
    If TipoTrans2 < 2 Then cboConcepto.ListIndex = 1
    chkAbonosSoloTransferencia.visible = IIf(TipoTrans2 = 0, True, False)
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
    
    If TipoTrans2 = 0 Then
        'Abonos.
        'Numero factura normal
        'Cliente mas grande
        J = 1200
        I = 5300
    Else
        'Numero factura normal
        'Cliente mas grande
        J = 1800
        I = 4800
        
    End If
    I = I - 250 'de las observaciones
    
    lwCobros.ColumnHeaders.Clear
    lwCobros.ColumnHeaders.Add , , "Serie", 800
    lwCobros.ColumnHeaders.Add , , "Factura", J
    lwCobros.ColumnHeaders.Add , , "Fecha", 1450
    
    If TipoTrans2 = 0 Or TipoTrans2 = 4 Then
        lwCobros.ColumnHeaders.Add , , "Vto", 699
        lwCobros.ColumnHeaders.Add , , "Fecha Vto", 1350
        lwCobros.ColumnHeaders.Add , , "Cliente", I
    Else
        lwCobros.ColumnHeaders.Add , , "Efec", 699
        lwCobros.ColumnHeaders.Add , , "Fecha Efec", 1350
        lwCobros.ColumnHeaders.Add , , "Proveedor", I
    End If
    
    
    lwCobros.ColumnHeaders.Add , , "Tipo", 1535
    
    
    lwCobros.ColumnHeaders.Add , , "Importe", 2035, 1
    lwCobros.ColumnHeaders.Add , , "ENTIDAD", 0, 1
    lwCobros.ColumnHeaders.Add , , "CtaProve", 0, 1

    'Para la ordenacion
    lwCobros.ColumnHeaders.Add , , "SeriFacord", 0, 1
    lwCobros.ColumnHeaders.Add , , "Fecfacord", 0, 1
    lwCobros.ColumnHeaders.Add , , "fecefecord", 0, 1
    lwCobros.ColumnHeaders.Add , , "Importeorden", 0, 1
    
    
    lwCobros.ColumnHeaders.Add , , "codmacta", 0, 1
    
    lwCobros.ColumnHeaders.Add , , "Obs", 600, 1
    
    
    lwCobros.SortKey = 11
    lwCobros.SortOrder = lvwAscending
    lwCobros.Sorted = True
    
    
End Sub


Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then Sql = CadenaSeleccion
        
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    
    Sql = CadenaSeleccion
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
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
Dim I As Integer
Dim Im2 As Currency

    Screen.MousePointer = vbHourglass
    
    cad = ""
    Im2 = 0
    For I = 1 To Me.lwCobros.ListItems.Count
        If lwCobros.ListItems(I).Selected Then
            cad = cad & "X"
            Im2 = Im2 + lwCobros.ListItems(I).SubItems(7)
        End If
    Next I
    If Len(cad) > 1 Then
        'Va a realizar la accion sobre  len(cad) vencimientos
        Sql = "Va a " & IIf(Index = 1, "seleccionar", "quitar la seleccion") & ":"
        Sql = Sql & vbCrLf & "Vencimientos: " & Len(cad) & vbCrLf & ""
        'msgbox
        For I = 1 To Me.lwCobros.ListItems.Count
            If lwCobros.ListItems(I).Selected Then
                If Index = 1 Then
                    If lwCobros.ListItems(I).ListSubItems(3).Tag <> "EMB" Then lwCobros.ListItems(I).Checked = True
                Else
                    lwCobros.ListItems(I).Checked = False
                End If
            End If
        Next I
    
    
    Else
        
        For I = 1 To Me.lwCobros.ListItems.Count
            If Index = 1 Then
                If lwCobros.ListItems(I).ListSubItems(3).Tag <> "EMB" Then lwCobros.ListItems(I).Checked = True
            Else
                lwCobros.ListItems(I).Checked = False
            End If
        Next I
    End If
    
    
    
    'El importe
    Importe = 0
    For I = 1 To Me.lwCobros.ListItems.Count
        'lwCobros2.ListItems(I).Checked = (Index = 1)
        If lwCobros.ListItems(I).Checked Then Importe = Importe + lwCobros.ListItems(I).SubItems(7)
        
        'If Index = 1 Then Importe = Importe + lwCobros2.ListItems(I).SubItems(6)
    Next I
    Text1(4).Tag = Importe
    If Importe <> 0 Then
        Text1(4).Text = Format(Importe, "###,###,##0.00")
    Else
        Text1(4).Text = ""
    End If
    If Im2 <> 0 Then PonleFoco lwCobros
    Screen.MousePointer = vbDefault




End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFecha_Click(Index As Integer)
    'FECHA FACTURA
    Indice = Index
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtFecha(Indice).Text <> "" Then frmF.Fecha = CDate(txtFecha(Indice).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtFecha(Indice)

End Sub

Private Sub imgCuentas_Click(Index As Integer)
    
    If Index = 2 Or Index = 3 Then
    
        Sql = ""
        Set frmBan = New frmBasico2
        AyudaBanco frmBan
        Set frmBan = Nothing
        If Sql <> "" Then
            txtCuentas(Index).Text = RecuperaValor(Sql, 1)
            txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
        End If
    
    Else
        Sql = ""
        AbiertoOtroFormEnListado = True
        Set frmCtas = New frmColCtas
        frmCtas.DatosADevolverBusqueda = True
        frmCtas.Show vbModal
        Set frmCtas = Nothing
        If Sql <> "" Then
            Me.txtCuentas(Index).Text = RecuperaValor(Sql, 1)
            Me.txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
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
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgSerie_Click(Index As Integer)
    IndCodigo = Index
    Sql = ""
    Set frmConta = New frmBasico
    AyudaContadores frmConta, txtSerie(Index), "tiporegi REGEXP '^[0-9]+$' = 0"
    Set frmConta = Nothing
    If Sql <> "" Then
        txtSerie(Index).Text = RecuperaValor(Sql, 1)
        txtNSerie(Index).Text = RecuperaValor(Sql, 2)
    End If
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
    
    If TipoTrans2 = 0 Then ' transferencias de abonos
        frmMens.Opcion = 55
    ElseIf TipoTrans2 = 3 Then ' confirming
        frmMens.Opcion = 66
    ElseIf TipoTrans2 = 2 Then ' Pago domiciliado
        frmMens.Opcion = 67
    ElseIf TipoTrans2 = 4 Then ' Recibo anticipado
        frmMens.Opcion = 68
    Else
        'transferencias de abonos
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

Private Sub lwCobros_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim co As Integer

    'Reordenar
    
    
     co = ColumnHeader.Index - 1
     If co = 4 Then co = 12 'fec vto
     If co = 0 Then co = 10 ' serie
     If co = 7 Then co = 13 'importe
     If co = 2 Then co = 11 'fecha factura
   
    If lwCobros.SortKey = co Then
        If lwCobros.SortOrder = lvwAscending Then
            lwCobros.SortOrder = lvwDescending
        Else
            lwCobros.SortOrder = lvwAscending
        End If
    Else
        lwCobros.SortKey = co
        lwCobros.SortOrder = lvwAscending
    End If
End Sub

Private Sub lwCobros_DblClick()
 If lwCobros.ListItems.Count > 0 Then
        If Not lwCobros.SelectedItem Is Nothing Then
            
            If lwCobros.SelectedItem.SubItems(15) = "*" Then
            
                If TipoTrans2 = 0 Or TipoTrans2 = 4 Then ' transferencia de abonos
                    Sql = "numserie ='" & lwCobros.SelectedItem.Text & "' AND numfactu =" & lwCobros.SelectedItem.SubItems(1)
                    Sql = Sql & " AND fecfactu =" & DBSet(lwCobros.SelectedItem.SubItems(2), "F") & " AND numorden "
                    Sql = DevuelveDesdeBD("observa", "cobros", Sql, lwCobros.SelectedItem.SubItems(3))
                Else
                    Sql = "numserie ='" & lwCobros.SelectedItem.Text & "' AND numfactu =" & lwCobros.SelectedItem.SubItems(1)
                    Sql = Sql & " AND fecfactu =" & DBSet(lwCobros.SelectedItem.SubItems(2), "F") & " AND numorden "
                    Sql = DevuelveDesdeBD("observa", "pagos", Sql, lwCobros.SelectedItem.SubItems(3))
                End If
                
                frmZoom.pValor = Sql
                frmZoom.pModo = 2
                frmZoom.Caption = "Observaciones cobro"
                frmZoom.Show vbModal
                
            End If
        End If
    End If
End Sub

Private Sub lwCobros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'Dim C As Currency
'Dim Cobro As Boolean

 '   Cobro = True
 '   C = Item.Tag
    
    If Item.Checked Then
        If Item.ListSubItems(3).Tag = "EMB" Then Item.Checked = False
    End If
    
    
    Importe = 0
    For I = 1 To lwCobros.ListItems.Count
        If lwCobros.ListItems(I).Checked Then Importe = Importe + lwCobros.ListItems(I).SubItems(7)
    Next I
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
            
            frmTESTransList.EsTransfAbonos = (TipoTrans2 = 0 Or TipoTrans2 = 4)
            If TipoTrans2 > 0 Then
                frmTESTransList.TipoTransfPagos = (TipoTrans2 - 1) '3
            Else
                'si son abonos me da igual este parametro
                frmTESTransList.TipoTransfPagos = 0
            End If
                    
            frmTESTransList.Show vbModal
            
            Set frmTESTransList = Nothing

    End Select
    
End Sub

Private Function SepuedeBorrar() As Boolean
Dim Sql As String
    
    SepuedeBorrar = False

    If lw1.SelectedItem.SubItems(9) = "Q" Then
        MsgBox "No se pueden modificar ni eliminar " & nTipo & "s en situación abonada.", vbExclamation
        Exit Function
    End If
    
    
    If TipoTrans2 = 2 Then
        'Veremos si algun vencimiento para ese Pago domiciliado ha sido contabilizado.
        
        Sql = "nrodocum = " & lw1.SelectedItem.Text & " And imppagad > 0 And anyodocum"
        Sql = DevuelveDesdeBD("count(*)", "pagos", Sql, lw1.SelectedItem.SubItems(1))
        If Val(Sql) > 0 Then
            Sql = "Ya ha contabilzado pagos de este documento"
            
            If vUsu.Nivel > 0 Then
                MsgBoxA Sql, vbExclamation
                Exit Function
            Else
                Sql = Sql & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBoxA(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            End If
        End If
    End If
    
    
    SepuedeBorrar = True

End Function


Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    
    If Me.lw1.SelectedItem Is Nothing Then Exit Sub
    If Me.lw1.SelectedItem = "" Then Exit Sub
        
    If Not SepuedeBorrar Then Exit Sub
        
        
    '*************** canviar els noms i el DELETE **********************************
    Select Case TipoTrans2
        Case 0, 1
            Sql = "¿Seguro que desea eliminar la Transferencia?"
        Case 2
            Sql = "¿Seguro que desea eliminar el Pago Domiciliado?"
        Case 3
            Sql = "¿Seguro que desea eliminar el Confirming?"
    End Select
    Sql = Sql & vbCrLf & " Código: " & lw1.SelectedItem.Text
    Sql = Sql & vbCrLf & " Fecha: " & lw1.SelectedItem.SubItems(2)
    Sql = Sql & vbCrLf & " Banco: " & lw1.SelectedItem.SubItems(5)
    Sql = Sql & vbCrLf & " Importe: " & lw1.SelectedItem.SubItems(8)
    
    
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = lw1.SelectedItem.Text
        If TipoTrans2 = 0 Or TipoTrans2 = 4 Then
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
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim FecUltCob As String
Dim Importe As Currency
Dim NumLinea As Integer


    ModificarCobros = False
    
    Conn.BeginTrans
    

    Sql = "select * from cobros where transfer = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    Sql = Sql & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        ' antes lo sumaba de los cobros_realizados
        ' ahora lo dejo todo a nulo
    
    
        FecUltCob = ""
        Importe = 0
    
        Sql = "update cobros set fecultco = " & DBSet(FecUltCob, "F", "S")
        If Importe = 0 Then
            Sql = Sql & " , impcobro = " & ValorNulo
        Else
            Sql = Sql & " , impcobro = " & DBSet(Importe, "N", "S")
        End If
        Sql = Sql & ", tiporem = " & ValorNulo
        Sql = Sql & ", transfer = " & ValorNulo
        Sql = Sql & ", anyorem = " & ValorNulo
        Sql = Sql & ", siturem = " & ValorNulo
        Sql = Sql & ", situacion = 0 , reciboanticipado =0 "
        
        Sql = Sql & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and "
        Sql = Sql & " numfactu = " & DBSet(Rs!numfactu, "N") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and "
        Sql = Sql & " numorden = " & DBSet(Rs!numorden, "N")
                    
        Conn.Execute Sql
    
        Rs.MoveNext
    Wend

    Sql = "delete from transferencias where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    Sql = Sql & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Conn.Execute Sql

    Set Rs = Nothing
    ModificarCobros = True
    Conn.CommitTrans
    Exit Function
    
eModificarCobros:
    Conn.RollbackTrans
    MuestraError Err.Number, "Modificar Cobros", Err.Description
End Function

Private Function ModificarPagos() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim FecUltCob As String
Dim Importe As Currency
Dim NumLinea As Integer


    ModificarPagos = False
    
    Conn.BeginTrans
    

    Sql = "select * from pagos where nrodocum = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    Sql = Sql & " and anyodocum = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        ' antes lo sumaba de los cobros_realizados
        ' ahora lo dejo todo a nulo
    
    
        FecUltCob = ""
        Importe = 0
    
        Sql = "update pagos set fecultpa = " & DBSet(FecUltCob, "F", "S")
        If Importe = 0 Then
            Sql = Sql & " , imppagad = " & ValorNulo
        Else
            Sql = Sql & " , imppagad = " & DBSet(Importe, "N", "S")
        End If
        Sql = Sql & ", nrodocum = " & ValorNulo
        Sql = Sql & ", anyodocum = " & ValorNulo
        Sql = Sql & ", situdocum = " & ValorNulo
        Sql = Sql & ", situacion = 0 "
        Sql = Sql & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and "
        Sql = Sql & " numfactu = " & DBSet(Rs!numfactu, "T") & " and fecfactu = " & DBSet(Rs!FecFactu, "F") & " and "
        Sql = Sql & " numorden = " & DBSet(Rs!numorden, "N") & " AND codmacta = " & DBSet(Rs!codmacta, "T")
                    
        Conn.Execute Sql
    
        Rs.MoveNext
    Wend

    Sql = "delete from transferencias where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
    Sql = Sql & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    
    Conn.Execute Sql

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

    txtFecha(4).Text = Format(Now, "dd/mm/yyyy")

    txtCuentas(2).Text = BancoPropio
    If txtCuentas(2).Text <> "" Then
        txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "bancos", "descripcion", "codmacta", txtCuentas(2), "T")
        If txtNCuentas(2).Text = "" Then txtNCuentas(2).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtCuentas(2).Text, "T")
    End If

    PonleFoco txtFecha(2)
    
    Label2.Caption = ""

    Me.Label3(8).Caption = "Fecha factura"   '"Fecha Pago"
    Label1(1).Caption = "Banco"

    

End Sub

Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    
    Me.lwCobros.ListItems.Clear
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub BotonModificar()
Dim Sql As String
    
 
    If lw1.SelectedItem Is Nothing Then Exit Sub
    If lw1.SelectedItem = 0 Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    ModoInsertar = True
    
    LimpiarCampos
    
    CadenaDesdeOtroForm = ""
    
    Modo = 4
    PonerModo Modo

    txtFecha(5).Text = Format(lw1.SelectedItem.SubItems(2), "dd/mm/yyyy")
    txtCuentas(3).Text = lw1.SelectedItem.SubItems(5)
    txtNCuentas(3).Text = lw1.SelectedItem.SubItems(7)
    Text2.Text = lw1.SelectedItem.SubItems(7)

    Select Case TipoTrans2
        Case 0, 1
            Label3(12).Caption = "Transferencia: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""
        Case 2
            Label3(12).Caption = "Pago Domiciliado: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""
        Case 3
            Label3(12).Caption = "Confirming: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""
        Case 4
            Label3(12).Caption = "Anticipar vencimientos: " & lw1.SelectedItem.Text & "/" & lw1.SelectedItem.SubItems(1) & ""
    End Select
        
        

    Me.Label3(8).Caption = "Fecha factura"   '"Fecha Pago"
    Label1(1).Caption = "Banco"
    
    If TipoTrans2 = 0 Or TipoTrans2 = 4 Then
        Sql = "from cobros, formapago,tipofpago where transfer = " & DBSet(lw1.SelectedItem.Text, "N") & " and anyorem = " & lw1.SelectedItem.SubItems(1)
        Sql = Sql & " and cobros.codforpa = formapago.codforpa AND tipofpago.tipoformapago = formapago.tipforpa  "
    
        PonerVtosTransferencia Sql, True
    Else
        Sql = "from pagos left join cuentas on pagos.codmacta=cuentas.codmacta where nrodocum = " & DBSet(lw1.SelectedItem.Text, "N") & " and anyodocum = " & lw1.SelectedItem.SubItems(1)
    
        PonerVtosTransferenciaPagos Sql, True
    End If
    
    If TipoTrans2 = 0 Or TipoTrans2 = 1 Then
    
        PosicionarCombo Me.cboConcepto2, lw1.SelectedItem.SubItems(11)
        
    Else
    
        chkFecha.Value = lw1.SelectedItem.SubItems(11)
        If TipoTrans2 = 3 Then
            
            Me.chkConfirmingPP2.Value = IIf(lw1.SelectedItem.SubItems(12) = "*", 1, 0)
        End If
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
        lwCobros.Height = 4104
        FrameCreacionRemesa.Refresh
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
Dim OmitirPregunta As Boolean
Dim VtosAgrupados As Integer

    Select Case Boton
        Case 1
            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
        
        
            CadenaDesdeOtroForm = ""
            If Asc(UCase(lw1.SelectedItem.SubItems(9))) > Asc("B") Then
                Select Case TipoTrans2
                    Case 0, 1
                        CadenaDesdeOtroForm = "No se puede modificar una transferencia " & lw1.SelectedItem.SubItems(3)
                    Case 2
                        CadenaDesdeOtroForm = "No se puede modificar un pago domiciliado " & lw1.SelectedItem.SubItems(3)
                    Case 3
                        CadenaDesdeOtroForm = "No se puede modificar un confirming " & lw1.SelectedItem.SubItems(3)
                    Case 4
                        CadenaDesdeOtroForm = "No se puede modificar el anticpo de facturas " & lw1.SelectedItem.SubItems(3)
                End Select
            End If
            
            If CadenaDesdeOtroForm <> "" Then
                MsgBox CadenaDesdeOtroForm, vbExclamation
                Exit Sub
            End If
        
            
            'Ene 2018
            VtosAgrupados = 0
            If TipoTrans2 < 2 Then
                'Transferencias /abonos. Podemos agrupar vtos en fichero
                CadenaDesdeOtroForm = ""
                frmMensajes.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & TipoTrans2 & "|"
                frmMensajes.Opcion = 64
                frmMensajes.Show vbModal
                If CadenaDesdeOtroForm = "" Then Exit Sub
                Fecha = RecuperaValor(CadenaDesdeOtroForm, 1)
                
                If RecuperaValor(CadenaDesdeOtroForm, 2) = "1" Then
                    VtosAgrupados = ComprobacionesAgrupaFicheroTransfer(TipoTrans2 = 0)
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
            OmitirPregunta = False
            
            Sql = "No se debe contabilizar "
            Select Case TipoTrans2
                Case 0, 1
                    Sql = Sql & "una transferencia abierta. Sin llevar al banco."
                Case 2
                    Sql = Sql & "un pago domiciliado abierto. Sin llevar al banco."
                Case 3
                    
                    Sql = Sql & "un confirming abierto. Sin llevar al banco."
                Case 4
                    'BA
                    Sql = Sql & "un anticipo sin llevar al banco."
                    OmitirPregunta = True
            End Select
            
            
            CadenaDesdeOtroForm = ""
            If lw1.SelectedItem.SubItems(9) = "A" Then
                If TipoTrans2 = 3 Then
                    Msg = DevuelveDesdeBD("ctaConfirming", "bancos", "codmacta", lw1.SelectedItem.SubItems(5), "T")
                    If Msg <> "" Then
                        'Va por confirming, con cuenta, dias palazamiento...
                        CadenaDesdeOtroForm = Sql
                    End If
                ElseIf TipoTrans2 = 4 Then
                    Msg = DevuelveDesdeBD("CtaAnticipoRecibos", "bancos", "codmacta", lw1.SelectedItem.SubItems(5), "T")
                    If Msg <> "" Then CadenaDesdeOtroForm = Sql
                Else
                    CadenaDesdeOtroForm = Sql
                End If
            End If
            
            If lw1.SelectedItem.SubItems(9) = "Q" Then
                Sql = "No se puede contabilizar "
                Select Case TipoTrans2
                    Case 0, 1
                        CadenaDesdeOtroForm = Sql & "Transferencia abonada."
                    Case 2
                        CadenaDesdeOtroForm = Sql & "Pago domiciliado abonado."
                    Case 3
                        CadenaDesdeOtroForm = Sql & "Confirming abonado."
                    Case 4
                        CadenaDesdeOtroForm = Sql & "Anticipo facturas YA abonado."
                        OmitirPregunta = False
               End Select
            End If
            
            
            Sql = ""
            If CadenaDesdeOtroForm <> "" Then
                If TipoTrans2 < 2 Or TipoTrans2 = 4 Then
                    If vUsu.Nivel = 0 Then
                        Sql = "S"
                        If lw1.SelectedItem.SubItems(9) <> "Q" Then Sql = "S"
                    Else
                        Sql = "N"
                    End If
                Else
                    Sql = "N"
                End If
            End If
            
            If Sql <> "" Then
                If Sql = "N" Then
                    MsgBox CadenaDesdeOtroForm, vbExclamation
                    CadenaDesdeOtroForm = ""
                    Exit Sub
                Else
                    If Not OmitirPregunta Then
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & vbCrLf & vbCrLf & "¿Seguro que desea contabilizarla?" & vbCrLf
                        CadenaDesdeOtroForm = String(70, "*") & vbCrLf & CadenaDesdeOtroForm & vbCrLf & String(70, "*")
                        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                    End If
                End If
            End If
            
             'Lo que habia
            Sql = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(4) & "|" & lw1.SelectedItem.SubItems(5) & "|" & lw1.SelectedItem.SubItems(8) & "|"
            If TipoTrans2 = 3 Then
                'Confirming
                CadenaDesdeOtroForm = DevuelveDesdeBD("ctaconfirming", "bancos", "codmacta", lw1.SelectedItem.SubItems(5), "T")
                If CadenaDesdeOtroForm = "" Then
               
                Else
                     Sql = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(4) & "|" & CadenaDesdeOtroForm & "|" & lw1.SelectedItem.SubItems(8) & "|"
                End If
            End If
            CadenaDesdeOtroForm = ""
            frmTESTransferenciasCont.Opcion = 8
            frmTESTransferenciasCont.Cobros = (TipoTrans2 = 0 Or TipoTrans2 = 4)
            frmTESTransferenciasCont.NumeroDocumento = CStr(Sql)
            frmTESTransferenciasCont.TipoTrans2 = TipoTrans2
            frmTESTransferenciasCont.Show vbModal
         
            'Hay que poner en el formualrio de arriba valor a cadenadesdeotroform si ha modificado
            If HaHabidoCambios Then CargaList
    Case 3
    
            If TipoTrans2 > 1 Then Exit Sub
            
            Sql = "0|0|"
            If Not lw1.SelectedItem Is Nothing Then
                If lw1.SelectedItem.SubItems(9) = "Q" Then Sql = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|"                       'Numtrans Anotrans
            End If
            frmTESTransferDev.Numtrans = CInt(RecuperaValor(Sql, 1))
            frmTESTransferDev.Anotrans = CInt(RecuperaValor(Sql, 2))
            frmTESTransferDev.Cobros = TipoTrans2 = 0
            frmTESTransferDev.Show vbModal

        
    
    End Select
End Sub


Private Function GeneraNormaBancaria(AgrupaVtos As Integer) As Boolean
Dim B As Boolean
Dim NumF As String
Dim NIF As String
Dim IdFich As String
Dim Tipoconfirming As Integer  '0 Stand    1Caixa    2Santander


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
    
   
    If TipoTrans2 < 2 Then
        If Not comprobarCuentasBancariasPagos(lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(1), TipoTrans2 = 1) Then
            Set miRsAux = Nothing
            Exit Function
        End If
    End If
    
    If Not ComprobarNifDatosProveedor Then Exit Function
        
   
    'Si es Norma34, transferenia, ofertaremos si queremos que el fichero sea
   
    Set miRsAux = Nothing
        'B = GeneraFicheroNorma34(NIF, Adodc1.Recordset!Fecha, Adodc1.Recordset!codmacta, "9", Adodc1.Recordset!Codigo, Adodc1.Recordset!descripcion, TipoDeFrm <> 0)
        
    If TipoTrans2 < 2 Then
            'Estaba como fecha: lw1.SelectedItem.SubItems(2)
        B = GeneraFicheroNorma34(NIF, Fecha, lw1.SelectedItem.SubItems(5), CStr(lw1.SelectedItem.SubItems(11)), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), TipoTrans2 = 1, lw1.SelectedItem.SubItems(1), IdFich, AgrupaVtos > 0)
        
    Else
         If TipoTrans2 = 3 Then ' si es caixa confirmning
            'Van por una "norma" de la caixa. De momento picassent
             Sql = DevuelveDesdeBD("TipoFichConfi", "bancos", "codmacta", lw1.SelectedItem.SubItems(5), "T")
             Tipoconfirming = Val(Sql)
             
             Select Case Tipoconfirming
             Case vbConfirmingCaixa
                B = GeneraFicheroCaixaConfirming(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
            
            Case vbConfirmingSantander
                B = GeneraFicheroGrSantanderConfirming(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
                'B = GeneraFicheroGrSantanderConfirmingC12(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
            Case vbSabadell
                B = GeneraFicheroSabadellConfirming(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
            
            Case vbConfirmingRural
                B = GeneraFicheroCaixaRural(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
                
            Case vbCaixaPopular
                B = GeneraFicheroCaixaPipular(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
                
            Case vbBancaMarch
                B = GeneraFicheroBancaMarch(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
                
            Case Else
                B = GeneraFicheroConfirmingSt(NIF, lw1.SelectedItem.SubItems(2), lw1.SelectedItem.SubItems(5), lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(7), lw1.SelectedItem.SubItems(1))
            End Select
            
        ElseIf TipoTrans2 = 4 Then
            MsgBox "No aceptar intercambio de ficheros con el banco", vbExclamation
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
            CopiarFicheroNormaBancaria CByte(TipoTrans2), cd1.FileName
            
            CadenaDesdeOtroForm = "OK"
                        
            
            If TipoTrans2 < 2 Then
                Set miRsAux = New ADODB.Recordset
                Sql = "Select * from transferencias where codigo = " & lw1.ListItems(lw1.SelectedItem.Index).Text
                Sql = Sql & " and anyo = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then
                    'Si a estaba generado el fichero, guardo en LOG
                    If Asc(DBLet(miRsAux!Situacion, "T")) >= Asc("B") Then
                        'YA se habia generado
                        Sql = "Transfer: " & Format(miRsAux!Codigo, "0000") & "/" & miRsAux!Anyo & "   Fecha:" & miRsAux!Fecha & vbCrLf
                        Sql = Sql & DBLet(miRsAux!Descripcion, "T") & vbCrLf
                        
                        Sql = Sql & "Banco: " & miRsAux!codmacta & " " & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux!codmacta, "T") & vbCrLf
                        Sql = Sql & "Situacion: " & miRsAux!Situacion & " Importe : " & miRsAux!Importe & "€ "
                        If TipoTrans2 = 0 Then
                            RC = "transfer= " & miRsAux!Codigo & " AND anyorem =" & miRsAux!Anyo & " AND 1"
                            RC = DevuelveDesdeBD("count(*)", "cobros", RC, "1")
                        Else
                            RC = "nrodocum= " & miRsAux!Codigo & " AND anyodocum =" & miRsAux!Anyo & " AND 1"
                            RC = DevuelveDesdeBD("count(*)", "pagos", RC, "1")
                        End If
                        Sql = Sql & " Vtos: " & RC
                        If DBLet(miRsAux!LlevaAgrupados, "N") > 0 Then Sql = Sql & " Cli/Prov agr: " & miRsAux!LlevaAgrupados
                        Sql = Sql & vbCrLf
                        RC = DBLet(miRsAux!IdFicheroSEPA, "T")
                        If RC <> "" Then
                            RC = "ID_Fich: " & RC & "       Usuario: " & DBLet(miRsAux!usurioFich, "T")
                            Sql = Sql & RC
                        End If
                        RC = ""
                        vLog.Insertar 31, vUsu, Sql
                    End If
                End If
                miRsAux.Close
                Set miRsAux = Nothing
                
            End If
            
            
            
            
            NIF = "update transferencias set situacion = 'B' "
            If TipoTrans2 < 2 Then
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
                If TipoTrans2 = 0 Then
                    If Not UpdatearCobrosTransferencia Then MsgBox "Error updateando cobros transferencia", vbExclamation
                Else
                    If Not UpdatearPagosTransferencia Then
                        Select Case TipoTrans2
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
    
    Sql = "Select * from cobros WHERE transfer=" & DBSet(lw1.SelectedItem.Text, "N")
    Sql = Sql & " AND anyorem =" & DBSet(lw1.SelectedItem.SubItems(1), "N")
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            While Not miRsAux.EOF
                Sql = "UPDATE cobros SET fecultco = '" & Format(lw1.SelectedItem.SubItems(2), FormatoFecha) & "', impcobro = "
                Im = miRsAux!ImpVenci
                If Not IsNull(miRsAux!Gastos) Then Im = Im + miRsAux!Gastos
                Sql = Sql & TransformaComasPuntos(CStr(Im))
                
                Sql = Sql & " ,siturem = 'B'"
                
                
                'WHERE
                Sql = Sql & " WHERE numserie='" & miRsAux!NUmSerie
                Sql = Sql & "' AND  numfactu =  " & miRsAux!numfactu
                Sql = Sql & "  AND  fecfactu =  '" & Format(miRsAux!FecFactu, FormatoFecha)
                Sql = Sql & "' AND  numorden =  " & miRsAux!numorden
                'Muevo siguiente
                miRsAux.MoveNext
                
                'Ejecuto SQL
                If Not Ejecuta(Sql) Then MsgBox "Error: " & Sql, vbExclamation
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
    
    Sql = "Select * from pagos WHERE nrodocum=" & DBSet(lw1.SelectedItem.Text, "N")
    Sql = Sql & " AND anyodocum =" & DBSet(lw1.SelectedItem.SubItems(1), "N")
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            While Not miRsAux.EOF
                Sql = "UPDATE pagos SET situdocum = 'B'"
                If TipoTrans2 <> 2 Then
                    Sql = Sql & ", fecultpa = '" & Format(lw1.SelectedItem.SubItems(2), FormatoFecha) & "', imppagad = "
                    Im = miRsAux!ImpEfect
                    Sql = Sql & TransformaComasPuntos(CStr(Im))
                
                
                End If
                
                'WHERE
                Sql = Sql & " WHERE numserie='" & miRsAux!NUmSerie
                Sql = Sql & "' AND  numfactu =  " & DBSet(miRsAux!numfactu, "T")
                Sql = Sql & "  AND  fecfactu =  '" & Format(miRsAux!FecFactu, FormatoFecha)
                Sql = Sql & "' AND  numorden =  " & miRsAux!numorden
                Sql = Sql & " and codmacta = " & DBSet(miRsAux!codmacta, "T")
                'Muevo siguiente
                miRsAux.MoveNext
                
                'Ejecuto SQL
                If Not Ejecuta(Sql) Then MsgBox "Error: " & Sql, vbExclamation
            Wend
    End If
    miRsAux.Close
                    
                    
                    
    UpdatearPagosTransferencia = True
    Exit Function
EUpdatearCobrosTransferencia:
    
End Function





Private Function ComprobarNifDatosProveedor() As Boolean
Dim Sql As String

    ComprobarNifDatosProveedor = False
    
    If TipoTrans2 = 0 Then
    
        Sql = "select nifclien, codmacta, nomclien from cobros where transfer = " & lw1.SelectedItem.Text & " and anyorem = " & DBSet(lw1.SelectedItem.SubItems(1), "N")
        Sql = Sql & " GROUP BY 1"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        While Not miRsAux.EOF
            
            If Trim(DBLet(miRsAux!nifclien, "T")) = "" Then Sql = Sql & "- " & miRsAux!codmacta & " " & miRsAux!nomclien & vbCrLf
            miRsAux.MoveNext
        
        Wend
        
    Else
        
        Sql = "select nifprove, codmacta, nomprove from pagos where nrodocum = " & lw1.SelectedItem.Text & " and anyodocum = " & DBSet(lw1.SelectedItem.SubItems(1), "N")
        Sql = Sql & " GROUP BY 1"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        While Not miRsAux.EOF
            
            If Trim(DBLet(miRsAux!NifProve, "T")) = "" Then Sql = Sql & "- " & miRsAux!codmacta & " " & miRsAux!nomprove & vbCrLf
            miRsAux.MoveNext
        
        Wend
    
    End If
        
    miRsAux.Close
    If Sql <> "" Then
        MsgBox "Error en NIFs: " & vbCrLf & Sql, vbExclamation
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
                If CuentaCorrectaUltimoNivel(Cta, Sql) Then
                    Text1(Index).Text = Cta
                    Text1(3).Text = Sql
                Else
                    MsgBox Sql, vbExclamation
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
Dim I2 As Currency


    lwCobros.ListItems.Clear
    If Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
    
'    Set Me.lwCobros.SmallIcons = frmPpal.ImgListviews
    Set lwCobros.SmallIcons = frmppal.imgListComun16
    
    
    Set miRsAux = New ADODB.Recordset
    
    cad = "Select cobros.*,nomforpa,siglas " & vSql
    cad = cad & " ORDER BY fecvenci"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!numfactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.ListSubItems(3).Tag = ""
        
        
        If Not ComprobarIBANCuentaBancaria2(DBLet(miRsAux!IBAN, "T"), cad) Then
            IT.ListSubItems(3).Tag = "NO"
            IT.ListSubItems(3).Bold = True
            IT.ListSubItems(3).ForeColor = vbRed
            IT.ListSubItems(3).ToolTipText = cad
        End If
        
        IT.SubItems(4) = miRsAux!FecVenci
        If IsNull(miRsAux!nomclien) Then
            IT.SubItems(5) = miRsAux!codmacta
        Else
            IT.SubItems(5) = miRsAux!nomclien
        End If
    
        IT.SubItems(6) = miRsAux!siglas
    
    
        If Modificar Then IT.Checked = True
    
    
            
    
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not Modificar Then
            If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        End If
        
        IT.SubItems(7) = Format(Importe, FormatoImporte)
        
        If Modificar Then ImporteTot = ImporteTot + Importe

        IT.Tag = Abs(Importe)  'siempre valor absoluto
            
        If DBLet(miRsAux!Devuelto, "N") = 1 Then
            IT.SmallIcon = 42
        End If
            
        If Modificar Then
            IT.SubItems(8) = txtCuentas(3).Text
        Else
            IT.SubItems(8) = txtCuentas(2).Text
        End If
        
        'Para el orden
        cad = Mid(miRsAux!NUmSerie & "  ", 1, 3) & Format(miRsAux!numfactu, "0000000")
        IT.SubItems(10) = cad
        IT.SubItems(11) = Format(miRsAux!FecFactu, "yyyymmdd") & cad
        IT.SubItems(12) = Format(miRsAux!FecVenci, "yyyymmdd") & cad
        I2 = (Importe * 100)
        I2 = 1000000 + I2
        IT.SubItems(13) = Format(I2, "0000000000")
        'Debug.Print IT.SubItems(12)
    
        IT.SubItems(14) = miRsAux!codmacta
        
        If DBLet(miRsAux!observa, "T") <> "" Then
          IT.SubItems(15) = "*"
          cad = miRsAux!observa
          If Len(cad) > 60 Then cad = Mid(cad, 1, 57) & "..."
          IT.ListSubItems(15).ToolTipText = cad
        Else
            IT.SubItems(15) = " "
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
    BloqueaTXT Text1(4), True

End Sub


Private Sub PonerVtosTransferenciaPagos(vSql As String, Modificar As Boolean)
Dim IT
Dim ImporteTot As Currency
Dim I2 As Currency
Dim Icono As Integer

    lwCobros.ListItems.Clear
    If Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
    
'    Set Me.lwCobros.SmallIcons = frmPpal.ImgListviews
    Set lwCobros.SmallIcons = frmppal.imgListComun16
    
    
    Set miRsAux = New ADODB.Recordset
    
    cad = ""
    If InStr(1, LCase(vSql), "tipofpago") = 0 Then cad = " '' "
    
    
    cad = "Select pagos.*," & cad & " siglas " & ",  embargo " & vSql
    cad = cad & " ORDER BY fecefect"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = miRsAux!numfactu
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        
         IT.ListSubItems(3).Tag = ""
        
        Icono = 0
        If Not ComprobarIBANCuentaBancaria2(DBLet(miRsAux!IBAN, "T"), cad) Then
            IT.ListSubItems(3).Tag = "NO"
            IT.ListSubItems(3).Bold = True
            IT.ListSubItems(3).ForeColor = vbRed
            IT.ListSubItems(3).ToolTipText = cad
            Icono = 42
        End If
       
        
        
        IT.SubItems(4) = miRsAux!fecefect
        IT.SubItems(5) = IIf(IsNull(miRsAux!nomprove), "N/D", miRsAux!nomprove)
    
    
    
        If DBLet(miRsAux!embargo, "N") = 1 Then
            IT.ListSubItems(5).ForeColor = vbRed
            IT.ListSubItems(5).ToolTipText = "** EMBARGADO **"
            IT.ListSubItems(3).Tag = "EMB"
            Icono = 34
                        
        End If
        
    
    
    
        IT.SubItems(6) = miRsAux!siglas
        If Modificar Then IT.Checked = True
    
        Importe = 0
        Importe = Importe + miRsAux!ImpEfect
        
        'Si ya he cobrado algo
        If Not Modificar Then
            If Not IsNull(miRsAux!imppagad) Then Importe = Importe - miRsAux!imppagad
        End If
        
        IT.SubItems(7) = Format(Importe, FormatoImporte)
        
        ImporteTot = ImporteTot + Importe

        IT.Tag = Abs(Importe)  'siempre valor absoluto
            
'        If DBLet(miRsAux!Devuelto, "N") = 1 Then
'            IT.SmallIcon = 42
'        End If
            
        If Modificar Then
            IT.SubItems(8) = txtCuentas(3).Text
        Else
            IT.SubItems(8) = txtCuentas(2).Text
        End If
        
        IT.SubItems(9) = miRsAux!codmacta
        
        
        'Para el orden
        cad = Mid(miRsAux!NUmSerie & "  ", 1, 3) & Format(miRsAux!numfactu, "0000000")
        IT.SubItems(10) = cad
        IT.SubItems(11) = Format(miRsAux!FecFactu, "yyyymmdd") & cad
        IT.SubItems(12) = Format(miRsAux!fecefect, "yyyymmdd") & cad
        I2 = (Importe * 100)
        I2 = 1000000 + I2
        IT.SubItems(13) = Format(I2, "0000000000")
        
        IT.SubItems(14) = miRsAux!codmacta
        
        
        'Si tiene observaciones
        If DBLet(miRsAux!observa, "T") <> "" Then
          IT.SubItems(15) = "*"
          cad = miRsAux!observa
          If Len(cad) > 60 Then cad = Mid(cad, 1, 57) & "..."
          IT.ListSubItems(15).ToolTipText = cad
        Else
            IT.SubItems(15) = " "
        End If
        
        
        
        
        If Icono > 0 Then IT.SmallIcon = Icono
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    If Not Modificar Then
        Text1(4).Text = "" 'Format(ImporteTot, "###,###,##0.00")
    Else
        '###revisar
        Text1(4).Text = Format(Me.lw1.SelectedItem.SubItems(8), "###,###,##0.00")
    End If
    

End Sub


'''''Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
'''''Dim Insertar As Boolean
'''''    SQL = ""
'''''    For i = 1 To Me.lwCobros.ListItems.Count
'''''        If Me.lwCobros.ListItems(i).Checked Then
'''''
'''''            Insertar = True
'''''            If Me.lwCobros.ListItems(i).Bold Then
'''''                RegistroDestino = i
'''''                If SinDestino Then Insertar = False
'''''            End If
'''''            If Insertar Then
'''''                SQL = SQL & ", ('" & lwCobros.ListItems(i).Text & "'," & lwCobros.ListItems(i).SubItems(1)
'''''                SQL = SQL & ",'" & Format(lwCobros.ListItems(i).SubItems(2), FormatoFecha) & "'," & lwCobros.ListItems(i).SubItems(3) & ")"
'''''            End If
'''''
'''''        End If
'''''    Next
'''''    SQL = Mid(SQL, 2)
'''''
'''''End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    Dim B As Boolean
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = False 'DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N")
    
        B = DBLet(Rs!Especial, "N") And Not (lw1.SelectedItem Is Nothing)
        If B Then B = TipoTrans2 <> 4
        Toolbar2.Buttons(1).Enabled = B
        
        If TipoTrans2 <> 2 Then Toolbar2.Buttons(2).Enabled = DBLet(Rs!Especial, "N") And Not (lw1.SelectedItem Is Nothing)
        Toolbar2.Buttons(3).Enabled = DBLet(Rs!Especial, "N") 'And Not (lw1.SelectedItem Is Nothing)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub



Private Sub CargaList()
Dim IT
Dim PriVezCol As Boolean

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    cad = "Select transferencias.codigo,transferencias.anyo, transferencias.fecha, "
    'CASE situacion WHEN 0 THEN 'ABIERTA' WHEN 1 THEN 'GENERADO FICHERO' WHEN 2 THEN 'CONTABILIZADA' END as descsituacion,"
    cad = cad & " wtiposituacionrem.descsituacion, "
    If TipoTrans2 < 2 Then
        cad = cad & " CASE concepto WHEN 0 THEN 'PENSION' WHEN 1 THEN 'NOMINA' WHEN 9 THEN 'ORDINARIA' END as desconcepto, "
    Else
        ' pagos domiciliados como caixa confirming
        cad = cad & " if (concepto = 0,'Vencimiento','Fecha intro.') as desconcepto, "
    End If
    cad = cad & " transferencias.codmacta,cuentas.nommacta,"
    cad = cad & " transferencias.descripcion, Importe , transferencias.tipotrans, situacion, transferencias.concepto,solopago "
    cad = cad & " from cuentas,transferencias,usuarios.wtiposituacionrem where transferencias.codmacta=cuentas.codmacta"
    cad = cad & " and wtiposituacionrem.situacio = transferencias.situacion "
    
    cad = cad & PonerFiltro
    
    If CampoOrden = "" Then
        CampoOrden = "transferencias.anyo, transferencias.codigo "
    ElseIf CampoOrden = "transferencias.fecha" Then
        CampoOrden = CampoOrden & " DESC , transferencias.codigo "
    End If
    cad = cad & " ORDER BY " & CampoOrden ' transferencias.anyo desc,
    If Orden Then cad = cad & " DESC"
    
    PriVezCol = True
    If lw1.ColumnHeaders.Count > 7 Then PriVezCol = False
    
    lw1.ColumnHeaders.Clear
    
    lw1.ColumnHeaders.Add , , "Código", 1050
    lw1.ColumnHeaders.Add , , "Año", 700
    lw1.ColumnHeaders.Add , , "Fecha", 1350
    lw1.ColumnHeaders.Add , , "Situación", 1540
    lw1.ColumnHeaders.Add , , "Concepto", 1500
    lw1.ColumnHeaders.Add , , "Cuenta", 1440
    
    lw1.ColumnHeaders.Add , , "Nombre", 2940
    
    lw1.ColumnHeaders.Add , , "Descripción", 2840
    lw1.ColumnHeaders.Add , , "Importe", IIf(TipoTrans2 = 3, 1200, 1840), 1
    lw1.ColumnHeaders.Add , , "S", 0, 1
    lw1.ColumnHeaders.Add , , "T", 0, 1
    lw1.ColumnHeaders.Add , , "C", 0, 1
    lw1.ColumnHeaders.Add , , "PP", IIf(TipoTrans2 = 3, 600, 1), 1  'Solo pago
    
    If PriVezCol Then
        Me.Refresh
        DoEvent2
        Screen.MousePointer = vbHourglass
    End If
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = DBLet(miRsAux!Codigo, "N")
        IT.SubItems(1) = DBLet(miRsAux!Anyo, "N")
        IT.SubItems(2) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        IT.SubItems(3) = DBLet(miRsAux!descsituacion, "T")
        IT.ListSubItems(3).ToolTipText = DBLet(miRsAux!descsituacion, "T")
        
        If TipoTrans2 = 4 Then
            IT.SubItems(4) = " "
        Else
            IT.SubItems(4) = DBLet(miRsAux!desconcepto, "T")
            IT.ListSubItems(4).ToolTipText = DBLet(miRsAux!desconcepto, "T")
        End If
        
        IT.SubItems(5) = miRsAux!codmacta
        IT.SubItems(6) = DBLet(miRsAux!Nommacta, "T")
        IT.ListSubItems(6).ToolTipText = DBLet(miRsAux!Nommacta, "T")
        IT.SubItems(7) = DBLet(miRsAux!Descripcion, "T")
        IT.ListSubItems(7).ToolTipText = DBLet(miRsAux!Descripcion, "T")
        IT.SubItems(8) = Format(miRsAux!Importe, "###,###,##0.00")
        IT.SubItems(9) = miRsAux!Situacion
        IT.SubItems(10) = miRsAux!tipotrans
        IT.SubItems(11) = miRsAux!Concepto
        If TipoTrans2 = 3 Then
            IT.SubItems(12) = IIf(DBLet(miRsAux!solopago, "N") = 1, "*", " ") 'para los confirming es
        Else
            IT.SubItems(12) = " "
        End If
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
Dim C2 As String
    'Filtro
    Select Case TipoTrans2
        Case 0 ' abonos
            C = " and transferencias.tipotrans = 1"
        Case 1 ' pagos
            C = " and transferencias.tipotrans = 0 and transferencias.subtipo = 0"
        Case 2 ' pagos domiciliados
            C = " and transferencias.tipotrans = 0 and transferencias.subtipo = 1"
        Case 3 ' confirming
            C = " and transferencias.tipotrans = 0 and transferencias.subtipo = 2"
        Case 4
            'Anticipos de credito
            C = " and transferencias.tipotrans = 2"
    End Select
    
    If Me.cboFiltro.ListIndex > 0 Then
        If cboFiltro.ListIndex = 3 Then
            'Ejercciios siguiente
            C2 = DateAdd("d", 1, vParam.fechafin)
        Else
            'Desde fecha incio
            C2 = vParam.fechaini
        End If
        C2 = " transferencias.fecha >= " & DBSet(C2, "F")
        
        If cboFiltro.ListIndex = 2 Then C2 = C2 & " AND transferencias.fecha <= " & DBSet(vParam.fechafin, "F")
        
        
        If C <> "" Then C = C & " AND "
        C = C & C2
    End If

    
    
    PonerFiltro = C
End Function


Private Sub NuevaTransf()

Dim forpa As String
Dim cad As String
Dim Impor As Currency
Dim Sql2 As String

    Sql = " (1=1) "
    
    
    
    'Del vto
    If txtFecha(2).Text <> "" Then Sql = Sql & " AND cobros.fecvenci >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
    If txtFecha(3).Text <> "" Then Sql = Sql & " AND cobros.fecvenci <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"
   
    
    'Si ha puesto importe desde Hasta
    If txtImporte(0).Text <> "" Then Sql = Sql & " AND impvenci >= " & TransformaComasPuntos(ImporteFormateado(txtImporte(0).Text))
    If txtImporte(1).Text <> "" Then Sql = Sql & " AND impvenci <= " & TransformaComasPuntos(ImporteFormateado(txtImporte(1).Text))
    
    
    'Desde hasta cuenta
    If Me.txtCuentas(0).Text <> "" Then Sql = Sql & " AND cobros.codmacta >= '" & txtCuentas(0).Text & "'"
    If Me.txtCuentas(1).Text <> "" Then Sql = Sql & " AND cobros.codmacta <= '" & txtCuentas(1).Text & "'"
    
    'El importe
    If Me.TipoTrans2 = 0 Then
        'abonos
        Sql = Sql & " AND (impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) < 0"
    Else
        'recibo anticipado
        Sql = Sql & " AND (impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) > 0"
    End If

    'serie
    If txtSerie(0).Text <> "" Then _
        Sql = Sql & " AND cobros.numserie >= '" & txtSerie(0).Text & "'"
    If txtSerie(1).Text <> "" Then _
        Sql = Sql & " AND cobros.numserie <= '" & txtSerie(1).Text & "'"
    
    'Fecha factura
    If txtFecha(0).Text <> "" Then _
        Sql = Sql & " AND cobros.fecfactu >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    If txtFecha(1).Text <> "" Then _
        Sql = Sql & " AND cobros.fecfactu <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    
    'Codigo factura
    If txtNumFac(0).Text <> "" Then _
        Sql = Sql & " AND cobros.numfactu >= '" & txtNumFac(0).Text & "'"
    If txtNumFac(1).Text <> "" Then _
        Sql = Sql & " AND cobros.numfactu <= '" & txtNumFac(1).Text & "'"
    
    
    Sql = Sql & " and situacion = 0 "
     
    If Me.TipoTrans2 = 0 Then
        If Me.chkAbonosSoloTransferencia.Value = 1 Then Sql = Sql & " AND formapago.tipforpa=1"
        
    Else
        'Recibos anticipados
        Sql = Sql & " AND formapago.tipforpa IN (1,2,3,5,7)"   'trans .tal, pag confir y domic
        Sql = Sql & " AND recedocu=0"
    End If
    
    
    
    CadenaDesdeOtroForm = ""

    Sql2 = Sql & " and not cobros.impcobro is null and cobros.impcobro <> 0 and cobros.codmacta=cuentas.codmacta AND (transfer is null) "
    
    Sql2 = "select cobros.* FROM cobros,cuentas,formapago  WHERE cobros.codforpa=formapago.codforpa AND " & Sql2
    
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
    
    Sql = Sql & " and (cobros.impcobro is null or cobros.impcobro = 0)"
        
     
    Screen.MousePointer = vbHourglass
    Set Rs = New ADODB.Recordset
    
        
    
    'Que la cuenta NO este bloqueada
    I = 0
    
    cad = " FROM cobros,formapago,cuentas WHERE cobros.codforpa = formapago.codforpa AND transfer is null AND situacion = 0 and "
    cad = cad & " cobros.codmacta=cuentas.codmacta AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtFecha(4).Text), FormatoFecha) & "') AND "
    cad = "Select cobros.codmacta,nommacta,fecbloq" & cad & Sql & " GROUP BY 1 ORDER BY 1"
        
    
    
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        cad = ""
        I = 1
        While Not Rs.EOF
            cad = cad & Rs!codmacta & " - " & Rs!Nommacta & " : " & Rs!FecBloq & vbCrLf
            Rs.MoveNext
        Wend
    End If

    Rs.Close
    
    If I > 0 Then
        cad = "Las siguientes cuentas estan bloqueadas." & vbCrLf & String(60, "-") & vbCrLf & cad
        MsgBox cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
        Exit Sub
    End If
    
    
    cad = " FROM cobros,formapago,cuentas,tipofpago WHERE cobros.codforpa = formapago.codforpa and tipofpago.tipoformapago=formapago.tipforpa AND transfer is null AND "
    cad = cad & " cobros.codmacta=cuentas.codmacta AND situacion = 0 and "
    
    'Hacemos un conteo
    Rs.Open "SELECT Count(*) " & cad & Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        I = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    cad = cad & Sql
    
    
    
    If I > 0 Then
        I = 1  'Para que siga por abajo
    End If
    
    

    'La suma
    If I > 0 Then
        Sql = "select sum(impvenci),sum(impcobro),sum(gastos) " & cad
        Impor = 0
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N") + DBLet(Rs.Fields(2), "N")
        Rs.Close
        If Impor = 0 Then I = 0
    End If
        

    Set Rs = Nothing
    
    If I = 0 Then
        MsgBox "Ningun dato para generera con esos valores", vbExclamation
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
    Else
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        Sql = "Delete from tmpcierre1 where codusu =" & vUsu.Codigo
        Conn.Execute Sql
        
        CadenaDesdeOtroForm = ""
        
        Sql = cad 'Le paso el SELECT
        
        PonerVtosTransferencia Sql, False
        
        Dim CadAux As String
        
        CadAux = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES (" & vUsu.Codigo
        CadAux = CadAux & ",'" & txtCuentas(2).Text & "','" & txtNCuentas(2).Text & "'," & DBSet(Text1(4).Text, "N") & ")"
        If Not Ejecuta(CadAux) Then Exit Sub
        
        CadenaDesdeOtroForm = "'" & Trim(txtCuentas(2).Text) & "'"
                
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub NuevaTransfPagos()

Dim forpa As String
Dim cad As String
Dim Impor As Currency
Dim Sql2 As String
Dim colCtas As Collection


    If TipoTrans2 = 1 Then
        ' transferencias
        Sql = "formapago.tipforpa = " & vbTransferencia
    
    
   
        
    Else
        ' pagos domiciliados o confirming
        If TipoTrans2 = 2 Then
            Sql = "formapago.tipforpa = " & vbPagoDomiciliado
        Else
            Sql = "formapago.tipforpa = " & vbConfirming
        End If
    End If
    
    'Del vto
    If txtFecha(2).Text <> "" Then Sql = Sql & " AND pagos.fecefect >= '" & Format(txtFecha(2).Text, FormatoFecha) & "'"
    If txtFecha(3).Text <> "" Then Sql = Sql & " AND pagos.fecefect <= '" & Format(txtFecha(3).Text, FormatoFecha) & "'"
   
    
    'Si ha puesto importe desde Hasta
    If txtImporte(0).Text <> "" Then Sql = Sql & " AND impefect >= " & TransformaComasPuntos(ImporteFormateado(txtImporte(0).Text))
    If txtImporte(1).Text <> "" Then Sql = Sql & " AND impefect <= " & TransformaComasPuntos(ImporteFormateado(txtImporte(1).Text))
    
    
    'Desde hasta cuenta
    If Me.txtCuentas(0).Text <> "" Then Sql = Sql & " AND pagos.codmacta >= '" & txtCuentas(0).Text & "'"
    If Me.txtCuentas(1).Text <> "" Then Sql = Sql & " AND pagos.codmacta <= '" & txtCuentas(1).Text & "'"
    
    'El importe POSITIVO
    Sql2 = "SI"
    If TipoTrans2 = 3 Then
        If Me.chkIncluirAbonos.Value = 1 Then Sql2 = ""
    End If
    
    If Sql2 = "SI" Then
        Sql = Sql & " AND (impefect - coalesce(imppagad,0)) > 0"    'Solo positivos
    Else
        Sql = Sql & " AND (impefect - coalesce(imppagad,0)) <> 0"   'PUEDEN entrar negativos
    End If
    Sql2 = ""
    

    'serie
    If txtSerie(0).Text <> "" Then _
        Sql = Sql & " AND pagos.numserie >= '" & txtSerie(0).Text & "'"
    If txtSerie(1).Text <> "" Then _
        Sql = Sql & " AND pagos.numserie <= '" & txtSerie(1).Text & "'"
    
    'Fecha factura
    If txtFecha(0).Text <> "" Then _
        Sql = Sql & " AND pagos.fecfactu >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    If txtFecha(1).Text <> "" Then _
        Sql = Sql & " AND pagos.fecfactu <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    
    'Codigo factura
    If txtNumFac(0).Text <> "" Then _
        Sql = Sql & " AND pagos.numfactu >= " & DBSet(txtNumFac(0).Text, "T")
    If txtNumFac(1).Text <> "" Then _
        Sql = Sql & " AND pagos.numfactu <= " & DBSet(txtNumFac(1).Text, "T")
    
    
    Sql = Sql & " and situacion = 0 "
     
    
    'JUNIO18
    'Confirming.
    'Se genera "la transferencia", el fichero..
    'CUando se conmtabiliza, la deuda pasa a ser del banco, con lo cual codmacta se updatea a banco.ctaconfiming
    ' y en el pago, en la columna(oculta) pagos.ctaconfirm ponemos el proveedor
    ' Con lo cual, para las nuevas remesas, no pondremos aquellos pagos que ya hayamos hecho confirmings
    If TipoTrans2 = 3 Then Sql = Sql & " and pagos.ctaconfirm is null "
    
    
    
    
    CadenaDesdeOtroForm = ""

    Sql2 = Sql & " and not pagos.imppagad is null and pagos.imppagad <> 0 and pagos.codmacta=cuentas.codmacta AND (nrodocum is null) AND pagos.codforpa = formapago.codforpa AND tipofpago.tipoformapago = formapago.tipforpa  "
   
    
    Sql2 = "select pagos.*,siglas FROM pagos,cuentas,formapago,tipofpago  WHERE " & Sql2
    
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
    
    Sql = Sql & " and (pagos.imppagad is null or pagos.imppagad = 0)"
        
     
    Screen.MousePointer = vbHourglass
    Set Rs = New ADODB.Recordset

    
        
    
    'Que la cuenta NO este bloqueada
    I = 0
    
    cad = " FROM pagos,formapago,cuentas,tipofpago WHERE pagos.codforpa = formapago.codforpa  AND tipofpago.tipoformapago = formapago.tipforpa  AND nrodocum is null AND situacion = 0 and "
    cad = cad & " pagos.codmacta=cuentas.codmacta AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(txtFecha(4).Text), FormatoFecha) & "') AND "
    cad = "Select pagos.codmacta,nommacta,fecbloq" & cad & Sql & " GROUP BY 1 ORDER BY 1"
        
    
    
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        cad = ""
        I = 1
        While Not Rs.EOF
            cad = cad & Rs!codmacta & " - " & Rs!Nommacta & " : " & Rs!FecBloq & vbCrLf
            Rs.MoveNext
        Wend
    End If

    Rs.Close
    
    If I > 0 Then
        cad = "Las siguientes cuentas estan bloqueadas." & vbCrLf & String(60, "-") & vbCrLf & cad
        MsgBox cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
        Exit Sub
    End If
    
        
    'Julio 2018. Para lanzar abonos

    'Vemos las cuentas que vamos a girar . Sacaremos codmacta
   
    
    cad = " FROM pagos,formapago,cuentas WHERE pagos.codforpa = formapago.codforpa AND nrodocum is null AND situacion = 0 and "
    cad = cad & " pagos.codmacta=cuentas.codmacta AND  nrodocum is null AND situacion = 0 and " & Sql
    
    cad = "Select distinct pagos.codmacta " & cad
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set colCtas = New Collection
    While Not Rs.EOF
        colCtas.Add CStr(Rs!codmacta)
        Rs.MoveNext
    Wend
    Rs.Close
    

    'Ahora veremos los negativos, de las cuentas que vamos a girar
    'Sol el select de los negativos , sin numserie ni na de na
    cad = "(impefect  - coalesce(imppagad,0)) < 0"
    cad = "  (nrodocum is null)  AND pagos.situacion = 0 AND " & cad
    cad = "Select codmacta,nomprove as nommacta,numserie,numfactu,impefect FROM pagos WHERE " & cad
    'Los confirming permitimos NEGATIVOS, si esta el check
    'Con lo cual , para que no de el aviso y/o lance las comepensciones, pongo false
    If TipoTrans2 = 3 Then
        If Me.chkIncluirAbonos.Value = 1 Then cad = cad & " AND FALSE"
    End If
    
    
    If colCtas.Count > 0 Then
        cad = cad & " AND pagos.codmacta IN ("
        For I = 1 To colCtas.Count
            If I > 1 Then cad = cad & ","
            cad = cad & "'" & colCtas.Item(I) & "'"
        Next
        cad = cad & ") ORDER BY codmacta,numfactu"
    
   
        Set colCtas = Nothing
        Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cad = ""
        I = 0
        Set colCtas = New Collection
        While Not Rs.EOF
            If I < 8 Then
                cad = cad & vbCrLf & Rs!codmacta & " " & Rs!Nommacta & "  " & Rs!NUmSerie & Format(Rs!numfactu, "000000") & "   -> " & Format(Rs!ImpEfect, FormatoImporte)
            End If
            I = I + 1
            colCtas.Add CStr(Rs!codmacta)
            Rs.MoveNext
        Wend
        Rs.Close
        
        If cad <> "" Then
            If chkCompensarAbonos.Value = 0 Then
            
                If I >= 9 Then cad = cad & vbCrLf & "....  y " & I & " vencimientos más"
                cad = "Proveedores con abonos. " & vbCrLf & cad & vbCrLf & " ¿Continuar?"
                If MsgBoxA(cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                    Set Rs = Nothing
                    Set colCtas = Nothing
                    cmdAceptar(0).Caption = "&Aceptar"
                    ModoInsertar = False
                    Exit Sub
                End If
                        
            Else
                '-------------------------------------------------------------------------
                CadenaDesdeOtroForm = ""
                For I = 1 To colCtas.Count
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "'" & colCtas.Item(I) & "',"
                Next
                frmTESCompensaAboPro.Show vbModal
                
                
                CadenaDesdeOtroForm = ""
                
                'Actualice BD
                Screen.MousePointer = vbHourglass
                espera 1
                Screen.MousePointer = vbHourglass
                Conn.Execute "commit"
                espera 1
                
            End If
        End If 'cad
    End If
    Set colCtas = Nothing

    
    
    
    
    

    
    
    
    
    
    
    
    
    
        
        
        
        
        
        
        
        
        
    'Para contar si ha o no vtos
    
    cad = " FROM pagos,formapago,cuentas,tipofpago WHERE pagos.codforpa = formapago.codforpa  AND tipofpago.tipoformapago = formapago.tipforpa   AND nrodocum is null AND "
    cad = cad & " pagos.codmacta=cuentas.codmacta AND situacion = 0 and "
    
    'Hacemos un conteo
    Rs.Open "SELECT Count(*) " & cad & Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        I = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    cad = cad & Sql
    
    
    
    If I > 0 Then
        I = 1  'Para que siga por abajo
    End If
    
    

    'La suma
    If I > 0 Then
        Sql = "select sum(impefect),sum(imppagad) " & cad
        Impor = 0
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N")
        Rs.Close
        If Impor = 0 Then I = 0
    End If
        

    Set Rs = Nothing
    
    If I = 0 Then
        MsgBox "Ningun dato  con esos valores", vbExclamation
        
        ModoInsertar = False
        cmdAceptar(0).Caption = "&Aceptar"
        
    Else
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        Sql = "Delete from tmpcierre1 where codusu =" & vUsu.Codigo
        Conn.Execute Sql
        
        CadenaDesdeOtroForm = ""
        
        Sql = cad 'Le paso el SELECT
        
        PonerVtosTransferenciaPagos Sql, False
        
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
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
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
            
            B = CuentaCorrectaUltimoNivel(Cta, Sql)
            
            
            If B And (Index = 2 Or Index = 3) Then
                'Bancos
                Cta = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Cta, "T")
                If Cta = "" Then
                   Sql = "No es una cuenta en bancos: " & txtCuentas(Index).Text & "  - " & Sql
                   B = False
                End If
            End If
            
            If Not B Then
                MsgBox Sql, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
                PonleFoco txtCuentas(Index)
            Else
            
            
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = Sql
                If B = 1 Then
                    txtNCuentas(Index).Tag = ""
                Else
                    txtNCuentas(Index).Tag = Sql
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
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtFecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtFecha(Index)
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
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
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
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    If TipoTrans2 > 0 Then Exit Sub 'Solo es numer i
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
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
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
Dim Sql As String

    VencimientosPorEntidadBancaria = ""

    Sql = " and length(cobros.iban) <> 0 and mid(cobros.iban,5,4) = (select mid(iban,5,4) from bancos where codmacta = " & DBSet(txtCuentas(2).Text, "T") & ")"
    
    VencimientosPorEntidadBancaria = Sql

End Function


Private Function GenerarTransferencia(Opcion As Integer) As Boolean
Dim C As String
Dim NumeroRemesa As Long
Dim Rs As ADODB.Recordset
Dim J As Integer
Dim I As Integer
Dim ImporteQueda As Currency

    On Error GoTo eGenerarTransferencia
    
    GenerarTransferencia = False
    
    
    'EMBARGO
    cad = ""
    For I = 1 To Me.lwCobros.ListItems.Count
        If lwCobros.ListItems(I).Checked Then
            If InStr(1, cad, lwCobros.ListItems(I).SubItems(14)) = 0 Then
                cad = cad & ", " & DBSet(lwCobros.ListItems(I).SubItems(14), "T")
            End If
        End If
    Next I

    If cad <> "" Then
        'No deberia ser ""
        cad = Mid(cad, 2)
        Set miRsAux = New ADODB.Recordset
        cad = "Select codmacta,nommacta from cuentas where embargo=1 AND codmacta in (" & cad & ")"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        While Not miRsAux.EOF
            cad = cad & miRsAux!codmacta & " " & miRsAux!Nommacta & vbCrLf
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        If cad <> "" Then
            MsgBox "Cuentas en situacion de embargo:" & vbCrLf & cad, vbExclamation
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
    
    I = 0
    For J = 1 To lwCobros.ListItems.Count
        If lwCobros.ListItems(J).Checked Then
            I = J
            Exit For
        End If
    Next J
    If I = 0 Then
        If TipoTrans2 = 0 Or TipoTrans2 = 4 Then
            MsgBox "No se ha seleccionado cobros. Revise.", vbExclamation
        Else
            MsgBox "No se ha seleccionado pagos. Revise.", vbExclamation
        End If
        If Opcion = 0 Then BloqueoManual False, "Transferencias", ""
        Exit Function
    End If
    
    
    'A partir de la fecha generemos leemos k remesa corresponde
    If Opcion = 0 Then
        Sql = "select max(codigo) from transferencias where anyo=" & Year(CDate(txtFecha(4).Text))
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        NumeroRemesa = 0
        If Not miRsAux.EOF Then
            NumeroRemesa = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        
        
        NumeroRemesa = NumeroRemesa + 1
    Else
        NumeroRemesa = lw1.SelectedItem.Text
        txtFecha(4).Text = lw1.SelectedItem.SubItems(2)
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    Conn.BeginTrans
    
    
    Set Rs = New ADODB.Recordset
    cad = "Select * from tmpcierre1 where codusu =" & vUsu.Codigo
    If CadenaDesdeOtroForm <> "" Then cad = cad & " and cta in (" & CadenaDesdeOtroForm & ")"
    
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    
        Select Case TipoTrans2
            Case 0, 1
                Label2.Caption = "Generando Transferencia " & NumeroRemesa & " del banco " & Rs!Cta
            Case 2
                Label2.Caption = "Generando Pago Domiciliado " & NumeroRemesa & " del banco " & Rs!Cta
            Case 3
                Label2.Caption = "Generando Confirming " & NumeroRemesa & " del banco " & Rs!Cta
            Case 4
                Label2.Caption = "Generando cobros anticipados: " & NumeroRemesa & " del banco " & Rs!Cta
        End Select
        Me.Refresh
        DoEvent2
    
    
        If Opcion = 0 Then
            'Ahora insertamos la remesa
            cad = "INSERT INTO transferencias (tipotrans, codigo, anyo, fecha,situacion,codmacta,descripcion,subtipo,concepto,solopago) "
            If TipoTrans2 = 0 Then
                cad = cad & " VALUES (1,"
            ElseIf TipoTrans2 = 4 Then
                'Anticipados
                cad = cad & " VALUES (2,"
            Else
                cad = cad & " VALUES (0,"
            End If
            cad = cad & NumeroRemesa & "," & Year(CDate(txtFecha(4).Text)) & ",'" & Format(txtFecha(4).Text, FormatoFecha) & "','A','"
            cad = cad & Rs!Cta & "','" & DevNombreSQL(txtRemesa.Text) & "',"
            Select Case TipoTrans2
                Case 0, 1 ' transf.abonos y transf.pagos
                    cad = cad & "0,"
                    cad = cad & DBSet(cboConcepto2.ItemData(cboConcepto.ListIndex), "N") & ",0)"
                Case 2 ' transf.pagos domiciliados
                    cad = cad & "1,"
                    If chkFecha.Value = 0 Then
                        cad = cad & "0,0)"
                    Else
                        cad = cad & "1,0)"
                    End If
                Case 3 ' pagos confirming
                    cad = cad & "2,"
                    If chkFecha.Value = 0 Then
                        cad = cad & "0"
                    Else
                        cad = cad & "1"
                    End If
                    cad = cad & "," & IIf(Me.chkConfirmingPP.Value = 1, 1, 0) & ")"
                    
                Case 4
                    'Cobros anticipados
                    cad = cad & "0,0,0)"

            End Select
                
            Conn.Execute cad
            
        Else
            'Paso la remesa a estado: A
            'Vuelvo a poner los vecnimientos a NULL para poder
            'meterlos luego
            
            '---remesa estado A
            
            cad = "UPDATE transferencias SET Situacion = 'A'"
            cad = cad & ", descripcion ='" & DevNombreSQL(Text2.Text) & "'"
            cad = cad & ", fecha= " & DBSet(txtFecha(5).Text, "F")
            cad = cad & ", codmacta= " & DBSet(txtCuentas(3).Text, "T")
            
            If TipoTrans2 = 3 Then cad = cad & ", solopago= " & Me.chkConfirmingPP2.Value
            
            cad = cad & " WHERE codigo=" & NumeroRemesa
            cad = cad & " AND anyo =" & Year(CDate(txtFecha(4).Text))
            If Not Ejecuta(cad) Then Exit Function
            
            
            If TipoTrans2 = 0 Or TipoTrans2 = 4 Then ' transferencias de abonos
                cad = "UPDATE cobros SET siturem=NULL, transfer=NULL, anyorem=NULL ,tiporem =NULL "
                cad = cad & " ,fecultco=NULL, impcobro = NULL "
                cad = cad & " WHERE transfer = " & NumeroRemesa
                cad = cad & " AND anyorem=" & Year(CDate(txtFecha(4).Text)) '& " AND tiporem = 1"
                If Not Ejecuta(cad) Then Exit Function
            Else
                ' transferencias de pagos
                cad = "UPDATE pagos SET situdocum=NULL,nrodocum=NULL,anyodocum=NULL "
                cad = cad & " ,fecultpa=NULL, imppagad = NULL "
                cad = cad & " WHERE nrodocum = " & NumeroRemesa
                cad = cad & " AND anyodocum=" & Year(CDate(txtFecha(4).Text))
                If Not Ejecuta(cad) Then Exit Function
            
            End If
        End If
        
        'Ahora cambiamos los cobros y les ponemos la remesa
        ' transferencias de abonos y anticipos
        If TipoTrans2 = 0 Or TipoTrans2 = 4 Then
            If Opcion = 0 Then
                cad = "UPDATE cobros SET siturem = 'A', transfer= " & NumeroRemesa & ", anyorem =" & Year(CDate(txtFecha(4).Text))
            Else
                cad = "UPDATE cobros SET siturem = 'A', transfer= " & NumeroRemesa & ", anyorem =" & Year(CDate(txtFecha(5).Text))
            End If
            If TipoTrans2 = 4 Then cad = cad & ", reciboanticipado=1 "
        
        
            'Para cada cobro UPDATE
            For J = 1 To lwCobros.ListItems.Count
               With lwCobros.ListItems(J)
                    If .Checked And .SubItems(8) = Rs!Cta Then   ' si el subitem es del banco
                        C = " WHERE numserie = '" & .Text & "' and numfactu = "
                        C = C & Val(.SubItems(1)) & " and fecfactu ='" & Format(.SubItems(2), FormatoFecha)
                        C = C & "' AND numorden =" & .SubItems(3)
                        C = cad & C
                        Conn.Execute C
                    Else
                   
                    End If
               End With
            Next J
            espera 0.5
            
            If Opcion = 0 Then
                'Hacemos un select sum para el importe
                cad = "Select sum(impvenci),sum(coalesce(impcobro,0)),sum(coalesce(gastos,0)) from cobros "
                cad = cad & " WHERE transfer=" & NumeroRemesa
                cad = cad & " AND anyorem =" & Year(CDate(txtFecha(4).Text))
                
                miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                cad = "UPDATE pagos SET situdocum = 'A', nrodocum= " & NumeroRemesa & ", anyodocum =" & Year(CDate(txtFecha(4).Text))
                cad = cad & ", ctabanc1 ='" & txtCuentas(2).Text & "'"
                'Pagos domiciliados, pude pagar por otro banco. Ponemos el banco seleccionado aqui
                If TipoTrans2 = 2 Then cad = cad & " , emitdocum = 1"
            Else
                cad = "UPDATE pagos SET situdocum = 'A', nrodocum= " & NumeroRemesa & ", anyodocum =" & Year(CDate(txtFecha(5).Text))
                cad = cad & ", ctabanc1 ='" & txtCuentas(3).Text & "'"
                If TipoTrans2 = 2 Then cad = cad & " , emitdocum = 1"
            End If
        
            'Para cada cobro UPDATE
            For J = 1 To lwCobros.ListItems.Count
               With lwCobros.ListItems(J)
                    If .Checked And .SubItems(8) = Rs!Cta Then   ' si el subitem es del banco
                        C = " WHERE numserie = '" & .Text & "' and numfactu = "
                        C = C & DBSet(.SubItems(1), "T") & " and fecfactu ='" & Format(.SubItems(2), FormatoFecha)
                        C = C & "' AND numorden =" & .SubItems(3)
                        C = C & " and codmacta = " & DBSet(.SubItems(9), "T")
                    
                        C = cad & C
                        Conn.Execute C
                    Else
     
                    End If
               End With
            Next J
            espera 0.5
            
            If Opcion = 0 Then
                'Hacemos un select sum para el importe
                cad = "Select sum(impefect),sum(coalesce(imppagad,0)) from pagos "
                cad = cad & " WHERE nrodocum=" & NumeroRemesa
                cad = cad & " AND anyodocum =" & Year(CDate(txtFecha(4).Text))
                
                miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            cad = "UPDATE transferencias SET importe=" & C
            cad = cad & " WHERE codigo=" & NumeroRemesa
            cad = cad & " AND anyo =" & Year(CDate(txtFecha(4).Text))
            Conn.Execute cad
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
    Select Case TipoTrans2
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
        Sql = "select codmacta,count(*) from cobros where"
        Sql = Sql & " transfer = " & lw1.ListItems(lw1.SelectedItem.Index).Text
        Sql = Sql & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)

    Else
        Sql = "select codmacta,count(*) from pagos where nrodocum = " & lw1.ListItems(lw1.SelectedItem.Index).Text
        Sql = Sql & " and anyodocum = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1)
    End If
    Sql = Sql & " group by codmacta having count(*) >1"
     miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Msg = ""
    RC = ""
    I = 0
    While Not miRsAux.EOF
        I = I + 1
        Msg = Msg & miRsAux!codmacta & "|"
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    
    If I = 0 Then
        ComprobacionesAgrupaFicheroTransfer = 0
        Exit Function
    End If
    
    'Comprobare la cuenta IBAN
    MsgErr = ""
    K = 0
    NumRegElim = 0
    For J = 1 To I
        Sql = " AND codmacta = '" & RecuperaValor(Msg, CInt(J)) & "' GROUP BY iban "
        If Abonos Then
            Sql = " transfer = " & lw1.ListItems(lw1.SelectedItem.Index).Text & " and anyorem = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1) & Sql
            Sql = "select  iban,count(*) ctos from cobros where  " & Sql
            
        Else
            Sql = " nrodocum = " & lw1.ListItems(lw1.SelectedItem.Index).Text & " and anyodocum = " & lw1.ListItems(lw1.SelectedItem.Index).SubItems(1) & Sql
            Sql = "select iban,count(*) ctos from pagos where  " & Sql
        End If
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        If miRsAux.EOF Then
            Err.Raise 513, , "no se encuentran vtos. en select: " & Mid(miRsAux.Source, 30)
        Else
            C = DevuelveDesdeBD("concat(codmacta,' ',nommacta)", "cuentas", "codmacta", RecuperaValor(Msg, CInt(J)), "T") & "    Nº:" & Format(miRsAux!ctos, "0000")
            Sql = miRsAux!IBAN
            miRsAux.MoveNext
            If miRsAux.EOF Then
                'perfeto. SOLO un IBAN
                RC = RC & "X"
                NumRegElim = NumRegElim + Val(Right(C, 4))
                Sql = ""
            Else
                Sql = C & vbCrLf & Sql & " // " & miRsAux!IBAN & vbCrLf
            End If
        End If
        If Sql <> "" Then MsgErr = MsgErr & Sql & vbCrLf
        miRsAux.Close
    Next J
    
    
    Set miRsAux = Nothing
    
    If MsgErr <> "" Then
        MsgErr = "IBAN distinto para agrupacion por cliente:" & vbCrLf & MsgErr
        MsgBox MsgErr, vbExclamation
        
    Else
        
        Sql = "Fecha de pago: " & RecuperaValor(CadenaDesdeOtroForm, 1) & vbCrLf
        RC = IIf(TipoTrans2 = 0, "Clientes", "Proveedores") & " agrupados: " & Len(RC) & vbCrLf
        RC = RC & "Total vencimientos agrupados: " & NumRegElim & vbCrLf
        Sql = Sql & RC & " ¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then ComprobacionesAgrupaFicheroTransfer = I
    End If
eComrpobacionesAgrupacionFichero:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    RC = ""
End Function

