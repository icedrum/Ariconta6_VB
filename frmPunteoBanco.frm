VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPunteoBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punteo bancario"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17250
   Icon            =   "frmPunteoBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   17250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   1590
      TabIndex        =   44
      Top             =   120
      Width           =   13065
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar"
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
         Left            =   11640
         TabIndex        =   47
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1500
         TabIndex        =   46
         Top             =   240
         Width           =   7815
      End
      Begin VB.CheckBox chkElimmFich 
         Caption         =   "Eliminar fichero "
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
         Left            =   9510
         TabIndex        =   45
         Top             =   240
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   1740
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1200
         Picture         =   "frmPunteoBanco.frx":000C
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
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
         Index           =   7
         Left            =   180
         TabIndex        =   48
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   3180
      TabIndex        =   49
      Top             =   2100
      Width           =   11535
      Begin VB.TextBox txtDatos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5955
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   52
         Text            =   "frmPunteoBanco.frx":685E
         Top             =   180
         Width           =   11235
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Integrar"
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
         Left            =   9120
         TabIndex        =   51
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Volver"
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
         Left            =   10320
         TabIndex        =   50
         Top             =   6240
         Width           =   1095
      End
   End
   Begin VB.Frame FrameGenera 
      Height          =   4935
      Left            =   4620
      TabIndex        =   20
      Top             =   2040
      Width           =   7665
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MaxLength       =   15
         TabIndex        =   25
         Text            =   "000000000000000"
         Top             =   2880
         Width           =   1965
      End
      Begin VB.CommandButton cmdAtoCancelar 
         Caption         =   "Cancelar"
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
         Left            =   6240
         TabIndex        =   29
         Top             =   4440
         Width           =   1155
      End
      Begin VB.CommandButton cmdAstoAceptar 
         Caption         =   "Aceptar"
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
         Left            =   4980
         TabIndex        =   28
         Top             =   4440
         Width           =   1155
      End
      Begin VB.TextBox txtFec 
         BeginProperty Font 
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
         TabIndex        =   22
         Text            =   "99/99/9999"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox Text10 
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   1440
         Width           =   6225
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1440
         Width           =   885
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2880
         Width           =   5085
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   2160
         Width           =   6225
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2160
         Width           =   885
      End
      Begin VB.TextBox Text5 
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   3600
         Width           =   5625
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3600
         Width           =   1405
      End
      Begin VB.Label Label10 
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
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1680
         Picture         =   "frmPunteoBanco.frx":6864
         Top             =   3330
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1230
         Picture         =   "frmPunteoBanco.frx":D0B6
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   900
         Picture         =   "frmPunteoBanco.frx":13908
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Si no pone contrapartida podrá añadir más de una línea en el asiento"
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
         TabIndex        =   38
         Top             =   3960
         Width           =   7275
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
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   37
         Top             =   540
         Width           =   705
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmPunteoBanco.frx":1A15A
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Ampliación"
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
         Left            =   2280
         TabIndex        =   36
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label5 
         Caption         =   "Contrapartida"
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
         TabIndex        =   32
         Top             =   3360
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   21
         Top             =   180
         Width           =   5505
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   41
      Top             =   120
      Width           =   1095
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   42
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Crear Asiento"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameIntro 
      Height          =   885
      Left            =   150
      TabIndex        =   4
      Top             =   1080
      Width           =   16935
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar punteados"
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
         Left            =   14400
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
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
         Left            =   1560
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   330
         Width           =   1575
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
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   330
         Width           =   5175
      End
      Begin VB.TextBox txtFec 
         BeginProperty Font 
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
         Left            =   12930
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   330
         Width           =   1275
      End
      Begin VB.TextBox txtFec 
         BeginProperty Font 
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
         Left            =   10170
         TabIndex        =   1
         Text            =   "99/99/9999"
         Top             =   330
         Width           =   1275
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   12630
         Picture         =   "frmPunteoBanco.frx":1A1E5
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   9870
         Picture         =   "frmPunteoBanco.frx":1A270
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Left            =   1140
         Picture         =   "frmPunteoBanco.frx":1A2FB
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fin"
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
         Left            =   11550
         TabIndex        =   6
         Top             =   360
         Width           =   1065
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
         Left            =   8550
         TabIndex        =   5
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
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
         Left            =   240
         TabIndex        =   8
         Top             =   330
         Width           =   915
      End
   End
   Begin VB.Frame FrameDatos 
      Height          =   7515
      Left            =   180
      TabIndex        =   9
      Top             =   2040
      Width           =   16905
      Begin VB.TextBox Text3 
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
         Index           =   2
         Left            =   14550
         TabIndex        =   16
         Top             =   7050
         Width           =   2200
      End
      Begin VB.TextBox Text3 
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
         Left            =   12210
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   7050
         Width           =   2200
      End
      Begin VB.TextBox Text3 
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
         Left            =   9930
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   7050
         Width           =   2200
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   8345
         _ExtentX        =   14711
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "D/H"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Concepto"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6255
         Left            =   8490
         TabIndex        =   11
         Top             =   540
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "D/H"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ampliacion"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "Doble click busca el importe en el otro lado"
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
         Left            =   330
         TabIndex        =   40
         Top             =   7020
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "DIFERENCIA"
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
         Left            =   14550
         TabIndex        =   19
         Top             =   6780
         Width           =   1845
      End
      Begin VB.Label Label3 
         Caption         =   "CONTABILIDAD"
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
         Left            =   12210
         TabIndex        =   18
         Top             =   6780
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "BANCO"
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
         Left            =   9930
         TabIndex        =   17
         Top             =   6780
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Extracto bancario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   4
         Left            =   8490
         TabIndex        =   12
         Top             =   180
         Width           =   5775
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   16680
      TabIndex        =   43
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
End
Attribute VB_Name = "frmPunteoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 314

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmBasico2
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmCo As frmConceptos
Attribute frmCo.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmCC As frmColCtas
Attribute frmCC.VB_VarHelpID = -1



Dim SQL As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Importe As Currency
Dim i As Integer
Dim PrimeraSeleccion As Boolean
Dim ClickAnterior As Byte '0 Empezar 1.-Debe 2.-Haber
    
'Con estas dos variables
Dim ContadorBus As Integer
Dim Checkear As Boolean
Dim De As Currency
Dim Ha As Currency
Dim EstaLW1 As Boolean

Dim CuentaAnterior As String
Dim FechaAnterior As String

Dim NF As Integer
Dim FicheroPpal As String
Dim Cta As String
Dim Saldo As Currency
Dim cad As String


Private Sub Check1_Click()
    CuentaAnterior = Text1.Text
    ConfirmarDatos False
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    CuentaAnterior = Text1.Text
    ConfirmarDatos False
    KEYpress KeyAscii
End Sub

Private Sub ConfirmarDatos(DesdeCuenta As Boolean)
    Screen.MousePointer = vbHourglass
    If Text1.Text <> "" Then
        If CuentaAnterior <> "" Then BloqueoManual False, "PUNTEOB", CuentaAnterior
    
        'Tiene cta.
        'Veamos si la cuenta esta definida en ctas bancarias o no
        SQL = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Text1.Text, "T")
        If SQL <> "" Then
            'Bloqueamos manualamente la tabla, con esa cuenta
            If Not BloqueoManual(True, "PUNTEOB", Text1.Text) Then
                MsgBox "Imposible acceder a puntear la cuenta. Esta bloqueada"
            Else
                Text3(0).Text = "": Text3(1).Text = "": Text3(2).Text = ""
                'Datos ok. Vamos a ver los resultados
                Label1(4).Caption = Text1.Text & " - " & Text2.Text
   '             PonerTamanyo True
                Me.Refresh
                CargarDatosLw True
            End If
        Else
            MsgBox "La cuenta no esta asociada a una cuenta bancaria.", vbExclamation
        End If
    Else
        MsgBox "Introduzca la cuenta ", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargarDatosLw(BorrarImportes As Boolean)

    If txtFec(0).Text = "" Or txtFec(1).Text = "" Then Exit Sub



       'Resetamos importes punteados
       If BorrarImportes Then
            De = 0
            Ha = 0
            Text3(0).Text = "": Text3(1).Text = "": Text3(2).Text = ""
        End If
        PrimeraSeleccion = True
                    
        'Cargamos los datos
        SQL = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
        Conn.Execute SQL
            
        SQL = "DELETE from tmpconext where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        
        SQL = "fechaent >= '" & Format(txtFec(0).Text, FormatoFecha)
        SQL = SQL & "' AND fechaent <= '" & Format(txtFec(1).Text, FormatoFecha) & "'"
        
        CargaDatosConExt Text1.Text, txtFec(0).Text, txtFec(1).Text, SQL, Text2.Text

                    
                    
                    
        Me.Refresh
        CargaBancario
        Me.Refresh
        CargaLineaApuntes
        
        FrameBotonGnral.Enabled = True
        
        Me.Refresh
End Sub



Private Sub cmdAstoAceptar_Click()
Dim NA As Long
Dim SQL As String
Dim Sql1 As String


    If txtFec(2).Text = "" Or Text9.Text = "" Or Text7.Text = "" Then
        MsgBox "Todos los campos, excepto la contrapartida, son obligados", vbExclamation
        Exit Sub
    End If
    
    'Generamos el asiento en errores
    If Not IsDate(txtFec(2).Text) Then
        MsgBox "Fecha incorrecta", vbExclamation
        Exit Sub
    End If
    
    varFecOk = FechaCorrecta2(CDate(txtFec(2).Text))
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            SQL = varTxtFec
        Else
            SQL = "Fechas fuera de ejercicio actual/siguiente"
        End If
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    ' cogemos el nro de asiento dependiendo de la fecha
    Dim Mc As Contadores
    
    Set Mc = New Contadores
    If Mc.ConseguirContador(0, txtFec(2).Text <= vParam.fechafin, False) = 0 Then
        NA = Mc.Contador
    Else
        MsgBox "Error al obtener contador", vbExclamation
        Set Mc = Nothing
        Exit Sub
    End If
    Set Mc = Nothing
    
    'Ahora generemos la cabecera de apunte
    Screen.MousePointer = vbHourglass
    If GenerarCabecera(NA) Then
        CadenaDesdeOtroForm = ""
        If Text4.Text <> "" Then
            frmAsientosHco.DesdeNorma43 = 2
        Else
            frmAsientosHco.DesdeNorma43 = 1
        End If
        frmAsientosHco.ASIENTO = Text9.Text & "|" & txtFec(2).Text & "|" & NA & "|"
        frmAsientosHco.Show vbModal
    End If
    
    ' si el asiento está descuadrado hemos de eliminarlo
    SQL = "select sum(coalesce(timported,0) - coalesce(timporteh,0)) from hlinapu where numasien = " & DBSet(NA, "N") & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
    Sql1 = "select count(*) from hlinapu where numasien = " & NA & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
    If DevuelveValor(SQL) <> 0 Or DevuelveValor(Sql1) = 0 Then
        'Borramos las lineas del apunte
        Screen.MousePointer = vbHourglass
        SQL = "Delete from hlinapu where numasien = " & NA & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
        Conn.Execute SQL
        SQL = "Delete from hcabapu where numasien = " & NA & " and numdiari = " & DBSet(Text9.Text, "N") & " and fechaent = " & DBSet(txtFec(2).Text, "F")
        Conn.Execute SQL
    
        'devolvemos el contador
        Set Mc = New Contadores
        Mc.DevolverContador 0, txtFec(2).Text <= vParam.fechafin, NA
        Set Mc = Nothing
    
    Else
    
    
        'Aumentamos los importes punteados
        Importe = CCur(ListView1.SelectedItem.SubItems(1))
        De = De + Importe
        Ha = Ha + Importe
        PonerImportes
    
        'Puntemos el extracto
        SQL = "UPDATE norma43 SET punteada= 1 WHERE codigo=" & ListView1.SelectedItem.Tag
        Conn.Execute SQL
    
        'Para buscarlo
        NA = ListView1.SelectedItem.Tag
        'Volvemos a cargar todo
        CargarDatosLw False
        'Volvemos a siutar el select item
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Tag = NA Then
                Set ListView1.SelectedItem = ListView1.ListItems(i)
                ListView1.SelectedItem.EnsureVisible
                ListView1_DblClick
                Exit For
            End If
        Next i
    End If
    Me.FrameGenera.visible = False
    Me.FrameDatos.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAtoCancelar_Click()
    Me.FrameGenera.visible = False
    Me.FrameDatos.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub PonerTamanyo(Punteo As Boolean)
    Me.FrameDatos.visible = Punteo
    Me.FrameIntro.visible = Not Punteo
    If Punteo Then
        Me.Height = FrameDatos.Height + 400
        Me.Width = FrameDatos.Width + 100
        If Screen.Width > 12300 Then
            Me.top = 800
            Me.Left = 800
        Else
            Me.top = 0
            Me.Left = 0
        End If
    
    Else
        Me.Height = FrameIntro.Height + 400
        Me.Width = FrameIntro.Width + 100
        If Screen.Width > 12300 Then
            Me.top = 4000
            Me.Left = 4000
        Else
            Me.top = 1000
            Me.Left = 1000
        End If
    End If
          
End Sub


Private Sub CrearAsiento()
    'Crear asiento
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem.Checked Then
        MsgBox "Extracto ya esta punteado", vbExclamation
        Exit Sub
    End If
    
    'Deshabilitamos
    Me.FrameDatos.Enabled = False
    'Limpiamos y ponemos datos
    Me.txtFec(2).Text = Format(ListView1.SelectedItem.Text, "dd/mm/yyyy")
    
    'dIARIO POR DEFECTO DE PARAMETROS
    'Veremos si hay parametros
    SQL = DevuelveDesdeBD("diario43", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    Text9.Text = SQL
    If Text9.Text <> "" Then SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text9.Text, "N")
    Text10.Text = SQL
    
    'Concepto por defecto desde parametros
    SQL = DevuelveDesdeBD("conce43", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    Text6.Text = SQL
    If Text6.Text <> "" Then SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", Text6.Text, "N")
    Text7.Text = SQL
    
    'La ampliacion del concepto viene del extracto bancario
    Text8.Text = ListView1.SelectedItem.SubItems(4)
    
    Text4.Text = "": Text5.Text = ""
    Text11.Text = ""
    Label1(5).Caption = Label1(4).Caption
    'Ponemos visible
    Me.FrameGenera.visible = True
    'Ponemos el foco en doc
    Text11.SetFocus
    
End Sub




Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Command1_Click()
    If EstaLW1 Then
        ListView1_DblClick
    Else
        ListView2_DblClick
    End If
End Sub

Private Sub cmdImportar_Click(Index As Integer)
    Text12.Text = Trim(Text12.Text)
    If Text12.Text = "" Then
        MsgBox "Debes poner el nombre de archivo", vbExclamation
        Exit Sub
    End If
    If Dir(Text12.Text, vbArchive) = "" Then
        MsgBox "Fichero NO existe", vbExclamation
        Exit Sub
    End If
    'Borramos los temporales
    SQL = "Delete from tmpnorma43 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    Screen.MousePointer = vbHourglass
    If ProcesarFichero Then
        NumRegElim = 1
        'Ahora procesamos los datos
        ProcesarDatos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    Me.Icon = frmppal.Icon

    'La toolbar
    With Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 44
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    

    FrameGenera.visible = False
    FrameIntro.Enabled = True
    FrameBotonGnral.Enabled = False
    Frame1.Enabled = True
    Frame2.visible = False
    
    
    Text1.Text = ""
    Text2.Text = ""
    txtFec(0).Text = ""
    txtFec(1).Text = ""
    
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub





Private Sub Form_Unload(Cancel As Integer)
    'Desbloqueamos
    BloqueoManual False, "PUNTEOB", Text1.Text
End Sub



Private Sub frmC_Selec(vFecha As Date)
    txtFec(CInt(txtFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    Text4.Text = RecuperaValor(CadenaSeleccion, 1)
    Text5.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCo_DatoSeleccionado(CadenaSeleccion As String)
    Text6.Text = RecuperaValor(CadenaSeleccion, 1)
    Text7.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text1.Text = RecuperaValor(CadenaSeleccion, 1)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    Text9.Text = RecuperaValor(CadenaSeleccion, 1)
    Text10.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click()
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub

Private Sub Image2_Click()
    Set frmCo = New frmConceptos
    frmCo.DatosADevolverBusqueda = "0|1|"
    frmCo.Show vbModal
    Set frmCo = Nothing
End Sub

Private Sub Image3_Click()
    Set frmCC = New frmColCtas
    frmCC.DatosADevolverBusqueda = "0|1"
    frmCC.ConfigurarBalances = 3  'NUEVO
    frmCC.Show vbModal
    Set frmCC = Nothing
End Sub

Private Sub Image4_Click()

    cd1.CancelError = False
    cd1.DialogTitle = "Archivo banco NORMA 43"
    cd1.ShowOpen
    If cd1.FileName <> "" Then Text12.Text = cd1.FileName
    
End Sub

Private Sub imgCuentas_Click()
    
    Set frmCta = New frmBasico2
    AyudaCuentasBancarias frmCta
    Set frmCta = Nothing
    
    PonerFoco Text1
End Sub

Private Sub imgppal_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    txtFec(0).Tag = Index
    If txtFec(Index).Text <> "" Then
        If IsDate(txtFec(Index).Text) Then frmC.Fecha = CDate(txtFec(Index).Text)
    End If
    frmC.Show vbModal
    Set frmC = Nothing
End Sub


Private Sub ListView1_Click()
    EstaLW1 = True
End Sub

Private Sub ListView1_DblClick()
Dim J As Integer
Dim Find As Boolean
Dim Fin As Long

    EstaLW1 = True
    If ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    J = ListView2.SelectedItem.Index
    Find = False
    Fin = ListView2.ListItems.Count
    Do
        For i = J To Fin
            If ListView2.ListItems(i).SubItems(1) = ListView1.SelectedItem.SubItems(1) Then
                If ListView2.ListItems(i).SubItems(2) <> ListView1.SelectedItem.SubItems(2) Then
                    'Ha encontrado con el mismo importe y signos distintos D-H
                    Set ListView2.SelectedItem = ListView2.ListItems(i)
                    ListView2.SelectedItem.EnsureVisible
                    Find = True
                    Exit For
                End If
            End If
        Next i
        If Not Find Then
            If J > 1 Then
                Fin = J
                J = 1
            Else
                Find = True
            End If
        End If
                
    Loop Until Find
End Sub


Private Sub ListView2_Click()
    EstaLW1 = False
End Sub

Private Sub ListView2_DblClick()
Dim J As Integer
Dim Find As Boolean
Dim Fin As Long

    EstaLW1 = False
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then
        J = 0
    Else
        J = ListView1.SelectedItem.Index + 1
    End If
    Find = False
    Fin = ListView1.ListItems.Count
    Do
        For i = J To Fin
            If ListView1.ListItems(i).SubItems(1) = ListView2.SelectedItem.SubItems(1) Then
                If ListView1.ListItems(i).SubItems(2) <> ListView2.SelectedItem.SubItems(2) Then
                    'Ha encontrado con el mismo importe y signos distintos D-H
                    Set ListView1.SelectedItem = ListView1.ListItems(i)
                    ListView1.SelectedItem.EnsureVisible
                    Find = True
                    Exit For
                End If
            End If
        Next i
        If Not Find Then
            If J > 1 Then
                Fin = J
                J = 1
            Else
                Find = True
            End If
        End If
                
    Loop Until Find

End Sub

Private Sub Text1_GotFocus()
    PonFoco Text1
    CuentaAnterior = Text1.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Or KeyCode = 187 Then
        KeyCode = 0
        Text1.Text = ""
        imgCuentas_Click
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus()
Dim RC As String


    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "+" Then Text1.Text = ""
    If Text1.Text = "" Then
        Text2.Text = ""
        Exit Sub
    Else
         RC = Text1.Text
         If CuentaCorrectaUltimoNivel(RC, SQL) Then
             Text1.Text = RC
             Text2.Text = SQL
             
             ConfirmarDatos True
             CuentaAnterior = Text1.Text
         Else
             MsgBox SQL, vbExclamation
             Text2.Text = ""
         End If
         If Text2.Text = "" Then PonerFoco Text1
         
    End If
             
End Sub


Private Sub PonerFoco(Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub Text11_GotFocus()
    PonFoco Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub



Private Sub Text4_GotFocus()
    PonFoco Text4
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Or KeyCode = 187 Then
        KeyCode = 0
        Text1.Text = ""
        Image3_Click
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text4_LostFocus()
Dim RC As String

    Text4.Text = Trim(Text4.Text)
    If Text4.Text = "+" Then Text4.Text = ""
    If Text4.Text = "" Then
        Text5.Text = ""
    Else
        RC = Text4.Text
        If CuentaCorrectaUltimoNivel(RC, SQL) Then
            Text4.Text = RC
            Text5.Text = SQL
        Else
            MsgBox SQL, vbExclamation
            Text5.Text = ""
            Text4.Text = ""
            Text4.SetFocus
        End If
    End If
End Sub



Private Sub Text6_GotFocus()
    PonFoco Text6
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text6_LostFocus()
   With Text6
        .Text = Trim(.Text)
        i = 1
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "El valor debe ser numérico: " & .Text, vbExclamation
            Else
                 If Val(.Text) >= 900 Then
                    MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
                Else
                    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", .Text, "N")
                    If SQL = "" Then
                        MsgBox "Concepto NO encontrado: " & .Text, vbExclamation
                    Else
                        Text7.Text = SQL
                        i = 0
                    End If
                End If
            End If
        Else
            'Igual a "" luego pasamos a otro campo en la tabulacion
            i = 2
        End If
        If i > 0 Then
            .Text = ""
            Text7.Text = ""
            If i = 1 Then Text6.SetFocus
        End If
    End With
End Sub

Private Sub Text8_GotFocus()
    PonFoco Text8
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text9_GotFocus()
    PonFoco Text9
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text9_LostFocus()
    With Text9
        .Text = Trim(.Text)
        i = 1
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "El valor debe ser numérico: " & .Text, vbExclamation
            Else
                SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", .Text, "N")
                If SQL = "" Then
                    MsgBox "Concepto NO encontrado: " & .Text, vbExclamation
                Else
                    Text10.Text = SQL
                    i = 0
                End If
            End If
        Else
            'Igual a "" luego pasamos a otro campo
            i = 2
        End If
        If i > 0 Then
            .Text = ""
            Text10.Text = ""
            If i = 1 Then Text9.SetFocus
        End If
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    CrearAsiento
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select

End Sub

Private Sub txtfec_GotFocus(Index As Integer)
    PonFoco txtFec(Index)
    FechaAnterior = txtFec(Index).Text
End Sub
'++
Private Sub txtfec_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFecha KeyAscii, 0
            Case 1: KEYFecha KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgppal_Click (Indice)
End Sub

'++

Private Sub txtfec_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtFec(Index).Text = Trim(txtFec(Index).Text)
    Mal = True

    If txtFec(Index).Text = "" Then Exit Sub

        If Not EsFechaOK(txtFec(Index)) Then
            MsgBox "No es una fecha correcta", vbExclamation
        Else
            Mal = False
        End If
    If Mal Then
        PonerFoco txtFec(Index)
    Else
        If txtFec(Index).Text <> FechaAnterior Then ConfirmarDatos True
    End If
    
End Sub



Private Sub CargaBancario()

    ListView1.ListItems.Clear
    SQL = "Select * from norma43 where"
    SQL = SQL & " codmacta ='" & Text1.Text & "'"
    SQL = SQL & " AND fecopera >='" & Format(txtFec(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecopera <='" & Format(txtFec(1).Text, FormatoFecha) & "'"
    'OCultar/mostrar punteados
    If Check1.Value = 0 Then
        'Ocultar los ya puntedos
        SQL = SQL & " AND Punteada = 0 "
    End If
    SQL = SQL & " ORDER BY fecopera,codigo"

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = Format(Rs!fecopera, "dd/mm/yyyy")
        'Importe Debe
        If Not IsNull(Rs!ImporteD) Then
            Importe = Rs!ImporteD
            SQL = "D"
        Else
            'Importe HABER
            If Not IsNull(Rs!ImporteH) Then
                Importe = Rs!ImporteH
                SQL = "H"
            Else
                SQL = "XX"
            End If
        End If
        ItmX.SubItems(1) = Format(Importe, FormatoImporte)
        ItmX.SubItems(2) = SQL
        ItmX.SubItems(3) = Format(Rs!Saldo, FormatoImporte)
        ItmX.SubItems(4) = Rs!Concepto
        ItmX.ListSubItems(4).ToolTipText = DBLet(Rs!Concepto, "T")
        
        ItmX.Tag = Rs!Codigo
        ItmX.Checked = (Rs!punteada = 1)
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
End Sub


Private Sub CargaLineaApuntes()

    ListView2.ListItems.Clear
    SQL = "Select numasien,fechaent,numdiari,linliapu,ampconce,timported,timporteh,punteada,saldo FROM tmpconext"
    SQL = SQL & " WHERE codusu = " & vUsu.Codigo
    
    If Check1.Value = 0 Then
        'Ocultar los ya puntedos
        SQL = SQL & " AND Punteada = '' "
    End If
    SQL = SQL & " ORDER BY pos"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set ItmX = ListView2.ListItems.Add()
        ItmX.Text = Format(Rs!FechaEnt, "dd/mm/yyyy")
        'Importe Debe
        SQL = " "
        If Not IsNull(Rs!timported) Then
            Importe = Format(Rs!timported, FormatoImporte)
            SQL = "D"
        Else
            'Importe HABER
            If Not IsNull(Rs!timporteH) Then
                Importe = Rs!timporteH
                SQL = "H"
            Else
                Importe = 0
                SQL = "XX"
            End If
        End If
        ItmX.SubItems(1) = Format(Importe, FormatoImporte)
        ItmX.SubItems(2) = SQL
        ItmX.SubItems(3) = Format(Rs!Saldo, FormatoImporte)
        ItmX.SubItems(4) = DBLet(Rs!Ampconce, "T")
        ItmX.ListSubItems(4).ToolTipText = DBLet(Rs!Ampconce, "T")

        
        ItmX.Tag = Rs!NumAsien & "|" & Rs!NumDiari & "|" & Rs!Linliapu & "|"
        ItmX.Checked = (Rs!punteada <> "")
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub




'----------------- PUNTEOS

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
EstaLW1 = True
Screen.MousePointer = vbHourglass
    Set ListView1.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, True
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 1 Then
            If PrimeraSeleccion Then
                BusquedaEnHaber
                PrimeraSeleccion = False
                ClickAnterior = 1
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault

End Sub


Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Screen.MousePointer = vbHourglass
    EstaLW1 = False
    Set ListView2.SelectedItem = Item
    'Ponemos a true o a false
    PunteaEnBD Item, False
    'Comprobamos primero que esta checkeado
    If Item.Checked = False Then
        PrimeraSeleccion = True
        ClickAnterior = 0
    Else
        If ClickAnterior <> 2 Then
            If PrimeraSeleccion Then
                BusquedaEnDebe
                PrimeraSeleccion = False
                ClickAnterior = 2
            Else
                PrimeraSeleccion = True
                ClickAnterior = 0
            End If
        Else
            PrimeraSeleccion = True
        End If
    
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub BusquedaEnHaber()
    ContadorBus = 1
    Checkear = False
    Do
        i = 1
        While i <= ListView2.ListItems.Count
            'Comprobamos k no esta chekeado
            If Not ListView2.ListItems(i).Checked Then
                'K tiene el mismo importe
                If ListView1.SelectedItem.SubItems(1) = ListView2.ListItems(i).SubItems(1) Then
                    'K no sean DEBE o HABER los dos
                    Checkear = (ListView1.SelectedItem.SubItems(2) <> ListView2.ListItems(i).SubItems(2))

                    If Checkear Then
                        'Tiene el mismo importe y no esta chequeado
                        Set ListView2.SelectedItem = ListView2.ListItems(i)
                        ListView2.SelectedItem.EnsureVisible
                        ListView2.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            i = i + 1
        Wend
        ContadorBus = ContadorBus + 1
        Loop Until ContadorBus > 2
End Sub



Private Sub BusquedaEnDebe()
    ContadorBus = 1
    Checkear = False
    Do
        i = 1
        While i <= ListView1.ListItems.Count
            If ListView2.SelectedItem.SubItems(1) = ListView1.ListItems(i).SubItems(1) Then
                'Lo hemos encontrado. Comprobamos que no esta chequeado
                If Not ListView1.ListItems(i).Checked Then
                    'Tiene el mismo importe y no son debe o haber
                    Checkear = (ListView2.SelectedItem.SubItems(2) <> ListView1.ListItems(i).SubItems(2))

                    If Checkear Then
                        Set ListView1.SelectedItem = ListView1.ListItems(i)
                        ListView1.SelectedItem.EnsureVisible
                        ListView1.SetFocus
                        Beep
                        Exit Sub
                    End If
                End If
            End If
            i = i + 1
        Wend
        ContadorBus = ContadorBus + 1
    Loop Until ContadorBus > 2
End Sub



Private Sub PunteaEnBD(ByRef IT As ListItem, EnDEBE As Boolean)
Dim RC As String
On Error GoTo EPuntea
    
    
    If Not EnDEBE Then
        'ASientos
        'Actualizamos en DOS tablas, en la tmp y en la hcoapuntes
        SQL = "UPDATE hlinapu SET "
        If IT.Checked Then
            RC = "1"
            Importe = 1
            Else
            RC = "0"
            Importe = -1
        End If
        Importe = Importe * CSng(IT.SubItems(1))
        If EnDEBE Then
            De = De + Importe
        Else
            Ha = Ha + Importe
        End If
        SQL = SQL & " punteada = " & RC
        SQL = SQL & " WHERE fechaent='" & Format(IT.Text, FormatoFecha) & "'"
        SQL = SQL & " AND numasien="
        RC = RecuperaValor(IT.Tag, 1)
        SQL = SQL & RC & " AND numdiari ="
        RC = RecuperaValor(IT.Tag, 2)
        SQL = SQL & RC & " AND linliapu ="
        RC = RecuperaValor(IT.Tag, 3)
        SQL = SQL & RC
        
        
        
        
    Else
        'En Norma 43
        
        If IT.Checked Then
            RC = "1"
            Importe = 1
            Else
            RC = "0"
            Importe = -1
        End If
        Importe = Importe * CSng(IT.SubItems(1))
        If EnDEBE Then
            De = De + Importe
        Else
            Ha = Ha + Importe
        End If
        SQL = "UPDATE norma43 SET punteada= " & RC & " WHERE codigo=" & IT.Tag
        
    End If
    
    Conn.Execute SQL
    
    'Ponemos los importes
    PonerImportes

    
    Exit Sub
EPuntea:
    MuestraError Err.Number, "Accediendo BD para puntear", Err.Description
End Sub

Private Sub PonFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Function GenerarCabecera(NumAsi As Long) As Boolean
Dim cad As String

    On Error GoTo EGenerarCabecera
    GenerarCabecera = False
    
    '-------------------------------------------------------------------------
    'Insertamos cabecera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
    'Ejemplo
    ' 1, '2003-11-25', 1, 1, NULL, 'misobs')
    SQL = SQL & Text9.Text & ",'" & Format(CDate(txtFec(2).Text), FormatoFecha) & "'," & NumAsi & ","
    'Observaciones
    SQL = SQL & "'Asiento generado desde punteo bancario por " & vUsu.Nombre & " el " & Format(Now, "dd/mm/yyyy") & "',"
    '
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Punteo Bancario')"
    Conn.Execute SQL
    
    '-----------------------------------------------------------------------------
    'La linea del asiento
    'Hemos puesto hlinapu mas atras para poder cambiarla
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, numdocum,"
    SQL = SQL & " ampconce, codconce, linliapu, codmacta, timporteD, timporteH, ctacontr, codccost, idcontab, punteada) VALUES ("
    
    'Ejemplo valores
    '1, '2001-01-20', 0, 0, '0', NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, 0)"
    SQL = SQL & Text9.Text & ",'" & Format(CDate(txtFec(2).Text), FormatoFecha) & "'," & NumAsi & ","
    '          dcumento
    SQL = SQL & DBSet(Text11.Text, "T") & ","
    
    'Ampliacion concepto
    cad = Mid(Text7.Text & " " & Text8.Text, 1, 30)
    SQL = SQL & DBSet(cad, "T") & ","
    
    'Concepto
    SQL = SQL & Text6.Text & ","
    
    'El importe
    Importe = CCur(ListView1.SelectedItem.SubItems(1))
    cad = "1,'" & Text1.Text & "',"
    If ListView1.SelectedItem.SubItems(2) = "H" Then
        'Va al debe
        cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
    Else
        cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
    End If
    
    'Contrapartida
    If Text4.Text <> "" Then
        cad = cad & ",'" & Text4.Text & "'"
    Else
        cad = cad & ",NULL"
    End If
    
    'y la punteamos
    cad = SQL & cad & ",NULL,'CONTAB',1)"
    Conn.Execute cad
    
    'Si tiene contrapartida entonces genero la segunda linea de apuntes
    ' k sera la de la contrapartida, con el importe el mismo al lado contrario
    ' el mismo concepto
    If Text4.Text <> "" Then
        'SI TIENE
            cad = "2,'" & Text4.Text & "',"
            'En la de arriba es igual a H
            If ListView1.SelectedItem.SubItems(2) = "D" Then
                'Va al debe
                cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
            Else
                cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
            End If
            
            'Contrapartida es la del banco
            cad = cad & ",'" & Text1.Text & "'"
            
            'y NO la punteamos
            cad = SQL & cad & ",NULL,'CONTAB',0)"
            Conn.Execute cad
    End If
    GenerarCabecera = True
    Exit Function
EGenerarCabecera:
    MuestraError Err.Number, Err.Description
End Function



Private Sub PonerImportes()

    If De <> 0 Then
        Text3(0).Text = Format(De, FormatoImporte)
        Else
        Text3(0).Text = ""
    End If
    If Ha <> 0 Then
        Text3(1).Text = Format(Ha, FormatoImporte)
        Else
        Text3(1).Text = ""
    End If
    Importe = De - Ha
    If Importe <> 0 Then
        Text3(2).Text = Format(Importe, FormatoImporte)
        Else
        Text3(2).Text = ""
    End If
End Sub

'############################################################
'  PARTE CORRESPONDIENTE A LA IMPORTACION DE DATOS NORMA 34
'############################################################

Private Function ProcesarFichero() As Boolean
Dim Fin As Boolean
Dim cad As String

On Error GoTo EProcesarFichero
    'Abrimos el fichero para lectura
    ProcesarFichero = False
    NF = FreeFile
    FicheroPpal = "|"
    Open Text12.Text For Input As #NF
    While Not EOF(NF)
        Line Input #NF, SQL
        If SQL <> "" Then
                                        'Separador de lineas
            FicheroPpal = FicheroPpal & SQL & "|"
        End If
    Wend
    Close #NF
    ProcesarFichero = True
    Exit Function
EProcesarFichero:
    MuestraError Err.Number
End Function


Private Sub ProcesarDatos()
Dim i As Long
Dim CONT As Long
Dim NF As Long
Dim Linea As String
Dim Fichero As String
Dim Primer23 As Boolean
Dim Num22 As Integer  'Para conrolar los asientos k se han realizado
Dim Ampliacion As String
Dim RegistroInsertado As Boolean
Dim Comienzo As Long   'Para cuando vienen varios bancos
Dim Fecha As String   'Fecha importacion datos

Dim ContadorMYSQL As Integer
Dim ContadorRegistrosBanco As Integer

    'Vemos cuantas cuentas trae el extracto
    i = 0
    CONT = 0
    Do
        NF = i + 1
        i = InStr(NF, FicheroPpal, "|11")  'los registros empiezan por 11 para las cuentas
        If i > 0 Then CONT = CONT + 1
    Loop Until i = 0
        
    If CONT = 0 Then
        MsgBox "Error en el fichero. No se ha encontrado registro 11", vbExclamation
        Exit Sub
    End If

    
    
    txtDatos.Text = ""
    Comienzo = 2
    ContadorMYSQL = 1
    ContadorRegistrosBanco = 0
    Cta = ""
    'Ya sabemos cuantas cont hay k tratar
    For i = 1 To CONT
        If i <> CONT Then
            Linea = "|11"
            'Hay mas de un |11 o cuenta bancaria
        Else
            'Una unica cta bancaria en este fichero
            Linea = "|88"
        End If
        
        NF = InStr(Comienzo, FicheroPpal, Linea)
        If NF = 0 Then
            MsgBox "imposible situar datos."
            Exit Sub
        End If
        
        Fichero = Mid(FicheroPpal, Comienzo, NF - 1)
        
        Comienzo = NF + 1
                
        'Fecha
        Fecha = ""
        Linea = Mid(Fichero, 31, 2) & "/" & Mid(Fichero, 29, 2) & "/" & Mid(Fichero, 27, 2)
        If IsDate(Linea) Then
            Fecha = "Fecha: " & Space(18) & Format(Linea, "dd/mm/yyyy")
        Else
            Fecha = "Fecha: " & Space(18) & "Error obteniendo fecha"
        End If
        Fecha = Fecha & vbCrLf
                
        'ANTES
        NF = InStr(1, Fichero, "|") 'Es el fin de la primera linea
        
        'Primara linea, la de la cuenta
        Linea = Mid(Fichero, 1, NF - 1) 'pq quitamos el pipe del principio y del final
        
        'De la primera linea obtenemos el numero de cuenta
        Ampliacion = Cta
        FijarCtaContable (Linea)
        If Ampliacion <> Cta Then
            If Ampliacion <> "" Then
                'HA CAMBIADO DE CUENTA DEEEENTRO DEL MISMO Fichero
                ContadorRegistrosBanco = 0
            End If
        End If
        
        If Cta = "" Then
            
            MsgBox "Error obteniendo la cuenta contable asociada. Linea: " & Linea, vbExclamation
            Exit Sub
        Else
            SQL = ""
            If ContadorRegistrosBanco = 0 Then
                If txtDatos.Text <> "" Then txtDatos.Text = txtDatos.Text & SQL & vbCrLf
                For NF = 1 To 98
                    SQL = SQL & "="
                Next NF
                txtDatos.Text = txtDatos.Text & SQL & vbCrLf
                SQL = Mid(Linea, 3, 4) & " " & Mid(Linea, 7, 4) & " ** " & Mid(Linea, 11, 10)
                txtDatos.Text = txtDatos.Text & "Cuenta bancaria: " & SQL & vbCrLf
                Fecha = Fecha & "Cuenta bancaria:   " & SQL & vbCrLf
                txtDatos.Text = txtDatos.Text & "Cuenta contable:   " & Cta & vbCrLf
                Fecha = Fecha & "Cuenta contable:    " & Cta & vbCrLf
                txtDatos.Text = txtDatos.Text & "Linea  F.Opercion   F.Valor         Debe            Haber          Concepto" & vbCrLf
                SQL = ""
                For NF = 1 To 98
                    SQL = SQL & "-"
                Next NF
                txtDatos.Text = txtDatos.Text & SQL & vbCrLf
            Else
                'Es otro trozo de fichero 11| pero de la misma cuenta
                txtDatos.Text = txtDatos.Text & String(98, "=") & vbCrLf
            End If
        End If
        
        'Fijaremos el saldo incial
        SQL = Mid(Linea, 34, 14)
        If Not IsNumeric(SQL) Then
            MsgBox "Error. Se esperaba un importe: " & SQL, vbExclamation
            Exit Sub
        End If
        Saldo = Val(SQL) / 100
        
        'ANTES 25 Noviembre
        'Se trabaja al reves
        'Signo del saldo
        If Mid(Linea, 33, 1) = "1" Then Saldo = Saldo * -1
        
        NF = InStr(1, Fichero, "|") 'Es el fin de la primera linea
        Fichero = Mid(Fichero, NF + 1) '+1 y le quito el pipe
        
        RegistroInsertado = False
        Ampliacion = ""
        Num22 = 0
        'Ya tenemos los primeros datos. Ahora a por los apuntes
        Do
            NF = InStr(1, Fichero, "|")
            Linea = Mid(Fichero, 1, NF - 1)
            Fichero = Mid(Fichero, NF + 1)
            
            SQL = Mid(Linea, 1, 2)
          
            
            If SQL = "22" Then
                If Num22 > 0 Then
                    If Not RegistroInsertado Then
                        If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                    End If
                End If
            
                'Primera parte de la linea de apunte
                If Not ProcesaLineaASiento(Linea, Ampliacion) Then Exit Sub
                RegistroInsertado = False
                Primer23 = True
                Num22 = Num22 + 1
            Else
                If SQL = "23" Then
                    If Primer23 Then
                        Primer23 = False
                        'Insertaremos
                        Ampliacion = ProcesaAmpliacion2(Linea)
                        If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                        RegistroInsertado = True
                    End If
                    
                    
                Else
                    If SQL = "33" Then
                        If Not RegistroInsertado Then
                            If Num22 > 0 Then
                                If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                            End If
                        End If
                        'Fin CTA. Hacer comprobaciones
                        
                        If Not HacerComprobaciones(Linea, ContadorRegistrosBanco, ContadorMYSQL) Then
                            Exit Sub
                        End If
                        Fichero = ""
                       
                    Else
                        'Cualquier otro caso no esta tratado
                        Fichero = ""
                    End If
                End If
            End If
        Loop Until Fichero = ""
        'Kitamos de ppal el valor
    Next i
    
    'Si llega aqui es k ha ido bien.Si no inserta nada, NO muestro los datos
    If ContadorMYSQL > 1 Then PonerModo 1
End Sub

Private Sub FijarCtaContable(ByRef Lin As String)
    SQL = "Select codmacta from bancos"
    SQL = SQL & " where mid(iban,5,4) = " & Mid(Lin, 3, 4) ' entidad
    SQL = SQL & " AND mid(iban,9,4) = " & Mid(Lin, 7, 4) ' oficina
    SQL = SQL & " AND mid(iban,15,10) = '" & Mid(Lin, 11, 10) & "'" ' cuentaba
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cta = ""
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Cta = Rs.Fields(0)
    End If
    Rs.Close
    Set Rs = Nothing
    If Cta = "" Then
        SQL = "Fichero pertenece a la cuenta bancaria:  " & Mid(Lin, 3, 4) & "  " & Mid(Lin, 7, 4) & " ** " & Mid(Lin, 11, 10) & vbCrLf
        SQL = SQL & vbCrLf & "No esta asociada a ninguna cuenta contable."
        MsgBox SQL, vbExclamation
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
Dim SQ As String

    If Index = 1 Then
        PonerModo 0
        Exit Sub
    End If
    
    'Comprobaremos que hay datos para traspasar
    If txtDatos.Text = "" Then
        MsgBox "Datos vacios", vbExclamation
        Exit Sub
    End If
    
    'COntamos los saltos de linea
    NumRegElim = 1
    SQ = txtDatos.Text
    NF = 0
    Do
        NumRegElim = InStr(1, SQ, vbCrLf)
        If NumRegElim > 0 Then
            SQ = Mid(SQ, NumRegElim + 2)  'vbcrlf son DOS caracteres
            NF = NF + 1
            If NF > 5 Then NumRegElim = 0 'Hay mas lineas que las del encabezado
        End If
    Loop Until NumRegElim = 0
    'Fichero comprobacion de saldos
    If NF <= 5 Then
        txtDatos.Text = ""
        If chkElimmFich.Value = 1 Then
            If Dir(Text12.Text, vbArchive) <> "" Then Kill Text12.Text
        End If
        Exit Sub
    End If
    'Comprobamos que no existen datos entre las fechas
    Screen.MousePointer = vbHourglass
    SQ = ""
    Set Rs = New ADODB.Recordset
    SQL = "Select min(fecopera) from tmpnorma43 where codusu = " & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then SQ = " fecopera >='" & Format(Rs.Fields(0), FormatoFecha) & "'"
    End If
    Rs.Close
    SQL = "Select max(fecopera) from tmpnorma43 where codusu = " & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then SQ = SQ & " and fecopera <='" & Format(Rs.Fields(0), FormatoFecha) & "'"
    End If
    Rs.Close
    SQL = "Select count(*) from norma43 where " & SQ
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Not Rs.EOF Then
        NF = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Set Rs = Nothing
    
    If NF > 0 Then
        SQL = "Se han encontrado datos entre las fechas importadas." & vbCrLf
        SQL = SQL & "( " & SQ & " )" & vbCrLf & vbCrLf
        SQL = SQL & "Puede duplicar los datos. ¿ Desea continuar ? " & vbCrLf
        If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If MsgBox("¿Los datos serán importados. ¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
    End If
    
    'Haremos la insercion del registro del banco
    If BloqueoManual(True, "norma43", "clave") Then
        InsertarHcoBanco
        BloqueoManual False, "norma43", ""
        PonerModo 0
        Text1_LostFocus
    Else
        MsgBox "Tabla bloqueada por otro usuario.", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Function InsertarRegistro(Ampliacion As String, ByRef ContadorMYSQL As Integer, ByRef ContadorRegistrosDeUnBanco As Integer) As Boolean
Dim vSql As String
Dim L As String

    On Error GoTo EProcesaAmpliacion
    InsertarRegistro = False
        
    vSql = "INSERT INTO tmpnorma43 (codusu,orden, codmacta, fecopera,"
    vSql = vSql & "fecvalor, importeD, importeH,  concepto,"
    vSql = vSql & "numdocum, saldo) VALUES (" & vUsu.Codigo & "," & ContadorMYSQL & ",'"
    'Numero de apunte
    txtDatos.Text = txtDatos.Text & Right("     " & NumRegElim, 5)
    'Fecha operacion
    L = RecuperaValor(CadenaDesdeOtroForm, 1)
    txtDatos.Text = txtDatos.Text & "  " & Format(L, "dd/mm/yyyy")
    vSql = vSql & Cta & "','" & L
    'Fc Valor
    L = RecuperaValor(CadenaDesdeOtroForm, 2)
    txtDatos.Text = txtDatos.Text & " " & Format(L, "dd/mm/yyyy")
    vSql = vSql & "','" & L
    'Importe DEBE/HABER
    vSql = vSql & "'," & RecuperaValor(CadenaDesdeOtroForm, 3)
    L = RecuperaValor(CadenaDesdeOtroForm, 3)
    NF = 0
    If L = "NULL" Then
        NF = 1
        L = RecuperaValor(CadenaDesdeOtroForm, 4)
    End If
    
    L = TransformaPuntosComas(L)
    L = Format(L, FormatoImporte)
    cad = "              "
    If NF = 0 Then
        'Debe
        txtDatos.Text = txtDatos.Text & "  " & Right("              " & L, 14) & "    " & cad
    Else
        txtDatos.Text = txtDatos.Text & "  " & cad & "    " & Right("              " & L, 14)
    End If
    vSql = vSql & "," & RecuperaValor(CadenaDesdeOtroForm, 4)
    
    'El concepto lo saco de la linea de aqui
    cad = DevNombreSQL(Trim(Ampliacion))  '30 como mucho
    vSql = vSql & ",'" & cad & "',"
    txtDatos.Text = txtDatos.Text & "    " & Ampliacion & vbCrLf
        
    'NumDocum
    vSql = vSql & "'" & RecuperaValor(CadenaDesdeOtroForm, 5) & "'"
    Saldo = Saldo - Importe
    cad = TransformaComasPuntos(CStr(Saldo))
    vSql = vSql & "," & cad & ")"
    'Para la BD
    ContadorMYSQL = ContadorMYSQL + 1
    
    'Para comprobar los regisitros
    ContadorRegistrosDeUnBanco = ContadorRegistrosDeUnBanco + 1
    'El que habia.
    NumRegElim = NumRegElim + 1 'Contador mas uno
    Conn.Execute vSql
    
    InsertarRegistro = True
    Exit Function
EProcesaAmpliacion:
    MuestraError Err.Number, Err.Description & vbCrLf & vSql
       
End Function

'Metere en CadenaDesdeOtroForm, empipado
' Fecha operacion, fecha valor, importeDebe, importe haber, numdocum
Private Function ProcesaLineaASiento(ByRef Lin As String, vAmpliacion As String) As Boolean
Dim Debe As Boolean


    ProcesaLineaASiento = False
    CadenaDesdeOtroForm = ""
    'Fecha operacion
    cad = Mid(Lin, 11, 6)
    cad = "20" & Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5, 2)
    If Not IsDate(cad) Then
        MsgBox "Formato fecha incorrecto", vbExclamation
        Exit Function
    End If
    CadenaDesdeOtroForm = Format(cad, FormatoFecha) & "|"
    
    'Fecha valor
    cad = Mid(Lin, 17, 6)
    cad = "20" & Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5, 2)
    If Not IsDate(cad) Then
        MsgBox "Formato fecha incorrecto", vbExclamation
        Exit Function
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(cad, FormatoFecha) & "|"
    
    
    'Importe
    cad = Mid(Lin, 28, 1)
    Debe = cad = "1"
    cad = Mid(Lin, 29, 14)
    If Not IsNumeric(cad) Then
        MsgBox "Importe registro 22 incorrecto: " & cad, vbExclamation
        Exit Function
    End If
    Importe = Val(cad) / 100
    cad = TransformaComasPuntos(CStr(Importe))
    
    'Importe debe / haber
    If Debe Then
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & cad & "|NULL|"
    Else
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "NULL|" & cad & "|"
    End If
    
    
    'Posible ampliacion
    If Len(Lin) > 53 Then
        vAmpliacion = Trim(Mid(Lin, 53))
        If Len(vAmpliacion) > 30 Then vAmpliacion = Mid(vAmpliacion, 1, 30)
    Else
        vAmpliacion = ""
    End If
    
  '  'Para el arrastrado
  '  'Esto va al reves de la contbiliad, ya k trabajamos con la cuenta del banoc
  '  'ANTES del 25 de Novi
    If Not Debe Then Importe = Importe * -1
  '  If Debe Then Importe = Importe * -1
    'Num docum
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Mid(Lin, 43, 10)
    ProcesaLineaASiento = True
End Function

Private Function ProcesaAmpliacion2(miLinea As String) As String
Dim CADENA As String
Dim C2 As String
Dim Blanco As Boolean
Dim i As Integer

    CADENA = ""
    Blanco = False
    For i = 5 To Len(miLinea)
        C2 = Mid(miLinea, i, 1)
        If C2 = " " Then
             If Not Blanco Then
                CADENA = CADENA & C2
                Blanco = True
            End If
        Else
            Blanco = False
            CADENA = CADENA & C2
        End If
    Next i
    If Len(CADENA) > 30 Then CADENA = Mid(CADENA, 1, 30)
    ProcesaAmpliacion2 = CADENA
End Function

Private Function HacerComprobaciones(ByRef Lin As String, ContadorRegistrosBanco As Integer, TotalRegistrosInsertados As Integer) As Boolean
Dim Ok As Boolean
Dim InsercionesActuales As Integer
    Set Rs = New ADODB.Recordset
    HacerComprobaciones = False
    InsercionesActuales = NumRegElim - 1
    cad = "Select max(orden) from tmpnorma43 where codusu =" & vUsu.Codigo
    cad = cad & " AND codmacta ='" & Cta & "'"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then NF = Rs.Fields(0)
    End If
    Rs.Close
    
    'Numero de lineas insertadas
    Ok = False
    'Total registros en BD
    If NF = ContadorRegistrosBanco Then
        'Coinciden los contadores de insercion parcial
        
        NF = Val(Mid(Lin, 21, 5)) + Val(Mid(Lin, 40, 5))
        If NF = NumRegElim - 1 Then Ok = True
    End If
    If Not Ok Then
        'Error en contadores de registros
        MsgBox "Error en contadores de registo", vbExclamation
        NumRegElim = 0
    End If
    
    
    
    If NumRegElim > 0 Then
        'Obtengo la suma de importes
        cad = "Select sum(importeD)as debe,sum(importeH) as haber,sum(importeD)-sum(importeH) from tmpnorma43 where codusu = " & vUsu.Codigo
        cad = cad & " AND codmacta ='" & Cta & "'"
        'Enero 2009.
        'Estamos admitiendo ficheros que , aun siendo de la misma cuenta, tran mas de una entrada 11| (cabecera de cuenta
        NF = ContadorRegistrosBanco - InsercionesActuales
        cad = cad & " AND orden >" & NF
        Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            cad = CStr(Val(Mid(Lin, 26, 14)) / 100)
            CadenaDesdeOtroForm = DBLet(Rs.Fields(0), "N")
            Ok = (cad = CadenaDesdeOtroForm)
            If Ok Then
                cad = CStr(Val(Mid(Lin, 45, 14)) / 100)
                CadenaDesdeOtroForm = DBLet(Rs.Fields(1), "N")
                Ok = (cad = CadenaDesdeOtroForm)
            End If
            If Ok Then
                Importe = Val(Mid(Lin, 60, 14)) / 100
                If Mid(Lin, 59, 1) = "2" Then Importe = Importe * -1
                
                If ContadorRegistrosBanco = 0 Then
                    cad = "Fichero de comprobación de saldos: " & vbCrLf & vbCrLf
                    cad = cad & "Saldo: " & CStr(Importe)
                    cad = cad & vbCrLf & vbCrLf & vbCrLf
                    cad = cad & "¿Desea eliminar el archivo?"
                    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                        If Dir(Text12.Text, vbArchive) <> "" Then
                            Kill Text12.Text
                            Text12.Text = ""
                        End If
                    End If
                End If
                
            End If
        End If
        Rs.Close
        If Ok Then
            NumRegElim = 1
        Else
            NumRegElim = 0
        End If
    End If
    
    'Si llegamos aqui y numregelim>0 esta bien
    If NumRegElim > 0 Then HacerComprobaciones = True
    Set Rs = Nothing
    
End Function


Private Sub PonerModo(vModo As Byte)
    Select Case vModo
    Case 0
        'Primer frame
        Frame1.Enabled = True
        Frame2.visible = False
    Case 1
        Frame2.visible = True
        Frame1.Enabled = False
    End Select
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 150
    Me.Refresh
End Sub


Private Sub InsertarHcoBanco()
Dim Codigo As Long
    
    Set Rs = New ADODB.Recordset
    Codigo = 0
    SQL = "Select max(codigo) from norma43"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Codigo = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Codigo = Codigo + 1
    
    SQL = "Select * from tmpnorma43 where codusu = " & vUsu.Codigo & " ORDER By Orden"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Cadena de insercion
    SQL = "INSERT INTO norma43 (codigo, codmacta, fecopera, fecvalor, importeD,"
    SQL = SQL & "importeH, concepto, numdocum, saldo, punteada) VALUES ("
    While Not Rs.EOF
        cad = Codigo & ",'" & Rs!codmacta & "','" & Format(Rs!fecopera, FormatoFecha)
        cad = cad & "','" & Format(Rs!fecvalor, FormatoFecha) & "',"
        If IsNull(Rs!ImporteD) Then
            cad = cad & "NULL," & TransformaComasPuntos(CStr(Rs!ImporteH))
        Else
            cad = cad & TransformaComasPuntos(CStr(Rs!ImporteD)) & ",NULL"
        End If
        cad = cad & ",'" & DevNombreSQL(DBLet(Rs!Concepto)) & "','" & Rs!Numdocum & "',"
        cad = cad & TransformaComasPuntos(CStr(Rs!Saldo)) & ",0);"
        cad = SQL & cad
        'Ejecutamos SQL
        Conn.Execute cad
        Codigo = Codigo + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    'Ahora deberiamos eliminar el archivo
    If chkElimmFich.Value = 1 Then
        If Dir(Text12.Text, vbArchive) <> "" Then Kill Text12.Text
         MsgBox "Importación finalizada", vbInformation
    Else
        MsgBox "Proceso finalizado. El fichero NO será eliminado", vbExclamation
    End If
End Sub



