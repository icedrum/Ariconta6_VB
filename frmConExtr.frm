VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConExtr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta extractos"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17445
   Icon            =   "frmConExtr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   17445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   525
      Left            =   3210
      TabIndex        =   35
      Top             =   240
      Width           =   13275
      Begin VB.CheckBox chkCtaSaldo 
         Caption         =   "Sólo ctas. con saldo"
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
         Left            =   5400
         TabIndex        =   39
         Top             =   120
         Width           =   2505
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "Resaltar total"
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
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   2025
      End
      Begin VB.CheckBox chkCtaConMov 
         Caption         =   "Sólo Cuentas con movimiento en el ejercicio"
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
         Left            =   8160
         TabIndex        =   37
         Top             =   120
         Width           =   4995
      End
      Begin VB.CheckBox chkPunteo 
         Caption         =   "Sólo Apuntes sin puntear"
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
         Left            =   2280
         TabIndex        =   36
         Top             =   120
         Width           =   2865
      End
   End
   Begin VB.TextBox Text6 
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
      Index           =   11
      Left            =   12810
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text6"
      Top             =   8850
      Width           =   1435
   End
   Begin VB.TextBox Text6 
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
      Index           =   10
      Left            =   11385
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text6"
      Top             =   8850
      Width           =   1435
   End
   Begin VB.TextBox Text6 
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
      Left            =   9945
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text6"
      Top             =   8850
      Width           =   1435
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   240
      TabIndex        =   16
      Top             =   60
      Width           =   2775
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Otra Consulta"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Grid"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cta Anterior"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cta Siguiente"
               Object.Tag             =   "0"
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox Text6 
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
      Left            =   12780
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   8340
      Width           =   1435
   End
   Begin VB.TextBox Text6 
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
      Index           =   7
      Left            =   11340
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   8340
      Width           =   1435
   End
   Begin VB.TextBox Text6 
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
      Left            =   9930
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text6"
      Top             =   8340
      Width           =   1435
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   270
      TabIndex        =   12
      Top             =   840
      Width           =   17130
      Begin VB.CommandButton cmdVer 
         Height          =   375
         Left            =   14040
         Picture         =   "frmConExtr.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ver datos"
         Top             =   90
         Width           =   375
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
         Left            =   12540
         TabIndex        =   2
         Text            =   "0000000000"
         Top             =   90
         Width           =   1305
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
         Left            =   9360
         TabIndex        =   1
         Text            =   "0000000000"
         Top             =   90
         Width           =   1275
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
         Index           =   2
         Left            =   1320
         TabIndex        =   0
         Text            =   "0000000000"
         Top             =   90
         Width           =   1305
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
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   90
         Width           =   4725
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Top             =   120
         Width           =   855
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Left            =   1020
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   12270
         Picture         =   "frmConExtr.frx":0A0E
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   9120
         Picture         =   "frmConExtr.frx":0A99
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   3
         Left            =   11220
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicio"
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
         Left            =   7830
         TabIndex        =   14
         Top             =   105
         Width           =   1305
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   16680
      TabIndex        =   18
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   6435
      Left            =   240
      TabIndex        =   19
      Top             =   1830
      Width           =   17040
      _ExtentX        =   30057
      _ExtentY        =   11351
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Asiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ampliacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contra."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "C.C."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Debe"
         Object.Width           =   2559
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Haber"
         Object.Width           =   2559
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Saldo"
         Object.Width           =   2559
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Punteada"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   " Saldo Período"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   33
      Top             =   8850
      Width           =   1575
   End
   Begin VB.Label LabelCab 
      Caption         =   "Contrapartida"
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
      Index           =   4
      Left            =   7530
      TabIndex        =   29
      Top             =   1440
      Width           =   1530
   End
   Begin VB.Label LabelCab 
      Alignment       =   1  'Right Justify
      Caption         =   "Debe"
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
      Index           =   6
      Left            =   11160
      TabIndex        =   28
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label LabelCab 
      Alignment       =   1  'Right Justify
      Caption         =   "Haber"
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
      Index           =   7
      Left            =   12780
      TabIndex        =   27
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label LabelCab 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Saldo "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   15090
      TabIndex        =   26
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label LabelCab 
      Caption         =   "C.C."
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
      Index           =   5
      Left            =   10350
      TabIndex        =   25
      Top             =   1440
      Width           =   705
   End
   Begin VB.Label LabelCab 
      Caption         =   "Documento"
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
      Index           =   2
      Left            =   2790
      TabIndex        =   24
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Label LabelCab 
      Caption         =   "Ampliación"
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
      Index           =   3
      Left            =   5010
      TabIndex        =   23
      Top             =   1440
      Width           =   2160
   End
   Begin VB.Label LabelCab 
      Alignment       =   1  'Right Justify
      Caption         =   "Asiento"
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
      Index           =   1
      Left            =   1560
      TabIndex        =   22
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label LabelCab 
      Caption         =   "Fecha"
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
      Index           =   0
      Left            =   330
      TabIndex        =   21
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Label LabelCab 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "P"
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
      Index           =   9
      Left            =   16230
      TabIndex        =   20
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label11 
      Caption         =   "Cargando datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   4000
      TabIndex        =   9
      Top             =   4000
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   " Saldo Actual"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   7410
      TabIndex        =   7
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label Label101 
      Caption         =   "1990 de 1000"
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
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   8340
      Width           =   3165
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   3480
      TabIndex        =   11
      Top             =   8340
      Width           =   3405
   End
   Begin VB.Label Label100 
      Caption         =   "Leyendo BD ........."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   8340
      Visible         =   0   'False
      Width           =   6645
   End
End
Attribute VB_Name = "frmConExtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 303

Public Cuenta As String   'Si es con cuenta
Public EjerciciosCerrados As Boolean


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1

Dim Sql As String
Dim vSql As String
Dim RC As String
Dim Mostrar As Boolean
Dim anc As Integer

Dim PrimeraVez As String

Dim DebePeriodo As Currency
Dim HaberPeriodo As Currency
Dim TotalPeriodo As Currency


Dim RT As Recordset
Private VieneDeIntroduccion As Boolean
Dim AnyoInicioEjercicio As String

Dim QuedanLineasDespuesModificar As Boolean



' ++ las he quitado de cargalistview
Dim ImpD As Currency
Dim ImpH As Currency



Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    Label10.Caption = DBLet(adodc1.Recordset!Nommacta, "T")
    If Err.Number <> 0 Then
        Err.Clear
        Label10.Caption = ""
         
    End If
    Label10.Refresh
End Sub

Private Sub RefrescarDatos2()
'CONSULTA EXTRACTOS
Dim F As Date
    If Text3(2).Text = "" Then
        MsgBox "Introduzca una cuenta", vbExclamation
        Text5.Text = ""
        PonerFoco Text3(2)
        Exit Sub
    End If
    If Text3(0).Text = "" Or Text3(1).Text = "" Then
        MsgBox "Introduce las fechas de consulta de extractos", vbExclamation
        Exit Sub
    End If
    
    If Text3(0).Text <> "" And Text3(1).Text <> "" Then
        If CDate(Text3(0).Text) > CDate(Text3(1).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Sub
        End If
    End If

    Sql = ""
    'Llegados aqui. Vemos la fecha y demas
    If Text3(0).Text <> "" Then
        Sql = " fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    End If
    
    If Text3(1).Text <> "" Then
        If Sql <> "" Then Sql = Sql & " AND "
        Sql = Sql & " fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    End If
    
    If Me.chkPunteo.Value = 1 Then
        Sql = Sql & " and punteada = 0"
    End If
    
    
    Text3(0).Tag = Sql  'Para las fechas
    

    'Para ver si la cuenta tiene movimientos o no
    vSql = "Select count(*) from hlinapu"
    If EjerciciosCerrados Then vSql = vSql & "1"
    vSql = vSql & " WHERE  fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    vSql = vSql & " AND fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    
    If Me.chkPunteo.Value = 1 Then
        vSql = vSql & " and punteada = 0"
    End If
    
    vSql = vSql & " AND codmacta ='"


    'Fijamos el año de incio de jercicio si es CERRADO
    F = CDate(Text3(0).Text)

    If Month(F) >= Month(vParam.fechaini) Then
        AnyoInicioEjercicio = Year(F)
    Else
        AnyoInicioEjercicio = Year(F) - 1
    End If


   
    Screen.MousePointer = vbHourglass
    If Not TieneMovimientos(Text3(2).Text) Then
    
        If Me.chkCtaConMov.Value = 1 Then
        
            MsgBox "La cuenta " & Text5.Text & " NO tiene movimientos en las fechas", vbExclamation
            
            
        End If
            
        Screen.MousePointer = vbDefault
        
        
        
        
        Cuenta = ""
        PonerModoUsuarioGnral 0, "ariconta"
        
        CargarDatos False
        PonerFoco Text3(2)
        Exit Sub
    End If
    
    Me.Refresh
    Screen.MousePointer = vbHourglass
    CargarDatos False

'    If Cuenta = "" Then PonerFoco DataGrid1
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Imprimir()
Dim MostrarAnterior As Byte

    frmConExtrList.Cuenta = Text3(2).Text
    frmConExtrList.Descripcion = Text5.Text
    frmConExtrList.FecDesde = Text3(0).Text
    frmConExtrList.FecHasta = Text3(1).Text
    
    frmConExtrList.Show vbModal
    
End Sub

Private Sub OtraCuenta(Index As Integer)
Dim I As Integer
Dim B As Boolean

    If Cuenta <> "" Then Exit Sub

    'Obtener la cuenta
    Screen.MousePointer = vbHourglass
    B = ObtenerCuenta(Index = 0)
    Screen.MousePointer = vbDefault
    If Not B Then Exit Sub

    'Poner datos
    Screen.MousePointer = vbHourglass
    
    'Ponemos los text a blanco
        For I = 6 To 8
            Text6(I).Text = ""
        Next I
    Label100.visible = True
    Label101.Caption = ""
    Label10.Caption = ""
    Label11.visible = True
    Me.Refresh
    DoEvents
    Screen.MousePointer = vbHourglass
    CargarDatos False
    Label100.visible = False
    Screen.MousePointer = vbDefault
End Sub



Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub chkCtaConMov_Click()
    If chkCtaConMov.Value = 0 Then Me.chkCtaSaldo.Value = 0
End Sub

Private Sub chkCtaSaldo_Click()
    If chkCtaSaldo.Value = 1 Then
        If Me.chkCtaConMov.Value = 0 Then
            MsgBoxA "No puede marcar cuenta con saldo, y no marcar movimientos periodo", vbExclamation
            chkCtaSaldo.Value = 0
        End If
    End If
End Sub

Private Sub chkPunteo_Click()
   ' RefrescarDatos2
End Sub

Private Sub cmdVer_Click()
    RefrescarDatos2
End Sub

Private Sub cmdVer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Activate()
   If Text3(2).Text = "" Then PonFoco Text3(2)
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim J As Integer
Dim I As Integer


    Me.Icon = frmppal.Icon

    Limpiar Me
    
    If EjerciciosCerrados Then
        Sql = "-1"
    Else
        AnyoInicioEjercicio = ""
        Sql = "0"
    End If
    
    CargarColumnas
    
    
    'Situaremos los textos
    '--------------------------------
    If vParam.autocoste Then
        anc = 1370
    Else
        anc = 1500
    End If
    
    '?? he sumado a todos los left 3000 unidades
    If vParam.autocoste Then
        For I = 2 To 3
            J = I * 3
            Text6(0 + J).Left = 8815 + 270
            Text6(0 + J).Width = anc - 15
            Text6(1 + J).Left = Text6(0 + J).Left + anc + 15
            Text6(1 + J).Width = anc - 15
            Text6(2 + J).Left = Text6(1 + J).Left + anc + 15
            Text6(2 + J).Width = anc - 15 + 100
        Next I
    Else
        For I = 2 To 3
            J = I * 3
            Text6(0 + J).Left = 9380
            Text6(0 + J).Width = 1500
            Text6(1 + J).Left = 10890
            Text6(1 + J).Width = 1500
            Text6(2 + J).Left = 12400
            Text6(2 + J).Width = 1600
        Next I
    End If
    
   
    imgCuentas.Picture = frmppal.imgIcoForms.ListImages(1).Picture


    ' añadido por el tema del listview
    For I = 6 To 11
        Text6(I).Width = 1850
    Next I
    
    For I = 2 To 3
        Text6(I * 3).Left = ListView1.ColumnHeaders(7).Left + 300
        Text6((I * 3) + 1).Left = ListView1.ColumnHeaders(8).Left + 300
        Text6((I * 3) + 2).Left = ListView1.ColumnHeaders(9).Left + 300
    Next I

    If EjerciciosCerrados Then
        I = -1
    Else
        I = 0
    End If
    
    
    Text3(0).Text = Format(DateAdd("yyyy", I, vParam.fechaini), "dd/mm/yyyy")
    If Not vParam.FecEjerAct Then I = I + 1
    Text3(1).Text = Format(DateAdd("yyyy", I, vParam.fechafin), "dd/mm/yyyy")
    
    VieneDeIntroduccion = False
    If Cuenta <> "" Then
        VieneDeIntroduccion = True
        Text3(2).Text = Cuenta
        Sql = ""
        CuentaCorrectaUltimoNivel Cuenta, Sql
        Text5.Text = Sql
        RefrescarDatos2
    Else
        CargaGrid
        PonFoco Text3(2)
    End If
    
    
    With Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 18
        .Buttons(2).Image = 16
        .Buttons(3).Image = 30
        .Buttons(5).Image = 7
        .Buttons(6).Image = 8
        
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
    chkTotal.Value = 0
    chkTotal.visible = False
    If vParam.RemarcaTotalExtr Then
        chkTotal.visible = True
        chkTotal.Value = CheckValueLeer(Me.Name)
        
    End If
    
    
    chkCtaConMov.Value = CheckValueLeer(Name)
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If vParam.RemarcaTotalExtr Then CheckValueGuardar Me.Name, Me.chkTotal
    CheckValueGuardar Me.Name, Me.chkCtaConMov.Value
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CByte(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    Text3(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text5.Text = RecuperaValor(CadenaSeleccion, 2)
    
    'no refrescaba
    'If ValorAnterior <> Text3(2).Text Then RefrescarDatos
    ValorAnterior = Text3(2).Text
    
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    Image1(0).Tag = Index
    If Text3(Index).Text <> "" Then
        frmC.Fecha = CDate(Text3(Index).Text)
    Else
        frmC.Fecha = Now
    End If
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub imgCuentas_Click()
    Set frmCta = New frmColCtas
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub



Private Sub ListView1_DblClick()
Dim NumAsien As Long
Dim Sql2 As String

If ListView1.SelectedItem Is Nothing Then Exit Sub


If Not VieneDeIntroduccion Then
    If Trim(ListView1.SelectedItem.Text) <> "" Then
        Screen.MousePointer = vbHourglass
        AsientoConExtModificado = 0
        Sql = ListView1.SelectedItem.ToolTipText & "|" & ListView1.SelectedItem.Text & "|" & ListView1.SelectedItem.SubItems(1) & "|"
        Sql2 = Sql
        

        frmAsientosHco.ASIENTO = Sql
        frmAsientosHco.vLinapu = ListView1.SelectedItem.Tag
        frmAsientosHco.Show vbModal
        
        espera 0.1
        If AsientoConExtModificado = 1 Then
            Screen.MousePointer = vbHourglass
            
            CargarDatos False
            
            Screen.MousePointer = vbDefault
            
            SituarAsiento Sql2
        Else
        
        End If
    End If
Else
    MsgBox "Esta en la introduccion de apuntes.", vbExclamation

End If

End Sub

'++
Private Sub SituarAsiento(vCadena As String)
Dim Sql As String

    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).ToolTipText = RecuperaValor(vCadena, 1) And ListView1.ListItems(I).Text = RecuperaValor(vCadena, 2) And ListView1.ListItems(I).SubItems(1) = RecuperaValor(vCadena, 3) Then
            Set ListView1.SelectedItem = ListView1.ListItems(I)
            ListView1.SelectedItem.EnsureVisible
            Exit For
        End If
    Next I

End Sub


Private Sub Text3_GotFocus(Index As Integer)
    With Text3(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    ValorAnterior = Text3(Index).Text
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 2 Then
        If KeyCode < 41 Then TeclaPulsada KeyCode
        If KeyCode <> 27 Then PonerFoco Text3(Index)
    End If
End Sub

'++
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
            Case 1:  KEYFecha KeyAscii, 1
            Case 2:  KEYBusqueda KeyAscii, 0
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCuentas_Click
End Sub

'++



Private Sub Text3_LostFocus(Index As Integer)
    
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 0, 1 ' fechas
            If Not EsFechaOK(Text3(Index)) Then
                MsgBox "Fecha incorrecta: " & Text3(Index).Text, vbExclamation
                Text3(Index).Text = ""
                PonerFoco Text3(Index)
                Exit Sub
            Else
                'If ValorAnterior <> Text3(Index).Text Then RefrescarDatos
            End If
    
        Case 2 ' cuentas
            RC = Trim(Text3(2).Text)
            If RC = "" Then
                Text5.Text = ""
                Exit Sub
            End If
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                Text3(2).Text = RC
                Text5.Text = Sql
                
                'If ValorAnterior <> Text3(Index).Text Then RefrescarDatos
            Else
                
                If LCase(Mid(Sql, 1, 19)) <> "no existe la cuenta" Then RC = ""
                
                MsgBox Sql, vbExclamation
                Text3(2).Text = ""
                Text5.Text = ""
                ListView1.ListItems.Clear
                If RC <> "" Then
                    Text3(2).Text = RC
                    OtraCuenta 0
                
                End If
                
                If ListView1.ListItems.Count = 0 Then PonerFoco Text3(2)
                If Text3(2).Text = "" Then PonerModoUsuarioGnral 0, "ariconta"
                    
            End If
     End Select
End Sub


Private Function CargarDatos(DesdeModificarLinea As Boolean)
Dim N As Long
On Error GoTo ECargaDatos


    Label101.Caption = ""
    Me.ListView1.ListItems.Clear
    Me.ListView1.Refresh
    Label100.visible = True
    Label100.Refresh
    
    Sql = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo & " AND Cta = '" & Text3(2).Text & "'"
    Conn.Execute Sql
        
    Sql = "DELETE from tmpconext where codusu= " & vUsu.Codigo & " AND Cta = '" & Text3(2).Text & "'"
    Conn.Execute Sql
    
    CargaDatosConExt Text3(2).Text, Text3(0).Text, Text3(1).Text, Text3(0).Tag, Text5.Text
    
    If DesdeModificarLinea Then
        'Compruebo que haya ALGUN datos, si no explota
        Sql = "cta = '" & Text3(2).Text & "' AND codusu"
        N = DevuelveDesdeBD("count(*)", "tmpconext", Sql, vUsu.Codigo)
        If N = 0 Then
            QuedanLineasDespuesModificar = False
            Exit Function
        End If
    End If

    CargaImportes
    
    CargaListview
    
    
    Label100.visible = False
    Exit Function
ECargaDatos:
        MuestraError Err.Number, "Datos cuenta"
        Label100.visible = False
End Function





Private Sub CargaGrid()


    If Text3(0).Tag = "" Then
        Sql = ""
        If Text3(0).Text <> "" Then Sql = " AND fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
        If Text3(1).Text <> "" Then Sql = Sql & " AND fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
        If Me.chkPunteo.Value = 1 Then Sql = Sql & " and punteada = 0"
        If Sql <> "" Then Text3(0).Tag = Mid(Sql, 5)
    End If
    
    



    adodc1.ConnectionString = Conn
    Sql = " codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, contra, ampconce, timporteD, timporteH, saldo,ccost, Punteada"
    If Text3(2).Text <> "" Then
        Sql = Sql & ",nommacta"
        Sql = "Select " & Sql & " from tmpconext left join cuentas on tmpconext.contra=cuentas.codmacta  WHERE codusu = " & vUsu.Codigo
    Else
        'Si esta a "" pongo otro select para que no de error
        Sql = Sql & ",linliapu"
        Sql = "Select " & Sql & " from tmpconext where codusu = " & vUsu.Codigo
    End If
    Sql = Sql & " AND cta = '" & Text3(2).Text & "' ORDER BY POS"
    
    adodc1.RecordSource = Sql
    adodc1.Refresh
    
    
    
    Label101.Caption = "Total lineas:   "
    Label101.Caption = Label101.Caption & Me.adodc1.Recordset.RecordCount
    
End Sub

Private Sub CargaListview()
Dim F1 As Date
Dim F2 As Date
Dim IT As ListItem
Dim PrimeraLineaNormal As Boolean
Dim Pinta As Boolean

Dim NumAto As Long  'el numero de asiento por si viene de los asientos

Dim cad As String
Dim miRsAux As ADODB.Recordset

    Me.ListView1.ListItems.Clear
    F1 = CDate(Text3(0).Text)
    F2 = CDate(Text3(1).Text)
    
    If Me.Cuenta <> "" Then NumAto = Val(RecuperaValor(Cuenta, 3))
    
    Set miRsAux = New ADODB.Recordset
    
    If Text3(2).Text = "" Then Exit Sub
       
    
    
    cad = " numasien,fechaent,cta codmacta,nomdocum numdocum,ampconce,timporteD impdebe,timporteH imphaber,ccost codccost"
    cad = cad & ",if(punteada='',' ','*') punteada,nommacta,contra ctacontr,linliapu numlinea, numdiari "
    If Text3(2).Text <> "" Then
        cad = "Select " & cad & " from tmpconext left join cuentas on tmpconext.contra=cuentas.codmacta  WHERE codusu = " & vUsu.Codigo
        If Me.chkPunteo.Value = 1 Then
            cad = cad & " and punteada = 0"
        End If
    Else
        cad = "Select " & cad & " from tmpconext left join cuentas on tmpconext.contra=cuentas.codmacta  where codusu = " & vUsu.Codigo
        If Me.chkPunteo.Value = 1 Then
            cad = cad & " and punteada = 0"
        End If
    End If
    cad = cad & " AND cta = '" & Text3(2).Text & "' ORDER BY fechaent,numasien,linliapu  " ' ORDER BY POS"
    
     
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRsAux.EOF Then
    
    Else
        'Deberiamos hacer TRES selects
        
        
        'De momento
        PrimeraLineaNormal = True  'Antes de imprimir la primera linea normal
        
        
        While Not miRsAux.EOF
     
                
            If miRsAux!FechaEnt < F1 Then
                'No hace nada
                Pinta = False
                
                
            Else
                If miRsAux!FechaEnt > F2 Then
                    Pinta = False
                
                Else
                    Pinta = True
                    
                End If
                
                If PrimeraLineaNormal Then PintaPrimeraLineaSaldo PrimeraLineaNormal, ImpD, ImpH, IT

            End If
                
            ImpD = ImpD + DBLet(miRsAux!impdebe, "N")
            ImpH = ImpH + DBLet(miRsAux!imphaber, "N")
            
            If Pinta Then
            
                'Pintamos la linea
                Set IT = ListView1.ListItems.Add()
                IT.Text = Format(miRsAux!FechaEnt, "dd/mm/yyyy")
                IT.SubItems(1) = DBLet(miRsAux!NumAsien, "N")
                IT.SubItems(2) = DBLet(miRsAux!Numdocum, "T")
                
                
                IT.SubItems(3) = DBLet(miRsAux!Ampconce, "T") & " "
                IT.ListSubItems(3).ToolTipText = DBLet(miRsAux!Ampconce, "T")
                If IsNull(miRsAux!ctacontr) Then
                    IT.SubItems(4) = " "
                Else
                    IT.SubItems(4) = DBLet(miRsAux!ctacontr, "T") & " "
                    IT.ListSubItems(4).ToolTipText = DBLet(miRsAux!Nommacta, "T")
                End If
                IT.SubItems(5) = DBLet(miRsAux!CodCcost, "T") & " "
                
                If IsNull(miRsAux!impdebe) Then
                    IT.SubItems(6) = " "
                    IT.SubItems(7) = Format(DBLet(miRsAux!imphaber, "N"), FormatoImporte)
                Else
                    IT.SubItems(7) = " "
                    IT.SubItems(6) = Format(DBLet(miRsAux!impdebe, "N"), FormatoImporte)
                End If
                IT.SubItems(8) = Format(ImpD - ImpH, FormatoImporte)
                IT.SubItems(9) = miRsAux!punteada 'en el select lleva si " " o "Si"
                IT.Tag = miRsAux!NumLinea  'para poder abrir el apunte
                IT.ToolTipText = miRsAux!NumDiari ' para poder abrir el diario
                
                If miRsAux!NumAsien = NumAto Then
                    IT.EnsureVisible
                    ListView1.SelectedItem = IT
                End If
                
            End If


            miRsAux.MoveNext
        
        Wend
        
        'No ha tenido movivimentos del periodo
        If PrimeraLineaNormal Then PintaPrimeraLineaSaldo PrimeraLineaNormal, ImpD, ImpH, IT

        PintaUltimaLineaSaldo IT
        
    End If
        
    miRsAux.Close
        
    Dim Rs As ADODB.Recordset
    cad = "SELECT codmacta, sum(coalesce(timporteD,0)) impdebe,sum(coalesce(timporteH,0)) imphaber"
    cad = cad & " from hlinapu "
    cad = cad & " where hlinapu.codmacta=" & DBSet(Text3(2).Text, "T") & " AND fechaent>=" & DBSet(vParam.fechaini, "F") '& " and fechaent <= " & DBSet(F2, "F")  '2013-01-01'"
    cad = cad & " group by 1 "
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Me.Text6(6).Text = Format(DBLet(Rs!impdebe, "N"), FormatoImporte)
        Me.Text6(7).Text = Format(DBLet(Rs!imphaber, "N"), FormatoImporte)
        Me.Text6(8).Text = Format(DBLet(Rs!impdebe, "N") - DBLet(Rs!imphaber, "N"), FormatoImporte)
    Else
        Me.Text6(6).Text = ""
        Me.Text6(7).Text = ""
        Me.Text6(8).Text = ""
    End If
    
    
    
    
End Sub


Private Sub PintaPrimeraLineaSaldo(ByRef LinaeaSaldoAnteriorPintada As Boolean, ByRef d As Currency, ByRef H As Currency, ByRef IT As ListItem)
Dim I As Integer
    
    LinaeaSaldoAnteriorPintada = False
    If d = 0 And H = 0 Then Exit Sub

    
    Set IT = ListView1.ListItems.Add(, "ANTERIOR")
    IT.Text = " "
    
    For I = 1 To 9
        IT.SubItems(I) = " "
    Next
    IT.SubItems(3) = "SALDO ANTERIOR AL PERIODO"
    IT.ListSubItems(3).ForeColor = vbBlack
    IT.ListSubItems(3).Bold = True
    If d <> 0 Then IT.SubItems(6) = Format(d, FormatoImporte)
    If H <> 0 Then IT.SubItems(7) = Format(H, FormatoImporte)
    IT.SubItems(8) = Format(d - H, FormatoImporte)
    ListView1.Refresh
    Screen.MousePointer = vbHourglass
End Sub



Private Sub PintaUltimaLineaSaldo(ByRef IT As ListItem)
Dim I As Integer
    
    
    If DebePeriodo = 0 And HaberPeriodo = 0 Then Exit Sub

    
    Set IT = ListView1.ListItems.Add(, "TOTAL")
    IT.Text = " "
    
    For I = 1 To 9
        IT.SubItems(I) = " "
    Next
    IT.SubItems(3) = "TOTAL"
    IT.ListSubItems(3).Bold = True
    If DebePeriodo <> 0 Then IT.SubItems(6) = Format(DebePeriodo, FormatoImporte)
    If HaberPeriodo <> 0 Then IT.SubItems(7) = Format(HaberPeriodo, FormatoImporte)
    IT.SubItems(8) = Format(TotalPeriodo, FormatoImporte)
    
        IT.SubItems(3) = " ***  TOTAL *** "
        
        
        If chkTotal.Value = 1 Then
            
            For I = 3 To 9
                IT.ListSubItems(I).Bold = True
            Next
            
            
            IT.EnsureVisible
            If TotalPeriodo < 0 Then
                IT.ListSubItems(8).ForeColor = vbRed
            Else
                If TotalPeriodo > 0 Then
                    IT.ListSubItems(8).ForeColor = vbBlue
                Else
                    IT.SubItems(8) = " "
                End If
            End If
        Else
            
            IT.ListSubItems(3).Bold = True
        
        End If
    
    
    
    
End Sub

Private Sub CargarColumnas()
Dim I As Integer
Dim cad As String

    
    cad = "1300|1150|2005|3714|1500|820|1950|1950|1950|350|"  '0|0|0|"
    'tieneanalitica
    Me.LabelCab(5).visible = (vParam.autocoste)
    
    
    For I = 1 To Me.ListView1.ColumnHeaders.Count
        ListView1.ColumnHeaders.Item(I).Width = RecuperaValor(cad, I)
        If I > 6 Then Me.LabelCab(I - 1).Width = ListView1.ColumnHeaders(I).Width

        Me.LabelCab(I - 1).Left = ListView1.ColumnHeaders.Item(I).Left + 120
    Next
    Me.LabelCab(9).Left = ListView1.ColumnHeaders.Item(10).Left + 300 '180
    
    Me.LabelCab(0).Left = ListView1.ColumnHeaders.Item(1).Left + 300
    Me.LabelCab(2).Left = ListView1.ColumnHeaders.Item(3).Left + 300
    Me.LabelCab(3).Left = ListView1.ColumnHeaders.Item(4).Left + 300
    Me.LabelCab(4).Left = ListView1.ColumnHeaders.Item(5).Left + 300
    Me.LabelCab(5).Left = ListView1.ColumnHeaders.Item(6).Left + 550
    
    
    Me.Width = ListView1.Left + ListView1.Width + 240
    
End Sub

Private Function ObtenerCuenta(Siguiente As Boolean) As Boolean
 
    Label101.Caption = ""
    Label100.visible = True
    Label100.Refresh
    
        
    ObtenerCuenta = ObtenerCuentaNormal(chkCtaConMov.Value = 1, Siguiente)
    
    Label100.visible = False
    
End Function




Private Function ObtenerCuentaNormal(ConMovimientos As Boolean, Siguiente As Boolean) As Boolean
    
    If ConMovimientos Then
           
           
           
           
        Sql = "select codmacta from hlinapu"
        If EjerciciosCerrados Then Sql = Sql & "1"
        Sql = Sql & " WHERE "
        Sql = Sql & " fechaent >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
        Sql = Sql & " AND fechaent <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
        
        If Me.chkCtaSaldo.Value = 0 Then
            If Me.chkCtaConMov.Value = 1 Then Sql = Sql & " and codconce < 900"
        End If
        
    Else
        Sql = "select codmacta from cuentas WHERE apudirec='S'  "
    End If
        
    Sql = Sql & " AND codmacta "
    If Siguiente Then
        Sql = Sql & ">"
    Else
        Sql = Sql & "<"
    End If
    Sql = Sql & " '" & Text3(2).Text & "'"
   
    If ConMovimientos Then
        Sql = Sql & " group by codmacta "
        If Me.chkCtaSaldo.Value = 1 Then Sql = Sql & " Having Sum(coalesce(timported, 0)) - Sum(coalesce(timporteh, 0)) <> 0"
   
    End If
   
    
    Sql = Sql & " ORDER BY codmacta"
    If Siguiente Then
        Sql = Sql & " ASC"
    Else
        Sql = Sql & " DESC"
    End If
    
    'Para optimizar la velocidad
    Sql = Sql & " limit 0, 1"

    
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        Sql = "No se ha obtenido la cuenta "
        If Siguiente Then
            Sql = Sql & "siguiente"
        Else
            Sql = Sql & "anterior"
        End If
        'Sql = Sql & " con movimientos en el periodo."
        MsgBox Sql, vbExclamation
        ObtenerCuentaNormal = False
    Else

    
        Text3(2).Text = RT!codmacta
        Text5.Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RT!codmacta, "T")
        ObtenerCuentaNormal = True
    End If
    RT.Close
    Set RT = Nothing
    
End Function



Private Sub CargaImportes()
Dim I As Integer
Dim Im1 As Currency
Dim Im2 As Currency


    Sql = "Select * from tmpconextcab where codusu=" & vUsu.Codigo & " and cta='" & Text3(2).Text & "'"
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        'Limpiaremos
        For I = 6 To 11
            Text6(I).Text = ""
        Next I
        ImpD = 0
        ImpH = 0
    Else
        Im1 = 0: Im2 = 0
        For I = 6 To 8
            Text6(I).Text = Format(RT.Fields(I + 4), FormatoImporte)
        Next I
        
        DebePeriodo = RT.Fields(7)
        HaberPeriodo = RT.Fields(8)
        TotalPeriodo = RT.Fields(9)
        
        'Importes calculado del periodo
        Im1 = RT.Fields(7) - RT.Fields(4)
        Im2 = RT.Fields(8) - RT.Fields(5)
        Text6(9).Text = Format(Im1, FormatoImporte)
        Text6(10).Text = Format(Im2, FormatoImporte)
        Im1 = Im1 - Im2
        Text6(11).Text = Format(Im1, FormatoImporte)
        
        
        ' ++
        ImpD = RT.Fields(4)
        ImpH = RT.Fields(5)
        
    End If
    RT.Close
End Sub

Private Sub TeclaPulsada(Codigo As Integer)
    Select Case Codigo
    Case 37 To 40
        If Codigo = 39 Or Codigo = 40 Then
            OtraCuenta 0
        Else
            OtraCuenta 1
        End If
    Case 13
        
    Case 27
        Unload Me
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(ButtonIndex As Integer)
Select Case ButtonIndex
Case 1
    Screen.MousePointer = vbHourglass
    
    Me.Frame2.Enabled = True
    
    Text3(2).SetFocus
    
    'Pongo a "" para cargar el grid a vacio
    Text3(2).Text = ""
    CargaGrid
    
    Screen.MousePointer = vbDefault
Case 2
    Imprimir
Case 3
    ListView1.GridLines = Not ListView1.GridLines
Case 5
    OtraCuenta 1
Case 6
    OtraCuenta 0
Case 9
    Unload Me
End Select
End Sub


Private Function TieneMovimientos(Cuenta As String) As Boolean
Dim RT As ADODB.Recordset
    
    Set RT = New ADODB.Recordset
    RT.Open vSql & Cuenta & "'", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TieneMovimientos = False
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then
            If Val(RT.Fields(0)) > 0 Then TieneMovimientos = True
        End If
    End If
    RT.Close
    Set RT = Nothing
End Function


Private Sub PonerFoco(ByRef T As Object)
On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
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
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) And Text3(2).Text <> "" And Cuenta = ""
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) And Text3(2).Text <> "" And Cuenta = ""
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


