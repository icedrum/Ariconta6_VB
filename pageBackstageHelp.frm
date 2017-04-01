VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.CommandBars.v17.2.0.ocx"
Begin VB.Form pageBackstageHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Version 17.2.0"
   ClientHeight    =   10410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17115
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   4
      Left            =   6120
      ScaleHeight     =   7215
      ScaleWidth      =   10815
      TabIndex        =   45
      Top             =   1320
      Visible         =   0   'False
      Width           =   10815
      Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator111 
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   4695
         _Version        =   1114114
         _ExtentX        =   8281
         _ExtentY        =   450
         _StockProps     =   2
         MarkupText      =   ""
      End
      Begin MSComctlLib.ListView ListViewEmpresa 
         Height          =   6015
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resumido"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Inicio"
            Object.Width           =   317
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fin"
            Object.Width           =   317
         EndProperty
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label Label111 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cambiar empresa"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   240
         Width           =   2535
      End
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   1095
      Index           =   3
      Left            =   360
      TabIndex        =   0
      Top             =   7800
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":00AA
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   4800
      Width           =   5415
      _Version        =   1114114
      _ExtentX        =   9551
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":4114
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":41DC
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   5415
      _Version        =   1114114
      _ExtentX        =   9551
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":8246
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":8325
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   5415
      _Version        =   1114114
      _ExtentX        =   9551
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":C38F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":C448
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   1095
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   9000
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":104B2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   64
      Icon            =   "pageBackstageHelp.frx":10576
   End
   Begin XtremeCommandBars.BackstageSeparator BackstageSeparator1 
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   1320
      Width           =   10335
      _Version        =   1114114
      _ExtentX        =   18230
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator4 
      Height          =   10095
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   135
      _Version        =   1114114
      _ExtentX        =   238
      _ExtentY        =   17806
      _StockProps     =   2
      Vertical        =   -1  'True
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageSeparator lblBackstageSeparator2 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   7440
      Width           =   5175
      _Version        =   1114114
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   5
      Left            =   240
      TabIndex        =   29
      Top             =   720
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":145E0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   48
      Icon            =   "pageBackstageHelp.frx":146A5
   End
   Begin XtremeCommandBars.BackstageSeparator BackstageSeparator2 
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   360
      Width           =   5175
      _Version        =   1114114
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin XtremeCommandBars.BackstageButton btnAcciones 
      Height          =   975
      Index           =   6
      Left            =   240
      TabIndex        =   32
      Top             =   1920
      Width           =   5295
      _Version        =   1114114
      _ExtentX        =   9340
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   $"pageBackstageHelp.frx":16CC5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   4
      EnableMarkup    =   -1  'True
      ImageAlignment  =   0
      IconWidth       =   48
      Icon            =   "pageBackstageHelp.frx":16D86
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   1
      Left            =   6240
      ScaleHeight     =   6375
      ScaleWidth      =   10575
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   10575
      Begin MSComctlLib.ListView ListView4 
         Height          =   5145
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   9075
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuarios conectados"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   6120
         Width           =   4695
      End
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   0
      Left            =   6120
      ScaleHeight     =   6375
      ScaleWidth      =   10695
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   10695
      Begin MSComctlLib.ListView ListView3 
         Height          =   4905
         Left            =   240
         TabIndex        =   11
         Top             =   1065
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
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
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   6120
         Width           =   4695
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ejercicio Actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2985
         TabIndex        =   16
         Top             =   360
         Width           =   3705
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ejercicio Siguiente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6735
         TabIndex        =   15
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5685
         TabIndex        =   14
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9015
         TabIndex        =   13
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8145
         TabIndex        =   12
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Informacion BBDD"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   2
      Left            =   6240
      ScaleHeight     =   6375
      ScaleWidth      =   10575
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   10575
      Begin MSComctlLib.ListView ListView1 
         Height          =   4035
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7117
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Documentos de interes"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.PictureBox tabPage 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   3
      Left            =   6120
      ScaleHeight     =   6855
      ScaleWidth      =   10575
      TabIndex        =   33
      Top             =   1440
      Width           =   10575
      Begin XtremeCommandBars.BackstageButton btnProtectDocument 
         Height          =   1230
         Left            =   3120
         TabIndex        =   34
         Top             =   4800
         Width           =   1290
         _Version        =   1114114
         _ExtentX        =   2275
         _ExtentY        =   2170
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableMarkup    =   -1  'True
         TextImageRelation=   1
         ShowShadow      =   -1  'True
         IconWidth       =   32
         Icon            =   "pageBackstageHelp.frx":191F0
      End
      Begin XtremeCommandBars.BackstageButton btnManageVersions 
         Height          =   1230
         Left            =   8280
         TabIndex        =   35
         Top             =   4800
         Width           =   1290
         _Version        =   1114114
         _ExtentX        =   2275
         _ExtentY        =   2170
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableMarkup    =   -1  'True
         TextImageRelation=   1
         ShowShadow      =   -1  'True
         IconWidth       =   32
         Icon            =   "pageBackstageHelp.frx":1A25A
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Acerca de ..."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   204
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   4920
         Width           =   4695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tel: +34 963 805 579"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   42
         Top             =   3120
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "46007 Valencia"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   41
         Top             =   2640
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pasaje Ventura Feliu 13, Entresuelo 2 Izquierda"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   40
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Licencia usuario final"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   375
         Left            =   720
         TabIndex        =   39
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005B5B5B&
         Height          =   375
         Left            =   4800
         TabIndex        =   38
         Top             =   4320
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "There are no previous versions of this file"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   4800
         TabIndex        =   37
         Top             =   4800
         Width           =   4215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ariadna Software "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   720
         TabIndex        =   36
         Top             =   1320
         Width           =   3975
      End
   End
   Begin XtremeCommandBars.BackstageSeparator BackstageSeparator3 
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   3240
      Width           =   5175
      _Version        =   1114114
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   2
      MarkupText      =   ""
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   12840
      Picture         =   "pageBackstageHelp.frx":1B2C4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acerca de ..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   9.75
         Charset         =   204
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   6360
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leyendo datos"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Base de datos"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005B5B5B&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   7200
      Width           =   2535
   End
End
Attribute VB_Name = "pageBackstageHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnProtectDocument_Click()
    frmppal.OpcionesMenuInformacion ID_Licencia_Usuario_Final_web
End Sub

Private Sub Form_Load()
    Label9.Caption = vEmpresa.nomempre
    Label6(0).Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lblBackstageSeparator4.Height = Me.ScaleHeight
    i = Me.Width - tabPage(4).Left - 240
    If i < 300 Then i = 300
    tabPage(4).Width = i
    
    
    i = i - ListViewEmpresa.Left - 120
    If i < 100 Then i = 100
    ListViewEmpresa.Width = i
    
    
    Me.Image1.Left = Me.Width - Image1.Width - 120
End Sub


Private Sub btnAcciones_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    tabPage(0).Visible = False
    tabPage(1).Visible = False
    tabPage(2).Visible = False
    tabPage(3).Visible = False
    tabPage(4).Visible = False
    
    Label9.Visible = False
    Select Case Index
       
        Case 2
            'CAlendario del contribuyente
            LanzaVisorMimeDocumento Me.hwnd, "http://www.agenciatributaria.es/AEAT.internet/Bibl_virtual/folletos/calendario_contribuyente.shtml"
            tabPage(3).Visible = True
            
        Case 1 ' documentos
            
            Label3.Visible = True
            Label3.Refresh
            Cargadocumentos
            ListView1.Refresh
            tabPage(2).Visible = True
            Label3.Visible = False
        Case 0 ' ayuda
            tabPage(3).Visible = True
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "/Ariconta-6.html"  ' "http://www.ariadnasw.com/clientes/"
            
        Case 10 ' arimailges.exe
       '
        Case 3 ' Informacion de la base de datos
            Label3.Visible = True
            Label3.Refresh
            If CargarInformacionBBDD Then
                CargaInformeBBDD
                Label9.Visible = True
                tabPage(0).Visible = True
            End If
            Label3.Visible = False
            
        Case 4 'Usuarios activos
            Label3.Visible = True
            Label3.Refresh
            CargaShowProcessList
            Label9.Visible = True
            tabPage(1).Visible = True
            Label3.Visible = False
 
        Case 5
            
             tabPage(3).Visible = True
 
 
        Case 6
            'Label3.Visible = True
            'Label3.Refresh
            BuscaEmpresas
            'Label9.Visible = True
            tabPage(4).Visible = True
            Label3.Visible = False
        
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Function CargarInformacionBBDD() As String
Dim Sql As String
Dim Sql2 As String
Dim CadValues As String
Dim NroRegistros As Long
Dim NroRegistrosSig As Long
Dim NroRegistrosTot As Long
Dim NroRegistrosTotSig As Long
Dim FecIniSig As Date
Dim FecFinSig As Date
Dim Porcen1 As Currency
Dim Porcen2 As Currency
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarInformacionBBDD
    
    CargarInformacionBBDD = False
    
    Sql = "delete from tmpinfbbdd where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    FecIniSig = DateAdd("yyyy", 1, vParam.fechaini)
    FecFinSig = DateAdd("yyyy", 1, vParam.fechafin)
    
    Sql2 = "insert into tmpinfbbdd (codusu,posicion,concepto,nactual,poractual,nsiguiente,porsiguiente) values "
    
    'asientos
    Sql = "select count(*) from hcabapu where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(Sql)
    Sql = "select count(*) from hcabapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(Sql)
    
    CadValues = "(" & vUsu.Codigo & ",1,'Asientos'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
    Conn.Execute Sql2 & CadValues
    
    'apuntes
    Sql = "select count(*) from hlinapu where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    NroRegistros = DevuelveValor(Sql)
    Sql = "select count(*) from hlinapu where fechaent between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
    NroRegistrosSig = DevuelveValor(Sql)
    
    CadValues = "(" & vUsu.Codigo & ",2,'Apuntes'," & DBSet(NroRegistros, "N") & ",0," & DBSet(NroRegistrosSig, "N") & ",0)"
    Conn.Execute Sql2 & CadValues
    
    
    If vEmpresa.TieneContabilidad Then
            'facturas de venta
            Sql = "select count(*) from factcli where "
            Sql = Sql & " fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
            
            NroRegistrosTot = DevuelveValor(Sql)
            
            
            Sql = "select count(*) from factcli where "
            Sql = Sql & " fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
            
            NroRegistrosTotSig = DevuelveValor(Sql)
            
            i = 3
            
            Sql = "select * from contadores where not tiporegi in ('0','1')"
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs.EOF
            
                Sql = "select count(*) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
                Sql = Sql & " and fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
            
                NroRegistros = DevuelveValor(Sql)
                Porcen1 = 0
                If NroRegistrosTot <> 0 Then
                    Porcen1 = Round(NroRegistros * 100 / NroRegistrosTot, 2)
                End If
                
                Sql = "select count(*) from factcli where numserie = " & DBSet(Rs!tiporegi, "T")
                Sql = Sql & " and fecfactu between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
                
                NroRegistrosSig = DevuelveValor(Sql)
                Porcen2 = 0
                If NroRegistrosTotSig <> 0 Then
                    Porcen2 = Round(NroRegistrosSig * 100 / NroRegistrosTotSig, 2)
                End If
            
                CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(Rs!nomregis, "T") & "," & DBSet(NroRegistros, "N") & "," & DBSet(Porcen1, "N") & ","
                CadValues = CadValues & DBSet(NroRegistrosSig, "N") & "," & DBSet(Porcen2, "N") & ")"
                Conn.Execute Sql2 & CadValues
                
                i = i + 1
            
                Rs.MoveNext
            Wend
            
            Set Rs = Nothing
            
            'facturas de proveedor
            i = i + 1
            
            Sql = "select count(*) from factpro where fecharec between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
            NroRegistros = DevuelveValor(Sql)
            Sql = "select count(*) from factpro where fecharec between " & DBSet(FecIniSig, "F") & " and " & DBSet(FecFinSig, "F")
            NroRegistrosSig = DevuelveValor(Sql)
            
            CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & ",'Facturas Proveedores'," & DBSet(NroRegistros, "N") & ",0,"
            CadValues = CadValues & DBSet(NroRegistrosSig, "N") & ",0)"
            
            Conn.Execute Sql2 & CadValues
    
    
    End If
    
    
    CargarInformacionBBDD = True
    Exit Function


eCargarInformacionBBDD:
    MuestraError Err.Number, "Cargar Temporal de BBDD", Err.Description
End Function






Private Sub CargaInformeBBDD()
Dim IT As ListItem
Dim TotalArray  As Long
    On Error GoTo ECargaInformeBBDD
    
    Set Rs = New ADODB.Recordset
    
    Sql = "select * from tmpinfbbdd where codusu = " & vUsu.Codigo & " order by posicion "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView3.ListItems.Clear
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "CONCEPTO", 3500.0631
    ListView3.ColumnHeaders.Add , , "count ACTUAL", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen ACTUAL", 1000.2522, 1
    ListView3.ColumnHeaders.Add , , "count siguiente", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen siguiente", 1000.2522, 1
    
    
    
    
    TotalArray = 0
    While Not Rs.EOF
        Set IT = ListView3.ListItems.Add
        
        IT.Text = UCase(DBLet(Rs!Concepto, "T"))
        
        If DBLet(Rs!posicion, "N") > 2 Then
            IT.SubItems(1) = Format(DBLet(Rs!nactual, "N"), "###,###,###,##0")
            IT.SubItems(2) = Format(DBLet(Rs!Poractual, "N"), "##0.00") & "%"
            IT.SubItems(3) = Format(DBLet(Rs!nsiguiente, "N"), "###,###,###,##0")
            IT.SubItems(4) = Format(DBLet(Rs!Porsiguiente, "N"), "##0.00") & "%"
        Else
            IT.SubItems(1) = Format(DBLet(Rs!nactual, "N"), "###,###,###,##0")
            IT.SubItems(3) = Format(DBLet(Rs!nsiguiente, "N"), "###,###,###,##0")
        End If
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    
    Label5.Caption = Format(Now, "dd/mm/yyyy hh:nn:ss")
    Exit Sub
    
ECargaInformeBBDD:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Sub




Private Sub CargaShowProcessList()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim cad As String
Dim Equipo As String

    On Error GoTo ECargaShowProcessList
    
    Set Rs = New ADODB.Recordset
    
    ListView4.ListItems.Clear
    ListView4.ColumnHeaders.Clear
    
    ListView4.ColumnHeaders.Add , , "ID", 1500.0631
    ListView4.ColumnHeaders.Add , , "User", 2250.2522, 1
    ListView4.ColumnHeaders.Add , , "Host", 3000.2522, 1
    ListView4.ColumnHeaders.Add , , "Tiempo espera", 3050.2522, 1
    
    
    Set Rs = New ADODB.Recordset
    
    SERVER = Mid(Conn.ConnectionString, InStr(LCase(Conn.ConnectionString), "server=") + 7)
    SERVER = Mid(SERVER, 1, InStr(1, SERVER, ";"))
    
    EquipoConBD = (UCase(vUsu.PC) = UCase(SERVER)) Or (LCase(SERVER) = "localhost")
    
    cad = "show full processlist"
    Rs.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not Rs.EOF
        If Not IsNull(Rs.Fields(3)) Then
            If InStr(1, Rs.Fields(3), "ariconta") <> 0 Then
                If UCase(Rs.Fields(3)) = UCase(vUsu.CadenaConexion) Then
                    Equipo = Rs.Fields(2)
                    'Primero quitamos los dos puntos del puerto
                    NumRegElim = InStr(1, Equipo, ":")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    'El punto del dominio
                    NumRegElim = InStr(1, Equipo, ".")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    Equipo = UCase(Equipo)
                    
                    
                    Set IT = ListView4.ListItems.Add
                    
                    IT.Text = Rs.Fields(0)
                    IT.SubItems(1) = Rs.Fields(1)
                    IT.SubItems(2) = Equipo
                    
                    'tiempo de espera
                    Dim FechaAnt As Date
                    FechaAnt = DateAdd("s", Rs.Fields(5), Now)
                    IT.SubItems(3) = Format((Now - FechaAnt), "hh:mm:ss")
                End If
            End If
        End If
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    'Label6.Caption = Format(Now, "dd/mm/yyyy hh:nn:ss")
    
    
    Exit Sub
    
ECargaShowProcessList:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Sub


Private Sub Cargadocumentos()
Dim RN As ADODB.Recordset
Dim cad As String
Dim IT As ListItem

    Set Me.ListView1.SmallIcons = frmppal.ImageList1 'frmPpal.ImageListDocumentos
    On Error GoTo eCargadocumentos
    cad = "select iddocumento,nombrearchi from usuarios.wfichdocs WHERE aplicacion='ariconta' order by orden "
    Set RN = New ADODB.Recordset
    RN.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
        ListView1.ListItems.Add , "D" & Format(RN!iddocumento, "00000"), RN!nombrearchi, , 7   '1:PDF
       
        RN.MoveNext
    Wend
    RN.Close
    
eCargadocumentos:
    Err.Clear
    If ListView1.ListItems.Count = 0 Then
        Label8.Caption = "No hay documentacion disponible"
        Label8.Visible = True
    Else
        Label8.Visible = False
    End If
End Sub

Private Sub ListView1_DblClick()
Dim Abrir As Boolean

    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Abrir = False 'antes \ImgFicFT
    If Dir(App.Path & "\temp\" & ListView1.SelectedItem & ".pdf", vbArchive) = "" Then
        adodc1.ConnectionString = Conn
        adodc1.RecordSource = "Select * from usuarios.wfichdocs where idDocumento=" & Mid(ListView1.SelectedItem.Key, 2)
        adodc1.Refresh

        If LeerBinary(adodc1.Recordset!Campo, App.Path & "\temp\" & ListView1.SelectedItem.Text & ".pdf") Then Abrir = True
    Else
        Abrir = True
        
    End If
    
    If Abrir Then LanzaVisorMimeDocumento Me.hwnd, App.Path & "\temp\" & ListView1.SelectedItem & ".pdf"
        
End Sub




Private Sub BuscaEmpresas()
Dim Prohibidas As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Sql As String
Dim ItmX As ListItem

ListView1.ListItems.Clear

Dim i As Integer


    If Me.ListViewEmpresa.Tag = "1" Then Exit Sub
    
    
    'Ajustamos las columnas "Inamovibles"
    ListViewEmpresa.ColumnHeaders(1).Width = 800
    ListViewEmpresa.ColumnHeaders(4).Width = 1400
    ListViewEmpresa.ColumnHeaders(5).Width = 1400
    
    i = ListViewEmpresa.Width - 3840
    If i < 0 Then i = 180
    i = CInt(i / 4)
    ListViewEmpresa.ColumnHeaders(2).Width = i * 3
    ListViewEmpresa.ColumnHeaders(3).Width = i
        
 
    
    
    
    
    'Cargamos las prohibidas
    Prohibidas = DevuelveProhibidas
    
    'Cargamos las empresas
    
    Set Rs = New ADODB.Recordset
    
    '[Monica]11/04/2014: solo debe de salir las ariconta
    Rs.Open "Select * from usuarios.empresasariconta where conta like 'ariconta%' ORDER BY Codempre", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF
        cad = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, cad) = 0 Then
            cad = Rs!nomempre
            Set ItmX = ListViewEmpresa.ListItems.Add()
            ItmX.Text = Rs!codempre
            
            
            ItmX.SubItems(1) = Rs!nomempre
            ItmX.SubItems(2) = Rs!nomresum
            cad = "fechafin"
            Sql = DevuelveDesdeBD("fechaini", "ariconta" & Rs!codempre & ".parametros", "1", "1", "N", cad)
            ItmX.SubItems(3) = Sql
            ItmX.SubItems(4) = cad
            
                
            cad = Rs!CONTA & "|" & Rs!nomresum '& "|" & Rs!Usuario & "|" & Rs!Pass & "|"
            
            If Rs!codempre = vEmpresa.codempre Then
                ItmX.Bold = True
                Set ListView1.SelectedItem = ItmX
            End If
                
           ' ItmX.Tag = Cad
           ' ItmX.ToolTipText = Rs!CONTA
            
            
            'Si el codconta > 100 son empresas que viene del cambio del plan contable.
            'Atenuare su visibilidad
            If Rs!codempre > 100 Then
                ItmX.ForeColor = &H808080
                ItmX.ListSubItems(1).ForeColor = &H808080
                ItmX.ListSubItems(2).ForeColor = &H808080
                ItmX.ListSubItems(3).ForeColor = &H808080
                'ItmX.SmallIcon = 2
            Else
                'normal
                'ItmX.SmallIcon = 1
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    ListViewEmpresa.Tag = "1"
End Sub


Private Function DevuelveProhibidas() As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim i As Integer
    On Error GoTo EDevuelveProhibidas
    DevuelveProhibidas = ""
    Set Rs = New ADODB.Recordset
    i = vUsu.Codigo Mod 1000
    Rs.Open "Select * from usuarios.usuarioempresasariconta WHERE codusu =" & i, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    cad = ""
    While Not Rs.EOF
        cad = cad & Rs.Fields(1) & "|"
        Rs.MoveNext
    Wend
    If cad <> "" Then cad = "|" & cad
    Rs.Close
    DevuelveProhibidas = cad
EDevuelveProhibidas:
    Err.Clear
    Set Rs = Nothing
End Function


Private Sub ListViewEmpresa_DblClick()
    ' If Not ItemCheck Then Exit Sub
    If ListViewEmpresa.SelectedItem Is Nothing Then Exit Sub
   frmppal.CambiarEmpresa CInt(ListViewEmpresa.SelectedItem.Text)
End Sub
