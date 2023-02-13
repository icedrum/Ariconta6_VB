VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14700
   Icon            =   "frmCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameDescierre 
      BorderStyle     =   0  'None
      Caption         =   $"frmCierre.frx":000C
      Height          =   3975
      Left            =   30
      TabIndex        =   21
      Top             =   -30
      Width           =   5565
      Begin VB.TextBox Text3 
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
         Left            =   180
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   2160
         Width           =   5085
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   5
         Left            =   4200
         TabIndex        =   23
         Top             =   3240
         Width           =   1035
      End
      Begin VB.CommandButton cmdDescerrar 
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
         Left            =   3000
         TabIndex        =   22
         Top             =   3240
         Width           =   1125
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   2
         Left            =   4890
         TabIndex        =   95
         Top             =   90
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
      Begin VB.Label Label19 
         Caption         =   "descrip"
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
         Left            =   240
         TabIndex        =   102
         Top             =   2640
         Width           =   5175
      End
      Begin VB.Label Label19 
         Caption         =   "-Haber apuntes ejercicio siguiente"
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
         Left            =   600
         TabIndex        =   101
         Top             =   1440
         Width           =   3510
      End
      Begin VB.Label Label19 
         Caption         =   "-Estar descuadrada"
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
         Left            =   600
         TabIndex        =   100
         Top             =   1080
         Width           =   3510
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Para deshacer cierre no debe :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   96
         Top             =   240
         Width           =   3765
      End
      Begin VB.Label Label19 
         Caption         =   "-Trabajar nadie en esta contabilidad"
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
         Left            =   600
         TabIndex        =   24
         Top             =   720
         Width           =   3615
      End
   End
   Begin VB.Frame fPyG 
      BorderStyle     =   0  'None
      Height          =   9825
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   14055
      Begin VB.Frame FrameAmort 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   675
         Left            =   7080
         TabIndex        =   97
         Top             =   8220
         Width           =   4245
         Begin VB.TextBox Text1 
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
            Index           =   15
            Left            =   2880
            TabIndex        =   98
            Text            =   "Text1"
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha última amortización"
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
            Left            =   150
            TabIndex        =   99
            Top             =   240
            Width           =   2745
         End
      End
      Begin VB.CommandButton cmdSimula 
         Caption         =   "&Imprimir"
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
         Left            =   120
         TabIndex        =   93
         Top             =   9150
         Width           =   1335
      End
      Begin VB.Frame FrameTipoSalida 
         Caption         =   "Tipo de salida"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   90
         TabIndex        =   82
         Top             =   6270
         Width           =   6915
         Begin VB.OptionButton optTipoSal 
            Caption         =   "Impresora"
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
            TabIndex        =   92
            Top             =   720
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "Archivo csv"
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
            TabIndex        =   91
            Top             =   1200
            Width           =   1515
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "PDF"
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
            TabIndex        =   90
            Top             =   1680
            Width           =   975
         End
         Begin VB.OptionButton optTipoSal 
            Caption         =   "eMail"
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
            Left            =   240
            TabIndex        =   89
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtTipoSalida 
            BeginProperty Font 
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "Text1"
            Top             =   720
            Width           =   3345
         End
         Begin VB.TextBox txtTipoSalida 
            BeginProperty Font 
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   1200
            Width           =   4665
         End
         Begin VB.TextBox txtTipoSalida 
            BeginProperty Font 
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   1680
            Width           =   4665
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   6450
            TabIndex        =   85
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   6450
            TabIndex        =   84
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton PushButtonImpr 
            Caption         =   "Propiedades"
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
            Left            =   5190
            TabIndex        =   83
            Top             =   720
            Width           =   1515
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Apertura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   7080
         TabIndex        =   70
         Top             =   3810
         Width           =   6885
         Begin VB.TextBox txtDiario 
            BeginProperty Font 
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
            TabIndex        =   76
            Top             =   450
            Width           =   1005
         End
         Begin VB.TextBox txtDescDiario 
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
            Left            =   2370
            TabIndex        =   75
            Top             =   450
            Width           =   4275
         End
         Begin VB.TextBox Text1 
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
            Left            =   2370
            TabIndex        =   74
            Text            =   "Text1"
            Top             =   930
            Width           =   1275
         End
         Begin VB.TextBox Text1 
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
            Left            =   1200
            TabIndex        =   73
            Text            =   "Text1"
            Top             =   1950
            Width           =   1125
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
            Index           =   3
            Left            =   2370
            TabIndex        =   72
            Top             =   1950
            Width           =   4245
         End
         Begin VB.TextBox Text1 
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
            Left            =   2370
            TabIndex        =   71
            Text            =   "Text1"
            Top             =   1470
            Width           =   1275
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   930
            Picture         =   "frmCierre.frx":0097
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Index           =   12
            Left            =   180
            TabIndex        =   80
            Top             =   450
            Width           =   675
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha contabilización"
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
            Left            =   180
            TabIndex        =   79
            Top             =   990
            Width           =   2445
         End
         Begin VB.Label Label5 
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
            Height          =   225
            Index           =   10
            Left            =   180
            TabIndex        =   78
            Top             =   2010
            Width           =   945
         End
         Begin VB.Label Label5 
            Caption         =   "Número de asiento"
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
            TabIndex        =   77
            Top             =   1530
            Width           =   2445
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Regularización 8 y 9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   7080
         TabIndex        =   57
         Top             =   390
         Width           =   6885
         Begin VB.TextBox Text1 
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
            Index           =   14
            Left            =   1200
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   1440
            Width           =   5445
         End
         Begin VB.TextBox txtDiario 
            BeginProperty Font 
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
            TabIndex        =   63
            Top             =   450
            Width           =   1005
         End
         Begin VB.TextBox txtDescDiario 
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
            Left            =   2340
            TabIndex        =   62
            Top             =   450
            Width           =   4305
         End
         Begin VB.TextBox Text1 
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
            Index           =   13
            Left            =   2340
            TabIndex        =   61
            Text            =   "Text1"
            Top             =   960
            Width           =   1275
         End
         Begin VB.TextBox Text1 
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
            Index           =   12
            Left            =   150
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   2160
            Width           =   1125
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
            Index           =   4
            Left            =   1350
            TabIndex        =   59
            Top             =   2160
            Width           =   5295
         End
         Begin VB.TextBox Text1 
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
            Left            =   5760
            TabIndex        =   58
            Text            =   "Text1"
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label5 
            Caption         =   "Nº asiento"
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
            Left            =   4680
            TabIndex        =   69
            Top             =   990
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "ESTADO"
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
            Left            =   150
            TabIndex        =   68
            Top             =   1500
            Width           =   1545
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   3
            Left            =   870
            Picture         =   "frmCierre.frx":0A99
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Index           =   17
            Left            =   150
            TabIndex        =   67
            Top             =   420
            Width           =   585
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha contabilización"
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
            Left            =   150
            TabIndex        =   66
            Top             =   990
            Width           =   2415
         End
         Begin VB.Label Label5 
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
            Height          =   225
            Index           =   15
            Left            =   150
            TabIndex        =   65
            Top             =   1890
            Width           =   1545
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cierre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   90
         TabIndex        =   46
         Top             =   3810
         Width           =   6915
         Begin VB.TextBox txtDiario 
            BeginProperty Font 
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
            TabIndex        =   52
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox txtDescDiario 
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
            Left            =   2400
            TabIndex        =   51
            Top             =   360
            Width           =   4365
         End
         Begin VB.TextBox Text1 
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
            Left            =   2400
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   840
            Width           =   1275
         End
         Begin VB.TextBox Text1 
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
            Left            =   1230
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   1860
            Width           =   1125
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
            Index           =   2
            Left            =   2400
            TabIndex        =   48
            Top             =   1860
            Width           =   4365
         End
         Begin VB.TextBox Text1 
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
            Left            =   2400
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   1380
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "Número de asiento"
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
            Left            =   180
            TabIndex        =   56
            Top             =   1410
            Width           =   2235
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   870
            Picture         =   "frmCierre.frx":149B
            Top             =   330
            Width           =   240
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Index           =   9
            Left            =   180
            TabIndex        =   55
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha contabilización"
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
            Left            =   180
            TabIndex        =   54
            Top             =   900
            Width           =   2385
         End
         Begin VB.Label Label5 
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
            Index           =   6
            Left            =   180
            TabIndex        =   53
            Top             =   1920
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pérdidas y Ganancias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3405
         Left            =   90
         TabIndex        =   30
         Top             =   390
         Width           =   6915
         Begin VB.TextBox txtDiario 
            BeginProperty Font 
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
            Left            =   1320
            TabIndex        =   39
            Top             =   420
            Width           =   1005
         End
         Begin VB.TextBox txtDescDiario 
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
            Left            =   2430
            TabIndex        =   38
            Top             =   420
            Width           =   4335
         End
         Begin VB.TextBox Text1 
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
            Left            =   2430
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   930
            Width           =   1275
         End
         Begin VB.TextBox Text1 
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
            Left            =   180
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   2160
            Width           =   1350
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
            Index           =   0
            Left            =   1590
            TabIndex        =   35
            Top             =   2160
            Width           =   5175
         End
         Begin VB.TextBox Text1 
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
            Left            =   180
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   2880
            Width           =   1095
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
            Index           =   1
            Left            =   1380
            TabIndex        =   33
            Top             =   2880
            Width           =   5385
         End
         Begin VB.TextBox Text1 
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
            Left            =   2430
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   1440
            Width           =   885
         End
         Begin VB.TextBox Text1 
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
            Left            =   5490
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   930
            Width           =   1275
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   900
            Picture         =   "frmCierre.frx":1E9D
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   45
            Top             =   390
            Width           =   585
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha contabilización"
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
            Left            =   180
            TabIndex        =   44
            Top             =   960
            Width           =   2235
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta Perdidas y Ganancias"
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
            Index           =   2
            Left            =   180
            TabIndex        =   43
            Top             =   1860
            Width           =   3015
         End
         Begin VB.Label Label5 
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
            Height          =   225
            Index           =   3
            Left            =   180
            TabIndex        =   42
            Top             =   2610
            Width           =   1545
         End
         Begin VB.Label Label5 
            Caption         =   "Número de asiento"
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
            Left            =   180
            TabIndex        =   41
            Top             =   1440
            Width           =   2115
         End
         Begin VB.Label Label5 
            Caption         =   "Grupo excepción"
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
            Left            =   3810
            TabIndex        =   40
            Top             =   960
            Width           =   1755
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   7080
         TabIndex        =   27
         Top             =   6300
         Width           =   6885
         Begin VB.OptionButton Option2 
            Caption         =   "Cierre real"
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
            Left            =   4080
            TabIndex        =   29
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Simulación"
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
            Left            =   1020
            TabIndex        =   28
            Top             =   270
            Width           =   1425
         End
      End
      Begin VB.CommandButton cmdSimula 
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
         Height          =   375
         Index           =   1
         Left            =   11040
         TabIndex        =   2
         Top             =   9180
         Width           =   1365
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   4
         Left            =   12480
         TabIndex        =   3
         Top             =   9180
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   12480
         TabIndex        =   4
         Top             =   9180
         Width           =   1275
      End
      Begin VB.CommandButton cmdCierreEjercicio 
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
         Left            =   11040
         TabIndex        =   1
         Top             =   9180
         Width           =   1275
      End
      Begin MSComctlLib.ProgressBar pb3 
         Height          =   375
         Left            =   6000
         TabIndex        =   16
         Top             =   9210
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   1
         Left            =   13500
         TabIndex        =   94
         Top             =   30
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
      Begin VB.Line GELine 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   9330
         X2              =   11970
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label GElabel1 
         Caption         =   "REGULARIZACION 8 y 9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7140
         TabIndex        =   26
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "S  I  M  U  L  A  C  I  O  N"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   420
         Left            =   7110
         TabIndex        =   20
         Top             =   7620
         Width           =   6750
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Top             =   6720
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "CIERRE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   4200
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "APERTURA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7080
         TabIndex        =   18
         Top             =   4170
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1080
         X2              =   5850
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   8010
         X2              =   11760
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
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
         Left            =   1530
         TabIndex        =   17
         Top             =   9210
         Width           =   4275
      End
      Begin VB.Label Label7 
         Caption         =   "PÉRDIDAS Y GANANCIAS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   600
         Width           =   3015
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   2880
         X2              =   5850
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "C  I  E  R  R  E"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   420
         Left            =   7110
         TabIndex        =   13
         Top             =   7590
         Width           =   6750
      End
   End
   Begin VB.Frame fRenumeracion 
      BorderStyle     =   0  'None
      Height          =   4245
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   5745
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   285
         Left            =   270
         TabIndex        =   10
         Top             =   2700
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.CommandButton cmdRenumera 
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
         Left            =   2970
         TabIndex        =   9
         Top             =   3420
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   4350
         TabIndex        =   8
         Top             =   3420
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ejercicio SIGUIENTE"
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
         Left            =   3210
         TabIndex        =   7
         Top             =   1860
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ejercicio ACTUAL"
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
         Left            =   270
         TabIndex        =   6
         Top             =   1860
         Width           =   2085
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   0
         Left            =   5160
         TabIndex        =   81
         Top             =   120
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
      Begin VB.Label Label2 
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
         Left            =   330
         TabIndex        =   11
         Top             =   3120
         Width           =   5115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "No debe haber nadie más trabajando contra esta contabilidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   5
         Top             =   630
         Width           =   4875
      End
   End
End
Attribute VB_Name = "frmCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0.- Renumeracion
    '1.- Perdidas y ganancias Y Cierre
    '4.- Simulacion de cierre, es decir, mostrará un listado con
    '    los apuntes de pyg, cierre, y apertura
    '5.- DESCIERRRE
    
    
    
    
    '--------------------------------------------------------------
    'REGULARIZACION grupos 8 9: Concepto=961
        
    
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
    
Private IdPrograma As Long
    '1301- RENUMERACION DE ASIENTOS
    '1303- CIERRE DE EJERCICIO
    
Private PrimeraVez As Boolean
Dim Cad As String
Dim SQL As String
Dim Rs As Recordset

Dim i As Integer
Dim NumeroRegistros As Long
Dim MaxAsiento As Long
Dim ImporteTotal As Currency
Dim ImportePyG As Currency


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub cmdCierreEjercicio_Click()
    cmdCierreEjercicio.Enabled = False
    DoEvent2
    Screen.MousePointer = vbHourglass
    HazCierreEjercicio
    Label10.Caption = ""
    Screen.MousePointer = vbDefault
    cmdCierreEjercicio.Enabled = True
End Sub

Private Sub HazCierreEjercicio()
Dim Ok As Boolean

    Screen.MousePointer = vbHourglass
    Ok = True
    If Ok Then
        'QUITAR EL COMETNARIO
        Label10.Caption = "Comprobar Asientos Descuadrados"
        Label10.Refresh
        Ok = ComprobarAsientosDescuadrados
    End If
    
    If Not Ok Then Exit Sub

    If Ok Then
        Label10.Caption = "Comprobar Facturas sin Asientos"
        Label10.Refresh
        Ok = ComprobarFacturasSinAsientos
    End If
    If Not Ok Then Exit Sub
    
    ' nuevo, comprobamos que no exista ya el asiento nro 1 en el año siquiente, si lo hay dar aviso
    Label10.Caption = "Comprobar que no exista el asiento 1 en el año siguiente"
    Label10.Refresh
    Ok = ExisteAsientoUnoLibre
    If Not Ok Then
        MsgBox "Asiento Nº 1 reservado para la apertura, está ocupado. Renumere ejercicio siguiente o llame a Ariadna.", vbExclamation
        Exit Sub
    End If
    
'    Ok = True
    For i = 0 To 3
        If txtDiario(i).Text = "" Then
           If i < 2 Or (i = 3 And vParam.GranEmpresa And vParam.NuevoPlanContable) Then
                MsgBox "Seleccione el diario.", vbExclamation
                Ok = False
                Exit For
            End If
        End If
    Next i
    Label10.Caption = ""
    If Not Ok Then Exit Sub
    
    
    'Coamprobamos las cuentas 8 9
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        If Not ComprobarCierreCuentas8y9 Then Exit Sub
    End If
    
    
    
    Ok = UsuariosConectados("")
    If Not Ok Then
        SQL = "Seguro que desea cerrar el ejercicio?"
        If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
        
    Else
        'Hay usuarios conectados
        If vUsu.Nivel > 1 Then
            'NO TIENE PERMISOS
            Exit Sub
        Else
            SQL = "No es recomendado, pero, ¿desea continuar con el proceso?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
        End If
    End If
    
    
   'BLOQUEAMOS LA BD
   'If Not Bloquear_DesbloquearBD(True) Then
   '     MsgBox "No se ha podido bloquear a nivel de BD.", vbExclamation
   '     Exit Sub
   ' End If
    If UsuariosConectados("Hacer cierre" & vbCrLf, True) Then Exit Sub
    
    
    
    'Mensaje de lineas en introduccion de asientos
    Ok = True
    Screen.MousePointer = vbHourglass
    

'--
'    'Comprobar cuadre
    If Ok Then
        'QUITAR EL COMETNARIO
        Label10.Caption = "Comprobar cuadre"
        Label10.Refresh
        Ok = ComprobarCuadre
    End If
    cmdCierreEjercicio.Enabled = False
    Me.Refresh
    espera 0.3
    Me.Refresh
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Perdidas y Ganancias"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        Ok = ASientoPyG
    End If
    
    
    If Ok Then
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            pb3.Value = 0
            pb3.visible = True
            Label10.Caption = "Regularizacion 8 y 9"
            Label10.Refresh
            Label3.Caption = ""
            Label3.Refresh
            Ok = ASiento8y9
        End If
    End If
    
    
    Me.Refresh
    DoEvent2
    espera 0.2
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Cierre"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        'Hacer el cierre
        Ok = HacerElCierre
    End If
    

    If Ok Then
        'GRABO LOG
        
        vLog.Insertar 18, vUsu, "Cierre Ejercicio: " & DateAdd("d", -1, vParam.fechaini)
        
    End If
    
    Screen.MousePointer = vbHourglass
    Label10.Caption = ""
    Me.Refresh
    Screen.MousePointer = vbHourglass
    'Bloquear_DesbloquearBD False
    pb3.visible = False
    
    If Ok Then Unload Me
    Screen.MousePointer = vbDefault
End Sub



Private Function ComprobarCuadre() As Boolean
    Screen.MousePointer = vbHourglass
    ComprobarCuadre = True
    CadenaDesdeOtroForm = ""
    frmMensajes.Opcion = 5
    frmMensajes.Show vbModal
    If CadenaDesdeOtroForm <> "" Then
        ComprobarCuadre = False
        If vUsu.Codigo > 0 Then
            MsgBox "Error en la comprobación del cuadre.", vbExclamation
        Else
            If MsgBox("Error en cuadre. ¿Continuar igualmente pese al riesgo?", vbQuestion + vbYesNoCancel) = vbYes Then ComprobarCuadre = True
        End If
    End If
    Me.Refresh
End Function


Private Function ComprobarAsientosDescuadrados() As Boolean
Dim SQL As String
Dim SqlInsert As String
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset
Dim F As Date
    
    On Error GoTo eComprobarAsientosDescuadrados
    
    Screen.MousePointer = vbHourglass
    
    
    ComprobarAsientosDescuadrados = False
    
    SQL = "DELETE FROM tmphistoapu where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    SqlInsert = "insert into tmphistoapu (codusu, numdiari, numasien, fechaent, timported, timporteh) "
    
    'Febrero 2018
    ' En instalaciones grandes , proceso MUY lento
    ' ACeleracion:
    '   Cargaremos la tmp SI o SI
    '   Haremos en vez de el proceso de un anyo, 12 proesos de un mes
    F = vParam.fechaini
    
    For i = 1 To 12
    
        Label10.Caption = "Asientos periodo " & F & " "
        
    
    
        SQL = " select " & vUsu.Codigo & ", numdiari, numasien, fechaent, sum(coalesce(timported,0)), sum(coalesce(timporteh,0)) "
        SQL = SQL & " from hlinapu "
        'SQL = SQL & " where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
        SQL = SQL & " where fechaent between " & DBSet(F, "F") & " and "
        
        F = DateAdd("m", 1, F)
        F = DateAdd("d", -1, F)
        If F > vParam.fechafin Then
            F = vParam.fechafin
            i = 14 'se salga
        End If
        SQL = SQL & DBSet(vParam.fechafin, "F")
        SQL = SQL & " group by numdiari, numasien, fechaent "
        SQL = SQL & " having sum(coalesce(timported,0)) <> sum(coalesce(timporteh,0))  "
        SQL = SQL & " order by numdiari, numasien, fechaent "
        Label10.Caption = Label10.Caption & F
        Label10.Refresh
        
        Conn.Execute SqlInsert & SQL
        
        
        F = DateAdd("d", 1, F) 'Siguiente mes
    Next i
        
    'Borramos los que sean =0 D-H
    
    Label10.Caption = "Asientos Descuadrados D-H"
    Label10.Refresh
    SQL = "DELETE FROM tmphistoapu where codusu = " & vUsu.Codigo & " AND timported - timporteh=0"
    Conn.Execute SQL
    espera 0.2
    SQL = "SELECT codusu FROM tmphistoapu where codusu = " & vUsu.Codigo
        
    If TotalRegistrosConsulta(SQL) <> 0 Then
        
       
        
        frmMensajes.Opcion = 30
        frmMensajes.Show vbModal
        ComprobarAsientosDescuadrados = False
            
    Else
        ComprobarAsientosDescuadrados = True
    
    End If
    Label10.Caption = "Asientos descuadrados"
    Label10.Refresh
    Exit Function
    
eComprobarAsientosDescuadrados:
    MuestraError Err.Number, "Comprobar Asientos Descuadrados", Err.Description
End Function

Private Function ExisteAsientoUnoLibre() As Boolean
Dim SQL As String

    SQL = "select * from hcabapu where numasien = 1 and fechaent between " & DBSet(DateAdd("yyyy", 1, CDate(vParam.fechaini)), "F") & " and " & DBSet(DateAdd("yyyy", 1, CDate(vParam.fechafin)), "F")
    ExisteAsientoUnoLibre = (TotalRegistrosConsulta(SQL) = 0)

End Function



Private Function ComprobarFacturasSinAsientos() As Boolean
Dim SQL As String
Dim SqlInsert As String
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset

    On Error GoTo eComprobarFacturasSinAsientos


    Screen.MousePointer = vbHourglass
    ComprobarFacturasSinAsientos = False
    
    
    SqlInsert = "insert into tmpfaclin (codusu,numserie,nomserie,Numfac,Fecha, total) "
    SQL = " select " & vUsu.Codigo & ", numserie, contadores.nomregis, numfactu, fecfactu, totfaccl "
    SQL = SQL & " from factcli inner join contadores on factcli.numserie = contadores.tiporegi "
    SQL = SQL & " where fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    SQL = SQL & " and (numasien is null or numasien = 0)"
    SQL = SQL & " union "
    SQL = SQL & " select " & vUsu.Codigo & ", numserie, contadores.nomregis, numfactu, fecharec, totfacpr "
    SQL = SQL & " from factpro inner join contadores on factpro.numserie = contadores.tiporegi "
    SQL = SQL & " where fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    SQL = SQL & " and (numasien is null or numasien = 0)"
    SQL = SQL & " order by 1,2,3,4 "
    
    
    If TotalRegistrosConsulta(SQL) <> 0 Then
    
        Conn.Execute "delete from tmpfaclin where codusu = " & vUsu.Codigo
        
        Conn.Execute SqlInsert & SQL
        
        frmMensajes.Opcion = 31
        frmMensajes.Show vbModal
        ComprobarFacturasSinAsientos = False
            
    Else
        ComprobarFacturasSinAsientos = True
    
    End If
    Exit Function
    
eComprobarFacturasSinAsientos:
    MuestraError Err.Number, "Comprobar Facturas sin Asientos", Err.Description
End Function



Private Sub cmdDescerrar_Click()
Dim Ok As Boolean
On Error GoTo EDescierre

    Label10.Caption = ""

    SQL = "Seguro que desea deshacer el cierre?"
    If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados("Deshacer cierre.", True) Then Exit Sub
    
    
    If Not ExisteAsientosDescerrar Then
        MsgBox "No existe el asiento de apertura, luego no existe el cierre para el ejercicio anterior", vbExclamation
        Exit Sub
    End If
      
   
   
   
    'Veremos si existen los apuntes en hco .
    'Queremos ver si no se han llevado a ejercicios cerrados
    If Not ExistenApuntesEjercicioAnterior Then
        MsgBox "No se han encontrado apuntes del ejercicios anteriores", vbExclamation
        Exit Sub
    End If
   
   '---------------
    'Veremos si existen los apuntes de ejercicio siguiente
    If ExistenApuntesEjercicioSiguiente Then
        MsgBox "Hay apuntes del ejercicio siguiente ", vbExclamation
        Exit Sub
    End If
   
    'pASSSWORD MOMENTANEO
    Cad = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(Cad) <> "ARIADNA" Then
        If Cad <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    'Mensaje de lineas en introduccion de asientos
    Ok = True
    
   'BLOQUEAMOS LA BD
'   If Not Bloquear_DesbloquearBD(True) Then
'        MsgBox "No se ha podido bloquear a nivel de BD.", vbExclamation
'        Exit Sub
'    End If
    If UsuariosConectados("Deshacer cierre" & vbCrLf, True) Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    cmdDescerrar.Enabled = False
    cmdCancel(5).Enabled = False
    Me.Refresh
    
    Me.Refresh
    espera 0.3
    Me.Refresh
    If Ok Then
        Label19(3).Caption = "Eliminar asientos"
        Label19(3).Refresh
        Ok = HacerDescierre
    End If
    
    'Desbloqueamos BD
    'Bloquear_DesbloquearBD False
    
    If Ok Then
    
        
        vLog.Insertar 19, vUsu, "Cierre: " & DateAdd("d", 1, vParam.fechafin)
        
    
    
        Unload Me
    Else
        cmdDescerrar.Enabled = False
        cmdCancel(5).Enabled = True
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EDescierre:
    MuestraError Err.Number, "Descierre ejercicio:"
End Sub

Private Sub cmdRenumera_Click()
Dim Ok As Boolean
    'Comprobacion si hay alguien trabajando
    
    
    
    If UsuariosConectados("", (vUsu.Nivel = 0)) Then Exit Sub
    
    
    
    SQL = "Deberia hacer una copia de seguridad." & vbCrLf & vbCrLf
    SQL = SQL & "¿ Desea continuar igualmente ?" & vbCrLf
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    'BLOQUEAMOS LA BD
    'If Not Bloquear_DesbloquearBD(True) Then
    '    MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
    '    Exit Sub
    'End If
    If UsuariosConectados("Renumerando asientos" & vbCrLf, True) Then Exit Sub
    Ok = True
    
    
        
    If Ok Then
        'Hemos bloqueado la tbla y esta preparado para la renumeración. No hay nadie trabajando, ni lo va a haber
        Screen.MousePointer = vbHourglass
        'LOG
        If Option1(0).Value Then
            SQL = "actual "
        Else
            SQL = "siguiente "
        End If
        
        SQL = "Ejercicio " & SQL & Format(vParam.fechaini, "dd/mm/yyyy") & " - " & Format(vParam.fechafin, "dd/mm/yyyy")
        
        vLog.Insertar 17, vUsu, SQL
        SQL = ""
        
        
        pb1.visible = True
        Label2.Caption = ""
        'Ocultanmos el del fondo , para que no pegue pantallazos
        Me.Hide
        frmppal.Hide
        Me.Show
        
        
        'Renumeramos aqui dentro
        RenumerarAsientos
        
        'Volvemos a mostrar
        Me.Hide
        frmppal.Show
        
        
        
        
        
        pb1.visible = False
        Screen.MousePointer = vbDefault
    End If
    
    'Bloquear_DesbloquearBD False
    If Ok Then Unload Me
End Sub

Private Sub cmdSimula_Click(Index As Integer)
    Frame1.Enabled = False
    Screen.MousePointer = vbHourglass
    Simula Index
    Frame1.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Simula(Index As Integer)
Dim Ok As Boolean

    Ok = True
    If Ok Then
        'QUITAR EL COMETNARIO
        Label10.Caption = "Comprobar Asientos Descuadrados"
        Label10.Refresh
        Ok = ComprobarAsientosDescuadrados
    End If
    
    If Not Ok Then Exit Sub

    If Ok Then
        Label10.Caption = "Comprobar Facturas sin Asientos"
        Label10.Refresh
        Ok = ComprobarFacturasSinAsientos
    End If
    If Not Ok Then Exit Sub

    For i = 0 To 3
        If txtDiario(i).Text = "" Then
           If i < 2 Or (i = 3 And vParam.GranEmpresa) Then
                MsgBox "Seleccione el diario.", vbExclamation
                Ok = False
                Exit For
            End If
        End If
    Next i
    If Not Ok Then Exit Sub
    
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        If Not ComprobarCierreCuentas8y9 Then Exit Sub
    End If
    
    Ok = True
    ImporteTotal = 0
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Perdidas y Ganancias"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        Ok = ASientoPyG
    End If
    Screen.MousePointer = vbHourglass
    Me.Refresh
    espera 0.2
  
    
    
    If Ok Then
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            pb3.Value = 0
            pb3.visible = True
            Label10.Caption = "Regularizacion 8 y 9"
            Label10.Refresh
            Label3.Caption = ""
            Label3.Refresh
            Ok = ASiento8y9
        End If
    End If
        
    
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If Ok Then
        pb3.Value = 0
        pb3.visible = True
        Label10.Caption = "Cierre"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        'Hacer el cierre
        Ok = SimulaCierreApertura
    End If
    
    pb3.visible = False
    Label10.Caption = ""
    Label10.Refresh
    Label3.Caption = ""
    Label3.Refresh
    Me.Refresh
    
    If Ok Then
    
        'Exportacion a PDF
        If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
            If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
        End If
        
        InicializarVbles True
        
        If Not HayRegParaInforme("tmphistoapu", "tmphistoapu.codusu = " & vUsu.Codigo) Then Exit Sub
        
        If optTipoSal(1).Value Then
            'EXPORTAR A CSV
            AccionesCSV
        
        Else
            'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
        
            If optTipoSal(2).Value Or optTipoSal(3).Value Then
                ExportarPDF = True 'generaremos el pdf
            Else
                ExportarPDF = False
            End If
            SoloImprimir = False
            If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
            
            AccionesCrystal
        End If
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    SQL = "Select  numasien Asiento, fechaent Fecha, numdiari Diario, codmacta Cuenta, nommacta Descripción, timported Debe, timporteh Haber "
    SQL = SQL & " FROM tmphistoapu "
    
    If cadselect <> "" Then SQL = SQL & " WHERE codusu = " & vUsu.Codigo
    
    SQL = SQL & " ORDER BY 1,2,3,4"
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = IdPrograma & "-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "AsientosHco.rpt"

    'si se imprime el nif o la cuenta de cliente
    cadParam = cadParam & "pTitulo=""" & Me.Caption & """|"
    numParam = numParam + 1
    
    cadFormula = "{tmphistoapu.codusu} =" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 22
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub






Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False
    DoEvent2
    cmdCancel(Opcion).Cancel = True
    Select Case Opcion
    Case 1, 4
        PonerDatosPyG
        PonerDatosCierre
    
    End Select
    Screen.MousePointer = vbDefault
End If

End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer

Me.Icon = frmppal.Icon

Limpiar Me
PrimeraVez = True
Me.fRenumeracion.visible = False
Me.fPyG.visible = False
Me.frameDescierre.visible = False
Select Case Opcion
Case 0
    Me.fRenumeracion.visible = True
    H = fRenumeracion.Height
    W = fRenumeracion.Width
    Label2.Caption = ""
    pb1.visible = False
    Caption = "Renumeración de asientos"
    
    
    
    IdPrograma = 1301
    ' La Ayuda
    With Me.ToolbarAyuda(0)
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
Case 1, 4
    fPyG.visible = True
    H = fPyG.Height + 220 '+ 120
    W = fPyG.Width
    Label3.Caption = ""
    Label10.Caption = ""
    pb3.visible = False
    
    Me.Option2(0).Value = True
    Option2_Click 0
    Opcion = 4
'    PonerGrandesEmpresas
    
    IdPrograma = 1303
    ' La Ayuda
    With Me.ToolbarAyuda(1)
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
        
Case 5
    frameDescierre.visible = True
    H = frameDescierre.Height + 400
    W = frameDescierre.Width
    Caption = "Deshacer cierre"
    Label19(3).Caption = ""
    Text3.Text = "Ejercicio actual: " & Format(vParam.fechaini, "dd/mm/yyyy") & " - " & Format(vParam.fechafin, "dd/mm/yyyy")


    IdPrograma = 1304
    ' La Ayuda
    With Me.ToolbarAyuda(2)
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With

End Select


Me.Height = H + 100
Me.Width = W + 100
End Sub



Private Function CadenaFechasActuralSiguiente(Actual As Boolean) As String
Dim SQL As String
    If Actual Then
        'ACTUAL
        SQL = "fechaent >='" & Format(vParam.fechaini, FormatoFecha) & "' AND "
        SQL = SQL & "fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
    Else
        'SIGUIENTE
        Cad = Format(DateAdd("yyyy", 1, vParam.fechaini), FormatoFecha)
        SQL = "fechaent >='" & Cad & "' AND "
        Cad = Format(DateAdd("yyyy", 1, vParam.fechafin), FormatoFecha)
        SQL = SQL & "fechaent <='" & Cad & "'"
    End If
    CadenaFechasActuralSiguiente = SQL
End Function

Private Sub RenumerarAsientos()
Dim ContAsientos As Long
Dim NumeroAntiguo As Long
Dim Fec As String
Dim RA As Recordset


    Set RA = New ADODB.Recordset
    
   
    'obtner el maximo
    Cad = CadenaFechasActuralSiguiente(Option1(0).Value)
    SQL = "Select max(numasien) from hcabapu where " & Cad
    RA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    MaxAsiento = 0
    If Not RA.EOF Then
        MaxAsiento = DBLet(RA.Fields(0), "N")
    End If
    RA.Close


    'Obtener contador
    SQL = "Select count(numasien) from hcabapu where " & Cad
    RA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ContAsientos = 0
    If Not RA.EOF Then
        ContAsientos = DBLet(RA.Fields(0), "N")
    End If
    RA.Close

    
    If MaxAsiento + ContAsientos > 99999999 Then
        MsgBox "La aplicación no tiene espacio suficiente para renumerar. Numero registros posibles mayor que la capacidad disponible.", vbCritical
        Exit Sub
    End If

    
    '[Monica]19/01/2017
    Conn.BeginTrans
    
    
    'Para la progresbar
    NumeroRegistros = ContAsientos

    
    'Tendremos el incremeto
    MaxAsiento = MaxAsiento + ContAsientos + 1
    
    
    Label2.Caption = "Preparación"
    Me.Refresh
    
    'actualizamos todas las tablas sumandole maxasiento al numero de asiento donode proceda
    'es decir en el ejercicio y si tiene asiento
    If Not PreparacionAsientos(MaxAsiento) Then
        Conn.RollbackTrans
        Exit Sub
    End If
    

    
    DoEvent2
    Me.Refresh
    espera 0.01
    
    '-----------------------------------------------------------------
    ' Ahora iremos cogiendo cada registro y los iremos actualizando con
    ' los nuevos valores de numasien, tb para las tblas relacionadas
    ' Solo cambia NUMASIEN
    Cad = CadenaFechasActuralSiguiente(Option1(0).Value)
    SQL = "Select numasien,fechaent,numdiari from hcabapu where " & Cad
    SQL = SQL & " ORDER BY fechaent,numasien"
    RA.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ContAsientos = MaxAsiento
    'Y maxasiento lo utilizo como contador
    If Option1(1).Value Then
        MaxAsiento = 2
        NumeroRegistros = NumeroRegistros + 1
    Else
        MaxAsiento = 1
    End If



    While Not RA.EOF
    
        
        
    
          NumeroAntiguo = RA.Fields(0)
          Fec = "'" & Format(RA.Fields(1), FormatoFecha) & "'"
          If Not CambiaNumeroAsiento(NumeroAntiguo, MaxAsiento, Fec, RA!NumDiari) Then
               MsgBox "Error muy grave. Se ha producido un error en la renumeracion del asiento: " & vbCrLf & NumeroAntiguo & "  --> " & MaxAsiento & " // Fecha: " & Fec, vbExclamation
               Conn.RollbackTrans
               End
          End If
          
          
          'progressbar
          Label2.Caption = NumeroAntiguo & " / " & RA.Fields(1)
          Label2.Refresh
          i = Int((MaxAsiento / NumeroRegistros) * pb1.Max)
          pb1.Value = i
          DoEvent2
          
        
          'Siguiente
          MaxAsiento = MaxAsiento + 1
          ContAsientos = ContAsientos + 1
          RA.MoveNext
          
          If (MaxAsiento Mod 50) = 0 Then
               DoEvent2
               Me.Refresh
               espera 0.01
           End If
           
           
        
           
    Wend
    RA.Close
    Set RA = Nothing
    
    
    'En contadores ponemos el contador al numero k le toca
    MaxAsiento = NumeroRegistros
    SQL = "UPDATE contadores set "
    If (Option1(0).Value) Then
        SQL = SQL & " contado1=" & MaxAsiento
    Else
        If MaxAsiento = 1 Then MaxAsiento = 2
        SQL = SQL & " contado2=" & MaxAsiento
    End If
    SQL = SQL & " WHERE TipoRegi = '0'"
    Conn.Execute SQL

    Conn.CommitTrans
    Exit Sub

eRenumerarAsientos:
    MuestraError Err.Number, "Renumerar Asientos", Err.Description
    Conn.RollbackTrans
End Sub



Private Function CambiaNumeroAsiento(Antiguo As Long, Nuevo As Long, Fecha As String, NuDi As Integer) As Boolean

On Error GoTo ECambia
    CambiaNumeroAsiento = False
    
    'AUX--> Cudiado. No taocar " SET numasien"
    Cad = " SET numasien = " & Nuevo & " WHERE numasien = " & Antiguo
    Cad = Cad & " AND fechaent = " & Fecha & " AND numdiari = " & NuDi
    
   
    Conn.Execute "set foreign_key_checks = 0"
   
    'Actualizamos el registro de facturas
    SQL = "UPDATE factcli" & Cad
    Conn.Execute SQL
    
    
    'Actualizamos el registro de facturas
    SQL = " SET fecregcontable = fecregcontable , " & Mid(Cad, 5)
    SQL = "UPDATE factpro" & SQL
    Conn.Execute SQL
    
    'lineas
    SQL = "UPDATE hlinapu" & Cad
    Conn.Execute SQL
    
    'lineas de ficheros
    SQL = "UPDATE hcabapu_fichdocs" & Cad
    Conn.Execute SQL
    
    'cabeceras
    SQL = "UPDATE hcabapu" & Cad
    Conn.Execute SQL
    
    'hco de liquidaciones
    SQL = "UPDATE liqiva" & Cad
    Conn.Execute SQL
    
    Conn.Execute "set foreign_key_checks = 1"
    
    
    CambiaNumeroAsiento = True
    Exit Function
ECambia:
    MuestraError Err.Number, "Renumeracion tipo 1, asiento: " & Antiguo

End Function


Private Function PreparacionAsientos(Suma As Long) As Boolean

On Error GoTo EPreparacionAsientos


    PreparacionAsientos = False

    SQL = CadenaFechasActuralSiguiente(Option1(0).Value)
    Cad = " Set NumASien = NumASien + " & Suma & " @@@@"   'para frapro pondra fecregcontable
    Cad = Cad & " WHERE numasien>0 AND " & SQL
    
    pb1.Max = 6
    
    'Facturas clientes
    Label2.Caption = "Facturas clientes"
    Label2.Refresh
    pb1.Value = 1
    SQL = "UPDATE factcli" & Cad
    SQL = Replace(SQL, "@@@@", " ")
    Conn.Execute SQL
    
    'Facturas proveedores
    Label2.Caption = "Facturas proveedores"
    Label2.Refresh
    pb1.Value = 2
    
    SQL = "UPDATE factpro" & Cad
    SQL = Replace(SQL, "@@@@", " , fecregcontable = fecregcontable")
    
    
    Conn.Execute SQL
    
    Conn.Execute "set foreign_key_checks = 0"
    
    
    'Lineas hco asiento
    Label2.Caption = "Lineas asientos"
    Label2.Refresh
    pb1.Value = 3
    SQL = CadenaFechasActuralSiguiente(Option1(0).Value)
    SQL = "Select distinct(numasien) from hlinapu WHERE " & SQL
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = CadenaFechasActuralSiguiente(Option1(0).Value)
    Cad = " Set NumASien = NumASien + " & Suma
    Cad = Cad & " WHERE numasien>0 AND " & SQL
    
    'Ejecutaremos esto
    SQL = "UPDATE hlinapu " & Cad
    Conn.Execute SQL
    
    'ASientos
    Label2.Caption = "Cabeceras asientos"
    Label2.Refresh
    pb1.Value = 4
    SQL = "UPDATE hcabapu " & Cad
    Conn.Execute SQL
    
    'asientos ficheros
    Label2.Caption = "Líneas asientos ficheros"
    Label2.Refresh
    pb1.Value = 5
    SQL = "UPDATE hcabapu_fichdocs " & Cad
    Conn.Execute SQL
    
    'Liquidaciones de Iva
    Label2.Caption = "Liquidaciones IVA"
    Label2.Refresh
    pb1.Value = 6
    SQL = "UPDATE liqiva " & Cad
    Conn.Execute SQL
    
    Conn.Execute "set foreign_key_checks = 1"
    
    pb1.Max = 1000
    PreparacionAsientos = True
    Exit Function

EPreparacionAsientos:
    MsgBox "Error grave. Soporte técnico", vbExclamation
    Set Rs = Nothing
End Function



Private Sub PonerDatosPyG()

    'Fecha siempre la del final de ejercicio
    Text1(0).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    '8y9
    Text1(13).Text = Text1(0).Text


    NumeroRegistros = 1
    SQL = DevuelveDesdeBD("contado1", "contadores", "tiporegi", "0", "T")
    If SQL = "" Then
        MsgBox "Error obteniendo numero de asiento."
        cmdCierreEjercicio.Enabled = False
    Else
        Text1(3).Text = Val(SQL) + 1
    End If
    If vParam.GranEmpresa Then Text1(11).Text = Val(Text1(3).Text) + 1
    
    'PyG
    SQL = CuentaCorrectaUltimoNivel(vParam.ctaperga, Cad)
    If SQL = "" Then
        MsgBox "Error en la cuenta de pérdidas y ganancias de parametros.", vbExclamation
    Else
        Text1(1).Text = vParam.ctaperga
        Text2(0).Text = Cad
    End If
    
    'Concepto  --> Siempre sera nuestro 960
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "960")
    If SQL = "" Then
        MsgBox "No existe el concepto 960.", vbExclamation
    Else
        Text1(2).Text = "960"
        Text2(1).Text = SQL
    End If
    
    'Concepto para grandes empresas
    If vParam.GranEmpresa Then
        SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "961")
        If SQL = "" Then
            MsgBox "No existe el concepto 961.", vbExclamation
        Else
            Text1(12).Text = "961"
            Text2(4).Text = SQL
        End If
    End If

    
    'Simulacion nos salimos ya
    If Opcion = 4 Then Exit Sub
    
    
    'Si ya hay un 960 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    SQL = "Select numasien from hlinapu WHERE codconce=960 and fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        MsgBox "Ya se ha efectuado el asiento de Perdidas y ganancias : " & Rs.Fields(0), vbExclamation
        cmdCierreEjercicio.Enabled = False
    End If
    Rs.Close
    
    
    'Si ya hay un 961 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    If vParam.GranEmpresa Then
        SQL = "Select numasien from hlinapu WHERE codconce=961 and fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            MsgBox "Ya se ha efectuado el asiento de regularizacion : " & Rs.Fields(0), vbExclamation
            cmdCierreEjercicio.Enabled = False
        End If
        Rs.Close
    End If
    
    
    'Comprobamos k tampoc haya asiento 1 en ejercicio siguiente
    SQL = "Select numasien from hcabapu WHERE fechaent>'" & Format(vParam.fechafin, FormatoFecha)
    SQL = SQL & "' and numasien=1"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        MsgBox "Ya existe el asiento numero 1 para el año siguiente.", vbExclamation
        cmdCierreEjercicio.Enabled = False
    End If
    Rs.Close
    Set Rs = Nothing
End Sub



Private Sub PonerDatosCierre()
Dim Ok As Boolean

    'Fecha CIERRE siempre la del final de ejercicio
    Text1(7).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    Text1(9).Text = Format(DateAdd("d", 1, vParam.fechafin), "dd/mm/yyyy")
    
    'NUmero asiento
    'Es uno mas k Perdidas y ganancias
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        Text1(4).Text = Val(Text1(3).Text) + 2
    Else
        Text1(4).Text = Val(Text1(3).Text) + 1
    End If


    
    'Concepto  --> Siempre sera nuestro 980  CIERRE
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "980")
    If SQL = "" Then
        MsgBox "No existe el concepto 980.", vbExclamation
    Else
        Text1(5).Text = "980"
        Text2(2).Text = SQL
    End If
    
    'Nº asiento apertura.---- El uno
    Text1(6).Text = 1
    
    'Concepto  --> Siempre sera nuestro 970  apertura
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "970")
    If SQL = "" Then
        MsgBox "No existe el concepto 970.", vbExclamation
    Else
        Text1(8).Text = "970"
        Text2(3).Text = SQL
    End If
    
    
    
    'Si es simulacion busco el numero de diario mas pequeño
    If Opcion = 4 Then
        SQL = "Select numdiari,desdiari from tiposdiario order by numdiari"
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            For i = 0 To 3
                txtDescDiario(i).Text = Rs.Fields(1)
                txtDiario(i).Text = Rs.Fields(0)
            Next i
        End If
        Rs.Close
        Set Rs = Nothing
        'Ponemos el control en simula
        If Not PrimeraVez Then cmdSimula(1).SetFocus
    End If
    
    
    
    'Simulacion nos salimos ya
    If Opcion = 4 Then Exit Sub
    
    
    'Si ya hay un 980 y/o 970 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Cad = CadenaFechasActuralSiguiente(True)
    SQL = "Select numasien from hlinapu WHERE codconce=980"
    SQL = SQL & " AND " & Cad
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        MsgBox "Ya se ha efectuado el asiento de cierre de ejercicio : " & Rs.Fields(0), vbExclamation
        Ok = False
    Else
        Ok = True
    End If
    Rs.Close
    
    'Apertura
    If Ok Then
            Cad = CadenaFechasActuralSiguiente(False)
            SQL = "Select numasien from hlinapu WHERE codconce=980"
            SQL = SQL & " AND " & Cad
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                MsgBox "Ya se ha efectuado el asiento de cierre de ejercicio : " & Rs.Fields(0), vbExclamation
                Ok = False
            Else
                Ok = True
            End If
            Rs.Close
    End If
    Set Rs = Nothing
    cmdCierreEjercicio.Enabled = Ok
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    txtDiario(i).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescDiario(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    i = Index
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|1|"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub

Private Sub Option2_Click(Index As Integer)
Dim vFecha As String

    If Index = 0 Then
        Opcion = 4
    Else
        Opcion = 1
    End If

    cmdSimula(0).visible = (Opcion = 4)
    cmdSimula(1).visible = (Opcion = 4)
    Me.cmdCancel(4).visible = (Opcion = 4)
    Me.cmdCierreEjercicio.visible = (Opcion = 1)
    Me.cmdCancel(1).visible = (Opcion = 1)
    Label6.visible = (Opcion = 4)
    Label4.visible = Not Label6.visible
    If Opcion = 1 Then
        Me.Caption = "Cierre de ejercicio"  '"Generar asiento de pérdidas y ganancias y el cierre de ejercicio"
    Else
        Me.Caption = "Simulación cierre"
    End If

    If TieneInmovilizado Then
        Me.FrameAmort.visible = (Opcion = 1)
        
        vFecha = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", 1, "N")
        Text1(15).Text = Format(vFecha, "dd/mm/yyyy")
    End If
    

'    Caption = "Cierre de Ejercicio"

    PonerGrandesEmpresas
        
    If Opcion = 1 Then
        For i = 0 To txtDiario.Count - 1
            txtDiario(i).Text = ""
        Next i
        
        SQL = "select numdiari, desdiari from tiposdiario "
        If TotalRegistrosConsulta(SQL) = 1 Then
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                For i = 0 To txtDiario.Count - 1
                    txtDiario(i).Text = DBLet(Rs.Fields(0), "N")
                    txtDescDiario(i).Text = DBLet(Rs.Fields(1), "T")
                Next i
            End If
            Set Rs = Nothing
        End If
    End If

    
    Me.FrameTipoSalida.Enabled = (Opcion = 4)
    If Opcion = 4 Then
        PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
        ponerLabelBotonImpresion cmdSimula(1), cmdSimula(0), 0
    End If


'    ' lo del form activate
    PonerDatosPyG
    PonerDatosCierre

End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub ToolbarAyuda_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtDiario_GotFocus(Index As Integer)
    PonFoco txtDiario(Index)
End Sub

'++
Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYBusqueda KeyAscii, 0
            Case 1:  KEYBusqueda KeyAscii, 1
            Case 2:  KEYBusqueda KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub
'++

Private Sub txtDiario_LostFocus(Index As Integer)
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    Me.txtDescDiario(Index).Text = ""
    If txtDiario(Index).Text = "" Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtDiario(Index).Text) Then
        MsgBox "Diario debe ser numérico: " & txtDiario(Index).Text, vbExclamation
        txtDiario(Index).Text = ""
        txtDiario(Index).SetFocus
        Exit Sub
    End If
    SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text)
    If SQL = "" Then
        MsgBox "No existe el diario : " & txtDiario(Index).Text, vbExclamation
        txtDiario(Index).Text = ""
        txtDiario(Index).SetFocus
    Else
        Me.txtDescDiario(Index).Text = SQL
    End If
    
    
End Sub




Private Function ASientoPyG() As Boolean
Dim Ok As Boolean
Dim NoTieneLineas As Boolean
Dim Cuantos As Long
    
    On Error GoTo EASientoPyG
    
    ASientoPyG = False

    'Generamos los  apuntes, sobre hcabapu
    If Opcion <> 4 Then
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(0).Text
        Cad = SQL & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(3).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Asiento Pérdidas y Ganancias')"
        
    Else
        'Estamo simulando
        'Borramos los datos de tmp
        Cad = "Delete from tmphistoapu where codusu = " & vUsu.Codigo
    End If
        
    Conn.Execute Cad
    Ok = Cuentas6y7
    
    If Ok Then
        
        'Entonces hacemos las cuentas de otros grupos de analitica
        If Text1(10).Text <> "" Then
            Ok = Cuentas9
        End If
    End If
    
    If Ok Then
        'Cuadramos el asiento
        Ok = CuadrarAsiento
    End If
    
    
    If Opcion = 1 And Ok Then
        
        'Veremos si hay algun registro insertado
        
        SQL = "numdiari=" & txtDiario(0).Text & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien  "
        espera 0.2
        SQL = DevuelveDesdeBD("count(*)", "hlinapu", SQL, Text1(3).Text)
        If SQL = "" Then SQL = "0"
        NoTieneLineas = Val(SQL) = 0
        
        
        If NoTieneLineas Then
            'NO HA INSERTADO NI UNA SOLA LINEA EN  LINAPU. CON LO CUAL TENDREMOS QUE AVISAR Y BORRAR
            If MsgBox("Ningún apunte en lineas de perdidas y ganancias. ¿Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then
                Ok = False

            Else
                'COMO NO genera el apunte pq no hay lineas 6 y 7 entonces el contador de cierre disminuye
                If vParam.NuevoPlanContable And vParam.GranEmpresa Then
                    'El del 7 y 8 pasa al del CIERRE
                    Text1(4).Text = Text1(11).Text
                    'El contador del py g pasa al de 8 y 9
                    Text1(11).Text = Text1(3).Text
                    
                Else
                    Text1(4).Text = Text1(3).Text
                End If
            End If
            
            'De cualquier modo hay que borrar la cabecera que ha creado
            SQL = " where numdiari=" & txtDiario(0).Text
            SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
    
            'Borramos por si acaso ha insertado lineas
            Cad = "Delete FROM hcabapu" & SQL
            Conn.Execute Cad
        End If
    End If
    
    
    If Opcion = 4 Then
        ASientoPyG = Ok
        Exit Function
    End If
    
    If Not Ok Then
        'Comun lineas y cabeceras
        SQL = " where numdiari=" & txtDiario(0).Text
        SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
        
        'Borramos por si acaso ha insertado lineas
        Cad = "Delete FROM hlinapu" & SQL
        Conn.Execute Cad
    
        'Borramos la cabcecera del apunte
        Cad = "DELETE FROM hcabapu" & SQL
        Conn.Execute Cad
        
        Label3.Caption = ""
        Exit Function
    End If
    ASientoPyG = True
    
    Exit Function
EASientoPyG:
    MuestraError Err.Number, "Error en procedmiento ASientoPyG"
End Function


Private Function Cuentas6y7() As Boolean
On Error GoTo ECuentas6y7


    Cuentas6y7 = False
    Set Rs = New ADODB.Recordset
    
    'Para todas las cuentas de los grupos 6 y 7  ----> Vienen en parametros
    ' calculamos su saldo y si es distinto de 0 lo insertamos en hlinapu
    MaxAsiento = 1
    If vParam.GrupoGto <> "" Then
        If Not Subgrupo(vParam.GrupoGto, "") Then Exit Function
    End If
    
    If vParam.GrupoVta <> "" Then
        If Not Subgrupo(vParam.GrupoVta, "") Then Exit Function
    End If
        
        
        
    pb3.Value = 0
    pb3.Max = 1000
        
        
    Set Rs = Nothing
    Cuentas6y7 = True
    Exit Function
ECuentas6y7:
    MuestraError Err.Number, "Cuentas   ventas / gastos "
End Function



Private Function Cuentas9() As Boolean
On Error GoTo ECuentas9


    Cuentas9 = False
    Set Rs = New ADODB.Recordset
    
    'Para todas las cuentas de los grupos 6 y 7  ----> Vienen en parametros
    ' calculamos su saldo y si es distinto de 0 lo insertamos en hlinapu
    If vParam.GrupoOrd <> "" And Text1(10).Text <> "" Then
        If Not Subgrupo(vParam.GrupoOrd, Text1(10).Text) Then Exit Function
    End If
    

        
    Set Rs = Nothing
    Cuentas9 = True
    Exit Function
ECuentas9:
    MuestraError Err.Number, "Cuentas   ventas / gastos "
End Function


'Febrero 2018
'Hacemos otro procedimiento
' NO lo borro por si hay que tirar p'atras

''''''Private Function Subgrupo(Primera As String, Excepcion As String) As Boolean
''''''Dim CONT As Integer
''''''Dim Importe As Currency
''''''Dim AUX3 As String
''''''Dim vCta As String
''''''
''''''    Subgrupo = False
''''''
''''''
''''''    Label10.Caption = "Saldos " & Primera & "*"
''''''    Label10.Refresh
''''''    pb3.Value = 0
''''''    pb3.Max = 12
''''''
''''''    'ATENCION AÑOS PARTIDOS
''''''    cad = Mid("__________", 1, vEmpresa.DigitosUltimoNivel - 1)
''''''    AUX3 = Primera & cad
''''''    cad = " from hlinapu" ' antes hsaldos
''''''
''''''    'Necesito sbaer tb nombre cta
''''''    If Opcion = 4 Then cad = cad & ",cuentas"
''''''
''''''    cad = cad & " WHERE "
''''''    If Opcion = 4 Then cad = cad & " cuentas.codmacta = hlinapu.codmacta AND "
''''''    'Por la ambiguedad del nombre
''''''    vCta = " ("
''''''    If Opcion = 4 Then vCta = vCta & " cuentas."
''''''    vCta = vCta & "codmacta like '" & AUX3 & "')"
''''''    If Excepcion <> "" Then
''''''        Excepcion = Mid(Excepcion & "__________", 1, vEmpresa.DigitosUltimoNivel)
''''''        vCta = vCta & " and not ("
''''''        If Opcion = 4 Then vCta = vCta & " cuentas."
''''''        vCta = vCta & "codmacta like '" & Excepcion & "')"
''''''    End If
''''''
''''''
''''''    cad = cad & vCta
''''''
''''''
''''''    cad = cad & " and fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
''''''
''''''    If Opcion = 4 Then
''''''        cad = "cuentas.codmacta,nommacta " & cad & " GROUP BY cuentas.codmacta ORDER BY cuentas.codmacta"
''''''    Else
''''''        cad = " codmacta " & cad & " GROUP BY codmacta ORDER BY codmacta"
''''''    End If
''''''
''''''    '
''''''    SQL = "Select sum(coalesce(timported,0))-sum(coalesce(timporteh,0)), " & cad
''''''    Rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
''''''
''''''
''''''    Label3.Caption = ""
''''''    Label3.Refresh
''''''
''''''
''''''    NumeroRegistros = 0
''''''    If Rs.EOF Then
''''''        Rs.Close
''''''        'Puede k asi este bien
''''''        Subgrupo = True
''''''        Exit Function
''''''    End If
''''''
''''''    'Contador
''''''    While Not Rs.EOF
''''''        NumeroRegistros = NumeroRegistros + 1
''''''        Rs.MoveNext
''''''    Wend
''''''
''''''    'Preparamos el SQL para la insercion de lineas de apunte
''''''    'Montamos la cadena casi al completo
''''''    CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
''''''
''''''    Rs.MoveFirst
''''''    CONT = 1
''''''    AUX3 = "'" & Text1(1).Text & "'"
''''''    While Not Rs.EOF
''''''
''''''        Label3.Caption = Rs.Fields(1)
''''''        Label3.Refresh
''''''        i = Int((CONT / NumeroRegistros) * pb1.Max)
''''''        pb1.Value = i
''''''        Importe = Rs.Fields(0)
''''''        If Importe <> 0 Then
''''''            If Opcion = 4 Then
''''''                cad = SQL & "," & MaxAsiento + CONT & ",'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs.Fields(2)) & "','1','" & Text2(1).Text & "',"
''''''            Else
''''''                cad = SQL & "," & MaxAsiento + CONT & ",'" & Rs.Fields(1) & "','',960,'" & Text2(1).Text & "',"
''''''            End If
''''''            InsertarLineasDeAsientos Importe, AUX3
''''''        End If
''''''        'Sig
''''''        CONT = CONT + 1
''''''        Rs.MoveNext
''''''    Wend
''''''    Rs.Close
''''''    MaxAsiento = MaxAsiento + CONT - 1
''''''    Subgrupo = True
''''''End Function

Private Function Subgrupo(Primera As String, Excepcion As String) As Boolean
Dim CONT As Integer
Dim Importe As Currency
Dim AUX3 As String
Dim vCta As String
Dim F As Date

    Subgrupo = False
    
    Label10.Caption = "Saldos " & Primera & "*"
    Label10.Refresh
    pb3.Value = 0
    pb3.Max = 12
    
    
    
    
    'Febrero 2018
    ' En instalaciones grandes , proceso MUY lento
    ' ACeleracion:
    '   Cargaremos la tmp SI o SI
    '   Haremos en vez de el proceso de un anyo, 12 proesos de un mes
    F = vParam.fechaini
    
    Conn.Execute "DELETE FROM tmpevolsal WHERE codusu = " & vUsu.Codigo
    pb3.Value = 0
    pb3.Max = 12
    For i = 1 To 12
    
        Label10.Caption = Primera & "*     mes: " & Format(F, "mmmm")
        Label10.Refresh
        
        
    
    
        AUX3 = Primera & "%"
        Cad = " from hlinapu" ' antes hsaldos
        
        
        
        Cad = Cad & " WHERE "
        
        'Por la ambiguedad del nombre
        vCta = " ("
        vCta = vCta & "codmacta like '" & AUX3 & "')"
        If Excepcion <> "" Then
            Excepcion = Mid(Excepcion & "__________", 1, vEmpresa.DigitosUltimoNivel)
            vCta = vCta & " and not ("
            If Opcion = 4 Then vCta = vCta & " cuentas."
            vCta = vCta & "codmacta like '" & Excepcion & "')"
        End If
        
        
        Cad = Cad & vCta
        
        
        Cad = Cad & " and fechaent between " & DBSet(F, "F") & " and "
        F = DateAdd("m", 1, F)
        F = DateAdd("d", -1, F)
        If F > vParam.fechafin Then
            F = vParam.fechafin
            i = 14 'se salga
        End If
        Cad = Cad & DBSet(F, "F")
        
        
        
        
        
        Cad = " codmacta " & Cad & " GROUP BY codmacta ORDER BY codmacta"
        
        
    
        'tmpevolsal(codusu,codmacta,apertura,importemes1)
        'Contador
        SQL = "Select " & vUsu.Codigo & ",codmacta," & i & ", (sum(coalesce(timported,0))-sum(coalesce(timporteh,0))) " & Cad
        SQL = "INSERT INTO tmpevolsal(codusu,codmacta,apertura,importemes1) " & SQL
        Conn.Execute SQL
        
        
        
        F = DateAdd("d", 1, F) 'Siguiente mes
        
        pb3.Value = pb3.Value + 1
    Next i
     
   
    
    
     
    'Necesito sbaer tb nombre cta
    SQL = "SELECT sum(importemes1),tmpevolsal.codmacta "
    If Opcion = 4 Then SQL = SQL & " , cuentas.nommacta"
    SQL = SQL & " FROM tmpevolsal "
    If Opcion = 4 Then SQL = SQL & ",cuentas"
    SQL = SQL & " where tmpevolsal.codusu =" & vUsu.Codigo
    If Opcion = 4 Then SQL = SQL & " AND tmpevolsal.codmacta = cuentas.codmacta "
    SQL = SQL & " group by tmpevolsal.codmacta"
 
    
    
    '
    'ANTES SQL = "Select sum(coalesce(timported,0))-sum(coalesce(timporteh,0)), " & cad
    Rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    Label3.Caption = ""
    Label3.Refresh
    
    
    NumeroRegistros = 0
    If Rs.EOF Then
        Rs.Close
        'Puede k asi este bien
        Subgrupo = True
        Exit Function
    End If
    
    'Contador
    While Not Rs.EOF
        NumeroRegistros = NumeroRegistros + 1
        Rs.MoveNext
    Wend
    
    'Preparamos el SQL para la insercion de lineas de apunte
    'Montamos la cadena casi al completo
    CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
    
    Rs.MoveFirst
    CONT = 1
    AUX3 = "'" & Text1(1).Text & "'"
    While Not Rs.EOF
    
        Label3.Caption = Rs.Fields(1)
        Label3.Refresh
        i = Int((CONT / NumeroRegistros) * pb1.Max)
        pb1.Value = i
        Importe = Rs.Fields(0)
        If Importe <> 0 Then
            If Opcion = 4 Then
                Cad = SQL & "," & MaxAsiento + CONT & ",'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs.Fields(2)) & "','1','" & Text2(1).Text & "',"
            Else
                Cad = SQL & "," & MaxAsiento + CONT & ",'" & Rs.Fields(1) & "','',960,'" & Text2(1).Text & "',"
            End If
            InsertarLineasDeAsientos Importe, AUX3
        End If
        'Sig
        CONT = CONT + 1
        Rs.MoveNext
    Wend
    Rs.Close
    MaxAsiento = MaxAsiento + CONT - 1
    Subgrupo = True
End Function














Private Sub CadenaLINAPU(Diario As String, Fecha As Date, Num As String)

    If Opcion = 4 Then
        SQL = "INSERT INTO tmphistoapu (codusu, numdiari, desdiari, fechaent, numasien,"
        SQL = SQL & "linliapu, codmacta, nommacta, numdocum, ampconce, timporteD, "
        SQL = SQL & "codccost,timporteH) VALUES ("
        SQL = SQL & vUsu.Codigo & ","
        SQL = SQL & Diario & ",'" & Me.txtDescDiario(1).Text & "'"
        SQL = SQL & ",'" & Format(Fecha, FormatoFecha) & "'," & Num
    Else
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
        SQL = SQL & "codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada) VALUES ("
        SQL = SQL & Diario
        SQL = SQL & ",'" & Format(Fecha, FormatoFecha) & "'," & Num
    End If
    'Primera parte es fija

End Sub

'////////////////////////////////////////////////////////////////////////
'///
'///
'/// Insertamos la linea de asiento correspondiente
Private Function InsertarLineasDeAsientos(ByRef Importe As Currency, ByRef Ctrapar As String) As Boolean
Dim Aux As String
    
        Aux = TransformaComasPuntos(CStr(Abs(Importe)))
        'Deb  centro coste haber
        If Importe < 0 Then     'si es negativo lo pongo al DEBE, si + al haber
            Aux = Aux & ",NULL,NULL"
        Else
            Aux = "NULL,NULL," & Aux
        End If
        ImporteTotal = ImporteTotal + Importe
        '                       contrapartida
        If Opcion = 4 Then
            
            Cad = Cad & Aux & ")"
        Else
            Cad = Cad & Aux & "," & Ctrapar & ",'CONTAB',0)"
        End If
        'Ejecutamos
        Conn.Execute Cad
    
End Function


'//////////////////////////////////////////////////////////////////7
'///  Cuadramos el asiento de perdidas y ganancias
'///
Private Function CuadrarAsiento() As Boolean
Dim Importe As Currency

On Error GoTo ECua
    CuadrarAsiento = False
    
    If Opcion < 4 Then
        SQL = "select sum(timporteD),  sum(timporteH) from hlinapu "
        SQL = SQL & " where numdiari=" & txtDiario(0).Text
        SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Importe = 0
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then Importe = Rs.Fields(0)
            If Not IsNull(Rs.Fields(1)) Then Importe = Importe - Rs.Fields(1)
            
        End If
        Rs.Close
        Set Rs = Nothing
        
        If Importe <> 0 Then
            CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
            Cad = SQL & "," & MaxAsiento + 1 & ",'" & Text1(1).Text & "','',960,'" & Text2(1).Text & "',"
            InsertarLineasDeAsientos Importe, "NULL"
        End If
    Else
        
        'Simulación
        ' "linliapu, codmacta, nommacta, numdocum, ampconce "
        ImporteTotal = ImporteTotal * -1 'Para k cuadre
        ImportePyG = ImporteTotal
        CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
        Cad = SQL & "," & MaxAsiento + 1 & ",'" & Text1(1).Text & "','"
        ' lo meto en el uno, NO en maxasiento
        Cad = SQL & ",1,'" & Text1(1).Text & "','"
        Cad = Cad & Text2(0).Text & "','1','" & Text2(1).Text & "',"
        InsertarLineasDeAsientos ImporteTotal, ""
    End If
    CuadrarAsiento = True
    Exit Function
ECua:
    MuestraError Err.Number, "Cuadrando asiento"
End Function


Private Function HacerElCierre() As Boolean
Dim Ok As Boolean
On Error GoTo EHacerElCierre
    
    HacerElCierre = False
    Set Rs = New ADODB.Recordset
    Conn.Execute "Delete from tmpcierre"  ' no hace falta codusu pq solo puede haber trabajando uno a la vez
    
    Label10.Caption = "Leyendo datos"
    Label10.Refresh
    If Not GeneraTmpCierre Then Exit Function
    
    'Esta grabado el fichero tmpcierre con los importes
    'Fijamos la pb3
    SQL = "Select count(*) from tmpcierre where importe<>0"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroRegistros = 0
    If Not Rs.EOF Then NumeroRegistros = DBLet(Rs.Fields(0), "N")
    Rs.Close
    If NumeroRegistros = 0 Then Exit Function
    NumeroRegistros = NumeroRegistros * 2   'Apertura y cierre
    
    MaxAsiento = 0
    Label10.Caption = "Asiento cierre"
    Label10.Refresh
    Ok = GeneraAsientoCierre
    
    If Ok Then
        Label10.Caption = "Asiento apertura"
        Label10.Refresh
        Ok = GeneraAsientoApertura
    End If
    
    'Si no se ha generado el asiento de cierre y o apertura tenemos k borrarlo
    If Not Ok Then
        'BORRAMOS CIERRE
        SQL = "DELETE FROM hcabapu where fechaent = '" & Format(vParam.fechafin, FormatoFecha) & "' AND Numasien = " & Text1(4).Text
        Conn.Execute SQL
        SQL = "DELETE FROM hlinapu where fechaent = '" & Format(vParam.fechafin, FormatoFecha) & "' AND Numasien = " & Text1(4).Text
        Conn.Execute SQL

        'BORRAMOS APERTURA
        SQL = "DELETE FROM hcabapu where fechaent = '" & Format(Text1(9).Text, FormatoFecha) & "' AND Numasien = 1"
        Conn.Execute SQL
        SQL = "DELETE FROM hlinapu where fechaent = '" & Format(Text1(9).Text, FormatoFecha) & "' AND Numasien = 1"
        Conn.Execute SQL
        
    Else
        Me.cmdCierreEjercicio.Enabled = False
        
        'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
        Cad = Format(DateAdd("yyyy", 1, vParam.fechaini), FormatoFecha)
        SQL = "UPDATE parametros SET fechaini= '" & Cad
        
        vParam.fechafin = DateAdd("yyyy", 1, vParam.fechafin)
        vParam.FechaActiva = DateAdd("yyyy", 1, vParam.FechaActiva)
        If vParam.FechaActiva > vParam.fechafin Then vParam.FechaActiva = vParam.fechaini
        If vParam.FechaActiva < DateAdd("yyyy", 1, vParam.fechaini) Then vParam.FechaActiva = DateAdd("yyyy", 1, vParam.fechaini)
        Cad = Format(vParam.fechafin, FormatoFecha)
        SQL = SQL & "' , fechafin='" & Cad & "'"
        Cad = Format(vParam.FechaActiva, FormatoFecha)
        SQL = SQL & " , fechaactiva='" & Cad & "'"
        
        SQL = SQL & " WHERE fechaini='" & Format(vParam.fechaini, FormatoFecha) & "'"
        
        'ANTES
        'Conn.Execute SQL
        If Not EjecutaSQL(SQL) Then MsgBox "Se ha producido actualizando fechas ejercicio parametros." & vbCrLf & "Cuando finalice avise a soporte técnico de Ariadna Software.", vbExclamation
            
        vParam.fechaini = DateAdd("yyyy", 1, vParam.fechaini)
       
        'los contadores
        'UPDATEAMOS LOS CONTADORES
        'con los nuevos valores
        'Es decir Contsiguiente pasa a actual, y en siguiente ponemos un 2, puesto k el 1 lo reservamos para apertura
        SQL = "UPDATE contadores SET contado1 =  contado2"
        Conn.Execute SQL
        
        'Ponemos en todos un 0
        SQL = "UPDATE contadores SET contado2 = 0"
        Conn.Execute SQL
        
        'Menos en asientos k podremos un 1, ya que se reservara para el cierre
        'del año siguiente
        SQL = "UPDATE contadores SET contado2 = 1 WHERE tiporegi='0'"
        Conn.Execute SQL
        
        
        'Diciembre 2020
        If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
            i = Year(vParam.fechafin)
            Cad = ""
            If i >= 2010 Then
                    i = i - 2000
                    If i >= 10 And i < 100 Then Cad = "OK"
            
            End If
            If Cad = "" Then
                MsgBoxA "****    REVISE CONTADORES FACTURAS PROVEEDOR     *****", vbCritical
            Else
            
                SQL = i & "00000"
                
                SQL = "UPDATE contadores SET contado2 = " & SQL & " WHERE tiporegi='1'"
                If Not EjecutaSQL(SQL) Then MsgBox "Error updateando contador proveeedores: " & vbCrLf & SQL, vbExclamation
            End If
        End If
        
    End If
    
    HacerElCierre = True
    
EHacerElCierre:
    If Err.Number <> 0 Then MuestraError Err.Number
    vParam.Leer
    vParam.FijarAplicarFiltrosEnCuentas vEmpresa.nomempre
    If vEmpresa.TieneTesoreria Then vParamT.Leer
    Set Rs = Nothing
End Function

Private Function SimulaCierreApertura() As Boolean


    SimulaCierreApertura = False
    Set Rs = New ADODB.Recordset
    Conn.Execute "Delete from tmpcierre"  ' no hace falta codusu pq solo puede haber trabajando uno a la vez
    
    Label10.Caption = "Leyendo datos"
    Label10.Refresh
    If Not GeneraTmpCierre Then Exit Function
    
    'Esta grabado el fichero tmpcierre con los importes
    'Fijamos la pb3
    SQL = "Select count(*) from tmpcierre where importe<>0"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroRegistros = 0
    If Not Rs.EOF Then NumeroRegistros = DBLet(Rs.Fields(0), "N")
    Rs.Close
    If NumeroRegistros = 0 Then Exit Function
    NumeroRegistros = NumeroRegistros * 2   'Apertura y cierre
    
    MaxAsiento = 0
    Label10.Caption = "Asiento cierre"
    Label10.Refresh
    
    CadenaLINAPU txtDiario(1).Text, vParam.fechafin, Text1(4).Text
    If Not GeneraLineasSimulacionCierre(True) Then Exit Function

    

    CadenaLINAPU txtDiario(2).Text, CDate(Text1(9).Text), Text1(6).Text
    If Not GeneraLineasSimulacionCierre(False) Then Exit Function
    SimulaCierreApertura = True

End Function






Private Function GeneraTmpCierre() As Boolean
Dim Importe As Currency
Dim vSql As String
Dim B As Boolean
On Error GoTo EGeneraTmpCierre
Dim F As Date

    GeneraTmpCierre = False
    
    'ANTES Febrero 2018
'            cad = vSql & " fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
'            cad = " from hlinapu where " & cad & " GROUP BY codmacta ORDER BY codmacta"
'
'            'Contador
'            SQL = "Select -1 * (sum(coalesce(timported,0))-sum(coalesce(timporteh,0))),codmacta " & cad
'
'            SQL = "INSERT INTO tmpcierre " & SQL
'            Conn.Execute SQL
    
   F = vParam.fechaini
    Conn.Execute "DELETE FROM tmpevolsal WHERE codusu = " & vUsu.Codigo
    pb3.Value = 0
    pb3.Max = 12
    For i = 1 To 12
    
        Label10.Caption = "Calculo mes " & Format(F, "mmmm")
        Label10.Refresh
        pb3.Value = pb3.Value + 1
        
    
        Cad = vSql & " fechaent between " & DBSet(F, "F") & " and "
        
        F = DateAdd("m", 1, F)
        F = DateAdd("d", -1, F)
        If F > vParam.fechafin Then
            F = vParam.fechafin
            i = 14 'se salga
        End If
        Cad = Cad & DBSet(F, "F")
        Cad = " from hlinapu where " & Cad & " GROUP BY codmacta ORDER BY codmacta"
        
        'tmpevolsal(codusu,codmacta,apertura,importemes1)
        'Contador
        SQL = "Select " & vUsu.Codigo & ",codmacta," & i & ", -1 * (sum(coalesce(timported,0))-sum(coalesce(timporteh,0))) " & Cad
        SQL = "INSERT INTO tmpevolsal(codusu,codmacta,apertura,importemes1) " & SQL
        Conn.Execute SQL
        
        
        
        F = DateAdd("d", 1, F) 'Siguiente mes
    Next i
     
    pb3.Value = 0
    pb3.Max = 1000
     
    
    'tmpevolsal(codusu,codmacta,apertura,importemes1)
    Label10.Caption = "Saldos cierre "
    Label10.Refresh
    SQL = " select sum(importemes1),codmacta from tmpevolsal where codusu = " & vUsu.Codigo & " group by codmacta"
    SQL = "INSERT INTO tmpcierre " & SQL
    Conn.Execute SQL

    
    'Y si es simulacion, enonces borramos las cuentas de perdidas ganancias
    If Opcion = 4 Then
        SQL = "Delete from tmpcierre where cta like '" & vParam.GrupoGto & "%'"
        SQL = SQL & " OR  cta like '" & vParam.GrupoVta & "%'"
        Conn.Execute SQL
        
            
        If vParam.GrupoOrd <> "" Then
            SQL = "Delete from tmpcierre where cta like '" & vParam.GrupoOrd & "%'"
            'Excepcion
            If Text1(10).Text <> "" Then SQL = SQL & " AND not (cta like '" & Text1(10).Text & "%')"
            Conn.Execute SQL
        End If
            
        'Si es gran empresa me cargo tb las 8% y 9%
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            SQL = "Delete from tmpcierre where cta like '8%' OR  cta like '9%'"
            Conn.Execute SQL
        End If
        
        'Comprobamos si existe el parametro
        Set miRsAux = New ADODB.Recordset
        SQL = "Select importe from tmpcierre where cta ='" & Text1(1).Text & "'"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        Importe = 0
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then
                SQL = "E"
                Importe = miRsAux.Fields(0)
            End If
        End If
        miRsAux.Close
        
        
        If SQL = "" Then
            'Metemos a la cta 129 de perdias y gancias las perdidas y gancias generadas
            SQL = "INSERT INTO tmpcierre (cta,importe) values ('" & Text1(1).Text & "'," & TransformaComasPuntos(CStr(ImportePyG)) & ")"
        Else
            ImportePyG = ImportePyG + Importe
            SQL = "UPDATE tmpcierre SET Importe= " & TransformaComasPuntos(CStr(ImportePyG)) & " WHERE cta='" & Text1(1).Text & "'"
        End If
        Conn.Execute SQL
        
        
        
        
        'Si es gran empresa, puede que haya saldado las cuentas 8 y 9.
        'Para ello haremos lo siguiente.
        'Iremos a zhistoapu y cogeremos los apuntes que esten relacionados
        'con la regularizacion8y9, es decir, o los que en numdocum pone un 2
        ' y/o numero asiento es el indicado en el txtbox de la regularizacion
        '
        'Agruparemos por cta, sum (importe), para saber cual sera el importe resultante
        'Updatearemos (o crearemos) la linea de apunte , al igual que con la 129
        If vParam.GranEmpresa Then
            Dim L As Collection
            
            SQL = "select codusu,codmacta,sum(timporteD) as d,sum(timporteH) as h from tmphistoapu "
            SQL = SQL & " where codusu = " & vUsu.Codigo & " and numdocum=2 and codmacta <'8' group by 1,2"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Set L = New Collection
            While Not miRsAux.EOF
                
                Importe = DBLet(miRsAux!d, "N")
                If Not IsNull(miRsAux!H) Then Importe = Importe - miRsAux!H
                SQL = miRsAux!codmacta & "|" & CStr(Importe) & "|"
                L.Add SQL
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            'Ahora ya tengo todas las cuentas a actualizar/crear
            For i = 1 To L.Count
                Cad = RecuperaValor(L.Item(i), 1)
                ImporteTotal = CCur(RecuperaValor(L.Item(i), 2))
                SQL = "Select importe from tmpcierre where cta ='" & Cad & "'"
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = ""
                Importe = 0
                If Not miRsAux.EOF Then
                    If Not IsNull(miRsAux.Fields(0)) Then
                        SQL = "E"
                        Importe = miRsAux.Fields(0)
                    End If
                End If
                miRsAux.Close
                
                
                If SQL = "" Then
                    'Metemos a la cta 129 de perdias y gancias las perdidas y gancias generadas
                    SQL = "INSERT INTO tmpcierre (cta,importe) values ('" & Cad & "'," & TransformaComasPuntos(CStr(ImportePyG)) & ")"
                Else
                    ImporteTotal = ImporteTotal + Importe
                    SQL = "UPDATE tmpcierre SET Importe= " & TransformaComasPuntos(CStr(ImporteTotal)) & " WHERE cta='" & Cad & "'"
                End If
                Conn.Execute SQL
            Next i
        End If  'granempresa
        
        Set miRsAux = Nothing
    End If
    
    
    
    GeneraTmpCierre = True
    Exit Function
EGeneraTmpCierre:
    MuestraError Err.Number, "Genera TmpCierre"
End Function




Private Function GeneraAsientoCierre() As Boolean
    On Error GoTo EGeneraAsientoCierre

    GeneraAsientoCierre = False

    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien,  obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(1).Text
    Cad = SQL & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(4).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Generación Asiento Cierre')"
    Conn.Execute Cad

    
    CadenaLINAPU txtDiario(1).Text, vParam.fechafin, Text1(4).Text
    If Not GeneraLineasCierre Then Exit Function

    
    GeneraAsientoCierre = True
    Exit Function
EGeneraAsientoCierre:
    MuestraError Err.Number
End Function



Private Function GeneraLineasCierre() As Boolean
Dim CONT As Integer
Dim Importe As Currency
Dim Aux As String
    On Error GoTo EGeneraLineasCierre
    GeneraLineasCierre = False
    Rs.Open "SELECT * from tmpcierre WHERE importe <>0 ORDER By Cta", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    CONT = 1
    
    ' hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
    ' codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada
    
    
    While Not Rs.EOF
        Cad = SQL & "," & CONT & ",'" & Rs.Fields(1) & "','',980,'" & Text2(2).Text & "',"
        Importe = Rs.Fields(0)
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        Cad = Cad & Aux
        Cad = Cad & ",NULL,'contab',0)"
        
        i = Int(((CONT + MaxAsiento) / NumeroRegistros) * pb3.Max)
        pb3.Value = i
        
        Conn.Execute Cad
        
        'Siguiente
        CONT = CONT + 1
        Rs.MoveNext
    Wend
    Rs.Close
    GeneraLineasCierre = True
    MaxAsiento = CONT - 1
    Exit Function
EGeneraLineasCierre:
    MuestraError Err.Number
End Function



Private Function GeneraAsientoApertura() As Boolean
    On Error GoTo EGeneraAsientoApertura

    GeneraAsientoApertura = False

    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(2).Text
    Cad = SQL & ",'" & Format(Text1(9).Text, FormatoFecha) & "'," & Text1(6).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Generación Asiento Apertura')"
    Conn.Execute Cad
    
    CadenaLINAPU txtDiario(2).Text, CDate(Text1(9).Text), Text1(6).Text
    If Not GeneraLineasApertura Then Exit Function

    GeneraAsientoApertura = True
    Exit Function
EGeneraAsientoApertura:
    MuestraError Err.Number
End Function



Private Function GeneraLineasApertura() As Boolean
Dim CONT As Integer
Dim Importe As Currency
Dim Aux As String
    On Error GoTo EGeneraLineasApertura
    GeneraLineasApertura = False
    Rs.Open "SELECT * from tmpcierre WHERE importe <>0 ORDER BY Cta ", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    CONT = 1
    
    ' hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
    ' codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada
    
    
    While Not Rs.EOF
        Cad = SQL & "," & CONT & ",'" & Rs.Fields(1) & "','',970,'" & Text2(3).Text & "',"
        Importe = Rs.Fields(0)
        Importe = Importe * -1
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        Cad = Cad & Aux
        Cad = Cad & ",NULL,'contab',0)"
        
        i = Int(((CONT + MaxAsiento) / NumeroRegistros) * pb3.Max)
        pb3.Value = i
        
        Conn.Execute Cad
        
        'Siguiente
        CONT = CONT + 1
        Rs.MoveNext
    Wend
    Rs.Close
    GeneraLineasApertura = True
    Exit Function
EGeneraLineasApertura:
    MuestraError Err.Number
End Function


Private Function ActualizarAsientoCierreApertura() As Boolean
Screen.MousePointer = vbHourglass

            ActualizarAsientoCierreApertura = False
            
            
            'CIERRE
            AlgunAsientoActualizado = False
            Me.Refresh
            If Not AlgunAsientoActualizado Then Exit Function
            

            'Apertura
            AlgunAsientoActualizado = False
            Me.Refresh
            If Not AlgunAsientoActualizado Then Exit Function

            Me.Refresh
            ActualizarAsientoCierreApertura = True
End Function

'Comprobamos k ha fecha fin ejercicio anterior, no hay cierre
Private Function NoHayCierre(FechaCierre As Date) As Boolean

    On Error GoTo ENoHayCierre
    NoHayCierre = True
    
    SQL = "Select numasien from hlinapu WHERE codconce=980 AND fechaent='" & Format(FechaCierre, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then NoHayCierre = False
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ENoHayCierre:
    MuestraError Err.Number, "Comprobar cierre anterior"
End Function




Private Function GeneraLineasSimulacionCierre(EsCierre As Boolean) As Boolean
Dim CONT As Integer
Dim Importe As Currency
Dim Aux As String
    On Error GoTo EGeneraLineasSimulacionCierre
    GeneraLineasSimulacionCierre = False
    Cad = "select tmpcierre.*,nommacta from tmpcierre,cuentas where tmpcierre.cta=cuentas.codmacta"
    Cad = Cad & " AND importe <>0 ORDER By Cta"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Cont = 1
    CONT = 2
    'linliapu, codmacta, nommacta, numdocum, ampconce, timporteD, "
    'codccost,timporteH) VALUES ("
    
    While Not Rs.EOF
        
        Importe = Rs.Fields(0)
        If EsCierre Then
            Cad = SQL & "," & CONT & ",'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs!Nommacta) & "','3','" & Text2(2).Text & "',"
        Else
            Cad = SQL & "," & CONT & ",'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs!Nommacta) & "','4','" & Text2(3).Text & "',"
            Importe = Importe * -1
        End If
        
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        Cad = Cad & Aux & ")"
    
        i = Int(((CONT + MaxAsiento) / (NumeroRegistros + 3)) * pb3.Max)
        pb3.Value = i
        
        Conn.Execute Cad
        
        'Siguiente
        CONT = CONT + 1
        Rs.MoveNext
    Wend
    Rs.Close
    GeneraLineasSimulacionCierre = True
    MaxAsiento = CONT - 1
    Exit Function
EGeneraLineasSimulacionCierre:
    MuestraError Err.Number
End Function



Private Function ExisteAsientosDescerrar() As Boolean

'Tiene k existe el asiento de apertura del año siguient
    Cad = CadenaFechasActuralSiguiente(True)
    SQL = "Select numasien from hlinapu WHERE codconce=970"
    SQL = SQL & " AND " & Cad
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ExisteAsientosDescerrar = True
    Else
        ExisteAsientosDescerrar = False
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Private Function ExistenApuntesEjercicioAnterior() As Boolean

'Tiene k existe el asiento de apertura del año siguient
    SQL = "Select count(*) from hlinapu WHERE "
    SQL = SQL & " Fechaent < '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ExistenApuntesEjercicioAnterior = False
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then ExistenApuntesEjercicioAnterior = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Private Function ExistenApuntesEjercicioSiguiente() As Boolean

'Tiene k existe el asiento de apertura del año siguient
    SQL = "Select count(*) from hlinapu WHERE "
    SQL = SQL & " Fechaent > '" & Format(vParam.fechafin, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ExistenApuntesEjercicioSiguiente = False
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then ExistenApuntesEjercicioSiguiente = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function




Private Function HacerDescierre() As Boolean
Dim N As Long
Dim MaxAsien As Long
Dim F As Date
Dim Sql1 As String
Dim Sql2 As String



    


On Error GoTo EHacerDescierre
    HacerDescierre = False
    Screen.MousePointer = vbHourglass
    
    'Si ya hay un 960 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    'P y G
    MaxAsien = 0
    Label19(3).Caption = "Perdidas y ganancias"
    Label19(3).Refresh

    SQL = ""
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "'"
    
    Cad = "(select count(*) atos from hlinapu where 1=1"
    Cad = Cad & SQL & " AND codconce =960"
    Cad = Cad & " group by numasien) as sel "

    
    
    
    SQL = DevuelveDesdeBD("count(*)", Cad, " 1", "1 ")
    If Val(SQL) > 1 Then Err.Raise 513, , "Mas de un asiento con concepto perdidas y ganancias"
         

   
    SQL = " codconce=960 "
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "' AND 1 "
    Cad = DevuelveDesdeBD("concat(numasien,'|',fechaent,'|')", "hlinapu", SQL, "1")
    If Cad <> "" Then
    
    
            MaxAsien = Val(RecuperaValor(Cad, 1))
            SQL = RecuperaValor(Cad, 2)
            F = CDate(SQL)
            ' lineas
            SQL = "DELETE FROM hlinapu WHERE numasien= " & MaxAsien
            SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
            Conn.Execute SQL
            
            SQL = "DELETE FROM hcabapu WHERE numasien= " & MaxAsien
            SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
            Conn.Execute SQL
            
            
           
            
    End If
   
   




    Me.Refresh
    espera 0.25
    Me.Refresh



    Label19(3).Caption = "Regularización 8 y 9"
    Label19(3).Refresh

    

    SQL = " codconce=961 "
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "' AND 1 "
    Cad = DevuelveDesdeBD("concat(numasien,'|',fechaent,'|')", "hlinapu", SQL, "1")
    If Cad <> "" Then
    
            MaxAsien = Val(RecuperaValor(Cad, 1))
            SQL = RecuperaValor(Cad, 2)
            F = CDate(SQL)
            ' lineas
            SQL = "DELETE FROM hlinapu WHERE numasien= " & MaxAsien
            SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
            Conn.Execute SQL
            
            SQL = "DELETE FROM hcabapu WHERE numasien= " & MaxAsien
            SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
            Conn.Execute SQL
            
            
           
            
    End If
            
            
    Me.Refresh
    espera 0.25
    Me.Refresh



    'Cierre
    'Si hay un 980  en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Label19(3).Caption = "Cierre"
    Label19(3).Refresh

    SQL = " codconce=980 "
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "' AND 1 "
    Cad = DevuelveDesdeBD("concat(numasien,'|',fechaent,'|')", "hlinapu", SQL, "1")
    If Cad = "" Then Err.Raise 513, , "Sin cierre"
    
    MaxAsien = Val(RecuperaValor(Cad, 1))
    SQL = RecuperaValor(Cad, 2)
    F = CDate(SQL)
    ' lineas
    SQL = "DELETE FROM hlinapu WHERE numasien= " & MaxAsien
    SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
    Conn.Execute SQL
    
    SQL = "DELETE FROM hcabapu WHERE numasien= " & MaxAsien
    SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
    Conn.Execute SQL
    
    Me.Refresh
    espera 0.25
    Me.Refresh
        
        
    Label19(3).Caption = "Apertura"
    Label19(3).Refresh
    'Si ya hay un  970 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    SQL = " codconce=970 "
    Cad = Format(vParam.fechaini, FormatoFecha)
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent >='" & Cad & "'"
    Cad = Format(vParam.fechafin, FormatoFecha)
    Cad = Format(Cad, FormatoFecha)
    SQL = SQL & " AND fechaent <='" & Cad & "' AND 1 "
    Cad = DevuelveDesdeBD("concat(numasien,'|',fechaent,'|')", "hlinapu", SQL, "1")
    If Cad = "" Then Err.Raise 513, , "Sin cierre"
    
    MaxAsien = Val(RecuperaValor(Cad, 1))
    SQL = RecuperaValor(Cad, 2)
    F = CDate(SQL)
    
    
    ' lineas
    SQL = "DELETE FROM hlinapu WHERE numasien= " & MaxAsien
    SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
    Conn.Execute SQL
    
    SQL = "DELETE FROM hcabapu WHERE numasien= " & MaxAsien
    SQL = SQL & " AND fechaent ='" & Format(F, FormatoFecha) & "'"
    Conn.Execute SQL
    
    Me.Refresh
    espera 0.25
    Me.Refresh


    'Hay k bajar una año las fechas de parametros de INICIO y FIN ejerecicio
    'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
    'Reestablecemos las fechas
    'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
    Label19(3).Caption = "Contadores"
    Label19(3).Refresh
    i = -1
    Cad = Format(DateAdd("yyyy", i, vParam.fechaini), FormatoFecha)
    SQL = "UPDATE parametros SET fechaini= '" & Cad
    Cad = Format(DateAdd("yyyy", i, vParam.fechafin), FormatoFecha)
    SQL = SQL & "' , fechafin='" & Cad & "'"
    'Fechaactiva
    Cad = Format(DateAdd("yyyy", i, vParam.FechaActiva), FormatoFecha)
    SQL = SQL & " , FechaActiva='" & Cad & "'"
    
    SQL = SQL & " WHERE fechaini='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Conn.Execute SQL
    
    vParam.fechaini = DateAdd("yyyy", i, vParam.fechaini)
    vParam.fechafin = DateAdd("yyyy", i, vParam.fechafin)
    'Fecha activa
    vParam.FechaActiva = DateAdd("yyyy", i, vParam.FechaActiva)


    Set Rs = New ADODB.Recordset
    
    SQL = "SELECT tiporegi, nomregis, contado1, contado2 from Contadores order by tiporegi"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    
    While Not Rs.EOF
        If DBLet(Rs!tiporegi, "T") = "0" Then ' asientos
            SQL = "select max(numasien) from hlinapu where fechaent between  " & DBSet(vParam.fechaini, "F") & " and  " & DBSet(vParam.fechafin, "F")
            
            Sql1 = "select max(numasien) from hlinapu where fechaent > " & DBSet(vParam.fechafin, "F")
            If DevuelveValor(Sql1) = 0 Then
                Sql1 = "select 1 from hlinapu "
            End If
        Else
            If IsNumeric(DBLet(Rs!tiporegi, "T")) Then ' facturas de proveedor
                SQL = "select max(numregis) from factpro where fecharec between  " & DBSet(vParam.fechaini, "F") & " and  " & DBSet(vParam.fechafin, "F")
                SQL = SQL & " and numserie = " & DBSet(Rs!tiporegi, "T")
                Sql1 = "select max(numregis) from factpro where fecharec > " & DBSet(vParam.fechafin, "F")
                Sql1 = Sql1 & " and numserie = " & DBSet(Rs!tiporegi, "T")
            Else ' facturas de cliente
                SQL = "select max(numfactu) from factcli where fecfactu between  " & DBSet(vParam.fechaini, "F") & " and  " & DBSet(vParam.fechafin, "F")
                SQL = SQL & " and numserie = " & DBSet(Rs!tiporegi, "T")
                
                Sql1 = "select max(numfactu) from factcli where fecfactu > " & DBSet(vParam.fechafin, "F")
                Sql1 = Sql1 & " and numserie = " & DBSet(Rs!tiporegi, "T")
            End If
        End If
        
        'actualizamos
        Sql2 = "update contadores set contado1 = " & DevuelveValor(SQL) & ",contado2 = " & DevuelveValor(Sql1)
        Sql2 = Sql2 & " where tiporegi = " & DBSet(Rs!tiporegi, "T")
        Conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    

    
    HacerDescierre = True
    Exit Function
EHacerDescierre:
    MuestraError Err.Number, "Proc. HacerDescierre", Err.Description
End Function




Private Sub EliminarAsiento()
        Screen.MousePointer = vbHourglass
        AlgunAsientoActualizado = False
End Sub

Private Sub PonerGrandesEmpresas()
Dim B As Boolean

    B = False
    If vParam.NuevoPlanContable Then B = vParam.GranEmpresa
    'Cuando sean autmocion en plan nuevo ya veremos como resolvemos esto
    'FALTA###
    
    
    GElabel1.visible = B
    Me.GELine.visible = B
    Label5(17).visible = B
    Label5(16).visible = B
    Label5(16).visible = B
    Label5(15).visible = B
    Label5(14).visible = False 'vEmpresa.GranEmpresa
    Label5(18).visible = B
    Text2(4).visible = B
    Me.Image1(3).visible = B
    Me.txtDiario(3).visible = B
    Me.txtDescDiario(3).visible = B
    Text1(13).visible = B
    Text1(14).visible = False 'vEmpresa.GranEmpresa
    Text1(11).visible = B
    Text1(12).visible = B
    Frame4.visible = B
    
    If Not B Then
        'SOlo para la antigua
        '---------------------------------------------------------
        'Si tiene el otro grupo de perdidas y ganacias entonces
        'tenemos k solicitar la excepcion a digitos de tercer nivel
        'ofertando el de parametros
        B = (vParam.GrupoOrd <> "") And (vParam.Automocion <> "")
        Label5(13).visible = B
        Text1(10).visible = B
        If B Then Text1(10).Text = vParam.Automocion
    End If
    
End Sub



Private Function ComprobarCierreCuentas8y9() As Boolean

''?????????????????????????
'??????????????CAMBIARLO POR HLINAPU

    On Error GoTo EComprobarCierreCuentas8y9
    ComprobarCierreCuentas8y9 = False

    Conn.Execute "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo

    Set Rs = New ADODB.Recordset
    
    Cad = "select " & vUsu.Codigo & ",codmacta,'T',sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) from hlinapu"
    Cad = Cad & " WHERE mid(codmacta,1,1) in ('8','9') AND "
    Cad = Cad & "  fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    
    
    
    Cad = Cad & " GROUP BY codmacta"
    'Insertamos en tmpcierre
    Cad = "INSERT INTO TMPCIERRE1 " & Cad
    Conn.Execute Cad
    
    'COJEREMOS TODAS LAS CUENTAS 8 y 9 a tres digitos y comprobaremos que en
    'la configuracion tienen puesto en el campo cuentaba
    '
    'Cruzamos tmpcierr1 con codusu = vusu y left join con cuentas
    'Veremos si hay null con lo cual esta mal y si no, updatearemos tmpcierre
    Cad = "select cta,nommacta,mid(iban, 15,10) cuentaba from tmpcierre1,cuentas where codusu = " & vUsu.Codigo & " and cta = codmacta"
  
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    i = 0
    While Not Rs.EOF
        If DBLet(Rs!cuentaba, "T") = "" Then
            
            i = i + 1
            Cad = Cad & "     " & Rs!Cta
            If (i Mod 5) = 0 Then Cad = Cad & vbCrLf
        
        Else
        
            'ASi, tanto para la simulacion, como para el cierre ya se contra que cuentas saldan las del 8 9
            Conn.Execute "UPDATE tmpcierre1 SET nomcta = '" & Rs!cuentaba & "' WHERE codusu = " & vUsu.Codigo & " and cta = '" & Rs!Cta & "'"
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    If i > 0 Then
        Cad = "Cuentas sin configurar el cierre: " & vbCrLf & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Set Rs = Nothing
        Exit Function
    End If
    
    
    'OK tiene todas las cuentas configuradas
    Cad = "Select tmpcierre1.nomcta, cuentas.codmacta from tmpcierre1 left join cuentas on nomcta = cuentas.codmacta where tmpcierre1.codusu=" & vUsu.Codigo
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    i = 0
    While Not Rs.EOF
        
        If IsNull(Rs!codmacta) Then
            i = i + 1
            Cad = "    " & Rs!nomcta
            If (i Mod 5) = 0 Then Cad = Cad & vbCrLf
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    If i > 0 Then
        Cad = "Cuentas de cierre configurada, pero no existen: " & vbCrLf & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Set Rs = Nothing
        Exit Function
    End If
    
            
                
            
    
    
    ComprobarCierreCuentas8y9 = True
    Set Rs = Nothing
    
    
    Exit Function
EComprobarCierreCuentas8y9:
    Set Rs = Nothing
    MuestraError Err.Number, "Comprobar Cierre Cuentas 8y9"
End Function

Private Function ASiento8y9() As Boolean
Dim Ok As Byte

    ASiento8y9 = False
    
        
    Ok = GenerarASiento8y9     '0.- MAL   '1,. Bien pero sin generar apunte(para que no contabilice)    2.- Bien
    
    If Opcion = 4 Then
        ASiento8y9 = (Ok > 0)
        Exit Function
    End If
    
    
    'Llamamos a actualizar el asiento, para pasarlo a hco
    If Ok = 2 Then
            Screen.MousePointer = vbHourglass
            AlgunAsientoActualizado = False
            Me.Refresh
            If AlgunAsientoActualizado Then
                Ok = 1
            Else
                Ok = 0
            End If
    End If
        
        
    'Si entra mal hay que borrar los apuntes que pudieran haberse creado
    If Ok = 0 Then
        'Comun lineas y cabeceras
        SQL = " where numdiari=" & txtDiario(2).Text
        SQL = SQL & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(11).Text & ""
        
        'Borramos por si acaso ha insertado lineas
        Cad = "Delete FROM hlinapu" & SQL
        Conn.Execute Cad
    
        'Borramos la cabcecera del apunte
        Cad = "DELETE FROM hcabapu" & SQL
        Conn.Execute Cad
        
        Label3.Caption = ""
    
    End If
    
    ASiento8y9 = Ok > 0
    

End Function


Private Function GenerarASiento8y9() As Byte
Dim CuentaSaldo As String
Dim Importe As Currency
Dim ImpComprobacion As Currency
Dim RT As ADODB.Recordset
Dim CONT As Long

    On Error GoTo EASiento8y9
    GenerarASiento8y9 = 0
    
    'Generamos los  apuntes, sobre cabapu, y luego los actualizamos

    
    'Cogemos tmpcierre1 que tendra cta, cta saldada
    'Ejemlo   codusu, cta, saldar
    '                 910   1290003
    '                 920   1290003
    '                 830   1200002
    'Entonces cogeremos los saldos para estas cuentas y las iremos saldando
    'Cargaremos RT con los saldos a 3 digitos (como no ocupara mucho... NO problemo
    Set Rs = New ADODB.Recordset
    
    SQL = "Select count(*) from tmpcierre1 where codusu = " & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroRegistros = 0
    
    If Not Rs.EOF Then NumeroRegistros = DBLet(Rs.Fields(0), "N")
    Rs.Close
    If NumeroRegistros = 0 Then
    
    
        'Automaticamente el numero de registro que se le iba asignar pasa al cierre
        Text1(4).Text = Text1(11).Text
    
    
    
        'OK
        Set Rs = Nothing
        GenerarASiento8y9 = 1
        Exit Function
    End If
    
    If Opcion <> 4 Then
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(3).Text
        Cad = SQL & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(11).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Generación Asiento 8 y 9')"
        
        Conn.Execute Cad
    End If
    
    
    
    
    NumeroRegistros = NumeroRegistros + 1 'Para que no desborde
    
    SQL = "Select tmpcierre1.* from tmpcierre1 where  codusu = " & vUsu.Codigo & " ORDER BY nomcta,cta"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CuentaSaldo = ""
    MaxAsiento = 1

    
    
    'Preparamos el SQL para la insercion de lineas de apunte
    'Montamos la cadena casi al completo
    CadenaLINAPU txtDiario(3).Text, vParam.fechafin, Text1(11).Text

    Set RT = New ADODB.Recordset
    
    While Not Rs.EOF
        If CuentaSaldo <> Rs!nomcta Then
            
                'SALDAMOS
            If CuentaSaldo <> "" Then SaldarCuenta8y9 CuentaSaldo
 
            CuentaSaldo = Rs!nomcta   'Cuenta saldo
            ImporteTotal = 0
        End If
    
        
        
        'Progress y label
        Label3.Caption = Rs!Cta & " - " & Rs!nomcta
        Label3.Refresh
        CONT = CONT + 1
        i = Int((CONT / NumeroRegistros) * pb1.Max)
        pb1.Value = i
        
        
        
        '$$$
        Cad = Mid(Rs!Cta & "__________", 1, vEmpresa.DigitosUltimoNivel)
        Cad = " WHERE hlinapu.codmacta=cuentas.codmacta AND hlinapu.codmacta like '" & Cad & "' AND "
        Cad = "select hlinapu.codmacta,sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) as miImporte,nommacta from hlinapu,cuentas" & Cad
        Cad = Cad & " fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
        
        Cad = Cad & " GROUP BY codmacta"
        
        ImpComprobacion = 0
        RT.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RT.EOF
            Importe = RT!miImporte  'importe
            
            If Importe <> 0 Then
                If Opcion = 4 Then
                    ' "linliapu, codmacta, nommacta, numdocum, ampconce "
                    Cad = SQL & "," & MaxAsiento & ",'" & RT!codmacta & "','" & DevNombreSQL(RT!Nommacta) & "','2','" & Text2(4).Text & "',"
                Else
                    Cad = SQL & "," & MaxAsiento & ",'" & RT!codmacta & "','',961,'" & Text2(4).Text & "',"
                End If
                InsertarLineasDeAsientos Importe, "NULL"
            End If
        
            'Sig
            ImpComprobacion = ImpComprobacion + Importe
            MaxAsiento = MaxAsiento + 1
            RT.MoveNext
        Wend
        RT.Close
        
        
        Importe = Rs!acumPerD   'Para combrobar que a ultimo nivel suma igual que a 3 digitos
        
        If ImpComprobacion <> Importe Then
            Cad = "Error obteniendo saldos. " & vbCrLf & "Subgrupo: " & Rs!Cta & vbCrLf
            Cad = Cad & "Imp 3 digitos: " & Importe & vbCrLf & "Ultimo nivel: " & ImpComprobacion
            If Opcion <> 4 Then Cad = Cad & vbCrLf & vbCrLf & " No puede continuar con el cierre"
            MsgBox Cad, vbExclamation
            If Opcion <> 4 Then
                Rs.Close
                Exit Function
            End If
        End If
            
            
        'Sigueinte subgrupoo
        Rs.MoveNext
    Wend
            
    Rs.Close
    Set Rs = Nothing
    Set RT = Nothing
    If CuentaSaldo <> "" Then SaldarCuenta8y9 CuentaSaldo
    ImporteTotal = 0
    
    GenerarASiento8y9 = 2

   Exit Function
EASiento8y9:
    Set Rs = Nothing
    MuestraError Err.Number, Err.Description
End Function


Private Sub SaldarCuenta8y9(LaCuenta As String)
Dim C As String
    ImporteTotal = ImporteTotal * -1 'Para k cuadre
    If ImporteTotal <> 0 Then
        If Opcion = 4 Then
            C = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", LaCuenta, "T")
            ' "linliapu, codmacta, nommacta, numdocum, ampconce "
            Cad = SQL & "," & MaxAsiento & ",'" & LaCuenta & "','" & DevNombreSQL(C) & "','2','" & Text2(4).Text & "',"
        Else
            Cad = SQL & "," & MaxAsiento & ",'" & LaCuenta & "','',961,'" & Text2(4).Text & "',"
        End If
        InsertarLineasDeAsientos ImporteTotal, "NULL"
        MaxAsiento = MaxAsiento + 1
    End If
End Sub




