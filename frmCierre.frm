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
            Caption         =   "Fecha �ltima amortizaci�n"
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
            Width           =   1125
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
            Width           =   1125
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   930
            Picture         =   "frmCierre.frx":000C
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
            Caption         =   "Fecha contabilizaci�n"
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
            Caption         =   "N�mero de asiento"
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
         Caption         =   "Regularizaci�n 8 y 9"
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
            Left            =   1200
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
            Caption         =   "N� asiento"
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
            Left            =   930
            Picture         =   "frmCierre.frx":0A0E
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
            Index           =   17
            Left            =   150
            TabIndex        =   67
            Top             =   420
            Width           =   585
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha contabilizaci�n"
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
            Width           =   4215
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
            Width           =   1125
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
            Index           =   4
            Left            =   2400
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   1380
            Width           =   1125
         End
         Begin VB.Label Label5 
            Caption         =   "N�mero de asiento"
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
            Picture         =   "frmCierre.frx":1410
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
            Caption         =   "Fecha contabilizaci�n"
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
         Caption         =   "P�rdidas y Ganancias"
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
            Left            =   1170
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
            Left            =   2250
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
            Left            =   2460
            TabIndex        =   37
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
            Width           =   5025
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
            Width           =   5235
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
            Left            =   5670
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   960
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
            Left            =   1920
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   1440
            Width           =   1125
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   870
            Picture         =   "frmCierre.frx":1E12
            Top             =   390
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
            Caption         =   "Fecha contabilizaci�n"
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
            Caption         =   "N� asiento"
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
            Left            =   4560
            TabIndex        =   41
            Top             =   960
            Width           =   1365
         End
         Begin VB.Label Label5 
            Caption         =   "Grupo excepci�n"
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
            Left            =   180
            TabIndex        =   40
            Top             =   1440
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
            Caption         =   "Simulaci�n"
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
         Height          =   315
         Left            =   1530
         TabIndex        =   17
         Top             =   9210
         Width           =   4275
      End
      Begin VB.Label Label7 
         Caption         =   "P�RDIDAS Y GANANCIAS"
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
         Caption         =   "No debe haber nadie m�s trabajando contra esta contabilidad"
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
   Begin VB.Frame frameDescierre 
      BorderStyle     =   0  'None
      Caption         =   $"frmCierre.frx":2814
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
         Top             =   1920
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
         Left            =   4290
         TabIndex        =   23
         Top             =   3270
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
         Left            =   3030
         TabIndex        =   22
         Top             =   3270
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
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "No debe haber nadie m�s trabajando contra esta contabilidad"
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
         Left            =   210
         TabIndex        =   96
         Top             =   720
         Width           =   4875
      End
      Begin VB.Label Label19 
         Caption         =   "Trabajar nadie en esta contabilidad"
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
         TabIndex        =   24
         Top             =   2490
         Width           =   5070
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
    '4.- Simulacion de cierre, es decir, mostrar� un listado con
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
Dim cad As String
Dim Sql As String
Dim Rs As Recordset

Dim I As Integer
Dim NumeroRegistros As Long
Dim MaxAsiento As Long
Dim ImporteTotal As Currency
Dim ImportePyG As Currency


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCierreEjercicio_Click()
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
    
    ' nuevo, comprobamos que no exista ya el asiento nro 1 en el a�o siquiente, si lo hay dar aviso
    Label10.Caption = "Comprobar que no exista el asiento 1 en el a�o siguiente"
    Label10.Refresh
    Ok = ExisteAsientoUnoLibre
    If Not Ok Then
        MsgBox "Asiento N� 1 reservado para la apertura, est� ocupado. Renumere ejercicio siguiente o llame a Ariadna.", vbExclamation
        Exit Sub
    End If
    
'    Ok = True
    For I = 0 To 3
        If txtDiario(I).Text = "" Then
           If I < 2 Or (I = 3 And vParam.GranEmpresa And vParam.NuevoPlanContable) Then
                MsgBox "Seleccione el diario.", vbExclamation
                Ok = False
                Exit For
            End If
        End If
    Next I
    If Not Ok Then Exit Sub
    
    
    'Coamprobamos las cuentas 8 9
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        If Not ComprobarCierreCuentas8y9 Then Exit Sub
    End If
    
    
    
    Ok = UsuariosConectados("")
    If Not Ok Then
        Sql = "Seguro que desea cerrar el ejercicio?"
        If MsgBox(Sql, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
        
    Else
        'Hay usuarios conectados
        If vUsu.Nivel > 1 Then
            'NO TIENE PERMISOS
            Exit Sub
        Else
            Sql = "No es recomendado, pero, �desea continuar con el proceso?"
            If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
        End If
    End If
    
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquear a nivel de BD.", vbExclamation
        Exit Sub
    End If
    
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
        pb3.Visible = True
        Label10.Caption = "Perdidas y Ganancias"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        Ok = ASientoPyG
    End If
    
    
    If Ok Then
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            pb3.Value = 0
            pb3.Visible = True
            Label10.Caption = "Regularizacion 8 y 9"
            Label10.Refresh
            Label3.Caption = ""
            Label3.Refresh
            Ok = ASiento8y9
        End If
    End If
    
    
    Me.Refresh
    DoEvents
    espera 0.2
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If Ok Then
        pb3.Value = 0
        pb3.Visible = True
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
    Bloquear_DesbloquearBD False
    pb3.Visible = False
    
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
            MsgBox "Error en la comprobaci�n del cuadre.", vbExclamation
        Else
            If MsgBox("Error en cuadre. �Continuar igualmente pese al riesgo?", vbQuestion + vbYesNoCancel) = vbYes Then ComprobarCuadre = True
        End If
    End If
    Me.Refresh
End Function


Private Function ComprobarAsientosDescuadrados() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset

    
    On Error GoTo eComprobarAsientosDescuadrados
    
    Screen.MousePointer = vbHourglass
    
    
    ComprobarAsientosDescuadrados = False
    
    
    
    SqlInsert = "insert into tmphistoapu (codusu, numdiari, numasien, fechaent, timported, timporteh) "
    Sql = " select " & vUsu.Codigo & ", numdiari, numasien, fechaent, sum(coalesce(timported,0)), sum(coalesce(timporteh,0)) "
    Sql = Sql & " from hlinapu "
    Sql = Sql & " where fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    Sql = Sql & " group by numdiari, numasien, fechaent "
    Sql = Sql & " having sum(coalesce(timported,0)) <> sum(coalesce(timporteh,0))  "
    Sql = Sql & " order by numdiari, numasien, fechaent "
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
    
        Conn.Execute "delete from tmphistoapu where codusu = " & vUsu.Codigo
        
        Conn.Execute SqlInsert & Sql
        
        frmMensajes.Opcion = 30
        frmMensajes.Show vbModal
        ComprobarAsientosDescuadrados = False
            
    Else
        ComprobarAsientosDescuadrados = True
    
    End If
    
    Exit Function
    
eComprobarAsientosDescuadrados:
    MuestraError Err.Number, "Comprobar Asientos Descuadrados", Err.Description
End Function

Private Function ExisteAsientoUnoLibre() As Boolean
Dim Sql As String

    Sql = "select * from hcabapu where numasien = 1 and fechaent between " & DBSet(DateAdd("yyyy", 1, CDate(vParam.fechaini)), "F") & " and " & DBSet(DateAdd("yyyy", 1, CDate(vParam.fechafin)), "F")
    ExisteAsientoUnoLibre = (TotalRegistrosConsulta(Sql) = 0)

End Function



Private Function ComprobarFacturasSinAsientos() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset

    On Error GoTo eComprobarFacturasSinAsientos


    Screen.MousePointer = vbHourglass
    ComprobarFacturasSinAsientos = False
    
    
    SqlInsert = "insert into tmpfaclin (codusu,numserie,nomserie,Numfac,Fecha, total) "
    Sql = " select " & vUsu.Codigo & ", numserie, contadores.nomregis, numfactu, fecfactu, totfaccl "
    Sql = Sql & " from factcli inner join contadores on factcli.numserie = contadores.tiporegi "
    Sql = Sql & " where fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    Sql = Sql & " and (numasien is null or numasien = 0)"
    Sql = Sql & " union "
    Sql = Sql & " select " & vUsu.Codigo & ", numserie, contadores.nomregis, numfactu, fecharec, totfacpr "
    Sql = Sql & " from factpro inner join contadores on factpro.numserie = contadores.tiporegi "
    Sql = Sql & " where fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    Sql = Sql & " and (numasien is null or numasien = 0)"
    Sql = Sql & " order by 1,2,3,4 "
    
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
    
        Conn.Execute "delete from tmpfaclin where codusu = " & vUsu.Codigo
        
        Conn.Execute SqlInsert & Sql
        
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

    Sql = "Seguro que desea deshacer el cierre?"
    If MsgBox(Sql, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
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
    cad = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(cad) <> "ARIADNA" Then
        If cad <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    'Mensaje de lineas en introduccion de asientos
    Ok = True
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquear a nivel de BD.", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    cmdDescerrar.Enabled = False
    CmdCancel(5).Enabled = False
    Me.Refresh
    
    Me.Refresh
    espera 0.3
    Me.Refresh
    If Ok Then
        Label19.Caption = "Eliminar asientos"
        Label19.Refresh
        Ok = HacerDescierre
    End If
    
    'Desbloqueamos BD
    Bloquear_DesbloquearBD False
    
    If Ok Then
    
        
        vLog.Insertar 19, vUsu, "Cierre: " & DateAdd("d", 1, vParam.fechafin)
        
    
    
        Unload Me
    Else
        cmdDescerrar.Enabled = False
        CmdCancel(5).Enabled = False
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
    
    
    
    Sql = "Deberia hacer una copia de seguridad." & vbCrLf & vbCrLf
    Sql = Sql & "� Desea continuar igualmente ?" & vbCrLf
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
    'BLOQUEAMOS LA BD
    If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    
    Ok = True
    
    
        
    If Ok Then
        'Hemos bloqueado la tbla y esta preparado para la renumeraci�n. No hay nadie trabajando, ni lo va a haber
        Screen.MousePointer = vbHourglass
        'LOG
        If Option1(0).Value Then
            Sql = "actual "
        Else
            Sql = "siguiente "
        End If
        
        Sql = "Ejercicio " & Sql & Format(vParam.fechaini, "dd/mm/yyyy") & " - " & Format(vParam.fechafin, "dd/mm/yyyy")
        
        vLog.Insertar 17, vUsu, Sql
        Sql = ""
        
        
        pb1.Visible = True
        Label2.Caption = ""
        'Ocultanmos el del fondo , para que no pegue pantallazos
        Me.Hide
        frmPpal.Hide
        Me.Show
        
        
        'Renumeramos aqui dentro
        RenumerarAsientos
        
        'Volvemos a mostrar
        Me.Hide
        frmPpal.Show
        
        
        
        
        
        pb1.Visible = False
        Screen.MousePointer = vbDefault
    End If
    
    Bloquear_DesbloquearBD False
    If Ok Then Unload Me
End Sub

Private Sub cmdSimula_Click(Index As Integer)
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

    For I = 0 To 3
        If txtDiario(I).Text = "" Then
           If I < 2 Or (I = 3 And vParam.GranEmpresa) Then
                MsgBox "Seleccione el diario.", vbExclamation
                Ok = False
                Exit For
            End If
        End If
    Next I
    If Not Ok Then Exit Sub
    
    If vParam.NuevoPlanContable And vParam.GranEmpresa Then
        If Not ComprobarCierreCuentas8y9 Then Exit Sub
    End If
    
    Ok = True
    ImporteTotal = 0
    If Ok Then
        pb3.Value = 0
        pb3.Visible = True
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
            pb3.Visible = True
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
        pb3.Visible = True
        Label10.Caption = "Cierre"
        Label10.Refresh
        Label3.Caption = ""
        Label3.Refresh
        'Hacer el cierre
        Ok = SimulaCierreApertura
    End If
    
    pb3.Visible = False
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
    Sql = "Select  numasien Asiento, fechaent Fecha, numdiari Diario, codmacta Cuenta, nommacta Descripci�n, timported Debe, timporteh Haber "
    Sql = Sql & " FROM tmphistoapu "
    
    If cadselect <> "" Then Sql = Sql & " WHERE codusu = " & vUsu.Codigo
    
    Sql = Sql & " ORDER BY 1,2,3,4"
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
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
    

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Public Sub InicializarVbles(A�adireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If A�adireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub






Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False
    DoEvents
    CmdCancel(Opcion).Cancel = True
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

Me.Icon = frmPpal.Icon

Limpiar Me
PrimeraVez = True
Me.fRenumeracion.Visible = False
Me.fPyG.Visible = False
Me.frameDescierre.Visible = False
Select Case Opcion
Case 0
    Me.fRenumeracion.Visible = True
    H = fRenumeracion.Height
    W = fRenumeracion.Width
    Label2.Caption = ""
    pb1.Visible = False
    Caption = "Renumeraci�n de asientos"
    
    
    
    IdPrograma = 1301
    ' La Ayuda
    With Me.ToolbarAyuda(0)
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
Case 1, 4
    fPyG.Visible = True
    H = fPyG.Height + 220 '+ 120
    W = fPyG.Width
    Label3.Caption = ""
    Label10.Caption = ""
    pb3.Visible = False
    
    Me.Option2(0).Value = True
    Option2_Click (0)
'    PonerGrandesEmpresas
    
    IdPrograma = 1303
    ' La Ayuda
    With Me.ToolbarAyuda(1)
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 26
    End With
        
Case 5
    frameDescierre.Visible = True
    H = frameDescierre.Height + 400
    W = frameDescierre.Width
    Caption = "Deshacer cierre"
    Label19.Caption = ""
    Text3.Text = "Ejercicio actual: " & Format(vParam.fechaini, "dd/mm/yyyy") & " - " & Format(vParam.fechafin, "dd/mm/yyyy")


    IdPrograma = 1304
    ' La Ayuda
    With Me.ToolbarAyuda(2)
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 26
    End With

End Select


Me.Height = H + 100
Me.Width = W + 100
End Sub



Private Function CadenaFechasActuralSiguiente(Actual As Boolean) As String
Dim Sql As String
    If Actual Then
        'ACTUAL
        Sql = "fechaent >='" & Format(vParam.fechaini, FormatoFecha) & "' AND "
        Sql = Sql & "fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
    Else
        'SIGUIENTE
        cad = Format(DateAdd("yyyy", 1, vParam.fechaini), FormatoFecha)
        Sql = "fechaent >='" & cad & "' AND "
        cad = Format(DateAdd("yyyy", 1, vParam.fechafin), FormatoFecha)
        Sql = Sql & "fechaent <='" & cad & "'"
    End If
    CadenaFechasActuralSiguiente = Sql
End Function

Private Sub RenumerarAsientos()
Dim ContAsientos As Long
Dim NumeroAntiguo As Long
Dim Fec As String
Dim RA As Recordset


    Set RA = New ADODB.Recordset
    
   
    'obtner el maximo
    cad = CadenaFechasActuralSiguiente(Option1(0).Value)
    Sql = "Select max(numasien) from hcabapu where " & cad
    RA.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    MaxAsiento = 0
    If Not RA.EOF Then
        MaxAsiento = DBLet(RA.Fields(0), "N")
    End If
    RA.Close


    

    'Obtener contador
    Sql = "Select count(numasien) from hcabapu where " & cad
    RA.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ContAsientos = 0
    If Not RA.EOF Then
        ContAsientos = DBLet(RA.Fields(0), "N")
    End If
    RA.Close


    
    If MaxAsiento + ContAsientos > 99999999 Then
        MsgBox "La aplicaci�n no tiene espacio suficiente para renumerar. Numero registros posibles mayor que la capacidad disponible.", vbCritical
        Exit Sub
    End If

    
    
    
    'Para la progresbar
    NumeroRegistros = ContAsientos

    
    'Tendremos el incremeto
    MaxAsiento = MaxAsiento + ContAsientos + 1
    
    
    Label2.Caption = "Preparaci�n"
    Me.Refresh
    
    'actualizamos todas las tablas sumandole maxasiento al numero de asiento donode proceda
    'es decir en el ejercicio y si tiene asiento
    PreparacionAsientos MaxAsiento

    
    DoEvents
    Me.Refresh
    espera 0.01
    
    '-----------------------------------------------------------------
    ' Ahora iremos cogiendo cada registro y los iremos actualizando con
    ' los nuevos valores de numasien, tb para las tblas relacionadas
    ' Solo cambia NUMASIEN
    cad = CadenaFechasActuralSiguiente(Option1(0).Value)
    Sql = "Select numasien,fechaent,numdiari from hcabapu where " & cad
    Sql = Sql & " ORDER BY fechaent,numasien"
    RA.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
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
                End
           End If
           
           
           'progressbar
           Label2.Caption = NumeroAntiguo & " / " & RA.Fields(1)
           Label2.Refresh
           I = Int((MaxAsiento / NumeroRegistros) * pb1.Max)
           pb1.Value = I
           DoEvents
           
         
           'Siguiente
           MaxAsiento = MaxAsiento + 1
           ContAsientos = ContAsientos + 1
           RA.MoveNext
           
           If (MaxAsiento Mod 50) = 0 Then
                DoEvents
                Me.Refresh
                espera 0.01
            End If
           
    Wend
    RA.Close
    Set RA = Nothing
    
    
    
    'En contadores ponemos el contador al numero k le toca
    MaxAsiento = NumeroRegistros
    Sql = "UPDATE contadores set "
    If (Option1(0).Value) Then
        Sql = Sql & " contado1=" & MaxAsiento
    Else
        If MaxAsiento = 1 Then MaxAsiento = 2
        Sql = Sql & " contado2=" & MaxAsiento
    End If
    Sql = Sql & " WHERE TipoRegi = '0'"
    Conn.Execute Sql

End Sub



Private Function CambiaNumeroAsiento(Antiguo As Long, Nuevo As Long, Fecha As String, NuDi As Integer) As Boolean

On Error GoTo ECambia
    CambiaNumeroAsiento = False
    
    'AUX
    cad = " SET numasien = " & Nuevo & " WHERE numasien = " & Antiguo
    cad = cad & " AND fechaent = " & Fecha & " AND numdiari = " & NuDi
    
   
    'Actualizamos el registro de facturas
    Sql = "UPDATE factcli" & cad
    Conn.Execute Sql
    
    
    'Actualizamos el registro de facturas
    Sql = "UPDATE factpro" & cad
    Conn.Execute Sql
    
    'lineas
    Sql = "UPDATE hlinapu" & cad
    Conn.Execute Sql
    
    'lineas de ficheros
    Sql = "UPDATE hcabapu_fichdocs" & cad
    Conn.Execute Sql
    
    'cabeceras
    Sql = "UPDATE hcabapu" & cad
    Conn.Execute Sql
    
    'hco de liquidaciones
    Sql = "UPDATE liqiva" & cad
    Conn.Execute Sql
    
    CambiaNumeroAsiento = True
    Exit Function
ECambia:
    MuestraError Err.Number, "Renumeracion tipo 1, asiento: " & Antiguo

End Function


Private Sub PreparacionAsientos(Suma As Long)
On Error GoTo EPreparacionAsientos

    Sql = CadenaFechasActuralSiguiente(Option1(0).Value)
    cad = " Set NumASien = NumASien + " & Suma
    cad = cad & " WHERE numasien>0 AND " & Sql
    
    pb1.Max = 6
    
    'Facturas clientes
    Label2.Caption = "Facturas clientes"
    Label2.Refresh
    pb1.Value = 1
    Sql = "UPDATE factcli" & cad
    Conn.Execute Sql
    
    'Facturas proveedores
    Label2.Caption = "Facturas proveedores"
    Label2.Refresh
    pb1.Value = 2
    Sql = "UPDATE factpro" & cad
    Conn.Execute Sql
    
    'Lineas hco asiento
    Label2.Caption = "Lineas asientos"
    Label2.Refresh
    pb1.Value = 3
    Sql = CadenaFechasActuralSiguiente(Option1(0).Value)
    Sql = "Select distinct(numasien) from hlinapu WHERE " & Sql
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql = CadenaFechasActuralSiguiente(Option1(0).Value)
    cad = " Set NumASien = NumASien + " & Suma
    cad = cad & " WHERE numasien>0 AND " & Sql
    
    'Ejecutaremos esto
    Sql = "UPDATE hlinapu " & cad
    Conn.Execute Sql
    
    'ASientos
    Label2.Caption = "Cabeceras asientos"
    Label2.Refresh
    pb1.Value = 4
    Sql = "UPDATE hcabapu " & cad
    Conn.Execute Sql
    
    'asientos ficheros
    Label2.Caption = "L�neas asientos ficheros"
    Label2.Refresh
    pb1.Value = 5
    Sql = "UPDATE hcabapu_fichdocs " & cad
    Conn.Execute Sql
    
    'Liquidaciones de Iva
    Label2.Caption = "Liquidaciones IVA"
    Label2.Refresh
    pb1.Value = 6
    Sql = "UPDATE liqiva " & cad
    Conn.Execute Sql
    

    pb1.Max = 1000
    Exit Sub
EPreparacionAsientos:
    MuestraError Err.Number
    MsgBox "Error grave. Soporte t�cnico", vbExclamation
    Set Rs = Nothing
End Sub



Private Sub PonerDatosPyG()

    'Fecha siempre la del final de ejercicio
    Text1(0).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    '8y9
    Text1(13).Text = Text1(0).Text


    NumeroRegistros = 1
    Sql = DevuelveDesdeBD("contado1", "contadores", "tiporegi", "0", "T")
    If Sql = "" Then
        MsgBox "Error obteniendo numero de asiento."
        cmdCierreEjercicio.Enabled = False
    Else
        Text1(3).Text = Val(Sql) + 1
    End If
    If vParam.GranEmpresa Then Text1(11).Text = Val(Text1(3).Text) + 1
    
    'PyG
    Sql = CuentaCorrectaUltimoNivel(vParam.ctaperga, cad)
    If Sql = "" Then
        MsgBox "Error en la cuenta de p�rdidas y ganancias de parametros.", vbExclamation
    Else
        Text1(1).Text = vParam.ctaperga
        Text2(0).Text = cad
    End If
    
    'Concepto  --> Siempre sera nuestro 960
    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "960")
    If Sql = "" Then
        MsgBox "No existe el concepto 960.", vbExclamation
    Else
        Text1(2).Text = "960"
        Text2(1).Text = Sql
    End If
    
    'Concepto para grandes empresas
    If vParam.GranEmpresa Then
        Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "961")
        If Sql = "" Then
            MsgBox "No existe el concepto 961.", vbExclamation
        Else
            Text1(12).Text = "961"
            Text2(4).Text = Sql
        End If
    End If
    

    
    'Simulacion nos salimos ya
    If Opcion = 4 Then Exit Sub
    
    
    'Si ya hay un 960 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Sql = "Select numasien from hlinapu WHERE codconce=960 and fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        MsgBox "Ya se ha efectuado el asiento de Perdidas y ganancias : " & Rs.Fields(0), vbExclamation
        cmdCierreEjercicio.Enabled = False
    End If
    Rs.Close
    
    
    'Si ya hay un 961 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    If vParam.GranEmpresa Then
        Sql = "Select numasien from hlinapu WHERE codconce=961 and fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            MsgBox "Ya se ha efectuado el asiento de regularizacion : " & Rs.Fields(0), vbExclamation
            cmdCierreEjercicio.Enabled = False
        End If
        Rs.Close
    End If
    
    
    'Comprobamos k tampoc haya asiento 1 en ejercicio siguiente
    Sql = "Select numasien from hcabapu WHERE fechaent>'" & Format(vParam.fechafin, FormatoFecha)
    Sql = Sql & "' and numasien=1"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        MsgBox "Ya existe el asiento numero 1 para el a�o siguiente.", vbExclamation
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
    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "980")
    If Sql = "" Then
        MsgBox "No existe el concepto 980.", vbExclamation
    Else
        Text1(5).Text = "980"
        Text2(2).Text = Sql
    End If
    
    'N� asiento apertura.---- El uno
    Text1(6).Text = 1
    
    'Concepto  --> Siempre sera nuestro 970  apertura
    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", "970")
    If Sql = "" Then
        MsgBox "No existe el concepto 970.", vbExclamation
    Else
        Text1(8).Text = "970"
        Text2(3).Text = Sql
    End If
    
    
    
    'Si es simulacion busco el numero de diario mas peque�o
    If Opcion = 4 Then
        Sql = "Select numdiari,desdiari from tiposdiario order by numdiari"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            For I = 0 To 3
                txtDescDiario(I).Text = Rs.Fields(1)
                txtDiario(I).Text = Rs.Fields(0)
            Next I
        End If
        Rs.Close
        Set Rs = Nothing
        'Ponemos el control en simula
        If Not PrimeraVez Then cmdSimula(1).SetFocus
    End If
    
    
    
    'Simulacion nos salimos ya
    If Opcion = 4 Then Exit Sub
    
    
    'Si ya hay un 980 y/o 970 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    cad = CadenaFechasActuralSiguiente(True)
    Sql = "Select numasien from hlinapu WHERE codconce=980"
    Sql = Sql & " AND " & cad
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        MsgBox "Ya se ha efectuado el asiento de cierre de ejercicio : " & Rs.Fields(0), vbExclamation
        Ok = False
    Else
        Ok = True
    End If
    Rs.Close
    
    'Apertura
    If Ok Then
            cad = CadenaFechasActuralSiguiente(False)
            Sql = "Select numasien from hlinapu WHERE codconce=980"
            Sql = Sql & " AND " & cad
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    txtDiario(I).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescDiario(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
    I = Index
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

    cmdSimula(0).Visible = (Opcion = 4)
    cmdSimula(1).Visible = (Opcion = 4)
    Me.CmdCancel(4).Visible = (Opcion = 4)
    Me.cmdCierreEjercicio.Visible = (Opcion = 1)
    Me.CmdCancel(1).Visible = (Opcion = 1)
    Label6.Visible = (Opcion = 4)
    Label4.Visible = Not Label6.Visible
    If Opcion = 1 Then
        Me.Caption = "Cierre de ejercicio"  '"Generar asiento de p�rdidas y ganancias y el cierre de ejercicio"
    Else
        Me.Caption = "Simulaci�n cierre"
    End If

    If TieneInmovilizado Then
        Me.FrameAmort.Visible = (Opcion = 1)
        
        vFecha = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", 1, "N")
        Text1(15).Text = Format(vFecha, "dd/mm/yyyy")
    End If
    

'    Caption = "Cierre de Ejercicio"

    PonerGrandesEmpresas
        
    If Opcion = 1 Then
        For I = 0 To txtDiario.Count - 1
            txtDiario(I).Text = ""
        Next I
        
        Sql = "select numdiari, desdiari from tiposdiario "
        If TotalRegistrosConsulta(Sql) = 1 Then
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                For I = 0 To txtDiario.Count - 1
                    txtDiario(I).Text = DBLet(Rs.Fields(0), "N")
                    txtDescDiario(I).Text = DBLet(Rs.Fields(1), "T")
                Next I
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
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
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
        MsgBox "Diario debe ser num�rico: " & txtDiario(Index).Text, vbExclamation
        txtDiario(Index).Text = ""
        txtDiario(Index).SetFocus
        Exit Sub
    End If
    Sql = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text)
    If Sql = "" Then
        MsgBox "No existe el diario : " & txtDiario(Index).Text, vbExclamation
        txtDiario(Index).Text = ""
        txtDiario(Index).SetFocus
    Else
        Me.txtDescDiario(Index).Text = Sql
    End If
    
    
End Sub




Private Function ASientoPyG() As Boolean
Dim Ok As Boolean
Dim NoTieneLineas As Boolean
Dim Cuantos As Long
    
    On Error GoTo EASientoPyG
    
    ASientoPyG = False

    'Generamos los  apuntes, sobre hcabapu, y luego los actualizamos
    If Opcion <> 4 Then
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(0).Text
        cad = Sql & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(3).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Asiento P�rdidas y Ganancias')"
        
    Else
        'Estamo simulando
        'Borramos los datos de tmp
        cad = "Delete from tmphistoapu where codusu = " & vUsu.Codigo
    End If
        
    Conn.Execute cad
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
        
        Sql = "numdiari=" & txtDiario(0).Text & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien  "
        espera 0.2
        Sql = DevuelveDesdeBD("count(*)", "hlinapu", Sql, Text1(3).Text)
        If Sql = "" Then Sql = "0"
        NoTieneLineas = Val(Sql) = 0
        
        
        If NoTieneLineas Then
            'NO HA INSERTADO NI UNA SOLA LINEA EN  LINAPU. CON LO CUAL TENDREMOS QUE AVISAR Y BORRAR
            If MsgBox("Ning�n apunte en lineas de perdidas y ganancias. �Continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then
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
            Sql = " where numdiari=" & txtDiario(0).Text
            Sql = Sql & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
    
            'Borramos por si acaso ha insertado lineas
            cad = "Delete FROM hcabapu" & Sql
            Conn.Execute cad
        End If
    End If
    
    
    If Opcion = 4 Then
        ASientoPyG = Ok
        Exit Function
    End If
    
    If Not Ok Then
        'Comun lineas y cabeceras
        Sql = " where numdiari=" & txtDiario(0).Text
        Sql = Sql & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
        
        'Borramos por si acaso ha insertado lineas
        cad = "Delete FROM hlinapu" & Sql
        Conn.Execute cad
    
        'Borramos la cabcecera del apunte
        cad = "DELETE FROM hcabapu" & Sql
        Conn.Execute cad
        
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
    If vParam.grupogto <> "" Then
        If Not Subgrupo(vParam.grupogto, "") Then Exit Function
    End If
    
    If vParam.grupovta <> "" Then
        If Not Subgrupo(vParam.grupovta, "") Then Exit Function
    End If
        
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
    If vParam.grupoord <> "" And Text1(10).Text <> "" Then
        If Not Subgrupo(vParam.grupoord, Text1(10).Text) Then Exit Function
    End If
    

        
    Set Rs = Nothing
    Cuentas9 = True
    Exit Function
ECuentas9:
    MuestraError Err.Number, "Cuentas   ventas / gastos "
End Function




Private Function Subgrupo(Primera As String, Excepcion As String) As Boolean
Dim CONT As Integer
Dim Importe As Currency
Dim AUX3 As String
Dim vCta As String

    Subgrupo = False
    'ATENCION A�OS PARTIDOS
    cad = Mid("__________", 1, vEmpresa.DigitosUltimoNivel - 1)
    AUX3 = Primera & cad
    cad = " from hlinapu" ' antes hsaldos
    
    'Necesito sbaer tb nombre cta
    If Opcion = 4 Then cad = cad & ",cuentas"
    
    cad = cad & " WHERE "
    If Opcion = 4 Then cad = cad & " cuentas.codmacta = hlinapu.codmacta AND "
    'Por la ambiguedad del nombre
    vCta = " ("
    If Opcion = 4 Then vCta = vCta & " cuentas."
    vCta = vCta & "codmacta like '" & AUX3 & "')"
    If Excepcion <> "" Then
        Excepcion = Mid(Excepcion & "__________", 1, vEmpresa.DigitosUltimoNivel)
        vCta = vCta & " and not ("
        If Opcion = 4 Then vCta = vCta & " cuentas."
        vCta = vCta & "codmacta like '" & Excepcion & "')"
    End If
    
    
    cad = cad & vCta
    
    
    cad = cad & " and fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    
    If Opcion = 4 Then
        cad = "cuentas.codmacta,nommacta " & cad & " GROUP BY cuentas.codmacta ORDER BY cuentas.codmacta"
    Else
        cad = " codmacta " & cad & " GROUP BY codmacta ORDER BY codmacta"
    End If
    
    'Contador
    Sql = "Select sum(coalesce(timported,0))-sum(coalesce(timporteh,0)), " & cad
    Rs.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
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
        I = Int((CONT / NumeroRegistros) * pb1.Max)
        pb1.Value = I
        Importe = Rs.Fields(0)
        If Importe <> 0 Then
            If Opcion = 4 Then
                cad = Sql & "," & MaxAsiento + CONT & ",'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs.Fields(2)) & "','1','" & Text2(1).Text & "',"
            Else
                cad = Sql & "," & MaxAsiento + CONT & ",'" & Rs.Fields(1) & "','',960,'" & Text2(1).Text & "',"
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
        Sql = "INSERT INTO tmphistoapu (codusu, numdiari, desdiari, fechaent, numasien,"
        Sql = Sql & "linliapu, codmacta, nommacta, numdocum, ampconce, timporteD, "
        Sql = Sql & "codccost,timporteH) VALUES ("
        Sql = Sql & vUsu.Codigo & ","
        Sql = Sql & Diario & ",'" & Me.txtDescDiario(1).Text & "'"
        Sql = Sql & ",'" & Format(Fecha, FormatoFecha) & "'," & Num
    Else
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
        Sql = Sql & "codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada) VALUES ("
        Sql = Sql & Diario
        Sql = Sql & ",'" & Format(Fecha, FormatoFecha) & "'," & Num
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
            
            cad = cad & Aux & ")"
        Else
            cad = cad & Aux & "," & Ctrapar & ",'CONTAB',0)"
        End If
        'Ejecutamos
        Conn.Execute cad
    
End Function


'//////////////////////////////////////////////////////////////////7
'///  Cuadramos el asiento de perdidas y ganancias
'///
Private Function CuadrarAsiento() As Boolean
Dim Importe As Currency

On Error GoTo ECua
    CuadrarAsiento = False
    
    If Opcion < 4 Then
        Sql = "select sum(timporteD),  sum(timporteH) from hlinapu "
        Sql = Sql & " where numdiari=" & txtDiario(0).Text
        Sql = Sql & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(3).Text & ""
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        Importe = 0
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then Importe = Rs.Fields(0)
            If Not IsNull(Rs.Fields(1)) Then Importe = Importe - Rs.Fields(1)
            
        End If
        Rs.Close
        Set Rs = Nothing
        
        If Importe <> 0 Then
            CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
            cad = Sql & "," & MaxAsiento + 1 & ",'" & Text1(1).Text & "','',960,'" & Text2(1).Text & "',"
            InsertarLineasDeAsientos Importe, "NULL"
        End If
    Else
        
        'Simulaci�n
        ' "linliapu, codmacta, nommacta, numdocum, ampconce "
        ImporteTotal = ImporteTotal * -1 'Para k cuadre
        ImportePyG = ImporteTotal
        CadenaLINAPU txtDiario(0).Text, vParam.fechafin, Text1(3).Text
        cad = Sql & "," & MaxAsiento + 1 & ",'" & Text1(1).Text & "','"
        ' lo meto en el uno, NO en maxasiento
        cad = Sql & ",1,'" & Text1(1).Text & "','"
        cad = cad & Text2(0).Text & "','1','" & Text2(1).Text & "',"
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
    Conn.Execute "Delete from tmpCierre"  ' no hace falta codusu pq solo puede haber trabajando uno a la vez
    
    Label10.Caption = "Leyendo datos"
    Label10.Refresh
    If Not GeneraTmpCierre Then Exit Function
    
    'Esta grabado el fichero tmpCierre con los importes
    'Fijamos la pb3
    Sql = "Select count(*) from tmpcierre where importe<>0"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        Sql = "DELETE FROM hcabapu where fechaent = '" & Format(vParam.fechafin, FormatoFecha) & "' AND Numasien = " & Text1(4).Text
        Conn.Execute Sql
        Sql = "DELETE FROM hlinapu where fechaent = '" & Format(vParam.fechafin, FormatoFecha) & "' AND Numasien = " & Text1(4).Text
        Conn.Execute Sql

        'BORRAMOS APERTURA
        Sql = "DELETE FROM hcabapu where fechaent = '" & Format(Text1(9).Text, FormatoFecha) & "' AND Numasien = 1"
        Conn.Execute Sql
        Sql = "DELETE FROM hlinapu where fechaent = '" & Format(Text1(9).Text, FormatoFecha) & "' AND Numasien = 1"
        Conn.Execute Sql
        
    Else
        Me.cmdCierreEjercicio.Enabled = False
        
        'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
        cad = Format(DateAdd("yyyy", 1, vParam.fechaini), FormatoFecha)
        Sql = "UPDATE parametros SET fechaini= '" & cad
        
        vParam.fechafin = DateAdd("yyyy", 1, vParam.fechafin)
        vParam.FechaActiva = DateAdd("yyyy", 1, vParam.FechaActiva)
        If vParam.FechaActiva >= vParam.fechafin Then vParam.FechaActiva = vParam.fechaini
        If vParam.FechaActiva < DateAdd("yyyy", 1, vParam.fechaini) Then vParam.FechaActiva = DateAdd("yyyy", 1, vParam.fechaini)
        cad = Format(vParam.fechafin, FormatoFecha)
        Sql = Sql & "' , fechafin='" & cad & "'"
        cad = Format(vParam.FechaActiva, FormatoFecha)
        Sql = Sql & " , fechaactiva='" & cad & "'"
        
        Sql = Sql & " WHERE fechaini='" & Format(vParam.fechaini, FormatoFecha) & "'"
        
        'ANTES
        'Conn.Execute SQL
        If Not EjecutaSQL(Sql) Then MsgBox "Se ha producido un error insertanto contadores en HCO." & vbCrLf & "Cuando finalice avise a soporte t�cnico de Ariadna Software.", vbExclamation
            
        vParam.fechaini = DateAdd("yyyy", 1, vParam.fechaini)
       
        'los contadores
        'UPDATEAMOS LOS CONTADORES
        'con los nuevos valores
        'Es decir Contsiguiente pasa a actual, y en siguiente ponemos un 2, puesto k el 1 lo reservamos para apertura
        Sql = "UPDATE contadores SET contado1 =  contado2"
        Conn.Execute Sql
        
        'Ponemos en todos un 0
        Sql = "UPDATE contadores SET contado2 = 0"
        Conn.Execute Sql
        
        'Menos en asientos k podremos un 1, ya que se reservara para el cierre
        'del a�o siguiente
        Sql = "UPDATE contadores SET contado2 = 1 WHERE tiporegi='0'"
        Conn.Execute Sql
        
    End If
    
    HacerElCierre = True
    
EHacerElCierre:
    If Err.Number <> 0 Then MuestraError Err.Number
    vParam.Leer
    If vEmpresa.TieneTesoreria Then vParamT.Leer
    Set Rs = Nothing
End Function

Private Function SimulaCierreApertura() As Boolean


    SimulaCierreApertura = False
    Set Rs = New ADODB.Recordset
    Conn.Execute "Delete from tmpCierre"  ' no hace falta codusu pq solo puede haber trabajando uno a la vez
    
    Label10.Caption = "Leyendo datos"
    Label10.Refresh
    If Not GeneraTmpCierre Then Exit Function
    
    'Esta grabado el fichero tmpCierre con los importes
    'Fijamos la pb3
    Sql = "Select count(*) from tmpcierre where importe<>0"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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


    GeneraTmpCierre = False
    
    cad = vSql & " fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    cad = " from hlinapu where " & cad & " GROUP BY codmacta ORDER BY codmacta"
    
    
    
    'Contador
    Sql = "Select -1 * (sum(coalesce(timported,0))-sum(coalesce(timporteh,0))),codmacta " & cad

    Sql = "INSERT INTO tmpCierre " & Sql
    Conn.Execute Sql
    
    
    'Y si es simulacion, enonces borramos las cuentas de perdidas ganancias
    If Opcion = 4 Then
        Sql = "Delete from tmpcierre where cta like '" & vParam.grupogto & "%'"
        Sql = Sql & " OR  cta like '" & vParam.grupovta & "%'"
        Conn.Execute Sql
        
            
        If vParam.grupoord <> "" Then
            Sql = "Delete from tmpcierre where cta like '" & vParam.grupoord & "%'"
            'Excepcion
            If Text1(10).Text <> "" Then Sql = Sql & " AND not (cta like '" & Text1(10).Text & "%')"
            Conn.Execute Sql
        End If
            
        'Si es gran empresa me cargo tb las 8% y 9%
        If vParam.NuevoPlanContable And vParam.GranEmpresa Then
            Sql = "Delete from tmpcierre where cta like '8%' OR  cta like '9%'"
            Conn.Execute Sql
        End If
        
        'Comprobamos si existe el parametro
        Set miRsAux = New ADODB.Recordset
        Sql = "Select importe from tmpcierre where cta ='" & Text1(1).Text & "'"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        Importe = 0
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then
                Sql = "E"
                Importe = miRsAux.Fields(0)
            End If
        End If
        miRsAux.Close
        
        
        If Sql = "" Then
            'Metemos a la cta 129 de perdias y gancias las perdidas y gancias generadas
            Sql = "INSERT INTO tmpcierre (cta,importe) values ('" & Text1(1).Text & "'," & TransformaComasPuntos(CStr(ImportePyG)) & ")"
        Else
            ImportePyG = ImportePyG + Importe
            Sql = "UPDATE tmpcierre SET Importe= " & TransformaComasPuntos(CStr(ImportePyG)) & " WHERE cta='" & Text1(1).Text & "'"
        End If
        Conn.Execute Sql
        
        
        
        
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
            
            Sql = "select codusu,codmacta,sum(timporteD) as d,sum(timporteH) as h from tmphistoapu "
            Sql = Sql & " where codusu = " & vUsu.Codigo & " and numdocum=2 and codmacta <'8' group by 1,2"
            miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Set L = New Collection
            While Not miRsAux.EOF
                
                Importe = DBLet(miRsAux!d, "N")
                If Not IsNull(miRsAux!H) Then Importe = Importe - miRsAux!H
                Sql = miRsAux!codmacta & "|" & CStr(Importe) & "|"
                L.Add Sql
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            'Ahora ya tengo todas las cuentas a actualizar/crear
            For I = 1 To L.Count
                cad = RecuperaValor(L.Item(I), 1)
                ImporteTotal = CCur(RecuperaValor(L.Item(I), 2))
                Sql = "Select importe from tmpcierre where cta ='" & cad & "'"
                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Sql = ""
                Importe = 0
                If Not miRsAux.EOF Then
                    If Not IsNull(miRsAux.Fields(0)) Then
                        Sql = "E"
                        Importe = miRsAux.Fields(0)
                    End If
                End If
                miRsAux.Close
                
                
                If Sql = "" Then
                    'Metemos a la cta 129 de perdias y gancias las perdidas y gancias generadas
                    Sql = "INSERT INTO tmpcierre (cta,importe) values ('" & cad & "'," & TransformaComasPuntos(CStr(ImportePyG)) & ")"
                Else
                    ImporteTotal = ImporteTotal + Importe
                    Sql = "UPDATE tmpcierre SET Importe= " & TransformaComasPuntos(CStr(ImporteTotal)) & " WHERE cta='" & cad & "'"
                End If
                Conn.Execute Sql
            Next I
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

    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien,  obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(1).Text
    cad = Sql & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(4).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Generaci�n Asiento Cierre')"
    Conn.Execute cad

    
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
    Rs.Open "SELECT * from tmpCierre WHERE importe <>0 ORDER By Cta", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    CONT = 1
    
    ' hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
    ' codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada
    
    
    While Not Rs.EOF
        cad = Sql & "," & CONT & ",'" & Rs.Fields(1) & "','',980,'" & Text2(2).Text & "',"
        Importe = Rs.Fields(0)
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        cad = cad & Aux
        cad = cad & ",NULL,'contab',0)"
        
        I = Int(((CONT + MaxAsiento) / NumeroRegistros) * pb3.Max)
        pb3.Value = I
        
        Conn.Execute cad
        
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

    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(2).Text
    cad = Sql & ",'" & Format(Text1(9).Text, FormatoFecha) & "'," & Text1(6).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Generaci�n Asiento Apertura')"
    Conn.Execute cad
    
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
    Rs.Open "SELECT * from tmpCierre WHERE importe <>0 ORDER BY Cta ", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    CONT = 1
    
    ' hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum,"
    ' codconce, ampconce, timporteD, codccost, timporteH, ctacontr, idcontab, punteada
    
    
    While Not Rs.EOF
        cad = Sql & "," & CONT & ",'" & Rs.Fields(1) & "','',970,'" & Text2(3).Text & "',"
        Importe = Rs.Fields(0)
        Importe = Importe * -1
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        cad = cad & Aux
        cad = cad & ",NULL,'contab',0)"
        
        I = Int(((CONT + MaxAsiento) / NumeroRegistros) * pb3.Max)
        pb3.Value = I
        
        Conn.Execute cad
        
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
    
    Sql = "Select numasien from hlinapu WHERE codconce=980 AND fechaent='" & Format(FechaCierre, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
    cad = "select tmpcierre.*,nommacta from tmpcierre,cuentas where tmpcierre.cta=cuentas.codmacta"
    cad = cad & " AND importe <>0 ORDER By Cta"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Cont = 1
    CONT = 2
    'linliapu, codmacta, nommacta, numdocum, ampconce, timporteD, "
    'codccost,timporteH) VALUES ("
    
    While Not Rs.EOF
        
        Importe = Rs.Fields(0)
        If EsCierre Then
            cad = Sql & "," & CONT & ",'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs!Nommacta) & "','3','" & Text2(2).Text & "',"
        Else
            cad = Sql & "," & CONT & ",'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs!Nommacta) & "','4','" & Text2(3).Text & "',"
            Importe = Importe * -1
        End If
        
        If Importe < 0 Then
            Aux = "NULL,NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
        Else
            Aux = TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL"
        End If
        cad = cad & Aux & ")"
    
        I = Int(((CONT + MaxAsiento) / (NumeroRegistros + 3)) * pb3.Max)
        pb3.Value = I
        
        Conn.Execute cad
        
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

'Tiene k existe el asiento de apertura del a�o siguient
    cad = CadenaFechasActuralSiguiente(True)
    Sql = "Select numasien from hlinapu WHERE codconce=970"
    Sql = Sql & " AND " & cad
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ExisteAsientosDescerrar = True
    Else
        ExisteAsientosDescerrar = False
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Private Function ExistenApuntesEjercicioAnterior() As Boolean

'Tiene k existe el asiento de apertura del a�o siguient
    Sql = "Select count(*) from hlinapu WHERE "
    Sql = Sql & " Fechaent < '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ExistenApuntesEjercicioAnterior = False
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then ExistenApuntesEjercicioAnterior = True
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Private Function ExistenApuntesEjercicioSiguiente() As Boolean

'Tiene k existe el asiento de apertura del a�o siguient
    Sql = "Select count(*) from hlinapu WHERE "
    Sql = Sql & " Fechaent > '" & Format(vParam.fechafin, FormatoFecha) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
Dim Sql1 As String
Dim Sql2 As String

On Error GoTo EHacerDescierre
    HacerDescierre = False
    Screen.MousePointer = vbHourglass
    
    'Si ya hay un 960 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    'P y G
    MaxAsien = 0
    Label19.Caption = "Perdidas y ganancias"
    Label19.Refresh

    ' cabecera
    Sql = "DELETE FROM hcabapu where (numdiari, fechaent, numasien) in ("
    Sql = Sql & "select numdiari, fechaent, numasien from hlinapu WHERE codconce=960 "
    cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"
    Sql = Sql & ")"
    Conn.Execute Sql

    ' lineas
    Sql = "DELETE FROM hlinapu WHERE codconce=960 "
    cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"

    Conn.Execute Sql


    Me.Refresh
    espera 0.5
    Me.Refresh


    'Compruebo si hay REGULARIZACON 8 y 9
    If vParam.GranEmpresa Then
        Label19.Caption = "Regularizaci�n 8 y 9"
        Label19.Refresh
    End If
    ' cabecera
    Sql = "DELETE FROM hcabapu where (numdiari, fechaent, numasien) in ("
    Sql = Sql & " select numdiari, fechaent, numasien FROM hlinapu WHERE codconce=961 "
    cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"
    Sql = Sql & ") "
    
    Conn.Execute Sql
    
    ' lineas
    Sql = "DELETE FROM hlinapu WHERE codconce=961 "
    cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"
    
    Conn.Execute Sql
    
    Me.Refresh
    espera 0.5
    Me.Refresh



    'Cierre
    'Si hay un 980  en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Label19.Caption = "Cierre"
    Label19.Refresh

    ' cabecera
    Sql = "DELETE FROM hcabapu where (numdiari, fechaent, numasien) in ("
    Sql = Sql & "SELECT numdiari, fechaent, numasien from hlinapu WHERE codconce=980"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"
    Sql = Sql & ")"
    Conn.Execute Sql


    ' lineas
    Sql = "DELETE FROM hlinapu WHERE codconce=980"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechaini))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = CStr(DateAdd("yyyy", -1, vParam.fechafin))
    cad = Format(cad, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"

    Conn.Execute Sql

    Me.Refresh
    espera 0.5
    Me.Refresh
        

    'Si ya hay un  970 en hcabapu, con esa fecha entonces es k ya esta hecho el cierre
    Label19.Caption = "Apertura"
    Label19.Refresh
    'cabecera
    Sql = "DELETE FROM hcabapu WHERE (numdiari, fechaent, numasien) in ("
    Sql = Sql & "SELECT numdiari, fechaent, numasien FROM hlinapu WHERE codconce=970"
    cad = Format(vParam.fechaini, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = Format(vParam.fechafin, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"
    Sql = Sql & ")"
    Conn.Execute Sql
    
    'lineas
    Sql = "DELETE FROM hlinapu WHERE codconce=970"
    cad = Format(vParam.fechaini, FormatoFecha)
    Sql = Sql & " AND fechaent >='" & cad & "'"
    cad = Format(vParam.fechafin, FormatoFecha)
    Sql = Sql & " AND fechaent <='" & cad & "'"
    
    Conn.Execute Sql
    
    Me.Refresh
    espera 0.5
    Me.Refresh


    'Hay k bajar una a�o las fechas de parametros de INICIO y FIN ejerecicio
    'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
    'Reestablecemos las fechas
    'Ahora, en parametros cambias ciertas cosas tales como fechas ejercicio
    Label19.Caption = "Contadores"
    Label19.Refresh
    I = -1
    cad = Format(DateAdd("yyyy", I, vParam.fechaini), FormatoFecha)
    Sql = "UPDATE parametros SET fechaini= '" & cad
    cad = Format(DateAdd("yyyy", I, vParam.fechafin), FormatoFecha)
    Sql = Sql & "' , fechafin='" & cad & "'"
    'Fechaactiva
    cad = Format(DateAdd("yyyy", I, vParam.FechaActiva), FormatoFecha)
    Sql = Sql & " , FechaActiva='" & cad & "'"
    
    Sql = Sql & " WHERE fechaini='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Conn.Execute Sql
    
    vParam.fechaini = DateAdd("yyyy", I, vParam.fechaini)
    vParam.fechafin = DateAdd("yyyy", I, vParam.fechafin)
    'Fecha activa
    vParam.FechaActiva = DateAdd("yyyy", I, vParam.FechaActiva)


    Set Rs = New ADODB.Recordset
    
    Sql = "SELECT tiporegi, nomregis, contado1, contado2 from Contadores order by tiporegi"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""
    
    While Not Rs.EOF
        If DBLet(Rs!tiporegi, "T") = "0" Then ' asientos
            Sql = "select max(numasien) from hlinapu where fechaent between  " & DBSet(vParam.fechaini, "F") & " and  " & DBSet(vParam.fechafin, "F")
            
            Sql1 = "select max(numasien) from hlinapu where fechaent > " & DBSet(vParam.fechafin, "F")
            If DevuelveValor(Sql1) = 0 Then
                Sql1 = "select 1 from hlinapu "
            End If
        Else
            If IsNumeric(DBLet(Rs!tiporegi, "T")) Then ' facturas de proveedor
                Sql = "select max(numregis) from factpro where fecharec between  " & DBSet(vParam.fechaini, "F") & " and  " & DBSet(vParam.fechafin, "F")
                Sql = Sql & " and numserie = " & DBSet(Rs!tiporegi, "T")
                Sql1 = "select max(numregis) from factpro where fecharec > " & DBSet(vParam.fechafin, "F")
                Sql1 = Sql1 & " and numserie = " & DBSet(Rs!tiporegi, "T")
            Else ' facturas de cliente
                Sql = "select max(numfactu) from factcli where fecfactu between  " & DBSet(vParam.fechaini, "F") & " and  " & DBSet(vParam.fechafin, "F")
                Sql = Sql & " and numserie = " & DBSet(Rs!tiporegi, "T")
                
                Sql1 = "select max(numfactu) from factcli where fecfactu > " & DBSet(vParam.fechafin, "F")
                Sql1 = Sql1 & " and numserie = " & DBSet(Rs!tiporegi, "T")
            End If
        End If
        
        'actualizamos
        Sql2 = "update contadores set contado1 = " & DevuelveValor(Sql) & ",contado2 = " & DevuelveValor(Sql1)
        Sql2 = Sql2 & " where tiporegi = " & DBSet(Rs!tiporegi, "T")
        Conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    'El contador de ejercicio actual es, MaxAsien
    If MaxAsien <> 0 Then
        'ESTO NO ES ASI Y NO LO VOY A TOCAR
        'SQL = "UPDATE Contadores SET contado1 = " & MaxAsien - 1 & " WHERE tiporegi ='0'"
        'EjecutaSQL SQL
    End If
    
    HacerDescierre = True
    Exit Function
EHacerDescierre:
    MuestraError Err.Number, "Proc. HacerDescierre"
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
    
    
    GElabel1.Visible = B
    Me.GELine.Visible = B
    Label5(17).Visible = B
    Label5(16).Visible = B
    Label5(16).Visible = B
    Label5(15).Visible = B
    Label5(14).Visible = False 'vEmpresa.GranEmpresa
    Label5(18).Visible = B
    Text2(4).Visible = B
    Me.Image1(3).Visible = B
    Me.txtDiario(3).Visible = B
    Me.txtDescDiario(3).Visible = B
    Text1(13).Visible = B
    Text1(14).Visible = False 'vEmpresa.GranEmpresa
    Text1(11).Visible = B
    Text1(12).Visible = B
    Frame4.Visible = B
    
    If Not B Then
        'SOlo para la antigua
        '---------------------------------------------------------
        'Si tiene el otro grupo de perdidas y ganacias entonces
        'tenemos k solicitar la excepcion a digitos de tercer nivel
        'ofertando el de parametros
        B = (vParam.grupoord <> "") And (vParam.Automocion <> "")
        Label5(13).Visible = B
        Text1(10).Visible = B
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
    
    cad = "select " & vUsu.Codigo & ",codmacta,'T',sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) from hlinapu"
    cad = cad & " WHERE mid(codmacta,1,1) in ('8','9') AND "
    cad = cad & "  fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    
    
    
    cad = cad & " GROUP BY codmacta"
    'Insertamos en tmpcierre
    cad = "INSERT INTO TMPCIERRE1 " & cad
    Conn.Execute cad
    
    'COJEREMOS TODAS LAS CUENTAS 8 y 9 a tres digitos y comprobaremos que en
    'la configuracion tienen puesto en el campo cuentaba
    '
    'Cruzamos tmpcierr1 con codusu = vusu y left join con cuentas
    'Veremos si hay null con lo cual esta mal y si no, updatearemos tmpcierre
    cad = "select cta,nommacta,mid(iban, 15,10) cuentaba from tmpcierre1,cuentas where codusu = " & vUsu.Codigo & " and cta = codmacta"
  
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    I = 0
    While Not Rs.EOF
        If DBLet(Rs!cuentaba, "T") = "" Then
            
            I = I + 1
            cad = cad & "     " & Rs!Cta
            If (I Mod 5) = 0 Then cad = cad & vbCrLf
        
        Else
        
            'ASi, tanto para la simulacion, como para el cierre ya se contra que cuentas saldan las del 8 9
            Conn.Execute "UPDATE tmpcierre1 SET nomcta = '" & Rs!cuentaba & "' WHERE codusu = " & vUsu.Codigo & " and cta = '" & Rs!Cta & "'"
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    If I > 0 Then
        cad = "Cuentas sin configurar el cierre: " & vbCrLf & vbCrLf & cad
        MsgBox cad, vbExclamation
        Set Rs = Nothing
        Exit Function
    End If
    
    
    'OK tiene todas las cuentas configuradas
    cad = "Select tmpcierre1.nomcta, cuentas.codmacta from tmpcierre1 left join cuentas on nomcta = cuentas.codmacta where tmpcierre1.codusu=" & vUsu.Codigo
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    I = 0
    While Not Rs.EOF
        
        If IsNull(Rs!codmacta) Then
            I = I + 1
            cad = "    " & Rs!nomcta
            If (I Mod 5) = 0 Then cad = cad & vbCrLf
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    If I > 0 Then
        cad = "Cuentas de cierre configurada, pero no existen: " & vbCrLf & vbCrLf & cad
        MsgBox cad, vbExclamation
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
        Sql = " where numdiari=" & txtDiario(2).Text
        Sql = Sql & " AND fechaent ='" & Format(vParam.fechafin, FormatoFecha) & "' AND numasien = " & Text1(11).Text & ""
        
        'Borramos por si acaso ha insertado lineas
        cad = "Delete FROM hlinapu" & Sql
        Conn.Execute cad
    
        'Borramos la cabcecera del apunte
        cad = "DELETE FROM hcabapu" & Sql
        Conn.Execute cad
        
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
    
    Sql = "Select count(*) from tmpcierre1 where codusu = " & vUsu.Codigo
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES (" & txtDiario(3).Text
        cad = Sql & ",'" & Format(vParam.fechafin, FormatoFecha) & "'," & Text1(11).Text & ",NULL," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Generaci�n Asiento 8 y 9')"
        
        Conn.Execute cad
    End If
    
    
    
    
    NumeroRegistros = NumeroRegistros + 1 'Para que no desborde
    
    Sql = "Select tmpcierre1.* from tmpcierre1 where  codusu = " & vUsu.Codigo & " ORDER BY nomcta,cta"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        I = Int((CONT / NumeroRegistros) * pb1.Max)
        pb1.Value = I
        
        
        
        '$$$
        cad = Mid(Rs!Cta & "__________", 1, vEmpresa.DigitosUltimoNivel)
        cad = " WHERE hlinapu.codmacta=cuentas.codmacta AND hlinapu.codmacta like '" & cad & "' AND "
        cad = "select hlinapu.codmacta,sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) as miImporte,nommacta from hlinapu,cuentas" & cad
        cad = cad & " fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
        
        cad = cad & " GROUP BY codmacta"
        
        ImpComprobacion = 0
        RT.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RT.EOF
            Importe = RT!miImporte  'importe
            
            If Importe <> 0 Then
                If Opcion = 4 Then
                    ' "linliapu, codmacta, nommacta, numdocum, ampconce "
                    cad = Sql & "," & MaxAsiento & ",'" & RT!codmacta & "','" & DevNombreSQL(RT!Nommacta) & "','2','" & Text2(4).Text & "',"
                Else
                    cad = Sql & "," & MaxAsiento & ",'" & RT!codmacta & "','',961,'" & Text2(4).Text & "',"
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
            cad = "Error obteniendo saldos. " & vbCrLf & "Subgrupo: " & Rs!Cta & vbCrLf
            cad = cad & "Imp 3 digitos: " & Importe & vbCrLf & "Ultimo nivel: " & ImpComprobacion
            If Opcion <> 4 Then cad = cad & vbCrLf & vbCrLf & " No puede continuar con el cierre"
            MsgBox cad, vbExclamation
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
            cad = Sql & "," & MaxAsiento & ",'" & LaCuenta & "','" & DevNombreSQL(C) & "','2','" & Text2(4).Text & "',"
        Else
            cad = Sql & "," & MaxAsiento & ",'" & LaCuenta & "','',961,'" & Text2(4).Text & "',"
        End If
        InsertarLineasDeAsientos ImporteTotal, "NULL"
        MaxAsiento = MaxAsiento + 1
    End If
End Sub




