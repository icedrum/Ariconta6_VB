VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESCobrosPdtesList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "frmTESCobrosPdtesList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   69
      Top             =   6360
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   72
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   71
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
         TabIndex        =   70
         Top             =   720
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenación"
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
      Left            =   7140
      TabIndex        =   51
      Top             =   6360
      Width           =   4785
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   150
         TabIndex        =   67
         Top             =   1110
         Width           =   4455
         Begin VB.OptionButton optVarios 
            Caption         =   "Nombre"
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
            Left            =   2340
            TabIndex        =   23
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optVarios 
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
            Height          =   240
            Index           =   3
            Left            =   510
            TabIndex        =   22
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   150
         TabIndex        =   66
         Top             =   360
         Width           =   4455
         Begin VB.OptionButton optVarios 
            Caption         =   "Tipo Pago"
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
            Left            =   2850
            TabIndex        =   21
            Top             =   270
            Width           =   1545
         End
         Begin VB.OptionButton optVarios 
            Caption         =   "Fecha Vto"
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
            Left            =   1350
            TabIndex        =   20
            Top             =   270
            Width           =   1365
         End
         Begin VB.OptionButton optVarios 
            Caption         =   "Cliente"
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
            Left            =   120
            TabIndex        =   19
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Formato extendido"
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
         Left            =   1980
         TabIndex        =   25
         Top             =   2070
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Resumen"
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
         Left            =   330
         TabIndex        =   24
         Top             =   2070
         Width           =   1335
      End
   End
   Begin VB.Frame FrameConcepto 
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
      Height          =   6345
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNAgente 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   4140
         Width           =   4155
      End
      Begin VB.TextBox txtNAgente 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   4530
         Width           =   4155
      End
      Begin VB.TextBox txtAgente 
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
         Tag             =   "imgAgente"
         Top             =   4140
         Width           =   1275
      End
      Begin VB.TextBox txtAgente 
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
         Tag             =   "imgAgente"
         Top             =   4530
         Width           =   1275
      End
      Begin VB.TextBox txtNDpto 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   3060
         Width           =   4155
      End
      Begin VB.TextBox txtNDpto 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   3450
         Width           =   4155
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
         TabIndex        =   0
         Tag             =   "imgFec"
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
         Index           =   3
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgFec"
         Top             =   1230
         Width           =   1305
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2310
         Width           =   4155
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1920
         Width           =   4155
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   5670
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
         Index           =   0
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   5280
         Width           =   4665
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
         Left            =   1230
         TabIndex        =   5
         Tag             =   "imgCuentas"
         Top             =   2310
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
         Index           =   0
         Left            =   1230
         TabIndex        =   4
         Tag             =   "imgCuentas"
         Top             =   1920
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
         Index           =   0
         Left            =   1230
         TabIndex        =   10
         Tag             =   "imgSerie"
         Top             =   5280
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
         Index           =   1
         Left            =   1230
         TabIndex        =   11
         Tag             =   "imgSerie"
         Top             =   5670
         Width           =   765
      End
      Begin VB.TextBox txtDpto 
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
         TabIndex        =   7
         Tag             =   "imgDpto"
         Top             =   3450
         Width           =   1275
      End
      Begin VB.TextBox txtDpto 
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
         TabIndex        =   6
         Tag             =   "imgDpto"
         Top             =   3060
         Width           =   1275
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
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "imgFec"
         Top             =   1260
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
         Index           =   0
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgFec"
         Top             =   840
         Width           =   1305
      End
      Begin VB.Image imgAgente 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   4140
         Width           =   255
      End
      Begin VB.Image imgAgente 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   4530
         Width           =   255
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
         Left            =   240
         TabIndex        =   63
         Top             =   4170
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
         Index           =   19
         Left            =   240
         TabIndex        =   62
         Top             =   4530
         Width           =   615
      End
      Begin VB.Image imgDpto 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   3060
         Width           =   255
      End
      Begin VB.Image imgDpto 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Agente"
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
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   59
         Top             =   3840
         Width           =   960
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
         Height          =   195
         Index           =   18
         Left            =   210
         TabIndex        =   58
         Top             =   510
         Width           =   2280
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
         Left            =   240
         TabIndex        =   57
         Top             =   840
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
         Index           =   16
         Left            =   240
         TabIndex        =   56
         Top             =   1230
         Width           =   615
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   2
         Left            =   945
         Top             =   810
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   3
         Left            =   945
         Top             =   1230
         Width           =   240
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
         Left            =   240
         TabIndex        =   45
         Top             =   1620
         Width           =   1890
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
         Left            =   240
         TabIndex        =   44
         Top             =   1950
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
         Index           =   9
         Left            =   240
         TabIndex        =   43
         Top             =   2340
         Width           =   615
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   2340
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   1950
         Width           =   255
      End
      Begin VB.Label lblNumFactu 
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
         Left            =   2580
         TabIndex        =   42
         Top             =   2700
         Width           =   4035
      End
      Begin VB.Label lblNumFactu 
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
         Left            =   2610
         TabIndex        =   41
         Top             =   2340
         Width           =   4035
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   3990
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   3990
         Top             =   840
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
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   3450
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
         TabIndex        =   39
         Top             =   3090
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
         Index           =   4
         Left            =   3300
         TabIndex        =   38
         Top             =   1260
         Width           =   615
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
         TabIndex        =   37
         Top             =   3990
         Width           =   4095
      End
      Begin VB.Label lblFecha 
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
         TabIndex        =   36
         Top             =   3630
         Width           =   4095
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
         Left            =   3300
         TabIndex        =   35
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Departamento"
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
         Index           =   7
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   1590
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
         Left            =   3300
         TabIndex        =   33
         Top             =   540
         Width           =   2280
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
         Left            =   270
         TabIndex        =   32
         Top             =   4980
         Width           =   960
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   5700
         Width           =   255
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   5280
         Width           =   255
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
         Left            =   240
         TabIndex        =   31
         Top             =   5700
         Width           =   735
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
         Left            =   240
         TabIndex        =   30
         Top             =   5280
         Width           =   780
      End
   End
   Begin VB.Frame frameConceptoDer 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6345
      Left            =   7140
      TabIndex        =   46
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox Check1 
         Caption         =   "Informe deuda vencida"
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
         Left            =   240
         TabIndex        =   81
         Top             =   6000
         Width           =   3075
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmTESCobrosPdtesList.frx":000C
         Left            =   2340
         List            =   "frmTESCobrosPdtesList.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pagado pendiente vencimiento"
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
         TabIndex        =   17
         Top             =   5250
         Width           =   4155
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
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4230
         Width           =   2055
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
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3330
         Width           =   2055
      End
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
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3780
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar Observaciones"
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
         Left            =   240
         TabIndex        =   18
         Top             =   5610
         Width           =   3075
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   4230
         TabIndex        =   47
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2430
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   720
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   4286
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Recibo anticipado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   180
         TabIndex        =   80
         Top             =   4770
         Width           =   1815
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   2160
         Picture         =   "frmTESCobrosPdtesList.frx":0010
         ToolTipText     =   "Seleccionar todos"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   1800
         Picture         =   "frmTESCobrosPdtesList.frx":015A
         ToolTipText     =   "Quita seleccion"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Situación Jurídica"
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
         Index           =   21
         Left            =   180
         TabIndex        =   68
         Top             =   4335
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   50
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Devuelto"
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
         Index           =   13
         Left            =   180
         TabIndex        =   49
         Top             =   3885
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Recibido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   180
         TabIndex        =   48
         Top             =   3390
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   10680
      TabIndex        =   28
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
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
      Left            =   9120
      TabIndex        =   26
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccion 
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
      Left            =   150
      TabIndex        =   27
      Top             =   9210
      Width           =   1335
   End
End
Attribute VB_Name = "frmTESCobrosPdtesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 602


' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
Public Legalizacion As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmAgen As frmBasico
Attribute frmAgen.VB_VarHelpID = -1
Private WithEvents frmDpto As frmBasico
Attribute frmDpto.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1

Private SQL As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim PrimeraVez As Boolean

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



Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Check1(0).Value = 1 Then
                Check1(2).Value = 1
                Check1_Click (2)
            End If
        Case 1
            If Check1(1).Value = 1 Then
                Check1(2).Value = 0
                Check1(0).Value = 0
            End If
        Case 2
            If Check1(2).Value = 1 Then
                Check1(1).Value = 0
            End If
        Case 4
            
            Combo1.ListIndex = IIf(Check1(4).Value = 1, 2, 0)
            Combo2.ListIndex = 2
            Combo3.ListIndex = IIf(Check1(4).Value = 1, 2, 0)
            Combo4.ListIndex = 2
            If Check1(4).Value = 1 Then
                For i = 1 To Me.ListView1(1).ListItems.Count
                    ListView1(1).ListItems(i).Checked = True
                Next
                
            End If
    End Select
End Sub

Private Sub cmdAccion_Click(Index As Integer)

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    tabla = "(cobros INNER JOIN formapago on cobros.codforpa = formapago.codforpa)  "
    tabla = tabla & " INNER JOIN tipofpago on formapago.tipforpa = tipofpago.tipoformapago "
    
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
    
    
    If Me.Check1(4).Value Then InsertaTmp
    
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
    
    If Legalizacion <> "" Then
        CadenaDesdeOtroForm = "OK"
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub





Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Legalizacion <> "" Then
            optTipoSal(2).Value = True
            optVarios(1).Value = True
            cmdAccion_Click (1)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Listado de Cobros Pendientes Clientes"

    For i = 0 To 1
        Me.imgSerie(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgDpto(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgAgente(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    For i = 0 To 3
        Me.ImgFec(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
    CargarListView 1
    CargarCombo
    
    Me.optVarios(0).Value = True
    
    optVarios_Click (0)
    
    Combo1.ListIndex = 0
    Combo2.ListIndex = 2
    Combo3.ListIndex = 0
    Combo4.ListIndex = 2
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    If Legalizacion <> "" Then
        txtfecha(2).Text = RecuperaValor(Legalizacion, 1)
        txtfecha(0).Text = RecuperaValor(Legalizacion, 2)
        txtfecha(1).Text = RecuperaValor(Legalizacion, 3)
    End If
    
    optVarios(0).Value = True
    optVarios(3).Value = True
    
    
End Sub



Private Sub frmAgen_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtAgente(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNAgente(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtSerie(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNSerie(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmDpto_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtDpto(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNDpto(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtfecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' tipos de forma de pago
        Case 0
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = False
            Next i
        Case 1
            For i = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(i).Checked = True
            Next i
    End Select
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub ImgAgente_Click(Index As Integer)
    IndCodigo = Index

    Set frmAgen = New frmBasico
    AyudaAgentes frmAgen, txtAgente(Index)
    Set frmAgen = Nothing
    
    PonFoco Me.txtSerie(Index)
End Sub



Private Sub imgDpto_Click(Index As Integer)
    IndCodigo = Index

    Set frmDpto = New frmBasico
    AyudaDepartamentos frmDpto, txtDpto(Index), "codmacta = " & DBSet(txtCuentas(0).Text, "T")
    Set frmDpto = Nothing
    
    PonFoco Me.txtDpto(Index)
End Sub



Private Sub imgSerie_Click(Index As Integer)
    IndCodigo = Index

    Set frmConta = New frmBasico
    AyudaContadores frmConta, txtSerie(Index), "tiporegi REGEXP '^[0-9]+$' = 0"
    Set frmConta = Nothing
    
    PonFoco Me.txtSerie(Index)
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2, 3
        IndCodigo = Index
    
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

Private Sub imgCuentas_Click(Index As Integer)
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
End Sub


Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_Click(Index As Integer)
'    Check1(1).Enabled = optVarios(1).Value
'    If Not Check1(1).Enabled Then Check1(1).Value = 0
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
  
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
        frmppal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmppal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmppal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmppal.cd1.FilterIndex = 1
    frmppal.cd1.ShowSave
    If frmppal.cd1.FileTitle <> "" Then
        If Dir(frmppal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmppal.cd1.FileName
    End If
End Sub

Private Sub PushButtonImpr_Click()
    frmppal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
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
        Case 0, 1 'cuentas
            Cta = (txtCuentas(Index).Text)
            B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
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
                    
                    'If txtCta(1).Text = "" Then 'ANTES solo lo hacia si el texto estaba vacio
                If Hasta >= 0 Then
                    txtCuentas(Hasta).Text = txtCuentas(Index).Text
                    txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
                End If
            End If
    
    
    End Select
    
    
    ' solo se puede introducir departamento si cuenta cliente desde y hasta son iguales
    txtDpto(0).Enabled = (txtCuentas(0).Text = txtCuentas(1).Text)
    imgDpto(0).Enabled = txtDpto(0).Enabled
    imgDpto(1).Enabled = txtDpto(1).Enabled
    If Not txtDpto(0).Enabled Then
        txtDpto(0).Text = ""
        txtNDpto(0).Text = ""
    End If
    txtDpto(1).Enabled = (txtCuentas(0).Text = txtCuentas(1).Text)
    If Not txtDpto(1).Enabled Then
        txtDpto(1).Text = ""
        txtNDpto(1).Text = ""
    End If
    

End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgSerie"
        imgSerie_Click Indice
    Case "imgFec"
        imgFec_Click Indice
    Case "imgCuentas"
        imgCuentas_Click Indice
    Case "imgAgente"
        ImgAgente_Click Indice
    Case "imgDpto"
        imgDpto_Click Indice
    End Select
End Sub


Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSerie_GotFocus(Index As Integer)
    ConseguirFoco txtSerie(Index), 3
End Sub

Private Sub txtSerie_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtSerie(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If

End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
        KEYpress KeyAscii
End Sub



Private Sub txtSerie_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
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
            If Index = 1 Then PonerFocoBtn Me.cmdAccion(1)
    End Select
    

    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


End Sub

'****** dpto
Private Sub txtDpto_GotFocus(Index As Integer)
    ConseguirFoco txtDpto(Index), 3
End Sub

Private Sub txtDpto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtDpto(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If

End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtDpto_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer

    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'departamentos
            If txtDpto(Index) <> "" And txtCuentas(0) <> "" Then
                txtNDpto(Index).Text = DevuelveDesdeBDNew(cConta, "departamentos", "descripcion ", "codmacta", txtCuentas(0), "T", , "Dpto", txtDpto(Index), "N")
                If txtNDpto(Index).Text <> "" Then txtDpto(Index).Text = Format(txtDpto(Index).Text, "0000")
            End If
    End Select

End Sub

'****** agentes
Private Sub txtAgente_GotFocus(Index As Integer)
    ConseguirFoco txtAgente(Index), 3
End Sub

Private Sub txtAgente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtAgente(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAgente_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer

    txtAgente(Index).Text = Trim(txtAgente(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'agentes
            txtNAgente(Index).Text = DevuelveDesdeBDNew(cConta, "agentes", "nombre", "codigo", txtAgente(Index), "N")
            If txtNAgente(Index).Text <> "" Then txtAgente(Index).Text = Format(txtAgente(Index).Text, "0000")
    End Select

End Sub


Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    
    If Me.Check1(4).Value = 1 Then
        SQL = "SELECT tmpevolsal.CODMACTA cliente,CUENTAS.NOMMACTA Nombre ,abs(APERTURA) Forpa,tmpevolsal.nomMACTA NomForpa,importemes1 MENOS30,importemes2 menos60,importemes3 menos90,importemes4 menos120,importemes5 mas120"
        SQL = SQL & " from tmpevolsal,cuentas where tmpevolsal.codmacta=cuentas.codmacta and tmpevolsal.codusu=" & vUsu.Codigo
        Sql2 = "tmpevolsal.CODMACTA"
        If optVarios(3).Value Then Sql2 = " CUENTAS.NOMMACTA"
        SQL = SQL & " ORDER BY " & Sql2 & " , tmpevolsal.nomMACTA"
    Else
        'O que habia
        
        SQL = "Select cobros.codmacta Cliente, cobros.nomclien Nombre, cobros.fecfactu FFactura, cobros.fecvenci FVenci, "
        SQL = SQL & " cobros.numorden Orden, cobros.gastos Gastos, cobros.impcobro Cobrado, cobros.impvenci ImpVenci, "
        SQL = SQL & " concat(cobros.numserie,' ', concat('0000000',cobros.numfactu)) Factura , cobros.codforpa FPago, "
        SQL = SQL & " formapago.nomforpa Descripcion, cobros.referencia Referenciasa, tipofpago.descformapago Tipo "
        
        If optVarios(0).Value Or optVarios(1).Value Then
            SQL = SQL & ", cobros.noremesar NoRemesar, cobros.situacionjuri SitJuridica, cobros.Devuelto Devuelto, cobros.recedocu Recepcion, cobros.observa Observaciones "
        End If
        
        
        RC = ""
        SQL_Combo Combo1, RC 'recibido
        SQL_Combo Combo2, RC 'devuelto
        SQL_Combo Combo3, RC    'situ juridica
        
        
        
        
        SQL = SQL & " FROM (cobros inner join formapago on cobros.codforpa = formapago.codforpa) "
        SQL = SQL & " inner join tipofpago on formapago.tipforpa = tipofpago.tipoformapago "
        SQL = SQL & " WHERE cobros.situacion = 0 and (cobros.impvenci + coalesce(cobros.gastos,0) - coalesce(cobros.impcobro,0)) <> 0 "
        If cadselect <> "" Then SQL = SQL & " AND " & cadselect
                
                
        
                
                
                
        Sql2 = ""
                
        If optVarios(0).Value Then
            If optVarios(3).Value Then Sql2 = Sql2 & " cobros.codmacta"
            If optVarios(4).Value Then Sql2 = Sql2 & " cobros.nomclien"
        End If
        
        If optVarios(1).Value Then
            Sql2 = Sql2 & " cobros.FecVenci"
            
            If optVarios(3).Value Then Sql2 = Sql2 & ",cobros.codmacta"
            If optVarios(4).Value Then Sql2 = Sql2 & ",cobros.nomclien"
        End If
    
        If optVarios(2).Value Then
            Sql2 = Sql2 & " tipofpago.descformapago"
            
            If optVarios(3).Value Then Sql2 = Sql2 & ",cobros.codmacta"
            If optVarios(4).Value Then Sql2 = Sql2 & ",cobros.nomclien"
        End If
    
    
    
    
    
        SQL = SQL & " ORDER BY " & Sql2
    
    End If '
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT_ As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT_ = "0602-00"
    
    

    If optVarios(0).Value Then
        If optVarios(3).Value Then cadParam = cadParam & "pGroup1={cobros.codmacta}|"
        If optVarios(4).Value Then cadParam = cadParam & "pGroup1={cobros.nomclien}|"
        
        numParam = numParam + 1
        
        cadParam = cadParam & "pTipo=1|"
        numParam = numParam + 1
        
        If Check1(1).Value Then cadParam = cadParam & "pResumen=1|"
        numParam = numParam + 1
                
         cadParam = cadParam & "pOrden={cobros.fecvenci}|"
         numParam = numParam + 1
        
        
    End If
    
    If optVarios(1).Value Then
        cadParam = cadParam & "pGroup1={cobros.fecvenci}|"
        
        If optVarios(3).Value Then cadParam = cadParam & "pOrden={cobros.codmacta}|"
        If optVarios(4).Value Then cadParam = cadParam & "pOrden={cobros.nomclien}|"
        numParam = numParam + 2
        
        cadParam = cadParam & "pTipo=2|"
        numParam = numParam + 1
    
        If Check1(1).Value Then
            cadParam = cadParam & "pResumen=1|"
            numParam = numParam + 1
        End If
    End If

    If optVarios(2).Value Then
        cadParam = cadParam & "pGroup1={tipofpago.descformapago}|"
        
        If optVarios(3).Value Then cadParam = cadParam & "pOrden={cobros.codmacta}|"
        If optVarios(4).Value Then cadParam = cadParam & "pOrden={cobros.nomclien}|"
        numParam = numParam + 2
    
        If Check1(1).Value Then
            cadParam = cadParam & "pResumen=1|"
            numParam = numParam + 1
        End If
    End If

    If Check1(0).Value Then
        cadParam = cadParam & "pObserva=1|"
        numParam = numParam + 1
    End If


    If optVarios(0).Value Or optVarios(1).Value Then
        'formato extendido
        If Check1(2).Value Then indRPT_ = "0602-01"
    End If
    If optVarios(2).Value Then
        indRPT_ = "0602-02"
        If Check1(2).Value Then indRPT_ = "0602-03"
    End If
    
        
    If Me.Check1(4).Value = 1 Then
        indRPT_ = "0602-04"
        SQL = "{tmpevolsal.codmacta}"
        If optVarios(4).Value Then cadParam = cadParam & "{cuentas.nommacta}"
        SQL = "pOrden2=" & SQL & "|"
        cadParam = cadParam & SQL
        numParam = numParam + 1
    End If
    If Not PonerParamRPT(indRPT_, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu

    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text, (Legalizacion <> "")) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 35
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub

Private Function MontaSQL() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim J As Integer


    MontaSQL = False
    
    If Not PonerDesdeHasta("cobros.NumSerie", "SER", Me.txtSerie(0), Me.txtNSerie(0), Me.txtSerie(1), Me.txtNSerie(1), "pDHSerie=""Serie ") Then Exit Function
    If Not PonerDesdeHasta("cobros.FecFactu", "F", Me.txtfecha(0), Me.txtfecha(0), Me.txtfecha(1), Me.txtfecha(1), "pDHFecha=""F. Factura ") Then Exit Function
    If Not PonerDesdeHasta("cobros.Fecvenci", "F", Me.txtfecha(2), Me.txtfecha(2), Me.txtfecha(3), Me.txtfecha(3), "pDHFecVto=""F.Vto: ") Then Exit Function
    If Not PonerDesdeHasta("cobros.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
    If Not PonerDesdeHasta("cobros.agente", "AGE", Me.txtAgente(0), Me.txtNAgente(0), Me.txtAgente(1), Me.txtNAgente(1), "pDHAgente=""Agente ") Then Exit Function
    If Not PonerDesdeHasta("cobros.departamento", "DPTO", Me.txtDpto(0), Me.txtNDpto(0), Me.txtDpto(1), Me.txtNDpto(1), "pDHDpto=""") Then Exit Function
            
    SQL = ""
    For J = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(J).Checked Then
            SQL = SQL & Me.ListView1(1).ListItems(J).SubItems(2) & ","
        Else
            If Me.Check1(4).Value Then SQL = SQL & Me.ListView1(1).ListItems(J).SubItems(2) & ","
        End If
    Next J
   
        
    If SQL <> "" Then
        ' quitamos la ultima coma
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        
        If Not AnyadirAFormula(cadselect, "formapago.tipforpa in (" & SQL & ")") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{formapago.tipforpa} in [" & SQL & "]") Then Exit Function
    Else
        If Not AnyadirAFormula(cadselect, "formapago.tipforpa is null") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "isnull({formapago.tipforpa})") Then Exit Function
    End If
    
    'Octubre 2018
    'NO HABIA NINGUN COMBO PUESTO
    RC = ""
    If Check1(4).Value = 0 Then
        SQL_Combo Combo1, RC 'recibido
        SQL_Combo Combo2, RC 'devuelto
        SQL_Combo Combo3, RC    'situ juridica
        SQL_Combo Combo4, RC
    End If
    
    If RC <> "" Then
        i = InStr(1, cadParam, "|pDHFecVto=""")
        If i = 0 Then
            cadParam = cadParam & "pDHFecVto=""" & RC & """|"
        Else
            J = InStr(i + 10, cadParam, "|")
            If J = 0 Then
                MsgBox "Imposible poner etiqueta en informe. El programa continuará. Avise soporte", vbExclamation
            Else
                RC2 = Mid(cadParam, J + 1)
                SQL = Mid(cadParam, i + 12, J - i - 13)
                SQL = SQL & "       " & RC
                SQL = "pDHFecVto=""Fecha vto. " & SQL & """|"
                RC = Mid(cadParam, 1, i)
                
                cadParam = RC & SQL & RC2
                
            End If
            
            
        End If
    End If

    
    
    ' Añadimos la condicion de que la situacion = 0 y que el importe pendiente <> 0
    If Check1(4).Value = 0 Then
        If Check1(3).Value = 0 Then
            'LO que habia, antes Junio18
            If Not AnyadirAFormula(cadselect, "cobros.situacion = 0") Then Exit Function
            If Not AnyadirAFormula(cadFormula, "{cobros.situacion} = 0") Then Exit Function
            
            If Not AnyadirAFormula(cadselect, "(cobros.impvenci + coalesce(cobros.gastos,0) - coalesce(cobros.impcobro,0)) <> 0") Then Exit Function
            If Not AnyadirAFormula(cadFormula, "{@Pendiente} <> 0") Then Exit Function
    
        Else
            'Riesgo
            RC = "{cobros.situacion} =0 or ({cobros.situacion}=1 AND {formapago.tipforpa} IN [4]) AND {cobros.Fecvenci} > Date(" & Format(Now, "yyyy,mm,dd") & ")"
            If Not AnyadirAFormula(cadFormula, RC) Then Exit Function
            RC = "cobros.situacion =0 or (cobros.situacion=1 and fecvenci >" & DBSet(Now, "F") & " AND formapago.tipforpa IN (4)) "
            If Not AnyadirAFormula(cadselect, RC) Then Exit Function
            
        
        End If
    
    Else
        'RC = "{cobros.situacion} =0 or {{cobros.situacion}=1 AND {formapago.tipforpa} IN [4]) AND {cobros.Fecvenci} > Date(" & Format(Now, "yyyy,mm,dd") & ")"""
        'If Not AnyadirAFormula(cadFormula, RC) Then Exit Function
        RC = "situacion=0 and (impvenci + coalesce(gastos,0) - coalesce(impcobro,0))<>0"
        If Not AnyadirAFormula(cadselect, RC) Then Exit Function
    End If
        
    
    
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
            
    MontaSQL = True
End Function

Private Sub SQL_Combo(ByRef Cbo As ComboBox, ByRef CadenaDH As String)
Dim Aux As String

    If Cbo.ListIndex <> 2 Then
        Aux = Right(Cbo.Name, 1)
        If Aux = "2" Then
            'Copmbo2
            Aux = "{cobros.Devuelto}"
        ElseIf Aux = "3" Then
            'Combo3
            Aux = "{cobros.situacionjuri}"
        ElseIf Aux = "4" Then
            'Combo 4
            Aux = "{cobros.reciboanticipado}"
        Else
            'Combo1
            Aux = "{cobros.recedocu}"
        End If
        
        
        Aux = Aux & " = " & IIf(Cbo.ListIndex = 1, 1, 0)
        If Not AnyadirAFormula(cadFormula, Aux) Then Exit Sub
        Aux = Replace(Aux, "{", "")
        Aux = Replace(Aux, "}", "")
        If Not AnyadirAFormula(cadselect, Aux) Then Exit Sub
        
    
    
        Aux = Right(Cbo.Name, 1)
        If Aux = "2" Then
            'Copmbo2
            Aux = "Devuelto:"
        ElseIf Aux = "3" Then
            'Combo3
            Aux = "Situacion jurídica:"
        ElseIf Aux = "4" Then
            'Combo4
            Aux = "Recibo anticipado:"
        Else
            'Combo1
            Aux = "Recepcion:"
        End If
        Aux = "    " & Aux & IIf(Cbo.ListIndex = 1, "Si", "No")
        CadenaDH = Trim(CadenaDH & Aux)
    End If
End Sub



Private Sub txtfecha_LostFocus(Index As Integer)
    txtfecha(Index).Text = Trim(txtfecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtfecha(Index)
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

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    
    DatosOK = True


End Function


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub CargarListView(Index As Integer)
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList
 
    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , " ", 300
    ListView1(Index).ColumnHeaders.Add , , "Descripción", 3200
    ListView1(Index).ColumnHeaders.Add , , " ", 0
    
    SQL = "Select * from tipofpago order by 1" 'descformapago"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        
        ItmX.Checked = True
        ItmX.Text = " " ' Rs.Fields(0).Value
        ItmX.SubItems(1) = Rs.Fields(1).Value
        ItmX.SubItems(2) = Rs.Fields(0).Value
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tipo de Forma de Pago.", Err.Description
    End If
End Sub

Private Sub CargarCombo()
    'recibido
    Combo1.Clear
    Combo1.AddItem "No"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Sí"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Todos"
    Combo1.ItemData(Combo1.NewIndex) = 2
    'devuelto
    Combo2.Clear
    Combo2.AddItem "No"
    Combo2.ItemData(Combo2.NewIndex) = 0
    Combo2.AddItem "Sí"
    Combo2.ItemData(Combo2.NewIndex) = 1
    Combo2.AddItem "Todos"
    Combo2.ItemData(Combo2.NewIndex) = 2
    'situacion jurídica
    Combo3.Clear
    Combo3.AddItem "No"
    Combo3.ItemData(Combo3.NewIndex) = 0
    Combo3.AddItem "Sí"
    Combo3.ItemData(Combo3.NewIndex) = 1
    Combo3.AddItem "Todos"
    Combo3.ItemData(Combo3.NewIndex) = 2

    'recibo anticipado
    Combo4.Clear
    Combo4.AddItem "No"
    Combo4.ItemData(Combo4.NewIndex) = 0
    Combo4.AddItem "Sí"
    Combo4.ItemData(Combo4.NewIndex) = 1
    Combo4.AddItem "Todos"
    Combo4.ItemData(Combo4.NewIndex) = 2


End Sub



Private Sub InsertaTmp()
Dim vi() As Currency
Dim Impor2 As Currency

    Conn.Execute "DELETE FROM tmpevolsal WHERE codusu =" & vUsu.Codigo
    ReDim vi(4)
    SQL = ""
    SQL = "select cobros.*,nomforpa"
    SQL = SQL & " from cobros,formapago where cobros.codforpa=formapago.codforpa AND " & cadselect
    SQL = SQL & "  order by codmacta,codforpa"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cadFormula = "INSERT INTO tmpevolsal(codusu,codmacta,nommacta,apertura,importemes1,importemes2,importemes3,importemes4,importemes5,importemes6) VALUES "
    RC = ""
    Do
        'NO ES EOF, ya lo hemos comprobado
        Msg = miRsAux!codmacta & Format(miRsAux!Codforpa, "0000")
        If RC <> Msg Then
            If RC <> "" Then
                Impor2 = 0
                For i = 0 To 4
                    SQL = SQL & ", " & DBSet(vi(i), "N")
                    Impor2 = Impor2 + vi(i)
                Next
                SQL = SQL & ", " & DBSet(Impor2, "N")
                    
                
                SQL = cadFormula & SQL & ")"
                Conn.Execute SQL
            End If
                
                
            For i = 0 To 4
                vi(i) = 0
            Next
            RC = Msg
            SQL = " (" & vUsu.Codigo & "," & DBSet(miRsAux!codmacta, "T") & "," & DBSet(miRsAux!nomforpa, "T") & "," & Abs(miRsAux!Codforpa)
            
        End If
        NumRegElim = DateDiff("d", miRsAux!FecVenci, Now())
        If NumRegElim <= 30 Then
            i = 0
        ElseIf NumRegElim <= 60 Then
                i = 1
        ElseIf NumRegElim <= 90 Then
                i = 2
        ElseIf NumRegElim <= 120 Then
                i = 3
        Else
            i = 4
        End If
        Impor2 = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
        vi(i) = vi(i) + Impor2
        
        miRsAux.MoveNext
    Loop Until miRsAux.EOF
    miRsAux.Close
    
    'El ultimo
     If RC <> "" Then
        Impor2 = 0
        For i = 0 To 4
            SQL = SQL & ", " & DBSet(vi(i), "N")
            Impor2 = Impor2 + vi(i)
        Next
        SQL = SQL & ", " & DBSet(Impor2, "N")
            
        
        SQL = cadFormula & SQL & ")"
        Conn.Execute SQL
    End If
    
    
    espera 0.5
    
    SQL = "select codmacta,sum(importemes6),min(nommacta) from tmpevolsal WHERE codusu = " & vUsu.Codigo & "  group by 1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = "UPDATE tmpevolsal SET importemes7 = " & DBSet(miRsAux.Fields(1), "N") & "  WHERE codusu=" & vUsu.Codigo
        SQL = SQL & " AND codmacta =" & DBSet(miRsAux!codmacta, "T") & " AND nommacta=" & DBSet(miRsAux.Fields(2), "T")
        Conn.Execute SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Set miRsAux = Nothing
    
    cadFormula = "{tmpevolsal.codusu} = " & vUsu.Codigo
End Sub
