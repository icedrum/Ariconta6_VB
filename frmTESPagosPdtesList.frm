VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESPagosPdtesList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "frmTESPagosPdtesList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
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
      TabIndex        =   46
      Top             =   4170
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   49
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   48
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
         TabIndex        =   47
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
      TabIndex        =   31
      Top             =   4170
      Width           =   4785
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   150
         TabIndex        =   43
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
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   150
         TabIndex        =   39
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
            Left            =   2880
            TabIndex        =   42
            Top             =   270
            Width           =   1515
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
            Left            =   1440
            TabIndex        =   41
            Top             =   270
            Width           =   1365
         End
         Begin VB.OptionButton optVarios 
            Caption         =   "Proveedor"
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
            TabIndex        =   40
            Top             =   270
            Width           =   1395
         End
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
         TabIndex        =   9
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
      Height          =   4155
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6915
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
         Index           =   3
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2340
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   3450
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
         TabIndex        =   32
         Top             =   3030
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
         Tag             =   "imgConcepto"
         Top             =   2340
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
         Tag             =   "imgConcepto"
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
         Left            =   1200
         TabIndex        =   6
         Tag             =   "imgConcepto"
         Top             =   3030
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
         Left            =   1200
         TabIndex        =   7
         Tag             =   "imgConcepto"
         Top             =   3450
         Width           =   765
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
         Tag             =   "imgConcepto"
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
         Tag             =   "imgConcepto"
         Top             =   840
         Width           =   1305
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
         Left            =   270
         TabIndex        =   38
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
         Left            =   270
         TabIndex        =   37
         Top             =   870
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
         Left            =   270
         TabIndex        =   36
         Top             =   1230
         Width           =   615
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   2
         Left            =   960
         Top             =   840
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   3
         Left            =   960
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Proveedor"
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
         TabIndex        =   25
         Top             =   1620
         Width           =   2370
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   2310
         Width           =   615
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   2370
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   1920
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         Top             =   870
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
         Index           =   4
         Left            =   3300
         TabIndex        =   20
         Top             =   1260
         Width           =   615
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   900
         Width           =   690
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
         TabIndex        =   17
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
         Left            =   180
         TabIndex        =   16
         Top             =   2730
         Width           =   960
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Index           =   1
         Left            =   900
         Top             =   3450
         Width           =   255
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Index           =   0
         Left            =   900
         Top             =   3030
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
         Left            =   180
         TabIndex        =   15
         Top             =   3450
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
         Left            =   180
         TabIndex        =   14
         Top             =   3030
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
      Height          =   4155
      Left            =   7140
      TabIndex        =   26
      Top             =   0
      Width           =   4815
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
         TabIndex        =   8
         Top             =   3570
         Width           =   2055
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   4230
         TabIndex        =   27
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
         Height          =   2310
         Index           =   1
         Left            =   180
         TabIndex        =   29
         Top             =   1020
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   4075
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4110
         Picture         =   "frmTESPagosPdtesList.frx":000C
         ToolTipText     =   "Puntear al Debe"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3750
         Picture         =   "frmTESPagosPdtesList.frx":0156
         ToolTipText     =   "Quitar al Debe"
         Top             =   720
         Width           =   240
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
         Left            =   180
         TabIndex        =   30
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "Emitido"
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
         Left            =   300
         TabIndex        =   28
         Top             =   3630
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
      TabIndex        =   12
      Top             =   6990
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
      TabIndex        =   10
      Top             =   6990
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
      TabIndex        =   11
      Top             =   6990
      Width           =   1335
   End
End
Attribute VB_Name = "frmTESPagosPdtesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 802


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

Private Sql As String
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


Private Sub cmdAccion_Click(Index As Integer)

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    tabla = "(pagos INNER JOIN formapago on pagos.codforpa = formapago.codforpa)  "
    tabla = tabla & " INNER JOIN tipofpago on formapago.tipforpa = tipofpago.tipoformapago "
    
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
    
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
    Me.Caption = "Listado de Pagos Pendientes Proveedores"

    For i = 0 To 1
        Me.imgSerie(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
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
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    optVarios(0).Value = True
    optVarios(3).Value = True
    
    
End Sub

Private Sub optVarios_Click(Index As Integer)
'    Check1(1).Enabled = optVarios(1).Value
'    If Not Check1(1).Enabled Then Check1(1).Value = 0
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtSerie(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNSerie(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
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


Private Sub imgSerie_Click(Index As Integer)
    IndCodigo = Index

    Set frmConta = New frmBasico
    AyudaContadores frmConta, txtSerie(Index), "tiporegi REGEXP '^[0-9]+$' <> 0 and cast(tiporegi as unsigned) > 0"
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
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgCuentas_Click(Index As Integer)
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
End Sub


Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
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
        Case 0, 1 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
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



Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    Sql = "Select cobros.codmacta Cliente, cobros.nomclien Nombre, cobros.fecfactu FFactura, cobros.fecvenci FVenci, "
    Sql = Sql & " cobros.numorden Orden, cobros.gastos Gastos, cobros.impcobro Cobrado, cobros.impvenci ImpVenci, "
    Sql = Sql & " concat(cobros.numserie,' ', concat('0000000',cobros.numfactu)) Factura , cobros.codforpa FPago, "
    Sql = Sql & " formapago.nomforpa Descripcion, cobros.referencia Referenciasa, tipofpago.descformapago Tipo "
    
    If optVarios(0).Value Or optVarios(1).Value Then
        Sql = Sql & ", cobros.noremesar NoRemesar, cobros.situacionjuri SitJuridica, cobros.Devuelto Devuelto, cobros.recedocu Recepcion, cobros.observa Observaciones "
    End If
    
    Sql = Sql & " FROM (cobros inner join formapago on cobros.codforpa = formapago.codforpa) "
    Sql = Sql & " inner join tipofpago on formapago.tipforpa = tipofpago.tipoformapago "
    Sql = Sql & " WHERE cobros.situacion = 0 and (cobros.impvenci + coalesce(cobros.gastos,0) - coalesce(cobros.impcobro,0)) <> 0 "
    If cadselect <> "" Then Sql = Sql & " AND " & cadselect
            
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

    Sql = Sql & " ORDER BY " & Sql2
            
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0802-00"
    

    If optVarios(0).Value Then
        If optVarios(3).Value Then cadParam = cadParam & "pGroup1={pagos.codmacta}|"
        If optVarios(4).Value Then cadParam = cadParam & "pGroup1={pagos.nomprove}|"
        
        numParam = numParam + 1
        
        cadParam = cadParam & "pTipo=1|"
        numParam = numParam + 1
        
        If check1(1).Value Then cadParam = cadParam & "pResumen=1|"
        numParam = numParam + 1
        
         cadParam = cadParam & "pOrden={pagos.fecefect}|"
         numParam = numParam + 1
    End If
    
    If optVarios(1).Value Then
        cadParam = cadParam & "pGroup1={pagos.fecefect}|"
        
        If optVarios(3).Value Then cadParam = cadParam & "pOrden={pagos.codmacta}|"
        If optVarios(4).Value Then cadParam = cadParam & "pOrden={pagos.nomprove}|"
        numParam = numParam + 2
        
        cadParam = cadParam & "pTipo=2|"
        numParam = numParam + 1
    
        If check1(1).Value Then
            cadParam = cadParam & "pResumen=1|"
            numParam = numParam + 1
        End If
    End If

    If optVarios(2).Value Then
        cadParam = cadParam & "pGroup1={tipofpago.descformapago}|"
        
        If optVarios(3).Value Then cadParam = cadParam & "pOrden={pagos.codmacta}|"
        If optVarios(4).Value Then cadParam = cadParam & "pOrden={pagos.nomprove}|"
        numParam = numParam + 2
    
        If check1(1).Value Then
            cadParam = cadParam & "pResumen=1|"
            numParam = numParam + 1
        End If
    End If


    ' ordenado por forma de pago
    If optVarios(2).Value Then
        indRPT = "0802-01"
    End If
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "CobrosPdtes.rpt"

    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 15
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub

Private Function CargarTemporal() As Boolean
Dim Sql As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    Sql = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "insert into tmpfaclin (codusu, codigo, numserie, nomserie, numfac, fecha, cta, cliente, nif, imponible, impiva, total, retencion,"
    Sql = Sql & " recargo, tipoopera, tipoformapago) "
    Sql = Sql & " select distinct " & vUsu.Codigo & ",0, factcli.numserie, contadores.nomregis, factcli.numfactu, factcli.fecfactu, factcli.codmacta, "
    Sql = Sql & " factcli.nommacta, factcli.nifdatos, factcli.totbases, factcli.totivas, factcli.totfaccl, factcli.trefaccl, "
    Sql = Sql & " factcli.totrecargo, tipofpago.descformapago , aa.denominacion"
    Sql = Sql & " from " & tabla
    Sql = Sql & " where " & cadselect
    
    Conn.Execute Sql
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
End Function

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim i As Integer
Dim TxtDHTipoPago As String  'el lw y llevara tambien emitido

    MontaSQL = False
    
    If Not PonerDesdeHasta("pagos.NumSerie", "SER", Me.txtSerie(0), Me.txtNSerie(0), Me.txtSerie(1), Me.txtNSerie(1), "pDHSerie=""") Then Exit Function
    If Not PonerDesdeHasta("pagos.FecFactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    If Not PonerDesdeHasta("pagos.fecefect", "F", Me.txtFecha(2), Me.txtFecha(2), Me.txtFecha(3), Me.txtFecha(3), "pDHFecVto=""") Then Exit Function
    If Not PonerDesdeHasta("pagos.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
    
    TxtDHTipoPago = ""
    Sql = ""
    RC = ""
    RC2 = ""
    For i = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then
            Sql = Sql & Me.ListView1(1).ListItems(i).SubItems(2) & ","
            RC2 = RC2 & "- " & Me.ListView1(1).ListItems(i).SubItems(3)
        Else
            RC = "N"
        End If
    Next i
    
    If Sql <> "" Then
        ' quitamos la ultima coma
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        
        If Not AnyadirAFormula(cadselect, "formapago.tipforpa in (" & Sql & ")") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{formapago.tipforpa} in [" & Sql & "]") Then Exit Function
        
        If RC <> "" Then
            'Significa que ALGUNA forma de pago NO ha sido seleccionada
                RC2 = Mid(RC2, 2)
                TxtDHTipoPago = "Tipo pago:  " & RC2
        End If
    Else
        If Not AnyadirAFormula(cadselect, "formapago.tipforpa is null") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "isnull({formapago.tipforpa})") Then Exit Function
    End If
    
    ' Añadimos la condicion de que la situacion = 0 y que el importe pendiente <> 0
    If Not AnyadirAFormula(cadselect, "pagos.situacion = 0") Then Exit Function
    If Not AnyadirAFormula(cadFormula, "{pagos.situacion} = 0") Then Exit Function
    
    If Not AnyadirAFormula(cadselect, "(pagos.impefect - coalesce(pagos.imppagad,0)) <> 0") Then Exit Function
    If Not AnyadirAFormula(cadFormula, "{@Pendiente} <> 0") Then Exit Function
    
    RC = ""
    Select Case Combo1.ListIndex
        Case 0 ' No
            If Not AnyadirAFormula(cadselect, "pagos.emitdocum = 0") Then Exit Function
            If Not AnyadirAFormula(cadFormula, "{pagos.emitdocum} = 0") Then Exit Function
            RC = "No"
        Case 1 ' Si
            If Not AnyadirAFormula(cadselect, "pagos.emitdocum = 1") Then Exit Function
            If Not AnyadirAFormula(cadFormula, "{pagos.emitdocum} = 1") Then Exit Function
            RC = "Si"
        Case 2 ' Todos
            
    End Select
    If RC <> "" Then
        If TxtDHTipoPago <> "" Then TxtDHTipoPago = TxtDHTipoPago & "   "
        TxtDHTipoPago = TxtDHTipoPago & "     Doc. emitido: " & RC
    End If
    
    cadParam = cadParam & "tipoPagoDH=""" & Trim(TxtDHTipoPago) & """|"
    
    
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
            
    MontaSQL = True
End Function

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtFecha(Index)
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
Dim Sql As String

    On Error GoTo ECargarList
 
    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

    ListView1(Index).ColumnHeaders.Add , , " ", 300
    ListView1(Index).ColumnHeaders.Add , , "Descripción", 3200
    ListView1(Index).ColumnHeaders.Add , , " ", 0
    ListView1(Index).ColumnHeaders.Add , , " ", 0  'Siglas
    
    Sql = "Select * from tipofpago order by 1" 'descformapago"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        
        ItmX.Checked = True
        ItmX.Text = " " ' Rs.Fields(0).Value
        ItmX.SubItems(1) = Rs.Fields(1).Value
        ItmX.SubItems(2) = Rs.Fields(0).Value
        ItmX.SubItems(3) = Rs.Fields(2).Value
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
End Sub
