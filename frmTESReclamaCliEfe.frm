VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESReclamaCliEfe 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "frmTESReclamaCliEfe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
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
      TabIndex        =   58
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   61
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   720
         Width           =   1515
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
      TabIndex        =   25
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         Tag             =   "imgConcepto"
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
         Tag             =   "imgConcepto"
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   3480
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   5700
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   44
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
         Left            =   1230
         TabIndex        =   10
         Tag             =   "imgConcepto"
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
         Tag             =   "imgConcepto"
         Top             =   5700
         Width           =   765
      End
      Begin VB.TextBox txtFPago 
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
         Tag             =   "imgConcepto"
         Top             =   3480
         Width           =   1275
      End
      Begin VB.TextBox txtFPago 
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
         Tag             =   "imgConcepto"
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
      Begin VB.Image imgAgente 
         Height          =   255
         Index           =   0
         Left            =   900
         Top             =   4140
         Width           =   255
      End
      Begin VB.Image imgAgente 
         Height          =   255
         Index           =   1
         Left            =   900
         Top             =   4590
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   4530
         Width           =   615
      End
      Begin VB.Image imgDpto 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   900
         Top             =   3060
         Width           =   255
      End
      Begin VB.Image imgDpto 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   900
         Top             =   3510
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
         TabIndex        =   51
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
         Left            =   270
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         Index           =   2
         Left            =   240
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Forma de Pago"
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
         TabIndex        =   30
         Top             =   2760
         Width           =   1770
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
         TabIndex        =   29
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
         Left            =   210
         TabIndex        =   28
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
         Left            =   210
         TabIndex        =   27
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
         Left            =   210
         TabIndex        =   26
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
      Height          =   9015
      Left            =   7140
      TabIndex        =   42
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtVarios 
         BeginProperty Font 
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
         Left            =   240
         TabIndex        =   17
         Top             =   6360
         Width           =   4215
      End
      Begin VB.TextBox txtVarios 
         BeginProperty Font 
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
         Left            =   270
         TabIndex        =   16
         Top             =   5700
         Width           =   4215
      End
      Begin VB.TextBox txtVarios 
         BeginProperty Font 
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
         TabIndex        =   15
         Top             =   5040
         Width           =   4215
      End
      Begin VB.TextBox txtVarios 
         Height          =   285
         Index           =   1
         Left            =   4980
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   5070
         Width           =   2775
      End
      Begin VB.TextBox txtCarta 
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
         Left            =   240
         TabIndex        =   14
         Tag             =   "imgConcepto"
         Top             =   4230
         Width           =   765
      End
      Begin VB.TextBox txtNCarta 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   4230
         Width           =   3405
      End
      Begin VB.CheckBox chkExcluirConEmail 
         Caption         =   "Excluir clientes con email (carta)"
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
         Left            =   270
         TabIndex        =   72
         Top             =   8520
         Width           =   4095
      End
      Begin VB.CheckBox chkInsertarReclamas 
         Caption         =   "Insertar registros reclamaciones"
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
         Left            =   270
         TabIndex        =   71
         Top             =   8100
         Width           =   4095
      End
      Begin VB.CheckBox chkMarcarUtlRecla 
         Caption         =   "Marcar fecha última reclamación"
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
         Left            =   270
         TabIndex        =   20
         Top             =   7680
         Width           =   4095
      End
      Begin VB.CheckBox chkMostrarCta 
         Caption         =   "Mostrar cuenta bancaria destinatario"
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
         Left            =   270
         TabIndex        =   19
         Top             =   7260
         Width           =   4275
      End
      Begin VB.TextBox txtDias 
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
         Left            =   2850
         TabIndex        =   13
         Top             =   3480
         Width           =   1485
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
         Left            =   240
         TabIndex        =   12
         Top             =   3510
         Width           =   1485
      End
      Begin VB.CheckBox chkReclamaDevueltos 
         Caption         =   "Sólo devueltos"
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
         Left            =   270
         TabIndex        =   18
         Top             =   6840
         Width           =   3075
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2190
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   780
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   3863
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
         Caption         =   "Cargo"
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
         Left            =   240
         TabIndex        =   78
         Top             =   6120
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Firmante"
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
         Index           =   21
         Left            =   270
         TabIndex        =   77
         Top             =   5400
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
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
         Left            =   270
         TabIndex        =   75
         Top             =   4740
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Carta"
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
         Left            =   240
         TabIndex        =   74
         Top             =   3960
         Width           =   630
      End
      Begin VB.Image imgCarta 
         Height          =   255
         Left            =   930
         Top             =   3930
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Días"
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
         Index           =   23
         Left            =   2850
         TabIndex        =   70
         Top             =   3210
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha reclamación"
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
         Index           =   22
         Left            =   270
         TabIndex        =   69
         Top             =   3210
         Width           =   1800
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   4
         Left            =   2130
         Picture         =   "frmTESReclamaCliEfe.frx":000C
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4110
         Picture         =   "frmTESReclamaCliEfe.frx":0097
         ToolTipText     =   "Puntear al Debe"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3750
         Picture         =   "frmTESReclamaCliEfe.frx":01E1
         ToolTipText     =   "Quitar al Debe"
         Top             =   480
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
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   1500
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
      TabIndex        =   24
      Top             =   9120
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
      Left            =   9150
      TabIndex        =   22
      Top             =   9120
      Width           =   1425
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
      TabIndex        =   23
      Top             =   9090
      Width           =   1335
   End
End
Attribute VB_Name = "frmTESReclamaCliEfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 608


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
Private WithEvents frmCarta As frmBasico
Attribute frmCarta.VB_VarHelpID = -1
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
Dim Fecha As Date

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
Dim Rs As ADODB.Recordset
Dim indRPT As String
Dim nomDocu As String


    If Not DatosOK Then Exit Sub
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    indRPT = "0608-01"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "CobrosPdtes.rpt"
    
    
    ' si es por email
    If Me.optTipoSal(3).Value Then
        
        cad = "DELETE FROM tmp347 WHERE codusu =" & vUsu.Codigo
        Conn.Execute cad
        
        Set Rs = New ADODB.Recordset
        
        cad = "SELECT fechaadq,maidatos,razosoci,nommacta FROM tmpentrefechas,cuentas WHERE"
        cad = cad & " fechaadq=codmacta AND    CodUsu = " & vUsu.Codigo
        cad = cad & " GROUP BY fechaadq ORDER BY maidatos"
        Rs.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        
        cad = "pFechaIMP= """ & txtFecha(4).Text & """|"
        cad = cad & "pverCCC= " & Abs(Me.chkMostrarCta) & "|"
        
        cadParam = cadParam & cad
        numParam = numParam + 2
        
        Sql = "{tmpentrefechas.codusu}=" & vUsu.Codigo
        NumRegElim = 0
        CONT = 0
        frmppal.Visible = False

        While Not Rs.EOF
            Me.Refresh
            espera 0.5
            RC = DBLet(Rs!maidatos, "T")
            If RC = "" Then
                
                If MsgBox("Sin mail para la cuenta: " & Rs!fechaadq & " - " & Rs!Nommacta & vbCrLf & "    ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then
                    CONT = 0
                    Rs.MoveLast
                End If
                
                Sql = "INSERT INTO tmp347 (codusu, cliprov, cta, nif, importe) VALUES (" & vUsu.Codigo
                Sql = Sql & ",0," & Rs!fechaadq & ",NULL,0)"
                '
                'AL meter la cuenta con el importe a 0, entonces no la leera para enviarala
                'Pero despues si k podremos NO actualizar sus pagosya que no se han enviado nada
                Conn.Execute Sql
            Else
                Screen.MousePointer = vbHourglass

                cadPDFrpt = cadNomRPT
                
                With frmVisReport
                    .Informe = App.Path & "\Informes\"
                    If ExportarPDF Then
                        'PDF
                        .Informe = .Informe & cadPDFrpt
                    Else
                        'IMPRIMIR
                        .Informe = .Informe & cadNomRPT
                    End If
                    .FormulaSeleccion = "{tmpentrefechas.codusu}=" & vUsu.Codigo & " AND {tmpentrefechas.nomconam}= """ & Rs.Fields(0) & """"
                    .SoloImprimir = False
                    .OtrosParametros = cadParam
                    .NumeroParametros = numParam
                    .ConSubInforme = True
            
                    .NumCopias2 = 1
            
                    .SoloImprimir = SoloImprimir
                    .ExportarPDF = True
                    .MostrarTree = vMostrarTree
                    .Show vbModal
                    'HaPulsadoImprimir = .EstaImpreso
                 
                 End With

                If CadenaDesdeOtroForm = "OK" Then
                    Me.Refresh
                    espera 0.5
                    CONT = CONT + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    Sql = Rs.Fields(0) & ".pdf"

                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Sql

                    'Insertamos en tmp347 la cuenta
                    Sql = "INSERT INTO tmp347(codusu, cliprov, cta, nif, importe) VALUES (" & vUsu.Codigo & ",0,'" & Rs.Fields(0) & "','" & Sql & "',1)"
                    Conn.Execute Sql
                End If

                
            End If
            Rs.MoveNext
        Wend
        Rs.Close

        If CONT > 0 Then
             
             espera 0.5
             
             Sql = "Reclamacion fecha: " & txtFecha(4).Text & "|"
             
             Sql = Sql & "Reclamación pago facturas efectuada el : " & txtFecha(4).Text & "|"
             
             'Escalona
             Sql = txtVarios(0).Text & "|Recuerde: En el archivo adjunto le enviamos información de su interés.|"

             LanzaProgramaAbrirOutlookMasivo 1, Sql

             If chkMarcarUtlRecla.Value = 1 Then
                If MsgBox("¿ Proceso realizado correctamente para actualizar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                   ActualizarRegistros
                End If
             End If

             If chkInsertarReclamas.Value Then
                If InsertarReclamaciones Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancelar_Click
                End If
             End If
             
             Exit Sub
            
        End If
        
    End If
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    tabla = "tmpentrefechas"
    
    
    If Not HayRegParaInforme(tabla, "tmpentrefechas.codusu=" & vUsu.Codigo) Then Exit Sub
    
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
    
    If chkMarcarUtlRecla.Value = 1 Then
        If MsgBox("¿ Proceso realizado correctamente para actualizar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
           ActualizarRegistros
        End If
    End If
    If chkInsertarReclamas.Value Then
        If InsertarReclamaciones Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            cmdCancelar_Click
        End If
    End If
End Sub

Private Function InsertarReclamaciones() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim CONT As Long

    On Error GoTo eInsertarReclamaciones

    InsertarReclamaciones = False


    Set Rs = New ADODB.Recordset

    'FINALMENTE GRABAMOS LA TABLA HCO
    Sql = "SELECT MAX(codigo) from reclama"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not Rs.EOF Then CONT = DBLet(Rs.Fields(0), "N")
    Rs.Close
    CONT = CONT + 1

    Sql = "SELECT tmpentrefechas.nomconam,cuentas.nommacta from tmpentrefechas, cuentas where codusu = " & vUsu.Codigo
    Sql = Sql & " and tmpentrefechas.nomconam = cuentas.codmacta "
    Sql = Sql & " group by 1"
    Sql = Sql & " order by 1"
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Sql = "INSERT INTO reclama (codigo, fecreclama, codmacta, nommacta, carta, importes, observaciones)"
        Sql = Sql & " VALUES (" & DBSet(CONT, "N") & "," & "'" & Format(txtFecha(4).Text, FormatoFecha) & "'," & DBSet(Rs!nomconam, "T")
        Sql = Sql & "," & DBSet(Rs!Nommacta, "T") & ","
    
        If optTipoSal(3).Value = 1 Then
            Sql = Sql & "1,"
        Else
            Sql = Sql & "0,"
        End If
    
        Sql = Sql & "0,"
        Sql = Sql & "'Reclamación de Fecha:" & txtFecha(4).Text & "')"
        
        Conn.Execute Sql
        
        ' grabamos las lineas
        Sql = "INSERT INTO reclama_facturas (codigo, numlinea, numserie, numfactu, fecfactu, numorden, impvenci) "
        Sql = Sql & " select " & DBSet(CONT, "N") & ", @numf:=@numf+1, codccost, nomccost, fecventa, conconam, "
        Sql = Sql & " COALESCE(impventa,0)+COALESCE(valoradq,0)-COALESCE(impperiodo,0) from tmpentrefechas, (SELECT @numf:=0) aaa where codusu = " & vUsu.Codigo
        Sql = Sql & " and nomconam = " & DBSet(Rs!nomconam, "T")
        
        Conn.Execute Sql
        
        ' acturalizamos el importe
        Sql = "update reclama set importes = (select sum(impvenci) from reclama_facturas where codigo = " & DBSet(CONT, "N") & ")"
        Sql = Sql & " where codigo = " & DBSet(CONT, "N")
        Conn.Execute Sql
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    InsertarReclamaciones = True
    Exit Function
    
eInsertarReclamaciones:
    MuestraError Err.Number, "Insertar Reclamaciones", Err.Description
End Function


Private Sub ActualizarRegistros()
Dim Sql As String

    On Error GoTo eActualizarRegistros
    
    ' si es por email borramos aquellos que no se hayan enviado
    If Me.optTipoSal(3).Value = 1 Then
        Sql = "delete from tmp347 where codusu = " & vUsu.Codigo & " and importe = 0"
        Conn.Execute Sql
    End If


    Sql = "update cobros, tmpentrefechas set cobros.ultimareclamacion = " & DBSet(txtFecha(4).Text, "F")
    Sql = Sql & " where tmpentrefechas.codusu = " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & " and tmpentrefechas.codccost = cobros.numserie "
    Sql = Sql & " and tmpentrefechas.nomccost = cobros.numfactu "
    Sql = Sql & " and tmpentrefechas.fecventa = cobros.fecfactu "
    Sql = Sql & " and tmpentrefechas.conconam = cobros.numorden "

    Conn.Execute Sql
    Exit Sub

eActualizarRegistros:
    MuestraError Err.Number, "Actualizar Registros", Err.Description
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Legalizacion <> "" Then
            optTipoSal(2).Value = True
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
    Me.Caption = "Efectuar Reclamaciones Clientes"

    For i = 0 To 1
        Me.imgSerie(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgDpto(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgAgente(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgCarta.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    For i = 0 To 3
        Me.ImgFec(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
     
     
    CargarListView 1
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    optTipoSal(1).Enabled = False
    txtTipoSalida(1).Enabled = False
    PushButton2(0).Enabled = False
    
End Sub



Private Sub frmAgen_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtAgente(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNAgente(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCarta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCarta.Text = RecuperaValor(CadenaSeleccion, 1)
        txtNCarta.Text = RecuperaValor(CadenaSeleccion, 2)
    End If
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

Private Sub imgCarta_Click()

    Set frmCarta = New frmBasico
    AyudaCartas frmCarta, txtCarta
    Set frmCarta = Nothing
    
    PonFoco Me.txtCarta

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
    Case 0, 1, 2, 3, 4
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
    If Index = 3 Then
        If optTipoSal(Index).Value = 1 Then
            chkExcluirConEmail.Enabled = False
            chkExcluirConEmail.Value = 0
        Else
            chkExcluirConEmail.Enabled = True
        End If
    End If
End Sub

Private Sub optVarios_Click(Index As Integer)
'    Check1(1).Enabled = optVarios(1).Value
'    If Not Check1(1).Enabled Then Check1(1).Value = 0
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
  
  
'    Check1(1).Enabled = optVarios(1).Value
    
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
                    
                    'If txtCta(1).Text = "" Then 'ANTES solo lo hacia si el texto estaba vacio
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
    Case "imgAgente"
        ImgAgente_Click Indice
    Case "imgCarta"
        imgCarta_Click
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

'****** Carta
Private Sub txtCarta_GotFocus()
    ConseguirFoco txtCarta, 3
End Sub

Private Sub txtCarta_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtcarta_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCarta_LostFocus()
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
Dim Hasta As Integer

    txtCarta.Text = Trim(txtCarta.Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If txtCarta <> "" Then
        txtNCarta.Text = DevuelveDesdeBDNew(cConta, "cartas", "descarta", "codcarta", txtCarta, "N")
        If txtNCarta.Text <> "" Then txtCarta.Text = Format(txtCarta.Text, "0000")
    End If

End Sub

'******


'****** agentes
Private Sub txtAgente_GotFocus(Index As Integer)
    ConseguirFoco txtAgente(Index), 3
End Sub

Private Sub txtAgente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAgente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAgente_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String
Dim Cta As String
Dim B As Boolean
Dim Sql As String
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
    Sql = "Select cobros.codmacta Cliente, cobros.nomclien Nombre, cobros.fecfactu FFactura, cobros.fecvenci FVenci, "
    Sql = Sql & " cobros.numorden Orden, cobros.gastos Gastos, cobros.impcobro Cobrado, cobros.impvenci ImpVenci, "
    Sql = Sql & " concat(cobros.numserie,' ', concat('0000000',cobros.numfactu)) Factura , cobros.codforpa FPago, "
    Sql = Sql & " formapago.nomforpa Descripcion, cobros.referencia Referenciasa, tipofpago.descformapago Tipo "
    
    Sql = Sql & " FROM (cobros inner join formapago on cobros.codforpa = formapago.codforpa) "
    Sql = Sql & " inner join tipofpago on formapago.tipforpa = tipofpago.tipoformapago "
    If cadselect <> "" Then Sql = Sql & " WHERE " & cadselect
            
            

    Sql = Sql & " ORDER BY " & Sql2

            
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = True
    
    indRPT = "0608-01"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "reclama.rpt"

    cadParam = cadParam & "pFechaIMP= """ & txtFecha(4).Text & """|"
    cadParam = cadParam & "pverCCC= " & Abs(Me.chkMostrarCta) & "|"
    cadParam = cadParam & "firmante= """ & txtVarios(2).Text & """|"
    cadParam = cadParam & "cargo= """ & txtVarios(3).Text & """|"
    cadParam = cadParam & "Asunto= """ & txtVarios(0).Text & """|"
    numParam = numParam + 4
    
    cadFormula = "{tmpentrefechas.codusu}=" & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 39
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub

Private Function CargarTemporal() As Boolean
Dim Rs As ADODB.Recordset
Dim i As Long
Dim Importe As Currency
Dim Dpto As Long

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    Screen.MousePointer = vbHourglass
    
    'Ahora haremos todo el proceso
    i = Val(txtDias.Text)
    i = i * -1
    Fecha = CDate(txtFecha(4).Text)
    Fecha = DateAdd("d", i, Fecha)
    
    'Ya tenemos en F la fecha a partir de la cual reclamamos
    'Montamos el SQL
    MontaSQLReclamacion
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
    
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    If i = 0 Then
        MsgBox "No hay valores entre las fechas", vbExclamation
        Exit Function
    End If
    
    
    
    
    'No enlazamos por NIF, si no k en NIF guardaremos codmacta
    'codinmov, nominmov, fechaadq, valoradq, amortacu, fecventa, impventa, impperiodo) VALUES (
    cad = "DELETE FROM tmpentrefechas WHERE codusu = " & vUsu.Codigo
    Conn.Execute cad
    
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = "INSERT INTO tmpentrefechas(codusu,codigo,codccost,nomccost,fecventa,conconam,fechaadq"
    Sql = Sql & ",nominmov,impventa,impperiodo,valoradq,codinmov ) VALUES (" & vUsu.Codigo & ","
    
    'Nuevo. Febrero 2010. Departamento ira en codinmov
    
    
    i = 1
    While Not Rs.EOF
    
        'Neuvo Febero 2010
        'Ademas de ver si me debe algo, si esta recibido NO lo puedo meter
        
        Importe = Rs!ImpVenci + DBLet(Rs!Gastos, "N") - DBLet(Rs!impcobro, "N")
        If DBLet(Rs!recedocu, "N") = 1 Then Importe = 0
        If Importe > 0 Then
            cad = i & ",'" & Rs!NUmSerie & "','"
            cad = cad & Rs!NumFactu & "','"
            cad = cad & Format(Rs!FecFactu, FormatoFecha) & "',"
            cad = cad & Rs!numorden & ",'"
            cad = cad & Rs!codmacta & "','"
            cad = cad & Rs!FecVenci & "',"
            cad = cad & TransformaComasPuntos(CStr(Rs!ImpVenci)) & ","
            If IsNull(Rs!impcobro) Then
                cad = cad & "NULL"
            Else
                cad = cad & TransformaComasPuntos(CStr(Rs!impcobro))
            End If
            'ValorADQ=GASTOS
            cad = cad & "," & TransformaComasPuntos(CStr(DBLet(Rs!Gastos, "N")))
            
            'Febrero 2010
            'Departamento
            cad = cad & "," & DBLet(Rs!departamento, "N")
            cad = Sql & cad & ")"
            Conn.Execute cad
            
            i = i + 1
            
        End If
        Rs.MoveNext
        
    Wend
    Rs.Close
    
    If i = 1 Then
        'Ningun valor con esa opcion
        MsgBox "No hay valores entre las fechas", vbExclamation
        Exit Function
    End If
    
    
    
    'Noviembre 2014
    'Comprobamos que todas las cuentas tienen email(si va por email)
    If optTipoSal(3).Value Then
        CadenaDesdeOtroForm = ""
        frmTESVarios.Opcion = 31
        frmTESVarios.Show vbModal
        
        If CadenaDesdeOtroForm = "" Then
            Screen.MousePointer = vbDefault
            Set Rs = Nothing
            Exit Function
        End If
    End If
    
    
    
    cad = "DELETE FROM tmpcuentas  where codusu = " & vUsu.Codigo
    Conn.Execute cad
    
    cad = "SELECT fechaadq,codinmov FROM tmpentrefechas WHERE codusu = " & vUsu.Codigo & " GROUP BY fechaadq,codinmov"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Datos contables
    Set miRsAux = New ADODB.Recordset
    CONT = 0
    While Not Rs.EOF
        'BUSCAMOS DATOS
        cad = "SELECT * from cuentas where codmacta='" & Rs.Fields(0) & "'"
    
        'Insertar datos en z347
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'Nuevo. Ya no llevamos NIF, llevaremos departamento
        RC = "" 'SERA EL NIF. Sera el DPTO
        i = 1
        If Not miRsAux.EOF Then
            'NIF -> codmacta
            RC = Rs.Fields(0)
            Dpto = Rs.Fields(1)
        Else
            'EOF
            i = 0
            MsgBox "No se encuentra la cuenta: " & Rs.Fields(0), vbExclamation
            'NOS SALIMOS
            Rs.Close
            Exit Function
        End If
        
        'NO es EOF y tiene NIF
        If i > 0 Then
            'Aumentamos el contador
            CONT = CONT + 1
            
            
            Sql = "INSERT INTO tmpcuentas (codusu, codmacta, nommacta,despobla,razosoci,dpto) VALUES (" & vUsu.Codigo & ",'" & RC & "','"
            Sql = Sql & DBLet(miRsAux!nifdatos, "T") & "','" 'En nommacta meto el NIF del cliente
                
                
            cad = DBLet(miRsAux!IBAN, "T")
            
            cad = cad & "','"
            'El dpto si tiene
        '    cad = cad & DevNombreSQL(DevuelveDesdeBD("descripcion", "departamentos", "codmacta = '" & miRsAux!codmacta & "' AND dpto", CStr(Dpto)))
        '    cad = cad & "'," & Dpto
        '    Ejecuta Sql & cad & ")"   'Lo pongo en funcion para que no me de error
            
            Sql = "UPDATE tmpentrefechas SET nomconam='" & RC & "' WHERE fechaadq = '" & Rs!fechaadq & "'"
            Sql = Sql & " AND codusu = " & vUsu.Codigo
            Conn.Execute Sql
            
            
        End If
        miRsAux.Close
            
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    
        
    If CONT = 0 Then
        MsgBox "Ningun dato devuelto para procesar por carta/mail", vbExclamation
        Exit Function
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    
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
Dim Rs As ADODB.Recordset
Dim i As Long
Dim Importe As Currency



    MontaSQL = False
    
    If Not PonerDesdeHasta("cobros.NumSerie", "SER", Me.txtSerie(0), Me.txtNSerie(0), Me.txtSerie(1), Me.txtNSerie(1), "pDHSerie=""") Then Exit Function
    If Not PonerDesdeHasta("cobros.FecFactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    If Not PonerDesdeHasta("cobros.Fecvenci", "F", Me.txtFecha(2), Me.txtFecha(2), Me.txtFecha(3), Me.txtFecha(3), "pDHFecVto=""") Then Exit Function
    If Not PonerDesdeHasta("cobros.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
    If Not PonerDesdeHasta("cobros.agente", "AGE", Me.txtAgente(0), Me.txtNAgente(0), Me.txtAgente(1), Me.txtNAgente(1), "pDHAgente=""") Then Exit Function
            
    Sql = ""
    For i = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then
            Sql = Sql & Me.ListView1(1).ListItems(i).SubItems(2) & ","
        End If
    Next i
    
    If Sql <> "" Then
        ' quitamos la ultima coma
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        
        If Not AnyadirAFormula(cadselect, "formapago.tipforpa in (" & Sql & ")") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{formapago.tipforpa} in [" & Sql & "]") Then Exit Function
    Else
        If Not AnyadirAFormula(cadselect, "formapago.tipforpa is null") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "isnull({formapago.tipforpa})") Then Exit Function
    End If
    
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
    If Not CargarTemporal Then Exit Function
    
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
    
    If txtFecha(4).Text = "" Then
        MsgBox "Introduzca la Fecha de Reclamación.", vbExclamation
        PonleFoco txtFecha(4)
        Exit Function
    End If
    If txtDias.Text = "" Then
        MsgBox "Introduzca los días de última Reclamación.", vbExclamation
        PonleFoco txtDias
        Exit Function
    End If
    If txtCarta.Text = "" Then
        MsgBox "Seleccione la carta a adjuntar.", vbExclamation
        PonleFoco txtCarta
        Exit Function
    End If
    'Si poner marcar como reclamacion entonces debe estar marcada la opcion
    'de insertar en las tablas de col reclamas
    If chkMarcarUtlRecla.Value = 1 Then
        If Me.chkInsertarReclamas.Value = 0 Then
            MsgBox "Debe marcar tambien la opcion de ' INSERTAR REGISTROS RECLAMACIONES '", vbExclamation
            Exit Function
        End If
    End If
       
    
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
    
    Sql = "Select * from tipofpago order by 1" 'descformapago"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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

Private Sub MontaSQLReclamacion()
    
    'Siempre hay que añadir el AND
    
    Sql = " and " & cadselect
    
    'Solo devueltos
    If chkReclamaDevueltos.Value = 1 Then
        Sql = Sql & " AND devuelto = 1"
    Else
        'QUE FALTE POR PAGAR
        Sql = Sql & " AND ((impvenci + coalesce(gastos,0) - coalesce(impcobro,0) > 0) or (devuelto = 1))"
    End If
    
    'Marzo2015
    If chkExcluirConEmail.Value = 1 Then Sql = Sql & " AND coalesce(maidatos,'')=''"
    
    'LA de la fecha
    Sql = Sql & " AND ((ultimareclamacion  is null) OR (ultimareclamacion <= '" & Format(Fecha, FormatoFecha) & "'))"
    
    
    'Select
    cad = "Select cobros.*, cuentas.codmacta FROM cobros,cuentas,formapago "
    cad = cad & " WHERE  formapago.codforpa=cobros.codforpa AND cobros.codmacta = cuentas.codmacta"
    cad = cad & " AND formapago.codforpa=cobros.codforpa "
    Sql = cad & Sql
    
    
End Sub

