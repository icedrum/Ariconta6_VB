VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInfBalSumSal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   6375
      Left            =   7110
      TabIndex        =   14
      Top             =   0
      Width           =   4455
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   3300
         TabIndex        =   54
         Top             =   5460
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cuentas Agrupadas"
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkBalIncioEjercicio 
         Caption         =   "Inventario inicial"
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
         Left            =   180
         TabIndex        =   53
         Top             =   5970
         Width           =   3795
      End
      Begin VB.TextBox txtPag2 
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
         Left            =   1110
         TabIndex        =   49
         Tag             =   "imgConcepto"
         Top             =   1920
         Width           =   1305
      End
      Begin VB.TextBox txtTitulo 
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
         Left            =   150
         MaxLength       =   25
         TabIndex        =   46
         Tag             =   "imgConcepto"
         Top             =   1410
         Width           =   4065
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
         Index           =   7
         Left            =   2040
         TabIndex        =   45
         Top             =   780
         Width           =   1485
      End
      Begin VB.Frame Frame2 
         Height          =   2145
         Left            =   150
         TabIndex        =   32
         Top             =   2340
         Width           =   4185
         Begin VB.CheckBox Check1 
            Caption         =   "9� nivel"
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
            Index           =   9
            Left            =   120
            TabIndex        =   43
            Top             =   1290
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "8� nivel"
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
            Index           =   8
            Left            =   2850
            TabIndex        =   42
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "7� nivel"
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
            Index           =   7
            Left            =   1470
            TabIndex        =   41
            Top             =   960
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "6� nivel"
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
            Index           =   6
            Left            =   120
            TabIndex        =   40
            Top             =   930
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "5� nivel"
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
            Index           =   5
            Left            =   2850
            TabIndex        =   39
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "4� nivel"
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
            Left            =   1470
            TabIndex        =   38
            Top             =   600
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "3� nivel"
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
            Left            =   120
            TabIndex        =   37
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2� nivel"
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
            TabIndex        =   36
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1er nivel"
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
            Left            =   1470
            TabIndex        =   35
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "�ltimo:  "
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
            Index           =   10
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Value           =   1  'Checked
            Width           =   1155
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1650
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Remarcar"
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
            Left            =   360
            TabIndex        =   44
            Top             =   1710
            Width           =   1005
         End
      End
      Begin VB.CheckBox chkAgrupacionCtasBalance 
         Caption         =   "Agrupar cuentas en balance"
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
         Left            =   180
         TabIndex        =   27
         Top             =   5520
         Width           =   3135
      End
      Begin VB.CheckBox chkApertura 
         Caption         =   "Desglosar el saldo de apertura"
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
         Left            =   180
         TabIndex        =   26
         Top             =   4710
         Width           =   3345
      End
      Begin VB.CheckBox chkMovimientos 
         Caption         =   "Acumulados y movimientos del periodo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   25
         Top             =   5010
         Value           =   1  'Checked
         Width           =   4125
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   24
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
      Begin VB.Label Label3 
         Caption         =   "1�P�gina"
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
         Index           =   1
         Left            =   150
         TabIndex        =   50
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Left            =   150
         TabIndex        =   48
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "T�tulo"
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
         Left            =   150
         TabIndex        =   47
         Top             =   1140
         Width           =   690
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   7
         Left            =   1680
         Picture         =   "frmInfBalSumSal.frx":0000
         Top             =   810
         Width           =   240
      End
   End
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selecci�n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3705
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNCta 
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
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1050
         Width           =   4185
      End
      Begin VB.TextBox txtNCta 
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
         Index           =   7
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1470
         Width           =   4185
      End
      Begin VB.TextBox txtAno 
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
         Index           =   1
         Left            =   3270
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2700
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmInfBalSumSal.frx":008B
         Left            =   1230
         List            =   "frmInfBalSumSal.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2670
         Width           =   1935
      End
      Begin VB.TextBox txtAno 
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
         Index           =   0
         Left            =   3270
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2220
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmInfBalSumSal.frx":008F
         Left            =   1230
         List            =   "frmInfBalSumSal.frx":0091
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2220
         Width           =   1935
      End
      Begin VB.TextBox txtCta 
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
         Left            =   1230
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1050
         Width           =   1275
      End
      Begin VB.TextBox txtCta 
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
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   6
         Left            =   990
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   7
         Left            =   990
         Top             =   1500
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
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   1440
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
         TabIndex        =   22
         Top             =   1080
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
         Left            =   240
         TabIndex        =   21
         Top             =   2640
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
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta"
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
         TabIndex        =   19
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Mes / A�o"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1410
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
      Left            =   10320
      TabIndex        =   4
      Top             =   6570
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
      Left            =   8730
      TabIndex        =   2
      Top             =   6570
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
      Left            =   120
      TabIndex        =   3
      Top             =   6570
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
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   6915
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
         TabIndex        =   17
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   15
         Top             =   1200
         Width           =   255
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
         TabIndex        =   12
         Top             =   1680
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
         Index           =   1
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   11
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
         Index           =   0
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   3345
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
         TabIndex        =   9
         Top             =   2160
         Width           =   975
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
         TabIndex        =   8
         Top             =   1680
         Width           =   975
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
         TabIndex        =   7
         Top             =   1200
         Width           =   1515
      End
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
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   1830
      TabIndex        =   51
      Top             =   6600
      Visible         =   0   'False
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.CommandButton cmdCancelarAccion 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   52
      Top             =   6570
      Width           =   1215
   End
End
Attribute VB_Name = "frmInfBalSumSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 306

Public Opcion As Byte
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


Public Cuenta As String
Public Descripcion As String
Public FecDesde As String
Public FecHasta As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon  As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private frmCtas As frmCtasAgrupadas

Private SQL As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim RS As ADODB.Recordset

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean

Public Legalizacion As String   'Datos para la legalizacion

Dim HanPulsadoSalir As Boolean

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



Private Sub chkAgrupacionCtasBalance_Click()
    Toolbar1.Buttons(1).Enabled = (chkAgrupacionCtasBalance.Value = 1)
End Sub

Private Sub chkBalIncioEjercicio_Click()
    If chkBalIncioEjercicio.Value = 1 Then
        txtTitulo(0).Text = "Inventario Inicial"
    Else
        txtTitulo(0).Text = "Balance de Sumas y Saldos"
    End If
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    PulsadoCancelar = False
    Me.cmdCancelarAccion.Visible = True
    Me.cmdCancelarAccion.Enabled = True
    
    Me.cmdCancelar.Visible = False
    Me.cmdCancelar.Enabled = False
        
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
'++
    'Balance a inicio de ejerecicio
    If Me.chkBalIncioEjercicio.Value = 1 Then
        'Es un balance especial.
        Screen.MousePointer = vbHourglass
        
        HacerBalanceInicio
        Screen.MousePointer = vbDefault
        
    Else
        If Not ComprobarCuentas(6, 7) Then Exit Sub
    
    
        'Ahora, si ha puesto desde hasta cuenta, no puede seleccionar
        'un desde.
        If Me.txtCta(6).Text <> "" Or txtCta(7).Text <> "" Then
            If CONT > 1 Then
                If vUsu.Nivel < 2 Then
                    cad = "debe"
                Else
                    cad = "puede"
                End If
                cad = "No " & cad & " pedir un balance a distintos niveles poniendo desde/hasta cuenta"
                MsgBox cad, vbExclamation
                If vUsu.Nivel > 1 Then Exit Sub
            End If
        End If
    
    
        If txtAno(0).Text = "" Or txtAno(1).Text = "" Then
            MsgBox "Introduce las fechas(a�os) de consulta", vbExclamation
            Exit Sub
        End If
        
        If Not ComparaFechasCombos(0, 1, 0, 1) Then Exit Sub
        
        If Abs(Val(txtAno(1).Text) - Val(txtAno(0).Text)) > 2 Then
            MsgBox "Fechas pertenecen a ejercicios distintos.", vbExclamation
            Exit Sub
        End If
        
        
    
        'Trabajaresmos contra ejercicios cerrados
        'Si el mes es mayor o igual k el de inicio, significa k la feha
        'de inicio de aquel ejercicio fue la misma k ahora pero de aquel a�o
        'si no significa k fue la misma de ahora pero del a�o anterior
        I = cmbFecha(0).ListIndex + 1
        If I >= Month(vParam.fechaini) Then
            CONT = Val(txtAno(0).Text)
        Else
            CONT = Val(txtAno(0).Text) - 1
        End If
        cad = Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & CONT
        FechaIncioEjercicio = CDate(cad)
        
        I = cmbFecha(1).ListIndex + 1
        If I <= Month(vParam.fechafin) Then
            CONT = Val(txtAno(1).Text)
        Else
            CONT = Val(txtAno(1).Text) + 1
        End If
        cad = Day(vParam.fechafin) & "/" & Month(vParam.fechafin) & "/" & CONT
        FechaFinEjercicio = CDate(cad)
    
        
        'Veamos si pertenecen a un mismo a�o
        If Abs(DateDiff("d", FechaFinEjercicio, FechaIncioEjercicio)) > 365 Then
            MsgBox "Las fechas son incorrectas. Abarca mas de un ejercicio", vbExclamation
            Exit Sub
        End If
     
        
        'Fecha informe
        If txtFecha(7).Text = "" Then txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
        
        
        EmpezarBalanceNuevo -1, pb2  'La -1 es balance normal
            
            
    End If

    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True

    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmpbalancesumas", "codusu=" & vUsu.Codigo) Then Exit Sub
    
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
    If Me.cmdCancelarAccion.Visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub Form_Activate()
Dim CONT As Integer

    If PrimeraVez Then
        PrimeraVez = False
        If Legalizacion <> "" Then
            optTipoSal(2).Value = True
            
            'Solo check enivado
            CONT = Val(RecuperaValor(Legalizacion, 4))
            For I = 1 To 10
                If I = CONT Then
                    Check1(I).Value = 1
                Else
                    Check1(I).Value = 0
                End If
            Next

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
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Balance de Sumas y Saldos"

    For I = 6 To 7
        Me.imgCuentas(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    PrimeraVez = True
     
    CargarComboFecha
     

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 29
    End With
    
    'Fecha informe
    txtFecha(7).Text = Format(Now, "dd/mm/yyyy")
    'Fecha inicial
    cmbFecha(0).ListIndex = Month(vParam.fechaini) - 1
    cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1
    
    txtAno(0).Text = Year(vParam.fechaini)
    txtAno(1).Text = Year(vParam.fechafin)
   
    txtTitulo(0).Text = Me.Caption
   
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    Me.Toolbar1.Buttons(1).Enabled = (chkAgrupacionCtasBalance.Value = 1)
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
    PonerNiveles
    
    If Legalizacion <> "" Then
        txtFecha(7).Text = RecuperaValor(Legalizacion, 1)     'Fecha informe
            
        txtAno(0).Text = Year(CDate(RecuperaValor(Legalizacion, 2)))     'Inicio
        txtAno(1).Text = Year(CDate(RecuperaValor(Legalizacion, 3)))     'Fin
        
        cmbFecha(0).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 2))) - 1
        cmbFecha(1).ListIndex = Month(CDate(RecuperaValor(Legalizacion, 3))) - 1
        
        ' Inventario
        If RecuperaValor(Legalizacion, 5) = 1 Then
            txtTitulo(0).Text = "Inventario final cierre"
            cad = "5"
            For I = 2 To vEmpresa.DigitosUltimoNivel
                cad = cad & "9"
            Next
            txtCta(7).Text = cad
        End If
    End If
    
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub




Private Sub Image1_Click(Index As Integer)

    Select Case Index
        Case 0 'cuentas agrupadas
            Set frmCtas = New frmCtasAgrupadas
            frmCtas.Show vbModal
            Set frmCtas = Nothing
    End Select

End Sub

Private Sub imgCuentas_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing

    PonFoco txtCta(Index)

End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 7
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




Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
        frmPpal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmPpal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmPpal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmPpal.cd1.FilterIndex = 1
    frmPpal.cd1.ShowSave
    If frmPpal.cd1.FileTitle <> "" Then
        If Dir(frmPpal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmPpal.cd1.FileName
    End If
End Sub

Private Sub PushButtonImpr_Click()
    frmPpal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Set frmCtas = New frmCtasAgrupadas
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
End Sub


Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgCuentas"
        imgCuentas_Click Indice
    Case "imgFecha"
        imgFec_Click Indice
    End Select
    
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        txtNCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser num�rica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        txtNCta(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 6, 7 'Cuentas
            
            RC = txtCta(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, SQL) Then
                txtCta(Index) = RC
                txtNCta(Index).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtCta(Index).Text = ""
                txtNCta(Index).Text = ""
                PonFoco txtCta(Index)
            End If
            
            If Index = 0 Then Hasta = 1
            If Hasta >= 1 Then
                txtCta(Hasta).Text = txtCta(Index).Text
                txtNCta(Hasta).Text = txtNCta(Index).Text
            End If
    End Select

End Sub



Private Sub AccionesCSV()
Dim SQL2 As String
Dim Tipo As Byte
    '1.- Sin apertura y sin movimientos
    '2.- Sin apertura y con movimientos
    '3.- Con apertura y sin movimientos
    '4.- Con apertura y con movimientos
    
    If Me.chkApertura = 1 Then
        Tipo = 3                   'Con apertura 3 y 4
    Else
        Tipo = 1                   'Sin apertura 1,2
    End If
    If chkMovimientos = 1 Then
        Tipo = Tipo + 1        'Con movimientos
    Else
        Tipo = Tipo + 0        ' Sin movimientos
    End If
            
    If Me.chkBalIncioEjercicio = 1 Then Tipo = 4

    Select Case Tipo
        Case 1
            SQL = "select cta Cuenta , nomcta Titulo, totald Saldo_deudor, totalh Saldo_acreedor from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
        Case 2
            SQL = "select cta Cuenta , nomcta Titulo, acumantd AcumAnt_deudor, acumanth AcumAnt_acreedor, acumperd AcumPer_deudor, acumperh AcumPer_acreedor, totald Saldo_deudor, totalh Saldo_acreedor from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
        Case 3
            SQL = "select cta Cuenta , nomcta Titulo, aperturad Apertura_deudor, aperturah Apertura_acreedor,  totald Saldo_deudor, totalh Saldo_acreedor from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
        Case 4
            SQL = "select cta Cuenta , nomcta Titulo, aperturad, aperturah, case when coalesce(aperturad,0) - coalesce(aperturah,0) > 0 then concat(coalesce(aperturad,0) - coalesce(aperturah,0),'D') when coalesce(aperturad,0) - coalesce(aperturah,0) < 0 then concat(coalesce(aperturah,0) - coalesce(aperturad,0),'H') when coalesce(aperturad,0) - coalesce(aperturah,0) = 0 then 0 end Apertura, "
            SQL = SQL & " acumantd AcumAnt_deudor, acumanth AcumAnt_acreedor, acumperd AcumPer_deudor, acumperh AcumPer_acreedor, "
            SQL = SQL & " totald Saldo_deudor, totalh Saldo_acreedor, case when coalesce(totald,0) - coalesce(totalh,0) > 0 then concat(coalesce(totald,0) - coalesce(totalh,0),'D') when coalesce(totald,0) - coalesce(totalh,0) < 0 then concat(coalesce(totalh,0) - coalesce(totald,0),'H') when coalesce(totald,0) - coalesce(totalh,0) = 0 then 0 end Saldo"
            SQL = SQL & " from tmpbalancesumas where codusu = " & vUsu.Codigo
            SQL = SQL & " order by 1 "
        
    End Select
    


        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    '1.- Sin apertura y sin movimientos
    '2.- Sin apertura y con movimientos
    '3.- Con apertura y sin movimientos
    '4.- Con apertura y con movimientos
    
    If Me.chkApertura = 1 Then
        Tipo = 3                   'Con apertura 3 y 4
    Else
        Tipo = 1                   'Sin apertura 1,2
    End If
    If chkMovimientos = 1 Then
        Tipo = Tipo + 1        'Con movimientos
    Else
        Tipo = Tipo + 0        ' Sin movimientos
    End If
            
    If Me.chkBalIncioEjercicio = 1 Then Tipo = 4
            
    cadParam = cadParam & "pTipo=" & Tipo & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pFecha=""" & txtFecha(7).Text & """|"
    
    'Numero de p�gina
    If txtPag2(0).Text <> "" Then
        cadParam = cadParam & "NumPag=" & txtPag2(0).Text - 1 & "|"
    Else
        cadParam = cadParam & "NumPag=0|"
    End If
    numParam = numParam + 2
    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & txtAno(0).Text & " a " & cmbFecha(1).Text & " " & txtAno(1).Text & """|"
    numParam = numParam + 1
    
    
    'Salto
    If Combo2.ListIndex >= 0 Then
        cadParam = cadParam & "Salto= " & Combo2.ItemData(Combo2.ListIndex) & "|"
    Else
        cadParam = cadParam & "Salto= 11|"
    End If
    numParam = numParam + 1
    
    
    'Titulo
    txtTitulo(0).Text = Trim(txtTitulo(0).Text)
    If txtTitulo(0).Text = "" Then
        cad = "Balance de sumas y saldos"
    Else
        cad = txtTitulo(0).Text
    End If
    cadParam = cadParam & "Titulo= """ & cad & """|"
    numParam = numParam + 1
        
        
    'Numero de digitos de ultimo nivel
    cadParam = cadParam & "pDigUltNivel=" & vEmpresa.DigitosUltimoNivel & "|"
    numParam = numParam + 1
        
        
    '------------------------------
    'Numero de niveles
    'Para cada nivel marcado veremos si tiene cuentas en la tmp
    CONT = 0
    UltimoNivel = 0
    For I = 1 To 10
        If Check1(I).Visible Then
'                If Check2(I).Value = 1 Then Cont = Cont + 1
            If Check1(I).Value = 1 Then
                If I = 10 Then
                    cad = vEmpresa.DigitosUltimoNivel
                Else
                    cad = CStr(DigitosNivel(I))
                End If
                If TieneCuentasEnTmpBalance(cad) Then
                    CONT = CONT + 1
                    UltimoNivel = CByte(cad)
                End If
            End If
        End If
    Next I
    cad = "numeroniveles= " & CONT & "|"
    SQL = SQL & cad
    'Otro parametro mas
    cad = "vUltimoNivel= " & UltimoNivel & "|"
    
    cadParam = cadParam & cad
    numParam = numParam + 2

    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "0306-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"SumasySaldos.rpt"

    cadFormula = "{tmpbalancesumas.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
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
        
        LanzaFormAyuda "imgFecha", Index
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    CONT = 0
    For I = 1 To 10
        If Check1(I).Value = 1 Then CONT = CONT + 1
    Next I
    If CONT = 0 Then
        MsgBox "Seleccione como m�nimo un nivel contable", vbExclamation
        Exit Function
    End If


    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer


QueCombosFechaCargar "0|1|"


For I = 1 To vEmpresa.numnivel - 1
    J = DigitosNivel(I)
    Check1(I).Visible = True
    Check1(I).Caption = "Digitos: " & J
Next I



'Cargamos le combo de resalte de fechas
Combo2.Clear

Combo2.AddItem "Sin remarcar"
Combo2.ItemData(Combo2.NewIndex) = 1000
For I = 1 To vEmpresa.numnivel - 1
    Combo2.AddItem "Nivel " & I
    Combo2.ItemData(Combo2.NewIndex) = I
Next I
End Sub


Private Sub EmpezarBalanceNuevo(vConta As Integer, ByRef PB As ProgressBar)
Dim Cade As String
Dim Apertura As Boolean
Dim QuitarSaldos2 As Byte
Dim Agrupa As Boolean
Dim IndiceCombo As Integer
Dim vOpcion As Byte
Dim Resetea6y7 As Boolean
Dim C1 As Long
Dim UltimoNivel As Byte
Dim PreCargarCierre

Dim InicioPeriodo As Date
Dim FinPeriodo As Date
Dim LenPrimerNivelCalculado As Byte
Dim LenNivelCalculado As Byte
Dim Fec As Date
Dim J As Integer



    Screen.MousePointer = vbHourglass
    
    
    'Nuevo agosto 2014
    '------------------------------------------------------------------------
    '  Haremos los calculos para el nivel mas alto de los selecionados.
    '  Para los inferiores sera tan facil como teneindo estos datos ya calculados
    ' agruparlos convenientemente, por length del nivel
    '
    '
    
    If vConta = -1 Then
        IndiceCombo = 0
        I = 6
    Else
        IndiceCombo = 14
        I = 18
    End If
    Cade = ""
    
    
    'Desde hasta cuentas
    If txtCta(I).Text <> "" Then Cade = Cade & " AND ((cuentas.codmacta)>='" & txtCta(I).Text & "')"
    If txtCta(I + 1).Text <> "" Then Cade = Cade & " AND ((cuentas.codmacta)<='" & txtCta(I + 1).Text & "')"
    
    'Fechas
    'Del ejercicio solicitado
    If Opcion = 24 Then
        I = 14
    Else
        I = 0
    End If
    Fec = "01/" & Right("00" & CStr(Me.cmbFecha(I).ListIndex + 1), 2) & "/" & txtAno(I).Text
    J = 1
    If Fec > vParam.fechafin Then
        'Siguiente
        FechaIncioEjercicio = DateAdd("yyyy", 1, vParam.fechaini)
    Else
    
        FechaIncioEjercicio = vParam.fechaini
    
        While J = 1
            If FechaIncioEjercicio <= Fec Then
                J = 0
            Else
                FechaIncioEjercicio = DateAdd("yyyy", -1, FechaIncioEjercicio)
            End If
        Wend
    End If
    FechaFinEjercicio = DateAdd("yyyy", 1, FechaIncioEjercicio)
    FechaFinEjercicio = DateAdd("d", -1, FechaFinEjercicio)

    Cade = " fechaent between " & DBSet(FechaIncioEjercicio, "F") & " AND " & DBSet(FechaFinEjercicio, "F") & Cade

    
    
    'Nivel del que vamos a calcular los datos
    LenPrimerNivelCalculado = 0
    If Check1(10).Value Then
        LenPrimerNivelCalculado = vEmpresa.DigitosUltimoNivel
    Else
        For I = vEmpresa.numnivel - 1 To 1 Step -1
            
            If vConta = -1 Then
                'Balance normal
                If Check1(I).Value = 1 Then LenPrimerNivelCalculado = DigitosNivel(I)
                
            Else
'                'Balance consolidado
'                If Me.ChkConso(i).Value = 1 Then LenPrimerNivelCalculado = DigitosNivel(i)
'
            End If
            If LenPrimerNivelCalculado > 0 Then Exit For
        Next
    End If
   
    'Como queremos tener ya el n pmbre de la cuenta, por eso todo el sql "complicado2
    'La otra opcion seria sacar solamente el codmacta(si lefjoins ni leches)
    ' y cargar a posteriori el nommacta. De velocidad no se va mucho, decimas
    cad = "SELECT substring(line.codmacta,1," & LenPrimerNivelCalculado
    cad = cad & ") as codmacta,coalesce(nommacta,'ERROR##') nommacta From "
    If vConta >= 0 Then cad = cad & "ariconta" & vConta & "."
    cad = cad & "hlinapu"
    cad = cad & " as line LEFT JOIN "
    If vConta >= 0 Then cad = cad & "ariconta" & vConta & "."
    cad = cad & "cuentas ON  "
    cad = cad & "cuentas.codmacta = substring(line.codmacta,1," & LenPrimerNivelCalculado & ")"
    cad = cad & " WHERE " & Cade

    cad = cad & " GROUP BY 1 ORDER By 1"

    Apertura = (Me.chkApertura.Value = 1) And vConta < 0
    
    
    'Para luego veremos que opciones ponemos
    If Me.chkApertura = 1 Then
        vOpcion = 3                   'Con apertura 3 y 4
    Else
        vOpcion = 1                   'Sin apertura 1,2
    End If
    If chkMovimientos = 1 Then
        vOpcion = vOpcion + 1        'Con movimientos
    Else
        vOpcion = vOpcion + 0        ' Sin movimientos
    End If
            
    '1.- Sin apertura y sin movimientos
    '2.- Sin apertura y con movimientos
    '3.- Con apertura y sin movimientos
    '4.- Con apertura y con movimientos
    

    
    Set RS = New ADODB.Recordset
    
    'Comprobamos si hay que quitar el pyg y el cierre
    QuitarSaldos2 = 0   'No hay k kitar
    CONT = 0
    
        CONT = 1  'Ambos
        'Si el mes contiene el cierre, entonces adelante
        If Month(vParam.fechafin) = Me.cmbFecha(IndiceCombo + 1).ListIndex + 1 Then
            'Si estamos en ejercicios cerrados seguro que hay asiento de cierre y p y g
                'Si no lo comprobamos. Concepto=960 y 980
                Agrupa = HayAsientoCierre((Me.cmbFecha(IndiceCombo + 1).ListIndex + 1), CInt(txtAno(IndiceCombo + 1).Text))
                If Agrupa Then QuitarSaldos2 = CONT
        End If
    
    'Agruparemos si esta seleccionado el chekc de agrupar y esta seleccionado
    'ultimo nivel y hay moivmientos para agrupar
    Agrupa = False
    If vConta < 0 And Me.chkAgrupacionCtasBalance.Value = 1 Then 'chekc de agrupar
        If Check1(10).Value = 1 Then                 'sheck de ultimo nivel
            RS.Open "Select * from ctaagrupadas", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then Agrupa = True
            RS.Close
        End If
    End If
    
    'Para los balances de ejercicios siguientes existe la opcion
    ' de que si la cuenta esta en el grupo gto o grupo venta, resetear el importe a 0
    
    Resetea6y7 = CDate("01/" & cmbFecha(0).ListIndex + 1 & "/" & txtAno(0).Text) > vParam.fechafin
    
    
    
    RS.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If RS.EOF Then
        'NO hay registros a mostrar
        If vConta < 0 Then
            MsgBox "Ningun dato en los valores seleccionados.", vbExclamation
            Screen.MousePointer = vbDefault
        End If
        Agrupa = False
        CONT = -1
    Else
    
        'Voy a ver si precargamos el RScon los datos para el cierr/pyg apertura
        'Veamos si precargamos los
        SQL = ""
        If Check1(10).Value Then
            'Esta chequeado ultimo nivel
            'Veamos si tiene seleccionado alguno mas
            SQL = "1"
            For CONT = 1 To 9
                If Check1(CInt(CONT)).Value = 1 Then SQL = SQL & "1"
            Next CONT
        End If
        PreCargarCierre = Len(SQL) = 1
    
        'Mostramos el frame de resultados
        CONT = 0
        While Not RS.EOF
            CONT = CONT + 1
            RS.MoveNext
        Wend
        PB.Visible = True
        PB.Value = 0
        Me.Refresh
        
        
        'Obtengo el periodo
        InicioPeriodo = "01/" & CStr(cmbFecha(IndiceCombo).ListIndex + 1) & "/" & CInt(txtAno(IndiceCombo).Text)
        I = DiasMes(cmbFecha(IndiceCombo + 1).ListIndex + 1, CInt(txtAno(IndiceCombo + 1).Text))
        FinPeriodo = I & "/" & CStr(cmbFecha(IndiceCombo + 1).ListIndex + 1) & "/" & CInt(txtAno(IndiceCombo + 1).Text)


        'Borramos los temporales
        SQL = "DELETE from tmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute SQL
        
        
        'Nuevo  13 Enero 2005
        ' Pondremos el frame a disabled, veremos el boton de cancelar
        ' y dejaremos k lo pulse
        ' Si lo pulsa cancelaremos y no saldremos
        PulsadoCancelar = False
        Me.cmdCancelarAccion.Visible = Legalizacion = ""
        HanPulsadoSalir = False
        Me.Refresh
        
                                                           ' antes ejercicioscerrados, ahora false
        If PreCargarCierre Then PrecargaPerdidasyGanancias False, FechaIncioEjercicio, FechaFinEjercicio, QuitarSaldos2
        
        
        'Dim t1 As Single
        C1 = 0
        RS.MoveFirst
        't1 = Timer
        While Not RS.EOF
                                                                                                                                       ' antes ejercicioscerrados, ahora false
            CargaBalanceNuevo RS.Fields(0), RS.Fields(1), Apertura, InicioPeriodo, FinPeriodo, FechaIncioEjercicio, FechaFinEjercicio, False, QuitarSaldos2, vConta, False, Resetea6y7, CBool(PreCargarCierre)
            
            PB.Value = Round((C1 / CONT), 3) * 1000
            PB.Refresh
            DoEvents
            If PulsadoCancelar Then RS.MoveLast
            'Siguiente cta
            C1 = C1 + 1
            RS.MoveNext
        Wend
        
        If PreCargarCierre Then CerrarPrecargaPerdidasyGanancias
        
        'Reestablecemos
        PonleFoco cmdCancelar
        Me.cmdCancelar.Visible = True
        Me.cmdCancelarAccion.Visible = False
        HanPulsadoSalir = True
        If PulsadoCancelar Then
            RS.Close
            Screen.MousePointer = vbDefault
            PB.Visible = False
            Exit Sub
        End If
        
    End If
    RS.Close
    
    
    'Siguientes subniveles, si es ue los ha pedido
    LenNivelCalculado = 0
    For I = LenPrimerNivelCalculado - 1 To 1 Step -1
        If vConta = -1 Then
            'Balance normal
            If Check1(I).Value = 1 Then LenNivelCalculado = DigitosNivel(I)
            
        Else
'                'Balance consolidado
'                If Me.ChkConso(i).Value = 1 Then LenPrimerNivelCalculado = DigitosNivel(i)
        End If
    
        If LenNivelCalculado <> 0 Then
            SQL = "insert into tmpbalancesumas (codusu,cta,nomcta,aperturaD,aperturaH,acumAntD,acumAntH,acumPerD,acumPerH,TotalD,TotalH) "
            SQL = SQL & " select " & vUsu.Codigo & ",substring(line.cta,1," & LenNivelCalculado & ") as codmacta,coalesce(nommacta,'ERROR##') nommacta,"
            SQL = SQL & " sum(coalesce(aperturad,0)) aperd,  sum(coalesce(aperturah,0)) aperh, sum(coalesce(acumAntD,0)) acumantd, sum(coalesce(acumAntH,0)) acumanth,"
            SQL = SQL & " sum(coalesce(acumperd,0)) acumperd,  sum(coalesce(acumperh,0)) acumperh, sum(coalesce(totalD,0)) totald, sum(coalesce(totalH,0)) totalh from "
            If vConta >= 0 Then SQL = SQL & "ariconta" & vConta & "."
            SQL = SQL & "tmpbalancesumas line LEFT JOIN "
            If vConta >= 0 Then SQL = SQL & "ariconta" & vConta & "."
            SQL = SQL & "cuentas On cuentas.codmacta = substring(line.cta,1," & LenNivelCalculado & ") "
            SQL = SQL & " where line.codusu = " & vUsu.Codigo & " and length(line.cta) = " & LenPrimerNivelCalculado
            SQL = SQL & " group by 1,2 "
            SQL = SQL & " order by 1,2 "
            
            Conn.Execute SQL
            
            LenNivelCalculado = 0
        End If
    Next I
    
    
    
    
    
    
    'Ninguna entrada
    If CONT <= 0 Then Exit Sub
    
    'Realizar agrupacion
    If Agrupa Then
        PB.Value = 0
        RS.Open "Select count(*) from ctaagrupadas", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CONT = DBLet(RS.Fields(0), "N")
        RS.Close
        If CONT > 0 Then
            SQL = "Select ctaagrupadas.codmacta,nommacta from ctaagrupadas,cuentas where ctaagrupadas.codmacta =cuentas.codmacta "
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 0
            While Not RS.EOF
                If AgrupacionCtasBalance(RS.Fields(0), RS.Fields(1)) Then
                    I = I + 1
                    PB.Value = Round((I / CONT), 3) * 1000
                    RS.MoveNext
                Else
                    RS.Close
                    Exit Sub
                End If
            Wend
        End If
    End If
    
    
    
    
    'Quitamos progress
    PB.Value = 0
    PB.Visible = False
    Me.Refresh
    
    
    '--------------------
    'Balance consolidado
    If vConta >= 0 Then
        
        SQL = "Select nomempre from Usuarios.empresasariconta where codempre =" & vConta
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        cad = ""
        If Not RS.EOF Then cad = DBLet(RS.Fields(0))
        RS.Close
        If cad = "" Then
            MsgBox "Error leyendo datos empresa: Codempre=" & vConta
            Exit Sub
        End If
        
        SQL = "Select count(*) from tmpbalancesumas where codusu = " & vUsu.Codigo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        CONT = DBLet(RS.Fields(0), "N")
        RS.Close
    
        
        SQL = "Select * from tmpbalancesumas where codusu = " & vUsu.Codigo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        I = 0
        PB.Value = 0
        Me.Refresh
        SQL = "INSERT INTO Usuarios.ztmpbalanceconsolidado (codempre, nomempre, codusu, cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES ("
        SQL = SQL & vConta & ",'" & cad & "',"
        While Not RS.EOF
            PB.Value = Round((I / CONT), 3) * 1000
            BACKUP_Tabla RS, Cade
            Cade = Mid(Cade, 2)
            Cade = SQL & Cade
            Conn.Execute Cade
            'Sig
            RS.MoveNext
            I = I + 1
        Wend
        RS.Close
        'Ponemos CONT=0 para k no entre en el if de abajo
        CONT = 0
    End If
    
    Set RS = Nothing
    
    'Si hay datos los mostramos
    If CONT > 0 Then
        'Las fechas
        SQL = "Fechas= ""Desde " & cmbFecha(0).ListIndex + 1 & "/" & txtAno(0).Text & "   hasta "
        SQL = SQL & cmbFecha(1).ListIndex + 1 & "/" & txtAno(1).Text & """|"
        'Si tiene desde hasta codcuenta
        cad = ""
        If txtCta(6).Text <> "" Then cad = cad & "Desde " & txtCta(6).Text & " - " & txtNCta(6).Tag
        If txtCta(7).Text <> "" Then
            If cad <> "" Then
                cad = cad & "    h"
            Else
                cad = "H"
            End If
            cad = cad & "asta " & txtCta(7).Text & " - " & txtNCta(7).Tag
        End If
        If cad = "" Then cad = " "
        SQL = SQL & "Cuenta= """ & cad & """|"
        
        'Fecha de impresion
        SQL = SQL & "FechaImp= """ & txtFecha(7).Text & """|"
        
        
        'Salto
        If Combo2.ListIndex >= 0 Then
            SQL = SQL & "Salto= " & Combo2.ItemData(Combo2.ListIndex) & "|"
            Else
            SQL = SQL & "Salto= 11|"
        End If
        
        'Titulo
        txtTitulo(0).Text = Trim(txtTitulo(0).Text)
        If txtTitulo(0).Text = "" Then
            cad = "Balance de sumas y saldos"
        Else
            cad = txtTitulo(0).Text
        End If
        SQL = SQL & "Titulo= """ & cad & """|"
        
        'Numero de p�gina
        If txtPag2(0).Text = "" Then
            I = 1
        Else
            I = Val(txtPag2(0).Text)
        End If
        If I > 0 Then I = I - 1
        
        cad = "NumPag= " & I & "|"
        SQL = SQL & cad
        
        
        '------------------------------
        'Numero de niveles
        'Para cada nivel marcado veremos si tiene cuentas en la tmp
        CONT = 0
        UltimoNivel = 0
        For I = 1 To 10
            If Check1(I).Visible Then
'                If Check2(I).Value = 1 Then Cont = Cont + 1
                If Check1(I).Value = 1 Then
                    If I = 10 Then
                        cad = vEmpresa.DigitosUltimoNivel
                    Else
                        cad = CStr(DigitosNivel(I))
                    End If
                    If TieneCuentasEnTmpBalance(cad) Then
                        CONT = CONT + 1
                        UltimoNivel = CByte(cad)
                    End If
                End If
            End If
        Next I
        cad = "numeroniveles= " & CONT & "|"
        SQL = SQL & cad
        'Otro parametro mas
        cad = "vUltimoNivel= " & UltimoNivel & "|"
        SQL = SQL & cad
    End If
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    cad = RecuperaValor(Lista, L)
    If cad <> "" Then
        I = Val(cad)
        With cmbFecha(I)
            .Clear
            For CONT = 1 To 12
                RC = "25/" & CONT & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next CONT
        End With
    End If
    L = L + 1
Loop Until cad = ""
End Sub


Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function

Private Function ComparaFechasCombos(Indice1 As Integer, Indice2 As Integer, InCombo1 As Integer, InCombo2 As Integer) As Boolean
    ComparaFechasCombos = False
    If txtAno(Indice1).Text <> "" And txtAno(Indice2).Text <> "" Then
        If Val(txtAno(Indice1).Text) > Val(txtAno(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        Else
            If Val(txtAno(Indice1).Text) = Val(txtAno(Indice2).Text) Then
                If Me.cmbFecha(InCombo1).ListIndex > Me.cmbFecha(InCombo2).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    ComparaFechasCombos = True
End Function


'Siempre k la fecha no este en fecha siguiente
Private Function HayAsientoCierre(Mes As Byte, Anyo As Integer, Optional Contabilidad As String) As Boolean
Dim C As String
    HayAsientoCierre = False
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierre = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            C = "Select count(*) from " & Contabilidad
            C = C & " hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then
                    If RS.Fields(0) > 0 Then HayAsientoCierre = True
                End If
            End If
            RS.Close
        End If
    End If
End Function



Private Function TieneCuentasEnTmpBalance(DigitosNivel As String) As Boolean
Dim RS As ADODB.Recordset
Dim C As String

    Set RS = New ADODB.Recordset
    TieneCuentasEnTmpBalance = False
    C = Mid("__________", 1, CInt(DigitosNivel))
    C = "Select count(*) from tmpbalancesumas  where cta like '" & C & "'"
    C = C & " AND codusu = " & vUsu.Codigo
    RS.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then TieneCuentasEnTmpBalance = True
        End If
    End If
    RS.Close
End Function

Private Sub PonerNiveles()
Dim I As Integer
Dim J As Integer


    Frame2.Visible = True
    Combo2.Clear
    Check1(10).Visible = True
    For I = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(I)
        cad = "Digitos: " & J
        Check1(I).Visible = True
        Me.Check1(I).Caption = cad
        
        Combo2.AddItem "Nivel :   " & I
        Combo2.ItemData(Combo2.NewIndex) = J
    Next I
    For I = vEmpresa.numnivel To 9
        Check1(I).Visible = False
    Next I
    
    
End Sub

Private Sub HacerBalanceInicio()
    
    
        'Numero de niveles
        'Para cada nivel marcado veremos si tiene cuentas en la tmp
        RC = ""
        For I = 1 To 10
            SQL = "0"
            If Check1(I).Visible Then
                If Check1(I).Value = 1 Then SQL = "1"
            End If
            RC = RC & SQL
        Next I
    
        
        'Borramos los temporales
        SQL = "DELETE from tmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute SQL
    
        'Precargamos el cierre
        PrecargaApertura  'Carga en ur RS la apertura
    
        CONT = 1
        If Not CargaBalanceInicioEjercicio(RC) Then CONT = 0
        CerrarPrecargaPerdidasyGanancias
        If CONT = 0 Then Exit Sub
                
        SQL = "Titulo= ""Balance inicio ejercicio""|"
        SQL = SQL & "NumPag= 0|"
        
        
        '------------------------------
        'Numero de niveles
        'Para cada nivel marcado veremos si tiene cuentas en la tmp
        CONT = 0
        For I = 1 To 10
            If Check1(I).Visible Then
'                If Check2(I).Value = 1 Then Cont = Cont + 1
                If Check1(I).Value = 1 Then
                    If I = 10 Then
                        cad = vEmpresa.DigitosUltimoNivel
                    Else
                        cad = CStr(DigitosNivel(I))
                    End If
                    If TieneCuentasEnTmpBalance(cad) Then CONT = CONT + 1
                End If
            End If
        Next I
        cad = "numeroniveles= " & CONT & "|"
        SQL = SQL & cad
        
        
        'Fecha de impresion
        SQL = SQL & "FechaImp= """ & txtFecha(7).Text & """|"
        
        
        'Remarcar
        If Combo2.ListIndex >= 0 Then
            SQL = SQL & "Salto= " & Combo2.ItemData(Combo2.ListIndex) & "|"
            Else
            SQL = SQL & "Salto= 11|"
        End If

        
End Sub


Private Sub txtPag2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtTitulo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
