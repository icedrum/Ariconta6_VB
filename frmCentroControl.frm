VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCentroControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "frmCentroControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRenumFRAPRO 
      Height          =   4965
      Left            =   0
      TabIndex        =   68
      Top             =   -60
      Width           =   7155
      Begin VB.TextBox txtRenumFrapro 
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
         Left            =   5220
         TabIndex        =   94
         Top             =   2730
         Width           =   1305
      End
      Begin VB.TextBox txtRenumFrapro 
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
         Left            =   1980
         TabIndex        =   93
         Top             =   2730
         Width           =   1275
      End
      Begin VB.TextBox Text3 
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   2040
         Width           =   3975
      End
      Begin VB.CommandButton cmdRenumFra 
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
         Left            =   4470
         TabIndex        =   70
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtRenumFrapro 
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
         Left            =   1620
         TabIndex        =   69
         Top             =   2040
         Width           =   855
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
         Index           =   5
         Left            =   5670
         TabIndex        =   71
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   330
         TabIndex        =   76
         Top             =   660
         Width           =   6075
         Begin VB.OptionButton optFrapro 
            Caption         =   "Siguiente"
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
            Left            =   3300
            TabIndex        =   78
            Top             =   300
            Width           =   1395
         End
         Begin VB.OptionButton optFrapro 
            Caption         =   "Actual"
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
            Left            =   630
            TabIndex        =   77
            Top             =   300
            Value           =   -1  'True
            Width           =   1395
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   2
         Left            =   6570
         TabIndex        =   83
         Top             =   270
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
      Begin VB.Label Label11 
         Caption         =   "Primer registro"
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
         Left            =   3510
         TabIndex        =   95
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "frmCentroControl.frx":000C
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   1320
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label LabelIndF 
         Caption         =   "Label1"
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
         Left            =   360
         TabIndex        =   75
         Top             =   3900
         Width           =   3255
      End
      Begin VB.Label LabelIndF 
         Caption         =   "Label1"
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
         Left            =   360
         TabIndex        =   74
         Top             =   3540
         Width           =   3255
      End
      Begin VB.Label Label11 
         Caption         =   "Serie"
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
         Left            =   360
         TabIndex        =   73
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Index           =   0
         Left            =   360
         TabIndex        =   72
         Top             =   2760
         Width           =   1575
      End
   End
   Begin VB.Frame FrameMovCtas 
      Height          =   5205
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   7785
      Begin VB.CheckBox chkActualizarTesoreria 
         Caption         =   "Actualizar cobros/pagos"
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
         Left            =   360
         TabIndex        =   5
         Top             =   4140
         Width           =   3015
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   3750
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3540
         Width           =   1545
      End
      Begin VB.CommandButton cmdMovercta 
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
         Left            =   5010
         TabIndex        =   6
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox DtxtCta 
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1410
         Width           =   3975
      End
      Begin VB.TextBox txtcta 
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
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1410
         Width           =   1575
      End
      Begin VB.TextBox DtxtCta 
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   930
         Width           =   3975
      End
      Begin VB.TextBox txtcta 
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
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   930
         Width           =   1575
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2190
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   1
         Left            =   6270
         TabIndex        =   7
         Top             =   4560
         Width           =   1095
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   0
         Left            =   7110
         TabIndex        =   79
         Top             =   270
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
         Index           =   1
         Left            =   300
         TabIndex        =   82
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   81
         Top             =   1890
         Width           =   960
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   3480
         Picture         =   "frmCentroControl.frx":0097
         Top             =   3540
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Bloquear cuenta de ORIGEN"
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
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   22
         Top             =   3570
         Width           =   2850
      End
      Begin VB.Label Label16 
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
         Left            =   300
         TabIndex        =   21
         Top             =   4620
         Width           =   4065
      End
      Begin VB.Label Label2 
         Caption         =   "Mover cuentas "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   0
         Left            =   330
         TabIndex        =   20
         Top             =   60
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ORIGEN"
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
         Height          =   240
         Index           =   8
         Left            =   450
         TabIndex        =   19
         Top             =   930
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DESTINO"
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
         Height          =   240
         Index           =   9
         Left            =   450
         TabIndex        =   18
         Top             =   1410
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   17
         Top             =   2640
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   16
         Top             =   2190
         Width           =   600
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   1
         Left            =   1410
         Top             =   1410
         Width           =   240
      End
      Begin VB.Image imgCta 
         Height          =   240
         Index           =   0
         Left            =   1410
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1410
         Picture         =   "frmCentroControl.frx":0122
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1410
         Picture         =   "frmCentroControl.frx":01AD
         Top             =   2190
         Width           =   240
      End
   End
   Begin VB.Frame FrameCeros 
      Height          =   3975
      Left            =   120
      TabIndex        =   45
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   2760
         TabIndex        =   90
         Text            =   "Text2"
         Top             =   2190
         Width           =   765
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   2760
         Visible         =   0   'False
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   2760
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   1710
         Width           =   765
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   2760
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   1230
         Width           =   765
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   2760
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   750
         Width           =   765
      End
      Begin VB.CommandButton cmdCeros 
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
         Left            =   2880
         TabIndex        =   47
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   3
         Left            =   4080
         TabIndex        =   46
         Top             =   3360
         Width           =   1095
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   3
         Left            =   4890
         TabIndex        =   84
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
      Begin VB.Label Label4 
         Caption         =   "Dígitos resultante"
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
         Index           =   0
         Left            =   360
         TabIndex        =   91
         Top             =   2235
         Width           =   2295
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   85
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   2520
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "Dígitos nivel anterior"
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
         Index           =   13
         Left            =   360
         TabIndex        =   50
         Top             =   1755
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Dígitos último nivel"
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
         Left            =   360
         TabIndex        =   49
         Top             =   1230
         Width           =   2115
      End
      Begin VB.Label Label4 
         Caption         =   "Nº Niveles"
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
         Index           =   5
         Left            =   360
         TabIndex        =   48
         Top             =   750
         Width           =   1095
      End
   End
   Begin VB.Frame frameNuevaEmpresa 
      Height          =   5955
      Left            =   0
      TabIndex        =   23
      Top             =   -60
      Width           =   7425
      Begin VB.CheckBox Check1 
         Caption         =   "Formas de pago"
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
         Left            =   3690
         TabIndex        =   35
         Top             =   4410
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   2
         Left            =   5160
         TabIndex        =   37
         Top             =   5370
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1770
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1890
         Width           =   1435
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
         Index           =   0
         Left            =   1770
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   630
         Width           =   4815
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
         Index           =   1
         Left            =   1770
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   1050
         Width           =   2085
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
         Index           =   2
         Left            =   1770
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   1470
         Width           =   825
      End
      Begin VB.CommandButton cmdNuevaEmpresa 
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
         Left            =   3840
         TabIndex        =   36
         Top             =   5370
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia plan contable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   3300
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia conceptos"
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
         Left            =   240
         TabIndex        =   30
         Top             =   3690
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia diarios"
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
         Left            =   240
         TabIndex        =   32
         Top             =   4050
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia Tipos IVA"
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
         Left            =   3690
         TabIndex        =   33
         Top             =   4050
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Asientos predefinidos"
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
         Left            =   3690
         TabIndex        =   29
         Top             =   3330
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia centros de coste"
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
         Left            =   3690
         TabIndex        =   31
         Top             =   3690
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Copia configuracion balances"
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
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   34
         Top             =   4410
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   1
         Left            =   6780
         TabIndex        =   80
         Top             =   600
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
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1500
         Picture         =   "frmCentroControl.frx":0238
         Top             =   1890
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre empresa"
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
         Index           =   11
         Left            =   270
         TabIndex        =   44
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre corto"
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
         Index           =   10
         Left            =   270
         TabIndex        =   43
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Número empresa"
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
         Index           =   7
         Left            =   270
         TabIndex        =   42
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Insertar datos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   5
         Left            =   240
         TabIndex        =   41
         Top             =   2610
         Width           =   2775
      End
      Begin VB.Label Label4 
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
         Height          =   285
         Index           =   6
         Left            =   270
         TabIndex        =   40
         Top             =   1890
         Width           =   1305
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   39
         Top             =   2760
         Width           =   5715
      End
      Begin VB.Label Label6 
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   4980
         Width           =   2835
      End
   End
   Begin VB.Frame FrDesbloq 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton cmdDesbloq 
         Caption         =   "Desbloquear"
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
         Left            =   4800
         TabIndex        =   12
         Top             =   3720
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Diario"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Nº Asiento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Obser"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   0
         Left            =   6360
         TabIndex        =   9
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   7170
         Picture         =   "frmCentroControl.frx":02C3
         ToolTipText     =   "Desmarcar todos"
         Top             =   450
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   6870
         Picture         =   "frmCentroControl.frx":0CC5
         ToolTipText     =   "Marcar todos"
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Desbloquear asientos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame FrameCambioIVA 
      Height          =   4575
      Left            =   60
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtFecha 
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
         Left            =   3570
         TabIndex        =   60
         Text            =   "00/00/0000"
         Top             =   2700
         Width           =   1245
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   1200
         TabIndex        =   59
         Text            =   "00/00/0000"
         Top             =   2700
         Width           =   1245
      End
      Begin VB.CommandButton cmdIVA 
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
         Left            =   2520
         TabIndex        =   61
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   4
         Left            =   3720
         TabIndex        =   62
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDescIVA 
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
         Left            =   1200
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   1860
         Width           =   3645
      End
      Begin VB.TextBox txtIVA 
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
         Left            =   480
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   1860
         Width           =   645
      End
      Begin VB.TextBox txtDescIVA 
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
         Left            =   1200
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   900
         Width           =   3645
      End
      Begin VB.TextBox txtIVA 
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
         Left            =   480
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   900
         Width           =   645
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Index           =   4
         Left            =   4500
         TabIndex        =   89
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
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   88
         Top             =   2370
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Destino"
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
         Index           =   4
         Left            =   180
         TabIndex        =   87
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Origen"
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
         Index           =   3
         Left            =   150
         TabIndex        =   86
         Top             =   540
         Width           =   780
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   3240
         Picture         =   "frmCentroControl.frx":7517
         Top             =   2730
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   18
         Left            =   2610
         TabIndex        =   67
         Top             =   2730
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   870
         Picture         =   "frmCentroControl.frx":75A2
         Top             =   2730
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   16
         Left            =   180
         TabIndex        =   66
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label lblIVA 
         Caption         =   "Label1"
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
         TabIndex        =   65
         Top             =   3510
         Width           =   4695
      End
      Begin VB.Image imgIVA 
         Height          =   240
         Index           =   1
         Left            =   1170
         Picture         =   "frmCentroControl.frx":762D
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image imgIVA 
         Height          =   240
         Index           =   0
         Left            =   1170
         Picture         =   "frmCentroControl.frx":802F
         Top             =   570
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCentroControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IdPrograma As Integer

Public Opcion As Byte
    '0.- Desbloquear asientos
    '1.- Mover ctas
    '2.- Crear empresa nueva
    '3.- Aumento ce deros
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmI As frmIVA
Attribute frmI.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Dim i As Integer
Dim SQL As String
Dim PrimeraVez As Boolean


Dim TablaAnt As String
Dim Tam2 As Long
Dim Tamanyo As Long
Dim NumTablas As Integer
Dim ParaElLog As String
Dim Insert As String
Dim Campos()

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkActualizarTesoreria_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCompruebaContab_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkSALTO_numerofactura_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkUpdateNumDocum_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCeros_Click()
Dim B As Boolean

    B = False

    If UsuariosConectados("Desbloqueando asientos" & vbCrLf, True) Then Exit Sub
    
    SQL = "Este programa aumentará el número de dígitos a último nivel" & vbCrLf
    SQL = SQL & vbCrLf & vbCrLf & "Debería hacer una copia de seguridad." & vbCrLf & vbCrLf
    SQL = SQL & "             ¿ Desea continuar?    "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    
    
    
    
        
    SQL = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(SQL) <> "ARIADNA" Then
        If SQL <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    If ComprobarOk(CByte(Text2(5).Text)) Then
        Label3(2).Caption = ""
        Label3(2).visible = True
        pb1.Value = 0
        Me.pb1.Max = 1000
        Me.pb1.visible = True
        B = HacerInsercionDigitoContable
        pb1.visible = False
        Label3(2).visible = False
        
        'Insertamos el LOG
        ParaElLog = "Nº nivel: " & Text2(3).Text & vbCrLf
        ParaElLog = ParaElLog & "Digitos último nivel: " & Text2(4).Text & vbCrLf
        ParaElLog = ParaElLog & "Digitos nivel anterior: " & Text2(5).Text & vbCrLf
        ParaElLog = "Aumentar CERO (" & CStr(B) & ")" & vbCrLf & ParaElLog
        vLog.Insertar 22, vUsu, ParaElLog
        ParaElLog = ""
        
    End If
    Screen.MousePointer = vbDefault
    
    If B Then Unload Me
    
End Sub

Private Sub cmdDesbloq_Click()
    SQL = "Seleccione algún asiento para desbloquear"
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            SQL = ""
            Exit For
        End If
    Next i
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    If UsuariosConectados("Desbloqueando asientos" & vbCrLf, True) Then Exit Sub
        
    SQL = InputBox("Escriba password de seguridad", "CLAVE")
    If UCase(SQL) <> "ARIADNA" Then
        If SQL <> "" Then MsgBox "Clave incorrecta", vbExclamation
        Exit Sub
    End If
    
    ParaElLog = ""
    For i = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(i).Checked Then
            SQL = "UPDATE hcabapu SET bloqactu = 0 WHERE numdiari =" & ListView1.ListItems(i).Text
            SQL = SQL & " AND fechaent = '" & Format(ListView1.ListItems(i).SubItems(1), FormatoFecha) & "'"
            SQL = SQL & " AND numasien = " & Val(ListView1.ListItems(i).SubItems(2))
            
            EjecutaSQL SQL
            'Para el LOG
            SQL = ListView1.ListItems(i).Text & "," & ListView1.ListItems(i).SubItems(1) & "," & Val(ListView1.ListItems(i).SubItems(2))
            
            ParaElLog = ParaElLog & ", [" & SQL & "]"
            ListView1.ListItems.Remove i
            
        End If
    Next i
    'Insertamos el LOG
    ParaElLog = "DESBLOQUEAR" & vbCrLf & ParaElLog
    vLog.Insertar 16, vUsu, ParaElLog
    ParaElLog = ""
    
    'Si nO queda ninguno cierro ventana
    If ListView1.ListItems.Count = 0 Then Unload Me
    
End Sub

Private Sub cmdIVA_Click()
Dim B As Boolean

        If txtIVA(0).Text = "" Or txtIVA(1).Text = "" Then
            MsgBox "IVA origen y destino requeridos", vbExclamation
            Exit Sub
        End If

        If txtIVA(0).Text = txtIVA(1).Text Then
            MsgBox "IVA origen no puede ser igual al IVA destino", vbExclamation
            Exit Sub
        End If
        SQL = "Debería tener una copia de seguridad." & vbCrLf & "El proceso puede tardar mucho tiempo" & vbCrLf
        SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub


        If UsuariosConectados("Cambiando IVA" & vbCrLf, True) Then Exit Sub


        SQL = InputBox("Password de seguridad")
        If UCase(SQL) <> "ARIADNA" Then Exit Sub
    
        Screen.MousePointer = vbHourglass
        B = HacerCambioIVA
        lblIVA.Caption = ""
        Screen.MousePointer = vbDefault
        If B Then
            SQL = "Proceso finalizado con éxito." & vbCrLf & vbCrLf & "¿Desea realizar otro cambio?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then
                Unload Me
            Else
                Limpiar Me
                PonleFoco txtIVA(0)
            End If
        End If
End Sub

Private Sub cmdMovercta_Click()

         'Hacemos lo que tengamos que hacer
        If txtCta(0).Text = "" Or txtCta(1).Text = "" Then
            MsgBox "Ponga cuentas contables", vbExclamation
            Exit Sub
        End If
        If txtCta(0).Text = txtCta(1).Text Then
            MsgBox "Misma cuenta origen destino", vbExclamation
            Exit Sub
        End If
        
        If txtFecha(0).Text = "" Then
            MsgBox "Ponga la fecha ""Desde""", vbExclamation
            Exit Sub
        End If
        
        
        'Diciemnre 2012
        'Pequeñas comprobaciones
        'Si tiene pagos cobros Preguntara
        If vEmpresa.TieneTesoreria Then
            If Me.chkActualizarTesoreria.Value = 1 Then
                
                    i = 0
                    SQL = DevuelveDesdeBD("count(*)", "cobros", "codmacta", txtCta(0).Text, "T")
                    If Val(SQL) > 0 Then Insert = "cobros"
                    SQL = DevuelveDesdeBD("count(*)", "pagos", "codmacta", txtCta(0).Text, "T")
                    If Val(SQL) > 0 Then Insert = Insert & " pagos"
                    If Insert <> "" Then
                        SQL = "Existen " & Insert & " relacionados con la cuenta. Continuar?"
                        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                            
                    End If
      
            End If
        End If
        Insert = ""
        
        SQL = "Deberia tener una copia de seguridad." & vbCrLf & "El proceso puede tardar mucho tiempo" & vbCrLf
        SQL = SQL & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub

        SQL = InputBox("Password de seguridad")
        If UCase(SQL) <> "ARIADNA" Then Exit Sub
    
        If HacerCambioCuenta Then Unload Me
         Label16.Caption = ""
End Sub

Private Sub cmdNuevaEmpresa_Click(Index As Integer)

Dim Ok As Boolean
Dim T As TextBox


    For Each T In Text2
        If T.visible Then
            T = Trim(T)
            If T = "" Then
                
                MsgBox "Todos los campos obligatorios", vbExclamation
                Exit Sub
            End If
        End If
    Next



    If Not IsNumeric(Text2(2).Text) Then
        MsgBox "Número de empresa tiene que ser numérico, obviamente", vbExclamation
        Exit Sub
    End If
    
    If Not IsDate(txtFecha(3).Text) Then
        MsgBox "Fecha inicio incorrecta", vbExclamation
        Exit Sub
    End If
    

    
    
    
    
    'Si marca el asipre tiene k tener marcados cuetas, y tal y tal
     Tam2 = Check1(0).Value + Check1(1).Value + Check1(2).Value + Check1(5).Value
     If Check1(4).Value = 1 Then
        If Tam2 <> 4 Then
            MsgBox "Si marca asientos predefinidos tiene que marcar cuentas, diarios, conceptos y centros de coste.", vbExclamation
            Exit Sub
        End If
    End If
    
    'Si marca IVA tiene que llevarse el plan contable, ya que los tipos de IVA estan
    'asociados a cuentas contables
    If Check1(3).Value Then
        If Check1(0).Value = 0 Then
            MsgBox "Los tipos de IVA estan asociados a cuentas contables de ultimo nivel.", vbExclamation
            Exit Sub
        End If
    End If
    
    SQL = "Va a generar una nueva empresa: " & Text2(0).Text
    SQL = SQL & vbCrLf & "Desea continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Ok = False
    Label6.Caption = "Generando estructura de BD"
    Me.Refresh
    
    Ok = GeneracionNuevaBD
    
    Label6.Caption = ""
    Screen.MousePointer = vbDefault
    If Ok Then
        MsgBox "Proceso finalizado correctamente.", vbExclamation
        Unload Me
    End If

End Sub



Private Sub cmdRenumFra_Click()
Dim Ok As Boolean


    If HayFacturasSinNroAsiento Then Exit Sub
    
    If Not ComprobarFRAPROContabilizadas(txtRenumFrapro(0), True) Then Exit Sub


    Me.LabelIndF(0).Caption = ""
    Me.LabelIndF(1).Caption = ""
    
    If MsgBox("Debería hacer una copia de seguridad." & vbCrLf & vbCrLf & vbCrLf & "El proceso puede durar muchísimo tiempo. ¿Desea continuar igualmente?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        
    If UsuariosConectados("Renumerar nºReg. en factura proveedor" & vbCrLf, True) Then Exit Sub


    SQL = InputBox("Password de seguridad")
    If UCase(SQL) <> "ARIADNA" Then Exit Sub
    
        
        
    'OK---------------------------------
    'A renumerar
    Screen.MousePointer = vbHourglass
    cmdRenumFra.Enabled = False
    
    Ok = HacerRenumeracionFacturas
    
    If Ok Then

        'Insertamos el LOG
'            ParaElLog = "actual"
'            If Me.optFrapro(1).Value Then ParaElLog = "siguiente"
'
'            ParaElLog = "Ejercicio " & ParaElLog & vbCrLf
'
'            ParaElLog = ParaElLog & "Nº registro " & txtRenumFrapro(2).Text & vbCrLf
'        ParaElLog = "Renumerar facturas proveedor." & vbCrLf & ParaElLog
'
'        vLog.Insertar 21, vUsu, ParaElLog
'
        
        ParaElLog = String(40, "*") & vbCrLf
        ParaElLog = ParaElLog & ParaElLog & ParaElLog
        ParaElLog = ParaElLog & vbCrLf & vbCrLf & "Compruebe el contador de facturas de proveedor" & vbCrLf & vbCrLf & vbCrLf & ParaElLog
        MsgBox ParaElLog, vbExclamation
        

        ParaElLog = ""
    End If
    
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
    Me.LabelIndF(0).Caption = ""
    Me.LabelIndF(1).Caption = ""
    cmdRenumFra.Enabled = True
    
    
    
    
    
    If Ok Then Unload Me
        
End Sub

Private Function HayFacturasSinNroAsiento() As Boolean
Dim SQL As String
Dim HayReg As Boolean
Dim Rs As ADODB.Recordset
Dim CadResult As String


    On Error GoTo eHayFacturasSinNroAsiento

    HayFacturasSinNroAsiento = False

    SQL = "select * from factpro where fecharec >= " & DBSet(txtRenumFrapro(0).Text, "F") & " and (numasien = 0 or numasien is null) "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    HayReg = False
    
    CadResult = "Las siguientes facturas no tienen número de asiento:" & vbCrLf & vbCrLf
    
    While Not Rs.EOF
        HayReg = True
        
        CadResult = CadResult & "Registro " & DBLet(Rs!Numregis, "N") & " de " & DBLet(Rs!fecharec, "F") & vbCrLf
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If HayReg Then MsgBox CadResult & vbCrLf & "Revise.", vbExclamation
    
    HayFacturasSinNroAsiento = HayReg
    Exit Function
    
eHayFacturasSinNroAsiento:
    MuestraError Err.Number, "Comprobacion Facturas sin número asiento", Err.Description
End Function


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 0
            CargarAsientosBloqueados
        Case 1
            txtFecha(0).Text = Format(vParam.fechaini, "dd/mm/yyyy")
            txtFecha(1).Text = Format(Now, "dd/mm/yyyy")
        Case 2
            SugerirValoresNuevaEmpresa
        Case 3
            'Cargo los valores
            Text2(3).Text = vEmpresa.numnivel
            Text2(4).Text = vEmpresa.DigitosUltimoNivel
            Text2(6).Text = vEmpresa.DigitosUltimoNivel + 1
            
                        
            i = vEmpresa.numnivel
            i = i - 1
            i = DigitosNivel(i)
            Text2(5).Text = i
            
            
            If vEmpresa.DigitosUltimoNivel = 10 Then
                MsgBox "La contabilidad ya está a 10 dígitos contables.", vbExclamation
                cmdCeros.Enabled = False
            End If
            
        Case 4
            PonleFoco txtIVA(0)
        Case 5
            '--
            'If vParam.CodiNume = 1 Then Me.chkUpdateNumDocum.Value = 1
            '++
            txtRenumFrapro(0).Text = vParam.fechaini
            txtRenumFrapro(1).Text = "1"
            Text3(0).Text = DevuelveDesdeBD("nomregis", "contadores", "tiporegi", txtRenumFrapro(1).Text, "T")
            
            
            PonerValoresCalculos
            
        End Select
    End If
End Sub

Private Sub PonerValoresCalculos()
Dim Mc As Contadores
Dim vContador As String
Dim DesdeInicio As Boolean
Dim Contador As Long
    DesdeInicio = ((txtRenumFrapro(0).Text = vParam.fechaini) Or (txtRenumFrapro(0).Text = DateAdd("yyyy", 1, vParam.fechaini)))

    If DesdeInicio Then

        Set Mc = New Contadores
        If Mc.ConseguirContador(txtRenumFrapro(1).Text, optFrapro(0), False) = 0 Then
            If Mc.Contador >= 1000000 Then
                vContador = Mid(CStr(Format(Mc.Contador, "0000000")), 1, 2)
                txtRenumFrapro(2).Text = (vContador * 1000000) + 1
            Else
                txtRenumFrapro(2).Text = 1
            End If
        End If
    Else
        SQL = "select max(numregis) from factpro where fecharec < " & DBSet(txtRenumFrapro(1).Text, "F")
        If optFrapro(0).Value Then
            SQL = SQL & " and fecharec >= " & DBSet(vParam.fechaini, "F")
        Else
            SQL = SQL & " and fecharec >= " & DBSet(DateAdd("yyyy", 1, vParam.fechaini), "F")
        End If
        SQL = SQL & " and numserie = " & DBSet(txtRenumFrapro(1).Text, "T")
        
        txtRenumFrapro(2).Text = DevuelveValor(SQL)
        If txtRenumFrapro(2).Text = 0 Then
            txtRenumFrapro(2).Text = 1
        End If
    End If

End Sub


Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim i As Integer


    Me.Icon = frmppal.Icon

    PrimeraVez = True
    Caption = "Herramientas"
    
    Me.Icon = frmppal.Icon
    
    Limpiar Me
    
    For i = 0 To 1
        Me.imgcta(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i

    ' La Ayuda
    For i = 0 To ToolbarAyuda.Count - 1
        With Me.ToolbarAyuda(i)
            .ImageList = frmppal.ImgListComun
            .Buttons(1).Image = 26
        End With
    Next i
    FrameMovCtas.visible = False
    Me.FrDesbloq.visible = False
    frameNuevaEmpresa.visible = False
    FrameCeros.visible = False
    FrameCambioIVA.visible = False
    FrameRenumFRAPRO.visible = False
    
    
    Select Case Opcion
    Case 0
            IdPrograma = 1407
            PonerFrameVisible Me.FrDesbloq, H, W
    Case 1
            IdPrograma = 1408
            PonerFrameVisible Me.FrameMovCtas, H, W
            Me.chkActualizarTesoreria.visible = vEmpresa.TieneTesoreria
            
            Me.Caption = "Traspaso de cuentas en apuntes"
            
            
    Case 2
            IdPrograma = 107
            PonerFrameVisible frameNuevaEmpresa, H, W
            Me.Caption = "Nueva Empresa"
            
    Case 3
            IdPrograma = 1410
            PonerFrameVisible FrameCeros, H, W
            pb1.visible = False
            Me.Caption = "Aumentar dígitos contables"
    Case 4
            IdPrograma = 1411
            lblIVA.Caption = ""
            PonerFrameVisible FrameCambioIVA, H, W
            Me.Caption = "Traspaso códigos de IVA"
                
    Case 5
            IdPrograma = 1409
            PonerFrameVisible FrameRenumFRAPRO, H, W
            Me.LabelIndF(0).Caption = ""
            Me.LabelIndF(1).Caption = ""
            Me.Caption = "Renumerar registros proveedor"
            
            'No puede actualizar el campo NUMDOCUM con el numregis si no esta
'--            'marcada la opcion Numeroregisro en documento(vParam.CodiNume = 1)
            'Ultimo periodo liquidado
            If vParam.perfactu > 0 Then
                If vParam.periodos = 1 Then
                    'IVA MENSUAL
                    i = vParam.perfactu
                Else
                    i = vParam.perfactu * 3
                End If
                NumTablas = DiasMes(CByte(i), vParam.anofactu)
                
            End If
    End Select
    
    Me.Height = H
    Me.Width = W
    Me.cmdCancelar(Opcion).Cancel = True

End Sub


Private Sub PonerFrameVisible(ByRef Fr As frame, ByRef He As Integer, ByRef Wi As Integer)
    Fr.top = 30
    Fr.Left = 30
    Fr.visible = True
    He = Fr.Height + 540
    Wi = Fr.Width + 120
End Sub


Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtRenumFrapro(1).Text = RecuperaValor(CadenaSeleccion, 1)
        Text3(0).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Me.txtRenumFrapro(0).Text = vFecha
End Sub

Private Sub frmI_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
    
End Sub


Private Sub imgCheck_Click(Index As Integer)

    For i = 1 To Me.ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = Index = 1
    Next
End Sub


Private Sub CargarAsientosBloqueados()
Dim IT As ListItem
    Set miRsAux = New ADODB.Recordset
    ListView1.ListItems.Clear
    SQL = "Select * from hcabapu WHERE bloqactu=1  ORDER BY fechaent,numdiari,fechaent"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = ListView1.ListItems.Add()
        IT.Text = miRsAux!NumDiari
        IT.SubItems(1) = Format(miRsAux!FechaEnt, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!NumAsien, "00000")
        SQL = DBLet(miRsAux!obsdiari, "T") & "  "
        SQL = Mid(SQL, 1, 20)
        IT.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Me.cmdDesbloq.Enabled = ListView1.ListItems.Count > 0
    
End Sub

Private Sub imgcta_Click(Index As Integer)
   Set frmCta = New frmColCtas
   SQL = ""
   frmCta.DatosADevolverBusqueda = "0|1|"
   frmCta.Show vbModal
   Set frmCta = Nothing
   If SQL <> "" Then
        txtCta(Index).Text = RecuperaValor(SQL, 1)
        DtxtCta(Index).Text = RecuperaValor(SQL, 2)
        SQL = ""
        PonFoco txtCta(Index)
    End If
    
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Set frmC = New frmCal
    SQL = ""
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index).Text)
    frmC.Show vbModal
    If SQL <> "" Then
        txtFecha(Index).Text = SQL
        SQL = ""
        PonFoco txtFecha(Index)
    End If
    Set frmC = Nothing
    
End Sub

Private Sub imgiva_Click(Index As Integer)
    SQL = ""
    Set frmI = New frmIVA
    frmI.DatosADevolverBusqueda = "0|1|"
    frmI.Show vbModal
    Set frmI = Nothing
    If SQL <> "" Then
        txtIVA(Index).Text = RecuperaValor(SQL, 1)
        txtDescIVA(Index).Text = RecuperaValor(SQL, 2)
        If Index = 0 Then
            PonFoco txtIVA(1)
        Else
            PonFoco txtFecha(4)
        End If
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        'FECHA
        
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtRenumFrapro(0).Text <> "" Then frmF.Fecha = CDate(txtRenumFrapro(0).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtRenumFrapro(0)
    
    End Select
    

End Sub

Private Sub optFrapro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 2 Then
        If Text2(2).Text <> "" Then
            If Not IsNumeric(Text2(2).Text) Then
                Text2(2).Text = ""
            Else
                Text2(2).Text = Val(Text2(2).Text)
                If Val(Text2(2).Text) > 99 Then
                    MsgBox "De uno a 99", vbExclamation
                    Text2(2).Text = ""
                End If
            End If
            If Text2(2).Text = "" Then PonleFoco Text2(2)
        End If
    End If
End Sub

Private Sub ToolbarAyuda_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select

End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
End Sub

Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 112 Then
        HacerF1
    Else
        If KeyCode = 107 Or KeyCode = 187 Then
            KeyCode = 0
            txtCta(Index).Text = ""
            imgcta_Click Index
        End If
    End If
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        Exit Sub
    End If
    
    Select Case Index
    Case 10000   'Las que no sean obligadas de ultimo nivel
        'NO hace falta que sean de ultimo nivel
        Cta = (txtCta(Index).Text)
                                '********
        B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
        If B = 0 Then
            MsgBox "NO existe la cuenta: " & txtCta(Index).Text, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
        Else
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            If B = 1 Then
                DtxtCta(Index).Tag = ""
            Else
                DtxtCta(Index).Tag = SQL
            End If
          
        End If
    Case Else
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, SQL) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If
    End Select
End Sub


Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HacerF1
End Sub

'++
Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
            Case 1:  KEYFecha KeyAscii, 1
            Case 2:  KEYFecha KeyAscii, 2
            Case 3:  KEYFecha KeyAscii, 3
            Case 4:  KEYFecha KeyAscii, 4
            Case 5:  KEYFecha KeyAscii, 5
            Case 6:  KEYFecha KeyAscii, 6
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtiva_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYIva KeyAscii, 0
            Case 1:  KEYIva KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
End Sub

Private Sub KEYIva(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgiva_Click (Indice)
End Sub


'++

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index))
    If txtFecha(Index) = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index), vbExclamation
        txtFecha(Index).Text = ""
        PonFoco txtFecha(Index)
    End If
End Sub



Private Sub HacerF1()
    Select Case Opcion
    Case 0
        
    Case 1
        
    End Select
End Sub


'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'
'   Cambio cuenta contable
'
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------

Private Function HacerCambioCuenta() As Boolean


Dim NombreArchivo As String
Dim NF As Integer
Dim Final As String

    On Error GoTo EHacerCambioCuenta
    
     HacerCambioCuenta = False
    
    'Veamos cuantos updates hay que hacer
    'Los fijos
            'bancos    4    Dicimebre 2012. La cta ppal bancaria NO se puede cambiar y se añaden ctaingreso,ctaefectosdesc,ctagastostarj, DIC 2012
            'inmovele         3
        
        
    'Variables
            'hlinapu        2
            'asipre_lineas      2
            'factpro    1
            'factcli        1
            'factcli_lineas 1
            'factpro_lineas 1
            'presupuestos   1
    
    'Total                  16
    Tamanyo = 16
    
    
    'Si tiene tesoreria
            'scaja          2

            'cobros         3
            'pagos          3
            'shacaja        2
            'sgastfij       2
            'stransfer      1
            'stransfercob   1
            'remesas        1
            'shcocob        1
            '_________________
            '               17
    If vEmpresa.TieneTesoreria Then Tamanyo = Tamanyo + 16
    Tam2 = 0
    Label16.Caption = "Comienzo proceso"
    Label16.visible = True
    'Los que no llevan fechas
    'bancos,inmovele
    PonerTabla "bancos"
    EjecutaSQLCambio "ctagastos", ""
    EjecutaSQLCambio "ctaingreso", ""
    EjecutaSQLCambio "ctaefectosdesc", ""
    EjecutaSQLCambio "ctagastostarj", ""
    
    
    PonerTabla "inmovele"
    EjecutaSQLCambio "codmact1", ""
    EjecutaSQLCambio "codmact2", ""
    EjecutaSQLCambio "codmact3", ""
    '++
    EjecutaSQLCambio "codprove", ""
    PonerTabla "inmovele_rep"
    EjecutaSQLCambio "codmacta2", ""
    
    
    'hlinapu        2
    'asipre_lineas      2
    NombreArchivo = "hlinapu|asipre_lineas|"
    For NF = 1 To 2
        PonerTabla RecuperaValor(NombreArchivo, NF)
        Final = "fechaent"
        If NF = 3 Then Final = ""
        EjecutaSQLCambio "codmacta", Final
        EjecutaSQLCambio "ctacontr", Final
    Next NF
    
    
    'Presupuestos
    PonerTabla "presupuestos"
    EjecutaSQLCambio "codmacta", ""
    
    'factpro    1
    'factcli        1
    PonerTabla "factcli"
    EjecutaSQLCambio "codmacta", "fecfactu"
    PonerTabla "factpro"
    EjecutaSQLCambio "codmacta", "fecharec"
    
    
    
    'Lineas de facturas
    PonerTabla "Lineas fracli"
    EjecutaSQLCambioLineasFras True, "fecfactu"
    PonerTabla "Lineas frapro"
    EjecutaSQLCambioLineasFras False, "fecharec"
    
    
    'Si tiene tesoreria
    'scaja,departamento,cobros,pagos,shacaja,shcobro,sgatfij,stransfer,stransfercob
    If vEmpresa.TieneTesoreria Then
        
        PonerTabla "cobros"
        EjecutaSQLCambio "codmacta", "fecvenci"
        EjecutaSQLCambio "ctabanc1", "fecvenci"
        EjecutaSQLCambio "ctabanc2", "fecvenci"
        
        PonerTabla "gastfijos"
        EjecutaSQLCambio "ctaprevista", ""
        EjecutaSQLCambio "contrapar", ""
        PonerTabla "shcocob"
        EjecutaSQLCambio "codmacta", ""
        
        
        PonerTabla "pagos"
        EjecutaSQLCambio "codmacta", "fecefect"
        EjecutaSQLCambio "ctabanc1", "fecefect"
        EjecutaSQLCambio "ctabanc2", "fecefect"
        
        
        PonerTabla "transferencias"
        EjecutaSQLCambio "codmacta", ""
        
        PonerTabla "remesas"
        EjecutaSQLCambio "codmacta", ""
        
    End If
    
'++ lo he añadido
        PonerTabla "tipoiva"
        EjecutaSQLCambio "cuentare", ""
        EjecutaSQLCambio "cuentarr", ""
        EjecutaSQLCambio "cuentaso", ""
        EjecutaSQLCambio "cuentasr", ""
        EjecutaSQLCambio "cuentasn", ""
        
        PonerTabla "parametros"
        EjecutaSQLCambio "ctahpdeudor", ""
        EjecutaSQLCambio "ctahpacreedor", ""
    
    
    If txtFecha(2).Text <> "" Then
        SQL = "UPDATE cuentas SET fecbloq = '" & Format(txtFecha(2).Text, FormatoFecha)
        SQL = SQL & "' WHERE codmacta = '" & Me.txtCta(0).Text & "'"
        Conn.Execute SQL
    End If
    
    ParaElLog = "Origen: " & txtCta(0).Text & " " & Me.DtxtCta(0).Text & vbCrLf
    ParaElLog = ParaElLog & "Destino: " & txtCta(1).Text & " " & Me.DtxtCta(1).Text & vbCrLf & vbCrLf
    ParaElLog = ParaElLog & "Fechas: " & txtFecha(0).Text & " - " & txtFecha(1).Text & vbCrLf
    If txtFecha(2).Text <> "" Then ParaElLog = ParaElLog & "Bloqueo: " & txtFecha(2).Text
    ParaElLog = "MOVER CTAS" & vbCrLf & ParaElLog
    vLog.Insertar 20, vUsu, ParaElLog
    ParaElLog = ""
    
    
    HacerCambioCuenta = True
    Exit Function
EHacerCambioCuenta:
    MuestraError Err.Number, TablaAnt & vbCrLf & SQL
End Function

Private Sub PonerTabla(ByRef T As String)
    TablaAnt = T
    Label16.Caption = ""
    Me.Refresh
    DoEvents
End Sub

Private Function EjecutaSQLCambio(Campo As String, CampoFecha As String) As Boolean
    Tam2 = Tam2 + 1
    Label16.Caption = Campo & " - " & TablaAnt & "    (" & Tam2 & " / " & Tamanyo & ")"
    Label16.Refresh
    SQL = "UPDATE " & TablaAnt & " SET " & Campo & " = " & txtCta(1).Text & " WHERE "
    SQL = SQL & Campo & " = " & txtCta(0).Text
    'Si tiene fechas
    If CampoFecha <> "" Then
        
        SQL = SQL & " AND " & CampoFecha & " >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        If txtFecha(1).Text <> "" Then SQL = SQL & " AND " & CampoFecha & " <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    End If
    Conn.Execute SQL
End Function




Private Function EjecutaSQLCambioLineasFras(Clientes As Boolean, CampoFecha As String) As Boolean
    Tam2 = Tam2 + 1
    Label16.Caption = TablaAnt & "    (" & Tam2 & " / " & Tamanyo & ")"
    Label16.Refresh
    If Clientes Then
        SQL = "UPDATE factcli,factcli_lineas SET codmacta='" & txtCta(1).Text & "'"
        SQL = SQL & " where factcli.numserie=factcli_lineas.numserie and factcli.numfactu=factcli_lineas.numfactu and"
        SQL = SQL & " factcli.anofactu=factcli_lineas.anofactu"
    Else
        SQL = "UPDATE factpro,factpro_lineas SET codmacta='" & txtCta(1).Text & "'"
        SQL = SQL & " where factpro.numregis=factpro_lineas.numregis and"
        SQL = SQL & " factpro.anofactu = factpro_lineas.anofactu"
    End If
    'Si tiene fechas
    
    SQL = SQL & " AND " & CampoFecha & " >= '" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    If txtFecha(1).Text <> "" Then SQL = SQL & " AND " & CampoFecha & " <= '" & Format(txtFecha(1).Text, FormatoFecha) & "'"
     
    
    SQL = SQL & " AND codtbase = '" & txtCta(0).Text & "'"
    
    Conn.Execute SQL
End Function



'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'
'       NUEVA EMPRESA
'
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
Private Function GeneracionNuevaBD() As Boolean

    GeneracionNuevaBD = False
    If Not IsNumeric(Text2(2).Text) Then
        MsgBox "Nº BD debe ser campo numérico", vbExclamation
        Exit Function
    End If
    
    
    
    'Comprobamos k la clave no esta
    TablaAnt = "nomempre"
    SQL = DevuelveDesdeBD("codempre", "usuarios.empresasariconta", "codempre", Text2(2).Text, "T", TablaAnt)
    If SQL <> "" Then
        MsgBox "El codigo de empresa " & Text2(2).Text & " esta asociado a " & TablaAnt, vbExclamation
        Exit Function
    End If
        
    'Hago un SQL para que de error si no existe la BD
    SQL = "UPDATE ariconta" & Text2(2).Text & ".hcabapu SET numdiari=1 WHERE numdiari=-1"
    If EjecutaSQL(SQL) Then
        MsgBox "YA existe la BD ", vbExclamation
        Exit Function
    End If
    
    
    
    
    If Not GeneraNuevaBD Then Exit Function
    Screen.MousePointer = vbHourglass
    'Insertamos en tabla empresas
        SQL = "INSERT INTO usuarios.empresasariconta (codempre, nomempre, nomresum, Conta,Tesor) VALUES ("
        SQL = SQL & Text2(2).Text & ",'" & Text2(0).Text & "','" & Text2(1).Text
        SQL = SQL & "','ariconta" & Text2(2).Text & "'," & Abs(vEmpresa.TieneTesoreria) & ")"
        Conn.Execute SQL
    
    
   If Not CrearEstructura Then Exit Function
        
   If InsercionDatos Then GeneracionNuevaBD = True
        
    
    
    
    Screen.MousePointer = vbDefault
    
End Function



Private Function GeneraNuevaBD() As Boolean
On Error Resume Next
       GeneraNuevaBD = False
        SQL = "CREATE DATABASE conta" & Text2(2).Text
        Conn.Execute SQL
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Creando BD"
        Else
            GeneraNuevaBD = True
        End If
End Function


'--------------------------------------------------------------------
'
'                    Crear estructura BD
'
'--------------------------------------------------------------------

Private Function CrearEstructura() As Boolean
Dim ColTablas As Collection
Dim ColCreate As Collection
Dim Bucle As Integer

    CrearEstructura = False

    Set ColTablas = New Collection
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "SHOW TABLES", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        ColTablas.Add CStr(miRsAux.Fields(0))
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ya tengo todas las tablas. Ahora para cada tabla ire buscando el show create table
    Set ColCreate = New Collection
    For Tam2 = 1 To ColTablas.Count
        SQL = ColTablas.Item(Tam2)
        miRsAux.Open "SHOW CREATE TABLE " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        TablaAnt = miRsAux.Fields(1)
        ColCreate.Add SQL & "|" & TablaAnt & "|"
        miRsAux.Close
    Next
    
    'Ya tengo los create tables
    'AHora para un bucle de 10 veces
    Bucle = 1
    Do
        Tamanyo = ColCreate.Count
        For Tam2 = Tamanyo To 1 Step -1
            TablaAnt = ColCreate.Item(Tam2)
            SQL = RecuperaValor(TablaAnt, 2) 'create table...
            TablaAnt = RecuperaValor(TablaAnt, 1)
            'TEngo que añadir el conta text2 .
            'Le quito los `
            SQL = Replace(SQL, "`", "")
            SQL = Trim(Mid(SQL, 13))
            'LE añado el contax.
            SQL = "CREATE TABLE conta" & Text2(2).Text & "." & SQL
            
            Label6.Caption = "[" & Bucle & "]" & TablaAnt & " (" & Tam2 & " /" & Tamanyo & ")"
            Label6.Refresh
            
            If EjecutaSQL(SQL) Then ColCreate.Remove Tam2
            
        Next
        Me.Refresh
        espera 0.5
        Bucle = Bucle + 1
        If ColCreate.Count = 0 Then
            Label6.Caption = "Creacion finalizada. " & Bucle - 1
            Label6.Refresh
            Bucle = 11 'YA ESTA TODO CREADO
        End If
            
        
    Loop Until Bucle > 10   'Si en 10 iteraciones no ha acabado.... vamos mal
    ''
    'Aqui ya tiene que a ver finalizado
    If ColCreate.Count > 0 Then
        'Algo va mal
        MsgBox "ALGO HA IDO MAL. "
    Else
        CrearEstructura = True
    End If
    
End Function



Private Function InsercionDatos() As Boolean
Dim Rs As Recordset
Dim Linea As String
Dim Origen As String
Dim Insert As String
Dim F As Date



    On Error GoTo EInsercionDatos
    InsercionDatos = False
    
    Insert = "ariconta" & Text2(2).Text & "."
    Origen = "ariconta" & vEmpresa.codempre & "."
    
    
    'Datos basico
    Tam2 = 2
    TablaAnt = "contadores|scryst|"

    
    For i = 1 To Tam2
        SQL = RecuperaValor(TablaAnt, i)
        Label6.Caption = "Datos básicos: " & SQL & " (" & i & "/" & Tam2 & ")"
        Label6.Refresh
        Linea = SQL
        If EjecutaSQL(SQL) Then
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            If Not EjecutaSQL(SQL) Then
                SQL = "Error insertando en tabla " & Insert & Linea
            Else
                SQL = ""
            End If
        Else
            SQL = "Error borrando tabla" & Insert & Linea
        End If
        If SQL <> "" Then
            SQL = SQL & ": " & Insert & vbCrLf & "El proceso continuará"
            MsgBox SQL, vbExclamation
        End If
   Next
    
    
    
    
    'Cuentas
    i = 0
    If Check1(i).Value Then
        Linea = "cuentas"
        
        Label6.Caption = Check1(i).Caption
        Label6.Refresh
        Conn.Execute "DELETE FROM " & Insert & Linea
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        SQL = SQL & " WHERE apudirec='S'"
        Conn.Execute SQL
    End If
    
    Me.Refresh
    
    'Conceptos
    i = 1
    If Check1(i).Value Then
        Linea = "conceptos"
        Label6.Caption = Check1(i).Caption
        Label6.Refresh
        Conn.Execute "DELETE FROM " & Insert & Linea
        
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
    End If
    
    
    'Centros de coste
    i = 5
    If Check1(i).Value Then
        Linea = "ccoste"
        Label6.Caption = Check1(i).Caption
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
        Linea = "ccoste_lineas"
        Label6.Caption = Check1(i).Caption & " lineas"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
    End If
    
    
    
    
    
    'Asientos predefinidos
    i = 4
    'Se hara, aparte de si esta marcado, si estan las cuentas, conceptos,centros de coste
    Tam2 = Check1(0).Value + Check1(1).Value + Check1(2).Value + Check1(5).Value
    If Check1(i).Value Then
        If Tam2 = 4 Then
            Linea = "asipre"
            Label6.Caption = Check1(i).Caption
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
            
            Linea = "asipre_lineas"
            Label6.Caption = Check1(i).Caption & " lineas"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
        End If
    End If
    
    

        
    'Configuracion Balances
    i = 6
    If Check1(i).Value Then
        
            Linea = "balances"
            
            Label6.Caption = "Balances 1/3"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
            
            
            Linea = "balances_texto"
            Label6.Caption = "Balances 2/3"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL
            
            
            Linea = "balances_ctas"
            Label6.Caption = "Balances 3/3"
            Label6.Refresh
            SQL = "INSERT INTO " & Insert & Linea
            SQL = SQL & " SELECT * FROM " & Origen & Linea
            Conn.Execute SQL

    End If
        
    
    
    i = 7
    If Check1(i).Value Then
        
        Linea = "formapago"
        Label6.Caption = Check1(i).Caption
        Label6.Refresh
        Conn.Execute "DELETE FROM " & Insert & Linea
        
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
            

    End If
    
    
    
    '-----------------------------------------------------------
    'dATOS FIJOS COMO EMPRESA,EMPRESA2, PARAMETROS
        'Asientos predefinidos

    
    'Empresa
        Linea = "empresa"
        Label6.Caption = "Datos Empresa"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL

        
        
        Linea = "empresa2"
        Label6.Caption = "Datos Empresa"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
        
   'Plan contables y actualizar contadores
    Label6.Caption = "Subcuentas"
    Label6.Refresh
    Linea = "cuentas"
    SQL = "INSERT INTO " & Insert & Linea
    SQL = SQL & " SELECT * FROM " & Origen & Linea
    SQL = SQL & " WHERE apudirec<>'S'"
    Conn.Execute SQL
    
    
    
    'Contadores
    Label6.Caption = "Contadores"
    Label6.Refresh
    SQL = "UPDATE " & Insert & "contadores Set contado1=1, contado2=1 WHERE tiporegi='0'"
    Conn.Execute SQL
    SQL = "UPDATE " & Insert & "contadores Set contado1=0, contado2=0 WHERE tiporegi<>'0'"
    Conn.Execute SQL
        
        
        
        
        
    '----------
    'parametros
    '----------------------------------------
    'Los parametros solo podran ser insertado SI se piden ctas, conce y diarios
    
    SQL = ""
    If Check1(1).Value = 0 Or Check1(3).Value = 0 Then SQL = "1"
    If Check1(0).Value = 0 Then SQL = SQL & "1"
    If Len(SQL) = 0 Then
        
    
        Linea = "parametros"
        Label6.Caption = "Parámetros"
        Label6.Refresh
        SQL = "INSERT INTO " & Insert & Linea
        SQL = SQL & " SELECT * FROM " & Origen & Linea
        Conn.Execute SQL
        
        espera 0.5
    
    
        'En parametros
        F = CDate(txtFecha(3).Text)
        F = DateAdd("yyyy", 1, F)
        F = DateAdd("d", -1, F)

            SQL = "UPDATE " & Insert & "parametros SET fechaini='" & Format(txtFecha(3).Text, "yyyy-mm-dd")
            SQL = SQL & "', fechafin='" & Format(F, "yyyy-mm-dd") & "'"
            Conn.Execute SQL
      
        
        
     End If
        
    'Y actualizamos a los valores k nuevos
    SQL = "UPDATE " & Insert & "empresa SET nomempre= '" & Text2(0).Text & "', nomresum= '" & Text2(1).Text & "',codempre =" & Text2(2).Text
    Conn.Execute SQL
        
        
        
    InsercionDatos = True
        
    Exit Function
    
    
EInsercionDatos:
        MuestraError Err.Number, Label6.Caption
End Function


Private Sub SugerirValoresNuevaEmpresa()
    SQL = "Select max(codempre) from usuarios.empresasariconta where codempre<100"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Tam2 = DBLet(miRsAux.Fields(0), "N") + 1
    miRsAux.Close
    Set miRsAux = Nothing
    Text2(2).Text = Tam2
    
    Me.Check1(7).visible = vEmpresa.TieneTesoreria
    Check1(7).Value = Abs(vEmpresa.TieneTesoreria)
End Sub



'--------------------------------------------------------------------
'
'                    Subir un CERO el digitos
'
'--------------------------------------------------------------------

Private Function ComprobarOk(ByRef vNivelAnterior As Byte) As Boolean
Dim vE As String
Dim UltimoNivel As Byte
    On Error GoTo EComprobarOk
    ComprobarOk = False
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '----------------------------------------------------------------------
    '
    'Comprobamos k las tablas siguientes NO tiene registros
    '
    '
    'Comprobamos k el ultimo nivel no es 10
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "empresa", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    vE = ""
    If miRsAux.EOF Then
        vE = "No está definida la empresa."
    Else
        UltimoNivel = DBLet(miRsAux.Fields(3), "N")
        If UltimoNivel = 0 Then
            vE = "Si definir último nivel contable"
        Else
            NumTablas = DBLet(miRsAux.Fields(3 + UltimoNivel), "N")
            If NumTablas = 0 Then
                vE = "Último nivel es 0. Datos incorrectos"
            Else
                If NumTablas = 10 Then
                    vE = "No se puede ampliar el último nivel. Ya es 10"
                Else
                    'Fale vamos a devolver el nivel anterior al ultimo
                    vNivelAnterior = CByte(DBLet(miRsAux.Fields(3 + UltimoNivel - 1)))
                    If vNivelAnterior < 3 Or vNivelAnterior > 10 Then vE = "Error obteniendo nivel anterior"
                End If
            End If
        End If
    End If
    miRsAux.Close
    If vE <> "" Then
        MsgBox vE, vbExclamation
        Exit Function
    End If
    ComprobarOk = True
    Exit Function
EComprobarOk:
    MuestraError Err.Number, "ComprobarOk." & Err.Description
End Function





Private Function AgregarCuentasNuevas() As Boolean
Dim Izda As String
Dim Der As String

    Label3(2).Caption = "PGC"
    Label3(2).Refresh
    
    AgregarCuentasNuevas = False
   
    Set miRsAux = New ADODB.Recordset
    SQL = "Select count(*) from cuentas where apudirec='S'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim = 0 Then
        MsgBox "Ninguna cuenta de último nivel. La aplicación finalizará", vbCritical
        End
    End If
    NumRegElim = NumRegElim + 1
    
    SQL = "Select * from cuentas where apudirec='S'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    BACKUP_TablaIzquierda miRsAux, Izda
    Izda = "INSERT INTO cuentas " & Izda & " VALUES "
    Tamanyo = 0
    pb1.Value = 0
    While Not miRsAux.EOF
           Tamanyo = Tamanyo + 1
           PonerProgressBar (CLng(Tamanyo / NumRegElim * 1000))
           DatosTabla miRsAux, Der
           SQL = Izda & Der
           Conn.Execute SQL
           espera 0.001
           miRsAux.MoveNext
           If (Tamanyo \ 75) = 0 Then DoEvents
    Wend
    miRsAux.Close
    AgregarCuentasNuevas = True
    pb1.Value = 0
End Function


Private Sub DatosTabla(ByRef Rs As ADODB.Recordset, ByRef Derecha As String)
Dim i As Integer
Dim nexo As String
Dim Valor As String
Dim Tipo As Integer
    Derecha = ""
    nexo = ""
    For i = 0 To Rs.Fields.Count - 1
        Tipo = Rs.Fields(i).Type
        
        If IsNull(Rs.Fields(i)) Then
            Valor = "NULL"
        Else
        
            'pruebas
            Select Case Tipo
            'TEXTO
            Case 129, 200, 201
                Valor = Rs.Fields(i)
                NombreSQL Valor
                'Si el campo es el codmacta o apudirec lo cambiamos
                If i = 0 Then
                    Valor = CambioCta(Valor)
                Else
                    If i = 2 Then Valor = "P"                    'de PROVISIONAL
                End If
                Valor = "'" & Valor & "'"
            'Fecha
            Case 133
                Valor = CStr(Rs.Fields(i))
                Valor = "'" & Format(Valor, "yyyy-mm-dd") & "'"
                
            'Numero normal, sin decimales
            Case 2, 3, 16 To 19
                Valor = Rs.Fields(i)
            
            'Numero con decimales
            Case 131
                Valor = CStr(Rs.Fields(i))
                Valor = TransformaComasPuntos(Valor)
            Case Else
                Valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                Valor = Valor & vbCrLf & "SQL: " & Rs.Source
                Valor = Valor & vbCrLf & "Pos: " & i
                Valor = Valor & vbCrLf & "Campo: " & Rs.Fields(i).Name
                Valor = Valor & vbCrLf & "Valor: " & Rs.Fields(i)
                MsgBox Valor, vbExclamation
                MsgBox "El programa finalizara. Avise al soporte técnico.", vbCritical
                End
            End Select
                        
        End If
        Derecha = Derecha & nexo & Valor
        nexo = ","
    Next i
    Derecha = "(" & Derecha & ")"
End Sub

Private Function CambioCta(Cta As String) As String
Dim cad As String



    cad = Mid(Cta, 1, CInt(Text2(5).Text))
    cad = cad & "0" & Mid(Cta, CInt(Text2(5).Text) + 1)
    CambioCta = cad
End Function

Private Function HacerInsercionDigitoContable() As Boolean
    
    On Error GoTo EHacerInsercionDigitoContable
    HacerInsercionDigitoContable = False
    
    'Agregamos las cuentas nuevas con el numero correspondiente
    If AgregarCuentasNuevas Then
        Me.Refresh
        'Ahora hemos creado las cuentas con un digito mas
        'Ahora tendremos k ir tabla por tabla cambiando las cuentas a nivel nuevo
            
        
       'Facturas
       '------------------------------------------
       'Cabeceras
       CambiaTabla "factcli", "codmacta|cuereten|", 2
       CambiaTabla "factpro", "codmacta|cuereten|", 2
       
       
       'Linapus
       CambiaTabla "hlinapu", "codmacta|ctacontr|", 2
       CambiaTabla "asipre_lineas", "codmacta|ctacontr|", 2
    
        'Linea facturas
       CambiaTabla "factcli_lineas", "codmacta|", 1
       CambiaTabla "factpro_lineas", "codmacta|", 1
       
       
       CambiaTabla "departamentos", "codmacta|", 1
       CambiaTabla "parametros", "ctaperga|ctahpdeudor|ctahpacreedor|", 3
       CambiaTabla "presupuestos", "codmacta|", 1
       CambiaTabla "inmovele_rep", "codmacta2|", 1
       CambiaTabla "inmovele", "codprove|codmact1|codmact2|codmact3|", 4
       CambiaTabla "norma43", "codmacta|", 1
       CambiaTabla "bancos", "ctagastos|ctaingreso|ctaefectosdesc|ctagastostarj|", 4
       
       'Tipos de iva
       CambiaTabla "tiposiva", "cuentare|cuentarr|cuentaso|cuentasr|cuentasn|", 5
   
       
       
       
       If vEmpresa.TieneTesoreria Then

            '
            CambiaTabla "paramtesor", "RemesaCancelacion|RemesaConfirmacion|pagarecta|taloncta|pagarectaPRO|talonctaPRO|ctaefectcomerciales|", 7
            CambiaTabla "cobros", "codmacta|ctabanc1|", 2
            CambiaTabla "remesas", "codmacta|", 1
            CambiaTabla "talones", "codmacta|", 1
            CambiaTabla "gastosfijos", "ctaprevista|contrapar|", 2
            
            CambiaTabla "compensa", "codmacta|", 1
            'CambiaTabla "compensa_facturas", "codmacta|ctabanc1|ctabanc2|", 3    no lleva compensacion
            
            CambiaTabla "reclama", "codmacta|", 1
            
            
            CambiaTabla "pagos", "codmacta|ctabanc1|", 2
            CambiaTabla "transferencias", "codmacta|", 1
            
       End If
            
       'Quitamos las cuentas 'S'
       SQL = "Delete from Cuentas where apudirec='S'"
       Conn.Execute SQL
       
       'Las k eran apuntes directos P pasan a ser S
       SQL = "UPDATE Cuentas SET apudirec='S' where apudirec='P'"
       Conn.Execute SQL
       
       'Actualizamos en empresas
       AumentarEmpresaDigitoUltimoNivel
       
       
       
       'Creamos las cuentas de subnivel
       CrearSubNivel
       pb1.Value = 0
       Label3(2).Caption = ""
       Label3(2).Refresh
       vEmpresa.Leer vEmpresa.codempre
       vParam.Leer
       If vEmpresa.TieneTesoreria Then vParamT.Leer
       vParam.FijarAplicarFiltrosEnCuentas vEmpresa.nomempre
       
       HacerInsercionDigitoContable = True
       
       
    End If
    Exit Function
EHacerInsercionDigitoContable:
    MuestraError Err.Number, "Errorfatal." & vbCrLf & Err.Description
End Function



Private Function CambiaTabla(tabla As String, vCampos As String, NCampos As Integer)
Dim i As Integer

    ReDim Campos(NCampos)
    
    For i = 1 To NCampos
        Campos(i) = RecuperaValor(vCampos, i)
    Next i
    
    Label3(2).Caption = tabla
    pb1.Value = 0
    Me.Refresh
    CambiaValores tabla, NCampos

End Function



Private Sub AumentarEmpresaDigitoUltimoNivel()

    
    
    i = vEmpresa.numnivel
    SQL = "UPDATE empresa SET numdigi" & CStr(i) & " = "
    i = CInt(Text2(5).Text) + 1
    SQL = SQL & CStr(i)
    i = vEmpresa.numnivel + 1
    SQL = SQL & ", numdigi" & CStr(i) & " = " & vEmpresa.DigitosUltimoNivel + 1
    SQL = SQL & ", numnivel = numnivel +1"
    
   
    
    Conn.Execute SQL
End Sub


Private Function CambiaValores(tabla As String, numCta As Integer)
Dim SQL As String
Dim cad As String
Dim i As Integer
    cad = ""
    SQL = ""
    On Error GoTo ECambia
    
    For i = 1 To numCta
        'Para bonito
        Label3(2).Caption = tabla & " (" & i & " de " & numCta & ")"
        pb1.Value = 0
        Me.Refresh
        Tamanyo = 0
        'Contador  COUNT(distinct(codmacta))
        SQL = "SELECT COUNT(DISTINCT(" & Campos(i) & ")) from " & tabla
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not miRsAux.EOF Then Tamanyo = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        

        If Tamanyo > 0 Then
            'Updateamos la primera cta
            Tamanyo = Tamanyo + 1
            SQL = "SELECT " & Campos(i) & " FROM " & tabla & " GROUP BY " & Campos(i)
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            While Not miRsAux.EOF
                NumRegElim = NumRegElim + 1
                PonerProgressBar Val((NumRegElim / Tamanyo) * 1000)
                If Not IsNull(miRsAux.Fields(0)) Then
                    cad = CambioCta(miRsAux.Fields(0))
                    SQL = "UPDATE " & tabla & " SET " & Campos(i) & " = '" & cad & "'"
                    SQL = SQL & " WHERE " & Campos(i) & " = '" & miRsAux.Fields(0) & "'"
                    Conn.Execute SQL
                End If
                'Sig
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
    Next i
    Exit Function
ECambia:
    MuestraError Err.Number, Err.Description
End Function

Private Sub PonerProgressBar(Valor As Long)
    If Valor <= 1000 Then pb1.Value = Valor
End Sub


Private Sub CrearSubNivel()
Dim Col As Collection

    Label3(2).Caption = "Subniveles a crear (leyendo)"
    Label3(2).Refresh
    pb1.Value = 0
    i = CInt(Text2(5).Text) + 1
    SQL = "select substring(codmacta,1," & i & "),nommacta from cuentas where apudirec='S' group by 1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Col = New Collection
    While Not miRsAux.EOF
        Col.Add CStr(miRsAux.Fields(0)) & "|" & DBLet(miRsAux.Fields(1), "T") & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ya tengo los subniveles que tengo que crear
    
    Label3(2).Caption = "Subniveles a crear (insertando)"
    Label3(2).Refresh
    espera 0.3
    DoEvents
    
    i = CInt(Text2(5).Text)
    SQL = String(i, "_")
    SQL = "Select codmacta,nommacta from cuentas where codmacta like '" & SQL & "'"
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    ParaElLog = "INSERT INTO cuentas(apudirec,model347,codmacta,nommacta,razosoci) VALUES ('N',0,'"
    Tam2 = CInt(Text2(5).Text)
    For i = 1 To Col.Count
        PonerProgressBar CLng((i / Col.Count) * 1000)
        TablaAnt = RecuperaValor(Col.Item(i), 1)
        SQL = ""
        miRsAux.Find "Codmacta = '" & Mid(TablaAnt, 1, Tam2) & "'", , adSearchForward, 1
        If Not miRsAux.EOF Then SQL = DBLet(miRsAux!Nommacta, "T")
        If SQL = "" Then SQL = RecuperaValor(Col.Item(i), 1)
        If SQL = "" Then SQL = "Aumentando ceros"
        SQL = DevNombreSQL(SQL)
        TablaAnt = ParaElLog & TablaAnt & "','" & SQL & "','" & SQL & "')"
        Conn.Execute TablaAnt
    Next i
End Sub

Private Sub txtIVA_GotFocus(Index As Integer)
    PonFoco txtIVA(Index)
End Sub

Private Sub txtIVA_LostFocus(Index As Integer)
    SQL = ""
    txtIVA(Index).Text = Trim(txtIVA(Index).Text)
    If txtIVA(Index).Text <> "" Then
        If EsNumerico(txtIVA(Index).Text) Then
            SQL = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", txtIVA(Index).Text)
            If SQL = "" Then MsgBox "No existe el tipo de iva: " & txtIVA(Index).Text, vbExclamation
               
        End If
        If SQL = "" Then
            txtIVA(Index).Text = ""
            PonleFoco txtIVA(Index)
        End If
    End If
    txtDescIVA(Index).Text = SQL
End Sub



Private Function HacerCambioIVA() As Boolean

    HacerCambioIVA = False
    If CambioIVA(True) Then
        If CambioIVA(False) Then
            HacerCambioIVA = True
            
            'EL LOG
            ParaElLog = "IVA origen:     " & txtIVA(0).Text & "  -  " & txtDescIVA(0).Text & vbCrLf
            ParaElLog = ParaElLog & "IVA destino:    " & txtIVA(1).Text & "  -   " & txtDescIVA(1).Text & vbCrLf
            ParaElLog = ParaElLog & "Fechas: " & txtFecha(4).Text & " - " & txtFecha(5).Text
            ParaElLog = "CAMBIO IVA" & vbCrLf & ParaElLog
            vLog.Insertar 16, vUsu, ParaElLog
            ParaElLog = ""
            
        End If
            
    End If
    
End Function



Private Function CambioIVA(Clientes As Boolean) As Boolean
    
    'NO HACE FALTA transaccionar.
'    Conn.CommitTrans
    CambioIVA = False
'    For i = 1 To 3
        If Clientes Then
            lblIVA.Caption = "Clientes"
        Else
            lblIVA.Caption = "Proveedores"
        End If
        lblIVA.Caption = lblIVA.Caption & ".  Iva " & i
        lblIVA.Refresh
        
        If Clientes Then
            SQL = "UPDATE factcli_lineas SET codigiva = " & txtIVA(1).Text
            SQL = SQL & " WHERE codigiva = " & txtIVA(0).Text
            TablaAnt = "fecfactu"
        Else
            SQL = "UPDATE factpro_lineas SET codigiva = " & txtIVA(1).Text
            SQL = SQL & " WHERE codigiva = " & txtIVA(0).Text
            TablaAnt = "fecharec"
        End If
        If txtFecha(4).Text <> "" Then SQL = SQL & " AND " & TablaAnt & ">= '" & Format(txtFecha(4).Text, FormatoFecha) & "'"
        If txtFecha(5).Text <> "" Then SQL = SQL & " AND " & TablaAnt & "<= '" & Format(txtFecha(5).Text, FormatoFecha) & "'"
    
    
        If Not EjecutaSQL(SQL) Then
            'Se ha producido un error
            TablaAnt = "Error grave." & vbCrLf & "Cambiando IVA " & i & vbCrLf & vbCrLf
            TablaAnt = TablaAnt & "Desc: " & SQL & vbCrLf & "Avise a soporte técnico con el error"
            MsgBox TablaAnt, vbCritical
            Exit Function
        End If
 '   Next i
    
    CambioIVA = True
        
End Function


'-------------------------------------------------------------------
'
'RENUMERAR FRA PROVEEDORES
'
'-------------------------------------------------------------------
Private Sub txtRenumFrapro_GotFocus(Index As Integer)
    PonFoco txtRenumFrapro(Index)
End Sub

Private Sub txtRenumFrapro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtRenumFrapro_LostFocus(Index As Integer)
    txtRenumFrapro(Index).Text = Trim(txtRenumFrapro(Index).Text)
    If txtRenumFrapro(Index).Text = "" Then Exit Sub
    
    Select Case Index
        Case 2 ' contador
            PonerFormatoEntero txtRenumFrapro(Index)
        Case 0 ' fecha
            If PonerFormatoFecha(txtRenumFrapro(Index)) Then
                PonerValoresCalculos
            End If
            
        Case 1 ' serie
            Text3(0).Text = PonerNombreDeCod(txtRenumFrapro(Index), "contadores", "nomregis", "tiporegi", "T")
    End Select
    
End Sub




Private Function HacerRenumeracionFacturas() As Boolean
Dim Fecha As Date
Dim F2 As Date
Dim Finicio As Date
Dim AnoPartido As Boolean
Dim Ok As Boolean

    On Error GoTo EHacerRenumeracionFacturas
    HacerRenumeracionFacturas = False
    
    
    SQL = "Select fechaini,codinume from parametros"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Finicio = miRsAux!fechaini
    'Si no graba numdocum, SEGURO que no lo updateamos
'--
'    If miRsAux!CodiNume = 2 Then Me.chkUpdateNumDocum = 0
    miRsAux.Close

    'Fecha INCIO en actual o siguiente
    If Me.optFrapro(1).Value Then Finicio = DateAdd("yyyy", 1, Finicio)


    LabelIndF(0).Caption = "Realizando comprobaciones"
    LabelIndF(1).Caption = ""
    Me.Refresh
    DoEvents
    
    
    Fecha = Finicio
    F2 = DateAdd("yyyy", 1, Fecha)
    F2 = DateAdd("d", -1, F2)
    AnoPartido = Year(Fecha) <> Year(F2)
    
    'ContadorInserciones --> Numregelim
    NumRegElim = Val(txtRenumFrapro(2).Text)
    
    SQL = "Select count(*) from factpro where fecharec>='" & Format(txtRenumFrapro(0).Text, FormatoFecha) & "'"
    SQL = SQL & " AND fecharec<='" & Format(F2, FormatoFecha) & "'"
    SQL = SQL & " and numserie = " & DBSet(txtRenumFrapro(1).Text, "T")
'--
'    If Me.chkSALTO_numerofactura.Value = 1 Then Sql = Sql & " AND numregis >= " & txtRenumFrapro(0).Text
'
'    'Desde hasta
'    AnyadeDesdeHastaNumregis
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Tam2 = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    'Ya se cuantas facturas hay.
    '1.- Si hay 0 cierro y me largo
    '2.- Si hay mas de una veo si entre las fechas del ejerccio hay alguna factura con numregis entre los valores
    '   que voy a renumerar
    If Tam2 = 0 Then
        MsgBox "Ninguna factura a renumerar", vbExclamation
        Set miRsAux = Nothing
        Exit Function
    End If




    Tamanyo = 0
'-- quito el if
'    If Me.chkSALTO_numerofactura.Value = 0 Then
        'Proceso normal. No voy a partir de un numero de factura
            If AnoPartido Then
                '        AÑO PARTIDO
    
                SQL = "Select count(*) from factpro where anofactu = " & Year(Fecha)
                SQL = SQL & " and numserie = " & DBSet(txtRenumFrapro(1).Text, "T")
                SQL = SQL & " and fecharec>='" & Format(txtRenumFrapro(0).Text, FormatoFecha) & "'"
                SQL = SQL & " AND numregis >= " & NumRegElim & " and numregis<= " & NumRegElim + Tam2
'--                AnyadeDesdeHastaNumregis
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not miRsAux.EOF Then Tamanyo = DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
            
                SQL = "Select count(*) from factpro where anofactu = " & Year(F2)
                SQL = SQL & " and numserie = " & DBSet(txtRenumFrapro(1).Text, "T")
                SQL = SQL & " and fecharec>='" & Format(txtRenumFrapro(0).Text, FormatoFecha) & "'"
                SQL = SQL & " and fecharec <= " & DBSet(F2, "F")
                SQL = SQL & " AND numregis >= " & NumRegElim & " and numregis<= " & NumRegElim + Tam2
'--                AnyadeDesdeHastaNumregis
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not miRsAux.EOF Then Tamanyo = Tamanyo + DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
            
            
            Else
                'AÑO NORMAL
                SQL = "Select count(*) from factpro where anofactu = " & Year(Fecha)
                SQL = SQL & " and numserie = " & DBSet(txtRenumFrapro(1).Text, "T")
                SQL = SQL & " and fecharec>='" & Format(txtRenumFrapro(0).Text, FormatoFecha) & "'"
                SQL = SQL & " and fecharec <= " & DBSet(F2, "F")
                SQL = SQL & " AND numregis >= " & NumRegElim & " and numregis<= " & NumRegElim + Tam2
'                AnyadeDesdeHastaNumregis
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Tamanyo = 0
                If Not miRsAux.EOF Then Tamanyo = DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
                
            End If
                
    If Tamanyo > 0 Then
        MsgBox "Se solaparán números de factura", vbExclamation
        Exit Function
    End If
    


    'AQUI SE HACE LA RENUMERACION PROIPAMENTE DICHA
    'Proceso laaargo donde los haya
    'Puesto que hay que hacer
    '   Crear la factura 0
    '   UPDATEAR LAS lineas de FACTURA A LA 0
    '      "        la factura a su nuevo numero
    '       "    las lineas al nuevo numero
    '    Si procede, updatear NUMDOCUM
'--    Fecha = Finicio
    Fecha = txtRenumFrapro(0).Text
    Ok = RenumeraFacturas(Fecha)


    '-----------------------------
    'Insertamos el LOG
    SQL = "siguiente"
    If optFrapro(0).Value Then SQL = "actual"
    SQL = "Ejercicio " & SQL & vbCrLf
    
    ParaElLog = SQL & "Renumerar facturas proveedor: " & vbCrLf
    ParaElLog = ParaElLog & "Nº registro " & txtRenumFrapro(2).Text & vbCrLf
    SQL = ""
    
    ParaElLog = ParaElLog & "Registros: " & CStr(NumRegElim) & vbCrLf
    ParaElLog = ParaElLog & vbCrLf & "Desde Fecha: " & txtRenumFrapro(0).Text & vbCrLf
    ParaElLog = "Renumerar nºregistro " & vbCrLf & ParaElLog
    
    
    vLog.Insertar 21, vUsu, ParaElLog
    ParaElLog = ""


    If Ok Then HacerRenumeracionFacturas = True
    
    Exit Function
EHacerRenumeracionFacturas:
    MuestraError Err.Number, Err.Description
End Function

Private Sub AnyadeDesdeHastaNumregis()
    If txtRenumFrapro(1).Text <> "" Then SQL = SQL & " AND numregis >= " & txtRenumFrapro(1).Text
    If txtRenumFrapro(2).Text <> "" Then SQL = SQL & " AND numregis <= " & txtRenumFrapro(2).Text
End Sub

'Comprobaremos que todas las facturas que estan contbilizadas tiene asiento
Private Function ComprobarFRAPROContabilizadas(Fecha As Date, DesdePeriodo As Boolean) As Boolean
Dim F As Date
Dim NF As Integer
Dim Bucles As Byte

    On Error GoTo EComprobarFRAPROContabilizadas
    ComprobarFRAPROContabilizadas = False
    
    
    'Por velocidad dividremos el ejhercicio end tres cuatrimestres
    LabelIndF(0).Caption = "Comprobar contabilización facturas"
    F = Fecha
    Bucles = 1
    F = txtRenumFrapro(0)
    If optFrapro(0) Then
        Fecha = vParam.fechafin
    Else
        Fecha = DateAdd("yyyy", 1, vParam.fechafin)
    End If

    Set miRsAux = New ADODB.Recordset
    
    Insert = ""
    For NF = 1 To Bucles
        
        SQL = "select numregis,fecharec,factpro.numasien,hcabapu.numasien as na,factpro.numdiari,anofactu "
        SQL = SQL & " from factpro left join hcabapu"
        SQL = SQL & " on factpro.numasien=hcabapu.numasien and factpro.fechaent=hcabapu.fechaent and factpro.numdiari=hcabapu.numdiari"
        SQL = SQL & " where fecharec>='" & Format(F, FormatoFecha)
        LabelIndF(1).Caption = "Desde : " & F & "   "
        
        F = DateAdd("d", -1, Fecha)
        SQL = SQL & "' and fecharec<='" & Format(F, FormatoFecha) & "'"
        SQL = SQL & " and numserie = " & DBSet(txtRenumFrapro(1).Text, "T")
        
        LabelIndF(1).Caption = LabelIndF(1).Caption & "  hasta:  " & F
        
        
        F = Fecha
        
        Fecha = DateAdd("m", 4, Fecha)
        
        DoEvents
        
        
        'AHora tenog el res
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            If IsNull(miRsAux!NumAsien) Then
                'Factura sin contabilizar
                
            Else
                If IsNull(miRsAux!NA) Then
                    'ERRROR GRAVE
                    'La factura tiene numero asiento, pero el asiento NO existe
                    If miRsAux!NumAsien = 0 Then
                        'Es posible ya que hay frapro que no se contabilizan
                    
                    Else
                        Insert = Insert & miRsAux!Numregis & " / " & miRsAux!anofactu & ": " & Format(miRsAux!NumAsien, "00000") & ";"
                    End If
                End If
            End If
        
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        
    Next NF
    
    
    If Insert <> "" Then
        'HAY ERRORES
        NF = FreeFile
        SQL = App.Path & "\" & Format(Now, "yymmdd") & "_" & Format(Now, "hhmmss") & ".txt"
        Open SQL For Output As #NF
        Print #NF, Insert
        Close #NF
        Insert = "Se han producido errores. Vea el archivo: " & vbCrLf & vbCrLf & SQL
        Insert = Insert & " Desea continuar?"
        If MsgBox(Insert, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    If DesdePeriodo Then Set miRsAux = Nothing
    
    ComprobarFRAPROContabilizadas = True
    Exit Function
EComprobarFRAPROContabilizadas:
    MuestraError Err.Number, "ComprobarFRAPROContabilizadas"
End Function



Private Function RenumeraFacturas(Fec As Date) As Boolean
Dim Ok As Boolean

    On Error GoTo ERenumeraFacturas
    RenumeraFacturas = False
    
    'Creo la factura 0
    LabelIndF(0).Caption = "Generando factura 0 / 00001"
    LabelIndF(1).Caption = ""
    Me.Refresh
    SQL = "INSERT INTO factpro (numserie,numregis,anofactu,fecharec,numfactu,fecfactu,codconce340,codopera,"
    SQL = SQL & "observa,codmacta,codforpa,totbases,tiporeten,fecliqpr,nommacta,escorrecta) VALUES "
    SQL = SQL & "('1', 0, 1, '0000-00-00', '1','0000-00-00', '0', '0', "
    SQL = SQL & "'RENUM','1', 0, 0,0,'0000-00-00','',0)"
    Conn.Execute SQL
    
    
    
    
    'En esta function RENUMERA
    LabelIndF(0).Caption = "Renumerando"
    DoEvents
    
    Ok = RenumeracionReal(Fec)
    
    
    
    
    'Borro la factura
    SQL = "DELETE FROM factpro WHERE numregis=0 AND anofactu=1"
    Conn.Execute SQL
    
    If Ok Then
        MsgBox "Proceso finalizado con exito", vbInformation
        RenumeraFacturas = True
    End If
    
    Exit Function
ERenumeraFacturas:
    MuestraError Err.Number, Err.Description
End Function


Private Function RenumeracionReal(Fec As Date) As Boolean


    On Error GoTo ERenumeracionReal
    RenumeracionReal = False
    SQL = "Select numserie, numregis,anofactu,numasien,fechaent,numdiari from factpro where fecharec>='" & Format(Fec, FormatoFecha)
    Fec = DateAdd("yyyy", 1, Fec)
    Fec = DateAdd("d", -1, Fec)
    SQL = SQL & "' AND fecharec <='" & Format(Fec, FormatoFecha) & "' "
    SQL = SQL & " and numserie = " & DBSet(txtRenumFrapro(1).Text, "T")
    SQL = SQL & " ORDER BY fecharec,numregis"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    Tam2 = Val(Me.txtRenumFrapro(2).Text)
    While Not miRsAux.EOF
            
            LabelIndF(1).Caption = miRsAux!NUmSerie & " " & miRsAux!Numregis & " / " & miRsAux!anofactu & " --> " & Tam2
            LabelIndF(1).Refresh
            NumRegElim = NumRegElim + 1
            If NumRegElim > 60 Then
                NumRegElim = 0
                Me.Refresh
                DoEvents
            End If
            
            'Updateo las lineas a la 0/1
            SQL = "UPDATE factpro_lineas set numregis = 0 , anofactu=1 where numserie = " & DBSet(miRsAux!NUmSerie, "T") & " and numregis =" & miRsAux!Numregis & " AND anofactu =" & miRsAux!anofactu
            Conn.Execute SQL
            SQL = "UPDATE factpro_totales set numregis = 0 , anofactu=1 where numserie = " & DBSet(miRsAux!NUmSerie, "T") & " and numregis =" & miRsAux!Numregis & " AND anofactu =" & miRsAux!anofactu
            Conn.Execute SQL
            
            'Updateo la factura
            SQL = "UPDATE factpro set numregis = " & Tam2 & " where numserie = " & DBSet(miRsAux!NUmSerie, "T") & " and numregis =" & miRsAux!Numregis & " AND anofactu =" & miRsAux!anofactu
            Conn.Execute SQL
            
            'Reestablezco las lineas
            SQL = "UPDATE factpro_lineas set numregis = " & Tam2 & ", anofactu =" & miRsAux!anofactu & " where numserie = " & DBSet(miRsAux!NUmSerie, "T") & " and numregis = 0 AND anofactu = 1"
            Conn.Execute SQL
            SQL = "UPDATE factpro_totales set numregis = " & Tam2 & ", anofactu =" & miRsAux!anofactu & " where numserie = " & DBSet(miRsAux!NUmSerie, "T") & " and numregis = 0 AND anofactu = 1"
            Conn.Execute SQL
            
'            If Me.chkUpdateNumDocum.Value = 1 Then
            If vParam.CodiNume <> 2 Then
                If Not IsNull(miRsAux!NumAsien) And Not IsNull(miRsAux!NumDiari) Then
                    SQL = "UPDATE hlinapu set numdocum = '" & Format(Tam2, "0000000000") & "' WHERE numasien =" & miRsAux!NumAsien
                    SQL = SQL & " AND numdiari =" & miRsAux!NumDiari & " AND fechaent = '" & Format(miRsAux!FechaEnt, FormatoFecha) & "'"
                    Conn.Execute SQL
                End If
            End If
        
            miRsAux.MoveNext
            Tam2 = Tam2 + 1
    Wend
    miRsAux.Close
    RenumeracionReal = True
    Exit Function
    
ERenumeracionReal:
    Insert = "Error grave: " & Err.Number & vbCrLf & vbCrLf & SQL & vbCrLf & "Desc: " & Err.Description
    MsgBox Insert, vbCritical
    Insert = ""
End Function




