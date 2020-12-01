VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAVNVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Contable de Intereses"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   9255
   Icon            =   "frmAVNVarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   7395
      Left            =   105
      TabIndex        =   12
      Top             =   120
      Width           =   8985
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4350
         Left            =   90
         TabIndex        =   15
         Top             =   1395
         Width           =   8640
         Begin VB.TextBox txtNombre 
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
            Index           =   9
            Left            =   3375
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   3780
            Width           =   5025
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   3780
            Width           =   900
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   3375
            Width           =   1305
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   8
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   3375
            Width           =   4620
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   2955
            Width           =   1305
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   7
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2970
            Width           =   4620
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   2520
            Width           =   1305
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   6
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   2520
            Width           =   4620
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   5
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   2100
            Width           =   1305
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   2100
            Width           =   4620
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1680
            Width           =   900
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   3375
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1680
            Width           =   5025
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2430
            MaxLength       =   30
            TabIndex        =   3
            Top             =   1260
            Width           =   5940
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   810
            Width           =   1305
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   405
            Width           =   4665
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   4
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   405
            Width           =   1305
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   5
            Left            =   180
            TabIndex        =   34
            Top             =   3825
            Width           =   1980
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   2160
            ToolTipText     =   "Buscar diario"
            Top             =   3810
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2160
            ToolTipText     =   "Buscar Concepto"
            Top             =   3420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto Haber"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   32
            Top             =   3420
            Width           =   2025
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   2160
            ToolTipText     =   "Buscar Concepto"
            Top             =   3015
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto Debe"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   30
            Top             =   2994
            Width           =   2025
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta"
            Top             =   2565
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Gasto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   28
            Top             =   2570
            Width           =   2025
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta"
            Top             =   2160
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Retención"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   2146
            Width           =   2025
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2160
            ToolTipText     =   "Buscar forma pago"
            Top             =   1710
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de Pago"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   24
            Top             =   1722
            Width           =   1980
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   22
            Top             =   1298
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   874
            Width           =   1965
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   2160
            Picture         =   "frmAVNVarios.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   900
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   285
            Index           =   24
            Left            =   180
            TabIndex        =   17
            Top             =   450
            Width           =   1890
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta"
            Top             =   450
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   8655
         Begin VB.TextBox txtCodigo 
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
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   480
            Width           =   1305
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1140
            Picture         =   "frmAVNVarios.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label4 
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
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   270
            TabIndex        =   14
            Top             =   450
            Width           =   750
         End
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
         Left            =   7695
         TabIndex        =   11
         Top             =   6795
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptar 
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
         Left            =   6510
         TabIndex        =   10
         Top             =   6795
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   300
         Left            =   180
         TabIndex        =   18
         Top             =   5850
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
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
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   6210
         Width           =   8505
      End
      Begin VB.Label lblProgres 
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
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   6525
         Width           =   8445
      End
   End
   Begin VB.Frame FrameCancelacion 
      Height          =   5520
      Left            =   90
      TabIndex        =   35
      Top             =   225
      Width           =   7680
      Begin VB.CommandButton CmdAcepCancelacion 
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
         Left            =   5160
         TabIndex        =   53
         Top             =   4725
         Width           =   1065
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
         Left            =   6345
         TabIndex        =   55
         Top             =   4725
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   19
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   46
         Top             =   1935
         Width           =   1305
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   18
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   45
         Top             =   1545
         Width           =   1305
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   17
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   48
         Top             =   3420
         Width           =   5850
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   16
         Left            =   1575
         MaxLength       =   6
         TabIndex        =   44
         Top             =   660
         Width           =   865
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   15
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   47
         Top             =   2790
         Width           =   1305
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   16
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   675
         Width           =   4935
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   14
         Left            =   2835
         MaxLength       =   6
         TabIndex        =   49
         Top             =   3870
         Width           =   1050
      End
      Begin VB.Frame FrameResultado 
         Caption         =   "Cálculo a Ingresar"
         Enabled         =   0   'False
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
         Height          =   1680
         Left            =   3915
         TabIndex        =   36
         Top             =   1305
         Width           =   3480
         Begin VB.TextBox txtCodigo 
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
            Index           =   13
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   39
            Top             =   405
            Width           =   2130
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   12
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   38
            Top             =   810
            Width           =   2130
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   11
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   37
            Top             =   1215
            Width           =   2130
         End
         Begin VB.Label Label1 
            Caption         =   "Bruto"
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
            Left            =   135
            TabIndex        =   42
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Retención"
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
            Left            =   135
            TabIndex        =   41
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Neto"
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
            Left            =   135
            TabIndex        =   40
            Top             =   1215
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   51
         Top             =   3870
         Width           =   915
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1215
         Picture         =   "frmAVNVarios.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1215
         Picture         =   "frmAVNVarios.frx":01AD
         ToolTipText     =   "Buscar fecha"
         Top             =   1545
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   14
         Left            =   465
         TabIndex        =   61
         Top             =   1950
         Width           =   600
      End
      Begin VB.Label Label4 
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
         Index           =   15
         Left            =   465
         TabIndex        =   60
         Top             =   1545
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Período Liquidación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   8
         Left            =   270
         TabIndex        =   59
         Top             =   1215
         Width           =   2040
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   7
         Left            =   270
         TabIndex        =   58
         Top             =   3420
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Index           =   6
         Left            =   510
         TabIndex        =   57
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Avnics"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   4
         Left            =   270
         TabIndex        =   56
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Cancelación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   54
         Top             =   2475
         Width           =   2265
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1215
         Picture         =   "frmAVNVarios.frx":0238
         ToolTipText     =   "Buscar fecha"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1275
         MouseIcon       =   "frmAVNVarios.frx":02C3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   675
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Penalización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   2
         Left            =   270
         TabIndex        =   52
         Top             =   3915
         Width           =   2340
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje Retención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   3
         Left            =   4275
         TabIndex        =   50
         Top             =   3915
         Width           =   2115
      End
   End
   Begin VB.Frame FrameAyudaMod123 
      Height          =   3795
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   7815
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
         Index           =   2
         Left            =   6465
         TabIndex        =   74
         Top             =   3180
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   24
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   71
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1140
         Width           =   1305
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   23
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   70
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   780
         Width           =   1305
      End
      Begin VB.CommandButton CmdAcepAyuda123 
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
         Left            =   5280
         TabIndex        =   72
         Top             =   3180
         Width           =   1065
      End
      Begin VB.Frame FrameResultados 
         Caption         =   "Resultados"
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
         Height          =   1185
         Left            =   480
         TabIndex        =   63
         Top             =   1680
         Width           =   7050
         Begin VB.TextBox txtCodigo 
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
            Index           =   22
            Left            =   2355
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   66
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   600
            Width           =   1905
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   21
            Left            =   420
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   65
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   600
            Width           =   1605
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   20
            Left            =   4665
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   64
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   600
            Width           =   1905
         End
         Begin VB.Label Label4 
            Caption         =   "Base de Retenciones"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   11
            Left            =   2385
            TabIndex        =   69
            Top             =   360
            Width           =   2085
         End
         Begin VB.Label Label4 
            Caption         =   "Nro.Perceptores"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   10
            Left            =   420
            TabIndex        =   68
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "Retenciones"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   9
            Left            =   4680
            TabIndex        =   67
            Top             =   360
            Width           =   1875
         End
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   17
         Left            =   510
         TabIndex        =   76
         Top             =   375
         Width           =   870
      End
      Begin VB.Label Label4 
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
         Index           =   13
         Left            =   510
         TabIndex        =   75
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Index           =   12
         Left            =   510
         TabIndex        =   73
         Top             =   1140
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1275
         Picture         =   "frmAVNVarios.frx":0415
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1275
         Picture         =   "frmAVNVarios.frx":04A0
         ToolTipText     =   "Buscar fecha"
         Top             =   810
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmAVNVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Integer

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFpa As frmFormaPago 'formas de pago de la contabilidad
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmDi As frmTiposDiario
Attribute frmDi.VB_VarHelpID = -1
Private WithEvents frmavn As frmAVNAvnics 'Avnics
Attribute frmavn.VB_VarHelpID = -1



'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim IndCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Dim FFin As Date
Dim FIni As Date


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub CmdAcepAyuda123_Click()
Dim I As Byte
Dim Sql As String
Dim cadwhere As String
Dim NRegs As Long
Dim C As Object
Dim Rs As ADODB.Recordset
Dim nReceptores As Long
Dim Mens As String

        
    cadwhere = " (1 = 1) "
    If txtCodigo(23).Text <> "" Then cadwhere = cadwhere & " and movim.fechamov >= " & DBSet(txtCodigo(23).Text, "F")
    If txtCodigo(24).Text <> "" Then cadwhere = cadwhere & " and movim.fechamov <= " & DBSet(txtCodigo(24).Text, "F")

    Sql = "SELECT  count(*) from movim where " & cadwhere

    NRegs = TotalRegistros(Sql)

    If NRegs <> 0 Then
        Screen.MousePointer = vbHourglass
        
        Mens = "Cálculo del Total de perceptores: "
        nReceptores = TotalReceptores(cadwhere, Mens)
        
        If nReceptores <> 0 Then
            Set Rs = New ADODB.Recordset
            Sql = "select sum(timport1), sum(timport2) from movim where " & cadwhere
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
            If Not Rs.EOF Then
                txtCodigo(21).Text = Format(nReceptores, "###,###,##0")
                txtCodigo(22).Text = Format(Rs.Fields(0).Value, "###,###,##0.00")
                txtCodigo(20).Text = Format(Rs.Fields(1).Value, "###,###,##0.00")
                
                FrameResultados.visible = True
                ImgFec(5).Enabled = False
                ImgFec(6).Enabled = False
                CmdAcepAyuda123.visible = False
                CmdAcepAyuda123.Enabled = False
                cmdCancel(2).Caption = "Salir"
                
                Screen.MousePointer = vbDefault
            End If
        End If
    Else
        MsgBoxA "No hay registros para generar el fichero", vbExclamation
    End If


End Sub

Private Sub CmdAcepCancelacion_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte
Dim NReg As Long
Dim cadwhere As String
Dim cad As String

    If Not DatosOK Then Exit Sub
    
    ' el avnic no tiene que haber sido cancelado en el ejercicio
    cadwhere = "anoejerc = year(" & DBSet(txtCodigo(15).Text, "F") & ") and codialta <> 2 "
    cadwhere = cadwhere & " and codavnic = " & DBSet(txtCodigo(16).Text, "N")
    
    cad = "select count(*) from avnic where " & cadwhere
    NReg = TotalRegistros(cad)
    
    If NReg <> 0 Then
        If CalculoPenalizacion(cadwhere) Then
'            MsgBox "Proceso realizado correctamente", vbExclamation
            VisualizarFrameResultado (True)
'            cmdCancel_Click
        End If
    Else
        MsgBoxA "No existen datos entre esos límites. Reintroduzca.", vbExclamation
    End If

End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim I As Byte
Dim cadwhere As String

    If Not DatosOK Then Exit Sub
             
    Sql = "SELECT * " & _
          " FROM movim " & _
          "WHERE fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " intconta = 0"
             
    If TotalRegistrosConsulta(Sql) = 0 Then
        MsgBox "No existen datos a contabilizar a esa fecha.", vbExclamation
        Exit Sub
    End If
    
    cadwhere = " fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " intconta = 0"
               
    
    ContabilizarIntereses (cadwhere)
     'Eliminar la tabla TMP
    BorrarTMPErrComprob

    DesBloqueoManual ("CONINT") 'CONtabilizacion de CALculo
    
    
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de cierre de turno. Llame a soporte."
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click (0)
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 0 ' integracion contable de intereses
                PonFoco txtCodigo(0)
            Case 1 ' cancelacion
                PonFoco txtCodigo(16)
            Case 2 ' ayuda 123
                PonFoco txtCodigo(23)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next H
    
    Me.FrameCobros.visible = False
    Me.FrameCancelacion.visible = False
    Me.FrameAyudaMod123.visible = False
    
    Select Case OpcionListado
        Case 0 ' integracion contable de intereses
            Me.Caption = "Integración Contable de Intereses"
        
            txtCodigo(0).Text = Format(Now, "dd/mm/yyyy") ' fecha de movimiento
            txtCodigo(1).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento
         
            FrameCobrosVisible True, H, W
            Pb1.visible = False
        Case 1 ' cancelacion de avnics
            Me.Caption = "Cancelación AVNICS"
        
            FrameCancelacionVisible True, H, W
                    
            'fecha de cancelacion
            txtCodigo(15).Text = Format(Now, "dd/mm/yyyy")
            'txtCodigo(10).Text = Format(vParamAplic.Porcrete, "##0.00")
            
            'periodo de liquidacion
            Select Case Month(CDate(txtCodigo(15).Text))
                Case 1 To 3
                    txtCodigo(18).Text = "01/01/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
                    txtCodigo(19).Text = "31/03/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
                Case 4 To 6
                    txtCodigo(18).Text = "01/04/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
                    txtCodigo(19).Text = "30/06/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
                Case 7 To 9
                    txtCodigo(18).Text = "01/07/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
                    txtCodigo(19).Text = "30/09/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
                Case 10 To 12
                    txtCodigo(18).Text = "01/10/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
                    txtCodigo(19).Text = "31/12/" & Format(Year(CDate(txtCodigo(15).Text)), "0000")
            End Select
                        
            VisualizarFrameResultado False
                
        Case 2 ' ayuda modelo 123
            Me.Caption = "Ayuda Modelo 123"
            FrameAyudaMod123Visible True, H, W
            
            FrameResultados.visible = False
            
            
            
            
     End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmavn_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/MM/yyyy")
End Sub


Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
'Concepto
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmDi_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(9).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(9).Text = Format(txtCodigo(9).Text, "00")
    txtNombre(9).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtCodigo(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(IndCodigo).Text = Format(txtCodigo(IndCodigo).Text, "000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas

    Select Case Index
        Case 0, 1
            IndCodigo = Index
        Case 2
            IndCodigo = 15
        Case 3, 4
            IndCodigo = Index + 15
        Case 5, 6
            IndCodigo = Index + 18
    End Select

    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtCodigo(IndCodigo).Text <> "" Then frmC.Fecha = CDate(txtCodigo(IndCodigo).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    PonFoco txtCodigo(IndCodigo)

End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 3 ' forma de pago de la tesoreria
            AbrirFrmForpaConta (Index)
        Case 4 'cuenta contable
            AbrirFrmCuentas (Index)
        Case 0
            AbrirFrmCuentas (5)
        Case 1
            AbrirFrmCuentas (6)
            
        Case 2 'concepto
            IndCodigo = 7
            Set frmCon = New frmConceptos
            frmCon.DatosADevolverBusqueda = "0|"
            frmCon.Show vbModal
            Set frmCon = Nothing
            
        Case 5 'concepto
            IndCodigo = 8
            Set frmCon = New frmConceptos
            frmCon.DatosADevolverBusqueda = "0|"
            frmCon.Show vbModal
            Set frmCon = Nothing
        
        Case 6 'diario
            'Tipos diario
            Set frmDi = New frmTiposDiario
            frmDi.DatosADevolverBusqueda = "0"
            frmDi.Show vbModal
            Set frmDi = Nothing
            PonFoco txtCodigo(9)
        
        Case 7 'AVNICS
            AbrirFrmAvnics (16)
        
    End Select
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub AbrirFrmAvnics(Indice As Integer)
    IndCodigo = Indice
    Set frmavn = New frmAVNAvnics
    frmavn.DatosADevolverBusqueda = "0|4|"
    frmavn.DeConsulta = True
    frmavn.CodigoActual = txtCodigo(IndCodigo)
    frmavn.Show vbModal
    Set frmavn = Nothing
End Sub

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtcodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtcodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 23: KEYFecha KeyAscii, 5 'fecha
            Case 24: KEYFecha KeyAscii, 6 'fecha
            
            'cancelacion avnics
            Case 16: KEYBusqueda KeyAscii, 7 'codigo avnic
            Case 18: KEYFecha KeyAscii, 3 'fecha desde periodo liquidacion
            Case 19: KEYFecha KeyAscii, 4 'fecha hasta periodo liquidacion
            Case 15: KEYFecha KeyAscii, 2 'fecha cancelacion
            'contabilizacion intereres
            Case 0: KEYFecha KeyAscii, 0 'fecha
            Case 4: KEYBusqueda KeyAscii, 4 'cta prevista
            Case 1: KEYFecha KeyAscii, 1 'fecha vto
            Case 3: KEYBusqueda KeyAscii, 3 'forma de pago
            Case 5: KEYBusqueda KeyAscii, 0 'cta retencion
            Case 6: KEYBusqueda KeyAscii, 1 'cta gasto
            Case 7: KEYBusqueda KeyAscii, 2 'concepto debe
            Case 8: KEYBusqueda KeyAscii, 5 'concepto haber
            Case 9: KEYBusqueda KeyAscii, 6 'diario
            
            Case 4: KEYBusqueda KeyAscii, 4 'concepto al debe
            Case 0: KEYFecha KeyAscii, 0 'fecha de movimiento
            Case 1: KEYFecha KeyAscii, 1 'fecha de vencimiento
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub txtcodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim I As Integer
Dim Sql As String
Dim RC As String

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    If txtCodigo(Index) = "" Then Exit Sub
    
    If Not PerderFocoGnral(txtCodigo(Index), 3) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 3 ' FORMA DE PAGO DE LA CONTABILIDAD
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago no existe. Reintroduzca.", vbExclamation
            End If
            
        Case 4, 5, 6 ' CUENTA CONTABLE
            RC = txtCodigo(Index)
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                txtCodigo(Index).Text = RC
                txtNombre(Index).Text = Sql
            Else
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
                PonFoco txtCodigo(Index)
            End If

        Case 0, 1 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            If Index = 0 Then
                txtCodigo(1).Text = txtCodigo(0).Text
            End If
            
        Case 7, 8 'CONCEPTOS
                If Not IsNumeric(txtCodigo(Index).Text) Then
                    MsgBoxA "El concepto debe de ser numérico", vbExclamation
                    PonFoco txtCodigo(Index)
                    Exit Sub
                End If
                If Val(txtCodigo(Index).Text) >= 900 Then
                    MsgBoxA "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
                    txtNombre(Index).Text = ""
                    txtCodigo(Index).Text = ""
                    PonFoco txtCodigo(Index)
                    Exit Sub
                End If
                
                Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtCodigo(Index).Text, "N")
                If Sql = "" Then
                    MsgBoxA "Concepto NO encontrado: " & txtCodigo(Index).Text, vbExclamation
                    txtCodigo(Index).Text = ""
                    PonFoco txtCodigo(Index)
                Else
                    txtNombre(Index).Text = Sql
                End If
        
        Case 9 'diario
            If Not IsNumeric(txtCodigo(Index).Text) Then
                MsgBoxA "Tipo de diario no es numérico: " & txtCodigo(Index).Text, vbExclamation
                txtCodigo(Index).Text = ""
                txtNombre(Index).Text = ""
                PonFoco txtCodigo(Index)
                Exit Sub
            End If
             Sql = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtCodigo(Index).Text, "N")
             If Sql = "" Then
                    Sql = "Diario no encontrado: " & txtCodigo(Index).Text
                    txtCodigo(Index).Text = ""
                    txtNombre(Index).Text = ""
                    MsgBoxA Sql, vbExclamation
                    PonFoco txtCodigo(Index)
            End If
            txtCodigo(Index).Text = Val(txtCodigo(Index))
            txtNombre(Index).Text = Sql
            
            
        Case 16 'codigo
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "avnic", "nombrper", "codavnic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 15, 18, 19 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 10, 14 'porcentaje de penalizacion
            If txtCodigo(Index).Text <> "" Then PonerFormatoDecimal txtCodigo(Index), 4
                          
        Case 23, 24 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.top = -90
        Me.FrameCobros.Left = 0
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub


Private Sub FrameCancelacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCancelacion.visible = visible
    If visible = True Then
        Me.FrameCancelacion.top = -90
        Me.FrameCancelacion.Left = 0
        W = Me.FrameCancelacion.Width
        H = Me.FrameCancelacion.Height
    End If
    
End Sub


Private Sub FrameAyudaMod123Visible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameAyudaMod123.visible = visible
    If visible = True Then
        Me.FrameAyudaMod123.top = -90
        Me.FrameAyudaMod123.Left = 0
        W = Me.FrameAyudaMod123.Width
        H = Me.FrameAyudaMod123.Height
    End If
    
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub AbrirFrmCuentas(Indice As Integer)
    IndCodigo = Indice
    Set frmCtas = New frmColCtas
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(Indice As Integer)
    IndCodigo = Indice
    Set frmFpa = New frmFormaPago
    frmFpa.DatosADevolverBusqueda = "0|1|"
    frmFpa.Show vbModal
    Set frmFpa = Nothing
End Sub
 

Private Sub ContabilizarIntereses(cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTABLA As String

    Sql = "CONINT" 'contabilizar CALCULO DE INTERESES

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(True, Sql, "1") Then
        MsgBox "No se pueden Contabilizar Cálculo de Intereses. Hay otro usuario contabilizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    ' nuevo
    B = CrearTMPErrComprob()
    If Not B Then Exit Sub
    
    
    'comprobar que todas las CUENTAS de codigos avnic existen
    'en la Conta: savnic.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Retención ..."
    B = ComprobarCtaContable(cadTABLA, 1, "movim.fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and intconta = 0")
    IncrementarProgres Me.Pb1, 33
    Me.Refresh
    If Not B Then
        frmAVNInformes.OpcionListado = 2
        frmAVNInformes.Show vbModal
        Exit Sub
    End If

    
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Cierre: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Asiento en Contabilidad..."
    
    
    cadwhere = "fechamov = " & DBSet(txtCodigo(0).Text, "F") & " and intconta = 0"
    B = PasarCalculoAContab(cadwhere)
    
    If Not B Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.Opcion = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If
    
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Orden1 As String
Dim Orden2 As String

   B = True

   Select Case OpcionListado
        Case 0 ' calculo de intereses
            If txtCodigo(0).Text = "" And B Then
                 MsgBoxA "Introduzca la Fecha de movimiento a contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(0)
            Else
                 ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
                  Orden1 = ""
                  Orden1 = vParam.fechaini ' DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
             
                  Orden2 = ""
                  Orden2 = vParam.fechafin 'DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
                  FIni = CDate(Orden1)
                  FFin = CDate(Orden2)
                  If Not (CDate(Orden1) <= CDate(txtCodigo(0).Text) And CDate(txtCodigo(0).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
                     MsgBoxA "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
                     B = False
                     PonFoco txtCodigo(0)
                  End If
            End If
             
            If txtCodigo(1).Text = "" And B Then
                 MsgBoxA "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(1)
            End If
             
            If txtCodigo(3).Text = "" And B Then
                 MsgBoxA "Introduzca la Forma de Pago para contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(3)
            End If
            
            If txtCodigo(5).Text = "" And B Then
                 MsgBoxA "Introduzca la cuenta de retención para contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(5)
            End If
            If txtCodigo(6).Text = "" And B Then
                 MsgBoxA "Introduzca la cuenta de gasto para contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(6)
            End If
            If txtCodigo(7).Text = "" And B Then
                 MsgBoxA "Introduzca el Concepto al Debe para contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(7)
            End If
            If txtCodigo(8).Text = "" And B Then
                 MsgBoxA "Introduzca el Concepto al Haber para contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(8)
            End If
            If txtCodigo(9).Text = "" And B Then
                 MsgBoxA "Introduzca el Diario para contabilizar.", vbExclamation
                 B = False
                 PonFoco txtCodigo(9)
            End If
            
        Case 1 ' cancelacion
            ' introduce el codigo avnic
            If txtCodigo(16).Text = "" Then
                MsgBoxA "El código avnic debe de tener un valor. Reintroduzca.", vbExclamation
                B = False
                PonFoco txtCodigo(16)
            End If
            
            ' el periodo de liquidacion debe de tener un valor
            If txtCodigo(18).Text = "" Or txtCodigo(19).Text = "" Then
                MsgBoxA "El período de liquidación debe de tener valores. Reintroduzca.", vbExclamation
                B = False
                PonFoco txtCodigo(18)
            End If
            
            ' la fecha de cancelacion no debe de ser nula
            If txtCodigo(15).Text = "" Then
                MsgBoxA "La fecha de cancelación debe de tener un valor. Reintroduzca.", vbExclamation
                B = False
                PonFoco txtCodigo(15)
            Else
                ' la fecha de cancelacion ha de estar comprendida dentro del periodo de liquidacion
                If Not (CDate(txtCodigo(15).Text) >= CDate(txtCodigo(18).Text) And CDate(txtCodigo(15).Text) <= CDate(txtCodigo(19).Text)) Then
                    MsgBoxA "La fecha de cancelación del avnic debe de estar comprendida dentro del período de liquidación", vbExclamation
                    B = False
                    PonFoco txtCodigo(15)
                End If
            End If
   
   End Select
   DatosOK = B
   
End Function

Private Function PasarCalculoAContab(cadwhere As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim I As Integer
Dim NumLinea As Long
Dim Mc As Contadores
Dim Numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim cadMen As String
Dim cad As String
Dim CtaDifer As String
Dim codmacta As String

    On Error GoTo EPasarCal

    PasarCalculoAContab = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
    Sql = "SELECT count(*)" & _
          " FROM movim " & _
          "WHERE " & cadwhere
             
    NumLinea = TotalRegistros(Sql)
    
    If NumLinea = 0 Then Exit Function
    
    NumLinea = NumLinea * 3
    
    If NumLinea > 0 Then
        NumLinea = NumLinea + 1
        
        CargarProgres Me.Pb1, NumLinea
        
        Conn.BeginTrans
        
        Set Mc = New Contadores
        
        If Mc.ConseguirContador("0", (CDate(txtCodigo(0).Text) <= CDate(FFin)), True) = 0 Then
        
        Obs = "Contabilizacion de Cálculo de Intereses AVNICS de fecha " & Format(txtCodigo(0).Text, "dd/mm/yyyy")

    
        'Insertar en la conta Cabecera Asiento
        B = InsertarCabAsientoDia(txtCodigo(9), Mc.Contador, txtCodigo(0).Text, Obs, cadMen, cConta)
        cadMen = "Insertando Cab. Asiento: " & cadMen
        
        If B Then
            Sql = "SELECT codavnic, timporte, timport1, timport2 " & _
                  " FROM movim " & _
                  " WHERE " & cadwhere
            
            Set Rs = New ADODB.Recordset
            
            Rs.Open Sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
            
            I = 0
            ImporteD = 0
            ImporteH = 0
            
            Ampliacion = "Int.AVNICS "
            ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtCodigo(7), "N")) & " " & Ampliacion
            ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtCodigo(8), "N")) & " " & Ampliacion
            
            
            If Not Rs.EOF Then Rs.MoveFirst
            While Not Rs.EOF And B
                codmacta = ""
                codmacta = DevuelveDesdeBDNew(cConta, "avnic", "codmacta", "codavnic", Rs.Fields(0).Value, "N", , "anoejerc", Year(CDate(txtCodigo(0).Text)), "N")
                
                Numdocum = "Av-" & Format(DBLet(Rs!codavnic, "N"), "000000")
                ' ******************IMPORTE BRUTO
                I = I + 1
                
                cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(txtCodigo(6), "T") & "," & DBSet(Numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If Rs.Fields(2).Value > 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(txtCodigo(7), "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteD = ImporteD + CCur(Rs.Fields(2).Value)
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(txtCodigo(8), "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet((Rs.Fields(2).Value * -1), "N") & "," & ValorNulo & "," & DBSet(codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(Rs.Fields(2).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & I
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & I & " de " & NumLinea & ")"
                Me.Refresh
                
                ' ******************RETENCION
                I = I + 1
                
                cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(txtCodigo(5), "T") & "," & DBSet(Numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If Rs.Fields(3).Value > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(txtCodigo(8), "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(Rs.Fields(3).Value, "N") & "," & ValorNulo & "," & DBSet(codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(Rs.Fields(3).Value)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(txtCodigo(7), "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((Rs.Fields(3).Value * -1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & DBSet(codmacta, "T") & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(Rs.Fields(3).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & I
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & I & " de " & NumLinea & ")"
                Me.Refresh
                
                ' ******************IMPORTE NETO
                I = I + 1
                
                cad = "1," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(I, "N") & "," & DBSet(codmacta, "T") & "," & DBSet(Numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If Rs.Fields(1).Value > 0 Then
                    ' importe al haber en positivo
                    cad = cad & DBSet(txtCodigo(8), "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet(Rs.Fields(1).Value, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + CCur(Rs.Fields(1).Value)
                Else
                    ' importe al debe en positivo, cambiamos el signo
                    cad = cad & DBSet(txtCodigo(7), "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((Rs.Fields(1).Value * -1), "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(Rs.Fields(1).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                B = InsertarLinAsientoDia(cad, cadMen, cConta)
                cadMen = "Insertando Lin. Asiento: " & I
            
                B = InsertarEnTesoreriaNew(txtCodigo(0).Text, txtCodigo(1).Text, Rs.Fields(0).Value, Year(CDate(txtCodigo(0).Text)), txtCodigo(4).Text, txtCodigo(2).Text, txtCodigo(3).Text, cadMen)
                cadMen = "Insertando en Tesoreria: "
               
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & I & " de " & NumLinea & ")"
                Me.Refresh
                
            
                Rs.MoveNext
            Wend
            Rs.Close
            
            If B Then
                'Poner intconta=1 en ariagroutil.movim
                B = ActualizarMovimientos(cadwhere, cadMen)
                cadMen = "Actualizando Movimientos: " & cadMen
            End If
        End If
    End If
   End If
   
EPasarCal:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Integrando Asiento a Contabilidad", Err.Description
    End If
    If B Then
        Conn.CommitTrans
        PasarCalculoAContab = True
    Else
        Conn.RollbackTrans
        PasarCalculoAContab = False
    End If
End Function


Private Function CalculoPenalizacion(cadwhere As String) As Boolean
Dim Sql As String
Dim DiasPen As Integer
Dim DiasInt As Integer
Dim Intereses As Currency
Dim Penalizacion As Currency
Dim bruto As Currency
Dim neto As Currency
Dim retencion As Currency

Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim Mens As String

    On Error GoTo eCalculoPenalizacion

    Set Rs = New ADODB.Recordset

    CalculoPenalizacion = False
    
    Conn.BeginTrans
    B = True
    Sql = "select * from avnic where " & cadwhere
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    txtCodigo(13).Text = Format(0, "###,###,##0.00")
    txtCodigo(12).Text = Format(0, "###,###,##0.00")
    txtCodigo(11).Text = Format(0, "###,###,##0.00")

    If Not Rs.EOF Then
        DiasInt = CDate(txtCodigo(15).Text) - CDate(txtCodigo(18).Text) + 1
        DiasPen = CDate(txtCodigo(19).Text) - CDate(txtCodigo(15).Text)
            
        Intereses = Round2(DBLet(Rs!Importes, "N") * DBLet(Rs!Porcinte, "N") / 100 * DiasInt / 365, 2)
        Penalizacion = Round2(DBLet(Rs!Importes, "N") * ImporteSinFormato(txtCodigo(14).Text) / 100 * DiasPen / 365, 2)
            
        bruto = Intereses - Penalizacion
        If bruto > 0 Then
            retencion = Round2(bruto * ImporteSinFormato(txtCodigo(10).Text) / 100, 2)
            neto = bruto - retencion
            
            'cargamos las variables para visualizarlas posteriormente
            txtCodigo(13).Text = Format(bruto, "###,###,##0.00")
            txtCodigo(12).Text = Format(retencion, "###,###,##0.00")
            txtCodigo(11).Text = Format(neto, "###,###,##0.00")
            
            Mens = "Insertar movimiento y actualizar avnic "
            B = InsertarMovimiento(Rs!codavnic, Rs!anoejerc, bruto, retencion, neto, Mens)
        End If
        If B Then
            Sql = "update avnic set codialta = 2 "
            Sql = Sql & "where " & cadwhere
            Conn.Execute Sql
        End If
    End If
    If B Then
        CalculoPenalizacion = B
        Conn.CommitTrans
        Exit Function
    End If
eCalculoPenalizacion:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, Mens
        Conn.RollbackTrans
    End If
End Function

Private Function InsertarMovimiento(codavnic As Long, anoejerc As Integer, bruto As Currency, retencion As Currency, neto As Currency, ByRef Mens As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarMovimiento
    
    InsertarMovimiento = False
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cConta, "movim", "codavnic", "codavnic", CStr(codavnic), "N", , "fechamov", txtCodigo(15).Text, "F", "anoejerc", CStr(anoejerc), "N")
    If Sql = "" Then
        Sql = "insert into movim (codavnic, fechamov, concepto, timporte, intconta, anoejerc, timport1, timport2) "
        Sql = Sql & "values (" & DBSet(codavnic, "N") & "," & DBSet(txtCodigo(15).Text, "F") & "," & DBSet(txtCodigo(17).Text, "T") & ","
        Sql = Sql & DBSet(neto, "N") & ",0," & DBSet(anoejerc, "N") & ","
        Sql = Sql & DBSet(bruto, "N") & "," & DBSet(retencion, "N") & ")"
        
        Conn.Execute Sql
    End If
    Sql = "update avnic set imporper = imporper + " & DBSet(neto, "N") & ","
    Sql = Sql & " imporret = imporret + " & DBSet(retencion, "N")
    Sql = Sql & " where codavnic = " & DBSet(codavnic, "N") & " and anoejerc = " & DBSet(anoejerc, "N")
    
    Conn.Execute Sql
    
    InsertarMovimiento = True
    Exit Function
    
eInsertarMovimiento:
    If Err.Number <> 0 Then
        Mens = Mens & vbCrLf & Err.Description
    End If
End Function


Private Sub VisualizarFrameResultado(B As Boolean)
Dim I As Integer

    FrameResultado.visible = B
    FrameResultado.Enabled = B
    CmdAcepCancelacion.visible = Not B
    CmdAcepCancelacion.Enabled = Not B
    If B Then
        cmdCancel(1).Caption = "Salir"
        For I = 10 To 19
            txtCodigo(I).Enabled = False
        Next I
        imgBuscar(7).Enabled = False
        ImgFec(2).Enabled = False
        ImgFec(3).Enabled = False
        ImgFec(4).Enabled = False
    Else
        cmdCancel(1).Caption = "Cancelar"
    End If
    
End Sub


Private Function TotalReceptores(cadwhere As String, Mens As String) As Long
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String

    On Error GoTo eTotalReceptores
    
    
    TotalReceptores = 0
    BorrarTMPavnics
    
    If CrearTMPavnics Then
        Set Rs = New ADODB.Recordset
        Sql = "select distinct nifperso, nifrepre from movim, avnic where " & cadwhere
        Sql = Sql & " and movim.codavnic = avnic.codavnic and movim.anoejerc = avnic.anoejerc "
        
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            If DBLet(Rs.Fields(0).Value, "T") <> "" Then
                Sql2 = "insert into tmpavnics (nif) values (" & DBSet(Rs.Fields(0).Value, "T") & ")"
                Conn.Execute Sql2
            End If
            If DBLet(Rs.Fields(1).Value, "T") <> "" Then
                Sql2 = "insert into tmpavnics (nif) values (" & DBSet(Rs.Fields(1).Value, "T") & ")"
                Conn.Execute Sql2
            End If
            
            Rs.MoveNext
        Wend
        Rs.Close
        Sql = "select count(distinct nif) from tmpavnics"
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        If Not Rs.EOF Then TotalReceptores = DBLet(Rs.Fields(0).Value, "N")
        
        Set Rs = Nothing
    End If

eTotalReceptores:
    If Err.Number <> 0 Then
        Mens = "Error en el calculo de Total de Receptores " & Err.Description
    End If
End Function
        
Public Sub BorrarTMPavnics()
On Error Resume Next
    Conn.Execute " DROP TABLE IF EXISTS tmpavnics;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CrearTMPavnics() As Boolean
'Crea una temporal donde insertara los nifs tanto de personas como de representantes
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPavnics = False
    
    Sql = "CREATE TEMPORARY TABLE tmpavnics ( nif varchar(9) )"
    Conn.Execute Sql
     
    CrearTMPavnics = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPavnics = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpavnics;"
        Conn.Execute Sql
    End If
End Function

Private Sub ValoresPorDefecto()
Dim MesAnt As Byte
Dim AnoAnt As Integer
Dim Fec As Date

    AnoAnt = Year(Now)
    MesAnt = Month(Now) - 1
    If MesAnt = 0 Then
        MesAnt = 12
        AnoAnt = AnoAnt - 1
    End If
    
    txtCodigo(0).Text = "01/" & Format(MesAnt, "00") & "/" & Format(AnoAnt, "0000")
    ' calculo para el mes siguiente
    MesAnt = MesAnt + 1
    If MesAnt = 13 Then
        MesAnt = 1
        AnoAnt = AnoAnt + 1
    End If
    Fec = CDate("01/" & Format(MesAnt, "00") & "/" & Format(AnoAnt, "0000"))
    ' al primer dia de este mes le quitamos 1 para que nos de el ultimo dia del mes anterior
    Fec = Fec - 1
    
    txtCodigo(1).Text = Format(Fec, "dd/mm/yyyy")
    
End Sub

