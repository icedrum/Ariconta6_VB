VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAVNModelo193 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3240
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtCodigo 
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
         Index           =   8
         Left            =   1350
         TabIndex        =   27
         Tag             =   "imgConcepto"
         Top             =   2700
         Width           =   765
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
         Index           =   0
         Left            =   1335
         MaxLength       =   6
         TabIndex        =   24
         Top             =   765
         Width           =   830
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
         Left            =   1335
         MaxLength       =   6
         TabIndex        =   25
         Top             =   1140
         Width           =   830
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
         Index           =   0
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   765
         Width           =   4170
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
         Index           =   1
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1140
         Width           =   4170
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   6
         Left            =   1350
         TabIndex        =   26
         Tag             =   "imgConcepto"
         Top             =   1935
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Porcentaje Retencion"
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
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   48
         Top             =   2385
         Width           =   2400
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
         Index           =   0
         Left            =   405
         TabIndex        =   40
         Top             =   1935
         Width           =   690
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
         Left            =   360
         TabIndex        =   29
         Top             =   765
         Width           =   690
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
         Left            =   360
         TabIndex        =   28
         Top             =   1140
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1050
         MouseIcon       =   "frmAVNModelo193.frx":0000
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar avnic"
         Top             =   765
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1035
         MouseIcon       =   "frmAVNModelo193.frx":0152
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar avnic"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label lblCuentas 
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
         Left            =   2520
         TabIndex        =   18
         Top             =   5190
         Width           =   4095
      End
      Begin VB.Label lblCuentas 
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
         Left            =   2520
         TabIndex        =   17
         Top             =   4800
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Código Avnics"
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
         Height          =   330
         Index           =   7
         Left            =   315
         TabIndex        =   16
         Top             =   450
         Width           =   1770
      End
      Begin VB.Label Label3 
         Caption         =   "Ejercicio"
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
         Height          =   330
         Index           =   6
         Left            =   315
         TabIndex        =   15
         Top             =   1575
         Width           =   960
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
      Height          =   5910
      Left            =   7050
      TabIndex        =   19
      Top             =   0
      Width           =   5535
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   5190
         Left            =   135
         TabIndex        =   20
         Top             =   360
         Width           =   5265
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
            Index           =   5
            Left            =   2250
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   30
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   90
            Width           =   555
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tipo de Soporte"
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
            Height          =   1140
            Left            =   45
            TabIndex        =   42
            Top             =   540
            Width           =   4650
            Begin VB.OptionButton Option1 
               Caption         =   "Presentación en Diskette"
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
               Index           =   0
               Left            =   585
               TabIndex        =   32
               Top             =   405
               Width           =   3750
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Presentación Telemática"
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
               Index           =   1
               Left            =   585
               TabIndex        =   33
               Top             =   720
               Width           =   3705
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Tipo de Declaración"
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
            Height          =   1545
            Left            =   0
            TabIndex        =   41
            Top             =   1800
            Width           =   4695
            Begin VB.OptionButton Option2 
               Caption         =   "Complementaria"
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
               Index           =   1
               Left            =   630
               TabIndex        =   35
               Top             =   780
               Width           =   2940
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Primera declaración"
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
               Height          =   195
               Index           =   0
               Left            =   630
               TabIndex        =   34
               Top             =   450
               Width           =   3705
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Sustitutiva"
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
               Left            =   630
               TabIndex        =   36
               Top             =   1125
               Width           =   2040
            End
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   540
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   37
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   4050
            Width           =   555
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
            Index           =   2
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   39
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   4635
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
            Index           =   3
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   31
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   90
            Width           =   1305
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   7
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   38
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   4050
            Width           =   1305
         End
         Begin MSComctlLib.Toolbar ToolbarAyuda 
            Height          =   390
            Left            =   4725
            TabIndex        =   45
            Top             =   45
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
            Caption         =   "Número Justificación Declaración Sustitutiva"
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
            Index           =   2
            Left            =   45
            TabIndex        =   44
            Top             =   3690
            Width           =   4920
         End
         Begin VB.Label Label4 
            Caption         =   "Teléfono"
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
            Index           =   3
            Left            =   45
            TabIndex        =   43
            Top             =   4635
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Número Justificante"
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
            Left            =   135
            TabIndex        =   21
            Top             =   135
            Width           =   2415
         End
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
      Left            =   11370
      TabIndex        =   2
      Top             =   6210
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
      Left            =   9810
      TabIndex        =   0
      Top             =   6210
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
      Left            =   90
      TabIndex        =   1
      Top             =   6120
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
      Left            =   45
      TabIndex        =   3
      Top             =   3300
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
         TabIndex        =   14
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   13
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   12
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "A.E.A.T."
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   60
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   255
      Left            =   1485
      TabIndex        =   46
      Top             =   6075
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   450
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
      Height          =   240
      Index           =   1
      Left            =   1485
      TabIndex        =   47
      Top             =   6390
      Width           =   7950
   End
End
Attribute VB_Name = "frmAVNModelo193"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 411


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmavn As frmAVNAvnics 'Avnics
Attribute frmavn.VB_VarHelpID = -1



Private Sql As String
Dim Cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String
Dim Tablas As String

Dim UltimoPeriodoLiquidacion As Boolean
Dim C2 As String

Dim FechaI As String
Dim FechaF As String
Dim Rs As ADODB.Recordset
Dim Importe As Currency
Dim PrimeraVez As Boolean

Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadSelect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False

End Sub


Private Sub cmdAccion_Click(Index As Integer)
Dim b As Boolean
Dim ConCli As Integer 'Clientes
Dim ConPro As Integer  'proveedores

Dim indRPT As String
Dim nomDocu As String
Dim cadWhere As String
Dim NRegs As Long

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    
    Screen.MousePointer = vbHourglass

    Sql = "SELECT  count(*) "
    Sql = Sql & " from avnic where "
    cadWhere = "anoejerc = " & DBSet(txtCodigo(6).Text, "N")
    If txtCodigo(0).Text <> "" Then cadWhere = cadWhere & " and avnic.codavnic >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then cadWhere = cadWhere & " and avnic.codavnic <= " & DBSet(txtCodigo(1).Text, "N")
    Sql = Sql & cadWhere
    
    NRegs = TotalRegistros(Sql)

    If NRegs <> 0 Then
        If GeneraFichero(cadWhere) Then
            CopiarFicheroHaciend3
        End If
    Else
        MsgBoxA "No hay datos entre esos límites.", vbExclamation
    End If
    
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Function GeneraFichero(cadWhere As String) As Boolean
Dim NFich1 As Integer
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Sql As String
Dim v_Hayreg As Integer
Dim NRegs As Long
Dim b As Boolean
Dim Mens As String

Dim impbase As Currency
Dim ImpReten As Currency
Dim TotalReg As Currency

Dim v_import As String
Dim v_impret As String
Dim t_import As Currency
Dim t_impret As Currency

Dim RsLin As ADODB.Recordset
Dim Linea As String


    On Error GoTo EGen
    
    GeneraFichero = False

    Mens = "Cargando la tabla temporal."
    b = CrearTMPavnicsNew(cadWhere, Mens)
    
    If b Then
        Mens = "Calculando totales."
        b = CalcularTotalesNew(impbase, ImpReten, TotalReg, Mens)
    End If
    
    If b Then
        Me.lblProgres(1).Caption = "Cargando fichero..."
        NRegs = TotalRegistros("select count(*) from tmptempo")
        
        Pb1.visible = True
        Pb1.Max = NRegs
        Pb1.Value = 0
                        
                        
        Set RsLin = New ADODB.Recordset
                        
        Linea = "Select * from empresa2"
        RsLin.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        
        NFich1 = FreeFile
        Open App.Path & "\mod193.txt" For Output As #NFich1
    
        Cad = "1193"
        Cad = Cad & txtCodigo(6).Text ' anoejerc
        Cad = Cad & RellenaABlancos(RsLin!nifempre, True, 9)
        Cad = Cad & RellenaABlancos(vEmpresa.NombreEmpresaOficial, True, 40)
        
        If Option1(0).Value Then 'tipo de documento
            Cad = Cad & "D"
        Else
            Cad = Cad & "T"
        End If
        Cad = Cad & Format(CCur(txtCodigo(2).Text), "000000000") ' telefono
        '[Monica]25/01/2012: cambiado antes el nompresi era de 39, ahora de 40
        Cad = Cad & RellenaABlancos(DBLet(RsLin!contacto, "T"), True, 40) ' nompresi
        
        Cad = Cad & Format(txtCodigo(5).Text, "000") 'numero
        
        Cad = Cad & Format(ComprobarCero(txtCodigo(3).Text), "0000000000") 'justificante
        
        'tipo de declaracion
        If Option2(0).Value Then
            Cad = Cad & "  " & String(13, "0")
        End If
        If Option2(1).Value Then
            Cad = Cad & "C " & String(13, "0")
        End If
        If Option2(2).Value Then
            Cad = Cad & " S" & Format(CCur(txtCodigo(4).Text), "000") & Format(CCur(txtCodigo(7).Text), "0000000000")
        End If
            
        Cad = Cad & Format(TotalReg, "000000000")
        Cad = Cad & Format(Round2(impbase * 100, 0), "000000000000000")
        Cad = Cad & Format(Round2(ImpReten * 100, 0), "000000000000000")
        Cad = Cad & Format(Round2(ImpReten * 100, 0), "000000000000000")
        
        '[Monica]25/01/2012: ahora son blancos
        Cad = Cad & Space(30) ' 30 blancos
        Cad = Cad & String(15, "0") ' gastos
        Cad = Cad & " " ' naturaleza del declarante
        Cad = Cad & Space(265)
        
        
        Print #NFich1, Cad
    
    
        Set Rs = New ADODB.Recordset
        
        'partimos de la tabla de historico de facturas
        Sql = "SELECT * from tmptempo order by nifperso " ' codavnic, tipocodi"
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        b = True
        v_Hayreg = 0
        While Not Rs.EOF And b
            v_Hayreg = 1
            
            Pb1.Value = Pb1.Value + 1
            
            
            Cad = "2193"
            Cad = Cad & txtCodigo(0).Text 'p.5
            Cad = Cad & RellenaABlancos(DBLet(RsLin!nifempre, "T"), True, 9) 'p.9
            
            If Trim(DBLet(Rs!nifrepre, "T")) <> "" And Not IsNull(Rs!nifrepre) Then 'p.18
                Cad = Cad & Space(9) 'p.27
                Cad = Cad & RellenaABlancos(DBLet(Rs!nifrepre, "T"), True, 9)
            Else
                Cad = Cad & RellenaABlancos(DBLet(Rs!nifperso, "T"), True, 9) 'p.18
                Cad = Cad & Space(9) 'p.27
            End If
                
            Cad = Cad & RellenaABlancos(DBLet(Rs!nombrper, "T"), True, 40) 'p.36
            Cad = Cad & " " 'p.76
            Cad = Cad & Mid(RellenaABlancos(DBLet(Rs!CodPobla, "T"), True, 6), 1, 2) 'p.77
            Cad = Cad & "1" 'p.79
            Cad = Cad & RellenaABlancos(DBLet(RsLin!nifempre, "T"), True, 12) 'p.80
            Cad = Cad & "B" 'p.92
            Cad = Cad & "06" 'p.93
            Cad = Cad & "1" 'p.95
            Cad = Cad & "O" 'p.96
            Cad = Cad & Space(20) 'p.97
            Cad = Cad & " " 'p.117
            Cad = Cad & "0000" 'p.118 '[Monica]25/01/2012: antes esto:"0" & Mid(txtCodigo(0).Text, 2, 3)
            Cad = Cad & "1" 'p.122
            Cad = Cad & Format(Round2(DBLet(Rs!ImporPer, "N") * 100, 0), "0000000000000") 'p.123
            Cad = Cad & Space(3) '[Monica]25/01/2012: antes esto:"000" 'p.136
            Cad = Cad & String(13, "0") 'p.139
            Cad = Cad & Format(Round2(DBLet(Rs!ImporPer, "N") * 100, 0), "0000000000000") 'p.152
            Cad = Cad & Format(Round2(ImporteSinFormato(txtCodigo(8)) * 100, 0), "0000") 'p.165
            Cad = Cad & Format(Round2(DBLet(Rs!ImporRet, "N") * 100, 0), "0000000000000") 'p.169
            '[Monica]20/01/2017:
            Cad = Cad & String(11, "0") 'p.182 penalizaciones
            Cad = Cad & Space(15) '[Monica]25/01/2012: antes esto: Repeat("0", 13) 'p.193
            
            '++monica:24/01/2008 añadido por un tema de Alzira
            '[Monica]20/01/2017: camio en el modelo 193
            Cad = Cad & " "  'p.208
            Cad = Cad & String(8, "0") 'p.209
            Cad = Cad & String(8, "0") 'p.217
            Cad = Cad & String(12, "0") 'p.225 importe compensaciones
            Cad = Cad & String(12, "0") 'p.237 imprte garantias
            Cad = Cad & Space(252) 'p.249-500 relleno a blancos
            
            '++
            
            '[Monica]25/01/2012: ahora la longitud del registro es hasta la 500 (antes 250)
'            Cad = Cad & Space(250)
            'fin
            
            Print #NFich1, Cad
                
            Rs.MoveNext
        Wend
        
        Set RsLin = Nothing
        
    End If
    
EGen:
    Close (NFich1)
    Set Rs = Nothing
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, Err.Description & vbCrLf & Mens
    Else
        GeneraFichero = True
    End If

End Function


Private Function CrearTMPavnicsNew(cadWhere As String, ByRef Mens As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim ImporPer As Currency
Dim ImporRet As Currency
Dim Existe As String


    On Error GoTo ECrear
    
    CrearTMPavnicsNew = False
        
    Me.lblProgres(1).Caption = "Cargando la tabla temporal..."
        
    'Borrar la tabla temporal
    Sql = " DROP TABLE IF EXISTS tmptempo;"
    Conn.Execute Sql
    
    Sql = "CREATE TEMPORARY TABLE tmptempo ( "
'    SQL = "CREATE TABLE tmptempo ( "
    Sql = Sql & "nombrper char(40), "
    Sql = Sql & "nifperso char(9) ,"
    Sql = Sql & "nifrepre char(9) ,"
    Sql = Sql & "codpobla char(6) ,"
    Sql = Sql & "imporper decimal(12,2),"
    Sql = Sql & "imporret decimal(12,2))"
    
    Conn.Execute Sql
    
    Sql = "select nombrper, nifperso, nifrepre, codposta, imporper, imporret, codavnic, "
    Sql = Sql & "nifpers1, nombper1, codposta from avnic where " & cadWhere
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
'        If Trim(DBLet(RS!nifrepre, "T")) <> "" Then
        If Not IsNull(Rs!nifrepre) And Trim(DBLet(Rs!nifrepre, "T")) <> "" Then
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cConta, "tmptempo", "nifperso", "nifperso", Rs.Fields(1).Value, "T", , "nifrepre", Rs.Fields(2).Value, "T")
        
            If Existe = "" Then
                Sql = "insert into tmptempo ( nombrper, nifperso, codpobla, imporper,"
                Sql = Sql & " imporret) values ("
                Sql = Sql & DBSet(Rs.Fields(0).Value, "T") & "," 'nombrper
                Sql = Sql & DBSet(Rs.Fields(1).Value, "T") & "," 'nifperso
                Sql = Sql & DBSet(Rs.Fields(2).Value, "T") & "," 'nifrepre
                Sql = Sql & DBSet(Rs.Fields(3).Value, "T") & "," 'codpobla
                Sql = Sql & DBSet(Rs.Fields(4).Value, "N") & "," 'imporper
                Sql = Sql & DBSet(Rs.Fields(5).Value, "N") & ")" 'imporret
            Else
                Sql = "update tmptempo set imporper = imporper + " & DBSet(Rs.Fields(4).Value, "N")
                Sql = Sql & ", imporret = imporret + " & DBSet(Rs.Fields(5).Value, "N")
                Sql = Sql & " where nifperso = " & DBSet(Rs.Fields(1).Value, "T")
                Sql = Sql & " and nifrepre = " & DBSet(Rs.Fields(2).Value, "T")
            End If
            
            Conn.Execute Sql
'       ElseIf Trim(DBLet(RS!nifpers1, "T")) = "" Then
       ElseIf (IsNull(Rs!nifpers1) Or Trim(DBLet(Rs!nifpers1, "T")) = "") Then
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cConta, "tmptempo", "nifperso", "nifperso", Rs.Fields(1).Value, "T")
        
            If Existe = "" Then
                Sql = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                Sql = Sql & " imporret) values ("
                Sql = Sql & DBSet(Rs.Fields(0).Value, "T") & "," 'nombrper
                Sql = Sql & DBSet(Rs.Fields(1).Value, "T") & "," 'nifperso
                Sql = Sql & ValorNulo & "," 'nifrepre
                Sql = Sql & DBSet(Rs.Fields(3).Value, "T") & "," 'codpobla
                Sql = Sql & DBSet(Rs.Fields(4).Value, "N") & "," 'imporper
                Sql = Sql & DBSet(Rs.Fields(5).Value, "N") & ")" 'imporret
            Else
                Sql = "update tmptempo set imporper = imporper + " & DBSet(Rs.Fields(4).Value, "N")
                Sql = Sql & ", imporret = imporret + " & DBSet(Rs.Fields(5).Value, "N")
                Sql = Sql & " where nifperso = " & DBSet(Rs.Fields(1).Value, "T")
            End If
            
            Conn.Execute Sql
       Else
            ImporPer = Round2(DBLet(Rs!ImporPer, "N") * 0.5, 2)
            ImporRet = Round2(DBLet(Rs!ImporRet, "N") * 0.5, 2)
        
        Debug.Print Rs!nifrepre & "-" & Rs!nifperso & "-" & Rs!nifpers1
            
            '++monica: añadida la condicion de no añadir nifs duplicados
            Existe = ""
            Existe = DevuelveDesdeBDNew(cConta, "tmptempo", "nifperso", "nifperso", Rs.Fields(1).Value, "T")
        
            If Existe = "" Then
                Sql = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                Sql = Sql & " imporret) values ("
                Sql = Sql & DBSet(Rs.Fields(0).Value, "T") & "," 'nombrper
                Sql = Sql & DBSet(Rs.Fields(1).Value, "T") & "," 'nifperso
                Sql = Sql & ValorNulo & "," 'nifrepre
                Sql = Sql & DBSet(Rs.Fields(3).Value, "T") & "," 'codpobla
                Sql = Sql & DBSet(ImporPer, "N") & "," 'imporper
                Sql = Sql & DBSet(ImporRet, "N") & ")" 'imporret
            Else
                Sql = "update tmptempo set imporper = imporper + " & DBSet(ImporPer, "N")
                Sql = Sql & ", imporret = imporret + " & DBSet(ImporRet, "N")
                Sql = Sql & " where nifperso = " & DBSet(Rs.Fields(1).Value, "T")
            End If
            
            Conn.Execute Sql
            
            Existe = ""
            Existe = DevuelveDesdeBDNew(cConta, "tmptempo", "nifperso", "nifperso", Rs.Fields(7).Value, "T")
       
            If Existe = "" Then
                Sql = "insert into tmptempo ( nombrper, nifperso, nifrepre, codpobla, imporper,"
                Sql = Sql & " imporret) values ("
                Sql = Sql & DBSet(Rs.Fields(8).Value, "T") & "," 'nombrper
                Sql = Sql & DBSet(Rs.Fields(7).Value, "T") & "," 'nifpers1
                Sql = Sql & ValorNulo & "," 'nifrepre
                Sql = Sql & DBSet(Rs.Fields(3).Value, "T") & "," 'codpobla
                Sql = Sql & DBSet(ImporPer, "N") & "," 'imporper
                Sql = Sql & DBSet(ImporRet, "N") & ")" 'imporret
            Else
                Sql = "update tmptempo set imporper = imporper + " & DBSet(ImporPer, "N")
                Sql = Sql & ", imporret = imporret + " & DBSet(ImporRet, "N")
                Sql = Sql & " where nifperso = " & DBSet(Rs.Fields(7).Value, "T")
            End If
            
            Conn.Execute Sql
       
       
       End If
       Rs.MoveNext
    Wend

    CrearTMPavnicsNew = True
    Exit Function
ECrear:
     If Err.Number <> 0 Then
        Mens = Err.Description
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmptempo;"
        Conn.Execute Sql
    End If
End Function

Private Function CalcularTotalesNew(ByRef impbase As Currency, ByRef ImpReten As Currency, ByRef TotalReg As Currency, ByRef Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim v_import As String
Dim v_impret As String


    On Error GoTo eCalcularTotales

    CalcularTotalesNew = False
    
    impbase = 0
    ImpReten = 0
    TotalReg = 0
    
    Me.lblProgres(1).Caption = "Calculando totales..."
    
    Sql = "select * from tmptempo"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        TotalReg = TotalReg + 1
        impbase = impbase + DBLet(Rs!ImporPer, "N")
        ImpReten = ImpReten + DBLet(Rs!ImporRet, "N")
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    CalcularTotalesNew = True
    Exit Function
    
eCalcularTotales:
    If Err.Number <> 0 Then
        Mens = Err.Description
    End If
End Function





Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        optTipoSal(0).Enabled = False
        optTipoSal(2).Enabled = False
        optTipoSal(3).Enabled = False
        
        
        PushButtonImpr.Enabled = False
        PushButton2(1).Enabled = False
    
        Option1(1).Value = True
        Option2(0).Value = True
      
        
        optTipoSal(1).Value = True
      
      
        PonFoco txtCodigo(0)
    
    End If
    Screen.MousePointer = vbDefault

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
    Me.Caption = "Modelo 193"

    ' Datos de si es sustitutiva
    Label4(2).visible = False
    txtCodigo(4).visible = False
    txtCodigo(7).visible = False
     
    Me.Pb1.visible = False
    
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture 'frmppal.imgListImages16.ListImages(1).Picture
    Next i
    
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
    txtCodigo(6).Text = Format(Year(Now) - 1, "0000")
     
    '[Monica]12/06/2020: a piñon en ariagroutil
    txtCodigo(4).Text = 173
    txtCodigo(5).Text = 173
     
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    
    PonerDatosFicheroSalida
    
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub


Private Sub PonerDatosFicheroSalida()
Dim CADENA As String

    txtTipoSalida(1).Text = App.Path & "\Exportar\Mod193_" & Format(Me.txtCodigo(6), "0000") & ".txt"

End Sub




Private Sub frmavn_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'codigo avnics
            AbrirFrmAvnics (Index)
        
    End Select
    PonFoco txtCodigo(IndCodigo)

End Sub

Private Sub AbrirFrmAvnics(indice As Integer)
    IndCodigo = indice
    Set frmavn = New frmAVNAvnics
    frmavn.DatosADevolverBusqueda = "0|4|"
    frmavn.DeConsulta = True
    frmavn.CodigoActual = txtCodigo(IndCodigo)
    frmavn.Show vbModal
    Set frmavn = Nothing
End Sub



Private Sub Option2_Click(Index As Integer)
    If Index = 2 Then
        Label4(2).visible = (Option2(Index).Value = 1)
        txtCodigo(4).visible = (Option2(Index).Value = 1)
        txtCodigo(4).Enabled = (Option2(Index).Value = 1)
        txtCodigo(7).visible = (Option2(Index).Value = 1)
        txtCodigo(7).Enabled = (Option2(Index).Value = 1)
        
        If txtCodigo(4).visible Then
            txtCodigo(4).Text = ""
            txtCodigo(7).Text = ""
        End If
    End If
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
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
    Sql = Sql & " where " & cadSelect
    
    Conn.Execute Sql
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
End Function

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtCodigo(6).Text = "" Then
        MsgBoxA "Debe introducir un valor en el Ejercicio de la declaración.", vbExclamation
        PonFoco txtCodigo(6)
        DatosOK = False
        Exit Function
    End If
    If txtCodigo(2).Text = "" Then
        MsgBoxA "Debe introducir un valor en el campo Teléfono.", vbExclamation
        PonFoco txtCodigo(2)
        DatosOK = False
        Exit Function
    End If

    If txtCodigo(8).Text = "" Then
        MsgBoxA "Debe introducir el porcentaje de retención.", vbExclamation
        PonFoco txtCodigo(8)
        DatosOK = False
        Exit Function
    End If
   
    DatosOK = True


End Function


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5: KEYBusqueda KeyAscii, 5 'codigo desde
            Case 6: KEYBusqueda KeyAscii, 6 'codigo hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 6
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            
            optTipoSal_Click (1)
            
        Case 2
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000000")
        Case 3, 7
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000000")
        Case 1, 4
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        Case 0, 1 'codigo de avnics
            txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "avnic", "nombrper", "codavnic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 8 ' porcentaje de retencion
            PonerFormatoDecimal txtCodigo(Index), 4
            
    End Select
End Sub


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Sub CopiarFicheroHaciend3()
    On Error GoTo ECopiarFichero193
    Pb1.visible = False
    Me.lblProgres(1).visible = False
    MsgBoxA "El archivo se ha generado con exito.", vbInformation
    Sql = ""
    cd1.CancelError = True
    cd1.FileName = txtTipoSalida(1)
    cd1.ShowSave
    
    Sql = App.Path & "\mod193.txt"
    
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBoxA("El fichero ya existe. ¿Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Sql = ""
        End If
        If Sql <> "" Then
            FileCopy Sql, cd1.FileName
            MsgBoxA Space(20) & "Copia efectuada correctamente" & Space(20), vbInformation
        End If
    End If
    Exit Sub
ECopiarFichero193:
    If Err.Number <> 32755 Then MuestraError Err.Number, "Copiar fichero 193"
End Sub

