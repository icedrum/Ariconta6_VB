VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAVNInformes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Avnics"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12300
   Icon            =   "frmAVNInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameErrores 
      Height          =   5505
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   8835
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
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
         Left            =   7335
         TabIndex        =   49
         Top             =   4830
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4155
         Left            =   210
         TabIndex        =   50
         Top             =   495
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   7329
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
      Begin VB.Label Label5 
         Caption         =   "Errores de Comprobación"
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
         Left            =   225
         TabIndex        =   52
         Top             =   180
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   7
         Left            =   450
         TabIndex        =   51
         Top             =   1470
         Width           =   3555
      End
   End
   Begin VB.Frame FrameRenovacion 
      Height          =   4200
      Left            =   5490
      TabIndex        =   37
      Top             =   180
      Width           =   6720
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
         Left            =   2850
         MaxLength       =   4
         TabIndex        =   39
         Top             =   1755
         Width           =   735
      End
      Begin VB.CommandButton CmdCancelar 
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
         Left            =   5205
         TabIndex        =   45
         Top             =   3390
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
         Left            =   4035
         TabIndex        =   43
         Top             =   3390
         Width           =   1065
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
         Index           =   6
         Left            =   2850
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1275
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cálculo utilizando la fecha del Avnic"
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
         Left            =   585
         TabIndex        =   40
         Top             =   2190
         Width           =   4695
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
         Left            =   2835
         MaxLength       =   4
         TabIndex        =   41
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Renovación de AVNICS"
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
         Height          =   345
         Index           =   0
         Left            =   585
         TabIndex        =   47
         Top             =   495
         Width           =   5925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Ejercicio"
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
         TabIndex        =   46
         Top             =   1755
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   585
         TabIndex        =   44
         Top             =   1275
         Width           =   1875
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2535
         Picture         =   "frmAVNInformes.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de Meses"
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
         TabIndex        =   42
         Top             =   2685
         Width           =   1740
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   7815
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   12090
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
         Left            =   225
         TabIndex        =   9
         Top             =   7065
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
         Left            =   225
         TabIndex        =   26
         Top             =   4230
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   1680
            Width           =   4665
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   0
            Left            =   6450
            TabIndex        =   29
            Top             =   1200
            Width           =   255
         End
         Begin VB.CommandButton PushButton2 
            Caption         =   ".."
            Height          =   315
            Index           =   1
            Left            =   6450
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   720
            Width           =   1515
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
         Height          =   6630
         Left            =   7335
         TabIndex        =   23
         Top             =   270
         Width           =   4455
         Begin VB.Frame Frame4 
            Caption         =   "Tipo Avnics"
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
            Height          =   1800
            Left            =   270
            TabIndex        =   25
            Top             =   675
            Width           =   3810
            Begin VB.CheckBox ChkTipoDocu 
               Caption         =   "Antiguos"
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
               TabIndex        =   5
               Top             =   405
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox ChkTipoDocu 
               Caption         =   "Nuevos"
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
               TabIndex        =   6
               Top             =   840
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox ChkTipoDocu 
               Caption         =   "Cancelados"
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
               TabIndex        =   7
               Top             =   1275
               Value           =   1  'Checked
               Width           =   1935
            End
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
         Height          =   3840
         Left            =   225
         TabIndex        =   12
         Top             =   270
         Width           =   6915
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   2
            Top             =   2040
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
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   3
            Top             =   2400
            Width           =   1305
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
            Index           =   4
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   4
            Top             =   3240
            Width           =   690
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
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "Text5"
            Top             =   1320
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
            Index           =   0
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "Text5"
            Top             =   945
            Width           =   4170
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   1
            Top             =   1320
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
            Index           =   0
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   0
            Top             =   945
            Width           =   830
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   885
            Picture         =   "frmAVNInformes.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   2400
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   885
            Picture         =   "frmAVNInformes.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   2040
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
            Left            =   225
            TabIndex        =   22
            Top             =   2400
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
            Left            =   225
            TabIndex        =   21
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Index           =   16
            Left            =   225
            TabIndex        =   20
            Top             =   1710
            Width           =   2490
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Index           =   0
            Left            =   225
            TabIndex        =   19
            Top             =   2955
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Año"
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
            Left            =   315
            TabIndex        =   18
            Top             =   3240
            Width           =   465
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   900
            MouseIcon       =   "frmAVNInformes.frx":01AD
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   915
            MouseIcon       =   "frmAVNInformes.frx":02FF
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar cliente"
            Top             =   945
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
            Index           =   12
            Left            =   225
            TabIndex        =   17
            Top             =   1320
            Width           =   645
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
            Left            =   225
            TabIndex        =   16
            Top             =   945
            Width           =   690
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
            Height          =   375
            Index           =   8
            Left            =   180
            TabIndex        =   13
            Top             =   540
            Width           =   3120
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
         Left            =   10440
         TabIndex        =   10
         Top             =   7110
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
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
         Index           =   1
         Left            =   8955
         TabIndex        =   8
         Top             =   7110
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAVNInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
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


Private WithEvents frmavn As frmAVNAvnics 'Avnics
Attribute frmavn.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


Dim IndCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

Dim cDesde As String
Dim cHasta As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("avnic.fechavto", "F", Me.txtCodigo(2), Me.txtCodigo(2), Me.txtCodigo(3), Me.txtCodigo(3), "pDHfechaVto=""") Then Exit Function
    If Not PonerDesdeHasta("avnic.codavnic", "COD", Me.txtCodigo(0), Me.txtNombre(0), Me.txtCodigo(1), Me.txtNombre(1), "pDHavnics=""") Then Exit Function
    If Not PonerDesdeHasta("avnic.anoejerc", "COD", Me.txtCodigo(4), Me.txtCodigo(4), Me.txtCodigo(4), Me.txtCodigo(4), "pDHejercici=""") Then Exit Function
    
    'Tipo de Avnics
    Codigo = "{avnic.codialta}"
    cDesde = ""
    cHasta = ""
    For i = 0 To Me.ChkTipoDocu.Count - 1
         If Me.ChkTipoDocu(i).Value = 1 Then 'seleccionado
            If cDesde = "" Then
                cDesde = Codigo & "=" & i
                cHasta = ChkTipoDocu(i).Caption
            Else
                cDesde = cDesde & " OR " & Codigo & "=" & i
                cHasta = cHasta & ", " & ChkTipoDocu(i).Caption
            End If
         End If
    Next i
    cDesde = "(" & cDesde & ")"
    AnyadirAFormula cadFormula, cDesde
    AnyadirAFormula cadSelect, Replace(Replace(cDesde, "{", ""), "}", "")
        
    cadParam = cadParam & "pTipoDoc=""" & cHasta & """|"
    numParam = numParam + 1

        
    
    MontaSQL = True
    
End Function

Private Sub Check1_Click()
    If Check1.Value <> 1 Then
        Label1(1).visible = False
        txtCodigo(5).visible = False
    Else
        Label1(1).visible = True
        txtCodigo(5).visible = True
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim cDesde As String
Dim fDesde As String
    
    Select Case OpcionListado
        Case 1 ' renovacion de avnics
    
            If Not DatosOK Then Exit Sub
        
            Sql = ""
            'Valores para Formula seleccion del informe
            cDesde = Trim(txtCodigo(7).Text)
            fDesde = Trim(txtCodigo(6).Text)
            
            Sql = "WHERE avnic.codialta <> 2"
            
            RenovarAvnics Sql
            
        Case 2
        
    End Select

End Sub

Private Function DatosOK() As Boolean
Dim Sql As String
Dim b As Boolean
    b = True
    If Me.Check1.Value = 1 Then
        If txtCodigo(5).Text = "" Then
            MsgBoxA "Debe introducir un valor en el número de meses que se va a incrementar.", vbExclamation
            b = False
        Else
            If CInt(txtCodigo(5).Text) < 1 Or CInt(txtCodigo(5).Text) > 12 Then
                MsgBoxA "El rango de meses que podemos incremetar es de uno a doce meses.", vbExclamation
                b = False
            End If
        End If
    End If
    DatosOK = b
End Function



Private Function RenovarAvnics(cadW As String) As Boolean
'Eliminar Albaranes de Fecha y Turno: Tabla (scaalb)
'que cumplan los criterios seleccionados en la cadena WHERE cadW

Dim Cad As String, Sql As String
Dim Rs As ADODB.Recordset
Dim todasElim As Boolean

    On Error GoTo EEliminar

    Cad = vbCrLf & "Va a renovar los AVNICS."
    Cad = Cad & "   ¿Desea Comenzar? "
    
    If MsgBoxA(Cad, vbQuestion + vbYesNoCancel) = vbYes Then     'Empezamos
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        
        If CrearAvnics(txtCodigo(6).Text, txtCodigo(7).Text) Then
            MsgBoxA "El proceso se realizó correctamente.", vbInformation
            Unload Me
        Else
            MsgBoxA "ATENCIÓN: Se ha producido error en el proceso.", vbExclamation
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Renovar Avnics", Err.Description
End Function

Private Function CrearAvnics(FecVto As String, AnoEje As String) As Boolean
'Eliminar las lineas y la Cabecera de un Caja. Tablas: cajascab, cajaslin
Dim Sql As String
Dim Sql2 As String
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim CADENA As String
Dim Mes As Integer
Dim Ano As Integer
Dim fecha1 As String
Dim Fecha As Date

    On Error GoTo ECrearAvnics
    CrearAvnics = False
    b = False
    
    Conn.BeginTrans
    
    If Me.Check1.Value = False Then
        Sql = "insert into `avnic` (`codavnic`,`nombrper`,`nifperso`,`nifrepre`,"
        Sql = Sql & "`codposta`,`nomcalle`,`poblacio`,`provinci`,`codialta`,`codbanco`,`codsucur`,"
        Sql = Sql & "`cuentaba`,`digcontr`,`imporper`,`imporret`,`anoejerc`,`nifpers1`,`fechalta`,`nombper1`,"
        Sql = Sql & "`nomcall1`,`poblaci1`,`provinc1`,`codpost1`,`fechavto`,`porcinte`,"
        Sql = Sql & "`importes`,`codmacta`,`observac`,`iban`) "
        Sql = Sql & "SELECT codavnic,nombrper,nifperso,nifrepre,"
        Sql = Sql & "codposta,nomcalle,poblacio,provinci,0,codbanco,codsucur,"
        Sql = Sql & "cuentaba,digcontr,0,0," & AnoEje & ",nifpers1,fechalta,nombper1,"
        Sql = Sql & "nomcall1,poblaci1,provinc1,codpost1,'" & Format(FecVto, FormatoFecha) & "',porcinte,"
        Sql = Sql & "importes,codmacta,observac, iban FROM avnic "
        Sql = Sql & "WHERE codialta <> 2 AND anoejerc =" & AnoEje - 1
        'AVNICS
        Conn.Execute Sql
    Else
        Sql = "SELECT codavnic,nombrper,nifperso,nifrepre,"
        Sql = Sql & "codposta,nomcalle,poblacio,provinci,0,codbanco,codsucur,"
        Sql = Sql & "cuentaba,digcontr,0,0," & AnoEje & ",nifpers1,fechalta,nombper1,"
        Sql = Sql & "nomcall1,poblaci1,provinc1,codpost1,fechavto,porcinte,"
        Sql = Sql & "importes,codmacta,observac, iban FROM avnic "
        Sql = Sql & "WHERE codialta <> 2 AND anoejerc =" & AnoEje - 1
    
        
        Sql2 = "insert into `avnic` (`codavnic`,`nombrper`,`nifperso`,`nifrepre`,"
        Sql2 = Sql2 & "`codposta`,`nomcalle`,`poblacio`,`provinci`,`codialta`,`codbanco`,`codsucur`,"
        Sql2 = Sql2 & "`cuentaba`,`digcontr`,`imporper`,`imporret`,`anoejerc`,`nifpers1`,`fechalta`,`nombper1`,"
        Sql2 = Sql2 & "`nomcall1`,`poblaci1`,`provinc1`,`codpost1`,`fechavto`,`porcinte`,"
        Sql2 = Sql2 & "`importes`,`codmacta`,`observac`,`iban`) values "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            fecha1 = CStr(DateAdd("m", CInt(txtCodigo(5).Text), Rs!fechavto)) ' le suma el numero de meses a la fecha de vto
            
            CADENA = "(" & DBSet(Rs!codavnic, "N") & "," & DBSet(Rs!nombrper, "T") & "," & DBSet(Rs!nifperso, "T") & "," & DBSet(Rs!nifrepre, "T") & ","
            CADENA = CADENA & DBSet(Rs!codposta, "T") & "," & DBSet(Rs!nomcalle, "T") & "," & DBSet(Rs!poblacio, "T") & "," & DBSet(Rs!provinci, "T") & ","
            CADENA = CADENA & "0" & "," & DBSet(Rs!codbanco, "N") & "," & DBSet(Rs!codsucur, "N") & ","
            CADENA = CADENA & DBSet(Rs!cuentaba, "T") & "," & DBSet(Rs!digcontr, "T") & ",0,0," & AnoEje & "," & DBSet(Rs!nifpers1, "T") & ","
            CADENA = CADENA & DBSet(Rs!fechalta, "F") & "," & DBSet(Rs!nombper1, "T") & ","
            CADENA = CADENA & DBSet(Rs!nomcall1, "T") & "," & DBSet(Rs!poblaci1, "T") & "," & DBSet(Rs!provinc1, "T") & "," & DBSet(Rs!codpost1, "T") & ","
            CADENA = CADENA & DBSet(fecha1, "F") & "," & DBSet(Rs!Porcinte, "N") & ","
            CADENA = CADENA & DBSet(Rs!Importes, "N") & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(Rs!observac, "T") & "," & DBSet(Rs!IBAN, "T") & ")"
        
            Sql2 = Sql2 & CADENA & ","
        
            Rs.MoveNext
        Wend
    
        ' quitamos la ultima coma del ultimo registro
        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
        
        Conn.Execute Sql2
        Set Rs = Nothing
    
    End If
    
    CrearAvnics = True
    b = True
    
ECrearAvnics:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description, "Insertar registros"
        b = False
    End If
    
    If Not b Then
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
    End If
    CrearAvnics = b
End Function




Private Sub cmdAccion_Click(Index As Integer)

    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("avnic", cadSelect) Then Exit Sub
    
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

End Sub

Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    vMostrarTree = False
    conSubRPT = False
        
'    cadParam = cadParam & "pTitulo=""" & txtTitulo(0).Text & """|"
'    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
'    cadParam = cadParam & "pTotalAsiento=" & chkTotalAsiento.Value & "|"
'
'    numParam = numParam + 3
    
    
    
    indRPT = "1416-00" '"rInfAvnics.rpt"

    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Sub AccionesCSV()
Dim Sql As String

    'Monto el SQL
    Sql = "Select codavnic AS Código,fechalta as FAlta,fechavto as FVto,nombrper as ApeNombre,nifperso as NIF,nombper1 as ApeNombre1,nifpers1 as NIF1,concat(iban,lpad(codbanco,4,'0'),lpad(codsucur,4,'0'),lpad(digcontr,2,'0'),lpad(cuentaba,10,'0')) as DatosBanco,importes as Importe,imporper as Percepción,imporret as Retención"
    Sql = Sql & " From avnic "
    
    If cadSelect <> "" Then Sql = Sql & " WHERE " & cadSelect
    
    Sql = Sql & " ORDER BY 1 "
        
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 0
                PonFoco txtCodigo(0)
            Case 1 ' renovacion
                PonFoco txtCodigo(6)
        
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
     Me.imgBuscar(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
     
     
    FrameCobros.visible = False
    FrameRenovacion.visible = False
    Me.FrameErrores.visible = False
    
    '###Descomentar
'    CommitConexion
    Select Case OpcionListado
        Case 0 ' listado de avnics
            FrameCobrosVisible True, H, W
            indFrame = 5
            tabla = "avnic"
                
            PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
            ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
            
        Case 1 ' renovacion de avnics
            tabla = "avnic"
            
            Label1(1).visible = False
            txtCodigo(5).visible = False
            
            FrameRenovacionVisible True, H, W
        
        Case 2 ' visualizacion de errores de la contabilizacion
            
            PonerFrameErroresVisible True, 1000, 2000
            CargarListaErrComprobacion
            Me.Caption = "Errores de Comprobacion: "
            PonerFocoBtn Me.CmdSalir
                
        
        
    End Select
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(IndCodigo).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmavn_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(IndCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object

    IndCodigo = Index
    Select Case Index
        Case 0
            IndCodigo = 6
    End Select
    
    'FECHA
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtCodigo(IndCodigo).Text <> "" Then frmC.Fecha = CDate(txtCodigo(Index).Text)
    frmC.Show vbModal
    Set frmC = Nothing
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'AVNICS
            AbrirFrmAvnics (Index)
        
    End Select
    PonFoco txtCodigo(IndCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
'        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
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
            Case 0: KEYBusqueda KeyAscii, 0 'avnics desde
            Case 1: KEYBusqueda KeyAscii, 1 'avnics hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            
            Case 6: KEYFecha KeyAscii, 0 'renovacion de avnics
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'AVNICS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "avnic", "nombrper", "codavnic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4 'EJERCICIO
              txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
              
        Case 7 'EJERCICIO
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)

        Case 6 'FECHA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
              
              
  End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.top = -90
        Me.FrameCobros.Left = 0
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub FrameRenovacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameRenovacion.visible = visible
    If visible = True Then
        Me.FrameRenovacion.top = -90
        Me.FrameRenovacion.Left = 0
        W = Me.FrameRenovacion.Width
        H = Me.FrameRenovacion.Height
    End If
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
 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub


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
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub

Private Sub CargarListaErrComprobacion()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarListErrComprobacion

    Sql = " SELECT  * "
    Sql = Sql & " FROM tmperrcomprob "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'Los encabezados
        ListView2.ColumnHeaders.Clear

        ListView2.ColumnHeaders.Add , , "Error en cuentas contables", 6000
        
    
        While Not Rs.EOF
            Set ItmX = ListView2.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarListErrComprobacion:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub



Private Sub PonerFrameErroresVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameErrores.visible = visible
    If visible = True Then
        Me.FrameErrores.top = -90
        Me.FrameErrores.Left = 0
        W = Me.FrameErrores.Width
        H = Me.FrameErrores.Height
    End If
    
End Sub


