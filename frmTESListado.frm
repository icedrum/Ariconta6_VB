VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESListado 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   Icon            =   "frmTESListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCompensaciones 
      Height          =   6045
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   8235
      Begin VB.CheckBox chkCompensa 
         Caption         =   "Dejar sólo importe compensacion"
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
         Left            =   960
         TabIndex        =   13
         Top             =   5370
         Width           =   4005
      End
      Begin VB.Frame FrameCambioFPCompensa 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   7785
         Begin VB.TextBox txtDescFPago 
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
            Left            =   3000
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   240
            Width           =   4695
         End
         Begin VB.TextBox txtFPago 
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
            Left            =   2220
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago vto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   49
            Left            =   90
            TabIndex        =   29
            Top             =   240
            Width           =   1590
         End
         Begin VB.Image imgFP 
            Height          =   240
            Index           =   8
            Left            =   1920
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.ComboBox cboCompensaVto 
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
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   4245
      End
      Begin VB.TextBox txtConcpto 
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
         Left            =   2340
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   4440
         Width           =   645
      End
      Begin VB.TextBox txtDescConcepto 
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
         Left            =   3030
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4440
         Width           =   4785
      End
      Begin VB.CommandButton cmdContabCompensaciones 
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
         Left            =   5700
         TabIndex        =   14
         Top             =   5370
         Width           =   975
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
         Index           =   22
         Left            =   6780
         TabIndex        =   15
         Top             =   5370
         Width           =   975
      End
      Begin VB.TextBox txtDescConcepto 
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
         Left            =   3030
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3960
         Width           =   4785
      End
      Begin VB.TextBox txtConcpto 
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
         Left            =   2340
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3960
         Width           =   645
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
         Left            =   3030
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3240
         Width           =   4785
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
         Index           =   0
         Left            =   2340
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3240
         Width           =   645
      End
      Begin VB.TextBox txtCtaBanc 
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
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2040
         Width           =   1275
      End
      Begin VB.TextBox txtDescBanc 
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
         Left            =   3660
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2040
         Width           =   4125
      End
      Begin VB.TextBox Text3 
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
         Left            =   2370
         TabIndex        =   6
         Top             =   840
         Width           =   1485
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   0
         Left            =   480
         Top             =   5370
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Compensa sobre Vto."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   47
         Left            =   210
         TabIndex        =   26
         Top             =   1440
         Width           =   2160
      End
      Begin VB.Label Label6 
         Caption         =   "Pagos"
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
         Left            =   960
         TabIndex        =   25
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "Cobros"
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
         Index           =   20
         Left            =   960
         TabIndex        =   24
         Top             =   3960
         Width           =   765
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   1
         Left            =   2040
         Picture         =   "frmTESListado.frx":000C
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   46
         Left            =   210
         TabIndex        =   22
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmTESListado.frx":685E
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmTESListado.frx":D0B0
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   45
         Left            =   210
         TabIndex        =   20
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   44
         Left            =   210
         TabIndex        =   18
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   2
         Left            =   2040
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   23
         Left            =   2040
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha contab."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   43
         Left            =   210
         TabIndex        =   16
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Contabilización compensaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   12
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   5370
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameProgreso 
      Height          =   1935
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label lbl2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblPPAL 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameRecepcionDocumentos 
      Height          =   4815
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   8115
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
         Left            =   3180
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   960
         Width           =   4425
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
         Index           =   1
         Left            =   2160
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   960
         Width           =   885
      End
      Begin VB.TextBox txtNConcepto 
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
         Left            =   3180
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   1680
         Width           =   4425
      End
      Begin VB.TextBox txtConcepto 
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
         Left            =   2160
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   1680
         Width           =   885
      End
      Begin VB.TextBox txtConcepto 
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
         Left            =   2160
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   2160
         Width           =   885
      End
      Begin VB.TextBox txtNConcepto 
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
         Left            =   3180
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   2160
         Width           =   4425
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
         Index           =   23
         Left            =   6840
         TabIndex        =   37
         Top             =   2820
         Width           =   975
      End
      Begin VB.CommandButton cmdRecepDocu 
         Caption         =   "Aceptar"
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
         Left            =   5640
         TabIndex        =   36
         Top             =   4320
         Width           =   975
      End
      Begin VB.CheckBox chkAgruparCtaPuente 
         Caption         =   "Agrupa apuntes cta puente"
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
         Left            =   600
         TabIndex        =   35
         Top             =   2640
         Width           =   4365
      End
      Begin VB.TextBox txtCta 
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
         Left            =   2190
         TabIndex        =   34
         Top             =   3480
         Width           =   1365
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
         Index           =   14
         Left            =   3690
         TabIndex        =   33
         Text            =   "Text5"
         Top             =   3480
         Width           =   3885
      End
      Begin VB.TextBox txtDescCCoste 
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
         Left            =   3450
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   3900
         Width           =   4125
      End
      Begin VB.TextBox txtCCost 
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
         Left            =   2190
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3900
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   13
         Left            =   480
         TabIndex        =   51
         Top             =   240
         Width           =   6990
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   1
         Left            =   1890
         Picture         =   "frmTESListado.frx":13902
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   50
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Debe"
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
         Index           =   22
         Left            =   600
         TabIndex        =   49
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   51
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   2
         Left            =   1890
         Picture         =   "frmTESListado.frx":1A154
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Haber"
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
         Index           =   23
         Left            =   600
         TabIndex        =   47
         Top             =   2160
         Width           =   795
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   3
         Left            =   1890
         Picture         =   "frmTESListado.frx":209A6
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   55
         Left            =   120
         TabIndex        =   46
         Top             =   3120
         Width           =   1440
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   14
         Left            =   1890
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Index           =   28
         Left            =   480
         TabIndex        =   45
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Centro Coste"
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
         Index           =   29
         Left            =   480
         TabIndex        =   44
         Top             =   3900
         Width           =   1365
      End
      Begin VB.Image imgCCoste 
         Height          =   240
         Index           =   0
         Left            =   1890
         Top             =   3930
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTESListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Public Opcion As Byte
    '1.- Cobros pendientes por cliente
    
    '3.- Reclamaciones por mail
    
    '4.- lISTADO agentes
    '5.- Departamentos
    
    '6.- Listado remesas
    
    '8.- Listado caja
    
    '9-  Devol remesas
    
    '10.- Listado formas de pago

    
    '11.- Transferencias PRovee   (o confirmings (domicilados o caixaconfirming)
    
    '12.- Listado previsional de gstos/pagos
    
    '13.- Transferencias ABONOS
    
    
    'Operaciones aseguradas
    '----------------------------
    '15.-  datos basicos
    '16.-  listado facturacion
    '17.-  Impagados asegurados
    
    
    '20.- Pregunta cuenta COBRO GENERICO
    '       La pongo aqui pq tengo implemntado todo
    
    
    '22.- Datos para la contabilizacion de las compensaciones
        
    '23.- Datos para la contbailiacion de la recpcion de documentos
    
    
    '24.-  Listado de documento(tal/pag) recibidos
    
    '25.-  Listado de pagos ordenados por banco  **** AHORA NO DEBERIA ENTRAR AQUI
    
    '26.-  Cancel remesa TAL/PAG.  Cando los importe no coinden. Solicitud cta y cc
    '27.-  Divide el vencimiento en dos vtos a partir del importe introducido en el text
        
        
    '30.-  Historico RECLAMACIONES
    '31.-   Gastos fijos
        
    '33.-  ASEGURADOS.  Listados avisos falta pago, avisos prorroga, aviso siniestro
        
    '34.-  Eliminar una recepcion de documentos, que ya ha sido contb con la puente
        
    '35.-  Gastos transferencias
        
    '36.-  Compensar ABONOS cobros
            
    '38.-  Recaudacion ejecutiva
        
    '39.-   Informe de comunicacion al seguro
    '40.-    Fras pendientes operaciones aseguradas
    
    '42.-   IMportar fichero norma 57 (recibos al cobro en ventanilla)
    
    '43.-   Confirmings
    '44.-   Caixaconfirming   igual que el de arriba
    
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
'--monica
'Private WithEvents frmB As frmBuscaGrid
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmS As frmBasico
Attribute frmS.VB_VarHelpID = -1
Private WithEvents frmCCos As frmBasico
Attribute frmCCos.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1

Dim Sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim i As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String
Dim IndCodigo As Integer


Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function

Private Sub cboCobro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboCompensaVto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAgruparCtaPuente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkCompensa_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 20 Or Index = 23 Or Index >= 26 Then
        CadenaDesdeOtroForm = "" 'Por si acaso. Tiene que devolve "" para que no haga nada
    End If
    Unload Me
End Sub


Private Sub cmdContabCompensaciones_Click()

    'COmprobaciones y leches
    If Me.txtConcpto(0).Text = "" Or txtDiario(0).Text = "" Or Text3(23).Text = "" Or _
        Me.txtConcpto(1).Text = "" Then
        MsgBox "Todos los campos de contabilizacion  son obligatorios", vbExclamation
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        If Me.txtCtaBanc(2).Text = "" Then
            MsgBox "Campo banco no puede estar vacio", vbExclamation
            Exit Sub
        End If
    Else
        If Me.txtFPago(8).Text <> "" Then
            RC = DevuelveDesdeBD("codforpa", "formapago", "codforpa", txtFPago(8).Text, "N")
            If RC = "" Then
                MsgBox "No existe la forma de pago", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If FechaCorrecta2(CDate(Text3(23).Text), True) > 1 Then
        PonFoco Text3(23)
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        'No compensa sobre ningun vencimiento.
        'No puede marcar la opcion del importe
        If chkCompensa.Value = 1 Then
            MsgBox "'Dejar sólo importe compensación' disponible cuando compense sobre un vencimiento", vbExclamation
            Exit Sub
        End If
    End If

    'Cargamos la cadena y cerramos
    CadenaDesdeOtroForm = Me.txtConcpto(0).Text & "|" & Me.txtConcpto(1).Text & "|" & txtDiario(0).Text & "|" & Text3(23).Text & "|" & Me.txtCtaBanc(2).Text & "|" & DevNombreSQL(txtDescBanc(2).Text) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.txtFPago(8).Text & "|" & Me.cboCompensaVto.ItemData(Me.cboCompensaVto.ListIndex) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.chkCompensa.Value & "|"
    Unload Me
End Sub








Private Sub cmdRecepDocu_Click()
    If txtDiario(1).Text = "" Or Me.txtConcepto(2).Text = "" Or txtConcepto(3).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    If Me.Label4(55).Visible Then
        If Me.txtCta(14).Text = "" Then
            MsgBox "Cuentas " & Label4(55).Caption & " requerida", vbExclamation
            Exit Sub
        End If
        Sql = ""
        If vParam.autocoste Then
            RC = Mid(txtCta(14).Text, 1, 1)
            If RC = 6 Or RC = 7 Then
                If txtCCost(0).Text = "" Then
                    MsgBox "Centro de coste requerido", vbExclamation
                    Exit Sub
                Else
                    Sql = txtCCost(0).Text
                End If
            End If
        End If
        txtCCost(0).Text = Sql
    Else
        txtCCost(0).Text = ""
        Me.txtCta(14).Text = ""
    End If
    
    
    
    
    i = 0
    If Me.chkAgruparCtaPuente(0).Visible Then
        If Me.chkAgruparCtaPuente(0).Value Then i = 1
    End If
    CadenaDesdeOtroForm = txtDiario(1).Text & "|" & Me.txtConcepto(2).Text & "|" & txtConcepto(3).Text & "|" & i & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtCta(14).Text & "|" & txtCCost(0).Text & "|"
    
    Unload Me

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion

        Case 22
            'Contabi efectos
            If CONT > 0 Then
                For i = 1 To Me.cboCompensaVto.ListCount
                    If Me.cboCompensaVto.ItemData(i) = CONT Then
                        CONT = i
                        Exit For
                    End If
                Next
            End If
            Me.cboCompensaVto.ListIndex = CONT
            PonFoco Text3(23)
        
        Case 23
            CadenaDesdeOtroForm = ""  'Para que  no devuelva nada
            
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmppal.Icon
    CargaImagenesAyudas Me.imgCtaBanc, 1, "Cuenta contable bancaria"
    CargaImagenesAyudas Image2, 2
    CargaImagenesAyudas Me.imgFP, 1, "Forma de pago"
    CargaImagenesAyudas Me.ImageAyudaImpcta, 3
    For Each Img In Me.imgConcepto
        Img.ToolTipText = "Concepto"
    Next
    For Each Img In Me.imgDiario
        Img.ToolTipText = "Diario"
    Next
    
    
    
    'Limpiamos el tag
    PrimeraVez = True
    FrameCompensaciones.Visible = False
    FrameRecepcionDocumentos.Visible = False
    
    CommitConexion
    
    Select Case Opcion
    Case 22
        
        
        For H = 0 To 1
            
            txtConcpto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 1)
            txtDescConcepto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 2)
        Next H
        Me.cboCompensaVto.Clear
        InsertaItemComboCompensaVto "No compensa sobre ningún vencimiento", 0
        
        'Veremos si puede sobre un Vto o no
        H = RecuperaValor(CadenaDesdeOtroForm, 5)
        CONT = 0
        If H = 1 Then CONT = RecuperaValor(CadenaDesdeOtroForm, 6)
        FrameCambioFPCompensa.Visible = CONT > 0
        CadenaDesdeOtroForm = ""
        H = FrameCompensaciones.Height + 120
        W = FrameCompensaciones.Width
        FrameCompensaciones.Visible = True
        Caption = "Compensacion efectos"
        Text3(23).Text = Format(Now, "dd/mm/yyyy")
        
        
        
        
    Case 23, 34
        '23.-  Contabilizar
        '34. Eliminar ya contabilizada
        
        
        
        
        'Tendremos el tipo de pago , talon o pagare
        Dim FP As Ctipoformapago
        Set FP = New Ctipoformapago
        
        If Opcion = 23 Then
            Label2(13).Caption = "Contabilizar recepción documentos"
            Caption = "Contabilizar"
        Else
            Label2(13).Caption = "Eliminar de recepción documentos"
            Caption = "Eliminar"
        End If
        
        'Cuenta beneficios gastos paras las diferencias si existieran
        'Si el total del talon es el total de las lineas entonces no mostrara los
        'datos del total. 0: igual   1  Mayor     2 Menor
        Sql = RecuperaValor(CadenaDesdeOtroForm, 2)
        i = CInt(Sql)
'        If CInt(SQL) > 0 Then
'            I = 1
'        Else
'            I = -1
'        End If
        
        Label4(55).Visible = i <> 0
        Image3(14).Visible = i <> 0
        txtCta(14).Visible = i <> 0
        DtxtCta(14).Visible = i <> 0
        Label6(28).Visible = i <> 0
        
        
        
        
        
        If i > 0 Then
            Sql = "Beneficios"
        Else
            Sql = "Pérdidas"
        End If
        
        If Opcion = 34 Then Sql = Sql & "(Deshacer apunte)"
        Label4(55).Caption = Sql

        


        '   No lleva ANALITICA
        If i <> 0 Then
            If Not vParam.autocoste Then i = 0
        End If
     
        Me.imgCCoste(0).Visible = i <> 0
        Me.txtCCost(0).Visible = i <> 0
        Label6(29).Visible = i <> 0
        Me.txtDescCCoste(0).Visible = i <> 0
     
        
        
        
        
        
        
        
        
        Sql = RecuperaValor(CadenaDesdeOtroForm, 1)
        i = CInt(Sql)
        If FP.Leer(i) = 0 Then
            If Opcion = 23 Then
                'Normal
                txtDiario(1).Text = FP.diaricli
                txtConcepto(2).Text = FP.condecli
                txtConcepto(3).Text = FP.conhacli
             Else
                'Eliminar. Iran cambiados
                txtDiario(1).Text = FP.diaricli
                txtConcepto(2).Text = FP.conhacli
                txtConcepto(3).Text = FP.condecli
                
                
             End If
                
            'Para que pinte la descripcion
            txtDiario_LostFocus 1
            txtConcepto_LostFocus 2
            txtConcepto_LostFocus 3
        End If
        
        
        
        
        H = 0
        If i = vbTalon Then
            Sql = "taloncta"
        Else
            Sql = "pagarecta"
        End If
        
        Sql = DevuelveDesdeBD(Sql, "paramtesor", "codigo", "1")
        If Len(Sql) = vEmpresa.DigitosUltimoNivel Then
            chkAgruparCtaPuente(0).Visible = True
            H = 1 '
        
            'Si esta configurado en parametrps, si la ultima vez lo marco seguira marcado
            If H = 1 Then H = CheckValueLeer("Agrup0")
            If H <> 1 Then H = 0
            chkAgruparCtaPuente(0).Value = H
            
        Else
            chkAgruparCtaPuente(0).Visible = False
        End If
        
        Set FP = Nothing
        
        If Label4(55).Visible Then '5055
            FrameRecepcionDocumentos.Height = 4815
            i = 4320
        Else
            FrameRecepcionDocumentos.Height = 3135
            i = 2640
        End If
        cmdRecepDocu.top = i
        cmdCancelar(23).top = i
        H = FrameRecepcionDocumentos.Height + 120
        W = FrameRecepcionDocumentos.Width
        FrameRecepcionDocumentos.Visible = True
        
        
            
        
        
        
        
        
    End Select
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    i = Opcion
    If Opcion = 13 Or i = 43 Or i = 44 Then i = 11
    
    'Aseguradas
    If Opcion >= 15 And Opcion <= 18 Then i = 15  'aseguradoas
    If Opcion = 33 Then i = 15 'aseguradoas
    If Opcion = 34 Then i = 23 'Eliminar recepcion documento
    If Opcion = 40 Then i = 39
    Me.cmdCancelar(i).Cancel = True
    
    PonerFrameProgreso

End Sub

Private Sub PonerFrameProgreso()
Dim i As Integer

    'Ponemos el frame al pricnipio de todo
    FrameProgreso.Visible = False
    FrameProgreso.ZOrder 0
    
    'lo ubicamos
    'Posicion horizintal WIDTH
    i = Me.Width - FrameProgreso.Width
    If i > 100 Then
        i = i \ 2
    Else
        i = 0
    End If
    FrameProgreso.Left = i
    'Posicion  VERTICAL HEIGHT
    i = Me.Height - FrameProgreso.Height
    If i > 100 Then
        i = i \ 2
    Else
        i = 0
    End If
    FrameProgreso.top = i
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtFPago(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescFPago(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub


Private Sub Image3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    RC = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub

Private Sub ImageAyudaImpcta_Click(Index As Integer)
Dim C As String
    Select Case Index
    Case 0
            C = "Compensaciones" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Cuando compense sobre un vencimiento al marcar la opción " & vbCrLf
            C = C & Space(10) & Me.chkCompensa.Caption & vbCrLf
            C = C & "se modificará el importe vencimiento poniendo el total a compensar  y en importe cobrado un cero"
            
    End Select
    MsgBox C, vbInformation

End Sub

Private Sub ImgCCoste_Click(Index As Integer)
    
    IndCodigo = Index
    
    Set frmCCos = New frmBasico
    
    AyudaCC frmCCos
    
    Set frmCCos = Nothing
    
    PonFoco Me.txtCCost(Index)
    AbiertoOtroFormEnListado = False
End Sub

Private Sub imgConcepto_Click(Index As Integer)
    Sql = ""
    AbiertoOtroFormEnListado = True
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = True
    frmCon.Show vbModal
    Set frmCon = Nothing
    If Sql <> "" Then
        Me.txtConcepto(Index).Text = RecuperaValor(Sql, 1)
        Me.txtNConcepto(Index).Text = RecuperaValor(Sql, 2)
    Else
        QuitarPulsacionMas Me.txtConcepto(Index)
    End If
    
    PonFoco Me.txtConcepto(Index)
    AbiertoOtroFormEnListado = False
End Sub

Private Sub imgCtaBanc_Click(Index As Integer)
    Sql = ""
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
    If Sql <> "" Then
        txtCtaBanc(Index).Text = RecuperaValor(Sql, 1)
        Me.txtDescBanc(Index).Text = RecuperaValor(Sql, 2)
    End If
End Sub

Private Sub imgDiario_Click(Index As Integer)
    LanzaBuscaGrid Index, 0
End Sub



Private Sub imgFP_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    'Set frmCta = New frmColCtas
    Set frmP = New frmFormaPago
    RC = Index
    frmP.DatosADevolverBusqueda = "0|1"
    frmP.Show vbModal
    Set frmP = Nothing
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub


Private Sub txtCCost_GotFocus(Index As Integer)
    PonFoco txtConcpto(Index)
End Sub

Private Sub txtCCost_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCCost_LostFocus(Index As Integer)
    Sql = ""
    txtCCost(Index).Text = Trim(txtCCost(Index).Text)
    If txtCCost(Index).Text <> "" Then
        

            txtConcpto(Index).Text = Val(txtConcpto(Index).Text)
            Sql = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtCCost(Index).Text, "T")
            If Sql = "" Then
                MsgBox "No existe el centro de coste: " & Me.txtCCost(Index).Text, vbExclamation
                Me.txtCCost(Index).Text = ""
            End If
        If txtCCost(Index).Text = "" Then SubSetFocus txtCCost(Index)
    End If
    Me.txtDescCCoste(Index).Text = Sql
End Sub

Private Sub txtConcepto_GotFocus(Index As Integer)
    ConseguirFoco txtConcepto(Index), 3
End Sub

Private Sub txtConcepto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtConcepto(Index).Tag, Index
    End If
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgConcepto"
        imgConcepto_Click Indice
    End Select
    
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 2, 3 'CONCEPTOS
            txtNConcepto(Index).Text = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index), "N")
            If txtConcepto(Index).Text <> "" Then txtConcepto(Index).Text = Format(txtConcepto(Index).Text, "000")
    End Select

End Sub

Private Sub txtDiario_GotFocus(Index As Integer)
    PonFoco txtDiario(Index)
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
    
    Sql = ""
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    If txtDiario(Index).Text <> "" Then
        
        If Not IsNumeric(txtDiario(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtDiario(Index).Text = ""
            SubSetFocus txtDiario(Index)
        Else
            txtDiario(Index).Text = Val(txtDiario(Index).Text)
            Sql = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text, "N")
            
            If Sql = "" Then
                MsgBox "No existe el diario: " & Me.txtDiario(Index).Text, vbExclamation
                Me.txtDiario(Index).Text = ""
                PonFoco txtDiario(Index)
            End If
        End If
    End If
    Me.txtDescDiario(Index).Text = Sql
     
End Sub




Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function





Private Sub txtCtaBanc_GotFocus(Index As Integer)
    PonFoco txtCtaBanc(Index)
End Sub

Private Sub txtCtaBanc_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaBanc_LostFocus(Index As Integer)
    txtCtaBanc(Index).Text = Trim(txtCtaBanc(Index).Text)
    If txtCtaBanc(Index).Text = "" Then
        txtDescBanc(Index).Text = ""
        Exit Sub
    End If
    
    Cad = txtCtaBanc(Index).Text
    i = CuentaCorrectaUltimoNivelSIN(Cad, Sql)
    If i = 0 Then
        MsgBox "NO existe la cuenta: " & txtCtaBanc(Index).Text, vbExclamation
        Sql = ""
        Cad = ""
    Else
        Cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Cad, "T")
        If Cad = "" Then
            MsgBox "Cuenta no asoaciada a ningun banco", vbExclamation
            Sql = ""
            i = 0
        End If
    End If
    
    txtCtaBanc(Index).Text = Cad
    Me.txtDescBanc(Index).Text = Sql
    If i = 0 Then PonFoco txtCtaBanc(Index)
    
End Sub

Private Sub txtFPago_GotFocus(Index As Integer)
    PonFoco txtFPago(Index)
End Sub

Private Sub txtFPago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtFPago_LostFocus(Index As Integer)
    If ComprobarCampoENlazado(txtFPago(Index), txtDescFPago(Index), "N") > 0 Then
        If txtFPago(Index).Text <> "" Then
            'Tiene valor.
            Sql = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", txtFPago(Index).Text, "N")
            If Sql = "" Then Sql = "Codigo no encontrado"
            txtDescFPago(Index).Text = Sql
        Else
            'Era un error
            SubSetFocus txtFPago(Index)
        End If
    End If
End Sub




Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


'Si tiene valor el campo fecha, entonces lo ponemos con el BD
Private Function CampoABD(ByRef T As TextBox, Tipo As String, CampoEnLaBD, Mayor_o_Igual As Boolean) As String

    CampoABD = ""
    If T.Text <> "" Then
        If Mayor_o_Igual Then
            CampoABD = " >= "
        Else
            CampoABD = " <= "
        End If
        Select Case Tipo
        Case "F"
            CampoABD = CampoEnLaBD & CampoABD & "'" & Format(T.Text, FormatoFecha) & "'"
        Case "T"
            CampoABD = CampoEnLaBD & CampoABD & "'" & T.Text & "'"
        Case "N"
            CampoABD = CampoEnLaBD & CampoABD & T.Text
        End Select
    End If
End Function



Private Function CampoBD_A_SQL(ByRef C As ADODB.Field, Tipo As String, Nulo As Boolean) As String

    If IsNull(C) Then
        If Nulo Then
            CampoBD_A_SQL = "NULL"
        Else
            If Tipo = "T" Then
                CampoBD_A_SQL = "''"
            Else
                CampoBD_A_SQL = "0"
            End If
        End If

    Else
    
        Select Case Tipo
        Case "F"
            CampoBD_A_SQL = "'" & Format(C.Value, FormatoFecha) & "'"
        Case "T"
            CampoBD_A_SQL = "'" & DevNombreSQL(C.Value) & "'"
        Case "N"
            CampoBD_A_SQL = TransformaComasPuntos(CStr(C.Value))
        End Select
    End If
End Function

Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.lblPPAL.Caption = TEXTO
        Me.lbl2.Caption = ""
        Me.ProgressBar1.Value = 0
        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub


'Para conceptos y diarios
'Opcion: 0- Diario
'        1- Conceptos
'        2- Centros de coste
'        3- Gastos fijos
'        4. Hco compensaciones
Private Sub LanzaBuscaGrid(Indice As Integer, OpcionGrid As Byte)


End Sub

                                       '                Para saber el index del listview
Public Sub InsertaItemComboCompensaVto(TEXTO As String, Indice As Integer)
    Me.cboCompensaVto.AddItem TEXTO
    Me.cboCompensaVto.ItemData(Me.cboCompensaVto.NewIndex) = Indice
End Sub




Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
'    If Index = 6 Then
'        'NO se ha cambiado nada de la cuenta
'        If txtCta(6).Text = txtCta(6).Tag Then
'
'            Exit Sub
'        Else
'            txtDpto(0).Text = ""
'            txtDpto(1).Text = ""
'            txtDescDpto(0).Text = ""
'            txtDescDpto(0).Text = ""
'        End If
'    End If
     
     
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
       ' txtCta(6).Tag = txtCta(6).Text
        Exit Sub
    End If
    
    If Index = 6 Then
        If txtCta(0).Text <> "" Or txtCta(1).Text <> "" Then
            MsgBox "Si selecciona desde / hasta cliente no podra seleccionar departamento", vbExclamation
            txtCta(6).Text = ""
            txtCta(6).Tag = txtCta(6).Text
            Exit Sub
        End If
        
    Else
        If Index = 0 Or Index = 1 Then
            If txtCta(6).Text <> "" Then
                MsgBox "Si seleciona departamento no puede seleccionar desde / hasta  cliente", vbExclamation
                txtCta(Index).Text = ""
                txtCta(6).Tag = txtCta(6).Text
                Exit Sub
            End If
        End If
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        txtCta(6).Tag = txtCta(6).Text
        PonFoco txtCta(Index)
        
        
        Exit Sub
    End If
    
    Select Case Index
    Case 0 To 7, 11, 12, 15, 16, 18, 19
        'NO hace falta que sean de ultimo nivel
        Cta = (txtCta(Index).Text)
                                '********
        B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
        If B = 0 Then
            MsgBox "NO existe la cuenta: " & txtCta(Index).Text, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
        Else
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = Sql
            If B = 1 Then
                DtxtCta(Index).Tag = ""
            Else
                DtxtCta(Index).Tag = Sql
            End If
            
            
            'Index=1. Cliente en listado de cobros. Si pongo el desde pongo el hasta lo mismo
            If Index = 1 Then
                
                If Len(Cta) = vEmpresa.DigitosUltimoNivel Then
                    txtCta(0).Text = Cta
                    DtxtCta(0).Text = DtxtCta(1).Text
                End If
            End If
            
        End If
    Case Else
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, Sql) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = Sql
            
            
        Else
            MsgBox Sql, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If
        
    End Select
    txtCta(Index).Tag = txtCta(Index).Text
End Sub


