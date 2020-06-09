VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTesCompensaAutomatica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compensacion automática"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   16140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   11040
      TabIndex        =   38
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompensar 
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
      Left            =   9720
      TabIndex        =   37
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame FrameTomaDatos 
      Height          =   4695
      Left            =   360
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   11415
      Begin VB.TextBox txConta 
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
         Left            =   5640
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   3960
         Width           =   885
      End
      Begin VB.TextBox txtDescConta 
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
         Left            =   6600
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   3960
         Width           =   4425
      End
      Begin VB.TextBox txConta 
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
         Left            =   5640
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   3480
         Width           =   885
      End
      Begin VB.TextBox txtDescConta 
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
         Left            =   6600
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   3480
         Width           =   4425
      End
      Begin VB.TextBox txConta 
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
         Left            =   5640
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   3000
         Width           =   885
      End
      Begin VB.TextBox txtDescConta 
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
         Left            =   6600
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   3000
         Width           =   4425
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1680
         TabIndex        =   12
         Top             =   3000
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   9720
         TabIndex        =   11
         Top             =   1920
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   8280
         TabIndex        =   10
         Top             =   1920
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   5640
         TabIndex        =   9
         Top             =   1920
         Width           =   1365
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
         Index           =   3
         Left            =   2160
         TabIndex        =   7
         Top             =   1920
         Width           =   585
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   4200
         TabIndex        =   8
         Top             =   1920
         Width           =   1365
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
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   1920
         Width           =   1185
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   9720
         TabIndex        =   5
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   8280
         TabIndex        =   4
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   5640
         TabIndex        =   3
         Top             =   840
         Width           =   1365
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
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   585
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   4200
         TabIndex        =   2
         Top             =   840
         Width           =   1365
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
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label lblConta 
         AutoSize        =   -1  'True
         Caption         =   "Concepto haber"
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
         Index           =   2
         Left            =   3840
         TabIndex        =   48
         Top             =   4020
         Width           =   1590
      End
      Begin VB.Image imgConta 
         Height          =   240
         Index           =   2
         Left            =   5400
         Top             =   4020
         Width           =   240
      End
      Begin VB.Image imgConta 
         Height          =   240
         Index           =   1
         Left            =   5400
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label lblConta 
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
         Index           =   0
         Left            =   3840
         TabIndex        =   44
         Top             =   3060
         Width           =   675
      End
      Begin VB.Image imgConta 
         Height          =   240
         Index           =   0
         Left            =   5400
         Top             =   3060
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   18
         Left            =   600
         TabIndex        =   36
         Top             =   3000
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   8
         Left            =   1320
         Picture         =   "frmTesCompensaAutomatica.frx":0000
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Apunte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Index           =   17
         Left            =   240
         TabIndex        =   35
         Top             =   2640
         Width           =   780
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
         Index           =   16
         Left            =   5640
         TabIndex        =   34
         Top             =   1680
         Width           =   570
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
         Index           =   15
         Left            =   9720
         TabIndex        =   33
         Top             =   1680
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   10320
         Picture         =   "frmTesCompensaAutomatica.frx":008B
         Top             =   1680
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
         Index           =   14
         Left            =   8280
         TabIndex        =   32
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha vto."
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
         Index           =   13
         Left            =   7200
         TabIndex        =   31
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   8880
         Picture         =   "frmTesCompensaAutomatica.frx":0116
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   6240
         Picture         =   "frmTesCompensaAutomatica.frx":01A1
         Top             =   1680
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
         Index           =   12
         Left            =   4200
         TabIndex        =   30
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   11
         Left            =   2160
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F. factura"
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
         Index           =   10
         Left            =   3120
         TabIndex        =   28
         Top             =   1920
         Width           =   990
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   4800
         Picture         =   "frmTesCompensaAutomatica.frx":022C
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Raiz cuenta"
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
         Left            =   600
         TabIndex        =   27
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pagos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   660
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
         Index           =   7
         Left            =   9720
         TabIndex        =   25
         Top             =   600
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   10320
         Picture         =   "frmTesCompensaAutomatica.frx":02B7
         Top             =   600
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
         Index           =   6
         Left            =   8280
         TabIndex        =   24
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha vto."
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
         Index           =   5
         Left            =   7200
         TabIndex        =   23
         Top             =   840
         Width           =   1080
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   8880
         Picture         =   "frmTesCompensaAutomatica.frx":0342
         Top             =   600
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
         Index           =   4
         Left            =   5640
         TabIndex        =   22
         Top             =   600
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   6240
         Picture         =   "frmTesCompensaAutomatica.frx":03CD
         Top             =   600
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
         Index           =   3
         Left            =   4200
         TabIndex        =   21
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F. factura"
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
         Left            =   3120
         TabIndex        =   19
         Top             =   840
         Width           =   990
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   4800
         Picture         =   "frmTesCompensaAutomatica.frx":0458
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobros"
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
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Raiz cuenta"
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
         Index           =   69
         Left            =   600
         TabIndex        =   17
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lblConta 
         AutoSize        =   -1  'True
         Caption         =   "Concepto debe"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   46
         Top             =   3540
         Width           =   1620
      End
   End
   Begin VB.Frame FrameVisualizacion 
      Height          =   8655
      Left            =   240
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   15855
      Begin VB.CommandButton cmdImprimir 
         Height          =   450
         Index           =   1
         Left            =   3840
         Picture         =   "frmTesCompensaAutomatica.frx":04E3
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Eliminar cuenta seleccionada"
         Top             =   180
         Width           =   450
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   450
         Index           =   0
         Left            =   4440
         Picture         =   "frmTesCompensaAutomatica.frx":0EE5
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   450
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   13920
         TabIndex        =   51
         Tag             =   "NIF|T|N|||||||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   12360
         TabIndex        =   50
         Tag             =   "NIF|T|N|||||||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDatos 
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
         Left            =   10800
         TabIndex        =   49
         Tag             =   "NIF|T|N|||||||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   7785
         Left            =   120
         TabIndex        =   40
         Top             =   735
         Width           =   15555
         _ExtentX        =   27437
         _ExtentY        =   13732
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   6085
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Proveedor"
            Object.Width           =   6085
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vtos Clie"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Vtos Prov"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cobros"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "pagos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Resultado"
            Object.Width           =   2346
         EndProperty
      End
      Begin VB.Label Label52 
         Caption         =   "Vencimientos a compensar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   41
         Top             =   240
         Width           =   3630
      End
   End
   Begin VB.Label lblIndicador 
      AutoSize        =   -1  'True
      Caption         =   "lblIndicador"
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
      Left            =   240
      TabIndex        =   42
      Top             =   4920
      Width           =   5925
   End
End
Attribute VB_Name = "frmTesCompensaAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmTDia As frmTiposDiario
Attribute frmTDia.VB_VarHelpID = -1

Dim PrimVez As Boolean
Dim Modo As Byte

Dim cad As String
Dim Mc As Contadores

Private Sub cmdCancelar_Click()
    If Modo = 0 Then
        Unload Me
        
    Else
        If Modo = 1 Then
            If cmdCompensar.visible Then
                If MsgBox("Desea cancelar el proceso?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            End If
            PonerModo 0
        End If
    End If
End Sub

Private Sub cmdCompensar_Click()

    If Not DatosOK Then Exit Sub
    
    
    If Modo = 0 Then
        Screen.MousePointer = vbHourglass
        If preparaCamposSQls Then
            PonerModo 1
            
            CargaLw
            
        End If
        lblIndicador.Caption = ""
    

        Screen.MousePointer = vbDefault
        
    Else
        If Modo = 1 Then
            cad = "Fecha: " & Me.txtcodigo(8).Text & vbCrLf & "Concepto debe: " & txConta(1).Text & " " & Me.txtDescConta(1).Text
            cad = cad & vbCrLf & "Concepto haber: " & txConta(2).Text & " " & Me.txtDescConta(2).Text & vbCrLf & vbCrLf & "¿Seguro que desea realizar la compensación?"
            
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
            If REalizaProcesoCompensacion Then
                MsgBox "Proceso  finalizado", vbInformation
                Unload Me
            End If
        End If
    End If
    
End Sub

Private Sub Command1_Click()
'Dim Importe As String
'Dim Cta As String
'Dim MasDe500 As Boolean
'Dim Inser As String
'
'    Inser = ""
'
'    For i = 1 To 1000
'        If False Then
'
'            'cobros(numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci,situacion) 4300 006490
'            Cta = "4300" & Format(i, "000000")
'            cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T")
'            If cad <> "" Then
'                NumRegElim = Int((10 * Rnd) + 1)   '10 %
'                If NumRegElim = 1 Then
'                    NumRegElim = Int((5 * Rnd) + 1)
'                    MasDe500 = NumRegElim > 4
'
'                    If MasDe500 Then
'                        NumRegElim = Int((100000 * Rnd) + 1)
'                    Else
'                        NumRegElim = Int((49999 * Rnd) + 1)
'                    End If
'                    Importe = NumRegElim / 100
'
'
'                    'Cad = "INSERT INTO cobros(numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci,situacion,ctabanc1) "
'                    cad = ", ('B'," & i + 1001 & ",'2020-04-06',1,'" & Cta & "',105,'2020-04-06'," & DBSet(Importe, "N") & ",0,'5720000001')"
'                    Inser = Inser & cad
'                End If
'            End If
'        Else
'            Cta = "4007" & Format(i, "000000")
'            cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T")
'            If cad <> "" Then
'
'                NumRegElim = Int((6 * Rnd) + 1)
'                If NumRegElim = 1 Then
'
'                    NumRegElim = Int((5 * Rnd) + 1)
'                    MasDe500 = NumRegElim > 3
'
'                    If MasDe500 Then
'                        NumRegElim = Int((100000 * Rnd) + 1)
'                    Else
'                        NumRegElim = Int((49999 * Rnd) + 1)
'                    End If
'                    Importe = NumRegElim / 100
'                    NumRegElim = Int((9 * Rnd) + 2)
'
'                    'Cad = "INSERT INTO pagos(numserie,codmacta,numfactu,fecfactu,numorden,codforpa,fecefect,impefect,ctabanc1) "
'                    cad = ", ('1','Liq" & i + 1001 & "-" & Format(NumRegElim, "00") & "','2020-04-01',1,'" & Cta & "',105,'2020-04-" & Format(NumRegElim + 10, "00") & "'," & DBSet(Importe, "N") & ",0,'5720000001')"
'                    Inser = Inser & cad
'                End If
'            End If
'
'
'
'
'
'        End If
'    Next i
'
'
'    If Inser <> "" Then
'        Inser = Mid(Inser, 2) 'primera coma
'        If False Then
'            cad = "INSERT INTO cobros(numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci,situacion,ctabanc1) VALUES " & Inser
'            Conn.Execute cad
'
'        Else
'            cad = " INSERT INTO pagos(numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecefect,impefect,situacion,ctabanc1) VALUES " & Inser
'            Conn.Execute cad
'        End If
'    End If
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    
    If Index = 0 Then
        Imprimir
    Else
        EliminarItem
    End If
End Sub

Private Sub Imprimir()
    InicializarVblesInformesGeneral True
    
    cad = "DELETE FROM tmptesoreriacomun WHERE codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = ""
    'Msg = "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,importe1,importe2,observa1,observa2) VALUES "
    For K = 1 To ListView5.ListItems.Count
        'tmptesoreriacomun (codusu,codigo,texto1,texto2,importe1,importe2,observa1,observa2)
        cad = cad & ", (" & vUsu.Codigo & "," & K & ",'" & txtCta(0).Text & ListView5.ListItems(K).Text & "'," & DBSet(ListView5.ListItems(K).SubItems(2), "T")
        cad = cad & "," & DBSet(ListView5.ListItems(K).SubItems(5), "N") & "," & DBSet(ListView5.ListItems(K).SubItems(6), "N") & ","
        Msg = SeparaVtos(ListView5.ListItems(K).Tag, True, True)
       
        cad = cad & DBSet(Trim(Msg), "T") & ","
        Msg = SeparaVtos(ListView5.ListItems(K).Tag, True, False)

        cad = cad & DBSet(Trim(Msg), "T") & ")"
        
        If Len(cad) > 3000 Then
            cad = Mid(cad, 2)
            Msg = "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,importe1,importe2,observa1,observa2) VALUES "
            cad = Msg & cad
            Ejecuta cad
            cad = ""
        End If
    Next K
    
    If Len(cad) > 0 Then
        cad = Mid(cad, 2)
        Msg = "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,importe1,importe2,observa1,observa2) VALUES "
        cad = Msg & cad
        Ejecuta cad
    End If

    
    'indRPT = IdPrograma & "-00"
    'If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    cadNomRPT = "CompensaAutom.rpt" ' "AsientosHco.rpt"
    cad = DevuelveSQL(2)
    cad = "Seleccion=""" & cad & """|"
    cadParam = cadParam & cad
    numParam = numParam + 1
    vMostrarTree = False
    conSubRPT = False
        

    'si se imprime el nif o la cuenta de cliente
    cadParam = cadParam & "pTitulo=""" & Me.Caption & """|"
    numParam = numParam + 1
    cadFormula = "{tmptesoreriacomun.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        PonerModo 0
        PonFoco txtCta(0)
        If Not BloqueoManual(True, Me.Tag, "1") Then
            cad = DBSet(Me.Tag, "T") & " AND 1"
            cad = DevuelveDesdeBD("codusu", "zbloqueos", cad, "1")
            If Val(cad) <> vUsu.Codigo Then Me.cmdCompensar.Enabled = False
        End If
        txtCta(3).Text = "1"
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimVez = True
    
    For I = 0 To 2
        Me.imgConta(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
    
    Me.Tag = "COMPENSA_AUTO"
    Me.Icon = frmppal.Icon
    lblIndicador.Caption = ""
    Limpiar Me
End Sub



Private Sub PonerFrameVisible(ByRef Fr As frame)
    Fr.top = 30
    Fr.Left = 30
    Fr.visible = True
    Me.Height = Fr.Height + 540 + 720
    Me.Width = Fr.Width + 120
    
    Me.cmdCancelar.top = Fr.Height + 180
    Me.cmdCompensar.top = cmdCancelar.top
    lblIndicador.top = cmdCancelar.top + 45
    cmdCancelar.Left = Fr.Width - cmdCancelar.Width - 240
    cmdCompensar.Left = cmdCancelar.Left - cmdCompensar.Width - 120
        
        
        
    
End Sub



Private Sub PonerModo(vModo As Byte)
    
    Modo = vModo
    If Modo = 0 Then
        PonerFrameVisible Me.FrameTomaDatos
        FrameVisualizacion.visible = False
        Me.cmdCompensar.visible = True

    Else
        FrameTomaDatos.visible = False
        PonerFrameVisible Me.FrameVisualizacion
        
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    BloqueoManual False, Me.Tag, ""
End Sub

Private Sub frmC_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
    cad = CadenaSeleccion
End Sub

Private Sub imgConta_Click(Index As Integer)
    cad = ""
    If Index = 0 Then
        Set frmTDia = New frmTiposDiario
        frmTDia.DatosADevolverBusqueda = "0|1|"
        frmTDia.Show vbModal
        Set frmTDia = Nothing
        
    Else
        Set frmCon = New frmConceptos
        frmCon.DatosADevolverBusqueda = "0|1|"
        frmCon.Show vbModal
        Set frmCon = Nothing
    End If
    If cad <> "" Then
        txConta(Index).Text = RecuperaValor(cad, 1)
        txtDescConta(Index).Text = RecuperaValor(cad, 2)
        PonFoco txConta(Index)
    End If
End Sub

Private Sub imgFecha_Click(Index As Integer)
    
    cad = Format(Now, "dd/mm/yyyy")
    If IsDate(txtcodigo(Index).Text) Then cad = Format(txtcodigo(Index).Text, "dd/mm/yyyy")
    Set frmC = New frmCal
    frmC.Fecha = CDate(cad)
    cad = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If cad <> "" Then txtcodigo(Index).Text = cad
End Sub


Private Sub ListView5_DblClick()
    If ListView5.ListItems.Count = 0 Then Exit Sub
    If ListView5.SelectedItem Is Nothing Then Exit Sub
    
    cad = "Cuentas. " & vbCrLf & "    Cli:   " & txtCta(0).Text & ListView5.SelectedItem.Text & " " & ListView5.SelectedItem.SubItems(1) & vbCrLf
    cad = cad & "    Pro:  " & txtCta(2).Text & ListView5.SelectedItem.Text & " " & ListView5.SelectedItem.SubItems(2) & vbCrLf & vbCrLf
    cad = cad & "Vencimientos." & vbCrLf
    Msg = SeparaVtos(ListView5.SelectedItem.Tag, True, True)
    cad = cad & "    Cob:" & Msg & vbCrLf
    Msg = SeparaVtos(ListView5.SelectedItem.Tag, True, False)
    cad = cad & "    Pag:" & Msg & vbCrLf & vbCrLf
    cad = cad & "Importes." & vbCrLf
    Msg = ListView5.SelectedItem.SubItems(5) & " - " & ListView5.SelectedItem.SubItems(6)
    cad = cad & "    Pendiente:  " & Msg & vbCrLf
    Msg = "cobrar"
    If InStr(1, ListView5.SelectedItem.SubItems(7), "-") > 0 Then Msg = "pagar"
    Msg = Replace(ListView5.SelectedItem.SubItems(7), "-", "") & " a " & Msg
    cad = cad & "    Resultado:     " & Msg & vbCrLf & vbCrLf
    MsgBox cad, vbInformation
    
'    cad = ""
'    For i = 1 To ListView5.ColumnHeaders.Count
'        cad = cad & ListView5.ColumnHeaders(i).Text & ": " & ListView5.ColumnHeaders(i).Width & vbCrLf
'    Next
'    MsgBox cad
    
End Sub

Private Sub txConta_GotFocus(Index As Integer)
    ConseguirFoco txConta(Index), 3
End Sub

Private Sub txConta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txConta_LostFocus(Index As Integer)
    cad = ""
    txConta(Index).Text = Trim(txConta(Index))
    If txConta(Index) <> "" Then
    
        If PonerFormatoEntero(txConta(Index)) Then
            If Index = 0 Then
                'diario
                cad = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txConta(Index).Text)
            Else
                'concepto
                cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txConta(Index).Text)
            End If
            If cad = "" Then
                cad = IIf(Index = 0, "diario", "concepto")
                cad = "No existe el " & cad & " en la BD"
                MsgBox cad, vbExclamation
                cad = ""
            End If
        End If
    End If
    Me.txtDescConta(Index).Text = cad
    If cad = "" Then
        If txConta(Index).Text <> "" Then
            txConta(Index).Text = ""
            PonFoco txConta(Index)
        End If
    End If
End Sub

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtcodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtcodigo_LostFocus(Index As Integer)
    
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    If txtcodigo(Index).Text = "" Then Exit Sub
    
    
    If Not PonerFormatoFecha(txtcodigo(Index)) Then txtcodigo(Index).Text = ""
    
End Sub



Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
    'No hacemos nada
    txtCta(Index).Text = UCase(txtCta(Index).Text)
End Sub


Private Function DatosOK() As Boolean
    
    DatosOK = False

    Select Case Modo
    Case 0
        'Pidiendo datos
        'Solo obligo cuenta
        cad = ""
        If txtCta(0).Text = "" Or txtCta(2).Text = "" Then cad = cad & vbCrLf & "Raiz cuenta cliente/proveedor "
        If txtcodigo(8).Text = "" Then cad = cad & vbCrLf & "Fecha apunte"
        If txtCta(3).Text = "" Then cad = cad & vbCrLf & "Serie pagos"
        If cad = "" Then
            If Len(txtCta(0).Text) <> 4 Or Len(txtCta(2).Text) <> 4 Then cad = "Raiz cuenta cliente/proveedor debe ser longitud 4"
        End If
        
        For I = 0 To 2
            If Me.txConta(I).Text = "" Or txtDescConta(I).Text = "" Then cad = cad & vbCrLf & Me.lblConta(I).Caption & " vacio"
        Next
        
        If cad <> "" Then
            cad = "Campos obligados" & vbCrLf & cad
            MsgBox cad, vbExclamation
            PonFoco txConta(0)
            Exit Function
        End If
        
        
        I = FechaCorrecta2(CDate(txtcodigo(8).Text), True)
        If I > 2 Then Exit Function
        
        
        
        
        
        DatosOK = True
    Case 1
        DatosOK = True
    End Select
End Function



Private Function DevuelveSQL(para As Byte) As String

    
    If para = 0 Then
        DevuelveSQL = " situacion=0 AND cobros.codmacta  like '" & Me.txtCta(0).Text & "%'"
        DevuelveSQL = DevuelveSQL & " AND codrem is null  and coalesce(transfer,0)=0"
        If txtCta(1).Text <> "" Then DevuelveSQL = DevuelveSQL & " AND cobros.numserie =" & DBSet(txtCta(1).Text, "T")
        
        
        For I = 1 To 4
            If Me.txtcodigo(I - 1).Text <> "" Then DevuelveSQL = DevuelveSQL & " AND " & IIf(I < 3, "cobros.fecfactu", "cobros.fecvenci") & IIf((I Mod 2) = 1, ">=", "<=") & DBSet(txtcodigo(I - 1).Text, "F")
        Next
    Else
        If para = 1 Then
            DevuelveSQL = " situacion =0 AND pagos.codmacta  like '" & Me.txtCta(2).Text & "%'"
            DevuelveSQL = DevuelveSQL & " AND nrodocum is null AND   emitdocum =0"
            If txtCta(3).Text <> "" Then DevuelveSQL = DevuelveSQL & " AND pagos.numserie =" & DBSet(txtCta(3).Text, "T")
            For I = 5 To 8
                If Me.txtcodigo(I - 1).Text <> "" Then DevuelveSQL = DevuelveSQL & " AND " & IIf(I < 7, "pagos.fecfactu", "pagos.fecefect") & IIf((I Mod 2) = 1, ">=", "<=") & DBSet(txtcodigo(I - 1).Text, "F")
            Next
        Else
            'Cadena seleccion para RPT
            'Cliente
            DevuelveSQL = "COBROS:   Raiz " & Me.txtCta(0).Text
            If txtCta(1).Text <> "" Then DevuelveSQL = DevuelveSQL & " Serie " & txtCta(1).Text
        
            Msg = ""
            For J = 0 To 1
                I = IIf(J = 0, 0, 2)
                Msg = Msg & ParaCadena(CInt(I)) & " "
                Msg = Msg & ParaCadena(CInt(I) + 1) & " "
                Msg = Trim(Msg)
                If Msg <> "" Then
                    Msg = IIf(I = 0, "F. fact ", "F.vto ") & Msg
                    DevuelveSQL = DevuelveSQL & "    " & Msg
                    Msg = ""
                End If
            Next
            DevuelveSQL = DevuelveSQL & """ + chr(13) + ""PAGOS:      Raiz " & Me.txtCta(2).Text
            If txtCta(3).Text <> "" Then DevuelveSQL = DevuelveSQL & " Serie " & txtCta(3).Text
            Msg = ""
            For J = 0 To 1
                I = IIf(J = 0, 4, 6)
                Msg = Msg & ParaCadena(CInt(I)) & " "
                Msg = Msg & ParaCadena(CInt(I) + 1) & " "
                Msg = Trim(Msg)
                If Msg <> "" Then
                    Msg = IIf(I = 4, "F. fact ", "F.vto ") & Msg
                    DevuelveSQL = DevuelveSQL & "    " & Msg
                    Msg = ""
                End If
            Next
        
        
        
        End If
    End If

      
End Function

Private Function ParaCadena(Indice As Integer) As String
    ParaCadena = ""
    If Me.txtcodigo(Indice).Text <> "" Then ParaCadena = IIf((Indice Mod 2) = 0, " desde ", " hasta ") & txtcodigo(Indice).Text
    
End Function

Private Function preparaCamposSQls() As Boolean

On Error GoTo epreparaCamposSQls
    preparaCamposSQls = False
    lblIndicador.Caption = "Cobros"
    lblIndicador.Refresh
    
    
    Conn.Execute "DELETE FROM tmpcompensaAuto WHERE codusu = " & vUsu.Codigo
    Conn.Execute "DELETE FROM tmpfaclin WHERE codusu = " & vUsu.Codigo
    
    
    cad = "select substring(codmacta,5) id," & vUsu.Codigo & ", sum(impvenci+coalesce(gastos,0)-coalesce(impcobro,0)) importe,"
    cad = cad & "  GROUP_CONCAT( substring(concat(numserie,'   '),1,3),numfactu,'(',fecfactu,')·',numorden separator '|')  facturas,1 cobro,count(*) cuantos"
    cad = cad & " from cobros where " & DevuelveSQL(0) & " GROUP BY 1"
    cad = "INSERT INTO tmpcompensaAuto(id,codusu,importe,facturas,cobro,cuantasFra) " & cad
    Conn.Execute cad
    
    lblIndicador.Caption = "Pagos"
    lblIndicador.Refresh

    cad = "SELECT substring(codmacta,5) id," & vUsu.Codigo & ",sum(impefect-coalesce(imppagad,0)) importe,GROUP_CONCAT( numfactu,'(',fecfactu,')','·',numorden separator '|') "
    cad = cad & " ,0 cobro ,count(*) cuantos"
    cad = cad & " from pagos where  " & DevuelveSQL(1) & " GROUP BY 1"
    cad = "INSERT INTO tmpcompensaAuto(id,codusu,importe,facturas,cobro,cuantasFra) " & cad
    Conn.Execute cad
        
    
    
    lblIndicador.Caption = "comprobaciones"
    lblIndicador.Refresh

    espera 0.3
    Set miRsAux = New ADODB.Recordset
    cad = "select sum(if(cobro=1,1,0)) cobros, sum(if(cobro=0,1,0)) pagos  from tmpcompensaAuto where codusu =" & vUsu.Codigo
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    If miRsAux.EOF Then
        cad = "N"
    Else
        If DBLet(miRsAux.Fields(0), "N") = 0 And DBLet(miRsAux.Fields(1), "N") = 0 Then cad = "N"
    End If
    miRsAux.Close
    If cad <> "" Then Err.Raise 513, , "Ningun dato generado"


    'Los añadire en una tabla para despues mostrar el resumen
    'Si que hay datos, Vamos a dejar solo lo los que comensan
    cad = "select id,sum(if(cobro=1,1,0)) cobros, sum(if(cobro=0,1,0)) pagos  "
    cad = cad & " from tmpcompensaAuto where codusu =" & vUsu.Codigo & " GROUP BY 1"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        If DBLet(miRsAux!Cobros, "N") = 0 Or DBLet(miRsAux!Pagos, "N") = 0 Then cad = cad & ", '" & miRsAux!Id & "'"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If cad <> "" Then
        cad = Mid(cad, 2) 'quitmoa las primera coma
        
        'Lo metemos en la tmp
        Msg = "INSERT INTO tmpfaclin(codusu,codigo,Numfac,Imponible,tipoiva) "
        Msg = Msg & " select codusu,id,concat(if(cobro=1,'" & Me.txtCta(0).Text & "','" & txtCta(2).Text & "'),id),importe,cobro"
        Msg = Msg & " From tmpcompensaAuto"
        Msg = Msg & " where codusu =" & vUsu.Codigo & " AND id IN (" & cad & ")"
        Conn.Execute Msg
        
        
        
        cad = "DELETE FROM  tmpcompensaAuto where codusu =" & vUsu.Codigo & " AND id IN (" & cad & ")"
        Conn.Execute cad
    End If

    cad = "Select count(*)  FROM  tmpcompensaAuto where codusu =" & vUsu.Codigo
    espera 0.5
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then If DBLet(miRsAux.Fields(0), "N") > 0 Then cad = ""
    miRsAux.Close
    If cad <> "" Then Err.Raise 513, , "Ningun dato para compensar"
    
    'Llegados aqui. SI se compensa
    preparaCamposSQls = True

epreparaCamposSQls:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
End Function







'Cargar datos sobre lw
Private Sub CargaLw()
Dim IT As ListItem
Dim rCta As ADODB.Recordset 'Las cuentas
Dim C2 As String
Dim Importe As Currency

On Error GoTo eCargaLw

    cmdCompensar.visible = False
    lblIndicador.Caption = "Leyendo cuentas contables"
    lblIndicador.Refresh
    
    For J = 0 To 2
        BloqueaTXT Me.txtDatos(J), True
        txtDatos(J).Tag = 0
        txtDatos(J).Text = ""
    Next
    
    
    DoEvent2
    Me.ListView5.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    
    
    
    cad = "select  id from tmpcompensaAuto where codusu =" & vUsu.Codigo & " GROUP BY 1"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    C2 = ""
    While Not miRsAux.EOF
        cad = cad & ", '" & txtCta(0).Text & miRsAux!Id & "'"
        C2 = C2 & ", '" & txtCta(2).Text & miRsAux!Id & "'"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If cad = "" Then Err.Raise 513, , "Error leyendo cuentas(cadena vacia). "
    
    
    Set rCta = New ADODB.Recordset
    C2 = Mid(C2, 2)
    cad = C2 & cad
    cad = "Select codmacta,nommacta from cuentas WHERE apudirec='S' and codmacta in (" & cad & ") ORDER BY codmacta"
    rCta.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If rCta.EOF Then Err.Raise 513, , "Error leyendo cuentas(cadena vacia). "
    
    lblIndicador.Caption = "Leyendo cuentas contables"
    lblIndicador.Refresh
    cad = "select  * from tmpcompensaAuto where codusu =" & vUsu.Codigo & " ORDER BY id,cobro desc"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    'Van a pares
    J = 0 'Si hay error en cuentas NO puede seguir
    While Not miRsAux.EOF
        
        lblIndicador.Caption = "Id " & miRsAux!Id
        lblIndicador.Refresh
        
        'El primero tiene que ser un cobros
        If miRsAux!Cobro = 0 Then Err.Raise 513, , "No existe registro cobros  para id: " & miRsAux!Id
        
        Set IT = ListView5.ListItems.Add(, "K" & miRsAux!Id)
        NumRegElim = Val(miRsAux!Id)
        IT.Text = miRsAux!Id
        
        
        
        'Cliente
        cad = "codmacta = '" & txtCta(0).Text & miRsAux!Id & "'"
        rCta.Find cad, , adSearchForward, 1
        If rCta.EOF Then
            cad = "error obteniendo cta "
            I = 1
        Else
            cad = rCta!Nommacta
            I = 0
        End If
        IT.SubItems(1) = CStr(cad)
        IT.SubItems(2) = " "
        Msg = miRsAux!facturas
        Msg = Replace(Msg, "·", ":")
        Msg = Replace(Msg, "|", "  ")
        If miRsAux!cuantasfra > 1 Then Msg = "(" & miRsAux!cuantasfra & ") " & Msg
        IT.Tag = miRsAux!facturas
        IT.SubItems(3) = Msg
        IT.SubItems(4) = " "
        Importe = miRsAux!Importe
        txtDatos(0).Tag = txtDatos(0).Tag + miRsAux!Importe
        
        IT.SubItems(5) = Format(Importe, FormatoImporte)
        IT.SubItems(6) = " "
        
        miRsAux.MoveNext
        If miRsAux.EOF Then Err.Raise 513, , "No existe registro prov para: " & NumRegElim
        If miRsAux!Cobro = 1 Then Err.Raise 513, , "No existe registro prov(2) para: " & NumRegElim
        
        
        'proveedor
        C2 = "codmacta = '" & txtCta(2).Text & miRsAux!Id & "'"
        rCta.Find C2, , adSearchForward, 1
        If rCta.EOF Then
            C2 = "error obteniendo cta "
            I = 1
        Else
            C2 = rCta!Nommacta
        End If
        IT.SubItems(2) = CStr(C2)
        
        
        
        'Quitamos comas, puntos espacioes
        If I = 0 Then
            ReemplazaCarcateres cad
            ReemplazaCarcateres C2
            If cad <> C2 Then I = 2
        End If
        
        'Facturas
        Msg = miRsAux!facturas
        Msg = Replace(Msg, "·", ":")
        Msg = Replace(Msg, "|", "  ")
        If miRsAux!cuantasfra > 1 Then Msg = "(" & miRsAux!cuantasfra & ") " & Msg
        IT.SubItems(4) = Msg
        IT.Tag = IT.Tag & "@@" & miRsAux!facturas
        IT.SubItems(6) = Format(miRsAux!Importe, FormatoImporte)
        Importe = Importe - miRsAux!Importe
        IT.SubItems(7) = Format(Importe, FormatoImporte)
        txtDatos(1).Tag = txtDatos(1).Tag + miRsAux!Importe
        If I > 0 Then
            If I = 1 Then
                IT.ForeColor = vbRed
                J = 1
            Else
                If I = 2 Then IT.ForeColor = vbBlue
            End If
        End If
        
        'Total
        txtDatos(2).Tag = txtDatos(2).Tag + Importe
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    Me.cmdCompensar.visible = J = 0
        
    For J = 0 To 2
        txtDatos(J).Text = Format(txtDatos(J).Tag, FormatoImporte)
    Next
    

    
eCargaLw:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    Set rCta = Nothing
    lblIndicador.Caption = ""
End Sub

Private Sub ReemplazaCarcateres(nombrecuenta As String)
    nombrecuenta = UCase(nombrecuenta)
    nombrecuenta = Replace(nombrecuenta, ".", "")
    nombrecuenta = Replace(nombrecuenta, ",", "")
    nombrecuenta = Replace(nombrecuenta, "º", "")
    nombrecuenta = Replace(nombrecuenta, "ª", "")
    nombrecuenta = Replace(nombrecuenta, " ", "")
End Sub



Private Function REalizaProcesoCompensacion() As Boolean
Dim N As Integer
Dim B As Boolean
Dim LineaApu As Integer

    'Para cada cli/prove vamos a
    '   1.- Crear un texto para las amplaciones de cobro / pago
    '   2.- Dar por cobrado/pagado los que hay que dar por pagados/cobrados
    '       Insertaremos en apunte
    '

    REalizaProcesoCompensacion = True
    
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", CDate(txtcodigo(8).Text) <= vParam.fechafin, False) = 0 Then

        cad = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        
        cad = cad & txConta(0).Text & ",'" & Format(CDate(txtcodigo(8).Text), FormatoFecha) & "'," & Mc.Contador & ","
        'Observaciones
        cad = cad & "'Compensacion automática.  Por " & vUsu.Nombre & " el " & Format(Now, "dd/mm/yyyy") & "',"
        '
        cad = cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: compensacion automatica')"
        Conn.Execute cad
        LineaApu = 0
        
        'Hco compensaciones
        cad = DevuelveDesdeBD("max(codigo)", "compensaclipro", "1", "1")
        If cad = "" Then cad = "0"
        NumRegElim = Val(cad)
        
         
        
        
                
        For N = 1 To ListView5.ListItems.Count
            Conn.BeginTrans
            Me.lblIndicador.Caption = I & " de " & ListView5.ListItems.Count & " - " & ListView5.ListItems(N).SubItems(1)
            lblIndicador.Refresh
            
            
            B = CompensarUnCliProv(N, LineaApu)
            
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
                Ejecuta "DELETE FROM compensaclipro where codigo=" & NumRegElim
                Ejecuta "DELETE FROM compensaclipro_facturas where codigo=" & NumRegElim
                NumRegElim = NumRegElim - 1
            End If
            If (N Mod 10) = 0 Then
                Me.Refresh
                DoEvent2
                Screen.MousePointer = vbHourglass
            End If
            
        Next N
        
                    
    Else
        REalizaProcesoCompensacion = False
    End If
    Set Mc = Nothing
    Set miRsAux = Nothing
    lblIndicador.Caption = ""
End Function





Private Function CompensarUnCliProv(Linea As Integer, ByRef LinApu As Integer) As Boolean
Dim ImporteFinal As Currency
Dim FraP As String
Dim FraC As String
Dim Hlinapu As String
Dim LineasCompensa As String
Dim ContadorLinCompensa As Integer
Dim ImporAux As Currency
Dim Aux As String
Dim C1 As String
Dim C2 As String
Dim c3 As String
Dim CierraVto As Boolean
Dim VtoCompensa As String
Dim Cuenta As String

    On Error GoTo eCompensarUnCliProv
    Set miRsAux = New ADODB.Recordset
    CompensarUnCliProv = False
    LineasCompensa = ""
    ContadorLinCompensa = 0
    Hlinapu = ""
    
    'compensaclipro_facturas codigo,linea,EsCobro,numserie,numfactu,fecfactu,numorden,importe,gastos,impcobro,fechavto ,compensado,destino
    
    ImporteFinal = ImporteFormateado(ListView5.ListItems(Linea).SubItems(7))
    
    'historico compensaciones
    NumRegElim = NumRegElim + 1
    cad = "INSERT INTO compensaclipro(codigo,autom,fecha,login,PC,codmacta,nommacta,resultado,fechahora) VALUES ("
    cad = cad & NumRegElim & ",1," & DBSet(txtcodigo(8).Text, "F") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(vUsu.PC, "T")
    cad = cad & ",'" & txtCta(0).Text & ListView5.ListItems(Linea).Text & "'," & DBSet(ListView5.ListItems(Linea).SubItems(1), "T")
    cad = cad & "," & DBSet(ListView5.ListItems(Linea).SubItems(7), "N") & "," & DBSet(Now, "FH") & ")"
    Conn.Execute cad
    
    
    
    '-----------------------------------------------------------------------------
    'La linea del asiento
    'Hemos puesto hlinapu mas atras para poder cambiarla
    'Vemos los cobros y updateamos y memtemos en hlinapu
    
    
    
    '.---------------------------------------------------------------------------------------
    '.---------------------------------------------------------------------------------------
    ' CLIENTES
    cad = SeparaVtos(ListView5.ListItems(Linea).Tag, False, True)
    If cad = "" Then Err.Raise 513, , "Obteninedo cobros " & ListView5.ListItems(Linea).Tag
    cad = Trim(Replace(cad, vbCrLf, ", "))
    'QUITAMOS  la ultima coma
    cad = Mid(cad, 1, Len(cad) - 1)
    Msg = "Select  NUmSerie , NumFactu, FecFactu, numorden ,impvenci,gastos,codmacta,fecultco,impcobro,observa,fecvenci "
    Msg = Msg & " from cobros WHERE (NUmSerie , NumFactu, FecFactu, numorden) IN ("
    Msg = Msg & cad & ") ORDER by fecfactu,fecvenci,numserie,numfactu"
    
    miRsAux.Open Msg, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    VtoCompensa = ""
    If ImporteFinal > 0 Then
        While Not miRsAux.EOF
            ImporAux = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
            If ImporAux > ImporteFinal Then VtoCompensa = miRsAux!NUmSerie & Format(miRsAux!NumFactu, "0000000") & "-" & miRsAux!numorden
            
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        If VtoCompensa = "" Then Err.Raise 513, , "Imposible obtener vencimiento destino cobros. Vto final: " & ImporteFinal
        
    End If
    Cuenta = txtCta(0).Text & ListView5.ListItems(Linea).Text
    While Not miRsAux.EOF
        
        ImporAux = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
        If ImporAux = 0 Then Err.Raise 513, , "Cobrado y situacion pone pdte"
            
        'compensaclipro_facturas codigo,linea,EsCobro,codmacta,numserie,numfactu,fecfactu,numorden,importe,gastos,impcobro,fechavto ,compensado,destino
        ContadorLinCompensa = ContadorLinCompensa + 1
        C1 = ", (" & NumRegElim & "," & ContadorLinCompensa & ",1," & DBSet(Cuenta, "T") & "," & DBSet(miRsAux!NUmSerie, "T") & ",'" & Format(miRsAux!NumFactu, "000000") & "',"
        C1 = C1 & DBSet(miRsAux!FecFactu, "F") & "," & miRsAux!numorden & "," & DBSet(miRsAux!ImpVenci, "N") & ","
        C1 = C1 & DBSet(miRsAux!Gastos, "N", "N") & "," & DBSet(miRsAux!impcobro, "N", "N") & "," & DBSet(miRsAux!FecVenci, "F") & ",@#I@#,@#D@#)"
        
        '(numdiari, fechaent, numasien, numdocum, ampconce, codconce, linliapu, codmacta, timporteD, timporteH, ctacontr, codccost, idcontab, punteada
        'numserie numfaccl numfacpr fecfactu numorden tipforpa
        C2 = ", (" & txConta(0).Text & "," & DBSet(txtcodigo(8).Text, "F") & "," & Mc.Contador & ",'" & miRsAux!NUmSerie & Format(miRsAux!NumFactu, "000000") & "',"
        C2 = C2 & DBSet(miRsAux!NUmSerie & Format(miRsAux!NumFactu, "000000") & " (" & miRsAux!numorden & ") " & Format(miRsAux!FecFactu, "dd/mm/yyyy") & " Vto " & Format(miRsAux!FecVenci, "dd/mm/yyyy"), "T")
        LinApu = LinApu + 1   'Linea apunte
        C2 = C2 & "," & txConta(2).Text & "," & LinApu & ",'" & Me.txtCta(0).Text & ListView5.ListItems(Linea).Text & "',NULL,@#I@#"
        C2 = C2 & ",'" & Me.txtCta(2).Text & ListView5.ListItems(Linea).Text & "',NULL,'COBROS',0"
        C2 = C2 & "," & DBSet(miRsAux!NUmSerie, "T") & "," & miRsAux!NumFactu & ",null," & DBSet(miRsAux!FecFactu, "F") & "," & miRsAux!numorden & ",1)"
        
        'Para el update del cobro
        c3 = "numserie = " & DBSet(miRsAux!NUmSerie, "T") & " AND numfactu = " & miRsAux!NumFactu
        c3 = c3 & "  AND fecfactu = " & DBSet(miRsAux!FecFactu, "F") & " AND numorden = " & miRsAux!numorden
        
        CierraVto = True
        If VtoCompensa <> "" Then
            Msg = miRsAux!NUmSerie & Format(miRsAux!NumFactu, "0000000") & "-" & miRsAux!numorden
            If Msg = VtoCompensa Then CierraVto = False
        End If
        miRsAux.MoveNext
        
        If Not CierraVto Then
            J = 1
            cad = "   Pdte. ant/act: " & ImporAux & "   // " & Abs(ImporteFinal)
            Msg = ", situacion=0 , impcobro = null,  fecultco = null , impvenci=" & DBSet(Abs(ImporteFinal), "N")
            'Cuanto hemos compensado
            ImporAux = ImporAux - ImporteFinal
        Else
            J = 0
            cad = ""
            Msg = ", situacion=1 , impcobro = " & DBSet(ImporAux, "N") & ", fecultco = " & DBSet(Now, "F") & ""
            
        End If
        
        cad = "Compensa cli-prov: " & Format(NumRegElim, "00000") & " " & Now & cad
        Msg = " observa =" & DBSet(cad, "T") & Msg
        C1 = Replace(C1, "@#D@#", J) 'es vto destino
        C1 = Replace(C1, "@#I@#", DBSet(ImporAux, "N")) 'es vto destino
        cad = "UPDATE cobros set " & Msg & " WHERE " & c3
        Conn.Execute cad
        LineasCompensa = LineasCompensa & C1
        C2 = Replace(C2, "@#I@#", DBSet(ImporAux, "N")) 'es vto destino
        Hlinapu = Hlinapu & C2
        
       
    Wend
    miRsAux.Close
    
    '.---------------------------------------------------------------------------------------
    '.---------------------------------------------------------------------------------------
    'Proveeedores
    '
    cad = SeparaVtos(ListView5.ListItems(Linea).Tag, False, False)
    If cad = "" Then Err.Raise 513, , "Obteninedo pagos " & ListView5.ListItems(Linea).Tag
    cad = Trim(Replace(cad, vbCrLf, ", "))
    cad = Mid(cad, 1, Len(cad) - 1) 'QUITAMOS  la ultima coma
    C2 = "'" & Me.txtCta(2).Text & ListView5.ListItems(Linea).Text & "'"
    cad = Trim(Replace(cad, "#cta#", C2))
    
    
    Msg = "Select  numserie,codmacta,numfactu,fecfactu,numorden,fecefect,impefect,fecultpa,observa,imppagad "
    Msg = Msg & " from pagos WHERE numserie =" & txtCta(3).Text & " AND (NumFactu, FecFactu, numorden,codmacta) IN ("
    Msg = Msg & cad & ") ORDER by fecfactu,fecefect,codmacta"
    
    miRsAux.Open Msg, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    VtoCompensa = ""
    If ImporteFinal < 0 Then
        While Not miRsAux.EOF
            ImporAux = miRsAux!ImpEfect + DBLet(miRsAux!imppagad, "N")
            If ImporAux > Abs(ImporteFinal) Then VtoCompensa = miRsAux!NUmSerie & "-" & miRsAux!NumFactu & "-" & miRsAux!numorden & miRsAux!FecFactu
            
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        If VtoCompensa = "" Then Err.Raise 513, , "Imposible obtener vencimiento destino pagos. Vto final: " & ImporteFinal
        
    End If
    
    
    
    Cuenta = txtCta(2).Text & ListView5.ListItems(Linea).Text
    While Not miRsAux.EOF
        
        ImporAux = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
        If ImporAux = 0 Then Err.Raise 513, , "pago y situacion pone pdte"
            
        'compensaclipro_facturas codigo,linea,EsCobro,cuenta,numserie,numfactu,fecfactu,numorden,importe,gastos,impcobro,fechavto ,compensado,destino
        ContadorLinCompensa = ContadorLinCompensa + 1
        C1 = ", (" & NumRegElim & "," & ContadorLinCompensa & ",0," & DBSet(Cuenta, "T") & "," & txtCta(3).Text & "," & DBSet(miRsAux!NumFactu, "T") & ","
        C1 = C1 & DBSet(miRsAux!FecFactu, "F") & "," & miRsAux!numorden & "," & DBSet(miRsAux!ImpEfect, "N") & ",null,"
        C1 = C1 & DBSet(miRsAux!imppagad, "N", "N") & "," & DBSet(miRsAux!fecefect, "F") & ",@#I@#,@#D@#)"
        
        '(numdiari, fechaent, numasien, numdocum, ampconce, codconce, linliapu, codmacta, timporteD, timporteH, ctacontr, codccost, idcontab, punteada
        'numserie numfaccl numfacpr fecfactu numorden tipforpa
        C2 = ", (" & txConta(0).Text & "," & DBSet(txtcodigo(8).Text, "F") & "," & Mc.Contador & "," & DBSet(miRsAux!NumFactu, "T") & ","
        C2 = C2 & DBSet(miRsAux!NumFactu & " (" & miRsAux!numorden & ") " & Format(miRsAux!FecFactu, "dd/mm/yyyy") & " Vto " & Format(miRsAux!fecefect, "dd/mm/yyyy"), "T")
        LinApu = LinApu + 1
        C2 = C2 & "," & txConta(1).Text & "," & LinApu & ",'" & Me.txtCta(2).Text & ListView5.ListItems(Linea).Text & "',@#I@#,NULL"
        C2 = C2 & ",'" & Me.txtCta(0).Text & ListView5.ListItems(Linea).Text & "',NULL,'PAGOS',0,"
        C2 = C2 & DBSet(miRsAux!NUmSerie, "T") & ",null," & DBSet(miRsAux!NumFactu, "T") & "," & DBSet(miRsAux!FecFactu, "F") & "," & miRsAux!numorden & ",1)"
        
        'Para el update del cobro
        c3 = "numserie = " & DBSet(miRsAux!NUmSerie, "T") & " AND numfactu = " & DBSet(miRsAux!NumFactu, "T")
        c3 = c3 & "  AND fecfactu = " & DBSet(miRsAux!FecFactu, "F") & " AND numorden = " & miRsAux!numorden
        c3 = c3 & "  AND codmacta = '" & Me.txtCta(2).Text & ListView5.ListItems(Linea).Text & "'"
        
        
        CierraVto = True
        If VtoCompensa <> "" Then
            Msg = miRsAux!NUmSerie & "-" & miRsAux!NumFactu & "-" & miRsAux!numorden & miRsAux!FecFactu
            
            If Msg = VtoCompensa Then CierraVto = False
        End If
        miRsAux.MoveNext
        
        
        If CierraVto Then
            cad = ""
            J = 0
            Msg = ", situacion=1 , imppagad = " & DBSet(ImporAux, "N") & ",  fecultpa = " & DBSet(Now, "F") & ""
        Else
            J = 1
            cad = "   Pdte ant/act " & ImporAux & " // " & Abs(ImporteFinal)
            Msg = ", situacion=0 , imppagad = null,  fecultpa = null , impefect =" & DBSet(Abs(ImporteFinal), "N")
            'Cuanto hemos compensado
             ImporAux = ImporAux + ImporteFinal
        End If
        cad = "Compensa cli-prov: " & Format(NumRegElim, "00000") & " " & Now & cad
        Msg = " observa =" & DBSet(cad, "T") & Msg
        C1 = Replace(C1, "@#D@#", J) 'es vto destino
        C1 = Replace(C1, "@#I@#", DBSet(ImporAux, "N")) 'es vto destino
        cad = ""
        LineasCompensa = LineasCompensa & C1
                
                
        cad = "UPDATE pagos set " & Msg & " WHERE " & c3
        Conn.Execute cad
                
        C2 = Replace(C2, "@#I@#", DBSet(ImporAux, "N")) 'es vto destino
        Hlinapu = Hlinapu & C2
        
        
    Wend
    miRsAux.Close
    
    
    
    
    
    
    
    
    
    
    'Hco compensaciones
    cad = "INSERT INTO compensaclipro_facturas( codigo,linea,EsCobro,codmacta,numserie,numfactu,fecfactu,numorden,importe,gastos,impcobro,fechavto ,compensado,destino) VALUES "
    LineasCompensa = Mid(LineasCompensa, 2)
    cad = cad & LineasCompensa
    Conn.Execute cad
     
    
    
    'Insertamos en apuntes
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, numdocum, ampconce, codconce, linliapu, codmacta, timporteD, timporteH, ctacontr, "
    cad = cad & " codccost, idcontab, punteada,numserie ,numfaccl ,numfacpr ,fecfactu ,numorden ,tipforpa) VALUES "
    Hlinapu = Mid(Hlinapu, 2)
    cad = cad & Hlinapu
    Conn.Execute cad
    
    CompensarUnCliProv = True
    Set miRsAux = Nothing
    Exit Function
    
eCompensarUnCliProv:
    
    MuestraError Err.Number, , Err.Description
    Set miRsAux = Nothing
    
End Function


Private Function SeparaVtos(ByVal CADENA As String, Normal As Boolean, Cobros As Boolean) As String
Dim Aux As String
Dim C1 As String
    
    
    I = InStr(1, CADENA, "@@")
    If Cobros Then
        CADENA = Mid(CADENA, 1, I - 1)
    Else
        CADENA = Mid(CADENA, I + 2)
    End If
    SeparaVtos = ""
    While CADENA <> ""
        I = InStr(1, CADENA, "|")
        'No puede ser eof
        If I = 0 Then
            Aux = CADENA
            CADENA = ""
        Else
            Aux = Mid(CADENA, 1, I - 1)
            CADENA = Mid(CADENA, I + 1)
        End If
        
        I = InStr(1, Aux, "·")
        J = CInt(Mid(Aux, I + 1, 1))
        I = InStr(1, Aux, "(")
        C1 = Mid(Aux, I + 1)
        C1 = Mid(C1, 1, 10)
        Aux = Mid(Aux, 1, I - 1)
        If Cobros Then
            If Normal Then
                Aux = Trim(Mid(Aux, 1, 3)) & Trim(Mid(Aux, 4)) & "-" & J & " de " & C1
            Else
                Aux = "('" & Trim(Mid(Aux, 1, 3)) & "'," & Trim(Mid(Aux, 4)) & ",'" & C1 & "'," & J & ")"
            End If
        Else
            
            If Normal Then
                Aux = Aux & "-" & J & " de " & C1
            Else
                
                Aux = "(" & DBSet(Aux, "T") & ",'" & C1 & "'," & J & ",#cta#)"
            End If
        End If
        SeparaVtos = SeparaVtos & Aux & vbCrLf
    Wend
    
End Function


Private Sub EliminarItem()
Dim Cob As Currency
Dim Pag As Currency
    If Me.ListView5.ListItems.Count = 0 Then Exit Sub
    
    J = 0
    Cob = 0
    Pag = 0
    cad = ""
    For I = 1 To ListView5.ListItems.Count
        If ListView5.ListItems(I).Selected Then
            J = J + 1
            cad = cad & vbCrLf & "  -" & ListView5.ListItems(I).Text & " " & ListView5.ListItems(I).SubItems(1)
            Cob = Cob + ImporteFormateado(ListView5.ListItems(I).SubItems(5))
            Pag = Pag + ImporteFormateado(ListView5.ListItems(I).SubItems(6))
            NumRegElim = I 'para despues situar el lw
        End If
    Next I
    
    If J = 0 Then
        MsgBox "Seleccione algun elemento para eliminar", vbExclamation
        Exit Sub
    End If
    
    
    Msg = "Ctas a eliminar: " & J & "    Cobros : " & Format(Cob, FormatoImporte) & "    Pagos : " & Format(Pag, FormatoImporte) & vbCrLf
    If Len(cad) > 750 Then cad = cad & " ... ... .."
    cad = "Va a quitar de las compensaciones: " & vbCrLf & Msg & cad & vbCrLf & "¿Continuar?"
    If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    'QUitar significa borrar de tempcomensauto , y lo pasamos a tmpfaclin para que el sumatorio total cuadre
    cad = ""
    For I = 1 To ListView5.ListItems.Count
        If ListView5.ListItems(I).Selected Then
            cad = cad & ", " & ListView5.ListItems(I).Text
        End If
    Next I
    
    cad = Mid(cad, 2)
    
    Msg = "INSERT INTO tmpfaclin(codusu,codigo,Numfac,Imponible,tipoiva) "
    Msg = Msg & " select codusu,id,concat(if(cobro=1,'" & Me.txtCta(0).Text & "','" & txtCta(2).Text & "'),id),importe,cobro"
    Msg = Msg & " From tmpcompensaAuto"
    Msg = Msg & " where codusu =" & vUsu.Codigo & " AND id IN (" & cad & ")"
    Conn.Execute Msg
    
    Msg = "DELETE From tmpcompensaAuto"
    Msg = Msg & " where codusu =" & vUsu.Codigo & " AND id IN (" & cad & ")"
    Conn.Execute Msg
    
    
    CargaLw
    
    If ListView5.ListItems.Count > 0 Then
        If NumRegElim > ListView5.ListItems.Count Then NumRegElim = ListView5.ListItems.Count
        ListView5.ListItems(NumRegElim).EnsureVisible
        Set ListView5.SelectedItem = Nothing
        
    End If
End Sub
