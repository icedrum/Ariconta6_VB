VERSION 5.00
Begin VB.Form frmPresuGenerar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar presupuestos"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frmPresuGenerar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMasiva 
      Caption         =   "Generacion masiva"
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
      Height          =   255
      Left            =   3570
      TabIndex        =   82
      Top             =   6840
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CheckBox ChkEliminar 
      Caption         =   "Eliminar datos (si los tuviera)"
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
      Left            =   120
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7170
      TabIndex        =   17
      Top             =   6780
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
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
      Left            =   5970
      TabIndex        =   16
      Top             =   6780
      Width           =   975
   End
   Begin VB.Frame FrMasiva 
      Height          =   6615
      Left            =   120
      TabIndex        =   64
      Top             =   60
      Width           =   7995
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   1290
         TabIndex        =   78
         Top             =   1680
         Width           =   3735
         Begin VB.OptionButton optEjer 
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
            Left            =   240
            TabIndex        =   80
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optEjer 
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
            Left            =   2280
            TabIndex        =   79
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.OptionButton optIncre 
         Caption         =   "Presupuesto anterior"
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
         Left            =   4440
         TabIndex        =   77
         Top             =   3120
         Width           =   2565
      End
      Begin VB.OptionButton optIncre 
         Caption         =   "Ejercicio"
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
         Left            =   2880
         TabIndex        =   76
         Top             =   3120
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.TextBox txtInc 
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
         Left            =   1560
         TabIndex        =   74
         Top             =   3105
         Width           =   915
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
         Left            =   1530
         TabIndex        =   70
         Top             =   990
         Width           =   1350
      End
      Begin VB.TextBox txtDesCta 
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
         Left            =   2910
         TabIndex        =   69
         Top             =   990
         Width           =   4395
      End
      Begin VB.TextBox txtDesCta 
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
         Left            =   2910
         TabIndex        =   66
         Top             =   600
         Width           =   4395
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
         Left            =   1530
         TabIndex        =   65
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
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
         Left            =   240
         TabIndex        =   81
         Top             =   6240
         Width           =   7335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   75
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incremento"
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
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   73
         Top             =   2670
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   72
         Top             =   1800
         Width           =   900
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
         Index           =   4
         Left            =   510
         TabIndex        =   71
         Top             =   1005
         Width           =   735
      End
      Begin VB.Image imgcta 
         Height          =   240
         Index           =   2
         Left            =   1260
         Top             =   1020
         Width           =   240
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
         Index           =   24
         Left            =   510
         TabIndex        =   68
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   23
         Left            =   180
         TabIndex        =   67
         Top             =   360
         Width           =   780
      End
      Begin VB.Image imgcta 
         Height          =   240
         Index           =   1
         Left            =   1260
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Frame FrameN 
      Height          =   6615
      Left            =   120
      TabIndex        =   19
      Top             =   60
      Width           =   7995
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
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1350
      End
      Begin VB.TextBox txtDesCta 
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
         Left            =   1500
         TabIndex        =   20
         Top             =   360
         Width           =   4185
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ejercicio siguiente"
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
         Left            =   5760
         TabIndex        =   1
         Top             =   360
         Width           =   2145
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   4
         Top             =   2130
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   5
         Top             =   2490
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   6
         Top             =   2850
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   7
         Top             =   3210
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   8
         Top             =   3570
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   9
         Top             =   3930
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   10
         Top             =   4290
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   11
         Top             =   4650
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   12
         Top             =   5010
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   13
         Top             =   5370
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Left            =   5880
         TabIndex        =   14
         Top             =   5730
         Width           =   1605
      End
      Begin VB.TextBox txtN 
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
         Index           =   11
         Left            =   5880
         TabIndex        =   15
         Top             =   6090
         Width           =   1605
      End
      Begin VB.TextBox txtAnual 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox txtPorc 
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
         Left            =   5040
         TabIndex        =   3
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   765
      End
      Begin VB.Image imgcta 
         Height          =   240
         Index           =   0
         Left            =   1200
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "ENERO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   62
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "FEBRERO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   61
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "MARZO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   60
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "ABRIL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   59
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "MAYO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   58
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "JUNIO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   57
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "JULIO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   56
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "AGOSTO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   55
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "SEPTIEMBRE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   54
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "OCTUBRE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   53
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "NOVIEMBRE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   52
         Top             =   5760
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "DICIEMBRE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   51
         Top             =   6120
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   2160
         TabIndex        =   50
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   2160
         TabIndex        =   49
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   2160
         TabIndex        =   48
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   2160
         TabIndex        =   47
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   2160
         TabIndex        =   46
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   2160
         TabIndex        =   45
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   2160
         TabIndex        =   44
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   2160
         TabIndex        =   43
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   2160
         TabIndex        =   42
         Top             =   5760
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   2160
         TabIndex        =   41
         Top             =   6120
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   4080
         TabIndex        =   40
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   4080
         TabIndex        =   39
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   4080
         TabIndex        =   38
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   4080
         TabIndex        =   37
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   4080
         TabIndex        =   36
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   4080
         TabIndex        =   35
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   4080
         TabIndex        =   34
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   4080
         TabIndex        =   33
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   4080
         TabIndex        =   32
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   4080
         TabIndex        =   31
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   4080
         TabIndex        =   30
         Top             =   5760
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   4080
         TabIndex        =   29
         Top             =   6120
         Width           =   1500
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7560
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   1920
         Y1              =   1680
         Y2              =   6480
      End
      Begin VB.Line Line3 
         X1              =   3840
         X2              =   3840
         Y1              =   1680
         Y2              =   6480
      End
      Begin VB.Line Line4 
         X1              =   5760
         X2              =   5760
         Y1              =   1680
         Y2              =   6480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "MES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "ANTERIOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2070
         TabIndex        =   27
         Top             =   1680
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "ACTUAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4020
         TabIndex        =   26
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   6120
         TabIndex        =   25
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   24
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   2160
         TabIndex        =   23
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   0
         X2              =   7980
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Importe ANUAL"
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
         TabIndex        =   22
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Incremento %"
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
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPresuGenerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Ejercicio As Long
Public Opcion As Byte
    '0 - Normal
    '1.- Generacion masiva
    

Public Modo As Byte
    ' 0 = insertar
    ' 1 = modificar


Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1

Dim cad As String
Dim Rs As Recordset
Dim i As Integer
Dim VV As Currency


Private AntiguoText As String 'Para comprobar si ha cambiado cosas o no

Private Sub Check1_Click()
    'Ha cambiado actual seiguiente
End Sub

Private Sub Check1_GotFocus()
    AntiguoText = Check1.Value
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Check1_LostFocus()
    If CStr(Check1.Value) <> AntiguoText Then PonerDatos2
End Sub

Private Sub HacerClick()
Dim Aux As String

On Error GoTo EINs
    'Comprobamos que los campos son correctos
    
    If txtCta(0).Text = "" Then
        MsgBox "Introduzca la cuenta ", vbExclamation
        Exit Sub
    End If
    
    For i = 0 To 11
        If txtN(i).Text <> "" Then
            If Not IsNumeric(txtN(i).Text) Then
                MsgBox "Los valores de los importes deben de ser numéricos", vbExclamation
                Exit Sub
            End If
        End If
    Next i
    'Llegados aqui esta todo bien. Luego haremos dos cosas
    i = Year(vParam.fechaini)
    If Check1.Value Then i = i + 1

    i = 0
    If Check1.Value Then i = i + 1

    If Check1.Value Then Ejercicio = Ejercicio + 1

    Dim F1 As Date
    Dim F2 As Date
    
    If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
        F1 = Format(Ejercicio, "0000") & "-" & Right("00" & Month(vParam.fechaini), 2) & "-01" 'DateAdd("yyyy", i, vParam.fechaini)
        F2 = Format(Ejercicio + 1, "0000") & "-" & Format(Month(vParam.fechafin), "00") & "-" & Format(Day(vParam.fechafin), "00") 'DateAdd("yyyy", i, vParam.fechafin)
    Else
        F1 = Format(Ejercicio, "0000") & "-" & Right("00" & Month(vParam.fechaini), 2) & "-01" 'DateAdd("yyyy", i, vParam.fechaini)
        F2 = Format(Ejercicio, "0000") & "-" & Format(Month(vParam.fechafin), "00") & "-" & Format(Day(vParam.fechafin), "00") 'DateAdd("yyyy", i, vParam.fechafin)
    End If

    If ChkEliminar.Value = 1 Then
'        Cad = "DELETE FROM presupuestos WHERE codmacta='" & txtCta(0).Text & "' AND anopresu = " & i
        cad = "DELETE FROM presupuestos WHERE codmacta='" & txtCta(0).Text & "' AND date(concat(anopresu,'-',mespresu,'-01')) between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
        Conn.Execute cad
    End If
    
    
    cad = "INSERT INTO presupuestos (codmacta, anopresu, mespresu, imppresu) VALUES ('"
    cad = cad & txtCta(0).Text & "',"
    
    Dim Anyo As Integer
    Dim Mes As Integer
    If Ejercicio = 0 Then
        If Me.Check1.Value Then
            Anyo = Year(vParam.fechaini) + 1
        Else
            Anyo = Year(vParam.fechaini)
        End If
    Else
        Anyo = Ejercicio 'Year(vParam.fechaini)
    End If
    Mes = Month(vParam.fechaini) - 1
    For i = 0 To 11
        Mes = Mes + 1
        If Mes > 12 Then
            Mes = 1
            Anyo = Anyo + 1
        End If
    
        If txtN(i).Text <> "" Then
            Aux = TransformaComasPuntos(ImporteFormateado(txtN(i).Text))
            Conn.Execute cad & Anyo & "," & Mes & "," & Aux & ")"
        End If
    Next i
    
    'Llegados aqui dejamos que vuelva a poner valores para otras cuentas
    If MsgBox("Datos generados.     ¿Salir?", vbQuestion + vbYesNoCancel) = vbYes Then
        i = Year(vParam.fechaini)
        If Check1.Value Then i = i + 1
        CadenaDesdeOtroForm = " presupuestos.codmacta = '" & txtCta(0).Text & "' and anopresu = " & i
        Unload Me
    Else
        txtAnual.Text = ""
        txtPorc.Text = ""
    End If

    Exit Sub
EINs:
    MuestraError Err.Number, "Insertando nuevos valores"
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkMasiva_Click()
    Opcion = Abs(Me.chkMasiva.Value)
    PonerFrames
 
    PonleFoco txtCta(Opcion)
    
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    If Opcion = 0 Then
        HacerClick
    Else
        If GeneracionMasiva Then Unload Me
        CadenaDesdeOtroForm = ""
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()

    Unload Me

End Sub


Private Sub PonerFrames()
    Me.FrameN.visible = (Opcion = 0)
    FrMasiva.visible = (Opcion = 1)
End Sub

Private Sub Form_Activate()
    ' si venimos modificando
    If Opcion = 0 Then
        If txtCta(0).Text <> "" Then PonerDatos2
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmppal.Icon

'    opcion = 0
'    Limpiar Me
    LimpiarLabels
    Label5.Caption = ""
    PonerFrames
    
    For i = 0 To imgCta.Count - 1
        imgCta(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    If Opcion = 1 Then
        Me.Caption = "Generar datos cuentas presupuestarias"
    Else
        Me.chkMasiva.visible = False
        Me.chkMasiva.Enabled = False
        Me.chkMasiva.Value = False
        If Modo = 0 Then
            Me.Caption = "Insertar datos cuentas presupuestarias"
        Else
            Me.Caption = "Modificar datos cuentas presupuestarias"
            Me.Check1.visible = False
            Me.Check1.Enabled = False
            Me.Check1.Value = False
        End If
        
    End If
    
    Me.ChkEliminar.Value = 1
    
    Dim J As Integer
    For i = 0 To 11
        J = Month(vParam.fechaini) + i
        If J > 12 Then J = J - 12
        
        Label2(i).Caption = NombreMes(J)
    Next i
    
End Sub

Private Function NombreMes(Mes As Integer) As String

    Select Case Mes
        Case 1
            NombreMes = "Enero"
        Case 2
            NombreMes = "Febrero"
        Case 3
            NombreMes = "Marzo"
        Case 4
            NombreMes = "Abril"
        Case 5
            NombreMes = "Mayo"
        Case 6
            NombreMes = "Junio"
        Case 7
            NombreMes = "Julio"
        Case 8
            NombreMes = "Agosto"
        Case 9
            NombreMes = "Septiembre"
        Case 10
            NombreMes = "Octubre"
        Case 11
            NombreMes = "Noviembre"
        Case 12
            NombreMes = "Diciembre"
    End Select

End Function


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(i).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescta(i).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgcta_Click(Index As Integer)
    i = Index
    AntiguoText = txtCta(Index).Text
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.Show vbModal
    Set frmC = Nothing
        
End Sub


Private Sub txtAnual_GotFocus()
    PonFoco txtAnual
    AntiguoText = txtAnual.Text
End Sub

Private Sub txtAnual_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAnual_LostFocus()
    txtAnual.Text = Trim(txtAnual.Text)
    If AntiguoText = txtAnual.Text Then Exit Sub
        
    cad = ""
    If txtAnual.Text <> "" Then
        If Not IsNumeric(txtAnual.Text) Then
            MsgBox "Campo numerico", vbExclamation
            txtAnual.Text = ""
            PonleFoco txtAnual
        Else
            If InStr(1, txtAnual.Text, ",") > 0 Then
                VV = ImporteFormateado(txtAnual.Text)
            Else
                VV = CCur(TransformaPuntosComas(txtAnual.Text))
            End If
            VV = Round2((VV / 12), 2)
            For i = 0 To 11
                txtN(i).Text = VV
            Next i
            cad = "OK"
        End If
        If txtAnual.Text <> "" Then txtPorc.Text = ""
    End If
    If cad = "" Then LimpiarCampos
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

'++
Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:   KEYBusqueda KeyAscii, 0 ' cuenta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgcta_Click (Indice)
End Sub

'++


Private Sub txtCta_LostFocus(Index As Integer)

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = AntiguoText Then Exit Sub
    If txtCta(Index).Text = "" Then
        txtDescta(Index).Text = ""
    Else
        CadenaDesdeOtroForm = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(CadenaDesdeOtroForm, cad) Then
                txtCta(Index).Text = CadenaDesdeOtroForm
                txtDescta(Index).Text = cad
        Else
            MsgBox cad, vbExclamation
            txtDescta(Index).Text = cad
        End If
        CadenaDesdeOtroForm = ""
    End If
    If Opcion = 0 Then PonerDatos2
End Sub
         


Private Sub txtInc_GotFocus()
    PonFoco txtInc
End Sub

Private Sub txtInc_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtInc_LostFocus()
    txtInc.Text = Trim(txtInc.Text)
    If txtInc.Text = "" Then Exit Sub
    
End Sub

Private Sub txtN_GotFocus(Index As Integer)
    PonFoco txtN(Index)
    
End Sub

Private Sub txtN_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtN_LostFocus(Index As Integer)
With txtN(Index)
    .Text = Trim(.Text)
    If .Text = "" Then Exit Sub
    If Not IsNumeric(.Text) Then
        MsgBox "Los importes deben de ser numéricos.", vbExclamation
        Exit Sub
    End If
    If InStr(1, .Text, ",") > 0 Then
        VV = ImporteFormateado(.Text)
    Else
        VV = CCur(TransformaPuntosComas(.Text))
    End If
    .Text = Format(VV, FormatoImporte)
End With
End Sub

Private Sub txtPorc_GotFocus()
    PonFoco txtPorc
    AntiguoText = txtPorc.Text
End Sub

Private Sub txtPorc_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
For i = 0 To 11
     Me.txtN(i).Text = ""
Next i

End Sub

Private Sub PonerDatosCuenta()

End Sub

Private Sub PonerDatos2()

    If txtCta(0).Text = "" Then
        LimpiarLabels
        If txtPorc.Text <> "" Then LimpiarCampos
        
    Else
        ' Solo si estamos insertando comprobamos que no estén dados de alta actual y siguiente
        If Modo = 0 Then
            If TieneEjercicio(txtCta(0), False) And TieneEjercicio(txtCta(0), True) Then
                MsgBox "Esta cuenta ya tiene presupuestos para ejercicio actual y el siguiente. " & vbCrLf & vbCrLf & "Ir por modificar.", vbExclamation
                Unload Me
            End If
        End If
        'Pondremos los datos en los campos
        PonerValores True
        PonerValores False
        
    End If

End Sub



Private Sub PonerValores(Anterior As Boolean)
Dim cad As String
Dim vFecIni As String
Dim vFecFin As String

On Error GoTo EPonerValoresAnteriores

    If Modo = 0 Then Ejercicio = Year(vParam.fechaini)

    If Modo = 1 Then ' modificar
        If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
            Label3(1).Caption = Format(Ejercicio - 1, "0000") & "-" & Format(Ejercicio, "0000")
            Label3(2).Caption = Format(Ejercicio, "0000") & "-" & Format(Ejercicio + 1, "0000")
        Else
            Label3(1).Caption = Format(Ejercicio - 1, "0000") '& "-" & Format(Ejercicio, "0000")
            Label3(2).Caption = Format(Ejercicio, "0000")  '& "-" & Format(Ejercicio + 1, "0000")
        End If
    Else
        Label3(1).Caption = "ANTERIOR"
        Label3(2).Caption = "ACTUAL"
    End If
    
    
    If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
        If Anterior Then
            vFecIni = Format(Ejercicio - 1, "0000") & "-" & Right("00" & Month(vParam.fechaini), 2) & "-01"
            vFecFin = Format(Ejercicio, "0000") & "-" & Right("00" & Month(vParam.fechafin), 2) & "-" & Format(Day(vParam.fechafin), "00")
        Else
            vFecIni = Format(Ejercicio, "0000") & "-" & Right("00" & Month(vParam.fechaini), 2) & "-01"
            vFecFin = Format(Ejercicio + 1, "0000") & "-" & Right("00" & Month(vParam.fechafin), 2) & "-" & Format(Day(vParam.fechafin), "00")
        End If
    Else
        If Anterior Then
            vFecIni = Format(Ejercicio - 1, "0000") & "-" & Right("00" & Month(vParam.fechaini), 2) & "-01"
            vFecFin = Format(Ejercicio - 1, "0000") & "-" & Right("00" & Month(vParam.fechafin), 2) & "-" & Format(Day(vParam.fechafin), "00")
        Else
            vFecIni = Format(Ejercicio, "0000") & "-" & Right("00" & Month(vParam.fechaini), 2) & "-01"
            vFecFin = Format(Ejercicio, "0000") & "-" & Right("00" & Month(vParam.fechafin), 2) & "-" & Format(Day(vParam.fechafin), "00")
        End If
    End If
    
    
    cad = "Select * from presupuestos where codmacta='" & txtCta(0).Text & "' AND"
    cad = cad & " date(concat(anopresu,'-',right(concat('00',mespresu),2),'-01')) between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        While Not Rs.EOF
            'El mes viene en el campo "mespresu"
            'Entonces para el label , k van del 12 al 23
            If Anterior Then
                If Rs!mespresu < Month(vParam.fechaini) Then
                    i = Rs!mespresu + 12 - Month(vParam.fechaini) + 12
                Else
                    i = Rs!mespresu - Month(vParam.fechaini) + 12
                End If
            Else
                If Rs!mespresu < Month(vParam.fechaini) Then
                    i = Rs!mespresu + 12 - Month(vParam.fechaini) + 24
                Else
                    i = Rs!mespresu - Month(vParam.fechaini) + 24
                End If
            End If
            Label2(i).Caption = Format(Rs!imppresu, FormatoImporte)
            
            'Ponemos el futuro valor(NUEVO)
            If txtPorc.Text <> "" Then
                'Porcentual
                VV = Round2(Rs!imppresu * CCur(txtPorc.Text), 2)
                VV = VV / 100
                VV = VV + Rs!imppresu
                If Anterior Then
                    i = i - 12
                Else
                    i = i - 24
                End If
                txtN(i).Text = Format(VV, FormatoImporte)
            End If
            
            'Sig
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
Exit Sub
EPonerValoresAnteriores:
    MuestraError Err.Number, "Poner Valores"
End Sub


Private Sub PonerValoresActuales()
On Error GoTo EPonerValoresActauales

    cad = "Select * from presupuestos where codmacta='" & txtCta(0).Text & "' AND"
    cad = cad & " anopresu = "
    i = Year(vParam.fechaini)
    'Ejercicio siguiente
    If Check1.Value = 1 Then i = i + 1
    
    cad = cad & i & ";"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        While Not Rs.EOF
            'El mes viene en el campo "mespresu"
            'Entonces para el label , k van del 24 al 35
            i = Rs!mespresu + 23
            Label2(i).Caption = Format(Rs!imppresu, FormatoImporte)
            
            'Sig
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
Exit Sub
EPonerValoresActauales:
    MuestraError Err.Number, "Poner datos anterior"
End Sub

Private Sub LimpiarLabels()
    For i = 12 To 35
        Label2(i).Caption = ""
    Next i
    
End Sub

Private Sub txtPorc_LostFocus()
    txtPorc.Text = Trim(txtPorc.Text)
    If txtPorc.Text = AntiguoText Then Exit Sub
    
    If txtPorc.Text <> "" Then txtAnual.Text = ""
    PonerDatos2
    
End Sub


Private Function GeneracionMasiva() As Boolean
Dim SQL As String
Dim Incremento As Currency

    On Error GoTo EGeneracionMasiva
    GeneracionMasiva = False
    
    'Generacion masvia de datos presupuestarios
    If txtInc.Text = "" Then
        MsgBox "Indique el incremento(%) a aplicar", vbExclamation
        Exit Function
    End If
    
    If MsgBox("Desea continuar con la generacion de datos presupuestarios?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    
    Incremento = ImporteSinFormato(txtInc.Text) / 100
    Set Rs = New ADODB.Recordset
    
    
    'Obtenedre el SQL de  las cuentas
    
    If ChkEliminar.Value = 1 Then
        FijarSQLTablaPresu False
        SQL = "DELETE FROM presupuestos WHERE " & cad
        Conn.Execute SQL
    End If
    
    FijarSQLTablaPresu True
    If optIncre(1).Value Then
        'Cojera los datos del presupuesto anterior
        SQL = "Select codmacta, anopresu anyo , mespresu mes ,imppresu debe,0 haber FROM presupuestos WHERE " & cad
    
    Else
        cad = Replace(cad, "anopresu", "year(fechaEnt)")
        cad = Replace(cad, "mespresu", "month(FechaEnt)")
        SQL = "select codmacta,year(fechaent) anyo, month(fechaent) mes, sum(coalesce(timported,0)) debe, sum(coalesce(timporteh,0))  haber from hlinapu where " & cad
        'Añado codmacta ultimo nivel
        SQL = SQL & "   AND codmacta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
        SQL = SQL & " group by 1,2,3"
    End If
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CadenaDesdeOtroForm = "INSERT INTO presupuestos (codmacta, anopresu, mespresu, imppresu) VALUES "
    i = 0
    SQL = ""

    While Not Rs.EOF
        i = i + 1
        VV = DBLet(Rs!Debe, "N") - DBLet(Rs!Haber, "N")
        VV = Round2((VV * Incremento) + VV, 2)  'importe
        If optEjer(0).Value Then
            cad = ",('" & Rs!codmacta & "'," & Rs!Anyo & "," & Rs!Mes & "," & TransformaComasPuntos(CStr(VV)) & ")"
        Else
            cad = ",('" & Rs!codmacta & "'," & Rs!Anyo + 1 & "," & Rs!Mes & "," & TransformaComasPuntos(CStr(VV)) & ")"
        End If
        SQL = SQL & cad
        If (i Mod 25) = 0 Then
             SQL = CadenaDesdeOtroForm & Mid(SQL, 2)  'QUITO la PRIMERa coma
            Conn.Execute SQL
            SQL = ""
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    If SQL <> "" Then
        SQL = CadenaDesdeOtroForm & Mid(SQL, 2)  'QUITO la PRIMERa coma
         Conn.Execute SQL
    End If
    MsgBox "Proceso finalizado", vbInformation
    GeneracionMasiva = True
EGeneracionMasiva:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Set Rs = Nothing
End Function


Private Sub FijarSQLTablaPresu(PeriodoAnterior As Boolean)
    cad = ""
    'Ejercicio
    If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        'Año natural
        i = Year(vParam.fechaini)
        If optEjer(1).Value Then i = i + 1
        If PeriodoAnterior Then i = i - 1
        cad = "anopresu = " & i
    Else
        i = Year(vParam.fechaini)
        If optEjer(1).Value Then i = i + 1
        If PeriodoAnterior Then i = i - 1
        
        cad = "(( anopresu = " & i & " AND mespresu >= " & Month(vParam.fechaini) & ") OR "
        cad = cad & " ( anopresu = " & i + 1 & " AND mespresu <= " & Month(vParam.fechafin) & ")) " ' AND ("
        
    End If
    
    If txtCta(1).Text <> "" Then cad = cad & " AND codmacta >= '" & txtCta(1).Text & "'"
    If txtCta(2).Text <> "" Then cad = cad & " AND codmacta <= '" & txtCta(2).Text & "'"
    
End Sub


Private Function TieneEjercicio(Cta As String, Actual As Boolean) As Boolean
Dim SQL As String

    SQL = "select count(*) from presupuestos where codmacta = " & DBSet(Cta, "T") & " and date(concat(anopresu,'-',right(concat('00',mespresu),2),'-1')) "
    If Actual Then
        SQL = SQL & " between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    Else
        SQL = SQL & " between " & DBSet(DateAdd("yyyy", 1, vParam.fechaini), "F") & " and " & DBSet(DateAdd("yyyy", 1, vParam.fechafin), "F")
    End If

    TieneEjercicio = (TotalRegistros(SQL) <> 0)



End Function


