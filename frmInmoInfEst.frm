VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInmoInfEst 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
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
      Height          =   7065
      Left            =   7110
      TabIndex        =   27
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox CheckSeccion 
         Caption         =   "Agrupar por sección"
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
         Left            =   240
         TabIndex        =   46
         Top             =   3600
         Width           =   3495
      End
      Begin VB.ComboBox cboSubvencion 
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
         ItemData        =   "frmInmoInfEst.frx":0000
         Left            =   240
         List            =   "frmInmoInfEst.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3120
         Width           =   1440
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Elementos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2025
         Left            =   210
         TabIndex        =   32
         Top             =   660
         Width           =   4035
         Begin VB.CheckBox ChkTipo 
            Caption         =   "Totalmente amortizado"
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
            Left            =   180
            TabIndex        =   11
            Top             =   1590
            Width           =   3405
         End
         Begin VB.CheckBox ChkTipo 
            Caption         =   "Vendido"
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
            TabIndex        =   9
            Top             =   810
            Width           =   3405
         End
         Begin VB.CheckBox ChkTipo 
            Caption         =   "Baja"
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
            Left            =   180
            TabIndex        =   10
            Top             =   1200
            Width           =   3405
         End
         Begin VB.CheckBox ChkTipo 
            Caption         =   "Activo"
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
            Left            =   180
            TabIndex        =   8
            Top             =   390
            Value           =   1  'Checked
            Width           =   3405
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3870
         TabIndex        =   40
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
         Height          =   2145
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   4200
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
         Left            =   3930
         Picture         =   "frmInmoInfEst.frx":0020
         ToolTipText     =   "Puntear al Debe"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3570
         Picture         =   "frmInmoInfEst.frx":016A
         ToolTipText     =   "Quitar al Debe"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Ubicaciones"
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
         Index           =   12
         Left            =   240
         TabIndex        =   52
         Top             =   3960
         Width           =   1170
      End
      Begin VB.Label lblSubnvecnion 
         Caption         =   "X"
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
         Left            =   270
         TabIndex        =   45
         Top             =   2830
         Width           =   3180
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
      Height          =   4395
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtConcepto 
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
         Left            =   1110
         TabIndex        =   7
         Tag             =   "imgConcepto"
         Top             =   3840
         Width           =   1305
      End
      Begin VB.TextBox txtNConcepto 
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
         Index           =   3
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   3840
         Width           =   4095
      End
      Begin VB.TextBox txtConcepto 
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
         Left            =   1110
         TabIndex        =   6
         Tag             =   "imgConcepto"
         Top             =   3330
         Width           =   1305
      End
      Begin VB.TextBox txtNConcepto 
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
         Index           =   2
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   3330
         Width           =   4095
      End
      Begin VB.TextBox txtNConcepto 
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtNConcepto 
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1380
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   2520
         TabIndex        =   36
         Top             =   1800
         Width           =   4200
         Begin VB.TextBox txtNCCoste 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   2355
         End
         Begin VB.TextBox txtNCCoste 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   690
            Width           =   2355
         End
         Begin VB.TextBox txtCCoste 
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
            Left            =   840
            TabIndex        =   4
            Tag             =   "imgConcepto"
            Top             =   270
            Width           =   825
         End
         Begin VB.TextBox txtCCoste 
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
            Left            =   840
            TabIndex        =   5
            Tag             =   "imgConcepto"
            Top             =   690
            Width           =   825
         End
         Begin VB.Image imgCCoste 
            Height          =   255
            Index           =   0
            Left            =   600
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Centro de Coste"
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
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   2310
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
            Left            =   0
            TabIndex        =   38
            Top             =   360
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
            Index           =   2
            Left            =   0
            TabIndex        =   37
            Top             =   720
            Width           =   615
         End
         Begin VB.Image imgCCoste 
            Height          =   255
            Index           =   1
            Left            =   600
            Top             =   720
            Width           =   255
         End
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   2070
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
         Index           =   1
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   2490
         Width           =   1305
      End
      Begin VB.TextBox txtConcepto 
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
         Left            =   1080
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   960
         Width           =   1305
      End
      Begin VB.TextBox txtConcepto 
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
         Left            =   1080
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1380
         Width           =   1305
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   2
         Left            =   810
         Top             =   3360
         Width           =   255
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
         Index           =   11
         Left            =   120
         TabIndex        =   51
         Top             =   3360
         Width           =   780
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
         Index           =   10
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Width           =   735
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   3
         Left            =   810
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Seccion"
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
         Index           =   9
         Left            =   120
         TabIndex        =   48
         Top             =   3000
         Width           =   1110
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
         Index           =   8
         Left            =   120
         TabIndex        =   35
         Top             =   1770
         Width           =   960
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
         Left            =   120
         TabIndex        =   34
         Top             =   2130
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
         Left            =   120
         TabIndex        =   33
         Top             =   2490
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   810
         Picture         =   "frmInmoInfEst.frx":02B4
         Top             =   2100
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   810
         Picture         =   "frmInmoInfEst.frx":033F
         Top             =   2490
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto"
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
         Index           =   6
         Left            =   90
         TabIndex        =   31
         Top             =   630
         Width           =   1110
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   1
         Left            =   780
         Top             =   1380
         Width           =   255
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   0
         Left            =   780
         Top             =   960
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
         Left            =   90
         TabIndex        =   26
         Top             =   1380
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
         Left            =   90
         TabIndex        =   25
         Top             =   960
         Width           =   780
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
      Left            =   10290
      TabIndex        =   15
      Top             =   7290
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
      TabIndex        =   13
      Top             =   7290
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
      TabIndex        =   14
      Top             =   7230
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
      TabIndex        =   16
      Top             =   4410
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
         TabIndex        =   30
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   29
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   28
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInmoInfEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 505

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

Public NumAsien As String
Public NumDiari As String
Public FechaEnt As String


Private WithEvents frmCon As frmInmoConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCCentroCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmSec As frmInmoSeccion
Attribute frmSec.VB_VarHelpID = -1


Private sql As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer


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
Dim tabla As String

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    'si es entre fechas enlazamos con el hco de amortizaciones
    tabla = "inmovele"
    If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        tabla = "inmovele INNER JOIN inmovele_his ON inmovele.codinmov = inmovele_his.codinmov"
        
        If Not CargarTablaTemporal(tabla, cadselect) Then Exit Sub
        
        tabla = "tmpconextcab"
        
        cadselect = "tmpconextcab.codusu = " & vUsu.Codigo
        
    End If
    
    
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
    
End Sub

Private Function CargarTablaTemporal(tabla As String, cadselect As String) As Boolean
Dim sql As String

    On Error GoTo eCargarTablaTemporal
    
    CargarTablaTemporal = False

    sql = "delete from tmpconextcab where codusu = " & vUsu.Codigo
    Conn.Execute sql
    
    sql = "insert into tmpconextcab (codusu, cta, acumantD, acumantH) select " & vUsu.Codigo & ", inmovele.codinmov, inmovele.codinmov, sum(imporinm) from  " & tabla
    sql = sql & " where " & cadselect
    sql = sql & " group by 1, 2"
    sql = sql & " order by 1,2 "
    
    Conn.Execute sql
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function


Private Sub cmdCancelar_Click()
    Unload Me
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
    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Estadística de Inmovilizado"

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With


    For I = 0 To 1
        Me.imgConcepto(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgCCoste(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
     
    Me.ChkTipo(1).Value = 1
    Me.lblSubnvecnion.Caption = vParam.TextoInmoSubencionado
    Me.cboSubvencion.ListIndex = 0
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    Frame2.visible = vParam.autocoste
    Frame2.Enabled = vParam.autocoste
    
    cad = "select codubiin,nomubiin FROM inmovubicacion ORDER BY 1"
    CargaListviewCodigoDescripcion Me.ListView1(1), cad, True, 33
    
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    sql = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    sql = CadenaSeleccion
End Sub

Private Sub imgCheck_Click(Index As Integer)
    ListviewSelecDeselec Me.ListView1(1), Index = 1
End Sub

Private Sub imgConcepto_Click(Index As Integer)
    
    sql = ""
    AbiertoOtroFormEnListado = True
    
    If Index < 2 Then
    
        
        Set frmCon = New frmInmoConceptos
        frmCon.DatosADevolverBusqueda = True
        frmCon.Show vbModal
        Set frmCon = Nothing
    
    Else
        Set frmSec = frmInmoSeccion
        frmSec.DatosADevolverBusqueda = True
        frmSec.Show vbModal
        Set frmSec = Nothing
    
    End If
    If sql <> "" Then
        Me.txtConcepto(Index).Text = RecuperaValor(sql, 1)
        Me.txtNConcepto(Index).Text = RecuperaValor(sql, 2)
    Else
        QuitarPulsacionMas Me.txtConcepto(Index)
    End If
    
    PonFoco Me.txtConcepto(Index)
    AbiertoOtroFormEnListado = False

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    sql = CadenaSeleccion
End Sub

Private Sub ImgCCoste_Click(Index As Integer)
    
    sql = ""
    AbiertoOtroFormEnListado = True
    
    Set frmCC = New frmCCCentroCoste
    frmCC.DatosADevolverBusqueda = "0|1|"
    frmCC.Show vbModal
    Set frmCC = Nothing
    
    
    If sql <> "" Then
        Me.txtCCoste(Index).Text = RecuperaValor(sql, 1)
        Me.txtNCCoste(Index).Text = RecuperaValor(sql, 2)
    Else
        QuitarPulsacionMas Me.txtCCoste(Index)
    End If
    
    PonFoco Me.txtCCoste(Index)
    AbiertoOtroFormEnListado = False

End Sub


Private Sub imgFec_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
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
    Case "imgCCoste"
        ImgCCoste_Click Indice
    End Select
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'Tipos de concepto de inmovilizado
            txtNConcepto(Index).Text = DevuelveDesdeBD("nomconam", "inmovcon", "codconam", txtConcepto(Index), "N")
            If txtConcepto(Index).Text <> "" Then txtConcepto(Index).Text = Format(txtConcepto(Index).Text, "0000")
        Case 2, 3
            txtNConcepto(Index).Text = DevuelveDesdeBD("nomsecin", "inmovseccion", "codsecin", txtConcepto(Index), "N")
            If txtConcepto(Index).Text <> "" Then txtConcepto(Index).Text = Format(txtConcepto(Index).Text, "0000")
    End Select

End Sub



Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
   ' KEYpressGnral KeyAscii
End Sub

Private Sub txtCCoste_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim sql As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCCoste(Index).Text = Trim(txtCCoste(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'Centros de coste de inmovilizado
            txtNCCoste(Index).Text = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtCCoste(Index), "T")
    End Select
End Sub



Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        sql = "Select inmovele.conconam concepto, inmovcon.nomconam DescConcepto,inmovele.codinmov Código,inmovele.nominmov Descripcion,inmovele.codmact1 Cuenta,"
        sql = sql & " inmovele.fechaadq FechaAdquisicion,inmovele.valoradq ValorAdquisicion,inmovele.amortacu AmorAcumulada , tmpconextcab.acumantH AmortPeridodo , inmovele.valoradq -inmovele.amortacu Pendiente "
        sql = sql & " ,inmovele.impventa Venta, inmovele.fecventa FechaVenta"
        sql = sql & " FROM   (inmovele inmovele INNER JOIN inmovcon inmovcon ON inmovele.conconam=inmovcon.codconam)"
        sql = sql & " INNER JOIN tmpconextcab tmpconextcab ON inmovele.codinmov=tmpconextcab.acumantD"
    
        If cadselect <> "" Then sql = sql & " WHERE " & cadselect
        
        sql = sql & " ORDER BY 1,2" ' ordenado por codigo y por fecha de inmovilizado
    
    Else
        sql = "Select inmovele.codinmov Código,inmovele.nominmov Descripcion,inmovele.codmact1 Cuenta ,inmovele.fechaadq FechaAdquisicion,inmovele.valoradq ValorAdquisicion, inmovele.amortacu Amortizado, coalesce(inmovele.valoradq,0) - coalesce(inmovele.amortacu,0) Pendiente, inmovele.fecventa FechaVta, inmovele.impventa ImpVenta "
        sql = sql & " FROM inmovele "
        
        If cadselect <> "" Then sql = sql & " WHERE " & cadselect
        
        sql = sql & " ORDER BY 1,2,4" ' ordenado por codigo y por fecha de adquisicion
    End If
    
        
    'LLamos a la funcion
    GeneraFicheroCSV sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
Dim CADENA As String

    vMostrarTree = False
    conSubRPT = False
        
    
    If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        'indRPT = "0505-01"
        indRPT = IIf(Me.CheckSeccion.Value = 0, "0505-01", "0505-03")
        cadFormula = "{tmpconextcab.codusu} = " & vUsu.Codigo
    Else
        'indRPT = "0505-00"
        indRPT = IIf(Me.CheckSeccion.Value = 0, "0505-00", "0505-02")
    End If
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "fichaelto.rpt"

    If vParam.autocoste Then
        cadParam = cadParam & "pAnalitica=1|"
        numParam = numParam + 1
    End If


    ' tipos de elementos
    CADENA = ""
    If Me.ChkTipo(1).Value = 1 Then CADENA = CADENA & "Activo,"
    If Me.ChkTipo(2).Value = 1 Then CADENA = CADENA & "Baja,"
    If Me.ChkTipo(3).Value = 1 Then CADENA = CADENA & "Vendido,"
    If Me.ChkTipo(4).Value = 1 Then CADENA = CADENA & "Amort.,"
    
    
    If CADENA <> "" Then
        CADENA = Mid(CADENA, 1, Len(CADENA) - 1)
        If CADENA <> "" Then CADENA = "Situacion: " & CADENA
    End If
    If Me.cboSubvencion.ListIndex > 0 Then
        If CADENA <> "" Then CADENA = CADENA & "       "
        CADENA = CADENA & "       " & Me.lblSubnvecnion.Caption & ": " & cboSubvencion.Text
    End If
    
    
    
    cadParam = cadParam & "pTipo=""" & CADENA & """|"
    numParam = numParam + 1


    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 62
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim Situacion As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("inmovele.conconam", "COI", Me.txtConcepto(0), Me.txtNConcepto(0), Me.txtConcepto(1), Me.txtNConcepto(1), "pDHConcepto=""") Then Exit Function
    If Not PonerDesdeHasta("inmovele.codccost", "CCO", Me.txtCCoste(0), Me.txtNCCoste(0), Me.txtCCoste(1), Me.txtNCCoste(1), "pDHCCoste=""") Then Exit Function
    If Not PonerDesdeHasta("inmovele.seccion", "CCI", Me.txtConcepto(2), Me.txtConcepto(2), Me.txtConcepto(3), Me.txtConcepto(3), "pDHSecci=""") Then Exit Function
    
    If txtFecha(0).Text <> "" Or txtFecha(1).Text <> "" Then
        If Not PonerDesdeHasta("inmovele_his.fechainm", "FEC", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    End If
    
    
    Situacion = ""
    For I = 1 To 4
        If ChkTipo(I).Value Then Situacion = Situacion & I & ","
    Next I
    
    'quitamos la ultima coma
    If Situacion <> "" Then
        Situacion = Mid(Situacion, 1, Len(Situacion) - 1)
        If Not AnyadirAFormula(cadFormula, "{inmovele.situacio} in [" & Situacion & "]") Then Exit Function
        If Not AnyadirAFormula(cadselect, "inmovele.situacio in (" & Situacion & ")") Then Exit Function
    End If
            
            
    If Me.cboSubvencion.ListIndex > 0 Then
        RC = IIf(cboSubvencion.ListIndex = 1, 1, 0)
        If Not AnyadirAFormula(cadFormula, "{inmovele.subvencionado} = " & RC) Then Exit Function
        If Not AnyadirAFormula(cadselect, "inmovele.subvencionado = " & RC) Then Exit Function
    End If
                    
            
    'Ubicacion
    Situacion = ""
    RC = ""
    For I = 1 To Me.ListView1(1).ListItems.Count
        If ListView1(1).ListItems(I).Checked Then
            Situacion = Situacion & ", " & I
            RC = RC & "X"
        End If
    Next I
    If RC = "" Then
        MsgBox "Seleccione alguna ubicacion", vbExclamation
        Exit Function
    End If
    
    If Len(RC) <> Me.ListView1(1).ListItems.Count Then
        'NO ha selccionado todas
        Situacion = Mid(Situacion, 2)
        If Not AnyadirAFormula(cadFormula, "{inmovele.ubicacion} in [" & Situacion & "]") Then Exit Function
        If Not AnyadirAFormula(cadselect, "inmovele.ubicacion in (" & Situacion & ")") Then Exit Function
    
    End If
            
            
            
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
            
    MontaSQL = True
End Function


Private Sub txtCCoste_GotFocus(Index As Integer)
    ConseguirFoco txtCCoste(Index), 3
End Sub

Private Sub txtCCoste_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtCCoste(Index).Tag, Index
    End If
End Sub

Private Function DatosOK() As Boolean
Dim I As Integer
Dim CADENA As String

    DatosOK = False
    
    CADENA = ""
    For I = 1 To 4
        If ChkTipo(I).Value Then CADENA = CADENA & I
    Next I
    
    If CADENA = "" Then
        MsgBox "Debe de introducir algún tipo de elemento. Revise.", vbExclamation
        DatosOK = False
        Exit Function
    End If
    
    DatosOK = True


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
    End If
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
