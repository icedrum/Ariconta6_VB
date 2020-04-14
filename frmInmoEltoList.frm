VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInmoEltoList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
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
      Height          =   7365
      Left            =   7110
      TabIndex        =   21
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox Check2 
         Caption         =   "Agrupa sección"
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
         TabIndex        =   41
         Top             =   3360
         Width           =   3405
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
         ItemData        =   "frmInmoEltoList.frx":0000
         Left            =   210
         List            =   "frmInmoEltoList.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3960
         Width           =   1560
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
         Left            =   240
         TabIndex        =   31
         Top             =   360
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
            Top             =   390
            Value           =   1  'Checked
            Width           =   3405
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sólo resumen"
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
         TabIndex        =   2
         Top             =   3000
         Width           =   3405
      End
      Begin VB.CheckBox chkSaltoPag 
         Caption         =   "Salto de página por elemento"
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
         TabIndex        =   29
         Top             =   2640
         Width           =   3405
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2385
         Index           =   1
         Left            =   240
         TabIndex        =   46
         Top             =   4680
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   4207
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
         TabIndex        =   47
         Top             =   4440
         Width           =   1170
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3570
         Picture         =   "frmInmoEltoList.frx":0020
         ToolTipText     =   "Quitar al Debe"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   3930
         Picture         =   "frmInmoEltoList.frx":016A
         ToolTipText     =   "Puntear al Debe"
         Top             =   4440
         Width           =   240
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
         Left            =   240
         TabIndex        =   40
         Top             =   3720
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
      Height          =   4545
      Left            =   120
      TabIndex        =   18
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
         Index           =   2
         Left            =   1260
         TabIndex        =   5
         Tag             =   "imgConcepto"
         Top             =   3480
         Width           =   1215
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   3480
         Width           =   4185
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
         Index           =   3
         Left            =   1260
         TabIndex        =   6
         Tag             =   "imgConcepto"
         Top             =   3960
         Width           =   1215
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3960
         Width           =   4185
      End
      Begin VB.TextBox txtNElemento 
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2100
         Width           =   4185
      End
      Begin VB.TextBox txtNElemento 
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2520
         Width           =   4185
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   840
         Width           =   4185
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1260
         Width           =   4185
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
         Left            =   1260
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   840
         Width           =   1215
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
         Left            =   1260
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox txtElemento 
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
         Left            =   1260
         TabIndex        =   4
         Tag             =   "imgConcepto"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtElemento 
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
         Left            =   1260
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   2
         Left            =   960
         Top             =   3480
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
         Index           =   8
         Left            =   240
         TabIndex        =   45
         Top             =   3960
         Width           =   735
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   3
         Left            =   960
         Top             =   3960
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
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   43
         Top             =   3120
         Width           =   1110
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
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   3510
         Width           =   690
      End
      Begin VB.Image imgElemento 
         Height          =   255
         Index           =   1
         Left            =   960
         Top             =   2640
         Width           =   255
      End
      Begin VB.Image imgElemento 
         Height          =   255
         Index           =   0
         Left            =   960
         Top             =   2130
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
         Left            =   270
         TabIndex        =   28
         Top             =   2520
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
         Left            =   270
         TabIndex        =   27
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Elemento"
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
         Left            =   270
         TabIndex        =   26
         Top             =   1770
         Width           =   1110
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
         Left            =   240
         TabIndex        =   25
         Top             =   510
         Width           =   1080
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   1
         Left            =   960
         Top             =   1260
         Width           =   255
      End
      Begin VB.Image imgConcepto 
         Height          =   255
         Index           =   0
         Left            =   960
         Top             =   840
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
         Left            =   240
         TabIndex        =   20
         Top             =   1260
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
         Left            =   240
         TabIndex        =   19
         Top             =   840
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
      Left            =   10320
      TabIndex        =   9
      Top             =   7560
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
      Left            =   8760
      TabIndex        =   7
      Top             =   7560
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
      Left            =   150
      TabIndex        =   8
      Top             =   7500
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
      TabIndex        =   10
      Top             =   4680
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
         TabIndex        =   24
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   23
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   22
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInmoEltoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private WithEvents frmEle As frmInmoElto
Attribute frmEle.VB_VarHelpID = -1
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



Private Sub chkSaltoPag_Click()
    Me.Check1.Enabled = (chkSaltoPag.Value = 0)
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("inmovele", cadselect) Then Exit Sub
    
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
    Me.Caption = "Elementos de Inmovilizado"

    For I = 0 To 1
        Me.imgConcepto(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
        Me.imgElemento(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
     
    Me.ChkTipo(1).Value = 1
    Me.lblSubnvecnion.Caption = vParam.TextoInmoSubencionado
    Me.cboSubvencion.ListIndex = 0
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    cad = "select codubiin,nomubiin FROM inmovubicacion ORDER BY 1"
    CargaListviewCodigoDescripcion Me.ListView1(1), cad, True, 20

    
    
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    sql = CadenaSeleccion
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
    Set frmCon = New frmInmoConceptos
    frmCon.DatosADevolverBusqueda = True
    frmCon.Show vbModal
    Set frmCon = Nothing
    If sql <> "" Then
        Me.txtConcepto(Index).Text = RecuperaValor(sql, 1)
        Me.txtNConcepto(Index).Text = RecuperaValor(sql, 2)
    Else
        QuitarPulsacionMas Me.txtConcepto(Index)
    End If
    
    PonFoco Me.txtConcepto(Index)
    AbiertoOtroFormEnListado = False

End Sub

Private Sub frmEle_DatoSeleccionado(CadenaSeleccion As String)
    sql = CadenaSeleccion
End Sub

Private Sub imgElemento_Click(Index As Integer)
    
    sql = ""
    AbiertoOtroFormEnListado = True
    
    
    If Index < 2 Then
        
        Set frmEle = New frmInmoElto
        frmEle.DatosADevolverBusqueda = "0|1|"
        frmEle.Show vbModal
        Set frmEle = Nothing
    
    Else
        
        Set frmSec = frmInmoSeccion
        frmSec.DatosADevolverBusqueda = True
        frmSec.Show vbModal
        Set frmSec = Nothing
    
    End If
    
    If sql <> "" Then
        Me.txtElemento(Index).Text = RecuperaValor(sql, 1)
        Me.txtNElemento(Index).Text = RecuperaValor(sql, 2)
    Else
        QuitarPulsacionMas Me.txtElemento(Index)
    End If
    
    PonFoco Me.txtElemento(Index)
    AbiertoOtroFormEnListado = False

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
    Case "imgElemento"
        imgElemento_Click Indice
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

Private Sub txtElemento_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim sql As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtElemento(Index).Text = Trim(txtElemento(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'Tipos de elemento de inmovilizado
            txtNElemento(Index).Text = DevuelveDesdeBD("nominmov", "inmovele", "codinmov", txtElemento(Index), "N")
            If txtElemento(Index).Text <> "" Then txtElemento(Index).Text = Format(txtElemento(Index).Text, "000000")
    End Select
End Sub



Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    sql = "Select inmovele.codinmov Código,inmovele.nominmov Descripcion,inmovele.codmact1 Cuenta,inmovele.valoradq ValorAdquisicion ,inmovele.fechaadq FechaAdquisicion, inmovele_his.fechainm FechaAmortizacion, inmovele_his.imporinm ImporteAmortizacion, inmovele_his.porcinm PorcentajeAmortizacion "
    sql = sql & " FROM (inmovele inner join inmovele_his on inmovele.codinmov = inmovele_his.codinmov) "
    
    If cadselect <> "" Then sql = sql & " WHERE " & cadselect
    
    sql = sql & " ORDER BY 1,2,5" ' ordenado por codigo y por fecha de inmovilizado
        
    'LLamos a la funcion
    GeneraFicheroCSV sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
Dim CADENA As String

    vMostrarTree = False
    conSubRPT = False
        
    
    'indRPT = "0503-00"
    'If Check1.Value Then indRPT = "0503-01" ' resumido
    indRPT = IIf(Check2.Value = 1, "0503-02", "0503-00")
    If Check1.Value Then indRPT = IIf(Check2.Value = 1, "0503-03", "0503-01")
    
    
    
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "fichaelto.rpt"

    ' si no es resumido miramos si saltamos pagina o no
    If Check1.Value = 0 Then
        cadParam = cadParam & "pSaltoPag=" & chkSaltoPag.Value & "|"
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
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 55
        
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
    If Not PonerDesdeHasta("inmovele.codinmov", "INM", Me.txtElemento(0), Me.txtNElemento(0), Me.txtElemento(1), Me.txtNElemento(1), "pDHElemento=""") Then Exit Function
    If Not PonerDesdeHasta("inmovele.seccion", "CCI", Me.txtConcepto(2), Me.txtConcepto(2), Me.txtConcepto(3), Me.txtConcepto(3), "pDHSecci=""") Then Exit Function
    
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


Private Sub txtElemento_GotFocus(Index As Integer)
    ConseguirFoco txtElemento(Index), 3
End Sub

Private Sub txtElemento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtElemento(Index).Tag, Index
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

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
