VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturasProListado 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ordenaci�n"
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
      Left            =   7140
      TabIndex        =   53
      Top             =   6750
      Width           =   4455
      Begin VB.OptionButton optVarios 
         Caption         =   "Fecha liquidaci�n"
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
         Left            =   390
         TabIndex        =   18
         Top             =   1791
         Width           =   2295
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "N� Registro"
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
         Left            =   390
         TabIndex        =   15
         Top             =   600
         Width           =   1755
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Fecha Recepci�n"
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
         Left            =   390
         TabIndex        =   17
         Top             =   1394
         Width           =   2295
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Proveedor"
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
         Left            =   390
         TabIndex        =   19
         Top             =   2190
         Width           =   1725
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Fecha Factura"
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
         Left            =   390
         TabIndex        =   16
         Top             =   997
         Width           =   1755
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Resumen"
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
         Left            =   2460
         TabIndex        =   20
         Top             =   2220
         Width           =   1335
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   59
         Top             =   240
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
      Height          =   6705
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNCuentas 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   5160
         Width           =   4215
      End
      Begin VB.TextBox txtNCuentas 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   4740
         Width           =   4215
      End
      Begin VB.TextBox txtNSerie 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1380
         Width           =   4665
      End
      Begin VB.TextBox txtNSerie 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   960
         Width           =   4665
      End
      Begin VB.TextBox txtNIF 
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
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "imgConcepto"
         Top             =   5700
         Width           =   1455
      End
      Begin VB.TextBox txtCuentas 
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
         Left            =   1230
         TabIndex        =   7
         Tag             =   "imgCuenta"
         Top             =   5160
         Width           =   1275
      End
      Begin VB.TextBox txtCuentas 
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
         Left            =   1230
         TabIndex        =   6
         Tag             =   "imgCuenta"
         Top             =   4740
         Width           =   1275
      End
      Begin VB.TextBox txtSerie 
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
         Tag             =   "imgSerie"
         Top             =   960
         Width           =   765
      End
      Begin VB.TextBox txtSerie 
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
         Tag             =   "imgSerie"
         Top             =   1380
         Width           =   765
      End
      Begin VB.TextBox txtNumFactu 
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
         TabIndex        =   3
         Tag             =   "N� Registro|N|S|||factpro|numregis|0000000|S|"
         Top             =   2640
         Width           =   1275
      End
      Begin VB.TextBox txtNumFactu 
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
         TabIndex        =   2
         Tag             =   "N� Registro|N|S|||factpro|numregis|0000000|S|"
         Top             =   2250
         Width           =   1275
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
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "imgFecha"
         Top             =   3930
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
         Index           =   0
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "imgFecha"
         Top             =   3510
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "N.I.F."
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
         Index           =   14
         Left            =   270
         TabIndex        =   51
         Top             =   5730
         Width           =   960
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
         Index           =   11
         Left            =   240
         TabIndex        =   47
         Top             =   4380
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
         Index           =   10
         Left            =   240
         TabIndex        =   46
         Top             =   4770
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
         Index           =   9
         Left            =   240
         TabIndex        =   45
         Top             =   5130
         Width           =   615
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   5190
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   4740
         Width           =   255
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmFacturasProListado.frx":0000
         Top             =   3930
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmFacturasProListado.frx":008B
         Top             =   3540
         Width           =   240
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
         TabIndex        =   44
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
         Index           =   3
         Left            =   270
         TabIndex        =   43
         Top             =   2280
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
         Left            =   270
         TabIndex        =   42
         Top             =   3930
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
         Left            =   270
         TabIndex        =   41
         Top             =   3570
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Registro"
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
         Index           =   7
         Left            =   270
         TabIndex        =   40
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Recepci�n"
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
         Left            =   270
         TabIndex        =   39
         Top             =   3210
         Width           =   2670
      End
      Begin VB.Label Label3 
         Caption         =   "Serie"
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
         Left            =   240
         TabIndex        =   38
         Top             =   630
         Width           =   960
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Index           =   1
         Left            =   960
         Top             =   1380
         Width           =   255
      End
      Begin VB.Image imgSerie 
         Height          =   255
         Index           =   0
         Left            =   960
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
         Left            =   240
         TabIndex        =   34
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
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   780
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
      Height          =   6705
      Left            =   7140
      TabIndex        =   48
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmFacturasProListado.frx":0116
         Left            =   120
         List            =   "frmFacturasProListado.frx":0118
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   4110
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sustituye Factura por Fecha"
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
         Left            =   180
         TabIndex        =   12
         Top             =   1200
         Width           =   3915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sustituye cuenta por NIF"
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
         Left            =   180
         TabIndex        =   11
         Top             =   840
         Width           =   3075
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
         Left            =   1080
         TabIndex        =   9
         Tag             =   "imgFecha"
         Text            =   "99/99/9999"
         Top             =   390
         Width           =   1245
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
         Left            =   3720
         TabIndex        =   10
         Tag             =   "imgConcepto"
         Top             =   390
         Width           =   525
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2145
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   2640
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3784
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   1320
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   5280
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   2328
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
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
         Caption         =   "Retenci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   120
         TabIndex        =   61
         Top             =   5040
         Width           =   930
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   3810
         Picture         =   "frmFacturasProListado.frx":011A
         ToolTipText     =   "Puntear al Debe"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   3450
         Picture         =   "frmFacturasProListado.frx":0264
         ToolTipText     =   "Quitar al Debe"
         Top             =   5040
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Factura"
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
         TabIndex        =   58
         Top             =   1560
         Width           =   1380
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   3480
         Picture         =   "frmFacturasProListado.frx":03AE
         ToolTipText     =   "Quitar al Debe"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   3840
         Picture         =   "frmFacturasProListado.frx":04F8
         ToolTipText     =   "Puntear al Debe"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "IVA"
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
         Left            =   150
         TabIndex        =   52
         Top             =   2400
         Width           =   870
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   2
         Left            =   720
         Picture         =   "frmFacturasProListado.frx":0642
         Top             =   450
         Width           =   240
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
         Index           =   13
         Left            =   150
         TabIndex        =   50
         Top             =   473
         Width           =   690
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
         Index           =   12
         Left            =   2640
         TabIndex        =   49
         Top             =   443
         Width           =   870
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
      Left            =   10410
      TabIndex        =   23
      Top             =   9540
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
      Left            =   8850
      TabIndex        =   21
      Top             =   9540
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
      TabIndex        =   22
      Top             =   9480
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
      TabIndex        =   24
      Top             =   6750
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
         TabIndex        =   37
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   36
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   35
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmFacturasProListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 402


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
Public Legalizacion As String

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String
Dim PrimeraVez As Boolean

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



Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAccion_Click(Index As Integer)
Dim EsElListadoRetenciones As Boolean
Dim B As Boolean

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    'Si no ex exportacion, y la opcion seleccionada es PROVEEDOR NO sustituye el NIF por la cuenta
    If optTipoSal(0).Value Then
        If optVarios(1).Value Then
            If Me.Check1(0).Value = 1 Then MsgBox "Agrupando por proveedor siempre ser� por cuenta contable", vbExclamation
        End If
    End If
    
    
    EsElListadoRetenciones = False
    If ListView1(0).ListItems.Count > 0 Then
        If ListView1(0).ListItems(1).Checked = False Then EsElListadoRetenciones = True 'quiere el listado retenciones
    End If
    
    If EsElListadoRetenciones Then
        tabla = "DELETE from tmppresu1 where codusu =" & vUsu.Codigo
        Conn.Execute tabla
    
        tabla = "insert into tmppresu1(codusu,codigo,titulo)  select " & vUsu.Codigo & ",codigo,descripcion from usuarios.wtiporeten"
        Conn.Execute tabla
    End If
    
    InicializarVbles True
    
    tabla = "factpro inner join factpro_totales on factpro.numserie = factpro_totales.numserie and factpro.numregis = factpro_totales.numregis and factpro.anofactu = factpro_totales.anofactu "
    tabla = "(" & tabla & ") inner join contadores on factpro.numserie = contadores.tiporegi "
    tabla = "(" & tabla & ") inner join formapago on factpro.codforpa = formapago.codforpa "
    tabla = "(" & tabla & ") inner join tipofpago on formapago.tipforpa = tipofpago.tipoformapago "
    tabla = "(" & tabla & ") inner join usuarios.wtipopera aa on factpro.codopera = aa.codigo "
    
    B = True
    Screen.MousePointer = vbHourglass
    If Not MontaSQL Then B = False
    If B Then
        If Not HayRegParaInforme(tabla, cadselect) Then B = False
    End If
    If Not B Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV EsElListadoRetenciones
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
    
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal EsElListadoRetenciones
    End If
    If Legalizacion <> "" Then
        CadenaDesdeOtroForm = "OK"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Legalizacion <> "" Then
            optTipoSal(2).Value = True
            optVarios(2).Value = True
            cmdAccion_Click (1)
        End If
    End If
End Sub




Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Listado de Facturas de Proveedores"

    For i = 0 To 1
        Me.imgSerie(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
     
    For i = 0 To 1
        Me.imgCuentas(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
     
    CargarListView 1
    CargarListView 0
    Me.optVarios(2).Value = True
    
    
    optVarios(4).visible = vParam.IvaEnFechaPago
    optVarios_Click (0)
    
    txtFecha(0).Text = vParam.fechaini
    txtFecha(1).Text = vParam.fechafin
    If Not vParam.FecEjerAct Then
        txtFecha(1).Text = Format(DateAdd("yyyy", 1, vParam.fechafin), "dd/mm/yyyy")
    End If
    
    
    txtFecha(2).Text = Format(Now, "dd/mm/yyyy")
    txtPag2(0).Text = "1"
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    If Legalizacion <> "" Then
        txtFecha(2).Text = RecuperaValor(Legalizacion, 1)
        txtFecha(0).Text = RecuperaValor(Legalizacion, 2)
        txtFecha(1).Text = RecuperaValor(Legalizacion, 3)
    End If
    

    
    CargarCombo
    'todos
    Combo1(0).ListIndex = 0
   
    
End Sub



Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtSerie(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        txtNSerie(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim J As Integer


    Screen.MousePointer = vbHourglass
    
    J = 0
    If Index < 2 Then J = 1
    
    IndCodigo = 0
    If (Index Mod 2) = 1 Then IndCodigo = 1
    
    For i = 1 To ListView1(J).ListItems.Count
        ListView1(J).ListItems(i).Checked = IndCodigo = 1
    Next i
    
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub imgSerie_Click(Index As Integer)
    IndCodigo = Index
    
    Set frmConta = New frmBasico
    AyudaContadores frmConta, txtSerie(Index), "tiporegi REGEXP '^[0-9]+$' <> 0 and tiporegi > 0 "
    Set frmConta = Nothing
    
    PonFoco Me.txtSerie(Index)
End Sub

Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2
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

Private Sub imgCuentas_Click(Index As Integer)
    Sql = ""
    AbiertoOtroFormEnListado = True
    Set frmCtas = New frmColCtas
    frmCtas.DatosADevolverBusqueda = True
    frmCtas.Show vbModal
    Set frmCtas = Nothing
    If Sql <> "" Then
        Me.txtCuentas(Index).Text = RecuperaValor(Sql, 1)
        Me.txtNCuentas(Index).Text = RecuperaValor(Sql, 2)
    Else
        QuitarPulsacionMas Me.txtCuentas(Index)
    End If
    
    PonFoco Me.txtCuentas(Index)
    AbiertoOtroFormEnListado = False
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_Click(Index As Integer)
    Check1(1).Enabled = optVarios(1).Value
    If Not Check1(1).Enabled Then Check1(1).Value = 0
    
    If optVarios(0) Then
        Label3(8).Caption = "Fecha Factura"
    ElseIf optVarios(4) Then
        Label3(8).Caption = "Fecha liquidaci�n"
    Else
        Label3(8).Caption = "Fecha Recepci�n"
    End If
    
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
  
  
    Check1(1).Enabled = optVarios(1).Value
    
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

Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub

Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtCuentas(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
   ' KEYpressGnral KeyAscii
End Sub

Private Sub txtCuentas_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    
    
    If txtCuentas(Index).Text = "" Then
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCuentas(Index).Text) Then
        If InStr(1, txtCuentas(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser num�rica: " & txtCuentas(Index).Text, vbExclamation
        txtCuentas(Index).Text = ""
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, Sql)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
            Else
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = Sql
                If B = 1 Then
                    txtNCuentas(Index).Tag = ""
                Else
                    txtNCuentas(Index).Tag = Sql
                End If
                Hasta = -1
                If Index = 6 Then
                    Hasta = 7
                Else
                    If Index = 0 Then
                        Hasta = 1
                    Else
                        If Index = 5 Then
                            Hasta = 4
                        Else
                            If Index = 23 Then Hasta = 24
                        End If
                    End If
                    
                End If
                    
                If Hasta >= 0 Then
                    txtCuentas(Hasta).Text = txtCuentas(Index).Text
                    txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
                End If
            End If
    
    
    End Select

End Sub

Private Sub txtNIF_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtPag2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtSerie_GotFocus(Index As Integer)
    ConseguirFoco txtSerie(Index), 3
End Sub

Private Sub txtNumFactu_GotFocus(Index As Integer)
    ConseguirFoco txtNumFactu(Index), 3
End Sub

Private Sub txtNumFactu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtNumFactu(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgSerie"
        imgSerie_Click Indice
    Case "imgFecha"
        imgFec_Click Indice
    Case "imgCuentas"
        imgCuentas_Click Indice
    End Select
End Sub

Private Sub txtNumFactu_KeyPress(Index As Integer, KeyAscii As Integer)
   ' KEYpressGnral KeyAscii
End Sub

Private Sub txtNumFactu_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtNumFactu(Index).Text = Trim(txtNumFactu(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada

    If txtNumFactu(Index).Text = "" Then
        Exit Sub
    End If

    PonerFormatoEntero txtNumFactu(Index)


'    If Not IsNumeric(txtNumFactu(Index).Text) Then
'        If InStr(1, txtNumFactu(Index).Text, "+") = 0 Then MsgBox "El n�mero de factura debe ser num�rico: " & txtSerie(Index).Text, vbExclamation
'        txtNumFactu(Index).Text = ""
'        Exit Sub
'    End If
End Sub



Private Sub txtSerie_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtSerie(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If

End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
   ' KEYpressGnral KeyAscii
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim Sql As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

    txtSerie(Index).Text = UCase(Trim(txtSerie(Index).Text))
    
    If txtSerie(Index).Text = "" Then
        txtNSerie(Index).Text = ""
        Exit Sub
    End If
    
    Select Case Index
        Case 0, 1 'tipos de movimiento
            txtNSerie(Index).Text = DevuelveDesdeBD("nomregis", "contadores", "tiporegi", txtSerie(Index), "T")
    End Select
    

    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


End Sub



Private Sub AccionesCSV(EsListadoRetenciones As Boolean)
Dim Sql2 As String

    'Monto el SQL
    
    If EsListadoRetenciones Then
        
        Sql = " SELECT factpro.tiporeten,factpro.numserie Serie,factpro.numregis Registro,tmpfaclin.Numfac, tmpfaclin.codigo, tmpfaclin.Cliente, tmpfaclin.NIF, "
        Sql = Sql & " tmpfaclin.cta, tmpfaclin.ctabase,tmpfaclin.numfactura, tmpfaclin.Fecha, factpro.totbases, factpro.totbasesret, factpro.suplidos, factpro.totivas,"
        Sql = Sql & " factpro.totrecargo, factpro.totfacpr, factpro.trefacpr, tmppresu1.titulo"
        Sql = Sql & " FROM   ariconta5.tmpfaclin tmpfaclin INNER JOIN"
        Sql = Sql & " factpro factpro ON ((tmpfaclin.numserie=factpro.numserie) AND (tmpfaclin.Numfac=factpro.numregis))"
        Sql = Sql & " AND (tmpfaclin.codigo=factpro.anofactu))"
        Sql = Sql & " INNER JOIN ariconta5.tmppresu1 tmppresu1 ON (tmpfaclin.codusu=tmppresu1.codusu) AND (factpro.tiporeten=tmppresu1.codigo)"
        Sql = Sql & " Where tmpfaclin.CodUsu = " & vUsu.Codigo
        
        
    Else
        Sql = "Select factpro.numserie Serie, tmpfaclin.nomserie Descripcion, factpro.numfactu Factura,factpro.numregis Registro, factpro.fecfactu Fecha, factpro.fecharec FRecep ,factpro.codmacta Cuenta, "
        Sql = Sql & " factpro.nommacta Titulo,nifdatos NIF , tmpfaclin.tipoformapago TipoPago, "
        Sql = Sql & " tmpfaclin.tipoopera TOperacion, factpro.codconce340 TFra, factpro.trefacpr Retencion,factpro.suplidos Suplidos , "
        Sql = Sql & " factpro_totales.baseimpo BaseImp,factpro_totales.codigiva IVA,factpro_totales.porciva PorcIva,factpro_totales.porcrec PorcRec,"
        Sql = Sql & " factpro_totales.impoiva ImpIva,factpro_totales.imporec ImpRec "
        Sql = Sql & " FROM (factpro inner join factpro_totales on factpro.numserie = factpro_totales.numserie and factpro.numregis = factpro_totales.numregis "
        Sql = Sql & " and factpro.fecharec = factpro_totales.fecharec) "
        Sql = Sql & " inner join tmpfaclin ON factpro.numserie=tmpfaclin.numserie AND factpro.numregis=tmpfaclin.Numfac and "
        
        
        Sql = Sql & " tmpfaclin.fecha = "
        If Me.optVarios(0).Value Then
            Sql = Sql & " factpro.fecfactu"
        Else
            If Me.optVarios(4).Value Then
                Sql = Sql & " factpro.fecliqpr"
            Else
                Sql = Sql & " factpro.fecharec"
            End If
        End If
        
        
        
        Sql = Sql & " WHERE  tmpfaclin.codusu =  " & vUsu.Codigo
        Sql = Sql & " ORDER BY factpro.codmacta, factpro.nommacta, factpro_totales.numlinea "
    End If
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal(EsListadoRetenciones As Boolean)
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
        
            
        
    If EsListadoRetenciones Then
        indRPT = "0405-04"
    Else
        'Lo que habia
        indRPT = "0405-00"
        If optVarios(3).Value Then indRPT = "0405-03"  'Numero registro
        If optVarios(1).Value Then indRPT = "0405-01"
        If Check1(1).Value And optVarios(1).Value Then indRPT = "0405-02"
    End If
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '

    'si se imprime el nif o la cuenta de cliente
    cadParam = cadParam & "pConNIF=" & Check1(0).Value & "|"
    numParam = numParam + 1
    cadParam = cadParam & "pConFecha=" & Check1(2).Value & "|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pNumPag=" & ComprobarCero(txtPag2(0).Text) & "|"
    numParam = numParam + 1
    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
    numParam = numParam + 1
    
    
     If EsListadoRetenciones Then
        If optVarios(1).Value Then cadParam = cadParam & "pOrden={tmpfaclin.cliente}|" ' fecha de recepcion
    End If
    
    If optVarios(3).Value Then cadParam = cadParam & "pOrden={tmpfaclin.Numfac}|" ' nro de registro
    If optVarios(0).Value Then cadParam = cadParam & "pOrden={tmpfaclin.fecha}|"   'ctabase --> YA NO
    If optVarios(2).Value Then cadParam = cadParam & "pOrden={tmpfaclin.ctabase}|" ' fecha de recepcion
    
        
    numParam = numParam + 1
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text, (Legalizacion <> "")) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then
    
        cadParam = ""
        If Me.txtCuentas(0).Text <> "" Then
            If Me.txtCuentas(0).Text = txtCuentas(1).Text Then cadParam = DevuelveDesdeBD("maidatos", "cuentas", "codmacta", txtCuentas(0).Text, "T")
        End If
        LanzaProgramaAbrirOutlook 17, cadParam
    End If
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Function CargarTemporal() As Boolean
Dim Sql As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    Sql = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    
    
    Sql = "insert into tmpfaclin (codusu, codigo,numfactura, numserie, nomserie, numfac, fecha, cta, cliente, nif, imponible, impiva, total, retencion,"
    Sql = Sql & " recargo,  tipoformapago,tipoopera, tipoiva,ctabase,suplidos) "
    Sql = Sql & " select distinct " & vUsu.Codigo & ",factpro.anofactu,numfactu, factpro.numserie, contadores.nomregis, factpro.numregis,"
    If Me.optVarios(0).Value Then
        Sql = Sql & " factpro.fecfactu"
    Else
        If Me.optVarios(4).Value Then
            Sql = Sql & " factpro.fecliqpr"
        Else
            Sql = Sql & " factpro.fecharec"
        End If
    End If
    Sql = Sql & " ,factpro.codmacta,factpro.nommacta,factpro.nifdatos, "
    'Marzo 2020
    Sql = Sql & " factpro.totbases "    'Antes:factpro.totbases
    Sql = Sql & " , factpro.totivas, factpro.totfacpr, factpro.trefacpr, "
    Sql = Sql & " factpro.totrecargo, tipofpago.descformapago , "
    
    'Abril 2017
    'if(factpro.codopera = 0 or factpro.codopera = 3, 1,0) aaaa "
    'Sql = Sql & " aa.denominacion, " '
    Sql = Sql & " factpro.codopera, " ',
    
    Sql = Sql & " CASE factpro.codopera WHEN 0 THEN 0 WHEN 3 THEN 0 WHEN 1 THEN 1 WHEN 2 THEN 2 WHEN 4 THEN 3 WHEN 5 THEN 4 WHEN 6 THEN 5 END "
    'SQL = SQL & ", fecfactu   "  'DATE_FORMAT(fecfactu,'%d/%m/%Y") "
    
    Sql = Sql & ", DATE_FORMAT(" & IIf(optVarios(0).Value, "factpro.fecharec", "factpro.fecfactu") & ",'%d/%m/%Y') "
    
    Sql = Sql & ",suplidos"
    
    Sql = Sql & " from " & tabla
    Sql = Sql & " where " & cadselect
    
    Conn.Execute Sql
    
    

    
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
End Function

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String
Dim i As Integer


    MontaSQL = False
    
    If Not PonerDesdeHasta("factpro.NumSerie", "SER", Me.txtSerie(0), Me.txtNSerie(0), Me.txtSerie(1), Me.txtNSerie(1), "pDHSerie=""") Then Exit Function
    
    If Me.optVarios(0).Value Then
        If Not PonerDesdeHasta("factpro.fecfactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""factura:") Then Exit Function
    Else
        If Me.optVarios(4).Value Then
            If Not PonerDesdeHasta("factpro.fecliqpr", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""liquidaci�n: ") Then Exit Function
        Else
            If Not PonerDesdeHasta("factpro.fecharec", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""recepci�n: ") Then Exit Function
        End If
    End If
    
        
    If Not PonerDesdeHasta("factpro.Numregis", "REG", Me.txtNumFactu(0), Me.txtNumFactu(0), Me.txtNumFactu(1), Me.txtNumFactu(1), "pDHFactura=""") Then Exit Function
    If Not PonerDesdeHasta("factpro.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
            
            
    If txtNIF.Text <> "" Then
        If Not AnyadirAFormula(cadselect, "factpro.nifdatos = " & DBSet(txtNIF.Text, "T")) Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{factpro.nifdatos} = """ & Trim(txtNIF.Text) & "") Then Exit Function
    End If
    
    If Me.Combo1(0).ListIndex > 0 Then ' en todos y en -1 no hacemos nada
        If Not AnyadirAFormula(cadselect, "factpro.codconce340 = " & DBSet(Chr(Combo1(0).ItemData(Combo1(0).ListIndex)), "T")) Then Exit Function
        If Not AnyadirAFormula(cadFormula, "{factpro.codconce340} = """ & Chr(Combo1(0).ItemData(Combo1(0).ListIndex)) & "") Then Exit Function
    End If
            
            
            
    Sql = ""
    RC = ""
    RC2 = ""
    For i = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(i).Checked Then
            Sql = Sql & Me.ListView1(1).ListItems(i).Text & ","
            RC2 = RC2 & " - " & ListView1(1).ListItems(i).SubItems(1)
        Else
            RC = "N"
        End If
    Next i
    
    If Sql <> "" Then
        ' quitamos la ultima coma
        If RC = "N" Then
            Sql = Mid(Sql, 1, Len(Sql) - 1)
            
            If Not AnyadirAFormula(cadselect, "factpro_totales.codigiva in (" & Sql & ")") Then Exit Function
            If Not AnyadirAFormula(cadFormula, "{factpro_totales.codigiva} in [" & Sql & "]") Then Exit Function
        Else
            RC2 = ""
        End If
    Else
        If Not AnyadirAFormula(cadselect, "factpro_totales.codigiva is null") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "isnull({factpro_totales.codigiva})") Then Exit Function
        RC2 = "NULO"
    End If
    'Para el texto rpt
    If RC2 <> "" Then
        cadParam = cadParam & "pIvas=""" & RC2 & """|"
    End If
    
    'Retencion
    Sql = ""
    RC = ""
    RC2 = ""
    For i = 1 To Me.ListView1(0).ListItems.Count
        If Me.ListView1(0).ListItems(i).Checked Then
            Sql = Sql & Me.ListView1(0).ListItems(i).Text & ","
            RC2 = RC2 & " - " & ListView1(0).ListItems(i).SubItems(1)
        Else
            RC = "N"
        End If
    Next i
    If Sql <> "" Then
        If RC = "N" Then
        
            ' quitamos la ultima coma
            Sql = Mid(Sql, 1, Len(Sql) - 1)
            
            If Not AnyadirAFormula(cadselect, "factpro.tiporeten in (" & Sql & ")") Then Exit Function
            If Not AnyadirAFormula(cadFormula, "{factpro.tiporeten} in [" & Sql & "]") Then Exit Function
    
        Else
            'Estan todos. NO hay que separar
            RC2 = ""
        End If
    Else
        'Nnguno marcado
        If Not AnyadirAFormula(cadselect, "factpro.tiporeten is null") Then Exit Function
        If Not AnyadirAFormula(cadFormula, "isnull({factpro.tiporeten})") Then Exit Function
    End If
    'Para el texto rpt
    If RC2 <> "" Then
        cadParam = cadParam & "pReten=""" & RC2 & """|"
    End If
    
    
    
    
    
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
    If Not CargarTemporal Then Exit Function
    


    
    
    cadFormula = "{tmpfaclin.codusu} = " & vUsu.Codigo
    
            
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
        
        LanzaFormAyuda txtFecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    
    DatosOK = True


End Function


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear
    
    If Index = 1 Then
        ListView1(Index).ColumnHeaders.Add , , "C�digo", 900
        ListView1(Index).ColumnHeaders.Add , , "Descripci�n", 2200
    Else
        ListView1(Index).ColumnHeaders.Add , , "C�digo", 240
        ListView1(Index).ColumnHeaders.Add , , "Descripci�n", 3100
    End If
    If Index = 1 Then
        Sql = "SELECT codigiva, nombriva "
        Sql = Sql & " FROM tiposiva "
        Sql = Sql & " ORDER BY codigiva "
    Else
        Sql = "SELECT * FROM usuarios.wtiporeten ORDER BY codigo"
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        ItmX.Checked = True
        ItmX.Text = Rs.Fields(0).Value
        ItmX.SubItems(1) = Rs.Fields(1).Value
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, IIf(Index = 1, "Cargar Tipo de Iva.", "carga retencion"), Err.Description
    End If
End Sub


Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim J As Long

    For i = 1 To Combo1.Count - 1
        Combo1(i).Clear
    Next i

    'Tipo de factura
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wconce340 ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Combo1(0).AddItem "Todos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = Asc(1)
    
    i = 0
    While Not Rs.EOF
        Combo1(0).AddItem Rs!Descripcion
        Combo1(0).ItemData(Combo1(0).NewIndex) = Asc(Rs!Codigo)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    Set Rs = Nothing

End Sub

