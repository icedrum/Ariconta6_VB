VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacturProv 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmFacturasProv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framecabeceras 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   -60
      TabIndex        =   42
      Top             =   600
      Width           =   11895
      Begin VB.CheckBox Check2 
         Caption         =   "NO deducible"
         Height          =   255
         Left            =   9120
         TabIndex        =   9
         Tag             =   "No deducible|N|S|||cabfactprov|nodeducible|||"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Intracomunitaria"
         Height          =   255
         Left            =   7320
         TabIndex        =   8
         Tag             =   "Extranjero|N|S|||cabfactprov|extranje|||"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   555
         Left            =   3840
         TabIndex        =   79
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   29
         Left            =   3900
         TabIndex        =   7
         Tag             =   "Fecha liquidacion|F|N|||cabfactprov|fecliqpr|dd/mm/yyyy||"
         Text            =   "T"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   28
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Nº factura|T|N|||cabfactprov|numfacpr|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   27
         Left            =   120
         TabIndex        =   74
         Tag             =   "totalfactura|N|S|||cabfactprov|totfacpr||N|"
         Text            =   "Text1"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   26
         Left            =   120
         TabIndex        =   73
         Tag             =   "año factura|N|S|||cabfactprov|anofacpr||S|"
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   25
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "Observaciones(Concepto)|T|S|||cabfactprov|confacpr|||"
         Text            =   "DDDDDDDDDDDDDDD"
         Top             =   960
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   4800
         TabIndex        =   3
         Tag             =   "Fecha factura|F|N|||cabfactprov|fecfacpr|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Tag             =   "Fecha recepcion|F|N|||cabfactprov|fecrecpr|dd/mm/yyyy||"
         Text            =   "T"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEF7E4&
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1680
         TabIndex        =   1
         Tag             =   "Nº registro|N|S|0||cabfactprov|numregis||S|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   10800
         TabIndex        =   5
         Tag             =   "Numero serie|N|S|||cabfactprov|numasien|||"
         Text            =   "9999999999"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   7560
         TabIndex        =   55
         Text            =   "Text4"
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   6240
         TabIndex        =   4
         Tag             =   "Cuenta cliente|T|N|||cabfactprov|codmacta|||"
         Text            =   "0000000000"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   6
         Left            =   1680
         TabIndex        =   10
         Tag             =   "Base imponible 1|N|N|||cabfactprov|ba1facpr|#,###,###,##0.00||"
         Top             =   1755
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   3240
         TabIndex        =   11
         Tag             =   "Tipo IVA 1|N|N|0|100|cabfactprov|tp1facpr|||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   8
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   54
         Tag             =   "Porcentaje IVA 1|N|S|||cabfactprov|pi1facpr|#0.00||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   9
         Left            =   6960
         TabIndex        =   12
         Tag             =   "Importe IVA 1|N|N|||cabfactprov|ti1facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   10
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   53
         Tag             =   "Porcentaje recargo 1|N|S|||cabfactprov|pr1facpr|#0.00||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   11
         Left            =   9240
         TabIndex        =   13
         Tag             =   "Importe recargo 1|N|S|||cabfactprov|tr1facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   12
         Left            =   1680
         TabIndex        =   14
         Tag             =   "Base imponible 2|N|S|||cabfactprov|ba2facpr|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   3240
         TabIndex        =   15
         Tag             =   "Tipo IVA 2|N|S|0|100|cabfactprov|tp2facpr|||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   14
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   52
         Tag             =   "Porcentaje IVA 2|N|S|||cabfactprov|pi2facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   15
         Left            =   6960
         TabIndex        =   16
         Tag             =   "Importe IVA 2|N|S|||cabfactprov|ti2facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   16
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   51
         Tag             =   "Porcentaje recargo 2|N|S|||cabfactprov|pr2facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   17
         Left            =   9240
         TabIndex        =   17
         Tag             =   "Importe recargo 2|N|S|||cabfactprov|tr2facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   18
         Left            =   1680
         TabIndex        =   18
         Tag             =   "Base imponible 3|N|S|||cabfactprov|ba3facpr|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   19
         Left            =   3240
         TabIndex        =   19
         Tag             =   "Tipo IVA 3|N|S|0|100|cabfactprov|tp3facpr|||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   20
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   50
         Tag             =   "Porcentaje IVA 3|N|S|||cabfactprov|pi3facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   21
         Left            =   6960
         TabIndex        =   20
         Tag             =   "Importe IVA 3|N|S|||cabfactprov|ti3facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   22
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   49
         Tag             =   "Porcentaje recargo 3|N|S|||cabfactprov|pr3facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   23
         Left            =   9240
         TabIndex        =   21
         Tag             =   "Importe recargo 3|N|S|||cabfactprov|tr3facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   46
         Text            =   "Text4"
         Top             =   1755
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   3960
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   3960
         TabIndex        =   44
         Text            =   "Text4"
         Top             =   2805
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3960
         TabIndex        =   43
         Text            =   "Text4"
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   2520
         TabIndex        =   23
         Tag             =   "Cuenta retencion|T|S|||cabfactprov|cuereten|||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   24
         Left            =   1680
         TabIndex        =   22
         Tag             =   "Porcentaje retencion|N|S|||cabfactprov|retfacpr|#0.00||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   6960
         TabIndex        =   24
         Tag             =   "Cuenta retencion|N|S|||cabfactprov|trefacpr|#,##0.00||"
         Text            =   "Text2"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "123.123.123.123,11"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   7
         Left            =   4920
         Picture         =   "frmFacturasProv.frx":000C
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Liquidacion"
         Height          =   195
         Index           =   4
         Left            =   3900
         TabIndex        =   78
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Nº factura"
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   77
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Recepción"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   76
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   72
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   " Fecha"
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   71
         Top             =   120
         Width           =   495
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmFacturasProv.frx":010E
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   5280
         Picture         =   "frmFacturasProv.frx":0199
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº registro"
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   70
         Top             =   120
         Width           =   735
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   7080
         Picture         =   "frmFacturasProv.frx":0224
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   7
         Left            =   6240
         TabIndex        =   69
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento"
         Height          =   195
         Index           =   8
         Left            =   10800
         TabIndex        =   68
         Top             =   120
         Width           =   975
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   3705
         Picture         =   "frmFacturasProv.frx":0C26
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   3705
         Picture         =   "frmFacturasProv.frx":1628
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   3705
         Picture         =   "frmFacturasProv.frx":202A
         Top             =   2880
         Width           =   240
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   10440
         Y1              =   3165
         Y2              =   3165
      End
      Begin VB.Label Label3 
         Caption         =   "Base Imponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   67
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Importes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   66
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de I.V.A."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   65
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "% I.V.A."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   64
         Top             =   1515
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "T.R. equiv."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   9240
         TabIndex        =   63
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Importe IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6960
         TabIndex        =   62
         Top             =   1515
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "% Rec."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   8520
         TabIndex        =   61
         Top             =   1515
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Retención"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   3795
         Width           =   1455
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   3675
         Picture         =   "frmFacturasProv.frx":2A2C
         Top             =   3870
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   2
         Left            =   8640
         TabIndex        =   59
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label Label3 
         Caption         =   "Total Ret."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   7080
         TabIndex        =   58
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta retención"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2520
         TabIndex        =   57
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "% Ret."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1680
         TabIndex        =   56
         Top             =   3600
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   8160
      TabIndex        =   32
      Top             =   7200
      Width           =   195
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   -120
      Top             =   7320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   30
      Top             =   7200
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10680
      TabIndex        =   27
      Top             =   7440
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   3720
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   1
      Left            =   4800
      TabIndex        =   39
      Top             =   7200
      Width           =   1395
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   6900
      MaxLength       =   10
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7200
      Width           =   945
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   320
      Index           =   3
      Left            =   8040
      TabIndex        =   34
      Top             =   7200
      Width           =   885
   End
   Begin VB.TextBox txtaux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   8880
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   7200
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -120
      Top             =   7440
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10680
      TabIndex        =   37
      Top             =   7410
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   35
      Top             =   7320
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9480
      TabIndex        =   26
      Top             =   7440
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacturasProv.frx":342E
      Height          =   2295
      Left            =   1680
      TabIndex        =   38
      Top             =   4920
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Contabilizar factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8280
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "P R O V E E D O R E S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   3900
      TabIndex        =   75
      Top             =   7320
      Width           =   5175
   End
   Begin VB.Menu mnOpcionesAsiPre 
      Caption         =   "&Opciones"
      HelpContextID   =   11111
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "&Filtro ejercicios"
      Begin VB.Menu mnActuralySiguiente 
         Caption         =   "Actual y siguiente"
      End
      Begin VB.Menu mnActual 
         Caption         =   "Actual"
      End
      Begin VB.Menu mnSiguiente 
         Caption         =   "Siguiente"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSinFiltro 
         Caption         =   "&Sin filtro"
      End
   End
End
Attribute VB_Name = "frmFacturProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 404

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Public FACTURA As String  'Con pipes numdiari|fechanormal|numasien


Private Const NO = "No encontrado"
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCo As frmContadores
Attribute frmCo.VB_VarHelpID = -1
Private WithEvents frmCC As frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmI As frmIVA
Attribute frmI.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busquedaa
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'//////////////////////////////////
'//////////////////////////////////
'//////////////////////////////////
'   Nuevo modo --> Modificando lineas
'  5.- Modificando lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private SQL As String
Dim i As Integer
Dim ancho As Integer


'para cuando modifica factura, y vuelve a integrar para forzar el numero de asiento
Dim Numasien2 As Long
Dim NumDiario As Integer

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas



'Para pasar de lineas a cabeceras
Dim Linfac As Long
Private ModificandoLineas As Byte
'0.- A la espera       1.- Insertar     2.- Modificar

Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean
Dim Rs As Recordset
Dim AUx As Currency
Dim Base As Currency
Dim AUX2 As Currency
Dim SumaLinea As Currency
Dim AntiguoText1 As String

Dim FILTRO  As Byte
Dim CTA_Inmovilizado As String
Dim NuevaFactura As Boolean


'Por si esta en un periodo liquidado, que pueda modificar CONCEPTO , cuentas,
Private ModificaFacturaPeriodoLiquidado As Boolean

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Function ActualizaFactura() As Boolean
Dim B As Boolean
On Error GoTo EActualiza
ActualizaFactura = False

B = ModificaDesdeFormularioClaves(Me, SQL)
If Not B Then Exit Function

'Las lineas
If Not adodc1.Recordset.EOF Then
    SQL = "UPDATE linfactprov SET numregis =" & Text1(2).Text
    SQL = SQL & " ,anofacpr = " & Text1(26).Text
    SQL = SQL & " WHERE numregis='" & Data1.Recordset!NumRegis
    SQL = SQL & " AND anofacpr =" & Data1.Recordset!anofacpr
    Conn.Execute SQL
End If

ActualizaFactura = True
Exit Function
EActualiza:
    MuestraError Err.Number, "Modificando claves factura"
End Function


Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim i As Integer
    Dim RC As Boolean
    Dim Mc As Contadores
    Dim CadConsulta As String
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOK Then
            i = FechaCorrecta2(CDate(Text1(1).Text))
            If i > 1 Then
                If i = 2 Then
                    Cad = varTxtFec
                Else
                    Cad = "La fecha factura no pertenece al ejercicio actual ni al siguiente."
                End If
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
            Set Mc = New Contadores
            If Mc.ConseguirContador("1", i = 0, False) = 1 Then
                Set Mc = Nothing
                Exit Sub
            End If
            Text1(2).Text = Mc.Contador
                        
            '-----------------------------------------
            'Hacemos insertar
                If InsertarDesdeForm(Me) Then
        
                    'LOG
                    vLog.Insertar 7, vUsu, Text1(2).Text & " " & Text1(1).Text
                           
        
        
                    If SituarData1(1) Then
                        PonerModo 5
                        'Haremos como si pulsamo el boton de insertar nuevas lineas
                        'Ponemos el importe en AUX
                        AUx = ImporteFormateado(Text2(4).Text)
                        cmdCancelar.Caption = "Cabecera"
                        ModificandoLineas = 0
                        
                        'Primero insertaremos los vencimientos
   
'                        If vEmpresa.TieneTesoreria Then
'                            Screen.MousePointer = vbHourglass
'                            frmVto.Show vbModal
'                            Screen.MousePointer = vbDefault
'                        End If
                        'Luego las bases
                        AnyadirLinea True, False
                    Else
                        SQL = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                        SQL = SQL & vbCrLf & " CLAVE: FrmFacturas. cmdAceptar. SituarData1"
                        MsgBox SQL, vbCritical
                    End If
                Else
                    'NO inserta,
                    Mc.DevolverContador "1", i = 0, Mc.Contador
                    Text1(2).Text = ""
                End If
        End If
    Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hay que comprobar si ha modificado, o no la clave de la factura
                i = 1
                If Data1.Recordset!NumRegis = Text1(2).Text Then
                        If Data1.Recordset!anofacpr = Text1(26).Text Then
                            i = 0
                            'NO HA MODIFICADO NADA
                    End If
                End If
            
                'Hacemos MODIFICAR
                If i <> 0 Then
                    'Modificar claves
                    SQL = " numregis = " & Data1.Recordset!NumRegis
                    SQL = SQL & " AND anofacpr=" & Data1.Recordset!anofacpr
                    Conn.BeginTrans
                    RC = ActualizaFactura
                    If RC Then
                        Conn.CommitTrans
                    Else
                        Conn.RollbackTrans
                    End If
                Else
                    RC = ModificaDesdeFormulario(Me)
                End If
                    
                If RC Then
                    If Numasien2 > 0 Then
                        'Por que si la busqueda era por numasien, por ejemplo, ahora
                        'al no tener valor no situaria el datagrid correctamente
                        Data1.RecordSource = "Select * from " & NombreTabla & " where numregis = " & Text1(2).Text
                    End If
                        
                    'LOG
                    vLog.Insertar 8, vUsu, Text1(2).Text & " " & Text1(1).Text
                           
            
                        
                        
                    If SituarData1(0) Then
                        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                        PonerModo 2
                        If Numasien2 > 0 Then
                            If IntegrarFactura Then
                                
                                Text1(4).Text = Numasien2
                                Numasien2 = -1
                                NumDiario = 0
                                Data1.RecordSource = CadenaConsulta
                            
                                DoEvents
                                If Not SituarData1(0) Then PonerModo 0
                                
                            End If
                        End If
                    Else
                        PonerModo 0
                    End If
                    DesBloqueaRegistroForm Text1(0)
                End If
            End If
            
    Case 5
        Cad = AuxOK
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
        Else
            'Insertaremos, o modificaremos
            If InsertarModificar Then
                'Reestablecemos los campos
                'y ponemos el grid
                cmdAceptar.Visible = False
                DataGrid1.AllowAddNew = False
                CargaGrid True
                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    AnyadirLinea True, False
                    If AUx <> 0 Then PonerFoco txtAux(0)
                    
                Else
                    ModificandoLineas = 0
                    CamposAux False, 0, False
                    cmdCancelar.Caption = "Cabecera"
                End If
            End If
        End If
    Case 1
        HacerBusqueda
    End Select
        
Error1:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdAux_Click(Index As Integer)
If Index = 0 Then
    imgppal_Click 100
    HabilitarCentroCoste
Else
    Set frmCC = New frmCCoste
    frmCC.DatosADevolverBusqueda = "0|1|"
    frmCC.Show vbModal
    Set frmCC = Nothing
    If txtAux(2).Text <> "" Then PonerFoco txtAux(4)
    
End If
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3
        LimpiarCampos
        PonerModo 0
        'Contador de facturas
    Case 4
        Modo = 2
        If Numasien2 > 0 Then
            'Es decir. Anofacl
            Text1(1).Text = Data1.Recordset!fecrecpr
            Text1(26).Text = Data1.Recordset!anofacpr
            If Not IntegrarFactura Then Exit Sub
        End If
        PonerCampos
        Modo = 4  'Reestablezco el modo para que vuelva a hahacer ponercampos
        lblIndicador.Caption = ""
        PonerModo 2
        
        DesBloqueaRegistroForm Text1(0)
    Case 5
        CamposAux False, 0, False

        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
            AntiguoText1 = ""
            If adodc1.Recordset.EOF Then
                AntiguoText1 = "La factura no tiene lineas. ¿SEGURO que desea salir?"
                If MsgBox(AntiguoText1, vbQuestion + vbYesNoCancel) = vbYes Then
                    AntiguoText1 = ""
                Else
                    'Para k no muestre el siguiente punto de error
                    AntiguoText1 = "###"
                End If
            Else
                'Comprobamos que el total de factura es el de suma
               ObtenerSigueinteNumeroLinea
               If AUx <> 0 Then AntiguoText1 = "El importe de lineas no suma el importe facturas: " & Format(AUx, "###,##0.00")
            End If
            If AntiguoText1 <> "" Then
                If AntiguoText1 <> "###" Then MsgBox AntiguoText1, vbExclamation
                Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            DesBloqueaRegistroForm Text1(0)
            If Numasien2 > 0 Then
                If IntegrarFactura Then
                    Text1(4).Text = Numasien2
                    Numasien2 = 0
                    NumDiario = 0
                    
                Else
                    'Si tenia numero de asiento, lo perdera debeido al error
                    If Numasien2 > 0 Then
                        SQL = "Desea modificar la cabecera de factura ?"
                        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                            'Ponemos el modo y poco mas
                            cmdAceptar.Caption = "Modificar"
                            cmdCancelar.Caption = "Cancelar"
                            PonerModo 4

            
                            DespalzamientoVisible False
                            lblIndicador.Caption = "Modificar"
                            PonerFoco Text1(28)
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        Else
                            Text1(4).Text = ""
                            Numasien2 = 0
                            NumDiario = 0
                        End If
                    End If
                End If
            End If
            If NuevaFactura Then
            
                If vParam.ContabilizaFactura Then HacerToolbar1 11, True
                If vEmpresa.TieneTesoreria Then HacerToolbar1 12, False
            End If
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
            Screen.MousePointer = vbDefault
        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
                 DataGrid1.Refresh
            End If

            cmdAceptar.Visible = False
            cmdCancelar.Caption = "Cabeceras"
            ModificandoLineas = 0
        End If
    End Select
End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)


'-----------------------------------------------
'     DICIEMBRE 2005, 20.
'     Modificacioin para que desoues de insertar
'     solo carge la factura esta, la nueva
'       OPCIONINSERTAR:     0.- Todas
'                           1.- Insertar
Private Function SituarData1(OpcionInsertar As Byte) As Boolean
    Dim SQL As String
    Dim OldC As Byte
    On Error GoTo ESituarData1
    
    OldC = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    If OpcionInsertar <> 1 Then
        If FILTRO > 1 Then
            'Por si acaso pone la de una año u la de otro
            CadenaConsulta = "Select * from " & NombreTabla & " where fecrecpr >='" & Format(vParam.fechaini, FormatoFecha) & "' " & Ordenacion
            Data1.RecordSource = CadenaConsulta
        End If
    Else
        
        CadenaConsulta = "Select * from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE NumRegis = " & Text1(2).Text
        CadenaConsulta = CadenaConsulta & " AND anofacpr = " & Text1(26).Text
        Data1.RecordSource = CadenaConsulta
        
    End If
    espera 0.2
    Data1.Refresh
    
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!NumRegis) = Text1(2).Text Then
                If CStr(.Fields!anofacpr) = Text1(26).Text Then
                        SituarData1 = True
                        Screen.MousePointer = OldC
                        Exit Function
                End If
            End If
            .MoveNext
        Wend
    End With
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
        Screen.MousePointer = OldC
End Function


Private Function IntegrarFactura() As Boolean
IntegrarFactura = False
If Text1(2).Text = "" Then Exit Function
'Primero comprobamos que esta cuadrada
If IsNull(Data1.Recordset!totfacpr) Then
    MsgBox "La factura no tiene importes", vbExclamation
    Exit Function
End If
'Sumamos las bases
Base = 0
If Not IsNull(Data1.Recordset!ba1facpr) Then Base = Base + Data1.Recordset!ba1facpr
If Not IsNull(Data1.Recordset!ba2facpr) Then Base = Base + Data1.Recordset!ba2facpr
If Not IsNull(Data1.Recordset!ba3facpr) Then Base = Base + Data1.Recordset!ba3facpr
AUX2 = Base 'Sumatorio imponibles1

'Le sumamos los IVAS
If Not IsNull(Data1.Recordset!ti1facpr) Then Base = Base + Data1.Recordset!ti1facpr
If Not IsNull(Data1.Recordset!ti2facpr) Then Base = Base + Data1.Recordset!ti2facpr
If Not IsNull(Data1.Recordset!ti3facpr) Then Base = Base + Data1.Recordset!ti3facpr

'Los recargos
If Not IsNull(Data1.Recordset!tr1facpr) Then Base = Base + Data1.Recordset!tr1facpr
If Not IsNull(Data1.Recordset!tr2facpr) Then Base = Base + Data1.Recordset!tr2facpr
If Not IsNull(Data1.Recordset!tr3facpr) Then Base = Base + Data1.Recordset!tr3facpr

'La retencion( es en negativo)
If Not IsNull(Data1.Recordset!trefacpr) Then Base = Base - Data1.Recordset!trefacpr

If Base <> Data1.Recordset!totfacpr Then
    MsgBox "Total factura no coincide con la suma de importes.", vbExclamation
    Exit Function
End If

'Comprobamos que la suma de lineas es las base imponible
ObtenerSigueinteNumeroLinea
'En suma lineas tendremos la suma del los imports
If SumaLinea <> AUX2 Then
    MsgBox "La suma de las lineas no coincide con la suma de bases imponibles.", vbExclamation
    Exit Function
End If



'Esta "cuadrado"

With frmActualizar
    .OpcionActualizar = 8
    'NumAsiento     --> CODIGO FACTURA
    'NumDiari       --> AÑO FACTURA
    'NUmSerie       --> SERIE DE LA FACTURA
    'FechaAsiento   --> Fecha factura
    .NumFac = CLng(Text1(2).Text)     'La fecha es la de recepcion
    .NumDiari = CInt(Text1(26).Text)
    .NumSerie = Text1(1).Text
    .FechaAsiento = Text1(1).Text
    If Numasien2 < 0 Then
        If Not Text1(4).Enabled Then
            If Text1(4).Text <> "" Then
                Numasien2 = Text1(4).Text
            End If
        End If
    End If
    If NumDiario <= 0 Then NumDiario = vParam.numdiapr
    .DiarioFacturas = NumDiario
    .NumAsiento = Numasien2
    .Show vbModal
    Me.Refresh
    If AlgunAsientoActualizado Then IntegrarFactura = True
End With

If IntegrarFactura Then
    Data1.Refresh
    If Not SituarData1(0) Then
       'If Not Data1.Recordset.EOF Then
       If TieneRegistros Then Data1.Recordset.MoveFirst
'        Else
'            LimpiarCampos
'            PonerModo 0
'        End If
    End If
End If
End Function

Private Function TieneRegistros() As Boolean
    On Error Resume Next
    TieneRegistros = False
    If Data1.Recordset.RecordCount > 0 Then TieneRegistros = True
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    NuevaFactura = True
    Check2.Value = 0 'NO DEDUCIBLE
    Check1.Value = 0 'Intracomunitaria

    SQL = AnyadeCadenaFiltro
    If SQL <> "" Then SQL = " WHERE " & SQL
    CadenaConsulta = "Select * from " & NombreTabla & SQL & Ordenacion
    PonerCadenaBusqueda True

    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    '###A mano
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(1)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        cmdAceptar.Caption = "&Aceptar"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(1)
        Text1(1).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                PonerFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    DataGrid1.Enabled = False
    CargaGrid False
    SQL = AnyadeCadenaFiltro
    If chkVistaPrevia.Value = 1 Then
        
        MandaBusquedaPrevia SQL
    Else
        If SQL <> "" Then SQL = " WHERE " & SQL
        CadenaConsulta = "Select * from " & NombreTabla & SQL & Ordenacion
        PonerCadenaBusqueda False
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)

Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
    Case 1
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 3
        Data1.Recordset.MoveLast
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    NuevaFactura = False
    
    'Comprobamos la fecha pertenece al ejercicio. Fecha RECEPCION (JULIO 2011)
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk >= 2 And varFecOk <= 3 Then
        If varFecOk = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            MsgBox "La factura pertenece a un ejercicio cerrado.", vbExclamation
        End If
        Exit Sub
    End If
    
    If Not ComprobarPeriodo(29) Then Exit Sub
    
    
    'Comprobamos que no esta actualizada ya
    If Not IsNull(Data1.Recordset!NumAsien) Then
        Numasien2 = Data1.Recordset!NumAsien
        If Numasien2 = 0 Then
            MsgBox "Contabilización de facturas especial. No puede modificarse", vbExclamation
            Exit Sub
        End If
        NumDiario = Data1.Recordset!NumDiari
        SQL = "Esta factura ya esta contabilizada. Desea desactualizar para poder modificarla?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        'Tengo desintegre la factura del hco
        If Not Desintegrar Then Exit Sub
        Text1(4).Text = ""
    End If
        
    'Llegados aqui bloqueamos desde form
    If Not BloqueaRegistroForm(Me) Then Exit Sub

    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4

    
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    PonerFoco Text1(28)
End Sub

Private Sub BotonEliminar()
    Dim i As Long
    Dim Fec As Date
    Dim Mc As Contadores
    
    'Ciertas comprobaciones
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    DataGrid1.Enabled = False
        
    
    'Comprobamos que no esta actualizada ya
    SQL = ""
    If Not IsNull(Data1.Recordset!NumAsien) Then
        SQL = "Esta factura ya esta contabilizada. "
    End If
    SQL = SQL & vbCrLf & vbCrLf & "Va usted a eliminar la factura :" & vbCrLf
    SQL = SQL & "Numero : " & Data1.Recordset!NumRegis & vbCrLf
    SQL = SQL & "Fecha recepcion : " & Format(Data1.Recordset!fecrecpr, "dd/mm/yyyy") & vbCrLf
    SQL = SQL & "Proveedor : " & Data1.Recordset!codmacta & " - " & Text4(0).Text & vbCrLf
    SQL = SQL & vbCrLf & "          ¿Desea continuar ?" & vbCrLf
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    NumRegElim = Data1.Recordset.AbsolutePosition
    Screen.MousePointer = vbHourglass
    'Lo hara en actualizar
    i = 0
    If Not IsNull(Data1.Recordset!NumAsien) Then i = Data1.Recordset!NumAsien
       'Memorizamos el numero de asiento y la fechaent para ver si devolvemos el contador
        'de asientos
    If i > 0 Then
        Fec = Data1.Recordset!FechaEnt
    
        'La borrara desde actualizar
        AlgunAsientoActualizado = False
        With frmActualizar
            .OpcionActualizar = 9
            .NumAsiento = Data1.Recordset!NumAsien
            .NumFac = Data1.Recordset!NumRegis
            .FechaAsiento = Data1.Recordset!FechaEnt
            .NumSerie = Data1.Recordset!anofacpr
            .NumDiari = Data1.Recordset!NumDiari
            .Show vbModal
        End With
        
        
        If AlgunAsientoActualizado Then
            'Devuelvo el contador de asientos. Por si acaso la factura era el ultimo asiento contabilizado
            Set Mc = New Contadores
            Mc.DevolverContador "0", Fec <= vParam.fechafin, i
            Set Mc = Nothing
        End If
    Else
        'La borrara desde este mismo form
        i = Data1.Recordset!NumRegis
        Fec = Data1.Recordset!fecrecpr
        Conn.BeginTrans
        If BorrarFactura Then
        
        
            'LOG
            vLog.Insertar 9, vUsu, Format(i, "000000") & " " & Fec
          
        
            AlgunAsientoActualizado = True
            Conn.CommitTrans
            Set Mc = New Contadores
            Mc.DevolverContador "1", (Fec < vParam.fechafin), i
            Set Mc = Nothing
        Else
            AlgunAsientoActualizado = False
            Conn.RollbackTrans
        End If
    End If
    If Not AlgunAsientoActualizado Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid False
        PonerModo 0
        Else
            Data1.Recordset.MoveFirst
            NumRegElim = NumRegElim - 1
            If NumRegElim > 1 Then
                For i = 1 To NumRegElim - 1
                    Data1.Recordset.MoveNext
                Next i
            End If
            PonerCampos
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function BorrarFactura() As Boolean

    On Error GoTo EBorrar
    
    SQL = " WHERE numregis = " & Data1.Recordset!NumRegis
    SQL = SQL & " AND anofacpr= " & Data1.Recordset!anofacpr
    'Las lineas
    AntiguoText1 = "DELETE from linfactprov " & SQL
    Conn.Execute AntiguoText1
    'La factura
    AntiguoText1 = "DELETE from cabfactprov " & SQL
    Conn.Execute AntiguoText1
    
    
            
    
    
EBorrar:
    If Err.Number = 0 Then
        BorrarFactura = True
    Else
        MuestraError Err.Number, "Eliminar factura"
        BorrarFactura = False
    End If
End Function

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim AUx As String

'If Data1.Recordset.EOF Then
'    MsgBox "Ningún registro devuelto.", vbExclamation
'    Exit Sub
'End If
'
'Cad = ""
'i = 0
'Do
'    j = i + 1
'    i = InStr(j, DatosADevolverBusqueda, "|")
'    If i > 0 Then
'        AUX = Mid(DatosADevolverBusqueda, j, i - j)
'        j = Val(AUX)
'        Cad = Cad & Text1(j).Text & "|"
'    End If
'Loop Until i = 0
'RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub






Private Function DesvincularFactura() As Boolean
On Error Resume Next
Dim T1 As Single
    T1 = Timer
    SQL = "UPDATE cabfactprov set numasien=NULL, fechaent=NULL, numdiari=NULL"
    SQL = SQL & " WHERE numregis = " & Data1.Recordset!NumRegis
    SQL = SQL & " AND anofacpr=" & Data1.Recordset!anofacpr
    Numasien2 = Data1.Recordset!NumAsien
    NumDiario = Data1.Recordset!NumDiari
    Conn.Execute SQL

    If Err.Number <> 0 Then
        DesvincularFactura = False
        MuestraError Err.Number, "Desvincular factura"
    Else
        T1 = 1 - (Timer - T1)
        If T1 > 0 Then
            Conn.Execute "Commit"
            espera T1
        End If
    
        DesvincularFactura = True
    End If
End Function



'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False

        PonerModo CInt(Modo)
        CargaGrid (Modo = 2)
        If Modo <> 2 Then
            CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
            Data1.RecordSource = CadenaConsulta
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon

    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    
    LeerFiltro True
    PonerFiltro FILTRO
        
    Label4.Tag = ""
    CTA_Inmovilizado = "" 'Cuenta por si mete un elto de inmovilzado
    If vParam.Constructoras Then
       ancho = frameTapa.Left + frameTapa.Width + 100
    Else
        ancho = frameTapa.Left
    End If
    Check1.Left = ancho
    Me.Check2.Left = Check1.Left + Check1.Width + 120
    Text1(29).Enabled = vParam.Constructoras
    frameTapa.Visible = Not vParam.Constructoras
    
    Caption = "Registro facturas proveedores (" & vEmpresa.nomresum & ")"
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 17
        .Buttons(13).Image = 16
        .Buttons(14).Image = 15
        .Buttons(16).Image = 6
        .Buttons(17).Image = 7
        .Buttons(18).Image = 8
        .Buttons(19).Image = 9
        
        
        'Si tiene tesoreria entonces
        If vEmpresa.TieneTesoreria Then
            .Buttons(12).Style = tbrDefault
            .Buttons(12).ToolTipText = "Generar vencimientos"
            .Buttons(12).Image = 25
        Else
            .Buttons(12).Style = tbrSeparator
        End If

        
    End With
    
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
'        Me.Width = 12000
'        Me.Height = Screen.Height
    End If
    Me.Height = 8610
    'Los campos auxiliares
    CamposAux False, 0, True
    
    
    '## A mano
    NombreTabla = "cabfactprov"
    Ordenacion = " ORDER BY fecrecpr,numregis"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
'    Data1.UserName = vUsu.Login
'    Data1.password = vUsu.Passwd
'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login
    
    PonerOpcionesMenu
    
    'Maxima longitud cuentas
    txtAux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE anofacpr =-1"
    Data1.Refresh
    CadAncho = False
    Modo = 0
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    NuevaFactura = False
End Sub


'Private Sub Form_Resize()
'If Me.WindowState <> 0 Then Exit Sub
'If Me.Width < 11610 Then Me.Width = 11610
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    LeerFiltro False
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim AUx As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        AUx = ValorDevueltoFormGrid(Text1(26), CadenaDevuelta, 1)
        CadB = AUx
        
        AUx = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 2)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & AUx
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    'Cuentas
    SQL = RecuperaValor(CadenaSeleccion, 3)
    If SQL <> "" Then
        'Cuenta bloqueada
        If Text1(1).Text <> "" Then 'Hay fecha
            SQL = RecuperaValor(CadenaSeleccion, 1)
            If EstaLaCuentaBloqueada(SQL, CDate(Text1(1).Text)) Then
                MsgBox "Cuenta bloqueada: " & SQL, vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    Select Case cmdAux(0).Tag
    Case 2, 5
        'Cuenta normal
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(0).Text = RecuperaValor(CadenaSeleccion, 2)
    Case 3, 6
        Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(4).Text = RecuperaValor(CadenaSeleccion, 2)
    Case 100
        txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux(1).Text = RecuperaValor(CadenaSeleccion, 2)
    End Select
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
'Centro de coste
txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1)
txtAux(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCo_DatoSeleccionado(CadenaSeleccion As String)
Dim B As Boolean
If Text1(0).Text = "" Then
     MsgBox "No hay fecha seleccionada ", vbExclamation
     Exit Sub
End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(Linfac).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmI_DatoSeleccionado(CadenaSeleccion As String)
    'Solo me interesa el codigo
    i = CInt(AUx - 2)
    Text1(((i) * 6) + 1).Text = RecuperaValor(CadenaSeleccion, 1)
    If PonerValoresIva(i) Then
        CalcularIVA i
        TotalesRecargo
        TotalesIVA
        TotalFactura
    End If
    
End Sub

Private Sub imgppal_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0, 1, 7
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        Linfac = Index
        If Index = 7 Then
            Linfac = 29
        Else
            Linfac = Index
        End If
        If Text1(Linfac).Text <> "" Then frmF.Fecha = CDate(Text1(Linfac).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        Linfac = 0
    Case 2, 6, 100
        cmdAux(0).Tag = Index
        'Cliente y cta retencion
        Set frmC = New frmColCtas
        frmC.DatosADevolverBusqueda = "0|1|"
        frmC.ConfigurarBalances = 3
        frmC.Show vbModal
        Set frmC = Nothing
        
        'Lo vuelvo a posicionar ande toca
        If Index = 100 Then txtAux_LostFocus 0
        
    Case 3, 4, 5
        AUx = Index
        Set frmI = New frmIVA
        frmI.DatosADevolverBusqueda = "0|1|"
        frmI.Show vbModal
        Set frmI = Nothing
    End Select
    Screen.MousePointer = vbDefault
End Sub





Private Sub Label4_DblClick()
    If Label4.Tag <> "" Then
        If Text1(4).Text = "" Then
            Label4.Tag = InputBox("NA:")
            If Val(Label4.Tag) > 0 Then Text1(4).Text = Val(Label4.Tag)
        End If
        Label4.Tag = ""
    End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Shift = 1 Then
            Label4.Tag = "OK"
        End If
    End If
End Sub

Private Sub mnActual_Click()
    PonerFiltro 2
End Sub

Private Sub mnActuralySiguiente_Click()
    PonerFiltro 1
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub



Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    'Condiciones para NO salir
    If Modo = 5 Then Exit Sub
        
    PulsadoSalir = True
    Screen.MousePointer = vbHourglass
    DataGrid1.Enabled = False
    Unload Me
End Sub

Private Sub mnSiguiente_Click()
    PonerFiltro 3
End Sub

Private Sub mnSinFiltro_Click()
    PonerFiltro 0
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
        Else
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
            AntiguoText1 = Text1(Index).Text
    End If
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Han pulsado F1
    If KeyCode = 112 Then
        Text1_LostFocus Index
        cmdAceptar_Click
        
    Else
        If (Shift And vbCtrlMask) > 0 Then
            If UCase(Chr(KeyCode)) = "B" Then
                LanzaPantalla Index
            End If
        End If
    End If
End Sub
Private Sub LanzaPantalla(Index As Integer)
Dim miI As Integer
        '----------------------------------------------------
        '----------------------------------------------------
        '
        ' Dependiendo de index lanzaremos una opcion uotra
        '
        '----------------------------------------------------
        
        'De momento solo para el 5. Cliente
        miI = -1
        Select Case Index
        Case 0, 1
            'FECHA
            Text1(Index).Text = ""
            miI = Index
        Case 5
            Text1(5).Text = ""
            miI = 2
                
        Case 3
            Text1(3).Text = ""
            miI = 6
            
        Case 7, 13, 19
            Text1(Index).Text = ""
            If Index = 7 Then
                miI = 3
            Else
                If Index = 13 Then
                    miI = 4
                Else
                    miI = 5
                End If
            End If
                
        End Select
        If miI >= 0 Then imgppal_Click miI
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo <> 1 Then
        If KeyCode = 107 Or KeyCode = 187 Then
                KeyCode = 0
                LanzaPantalla Index
        End If
    End If
End Sub
'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)

Text1(Index).BackColor = vbWhite
'En AntiguoText1 tenemos el valor anterior
If Modo = 3 Or Modo = 4 Then
    PierdeFoco3 Index
    
      
    'Ahora, si no ha pueto Base2 lo pasamos a retencion
    'o si no pone retencion lo pasamos a boton aceptar
    If Index = 12 Then
        If Text1(12).Text = "" Then PonerFoco Text1(24)
    Else
        If Index = 24 Then
            If Text1(24).Text = "" Then PonerFoco cmdAceptar
        End If
    End If
    
Else
    If Modo = 1 Then
        If Index = 5 Or Index = 3 Then PierdeFoco3 Index
    End If

End If
End Sub


'Para cuando piede foco y estamos insertando o modificando
Private Sub PierdeFoco3(indice As Integer)
Dim RC As String
Dim Correcto As Boolean
Dim Valor As Currency
Dim L As Integer
Dim J As Integer
    Text1(indice).Text = Trim(Text1(indice).Text)
    If Text1(indice).Text = "" Then
        'Hemos puesto a blancos el campo, luego quitaremos
        'los valores asociados a el
        If Text1(indice) = AntiguoText1 Then Exit Sub
        Select Case indice
        Case 1
            'Ponemos a blanco tb el año de factura
            Text1(26).Text = ""
        Case 6 To 23
               
            If indice < 12 Then
                'PRIMERA LINEA
                L = 1
                'Numero de campo k ocupa
                i = indice - 6
            Else
                If indice < 18 Then
                    L = 2
                    i = indice - 12
                Else
                    L = 3
                    i = indice - 18
                End If
            End If
            
            'Ponemos IVA
            If i = 1 Then
                'Ha puesto a blanco el IVA. Borarmos el resto de campos
                J = (L * 6) + 5
                Text4(L).Text = ""
                For J = indice To J
                    Text1(J).Text = ""
                Next J
            End If
            'Ha cambiado la base o el iva. Luego hay k recalcular valores
            If i < 2 Then CalcularIVA CInt(L)
            TotalesRecargo
            TotalesIVA
            TotalesBase
            TotalFactura
        
        
        Case 3
            Text4(4).Text = ""
        Case 5
            Text4(0).Text = ""
        Case 24
            Text2(3).Text = ""
            TotalFactura
        End Select
    Else
        With Text1(indice)
            
           Select Case indice
           Case 1, 29
                If (Text1(indice).Text = AntiguoText1) And (Text1(26).Text <> "") Then Exit Sub
                If Not EsFechaOK(Text1(indice)) Then
                    MsgBox "Fecha incorrecta: " & .Text, vbExclamation
                    .Text = ""
                    Text1(indice).Text = ""
                    PonerFoco Text1(indice)
                    Exit Sub
                End If
                
                'Hay que comprobar que las fechas estan
                'en los ejercicios y si
                '       0 .- Año actual
                '       1 .- Siguiente
                '       2 .- Anterior al inicio
                '       3 .- Posterior al fin
                ModificandoLineas = FechaCorrecta2(CDate(.Text))
                If ModificandoLineas > 1 Then
                    If ModificandoLineas = 2 Then
                        RC = varTxtFec
                    Else
                        If ModificandoLineas = 3 Then
                            RC = "ya esta cerrado"
                        Else
                            RC = " todavia no ha sido abierto"
                        End If
                        RC = "La fecha pertenece a un ejercicio que " & RC
                    End If
                    MsgBox RC, vbExclamation
                    .Text = ""
                    PonerFoco Text1(indice)
                    Exit Sub
                End If
                
                
                .Text = Format(.Text, "dd/mm/yyyy")
                If indice = 1 Then Text1(26).Text = Year(CDate(.Text))
                
                'Si la fecha factura esta vacia entonces pongo esta
                If Text1(0).Text = "" Then Text1(0).Text = .Text
                
                'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
                'que el periodo de liquidacion del IVA no ha pasado.
                
                'ATENCION###
                'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
                'que el periodo de liquidacion del IVA no ha pasado.
                i = 0
                If vParam.Constructoras Then
                    If indice = 29 Then i = 1
                Else
                    If indice = 1 Then i = 1
                End If
                If i > 0 Then
                    If Not ComprobarPeriodo(indice) Then PonerFoco Text1(indice)
                End If
                
                
                
                
            Case 0
                If Not EsFechaOK(Text1(indice)) Then
                    MsgBox "Fecha incorrecta: " & .Text, vbExclamation
                    .Text = ""
                    PonerFoco Text1(indice)
                    Exit Sub
                End If
            Case 3, 5
                'Cuenta cliente
                If AntiguoText1 = .Text Then Exit Sub
                RC = .Text
                If indice = 3 Then
                    i = 4
                    Else
                    i = 0
                End If
                If CuentaCorrectaUltimoNivel(RC, SQL) Then
                    .Text = RC
                    Text4(i).Text = SQL
                    If Text1(1).Text <> "" Then
                        If Modo > 2 Then
                            If EstaLaCuentaBloqueada(RC, CDate(Text1(1).Text)) Then
                                MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                                .Text = ""
                                Text4(i).Text = ""
                            End If
                        End If
                    End If

                    
                    
                    RC = ""
                Else
                    'Si es k no existe la cuenta preguntaremos si kiere insertar
                    If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                            RC = RellenaCodigoCuenta(Text1(indice).Text)
                            SQL = "La cuenta: " & RC & " no existe. ¿Desea crearla?"
                            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                                CadenaDesdeOtroForm = RC
                                cmdAux(0).Tag = indice
                                Set frmC = New frmColCtas
                                frmC.DatosADevolverBusqueda = "0|1|"
                                frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                                frmC.Show vbModal
                                Set frmC = Nothing
                                If Text1(5).Text = RC Then SQL = "" 'Para k no los borre
                            End If
                    Else
                        'Cualquier otro error
                        'menos si no estamos buscando, k dejaremos
                        If Modo = 1 Then
                            SQL = ""
                        Else
                            MsgBox SQL, vbExclamation
                        End If
                    End If
                    
                    If SQL <> "" Then
                        .Text = ""
                        Text4(i).Text = ""
                        PonerFoco Text1(indice)
                    End If
                End If
                
            Case 7, 13, 19  'TIpos de iva
                i = ((indice - 1) / 6)
                'If Not IsNumeric(.Text) Then
                If Not EsNumerico(.Text) Then
                    MsgBox "Tipo de iva " & i & " incorrecto:  " & .Text
                    .Text = ""
                    Text4(i).Text = ""
                    PonerFoco Text1(indice)
                    Exit Sub
                End If
                If .Text = AntiguoText1 Then Exit Sub
                If PonerValoresIva(i) Then
                    CalcularIVA i
                    TotalesRecargo
                    TotalesIVA
                    TotalesBase
                    TotalFactura
                End If
            Case 6, 12, 18
                'BASES IMPONIBLES
                Correcto = True
                i = ((indice) / 6)
                'If Not IsNumeric(.Text) Then
                If Not EsNumerico(.Text) Then
                    'MsgBox "Importe debe de ser numérico: " & .Text, vbExclamation
                    .Text = ""
                    Correcto = False
                Else
                    If InStr(1, .Text, ",") > 0 Then
                        Valor = ImporteFormateado(.Text)
                    Else
                        Valor = CCur(TransformaPuntosComas(.Text))
                    End If
                    .Text = Format(Valor, "###,###,###,##0.00")
                    If .Text = AntiguoText1 Then Exit Sub
                End If
                'Recalculamos iva
                CalcularIVA i
                TotalesRecargo
                TotalesIVA
                TotalesBase
                TotalFactura
                If Not Correcto Then PonerFoco Text1(indice)
                
            Case 9, 15, 21
                If Not EsNumerico(.Text) Then
                    MsgBox "% de recargo debe de ser numérico: " & .Text, vbExclamation
                    .Text = ""
                Else
                    If InStr(1, .Text, ",") > 0 Then
                        Valor = ImporteFormateado(.Text)
                    Else
                        Valor = CCur(TransformaPuntosComas(.Text))
                    End If
                    .Text = Format(Valor, FormatoImporte)
                End If
                If .Text = AntiguoText1 Then Exit Sub
                TotalesRecargo
                TotalesIVA
                TotalFactura
            Case 24
                'If Not IsNumeric(.Text) Then
                If Not EsNumerico(.Text) Then
                    'MsgBox "% de recargo debe de ser numérico: " & .Text, vbExclamation
                    .Text = ""
                Else
                    If InStr(1, .Text, ",") > 0 Then
                        Valor = ImporteFormateado(.Text)
                    Else
                        Valor = CCur(TransformaPuntosComas(.Text))
                    End If
                    .Text = Format(Valor, "#0.00")
                End If
                If .Text = AntiguoText1 Then Exit Sub
                If Valor = 0 Then
                    .Text = ""
                    Text2(3).Text = ""
                Else
                    Base = ImporteFormateado(Text2(0).Text)
                    If Base = 0 Then
                        Text2(3).Text = ""
                    Else
                        Base = Round(Base * (Valor / 100), 2)
                        Text2(3).Text = Format(Base, FormatoImporte)
                    End If
                    TotalFactura
                End If
            End Select
        End With
End If


End Sub



Private Sub HacerBusqueda()
    Dim CadB As String
    CadB = ObtenerBusqueda(Me)
    
    SQL = AnyadeCadenaFiltro
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
        Else
            'Se muestran en el mismo form
            If CadB <> "" Then
                If SQL <> "" Then SQL = " AND (" & SQL & ")"
                CadB = CadB & SQL
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda False
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(1), 16, "Recepción: ")
        Cad = Cad & ParaGrid(Text1(26), 8, "Año: ")
        Cad = Cad & ParaGrid(Text1(2), 15, "Nº registro")
        Cad = Cad & ParaGrid(Text1(28), 15, "Nº factura")
        Cad = Cad & ParaGrid(Text1(0), 15, "Fecha fac:")
        Cad = Cad & ParaGrid(Text1(27), 15, "Total:")
        Cad = Cad & ParaGrid(Text1(5), 17, "Proveedor")
        
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NombreTabla
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "1|2|"
            frmB.vTitulo = "Facturas proveedores"
            frmB.vSelElem = 1
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                'If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
               ' Text1(kCampo).SetFocus
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda(Insertando As Boolean)
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Insertando Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla facturas proveedores con estos parámetros.", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
        Else
            PonerModo 2
            'Data1.Recordset.MoveLast
            Data1.Recordset.MoveFirst
            PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
        MuestraError Err.Number, "PonerCadenaBusqueda"
        PonerModo 0
        Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Por si modifica factura
    Numasien2 = -1
    NumDiario = 0
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True
    
    'En SQL almacenamos el importe
    Base = DBLet(Data1.Recordset!totfacpr)
'    If Not IsNull(Data1.Recordset!trefaccl) Then
'        Base = Base + Data1.Recordset!trefaccl
'    End If
    SQL = Base
    'Cargamos datos extras
    TotalesBase
    TotalesIVA
    TotalesRecargo
    TotalFactura
    If SQL <> CStr(AUx) Then
         MsgBox "Importe factura distinto Importe calculado: " & SQL & " - " & CStr(AUx), vbExclamation
    End If
    
    'Cliente
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(5).Text, "T")
    Text4(0).Text = SQL
    
    'IVAS
    For i = 1 To 3
        kCampo = (i * 6) + 1
        If Text1(kCampo).Text <> "" Then
            SQL = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", Text1(kCampo).Text, "N")
        Else
            SQL = ""
        End If
        Text4(i).Text = SQL
    Next i
    
    'Retencion
    If Text1(3).Text <> "" Then
        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(3).Text, "T")
    Else
        SQL = ""
    End If
    Text4(4).Text = SQL
        
        
    If Modo = 2 Then lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean
    If Modo = 1 Or Modo = 4 Then
        'Reestablecer colores
        For i = 0 To Text1.Count - 1
            Text1(i).BackColor = vbWhite
            Text1(i).Enabled = True
        Next i
        Text1(2).Enabled = False
        Text1(2).BackColor = &HFEF7E4
        
        For i = 3 To 5
            imgppal(i).Enabled = True
        Next i
        imgppal(0).Enabled = True
        
    Else

    End If
    
    Text1(4).Enabled = (Kmodo = 1)

    
    If Modo = 5 And Kmodo <> 5 Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 3
        Toolbar1.Buttons(6).ToolTipText = "Nuevo factura"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 4
        Toolbar1.Buttons(7).ToolTipText = "Modificar factura"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 5
        Toolbar1.Buttons(8).ToolTipText = "Eliminar factura"
    End If
    
    'ASIGNAR MODO
    Modo = Kmodo
    If Modo = 1 Then
        Text1(2).Enabled = True
        Text1(2).BackColor = vbWhite
    End If
    If Modo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
        Toolbar1.Buttons(6).Image = 12
        Toolbar1.Buttons(6).ToolTipText = "Nueva linea factura"
        '-- Modificar
        Toolbar1.Buttons(7).Image = 13
        Toolbar1.Buttons(7).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(8).Image = 14
        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea factura"
    End If
    B = (Modo < 5)
    chkVistaPrevia.Visible = B

    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B  'Lineas factur
    Toolbar1.Buttons(11).Enabled = B


    'El boton de vto sera enable si
    If vEmpresa.TieneTesoreria Then
        Toolbar1.Buttons(12).Enabled = B And vUsu.Nivel < 3
    End If



        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    cmdAceptar.Visible = B Or Modo = 1
   
    
    Me.framecabeceras.Enabled = B Or Modo = 1
    'Si es modiifcar y de periodo CERRADO
    If Modo = 4 Then
        If ModificaFacturaPeriodoLiquidado Then HabilitarTXTCabecerasAlModificar True
    End If
    

    '
    B = B Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    mnOpcionesAsiPre.Enabled = Not B
   
   
    If Modo = 1 Then
        Text2(4).Tag = "Importe|N|N|||cabfactprov|totfacpr|#,##0.00||"
    Else
        Text2(4).Tag = ""
    End If
    Text2(4).Locked = Not (Modo = 1)
   
   

    'El text
    B = (Modo = 2) Or (Modo = 5)
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B

   
   
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    
    B = Modo > 2 Or Modo = 1
    cmdCancelar.Visible = B
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    PonerOpcionesMenuGeneral Me
    
    
    'MAXLENGTH
    If Modo < 2 Then
        i = 0
    Else
        i = 10
    End If
    Text1(0).MaxLength = i
    Text1(1).MaxLength = i
    Text1(2).MaxLength = i
    Text1(5).MaxLength = i
    Text1(28).MaxLength = i
End Sub


Private Function DatosOK() As Boolean
'    Dim RS As ADODB.Recordset
    Dim B As Boolean
    
    
      
    'Si no es constructoras igualamos los campos fecfac y fecliquidacion
    If Not vParam.Constructoras Then Text1(29).Text = Text1(1).Text

    'Una comprobacion mas
    
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Len(Me.Text1(5).Text) <> vEmpresa.DigitosUltimoNivel Then
        MsgBox "Proveedor incorrecto", vbExclamation
        Exit Function
    End If
    
    
    'No pude tener Base imponible sin IVA
   If ((Text1(6).Text = "") Xor (Text1(7).Text = "")) Then
        B = False
        
   Else
    
            If Text1(7).Text = "" Then
                'Ningun campo puede estar puesto
                If ((Text1(9).Text <> "") Or (Text1(10).Text <> "") Or (Text1(11).Text <> "")) Then
                    MsgBox "Datos de IVA (1) sin poner el tipo", vbExclamation
                    Exit Function
                End If
            End If
    End If

   If ((Text1(12).Text = "") Xor (Text1(13).Text = "")) Then
        B = False
        
   Else
    
            If Text1(13).Text = "" Then
                'Ningun campo puede estar puesto
                If ((Text1(15).Text <> "") Or (Text1(16).Text <> "") Or (Text1(17).Text <> "")) Then
                    MsgBox "Datos de IVA (2) sin poner el tipo", vbExclamation
                    Exit Function
                End If
            End If
    End If

   If ((Text1(18).Text = "") Xor (Text1(19).Text = "")) Then
        B = False
    Else
    
            If Text1(19).Text = "" Then
                'Ningun campo puede estar puesto
                If ((Text1(21).Text <> "") Or (Text1(22).Text <> "") Or (Text1(23).Text <> "")) Then
                    MsgBox "Datos de IVA (3) sin poner el tipo", vbExclamation
                    Exit Function
                End If
            End If
    End If

        
   If Not B Then
        MsgBox "No puede tener base imponible sin iva.", vbExclamation
        Exit Function
    End If
    
    'No puede tener % de retencion sin cuenta de retencion
    If ((Text1(24).Text = "") Xor (Text1(3).Text = "")) Then
        MsgBox "No hay porcentaje de rentencion sin cuenta de retencion", vbExclamation
        B = False
        Exit Function
    End If
    
    If ((Text1(24).Text = "") Xor (Text2(3).Text = "")) Then
        MsgBox "No hay porcentaje de rentencion sin importe de retencion", vbExclamation
        B = False
        Exit Function
    End If
    
    
    
    'Cuando inserto una factura compruebo que para el año y cuenta
    'no existe la factura
    If Modo = 3 Then
        If Not ComprobarNuevaFactura Then Exit Function
    End If
    
    
    'Compruebo si hay fechas bloqueadas
    If vParam.CuentasBloqueadas <> "" Then
        If EstaLaCuentaBloqueada(Text1(5).Text, CDate(Text1(1).Text)) Then
            MsgBox "Cuenta bloqueada: " & Text1(5).Text, vbExclamation
            B = False
            Exit Function
        End If
        If Text1(3).Text <> "" Then
            If EstaLaCuentaBloqueada(Text1(3).Text, CDate(Text1(1).Text)) Then
                MsgBox "Cuenta bloqueada: " & Text1(3).Text, vbExclamation
                B = False
                Exit Function
            End If
        End If
    End If
    
    
    
    
    
    'Ahora. Si estamos modificando, y el año factura NO es el mismo, entonces
    'la estamos liando, y para evitar lios, NO dejo este tipo de modificacion
    If Modo = 4 Then
        If CDate(Text1(1).Text) <> Data1.Recordset!fecrecpr Then
            'HAN CAMBIADO LA FECHA. Veremos si dejo
            If Year(CDate(Text1(1).Text)) <> Data1.Recordset!anofacpr Then
                MsgBox "No puede cambiar de año la factura. ", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    
    
    DatosOK = B
End Function



Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim Valor As Currency
If Index = 3 Then
    'Voy a dejar modificar el importe de la retencion
    If Modo > 2 Then
        With Text2(3)
            .Text = Trim(.Text)
        
            If .Text <> "" Then
                If Not EsNumerico(.Text) Then
                    'MsgBox "Importe debe de ser numérico: " & .Text, vbExclamation
                    .Text = ""
                    PonerFoco Text2(3)
                Else
                    If InStr(1, .Text, ",") > 0 Then
                        Valor = ImporteFormateado(.Text)
                    Else
                        Valor = CCur(TransformaPuntosComas(.Text))
                    End If
                    .Text = Format(Valor, "###,###,###,##0.00")
                    
                End If
            End If
            TotalFactura
        End With
    End If
End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolbar1 Button.Index, False
End Sub





Private Sub HacerToolbar1(Indi As Integer, EsNuevaFactura As Boolean)
Dim N As Long
    Select Case Indi
    Case 1
        BotonBuscar
    Case 2
        BotonVerTodos
    Case 6
        If Modo <> 5 Then
            cmdCancelar.Caption = "Cancelar"
            BotonAnyadir
        Else
            'AÑADIR linea factura
            AnyadirLinea True, True
        End If
    Case 7
        If Modo <> 5 Then
            BotonModificar
        Else
            'MODIFICAR linea factura
            ModificarLinea
        End If
    Case 8
        If Modo <> 5 Then
            'If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
             'Modificar lineas
            varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
            If varFecOk = 2 Or varFecOk = 3 Then
                If varFecOk = 2 Then
                    MsgBox varTxtFec, vbExclamation
                Else
                    MsgBox "La factura pertenece a un ejercicio cerrado.", vbExclamation
                End If
                Exit Sub
            End If
            
            If Not ComprobarPeriodo(29) Then Exit Sub
            
            If BloqueaRegistroForm(Me) Then BotonEliminar
        Else
            'ELIMINAR linea factura
            EliminarLineaFactura
        End If
    Case 10
        
            varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
            If varFecOk = 2 Or varFecOk = 3 Then
                If varFecOk = 2 Then
                    MsgBox varTxtFec, vbExclamation
                Else
                    MsgBox "La factura pertenece a un ejercicio cerrado.", vbExclamation
                End If
                Exit Sub
            End If
        
        
        If Not ComprobarPeriodo(29) Then Exit Sub
        
        If Numasien2 > 0 Then
            espera 0.2
            CargaGrid True
            espera 0.1
        End If
        'Comprobamos que no esta actualizada ya
        If Not IsNull(Data1.Recordset!NumAsien) Then
            N = Data1.Recordset!NumAsien
            If N = 0 Then
                MsgBox "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
                Exit Sub
            End If
            
            
            SQL = "Esta factura ya esta contabilizada. Desea desactualizar para poder modificarla?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            Numasien2 = Data1.Recordset!NumAsien
            NumDiario = Data1.Recordset!NumDiari
            'Tengo desintegre la factura del hco
            If Not Desintegrar Then Exit Sub
            ObtenerSigueinteNumeroLinea
            Text1(4).Text = ""
        End If
            
        
   
        'If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
        If Not BloqueaRegistroForm(Me) Then Exit Sub
        
        PonerModo 5
        ModificandoLineas = 0
        'Si tiene numasien es k kiere modificar algo, luego se lo sugiero
        If Numasien2 > 0 Then
            If adodc1.Recordset.RecordCount = 1 Then ModificarLinea
        End If

        'Fuerzo que se vean las lineas
        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
    Case 11
        If Data1.Recordset.EOF Then Exit Sub
        
        If Not IsNull(Data1.Recordset!NumAsien) Then
            MsgBox "La factura ya esta contabilizada.", vbExclamation
            Exit Sub
        End If
        
        varFecOk = FechaCorrecta2(Data1.Recordset!fecrecpr)
        If varFecOk >= 2 Then
            If varFecOk = 2 Then
                MsgBox varTxtFec, vbExclamation
            Else
                MsgBox "No se puede contabilizar con esta fecha.", vbExclamation
            End If
            Exit Sub
        End If
        
        
                    
        If FacturaContabilizada("", Text1(2).Text, Text1(26).Text) Then
            MsgBox "Factura ya contabilizada(Step: 2). ", vbExclamation
            Exit Sub
        End If
                        
            
            
        
        
        If Not EsNuevaFactura Then
            SQL = "Va a contabilizar la factura" & vbCrLf & vbCrLf & "Numero:  " & _
                 Data1.Recordset!NumRegis & "       Fecha: " & Data1.Recordset!fecrecpr
            SQL = SQL & vbCrLf & vbCrLf & "     ¿Desea continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
            'Bloqueamos
            If Not BloqueaRegistroForm(Me) Then Exit Sub
        Else
            'Nueva factura
            'Estamos contabilizando automaticamente
        End If
        
        SQL = ""
        Screen.MousePointer = vbHourglass
        'Actualizar
        If IntegrarFactura Then
            If Data1.Recordset.EOF Then
                LimpiarCampos
                CargaGrid False
                PonerModo 0
            Else
                PonerCampos
                PonerModo 2
            End If
        End If
        If Not EsNuevaFactura Then DesBloqueaRegistroForm Text1(0)

        Screen.MousePointer = vbDefault
        
        
    Case 12
        If Data1.Recordset.EOF Then Exit Sub
        
        'Pongo la fecha de RECEPCION
        varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
        SQL = ""
        If varFecOk >= 2 Then
            If varFecOk = 2 Then
                SQL = varTxtFec
            Else
                SQL = "Fecha factura fuera de ejercicios."
            End If
            SQL = SQL & vbCrLf & "¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        SQL = ""



        frmVto.opcion = 1
        frmVto.Importe = Data1.Recordset!totfacpr
        'frmVto.Datos = "|AF23/1234|2005|10/06/05|4000001|Ariadna GIL|"
        frmVto.Datos = "|" & Data1.Recordset!numfacpr & "||" & Data1.Recordset!fecfacpr & "|" & Data1.Recordset!codmacta & "|" & Text4(0).Text & "|"
        frmVto.Show vbModal



    Case 13
        frmListado.opcion = 13
        frmListado.Show vbModal
    Case 14
        If Modo = 4 Or Modo = 3 Then If MsgBox("Esta editando la factura. ¿Salir?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        mnSalir_Click
    Case 16 To 19
        Desplazamiento (Indi - 16)
    Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    For i = 16 To 19
        Toolbar1.Buttons(i).Enabled = bol
        Toolbar1.Buttons(i).Visible = bol
    Next i
End Sub



Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = MontaSQLCarga(Enlaza)
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockPessimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    DataGrid1.Tag = "Asignando"
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-numaspre  1.- Lin aspre
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    
    DataGrid1.Columns(0).Caption = "Cuenta"
    DataGrid1.Columns(0).Width = 1040
    
    DataGrid1.Columns(1).Caption = "Título"
    DataGrid1.Columns(1).Width = 3300

    'Cuenta
    If vParam.autocoste Then
        DataGrid1.Columns(2).Caption = "C.C."
        DataGrid1.Columns(2).Width = 680
    
        DataGrid1.Columns(3).Caption = "centro coste"
        DataGrid1.Columns(3).Width = 2345
        anc = 0
        Else
        DataGrid1.Columns(2).Visible = False
        DataGrid1.Columns(3).Visible = False
        ' e incrementamos el ancho de la columna 1
        anc = 3025
    End If
    DataGrid1.Columns(1).Width = DataGrid1.Columns(1).Width + anc
    
    DataGrid1.Columns(4).Caption = "Importe"
    DataGrid1.Columns(4).Width = 2000
    DataGrid1.Columns(4).NumberFormat = FormatoImporte
    DataGrid1.Columns(4).Alignment = dbgRight
    

    DataGrid1.Columns(5).Visible = False   'nº linea
    'Fiajamos el cadancho
    If Not CadAncho Then
        DataGrid1.Tag = "Fijando ancho"
        anc = DataGrid1.Left
        txtAux(0).Left = anc + 330
        txtAux(0).Width = DataGrid1.Columns(0).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = anc + DataGrid1.Columns(1).Left
                
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width
        txtAux(1).Width = DataGrid1.Columns(1).Width - cmdAux(0).Width - 30
        
        If vParam.autocoste Then
            txtAux(2).Left = anc + DataGrid1.Columns(2).Left + 30
            txtAux(2).Width = DataGrid1.Columns(2).Width - 20
        
            cmdAux(1).Left = anc + DataGrid1.Columns(3).Left
            
            txtAux(3).Left = cmdAux(1).Left + cmdAux(1).Width
            txtAux(3).Width = DataGrid1.Columns(3).Width - cmdAux(0).Width - 30
        End If
           
        txtAux(4).Left = anc + DataGrid1.Columns(4).Left + 30
        txtAux(4).Width = DataGrid1.Columns(4).Width - 30
        
        
        If vParam.autocoste Then
            cmdAux(1).Visible = False
        
        End If
        CadAncho = True
    End If
        
    For i = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(i).AllowSizing = False
    Next i
   
'    For i = 0 To txtaux.Count - 1
'        txtaux(i).Visible = True
'        txtaux(i).Top = 6000
'    Next i
'    cmdAux(0).Top = 6000
'    cmdAux(0).Visible = True
    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub

Private Function MontaSQLCarga(Enlaza As Boolean) As String
    '--------------------------------------------------------------------
    ' MontaSQlCarga:
    '   Basándose en la información proporcionada por el vector de campos
    '   crea un SQl para ejecutar una consulta sobre la base de datos que los
    '   devuelva.
    ' Si ENLAZA -> Enlaza con el data1
    '           -> Si no lo cargamos sin enlazar a nngun campo
    '--------------------------------------------------------------------
    Dim SQL As String
    
    SQL = "SELECT linfactprov.codtbase, cuentas.nommacta, linfactprov.codccost, ccoste.nomccost, linfactprov.impbaspr, linfactprov.numlinea"
    SQL = SQL & " FROM (ccoste RIGHT JOIN linfactprov ON ccoste.codccost = linfactprov.codccost) INNER JOIN cuentas ON linfactprov.codtbase = cuentas.codmacta WHERE "
    If Enlaza Then
        SQL = SQL & " numregis = " & Data1.Recordset!NumRegis
        SQL = SQL & " AND anofacpr= " & Data1.Recordset!anofacpr
        Else
        SQL = SQL & " anofacpr = -1"
    End If
    SQL = SQL & " ORDER BY linfactprov.numlinea"
    MontaSQLCarga = SQL
End Function

Private Sub AnyadirLinea(Limpiar As Boolean, DesdeBoton As Boolean)
    Dim anc As Single
    Dim Preg As String
    
    If ModificandoLineas = 2 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    Linfac = ObtenerSigueinteNumeroLinea   'Fijamos en aux el importe que queda
    If AUx = 0 Then
        If DesdeBoton Then
            Preg = "La suma de las lineas coincide con el importe factura. ¿Seguro que desea insertar mas lineas?"
            If MsgBox(Preg, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
 
        Else
            LLamaLineas anc, 0, True
            cmdCancelar.Caption = "Cabecera"
            Exit Sub
        End If
    End If
    
    
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        adodc1.Recordset.MoveLast
        DataGrid1.HoldFields
       ' DataGrid1.Row = DataGrid1.Row + 1
    End If
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row + 1) + 15
    End If
    LLamaLineas anc, 1, Limpiar
    'Ponemos el importe
    
    txtAux(4).Text = AUx
    HabilitarCentroCoste
    'Ponemos el foco
    PonerFoco txtAux(0)
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    'If Not BloqueaRegistroForm(Me) Then Exit Sub
    
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    If vParam.autocoste Then
        txtAux(2).Text = DataGrid1.Columns(2).Text
        txtAux(3).Text = DataGrid1.Columns(3).Text
    End If
    txtAux(4).Text = adodc1.Recordset!impbaspr

    LLamaLineas anc, 2, False
    HabilitarCentroCoste
    PonerFoco txtAux(0)
End Sub

Private Sub EliminarLineaFactura()
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Lineas de factura." & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar la linea: " & vbCrLf
    SQL = SQL & adodc1.Recordset.Fields(0) & " - " & adodc1.Recordset.Fields(1) & ": " & adodc1.Recordset.Fields(4)
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = "Delete from linfactprov"
        SQL = SQL & " WHERE numlinea = " & adodc1.Recordset!NumLinea
        SQL = SQL & " AND anofacpr=" & Data1.Recordset!anofacpr
        SQL = SQL & " AND numregis = " & Data1.Recordset!NumRegis & ";"
        Conn.Execute SQL
        
        'LOG
        vLog.Insertar 8, vUsu, "Lin_e: " & Format(Data1.Recordset!NumRegis, "000000") & "  nº:" & adodc1.Recordset!NumLinea
        
        
        CargaGrid (Not Data1.Recordset.EOF)
    End If
End Sub


'Ademas de obtener el siguiente nº de linea, tb obtiene la suma de
'las lineas de factura, Y fijamos lo que falta en aux para luego ofertarlo

Private Function ObtenerSigueinteNumeroLinea() As Long
    Dim Rs As ADODB.Recordset
    Dim i As Long
    
    Set Rs = New ADODB.Recordset
    
    SQL = " WHERE linfactprov.numregis= " & Data1.Recordset!NumRegis
    SQL = SQL & " AND linfactprov.anofacpr=" & Data1.Recordset!anofacpr & ";"
    Rs.Open "SELECT Max(numlinea) FROM linfactprov" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    i = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then i = Rs.Fields(0)
    End If
    Rs.Close

    'La suma
    SumaLinea = 0
    If i > 0 Then
        Rs.Open "SELECT sum(impbaspr) FROM linfactprov" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then SumaLinea = Rs.Fields(0)
        End If
        Rs.Close
    End If
    Set Rs = Nothing
    
    'Lo que falta lo fijamos en aux. El importe es de la bASE IMPONIBLE si fuera del total seria Text2(4).Text
    AUx = ImporteFormateado(Text2(0).Text)
    AUx = AUx - SumaLinea
    ObtenerSigueinteNumeroLinea = i + 1
End Function




Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid
    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)
    cmdAceptar.Visible = Not B
    'cmdCancelar.Visible = Not b
    CamposAux Not B, alto, Limpiar
End Sub

Private Sub CamposAux(Visible As Boolean, Altura As Single, Limpiar As Boolean)
    
    DataGrid1.Enabled = Not Visible
    cmdAux(0).Visible = Visible
    cmdAux(0).Top = Altura
    If vParam.autocoste Then
        cmdAux(1).Visible = Visible
        txtAux(3).Visible = Visible
        txtAux(2).Visible = Visible
        cmdAux(1).Top = Altura
    Else
        txtAux(3).Visible = False
        txtAux(2).Visible = False
        txtAux(3).Text = ""
        txtAux(2).Text = ""
        cmdAux(1).Visible = False
    End If
    For i = 0 To txtAux.Count - 1
        If i < 2 Or i > 3 Then txtAux(i).Visible = Visible
        txtAux(i).Top = Altura
    Next i

    If Limpiar Then
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
With txtAux(Index)
    If Index <> 5 Then
         .SelStart = 0
        .SelLength = Len(.Text)
    Else
        .SelStart = Len(.Text)
    End If
End With

End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Sng As Double
    Dim CrearEltoInmov As Boolean
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtAux(Index).Text = Trim(txtAux(Index).Text)
    
        'Comun a todos
        If txtAux(Index).Text = "" Then
            Select Case Index
            Case 0
                txtAux(1).Text = ""
                HabilitarCentroCoste
            Case 2
                txtAux(3).Text = ""
            End Select
            Exit Sub
        End If
        
        Select Case Index
        Case 0
            'Cta
            CTA_Inmovilizado = ""
            RC = txtAux(0).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtAux(0).Text = RC
                txtAux(1).Text = SQL
                'Para el inmovilizado
                CTA_Inmovilizado = RC
                RC = ""
            Else
            
                SQL = "La cuenta no existe, desea crearla?"
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                        txtAux(0).Text = "" 'Pos si cancela
                        Set frmC = New frmColCtas
                        cmdAux(0).Tag = 100
                        CadenaDesdeOtroForm = RC
                        frmC.DatosADevolverBusqueda = "0|1|"
                        frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                        frmC.Show vbModal
                        Set frmC = Nothing
                        If txtAux(0).Text <> "" Then CTA_Inmovilizado = txtAux(0).Text
                End If
            
            End If
            
            
            If CTA_Inmovilizado <> "" Then
                CrearEltoInmov = False
                If vParam.NuevoPlanContable Then
                    If Mid(CTA_Inmovilizado, 1, 2) = "20" Or Mid(CTA_Inmovilizado, 1, 2) = "21" Then CrearEltoInmov = True
                Else
                    If Mid(CTA_Inmovilizado, 1, 2) = "21" Or Mid(CTA_Inmovilizado, 1, 2) = "22" Then CrearEltoInmov = True
                End If
                
                If CrearEltoInmov Then
                    SQL = "Desea crear un elemento de Inmovilizado ? "
                    RC = "NO"
                    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                        'Le pasaremos el codprove|nomprove|numfac|importe
                        'ANTES
                        'CadenaDesdeOtroForm = Text1(5).Text & "|" & Text4(0).Text & "|" & Text1(28).Text & "|" & Text1(0).Text & "|" & Text1(6).Text & "|"
                        CadenaDesdeOtroForm = Text1(5).Text & "|" & Text4(0).Text & "|" & Text1(28).Text & "|" & Text1(0).Text & "|" & Text2(0).Text & "|"
                        
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtAux(0).Text & "|" & txtAux(1).Text & "|"
                        frmInmoElto.Nuevo = CadenaDesdeOtroForm
                        CadenaDesdeOtroForm = ""
                        frmInmoElto.Show vbModal
 
                        Screen.MousePointer = vbDefault
                    End If
                End If
            End If
            
            HabilitarCentroCoste
            
            If CTA_Inmovilizado = "" Then
                txtAux(0).SetFocus
            Else
                If txtAux(2).Visible And txtAux(2).Enabled Then
                    txtAux(2).SetFocus
                Else
                    txtAux(4).SetFocus
                End If
            End If
        Case 2
            txtAux(2).Text = UCase(txtAux(2).Text)
            RC = "idsubcos"
            SQL = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtAux(2).Text, "T", RC)
            If SQL = "" Then
                MsgBox "Centro de coste no encontrado: " & txtAux(2).Text, vbExclamation
                txtAux(2).Text = ""
                txtAux(2).SetFocus
            End If
            txtAux(3).Text = SQL
            If SQL <> "" Then txtAux(4).SetFocus
        Case 4
            If Not EsNumerico(txtAux(4).Text) Then
                'MsgBox "Importe incorrecto: " & txtaux(4).Text, vbExclamation
                txtAux(4).Text = ""
                txtAux(4).SetFocus
            Else
                txtAux(4).Text = TransformaPuntosComas(txtAux(4).Text)
                'cmdAceptar.SetFocus
            End If
            
        End Select
End Sub


Private Function AuxOK() As String
    
    'Cuenta
    If txtAux(0).Text = "" Then
        AuxOK = "Cuenta no puede estar vacia."
        Exit Function
    End If
    If Len(txtAux(0).Text) <> vEmpresa.DigitosUltimoNivel Then
        AuxOK = "Longitud cuenta incorrecta"
        Exit Function
    End If
    If EstaLaCuentaBloqueada(txtAux(0).Text, CDate(Text1(1).Text)) Then
        AuxOK = "Cuenta bloqueada: " & txtAux(0).Text
        Exit Function
    End If
    
    'Importe
    If txtAux(4).Text = "" Then
        AuxOK = "El importe no puede estar vacio"
        Exit Function
    End If
        
    If txtAux(4).Text <> "" Then
        If Not IsNumeric(txtAux(4).Text) Then
            AuxOK = "El importe debe de ser numérico."
            Exit Function
        End If
    End If
    
    'cENTRO DE COSTE
    If txtAux(2).Visible Then
        If txtAux(2).Enabled Then
            If txtAux(2).Text = "" Then
                AuxOK = "Centro de coste no puede ser nulo"
                Exit Function
            End If
        End If
    End If
    
    
    'Vemos los importes
    '--------------------------
    'Total factura en AUX
    AUx = ImporteFormateado(Text2(4).Text)
    
    
    'Si tiene retencion
    AUX2 = 0
    If Text2(3).Text <> "" Then AUX2 = ImporteFormateado(Text2(3).Text)
    AUx = AUx + AUX2
    
    'El iVA
    AUX2 = 0
    If Text2(1).Text <> "" Then AUX2 = ImporteFormateado(Text2(1).Text)
    AUx = AUx - AUX2
    
    'La retencion
    AUX2 = 0
    If Text2(2).Text <> "" Then AUX2 = ImporteFormateado(Text2(2).Text)
    AUx = AUx - AUX2
    
    
    'Importe linea en aux2
    AUX2 = CCur(txtAux(4).Text)
    
    'Suma de las lineas mas lo que acabamos de poner
    AUX2 = AUX2 + SumaLinea
    
    'Si estamos insertando entonces la suma de lineas + aux2 no debe ser superior a toal fac
    If ModificandoLineas = 2 Then
        'Si estasmos insertando no hacemos nada puesto que el importe son las sumas directamente
       'Estamos modificando, hay que quitarle el importe que tenia
        AUX2 = AUX2 - adodc1.Recordset!impbaspr
    End If
'    If ModificandoLineas = 2 Then
'        If Aux > 0 Then
'            If AUX2 > Aux Then
'                    AuxOK = "El importe excede del total de factura"
'                    Exit Function
'            End If
'        Else
'            If AUX2 < Aux Then
'                    AuxOK = "El importe excede del total de factura"
'                    Exit Function
'            End If
'
'        End If
'    End If
    AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        SQL = "INSERT INTO linfactprov (numregis, anofacpr, numlinea, codtbase, impbaspr, codccost) VALUES ("
        ''R', 11, 2003, 1, '6000001', 1500, 'TIEN')
        SQL = SQL & Data1.Recordset!NumRegis & ","
        SQL = SQL & Data1.Recordset!anofacpr & "," & Linfac & ",'"
        'Cuenta
        SQL = SQL & txtAux(0).Text & "',"
        'Importe
        SQL = SQL & TransformaComasPuntos(txtAux(4).Text) & ","
   
        'Centro coste
        If txtAux(2).Text = "" Then
          SQL = SQL & ValorNulo
          Else
          SQL = SQL & "'" & txtAux(2).Text & "'"
        End If
        SQL = SQL & ")"
        
    Else
    
        'MODIFICAR
        'UPDATE asipre_lineas SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
        '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
        SQL = "UPDATE linfactprov SET "
        
        SQL = SQL & " codtbase = '" & txtAux(0).Text & "',"
        SQL = SQL & " impbaspr = "
        SQL = SQL & TransformaComasPuntos(txtAux(4).Text) & ","
        
        'Centro coste
        If txtAux(2).Text = "" Then
          SQL = SQL & " codccost = " & ValorNulo
          Else
          SQL = SQL & " codccost = '" & txtAux(2).Text & "'"
        End If
    
        SQL = SQL & " WHERE numregis= " & Data1.Recordset!NumRegis
        SQL = SQL & " AND anofacpr=" & Data1.Recordset!anofacpr
        SQL = SQL & " AND numlinea =" & adodc1.Recordset!NumLinea & ";"
        
        'LOG
        vLog.Insertar 8, vUsu, "Lin: " & Format(Data1.Recordset!NumRegis, "000000") & "  nº:" & adodc1.Recordset!NumLinea
        
    End If
    Conn.Execute SQL
    InsertarModificar = True
    Exit Function
EInsertarModificar:
        MuestraError Err.Number, "InsertarModificar linea asiento.", Err.Description
End Function
 


Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub CargaGrid(Enlaza As Boolean)
Dim B As Boolean
B = DataGrid1.Enabled
CargaGrid2 Enlaza
DataGrid1.Enabled = B
End Sub

Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub


Private Function PonerValoresIva(numero As Integer) As Boolean
Dim J As Integer
J = ((numero - 1) * 6) + 7
Set Rs = New ADODB.Recordset
Rs.Open "Select nombriva,porceiva,porcerec from tiposiva where codigiva =" & Text1(J).Text, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
If Rs.EOF Then
    MsgBox "Tipo de IVA incorrecto: " & Text1(J).Text, vbExclamation
    Text1(J).Text = ""
    Text4(numero).Text = ""
    PonerValoresIva = False
Else
    PonerValoresIva = True
    
    Text4(numero).Text = Rs.Fields(0)
    Text1(J + 1).Text = Format(Rs.Fields(1), "#0.00")
    AUx = DBLet(Rs.Fields(2), "N")
    If AUx = 0 Then
        Text1(J + 3).Text = ""
        Else
        Text1(J + 3).Text = Format(AUx, "#0.00")
    End If
End If
Rs.Close
End Function


'Indica en k fila estamos
Private Sub CalcularIVA(numero As Integer)
Dim J As Integer

On Error GoTo ECalIVa


    J = ((numero - 1) * 6) + 6
    Base = ImporteFormateado(Text1(J).Text)
    'EL iva
    AUx = ImporteFormateado(Text1(J + 2).Text) / 100
    If AUx = 0 Then
        If Text1(J + 2).Text = "" Then
            Text1(J + 3).Text = ""
        Else
            Text1(J + 3).Text = "0,00"
        End If
    Else
        Text1(J + 3).Text = Format(Round((AUx * Base), 2), FormatoImporte)
    End If
    
    'Recargo
    AUx = ImporteFormateado(Text1(J + 4).Text) / 100
    If AUx = 0 Then
        Text1(J + 5).Text = ""
    Else
        Text1(J + 5).Text = Format(Round((AUx * Base), 2), FormatoImporte)
    End If
    
Exit Sub
ECalIVa:
    Err.Clear
End Sub


Private Sub TotalesBase()
    'Base imponible
    AUx = 0
    For i = 1 To 3
        If Text1(i * 6).Text <> "" Then
            Base = ImporteFormateado(Text1(i * 6).Text)
            AUx = AUx + Base
        End If
    Next i
    If AUx = 0 Then
        Text2(0).Text = ""
    Else
        Text2(0).Text = Format(AUx, FormatoImporte)
    End If
End Sub


Private Sub TotalesIVA()
    'Total IVA
    AUx = 0
    For i = 1 To 3
        ancho = (i * 6) + 3
        If Text1(ancho).Text <> "" Then
            Base = ImporteFormateado(Text1(ancho).Text)
            AUx = AUx + Base
        End If
    Next i
    If AUx = 0 Then
        Text2(1).Text = ""
    Else
        Text2(1).Text = Format(AUx, FormatoImporte)
    End If
End Sub

Private Sub TotalesRecargo()
    'RECARGO
    AUx = 0
    For i = 1 To 3
        ancho = (i * 6) + 5
        If Text1(ancho).Text <> "" Then
            Base = ImporteFormateado(Text1(ancho).Text)
            AUx = AUx + Base
        End If
    Next i
    If AUx = 0 Then
        Text2(2).Text = ""
    Else
        Text2(2).Text = Format(AUx, FormatoImporte)
    End If
End Sub

Private Sub TotalFactura()
    'El total
    AUx = 0
    ' Base + iva + recargao   -  retencion
    For i = 0 To 2
        If Text2(i).Text = "" Then
   
        Else
            Base = ImporteFormateado(Text2(i).Text)
            AUx = AUx + Base
        End If
    Next i
    If Text2(3).Text = "" Then
        
    Else
        Base = ImporteFormateado(Text2(3).Text)
        AUx = AUx - Base
    End If
    
    If AUx = 0 Then
        Text2(4).Text = ""
    Else
        Text2(4).Text = Format(AUx, FormatoImporte)
    End If
    Text1(27).Text = TransformaComasPuntos(CStr(AUx))
End Sub


'Comprobara si el periodo esta liquidado o no.
'Si la fecha pertenece a un periodo liquidado entonces
'mostraremos un mensaje para preguntar si desea continuar o no
Private Function ComprobarPeriodo(indice As Integer) As Boolean
Dim Cerrado As Boolean

    
    'Primero pondremos la fecha a año periodo
    i = Year(CDate(Text1(indice).Text))
    If vParam.periodos = 0 Then
        'Trimestral
        ancho = ((Month(CDate(Text1(indice).Text)) - 1) \ 3) + 1
        Else
        ancho = Month(CDate((Text1(indice).Text)))
    End If
    Cerrado = False
    If i < vParam.anofactu Then
        Cerrado = True
    Else
        If i = vParam.anofactu Then
            'El mismo año. Comprobamos los periodos
            If vParam.perfactu >= ancho Then _
                Cerrado = True
        End If
    End If
    ComprobarPeriodo = True
    ModificaFacturaPeriodoLiquidado = False
    If Cerrado Then
        ModificaFacturaPeriodoLiquidado = True
        SQL = "La fecha "
        If indice = 0 Then
            SQL = SQL & "factura"
        Else
            SQL = SQL & "liquidacion"
        End If
        
        
        SQL = SQL & " corresponde a un periodo ya liquidado. " & vbCrLf
        SQL = SQL & vbCrLf & " ¿Desea continuar igualmente ?"
      
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then ComprobarPeriodo = False
    End If
        

End Function



Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtAux(0).Text <> "" Then
           hab = HayKHabilitarCentroCoste(txtAux(0).Text)
        Else
            txtAux(2).Text = ""
            txtAux(3).Text = ""
        End If
        If hab Then
            txtAux(2).BackColor = &H80000005
            Else
            txtAux(2).BackColor = &H80000018
        End If
    End If
    txtAux(2).Enabled = hab
    cmdAux(1).Enabled = hab
    Me.Refresh
End Sub




Private Function Desintegrar() As Boolean
        Desintegrar = False
        'Primero hay que desvincular la factura de la tabla de hco
        If DesvincularFactura Then
            frmActualizar.OpcionActualizar = 2  'Desactualizar para eliminar
            frmActualizar.NumAsiento = Data1.Recordset!NumAsien
            frmActualizar.FechaAsiento = Data1.Recordset!FechaEnt
            frmActualizar.NumDiari = Data1.Recordset!NumDiari
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            If AlgunAsientoActualizado Then Desintegrar = True
        End If
End Function





Private Sub LeerFiltro(Leer As Boolean)
    SQL = App.Path & "\filfacp.dat"
    If Leer Then
        FILTRO = 0
        If Dir(SQL) <> "" Then
            AbrirFicheroFiltro True
            If IsNumeric(SQL) Then FILTRO = CByte(SQL)
        End If
    Else
        AbrirFicheroFiltro False
    End If
End Sub


Private Sub AbrirFicheroFiltro(Leer As Boolean)
On Error GoTo eAbrir
    i = FreeFile
    If Leer Then
        Open SQL For Input As #i
        SQL = "0"
        Line Input #i, SQL
    Else
        Open SQL For Output As #i
        Print #i, FILTRO
    End If
    Close #i
    Exit Sub
eAbrir:
    Err.Clear
End Sub


Private Sub PonerFiltro(NumFilt As Byte)
    FILTRO = NumFilt
    Me.mnActual.Checked = (NumFilt = 2)
    Me.mnActuralySiguiente.Checked = (NumFilt = 1)
    Me.mnSiguiente.Checked = (NumFilt = 3)
    Me.mnSinfiltro.Checked = (NumFilt = 0)
End Sub


Private Function AnyadeCadenaFiltro() As String
Dim AUx As String

    AUx = ""
    If FILTRO <> 0 Then
        '-------------------------------- INICIO
        If FILTRO < 3 Then
            'INicio = actual
            AUx = " fecrecpr >='" & Format(vParam.fechaini, FormatoFecha) & "'"
        Else
            AUx = " fecrecpr >='" & Format(DateAdd("yyyy", 1, vParam.fechaini), FormatoFecha) & "'"
        End If
        
        If FILTRO = 2 Then
            AUx = AUx & " AND fecrecpr <='" & Format(vParam.fechafin, FormatoFecha) & "'"
        Else
            AUx = AUx & " AND fecrecpr <='" & Format(DateAdd("yyyy", 1, vParam.fechafin), FormatoFecha) & "'"
        End If
        
    End If  'filtro=0
    AnyadeCadenaFiltro = AUx
End Function


Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function ComprobarNuevaFactura() As Boolean
Dim C As String
    
    On Error GoTo EcomprobarNuevaFactura
    ComprobarNuevaFactura = False
    
    C = "Select numregis from cabfactprov where codmacta ='" & Text1(5).Text & "' and anofacpr =" & Year(CDate(Text1(1).Text))
    C = C & " AND numfacpr ='" & Text1(28).Text & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        C = "Ya existe la factura: " & Text1(28).Text & " para el este proveedor y año: " & Year(CDate(Text1(1).Text))
        C = C & vbCrLf & " ¿Continuar?"
        If MsgBox(C, vbQuestion + vbYesNo) = vbYes Then ComprobarNuevaFactura = True
    Else
        ComprobarNuevaFactura = True
    End If
    miRsAux.Close
    
    Set miRsAux = Nothing
    Exit Function
EcomprobarNuevaFactura:
    MuestraError Err.Number, "Comprobando factura / proveedor"
    Set miRsAux = Nothing
End Function



Private Sub HabilitarTXTCabecerasAlModificar(Preparando As Boolean)
Dim J As Integer

    'Si el usuario no tiene permisos le bloqueamos muchas cosas
    If vUsu.Nivel = 2 Then
        For J = 0 To 25
            'Solo dejamos enabled fecha, codclien, concepto
            'cta retencion.  Index: 0,5,25,3
            If Preparando Then
                If Not (J = 0 Or J = 5 Or J = 25 Or J = 3 Or J = 28) Then Text1(J).Enabled = False
            Else
                Text1(J).Enabled = True
            End If
        Next J
        
        If Preparando Then
            imgppal(0).Enabled = False
            For J = 3 To 5
                imgppal(J).Enabled = False
            Next J
        End If
        
    End If
End Sub
