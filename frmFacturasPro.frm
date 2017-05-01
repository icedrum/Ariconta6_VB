VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFacturasPro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   17655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDatosFiscales 
      Caption         =   "DATOS FISCALES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   390
      TabIndex        =   98
      Top             =   2520
      Visible         =   0   'False
      Width           =   16935
      Begin VB.TextBox Text4 
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
         Index           =   21
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text4"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text1 
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
         Left            =   1320
         TabIndex        =   105
         Tag             =   "País|T|S|||factpro|codpais|||"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text1 
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
         Left            =   10290
         TabIndex        =   100
         Tag             =   "Nif|T|S|||factpro|nifdatos|||"
         Top             =   390
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   10320
         TabIndex        =   104
         Tag             =   "Provincia|T|S|||factpro|desprovi|||"
         Top             =   1260
         Width           =   4020
      End
      Begin VB.TextBox Text1 
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
         Left            =   4020
         TabIndex        =   103
         Tag             =   "Poblacion|T|S|||factpro|despobla|||"
         Top             =   1260
         Width           =   5130
      End
      Begin VB.TextBox Text1 
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
         Left            =   1320
         TabIndex        =   102
         Tag             =   "CP|T|S|||factpro|codpobla|||"
         Top             =   1230
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   1320
         TabIndex        =   101
         Tag             =   "Dirección|T|S|||factpro|dirdatos|||"
         Top             =   810
         Width           =   7830
      End
      Begin VB.TextBox Text1 
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
         Left            =   1320
         TabIndex        =   99
         Tag             =   "Nombre|T|N|||factpro|nommacta|||"
         Top             =   390
         Width           =   7830
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   990
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "País"
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
         Height          =   255
         Index           =   18
         Left            =   300
         TabIndex        =   112
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
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
         Height          =   255
         Index           =   17
         Left            =   9330
         TabIndex        =   111
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
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
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   110
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "NIF"
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
         Height          =   255
         Index           =   15
         Left            =   9330
         TabIndex        =   109
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "C.Postal"
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
         Height          =   255
         Index           =   14
         Left            =   300
         TabIndex        =   108
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección"
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
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   107
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   106
         Top             =   450
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2145
      Left            =   9690
      TabIndex        =   84
      Top             =   4920
      Width           =   7725
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FCFCE2&
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
         Height          =   360
         Index           =   13
         Left            =   5640
         TabIndex        =   21
         Tag             =   "Total Factura|N|S|||factpro|totfacpr|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1590
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Index           =   12
         Left            =   5640
         TabIndex        =   20
         Tag             =   "Importe Retención|N|S|||factpro|trefacpr|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1050
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Left            =   1740
         TabIndex        =   19
         Tag             =   "Base Retención|N|S|||factpro|totbasesret|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Left            =   5640
         TabIndex        =   18
         Tag             =   "Importe Iva|N|S|||factpro|totivas|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   570
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Left            =   1740
         TabIndex        =   17
         Tag             =   "Base Imponible|N|S|||factpro|totbases|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   570
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
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
         Height          =   255
         Index           =   13
         Left            =   3780
         TabIndex        =   90
         Top             =   1650
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retención"
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
         Height          =   255
         Index           =   12
         Left            =   3780
         TabIndex        =   89
         Top             =   1110
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Base Retención"
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
         Height          =   255
         Index           =   11
         Left            =   180
         TabIndex        =   88
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
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
         Height          =   255
         Index           =   10
         Left            =   3780
         TabIndex        =   87
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   86
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "Totales Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   85
         Top             =   210
         Width           =   1980
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3330
      TabIndex        =   71
      Top             =   90
      Width           =   2565
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   72
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Datos Fiscales"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pagos"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Errores NºRegistro"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas sin Asiento"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   10170
      TabIndex        =   69
      Top             =   90
      Width           =   2415
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmFacturasPro.frx":0000
         Left            =   90
         List            =   "frmFacturasPro.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   2235
      End
   End
   Begin VB.Frame FrameAux2 
      Height          =   2145
      Left            =   270
      TabIndex        =   60
      Top             =   4920
      Width           =   9375
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   1890
         TabIndex        =   117
         Tag             =   "Fecha|F|N|||factpro_totales|fecharec|dd/mm/yyyy||"
         Text            =   "Fecha"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   10
         Left            =   8160
         TabIndex        =   94
         Tag             =   "Importe Rec|N|S|||factpro_totales|imporec|###,###,##0.00||"
         Text            =   "ImpRec"
         Top             =   1590
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   7260
         TabIndex        =   93
         Tag             =   "Importe Iva|N|S|||factpro_totales|impoiva|###,###,##0.00||"
         Text            =   "ImpIva"
         Top             =   1590
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   6390
         TabIndex        =   92
         Tag             =   "%Ret|N|S|||factpro_totales|porcrec|##0.00||"
         Text            =   "PorRec"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   5550
         TabIndex        =   67
         Tag             =   "%Iva|N|S|||factpro_totales|porciva|##0.00||"
         Text            =   "PorIva"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   6
         Left            =   4800
         TabIndex        =   66
         Tag             =   "Iva|N|S|||factpro_totales|codigiva|000||"
         Text            =   "Iva"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   5
         Left            =   4080
         TabIndex        =   65
         Tag             =   "Base Imponible|N|S|||factpro_totales|baseimpo|###,###,##0.00||"
         Text            =   "Base Imponible"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   3330
         TabIndex        =   64
         Tag             =   "Linea|N|N|||factpro_totales|numlinea|||"
         Text            =   "Linea"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   2580
         TabIndex        =   63
         Tag             =   "Año factura|N|N|||factpro_totales|anofactu||S|"
         Text            =   "Año"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   1110
         TabIndex        =   62
         Tag             =   "Nº registro|N|N|0||factpro_totales|numregis|0000000|S|"
         Text            =   "numregis"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   330
         TabIndex        =   61
         Tag             =   "Nº Serie|T|S|||factpro_totales|numserie||S|"
         Text            =   "Serie"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   2910
         Top             =   630
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
      Begin MSComctlLib.ListView lw1 
         Height          =   1545
         Left            =   150
         TabIndex        =   68
         Top             =   510
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   2725
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
         Caption         =   "Desglose Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   70
         Top             =   210
         Width           =   1980
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   13950
      TabIndex        =   53
      Top             =   270
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5970
      TabIndex        =   51
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   52
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   270
      TabIndex        =   48
      Top             =   90
      Width           =   3015
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   50
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4050
      Index           =   0
      Left            =   270
      TabIndex        =   37
      Top             =   870
      Width           =   17160
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
         Index           =   3
         ItemData        =   "frmFacturasPro.frx":0044
         Left            =   10590
         List            =   "frmFacturasPro.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1950
         Visible         =   0   'False
         Width           =   6270
      End
      Begin VB.TextBox Text1 
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
         Index           =   26
         Left            =   9450
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "Fecha|F|N|||factpro|fecfactu|dd/mm/yyyy|N|"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
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
         Index           =   25
         Left            =   7980
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Nº factura|T|N|||factpro|numfactu|||"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
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
         Left            =   15540
         TabIndex        =   16
         Tag             =   "Número Asiento|N|S|||factpro|numasien|00000000||"
         Text            =   "1234567890"
         Top             =   3270
         Width           =   1290
      End
      Begin VB.TextBox Text1 
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
         Left            =   14550
         TabIndex        =   6
         Tag             =   "Fecha Liquidacion|F|N|||factpro|fecliqpr|||"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   11340
         TabIndex        =   15
         Tag             =   "Porcentaje Retencion|N|S|||factpro|retfacpr|##0.00||"
         Text            =   "1234567890"
         Top             =   3270
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   4890
         TabIndex        =   14
         Tag             =   "Cuenta Retencion|T|S|||factpro|cuereten|||"
         Text            =   "1234567890"
         Top             =   3270
         Width           =   1350
      End
      Begin VB.TextBox Text4 
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
         Index           =   6
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text4"
         Top             =   3270
         Width           =   4785
      End
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
         Index           =   2
         ItemData        =   "frmFacturasPro.frx":0048
         Left            =   180
         List            =   "frmFacturasPro.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Tipo retencion|N|N|||factpro|tiporeten|||"
         Top             =   3270
         Width           =   4560
      End
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
         Index           =   1
         ItemData        =   "frmFacturasPro.frx":004C
         Left            =   7980
         List            =   "frmFacturasPro.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "Tipo operación|N|N|||factpro|codopera|||"
         Top             =   1950
         Width           =   2490
      End
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
         ItemData        =   "frmFacturasPro.frx":0050
         Left            =   7980
         List            =   "frmFacturasPro.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1260
         Width           =   8850
      End
      Begin VB.TextBox Text4 
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
         Index           =   5
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text4"
         Top             =   1950
         Width           =   6105
      End
      Begin VB.TextBox Text4 
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
         Index           =   4
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "Text4"
         Top             =   1260
         Width           =   6135
      End
      Begin VB.TextBox Text4 
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
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text4"
         Top             =   540
         Width           =   4245
      End
      Begin VB.TextBox Text1 
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
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "Observaciones|T|S|||factpro|observa|||"
         Top             =   2580
         Width           =   16635
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   210
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "Serie|T|N|||factpro|numserie||S|"
         Text            =   "123"
         Top             =   540
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FCFCE2&
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
         Left            =   5205
         TabIndex        =   2
         Tag             =   "Nº Registro|N|S|||factpro|numregis|0000000|S|"
         Top             =   540
         Width           =   1155
      End
      Begin VB.TextBox Text1 
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
         Left            =   6510
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Recepcion|F|N|||factpro|fecharec|dd/mm/yyyy||"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FCFCE2&
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
         Left            =   210
         TabIndex        =   7
         Tag             =   "Cuenta Proveedor|T|N|||factpro|codmacta|||"
         Text            =   "1234567890"
         Top             =   1260
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         TabIndex        =   9
         Tag             =   "Forma de pago|N|N|||factpro|codforpa|000||"
         Text            =   "1234567890"
         Top             =   1950
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   8010
         TabIndex        =   91
         Tag             =   "Año factura|N|N|||factpro|anofactu||S|"
         Text            =   "1234567890"
         Top             =   2580
         Width           =   1350
      End
      Begin VB.TextBox Text1 
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
         Left            =   10440
         MaxLength       =   30
         TabIndex        =   96
         Tag             =   "Tipo factura|T|N|||factpro|codconce340|||"
         Top             =   1260
         Width           =   1245
      End
      Begin VB.TextBox Text1 
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
         Left            =   15540
         TabIndex        =   114
         Tag             =   "Número Diario|N|S|||factpro|numdiari|00000000||"
         Text            =   "1234567890"
         Top             =   3270
         Width           =   1290
      End
      Begin VB.TextBox Text1 
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
         Index           =   27
         Left            =   8070
         MaxLength       =   30
         TabIndex        =   120
         Tag             =   "Tipo intracomunitaria|T|S|||factpro|codintra|||"
         Top             =   2580
         Width           =   1245
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo Intracomunitaria"
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
         Left            =   10590
         TabIndex        =   119
         Top             =   1650
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "F.Factura"
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
         Left            =   9480
         TabIndex        =   116
         Top             =   240
         Width           =   1020
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   9
         Left            =   10560
         Picture         =   "frmFacturasPro.frx":0054
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "NºFactura"
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
         Left            =   8010
         TabIndex        =   115
         Top             =   240
         Width           =   1020
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   8
         Left            =   1740
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   7
         Left            =   15630
         Picture         =   "frmFacturasPro.frx":00DF
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   16560
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Liq. "
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
         Height          =   255
         Index           =   2
         Left            =   14550
         TabIndex        =   97
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Asiento"
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
         Height          =   255
         Index           =   8
         Left            =   15600
         TabIndex        =   83
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "% Retención"
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
         Height          =   255
         Index           =   7
         Left            =   11340
         TabIndex        =   82
         Top             =   3000
         Width           =   1365
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   4890
         TabIndex        =   81
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   6780
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   180
         TabIndex        =   79
         Top             =   3000
         Width           =   1380
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Operación"
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
         Left            =   8010
         TabIndex        =   78
         Top             =   1650
         Width           =   1920
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
         Left            =   7980
         TabIndex        =   77
         Top             =   960
         Width           =   1380
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   1770
         Top             =   1650
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   210
         TabIndex        =   76
         Top             =   1650
         Width           =   1545
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   2100
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Proveedor"
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
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   73
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   7620
         Picture         =   "frmFacturasPro.frx":016A
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   810
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Observaciones"
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
         Left            =   180
         TabIndex        =   42
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label Label18 
         Caption         =   "Recepción"
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
         Left            =   6540
         TabIndex        =   41
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label4 
         Caption         =   "Nº Registro"
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
         Left            =   5220
         TabIndex        =   39
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   38
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   285
      TabIndex        =   43
      Top             =   7125
      Width           =   17190
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   2
         Left            =   2190
         TabIndex        =   118
         Tag             =   "Fecha|F|N|||factpro_lineas|fecharec|dd/mm/yyyy||"
         Text            =   "fecha"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   12
         Left            =   15630
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   95
         Text            =   "Nombre cuenta"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.CheckBox chkAux 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   14250
         TabIndex        =   35
         Tag             =   "Aplica Retencion|N|N|0|1|factpro_lineas|aplicret|||"
         Top             =   2190
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   11
         Left            =   13200
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "Importe Rec|N|S|||factpro_lineas|imporec|###,###,##0.00||"
         Text            =   "ImpRec"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   3
         Left            =   2910
         TabIndex        =   26
         Tag             =   "Año factura|N|N|||factpro_lineas|anofactu||S|"
         Text            =   "año"
         Top             =   2160
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   12
         Left            =   14520
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "CC|T|S|||factpro_lineas|codccost|||"
         Text            =   "CC"
         Top             =   2160
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   10
         Left            =   12090
         MaxLength       =   15
         TabIndex        =   31
         Tag             =   "Importe Iva|N|S|||factpro_lineas|impoiva|###,###,##0.00||"
         Text            =   "ImpIva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   9
         Left            =   10980
         MaxLength       =   15
         TabIndex        =   34
         Tag             =   "% Recargo|N|S|||factpro_lineas|porcrec|##0.00||"
         Text            =   "%rec"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   8
         Left            =   10140
         MaxLength       =   50
         TabIndex        =   33
         Tag             =   "% Iva|N|S|||factpro_lineas|porciva|##0.00||"
         Text            =   "%iva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   15420
         TabIndex        =   59
         ToolTipText     =   "Buscar concepto"
         Top             =   2130
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   6
         Left            =   8370
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Importe Base|N|N|||factpro_lineas|baseimpo|###,###,##0.00||"
         Text            =   "Importe"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   9780
         TabIndex        =   58
         ToolTipText     =   "Buscar cuenta"
         Top             =   2190
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   7
         Left            =   9150
         MaxLength       =   15
         TabIndex        =   30
         Tag             =   "Codigo Iva|N|N|||factpro_lineas|codigiva|000||"
         Text            =   "Iva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   60
         TabIndex        =   56
         Top             =   0
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   180
            TabIndex        =   57
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   5
         Left            =   4050
         MaxLength       =   15
         TabIndex        =   28
         Tag             =   "Cuenta|T|N|||factpro_lineas|codmacta|||"
         Text            =   "Cta Base"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   4
         Left            =   3330
         TabIndex        =   27
         Tag             =   "Linea|N|N|||factpro_lineas|numlinea||S|"
         Text            =   "linea"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   0
         Left            =   105
         TabIndex        =   24
         Tag             =   "Nº Serie|T|S|||factpro_lineas|numserie||S|"
         Text            =   "Serie"
         Top             =   2145
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   1
         Left            =   840
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "Nº registro|N|N|0||factpro_lineas|numregis|0000000|S|"
         Text            =   "numregis"
         Top             =   2145
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4800
         TabIndex        =   45
         ToolTipText     =   "Buscar cuenta"
         Top             =   2190
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   5
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   44
         Text            =   "Nombre cuenta"
         Top             =   2190
         Visible         =   0   'False
         Width           =   3285
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   3720
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "AdoAux(1)"
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Height          =   2040
         Index           =   1
         Left            =   45
         TabIndex        =   46
         Top             =   780
         Width           =   16770
         _ExtentX        =   29580
         _ExtentY        =   3598
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   285
      TabIndex        =   22
      Top             =   10290
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   120
         TabIndex        =   23
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   16380
      TabIndex        =   49
      Top             =   10350
      Width           =   1035
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
      Left            =   15090
      TabIndex        =   47
      Top             =   10350
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3720
      Top             =   10320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   16320
      TabIndex        =   54
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
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8190
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   16380
      TabIndex        =   40
      Top             =   10350
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacturasPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public FACTURA As String  'Con pipes numserie|numfactu|anofactu
Public vLinapu As Integer

Public DesdeNorma43 As Byte  'La uno y la 2 son validas
Public Datos As String  'Tendra, empipado, numero asiento  y demas

Private Const NO = "No encontrado"

Private Const IdPrograma = 404

Private WithEvents frmFact As frmFacturasProPrev
Attribute frmFact.VB_VarHelpID = -1
Private WithEvents frmFPag As frmBasico2
Attribute frmFPag.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmPais As frmBasico2
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmAgen As frmAgentes
Attribute frmAgen.VB_VarHelpID = -1
Private WithEvents frmDpto As frmBasico
Attribute frmDpto.VB_VarHelpID = -1

Private frmAsi As frmAsientosHco
Attribute frmAsi.VB_VarHelpID = -1

Private WithEvents frmTIva As frmBasico2 'frmIVA
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmCtasRet As frmColCtas
Attribute frmCtasRet.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico 'frmCCoste
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmPag As frmFacturasProPag ' pagos de tesoreria
Attribute frmPag.VB_VarHelpID = -1
Private WithEvents frmUtil As frmUtilidades
Attribute frmUtil.VB_VarHelpID = -1
Private frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1


Dim AntiguoText1 As String
Private CadenaAmpliacion As String
Private Sql As String


Dim PosicionGrid As Integer

Dim Linliapu As Long
Dim FicheroAEliminar As String

Dim Numasien2 As Long
Dim NumDiario As Integer


Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos

Dim CadB As String
Dim CadB1 As String
Dim CadB2 As String

Dim PulsadoSalir As Boolean
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim ActualizandoAsiento As Boolean   'Para k no devuelv el contador
Dim VieneDeDesactualizar As Boolean

Dim B As Boolean

Private BuscaChekc As String

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim VarieAnt As String
Private DiarioPorDefecto As String 'Si solo tiene un diario que lo ponga

Dim IT As ListItem


Dim cadFiltro As String
Dim i As Long
Dim Ancho As Integer

Private Mc As Contadores

Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

'Por si esta en un periodo liquidado, que pueda modificar CONCEPTO , cuentas,
Private ModificaFacturaPeriodoLiquidado As Boolean

Dim ModificarPagos As Boolean


Dim IvaCuenta As String
Dim CambiarIva As Boolean

Dim CtaBanco As String
Dim IBAN As String
Dim NomBanco As String

Dim Pagado As Byte
Dim FechaPago As String

Dim TipForpa As Integer
Dim FecFactuAnt As String
Dim NumFactuAnt As String
Dim CodmactaAnt As String


Dim FecRecepAnt As String ' recepcion




Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    If Modo = 0 Then Exit Sub
    'HacerBusqueda2
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAux_LostFocus(Index As Integer)
    If Not vParam.autocoste Then
        PonleFoco cmdAceptar
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim cad As String
    Dim i As Integer
    Dim Limp As Boolean
    Dim Mc As Contadores
    Dim B As Boolean
    Dim SqlLog As String
    

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                FecFactuAnt = Text1(26).Text
                FecRecepAnt = Text1(1).Text
                
                Set Mc = New Contadores
                i = FechaCorrecta2(CDate(Text1(1).Text))
                If Mc.ConseguirContador(Trim(Text1(2).Text), (i = 0), False) = 0 Then
                    'COMPROBAR NUMERO ASIENTO
                    Text1(0).Text = Mc.Contador
                    If ComprobarNumeroFactura(i = 0) Then
                        B = InsertarDesdeForm2(Me, 1)
                    Else
                        B = False
                    End If
                    
                    If B Then
                        Data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PosicionarData
                        PonerCampos
                        '[Monica]14/05/2015 añado numasien
                        Numasien2 = 0
                        BotonAnyadirLinea 1, True
                    Else
                        'SI NO INSERTA debemos devolver el contador
                        Mc.DevolverContador Trim(Text1(2).Text), (i = 0), Mc.Contador
                    End If
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                '-----------------------------------------
                'Hay que comprobar si ha modificado, o no la clave de la factura
                i = 1
                If Data1.Recordset!NUmSerie = Text1(2).Text Then
                    If Data1.Recordset!NumRegis = CLng(Text1(0).Text) Then
                        If Data1.Recordset!anofactu = Text1(14).Text Then
                            i = 0
                            'NO HA MODIFICADO NADA
                        End If
                    End If
                End If
            
                'Hacemos MODIFICAR
                Dim RC As Boolean
                If i <> 0 Then
                    MsgBox "No se puede cambiar campos clave  de la factura.", vbExclamation
                    RC = False
                Else
                    RC = ModificarFactura
                End If
                    
                If RC Then
                    '--DesBloqueaRegistroForm Me.Text1(0)
                    TerminaBloquear
                    
                    If Numasien2 > 0 Then
                        If IntegrarFactura Then
                            Text1(8).Text = Format(Numasien2, "0000000")
                            Numasien2 = -1
                            NumDiario = 0
                        Else
                            B = False
                        End If
                    End If
                    
                    If ModificarPagos Then PagosTesoreria
                    
                    'LOG
                    SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
                    SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text

                    vLog.Insertar 9, vUsu, SqlLog
                    
                    
                    'Nuevo. Si ahora tiene retencion, y antes NO tenia
                    ActualizarRetencionLineasSiNecesario
                    
                    
                    
                    
                    PosicionarData
                    
                End If
            End If
        
        Case 5 'LLÍNIES
            FecFactuAnt = Text1(26).Text
            FecRecepAnt = Text1(1).Text
            
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    ModificarLinea
                                        
                    '**** parte de contabilizacion de la factura
                    TerminaBloquear
                    
                    If Numasien2 > 0 Then
                        If IntegrarFactura Then
                            Text1(8).Text = Format(Numasien2, "0000000")
                            Numasien2 = -1
                            NumDiario = 0
                        Else
                            B = False
                        End If
                    End If
                
                    If ModificarPagos Then PagosTesoreria
                    
                    PosicionarData
                    
            End Select
            
    
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PagosTesoreria()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Actualizar As Boolean
Dim Aux As String


    On Error GoTo ePagosTesoreria

    If Not vEmpresa.TieneTesoreria Then Exit Sub
    
    
    
    
    ' si me cambian el nro de fra la cambio ya, SIEMPRE que no haayan pagos parciales
    
    Actualizar = False
    If Trim(Text1(25).Text) <> Trim(NumFactuAnt) Then Actualizar = True
    If Text1(4).Text <> CodmactaAnt Then Actualizar = True
    If Actualizar Then
    
        Sql = "numserie = " & DBSet(Text1(2).Text, "T")
        Sql = Sql & " and codmacta = " & DBSet(CodmactaAnt, "T") & " and numfactu = " & DBSet(NumFactuAnt, "T")
        Sql = Sql & " and fecfactu = " & DBSet(FecFactuAnt, "F") & " AND 1"
        Sql = DevuelveDesdeBD("imppagad", "pagos", Sql, "1")
        If Sql <> "" Then
            'Tiene pagos parciales efectuados. Debera ir a tesoreria
            MsgBox "Tiene pagos parciales realizados. Revise tesorería", vbExclamation
            Exit Sub
        End If
        Sql = "update pagos set numfactu = " & DBSet(Text1(25).Text, "T")
        Sql = Sql & ", codmacta = " & DBSet(Text1(4).Text, "T")
        
        'Datos fiscales
        If Text1(4).Text <> CodmactaAnt Then
            Set Rs = New ADODB.Recordset
            
            Aux = "Select Nommacta, dirdatos ,codposta ,desPobla, desProvi, nifdatos, codPAIS FROM cuentas where codmacta =" & DBSet(Text1(4).Text, "T")
            Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            'NO PUEDE SER EOG
            If Rs.EOF Then
                MsgBox "Cuenta proveedor incorrecta. ", vbCritical
            Else
                'nomprove  domprove   pobprove  cpprove nifprove codpais
                
                Sql = Sql & ", nomprove =" & DBSet(Rs!Nommacta, "T")
                Sql = Sql & ",domprove =" & DBSet(Rs!dirdatos, "T")
                Sql = Sql & ", cpprove=" & DBSet(Rs!codposta, "T")
                Sql = Sql & ", pobprove=" & DBSet(Rs!desPobla, "T")
                Sql = Sql & ", proprove=" & DBSet(Rs!desProvi, "T")
                Sql = Sql & ", nifprove=" & DBSet(Rs!nifdatos, "T")
                Sql = Sql & ", codPAIS=" & DBSet(Rs!codPAIS, "T")
                    
            End If
            Rs.Close
            Set Rs = Nothing
                        
        End If
        Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T")
        Sql = Sql & " and codmacta = " & DBSet(CodmactaAnt, "T") & " and numfactu = " & DBSet(NumFactuAnt, "T")
        Sql = Sql & " and fecfactu = " & DBSet(FecFactuAnt, "F")
        
        Conn.Execute Sql
        
        Sql = "update hlinapu set numfacpr = " & DBSet(Text1(25).Text, "T") & " where numserie = " & DBSet(Text1(2).Text, "T")
        Sql = Sql & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfacpr = " & DBSet(NumFactuAnt, "T")
        Sql = Sql & " and fecfactu = " & DBSet(FecFactuAnt, "F")
        
        Conn.Execute Sql
    End If
    
    '[Monica]12/09/2016: si la factura ha sido traspasada y no está en cartera, no hacemos nada en cartera
    If EsFraProTraspasada And Not ExisteAlgunPago(Text1(2).Text, CodmactaAnt, Text1(25).Text, FecFactuAnt, False) Then Exit Sub
    
    
    If ExisteAlgunPago(Text1(2).Text, CodmactaAnt, Text1(25).Text, FecFactuAnt, True) Then
        MsgBox "Hay algún efecto que ya ha sido pagado. Revise cartera de pagos.", vbExclamation

        Set frmMens = New frmMensajes

        frmMens.Opcion = 28
        frmMens.Parametros = Trim(Text1(2).Text) & "|" & CodmactaAnt & "|" & Trim(Text1(25).Text) & "|" & Text1(26).Text & "|"
        frmMens.Show vbModal

        Set frmMens = Nothing

        ContinuarPago = False

        Exit Sub
    
    End If
    

    Sql = "delete from tmppagos where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute Sql
    
    ContinuarPago = False
    
    If CargarPagosTemporal(Text1(5).Text, Text1(1).Text, ImporteFormateado(Text1(13).Text)) Then
        ' Insertamos
        If Not ExisteAlgunPago(Text1(2).Text, Text1(4).Text, Text1(25).Text, FecFactuAnt, False) Then
    '{{{{{{{{{
            Sql = "select ccc.ctabanco,ccc.iban, ddd.nommacta "
            Sql = Sql & " from cuentas ccc, cuentas ddd "
            Sql = Sql & " where ccc.codmacta = " & DBSet(Text1(4).Text, "T")
            Sql = Sql & " and ccc.ctabanco = ddd.codmacta "
            
            CtaBanco = ""
            IBAN = ""
            NomBanco = ""
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs.EOF Then
                CtaBanco = DBLet(Rs.Fields(0))
                IBAN = DBLet(Rs.Fields(1))
                NomBanco = DBLet(Rs.Fields(2))
            End If
        
            TipForpa = DevuelveValor("select formapago.tipforpa from formapago where codforpa = " & DBSet(Text1(5).Text, "N"))
            
            Set frmPag = frmFacturasProPag
            frmPag.CodigoActual = CtaBanco & "|" & "|" & "|" & "|" & "|" & IBAN & "|" & TipForpa & "|" & NomBanco & "|"
            frmPag.Show vbModal
            Set frmPag = Nothing
    
            If ContinuarPago Then
                CargarPagos
                If Pagado Then ContabilizarPagos
            End If
            
        Else
            Dim Nregs As Long
            Dim Nregs2 As Long
            Nregs = TotalRegistros("select count(*) from tmppagos where codusu = " & vUsu.Codigo)
            Nregs2 = TotalRegistros("select count(*) from pagos where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T") & " and fecfactu = " & DBSet(FecFactuAnt, "F") & " and codmacta = " & DBSet(Text1(4).Text, "T"))

            If Nregs = Nregs2 Then
                CargarPagos
            Else
                MsgBox "No coincide el número de pagos en tesoreria. Modifíquelos en cartera.", vbExclamation
                ' mandarlo al listview de cobros
            
                Set frmMens = New frmMensajes
                
                frmMens.Opcion = 28
                frmMens.Parametros = Trim(Text1(2).Text) & "|" & Trim(Text1(4).Text) & "|" & Trim(Text1(25).Text) & "|" & Text1(26).Text & "|"
                frmMens.Show vbModal
                
                Set frmMens = Nothing
            
            End If
        
        End If
    End If
    
    
    Exit Sub
    
ePagosTesoreria:
    MuestraError Err.Number, "Pagos Tesoreria", Err.Description
End Sub

Private Function ExisteAlgunPago(Serie As String, Cuenta As String, FACTURA As String, FecFactu As String, Pagado As Boolean) As Boolean
Dim Sql As String
    
    Sql = "select count(*) from pagos where numserie = " & DBSet(Serie, "T")
    Sql = Sql & " and codmacta = " & DBSet(Cuenta, "T")
    Sql = Sql & " and numfactu = " & DBSet(FACTURA, "T")
    Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
    
    If Pagado Then
' un pago lo damos como pagado si el importe de pago es <> 0
'[Monica]12/09/2016: quito la condicion: numasien is null pq puede tener nro de transferencia y no modificariamos el importe total de transferencia
        Sql = Sql & " and imppagad <> 0 and not imppagad is null " 'and numasien is null "
    End If
    
    ExisteAlgunPago = (TotalRegistros(Sql) <> 0)

End Function


Private Function PagosContabilizados(Serie As String, Cuenta As String, FACTURA As String, FecFactu As String) As String
Dim Sql As String
Dim CadResult As String
Dim Rs As ADODB.Recordset

    On Error GoTo ePagosContabilizados

    Sql = "select numasien, fechaent from hlinapu where numserie = " & DBSet(Serie, "T")
    Sql = Sql & " and codmacta = " & DBSet(Cuenta, "T")
    Sql = Sql & " and numfacpr = " & DBSet(FACTURA, "T")
    Sql = Sql & " and fecfactu = " & DBSet(FecFactu, "F")
    
    CadResult = ""
    
    If TotalRegistrosConsulta(Sql) = 0 Then
        CadResult = ""
    Else
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            CadResult = CadResult & Format(DBLet(Rs!NumAsien, "N"), "0000000") & " de " & Format(DBLet(Rs!FechaEnt), "dd/mm/yyyy") & vbCrLf
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    End If
    
    
    PagosContabilizados = CadResult
    
    Exit Function
    
ePagosContabilizados:
    MuestraError Err.Number, "Pagos contabilizados", Err.Description
End Function



Private Sub CargarPagos()
Dim Sql As String
Dim Mens As String

    If ExisteAlgunPago(Text1(2).Text, Text1(4).Text, Text1(25).Text, FecFactuAnt, False) Then
        B = ActualizarPagos(Mens)
        
        If B Then
            Sql = PagosContabilizados(Text1(2).Text, Text1(4).Text, Text1(25).Text, FecFactuAnt)
            If Sql <> "" Then
                MsgBox "La factura tiene asientos que ya están contabilizados. Revise y modifique en su caso los siguientes asientos: " & vbCrLf & vbCrLf & Sql, vbExclamation
            End If
        End If
    Else
        B = InsertarPagos(Mens)
    End If
    
    If B Then
'        MsgBox "Proceso realizado correctamente.", vbExclamation
    Else
        MuestraError 0, "Cargar Pagos", Mens
    End If

End Sub

Private Function UpdateaPagos(ByRef Rs As ADODB.Recordset, ByRef RS1 As ADODB.Recordset, ByRef i As Long, ByRef Mens As String) As Boolean
Dim Sql As String

    On Error GoTo eUpdateaPagos
    
    UpdateaPagos = False

    B = True

    While Not Rs.EOF And B
        Sql = "update pagos set codmacta = " & DBSet(Text1(4).Text, "T")
        
        Sql = Sql & ", codforpa = " & DBSet(Text1(5).Text, "N")
        Sql = Sql & ", fecefect = " & DBSet(RS1!FecVenci, "F")
        Sql = Sql & ", impefect = " & DBSet(RS1!ImpVenci, "N")
        If Modo <> 4 Then Sql = Sql & ", ctabanc1 = " & DBSet(CtaBanco, "T", "S")
        Sql = Sql & ", fecfactu = " & DBSet(Text1(26).Text, "F")
        
        If Pagado Then
            Sql = Sql & ", fecultpa = " & DBSet(FechaPago, "F") ' DBSet(Rs!FecVenci, "F")
            Sql = Sql & ", imppagad = " & DBSet(RS1!ImpVenci, "N")
        Else
            Sql = Sql & ", fecultpa = " & ValorNulo
            Sql = Sql & ", imppagad = " & ValorNulo
        End If
        Sql = Sql & ", iban = " & DBSet(IBAN, "T", "S")
        Sql = Sql & ", numorden = " & DBSet(RS1!numorden, "N")
        Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
        Sql = Sql & " and fecfactu = " & DBSet(FecFactuAnt, "F") & " and numorden = " & DBSet(Rs!numorden, "N")
        
        Conn.Execute Sql
        
        i = Rs!numorden ' me guardo el nro de orden para despues ir incrementandolo
        
        RS1.MoveNext
        Rs.MoveNext
        
        ' si no hay mas registros en la temporal salgo del bucle
        If RS1.EOF Then B = False
    Wend
    
    UpdateaPagos = True
    Exit Function

eUpdateaPagos:
    Mens = Mens & Err.Description
End Function

Private Function InsertaPagos(ByRef RS1 As ADODB.Recordset, ByRef i As Long, ByRef Mens As String) As Boolean
Dim CadInsert As String
Dim CadValues As String
Dim textCSB As String
Dim Sql As String
        
    On Error GoTo eInsertaPagos
        
    InsertaPagos = False
        
    CadInsert = "insert into pagos (numserie,codmacta,numfactu,fecfactu,numorden,codforpa,fecefect,impefect," & _
                "ctabanc1,fecultpa,imppagad,emitdocum," & _
                "text1csb,text2csb,nrodocum,referencia, iban,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais,situacion,codusu) values "
    CadValues = ""
    
    While Not RS1.EOF
        i = i + 1
        
        Sql = DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(4).Text, "T") & "," & DBSet(Text1(25).Text, "T") & "," & DBSet(Text1(26).Text, "F") & "," & DBSet(i, "N") & ","
        Sql = Sql & DBSet(Text1(5).Text, "N") & "," & DBSet(RS1!FecVenci, "F") & "," & DBSet(RS1!ImpVenci, "N") & ","
        Sql = Sql & DBSet(CtaBanco, "T", "S") & ","
        
        If Pagado Then
'            B = ContabilizarPago
            Sql = Sql & DBSet(FechaPago, "F") & "," & DBSet(RS1!ImpVenci, "N") & ","
        Else
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
        End If
        
        Sql = Sql & "0,"
        
        textCSB = "Factura " & Text1(25).Text & " de Fecha " & Text1(1).Text
        
        Sql = Sql & DBSet(textCSB, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(IBAN, "T", "S") & ","
        
        Sql = Sql & DBSet(Text1(15).Text, "T", "S") & "," & DBSet(Text1(16).Text, "T", "S") & "," & DBSet(Text1(17).Text, "T", "S") & "," & DBSet(Text1(18).Text, "T", "S") & ","
        Sql = Sql & DBSet(Text1(19).Text, "T", "S") & "," & DBSet(Text1(20).Text, "T", "S") & "," & DBSet(Text1(21).Text, "T", "S") & ","
        
        If Pagado Then
            Sql = Sql & "1"
        Else
            Sql = Sql & "0"
        End If
        
        ' la parte del codusu
        Sql = Sql & "," & DBSet(vUsu.Id, "N")
        
        
        CadValues = CadValues & "(" & Sql & "),"
    
        RS1.MoveNext
    Wend

    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute CadInsert & CadValues
    End If

    InsertaPagos = True
    Exit Function

eInsertaPagos:
    Mens = Mens & Err.Description
End Function

Private Function ActualizarPagos(ByRef Mens As String) As Boolean
Dim Sql As String
Dim Sql1 As String
Dim Nregs As Integer
Dim Nregs1 As Integer
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim CadInsert As String
Dim CadValues As String

    On Error GoTo eActualizarPagos

    ActualizarPagos = False


    Sql = "select * from pagos where numserie = " & DBSet(Text1(2).Text, "T") & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T") & " and fecfactu = " & DBSet(FecFactuAnt, "F")
    
    Sql = Sql & " order by numorden "
    Nregs = TotalRegistrosConsulta(Sql)
    
    Sql1 = "select * from tmppagos where codusu = " & vUsu.Codigo & " order by numorden "
    Nregs1 = TotalRegistrosConsulta(Sql1)
    
    If Nregs = Nregs1 Then
    ' Mismo nro de registros en pagos que en la temporal --> los actualizamos
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        i = 0
        
        Mens = "Actualizando Pagos: " & vbCrLf & vbCrLf
        B = UpdateaPagos(Rs, RS1, i, Mens)
        
        Set Rs = Nothing
        Set RS1 = Nothing
    
    ElseIf Nregs < Nregs1 Then
    ' Menos registros en pagos que en la temporal --> actualizamos e insertamos los no existentes
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        i = 0
        
        Mens = "Actualizando Pagos: " & vbCrLf & vbCrLf
        B = UpdateaPagos(Rs, RS1, i, Mens)
        
        Set Rs = Nothing ' cierro el de pagos
        
        ' sin cerrar el recordset de tmppagos, insertamos los restantes registros de la tmppagos
        Mens = "Insertando Pagos Restantes: " & vbCrLf & vbCrLf
        B = InsertaPagos(RS1, i, Mens)
        
        Set RS1 = Nothing
    
    Else
    ' Mas registros en pagos que en la temporal --> actualizamos y borramos los que sobran
    
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        i = 0
        
        Mens = "Actualizando Pagos: " & vbCrLf & vbCrLf
        B = UpdateaPagos(Rs, RS1, i, Mens)
        
        Set Rs = Nothing ' cierro el de pagos
        
        'borro los registros restantes de pagos
        Mens = "Eliminado Pagos restantes: " & vbCrLf & vbCrLf
        Sql = "delete from pagos "
        Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
        Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F") & " and numorden > " & DBSet(i, "N")
        
        Conn.Execute Sql
        
        Set RS1 = Nothing
    End If

    ActualizarPagos = B
    Exit Function

eActualizarPagos:
    Mens = Mens & Err.Description
End Function


Private Function InsertarPagos(ByRef Mens As String) As Boolean
Dim Sql As String
Dim textCSB As String
Dim CadInsert As String
Dim CadValues As String
Dim Rs As ADODB.Recordset
Dim i As Long

    On Error GoTo eInsertarPagos

    InsertarPagos = False

    Sql = "select * from tmppagos where codusu = " & DBSet(vUsu.Codigo, "N") & " order by numorden "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    i = 0
    Mens = "Insertando Pagos: " & vbCrLf & vbCrLf
    B = InsertaPagos(Rs, i, Mens)
    
    Set Rs = Nothing
    
    InsertarPagos = B
    Exit Function
    
eInsertarPagos:
    MuestraError Err.Number, "Insertar Pagos", Err.Description & " " & Mens
End Function

Private Function CargarPagosTemporal(Forpa As String, FecFactu As String, TotalFac As Currency) As Boolean
Dim Sql As String
Dim CadValues As String
Dim Rsvenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarPagos

    CargarPagosTemporal = False

    Sql = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set Rsvenci = New ADODB.Recordset
    Rsvenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    If Not Rsvenci.EOF Then
        If Rsvenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            i = 1
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(Rsvenci!primerve, "N"), FecVenci)
            '===
            
            'IMPORTE del Vencimiento
            If Rsvenci!numerove = 1 Then
                ImpVenci = TotalFac
            Else
                ImpVenci = Round(TotalFac / Rsvenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * Rsvenci!numerove <> TotalFac Then
                    ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * Rsvenci.Fields(0).Value), 2)
                End If
            End If
            CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For i = 2 To Rsvenci!numerove
                FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round(TotalFac / Rsvenci.Fields(0).Value, 2)
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next i
        End If
    End If
    
    Set Rsvenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO tmppagos (codusu, numorden, fecvenci, impvenci)"
        Sql = Sql & " VALUES " & Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute Sql
    End If
    
    CargarPagosTemporal = True
    Exit Function

eCargarPagos:

End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "numserie= " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 ' cuenta base
            cmdAux(0).Tag = 0
            LlamaContraPar
            If txtaux(5).Text <> "" Then
                txtAux_LostFocus 5
                If txtaux(5).Text <> "" Then PonFoco txtaux(6)
            Else
                PonFoco txtaux(5)
            End If
        Case 1 'tipo de iva
            cmdAux(0).Tag = 1
            
            Set frmTIva = New frmBasico2
            AyudaTiposIva frmTIva
            Set frmTIva = Nothing
            
            PonFoco txtaux(7)
        Case 2 'cento de coste
            If txtaux(12).Enabled Then
                Set frmCC = New frmBasico
                AyudaCC frmCC
                Set frmCC = Nothing
            End If

    End Select
End Sub

Private Sub LlamaContraPar()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing
    
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
    If Index = 2 And (Modo = 3 Or Modo = 4) Then
        If Combo1(Index).ListIndex = 0 Then
            Text1(7).Text = ""
            Text1(6).Text = ""
            Text4(6).Text = ""
        End If
    End If
    
    If Combo1(Index).ListIndex = -1 Then Exit Sub
    
    ' en el caso de que sea bienes de inversion se pone en ambos combos
    If Index = 0 And Modo <> 1 Then
        If Chr(Combo1(Index).ItemData(Combo1(Index).ListIndex)) = "I" Then
            Combo1(1).ListIndex = 4
        Else
            If Combo1(1).ListIndex = 4 Then Combo1(1).ListIndex = 0
        End If
    End If
    
    If Index = 1 And (Modo = 3 Or Modo = 4) Then
        If Combo1(1).ListIndex = 4 Then
            PosicionarCombo Combo1(0), Asc("I")
            Text1(22).Text = "I"
        Else
            Combo1(0).ListIndex = 0
            Text1(22).Text = "0"
        End If
    End If
    
    If Index = 0 And (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        If Combo1(0).ListIndex = 0 Then
            Text1(22).Text = "0"
        Else
            Text1(22).Text = Chr(Combo1(0).ItemData(Combo1(0).ListIndex))
        End If
    End If
    
    
    ' intracomunitario
    If Index = 1 And (Modo = 1 Or Modo = 2 Or Modo = 3 Or Modo = 4) Then
        If Combo1(1).ListIndex = 1 Then
            Combo1(3).Visible = True
            Label11.Visible = True
            Combo1(3).Enabled = True
            Label11.Enabled = True
            
            If Modo = 3 Then
                PosicionarCombo Combo1(3), Asc("A")
                Text1(27).Text = "A"
            End If
            
        Else
            Combo1(3).Visible = False
            Label11.Visible = False
            Combo1(3).Enabled = False
            Label11.Enabled = False
            
            Text1(27).Text = ""
        End If
    End If
    ' tipo de intracomunitario
    If Index = 3 And (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        If Combo1(3).ListIndex = -1 Then
            Text1(27).Text = ""
        Else
            Text1(27).Text = Chr(Combo1(3).ItemData(Combo1(3).ListIndex))
        End If
    End If
    
    
End Sub

Private Sub Form_Activate()
'    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    
    If PrimeraVez Then
        B = False
        If FACTURA <> "" Then
            B = True
            Modo = 2
            Sql = "Select * from factpro "
            Sql = Sql & " WHERE numserie = " & RecuperaValor(FACTURA, 1)
            Sql = Sql & " AND numregis =" & RecuperaValor(FACTURA, 2)
            Sql = Sql & " AND anofactu= " & RecuperaValor(FACTURA, 3)
            CadenaConsulta = Sql
            PonerCadenaBusqueda
            'BOTON lineas
            
            cboFiltro.ListIndex = 0
            
        Else
            Modo = 0
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE numserie is null"
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            
            cboFiltro.ListIndex = vUsu.FiltroFactPro
            
        End If
        
        CargarSqlFiltro
        
        PonerModo CInt(Modo)
        VieneDeDesactualizar = B
'        CargaGrid 1, (Modo = 2)
        If Modo <> 2 Then
            
            'ESTO LO HE CAMBIADO HOY 9 FEB 2006
            'Antes no estaba el IF
            If FACTURA <> "" Then
                MsgBox "Proceso de sistema. Stop. Frm_Activate"
            End If
        Else

        End If
        If FACTURA <> "" Then
            If vLinapu > 0 Then
                If Not (AdoAux(1).Recordset Is Nothing) Then
                    If Not AdoAux(1).Recordset.EOF Then
                        AdoAux(1).Recordset.Find "linliapu = " & vLinapu
                        If AdoAux(1).Recordset.EOF Then AdoAux(1).Recordset.MoveFirst
                    End If
                End If
            End If
            
            'Pulso botono pasar a lineas
            HacerToolBar 10
            
            If DesdeNorma43 > 0 Then
                ModoLineas = 0
                'Ponemos en marcha, la maquinaria. Desde variable DATOS extraemos
                If DesdeNorma43 = 1 Then
                    BotonAnyadirLinea 1, True
                Else
                    'Es TIPO 2. Es decir lo dejamos modificando lineas
'                    BotonModificarLinea 1
                End If
            End If
        
        End If
        Toolbar1.Enabled = True
        
        PrimeraVez = False
        
        
    End If
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    cadFiltro = ""
    
    Select Case Me.cboFiltro.ListIndex
        Case 0 ' sin filtro
            cadFiltro = "(1=1)"
        
        Case 1 ' ejercicios abiertos
            cadFiltro = "factpro.fecharec >= " & DBSet(vParam.fechaini, "F")
        
        Case 2 ' ejercicio actual
            cadFiltro = "factpro.fecharec between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
        
        Case 3 ' ejercicio siguiente
            cadFiltro = "factpro.fecharec > " & DBSet(vParam.fechafin, "F")
    
    End Select
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)

     If Modo > 2 Then Cancel = 1
        

    Screen.MousePointer = vbDefault
    
    vUsu.ActualizarFiltro "ariconta", IdPrograma, Me.cboFiltro.ListIndex
    
End Sub

Private Sub Form_Load()
Dim i As Integer

    Me.Icon = frmppal.Icon

    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False
    ActualizandoAsiento = False
    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 47
        .Buttons(2).Image = 44
        .Buttons(3).Image = 42
        .Buttons(4).Image = 36
    End With


    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
   
    With Me.ToolbarAux
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    For i = 0 To imgppal.Count - 1
        If i <> 0 And i <> 7 And i <> 9 Then imgppal(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    CargaFiltros
    
    
    Caption = "Facturas de Proveedor"
    
    NumTabMto = 1
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han llínies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "factpro"
    Ordenacion = " ORDER BY factpro.numserie, factpro.numregis , factpro.fecfactu"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numserie is null"
    Data1.Refresh
       
    
    ModoLineas = 0
    DiarioPorDefecto = ""
       
    CargarColumnas
    
    CargarCombo
    
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    'Maxima longitud cuentas
    txtaux(5).MaxLength = vEmpresa.DigitosUltimoNivel
    'CadAncho = False
    PulsadoSalir = False

End Sub

Private Sub CargarColumnas()
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Columnas = "Linea|Tipo|Descripcion|Base|IVA|Recargo|"
    Ancho = "0|800|2450|1800|1800|1800|"
    'vwColumnRight =1  left=0   center=2
    Alinea = "0|0|0|1|1|1|"
    'Formatos
    Formato = "|||###,###,##0.00|###,###,##0.00|###,###,##0.00|"
    Ncol = 6

    lw1.Tag = "5|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim


End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    For i = 0 To Combo1.Count - 1
        Me.Combo1(i).ListIndex = -1
    Next i

    Me.chkAux(0).Value = 0

    lw1.ListItems.Clear
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    If Not Data1.Recordset Is Nothing Then
        DespalzamientoVisible B And (Data1.Recordset.RecordCount > 1)
    End If
    
    Toolbar1.Buttons(8).Enabled = B
    
    B = Modo = 2 Or Modo = 0 Or Modo = 5
    
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = B
        If Modo <> 1 Then
            Text1(i).BackColor = vbWhite
        End If
    Next i
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).Locked = B
    Next i
    
    For i = 0 To imgppal.Count - 1
        imgppal(i).Enabled = Not B
    Next i
    imgppal(6).Enabled = (Text1(8).Text <> "")
    
    ' observaciones
    
    imgppal(8).Enabled = (Data1.Recordset.RecordCount <> 0)
    
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
       
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    
    Text1(0).Enabled = (Modo = 1)
    
    Text1(2).Enabled = (Modo = 3 Or Modo = 1)
    imgppal(1).Enabled = (Modo = 3 Or Modo = 1)
    
    B = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
            
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 1, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    
    DataGridAux(1).Enabled = B
        
    'lineas de factura
    Dim anc As Single
    anc = DataGridAux(1).top
    If DataGridAux(1).Row < 0 Then
        anc = anc + 230
    Else
        anc = anc + DataGridAux(1).RowTop(DataGridAux(1).Row) + 5
    End If
    If Modo = 1 Then
        LLamaLineas 1, Modo, anc
    Else
        LLamaLineas 1, 3, anc
    End If
    
    For i = 0 To txtaux.Count - 1
        txtaux(i).BackColor = vbWhite
    Next i
    
    Frame4.Enabled = (Modo = 1)
    
    
    txtaux(8).Enabled = (Modo = 1)
    txtaux(9).Enabled = (Modo = 1)
    
    ' numero de asiento
    Text1(8).Enabled = (Modo = 1)
    
    
    ' ponemos en azul clarito
    Text1(0).BackColor = vbMoreLightBlue  ' factura
    Text1(13).BackColor = vbMoreLightBlue ' total factura
    Text1(4).BackColor = vbMoreLightBlue ' codmacta del cliente
    
    
        
    
    
    PonerModoUsuarioGnral Modo, "ariconta"

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub Desplazamiento(Index As Integer)
    If Data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 1
            Data1.Recordset.MoveFirst
        Case 2
            Data1.Recordset.MovePrevious
            If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
        Case 3
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
        Case 4
            Data1.Recordset.MoveLast
    End Select
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 ' lineas de totales
            tabla = "factpro_totales"
            Sql = "SELECT factpro_totales.numserie, factpro_totales.numregis, factpro_totales.fecharec, factpro_totales.anofactu, factpro_totales.numlinea, factpro_totales.baseimpo, factpro_totales.codigiva, factpro_totales.porciva,"
            Sql = Sql & " factpro_totales.porcrec, factpro_totales.impoiva, factpro_totales.imporec "
            Sql = Sql & " FROM " & tabla
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "factpro", "factpro_totales")
            Else
                Sql = Sql & " WHERE factpro_totales.numlinea is null"
            End If
            Sql = Sql & " ORDER BY 1,2,3,4,5"
            
       
       
       Case 1 ' lineas de facturas
            tabla = "factpro_lineas"
            Sql = "SELECT factpro_lineas.numserie, factpro_lineas.numregis, factpro_lineas.fecharec, factpro_lineas.anofactu, factpro_lineas.numlinea, factpro_lineas.codmacta, cuentas.nommacta, factpro_lineas.baseimpo, factpro_lineas.codigiva,"
            Sql = Sql & " factpro_lineas.porciva, factpro_lineas.porcrec, factpro_lineas.impoiva, factpro_lineas.imporec, factpro_lineas.aplicret, IF(factpro_lineas.aplicret=1,'*','') as daplicret, factpro_lineas.codccost, ccoste.nomccost "
            Sql = Sql & " FROM (factpro_lineas LEFT JOIN ccoste ON factpro_lineas.codccost = ccoste.codccost) "
            Sql = Sql & " INNER JOIN cuentas ON factpro_lineas.codmacta = cuentas.codmacta "
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "factpro", "factpro_lineas")
            Else
                Sql = Sql & " WHERE factpro_lineas.numlinea is null"
            End If
            Sql = Sql & " ORDER BY 1,2,3,4,5"
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Sub frmAgen_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(26).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmPag_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        CtaBanco = RecuperaValor(CadenaSeleccion, 1)
        IBAN = RecuperaValor(CadenaSeleccion, 6)
        
        Pagado = RecuperaValor(CadenaSeleccion, 7)
        FechaPago = RecuperaValor(CadenaSeleccion, 8)
    End If
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
Dim vFe As String

    'Cuentas
    vFe = RecuperaValor(CadenaSeleccion, 3)
    If vFe <> "" Then
        vFe = RecuperaValor(CadenaSeleccion, 1)
        If EstaLaCuentaBloqueada(vFe, CDate(Text1(1).Text)) Then
            MsgBox "Cuenta bloqueada: " & vFe, vbExclamation
        End If
    End If

    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    Text4(4).Text = RecuperaValor(CadenaSeleccion, 2)
        


End Sub

Private Sub frmCtasRet_DatoSeleccionado(CadenaSeleccion As String)
Dim vFe As String

    'Cuenta de retencion
    vFe = RecuperaValor(CadenaSeleccion, 3)
    If vFe <> "" Then
        vFe = RecuperaValor(CadenaSeleccion, 1)
        If EstaLaCuentaBloqueada(vFe, CDate(Text1(1).Text)) Then
            MsgBox "Cuenta bloqueada: " & vFe, vbExclamation
        End If
    End If
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1)
    Text4(6).Text = RecuperaValor(CadenaSeleccion, 2)
        

End Sub

Private Sub frmDpto_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(25).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(25).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmFact_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "numserie = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T") & " and numregis = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "N") & " and anofactu = year(" & DBSet(RecuperaValor(CadenaSeleccion, 3), "F") & ") "
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.cmdAux(0).Tag + 2)
    txtaux(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(5).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
Dim RC As String

    'Tipos de Iva
    txtaux(7).Text = RecuperaValor(CadenaSeleccion, 1)
    RC = "porcerec"
    txtaux(8).Text = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", txtaux(7), "N", RC)
    PonerFormatoDecimal txtaux(8), 4
    If RC = 0 Then
        txtaux(9).Text = ""
    Else
        txtaux(9).Text = RC
    End If
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgFec_Click(Index As Integer)
       
       Screen.MousePointer = vbHourglass
       
       Dim esq As Long
       Dim dalt As Long
       Dim menu As Long
       Dim Obj As Object
    
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
    
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
Dim vFe As String

    'Cuentas
    vFe = RecuperaValor(CadenaSeleccion, 3)
    If vFe <> "" Then
        vFe = RecuperaValor(CadenaSeleccion, 1)
        If EstaLaCuentaBloqueada(vFe, CDate(Text1(1).Text)) Then
            MsgBox "Cuenta bloqueada: " & vFe, vbExclamation
            If cmdAux(0).Tag = "0" Then txtaux(4).Text = ""
            Exit Sub
        End If
    End If
    If cmdAux(0).Tag = 5 Then
        'Cuenta normal
        txtaux(5).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
        
        'Habilitaremos el ccoste
        HabilitarCentroCoste
        
    Else
        'contrapartida
        txtaux(5).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
    End If

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    'Centro de coste
    txtaux(12).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(12).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(21).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(21).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmUtil_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion = "" Then
        ' no hacemos nada
    Else
        cboFiltro.ListIndex = 0
        
        Sql = "Select * from factpro "
        Sql = Sql & " WHERE numserie = " & RecuperaValor(CadenaSeleccion, 1)
        Sql = Sql & " AND numfactu =" & RecuperaValor(CadenaSeleccion, 2)
        Sql = Sql & " AND anofactu= " & RecuperaValor(CadenaSeleccion, 3)
        
        CadenaConsulta = Sql
        PonerCadenaBusqueda
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    If (Modo = 2 Or Modo = 5 Or Modo = 0) And (Index <> 6) And (Index <> 8) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        'FECHA recepcion
        Indice = 1
        
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco Text1(1)
        
    Case 1 ' contadores
        Set frmConta = New frmBasico
        AyudaContadores frmConta, Text1(Index).Text, "tiporegi REGEXP '^[0-9]+$' <> 0 and cast(tiporegi as unsigned) > 0"
        Set frmConta = Nothing
        PonFoco Text1(1)
    
    Case 2
        'Cuentas cliente
        Set frmCtas = New frmColCtas
        frmCtas.DatosADevolverBusqueda = "0|1|2|"
        frmCtas.ConfigurarBalances = 3  'NUEVO
        frmCtas.Show vbModal
        Set frmCtas = Nothing
        PonFoco Text1(4)
        
    Case 3 ' forma de pago
        Set frmFPag = New frmBasico2
        AyudaFPago frmFPag
        Set frmFPag = Nothing
        PonFoco Text1(5)
    
    Case 4
        'Cuenta retencion
        Set frmCtasRet = New frmColCtas
        frmCtasRet.DatosADevolverBusqueda = "0|1|2|"
        frmCtasRet.ConfigurarBalances = 3  'NUEVO
        frmCtasRet.Show vbModal
        Set frmCtasRet = Nothing
        PonFoco Text1(6)
        
    Case 5
        'pais
        Set frmPais = New frmBasico2
        AyudaPais frmPais
        Set frmPais = Nothing
        
    Case 6
        ' vamos al historico de apuntes
        Set frmAsi = New frmAsientosHco
        
        frmAsi.ASIENTO = Data1.Recordset!NumDiari & "|" & Data1.Recordset!fecharec & "|" & Data1.Recordset!NumAsien & "|"
        frmAsi.SoloImprimir = True
        frmAsi.Show vbModal
        
        Set frmAsi = Nothing
        
    Case 7
        'Fecha de liquidacion
        Indice = 23
        
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(23).Text <> "" Then frmF.Fecha = CDate(Text1(23).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco Text1(23)
        
    Case 8
        ' observaciones
        Screen.MousePointer = vbDefault
        
        Indice = 3
        
        Set frmZ = New frmZoom
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
        frmZ.Caption = "Observaciones Facturas Proveedor"
        frmZ.Show vbModal
        Set frmZ = Nothing
        
    Case 9
        ' fecha de factura
        Indice = 26
        
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(26).Text <> "" Then frmF.Fecha = CDate(Text1(26).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        
        PonFoco Text1(Indice)
        
    End Select
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    'BotonEliminar
    HacerToolBar 8
End Sub


Private Sub mnModificar_Click()
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub


Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonFoco Text1(2) ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            PonFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    CadB1 = ObtenerBusqueda2(Me, , 2, "FrameAux1")
    
    HacerBusqueda2
    
End Sub

Private Sub HacerBusqueda2()

    CargarSqlFiltro
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia
    ElseIf CadB <> "" Or CadB1 <> "" Or cadFiltro <> "" Then
        CadenaConsulta = "select distinct factpro.* from (" & NombreTabla & " INNER JOIN cuentas ON factpro.codmacta = cuentas.codmacta)  "
        CadenaConsulta = CadenaConsulta & " left join factpro_lineas on factpro.numserie = factpro_lineas.numserie and factpro.numregis = factpro_lineas.numregis and factpro.anofactu = factpro_lineas.anofactu "
        CadenaConsulta = CadenaConsulta & " WHERE (1=1) "
        If CadB <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB & " "
        If CadB1 <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB1 & " "
        If cadFiltro <> "" Then CadenaConsulta = CadenaConsulta & " and " & cadFiltro & " "
        
        CadenaConsulta = CadenaConsulta & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonFoco Text1(0)
        ' **********************************************************************
    End If
    
End Sub


Private Sub MandaBusquedaPrevia()
Dim cWhere As String
Dim cWhere1 As String

    cWhere = "(numserie, numregis, anofactu) in (select factpro.numserie, factpro.numregis, factpro.anofactu from "
    cWhere = cWhere & "factpro LEFT JOIN factpro_lineas ON factpro.numserie = factpro_lineas.numserie and factpro.anofactu = factpro_lineas.anofactu and factpro.numregis = factpro_lineas.numregis "
    cWhere = cWhere & " WHERE (1=1) "
    
    cWhere1 = ""
    If CadB <> "" Then cWhere1 = cWhere1 & " and " & CadB & " "
    If CadB1 <> "" Then cWhere1 = cWhere1 & " and " & CadB1 & " "
    If cadFiltro <> "" Then cWhere1 = cWhere1 & " and " & cadFiltro & " "
    
    If Trim(cWhere1) <> "and (1=1)" Then
        cWhere = cWhere & cWhere1 & ")"
    Else
        cWhere = ""
    End If
    
    Set frmFact = New frmFacturasProPrev
    frmFact.cWhere = cWhere
    frmFact.DatosADevolverBusqueda = "0|1|2|"
    frmFact.Show vbModal
    
    Set frmFact = Nothing

        
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
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


Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    CadB1 = ""
    
    PonerModo 0
    
    HacerBusqueda2
    
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
    'Contador de facturas
    Set Mc = New Contadores
    
    PonerModo 3
    
    Combo1(0).ListIndex = 0
    Combo1(1).ListIndex = 0
    Combo1(2).ListIndex = 0
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(9).Text = "0,00"
    
    ' por defecto para todos cuando insertamos es 1
    Text1(2).Text = "1"
    Text1_LostFocus (2)
    
    FrameDatosFiscales.Visible = False
    
    Text1_LostFocus (1)
    PonFoco Text1(2)
    ' ***********************************************************
    
End Sub


Private Sub BotonModificar()

'    If Not SePuedeModificarAsiento(True) Then Exit Sub

    
    '---------
    'MODIFICAR
    '----------
    'Comprobamos la fecha pertenece al ejercicio
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk >= 2 Then
        If varFecOk = 2 Then
            Sql = varTxtFec
        Else
            Sql = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    
    If Not ComprobarPeriodo2(23) Then Exit Sub

    PonerModo 4

    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonFoco Text1(1)
    ' *********************************************************
    
    FecFactuAnt = Text1(26).Text
    FecRecepAnt = Text1(1).Text
    NumFactuAnt = Text1(25).Text
    CodmactaAnt = Text1(4).Text
    
    NumDiario = 0
    'Comprobamos que no esta actualizada ya
    If Not IsNull(Data1.Recordset!NumAsien) Then
        Numasien2 = Data1.Recordset!NumAsien
        If Numasien2 = 0 Then
            MsgBox "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
            Exit Sub
        End If

        Numasien2 = Data1.Recordset!NumAsien
        NumDiario = Data1.Recordset!NumDiari
    Else
        Numasien2 = -1
    End If
        
        
    'Si viene a esta factura buscando por un campo k no sea clave entonces no le dejo seguir
    If InStr(1, Data1.Recordset.Source, "numasien") Then
        MsgBox "Busque la factura por su numero de factura", vbExclamation
        Numasien2 = -1
        Exit Sub
    End If
    
    'Llegados aqui bloqueamos desde form
'--    If Not BloqueaRegistroForm(Me) Then Exit Sub

    If Numasien2 >= 0 Then
        'Tengo desintegrar la factura del hco
        If Not Desintegrar Then
'--            DesBloqueaRegistroForm Me.Text1(0)
            TerminaBloquear
            Exit Sub
        End If
        Text1(8).Text = ""
    End If
    
    If Mc Is Nothing Then Set Mc = New Contadores
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    DespalzamientoVisible False
    PonFoco Text1(1)
    
    
    
End Sub


Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim i As Long
    Dim Fec As Date
    Dim Mc As Contadores
    Dim SqlLog As String
    
    'Ciertas comprobaciones
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    DataGridAux(1).Enabled = False

    'Comprobamos la fecha pertenece al ejercicio
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk >= 2 Then
        If varFecOk = 2 Then
            Sql = varTxtFec
        Else
            Sql = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBox Sql, vbExclamation
        Exit Sub
    End If

    'Comprobamos si esta liquidado
    If Not ComprobarPeriodo2(23) Then Exit Sub
    
    'Comprobamos que no esta actualizada ya
    Sql = ""
    If Not IsNull(Data1.Recordset!NumAsien) Then
        Sql = "Esta factura ya esta contabilizada. "
    End If
    
    Sql = Sql & vbCrLf & vbCrLf & "Va usted a eliminar la factura :" & vbCrLf
    Sql = Sql & "Numero : " & Data1.Recordset!NumFactu & vbCrLf
    Sql = Sql & "Fecha  : " & Data1.Recordset!FecFactu & vbCrLf
    Sql = Sql & "Proveedor : " & Me.Data1.Recordset!codmacta & " - " & Text4(4).Text & vbCrLf
    Sql = Sql & vbCrLf & "          ¿Desea continuar ?" & vbCrLf
    
    If Not EliminarDesdeActualizar Then
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    Screen.MousePointer = vbHourglass
    'Lo hara en actualizar
    i = 0
    If Not IsNull(Data1.Recordset!NumAsien) Then i = Data1.Recordset!NumAsien
    If i > 0 Then
        
            'Memorizamos el numero de asiento y la fechaent para ver si devolvemos el contador
            'de asientos
            i = Data1.Recordset!NumAsien
            Fec = Data1.Recordset!FechaEnt
        
            'La borrara desde actualizar
            AlgunAsientoActualizado = False
           
            
            SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
            SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text
            SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text

            
            With frmActualizar
                .OpcionActualizar = 9
                .NumAsiento = Data1.Recordset!NumAsien
                .NumFac = Data1.Recordset!NumRegis
                .FACTURA = Data1.Recordset!NumFactu
                .Proveedor = Data1.Recordset!codmacta
                .FechaAsiento = Data1.Recordset!fecharec
                .FechaFactura = Data1.Recordset!FecFactu
                .NUmSerie = Data1.Recordset!NUmSerie & "|" & Data1.Recordset!anofactu & "|"
                .NumDiari = Data1.Recordset!NumDiari
                .FechaAnterior = Data1.Recordset!fecharec
                .SqlLog = SqlLog
                .Show vbModal
            End With
            Set Mc = New Contadores
            Mc.DevolverContador "0", Fec <= vParam.fechafin, i
            Set Mc = Nothing
        
    Else
        'La borrara desde este mismo form
        Conn.BeginTrans
        
        i = Data1.Recordset!NumRegis
        Fec = Data1.Recordset!fecharec
        If BorrarFactura Then
            'LOG
            SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
            SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text
            SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text
            
            vLog.Insertar 10, vUsu, SqlLog
        
            AlgunAsientoActualizado = True
            Conn.CommitTrans
            Set Mc = New Contadores
            Mc.DevolverContador CStr(DBLet(Data1.Recordset!NUmSerie)), (Fec <= vParam.fechafin), i
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
        CargaGrid 1, False
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
    Screen.MousePointer = vbDefault
End Sub

Private Function BorrarFactura() As Boolean
    
    On Error GoTo EBorrar
    Sql = " WHERE numserie = '" & Data1.Recordset!NUmSerie & "'"
    Sql = Sql & " AND numregis = " & Data1.Recordset!NumRegis
    Sql = Sql & " AND anofactu= " & Data1.Recordset!anofactu
    'Las lineas
    AntiguoText1 = "DELETE from factpro_totales " & Sql
    Conn.Execute AntiguoText1
    AntiguoText1 = "DELETE from factpro_lineas " & Sql
    Conn.Execute AntiguoText1
    'La factura
    AntiguoText1 = "DELETE from factpro " & Sql
    Conn.Execute AntiguoText1
    
    ComprobarContador Data1.Recordset!NUmSerie, CDate(Text1(1).Text), Data1.Recordset!NumRegis
    
EBorrar:
    If Err.Number = 0 Then
        BorrarFactura = True
    Else
        MuestraError Err.Number, "Eliminar factura"
        BorrarFactura = False
    End If
End Function


Private Function ComprobarContador(LEtra As String, Fecha As Date, NumeroFAC As Long)
Dim Mc As Contadores
Dim B As Byte
On Error GoTo EComr

    Set Mc = New Contadores
    B = FechaCorrecta2(Fecha)
    Mc.DevolverContador LEtra, B = 0, NumeroFAC
    
EComr:
    If Err.Number <> 0 Then MuestraError Err.Number, "Devolviendo contador."
    Set Mc = Nothing
End Function



Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For i = 1 To DataGridAux.Count ' - 1
        If i <> 3 Then
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
        End If
    Next i
    
    imgppal(6).Enabled = (Text1(8).Text <> "")
    imgppal(6).Visible = (Text1(8).Text <> "")
        
    Text4(2).Text = PonerNombreDeCod(Text1(2), "contadores", "nomregis", "tiporegi", "T")
    Text4(4).Text = PonerNombreDeCod(Text1(4), "cuentas", "nommacta", "codmacta", "T")
    Text4(6).Text = PonerNombreDeCod(Text1(6), "cuentas", "nommacta", "codmacta", "T")
    Text4(5).Text = PonerNombreDeCod(Text1(5), "formapago", "nomforpa", "codforpa", "N")
    Text4(21).Text = PonerNombreDeCod(Text1(21), "paises", "nompais", "codpais", "T")
    
    If Text1(22).Text = "0" Then
        Combo1(0).ListIndex = 0
    Else
        PosicionarCombo Combo1(0), Asc(Text1(22).Text)
    End If
    
    'intracomunitaria
    Combo1_Validate 1, False
    
    If Text1(27).Text = "" Then
        Combo1(3).ListIndex = -1
    Else
        PosicionarCombo Combo1(3), Asc(Text1(27).Text)
    End If
    
    CargaDatosLW

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
End Sub


Private Sub cmdCancelar_Click()
Dim i As Integer
Dim v

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
            'Contador de facturas
            If Modo = 3 Then
                'Intentetamos devolver el contador
                If Text1(0).Text <> "" Then
                    i = FechaCorrecta2(CDate(Text1(0).Text))
                    Mc.DevolverContador Mc.TipoContador, i = 0, Mc.Contador
                End If
            End If
            LimpiarCampos
            PonerModo 0
            Set Mc = Nothing

        Case 4  'Modificar
            Modo = 2   'Para que el lostfocus NO haga nada
            If Numasien2 > 0 Then
                'Ha cancelado. Tendre que situar los campos correctamente
                'Es decir. Anofacl
                Text1(1).Text = Data1.Recordset!fecharec
                Text1(0).Text = Data1.Recordset!NumRegis
                Text1(14).Text = Data1.Recordset!anofactu
                If Not IntegrarFactura Then
                    Modo = 4 'lo pongo por si acaso
                    Exit Sub
                End If
            End If
            PonerCampos
            Modo = 4  'Reestablezco el modo para que vuelva a hahacer ponercampos
            '--DesBloqueaRegistroForm Me.Text1(0)
            TerminaBloquear
            
            PonerModo 2
            'Contador de facturas
            Set Mc = Nothing
                
        Case 5 'LLÍNIES
            TerminaBloquear
        
            If ModoLineas = 1 Then 'INSERTAR
                ModoLineas = 0
                DataGridAux(1).AllowAddNew = False
                If Not AdoAux(1).Recordset.EOF Then AdoAux(1).Recordset.MoveFirst
                
                If AdoAux(1).Recordset.EOF Then
                    If MsgBox("No se permite una factura sin líneas " & vbCrLf & vbCrLf & "¿ Desea eliminar la factura ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        BotonEliminar True
                        Exit Sub
                    Else
                        ModoLineas = 1
                        cmdAceptar_Click
                        Exit Sub
                    End If
                End If
                
            End If
            ModoLineas = 0
            LLamaLineas 1, 0, 0
            
            Modo = 2   'Para que el lostfocus NO haga nada
            If Numasien2 > 0 Then
                'Ha cancelado. Tendre que situar los campos correctamente
                'Es decir. Anofacl
                Text1(1).Text = Data1.Recordset!fecharec
                Text1(0).Text = Data1.Recordset!NumRegis
                Text1(14).Text = Data1.Recordset!anofactu
                If Not IntegrarFactura Then
                    Modo = 4 'lo pongo por si acaso
                    Exit Sub
                End If
                PagosTesoreria
            Else
                ' cogemos un nro.de asiento para integrarlo
                Set Mc = New Contadores
                
                i = FechaCorrecta2(CDate(Text1(1).Text))
                If Mc.ConseguirContador("0", (i = 0), False) = 0 Then
                    Text1(8).Text = Format(Mc.Contador, "0000000")
                    Numasien2 = Mc.Contador
                    If ModificaDesdeFormulario2(Me, 2, "Frame2") Then
                        If Not IntegrarFactura Then
                            Modo = 4
                            Exit Sub
                        End If
                        PagosTesoreria
                    End If
                Else
                    Mc.DevolverContador "0", (i = 0), CLng(Text1(8).Text)
                End If
                
            End If
            
            PosicionarData
            PonerCampos
            
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Sql As String
Dim cad As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOK = False
    
    'fecha de liquidacion
    Text1(23).Text = Text1(1).Text
    
    If Combo1(0).ListIndex = 0 Then
        Text1(22).Text = "0"
    Else
        Text1(22).Text = Chr(Combo1(0).ItemData(Combo1(0).ListIndex))
    End If

    
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' NOV 2007
    ' NUEVA ambitode fecha activa
    '       0 .- Año actual
    '       1 .- Siguiente
    '       2 .- Fuera de ambito  !!! NUEVO !!!
    '       2 .- Anterior al inicio
    '       3 .- Posterior al fin
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            MsgBox "La fecha no pertenece al ejercicio actual ni al siguiente", vbExclamation
        End If
        B = False

    End If
    
    ' controles añadidos de la factura de david
    'No puede tener % de retencion sin cuenta de retencion
    If Combo1(2).ListIndex > 0 Then
       If ((Text1(6).Text = "") Or (Text1(7).Text = "")) Then
            MsgBox "Indique porcentaje y cuenta de rentención ", vbExclamation
            B = False
            PonFoco Text1(6)
            Exit Function
        End If
    Else
        If ((Text1(6).Text <> "") Or (Text1(7).Text <> "")) Then
            MsgBox "Ha indicado porcentaje y/o cuenta de rentención sin indicar el tipo", vbExclamation
            B = False
            PonFoco Text1(6)
            Exit Function
        End If
        
    End If
    
    'Compruebo si hay fechas bloqueadas
    If vParam.CuentasBloqueadas <> "" Then ' cuenta cliente
        If EstaLaCuentaBloqueada(Text1(4).Text, CDate(Text1(1).Text)) Then
            MsgBox "Cuenta bloqueada: " & Text1(4).Text, vbExclamation
            B = False
            Exit Function
        End If
        If Text1(6).Text <> "" Then ' cuenta de retencion
            If EstaLaCuentaBloqueada(Text1(6).Text, CDate(Text1(1).Text)) Then
                MsgBox "Cuenta bloqueada: " & Text1(6).Text, vbExclamation
                B = False
                Exit Function
            End If
        End If
    End If
    
    
    'Ahora. Si estamos modificando, y el año factura NO es el mismo, entonces
    'la estamos liando, y para evitar lios, NO dejo este tipo de modificacion
    If Modo = 4 Then
        If CDate(Text1(1).Text) <> Data1.Recordset!fecharec Then
            'HAN CAMBIADO LA FECHA. Veremos si dejo
            If Year(CDate(Text1(1).Text)) <> Data1.Recordset!anofactu Then
                MsgBox "No puede cambiar de año la factura. ", vbExclamation
                B = False
                Exit Function
            End If
            
            '[Monica]19/01/2017
            ' si hay alguna factura de la serie con numero de registro mayor y fecha de recepcion inferior a la que hemos introducido damos aviso
            Sql = "select count(*) from factpro where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis > " & DBSet(Text1(0).Text, "N") & " and fecharec < " & DBSet(Text1(1).Text, "F") & " and anofactu = " & Year(CDate(Text1(1).Text))
            If DevuelveValor(Sql) <> 0 Then
                If MsgBox("Existe alguna factura de la serie con nro.registro superior y fecha de recepción inferior a ésta." & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    B = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    
    'la forma de pago ha de existir
    If Text4(5).Text = "" And (Modo = 3 Or Modo = 4) Then
        MsgBox "No existe a forma de pago. Revise.", vbExclamation
        B = False
        PonFoco Text1(5)
        Exit Function
    End If
    
    'comprobamos que si la factura es intracomunitaria tiene que tener valor el tipo de intracomunitaria
    If Modo = 3 Or Modo = 4 Then
        If Combo1(1).ListIndex = 1 Then
            If Combo1(3).ListIndex = -1 Then
                MsgBox "Debe seleccionar un tipo de intracomunitaria. Revise.", vbExclamation
                B = False
                Combo1(3).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Modo = 3 Then
        Sql = "select count(*) from factpro where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
        Sql = Sql & " and anofactu = year(" & DBSet(Text1(1).Text, "F") & ") and codmacta = " & DBSet(Text1(4).Text, "T")
        If DevuelveValor(Sql) <> 0 Then
            MsgBox "Factura ya existe para esta serie proveedor año. Revise.", vbExclamation
            B = False
            Exit Function
        End If
    End If
    





    DatosOK = B

EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(numserie=" & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N") & ") "
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub


Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    Conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE (numasien=" & Trim(Text1(0).Text) & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N") & ") "
        ' ***********************************************************************
        
    Conn.Execute "DELETE FROM hlinapu " & vWhere
    
    Conn.Execute "DELETE FROM hcabapu_fichdocs " & vWhere

'    ' *******************************
    Conn.Execute "Delete from " & NombreTabla & vWhere
       
    'El LOG
    vLog.Insertar 3, vUsu, Sql
       
    PagosTesoreria
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

Dim RC As String
Dim Correcto As Boolean
Dim Valor As Currency
Dim L As Long
Dim i As Integer
Dim J As Integer



    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    If (Index = 13 Or Index = 0 Or Index = 4) And Modo = 1 Then
        Text1(Index).BackColor = vbMoreLightBlue ' azul clarito
    End If

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 4
            Text4(Index) = ""
        Case 5
            Text4(Index) = ""
            
        Case 6
            Text4(Index) = ""
            If Text1(Index).Text = "" Then Text1(7).Text = ""
        Case 21
            Text4(Index) = ""
    End Select
    
    If Text1(Index).Text = "" Then Exit Sub
    
    If Modo = 5 Then Exit Sub
    
    Select Case Index
        Case 0 'Nro de factura
            PonerFormatoEntero Text1(Index)

        Case 1, 23 '1 - fecha de factura
                   '23- fecha de liquidacion
            Sql = ""
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                If Index = 1 Then Text1(14).Text = ""
                PonFoco Text1(Index)
                Exit Sub
            End If
            ModificandoLineas = FechaCorrecta2(CDate(Text1(Index).Text))
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
                Text1(Index).Text = ""
                If Index = 1 Then Text1(14).Text = ""
                PonFoco Text1(Index)
                Exit Sub
            End If
            
            Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
            If Index = 1 Then Text1(14).Text = Year(CDate(Text1(Index).Text))
            
            If Index = 1 And Modo <> 1 Then Text1(23).Text = Text1(1).Text
            
            'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
            'que el periodo de liquidacion del IVA no ha pasado.
            i = 0
            If vParam.Constructoras Then
                If Index = 23 Then i = 1
            Else
                If Index = 1 Then i = 1
            End If
            If i > 0 Then
                If Not ComprobarPeriodo2(Index) Then PonFoco Text1(Indice)
            End If

        Case 2 ' Serie
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Debe ser un número: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                Text4(2).Text = ""
                PonFoco Text1(2)
                Exit Sub
            End If
            If Text1(Index).Text = AntiguoText1 Then Exit Sub

            Text4(2).Text = DevuelveValor("select nomregis from contadores where tiporegi = " & DBSet(Text1(2).Text, "T") & " and tiporegi REGEXP '^[0-9]+$' <> 0 and cast(tiporegi as unsigned) > 0 ")
            If Text4(2).Text = "0" Then
                Text4(2).Text = ""
                MsgBox "Letra de serie no existe o no es de facturas de proveedor. Reintroduzca.", vbExclamation
                Text1(2).Text = ""
                PonFoco Text1(2)
            End If
        
        Case 4, 6 ' cuenta de proveedor, cuenta de retencion
                'Cuenta proveedor
                If AntiguoText1 = Text1(Index).Text Then Exit Sub
                RC = Text1(Index).Text
                i = Index
                
                If CuentaCorrectaUltimoNivel(RC, Sql) Then
                    Text1(Index).Text = RC
                    Text4(i).Text = Sql
                    If Text1(1).Text <> "" Then
                        If Modo > 2 Then
                            If EstaLaCuentaBloqueada(RC, CDate(Text1(1).Text)) Then
                                MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                                Text1(Index).Text = ""
                                Text4(i).Text = ""
                                PonFoco Text1(Index)
                                Exit Sub
                            End If
                        End If
                    End If
                    If Index = 4 And (Modo = 3 Or (Modo = 4 And Trim(Text1(Index).Text) <> AntiguoText1)) Then
                        CargarDatosCuenta Text1(Index)
                    End If
                    RC = ""
                Else
                    
                    If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                            RC = RellenaCodigoCuenta(Text1(Index).Text)
                            Sql = "La cuenta: " & RC & " no existe. ¿Desea crearla?"
                            If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
                                CadenaDesdeOtroForm = RC
                                cmdAux(0).Tag = Indice
                                Set frmC = New frmColCtas
                                frmC.DatosADevolverBusqueda = "0|1|"
                                frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                                frmC.Show vbModal
                                Set frmC = Nothing
                                If Text1(4).Text = RC Then Sql = "" 'Para k no los borre
                            End If
                    Else
                        'Cualquier otro error
                        'menos si no estamos buscando, k dejaremos
                        If Modo = 1 Then
                            Sql = ""
                        Else
                            MsgBox Sql, vbExclamation
                        End If
                    End If
                    
                    If Sql <> "" Then
                        Text1(Index).Text = ""
                        Text4(i).Text = ""
                        PonFoco Text1(Index)
                    End If
                    
                    
                End If
        
        
        Case 5 ' forma de pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text4(Index).Text = PonerNombreDeCod(Text1(Index), "formapago", "nomforpa", "codforpa", "N")
                If Text4(Index).Text = "" Then
                    MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
        
        Case 7 ' % de retencion
            PonerFormatoDecimal Text1(Index), 4
        
        Case 21 ' codigo de pais
            If Text1(Index).Text <> "" Then
                Text4(Index).Text = PonerNombreDeCod(Text1(Index), "paises", "nompais", "codpais", "T")
                If Text4(Index) = "" Then
                    MsgBox "No existe el País. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
        
        Case 25 ' numero de factura
            
        Case 26 ' fecha de factura
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                PonFoco Text1(Index)
                Exit Sub
            End If
        
    End Select
End Sub

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 3 Then
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 1:  KEYBusqueda KeyAscii, 0 ' fecha de factura
                Case 4:  KEYBusqueda KeyAscii, 2 ' cuenta proveedor
                Case 6:  KEYBusqueda KeyAscii, 4 ' cuenta de retencion
                Case 5:  KEYBusqueda KeyAscii, 3 ' forma de pago
                Case 2:  KEYBusqueda KeyAscii, 1 ' serie
                Case 21: KEYBusqueda KeyAscii, 5 ' pais
                Case 26: KEYBusqueda KeyAscii, 10 ' fecha de factura
            End Select
         Else
            KEYpress KeyAscii
         End If
    Else
        If Index <> 3 Or (Index = 3 And Text1(Index) = "") Then KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgppal_Click (Indice)
End Sub
'++

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Datos Fiscales
            Me.FrameDatosFiscales.Visible = Not Me.FrameDatosFiscales.Visible
           
        Case 2 'Cartera de Cobros
            If Not Data1.Recordset.EOF Then
                Set frmMens = New frmMensajes
                
                frmMens.Opcion = 28
                frmMens.Parametros = Trim(Text1(2).Text) & "|" & Trim(Text1(4).Text) & "|" & Trim(Text1(25).Text) & "|" & Text1(26).Text & "|"
                frmMens.Show vbModal
                
                Set frmMens = Nothing
            End If
    
        Case 3
            Screen.MousePointer = vbHourglass
            
            Set frmUtil = New frmUtilidades
            
            frmUtil.Opcion = 6
            frmUtil.Show vbModal

            Set frmUtil = Nothing
            
        Case 4
            ComprobarFrasSinAsiento

            
    End Select

End Sub

Private Sub ComprobarFrasSinAsiento()
Dim Sql As String
Dim vCadena As String
Dim vCadena2 As String
Dim Rs As ADODB.Recordset
Dim IntegrarFactura As Boolean
Dim i As Integer
Dim Nregs As Long
Dim SqlLog As String

    
    Sql = "select * from factpro where (numasien = 0 or numasien is null or fechaent is null or numdiari is null) "
    If cadFiltro <> "" Then Sql = Sql & " and " & cadFiltro

    vCadena = ""
    vCadena2 = ""
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Nregs = 1
        
        While Not Rs.EOF
            vCadena = vCadena & "Fra.Reg. " & DBLet(Rs!NUmSerie) & " " & Format(DBLet(Rs!NumRegis), "0000000") & " " & DBLet(Rs!fecharec, "F")
            vCadena2 = vCadena2 & "(" & DBSet(Rs!NUmSerie, "T") & "," & Format(DBSet(Rs!NumRegis, "N"), "0000000") & "," & Year(DBLet(Rs!fecharec, "F")) & "),"
            
            If (Nregs Mod 2) = 0 Then
                vCadena = vCadena & vbCrLf
            Else
                vCadena = vCadena & "  "
            End If
            
            Nregs = Nregs + 1
            
            Rs.MoveNext
        Wend
        
        If MsgBox("Las siguientes facturas no tienen Asiento asociado: " & vbCrLf & vbCrLf & vCadena & vbCrLf & vbCrLf & " ¿ Asigna asiento ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            Rs.MoveFirst
            
            
            While Not Rs.EOF
                IntegrarFactura = False
                
                ' cogemos un nro.de asiento para integrarlo
                Set Mc = New Contadores
                
                i = FechaCorrecta2(CDate(DBLet(Rs!FecFactu, "F")))
                If Mc.ConseguirContador("0", (i = 0), False) = 0 Then
                    
                    Numasien2 = Mc.Contador
                
                    SqlLog = "Registro : " & Rs!NUmSerie & " " & Rs!NumRegis & " de fecha " & Rs!fecharec
                    SqlLog = SqlLog & vbCrLf & "Cuenta  : " & DBLet(Rs!codmacta, "T") & " " & DBLet(Rs!Nommacta, "T")
                    SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Rs!totfacpr, "N")
                    
                    With frmActualizar
                        .OpcionActualizar = 8
                        'NumAsiento     --> CODIGO FACTURA
                        'NumDiari       --> AÑO FACTURA
                        'NUmSerie       --> SERIE DE LA FACTURA
                        'FechaAsiento   --> Fecha factura
                        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
                        .NumFac = DBLet(Rs!NumRegis, "N")
                        .NumDiari = Year(DBLet(Rs!fecharec, "F"))
                        .NUmSerie = DBLet(Rs!NUmSerie, "T")
                        .FechaAsiento = DBLet(Rs!fecharec, "F")
                        .FechaAnterior = DBLet(Rs!fecharec, "F")
                        .SqlLog = "" 'SqlLog
                        
                        
                        If NumDiario <= 0 Then NumDiario = vParam.numdiapr
                        .DiarioFacturas = NumDiario
                        .NumAsiento = Numasien2
                        .Show vbModal
                        
                        If AlgunAsientoActualizado Then IntegrarFactura = True
                        
                        Screen.MousePointer = vbHourglass
                        Me.Refresh
                    End With
                
                    If IntegrarFactura Then
                        Sql = "update factpro set numdiari = " & DBSet(NumDiario, "N") & ", fechaent = " & DBSet(Rs!FecFactu, "F") & ", "
                        Sql = Sql & " numasien = " & DBSet(Numasien2, "N") & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and anofactu = year("
                        Sql = Sql & DBSet(Rs!fecharec, "F") & ") and numregis = " & DBSet(Rs!NumRegis, "N")
                    
                        Conn.Execute Sql
                        
                    End If
                End If
                
                Rs.MoveNext
            Wend
        
            vLog.Insertar 29, vUsu, vCadena
        
        End If
        
        
        Set Rs = Nothing
        
        CadB = "(factpro.numserie, factpro.numregis, factpro.anofactu) in (" & Mid(vCadena2, 1, Len(vCadena2) - 1) & ")"
        HacerBusqueda2
    
    Else
        MsgBox "No hay facturas sin asiento asignado.", vbExclamation
    End If

End Sub






'************* LLINIES: ****************************
Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LINASI As Long
Dim Ampliacion As String

    'Comprobamos la fecha pertenece al ejercicio
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk >= 2 Then
        If varFecOk = 2 Then
            Sql = varTxtFec
        Else
            Sql = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBox Sql, vbExclamation
        Exit Sub
    End If


    '**** parte correspondiente por si la factura está contabilizada
    NumDiario = 0
    'Comprobamos que no esta actualizada ya
    If Not IsNull(Data1.Recordset!NumAsien) Then
        Numasien2 = Data1.Recordset!NumAsien
        If Numasien2 = 0 Then
            MsgBox "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
            Exit Sub
        End If
            
        Numasien2 = Data1.Recordset!NumAsien
        NumDiario = Data1.Recordset!NumDiari
    Else
        Numasien2 = -1
    End If
    
    If Not ComprobarPeriodo2(23) Then Exit Sub
    
    'Llegados aqui bloqueamos desde form
    If Not BLOQUEADesdeFormulario2(Me, Data1, 1) Then Exit Sub
    
    FecFactuAnt = Text1(26).Text
    FecRecepAnt = Text1(1).Text
    

    If Numasien2 >= 0 Then
        'Tengo desintegrar la factura del hco
        If Not Desintegrar Then
            TerminaBloquear
            Exit Sub
        End If
        Text1(8).Text = ""
    End If
    ' ***** hasta aqui, si la factura estaba contabilizada


    'Fuerzo que se vean las lineas
    
    Select Case Button.Index
        Case 1
            'AÑADIR linea factura
            BotonAnyadirLinea 1, True
        Case 2
            'MODIFICAR linea factura
            BotonModificarLinea 1
        Case 3
            'ELIMINAR linea factura
            BotonEliminarLinea 1
            PagosTesoreria
            
    End Select


End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim SqlLog As String

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'linea de asiento
            Sql = "¿Seguro que desea eliminar la línea de la factura?"
            Sql = Sql & vbCrLf & "Serie: " & AdoAux(Index).Recordset!NUmSerie & " - " & AdoAux(Index).Recordset!NumRegis & " - " & AdoAux(Index).Recordset!fecharec & " - " & AdoAux(Index).Recordset!NumLinea
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM factpro_lineas "
                Sql = Sql & Replace(vWhere, "factpro", "factpro_lineas") & " and numlinea = " & DBLet(AdoAux(Index).Recordset!NumLinea, "N")
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute Sql
        
        RecalcularTotales
        
        '**** parte de contabilizacion de la factura
        TerminaBloquear
        
        If Numasien2 > 0 Then
            If IntegrarFactura Then
                Text1(8).Text = Format(Numasien2, "0000000")
                Numasien2 = -1
                NumDiario = 0
            Else
                B = False
            End If
        End If
        
        'LOG
        SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
        SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text
        SqlLog = SqlLog & vbCrLf & "Línea   : " & DBLet(Me.AdoAux(1).Recordset!NumLinea, "N")
        SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text
        
        
        vLog.Insertar 12, vUsu, SqlLog
        
        'Creo que no hace falta volver a situar el datagrid
        If True Then
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
            Data1.Refresh
            PonerModo 2
        Else
            PonerModo 0
        End If
        '**** hasta aqui
        
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
        ' ***************************************
'        If BLOQUEADesdeFormulario2(Me, data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto)
        ' ************************
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Sub BotonAnyadirLinea(Index As Integer, Limpia As Boolean)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer

    ModoLineas = 1 'Posem Modo Afegir Llínia

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vTabla = "factpro_lineas"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 1   'hlinapu
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = ""
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", Replace(vWhere, "factpro", "factpro_lineas"))
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)

            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 230 '248
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'lineas de factura
                    If Limpia Then
                        For i = 0 To txtaux.Count - 1
                            txtaux(i).Text = ""
                        Next i
                    End If
                    txtaux(0).Text = Text1(2).Text 'serie
                    txtaux(1).Text = Text1(0).Text 'numregis
                    txtaux(2).Text = Text1(1).Text 'fecharec
                    txtaux(3).Text = Text1(14).Text 'anofactura
                    
                    txtaux(4).Text = Format(NumF, "0000") 'linea contador
                    
                    
                    If Limpia Then
                        txtAux2(5).Text = ""
                        txtAux2(12).Text = ""
                    End If
                    
                    
                    
                    chkAux(0).Value = IIf(Combo1(2).ListIndex > 0, 1, 0)
                   
                    
                    If Limpia Then
                        PonFoco txtaux(5)
                    Else
                        PonFoco txtaux(5)
                    End If
            
                    ' traemos la cuenta de contrapartida habitual
                    PonFoco txtaux(5)

                    txtaux(5).Text = CuentaHabitual(Text1(4).Text)
                    If txtaux(5).Text <> "" Then
                        If EstaLaCuentaBloqueada(txtaux(5).Text, CDate(Text1(1).Text)) Then
                            txtaux(5).Text = ""
                        Else
                            If Not ExisteEnFactura(Text1(2).Text, Text1(0).Text, Text1(1).Text, txtaux(5).Text) Then
                                txtAux_LostFocus (5)
                                PonFoco txtaux(7)
                                txtAux_LostFocus (7)
                                PonFoco txtaux(6)
                            Else
                                txtaux(5).Text = ""
                                PonFoco txtaux(5)
                            End If
                        End If
                        
                    End If
            
            End Select

    End Select
End Sub

Private Function ExisteEnFactura(Serie As String, NumFactu As String, FecFactu As String, Cuenta As String) As Boolean
Dim Sql As String

    ExisteEnFactura = False
    
    If Serie = "" Or NumFactu = "" Or FecFactu = "" Or Cuenta = "" Then Exit Function

    Sql = "select count(*) from factpro_lineas where numserie = " & DBSet(Serie, "T") & " and numregis = " & DBSet(NumFactu, "N")
    Sql = Sql & " and fecharec = " & DBSet(FecFactu, "F") & " and codmacta = " & DBSet(Cuenta, "T")

    ExisteEnFactura = (TotalRegistros(Sql) <> 0)
    
End Function


Private Function CuentaHabitual(CtaOrigen As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    CuentaHabitual = ""
    
    Sql = "select codcontrhab from cuentas where codmacta = " & DBSet(CtaOrigen, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        CuentaHabitual = DBLet(Rs.Fields(0).Value)
    End If
    
End Function


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub


    ModoLineas = 2 'Modificar llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    Select Case Index
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If

            anc = DataGridAux(Index).top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

    End Select

    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 'lineas de facturas
            txtaux(0).Text = DataGridAux(Index).Columns(0).Text
            txtaux(1).Text = DataGridAux(Index).Columns(1).Text
            txtaux(2).Text = DataGridAux(Index).Columns(2).Text
            txtaux(3).Text = DataGridAux(Index).Columns(3).Text
            txtaux(4).Text = DataGridAux(Index).Columns(4).Text
            
            txtaux(5).Text = DataGridAux(Index).Columns(5).Text 'cuenta
            txtAux2(5).Text = DataGridAux(Index).Columns(6).Text 'denominacion
            txtaux(6).Text = DataGridAux(Index).Columns(7).Text 'baseimpo
            txtaux(7).Text = DataGridAux(Index).Columns(8).Text 'codigiva
            txtaux(8).Text = DataGridAux(Index).Columns(9).Text '%iva
            txtaux(9).Text = DataGridAux(Index).Columns(10).Text '%retencion
            txtaux(10).Text = DataGridAux(Index).Columns(11).Text 'importe iva
            txtaux(11).Text = DataGridAux(Index).Columns(12).Text 'importe retencion
            If DataGridAux(Index).Columns(13).Text = 1 Then
                chkAux(0).Value = 1 ' DataGridAux(Index).Columns(14).Text 'aplica retencion
            Else
                chkAux(0).Value = 0
            End If
            txtaux(12).Text = DataGridAux(Index).Columns(15).Text 'centro de coste
            txtAux2(12).Text = DataGridAux(Index).Columns(16).Text 'nombre centro de coste
            
            IvaCuenta = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtaux(5).Text, "N")
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    
    PonFoco txtaux(5)
    
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'lineas de factura
            For jj = 5 To txtaux.Count - 1
                txtaux(jj).Visible = B
                txtaux(jj).top = alto
            Next jj
            
            txtAux2(5).Visible = B
            txtAux2(5).top = alto
            txtAux2(12).Visible = B
            txtAux2(12).top = alto
            
            
            chkAux(0).Visible = B
            chkAux(0).top = alto
            
            For jj = 0 To 2
                cmdAux(jj).Visible = B
                cmdAux(jj).top = txtaux(5).top
                cmdAux(jj).Height = txtaux(5).Height
            Next jj
            
            If Not vParam.autocoste Then
                cmdAux(2).Visible = False
                cmdAux(2).Enabled = False
                txtaux(12).Visible = False
                txtaux(12).Enabled = False
                txtAux2(12).Visible = False
                txtAux2(12).Enabled = False
            End If
            If B Then
                'Aui es donde bloquamvamos los imprtes de IVA. Ahora NO los bloqueamos
                BloqueaTXT txtaux(10), Not B
                BloqueaTXT txtaux(11), Not B
            End If
    End Select
End Sub



Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim B As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte


    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And (Modo = 5 And ModoLineas = 1) Then  'insertar
    
    End If
    
    If B And Modo = 5 Then ' tanto si insertamos como si modificamos en lineas
        'Cuenta
        If txtaux(5).Text = "" Then
            MsgBox "Cuenta no puede estar vacia.", vbExclamation
            DatosOkLlin = False
            PonFoco txtaux(5)
            Exit Function
        End If
        
        If Not IsNumeric(txtaux(5).Text) Then
            MsgBox "Cuenta debe ser numrica", vbExclamation
            DatosOkLlin = False
            PonFoco txtaux(5)
            Exit Function
        End If
        
        If txtaux(5).Text = NO Then
            MsgBox "La cuenta debe estar dada de alta en el sistema", vbExclamation
            DatosOkLlin = False
            PonFoco txtaux(5)
            Exit Function
        End If
        
        If Not EsCuentaUltimoNivel(txtaux(5).Text) Then
            MsgBox "La cuenta no es de último nivel", vbExclamation
            DatosOkLlin = False
            PonFoco txtaux(5)
            Exit Function
        End If
        
        If IvaCuenta = "" Then
            CambiarIva = True
        Else
            If CInt(ComprobarCero(txtaux(7).Text)) <> CInt(ComprobarCero(IvaCuenta)) Then
                If MsgBox("El código de iva es distinto del de la cuenta. " & vbCrLf & " ¿ Desea modificarlo en la cuenta ? " & vbCrLf & vbCrLf, vbQuestion + vbYesNo) = vbYes Then
                    CambiarIva = True
                Else
                    CambiarIva = False
                End If
            End If
        End If
        
        'Centro de coste
        If txtaux(12).Visible Then
            If txtaux(12).Enabled Then
                If txtaux(12).Text = "" Then
                    MsgBox "Centro de coste no puede ser nulo", vbExclamation
                    PonFoco txtaux(12)
                    Exit Function
                End If
            End If
        End If
        
        ' en el caso de que sea exportacion - importacion el tipo de iva ha de ser cero
        If Combo1(1).ListIndex = 2 Then
            If ComprobarCero(txtaux(8).Text) <> 0 Then
                MsgBox "Código de iva incorrecto. Debe ser Iva a 0%. Revise.", vbExclamation
                PonFoco txtaux(7)
                Exit Function
            End If
        End If
        
    End If
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    
    SepuedeBorrar = True
End Function


' *********************************************************************************
Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 1 'lineas de facturas
                If DataGridAux(Index).Columns.Count > 2 Then
                End If
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
'    If numTab = 0 Then
'        SSTab1.Tab = 2
'    ElseIf numTab = 1 Then
'        SSTab1.Tab = 1
'    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Sub CargaFrame(Index As Integer, Enlaza As Boolean)
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub
' ****************************************


Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, Enlaza)

    B = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = Conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, Enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 350
    
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For i = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(i).AllowSizing = False
    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        
        Case 1 'lineas de asiento
            
            If vParam.autocoste Then
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(5)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(5)|T|Denominación|3995|;"
                tots = tots & "S|txtaux(6)|T|Importe|1905|;S|txtaux(7)|T|Iva|625|;S|cmdAux(1)|B|||;S|txtaux(8)|T|%Iva|765|;"
                tots = tots & "S|txtaux(9)|T|%Rec|765|;S|txtaux(10)|T|Importe Iva|1554|;S|txtaux(11)|T|Importe Rec|1554|;"
                tots = tots & "N||||0|;S|chkAux(0)|CB|Ret|400|;S|txtaux(12)|T|CC|710|;S|cmdAux(2)|B|||;S|txtAux2(12)|T|Nombre|2470|;"
            Else
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(5)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(5)|T|Denominación|5695|;"
                tots = tots & "S|txtaux(6)|T|Importe|2405|;S|txtaux(7)|T|Iva|625|;S|cmdAux(1)|B|||;S|txtaux(8)|T|%Iva|855|;"
                tots = tots & "S|txtaux(9)|T|%Rec|855|;S|txtaux(10)|T|Importe Iva|1954|;S|txtaux(11)|T|Importe Rec|1954|;"
                tots = tots & "N||||0|;S|chkAux(0)|CB|Ret|400|;N||||0|;N||||0|;"
            End If
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgLeft
            DataGridAux(Index).Columns(6).Alignment = dbgLeft
            DataGridAux(Index).Columns(14).Alignment = dbgCenter
            
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (Enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To 4
                    txtaux(i).Text = ""
                Next i
                txtAux2(5).Text = ""
                txtAux2(12).Text = ""
            End If
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************
      
    'Obtenemos las sumas
'    ObtenerSumas
    If Enlaza Then CargaDatosLW
    
    PonerModoUsuarioGnral Modo, "ariconta"

      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean
Dim Limp As Boolean
Dim cad As String



    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1"
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        Conn.BeginTrans
        
        B = True
        If CambiarIva Then B = ActualizarIva
    
        If B And InsertarDesdeForm2(Me, 2, nomframe) Then
        
            B = RecalcularTotales
            
            If B Then RecalcularTotalesFactura
        
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    
                    DataGridAux(1).AllowAddNew = False
                    
                    If Not AdoAux(1).Recordset.EOF Then PosicionGrid = DataGridAux(1).FirstRow
                    CargaGrid 1, True
                    Limp = True

                    txtaux(11).Text = ""
                    If Limp Then
                        txtAux2(5).Text = ""
                        txtAux2(12).Text = ""
                        For i = 0 To 11
                            txtaux(i).Text = ""
                        Next i
                    End If
                    ModoLineas = 0
                    If B Then
                            BotonAnyadirLinea NumTabMto, True
                    End If
            End Select
           
        Else
           Conn.RollbackTrans
        End If
    End If
End Sub

Private Function ActualizarIva() As Boolean
Dim Sql As String

    On Error GoTo eActualizarIva
    
    ActualizarIva = False
    
    Sql = "update cuentas set codigiva = " & DBSet(txtaux(7).Text, "N") & " where codmacta = " & DBSet(txtaux(5).Text, "T")
    Conn.Execute Sql
    
    ActualizarIva = True
    Exit Function
    
eActualizarIva:
    MuestraError Err.Number, "Actualizar Iva", Err.Description
End Function


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim v As Integer
Dim cad As String
Dim SqlLog As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'apuntes
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        Conn.BeginTrans
        
        B = True
        If CambiarIva Then B = ActualizarIva
    
        If B And ModificaDesdeFormulario2(Me, 2, nomframe) Then
        
            B = RecalcularTotales
            'LOG
            SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
            SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text
            SqlLog = SqlLog & vbCrLf & "Linea   : " & txtaux(4).Text
            vLog.Insertar 11, vUsu, SqlLog
        
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            ' ******************************************************
            ModoLineas = 0

            If NumTabMto <> 3 Then
                v = AdoAux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(3).Name & " =" & v)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
            
        Else
            Conn.RollbackTrans
        End If
    End If
        
End Sub




Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & "factpro.numserie=" & DBSet(Text1(2).Text, "T") & " and factpro.numregis=" & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
End Sub
' ***********************************************

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2) And DesdeNorma43 = 0
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N") And DesdeNorma43 = 0 And (Modo <> 0 And Modo <> 5)
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!especial, "N") And DesdeNorma43 = 0 And Modo = 2 And vEmpresa.TieneTesoreria
        Me.Toolbar2.Buttons(3).Enabled = DBLet(Rs!especial, "N") And DesdeNorma43 = 0 And (Modo = 2 Or Modo = 0)
        Me.Toolbar2.Buttons(4).Enabled = DBLet(Rs!especial, "N") And DesdeNorma43 = 0 And (Modo = 2 Or Modo = 0)
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        ToolbarAux.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        
        vUsu.LeerFiltros "ariconta", IdPrograma
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    AntiguoText1 = txtaux(Index).Text
    ConseguirFoco txtaux(Index), Modo
    
    If Index = 11 Then
        If ComprobarCero(txtaux(9).Text) = 0 Then
            PonerFocoChk Me.chkAux(0)
        End If
    End If
    
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5:  KEYImage KeyAscii, 0 ' cta base
            Case 7:  KEYImage KeyAscii, 1 ' iva
            Case 12:  KEYImage KeyAscii, 2 ' Centro Coste
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    cmdAux_Click (Indice)
End Sub
'++


Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Importe As Currency
        
        If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
        
        If txtaux(Index).Text = AntiguoText1 Then
             If Index = 12 And vParam.autocoste Then cmdAceptar.SetFocus
             Exit Sub
        End If
    
        
        Select Case Index
        Case 5
            RC = txtaux(5).Text
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                txtaux(5).Text = RC
                If Modo = 1 Then Exit Sub
                If EstaLaCuentaBloqueada(RC, CDate(Text1(1).Text)) Then
                    MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                    txtaux(5).Text = ""
                Else
                    txtAux2(5).Text = Sql
                    ' traemos el tipo de iva de la cuenta
                    txtaux(7).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtaux(5).Text, "N")
                    IvaCuenta = txtaux(7)
                    If txtaux(7).Text <> "" Then txtAux_LostFocus (7)
                    RC = ""
                End If
            Else
                If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                    txtaux(5).Text = RC
                    txtAux2(5).Text = ""
                    'NO EXISTE LA CUENTA, añado que debe de tener permiso de creacion de cuentas
                    If vUsu.PermiteOpcion("ariconta", 201, vbOpcionCrearEliminar) Then
                        Sql = Sql & " ¿Desea crearla?"
                        If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                            CadenaDesdeOtroForm = RC
                            cmdAux(0).Tag = Index
                            Set frmC = New frmColCtas
                            frmC.DatosADevolverBusqueda = "0|1|"
                            frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                            frmC.Show vbModal
                            Set frmC = Nothing
                            If txtaux(5).Text = RC Then
                                Sql = "" 'Para k no los borre
                                ' traemos el tipo de iva de la cuenta
                                txtaux(7).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtaux(5).Text, "N")
                                IvaCuenta = txtaux(7)
                                txtAux_LostFocus (7)
                            
                            End If
                        End If
                    Else
                        MsgBox Sql, vbExclamation
                    End If
                Else
                    MsgBox Sql, vbExclamation
                End If
                    
                If Sql <> "" Then
                  txtaux(5).Text = ""
                  txtAux2(5).Text = ""
                  RC = "NO"
                End If
            End If
            
            
            ''AHORA. Si el elemento es de inmvoilizado, entonces crearemos el elmento
            If txtAux2(5).Text <> "" Then
                If Modo = 5 Then
                    i = 0
                    If ModoLineas = 1 Then
                        i = 1
                    Else
                        If ModoLineas = 2 Then
                            'Solo si cambia la cuenta de lineas
                            If txtaux(5).Text <> DBLet(AdoAux(1).Recordset!codmacta, "T") Then i = 1
                        End If
                    End If
                    If i = 1 Then CrearElementoInmovilizado_ txtaux(5).Text
                End If
            End If
            
            HabilitarCentroCoste
            If RC <> "" Then PonFoco txtaux(5)
                
            If Modo = 5 And ModoLineas = 1 Then MostrarObservaciones txtaux(Index)
            
        Case 6
            PonerFormatoDecimal txtaux(Index), 1
            
        Case 7 ' iva
            RC = "porcerec"
            txtaux(8).Text = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", txtaux(7), "N", RC)
            If txtaux(8).Text = "" Then
                MsgBox "No existe el Tipo de Iva. Reintroduzca.", vbExclamation
                PonFoco txtaux(7)
            Else
                If RC = 0 Then
                    txtaux(9).Text = ""
                Else
                    txtaux(9).Text = RC
                End If
            End If
                
        Case 10, 11
           'LOS IMPORTES
            If PonerFormatoDecimal(txtaux(Index), 1) Then
                If Not vParam.autocoste Then cmdAceptar.SetFocus
            End If
                
        Case 12
            txtaux(12).Text = UCase(txtaux(12).Text)
            Sql = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtaux(12).Text, "T")
            txtAux2(12).Text = ""
            If Sql = "" Then
                MsgBox "Concepto NO encontrado: " & txtaux(12).Text, vbExclamation
                txtaux(12).Text = ""
            Else
                txtAux2(12).Text = Sql
            End If
            
            cmdAceptar.SetFocus
        End Select

        If Index = 5 Or Index = 6 Or Index = 7 Then CalcularIVA


End Sub

Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtaux(5).Text <> "" Then
            hab = HayKHabilitarCentroCoste(txtaux(5).Text)
        Else
            txtaux(12).Text = ""
        End If
        If hab Then
            txtaux(12).BackColor = &H80000005
            Else
            txtaux(12).BackColor = &H80000018
            txtaux(12).Text = ""
        End If
    End If
    txtaux(12).Enabled = hab
End Sub
'
''1.-Debe    2.-Haber   3.-Decide en asiento
'Private Sub HabilitarImportes(tipoConcepto As Byte)
'    Dim bDebe As Boolean
'    Dim bHaber As Boolean
'
'    'Vamos a hacer .LOCKED =
'    Select Case tipoConcepto
'    Case 1
'        bDebe = False
'        bHaber = True
'    Case 2
'        bDebe = True
'        bHaber = False
'    Case 3
'        bDebe = False
'        bHaber = False
'    Case Else
'        bDebe = True
'        bHaber = True
'    End Select
'
'    txtAux(9).Enabled = Not bDebe
'    txtAux(10).Enabled = Not bHaber
'
'    If bDebe Then
'        txtAux(9).BackColor = &H80000018
'    Else
'        txtAux(9).BackColor = &H80000005
'    End If
'    If bHaber Then
'        txtAux(10).BackColor = &H80000018
'    Else
'        txtAux(10).BackColor = &H80000005
'    End If
'End Sub
'

Private Sub FijarDiarioPorDefecto()
    
    AntiguoText1 = "Select * from tiposdiario"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open AntiguoText1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        DiarioPorDefecto = miRsAux!NumDiari & "|" & miRsAux!desdiari & "|"
        miRsAux.MoveNext
        If Not miRsAux.EOF Then AntiguoText1 = ""
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    'Si hay mas de un diario, NO hago nada
    If AntiguoText1 = "" Then DiarioPorDefecto = ""
    AntiguoText1 = ""
        
End Sub


Private Sub HacerToolBar(Boton As Integer)

    'Si viene desde hco solo podemos MODIFCAR, ELIMINAR, LINEAS, ACTUALIZAR,SALIR
    If VieneDeDesactualizar Then
        i = Boton
        Sql = ""
        If i < 6 Then
            Sql = "NO"
        Else
            If i > 15 Then
                Sql = "NO"
            Else
                'INSERTAR, pero no estamos en edicion lineas
                If i = 6 And Modo <> 5 Then
                    Sql = "NO"
                End If
            End If
        End If
        If Sql <> "" Then
            MsgBox "Esta modificando el asiento de historico. Finalice primero este proceso.", vbExclamation
            Exit Sub
        End If
    End If
    
    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
            BotonModificar
        Case 3
            BotonEliminar False
        Case 5
            BotonBuscar
        Case 6
            BotonVerTodos
        Case 8
            'Imprimir factura
            
            


    End Select
End Sub


Private Function ModificarFactura() As Boolean
Dim B1 As Boolean
Dim vC As Contadores

    On Error GoTo EModificar
         
        ModificarFactura = False
     
                    
        Conn.BeginTrans
        'Comun
        
        ActualizarRetencionLineasSiNecesario
        
        
        
        
        B = RecalcularTotalesFactura
        
        If B Then B = ModificaDesdeFormulario2(Me, 1)
        
        If B Then B = ModificaLineas
  
EModificar:
        If Err.Number <> 0 Or Not B Then
            MuestraError Err.Number
            Conn.RollbackTrans
            ModificarFactura = False
            B1 = False
        Else
            Conn.CommitTrans
            ModificarFactura = True
        End If
        
End Function

Private Function ModificaLineas() As Boolean
Dim Sql As String

    On Error GoTo eRecalcularTotalesFactura

    ModificaLineas = False

    
    Sql = "update factpro_lineas set fecharec = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numregis= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    Conn.Execute Sql
    
    Sql = "update factpro_totales set fecharec = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numregis= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    Conn.Execute Sql
    
    ModificaLineas = True
    Exit Function
    
eRecalcularTotalesFactura:
    MuestraError Err.Number, "Modifica Lineas Factura", Err.Description
End Function



'##### Nuevo para el ambito de fechas
Private Function AmbitoDeFecha(DesbloqueAsiento As Boolean) As Boolean
        AmbitoDeFecha = False
        varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
        If varFecOk > 1 Then
            If varFecOk = 2 Then
                MsgBox varTxtFec, vbExclamation
            Else
                MsgBox "El asiento pertenece a un ejercicio cerrado.", vbExclamation
            End If
        Else
            AmbitoDeFecha = True
        End If
    
'        If DesbloqueAsiento Then DesBloqAsien
End Function


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
        Case 0
            txtaux(0).Text = ""
            miI = 3
        Case 3
            txtaux(3).Text = ""
            miI = 0
        Case 4
            txtaux(4).Text = ""
            miI = 1
            
        Case 8
            txtaux(8).Text = ""
            miI = 2
        End Select
        If miI >= 0 Then cmdAux_Click miI
End Sub



Private Function AuxOK() As String
    
    'Cuenta
    If txtaux(4).Text = "" Then
        AuxOK = "Cuenta no puede estar vacia."
        Exit Function
    End If
    
    If Not IsNumeric(txtaux(4).Text) Then
        AuxOK = "Cuenta debe ser numérica"
        Exit Function
    End If
    
    If txtAux2(4).Text = NO Then
        AuxOK = "La cuenta debe estar dada de alta en el sistema"
        Exit Function
    End If
    
    If Not EsCuentaUltimoNivel(txtaux(4).Text) Then
        AuxOK = "La cuenta no es de último nivel"
        Exit Function
    End If
    
        
    'Codigo de iva
    If txtaux(4).Text = "" Then
        AuxOK = "El código de iva no puede estar vacio"
        Exit Function
    End If
        
    If txtaux(7).Text <> "" Then
        If Not IsNumeric(txtaux(7).Text) Then
            AuxOK = "El código de iva debe de ser numérico."
            Exit Function
        End If
    End If
    
    'Importe
    If txtaux(6).Text <> "" Then
        If Not EsNumerico(txtaux(6).Text) Then
            AuxOK = "El importe DEBE debe ser numérico"
            Exit Function
        End If
    End If
    
    
    'cENTRO DE COSTE
    If txtaux(12).Enabled Then
        If txtaux(12).Text = "" Then
            AuxOK = "Centro de coste no puede ser nulo"
            Exit Function
        End If
    End If
    
                                            'Fecha del asiento
    If EstaLaCuentaBloqueada(txtaux(5).Text, CDate(Text1(1).Text)) Then
        AuxOK = "Cuenta bloqueada: " & txtaux(5).Text
        Exit Function
    End If
    
    AuxOK = ""
End Function



Private Function ComprobarNumeroFactura(Actual As Boolean) As Boolean
Dim cad As String
Dim RT As ADODB.Recordset
        cad = " WHERE numregis=" & Text1(0).Text
        cad = cad & " and numserie = " & DBSet(Text1(2).Text, "T")
        
        If Actual Then
            i = 0
        Else
            i = 1
        End If
        
        cad = cad & " AND anofactu =" & DBSet(Text1(14).Text, "N")
        
        Set RT = New ADODB.Recordset
        ComprobarNumeroFactura = True
        i = 0
        RT.Open "Select numregis from factpro" & cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.EOF) Then
                ComprobarNumeroFactura = False
            End If
        End If
        RT.Close
        If ComprobarNumeroFactura Then
            i = 1
            RT.Open "Select numregis from factpro" & cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RT.EOF Then
                If Not IsNull(RT.EOF) Then
                    ComprobarNumeroFactura = False
                End If
            End If
            RT.Close
        End If
        Set RT = Nothing
        If Not ComprobarNumeroFactura Then
            cad = "Verifique los contadores. Ya existe la factura " & Text1(0).Text & vbCrLf
            MsgBox cad, vbExclamation
        End If
End Function

Private Function SituarData1(Insertar As Boolean) As Boolean
    Dim Sql As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        Sql = "Select * from factpro WHERE numserie =" & DBSet(Text1(2).Text, "T")
        Sql = Sql & " AND fecharec=" & DBSet(Text1(1).Text, "F") & " AND numregis = " & Text1(0).Text
        Data1.RecordSource = Sql
    'End If
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!NUmSerie) = Text1(2).Text Then
                If CStr(.Fields!NumRegis) = Text1(0).Text Then
                    If Format(CStr(.Fields!fecharec), "dd/mm/yyyy") = Text1(1).Text Then
                        SituarData1 = True
                        Exit Function
                    End If
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
End Function


'********************************************************
'
' FUNCIONES CORRESPONDIENTES A LA INSERCION DE DOCUMENTOS
'
'********************************************************


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo "
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim cad As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String
Dim C As String


    On Error GoTo ECargaDatosLW
    
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 5 ' imagenes
        cad = "select h.numlinea,  h.codigiva, tt.nombriva,  h.baseimpo, h.impoiva, h.imporec from factpro_totales h inner join tiposiva tt on h.codigiva = tt.codigiva  WHERE "
        cad = cad & " numserie=" & DBSet(Data1.Recordset!NUmSerie, "T")
        cad = cad & " and numregis=" & DBSet(Data1.Recordset!NumRegis, "N")
        cad = cad & " and fecharec=" & DBSet(Data1.Recordset!fecharec, "F")
        cad = cad & " and anofactu=" & Data1.Recordset!anofactu
        GroupBy = ""
        BuscaChekc = "numlinea"
        
    End Select
    
    
    cad = cad & " ORDER BY 1"
    
    lw1.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not Rs.EOF
        Set IT = lw1.ListItems.Add

        IT.Text = Rs!NumLinea
        IT.SubItems(1) = Format(Rs!codigiva, "000")
        IT.SubItems(2) = Rs!nombriva
        IT.SubItems(3) = Format(Rs!Baseimpo, "###,###,##0.00")
        IT.SubItems(4) = Format(Rs!Impoiva, "###,###,##0.00")
        If DBLet(Rs!ImpoRec) <> 0 Then
            IT.SubItems(5) = Format(Rs!ImpoRec, "###,###,##0.00")
        Else
            IT.SubItems(5) = " "
        End If
        
        Set IT = Nothing

        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
End Sub


Private Sub AnyadirAlListview(vpaz As String, DesdeBD As Boolean)
Dim J As Integer
Dim Aux As String
Dim IT As ListItem
Dim Contador As Integer
    If Dir(vpaz, vbArchive) = "" Then
'        MsgBox "No existe el archivo: " & vpaz, vbExclamation
    Else
        Set IT = lw1.ListItems.Add()

        IT.Text = Me.Adodc1.Recordset!Orden '"Nuevo " & Contador
        
        IT.SubItems(1) = Me.Adodc1.Recordset.Fields(5)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        IT.SubItems(2) = vpaz
        IT.SubItems(3) = Me.Adodc1.Recordset.Fields(0)
        
        Set IT = Nothing
    End If
End Sub




Private Sub CargaFiltros()
Dim Aux As String
    

    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Ejercicios Abiertos "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Ejercicio Actual "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2
    cboFiltro.AddItem "Ejercicio Siguiente "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 3

End Sub
    

Private Function SePuedeModificarAsiento(MostrarMensaje As Boolean) As Boolean
Dim CadFac As String

        CadFac = ""
        
        SePuedeModificarAsiento = False
      
        If Me.AdoAux(1).Recordset.EOF Then Exit Function
        
        'Primero comprobamos si esta cerrado el ejercicio
        varFecOk = FechaCorrecta2(AdoAux(1).Recordset!FechaEnt)
        If varFecOk >= 2 Then
            If varFecOk = 2 Then
                If MostrarMensaje Then MsgBox varTxtFec, vbExclamation
            Else
                If MostrarMensaje Then MsgBox "El asiento pertenece a un ejercicio cerrado.", vbExclamation
            End If
            Exit Function
        End If
        
        'Cojo prestado esta variabel un momento CadenaDesdeOtroForm
        If Not IsNull(AdoAux(1).Recordset!idcontab) Then
            If AdoAux(1).Recordset!idcontab = "FRACLI" Then
                CadFac = "FRACLI"
                CadenaDesdeOtroForm = " clientes "
            Else
                If AdoAux(1).Recordset!idcontab = "FRAPRO" Then
                    CadFac = "FRAPRO"
                    CadenaDesdeOtroForm = " proveedores "
                End If
            End If
        End If
        If CadFac <> "" Then
                If MostrarMensaje Then MsgBox "Este apunte pertenece a una factura de " & CadenaDesdeOtroForm & " y solo se puede modificar en el registro" & _
                    " de facturas de " & CadenaDesdeOtroForm & ".", vbExclamation
                i = -1
            Exit Function
        Else
            SePuedeModificarAsiento = True
        End If


End Function

Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String

    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i

    'Tipo de factura
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wconce340 ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        Combo1(0).AddItem Rs!Descripcion
        Combo1(0).ItemData(Combo1(0).NewIndex) = Asc(Rs!Codigo)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Tipo de operacion
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtipopera ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Combo1(1).AddItem Rs!denominacion
        Combo1(1).ItemData(Combo1(1).NewIndex) = Rs!Codigo
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Tipo de retencion
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtiporeten ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Combo1(2).AddItem Rs!Descripcion
        Combo1(2).ItemData(Combo1(2).NewIndex) = Rs!Codigo
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Tipo de intracomunitaria
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtipointra ORDER BY codintra"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        Combo1(3).AddItem Rs!nomintra
        Combo1(3).ItemData(Combo1(3).NewIndex) = Asc(Rs!Codintra)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

End Sub

Private Function ComprobarPeriodo2(Indice As Integer) As Boolean
Dim Cerrado As Boolean

    ModificarPagos = True


    'Primero pondremos la fecha a año periodo
    i = Year(CDate(Text1(Indice).Text))
    If vParam.periodos = 0 Then
        'Trimestral
        Ancho = ((Month(CDate(Text1(Indice).Text)) - 1) \ 3) + 1
        Else
        Ancho = Month(CDate((Text1(Indice).Text)))
    End If
    Cerrado = False
    If i < vParam.anofactu Then
        Cerrado = True
    Else
        If i = vParam.anofactu Then
            'El mismo año. Comprobamos los periodos
            If vParam.perfactu >= Ancho Then _
                Cerrado = True
        End If
    End If
    ComprobarPeriodo2 = True
    ModificaFacturaPeriodoLiquidado = False
    If Cerrado Then
        ModificaFacturaPeriodoLiquidado = True
        Sql = "La fecha "
        If Indice = 0 Then
            Sql = Sql & "factura"
        Else
            Sql = Sql & "liquidacion"
        End If
        Sql = Sql & " corresponde a un periodo ya liquidado. " & vbCrLf
        
        If vUsu.Nivel = 0 Then
            Sql = Sql & vbCrLf & " ¿Desea continuar igualmente ?"
  
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then ComprobarPeriodo2 = False
        Else
        
            MsgBox Sql, vbExclamation
            
            ComprobarPeriodo2 = False
        
        End If
    
        ModificarPagos = False
    
    End If
    
End Function


Private Sub CargarDatosCuenta(Cuenta As String)
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo eTraerDatosCuenta
    
    Sql = "select * from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(5).Text = ""
    Text4(5).Text = ""
    
    For i = 15 To 21
        Text1(i).Text = ""
    Next i
    
    If Not Rs.EOF Then
        Text1(5).Text = DBLet(Rs!Forpa, "N")
        Text4(5).Text = PonerNombreDeCod(Text1(5), "formapago", "nomforpa", "codforpa", "N")
        
        Text1(15).Text = DBLet(Rs!Nommacta, "T")
        Text1(16).Text = DBLet(Rs!dirdatos, "T")
        Text1(17).Text = DBLet(Rs!codposta, "T")
        Text1(18).Text = DBLet(Rs!desPobla, "T")
        Text1(19).Text = DBLet(Rs!desProvi, "T")
        Text1(20).Text = DBLet(Rs!nifdatos, "T")
        Text1(21).Text = DBLet(Rs!codPAIS, "T")
        Text4(21).Text = PonerNombreDeCod(Text1(21), "paises", "nompais", "codpais", "T")
    End If
    Exit Sub
    
eTraerDatosCuenta:
    MuestraError Err.Number, "Cargar Datos de Cuenta", Err.Description

End Sub


Private Function AnyadeCadenaFiltro() As String
Dim Aux As String

    Aux = ""
    If vUsu.FiltroFactCli <> 0 Then
        '-------------------------------- INICIO
        i = Year(vParam.fechaini)
        If vUsu.FiltroFactCli < 3 Then
            'INicio = actual
            Aux = " anofactu >= " & i
            Else
            Aux = " anofactu >=" & i + 1
        End If
        i = Year(vParam.fechafin)
        If vUsu.FiltroFactCli = 2 Then
            Aux = Aux & " AND anofactu <= " & i
        Else
            Aux = Aux & " AND anofactu <= " & i + 1
        End If
        
    End If  'filtro=0
    AnyadeCadenaFiltro = Aux
End Function

Private Sub CalcularIVA()
Dim J As Integer
Dim Base As Currency
Dim Aux As Currency

    Base = ImporteFormateado(txtaux(6).Text)
    
    'EL iva
    Aux = ImporteFormateado(txtaux(8).Text) / 100
    If Aux = 0 Then
        txtaux(10).Text = "0,00"
    Else
        txtaux(10).Text = Format(Round((Aux * Base), 2), FormatoImporte)
    End If
    
    'Recargo
    Aux = ImporteFormateado(txtaux(9).Text) / 100
    If Aux = 0 Then
        txtaux(11).Text = ""
    Else
        txtaux(11).Text = Format(Round((Aux * Base), 2), FormatoImporte)
    End If

End Sub

Private Function RecalcularTotales() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim SqlValues As String
Dim i As Long
Dim Rs As ADODB.Recordset

Dim Baseimpo As Currency
Dim Basereten As Currency
Dim Impoiva As Currency
Dim ImpoRec As Currency
Dim Imporeten As Currency
Dim TotalFactura As Currency

    On Error GoTo eRecalcularTotales

    RecalcularTotales = False

    Sql = "delete from factpro_totales where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    Conn.Execute Sql
    
    SqlInsert = "insert into factpro_totales (numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec) values "
    
    Sql = "select codigiva, porciva, porcrec, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec from factpro_lineas "
    Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    Sql = Sql & " group by 1,2,3"
    Sql = Sql & " order by 1,2,3"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    
    SqlValues = ""
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    ImpoRec = 0
    Imporeten = 0
    
    TotalFactura = 0
    
    While Not Rs.EOF
        Sql = "(" & DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & "," & DBSet(Text1(14).Text, "N") & ","
        Sql = Sql & DBSet(i, "N") & "," & DBSet(Rs!Baseimpo, "N") & "," & DBSet(Rs!codigiva, "N") & "," & DBSet(Rs!porciva, "N") & "," & DBSet(Rs!porcrec, "N") & ","
        Sql = Sql & DBSet(Rs!Imporiva, "N") & "," & DBSet(Rs!imporrec, "N") & "),"
        
        SqlValues = SqlValues & Sql
        
        Baseimpo = Baseimpo + DBLet(Rs!Baseimpo, "N")
        Impoiva = Impoiva + DBLet(Rs!Imporiva, "N")
        ImpoRec = ImpoRec + DBLet(Rs!imporrec, "N")
        
        i = i + 1
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert & SqlValues
    End If
    
    
    RecalcularTotales = RecalcularTotalesFactura
    Exit Function
    
eRecalcularTotales:
    MuestraError Err.Number, "Recalcular Totales", Err.Description
End Function


Private Function RecalcularTotalesFactura() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim SqlValues As String
Dim i As Long
Dim Rs As ADODB.Recordset

Dim Baseimpo As Currency
Dim Basereten As Currency
Dim Impoiva As Currency
Dim ImpoRec As Currency
Dim Imporeten As Currency
Dim TotalFactura As Currency
Dim PorcRet As Currency

Dim TipoRetencion As Integer

    On Error GoTo eRecalcularTotalesFactura

    RecalcularTotalesFactura = False

    TipoRetencion = DevuelveValor("select tipo from usuarios.wtiporeten where codigo = " & DBSet(Combo1(2).ListIndex, "N"))
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    Imporeten = 0
    ImpoRec = 0
    TotalFactura = 0
    
    Sql = "select aplicret, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec from factpro_lineas "
    Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    Sql = Sql & " group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Baseimpo = Baseimpo + DBLet(Rs!Baseimpo, "N")
        
        ' en el caso de inversion sujeto pasivo o intracomunitarias no se suma a totales los ivas y recargos
        If Combo1(1).ListIndex = 1 Or Combo1(1).ListIndex = 4 Then
        
        Else
            Impoiva = Impoiva + DBLet(Rs!Imporiva, "N")
            ImpoRec = ImpoRec + DBLet(Rs!imporrec, "N")
        End If
        
        If Rs!aplicret = 1 Then
            Basereten = Basereten + DBLet(Rs!Baseimpo, "N")
            
            If TipoRetencion = 1 Then
                If Combo1(1).ListIndex = 1 Or Combo1(1).ListIndex = 4 Then
                
                Else
                    Basereten = Basereten + DBLet(Rs!Imporiva, "N")
                End If
            End If
        End If
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    PorcRet = ImporteFormateado(Text1(7).Text)
    
    If PorcRet = 0 Then Basereten = 0
   
    
    If PorcRet = 0 Then
        Imporeten = 0
    Else
        Imporeten = Round((PorcRet * Basereten / 100), 2)
    End If
    
    TotalFactura = Baseimpo + Impoiva + ImpoRec - Imporeten
    
    Text1(9).Text = Format(Baseimpo, FormatoImporte)
    Text1(11).Text = Format(Basereten, FormatoImporte)
    Text1(10).Text = Format(Impoiva, FormatoImporte)
    Text1(12).Text = Format(Imporeten, FormatoImporte)
    Text1(13).Text = Format(TotalFactura, FormatoImporte)
    
    If PorcRet = 0 Then
        Text1(11).Text = ""
        Text1(12).Text = ""
    End If
    
    Sql = "update factpro set "
    Sql = Sql & " totbases = " & DBSet(Baseimpo, "N")
    Sql = Sql & ", totivas = " & DBSet(Impoiva, "N")
    Sql = Sql & ", totrecargo = " & DBSet(ImpoRec, "N")
    Sql = Sql & ", totfacpr = " & DBSet(TotalFactura, "N")
    Sql = Sql & ", totbasesret = " & DBSet(Basereten, "N", "S")
    Sql = Sql & ", trefacpr = " & DBSet(Imporeten, "N", "S")
    Sql = Sql & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numregis= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    Conn.Execute Sql
    
    
    RecalcularTotalesFactura = True
    Exit Function
    
eRecalcularTotalesFactura:
    MuestraError Err.Number, "Recalcular Totales Factura", Err.Description
End Function


Private Function IntegrarFactura() As Boolean

    IntegrarFactura = False
    
    With frmActualizar
        .OpcionActualizar = 8
        'NumAsiento     --> CODIGO FACTURA
        'NumDiari       --> AÑO FACTURA
        'NUmSerie       --> SERIE DE LA FACTURA
        'FechaAsiento   --> Fecha factura
        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
        .NumFac = CLng(Text1(0).Text)
        .NumDiari = CInt(Text1(14).Text)
        .NUmSerie = Text1(2).Text
        .FechaAsiento = Text1(1).Text
        .FechaAnterior = FecRecepAnt 'FecFactuAnt
        
        If Numasien2 < 0 Then
            
            If Not Text1(8).Enabled Then
                If Text1(8).Text <> "" Then
                    Numasien2 = Text1(8).Text
                End If
            End If
            
        End If
        If NumDiario <= 0 Then NumDiario = vParam.numdiapr
        .DiarioFacturas = NumDiario
        .NumAsiento = Numasien2
        .Show vbModal
        If AlgunAsientoActualizado Then IntegrarFactura = True
        Screen.MousePointer = vbHourglass
        Me.Refresh
    End With
    

End Function


Private Function Desintegrar() As Boolean
        Desintegrar = False
        'Primero hay que desvincular la factura de la tabla de hco
        If DesvincularFactura Then
            frmActualizar.OpcionActualizar = 2  'Desactualizar para eliminar
            frmActualizar.NumAsiento = Data1.Recordset!NumAsien
            frmActualizar.FechaAsiento = FecRecepAnt 'FecFactuAnt
            frmActualizar.NumDiari = Data1.Recordset!NumDiari
            frmActualizar.FechaAnterior = Data1.Recordset!FechaEnt
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            If AlgunAsientoActualizado Then Desintegrar = True
        End If
End Function


Private Function DesvincularFactura() As Boolean
On Error Resume Next
    Sql = "UPDATE factpro set numasien=NULL, fechaent=NULL, numdiari=NULL"
    Sql = Sql & " WHERE numregis = " & Data1.Recordset!NumRegis
    Sql = Sql & " AND numserie = '" & Data1.Recordset!NUmSerie & "'"
    Sql = Sql & " AND anofactu =" & Data1.Recordset!anofactu
    Numasien2 = Data1.Recordset!NumAsien
    NumDiario = Data1.Recordset!NumDiari
    Conn.Execute Sql
    If Err.Number <> 0 Then
        DesvincularFactura = False
        MuestraError Err.Number, "Desvincular factura"
    Else
        DesvincularFactura = True
    End If
End Function


Private Function TieneRegistros() As Boolean
    On Error Resume Next
    TieneRegistros = False
    If Data1.Recordset.RecordCount > 0 Then TieneRegistros = True
End Function



Private Function ContabilizarPagos() As Boolean
Dim Mc As Contadores
Dim FP As Ctipoformapago
Dim Sql As String
Dim Ampliacion As String
Dim Numdocum As String
Dim Conce As Integer
Dim LlevaContr As Boolean
Dim Im As Currency
Dim Debe As Boolean
Dim ElConcepto As Integer
Dim Linea As Integer
Dim TotImpo As Currency
Dim Sql1 As String
Dim Rs As ADODB.Recordset
Dim impo As Currency
Dim cad As String
Dim Sql4 As String
    
    On Error GoTo ECon
    
    ContabilizarPagos = False
    
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", CDate(FechaPago) <= vParam.fechafin, True) = 1 Then Exit Function

    Set FP = New Ctipoformapago
    
    Linea = DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(Text1(5), "N"))
    
    If FP.Leer(Linea) Then
        Set Mc = Nothing
        Set FP = Nothing
    End If
    
    Sql1 = "select * "
    Sql = " from pagos where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
    Sql = Sql & " and codmacta = " & DBSet(Text1(4).Text, "T")
    Sql = Sql & " and fecfactu = " & DBSet(Text1(26).Text, "F")
    Sql = Sql & " order by numorden"
    
    TotImpo = DevuelveValor("select sum(coalesce(impefect,0)) " & Sql)
    
    Sql = Sql1 & Sql
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Inserto cabecera de apunte
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, feccreacion, usucreacion, desdeaplicacion, obsdiari) VALUES ("
    Sql = Sql & FP.diaricli
    Sql = Sql & ",'" & Format(FechaPago, FormatoFecha) & "'," & Mc.Contador & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilización Pago Facturas Proveedor',"
    Sql1 = DBSet("Generado desde Facturas de Proveedor el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre, "T")
    If TotImpo < 0 Then Sql1 = Sql1 & "  (CARGO)"
    Conn.Execute Sql & Sql1 & ")"
    
    Linea = 0
    While Not Rs.EOF
        
        Linea = Linea + 1
        
        'importe
        impo = ImporteFormateado(DBLet(Rs!imppagad))
        
        'Inserto en las lineas de apuntes
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
        Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada,numserie,numfacpr,fecfactu,numorden,tipforpa ) VALUES ("
        Sql = Sql & FP.diaricli
        Sql = Sql & ",'" & Format(FechaPago, FormatoFecha) & "'," & Mc.Contador & ","
        
        
        'numdocum
        Numdocum = Text1(2).Text & "-" & Text1(25).Text  ' letra de serie y factura
        
        'Concepto y ampliacion del apunte
        Ampliacion = ""
        'Proveedores
        Debe = True
        If impo < 0 Then
            If Not vParam.abononeg Then Debe = False
        End If
        If Debe Then
            Conce = FP.ampdepro
            LlevaContr = FP.ctrdepro = 1
            ElConcepto = FP.condepro
        Else
            ElConcepto = FP.conhapro
            Conce = FP.amphapro
            LlevaContr = FP.ctrhapro = 1
        End If
               
        'Si el importe es negativo y no permite abonos negativos
        'como ya lo ha cambiado de lado (dbe <-> haber)
        If impo < 0 Then
            If Not vParam.abononeg Then impo = Abs(impo)
        End If
           
        If Conce = 2 Then
            Ampliacion = Ampliacion & DBLet(Rs!fecefect)  'Fecha vto
        ElseIf Conce = 4 Then
            'Contra partida
            Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaBanco, "T")
        Else
            
           If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
           Ampliacion = Ampliacion & Text1(2).Text & "/" & Text1(25).Text 'RecuperaValor(Vto, 1) & "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
        End If
        
        'Fijo en concepto el codconce
        Conce = ElConcepto
        cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
        Ampliacion = cad & " " & Ampliacion
        Ampliacion = Mid(Ampliacion, 1, 35)
        
        'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
        'Cuenta Cliente/proveedor
        cad = Linea & ",'" & Trim(Text1(4).Text) & "','" & Numdocum & "'," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
        'Importe cobro-pago
        ' nos lo dire "debe"
        If Not Debe Then
            cad = cad & "NULL," & TransformaComasPuntos(CStr(impo))
        Else
            cad = cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
        End If
        'Codccost
        cad = cad & ",NULL,"
        If LlevaContr Then
            cad = cad & "'" & CtaBanco & "'"
        Else
            cad = cad & "NULL"
        End If
        cad = cad & ",'PAGOS',0," & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & "," & TipForpa & ")"
        cad = Sql & cad
        Conn.Execute cad
        
        
        Rs.MoveNext
        
    Wend
    
    'El banco    *******************************************************************************
    '---------------------------------------------------------------------------------------------
    
    Linea = Linea + 1
    
    'Vuelvo a fijar los valores
     'Concepto y ampliacion del apunte
    Ampliacion = ""
    'CLIENTES
     'Si el apunte va al debe, el contrapunte va al haber
     If Not Debe Then
         Conce = FP.ampdepro
         LlevaContr = FP.ctrdepro = 1
         ElConcepto = FP.condepro
     Else
         ElConcepto = FP.conhapro
         Conce = FP.amphacli
         LlevaContr = FP.ctrhapro = 1
     End If
           
           
    If Conce = 2 Then
       Ampliacion = Ampliacion & DBLet(Rs!fecefect)  'Fecha efecto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(Text1(2).Text)
    Else
        
       If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
        Ampliacion = Ampliacion & Text1(2).Text & "/" & Text1(25).Text
    End If
    
    
    Conce = ElConcepto
    cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    cad = Linea & "," & DBSet(CtaBanco, "T") & ",'" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
    'Importe cliente
    'Si el cobro/pago va al debe el contrapunte ira al haber
    If Not Debe Then
        'al debe
        cad = cad & TransformaComasPuntos(CStr(TotImpo)) & ",NULL"
    Else
        'al haber
        cad = cad & "NULL," & TransformaComasPuntos(CStr(TotImpo))
    End If
    
    'Codccost
    cad = cad & ",NULL,"
    
    If LlevaContr Then
        cad = cad & "'" & Trim(Text1(4).Text) & "'"
    Else
        cad = cad & "NULL"
    End If
    cad = cad & ",'PAGOS',0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
    cad = Sql & cad
    Conn.Execute cad
    
    ContabilizarPagos = True

    Set Mc = Nothing
    Set FP = Nothing

    Exit Function
ECon:
    MuestraError Err.Number, "Contabilizar Pagos"
    Set Mc = Nothing
    Set FP = Nothing
End Function


Private Function EsFraProTraspasada() As Boolean
Dim Sql As String

    Sql = "select estraspasada from factpro where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = "
    Sql = Sql & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    EsFraProTraspasada = (DevuelveValor(Sql) = 1)
    

End Function




Private Sub CrearElementoInmovilizado_(CTA_Inmovilizado As String)
  Dim CrearEltoInmov As Boolean
  
        CrearEltoInmov = False
        If vParam.NuevoPlanContable Then
            If Mid(CTA_Inmovilizado, 1, 2) = "20" Or Mid(CTA_Inmovilizado, 1, 2) = "21" Then CrearEltoInmov = True
        Else
            If Mid(CTA_Inmovilizado, 1, 2) = "21" Or Mid(CTA_Inmovilizado, 1, 2) = "22" Then CrearEltoInmov = True
        End If
        If Not CrearEltoInmov Then Exit Sub
  
  
        Sql = DevuelveDesdeBD("codigo", "paramamort", "1", "1")
        If Trim(Sql) = "" Then Exit Sub
        
    
        
        
        If CrearEltoInmov Then
            Sql = "Desea crear un elemento de Inmovilizado ? "
            If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
                'Le pasaremos el codprove|nomprove|numfac|importe
                'ANTES
                'codprove    nombre    numfac     fecha adq     importe     Cuenta    Des. cuenta
                CadenaDesdeOtroForm = Text1(4).Text & "|" & Me.Text4(4).Text & "|"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(25).Text & "|" & Text1(1).Text & "||"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtaux(5).Text & "|" & txtAux2(5).Text & "|"
                
                frmInmoElto.Nuevo = CadenaDesdeOtroForm
                CadenaDesdeOtroForm = ""
                frmInmoElto.Show vbModal
    
                Screen.MousePointer = vbDefault
            End If
        End If
    
End Sub




Private Sub ActualizarRetencionLineasSiNecesario()
Dim Aux As String

    
    'Si antes no tenia y ahora si
    'Es dentro de una transaccion. Con lo cual updateamos sin problems
    If Combo1(2).ListIndex > 0 And Val(DBLet(Me.Data1.Recordset!tiporeten, "T")) = 0 Then
        'Si solo tiene una linea, la actualizo yo
        Aux = Replace(ObtenerWhereCab(False), "factpro", "factpro_lineas")
        Aux = DevuelveDesdeBD("count(*)", "factpro_lineas ", Aux & " AND aplicret", "0")
        If Val(Aux) = 1 Then
            'updateamos
            Aux = Replace(ObtenerWhereCab(True), "factpro", "factpro_lineas")
            Aux = "UPDATE factpro_lineas set aplicret=1 " & Aux
            Ejecuta Aux
        Else
            If Val(Aux) > 1 Then
                MsgBox "Indique las lineas que llevan retencion", vbInformation
            End If
        End If
    End If

End Sub
