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
      Height          =   2145
      Left            =   390
      TabIndex        =   90
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
         TabIndex        =   105
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
         TabIndex        =   97
         Tag             =   "Pa�s|T|S|||factpro|codpais|||"
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
         TabIndex        =   92
         Tag             =   "Nif|T|S|||factpro|nifdatos|||"
         Top             =   390
         Width           =   2070
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
         Tag             =   "Direcci�n|T|S|||factpro|dirdatos|||"
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
         TabIndex        =   91
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
         Caption         =   "Pa�s"
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
         TabIndex        =   104
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
         TabIndex        =   103
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Poblaci�n"
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Direcci�n"
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
         TabIndex        =   99
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
         TabIndex        =   98
         Top             =   450
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2385
      Left            =   9720
      TabIndex        =   79
      Top             =   4800
      Width           =   7725
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
         Index           =   31
         Left            =   5880
         TabIndex        =   120
         Tag             =   "Importe Iva|N|S|||factpro|suplidos|###,###,##0.00||"
         Text            =   "9.999.999,99"
         Top             =   1320
         Width           =   1575
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
         Index           =   30
         Left            =   2040
         TabIndex        =   118
         Tag             =   "Importe Iva|N|S|||factpro|totrecargo|###,###,##0.00||"
         Text            =   "9.999.999,99"
         Top             =   1320
         Width           =   1575
      End
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
         Left            =   5880
         TabIndex        =   23
         Tag             =   "Total Factura|N|S|||factpro|totfacpr|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1800
         Width           =   1575
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
         Left            =   5880
         TabIndex        =   22
         Tag             =   "Importe Retenci�n|N|S|||factpro|trefacpr|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   840
         Width           =   1575
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
         Left            =   5880
         TabIndex        =   21
         Tag             =   "Base Retenci�n|N|S|||factpro|totbasesret|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   360
         Width           =   1575
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
         Left            =   2040
         TabIndex        =   20
         Tag             =   "Importe Iva|N|S|||factpro|totivas|###,###,##0.00||"
         Text            =   "9.999.999,99"
         Top             =   840
         Width           =   1575
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
         Left            =   2040
         TabIndex        =   19
         Tag             =   "Base Imponible|N|S|||factpro|totbases|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Suplidos"
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
         Index           =   20
         Left            =   3960
         TabIndex        =   121
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Rec. Eq."
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
         Index           =   19
         Left            =   120
         TabIndex        =   119
         Top             =   1320
         Width           =   1875
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
         Left            =   3960
         TabIndex        =   85
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retenci�n"
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
         Left            =   3960
         TabIndex        =   84
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Base Retenci�n"
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
         Left            =   3960
         TabIndex        =   83
         Top             =   360
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
         Left            =   120
         TabIndex        =   82
         Top             =   840
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
         Left            =   120
         TabIndex        =   81
         Top             =   360
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
         Left            =   120
         TabIndex        =   80
         Top             =   0
         Width           =   1980
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3330
      TabIndex        =   66
      Top             =   90
      Width           =   3285
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   67
         Top             =   180
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Datos Fiscales"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pagos"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Errores N�Registro"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas sin Asiento"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Entrada factura"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   10170
      TabIndex        =   64
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
      Height          =   2385
      Left            =   270
      TabIndex        =   62
      Top             =   4800
      Width           =   9375
      Begin VB.CommandButton cmdAux3 
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
         Left            =   2520
         TabIndex        =   129
         ToolTipText     =   "Buscar cuenta"
         Top             =   1440
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame FrameModifIVA 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   120
         TabIndex        =   127
         Top             =   240
         Visible         =   0   'False
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAuxTot 
            Height          =   330
            Left            =   0
            TabIndex        =   128
            Top             =   0
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
      Begin VB.TextBox txtaux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3600
         TabIndex        =   126
         Tag             =   "Base Imponible|N|S|||factcli_totales|baseimpo|###,###,##0.00||"
         Text            =   "Base Imponible"
         Top             =   1680
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
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
         Left            =   1080
         TabIndex        =   125
         Tag             =   "Iva|N|S|||factcli_totales|codigiva|000||"
         Text            =   "Iva"
         Top             =   1680
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6540
         TabIndex        =   124
         Tag             =   "Importe Iva|N|S|||factcli_totales|impoiva|###,###,##0.00||"
         Text            =   "ImpIva"
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7440
         TabIndex        =   123
         Tag             =   "Importe Rec|N|S|||factcli_totales|imporec|###,###,##0.00||"
         Text            =   "ImpRec"
         Top             =   1710
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtaux3 
         Appearance      =   0  'Flat
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
         Left            =   2160
         TabIndex        =   122
         Tag             =   "Iva|N|S|||factcli_totales|codigiva|000||"
         Text            =   "Iva"
         Top             =   1680
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
         Height          =   1665
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   2937
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
         TabIndex        =   65
         Top             =   0
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
      TabIndex        =   55
      Top             =   270
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   6720
      TabIndex        =   53
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   54
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
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   270
      TabIndex        =   50
      Top             =   90
      Width           =   3015
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   52
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
      Height          =   3930
      Index           =   0
      Left            =   270
      TabIndex        =   39
      Top             =   870
      Width           =   17160
      Begin VB.TextBox txtPDF 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "Text4"
         Top             =   540
         Width           =   4260
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
         Index           =   4
         ItemData        =   "frmFacturasPro.frx":0044
         Left            =   13200
         List            =   "frmFacturasPro.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Situacion inmueble|N|S|||factcli|CatastralSitu|||"
         Top             =   1950
         Visible         =   0   'False
         Width           =   3810
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
         Index           =   29
         Left            =   10560
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "RCatas|T|S|||factcli|CatastralREF|||"
         Top             =   1950
         Visible         =   0   'False
         Width           =   2580
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
         Index           =   28
         Left            =   12840
         TabIndex        =   112
         Text            =   "1234567890"
         Top             =   3270
         Width           =   1530
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
         Index           =   3
         ItemData        =   "frmFacturasPro.frx":0048
         Left            =   10590
         List            =   "frmFacturasPro.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   13
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
         Left            =   9810
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
         Tag             =   "N� factura|T|N|||factpro|numfactu|||"
         Top             =   540
         Width           =   1635
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
         TabIndex        =   18
         Tag             =   "N�mero Asiento|N|S|||factpro|numasien|00000000||"
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
         Left            =   11190
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   75
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
         ItemData        =   "frmFacturasPro.frx":004C
         Left            =   180
         List            =   "frmFacturasPro.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   15
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
         ItemData        =   "frmFacturasPro.frx":0050
         Left            =   7980
         List            =   "frmFacturasPro.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "Tipo operaci�n|N|N|||factpro|codopera|||"
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
         ItemData        =   "frmFacturasPro.frx":0054
         Left            =   7980
         List            =   "frmFacturasPro.frx":0056
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   57
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
         TabIndex        =   14
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
         Tag             =   "N� Registro|N|S|||factpro|numregis|0000000|S|"
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
         TabIndex        =   86
         Tag             =   "A�o factura|N|N|||factpro|anofactu||S|"
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
         TabIndex        =   88
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
         TabIndex        =   106
         Tag             =   "N�mero Diario|N|S|||factpro|numdiari|00000000||"
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
         TabIndex        =   111
         Tag             =   "Tipo intracomunitaria|T|S|||factpro|codintra|||"
         Top             =   2580
         Width           =   1245
      End
      Begin VB.Image imgpdf 
         Height          =   240
         Index           =   0
         Left            =   14880
         Picture         =   "frmFacturasPro.frx":0058
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgpdf 
         Height          =   240
         Index           =   1
         Left            =   15240
         Picture         =   "frmFacturasPro.frx":0A5A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Documento asociado"
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
         Left            =   12720
         TabIndex        =   117
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Situaci�n inmueble"
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
         Index           =   23
         Left            =   13200
         TabIndex        =   115
         Top             =   1680
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.Label Label1 
         Caption         =   "Referencia catastral"
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
         Index           =   22
         Left            =   10560
         TabIndex        =   114
         Top             =   1680
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Iden. SII"
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
         Index           =   21
         Left            =   12840
         TabIndex        =   113
         Top             =   3000
         Width           =   915
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
         TabIndex        =   110
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
         Left            =   9840
         TabIndex        =   108
         Top             =   240
         Width           =   1020
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   9
         Left            =   10920
         Picture         =   "frmFacturasPro.frx":145C
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "N�Factura"
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
         Left            =   8010
         TabIndex        =   107
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
         Left            =   12270
         Picture         =   "frmFacturasPro.frx":14E7
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
         Left            =   11190
         TabIndex        =   89
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
         TabIndex        =   78
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "% Retenci�n"
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
         TabIndex        =   77
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Retenci�n"
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
         TabIndex        =   76
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
         Caption         =   "Retenci�n"
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
         TabIndex        =   74
         Top             =   3000
         Width           =   1380
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Operaci�n"
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   68
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   7620
         Picture         =   "frmFacturasPro.frx":1572
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
         TabIndex        =   44
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label Label18 
         Caption         =   "Recepci�n"
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
         TabIndex        =   43
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   5220
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   285
      TabIndex        =   45
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
         TabIndex        =   109
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
         TabIndex        =   87
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
         TabIndex        =   37
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
         TabIndex        =   34
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
         TabIndex        =   28
         Tag             =   "A�o factura|N|N|||factpro_lineas|anofactu||S|"
         Text            =   "a�o"
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
         TabIndex        =   38
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
         TabIndex        =   33
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   61
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
         TabIndex        =   31
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
         TabIndex        =   60
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
         TabIndex        =   32
         Tag             =   "Codigo Iva|N|N|||factpro_lineas|codigiva|000||"
         Text            =   "Iva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   60
         TabIndex        =   58
         Top             =   120
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   180
            TabIndex        =   59
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   26
         Tag             =   "N� Serie|T|S|||factpro_lineas|numserie||S|"
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
         TabIndex        =   27
         Tag             =   "N� registro|N|N|0||factpro_lineas|numregis|0000000|S|"
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   48
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
      TabIndex        =   24
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
         TabIndex        =   25
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
      TabIndex        =   51
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
      TabIndex        =   49
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
      TabIndex        =   56
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
      Left            =   12720
      Top             =   0
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
      TabIndex        =   42
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

Private WithEvents frmPag_ As frmFacturasProPag ' pagos de tesoreria
Attribute frmPag_.VB_VarHelpID = -1
Private WithEvents frmUtil As frmUtilidades
Attribute frmUtil.VB_VarHelpID = -1
Private frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1


Dim AntiguoText1 As String
Private CadenaAmpliacion As String
Private SQL As String


Dim PosicionGrid As Integer

Dim Linliapu As Long
Dim FicheroAEliminar As String

Dim Numasien2 As Long
Dim NumDiario As Integer
Dim ContabilizaApunte As Boolean

Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Ll�nies

Dim NumTabMto As Integer 'Indica quin n� de Tab est� en modo Mantenimient
Dim TituloLinea As String 'Descripci� de la ll�nia que est� en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de ll�nies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de b�squeda posar el valor de poblaci� seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el n� del Bot� PrimerRegistro en la Toolbar1
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

Private Sub cboFiltro_KeyPress(KeyAscii As Integer)
    If Modo = 0 And KeyAscii = 27 Then Unload Me
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
    Dim Cad As String
    Dim i As Integer
    Dim Limp As Boolean
    Dim Mc As Contadores
    Dim B As Boolean
    Dim SqlLog As String
    Dim otro As Boolean

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
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
                        data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PosicionarData
                        PonerCampos
                        '[Monica]14/05/2015 a�ado numasien
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
                If data1.Recordset!NUmSerie = Text1(2).Text Then
                    If data1.Recordset!Numregis = CLng(Text1(0).Text) Then
                        If data1.Recordset!anofactu = Text1(14).Text Then
                            i = 0
                            'NO HA MODIFICADO NADA
                        End If
                    End If
                End If
            
                'Hacemos MODIFICAR
                Dim RC As Boolean
                If i <> 0 Then
                    MsgBoxA "No se puede cambiar campos clave  de la factura.", vbExclamation
                    RC = False
                Else
                    RC = ModificarFactura
                End If
                    
                If RC Then
                    '--DesBloqueaRegistroForm Me.Text1(0)
                    TerminaBloquear
                    
                    If Numasien2 > 0 Then
                        If IntegrarFactura_(False) Then
                            Text1(8).Text = Format(Numasien2, "0000000")
                            Numasien2 = -1
                            NumDiario = 0
                        Else
                            B = False
                        End If
                    End If
                    
                    
                    If Not ModificarPagos Then
                        If Text1(25).Text <> DBLet(data1.Recordset!NumFactu, "N") Then ModificarPagos = True
                        If Val(Text1(5).Text) <> DBLet(data1.Recordset!Codforpa, "N") Then ModificarPagos = True
                        If Me.Text1(4).Text <> DBLet(data1.Recordset!codmacta, "T") Then ModificarPagos = True
                        If Text1(1).Text <> DBLet(data1.Recordset!fecharec, "F") Then ModificarPagos = True
                        If Text1(26).Text <> DBLet(data1.Recordset!FecFactu, "F") Then ModificarPagos = True
                        
                    End If
                    
                    If ModificarPagos Then
                        PagosTesoreria
                    Else
                        If Me.Text1(15).Text <> DBLet(data1.Recordset!Nommacta, "T") Then
                            
                            Cad = " WHERE numserie = " & DBSet(Text1(2).Text, "T")
                            Cad = Cad & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
                            Cad = Cad & " and fecfactu = " & DBSet(Text1(26).Text, "F")
                            
                            Cad = "UPDATE pagos set nomprove = " & DBSet(Text1(15).Text, "T") & Cad
                            Ejecuta Cad, False
                        
                        End If
                    End If
                    
                    'LOG
                    SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
                    SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text

                    vLog.Insertar 9, vUsu, SqlLog
                    
                    
                    'Nuevo. Si ahora tiene retencion, y antes NO tenia
                    ActualizarRetencionLineasSiNecesario
                    
                    
                    
                    
                    PosicionarData
                    
                End If
            End If
        
        Case 5 'LL�NIES
            FecFactuAnt = Text1(26).Text
            FecRecepAnt = Text1(1).Text
            
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    InsertarLinea
                Case 2 'modificar ll�nies
                    If ModificarLinea2 Then
                                            
                        '**** parte de contabilizacion de la factura
                        TerminaBloquear
                        
                        If Numasien2 > 0 Then
                            If IntegrarFactura_(False) Then
                                Text1(8).Text = Format(Numasien2, "0000000")
                                Numasien2 = -1
                                NumDiario = 0
                            Else
                                B = False
                            End If
                        End If
                    
                        If ModificarPagos Then PagosTesoreria
                    
                        PosicionarData
                    End If
            End Select
            
    Case 6
            'Ha a�adido /Modificado   IVA
        If AnyadirModificarIVA Then
            RecalcularTotalesFactura True
            
            LLamaLineas 2, 0
            PonerModo 2
            PosicionarData
            PonerCampos
        End If
    
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBoxA Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PagosTesoreria()
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Actualizar As Boolean
Dim Aux As String


    On Error GoTo ePagosTesoreria

    If Not vEmpresa.TieneTesoreria Then Exit Sub
    
    
    
    
    ' si me cambian el nro de fra la cambio ya, SIEMPRE que no haayan pagos parciales
    
    Actualizar = False
    
    If NumFactuAnt <> "" Or NumFactuAnt <> "" Then
    If Trim(Text1(25).Text) <> Trim(NumFactuAnt) Then Actualizar = True
    If Text1(4).Text <> CodmactaAnt Then Actualizar = True
    End If
    If Actualizar Then
    
        SQL = "numserie = " & DBSet(Text1(2).Text, "T")
        SQL = SQL & " and codmacta = " & DBSet(CodmactaAnt, "T") & " and numfactu = " & DBSet(NumFactuAnt, "T")
        SQL = SQL & " and fecfactu = " & DBSet(FecFactuAnt, "F") & " AND 1"
        SQL = DevuelveDesdeBD("imppagad", "pagos", SQL, "1")
        If SQL <> "" Then
            'Tiene pagos parciales efectuados. Debera ir a tesoreria
            MsgBoxA "Tiene pagos parciales realizados. Revise tesorer�a", vbExclamation
            Exit Sub
        End If
        SQL = "update pagos set numfactu = " & DBSet(Text1(25).Text, "T")
        SQL = SQL & ", codmacta = " & DBSet(Text1(4).Text, "T")
        
        'Datos fiscales
        If Text1(4).Text <> CodmactaAnt Then
            Set Rs = New ADODB.Recordset
            
            Aux = "Select razosoci, dirdatos ,codposta ,desPobla, desProvi, nifdatos, codPAIS FROM cuentas where codmacta =" & DBSet(Text1(4).Text, "T")
            Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            'NO PUEDE SER EOG
            If Rs.EOF Then
                MsgBoxA "Cuenta proveedor incorrecta. ", vbCritical
            Else
                'nomprove  domprove   pobprove  cpprove nifprove codpais
                
                SQL = SQL & ", nomprove =" & DBSet(Rs!razosoci, "T")
                SQL = SQL & ",domprove =" & DBSet(Rs!dirdatos, "T")
                SQL = SQL & ", cpprove=" & DBSet(Rs!codposta, "T")
                SQL = SQL & ", pobprove=" & DBSet(Rs!desPobla, "T")
                SQL = SQL & ", proprove=" & DBSet(Rs!desProvi, "T")
                SQL = SQL & ", nifprove=" & DBSet(Rs!nifdatos, "T")
                SQL = SQL & ", codPAIS=" & DBSet(Rs!codpais, "T")
                    
            End If
            Rs.Close
            Set Rs = Nothing
                        
        End If
        SQL = SQL & " where numserie = " & DBSet(Text1(2).Text, "T")
        SQL = SQL & " and codmacta = " & DBSet(CodmactaAnt, "T") & " and numfactu = " & DBSet(NumFactuAnt, "T")
        SQL = SQL & " and fecfactu = " & DBSet(FecFactuAnt, "F")
        
        Conn.Execute SQL
        
        SQL = "update hlinapu set numfacpr = " & DBSet(Text1(25).Text, "T") & " where numserie = " & DBSet(Text1(2).Text, "T")
        SQL = SQL & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfacpr = " & DBSet(NumFactuAnt, "T")
        SQL = SQL & " and fecfactu = " & DBSet(FecFactuAnt, "F")
        
        Conn.Execute SQL
    End If
    
    '[Monica]12/09/2016: si la factura ha sido traspasada y no est� en cartera, no hacemos nada en cartera
    If EsFraProTraspasada And Not ExisteAlgunPago(Text1(2).Text, CodmactaAnt, Text1(25).Text, FecFactuAnt, False) Then Exit Sub
    
    
    If ExisteAlgunPago(Text1(2).Text, CodmactaAnt, Text1(25).Text, FecFactuAnt, True) Then
        MsgBoxA "Hay alg�n efecto que ya ha sido pagado. Revise cartera de pagos.", vbExclamation

        Set frmMens = New frmMensajes

        frmMens.Opcion = 28
        frmMens.Parametros = Trim(Text1(2).Text) & "|" & CodmactaAnt & "|" & Trim(Text1(25).Text) & "|" & Text1(26).Text & "|"
        frmMens.Show vbModal

        Set frmMens = Nothing

        ContinuarPago = False

        Exit Sub
    
    End If
    

    SQL = "delete from tmppagos where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute SQL
    
    ContinuarPago = False
    
    'If CargarPagosTemporal(Text1(5).Text, Text1(1).Text, ImporteFormateado(Text1(13).Text)) Then
    If CargarPagosTemporal(Text1(5).Text, FecFactuAnt, ImporteFormateado(Text1(13).Text)) Then
        ' Insertamos  FecFactuAnt
        If Not ExisteAlgunPago(Text1(2).Text, Text1(4).Text, Text1(25).Text, FecFactuAnt, False) Then
   
            'MONIIIII, lo tenias asi. �Why?
            'SQL = "select ccc.ctabanco,ccc.iban, ddd.nommacta "
            'SQL = SQL & " from cuentas ccc, cuentas ddd "
            'SQL = SQL & " where ccc.codmacta = " & DBSet(Text1(4).Text, "T")
            'SQL = SQL & " and ccc.ctabanco = ddd.codmacta "
           
            SQL = "select ctabanco,cuentas.iban, bancos.descripcion  from cuentas left join  bancos on  ctabanco  = bancos.codmacta "
            SQL = SQL & " Where Cuentas.codmacta = " & DBSet(Text1(4).Text, "T")
            
            CtaBanco = ""
            IBAN = ""
            NomBanco = ""
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs.EOF Then
                CtaBanco = DBLet(Rs.Fields(0))
                IBAN = DBLet(Rs.Fields(1))
                NomBanco = DBLet(Rs.Fields(2))
            End If
        
            TipForpa = DevuelveValor("select formapago.tipforpa from formapago where codforpa = " & DBSet(Text1(5).Text, "N"))
            
            Set frmPag_ = frmFacturasProPag
            ContinuarPago = False
            If IsNull(data1.Recordset!totfacpr) Then
                'Insertando
                SQL = ObtenerWhereCab(False) & " AND 1"
                SQL = DevuelveDesdeBD("totfacpr", "factpro", SQL, "1")
                If SQL = "" Then SQL = "0"
                
            Else
                SQL = data1.Recordset!totfacpr
            End If
            frmPag_.CodigoActual = CtaBanco & "|" & "|" & "|" & "|" & "|" & IBAN & "|" & TipForpa & "|" & NomBanco & "|" & SQL & "|"
                        
            frmPag_.Show vbModal
            Set frmPag_ = Nothing
    
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
                '[Monica]26/07/2017: si hay pagos que no estan pagados preguntamos si modificamos
                If MsgBoxA("� Desea recalcular los vencimientos ?", vbQuestion + vbYesNoCancel + vbDefaultButton1) = vbYes Then
                    CargarPagos
                End If
            Else
                MsgBoxA "No coincide el n�mero de pagos en tesoreria. Modif�quelos en cartera.", vbExclamation
                ' mandarlo al listview de cobros
            
                Set frmMens = New frmMensajes
                
                frmMens.Opcion = 28
                frmMens.Parametros = Trim(Text1(2).Text) & "|" & Trim(Text1(4).Text) & "|" & Trim(Text1(25).Text) & "|" & Text1(26).Text & "|"
                frmMens.Show vbModal
                
                Set frmMens = Nothing
            
            End If
        
        End If
    End If
    
     Screen.MousePointer = vbDefault
    Exit Sub
    
ePagosTesoreria:
    MuestraError Err.Number, "Pagos Tesoreria", Err.Description
End Sub

Private Function ExisteAlgunPago(Serie As String, Cuenta As String, FACTURA As String, FecFactu As String, Pagado As Boolean) As Boolean
Dim SQL As String
    
    SQL = "select count(*) from pagos where numserie = " & DBSet(Serie, "T")
    SQL = SQL & " and codmacta = " & DBSet(Cuenta, "T")
    SQL = SQL & " and numfactu = " & DBSet(FACTURA, "T")
    SQL = SQL & " and fecfactu = " & DBSet(FecFactu, "F")
    
    If Pagado Then
' un pago lo damos como pagado si el importe de pago es <> 0
'[Monica]12/09/2016: quito la condicion: numasien is null pq puede tener nro de transferencia y no modificariamos el importe total de transferencia
        SQL = SQL & " and ((imppagad <> 0 and not imppagad is null) " 'and numasien is null "
        
        '[Monica]26/07/2017: o se ha emitido documento
        SQL = SQL & " or emitdocum = 1) "
        
    End If
    
    ExisteAlgunPago = (TotalRegistros(SQL) <> 0)

End Function


Private Function PagosContabilizados(Serie As String, Cuenta As String, FACTURA As String, FecFactu As String) As String
Dim SQL As String
Dim CadResult As String
Dim Rs As ADODB.Recordset

    On Error GoTo ePagosContabilizados

    SQL = "select numasien, fechaent from hlinapu where numserie = " & DBSet(Serie, "T")
    SQL = SQL & " and codmacta = " & DBSet(Cuenta, "T")
    SQL = SQL & " and numfacpr = " & DBSet(FACTURA, "T")
    SQL = SQL & " and fecfactu = " & DBSet(FecFactu, "F")
    
    CadResult = ""
    
    If TotalRegistrosConsulta(SQL) = 0 Then
        CadResult = ""
    Else
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
Dim SQL As String
Dim Mens As String

    If ExisteAlgunPago(Text1(2).Text, Text1(4).Text, Text1(25).Text, FecFactuAnt, False) Then
        B = ActualizarPagos(Mens)
        
        If B Then
            SQL = PagosContabilizados(Text1(2).Text, Text1(4).Text, Text1(25).Text, FecFactuAnt)
            If SQL <> "" Then
                MsgBoxA "La factura tiene asientos que ya est�n contabilizados. Revise y modifique en su caso los siguientes asientos: " & vbCrLf & vbCrLf & SQL, vbExclamation
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
Dim SQL As String
Dim Aux As String

    On Error GoTo eUpdateaPagos
    
    UpdateaPagos = False

    B = True

    While Not Rs.EOF And B
        SQL = "update pagos set codmacta = " & DBSet(Text1(4).Text, "T")
        
        SQL = SQL & ", codforpa = " & DBSet(Text1(5).Text, "N")
        SQL = SQL & ", fecefect = " & DBSet(RS1!FecVenci, "F")
        SQL = SQL & ", impefect = " & DBSet(RS1!ImpVenci, "N")
        If Modo < 4 Then SQL = SQL & ", ctabanc1 = " & DBSet(CtaBanco, "T", "S")
        SQL = SQL & ", fecfactu = " & DBSet(Text1(26).Text, "F")
        
        If Pagado Then
            SQL = SQL & ", fecultpa = " & DBSet(FechaPago, "F") ' DBSet(Rs!FecVenci, "F")
            SQL = SQL & ", imppagad = " & DBSet(RS1!ImpVenci, "N")
        Else
            SQL = SQL & ", fecultpa = " & ValorNulo
            SQL = SQL & ", imppagad = " & ValorNulo
        End If
        
        If Rs!codmacta <> Text1(4).Text Then SQL = SQL & ", iban = " & DBSet(IBAN, "T", "S")
        SQL = SQL & ", numorden = " & DBSet(RS1!numorden, "N")
        SQL = SQL & " where numserie = " & DBSet(Text1(2).Text, "T") & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
        SQL = SQL & " and fecfactu = " & DBSet(FecFactuAnt, "F") & " and numorden = " & DBSet(Rs!numorden, "N")
        
        Conn.Execute SQL
        
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
Dim SQL As String
        
    On Error GoTo eInsertaPagos
        
    InsertaPagos = False
        
    CadInsert = "insert into pagos (numserie,codmacta,numfactu,fecfactu,numorden,codforpa,fecefect,impefect," & _
                "ctabanc1,fecultpa,imppagad,emitdocum," & _
                "text1csb,text2csb,nrodocum,referencia, iban,nomprove,domprove,pobprove,cpprove,proprove,nifprove,codpais,situacion,codusu) values "
    CadValues = ""
    
    While Not RS1.EOF
        i = i + 1
        
        SQL = DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(4).Text, "T") & "," & DBSet(Text1(25).Text, "T") & "," & DBSet(Text1(26).Text, "F") & "," & DBSet(i, "N") & ","
        SQL = SQL & DBSet(Text1(5).Text, "N") & "," & DBSet(RS1!FecVenci, "F") & "," & DBSet(RS1!ImpVenci, "N") & ","
        SQL = SQL & DBSet(CtaBanco, "T", "S") & ","
        
        If Pagado Then
'            B = ContabilizarPago
            SQL = SQL & DBSet(FechaPago, "F") & "," & DBSet(RS1!ImpVenci, "N") & ","
        Else
            SQL = SQL & ValorNulo & "," & ValorNulo & ","
        End If
        
        SQL = SQL & "0,"
        
        textCSB = "Factura " & Text1(25).Text & " de Fecha " & Text1(26).Text
        
        SQL = SQL & DBSet(textCSB, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(IBAN, "T", "S") & ","
        
        SQL = SQL & DBSet(Text1(15).Text, "T", "S") & "," & DBSet(Text1(16).Text, "T", "S") & "," & DBSet(Text1(18).Text, "T", "S") & "," & DBSet(Text1(17).Text, "T", "S") & ","
        SQL = SQL & DBSet(Text1(19).Text, "T", "S") & "," & DBSet(Text1(20).Text, "T", "S") & "," & DBSet(Text1(21).Text, "T", "S") & ","
        
        If Pagado Then
            SQL = SQL & "1"
        Else
            SQL = SQL & "0"
        End If
        
        ' la parte del codusu
        SQL = SQL & "," & DBSet(vUsu.Id, "N")
        
        
        CadValues = CadValues & "(" & SQL & "),"
    
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
Dim SQL As String
Dim Sql1 As String
Dim Nregs As Integer
Dim Nregs1 As Integer
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim CadInsert As String
Dim CadValues As String

    On Error GoTo eActualizarPagos

    ActualizarPagos = False


    SQL = "select * from pagos where numserie = " & DBSet(Text1(2).Text, "T") & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T") & " and fecfactu = " & DBSet(FecFactuAnt, "F")
    
    SQL = SQL & " order by numorden "
    Nregs = TotalRegistrosConsulta(SQL)
    
    Sql1 = "select * from tmppagos where codusu = " & vUsu.Codigo & " order by numorden "
    Nregs1 = TotalRegistrosConsulta(Sql1)
    
    If Nregs = Nregs1 Then
    ' Mismo nro de registros en pagos que en la temporal --> los actualizamos
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
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
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        i = 0
        
        Mens = "Actualizando Pagos: " & vbCrLf & vbCrLf
        B = UpdateaPagos(Rs, RS1, i, Mens)
        
        Set Rs = Nothing ' cierro el de pagos
        
        'borro los registros restantes de pagos
        Mens = "Eliminado Pagos restantes: " & vbCrLf & vbCrLf
        SQL = "delete from pagos "
        SQL = SQL & " where numserie = " & DBSet(Text1(2).Text, "T") & " and codmacta = " & DBSet(Text1(4).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
        SQL = SQL & " and fecfactu = " & DBSet(Text1(1).Text, "F") & " and numorden > " & DBSet(i, "N")
        
        Conn.Execute SQL
        
        Set RS1 = Nothing
    End If

    ActualizarPagos = B
    Exit Function

eActualizarPagos:
    Mens = Mens & Err.Description
End Function

'Public para poder llamarlo desde VER pago
Public Function InsertarPagos(ByRef Mens As String) As Boolean
Dim SQL As String
Dim textCSB As String
Dim CadInsert As String
Dim CadValues As String
Dim Rs As ADODB.Recordset
Dim i As Long

    On Error GoTo eInsertarPagos

    InsertarPagos = False

    SQL = "select * from tmppagos where codusu = " & DBSet(vUsu.Codigo, "N") & " order by numorden "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    i = 0
    Mens = "Insertando Pagos: " & vbCrLf & vbCrLf
    B = InsertaPagos(Rs, i, Mens)
    If Not B Then MsgBox Mens, vbExclamation
    
    Set Rs = Nothing
    
    InsertarPagos = B
    Exit Function
    
eInsertarPagos:
    MuestraError Err.Number, "Insertar Pagos", Err.Description & " " & Mens
End Function

Public Function CargarPagosTemporal(Forpa As String, FecFactu As String, TotalFac As Currency) As Boolean
Dim SQL As String
Dim CadValues As String
Dim Rsvenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarPagos

    CargarPagosTemporal = False

    SQL = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set Rsvenci = New ADODB.Recordset
    Rsvenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
                ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * Rsvenci!numerove <> TotalFac Then
                    ImpVenci = Round2(ImpVenci + (TotalFac - ImpVenci * Rsvenci.Fields(0).Value), 2)
                End If
            End If
            CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For i = 2 To Rsvenci!numerove
                FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next i
        End If
    End If
    
    Set Rsvenci = Nothing
    
    If CadValues <> "" Then
        SQL = "INSERT INTO tmppagos (codusu, numorden, fecvenci, impvenci)"
        SQL = SQL & " VALUES " & Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute SQL
    End If
    
    CargarPagosTemporal = True
    Exit Function

eCargarPagos:

End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = "numserie= " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
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
            AntiguoText1 = "."
            If txtaux(7).Text <> "" Then txtAux_LostFocus 7
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


Private Sub cmdAux3_Click()
          
        Set frmTIva = New frmBasico2
        AyudaTiposIva frmTIva
        Set frmTIva = Nothing
        
        PonFoco txtaux3(0)
        If txtaux3(0).Text <> "" Then txtaux3_LostFocus 0
End Sub

Private Sub Combo1_Click(Index As Integer)
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
        ElseIf Combo1(1).ListIndex = 5 Then
            Combo1(0).ListIndex = 24
            Text1(22).Text = "X"
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
    

    If Combo1(0).ListIndex = 18 Then
        ReferenciaCatastral True
    Else
        ReferenciaCatastral False
    End If
    
    
    
    ' intracomunitario
    If Index = 1 And (Modo = 1 Or Modo = 2 Or Modo = 3 Or Modo = 4) Then
        If Combo1(1).ListIndex = 1 Then
            ReferenciaCatastral False
            Combo1(3).visible = True
            Label11.visible = True
            Combo1(3).Enabled = True
            Label11.Enabled = True
            
            If Modo = 3 Then
                PosicionarCombo Combo1(3), Asc("A")
                Text1(27).Text = "A"
            End If
            
        Else
            Combo1(3).visible = False
            Label11.visible = False
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

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
'    If Index = 2 And (Modo = 3 Or Modo = 4) Then
'        If Combo1(Index).ListIndex = 0 Then
'            Text1(7).Text = ""
'            Text1(6).Text = ""
'            Text4(6).Text = ""
'        End If
'    End If
'
'    If Combo1(Index).ListIndex = -1 Then Exit Sub
'
'    ' en el caso de que sea bienes de inversion se pone en ambos combos
'    If Index = 0 And Modo <> 1 Then
'        If Chr(Combo1(Index).ItemData(Combo1(Index).ListIndex)) = "I" Then
'            Combo1(1).ListIndex = 4
'        Else
'            If Combo1(1).ListIndex = 4 Then Combo1(1).ListIndex = 0
'        End If
'    End If
'
'    If Index = 1 And (Modo = 3 Or Modo = 4) Then
'        If Combo1(1).ListIndex = 4 Then
'            PosicionarCombo Combo1(0), Asc("I")
'            Text1(22).Text = "I"
'        ElseIf Combo1(1).ListIndex = 5 Then
'            Combo1(0).ListIndex = 24
'            Text1(22).Text = "X"
'        Else
'            Combo1(0).ListIndex = 0
'            Text1(22).Text = "0"
'        End If
'    End If
'
'    If Index = 0 And (Modo = 1 Or Modo = 3 Or Modo = 4) Then
'        If Combo1(0).ListIndex = 0 Then
'            Text1(22).Text = "0"
'        Else
'            Text1(22).Text = Chr(Combo1(0).ItemData(Combo1(0).ListIndex))
'        End If
'    End If
'
'
'    If Combo1(0).ListIndex = 18 Then
'        ReferenciaCatastral True
'    Else
'        ReferenciaCatastral False
'    End If
'
'
'
'    ' intracomunitario
'    If Index = 1 And (Modo = 1 Or Modo = 2 Or Modo = 3 Or Modo = 4) Then
'        If Combo1(1).ListIndex = 1 Then
'            ReferenciaCatastral False
'            Combo1(3).visible = True
'            Label11.visible = True
'            Combo1(3).Enabled = True
'            Label11.Enabled = True
'
'            If Modo = 3 Then
'                PosicionarCombo Combo1(3), Asc("A")
'                Text1(27).Text = "A"
'            End If
'
'        Else
'            Combo1(3).visible = False
'            Label11.visible = False
'            Combo1(3).Enabled = False
'            Label11.Enabled = False
'
'            Text1(27).Text = ""
'        End If
'    End If
'    ' tipo de intracomunitario
'    If Index = 3 And (Modo = 1 Or Modo = 3 Or Modo = 4) Then
'        If Combo1(3).ListIndex = -1 Then
'            Text1(27).Text = ""
'        Else
'            Text1(27).Text = Chr(Combo1(3).ItemData(Combo1(3).ListIndex))
'        End If
'    End If
'
    
End Sub

Private Sub Form_Activate()
'    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    
    If PrimeraVez Then
        B = False
        If FACTURA <> "" Then
            B = True
            Modo = 2
            SQL = "Select * from factpro "
            SQL = SQL & " WHERE numserie = " & RecuperaValor(FACTURA, 1)
            SQL = SQL & " AND numregis =" & RecuperaValor(FACTURA, 2)
            SQL = SQL & " AND anofactu= " & RecuperaValor(FACTURA, 3)
            CadenaConsulta = SQL
            PonerCadenaBusqueda
            'BOTON lineas
            'BOTON lineas
            If Combo1(0).ListIndex = 18 Then ReferenciaCatastral True
            cboFiltro.ListIndex = 0
            
        Else
            Modo = 0
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE false  " ' numserie is null"
            data1.RecordSource = CadenaConsulta
            data1.Refresh
            
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
                MsgBoxA "Proceso de sistema. Frm_Activate", vbExclamation
            End If
        Else

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
    ContabilizaApunte = True  'por defecto
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
        
        .Buttons(6).Image = 31
        .Buttons(7).Image = 28
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
    
    'Totales IVA
    With Me.ToolbarAuxTot
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    
    
    
    For i = 0 To imgppal.Count - 1
        If i <> 0 And i <> 7 And i <> 9 Then imgppal(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    CargaFiltros
    
    
    Caption = "Facturas de Proveedor"
    
    NumTabMto = 1
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han ll�nies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "factpro"
    Ordenacion = " ORDER BY factpro.numserie, factpro.numregis , factpro.fecfactu"
    '************************************************
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    data1.ConnectionString = Conn
    '***** canviar el nom de la PK de la cap�alera; repasar codEmpre *************
    data1.RecordSource = "Select * from " & NombreTabla & " where numserie is null"
    data1.Refresh
       
    
    ModoLineas = 0
    DiarioPorDefecto = ""
       
    CargarColumnas
    
    CargarCombo
    

    Label1(21).visible = vParam.SIITiene
    Text1(28).visible = vParam.SIITiene
    If vParam.SIITiene Then
        Text1(28).Tag = "ID|N|S|||factpro|SII_ID|00000000||"
    Else
        Text1(28).Tag = ""
    End If
    
    
    
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
    
    Limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    For i = 0 To Combo1.Count - 1
        Me.Combo1(i).ListIndex = -1
    Next i

    Me.chkAux(0).Value = 0

    lw1.ListItems.Clear
    If vParam.SIITiene Then Text1(28).BackColor = vbWhite
   
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    FrameModifIVA.visible = False
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
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
    If Not data1.Recordset Is Nothing Then
        DespalzamientoVisible B And (data1.Recordset.RecordCount > 1)
    End If
    
    Toolbar1.Buttons(8).Enabled = B
    
    B = Modo = 2 Or Modo = 0 Or Modo = 5
    
    For i = 0 To 27
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
    
    
    If vParam.SIITiene Then Text1(28).Locked = Modo <> 1
    B = Modo = 1 Or (vParam.IvaEnFechaPago And (Modo = 3 Or Modo = 4))
    BloqueaTXT Text1(23), Not B
    Text1(23).Enabled = B
    
    
    
    
    
    
    
    ' observaciones
    
    imgppal(8).Enabled = Modo <> 0
    
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
       
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    B = (Modo < 5)
    chkVistaPrevia.visible = B
    
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
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub Desplazamiento(Index As Integer)
    If data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 1
            data1.Recordset.MoveFirst
        Case 2
            data1.Recordset.MovePrevious
            If data1.Recordset.BOF Then data1.Recordset.MoveFirst
        Case 3
            data1.Recordset.MoveNext
            If data1.Recordset.EOF Then data1.Recordset.MoveLast
        Case 4
            data1.Recordset.MoveLast
    End Select
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, Enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informaci� proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enla�a en el data1
'           -> Si no el carreguem sense enlla�ar a cap camp
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 ' lineas de totales
            tabla = "factpro_totales"
            SQL = "SELECT factpro_totales.numserie, factpro_totales.numregis, factpro_totales.fecharec, factpro_totales.anofactu, factpro_totales.numlinea, factpro_totales.baseimpo, factpro_totales.codigiva, factpro_totales.porciva,"
            SQL = SQL & " factpro_totales.porcrec, factpro_totales.impoiva, factpro_totales.imporec "
            SQL = SQL & " FROM " & tabla
            If Enlaza Then
                SQL = SQL & Replace(ObtenerWhereCab(True), "factpro", "factpro_totales")
            Else
                SQL = SQL & " WHERE factpro_totales.numlinea is null"
            End If
            SQL = SQL & " ORDER BY 1,2,3,4,5"
            
       
       
       Case 1 ' lineas de facturas
            tabla = "factpro_lineas"
            SQL = "SELECT factpro_lineas.numserie, factpro_lineas.numregis, factpro_lineas.fecharec, factpro_lineas.anofactu, factpro_lineas.numlinea, factpro_lineas.codmacta, cuentas.nommacta, factpro_lineas.baseimpo, factpro_lineas.codigiva,"
            SQL = SQL & " factpro_lineas.porciva, factpro_lineas.porcrec, factpro_lineas.impoiva, factpro_lineas.imporec, factpro_lineas.aplicret, IF(factpro_lineas.aplicret=1,'*','') as daplicret, factpro_lineas.codccost, ccoste.nomccost "
            SQL = SQL & " FROM (factpro_lineas LEFT JOIN ccoste ON factpro_lineas.codccost = ccoste.codccost) "
            SQL = SQL & " INNER JOIN cuentas ON factpro_lineas.codmacta = cuentas.codmacta "
            If Enlaza Then
                SQL = SQL & Replace(ObtenerWhereCab(True), "factpro", "factpro_lineas")
            Else
                SQL = SQL & " WHERE factpro_lineas.numlinea is null"
            End If
            SQL = SQL & " ORDER BY 1,2,3,4,5"
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = SQL
End Function


Private Sub frmAgen_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(26).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(26).Text = RecuperaValor(CadenaSeleccion, 2)
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
            MsgBoxA "Cuenta bloqueada: " & vFe, vbExclamation
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
            MsgBoxA "Cuenta bloqueada: " & vFe, vbExclamation
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


Public Sub EstablecerValoresSeleccionPago(CadenaTexto As String)
    If CadenaTexto <> "" Then
        CtaBanco = RecuperaValor(CadenaTexto, 1)
        IBAN = Replace(RecuperaValor(CadenaTexto, 2), " ", "")
        
        Pagado = RecuperaValor(CadenaTexto, 3)
        FechaPago = RecuperaValor(CadenaTexto, 4)
        ContinuarPago = True
    End If

End Sub


Private Sub frmPag__DatoSeleccionado2(CadenaSeleccion As String)
    EstablecerValoresSeleccionPago CadenaSeleccion
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
Dim RC As String

      
    If Modo = 6 Then
        'TOTALES IVA
        txtaux3(0).Text = RecuperaValor(CadenaSeleccion, 1)
        
    Else
        'Lineas
        txtaux(7).Text = RecuperaValor(CadenaSeleccion, 1)
        RC = "porcerec"
        txtaux(8).Text = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", txtaux(7), "N", RC)
        PonerFormatoDecimal txtaux(8), 4
        If RC = 0 Then
            txtaux(9).Text = ""
        Else
            txtaux(9).Text = RC
        End If
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
            MsgBoxA "Cuenta bloqueada: " & vFe, vbExclamation
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
        
        SQL = "Select * from factpro "
        SQL = SQL & " WHERE numserie = " & RecuperaValor(CadenaSeleccion, 1)
        SQL = SQL & " AND numfactu =" & RecuperaValor(CadenaSeleccion, 2)
        SQL = SQL & " AND anofactu= " & RecuperaValor(CadenaSeleccion, 3)
        
        CadenaConsulta = SQL
        PonerCadenaBusqueda
    End If
End Sub

Private Sub imgpdf_Click(Index As Integer)
    If Modo <> 2 Then Exit Sub
    
    If Index = 0 Then
        If txtPDF.Text <> "" Then
            MsgBoxA "Ya tiene asiganado un documento", vbExclamation
            Exit Sub
        End If
        
        If InsertarDesdeFichero Then TieneDocumentoAsociado
        
    Else
        If txtPDF.Text = "" Then
            MsgBoxA "No tiene asiganado nig�n documento", vbExclamation
            Exit Sub
        End If
    
        If MsgBoxA("Va a eliminar el documento asociado." & vbCrLf & "�Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
            
        SQL = ObtenerWhereCP(True)
        SQL = "DELETE FROM factpro_fichdocs" & SQL
        Ejecuta SQL
        TieneDocumentoAsociado
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
Dim CuentaAntes As String
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
        CuentaAntes = Text1(4).Text
        Set frmCtas = New frmColCtas
        frmCtas.DatosADevolverBusqueda = "0|1|2|"
        frmCtas.ConfigurarBalances = 3  'NUEVO
        frmCtas.Show vbModal
        Set frmCtas = Nothing
        If Modo <> 1 Then
            If CuentaAntes <> Text1(4).Text Then Text1_LostFocus 4
        End If
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
        
        frmAsi.ASIENTO = data1.Recordset!NumDiari & "|" & data1.Recordset!FechaEnt & "|" & data1.Recordset!NumAsien & "|"
        frmAsi.SoloImprimir = True
        frmAsi.Show vbModal
        
        Set frmAsi = Nothing
        
    Case 7
        'Fecha de liquidacion
        Indice = 23
        If Text1(23).Enabled And Not Text1(23).Locked Then
            Set frmF = New frmCal
            frmF.Fecha = Now
            If Text1(23).Text <> "" Then frmF.Fecha = CDate(Text1(23).Text)
            frmF.Show vbModal
            Set frmF = Nothing
            PonFoco Text1(23)
        End If
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
    
    If BLOQUEADesdeFormulario2(Me, data1, 1) Then BotonModificar
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
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
    If Index = 3 Then
        If KeyCode = 38 Or KeyCode = 40 Then Exit Sub
    End If
    KEYdown KeyCode
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonFoco Text1(2) ' <===
        ' *** si n'hi han combos a la cap�alera ***
    Else
        HacerBusqueda
        If data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            PonFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    CadB1 = ObtenerBusqueda2(Me, , 2, "FrameAux1")
    
    
    If CadB = "" And CadB1 = "" Then Exit Sub
    
    HacerBusqueda3 True
    
End Sub

Private Sub HacerBusqueda3(AplicaFiltros As Boolean)

    If AplicaFiltros Then CargarSqlFiltro
    
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
        ' *** foco al 1r camp visible de la cap�alera que siga clau primaria ***
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
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    data1.RecordSource = CadenaConsulta
    data1.Refresh
    
    If data1.Recordset.RecordCount <= 0 Then
        MsgBoxA "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        data1.Recordset.MoveFirst
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
    
    HacerBusqueda3 True
    
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
    'Contador de facturas
    Set Mc = New Contadores
    
    PonerModo 3
    CodmactaAnt = ""
    NumFactuAnt = ""
    Combo1(0).ListIndex = 0
    Combo1(1).ListIndex = 0
    Combo1(2).ListIndex = 0
    
    If Now <= DateAdd("yyyy", 1, vParam.fechafin) Then Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(9).Text = "0,00"
    
    ' por defecto para todos cuando insertamos es 1
    Text1(2).Text = "1"
    Text1_LostFocus (2)
    
    FrameDatosFiscales.visible = False
    
    Text1_LostFocus (1)
    PonFoco Text1(2)
    ' ***********************************************************
    
End Sub


Private Sub BotonModificar()

    
    '---------
    'MODIFICAR
    '----------
    'Comprobamos la fecha pertenece al ejercicio
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk >= 2 Then
        If varFecOk = 2 Then
            SQL = varTxtFec
        Else
            SQL = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBoxA SQL, vbExclamation
        Exit Sub
    End If
    
    'Falta ver sii
    If Not ComprobarPeriodo2(23, 1) Then Exit Sub
    If Not ComprobarPeriodo2(1, 2) Then Exit Sub
    
    PonerModo 4

    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonFoco Text1(25)
    ' *********************************************************
    
    FecFactuAnt = Text1(26).Text
    FecRecepAnt = Text1(1).Text
    NumFactuAnt = Text1(25).Text
    CodmactaAnt = Text1(4).Text
    
    NumDiario = 0
    ContabilizaApunte = True
    'Comprobamos que no esta actualizada ya
    If Not IsNull(data1.Recordset!NumAsien) Then
        Numasien2 = data1.Recordset!NumAsien
        If Numasien2 = 0 Then
            MsgBoxA "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
            PonerModo 2
            Exit Sub
        End If
        If Val(DBLet(data1.Recordset!no_modifica_apunte, "N")) = 1 Then
            ContabilizaApunte = False
            Numasien2 = data1.Recordset!NumAsien
        Else
            Numasien2 = data1.Recordset!NumAsien
            NumDiario = data1.Recordset!NumDiari
        End If
    Else
        Numasien2 = -1
    End If
        
        
    'Si viene a esta factura buscando por un campo k no sea clave entonces no le dejo seguir
    If InStr(1, data1.Recordset.Source, "numasien") Then
        MsgBoxA "Busque la factura por su numero de factura", vbExclamation
        Numasien2 = -1
        PonerModo 2
        Exit Sub
    End If
    
    If Numasien2 >= 0 Then
        'Tengo desintegrar la factura del hco
        If Not Desintegrar Then
            TerminaBloquear
            Exit Sub
        End If
        Text1(8).Text = ""
        If Not ContabilizaApunte Then Text1(8).Text = Numasien2
        
    End If
    
    If Mc Is Nothing Then Set Mc = New Contadores
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    DespalzamientoVisible False
    'PonFoco Text1(1)
    ModificarPagos = False
    
    
End Sub


Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim i As Long
    Dim Fec As Date
    Dim Mc As Contadores
    Dim SqlLog As String
    
    'Ciertas comprobaciones
    If data1.Recordset Is Nothing Then Exit Sub
    If data1.Recordset.EOF Then Exit Sub
    DataGridAux(1).Enabled = False

    'Comprobamos la fecha pertenece al ejercicio
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk >= 2 Then
        If varFecOk = 2 Then
            SQL = varTxtFec
        Else
            SQL = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBoxA SQL, vbExclamation
        Exit Sub
    End If

    'Comprobamos si esta liquidado
    If Not ComprobarPeriodo2(23, 1) Then Exit Sub
    If Not ComprobarPeriodo2(1, 2) Then Exit Sub
    
    'Comprobamos que no esta actualizada ya
    SQL = ""
    If Not IsNull(data1.Recordset!NumAsien) Then
        SQL = "Esta factura ya esta contabilizada. "
    End If
    
    SQL = SQL & vbCrLf & vbCrLf & "Va usted a eliminar la factura :" & vbCrLf
    SQL = SQL & "Numero : " & data1.Recordset!NumFactu & vbCrLf
    SQL = SQL & "Fecha  : " & data1.Recordset!FecFactu & vbCrLf
    SQL = SQL & "Proveedor : " & Me.data1.Recordset!codmacta & " - " & Text4(4).Text & vbCrLf
    SQL = SQL & vbCrLf & "          �Desea continuar ?" & vbCrLf
    
    If Not EliminarDesdeActualizar Then
        If MsgBoxA(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    NumRegElim = data1.Recordset.AbsolutePosition
    Screen.MousePointer = vbHourglass
    'Lo hara en actualizar
    i = 0
    If Not IsNull(data1.Recordset!NumAsien) Then
        i = data1.Recordset!NumAsien
       If Val(DBLet(data1.Recordset!no_modifica_apunte, "N")) = 1 Then i = 0   'YA que nos e tratan los apuntes
    End If
    If i > 0 Then
        
            'Memorizamos el numero de asiento y la fechaent para ver si devolvemos el contador
            'de asientos
            i = data1.Recordset!NumAsien
            Fec = data1.Recordset!FechaEnt
        
            'La borrara desde actualizar
            AlgunAsientoActualizado = False
           
            
            SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
            SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text
            SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text

            
            With frmActualizar
                .OpcionActualizar = 9
                .NumAsiento = data1.Recordset!NumAsien
                .NumFac = data1.Recordset!Numregis
                .FACTURA = data1.Recordset!NumFactu
                .Proveedor = data1.Recordset!codmacta
                .FechaAsiento = data1.Recordset!fecharec
                .FechaFactura = data1.Recordset!FecFactu
                .NUmSerie = data1.Recordset!NUmSerie & "|" & data1.Recordset!anofactu & "|"
                .NumDiari = data1.Recordset!NumDiari
                .FechaAnterior = data1.Recordset!fecharec
                .SqlLog = SqlLog
                .Show vbModal
            End With
            Set Mc = New Contadores
            Mc.DevolverContador "0", Fec <= vParam.fechafin, i
            Set Mc = Nothing
        
    Else
        'La borrara desde este mismo form
        Conn.BeginTrans
        
        i = data1.Recordset!Numregis
        Fec = data1.Recordset!fecharec
        If BorrarFactura Then
            'LOG
            SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " " & Text1(1).Text
            SqlLog = SqlLog & vbCrLf & "Proveed.: " & Text1(4).Text & " " & Text4(4).Text
            SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text
            
            vLog.Insertar 10, vUsu, SqlLog
        
            AlgunAsientoActualizado = True
            Conn.CommitTrans
            Set Mc = New Contadores
            Mc.DevolverContador CStr(DBLet(data1.Recordset!NUmSerie)), (Fec <= vParam.fechafin), i
            Set Mc = Nothing
            
            
            'Mayo 2018
            'Avisa o borra vencimientos
            SQL = "select count(*) from pagos where numserie = " & DBSet(data1.Recordset!NUmSerie, "T") & " and codmacta = " & DBSet(data1.Recordset!codmacta, "T")
            SQL = SQL & " and numfactu = " & DBSet(data1.Recordset!NumFactu, "T") & " and fecfactu = " & DBSet(data1.Recordset!FecFactu, "F") & " and imppagad <> 0 and not imppagad is null "
        
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "Hay pagos que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
            Else
              
                SQL = "DELETE from pagos where numserie = " & DBSet(data1.Recordset!NUmSerie, "T") & " and codmacta = " & DBSet(data1.Recordset!codmacta, "T")
                SQL = SQL & " and numfactu = " & DBSet(data1.Recordset!NumFactu, "T") & " and fecfactu = " & DBSet(data1.Recordset!FecFactu, "F") & " and imppagad <> 0 and not imppagad is null "
        
                Ejecuta SQL
            End If
            
            
        Else
            AlgunAsientoActualizado = False
            Conn.RollbackTrans
        End If
    End If
    If Not AlgunAsientoActualizado Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    data1.Refresh
    If data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid 1, False
        PonerModo 0
        Else
            data1.Recordset.MoveFirst
            NumRegElim = NumRegElim - 1
            If NumRegElim > 1 Then
                For i = 1 To NumRegElim - 1
                    data1.Recordset.MoveNext
                Next i
            End If
            PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function BorrarFactura() As Boolean
    
    On Error GoTo EBorrar
    SQL = " WHERE numserie = '" & data1.Recordset!NUmSerie & "'"
    SQL = SQL & " AND numregis = " & data1.Recordset!Numregis
    SQL = SQL & " AND anofactu= " & data1.Recordset!anofactu
    'Las lineas
    AntiguoText1 = "DELETE from factpro_totales " & SQL
    Conn.Execute AntiguoText1
    AntiguoText1 = "DELETE from factpro_lineas " & SQL
    Conn.Execute AntiguoText1
    AntiguoText1 = "DELETE from factpro_fichdocs " & SQL
    Conn.Execute AntiguoText1
    
    'La factura
    AntiguoText1 = "DELETE from factpro " & SQL
    Conn.Execute AntiguoText1
    
    ComprobarContador data1.Recordset!NUmSerie, CDate(Text1(1).Text), data1.Recordset!Numregis
    
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

    If data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    For i = 1 To DataGridAux.Count ' - 1
        If i <> 3 Then
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
        End If
    Next i
    
    imgppal(6).Enabled = (Text1(8).Text <> "")
    imgppal(6).visible = (Text1(8).Text <> "")
        
    If Text1(30).Text = "0,00" Then Text1(30).Text = ""
    If Text1(31).Text = "0,00" Then Text1(31).Text = ""
        
        
    Text4(2).Text = PonerNombreDeCod(Text1(2), "contadores", "nomregis", "tiporegi", "T")
    Text4(4).Text = PonerNombreDeCod(Text1(4), "cuentas", "nommacta", "codmacta", "T")
    Text4(6).Text = PonerNombreDeCod(Text1(6), "cuentas", "nommacta", "codmacta", "T")
    Text4(5).Text = PonerNombreDeCod(Text1(5), "formapago", "nomforpa", "codforpa", "N")
    Text4(21).Text = PonerNombreDeCod(Text1(21), "paises", "nompais", "codpais", "T")
    
    If vParam.SIITiene Then Color_CampoSII
    
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
    TieneDocumentoAsociado
    
    
    FrameModifIVA.visible = False
    If Modo = 2 And vUsu.Nivel = 0 Then
        If Val(data1.Recordset!no_modifica_apunte) = 1 Then
            FrameModifIVA.visible = True
    
            ToolbarAuxTot.Buttons(2).Enabled = Me.lw1.ListItems.Count > 0
            ToolbarAuxTot.Buttons(3).Enabled = Me.lw1.ListItems.Count > 0
        End If
    End If

    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
    
    
End Sub


Private Sub cmdCancelar_Click()
Dim i As Integer
Dim v
Dim Recalcular As Boolean

    Select Case Modo
        Case 1, 3 'B�squeda, Insertar
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
                Text1(1).Text = data1.Recordset!fecharec
                Text1(0).Text = data1.Recordset!Numregis
                Text1(14).Text = data1.Recordset!anofactu
                If Not IntegrarFactura_(False) Then
                    Modo = 4 'lo pongo por si acaso
                    Exit Sub
                End If
            End If
            PonerCampos
            Modo = 4  'Reestablezco el modo para que vuelva a hahacer ponercampos
            '--DesBloqueaRegistroForm Me.Text1(0)
            TerminaBloquear
            
            
            PonerModo 2
            If vParam.SIITiene Then Color_CampoSII
            
            'Contador de facturas
            Set Mc = Nothing
                
        Case 5 'LL�NIES
            TerminaBloquear
        
            If ModoLineas = 1 Then 'INSERTAR
                ModoLineas = 0
                DataGridAux(1).AllowAddNew = False
                If Not AdoAux(1).Recordset.EOF Then AdoAux(1).Recordset.MoveFirst
                
                If AdoAux(1).Recordset.EOF Then
                    If MsgBox("No se permite una factura sin l�neas " & vbCrLf & vbCrLf & "� Desea eliminar la factura ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                        BotonEliminar True
                        Exit Sub
                    Else
                        ModoLineas = 1
                        cmdAceptar_Click
                        Exit Sub
                    End If
                End If
                
            End If
            
            Recalcular = True
            If ModoLineas = 2 Then Recalcular = False
            ModoLineas = 0
            LLamaLineas 1, 0, 0
            
            Modo = 2   'Para que el lostfocus NO haga nada
            If Numasien2 > 0 Then
                'Ha cancelado. Tendre que situar los campos correctamente
                'Es decir. Anofacl
                Text1(1).Text = data1.Recordset!fecharec
                Text1(0).Text = data1.Recordset!Numregis
                Text1(14).Text = data1.Recordset!anofactu
                If Not IntegrarFactura_(False) Then
                    Modo = 4 'lo pongo por si acaso
                    Exit Sub
                End If
                If Recalcular Then PagosTesoreria
            Else
                ' cogemos un nro.de asiento para integrarlo
                Set Mc = New Contadores
                
                i = FechaCorrecta2(CDate(Text1(1).Text))
                If Mc.ConseguirContador("0", (i = 0), False) = 0 Then
                    Text1(8).Text = Format(Mc.Contador, "0000000")
                    Numasien2 = Mc.Contador
                    ContabilizaApunte = True
                    If ModificaDesdeFormulario2(Me, 2, "Frame2") Then
                        If Not IntegrarFactura_(False) Then
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
    Case 6
                    
            ModoLineas = 0
            LLamaLineas 2, 0, 0
            PonerModo 2
            CargaDatosLW
            
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Cad As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOK = False
    
    'fecha de liquidacion
    If Not vParam.IvaEnFechaPago Then Text1(23).Text = Text1(1).Text
    
    
    
    
    
    
    If Combo1(0).ListIndex = 0 Then
        Text1(22).Text = "0"
    Else
        Text1(22).Text = Chr(Combo1(0).ItemData(Combo1(0).ListIndex))
    End If

    
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    
    'Fecha recepcion y liquidacion NO puede ser menor a fecha emision factura
    Cad = ""
    If CDate(Text1(1).Text) < CDate(Text1(26).Text) Then Cad = Cad & vbCrLf & " - Fecha recepci�n"
    If CDate(Text1(23).Text) < CDate(Text1(26).Text) Then Cad = Cad & vbCrLf & " - Fecha liquidaci�n"
       
    If Cad <> "" Then
        Cad = "Error en fechas factura proveedor. " & vbCrLf & Cad & vbCrLf & vbCrLf
        Cad = Cad & "No puede ser inferior a fecha emisi�n factura."
        MsgBox Cad, vbExclamation
        Exit Function
    End If
    
    If CDate(Text1(26).Text) < "01/01/2000" Or CDate(Text1(26).Text) > CDate("01/01/" & Year(vParam.fechafin) + 5) Then
        MsgBoxA "Fecha emisi�n factura incorrecta *****", vbExclamation
        Exit Function
    End If
    ' NOV 2007
    ' NUEVA ambitode fecha activa
    '       0 .- A�o actual
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
    
    ' controles a�adidos de la factura de david
    'No puede tener % de retencion sin cuenta de retencion
    If Combo1(2).ListIndex > 0 Then
       If ((Text1(6).Text = "") Or (Text1(7).Text = "")) Then
            MsgBox "Indique porcentaje y cuenta de rentenci�n ", vbExclamation
            B = False
            PonFoco Text1(6)
            Exit Function
        End If
    Else
        If ((Text1(6).Text <> "") Or (Text1(7).Text <> "")) Then
            MsgBox "Ha indicado porcentaje y/o cuenta de rentenci�n sin indicar el tipo", vbExclamation
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
    
    
    'Ahora. Si estamos modificando, y el a�o factura NO es el mismo, entonces
    'la estamos liando, y para evitar lios, NO dejo este tipo de modificacion
    If Modo = 4 Then
        If CDate(Text1(1).Text) <> data1.Recordset!fecharec Then
            'HAN CAMBIADO LA FECHA. Veremos si dejo
            If Year(CDate(Text1(1).Text)) <> data1.Recordset!anofactu Then
                MsgBox "No puede cambiar de a�o la factura. ", vbExclamation
                B = False
                Exit Function
            End If
            
            '[Monica]19/01/2017
            ' si hay alguna factura de la serie con numero de registro mayor y fecha de recepcion inferior a la que hemos introducido damos aviso
            SQL = "select count(*) from factpro where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis > " & DBSet(Text1(0).Text, "N") & " and fecharec < " & DBSet(Text1(1).Text, "F") & " and anofactu = " & Year(CDate(Text1(1).Text))
            If DevuelveValor(SQL) <> 0 Then
                If MsgBox("Existe alguna factura de la serie con nro.registro superior y fecha de recepci�n inferior a �sta." & vbCrLf & vbCrLf & "� Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
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
        SQL = "select count(*) from factpro where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
        SQL = SQL & " and anofactu = year(" & DBSet(Text1(1).Text, "F") & ") and codmacta = " & DBSet(Text1(4).Text, "T")
        If DevuelveValor(SQL) <> 0 Then
            MsgBox "Factura ya existe para esta serie proveedor a�o. Revise.", vbExclamation
            B = False
            Exit Function
        End If
    End If
    





    DatosOK = B

EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(numserie=" & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N") & ") "
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        If vParam.SIITiene Then Color_CampoSII
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
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE (numasien=" & Trim(Text1(0).Text) & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N") & ") "
        ' ***********************************************************************
        
    Conn.Execute "DELETE FROM hlinapu " & vWhere
    
    Conn.Execute "DELETE FROM hcabapu_fichdocs " & vWhere

'    ' *******************************
    Conn.Execute "Delete from " & NombreTabla & vWhere
       
    'El LOG
    vLog.Insertar 3, vUsu, SQL
       
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
            SQL = ""
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                If Index = 1 Then Text1(14).Text = ""
                Text1(Index).Text = ""
                PonFoco Text1(Index)
                Exit Sub
            End If
            ModificandoLineas = FechaCorrecta2(CDate(Text1(Index).Text))
            If Modo = 1 Then ModificandoLineas = 0
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
                MsgBoxA RC, vbExclamation
                Text1(Index).Text = ""
                If Index = 1 Then Text1(14).Text = ""
                PonFoco Text1(Index)
                Exit Sub
            End If
            
            Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
            If Index = 1 And Modo <> 1 Then Text1(14).Text = Year(CDate(Text1(Index).Text))
            
            i = 0  'No actualiza fecha liquidacion
            If Index = 1 Then
                If Modo = 3 Then
                    i = 1
                Else
                    If Modo = 4 Then
                        If Not vParam.IvaEnFechaPago Then i = 1
                    End If
                End If
            End If
            If i = 1 Then Text1(23).Text = Text1(1).Text
            
            'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
            'que el periodo de liquidacion del IVA no ha pasado.
            If Modo <> 1 Then If Not ComprobarPeriodo2(Index, IIf(Index = 1, 0, 1)) Then PonFoco Text1(Indice)
    

        Case 2 ' Serie
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Debe ser un n�mero: " & Text1(Index).Text, vbExclamation
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
        
        Case 3
            If Len(Text1(Index).Text) > 0 Then PonCursorInicio
        
        Case 4, 6 ' cuenta de proveedor, cuenta de retencion
                'Cuenta proveedor
                If AntiguoText1 = Text1(Index).Text Then Exit Sub
                RC = Text1(Index).Text
                i = Index
                
                If CuentaCorrectaUltimoNivel(RC, SQL) Then
                    Text1(Index).Text = RC
                    Text4(i).Text = SQL
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
                    
                    If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                            RC = RellenaCodigoCuenta(Text1(Index).Text)
                            SQL = "La cuenta: " & RC & " no existe. �Desea crearla?"
                            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                                CadenaDesdeOtroForm = RC
                                cmdAux(0).Tag = Indice
                                Set frmC = New frmColCtas
                                frmC.DatosADevolverBusqueda = "0|1|"
                                frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                                frmC.Show vbModal
                                Set frmC = Nothing
                                If Text1(4).Text = RC Then SQL = "" 'Para k no los borre
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
                    MsgBox "No existe el Pa�s. Reintroduzca.", vbExclamation
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

Private Sub PonCursorInicio()
    On Error Resume Next
    Text1(3).SelStart = 0
    If Err.Number <> 0 Then Err.Clear
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
Dim CarpetaAlbar As String
Dim CadFiltro_OLD As String

    Select Case Button.Index
    
        Case 1 'Datos Fiscales
            Me.FrameDatosFiscales.visible = Not Me.FrameDatosFiscales.visible
           
        Case 2 'Cartera de Cobros
            If Not data1.Recordset.EOF Then
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

            
            
            
            
        Case 6, 7
            If Not (Modo = 2 Or Modo = 0) Then Exit Sub
            
            
            
            If Button.Index = 6 Then
                'Albaranes
                If FijarCarpetaDestinoPendientes(False) Then
                    frmProAlbaranres.Carpeta = SQL
                    frmProAlbaranres.Show vbModal
                End If
            Else
            
                'Recepcion facturas proveedor.
                ' Sobre una carpeta compartida : PathFacturasProv , creara una estructura
                '   PathFacturasProv    -->  \00001 Nomempresa1\
                '                       -->  \00002 Nomempresa2\
                '                       -->  \00005 Nomempresa5\
                CarpetaAlbar = ""
                If Not FijarCarpetaDestinoPendientes(False) Then Exit Sub
                CarpetaAlbar = SQL
                SQL = ""
                If FijarCarpetaDestinoPendientes(True) Then
                
                    
                    
                    CadenaDesdeOtroForm = ""
                    frmAlfresQFRA.CarpetaAlbaranes = CarpetaAlbar
                    frmAlfresQFRA.CarpetaDestino = SQL
                    frmAlfresQFRA.Show vbModal
                    If CadenaDesdeOtroForm <> "" Then
                     
                        'OK. La hemos inseertado   ejmplo: 1|712|2018|

                        CadB = "factpro.numserie = " & RecuperaValor(CadenaDesdeOtroForm, 1) & " AND factpro.numregis=" & RecuperaValor(CadenaDesdeOtroForm, 2) & " AND factpro.anofactu = " & RecuperaValor(CadenaDesdeOtroForm, 3)
                        CadB1 = ""
                        CadFiltro_OLD = cadFiltro
                        cadFiltro = ""
                        PonerModo 0
                        HacerBusqueda3 False   'NO APLICA FILTROS
                        
                        If Me.data1.Recordset.EOF Then
                            MsgBox "No se encuantra factura insertada. " & CadenaDesdeOtroForm, vbExclamation
                        Else
                            'Primer paso. Asiento
                            FecRecepAnt = Text1(1).Text
                            IntegrarFactura_ False
                            FecFactuAnt = Text1(1).Text
                            PagosTesoreria
                        
                            CadB = "factpro.numserie = " & RecuperaValor(CadenaDesdeOtroForm, 1) & " AND factpro.numregis=" & RecuperaValor(CadenaDesdeOtroForm, 2) & " AND factpro.anofactu = " & RecuperaValor(CadenaDesdeOtroForm, 3)
                            CadB1 = ""
                            PonerModo 0
                            HacerBusqueda3 False
                        
                            'Si la cuenta es de innmovilizado, llamo al inmov
                            Dim Cade As String
                            Cade = "factpro_lineas.numserie = " & RecuperaValor(CadenaDesdeOtroForm, 1) & " AND factpro_lineas.numregis=" & RecuperaValor(CadenaDesdeOtroForm, 2) & " AND factpro_lineas.anofactu = " & RecuperaValor(CadenaDesdeOtroForm, 3)
                            Cade = Cade & " AND factpro_lineas.codmacta=cuentas.codmacta "
                            Cade = Cade & " AND factpro_lineas.codmacta like '2%' AND 1"
                            
                            Cade = DevuelveDesdeBD("concat(cuentas.codmacta,'|',nommacta,'|')", "factpro_lineas,cuentas", Cade, "1")
                            If Len(Cade) > 8 Then
                                txtaux(5).Text = RecuperaValor(Cade, 1)
                                txtAux2(5).Text = RecuperaValor(Cade, 2)
                                Cade = RecuperaValor(Cade, 1)
                                CrearElementoInmovilizado_ Cade
                            
                            End If
                            
                        End If
                        cadFiltro = CadFiltro_OLD
                    End If
            
                End If
            
            End If
    End Select

End Sub


 'Recepcion facturas proveedor o albaranes
                ' Sobre una carpeta compartida : PathFacturasProv , creara una estructura
                '   PathFacturasProv    -->  \00001 Nomempresa1\
                '                       -->  \00002 Nomempresa2\
                '                       -->  \00005 Nomempresa5\
Private Function FijarCarpetaDestinoPendientes(facturas As Boolean) As Boolean
Dim C As String
Dim Nombre As String


On Error GoTo eFijarCarpetaDestinoFacturasPendientes
    
    i = -1
    Msg = ""
    SQL = ""
    If facturas Then
        Nombre = vParam.PathFacturasProv
    Else
        Nombre = vParam.PathAlbaranesProv
    End If
    If Nombre = "" Then
        Msg = "Falta configurar parametros"
    Else
        If Dir(Nombre, vbDirectory) = "" Then
            Msg = "No existe la carpeta destino: " & Nombre
        Else
            i = 0
            C = Dir(Nombre & "\", vbDirectory)   ' Recupera la primera entrada.
            Do While C <> ""   ' Inicia el bucle.
               ' Ignora el directorio actual y el que lo abarca.
               If C <> "." And C <> ".." Then
                  ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
                  If (GetAttr(Nombre & "\" & C) And vbDirectory) = vbDirectory Then
                    Msg = Mid(C, 1, 5)
                    If IsNumeric(Msg) Then
                        If Msg = Format(vEmpresa.codempre, "00000") Then
                            'Bravooooooooo
                            'es esta la carpeta
                            i = 1
                            SQL = C
                        End If
                    End If
                    Msg = ""
                  End If   ' solamente si representa un directorio.
               End If
               If i = 0 Then
                    C = Dir   ' Obtiene siguiente entrada.
               Else
                    C = ""  'YA la ha encontrado
                End If
            Loop
            
            If i = 0 Then
                'NO existe la carpeta
                C = Format(vEmpresa.codempre, "00000") & " " & vEmpresa.nomresum
                MkDir Nombre & "\" & C
                
                'Si llega aqui... perfecto
                SQL = C
            Else
                'Ya lo ha asignado arriba
            End If
            SQL = Nombre & "\" & SQL
            Msg = ""
        End If
    End If
    If Msg <> "" Then
        MsgBox Msg, vbExclamation
        
    Else
        FijarCarpetaDestinoPendientes = True
    End If
    Exit Function
eFijarCarpetaDestinoFacturasPendientes:
    MuestraError Err.Number, Err.Description
End Function


Private Sub ComprobarFrasSinAsiento()
Dim SQL As String
Dim vCadena As String
Dim vCadena2 As String
Dim Rs As ADODB.Recordset
Dim IntegrarFactura As Boolean
Dim i As Integer
Dim Nregs As Long
Dim SqlLog As String

    
    SQL = "select * from factpro where (numasien = 0 or numasien is null or fechaent is null or numdiari is null) "
    If cadFiltro <> "" Then SQL = SQL & " and " & cadFiltro

    vCadena = ""
    vCadena2 = ""
    
    If TotalRegistrosConsulta(SQL) <> 0 Then
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Nregs = 1
        
        While Not Rs.EOF
            vCadena = vCadena & "Fra.Reg. " & DBLet(Rs!NUmSerie) & " " & Format(DBLet(Rs!Numregis), "0000000") & " " & DBLet(Rs!fecharec, "F")
            vCadena2 = vCadena2 & "(" & DBSet(Rs!NUmSerie, "T") & "," & Format(DBSet(Rs!Numregis, "N"), "0000000") & "," & Year(DBLet(Rs!fecharec, "F")) & "),"
            
            If (Nregs Mod 2) = 0 Then
                vCadena = vCadena & vbCrLf
            Else
                vCadena = vCadena & "  "
            End If
            
            Nregs = Nregs + 1
            
            Rs.MoveNext
        Wend
        
        If MsgBox("Las siguientes facturas no tienen Asiento asociado: " & vbCrLf & vbCrLf & vCadena & vbCrLf & vbCrLf & " � Asigna asiento ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            Rs.MoveFirst
            
            
            While Not Rs.EOF
                IntegrarFactura = False
                
                ' cogemos un nro.de asiento para integrarlo
                Set Mc = New Contadores
                
                i = FechaCorrecta2(CDate(DBLet(Rs!FecFactu, "F")))
                If Mc.ConseguirContador("0", (i = 0), False) = 0 Then
                    
                    Numasien2 = Mc.Contador
                
                    SqlLog = "Registro : " & Rs!NUmSerie & " " & Rs!Numregis & " de fecha " & Rs!fecharec
                    SqlLog = SqlLog & vbCrLf & "Cuenta  : " & DBLet(Rs!codmacta, "T") & " " & DBLet(Rs!Nommacta, "T")
                    SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Rs!totfacpr, "N")
                    
                    With frmActualizar
                        .OpcionActualizar = 8
                        'NumAsiento     --> CODIGO FACTURA
                        'NumDiari       --> A�O FACTURA
                        'NUmSerie       --> SERIE DE LA FACTURA
                        'FechaAsiento   --> Fecha factura
                        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
                        .NumFac = DBLet(Rs!Numregis, "N")
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
                        SQL = "update factpro set numdiari = " & DBSet(NumDiario, "N") & ", fechaent = " & DBSet(Rs!FecFactu, "F") & ", "
                        SQL = SQL & " numasien = " & DBSet(Numasien2, "N") & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and anofactu = year("
                        SQL = SQL & DBSet(Rs!fecharec, "F") & ") and numregis = " & DBSet(Rs!Numregis, "N")
                    
                        Conn.Execute SQL
                        
                    End If
                End If
                
                Rs.MoveNext
            Wend
        
            vLog.Insertar 29, vUsu, vCadena
        
        End If
        
        
        Set Rs = Nothing
        
        CadB = "(factpro.numserie, factpro.numregis, factpro.anofactu) in (" & Mid(vCadena2, 1, Len(vCadena2) - 1) & ")"
        HacerBusqueda3 True
    
    Else
        MsgBoxA "No hay facturas sin asiento asignado.", vbExclamation
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
            SQL = varTxtFec
        Else
            SQL = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBox SQL, vbExclamation
        Exit Sub
    End If

    'Marzo 2019
    ' Si no modifica el apunte, NO dejo tocar lineas
    If Val(DBLet(data1.Recordset!no_modifica_apunte, "N")) = 1 Then
        MsgBoxA "Factura integrada. Modifique el apunte", vbExclamation
        Exit Sub
    End If
    



    '**** parte correspondiente por si la factura est� contabilizada
    NumDiario = 0
    'Comprobamos que no esta actualizada ya
    If Not IsNull(data1.Recordset!NumAsien) Then
        Numasien2 = data1.Recordset!NumAsien
        If Numasien2 = 0 Then
            MsgBox "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
            Exit Sub
        End If
            
            
        'Creo que no es obligatorio, pero en clientes tambien esta asi
        ContabilizaApunte = True
        If Val(DBLet(data1.Recordset!no_modifica_apunte, "N")) = 1 Then ContabilizaApunte = False
       
            
            
        Numasien2 = data1.Recordset!NumAsien
        NumDiario = data1.Recordset!NumDiari
    Else
        Numasien2 = -1
    End If
    
    If Not ComprobarPeriodo2(23, 1) Then Exit Sub
    If Not ComprobarPeriodo2(1, 2) Then Exit Sub
    
        
    
    
    'Llegados aqui bloqueamos desde form
    If Not BLOQUEADesdeFormulario2(Me, data1, 1) Then Exit Sub
    
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
            'A�ADIR linea factura
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



Private Sub ToolbarAuxTot_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim alto As Integer
Dim Indc As Integer

    If Modo <> 2 Or vUsu.Nivel > 0 Then Exit Sub
    
    
    If Button.Index > 1 Then
        If Me.lw1.SelectedItem Is Nothing Then
            MsgBox "Seleccione algun IVA", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    '----------
    'Comprobamos la fecha pertenece al ejercicio
    varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
    If varFecOk >= 2 Then
        If varFecOk = 2 Then
            SQL = varTxtFec
        Else
            SQL = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBoxA SQL, vbExclamation
        Exit Sub
    End If
    
    
    If Not ComprobarPeriodo2(23, 1) Then Exit Sub
    If Not ComprobarPeriodo2(1, 2) Then Exit Sub
    
    If Button.Index = 3 Then
        'Eliminar blabla bla
        If MsgBoxA("�Seguro que desea eliminar la linea de IVA seleccionada?", vbQuestion + vbYesNoCancel) = vbYes Then
            SQL = ObtenerWhereCab(True)
            SQL = Replace(SQL, "factpro.", "factpro_totales.")
            SQL = SQL & " AND numlinea = " & lw1.SelectedItem.Text
            SQL = "DELETE from factpro_totales " & SQL
            
            If Ejecuta(SQL, False) Then
                RecalcularTotalesFactura True
                LLamaLineas 2, 0
                PonerModo 2
                PosicionarData
                PonerCampos
            End If
        End If
    
    Else
        txtaux3(0).Tag = -1
        If Button.Index = 1 Then
            lw1.ListItems.Add , "N", ""
            Indc = lw1.ListItems.Count
            lw1.ListItems(Indc).EnsureVisible
            Set lw1.SelectedItem = lw1.ListItems(Indc)
            ModoLineas = 1
        Else
            Indc = lw1.SelectedItem.Index
            txtaux3(0).Tag = lw1.SelectedItem.Text
            ModoLineas = 2
        End If
        
        alto = lw1.ListItems(Indc).top + lw1.top + 45
        PonerModo 6
        LLamaLineas 2, IIf(Button.Index = 1, 1, 2), CSng(alto)
        Set lw1.SelectedItem = Nothing
        
        alto = 2
        If Button.Index = 1 Then
            alto = 0
        Else
            txtaux3_LostFocus 0
        End If
        
        PonFoco txtaux3(alto)
        
    End If
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
Dim SQL As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim SqlLog As String

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Ll�nia
    
    If Modo = 4 Then 'Modificar Cap�alera
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
            SQL = "�Seguro que desea eliminar la l�nea de la factura?"
            SQL = SQL & vbCrLf & "Serie: " & AdoAux(Index).Recordset!NUmSerie & " - " & AdoAux(Index).Recordset!Numregis & " - " & AdoAux(Index).Recordset!fecharec & " - " & AdoAux(Index).Recordset!NumLinea
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM factpro_lineas "
                SQL = SQL & Replace(vWhere, "factpro", "factpro_lineas") & " and numlinea = " & DBLet(AdoAux(Index).Recordset!NumLinea, "N")
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute SQL
        
        RecalcularTotales
        
        '**** parte de contabilizacion de la factura
        TerminaBloquear
        
        If Numasien2 > 0 Then
            If IntegrarFactura_(False) Then
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
        SqlLog = SqlLog & vbCrLf & "L�nea   : " & DBLet(Me.AdoAux(1).Recordset!NumLinea, "N")
        SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text
        
        
        vLog.Insertar 12, vUsu, SqlLog
        
        'Creo que no hace falta volver a situar el datagrid
        If True Then
            lblIndicador.Caption = data1.Recordset.AbsolutePosition & " de " & data1.Recordset.RecordCount
            
            data1.Refresh
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

    ModoLineas = 1 'Posem Modo Afegir Ll�nia

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    ' *** posar el nom del les distintes taules de ll�nies ***
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
Dim SQL As String

    ExisteEnFactura = False
    
    If Serie = "" Or NumFactu = "" Or FecFactu = "" Or Cuenta = "" Then Exit Function

    SQL = "select count(*) from factpro_lineas where numserie = " & DBSet(Serie, "T") & " and numregis = " & DBSet(NumFactu, "N")
    SQL = SQL & " and fecharec = " & DBSet(FecFactu, "F") & " and codmacta = " & DBSet(Cuenta, "T")

    ExisteEnFactura = (TotalRegistros(SQL) <> 0)
    
End Function


Private Function CuentaHabitual(CtaOrigen As String) As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    CuentaHabitual = ""
    
    SQL = "select codcontrhab from cuentas where codmacta = " & DBSet(CtaOrigen, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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


    ModoLineas = 2 'Modificar ll�nia

    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    Select Case Index
        Case 0, 1 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
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

    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 1 'lineas de factura
            For jj = 5 To txtaux.Count - 1
                txtaux(jj).visible = B
                txtaux(jj).top = alto
            Next jj
            
            txtAux2(5).visible = B
            txtAux2(5).top = alto
            txtAux2(12).visible = B
            txtAux2(12).top = alto
            
            
            chkAux(0).visible = B
            chkAux(0).top = alto
            
            For jj = 0 To 2
                cmdAux(jj).visible = B
                cmdAux(jj).top = txtaux(5).top
                cmdAux(jj).Height = txtaux(5).Height
            Next jj
            
            If Not vParam.autocoste Then
                cmdAux(2).visible = False
                cmdAux(2).Enabled = False
                txtaux(12).visible = False
                txtaux(12).Enabled = False
                txtAux2(12).visible = False
                txtAux2(12).Enabled = False
            End If
            If B Then
                'Aui es donde bloquamvamos los imprtes de IVA. Ahora NO los bloqueamos
                BloqueaTXT txtaux(10), Not B
                BloqueaTXT txtaux(11), Not B
            End If
    
    
    Case 2
            
            
            For jj = 0 To txtaux3.Count - 1
                txtaux3(jj).visible = B
                txtaux3(jj).top = alto
                    
                If B Then
                    txtaux3(jj).Left = lw1.ColumnHeaders(jj + 2).Left + lw1.Left + 15
                    txtaux3(jj).Width = lw1.ColumnHeaders(jj + 2).Width - 6

                    txtaux3(jj).Text = lw1.SelectedItem.SubItems(jj + 1)
                     
                     If jj = 0 Then BloqueaTXT txtaux3(jj), xModo = 2
                     
                End If
            Next jj
            If B Then
                cmdAux3.Left = txtaux3(1).Left - 90
                cmdAux3.top = alto + 30
            End If
            cmdAux3.visible = B And ModoLineas = 1
            If B Then
                BloqueaTXT txtaux3(1), True
                For jj = 1 To lw1.ListItems.Count
                    lw1.ListItems(jj).Selected = False
                Next
            Else
                txtaux3(0).Tag = ""
                txtaux3(2).Tag = ""
                txtaux3(3).Tag = ""
            End If
    
    
    
    End Select
End Sub



Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim B As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Importe As Currency

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False


    'Si  no tiene analitica, garaantizo el CCOST a vacio
    If Not vParam.autocoste Then txtaux(12).Text = ""
    
    




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
            MsgBox "La cuenta no es de �ltimo nivel", vbExclamation
            DatosOkLlin = False
            PonFoco txtaux(5)
            Exit Function
        End If
        
        If IvaCuenta = "" Then
            CambiarIva = True
        Else
            If ModoLineas = 1 Then
            
                If CInt(ComprobarCero(txtaux(7).Text)) <> CInt(ComprobarCero(IvaCuenta)) Then
                    If MsgBox("El c�digo de iva es distinto del de la cuenta. " & vbCrLf & " � Desea modificarlo en la cuenta ? " & vbCrLf & vbCrLf, vbQuestion + vbYesNo) = vbYes Then
                        CambiarIva = True
                    Else
                        CambiarIva = False
                    End If
                End If
            End If
        End If
        
        'Centro de coste
        If txtaux(12).visible Then
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
                MsgBox "C�digo de iva incorrecto. Debe ser Iva a 0%. Revise.", vbExclamation
                PonFoco txtaux(7)
                Exit Function
            End If
        End If
        
    End If
    
    
    
    
    
    'Como puede modificar los IVA, hay que comprobar
    If B Then
        
        Importe = ImporteFormateado(txtaux(8).Text) / 100
        Importe = ImporteFormateado(txtaux(6).Text) * Importe
        
        
        
        If Abs(Importe - ImporteFormateado(txtaux(10).Text)) >= 0.1 Then
            Mens = "Iva calculado: " & Format(Importe, FormatoImporte) & vbCrLf
            Mens = Mens & "Iva introducido: " & txtaux(10).Text & vbCrLf
            Mens = "DIFERENCIAS EN IVA" & vbCrLf & vbCrLf & Mens & vbCrLf & "�Desea continuar igualmente?"
            
            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
        End If
        
        If B Then
            If Me.txtaux(9).Text <> "" Then
                'REGARCO
                
                Importe = ImporteFormateado(txtaux(9).Text) / 100
                Importe = ImporteFormateado(txtaux(6).Text) * Importe
                If Abs(Importe - ImporteFormateado(txtaux(11).Text)) >= 0.05 Then
                    Mens = "Iva calculado: " & Format(Importe, FormatoImporte) & vbCrLf
                    Mens = Mens & "Iva introducido: " & txtaux(11).Text & vbCrLf
                    Mens = "DIFERENCIAS EN RECARGO EQUIVALENCIA" & vbCrLf & vbCrLf & Mens & vbCrLf & "�Desea continuar igualmente?"
                    
                    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
                End If
                
            End If
        End If
    End If
    
    
    If B And ModoLineas = 2 Then
        'Si cambia BASE, codigo o importe, hay que recalcular pago factura
        
        If Val(txtaux(7).Text) <> DBLet(AdoAux(1).Recordset!codigiva, "N") Then
            ModificarPagos = True
        ElseIf ImporteFormateado(txtaux(6).Text) <> DBLet(AdoAux(1).Recordset!Baseimpo, "N") Then
            ModificarPagos = True
        ElseIf ImporteFormateado(txtaux(10).Text) <> DBLet(AdoAux(1).Recordset!Impoiva, "N") Then
            ModificarPagos = True
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
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(5)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(5)|T|Denominaci�n|3995|;"
                tots = tots & "S|txtaux(6)|T|Importe|1905|;S|txtaux(7)|T|Iva|625|;S|cmdAux(1)|B|||;S|txtaux(8)|T|%Iva|765|;"
                tots = tots & "S|txtaux(9)|T|%Rec|765|;S|txtaux(10)|T|Importe Iva|1554|;S|txtaux(11)|T|Importe Rec|1554|;"
                tots = tots & "N||||0|;S|chkAux(0)|CB|Ret|400|;S|txtaux(12)|T|CC|710|;S|cmdAux(2)|B|||;S|txtAux2(12)|T|Nombre|2470|;"
            Else
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(5)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(5)|T|Denominaci�n|5695|;"
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
      
    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
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
'Inserta registre en les taules de Ll�nies
Dim nomframe As String
Dim B As Boolean
Dim Limp As Boolean
Dim Cad As String



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
            
            If B Then RecalcularTotalesFactura False
        
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
            
            B = BLOQUEADesdeFormulario2(Me, data1, 1)
            
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
Dim SQL As String

    On Error GoTo eActualizarIva
    
    ActualizarIva = False
    
    SQL = "update cuentas set codigiva = " & DBSet(txtaux(7).Text, "N") & " where codmacta = " & DBSet(txtaux(5).Text, "T")
    Conn.Execute SQL
    
    ActualizarIva = True
    Exit Function
    
eActualizarIva:
    MuestraError Err.Number, "Actualizar Iva", Err.Description
End Function


Private Function ModificarLinea2() As Boolean
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim v As Integer
Dim Cad As String
Dim SqlLog As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'apuntes
    End Select
    ' **************************************************************
    ModificarLinea2 = False
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
                v = AdoAux(NumTabMto).Recordset.Fields(3) 'el 2 es el n� de llinia
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
            
            ModificarLinea2 = True
        Else
            Conn.RollbackTrans
        End If
    End If
        
End Function




Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & "factpro.numserie=" & DBSet(Text1(2).Text, "T") & " and factpro.numregis=" & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripci� ***
Private Sub LimpiarCamposFrame(Index As Integer)
End Sub
' ***********************************************

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim B As Boolean
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N") And (Modo <> 0 And Modo <> 5)
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!especial, "N") And Modo = 2 And vEmpresa.TieneTesoreria
        
        Me.Toolbar2.Buttons(3).Enabled = DBLet(Rs!especial, "N") And (Modo = 2 Or Modo = 0)
        Me.Toolbar2.Buttons(4).Enabled = DBLet(Rs!especial, "N") And (Modo = 2 Or Modo = 0)
        
        B = False
        If vParam.PathAlbaranesProv <> "" Then B = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Me.Toolbar2.Buttons(6).Enabled = B
        Me.Toolbar2.Buttons(7).Enabled = B
        
        
        
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2)
        ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2 And Me.data1.Recordset.RecordCount > 0)
        ToolbarAux.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.data1.Recordset.RecordCount > 0)
        
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
    Dim CalcularElIva As Boolean
    
        If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
        
        If txtaux(Index).Text = AntiguoText1 Then
             If Index = 12 And vParam.autocoste Then PonleFoco cmdAceptar
             Exit Sub
        End If
    
        CalcularElIva = True
        Select Case Index
        Case 5
            RC = txtaux(5).Text
            If RC = "" Then
                txtaux(5).Text = ""
                txtAux2(5).Text = ""
                Exit Sub
            End If
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(5).Text = RC
                If Modo = 1 Then Exit Sub
                If EstaLaCuentaBloqueada(RC, CDate(Text1(1).Text)) Then
                    MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                    txtaux(5).Text = ""
                Else
                    txtAux2(5).Text = SQL
                    ' traemos el tipo de iva de la cuenta
                    If ModoLineas = 1 Then
                        txtaux(7).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtaux(5).Text, "N")
                        IvaCuenta = txtaux(7)
                        If txtaux(7).Text <> "" Then txtAux_LostFocus (7)
                    Else
                        CalcularElIva = False
                    End If
                    RC = ""
                    
                End If
            Else
                If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                    txtaux(5).Text = RC
                    txtAux2(5).Text = ""
                    'NO EXISTE LA CUENTA, a�ado que debe de tener permiso de creacion de cuentas
                    If vUsu.PermiteOpcion("ariconta", 201, vbOpcionCrearEliminar) Then
                        SQL = SQL & " �Desea crearla?"
                        If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                            CadenaDesdeOtroForm = RC
                            cmdAux(0).Tag = Index
                            Set frmC = New frmColCtas
                            frmC.DatosADevolverBusqueda = "0|1|"
                            frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                            frmC.Show vbModal
                            Set frmC = Nothing
                            If txtaux(5).Text = RC Then
                                SQL = "" 'Para k no los borre
                                ' traemos el tipo de iva de la cuenta
                                txtaux(7).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtaux(5).Text, "N")
                                IvaCuenta = txtaux(7)
                                txtAux_LostFocus (7)
                            
                            End If
                        End If
                    Else
                        MsgBox SQL, vbExclamation
                    End If
                Else
                    MsgBox SQL, vbExclamation
                End If
                    
                If SQL <> "" Then
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
            If Not PonerFormatoDecimal(txtaux(Index), 1) Then
                txtaux(Index).Text = ""
            Else
                'Si modificando lienas, no cambia el importe NO recalculo iVA
                If Modo = 5 And ModoLineas = 2 Then
                    If ImporteFormateado(txtaux(Index).Text) = CCur(DBLet(AdoAux(1).Recordset!Baseimpo, "N")) Then CalcularElIva = False
                    
                End If
            End If
            
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
                
             If Modo = 5 And ModoLineas = 2 Then
                If txtaux(7).Text <> "" Then
                    If Val(txtaux(Index).Text) = Val(DBLet(AdoAux(1).Recordset!codigiva, "N")) Then CalcularElIva = False
                End If
            End If
                
        Case 10, 11
           'LOS IMPORTES
            If PonerFormatoDecimal(txtaux(Index), 1) Then
                If Not vParam.autocoste Then PonleFoco cmdAceptar
            End If
                
        Case 12
            txtaux(12).Text = UCase(txtaux(12).Text)
            SQL = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtaux(12).Text, "T")
            txtAux2(12).Text = ""
            If SQL = "" Then
                MsgBox "Concepto NO encontrado: " & txtaux(12).Text, vbExclamation
                txtaux(12).Text = ""
            Else
                txtAux2(12).Text = SQL
            End If
            
            PonleFoco cmdAceptar
        End Select

        If CalcularElIva Then
            If Index = 5 Or Index = 6 Or Index = 7 Then CalcularIVA
        End If
        Screen.MousePointer = vbDefault
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
Dim TotalDesdeLineas As Boolean

    On Error GoTo EModificar
         
        ModificarFactura = False
     
                    
        Conn.BeginTrans
        'Comun
        
        ActualizarRetencionLineasSiNecesario
        
        
        TotalDesdeLineas = False
        If Val(DBLet(data1.Recordset!no_modifica_apunte, "N")) Then TotalDesdeLineas = True
        
        B = RecalcularTotalesFactura(TotalDesdeLineas)
        
        
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
Dim SQL As String

    On Error GoTo eRecalcularTotalesFactura

    ModificaLineas = False

    
    SQL = "update factpro_lineas set fecharec = " & DBSet(Text1(1).Text, "F")
    SQL = SQL & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numregis= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    Conn.Execute SQL
    
    SQL = "update factpro_totales set fecharec = " & DBSet(Text1(1).Text, "F")
    SQL = SQL & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numregis= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    Conn.Execute SQL
    
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
        AuxOK = "Cuenta debe ser num�rica"
        Exit Function
    End If
    
    If txtAux2(4).Text = NO Then
        AuxOK = "La cuenta debe estar dada de alta en el sistema"
        Exit Function
    End If
    
    If Not EsCuentaUltimoNivel(txtaux(4).Text) Then
        AuxOK = "La cuenta no es de �ltimo nivel"
        Exit Function
    End If
    
        
    'Codigo de iva
    If txtaux(4).Text = "" Then
        AuxOK = "El c�digo de iva no puede estar vacio"
        Exit Function
    End If
        
    If txtaux(7).Text <> "" Then
        If Not IsNumeric(txtaux(7).Text) Then
            AuxOK = "El c�digo de iva debe de ser num�rico."
            Exit Function
        End If
    End If
    
    'Importe
    If txtaux(6).Text <> "" Then
        If Not EsNumerico(txtaux(6).Text) Then
            AuxOK = "El importe DEBE debe ser num�rico"
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
Dim Cad As String
Dim RT As ADODB.Recordset
        Cad = " WHERE numregis=" & Text1(0).Text
        Cad = Cad & " and numserie = " & DBSet(Text1(2).Text, "T")
        
        If Actual Then
            i = 0
        Else
            i = 1
        End If
        
        Cad = Cad & " AND anofactu =" & DBSet(Text1(14).Text, "N")
        
        Set RT = New ADODB.Recordset
        ComprobarNumeroFactura = True
        i = 0
        RT.Open "Select numregis from factpro" & Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.EOF) Then
                ComprobarNumeroFactura = False
            End If
        End If
        RT.Close
        If ComprobarNumeroFactura Then
            i = 1
            RT.Open "Select numregis from factpro" & Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RT.EOF Then
                If Not IsNull(RT.EOF) Then
                    ComprobarNumeroFactura = False
                End If
            End If
            RT.Close
        End If
        Set RT = Nothing
        If Not ComprobarNumeroFactura Then
            Cad = "Verifique los contadores. Ya existe la factura " & Text1(0).Text & vbCrLf
            MsgBox Cad, vbExclamation
        End If
End Function

Private Function SituarData1(Insertar As Boolean) As Boolean
    Dim SQL As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        SQL = "Select * from factpro WHERE numserie =" & DBSet(Text1(2).Text, "T")
        SQL = SQL & " AND fecharec=" & DBSet(Text1(1).Text, "F") & " AND numregis = " & Text1(0).Text
        data1.RecordSource = SQL
    'End If
    
    data1.Refresh
    With data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not data1.Recordset.EOF
            If CStr(.Fields!NUmSerie) = Text1(2).Text Then
                If CStr(.Fields!Numregis) = Text1(0).Text Then
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
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim IT As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String
Dim C As String


    On Error GoTo ECargaDatosLW
    
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 5 ' imagenes
        Cad = "select h.numlinea,  h.codigiva, tt.nombriva,  h.baseimpo, h.impoiva, h.imporec from factpro_totales h inner join tiposiva tt on h.codigiva = tt.codigiva  WHERE "
        Cad = Cad & " numserie=" & DBSet(data1.Recordset!NUmSerie, "T")
        Cad = Cad & " and numregis=" & DBSet(data1.Recordset!Numregis, "N")
        Cad = Cad & " and fecharec=" & DBSet(data1.Recordset!fecharec, "F")
        Cad = Cad & " and anofactu=" & data1.Recordset!anofactu
        GroupBy = ""
        BuscaChekc = "numlinea"
        
    End Select
    
    
    Cad = Cad & " ORDER BY 1"
    
    lw1.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
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




'
'Private Sub AnyadirAlListview(vpaz As String, DesdeBD As Boolean)
'Dim J As Integer
'Dim Aux As String
'Dim IT As ListItem
'Dim Contador As Integer
'    If Dir(vpaz, vbArchive) = "" Then
''        MsgBox "No existe el archivo: " & vpaz, vbExclamation
'    Else
'        Set IT = lw1.ListItems.Add()
'
'        IT.Text = Me.Adodc1.Recordset!Orden '"Nuevo " & Contador
'
'        IT.SubItems(1) = Me.Adodc1.Recordset.Fields(5)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
'        IT.SubItems(2) = vpaz
'        IT.SubItems(3) = Me.Adodc1.Recordset.Fields(0)
'
'        Set IT = Nothing
'    End If
'End Sub
'
'


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
    



Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim SQL As String

    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i

    'Tipo de factura
    Set Rs = New ADODB.Recordset
    SQL = "SELECT * FROM usuarios.wconce340 ORDER BY codigo"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    SQL = "SELECT * FROM usuarios.wtipopera ORDER BY codigo"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Combo1(1).AddItem Rs!denominacion
        Combo1(1).ItemData(Combo1(1).NewIndex) = Rs!Codigo
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Tipo de retencion
    Set Rs = New ADODB.Recordset
    SQL = "SELECT * FROM usuarios.wtiporeten ORDER BY codigo"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Combo1(2).AddItem Rs!Descripcion
        Combo1(2).ItemData(Combo1(2).NewIndex) = Rs!Codigo
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Tipo de intracomunitaria
    Set Rs = New ADODB.Recordset
    SQL = "SELECT * FROM usuarios.wtipointra ORDER BY codintra"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        Combo1(3).AddItem Rs!nomintra
        Combo1(3).ItemData(Combo1(3).NewIndex) = Asc(Rs!Codintra)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    
    SQL = "SELECT * FROM usuarios.wtipoinmueble ORDER BY codigo"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        Combo1(4).AddItem Rs!Descripcion
        Combo1(4).ItemData(Combo1(4).NewIndex) = Asc(Rs!Codigo)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
    
    Set Rs = Nothing

End Sub

'Tipo Verificacion
'  0. periodo IVA y SII
'  1- Solo periodo
'  2- Solo SII
Private Function ComprobarPeriodo2(Indice As Integer, TipoVerificacion As Byte) As Boolean
Dim Cerrado As Boolean
Dim MensajeSII As String
Dim Mostrar As Boolean
Dim ModEspecial As Boolean


  

    MensajeSII = ""
    If TipoVerificacion <> 1 Then
        If vParam.SIITiene Then
            'SI esta presentada...
            If Modo <> 3 And Modo <> 1 Then
                If DBLet(data1.Recordset!sii_id, "N") > 0 Then
                    'If Val(DBLet(data1.Recordset!sii _status, "N")) > 2 Then
                    If Text1(28).BackColor = &HC0FFC0 Or Text1(28).BackColor = &H80FF& Then
                    
                        'Si fecha >= fechaini
                        ModEspecial = False
                        If vUsu.Nivel <= 1 Then
                            If data1.Recordset!fecharec >= vParam.fechaini Then
                            
                                If Val(DBLet(data1.Recordset!no_modifica_apunte, "N")) = 0 Then ModEspecial = True
                            End If
                        End If
                        
                        If ModEspecial Then
                        
                            'Bloqueamos el registro
                            CadenaDesdeOtroForm = ""
                            Ampliacion = ""
                            Conn.Execute "DELETE from tmpfaclin WHERE codusu = " & vUsu.Codigo
                            With frmFacturaModificar
                                .Cliente = False
                                .Anyo = data1.Recordset!anofactu
                                .Codigo = data1.Recordset!Numregis
                                .NUmSerie = data1.Recordset!NUmSerie
                                .Fecha = data1.Recordset!fecharec
                                .Show vbModal
                            End With
                            
                            
                            'Si que ha modificado
                            Screen.MousePointer = vbHourglass
                            If CadenaDesdeOtroForm <> "" Or Ampliacion <> "" Then
                                
                                If ModificaFacturaSiiPresentada Then
                                    'CargaGrid 1, True
                                    PosicionarData
                                    PonerCampos
                                End If
                            End If
                            Screen.MousePointer = vbDefault
                        Else
                            MsgBox "La factura ya esta presentada en el sistema de SII de la AEAT.", vbExclamation
                            
                        End If
                        Exit Function
                    End If
                End If
            End If
            
            
            If Modo > 2 Then
                If UltimaFechaCorrectaSII(vParam.SIIDiasAviso, Now) > CDate(Text1(Indice).Text) Then
                    MensajeSII = ""  'String(70, "*") & vbCrLf
                    MensajeSII = MensajeSII & "SII." & vbCrLf & vbCrLf & "Excede del maximo dias permitido para comunicar la factura" & vbCrLf & MensajeSII
                End If
            End If
        
        End If
    End If




    'Primero pondremos la fecha a a�o periodo
    Cerrado = False
    If TipoVerificacion <> 2 Then
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
                'El mismo a�o. Comprobamos los periodos
                If vParam.perfactu >= Ancho Then _
                    Cerrado = True
            End If
        End If
        
    End If
    
    ComprobarPeriodo2 = True
    ModificaFacturaPeriodoLiquidado = False
    
    Mostrar = Cerrado
    If Not Cerrado Then
        If MensajeSII <> "" Then Mostrar = True
    End If
    
    If Mostrar Then
        ModificaFacturaPeriodoLiquidado = True
        If Cerrado Then
            SQL = "La fecha "
            If Indice = 0 Then
                SQL = SQL & "factura"
            Else
                SQL = SQL & "liquidacion"
            End If
            SQL = SQL & " corresponde a un periodo ya liquidado. " & vbCrLf
        Else
            SQL = ""
        End If
        If MensajeSII <> "" Then MensajeSII = MensajeSII & vbCrLf & vbCrLf
        SQL = MensajeSII & SQL
        
        
        If vUsu.Nivel = 0 Then
            SQL = SQL & vbCrLf & " �Desea continuar igualmente ?"
  
            If MsgBoxA(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then ComprobarPeriodo2 = False
        Else
        
            MsgBoxA SQL, vbExclamation
            
            ComprobarPeriodo2 = False
        
        End If
    
       
    
    End If
    
End Function


Private Sub CargarDatosCuenta(Cuenta As String)
Dim Rs As ADODB.Recordset
Dim PonDatos As Boolean
Dim SQL As String

    On Error GoTo eTraerDatosCuenta
    
    SQL = "select * from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    PonDatos = False
    If Modo = 3 Then
        PonDatos = True
    Else
        If DBLet(data1.Recordset!codmacta, "T") <> Rs!codmacta Then PonDatos = True
    End If
    
    If PonDatos Then
        Text1(5).Text = ""
        Text4(5).Text = ""
        
        For i = 15 To 21
            Text1(i).Text = ""
        Next i
    End If
        
    If Not Rs.EOF Then
        
        
        If PonDatos Then
            If Not IsNull(Rs!Forpa) Then
                Text1(5).Text = DBLet(Rs!Forpa, "N")
                Text4(5).Text = PonerNombreDeCod(Text1(5), "formapago", "nomforpa", "codforpa", "N")
            End If
        End If
        
        If Text1(15).Text = "" Then Text1(15).Text = DBLet(Rs!razosoci, "T")
        If Text1(16).Text = "" Then Text1(16).Text = DBLet(Rs!dirdatos, "T")
        If Text1(17).Text = "" Then Text1(17).Text = DBLet(Rs!codposta, "T")
        If Text1(18).Text = "" Then Text1(18).Text = DBLet(Rs!desPobla, "T")
        If Text1(19).Text = "" Then Text1(19).Text = DBLet(Rs!desProvi, "T")
        If Text1(20).Text = "" Then Text1(20).Text = DBLet(Rs!nifdatos, "T")
        If Text1(21).Text = "" Then
            Text1(21).Text = DBLet(Rs!codpais, "T")
            Text4(21).Text = PonerNombreDeCod(Text1(21), "paises", "nompais", "codpais", "T")
        End If
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
        txtaux(10).Text = Format(Round2((Aux * Base), 2), FormatoImporte)
    End If
    
    'Recargo
    Aux = ImporteFormateado(txtaux(9).Text) / 100
    If Aux = 0 Then
        txtaux(11).Text = ""
    Else
        txtaux(11).Text = Format(Round2((Aux * Base), 2), FormatoImporte)
    End If

End Sub

Private Function RecalcularTotales() As Boolean
Dim SQL As String
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

    SQL = "delete from factpro_totales where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    Conn.Execute SQL
    
    SqlInsert = "insert into factpro_totales (numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec) values "
    
    SQL = "select factpro_lineas.codigiva, porciva, porcrec, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec from factpro_lineas "
    SQL = SQL & " left join tiposiva on factpro_lineas.codigiva=tiposiva.codigiva "
    SQL = SQL & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    SQL = SQL & " group by 1,2,3"
    SQL = SQL & " order by 1,2,3"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    
    SqlValues = ""
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    ImpoRec = 0
    Imporeten = 0
    TotalFactura = 0
    
    While Not Rs.EOF
        SQL = "(" & DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & "," & DBSet(Text1(14).Text, "N") & ","
        SQL = SQL & DBSet(i, "N") & "," & DBSet(Rs!Baseimpo, "N") & "," & DBSet(Rs!codigiva, "N") & "," & DBSet(Rs!porciva, "N") & "," & DBSet(Rs!porcrec, "N") & ","
        SQL = SQL & DBSet(Rs!Imporiva, "N") & "," & DBSet(Rs!imporrec, "N") & "),"
        
        SqlValues = SqlValues & SQL
        
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
    
    
    RecalcularTotales = RecalcularTotalesFactura(False)
    Exit Function
    
eRecalcularTotales:
    MuestraError Err.Number, "Recalcular Totales", Err.Description
End Function

'DesdeA�adirLineaDeTotalesIVA.  Integraciones SAGE
Private Function RecalcularTotalesFactura(DesdeA�adirLineaDeTotalesIVA As Boolean) As Boolean
Dim SQL As String
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
Dim Suplidos As Currency
Dim TipoRetencion As Integer
Dim IvaModificable As Boolean

    On Error GoTo eRecalcularTotalesFactura

    RecalcularTotalesFactura = False

    TipoRetencion = DevuelveValor("select tipo from usuarios.wtiporeten where codigo = " & DBSet(Combo1(2).ListIndex, "N"))
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    Imporeten = 0
    ImpoRec = 0
    TotalFactura = 0
    Suplidos = 0
    
    SQL = ""
    SqlInsert = "factpro_lineas"
    If DesdeA�adirLineaDeTotalesIVA Then
        SQL = "0"
        SqlInsert = "factpro_totales"
    End If
    'Select
    SQL = "select " & SQL & " aplicret, tipodiva, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec"
    
    
    
    'SQL = "select aplicret,tipodiva, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec from factpro_lineas "
    SQL = SQL & " FROM ###### "
    SQL = SQL & " left join tiposiva on ######.codigiva=tiposiva.codigiva "

    SQL = SQL & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    SQL = SQL & " group by 1,2 order by 1"
    
    SQL = Replace(SQL, "######", SqlInsert)
    SqlInsert = ""
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        If Rs!TipoDIva = 4 Then
            'SUPLIDO
            Suplidos = Suplidos + Rs!Baseimpo
        Else
            Baseimpo = Baseimpo + DBLet(Rs!Baseimpo, "N")
        End If
        

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
        Imporeten = Round2((PorcRet * Basereten / 100), 2)
    End If
    
    TotalFactura = Baseimpo + Impoiva + ImpoRec + Suplidos - Imporeten
    
    Text1(9).Text = Format(Baseimpo, FormatoImporte)
    Text1(11).Text = Format(Basereten, FormatoImporte)
    Text1(10).Text = Format(Impoiva, FormatoImporte)
    Text1(12).Text = Format(Imporeten, FormatoImporte)
    Text1(13).Text = Format(TotalFactura, FormatoImporte)
    
    If ImpoRec <> 0 Then Text1(30).Text = Format(ImpoRec, FormatoImporte)
    If Suplidos <> 0 Then Text1(31).Text = Format(Suplidos, FormatoImporte)
    
    If PorcRet = 0 Then
        Text1(11).Text = ""
        Text1(12).Text = ""
    End If
    
    SQL = "update factpro set "
    SQL = SQL & " totbases = " & DBSet(Baseimpo, "N")
    SQL = SQL & ", totivas = " & DBSet(Impoiva, "N")
    SQL = SQL & ", totrecargo = " & DBSet(ImpoRec, "N")
    SQL = SQL & ", totfacpr = " & DBSet(TotalFactura, "N")
    SQL = SQL & ", totbasesret = " & DBSet(Basereten, "N", "S")
    SQL = SQL & ", trefacpr = " & DBSet(Imporeten, "N", "S")
    SQL = SQL & ", suplidos = " & DBSet(Suplidos, "N", "S")
    
    SQL = SQL & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numregis= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    Conn.Execute SQL
    
    
    
     'OCTUB 2017
    'Si ha ha cambiado la fecha factura updateamos
    If Modo = 4 Then
        If Text1(1).Text <> Format(data1.Recordset!fecharec, "dd/mm/yyyy") Then
             SQL = "UPDATE factpro_totales set fecharec = " & DBSet(Text1(1).Text, "F")
             SQL = SQL & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numregis= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
             Ejecuta SQL
        End If
    End If
    
    
    
    
    RecalcularTotalesFactura = True
    Exit Function
    
eRecalcularTotalesFactura:
    MuestraError Err.Number, "Recalcular Totales Factura", Err.Description
End Function


Private Function IntegrarFactura_(DentroTrans As Boolean) As Boolean

    If Not ContabilizaApunte Then
        ContabilizaApunte = True 'Lo dejo por defecto otra vez
        IntegrarFactura_ = True
        Exit Function
    End If
    

    IntegrarFactura_ = False
    
    With frmActualizar
        .OpcionActualizar = 8
        'NumAsiento     --> CODIGO FACTURA
        'NumDiari       --> A�O FACTURA
        'NUmSerie       --> SERIE DE LA FACTURA
        'FechaAsiento   --> Fecha factura
        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
        .NumFac = CLng(Text1(0).Text)
        .NumDiari = CInt(Text1(14).Text)
        .NUmSerie = Text1(2).Text
        .FechaAsiento = Text1(1).Text
        .FechaAnterior = FecRecepAnt 'FecFactuAnt
        .DentroBeginTrans = DentroTrans
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
        If AlgunAsientoActualizado Then
            IntegrarFactura_ = True
            'NOVIEMBRE 18
            'PUEDE HABER CAMBIADO LA FECHA
            If Not IsNull(data1.Recordset!FechaEnt) Then
                If Format(data1.Recordset!FechaEnt) <> Text1(1).Text Then
                    Ejecuta "UPDATE factpro set fechaent=" & DBSet(Text1(1).Text, "F") & ObtenerWhereCab(True), False
                    
                
                End If
            End If
        End If
        Screen.MousePointer = vbHourglass
        Me.Refresh
    End With
    

End Function


Private Function Desintegrar() As Boolean
        If Not ContabilizaApunte Then
            Desintegrar = True
            Exit Function
        End If


        Desintegrar = False
        'Primero hay que desvincular la factura de la tabla de hco
        If DesvincularFactura Then
            frmActualizar.OpcionActualizar = 2  'Desactualizar para eliminar
            frmActualizar.NumAsiento = data1.Recordset!NumAsien
            frmActualizar.FechaAsiento = FecRecepAnt 'FecFactuAnt
            frmActualizar.NumDiari = data1.Recordset!NumDiari
            frmActualizar.FechaAnterior = data1.Recordset!FechaEnt
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            If AlgunAsientoActualizado Then Desintegrar = True
        End If
End Function


Private Function DesvincularFactura() As Boolean
On Error Resume Next
    SQL = "UPDATE factpro set numasien=NULL, fechaent=NULL, numdiari=NULL"
    SQL = SQL & " WHERE numregis = " & data1.Recordset!Numregis
    SQL = SQL & " AND numserie = '" & data1.Recordset!NUmSerie & "'"
    SQL = SQL & " AND anofactu =" & data1.Recordset!anofactu
    Numasien2 = data1.Recordset!NumAsien
    NumDiario = data1.Recordset!NumDiari
    Conn.Execute SQL
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
    If data1.Recordset.RecordCount > 0 Then TieneRegistros = True
End Function


Public Function HayQueContabilizarDesdePantallaPagos() As Boolean
    HayQueContabilizarDesdePantallaPagos = False
    If Pagado = 1 Then
        HayQueContabilizarDesdePantallaPagos = True
    
    End If
End Function


Public Function ContabilizarPagos() As Boolean
Dim Mc As Contadores
Dim FP As Ctipoformapago
Dim SQL As String
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
Dim Cad As String
Dim Sql4 As String
Dim fecefect As Date
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
    SQL = " from pagos where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(25).Text, "T")
    SQL = SQL & " and codmacta = " & DBSet(Text1(4).Text, "T")
    SQL = SQL & " and fecfactu = " & DBSet(Text1(26).Text, "F")
    SQL = SQL & " order by numorden"
    
    TotImpo = DevuelveValor("select sum(coalesce(impefect,0)) " & SQL)
    
    SQL = Sql1 & SQL
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Inserto cabecera de apunte
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, feccreacion, usucreacion, desdeaplicacion, obsdiari) VALUES ("
    SQL = SQL & FP.diaricli
    SQL = SQL & ",'" & Format(FechaPago, FormatoFecha) & "'," & Mc.Contador & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizaci�n Pago Facturas Proveedor',"
    Sql1 = "Generado desde Facturas de Proveedor el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre
    If TotImpo < 0 Then Sql1 = Sql1 & "  (CARGO)"
    Conn.Execute SQL & DBSet(Sql1, "T") & ")"
    
    fecefect = CDate("2100/01/01") 'Para que coja el primero
    Linea = 0
    While Not Rs.EOF
        
        Linea = Linea + 1
        
        
        'Fecefc para el contapunte del banco(si ha elegido ese concepto)
        If Rs!fecefect < fecefect Then fecefect = Rs!fecefect
        
        'importe
        impo = ImporteFormateado(DBLet(Rs!imppagad))
        
        'Inserto en las lineas de apuntes
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,numserie,numfacpr,fecfactu,numorden,tipforpa ) VALUES ("
        SQL = SQL & FP.diaricli
        SQL = SQL & ",'" & Format(FechaPago, FormatoFecha) & "'," & Mc.Contador & ","
        
        
        'numdocum
        Numdocum = ""
        If Text1(2).Text <> "1" Then Numdocum = Text1(2).Text & "-"
        Numdocum = Numdocum & Text1(25).Text  ' letra de serie y factura
        
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
        Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
        Ampliacion = Cad & " " & Ampliacion
        Ampliacion = Mid(Ampliacion, 1, 35)
        
        'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
        'Cuenta Cliente/proveedor
        Cad = Linea & ",'" & Trim(Text1(4).Text) & "'," & DBSet(Numdocum, "T") & "," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
        'Importe cobro-pago
        ' nos lo dire "debe"
        If Not Debe Then
            Cad = Cad & "NULL," & TransformaComasPuntos(CStr(impo))
        Else
            Cad = Cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
        End If
        'Codccost
        Cad = Cad & ",NULL,"
        If LlevaContr Then
            Cad = Cad & "'" & CtaBanco & "'"
        Else
            Cad = Cad & "NULL"
        End If
        Cad = Cad & ",'PAGOS',0," & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "T") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & "," & TipForpa & ")"
        Cad = SQL & Cad
        Conn.Execute Cad
        
        
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
    If Not vParam.abononeg Then TotImpo = Abs(TotImpo) 'El lado(D/H) ya lo ha configurado arriba
           
    If Conce = 2 Then
           Ampliacion = Ampliacion & "Fec.Vto: " & Format(fecefect, "dd/mm/yyyy") 'Fecha efecto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(Text1(2).Text)
    Else
        
       If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
        Ampliacion = Ampliacion & Text1(2).Text & "/" & Text1(25).Text
    End If
    
    
    Conce = ElConcepto
    Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = Cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    Cad = Linea & "," & DBSet(CtaBanco, "T") & ",'" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
    'Importe cliente
    'Si el cobro/pago va al debe el contrapunte ira al haber
    If Not Debe Then
        'al debe
        Cad = Cad & TransformaComasPuntos(CStr(TotImpo)) & ",NULL"
    Else
        'al haber
        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(TotImpo))
    End If
    
    'Codccost
    Cad = Cad & ",NULL,"
    
    If LlevaContr Then
        Cad = Cad & "'" & Trim(Text1(4).Text) & "'"
    Else
        Cad = Cad & "NULL"
    End If
    Cad = Cad & ",'PAGOS',0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
    Cad = SQL & Cad
    Conn.Execute Cad
    
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
Dim SQL As String

    SQL = "select estraspasada from factpro where numserie = " & DBSet(Text1(2).Text, "T") & " and numregis = "
    SQL = SQL & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    EsFraProTraspasada = (DevuelveValor(SQL) = 1)
    

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
  
  
        SQL = DevuelveDesdeBD("codigo", "paramamort", "1", "1")
        If Trim(SQL) = "" Then Exit Sub
        
    
        
        
        If CrearEltoInmov Then
            SQL = "Desea crear un elemento de Inmovilizado ? "
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                'Le pasaremos el codprove|nomprove|numfac|importe
                'ANTES
                'codprove    nombre    numfac     fecha adq     importe     Cuenta    Des. cuenta
                CadenaDesdeOtroForm = Text1(4).Text & "|" & Me.Text4(4).Text & "|"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(25).Text & "|" & Text1(26).Text & "||"
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
    If Combo1(2).ListIndex > 0 And Val(DBLet(Me.data1.Recordset!tiporeten, "T")) = 0 Then
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


Private Sub Color_CampoSII()
Dim Color As Byte
Dim Aux As String

    If DBLet(data1.Recordset!sii_id, "N") = 0 Then
        Color = 0
    Else
        Aux = "concat(enviada,'|',coalesce(csv,''),'|',coalesce(resultado,''))"
        Aux = DevuelveDesdeBD(Aux, "aswsii.envio_facturas_recibidas", "IDEnvioFacturasRecibidas", CStr(data1.Recordset!sii_id))
        If Aux = "" Then
            Color = 2
        Else
            If RecuperaValor(Aux, 1) = 1 Then
                If RecuperaValor(Aux, 2) = "" Then
                    Color = 2
                Else
                    Color = 4
                End If
            Else
                Color = 3
            End If
        End If
    End If
    
    If Color = 0 Then
        Text1(28).BackColor = vbWhite
        Text1(28).ForeColor = vbBlack
        Text1(28).FontBold = False
    Else
        Text1(28).FontBold = True
        Text1(28).ForeColor = vbBlack
        If Color = 4 Then
            'OK
            Text1(28).BackColor = &HC0FFC0
        ElseIf Color = 3 Then
            Text1(28).BackColor = &H80FF&
        Else
            Text1(28).BackColor = &HFF&
        End If
    End If
End Sub



Private Function ModificaFacturaSiiPresentada() As Boolean
Dim C As String
On Error GoTo eModificaDesdeFormAux
    ModificaFacturaSiiPresentada = False
        
    Conn.BeginTrans
        
        
    'Borramos de linfact
    '
    If CadenaDesdeOtroForm <> "" Then
        C = ObtenerWhereCP(True)
        Conn.Execute "DELETE FROM factpro_lineas " & C
            
        
        'insertamos  dedesde tmpfaclin
        'factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost)
        C = "INSERT INTO factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost) VALUES "
        C = C & CadenaDesdeOtroForm
        Conn.Execute C
        
    End If
    
    If Ampliacion <> "" Then
        C = Trim(Mid(Ampliacion, 1, 10))
        C = "UPDATE factpro SET cuereten = " & DBSet(C, "T", "S")
        Ampliacion = Mid(Ampliacion, 11)
        C = C & " , observa = " & DBSet(Ampliacion, "T", "S")
        C = C & " WHERE numregis= " & data1.Recordset!Numregis & " AND numserie =" & DBSet(data1.Recordset!NUmSerie, "T") & " AND anofactu =" & data1.Recordset!anofactu
        Conn.Execute C
    End If
        
    'Borramos lineas apuntes
    C = Val(DBSet(data1.Recordset!no_modifica_apunte, "N"))
    If Val(C) = 0 Then
        
        Numasien2 = data1.Recordset!NumAsien
        NumDiario = data1.Recordset!NumDiari
        FecRecepAnt = Text1(1).Text
        If Numasien2 > 0 Then
            C = " WHERE (numasien=" & Numasien2 & " and fechaent = " & DBSet(FecRecepAnt, "F") & " and numdiari = " & DBSet(NumDiario, "N") & ") "
            Conn.Execute "DELETE FROM hlinapu " & C
            
            IntegrarFactura_ (True)
            
    
        End If
    End If
    
    'Si llega aqui. Todo bien
    Conn.CommitTrans
    ModificaFacturaSiiPresentada = True
    
    
    
    Exit Function
eModificaDesdeFormAux:
    MuestraError Err.Number, Err.Description
    Conn.RollbackTrans
End Function


Private Sub ReferenciaCatastral(visible As Boolean)
    Text1(29).visible = visible
    Combo1(4).visible = visible
    Label1(22).visible = visible
    Label1(23).visible = visible
End Sub






'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'
'               Adjunto PDF
'
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------



Private Sub TieneDocumentoAsociado()
    SQL = ObtenerWhereCP(False) & " AND 1"
    SQL = DevuelveDesdeBD("docum", "factpro_fichdocs", SQL, "1")
    txtPDF.Text = Trim(SQL)
End Sub

Private Sub txtPDF_DblClick()

    On Error GoTo etxtPDF_DblClick

    If Modo <> 2 Then Exit Sub
    If txtPDF.Text = "" Then Exit Sub
    


    If Dir(App.Path & "\Temp", vbDirectory) = "" Then MkDir App.Path & "\Temp"
    
    SQL = ObtenerWhereCP(True) & " AND orden =1"
    SQL = "Select campo,docum from factpro_fichdocs " & SQL
    
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = SQL
    Adodc1.Refresh

    If Adodc1.Recordset.EOF Then
        'NO HAY NINGUNA
        MsgBoxA "Ningun documento asociado a la factura", vbExclamation
    Else
        'LEEMOS LA IMAGEN

        
        SQL = App.Path & "\Temp\" & Adodc1.Recordset!DOCUM
        LeerBinary Adodc1.Recordset!Campo, SQL
        Adodc1.RecordSource = "Select codigo from factpro_fichdocs WHERE false "
        Adodc1.Refresh
        
        
        Call ShellExecute(Me.hwnd, "Open", SQL, "", "", 1)
        
    End If


    Exit Sub
etxtPDF_DblClick:
    MuestraError Err.Number, Err.Description
End Sub







Private Function InsertarDesdeFichero() As Boolean
Dim CADENA As String
Dim Carpeta As String
Dim Aux As String
Dim J As Integer
Dim C As String
Dim Rs As ADODB.Recordset
Dim L As Long
Dim Fichero  As String


    InsertarDesdeFichero = False

    Fichero = ""
    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    cd1.MaxFileSize = 1024 * 30
    cd1.Filter = "Archivos PDF|*.pdf|Archivos Jpg|*.jpg"
    cd1.ShowOpen
    cd1.MaxFileSize = 256
    cd1.CancelError = False
    
    If cd1.FileName = "" Then
        InsertarDesdeFichero = False
        Exit Function
    End If
    
    If FileLen(cd1.FileName) / 1000 > 1024 Then
        MsgBoxA "No se permite insertar ficheros de tama�o superior a 1 M", vbExclamation
        InsertarDesdeFichero = False
        Exit Function
    End If
    
    If LCase(Right(cd1.FileName, 3)) <> "pdf" Then
        MsgBoxA "Debe seleccionar documentos pdf", vbExclamation
        InsertarDesdeFichero = False
        Exit Function
    End If
    
'    '******* Cambiamos cursor
    Screen.MousePointer = vbHourglass

    J = InStr(1, cd1.FileName, Chr(0))
    CADENA = cd1.FileName
    
        
            
    Screen.MousePointer = vbDefault
    
    
    'De momento solo un documento
    'Fichero = ObtenerWhereCP(True)
    'Fichero = "select max(factpro_fichdocs) from hcabapu_fichdocs where " & SQL & """"
    'L = CLng(DevuelveValor(Fichero) + 1)
    C = 1
    
    Fichero = DevuelveDesdeBD("max(codigo)", "factpro_fichdocs", "1", "1")
    L = Val(Fichero) + 1
    
    
    Fichero = CADENA
    
    ' es nuevo
    C = "," & DBSet(Text1(2).Text, "T") & "," & C & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(cd1.FileTitle, "T") & ")"
    C = " (" & DBSet(L, "N") & "," & DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Year(Text1(1).Text), "N") & C
    
    C = "INSERT INTO factpro_fichdocs(codigo,numserie,numregis,anofactu,numfactu,orden,fechacrea,usucrea,docum) values" & C
    Conn.Execute C
    
    espera 0.2
    
    
    'Abro parar guardar el binary
    C = "Select * from factpro_fichdocs where codigo =" & L
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = C
    Adodc1.Refresh
'
    If Adodc1.Recordset.EOF Then
        'MAAAAAAAAAAAAL

    Else
        'Guardar
        GuardarBinary Adodc1.Recordset!Campo, Fichero
        Adodc1.Recordset.Update
        
        Adodc1.RecordSource = "Select * from factpro_fichdocs where false"
        Adodc1.Refresh
    End If
    InsertarDesdeFichero = True
End Function












Private Sub txtaux3_GotFocus(Index As Integer)
    ConseguirFoco txtaux3(Index), 3
End Sub

Private Sub txtaux3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYImage KeyAscii, 0 ' cta base
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtaux3_LostFocus(Index As Integer)
Dim Im As Currency
Dim C As Currency
    Dim RC As String
        
    
        Select Case Index
        
            
        Case 2, 3, 4
            If Not PonerFormatoDecimal(txtaux3(Index), 1) Then
                txtaux3(Index).Text = ""
            Else
                If Index = 2 Then
                    'Base imponible. Calculo IVA y recargo y voy a  aceptar
                    Im = ImporteFormateado(txtaux3(Index).Text)
                    C = 0
                    If txtaux3(2).Tag <> "" Then C = txtaux3(2).Tag
                    Im = Round((Im * C) / 100, 2)
                    txtaux3(3).Text = Format(Im, FormatoImporte)
                    'Recargo
                    C = 0
                    If txtaux3(3).Tag <> "" Then C = txtaux3(3).Tag
                    Im = ImporteFormateado(txtaux3(Index).Text)
                    Im = Round((Im * C) / 100, 2)
                    txtaux3(4).Text = Format(Im, FormatoImporte)
                    
                    PonerFocoBtn cmdAceptar
                End If
            End If
            
        Case 0 ' iva
            
            RC = "concat(porceiva,'|',porcerec,'|')"
            txtaux3(1).Text = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", txtaux3(0), "N", RC)
            If txtaux3(1).Text = "" Then
                
                If txtaux3(0).Text <> "" Then
                    MsgBoxA "No existe el Tipo de Iva. Reintroduzca.", vbExclamation
                    txtaux3(0).Text = ""
                    PonFoco txtaux3(0)
                End If
                txtaux3(2).Tag = 0
                txtaux3(3).Tag = 0
            Else
                txtaux3(2).Tag = TransformaPuntosComas(RecuperaValor(RC, 1))
                txtaux3(3).Tag = TransformaPuntosComas(RecuperaValor(RC, 2))
            End If
            
            
        End Select
            
End Sub



Private Function AnyadirModificarIVA() As Boolean

    AnyadirModificarIVA = False
    
    For NumRegElim = 0 To 4
        If txtaux3(NumRegElim).Text = "" Then
            MsgBoxA "Campos obligatorios", vbExclamation
            Exit Function
        End If
    Next
    
    'LLegados aqui, vamos a modificar o insertar
    CadB1 = ObtenerWhereCab(False)
    CadB1 = Replace(CadB1, "factpro.", "")
    If ModoLineas = 1 Then
        CadB = "INSERT INTO factpro_totales set "
        'Para el where
        
        CadB2 = CadB1
        'Numlinea
        CadB2 = DevuelveDesdeBD("max(numlinea)", "factpro_totales", CadB2, 1)
        CadB2 = Val(CadB2) + 1
        
        CadB1 = Replace(CadB1, " and ", ",")
        CadB1 = CadB1 & ", numlinea = " & CadB2
        'impoiva imporec
        CadB1 = CadB1 & ", porciva  = " & DBSet(txtaux3(2).Tag, "N", "N")
        CadB1 = CadB1 & ", porcrec  = " & DBSet(txtaux3(3).Tag, "N", "N")
        CadB1 = CadB1 & ", codigiva  = " & DBSet(txtaux3(0).Text, "N")
        CadB1 = CadB1 & ", fecharec  = " & DBSet(Text1(1).Text, "F")
        CadB = CadB & CadB1
        
        
        CadB2 = ""
    Else
        CadB2 = " WHERE " & CadB1 & " AND numlinea =" & txtaux3(0).Tag  'para el where
        
        CadB = "UPDATE factpro_totales set porciva=porciva "
    End If
    CadB1 = ""
    For NumRegElim = 2 To 4
        CadB1 = CadB1 & ", " & RecuperaValor("baseimpo|impoiva|imporec|", NumRegElim - 1) & "=" & DBSet(txtaux3(NumRegElim), "N", "N")
    Next
    CadB = CadB & CadB1 & CadB2
    
    If Ejecuta(CadB, False) Then AnyadirModificarIVA = True
    
End Function

