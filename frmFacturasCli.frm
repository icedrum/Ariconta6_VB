VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFacturasCli 
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
      Left            =   450
      TabIndex        =   95
      Top             =   2520
      Visible         =   0   'False
      Width           =   16935
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
         Index           =   5
         ItemData        =   "frmFacturasCli.frx":0000
         Left            =   13440
         List            =   "frmFacturasCli.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Tag             =   "Doc- ID|N|N|||factcli|TipoDocumentoId|||"
         Top             =   390
         Width           =   3210
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
         Index           =   21
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   111
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
         TabIndex        =   103
         Tag             =   "Pa�s|T|S|||factcli|codpais|||"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
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
         Index           =   20
         Left            =   10290
         TabIndex        =   97
         Tag             =   "Nif|T|S|||factcli|nifdatos|||"
         Text            =   "teetetete"
         Top             =   390
         Width           =   3030
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
         TabIndex        =   102
         Tag             =   "Provincia|T|S|||factcli|desprovi|||"
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
         TabIndex        =   101
         Tag             =   "Poblacion|T|S|||factcli|despobla|||"
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
         TabIndex        =   100
         Tag             =   "CP|T|S|||factcli|codpobla|||"
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
         TabIndex        =   99
         Tag             =   "Direcci�n|T|S|||factcli|dirdatos|||"
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
         TabIndex        =   96
         Tag             =   "Nombre|T|N|||factcli|nommacta|||"
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   108
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "NIF / ID"
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
         TabIndex        =   107
         Top             =   450
         Width           =   1005
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
         Top             =   450
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2385
      Left            =   9690
      TabIndex        =   84
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
         Left            =   5760
         TabIndex        =   128
         Tag             =   "Importe Iva|N|S|||factcli|suplidos|###,###,##0.00||"
         Text            =   "123456789012345"
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
         Left            =   1800
         TabIndex        =   126
         Tag             =   "Importe Iva|N|S|||factcli|totrecargo|###,###,##0.00||"
         Text            =   "123456789012345"
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
         Left            =   5760
         TabIndex        =   22
         Tag             =   "Total Factura|N|S|||factcli|totfaccl|###,###,##0.00||"
         Text            =   "123456789012345"
         Top             =   1920
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
         Left            =   5760
         TabIndex        =   21
         Tag             =   "Importe Retenci�n|N|S|||factcli|trefaccl|###,###,##0.00||"
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
         Left            =   5760
         TabIndex        =   20
         Tag             =   "Base Retenci�n|N|S|||factcli|totbasesret|###,###,##0.00||"
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
         Left            =   1800
         TabIndex        =   19
         Tag             =   "Importe Iva|N|S|||factcli|totivas|###,###,##0.00||"
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
         Index           =   9
         Left            =   1800
         TabIndex        =   18
         Tag             =   "Base Imponible|N|S|||factcli|totbases|###,###,##0.00||"
         Text            =   "1.023.654.99"
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
         Index           =   25
         Left            =   3720
         TabIndex        =   129
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
         Index           =   24
         Left            =   120
         TabIndex        =   127
         Top             =   1320
         Width           =   1755
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
         Left            =   3720
         TabIndex        =   90
         Top             =   1965
         Width           =   1935
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
         Left            =   3720
         TabIndex        =   89
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
         Left            =   3720
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
         Top             =   0
         Width           =   1980
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3930
      TabIndex        =   71
      Top             =   90
      Width           =   2985
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   72
         Top             =   180
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Datos Fiscales"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cobros"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Errores N�Factura"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas sin Asiento"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Sii. Modificar factura presentada"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   9960
      TabIndex        =   69
      Top             =   90
      Width           =   2445
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
         ItemData        =   "frmFacturasCli.frx":0004
         Left            =   120
         List            =   "frmFacturasCli.frx":0011
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
         TabIndex        =   132
         ToolTipText     =   "Buscar cuenta"
         Top             =   1440
         Visible         =   0   'False
         Width           =   195
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
         Left            =   2880
         TabIndex        =   67
         Tag             =   "Iva|N|S|||factcli_totales|codigiva|000||"
         Text            =   "Iva"
         Top             =   1560
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameModifIVA 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   240
         TabIndex        =   130
         Top             =   240
         Visible         =   0   'False
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAuxTot 
            Height          =   330
            Left            =   0
            TabIndex        =   131
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
         Index           =   4
         Left            =   8160
         TabIndex        =   66
         Tag             =   "Importe Rec|N|S|||factcli_totales|imporec|###,###,##0.00||"
         Text            =   "ImpRec"
         Top             =   1590
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
         Index           =   3
         Left            =   7260
         TabIndex        =   65
         Tag             =   "Importe Iva|N|S|||factcli_totales|impoiva|###,###,##0.00||"
         Text            =   "ImpIva"
         Top             =   1560
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
         Index           =   0
         Left            =   1800
         TabIndex        =   63
         Tag             =   "Iva|N|S|||factcli_totales|codigiva|000||"
         Text            =   "Iva"
         Top             =   1560
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
         Index           =   2
         Left            =   4320
         TabIndex        =   64
         Tag             =   "Base Imponible|N|S|||factcli_totales|baseimpo|###,###,##0.00||"
         Text            =   "Base Imponible"
         Top             =   1560
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
         Top             =   600
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
         Left            =   120
         TabIndex        =   70
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
      Left            =   6960
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
      Left            =   240
      TabIndex        =   50
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   52
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
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
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Height          =   3915
      Index           =   0
      Left            =   270
      TabIndex        =   39
      Top             =   870
      Width           =   17160
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
         Index           =   33
         Left            =   12960
         MaxLength       =   20
         TabIndex        =   125
         Tag             =   "Factura fin|T|S|||factcli|FraResumenFin|||"
         Top             =   1260
         Visible         =   0   'False
         Width           =   1860
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
         Index           =   32
         Left            =   10320
         MaxLength       =   20
         TabIndex        =   124
         Tag             =   "Factura ini|T|S|||factcli|FraResumenIni|||"
         Top             =   1260
         Visible         =   0   'False
         Width           =   1860
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
         ItemData        =   "frmFacturasCli.frx":0048
         Left            =   13080
         List            =   "frmFacturasCli.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Tag             =   "Situacion inmueble|N|S|||factcli|CatastralSitu|||"
         Top             =   1260
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
         Left            =   11040
         MaxLength       =   20
         TabIndex        =   120
         Tag             =   "RCatas|T|S|||factcli|CatastralREF|||"
         Top             =   1260
         Visible         =   0   'False
         Width           =   1860
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
         Left            =   12720
         TabIndex        =   16
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
         ItemData        =   "frmFacturasCli.frx":004C
         Left            =   10530
         List            =   "frmFacturasCli.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1260
         Visible         =   0   'False
         Width           =   6330
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   210
         TabIndex        =   11
         Tag             =   "Agente|N|S|||factcli|codagente|000||"
         Text            =   "1234567890"
         Top             =   2580
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
         Index           =   26
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Text4"
         Top             =   2580
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   210
         TabIndex        =   9
         Tag             =   "Departamento|N|S|||factcli|dpto|0000||"
         Text            =   "1234567890"
         Top             =   1950
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
         Index           =   25
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text4"
         Top             =   1950
         Width           =   6135
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
         TabIndex        =   17
         Tag             =   "N�mero Asiento|N|S|||factcli|numasien|00000000||"
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
         Left            =   8040
         TabIndex        =   4
         Tag             =   "Fecha Liquidacion|F|N|||factcli|fecliqcl|||"
         Top             =   570
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
         Left            =   11280
         TabIndex        =   15
         Tag             =   "Porcentaje Retencion|N|S|||factcli|retfaccl|##0.00||"
         Text            =   "1234567890"
         Top             =   3270
         Width           =   1335
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
         Tag             =   "Cuenta Retencion|T|S|||factcli|cuereten|||"
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
         ItemData        =   "frmFacturasCli.frx":0050
         Left            =   180
         List            =   "frmFacturasCli.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Tipo retencion|N|N|||factcli|tiporeten|||"
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
         ItemData        =   "frmFacturasCli.frx":0054
         Left            =   7950
         List            =   "frmFacturasCli.frx":0056
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Tipo operaci�n|N|N|||factcli|codopera|||"
         Top             =   1260
         Width           =   2250
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
         ItemData        =   "frmFacturasCli.frx":0058
         Left            =   9540
         List            =   "frmFacturasCli.frx":005A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   7320
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
         Left            =   9420
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text4"
         Top             =   1950
         Width           =   7425
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text4"
         Top             =   570
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
         Height          =   375
         Index           =   3
         Left            =   7950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "Observaciones|T|S|||factcli|observa|||"
         Top             =   2580
         Width           =   8895
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
         Left            =   240
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "Serie|T|N|||factcli|numserie||S|"
         Text            =   "123"
         Top             =   570
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
         Left            =   6525
         TabIndex        =   3
         Tag             =   "N� factura|N|S|0||factcli|numfactu|0000000|S|"
         Top             =   570
         Width           =   1395
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
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Fecha|F|N|||factcli|fecfactu|dd/mm/yyyy|N|"
         Top             =   570
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
         TabIndex        =   6
         Tag             =   "Cuenta Cliente|T|N|||factcli|codmacta|||"
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
         Left            =   7950
         TabIndex        =   10
         Tag             =   "Forma de pago|N|N|||factcli|codforpa|000||"
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
         Left            =   7950
         TabIndex        =   91
         Tag             =   "A�o factura|N|N|||factcli|anofactu||S|"
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
         Left            =   10860
         MaxLength       =   30
         TabIndex        =   93
         Tag             =   "Tipo factura|T|N|||factcli|codconce340|||"
         Top             =   570
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
         TabIndex        =   112
         Tag             =   "N�mero Diario|N|S|||factcli|numdiari|00000000||"
         Text            =   "1234567890"
         Top             =   2595
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
         Left            =   10560
         MaxLength       =   30
         TabIndex        =   117
         Tag             =   "Tipo intracomunitaria|T|S|||factcli|codintra|||"
         Top             =   2580
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Ticket final"
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
         Index           =   27
         Left            =   12960
         TabIndex        =   134
         Top             =   1005
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Ticket inicial"
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
         Index           =   26
         Left            =   10320
         TabIndex        =   133
         Top             =   1005
         Visible         =   0   'False
         Width           =   2235
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
         Left            =   13080
         TabIndex        =   122
         Top             =   990
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
         Left            =   10320
         TabIndex        =   121
         Top             =   990
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
         Left            =   12720
         TabIndex        =   119
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label Label9 
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
         Left            =   10530
         TabIndex        =   118
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Agente"
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
         Left            =   210
         TabIndex        =   116
         Top             =   2310
         Width           =   1545
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   10
         Left            =   1770
         Top             =   2310
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
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
         Left            =   210
         TabIndex        =   114
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   9
         Left            =   1770
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   8
         Left            =   9510
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   7
         Left            =   9060
         Top             =   270
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
         Left            =   8040
         TabIndex        =   94
         Top             =   270
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
         Caption         =   "% Retencion"
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
         Left            =   11280
         TabIndex        =   82
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
         TabIndex        =   79
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
         Left            =   7980
         TabIndex        =   78
         Top             =   990
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
         Left            =   9600
         TabIndex        =   77
         Top             =   270
         Width           =   1380
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   9510
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
         Left            =   7950
         TabIndex        =   76
         Top             =   1650
         Width           =   1545
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   1770
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Cliente"
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
         Top             =   990
         Width           =   1545
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   6150
         Picture         =   "frmFacturasCli.frx":005C
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   840
         Top             =   270
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
         Left            =   7950
         TabIndex        =   44
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   5190
         TabIndex        =   43
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "N� Factura"
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
         Left            =   6660
         TabIndex        =   41
         Top             =   270
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
         Left            =   240
         TabIndex        =   40
         Top             =   270
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
         TabIndex        =   92
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
         Tag             =   "Aplica Retencion|N|N|0|1|factcli_lineas|aplicret|||"
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
         TabIndex        =   36
         Tag             =   "Importe Rec|N|S|||factcli_lineas|imporec|###,###,##0.00||"
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
         Tag             =   "A�o factura|N|N|||factcli_lineas|anofactu||S|"
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
         Tag             =   "CC|T|S|||factcli_lineas|codccost|||"
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
         TabIndex        =   35
         Tag             =   "Importe Iva|N|S|||factcli_lineas|impoiva|###,###,##0.00||"
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
         Tag             =   "% Recargo|N|S|||factcli_lineas|porcrec|##0.00||"
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
         Tag             =   "% Iva|N|S|||factcli_lineas|porciva|##0.00||"
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
         MaxLength       =   12
         TabIndex        =   31
         Tag             =   "Importe Base|N|N|||factcli_lineas|baseimpo|###,###,##0.00||"
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
         Tag             =   "Codigo Iva|N|N|||factcli_lineas|codigiva|000||"
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
         Tag             =   "Cuenta|T|N|||factcli_lineas|codmacta|||"
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
         Tag             =   "Linea|N|N|||factcli_lineas|numlinea||S|"
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
         Index           =   2
         Left            =   2220
         TabIndex        =   27
         Tag             =   "Fecha|F|N|||factcli_lineas|fecfactu|dd/mm/yyyy||"
         Text            =   "fecha"
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
         TabIndex        =   25
         Tag             =   "N� Serie|T|S|||factcli_lineas|numserie||S|"
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
         TabIndex        =   26
         Tag             =   "N� factura|N|N|0||factcli_lineas|numfactu|000000|S|"
         Text            =   "factura"
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
      TabIndex        =   23
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
         TabIndex        =   24
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
      Left            =   720
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
Attribute VB_Name = "frmFacturasCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'Public DatosADevolverBusqueda As StringCombo1(0)    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Public FACTURA As String  'Con pipes nuwtipoperamserie|numfactu|anofactu




Private Const NO = "No encontrado"

Private Const IdPrograma = 401

Private WithEvents frmFact As frmFacturasCliPrev
Attribute frmFact.VB_VarHelpID = -1
Private WithEvents frmFPag As frmBasico2
Attribute frmFPag.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmPais As frmBasico2
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmAgen As frmBasico
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

Private WithEvents frmCob As frmFacturasCliCob ' cobros de tesoreria
Attribute frmCob.VB_VarHelpID = -1
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

Dim NumAsien As Long
Dim NumDiario_ As Integer
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
'   6.- Totales . Para facturas que vienen desde SAGE

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
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de ll�nies
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
Dim I As Long
Dim Ancho As Integer

Private Mc As Contadores

Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

'Por si esta en un periodo liquidado, que pueda modificar CONCEPTO , cuentas,
Private ModificaFacturaPeriodoLiquidado As Boolean


Dim IvaCuenta As String
Dim CambiarIva As Boolean

Dim CtaBanco As String
Dim IBAN As String
Dim NomBanco As String

Dim Cobrado As Byte
Dim FechaCobro As String

Dim TipForpa As Integer
Dim FecFactuAnt As String
Dim AntLetraSer As String

Dim ModificarCobros As Boolean

Dim HaCambiadoTipoIva As Boolean  'Modificando lineas. Si ya ha cambiado el tipo de iva una vez, siempre tendra que calcular el IVA

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
        'cmdAceptar.SetFocus
        PonleFoco cmdAceptar
    End If
End Sub

Private Sub cmdAceptar_Click()
     Dim cad As String
    Dim I As Integer
    Dim Limp As Boolean
    Dim Mc As Contadores
    Dim B As Boolean

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                FecFactuAnt = Text1(1).Text
                
                Set Mc = New Contadores
                I = FechaCorrecta2(CDate(Text1(1).Text))
                    If ComprobarNumeroFactura(I = 0) Then
                        B = InsertarDesdeForm2(Me, 1)
                    Else
                        B = False
                    End If
                    
                    If B Then
                        Data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PosicionarData
                        PonerCampos
                        NumAsien = 0
                        BotonAnyadirLinea 1, True
                    Else
                        'SI NO INSERTA debemos devolver el contador
                        Mc.DevolverContador Trim(Text1(2).Text), (I = 0), Mc.Contador
                    End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                '-----------------------------------------
                'Hay que comprobar si ha modificado, o no la clave de la factura
                I = 1
                If Data1.Recordset!NUmSerie = Text1(2).Text Then
                    If Data1.Recordset!numfactu = CLng(Text1(0).Text) Then
                        If Data1.Recordset!Anofactu = Text1(14).Text Then
                            I = 0
                            'NO HA MODIFICADO NADA
                        End If
                    End If
                End If
            
                'Hacemos MODIFICAR
                Dim RC As Boolean
                If I <> 0 Then
                    MsgBoxA "No se puede cambiar campos clave  de la factura.", vbExclamation
                    RC = False
                Else
                    RC = ModificarFactura
                End If
                    
                If RC Then
                    '--DesBloqueaRegistroForm Me.Text1(0)
                    TerminaBloquear
                    
                    If NumAsien > 0 Then  'ContabilizaApunte
                        
                        If IntegrarFactura(False) Then
                            Text1(8).Text = Format(NumAsien, "0000000")
                            NumAsien = -1
                            NumDiario_ = 0
                            
                        Else
                            B = False
                        End If
                    
                    End If
                    
                    If Not ModificarCobros Then
                        'Si ha cambiado CODMACTA , OBVIAMENTE, deberia cargar cobros denuevo
                        If Text1(4).Text <> DBLet(Data1.Recordset!codmacta, "T") Then ModificarCobros = True
                    End If
                    
                    If ModificarCobros Then
                        CobrosTesoreria
                    Else
                        'No ha modificado NADA respecto a cobro, pero si que ha actualizado Nommacta
                        'Si ha modificado el nombre en datos fiscales tendremos que updaear en cobors
                        If DBLet(Data1.Recordset!Nommacta, "T") <> Text1(15).Text Then
                            cad = "UPDATE cobros set nomclien = " & DBSet(Text1(15).Text, "T")
                            cad = cad & " where numserie = " & DBSet(Text1(2).Text, "T")
                            cad = cad & " and numfactu = " & DBSet(Text1(0).Text, "N")
                            cad = cad & " and fecfactu = " & DBSet(Text1(1).Text, "F")
                            Ejecuta cad, False
                            cad = ""
                        End If
                    
                    End If
                    'LOG
                    vLog.Insertar 5, vUsu, "Factura : " & Text1(2).Text & Text1(0).Text & " " & Text1(1).Text
                    'Creo que no hace falta volver a situar el datagrid
                    'If SituarData1(0) Then
                    PosicionarData
                    
                    
                    If FACTURA <> "" Then Unload Me
                    
                End If
            End If
        
        Case 5 'LL�NIES
            FecFactuAnt = Text1(1).Text
            
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    InsertarLinea
                Case 2 'modificar ll�nies
            
                        If ModificarLinea Then
                                                
                            '**** parte de contabilizacion de la factura
                            TerminaBloquear
                            
                            If NumAsien > 0 Then
                                If IntegrarFactura(False) Then
                                    Text1(8).Text = Format(NumAsien, "0000000")
                                    NumAsien = -1
                                    NumDiario_ = 0
                                Else
                                    B = False
                                End If
                            End If
                        
                            If ModificarCobros Then CobrosTesoreria
                            
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


Private Sub CobrosTesoreria()
Dim Sql As String
Dim Rs As ADODB.Recordset

    
    On Error GoTo eCobrosTesoreria

    If Not vEmpresa.TieneTesoreria Then Exit Sub
    
    '[Monica]12/09/2016: si la factura ha sido traspasada y no est� en cartera, no hacemos nada en cartera
    If EsFraCliTraspasada And Not ExisteAlgunCobro(Text1(2).Text, Text1(0).Text, FecFactuAnt, False) Then Exit Sub
    
    If ExisteAlgunCobro(Text1(2).Text, Text1(0).Text, FecFactuAnt, True) Then
        MsgBoxA "Hay alg�n efecto que ya ha sido cobrado. Revise cartera de cobros.", vbExclamation

        Set frmMens = New frmMensajes

        frmMens.Opcion = 27
        frmMens.Parametros = Trim(Text1(2).Text) & "|" & Trim(Text1(0).Text) & "|" & Text1(1).Text & "|"
        frmMens.Show vbModal

        Set frmMens = Nothing

        ContinuarCobro = False

        Exit Sub
    End If
    

    Sql = "delete from tmpcobros where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute Sql
    
    ContinuarCobro = False
    
    If CargarCobrosTemporal(Text1(5).Text, Text1(1).Text, ImporteFormateado(Text1(13).Text)) Then
        ' Insertamos
        If Not ExisteAlgunCobro(Text1(2).Text, Text1(0).Text, FecFactuAnt, False) Then
    
            Sql = "select ccc.ctabanco,ccc.iban, ddd.nommacta "
            Sql = Sql & " from cuentas ccc left join  cuentas ddd ON ccc.ctabanco = ddd.codmacta "
            Sql = Sql & " where ccc.codmacta = " & DBSet(Text1(4).Text, "T")
            
            
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
            Rs.Close
            
            Set frmCob = frmFacturasCliCob
            
             If IsNull(Data1.Recordset!totfaccl) Then
                'Insertando
                Sql = ObtenerWhereCab(False) & " AND 1"
                Sql = DevuelveDesdeBD("totfaccl", "factcli", Sql, "1")
                If Sql = "" Then Sql = "0"
                
            Else
                Sql = Data1.Recordset!totfaccl
            End If
            
            frmCob.CodigoActual = CtaBanco & "|" & "|" & "|" & "|" & "|" & IBAN & "|" & TipForpa & "|" & NomBanco & "|" & Sql & "|"
            frmCob.Show vbModal
            Set frmCob = Nothing
    
            If ContinuarCobro Then
                CargarCobros
                If Cobrado Then ContabilizarCobros
            End If
            
        Else
            Dim Nregs As Long
            Dim Nregs2 As Long
            Nregs = TotalRegistros("select count(*) from tmpcobros where codusu = " & vUsu.Codigo)
            Nregs2 = TotalRegistros("select count(*) from cobros where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(FecFactuAnt, "F"))

            If Nregs = Nregs2 Then
                CargarCobros
            Else
                MsgBoxA "No coincide el n�mero de cobros en tesoreria. Modif�quelos en cartera.", vbExclamation
                ' mandarlo al listview de cobros
            
                Set frmMens = New frmMensajes
                
                frmMens.Opcion = 27
                frmMens.Parametros = Trim(Text1(2).Text) & "|" & Trim(Text1(0).Text) & "|" & Text1(1).Text & "|"
                frmMens.Show vbModal
                
                Set frmMens = Nothing
            
            End If
        
        End If
    End If
    
    
    Exit Sub
    
eCobrosTesoreria:
    MuestraError Err.Number, "Cobros Tesoreria", Err.Description
End Sub

Private Function ExisteAlgunCobro(Serie As String, FACTURA As String, Fecfactu As String, Cobrado As Boolean) As Boolean
Dim Sql As String
    
    Sql = "select count(*) from cobros where numserie = " & DBSet(Serie, "T")
    Sql = Sql & " and numfactu = " & DBSet(FACTURA, "N")
    Sql = Sql & " and fecfactu = " & DBSet(Fecfactu, "F")
    
    If Cobrado Then
' un cobro lo damos como cobrado si el importe de cobro es <> 0
'[Monica]12/09/2016: quito la condicion: numasien is null pq puede tener nro de remesa y no modificariamos el importe total de remesa
        Sql = Sql & " and impcobro <> 0 and not impcobro is null " ' and numasien is null "
    End If
    
    ExisteAlgunCobro = (TotalRegistros(Sql) <> 0)

End Function


Private Function CobrosContabilizados(Serie As String, FACTURA As String, Fecfactu As String) As String
Dim Sql As String
Dim CadResult As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCobrosContabilizados

    Sql = "select numasien, fechaent from hlinapu where numserie = " & DBSet(Serie, "T")
    Sql = Sql & " and numfaccl = " & DBSet(FACTURA, "N")
    Sql = Sql & " and fecfactu = " & DBSet(Fecfactu, "F")
    
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
    
    
    CobrosContabilizados = CadResult
    
    Exit Function
    
eCobrosContabilizados:
    MuestraError Err.Number, "Cobros contabilizados", Err.Description
End Function



Private Sub CargarCobros()
Dim Sql As String
Dim Mens As String

    If ExisteAlgunCobro(Text1(2).Text, Text1(0).Text, FecFactuAnt, False) Then
        B = ActualizarCobros(Mens)
        
        If B Then
            Sql = CobrosContabilizados(Text1(2).Text, Text1(0).Text, FecFactuAnt)
            If Sql <> "" Then
                MsgBoxA "La factura tiene asientos que ya est�n contabilizados. Revise y modifique en su caso los siguientes asientos: " & vbCrLf & vbCrLf & Sql, vbExclamation
            End If
        End If
    Else
        B = InsertarCobros(Mens)
    End If
    
    If B Then
'        msgboxA "Proceso realizado correctamente.", vbExclamation
    Else
        MuestraError 0, "Cargar Cobros", Mens
    End If

End Sub

Private Function UpdateaCobros(ByRef Rs As ADODB.Recordset, ByRef RS1 As ADODB.Recordset, ByRef I As Long, ByRef Mens As String) As Boolean
Dim Sql As String

    On Error GoTo eUpdateaCobros
    
    UpdateaCobros = False

    B = True

    While Not Rs.EOF And B
        Sql = "update cobros set codmacta = " & DBSet(Text1(4).Text, "T")
        Sql = Sql & ", codforpa = " & DBSet(Text1(5).Text, "N")
        Sql = Sql & ", fecvenci = " & DBSet(RS1!FecVenci, "F")
        Sql = Sql & ", impvenci = " & DBSet(RS1!ImpVenci, "N")
        Sql = Sql & ", fecfactu = " & DBSet(Text1(1).Text, "F")
        
        If Cobrado Then
            Sql = Sql & ", fecultco = " & DBSet(FechaCobro, "F") ' DBSet(Rs!FecVenci, "F")
            Sql = Sql & ", impcobro = " & DBSet(RS1!ImpVenci, "N")
        Else
            Sql = Sql & ", fecultco = " & ValorNulo
            Sql = Sql & ", impcobro = " & ValorNulo
            If CtaBanco = "" Then CtaBanco = Rs!ctabanc1
            Sql = Sql & ", ctabanc1 = " & DBSet(CtaBanco, "T", "S")
        End If
        Sql = Sql & ", agente = " & DBSet(Text1(26).Text, "N", "S")
        Sql = Sql & ", departamento = " & DBSet(Text1(25).Text, "N", "S")
        Sql = Sql & ", nomclien = " & DBSet(Text1(15).Text, "T", "S")
        Sql = Sql & ", domclien = " & DBSet(Text1(16).Text, "T", "S")
        Sql = Sql & ", pobclien = " & DBSet(Text1(18).Text, "T", "S")
        Sql = Sql & ", cpclien = " & DBSet(Text1(17).Text, "T", "S")
        Sql = Sql & ", proclien = " & DBSet(Text1(19).Text, "T", "S")
        Sql = Sql & ", iban = " & DBSet(IBAN, "T", "S")
        Sql = Sql & ", numorden = " & DBSet(RS1!numorden, "N")
        Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and fecfactu = " & DBSet(FecFactuAnt, "F") & " and numorden = " & DBSet(Rs!numorden, "N")
        
        Conn.Execute Sql
        
        I = Rs!numorden ' me guardo el nro de orden para despues ir incrementandolo
        
        RS1.MoveNext
        Rs.MoveNext
        
        ' si no hay mas registros en la temporal salgo del bucle
        If RS1.EOF Then B = False
    Wend
    
    UpdateaCobros = True
    Exit Function

eUpdateaCobros:
    Mens = Mens & Err.Description
End Function


Private Function ActualizarCobros(ByRef Mens As String) As Boolean
Dim Sql As String
Dim Sql1 As String
Dim Nregs As Integer
Dim Nregs1 As Integer
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim CadInsert As String
Dim CadValues As String

    On Error GoTo eActualizarCobros

    ActualizarCobros = False


    Sql = "select * from cobros where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and fecfactu = " & DBSet(FecFactuAnt, "F")
    
    Sql = Sql & " order by numorden "
    Nregs = TotalRegistrosConsulta(Sql)
    
    Sql1 = "select * from tmpcobros where codusu = " & vUsu.Codigo & " order by numorden "
    Nregs1 = TotalRegistrosConsulta(Sql1)
    
    If Nregs = Nregs1 Then
    ' Mismo nro de registros en cobros que en la temporal --> los actualizamos
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        I = 0
        
        Mens = "Actualizando Cobros: " & vbCrLf & vbCrLf
        B = UpdateaCobros(Rs, RS1, I, Mens)
        
        Set Rs = Nothing
        Set RS1 = Nothing
    
    ElseIf Nregs < Nregs1 Then
    ' Menos registros en cobros que en la temporal --> actualizamos e insertamos los no existentes
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        I = 0
        
        Mens = "Actualizando Cobros: " & vbCrLf & vbCrLf
        B = UpdateaCobros(Rs, RS1, I, Mens)
        
        Set Rs = Nothing ' cierro el de cobros
        
        ' sin cerrar el recordset de tmpcobros, insertamos los restantes registros de la tmpcobros
        Mens = "Insertando Cobros Restantes: " & vbCrLf & vbCrLf
        B = InsertaCobros(RS1, I, Mens)
        
        Set RS1 = Nothing
    
    Else
    ' Mas registros en cobros que en la temporal --> actualizamos y borramos los que sobran
    
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        I = 0
        
        Mens = "Actualizando Cobros: " & vbCrLf & vbCrLf
        B = UpdateaCobros(Rs, RS1, I, Mens)
        
        Set Rs = Nothing ' cierro el de cobros
        
        'borro los registros restantes de cobros
        Mens = "Eliminado Cobros restantes: " & vbCrLf & vbCrLf
        Sql = "delete from cobros "
        Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F") & " and numorden > " & DBSet(I, "N")
        
        Conn.Execute Sql
        
        Set RS1 = Nothing
    End If

    ActualizarCobros = B
    Exit Function

eActualizarCobros:
    Mens = Mens & Err.Description
End Function





Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "numserie= " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
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
            If txtAux(4).Text <> "" Then
                PonFoco txtAux(5)
            Else
                PonFoco txtAux(4)
            End If
        Case 1 'tipo de iva
            cmdAux(0).Tag = 1
            
            Set frmTIva = New frmBasico2
            AyudaTiposIva frmTIva
            Set frmTIva = Nothing
            
            PonFoco txtAux(7)
            If txtAux(7).Text <> "" Then txtAux_LostFocus 7
        Case 2 'cento de coste
            If txtAux(12).Enabled Then
                Set frmCC = New frmBasico
                AyudaCC frmCC
                Set frmCC = Nothing
            End If

    End Select
'    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
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
        If txtaux3(0).Text <> "" Then TxtAux3_LostFocus 0
End Sub

Private Sub Combo1_Click(Index As Integer)
    If PrimeraVez Then Exit Sub
    If Index = 2 And (Modo = 3 Or Modo = 4) Then
        If Combo1(Index).ListIndex = 0 Then
            Text1(7).Text = ""
            Text1(6).Text = ""
            Text4(6).Text = ""
        End If
    End If
    If Index = 0 And (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        If Combo1(0).ListIndex = 0 Then
            Text1(22).Text = "0"
        Else
            If Combo1(0).ListIndex <> -1 Then Text1(22).Text = Chr(Combo1(0).ItemData(Combo1(0).ListIndex))
        End If
        
    End If
    If Combo1(0).ListIndex = 18 Then
        ReferenciaCatastral True
    Else
        ReferenciaCatastral False
    End If
    FacturaResumenIniFin Combo1(0).ListIndex = 2
    
        
    
    If Index = 1 And (Modo = 1 Or Modo = 2 Or Modo = 3 Or Modo = 4) Then
        
        If Combo1(1).ListIndex = 1 Then
            ReferenciaCatastral False
            FacturaResumenIniFin False
            Combo1(3).visible = True
            Label9.visible = True
            Combo1(3).Enabled = True
            Label9.Enabled = True
            
            If Modo = 3 Then
                PosicionarCombo Combo1(3), Asc("E")
                Text1(27).Text = "E"
            End If
        Else
            
            
            Combo1(3).visible = False
            Label9.visible = False
            Combo1(3).Enabled = False
            Label9.Enabled = False
            
            Combo1(3).ListIndex = -1
            
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
    'Lo he pasado a click
End Sub

Private Sub ReferenciaCatastral(visible As Boolean)
    Text1(29).visible = visible
    Combo1(4).visible = visible
    Label1(22).visible = visible
    Label1(23).visible = visible
End Sub


Private Sub FacturaResumenIniFin(visible As Boolean)
    Text1(32).visible = visible
    Text1(33).visible = visible
    Label1(26).visible = visible
    Label1(27).visible = visible
End Sub


Private Sub Form_Activate()
    
    If PrimeraVez Then
        B = False
        If FACTURA <> "" Then
            B = True
            Modo = 2
            Sql = "Select * from factcli "
            Sql = Sql & " WHERE numserie = " & DBSet(RecuperaValor(FACTURA, 1), "T")
            Sql = Sql & " AND numfactu =" & RecuperaValor(FACTURA, 2)
            Sql = Sql & " AND anofactu= " & RecuperaValor(FACTURA, 3)
            CadenaConsulta = Sql
            PonerCadenaBusqueda
            'BOTON lineas
            If Combo1(0).ListIndex = 18 Then ReferenciaCatastral True
            If Combo1(0).ListIndex = 2 Then FacturaResumenIniFin True
            cboFiltro.ListIndex = 0
            
        Else
            Modo = 0
            'CadenaConsulta = "Select * from " & NombreTabla & " WHERE numserie is null"
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE false"
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            
            cboFiltro.ListIndex = vUsu.FiltroFactCli
            
        End If
        
        CargarSqlFiltro
         
        PonerModo CInt(Modo)
        VieneDeDesactualizar = B
'        CargaGrid 1, (Modo = 2)
        If Modo <> 2 Then
 
            If FACTURA <> "" Then MsgBoxA "Proceso de sistema. Frm_Activate", vbExclamation
            
           
            
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
            cadFiltro = "factcli.fecfactu >= " & DBSet(vParam.fechaini, "F")
        
        Case 2 ' ejercicio actual
            cadFiltro = "factcli.fecfactu between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
        
        Case 3 ' ejercicio siguiente
            cadFiltro = "factcli.fecfactu > " & DBSet(vParam.fechafin, "F")
    
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
Dim I As Integer

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
        .Buttons(5).Image = 45
        .Buttons(5).visible = vParam.SIITiene
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
    
    With Me.ToolbarAuxTot
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
    
    For I = 0 To imgppal.Count - 1
        If I <> 0 And I <> 7 Then imgppal(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next I
    imgppal(7).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    CargaFiltros
    
    
    Caption = "Facturas de Cliente"
    
    NumTabMto = 1
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han ll�nies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "factcli"
    Ordenacion = " ORDER BY factcli.numserie, factcli.numfactu , factcli.fecfactu"
    '************************************************
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    '***** canviar el nom de la PK de la cap�alera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where false "   'numserie is null"
    Data1.Refresh
       
    
    ModoLineas = 0
    DiarioPorDefecto = ""
       
    CargarColumnas
    
    CargarCombo
    
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    
    
    Label1(21).visible = vParam.SIITiene
    Text1(28).visible = vParam.SIITiene
    If vParam.SIITiene Then
        Text1(28).Tag = "Status|N|S|||factcli|SII_ID|00000000||"
    Else
        Text1(28).Tag = ""
    End If
    
    
    'Maxima longitud cuentas
    txtAux(5).MaxLength = vEmpresa.DigitosUltimoNivel
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
Dim I As Integer

    On Error Resume Next
    
    Limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    For I = 0 To Combo1.Count - 1
        Me.Combo1(I).ListIndex = -1
    Next I

    Me.chkAux(0).Value = 0
    FrameModifIVA.visible = False
    lw1.ListItems.Clear
    If vParam.SIITiene Then
        Text1(28).BackColor = vbWhite
        Text1(21).ToolTipText = ""
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub Color_CampoSII()
Dim Color As Byte
Dim Aux As String
Dim CaptToolText As String
    
    'Ene 2021
    'Columna status_ID
    '  De momento si es <5  sigue como estaba ya que seria pdte, incorrecta, aceptada con errores
    '
    ' Si es mayor :     8. Modificandose y luego la modifican primero en el SII
    '                   9.  TODO OK
    CaptToolText = ""
    If DBLet(Data1.Recordset!SII_ID, "N") = 0 Then
        Color = 0
    Else
        
        If Val(Data1.Recordset!SII_estado) > 4 Then
            'NUEVO
            
            If Val(Data1.Recordset!SII_estado) = 8 Then
                'Modificandose para poder rectificarla en PORTAL
                Color = 6
                CaptToolText = "Mofic�ndose para SII"
            Else
                'OK
                Color = 4
            End If
        Else
    
            'Lo que habia
            Aux = "concat(enviada,'|',coalesce(csv,''),'|',coalesce(resultado,''),'|')"
            Aux = DevuelveDesdeBD(Aux, "aswsii.envio_facturas_emitidas", "IDEnvioFacturasEmitidas", CStr(Data1.Recordset!SII_ID))
            If Aux = "" Then Aux = "1|||"
            If RecuperaValor(Aux, 1) = 1 Then
                If RecuperaValor(Aux, 2) = "" Then
                    'Erronea
                    Color = 2
                    CaptToolText = "Erronea "
                Else
                    If RecuperaValor(Aux, 3) = "Correcto" Then
                        Color = 4
                    Else
                        If RecuperaValor(Aux, 3) = "AceptadoConErrores" Then
                            Color = 5
                            CaptToolText = "Aceptada con erroeres"
                        Else
                            MsgBox "Situcion SII    no definida. Llame a soporte t�cnico", vbExclamation
                            Color = 2
                        End If
                        
                    End If
                End If
            Else
                'En proceso de carga
                Color = 3
                CaptToolText = "Subiendo"
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
            Text1(28).BackColor = vbSiiOK
        ElseIf Color = 3 Then
            'en proceso
            Text1(28).BackColor = vbSiiEnProceso
        ElseIf Color = 2 Then
            'Erronea
            Text1(28).BackColor = vbSiiError
        ElseIf Color = 5 Then
            'Aceptada con errores
            Text1(28).BackColor = vbSiiAceptadaConErr
            
        ElseIf Color = 6 Then
            'Modificando
            Text1(28).BackColor = vbSiiMofificando
        End If
    End If
    Text1(28).ToolTipText = CaptToolText
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Integer, NumReg As Byte
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
    
    For I = 0 To 27    'MENOS SII
        Text1(I).Locked = B
        If Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    
    If vParam.SIITiene Then Text1(28).Locked = Modo <> 1
    'De momento FIJO
    Text1(23).Enabled = Modo = 1
    
    For I = 0 To Combo1.Count - 1
        Combo1(I).Locked = B
    Next I
    
    For I = 0 To imgppal.Count - 1
        If I <> 8 Then imgppal(I).Enabled = Not B
    Next I
    imgppal(8).Enabled = Modo <> 0
    imgppal(6).Enabled = (Text1(8).Text <> "")
    
    ' observaciones
    'imgppal(8).Enabled = (data1.Recordset.RecordCount > 1)
    
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
       
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    B = (Modo < 5)
    chkVistaPrevia.visible = B
    
    Text1(0).Enabled = (Modo = 1 Or Modo = 3)
    
    
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
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I
    
    Frame4.Enabled = (Modo = 1)
    
    
    txtAux(8).Enabled = (Modo = 1)
    txtAux(9).Enabled = (Modo = 1)
    
    ' numero de asiento
    Text1(8).Enabled = (Modo = 1) Or (Modo = 4 And vUsu.Login = "root")
    
    
    ' ponemos en azul clarito
    Text1(0).BackColor = vbMoreLightBlue  ' factura
    Text1(13).BackColor = vbMoreLightBlue ' total factura
    Text1(4).BackColor = vbMoreLightBlue ' codmacta del cliente
    
    
    PonerModoUsuarioGnral Modo, "ariconta"

EPonerModo:
    If Err.Number <> 0 Then MsgBoxA Err.Number & ": " & Err.Description, vbExclamation
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
'   Basant-se en la informaci� proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enla�a en el data1
'           -> Si no el carreguem sense enlla�ar a cap camp
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 ' lineas de totales
            tabla = "factcli_totales"
            Sql = "SELECT factcli_totales.numserie, factcli_totales.numfactu, factcli_totales.fecfactu, factcli_totales.anofactu, factcli_totales.numlinea, factcli_totales.baseimpo, factcli_totales.codigiva, factcli_totales.porciva,"
            Sql = Sql & " factcli_totales.porcrec, factcli_totales.impoiva, factcli_totales.imporec "
            Sql = Sql & " FROM " & tabla
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "factcli", "factcli_totales")
            Else
                'Sql = Sql & " WHERE factcli_totales.numlinea is null"
                Sql = Sql & " WHERE false "
            End If
            Sql = Sql & " ORDER BY 1,2,3,4,5"
            
       
       
       Case 1 ' lineas de facturas
            tabla = "factcli_lineas"
            Sql = "SELECT factcli_lineas.numserie, factcli_lineas.numfactu, factcli_lineas.fecfactu, factcli_lineas.anofactu, factcli_lineas.numlinea, factcli_lineas.codmacta, cuentas.nommacta, factcli_lineas.baseimpo, factcli_lineas.codigiva,"
            Sql = Sql & " factcli_lineas.porciva, factcli_lineas.porcrec, factcli_lineas.impoiva, factcli_lineas.imporec, factcli_lineas.aplicret, IF(factcli_lineas.aplicret=1,'*','') as daplicret, factcli_lineas.codccost, ccoste.nomccost "
            Sql = Sql & " FROM (factcli_lineas LEFT JOIN ccoste ON factcli_lineas.codccost = ccoste.codccost) "
            Sql = Sql & " INNER JOIN cuentas ON factcli_lineas.codmacta = cuentas.codmacta "
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "factcli", "factcli_lineas")
            Else
                'Sql = Sql & " WHERE factcli_lineas.numlinea is null"
                Sql = Sql & " WHERE false"
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

Private Sub frmCob_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        CtaBanco = RecuperaValor(CadenaSeleccion, 1)
        IBAN = RecuperaValor(CadenaSeleccion, 2)
        
        Cobrado = RecuperaValor(CadenaSeleccion, 3)
        FechaCobro = RecuperaValor(CadenaSeleccion, 4)
    End If
End Sub

Private Sub frmConta_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
    
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
Dim vFe As String

    'Cuentas
    vFe = RecuperaValor(CadenaSeleccion, 3)
    If vFe <> "" Then
        vFe = RecuperaValor(CadenaSeleccion, 1)
        If EstaLaCuentaBloqueada2(vFe, CDate(Text1(1).Text)) Then
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
        If EstaLaCuentaBloqueada2(vFe, CDate(Text1(1).Text)) Then
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
        CadB = "numserie = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "N") & " and anofactu = year(" & DBSet(RecuperaValor(CadenaSeleccion, 3), "F") & ") "
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
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
    
    If Modo = 6 Then
        'TOTALES IVA
        txtaux3(0).Text = RecuperaValor(CadenaSeleccion, 1)
        
    Else
        'Lineas
        txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1)
        RC = "porcerec"
        txtAux(8).Text = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", txtAux(7), "N", RC)
        PonerFormatoDecimal txtAux(8), 4
        If RC = 0 Then
            txtAux(9).Text = ""
        Else
            txtAux(9).Text = RC
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
        If EstaLaCuentaBloqueada2(vFe, CDate(Text1(1).Text)) Then
            MsgBoxA "Cuenta bloqueada: " & vFe, vbExclamation
            If cmdAux(0).Tag = "0" Then txtAux(4).Text = ""
            Exit Sub
        End If
    End If
    If cmdAux(0).Tag = 0 Then
        'Cuenta normal
        txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
        
        'Habilitaremos el ccoste
        HabilitarCentroCoste
        
    Else
        'contrapartida
        txtAux(6).Text = RecuperaValor(CadenaSeleccion, 1)
    End If

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    'Centro de coste
    txtAux(12).Text = RecuperaValor(CadenaSeleccion, 1)
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
        
        Sql = "Select * from factcli "
        Sql = Sql & " WHERE numserie = " & RecuperaValor(CadenaSeleccion, 1)
        Sql = Sql & " AND numfactu =" & RecuperaValor(CadenaSeleccion, 2)
        Sql = Sql & " AND anofactu= " & RecuperaValor(CadenaSeleccion, 3)
        
        CadenaConsulta = Sql
        PonerCadenaBusqueda
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
Dim CuentaAnt As String

    If (Modo = 2 Or Modo = 5 Or Modo = 0) And (Index <> 6) And (Index <> 8) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        'FECHA FACTURA
        Indice = 1
        
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco Text1(1)
        
    Case 1 ' contadores
        Set frmConta = New frmBasico
        AyudaContadores frmConta, Text1(Index).Text, "tiporegi REGEXP '^[0-9]+$' = 0"
        Set frmConta = Nothing
        If Sql <> "" Then
            Text1(2).Text = RecuperaValor(Sql, 1)
            Text4(2).Text = RecuperaValor(Sql, 2)
            Text1_LostFocus 2
            PonFoco Text1(1)
        End If
        
    
    Case 2
        'Cuentas cliente
        CuentaAnt = Text1(4).Text
        Set frmCtas = New frmColCtas
        frmCtas.FILTRO = 1
        frmCtas.DatosADevolverBusqueda = "0|1|2|"
        frmCtas.ConfigurarBalances = 3  'NUEVO
        frmCtas.Show vbModal
        Set frmCtas = Nothing
        If Modo <> 1 Then
            If CuentaAnt <> Text1(4).Text Then Text1_LostFocus 4
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
        
        
        If IsNull(Data1.Recordset!NumDiari) Or IsNull(Data1.Recordset!FechaEnt) Then
            'NADA
            MsgBox "Diario /o fecha de apunte son nulos", vbExclamation
        Else
            Set frmAsi = New frmAsientosHco
            frmAsi.Asiento = DBLet(Data1.Recordset!NumDiari, "N") & "|" & Data1.Recordset!Fecfactu & "|" & Data1.Recordset!NumAsien & "|"
            frmAsi.SoloImprimir = True
            frmAsi.Show vbModal
            Set frmAsi = Nothing
        End If
       
        
    Case 7
        'Fecha de liquidacion
        If Text1(23).Enabled Then
            Indice = 23
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
        frmZ.Caption = "Observaciones Facturas Cliente"
        frmZ.Show vbModal
        Set frmZ = Nothing
        
    Case 9
        ' departamento
        Indice = 25
        
        Set frmDpto = New frmBasico
        AyudaDepartamentos frmDpto, Text1(Indice).Text, "codmacta = " & DBSet(Text1(4).Text, "T")
        Set frmDpto = Nothing
        PonFoco Text1(Indice)
        
        
        
    Case 10
        ' agente
        Set frmAgen = New frmBasico
        AyudaAgentes frmAgen
        Set frmAgen = Nothing
        
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
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
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
Dim I As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonFoco Text1(2) ' <===
        ' *** si n'hi han combos a la cap�alera ***
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
    
    If CadB = "" And CadB1 = "" Then Exit Sub
    HacerBusqueda2
    
End Sub

Private Sub HacerBusqueda2()

    CargarSqlFiltro
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia
    ElseIf CadB <> "" Or CadB1 <> "" Or cadFiltro <> "" Then
        CadenaConsulta = "select distinct factcli.* from " & NombreTabla
        'CadenaConsulta = CadenaConsulta & " INNER JOIN cuentas ON factcli.codmacta = cuentas.codmacta  "
        CadenaConsulta = CadenaConsulta & " left join factcli_lineas on factcli.numserie = factcli_lineas.numserie and factcli.numfactu = factcli_lineas.numfactu and factcli.anofactu = factcli_lineas.anofactu "
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
    
'    CargaDatosLW

End Sub


Private Sub MandaBusquedaPrevia()
Dim cWhere As String
Dim cWhere1 As String
    
    Screen.MousePointer = vbHourglass
    cWhere = "(numserie, numfactu, fecfactu) in (select factcli.numserie, factcli.numfactu, factcli.fecfactu from "
    cWhere = cWhere & "factcli LEFT JOIN factcli_lineas ON factcli.numserie = factcli_lineas.numserie and factcli.fecfactu = factcli_lineas.fecfactu and factcli.numfactu = factcli_lineas.numfactu "
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


    Set frmFact = New frmFacturasCliPrev
    
    frmFact.DatosADevolverBusqueda = "0|1|2|"
    frmFact.cWhere = cWhere
    frmFact.Show vbModal
    
    Set frmFact = Nothing
    

        
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBoxA "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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
    Combo1(5).ListIndex = 0
    
    If Now <= DateAdd("yyyy", 1, vParam.fechafin) Then
        Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Else
        Text1(1).Text = Format(vParam.fechafin - 1, "dd/mm/yyyy")
    End If
    Text1(9).Text = "0,00"
    
    FrameDatosFiscales.visible = False
    
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
        MsgBoxA Sql, vbExclamation
        Exit Sub
    End If
    
    
    If Not ComprobarPeriodo2(23) Then Exit Sub
    
    PonerModo 4

    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonFoco Text1(1)
    ' *********************************************************
    
    FecFactuAnt = Text1(1).Text
    
    
    
    NumDiario_ = 0
    ContabilizaApunte = True
    'Comprobamos que no esta actualizada ya
    If Not IsNull(Data1.Recordset!NumAsien) Then
        NumAsien = Data1.Recordset!NumAsien
        If NumAsien = 0 Then
            MsgBoxA "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
            Exit Sub
        End If
      
        
        If Val(DBLet(Data1.Recordset!no_modifica_apunte, "N")) = 1 Then
            ContabilizaApunte = False
            NumAsien = Data1.Recordset!NumAsien
            
        Else
            NumAsien = Data1.Recordset!NumAsien
            NumDiario_ = Data1.Recordset!NumDiari
        End If
    Else
        NumAsien = -1
    End If
        
        
     'Si viene a esta factura buscando por un campo k no sea clave entonces no le dejo seguir
    If InStr(1, Data1.Recordset.Source, "numasien") Then
        MsgBoxA "Busque la factura por su numero de factura", vbExclamation
        NumAsien = -1
        
    End If
    

    If NumAsien >= 0 Then
       
        'Tengo desintegrar la factura del hco
        If Not Desintegrar Then
            TerminaBloquear
            Exit Sub
        End If
        
        Text1(8).Text = ""
        If Not ContabilizaApunte Then Text1(8).Text = NumAsien
    Else
        PonerModo 2
        Exit Sub
    End If
    
    If Mc Is Nothing Then Set Mc = New Contadores
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    DespalzamientoVisible False
    PonFoco Text1(1)
    AntiguoText1 = ""
    
    
End Sub


Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim I As Long
    Dim Fec As Date
    Dim Mc As Contadores
    Dim SqlLog As String

    Dim CSV_SII As String
    Dim Id_ModifSII As Long



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
        MsgBoxA Sql, vbExclamation
        Exit Sub
    End If

    'Comprobamos si esta liquidado
    If Not ComprobarPeriodo2(23) Then Exit Sub
    
    
    'Si esta en SII
    'Si esta eliminado una factura presentada, singnifica que YA le ha dado a MODIFICAR
    CSV_SII = ""
    If vParam.SIITiene Then
        'SI esta presentada...
     
            If DBLet(Data1.Recordset!SII_ID, "N") > 0 Then
                If Text1(28).BackColor = vbSiiMofificando Then
                    CadenaDesdeOtroForm = ""
                    BuscaChekc = "numserie = " & DBSet(Data1.Recordset!NUmSerie, "T") & " AND factura_regis =" & Data1.Recordset!numfactu & " AND anofactu=" & Data1.Recordset!Anofactu & " AND estado=0 AND esfacturacliente" '=1 cliente
                    BuscaChekc = DevuelveDesdeBD("ID", "modificarsii", BuscaChekc, "1")
                    Id_ModifSII = Val(BuscaChekc)
                    If Id_ModifSII = 0 Then
                        MsgBoxA "No se encuetra registro en la tabla de modificacion de facturas SII(" & Id_ModifSII & ")", vbExclamation
                        
                    Else
                        frmMensajes.Opcion = 70
                        frmMensajes.Show vbModal
                    End If
                    If CadenaDesdeOtroForm = "" Then Exit Sub
                    CSV_SII = CStr(CadenaDesdeOtroForm)
                    CadenaDesdeOtroForm = ""
                End If
            End If

    End If
    

    
    
    
    
    'Comprobamos que no esta actualizada ya
    Sql = ""
    If Not IsNull(Data1.Recordset!NumAsien) Then
        Sql = "Esta factura ya esta contabilizada. "
    End If
    
    Sql = Sql & vbCrLf & vbCrLf & "Va usted a eliminar la factura :" & vbCrLf
    Sql = Sql & "Numero : " & Data1.Recordset!numfactu & vbCrLf
    Sql = Sql & "Fecha  : " & Data1.Recordset!Fecfactu & vbCrLf
    Sql = Sql & "Cliente : " & Me.Data1.Recordset!codmacta & " - " & Text4(4).Text & vbCrLf
    Sql = Sql & vbCrLf & "          �Desea continuar ?" & vbCrLf
    
    If Not EliminarDesdeActualizar Then
        If CSV_SII = "" Then
            If MsgBoxA(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    Screen.MousePointer = vbHourglass
    'Lo hara en actualizar
    I = 0
    If Not IsNull(Data1.Recordset!NumAsien) Then
        I = Data1.Recordset!NumAsien
        If Val(DBLet(Data1.Recordset!no_modifica_apunte, "N")) = 1 Then I = 0   'YA que no se tratan los apuntes
    End If
    If I > 0 Then
        
            'Memorizamos el numero de asiento y la fechaent para ver si devolvemos el contador
            'de asientos
            I = Data1.Recordset!NumAsien
            Fec = Data1.Recordset!FechaEnt
    
            'La borrara desde actualizar
            AlgunAsientoActualizado = False
        
            SqlLog = "Factura : " & CStr(DBLet(Data1.Recordset!NUmSerie)) & Format(Data1.Recordset!numfactu, "000000") & " " & Fec
            SqlLog = SqlLog & vbCrLf & "Cliente : " & Text1(4).Text & " " & Text4(4).Text
            SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text
            SqlLog = SqlLog & vbCrLf & "Asiento : " & I
        
        
            With frmActualizar
                .OpcionActualizar = 7
                .NumAsiento = Data1.Recordset!NumAsien
                .NumFac = Data1.Recordset!numfactu
                .FechaAsiento = Data1.Recordset!Fecfactu
                .NUmSerie = Data1.Recordset!NUmSerie & "|" & Data1.Recordset!Anofactu & "|"
                .NumDiari = Data1.Recordset!NumDiari
                .FechaAnterior = Data1.Recordset!Fecfactu
                .SqlLog = SqlLog
                .Show vbModal
            End With
                Set Mc = New Contadores
            Mc.DevolverContador "0", Fec <= vParam.fechafin, I
            Set Mc = Nothing
        
        
                    
                    
             If Id_ModifSII > 0 And CSV_SII <> "" Then
                'Cierro el proceso de eliminacion de facturas
                Sql = "UPDATE modificarsii SET estado=1, FechaHoraCierreProceso =" & DBSet(Now, "FH")
                Sql = Sql & ", Observaciones=concat(coalesce(observaciones,''),'  ','Eliminada en portal SII-AEAT')"
                Sql = Sql & " WHERE ID = " & Id_ModifSII
                I = 0
                If Not Ejecuta(Sql, False) Then I = 1
                
                
                BuscaChekc = DevuelveDesdeBD("SII_ID", "modificarsii", "ID", CStr(Id_ModifSII))
                If BuscaChekc = "" Then
                    I = 20
                Else
                    'ACtualizo SII
                    Sql = "UPDATE aswsii.envio_facturas_emitidas SET "
                    Sql = Sql & " Resultado = 'Eliminado', CSV =" & DBSet(CSV_SII, "T")
                    Sql = Sql & " , Mensaje =concat('Eliminado en portal-Ariconta('," & Id_ModifSII & ",')    ',coalesce(mensaje,''))"
                    Sql = Sql & " WHERE IDEnvioFacturasEmitidas =" & BuscaChekc
                    If Not Ejecuta(Sql, False) Then I = I + 30
                End If
                
                
                If I > 0 Then MsgBoxA "Error actualizando tablas auxiliar SII. Llame a soporte. Codigo:  " & I, vbExclamation
                    
            End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    Else
        'La borrara desde este mismo form
        Conn.BeginTrans
        
        I = Data1.Recordset!numfactu
        Fec = Data1.Recordset!Fecfactu
        If BorrarFactura Then
            'LOG
            SqlLog = "Factura : " & CStr(DBLet(Data1.Recordset!NUmSerie)) & Format(I, "000000") & " de fecha " & Fec
            SqlLog = SqlLog & vbCrLf & "Cliente : " & Text1(4).Text & " " & Text4(4).Text
            SqlLog = SqlLog & vbCrLf & "Importe : " & Text1(13).Text
            
            vLog.Insertar 6, vUsu, SqlLog
        
            AlgunAsientoActualizado = True
            Conn.CommitTrans
            Set Mc = New Contadores
            Mc.DevolverContador CStr(DBLet(Data1.Recordset!NUmSerie)), (Fec <= vParam.fechafin), I
            Set Mc = Nothing
            
            
            
            
            'MAYO 2018
            Sql = "select count(*) from cobros where numserie = " & DBSet(Data1.Recordset!NUmSerie, "T") & " and"
            Sql = Sql & " numfactu = " & Data1.Recordset!numfactu & " and fecfactu = " & DBSet(Data1.Recordset!Fecfactu, "F")
            Sql = Sql & " and impcobro <> 0 and not impcobro is null "
            
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Hay cobros que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
            Else

                
                Sql = "DELETE from cobros where  numserie = " & DBSet(Data1.Recordset!NUmSerie, "T") & " and"
                Sql = Sql & " numfactu = " & Data1.Recordset!numfactu & " and fecfactu = " & DBSet(Data1.Recordset!Fecfactu, "F")
                Ejecuta Sql
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
                For I = 1 To NumRegElim - 1
                    Data1.Recordset.MoveNext
                Next I
            End If
            PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function BorrarFactura() As Boolean
    
    On Error GoTo EBorrar
    Sql = " WHERE numserie = '" & Data1.Recordset!NUmSerie & "'"
    Sql = Sql & " AND numfactu = " & Data1.Recordset!numfactu
    Sql = Sql & " AND anofactu= " & Data1.Recordset!Anofactu
    'Las lineas
    AntiguoText1 = "DELETE from factcli_totales " & Sql
    Conn.Execute AntiguoText1
    AntiguoText1 = "DELETE from factcli_lineas " & Sql
    Conn.Execute AntiguoText1
    'La factura
    AntiguoText1 = "DELETE from factcli " & Sql
    Conn.Execute AntiguoText1
    
    ComprobarContador Data1.Recordset!NUmSerie, CDate(Text1(1).Text), Data1.Recordset!numfactu
    
    
    
        
    
    
    
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
Dim I As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    For I = 1 To DataGridAux.Count ' - 1
        If I <> 3 Then
            CargaGrid I, True
            If Not AdoAux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
        End If
    Next I
    
    imgppal(6).Enabled = (Text1(8).Text <> "")
    imgppal(6).visible = (Text1(8).Text <> "")
    If Text1(30).Text = "0,00" Then Text1(30).Text = ""
    If Text1(31).Text = "0,00" Then Text1(31).Text = ""
    Text4(2).Text = PonerNombreDeCod(Text1(2), "contadores", "nomregis", "tiporegi", "T")
    Text4(4).Text = PonerNombreDeCod(Text1(4), "cuentas", "nommacta", "codmacta", "T")
    Text4(6).Text = PonerNombreDeCod(Text1(6), "cuentas", "nommacta", "codmacta", "T")
    Text4(5).Text = PonerNombreDeCod(Text1(5), "formapago", "nomforpa", "codforpa", "N")
    Text4(21).Text = PonerNombreDeCod(Text1(21), "paises", "nompais", "codpais", "T")
    Text4(25).Text = DevuelveDesdeBDNew(cConta, "departamentos", "descripcion", "codmacta", Text1(4).Text, "T", , "dpto", Text1(25).Text, "N")
    Text4(26).Text = PonerNombreDeCod(Text1(26), "agentes", "nombre", "codigo", "N")
    
    If vParam.SIITiene Then Color_CampoSII
        
    If Text1(22).Text = "0" Then
        Combo1(0).ListIndex = 0
    Else
        PosicionarCombo Combo1(0), Asc(Text1(22).Text)
    End If
    
    'Combo1_Validate 1, False
    Combo1_Click 1
    
    If Text1(27).Text = "" Then
        Combo1(3).ListIndex = -1
    Else
        PosicionarCombo Combo1(3), Asc(Text1(27).Text)
    End If
    
    
    CargaDatosLW

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
    FrameModifIVA.visible = False
    If Modo = 2 And vUsu.Nivel = 0 Then
        If Val(Data1.Recordset!no_modifica_apunte) = 1 Then
            FrameModifIVA.visible = True
    
            ToolbarAuxTot.Buttons(2).Enabled = Me.lw1.ListItems.Count > 0
            ToolbarAuxTot.Buttons(3).Enabled = Me.lw1.ListItems.Count > 0
        End If
    End If

    
    
    
End Sub


Private Sub cmdCancelar_Click()
Dim I As Integer
Dim v

    Select Case Modo
        Case 1, 3 'B�squeda, Insertar
            'Contador de facturas
            If Modo = 3 Then
                'Intentetamos devolver el contador
                If Text1(0).Text <> "" Then
                    I = 0
                    If Text1(1).Text <> "" Then I = FechaCorrecta2(CDate(Text1(1).Text))
                    
                    Mc.DevolverContador Mc.TipoContador, I = 0, Mc.Contador
                End If
            End If
            LimpiarCampos
            PonerModo 0
            Set Mc = Nothing

        Case 4  'Modificar
            Modo = 2   'Para que el lostfocus NO haga nada
            If NumAsien > 0 Then
                'Ha cancelado. Tendre que situar los campos correctamente
                'Es decir. Anofacl
                Text1(1).Text = Data1.Recordset!Fecfactu
                Text1(0).Text = Data1.Recordset!numfactu
                Text1(14).Text = Data1.Recordset!Anofactu
                If Not IntegrarFactura(False) Then
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
                    If MsgBoxA("No se permite una factura sin l�neas " & vbCrLf & vbCrLf & "� Desea eliminar la factura ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
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
            If NumAsien > 0 Then
                'Ha cancelado. Tendre que situar los campos correctamente
                'Es decir. Anofacl
                Text1(1).Text = Data1.Recordset!Fecfactu
                Text1(0).Text = Data1.Recordset!numfactu
                Text1(14).Text = Data1.Recordset!Anofactu
                If Not IntegrarFactura(False) Then
                    Modo = 4 'lo pongo por si acaso
                    Exit Sub
                End If
                CobrosTesoreria
            Else
                ' cogemos un nro.de asiento para integrarlo
                Set Mc = New Contadores
                
                I = FechaCorrecta2(CDate(Text1(1).Text))
                If Mc.ConseguirContador("0", (I = 0), False) = 0 Then
                    Text1(8).Text = Format(Mc.Contador, "0000000")
                    NumAsien = Mc.Contador
                    ContabilizaApunte = True
                    If ModificaDesdeFormulario2(Me, 2, "Frame2") Then
                        If Not IntegrarFactura(False) Then
                            Modo = 4
                            Exit Sub
                        End If
                        CobrosTesoreria
                    End If
                Else
                    Mc.DevolverContador "0", (I = 0), CLng(Text1(8).Text)
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
'    Text1(22).Text = Combo1(0).ListIndex
    
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
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
            MsgBoxA varTxtFec, vbExclamation
        Else
            MsgBoxA "La fecha no pertenece al ejercicio actual ni al siguiente", vbExclamation
        End If
        B = False

    End If
    
    ' controles a�adidos de la factura de david
    'No puede tener % de retencion sin cuenta de retencion
    If ((Text1(6).Text = "") Xor (Text1(7).Text = "")) And Combo1(2).ListIndex > 0 Then
        MsgBoxA "No hay porcentaje de rentenci�n sin cuenta de retenci�n", vbExclamation
        B = False
        Exit Function
    End If
    
    'Compruebo si hay fechas bloqueadas
    
        If EstaLaCuentaBloqueada2(Text1(4).Text, CDate(Text1(1).Text)) Then
            MsgBoxA "Cuenta bloqueada: " & Text1(4).Text, vbExclamation
            B = False
            Exit Function
        End If
        If Text1(6).Text <> "" Then ' cuenta de retencion
            If EstaLaCuentaBloqueada2(Text1(6).Text, CDate(Text1(1).Text)) Then
                MsgBoxA "Cuenta bloqueada: " & Text1(6).Text, vbExclamation
                B = False
                Exit Function
            End If
        End If
    
    
    
    'Ahora. Si estamos modificando, y el a�o factura NO es el mismo, entonces
    'la estamos liando, y para evitar lios, NO dejo este tipo de modificacion
    If Modo = 4 Then
        If CDate(Text1(1).Text) <> Data1.Recordset!Fecfactu Then
            'HAN CAMBIADO LA FECHA. Veremos si dejo
            If Year(CDate(Text1(1).Text)) <> Data1.Recordset!Anofactu Then
                MsgBoxA "No puede cambiar de a�o la factura. ", vbExclamation
                B = False
                Exit Function
            End If
        End If
    End If
    
    
    'la forma de pago ha de existir
    If Text4(5).Text = "" And (Modo = 3 Or Modo = 4) Then
        MsgBoxA "No existe a forma de pago. Revise.", vbExclamation
        B = False
        PonFoco Text1(5)
        Exit Function
    End If
    
    'comprobamos que si la factura es intracomunitaria tiene que tener valor el tipo de intracomunitaria
    If Modo = 3 Or Modo = 4 Then
        If Combo1(1).ListIndex = 1 Then
            If Combo1(3).ListIndex = -1 Then
                MsgBoxA "Debe seleccionar un tipo de intracomunitaria. Revise.", vbExclamation
                B = False
                PonleFoco Combo1(3)
                Exit Function
            End If
        End If
    End If

    'Ticket ini y fin si resumen factura ticjet
    If B Then
        If Combo1(0).ListIndex = 2 Then
            cad = ""
            If Text1(32).Text = "" Then cad = "Ticket inicial "
            If Text1(33).Text = "" Then cad = cad & vbCrLf & "Ticket final"
            If cad <> "" Then
                cad = "Deberia indicar campos para : " & Combo1(0).Text & vbCrLf & cad
            Else
                If Len(Text1(32).Text) <> Len(Text1(33).Text) Then
                    cad = "Distinta longitud de factura ticket inicial /final"
                Else
                    If Text1(32).Text = Text1(33).Text Then cad = "Ticket inicial /final no pueden ser iguales"
                End If
            End If
            If cad <> "" Then
                cad = cad & vbCrLf & vbCrLf & "�Continuar?"
                If MsgBoxA(cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                    PonFoco Text1(33)
                    B = False
                End If
            Else
                I = Len(Text1(33).Text)
                If I > 6 Then
                    If Val(Mid(Text1(32).Text, 1, 3)) > Val(Mid(Text1(33).Text, 1, 3)) Then
                        cad = "z"
                    Else
                        If Val(Mid(Text1(32).Text, 4)) > Val(Mid(Text1(33).Text, 4)) Then cad = "K"
                    End If
                End If
                If cad <> "" Then
                    cad = "Ticket inicial parece mayor que el ticket final. �Continuar?"
                    If MsgBoxA(cad, vbQuestion + vbYesNoCancel) <> vbYes Then B = False
  
                End If
            End If
        End If
    End If
    
    
    If B And vParam.SIITiene And Modo = 4 Then
        ' modificando una FACTURA que hubiera sido presentada NO puede cambiar el Tipo De operacion Si est; es
        ' Factura resumen de ticket. Ya que cuando sube una factura esumen, el numero de factura realmente es el
        ' de ticket inicial
        
        cad = Chr(Combo1(0).ItemData(Combo1(0).ListIndex))
        
        
        If Data1.Recordset!codconce340 <> cad Then
            If Data1.Recordset!codconce340 = "B" Then
            
                'SOLO PARA TAXCO.
                'Alzira si que quiere cambiar de tipo de factura y pasarla a J-Ticket
                cad = ""
                If Not vParamT Is Nothing Then
                    '     de momento esto es TAXCO. Deberiamos poner un parametros que fuera "permite modificar fra resumen tickets
                    If vParamT.ImpresionRecibosConObserva Then cad = "NO"
                End If
                If cad <> "" Then
                    MsgBox "La factura era una FACTURA RESUMEN DE TICKETS." & vbCrLf & "No puede cambiarse", vbExclamation
                    B = False
                End If
            End If
        End If
    End If
    DatosOK = B

EDatosOK:
    If Err.Number <> 0 Then MsgBoxA Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(numserie=" & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N") & ") "
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
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
    vLog.Insertar 3, vUsu, Sql
       
    CobrosTesoreria
    
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
    If Index = 2 Then AntLetraSer = Text1(2).Text
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

Dim RC As String
Dim Correcto As Boolean
Dim Valor As Currency
Dim L As Long
Dim I As Integer
Dim J As Integer
Dim LeerCCuenta As Boolean
Dim Rs As ADODB.Recordset


    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    If (Index = 13 Or Index = 0 Or Index = 4) And Modo = 1 Then
        Text1(Index).BackColor = vbMoreLightBlue ' azul clarito
    End If

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 4, 5, 6
            
            Text4(Index) = ""
            If Index = 6 And Text1(Index).Text = "" Then Text1(7).Text = ""
            If Index = 4 And Text1(Index).Text = "" Then AntiguoText1 = ""
        Case 21
            Text4(Index) = ""
        Case 25
            Text4(Index) = ""
        Case 26
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
                MsgBoxA "Fecha incorrecta", vbExclamation
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
            If Index = 1 Then Text1(14).Text = Year(CDate(Text1(Index).Text))
            
            If Index = 1 And Modo <> 1 Then Text1(23).Text = Text1(1).Text
            
            'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
            'que el periodo de liquidacion del IVA no ha pasado.
            I = 0
            If vParam.IvaEnFechaPago Then
                If Index = 23 Then I = 1
                I = 1
            Else
                If Index = 1 Then I = 1
            End If
            If I > 0 And Modo <> 1 Then
                If Not ComprobarPeriodo2(Index) Then PonFoco Text1(Indice)
            End If

        Case 2 ' Serie
            If Modo = 1 Then Exit Sub
            If IsNumeric(Text1(Index).Text) Then
                MsgBoxA "Debe ser una letra: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                PonFoco Text1(2)
            End If
            Text1(Index).Text = UCase(Text1(Index).Text)
            If Text1(Index).Text = AntiguoText1 Then Exit Sub

            Text4(2).Text = DevuelveValor("select nomregis from contadores where tiporegi = " & DBSet(Text1(2).Text, "T") & " and tiporegi REGEXP '^[0-9]+$' = 0")
            If Text4(2).Text = "0" Then
                MsgBoxA "Letra de serie no existe o no es de facturas de cliente. Reintroduzca.", vbExclamation
                Text4(2).Text = ""
                Text1(2).Text = ""
                PonFoco Text1(2)
            Else
                If Modo = 3 Then
                    ' traemos el contador
                    If Text1(2).Text <> AntLetraSer Then
                        If Text1(1).Text <> "" Then I = FechaCorrecta2(CDate(Text1(1).Text))
                        If Mc.ConseguirContador(Trim(Text1(2).Text), (I = 0), False) = 0 Then
                            'COMPROBAR NUMERO ASIENTO
                            Text1(0).Text = Mc.Contador
                                        
                    
                            Sql = "select codconce340 from contadores where tiporegi = " & DBSet(Text1(2).Text, "T")
                            Set Rs = New ADODB.Recordset
                            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If Not Rs.EOF Then
                                If DBLet(Rs.Fields(0).Value, "T") <> "" Then
                                    PosicionarCombo Combo1(0), Asc(DBLet(Rs.Fields(0).Value, "T"))
                                Else
                                    Combo1(0).ListIndex = 0
                                End If
                            End If
                            Set Rs = Nothing
                        End If
                    End If
                End If
            End If
        Case 3
            If Len(Text1(Index).Text) > 0 Then PonCursorInicio
        Case 4, 6 ' cuenta de cliente, cuenta de retencion
                'Cuenta cliente
                If AntiguoText1 = Text1(Index).Text Then
                    If Text4(4).Text <> "" Then Exit Sub
                End If
                RC = Text1(Index).Text
                I = Index
                
                
                If CuentaCorrectaUltimoNivel(RC, Sql) Then
                    If Modo = 1 Then Sql = ""
                    Text1(Index).Text = RC
                    Text4(I).Text = Sql
                    If Text1(1).Text <> "" Then
                        If Modo > 2 Then
                            If EstaLaCuentaBloqueada2(RC, CDate(Text1(1).Text)) Then
                                MsgBoxA "Cuenta bloqueada: " & RC, vbExclamation
                                Text1(Index).Text = ""
                                Text4(I).Text = ""
                                PonFoco Text1(Index)
                                Exit Sub
                            End If
                        End If
                    End If
                    If Index = 4 Then
                        LeerCCuenta = False
                        If Modo = 3 Then
                            If Text1(Index).Text <> AntiguoText1 Then LeerCCuenta = True
                        Else
                            If Modo = 4 Then
                                If AntiguoText1 = "" Then
                                    If Text1(Index).Text <> Data1.Recordset!codmacta Then LeerCCuenta = True
                                Else
                                    If Trim(Text1(Index).Text) <> AntiguoText1 Then LeerCCuenta = True
                                End If
                            End If
                        End If
                        If LeerCCuenta Then
                            CargarDatosCuenta Text1(Index)
                            AntiguoText1 = Text1(Index).Text
                        End If
                    End If
                    RC = ""
                Else
                    
                    If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                            RC = RellenaCodigoCuenta(Text1(Index).Text)
                            Sql = "La cuenta: " & RC & " no existe.       �Desea crearla?"
                            If MsgBoxA(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
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
                            MsgBoxA Sql, vbExclamation
                        End If
                    End If
                    
                    If Sql <> "" Then
                        Text1(Index).Text = ""
                        Text4(I).Text = ""
                        PonFoco Text1(Index)
                    End If
                    
                    
                End If
        
        
        Case 5 ' forma de pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text4(Index).Text = PonerNombreDeCod(Text1(Index), "formapago", "nomforpa", "codforpa", "N")
                If Text4(Index).Text = "" Then
                    MsgBoxA "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
        
        Case 7 ' % de retencion
            PonerFormatoDecimal Text1(Index), 4
        
        Case 21 ' codigo de pais
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                Text4(Index).Text = PonerNombreDeCod(Text1(Index), "paises", "nompais", "codpais", "T")
                If Text4(Index) = "" Then
                    MsgBoxA "No existe el Pa�s. Reintroduzca.", vbExclamation
                    Text1(Index).Text = ""
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
        
        Case 25 ' departamento
            If Text1(Index).Text <> "" Then
                Text4(Index).Text = DevuelveDesdeBDNew(cConta, "departamentos", "descripcion", "codmacta", Text1(4).Text, "T", , "dpto", Text1(25).Text, "N")
                If Text4(Index) = "" Then
                    MsgBoxA "No existe el Departamento de este Cliente. Reintroduzca.", vbExclamation
                    Text1(Index).Text = ""
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
            End If
            
        Case 26 ' agente
            If Text1(Index).Text <> "" Then
                Text4(Index).Text = PonerNombreDeCod(Text1(Index), "agentes", "nombre", "codigo", "N")
                If Text4(Index) = "" Then
                    MsgBoxA "No existe el Agente. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text4(Index).Text = ""
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
                Case 4:  KEYBusqueda KeyAscii, 2 ' cuenta cliente
                Case 6:  KEYBusqueda KeyAscii, 4 ' cuenta de retencion
                Case 5:  KEYBusqueda KeyAscii, 3 ' forma de pago
                Case 2:  KEYBusqueda KeyAscii, 1 ' serie
                Case 21: KEYBusqueda KeyAscii, 5 ' pais
                Case 25: KEYBusqueda KeyAscii, 9 ' departamento
                Case 26: KEYBusqueda KeyAscii, 10 ' agente
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
Dim Sigo As Byte ' 0 salir   1. Modificar   2 Subir de nuevo

    Select Case Button.Index
    
        Case 1 'Datos Fiscales
            Me.FrameDatosFiscales.visible = Not Me.FrameDatosFiscales.visible
           
        Case 2 'Cartera de Cobros
            If Not Data1.Recordset.EOF Then
                Set frmMens = New frmMensajes
                
                frmMens.Opcion = 27
                frmMens.Parametros = Trim(Text1(2).Text) & "|" & Trim(Text1(0).Text) & "|" & Text1(1).Text & "|"
                frmMens.Show vbModal
                
                Set frmMens = Nothing
            End If
    
        Case 3
            Screen.MousePointer = vbHourglass
            
            Set frmUtil = New frmUtilidades
            
            frmUtil.Opcion = 5
            frmUtil.Show vbModal

            Set frmUtil = Nothing
            
        Case 4
'            CadFacErr = "(numasien = 0 or numasien is null or fechaent is null or numdiari is null)"
'
'            HacerBusqueda
            ComprobarFrasSinAsiento
            
            
        Case 5
            'Quiere MODIFICAR factura presentada en SII
            
            If Modo <> 2 Then Exit Sub
            If vUsu.Nivel > 1 Then Exit Sub
            
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

            Sigo = 0
            If Text1(28).BackColor = vbSiiMofificando Then
                'Dejo pasar pq va a cerrar el proceso y volver a subir al SII
                Sigo = 2
            Else
                If Text1(28).BackColor = vbSiiOK Then Sigo = 1
                If Text1(28).BackColor = vbSiiAceptadaConErr Then Sigo = 1
            End If
            
            If Sigo = 0 Then Exit Sub
            
            
            
            
            
            BuscaChekc = "numserie = " & DBSet(Data1.Recordset!NUmSerie, "T") & " AND numfactu =" & Data1.Recordset!numfactu & " AND anofactu=" & Data1.Recordset!Anofactu
            
            frmSiiPreparaModificar.Escliente = True
            frmSiiPreparaModificar.where = BuscaChekc
            frmSiiPreparaModificar.AbrirProceso = Sigo = 1
            CadenaDesdeOtroForm = ""
            frmSiiPreparaModificar.Show vbModal
            
            If CadenaDesdeOtroForm <> "" Then
                'OK Se han modificado cosas. Modificamos y volvemos a cargar
                Screen.MousePointer = vbHourglass
                PosicionarData
                Screen.MousePointer = vbDefault
            End If
            
    End Select

End Sub


Private Sub ComprobarFrasSinAsiento()
Dim Sql As String
Dim vCadena As String
Dim vCadena2 As String
Dim Rs As ADODB.Recordset
Dim IntegrarFactura As Boolean
Dim I As Integer
Dim Nregs As Long
Dim SqlLog As String

    
    Sql = "select * from factcli where (numasien = 0 or numasien is null or fechaent is null or numdiari is null) "
    Sql = Sql & " AND fecfactu >=" & DBSet(vParam.fechaini, "F")
    If cadFiltro <> "" Then Sql = Sql & " and " & cadFiltro

    vCadena = ""
    vCadena2 = ""
    NumDiario_ = 0
    If TotalRegistrosConsulta(Sql) <> 0 Then
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Nregs = 1
        
        While Not Rs.EOF
            vCadena = vCadena & "Fra. " & DBLet(Rs!NUmSerie) & " " & Format(DBLet(Rs!numfactu), "0000000") & " " & DBLet(Rs!Fecfactu, "F")
            vCadena2 = vCadena2 & "(" & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!numfactu, "N") & "," & Year(DBLet(Rs!Fecfactu, "F")) & "),"
            
            If (Nregs Mod 2) = 0 Then
                vCadena = vCadena & vbCrLf
            Else
                vCadena = vCadena & "  "
            End If
            
            Nregs = Nregs + 1
            
            Rs.MoveNext
        Wend
        
        If MsgBoxA("Las siguientes facturas no tienen Asiento asociado: " & vbCrLf & vbCrLf & vCadena & vbCrLf & vbCrLf & " � Asigna asiento ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            Rs.MoveFirst
            
            While Not Rs.EOF
                IntegrarFactura = False
                
                ' cogemos un nro.de asiento para integrarlo
                Set Mc = New Contadores
                
                I = FechaCorrecta2(CDate(DBLet(Rs!Fecfactu, "F")))
                If Mc.ConseguirContador("0", (I = 0), False) = 0 Then
                    
                    NumAsien = Mc.Contador
                
                    SqlLog = "Factura : " & Rs!NUmSerie & " " & Rs!numfactu & " de fecha " & Rs!Fecfactu
                    SqlLog = SqlLog & vbCrLf & "Cuenta  : " & DBLet(Rs!codmacta, "T") & " " & DBLet(Rs!Nommacta, "T")
                    SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Rs!totfaccl, "N")
                    
                    With frmActualizar
                        .OpcionActualizar = 6
                        'NumAsiento     --> CODIGO FACTURA
                        'NumDiari       --> A�O FACTURA
                        'NUmSerie       --> SERIE DE LA FACTURA
                        'FechaAsiento   --> Fecha factura
                        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
                        .NumFac = DBLet(Rs!numfactu, "N")
                        .NumDiari = Year(DBLet(Rs!Fecfactu, "F"))
                        .NUmSerie = Rs!NUmSerie
                        .FechaAsiento = DBLet(Rs!Fecfactu, "F")
                        .FechaAnterior = DBLet(Rs!Fecfactu, "F")
                        .SqlLog = "" 'SqlLog
                        
                        If NumDiario_ <= 0 Then NumDiario_ = vParam.numdiacl
                        .DiarioFacturas = NumDiario_
                        .NumAsiento = NumAsien
                        .Show vbModal
                        
                        If AlgunAsientoActualizado Then IntegrarFactura = True
                        
                        Screen.MousePointer = vbHourglass
                        Me.Refresh
                    End With
                
                    If IntegrarFactura Then
                        Sql = "update factcli set numdiari = " & DBSet(NumDiario_, "N") & ", fechaent = " & DBSet(Rs!Fecfactu, "F") & ", "
                        Sql = Sql & " numasien = " & DBSet(NumAsien, "N") & " where numserie = " & DBSet(Rs!NUmSerie, "T") & " and anofactu = year("
                        Sql = Sql & DBSet(Rs!Fecfactu, "F") & ") and numfactu = " & DBSet(Rs!numfactu, "N")
                    
                        Conn.Execute Sql
                        
                        
                    End If
                End If
                
                Rs.MoveNext
            Wend
        
            vLog.Insertar 28, vUsu, vCadena
        
        End If
        
        CadB = "(factcli.numserie,factcli.numfactu,factcli.anofactu) in (" & Mid(vCadena2, 1, Len(vCadena2) - 1) & ")"
        HacerBusqueda2
        
        Set Rs = Nothing
    
    Else
        MsgBoxA "No hay facturas sin asiento asignado.", vbInformation
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
        MsgBoxA Sql, vbExclamation
        Exit Sub
    End If
    
    'Marzo 2019
    ' Si no modifica el apunte, NO dejo tocar lineas
    If Val(DBLet(Data1.Recordset!no_modifica_apunte, "N")) = 1 Then
        MsgBoxA "Factura integrada. Modifique el apunte", vbExclamation
        Exit Sub
    End If
    
    '**** parte correspondiente por si la factura est� contabilizada
    NumDiario_ = 0
    'Comprobamos que no esta actualizada ya
    If Not IsNull(Data1.Recordset!NumAsien) Then
        NumAsien = Data1.Recordset!NumAsien
        If NumAsien = 0 Then
            MsgBoxA "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
            Exit Sub
        End If
        
        ContabilizaApunte = True
        If Val(DBLet(Data1.Recordset!no_modifica_apunte, "N")) = 1 Then ContabilizaApunte = False
        
        NumAsien = Data1.Recordset!NumAsien
        NumDiario_ = DBLet(Data1.Recordset!NumDiari, "N")
    Else
        NumAsien = -1
    End If
    
    If Not ComprobarPeriodo2(23) Then Exit Sub
    
    'Llegados aqui bloqueamos desde form
    '--If Not BloqueaRegistroForm(Me) Then Exit Sub
    If Not BLOQUEADesdeFormulario2(Me, Data1, 1) Then Exit Sub
    
    FecFactuAnt = Text1(1).Text
    

    If NumAsien >= 0 Then
        'Tengo desintegrar la factura del hco
        If Not Desintegrar Then
            '--DesBloqueaRegistroForm Me.Text1(0)
            TerminaBloquear
            Exit Sub
        End If
        Text1(8).Text = ""
        If Not ContabilizaApunte Then Text1(8).Text = NumAsien
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
            CobrosTesoreria
            
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
                Sql = varTxtFec
            Else
                Sql = "La factura pertenece a un ejercicio cerrado."
            End If
            MsgBoxA Sql, vbExclamation
            Exit Sub
        End If
        
        
        If Not ComprobarPeriodo2(23) Then Exit Sub
    
    If Button.Index = 3 Then
        'Eliminar blabla bla
        If MsgBoxA("�Seguro que desea eliminar la linea de IVA seleccionada?", vbQuestion + vbYesNoCancel) = vbYes Then
            Sql = ObtenerWhereCab(True)
            Sql = Replace(Sql, "factcli.", "factcli_totales.")
            Sql = Sql & " AND numlinea = " & lw1.SelectedItem.Text
            Sql = "DELETE from factcli_totales " & Sql
            
            If Ejecuta(Sql, False) Then
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
            TxtAux3_LostFocus 0
        End If
        
        PonFoco txtaux3(alto)
        
    End If
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    If False Then
        Caption = Linliapu + 1
        For I = 0 To 4
            Toolbar2.Buttons(I + 1).Image = Linliapu + I
        Next
        Linliapu = Linliapu + 5
        Exit Sub
    End If
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
            Sql = "�Seguro que desea eliminar la l�nea de la factura?"
            Sql = Sql & vbCrLf & "Serie: " & AdoAux(Index).Recordset!NUmSerie & " - " & AdoAux(Index).Recordset!numfactu & " - " & AdoAux(Index).Recordset!Fecfactu & " - " & AdoAux(Index).Recordset!NumLinea
            If MsgBoxA(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM factcli_lineas "
                Sql = Sql & Replace(vWhere, "factcli", "factcli_lineas") & " and numlinea = " & DBLet(AdoAux(Index).Recordset!NumLinea, "N")
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute Sql
        
        RecalcularTotales
        
        '**** parte de contabilizacion de la factura
        '--DesBloqueaRegistroForm Me.Text1(0)
        TerminaBloquear
        
        If NumAsien > 0 Then
            If IntegrarFactura(False) Then
                Text1(8).Text = Format(NumAsien, "0000000")
                NumAsien = -1
                NumDiario_ = 0
            Else
                B = False
            End If
        End If
        
        'LOG
        
        SqlLog = "Factura : " & Text1(2).Text & Text1(0).Text & " " & Text1(1).Text & " L�nea : " & DBLet(Me.AdoAux(1).Recordset!NumLinea, "N")
        SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Me.AdoAux(1).Recordset!Baseimpo, "N")
        
        vLog.Insertar 8, vUsu, SqlLog
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
     Screen.MousePointer = vbDefault
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Sub BotonAnyadirLinea(Index As Integer, Limpia As Boolean)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer

    ModoLineas = 1 'Posem Modo Afegir Ll�nia

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 1: vTabla = "factcli_lineas"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 1   'hlinapu
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = ""
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", Replace(vWhere, "factcli", "factcli_lineas"))
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
                        For I = 0 To txtAux.Count - 1
                            txtAux(I).Text = ""
                        Next I
                    End If
                    txtAux(0).Text = Text1(2).Text 'serie
                    txtAux(1).Text = Text1(0).Text 'numfactu
                    txtAux(2).Text = Text1(1).Text 'fecha
                    txtAux(3).Text = Text1(14).Text 'anofactura
                    
                    txtAux(4).Text = Format(NumF, "0000") 'linea contador
                    
                    
                    If Limpia Then
                        txtAux2(5).Text = ""
                        txtAux2(12).Text = ""
                    End If
                    
                    ' antes si hay retencion se marca como que hay que aplicarle retencion
                    chkAux(0).Value = 1
                    
                    If Limpia Then
                        PonFoco txtAux(5)
                    Else
                        PonFoco txtAux(5)
                    End If
            
                    ' traemos la cuenta de contrapartida habitual
                    PonFoco txtAux(5)

                    txtAux(5).Text = CuentaHabitual(Text1(4).Text)
                    If txtAux(5).Text <> "" Then
                        If EstaLaCuentaBloqueada2(txtAux(5).Text, CDate(Text1(1).Text)) Then
                            txtAux(5).Text = ""
                        Else
                            If Not ExisteEnFactura(Text1(2).Text, Text1(0).Text, Text1(1).Text, txtAux(5).Text) Then
                                txtAux_LostFocus (5)
                                PonFoco txtAux(7)
                                txtAux_LostFocus (7)
                                PonFoco txtAux(6)
                            Else
                                txtAux(5).Text = ""
                                PonFoco txtAux(5)
                            End If
                        End If
                        
                    End If
            
            End Select

    End Select
End Sub

Private Function ExisteEnFactura(Serie As String, numfactu As String, Fecfactu As String, Cuenta As String) As Boolean
Dim Sql As String

    ExisteEnFactura = False
    
    If Serie = "" Or numfactu = "" Or Fecfactu = "" Or Cuenta = "" Then Exit Function

    Sql = "select count(*) from factcli_lineas where numserie = " & DBSet(Serie, "T") & " and numfactu = " & DBSet(numfactu, "N")
    Sql = Sql & " and fecfactu = " & DBSet(Fecfactu, "F") & " and codmacta = " & DBSet(Cuenta, "T")

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
    Dim I As Integer
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
    ' *** bloqueje la clau primaria de la cap�alera ***
'    BloquearTxt Text1(0), True
    ' *********************************

    Select Case Index
        Case 0, 1 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
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
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux(3).Text = DataGridAux(Index).Columns(3).Text
            txtAux(4).Text = DataGridAux(Index).Columns(4).Text
            
            txtAux(5).Text = DataGridAux(Index).Columns(5).Text 'cuenta
            txtAux2(5).Text = DataGridAux(Index).Columns(6).Text 'denominacion
            txtAux(6).Text = DataGridAux(Index).Columns(7).Text 'baseimpo
            txtAux(7).Text = DataGridAux(Index).Columns(8).Text 'codigiva
            txtAux(8).Text = DataGridAux(Index).Columns(9).Text '%iva
            txtAux(9).Text = DataGridAux(Index).Columns(10).Text '%retencion
            txtAux(10).Text = DataGridAux(Index).Columns(11).Text 'importe iva
            txtAux(11).Text = DataGridAux(Index).Columns(12).Text 'importe retencion
            If DataGridAux(Index).Columns(13).Text = 1 Then
                chkAux(0).Value = 1 ' DataGridAux(Index).Columns(14).Text 'aplica retencion
            Else
                chkAux(0).Value = 0
            End If
            txtAux(12).Text = DataGridAux(Index).Columns(15).Text 'centro de coste
            txtAux2(12).Text = DataGridAux(Index).Columns(16).Text 'nombre centro de coste
                 
            IvaCuenta = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtAux(5).Text, "N")
            HaCambiadoTipoIva = False
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    HabilitarCentroCoste
    
    PonFoco txtAux(5)
    
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 2 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 1 'lineas de factura
            For jj = 5 To txtAux.Count - 1
                txtAux(jj).visible = B
                txtAux(jj).top = alto
            Next jj
            
            txtAux2(5).visible = B
            txtAux2(5).top = alto
            txtAux2(12).visible = B
            txtAux2(12).top = alto
            
            chkAux(0).visible = B
            chkAux(0).top = alto
            
            For jj = 0 To 2
                cmdAux(jj).visible = B
                cmdAux(jj).top = txtAux(5).top
                cmdAux(jj).Height = txtAux(5).Height
            Next jj
            
            If Not vParam.autocoste Then
                cmdAux(2).visible = False
                cmdAux(2).Enabled = False
                txtAux(12).visible = False
                txtAux(12).Enabled = False
                txtAux2(12).visible = False
                txtAux2(12).Enabled = False
            End If
            
            'Los campos importes de IVA siempre bloqueados excepto que el parametro indique lo contrario
            B = Not B
            If Not B Then
                If Not vParam.ModificarIvaLineasFraCli Then B = True
            End If
            
            BloqueaTXT txtAux(10), B
            BloqueaTXT txtAux(11), B
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
Dim Sql As String
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
    If Not vParam.autocoste Then txtAux(12).Text = ""

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And (Modo = 5 And ModoLineas = 1) Then  'insertar
    
    End If
    
    If B And Modo = 5 Then ' tanto si insertamos como si modificamos en lineas
        'Cuenta
        If txtAux(5).Text = "" Then
            MsgBoxA "Cuenta no puede estar vacia.", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(5)
            Exit Function
        End If
        
        If Not IsNumeric(txtAux(5).Text) Then
            MsgBoxA "Cuenta debe ser numrica", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(5)
            Exit Function
        End If
        
        If txtAux(5).Text = NO Then
            MsgBoxA "La cuenta debe estar dada de alta en el sistema", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(5)
            Exit Function
        End If
        
        If Not EsCuentaUltimoNivel(txtAux(5).Text) Then
            MsgBoxA "La cuenta no es de �ltimo nivel", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(5)
            Exit Function
        End If
        
        If IvaCuenta = "" Then
            CambiarIva = True
        Else
        
             If ModoLineas = 1 Then
                'Enero 2020
                'QUito esta pregunta
                CambiarIva = False
                If False Then
                    If CInt(ComprobarCero(txtAux(7).Text)) <> CInt(ComprobarCero(IvaCuenta)) Then
                        If MsgBoxA("El c�digo de iva es distinto del de la cuenta. " & vbCrLf & " � Desea modificarlo en la cuenta ? " & vbCrLf & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            CambiarIva = True
                        Else
                            CambiarIva = False
                        End If
                    End If
                End If
            End If
        End If
        
        'Centro de coste
        If txtAux(12).visible Then
            If txtAux(12).Enabled Then
                If txtAux(12).Text = "" Then
                    MsgBoxA "Centro de coste no puede ser nulo", vbExclamation
                    PonFoco txtAux(12)
                    Exit Function
                End If
            End If
        End If
        
        
    End If
    
    
    
    
    
    'Como puede modificar los IVA, hay que comprobar
    If B And vParam.ModificarIvaLineasFraCli Then
        
        Importe = ImporteFormateado(txtAux(8).Text) / 100
        Importe = ImporteFormateado(txtAux(6).Text) * Importe
        
        
        
        If Abs(Importe - ImporteFormateado(txtAux(10).Text)) >= 0.1 Then
            Mens = "Iva calculado: " & Format(Importe, FormatoImporte) & vbCrLf
            Mens = Mens & "Iva introducido: " & txtAux(10).Text & vbCrLf
            Mens = "DIFERENCIAS EN IVA" & vbCrLf & vbCrLf & Mens & vbCrLf & "�Desea continuar igualmente?"
            
            If MsgBoxA(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
        End If
        
        If B Then
            If Me.txtAux(9).Text <> "" Then
                'REGARCO
                
                Importe = ImporteFormateado(txtAux(9).Text) / 100
                Importe = ImporteFormateado(txtAux(6).Text) * Importe
                If Abs(Importe - ImporteFormateado(txtAux(11).Text)) >= 0.05 Then
                    Mens = "Iva calculado: " & Format(Importe, FormatoImporte) & vbCrLf
                    Mens = Mens & "Iva introducido: " & txtAux(11).Text & vbCrLf
                    Mens = "DIFERENCIAS EN RECARGO EQUIVALENCIA" & vbCrLf & vbCrLf & Mens & vbCrLf & "�Desea continuar igualmente?"
                    
                    If MsgBoxA(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
                End If
                
            End If
        End If
    End If
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBoxA Err.Number & ": " & Err.Description, vbExclamation
End Function

    
Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    SepuedeBorrar = True
End Function


' *********************************************************************************
Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

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
Dim I As Byte
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

    For I = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(I).AllowSizing = False
    Next I
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
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
                For I = 0 To 4
                    txtAux(I).Text = ""
                Next I
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
            
            If B Then RecalcularTotalesFactura (False)
        
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

                    txtAux(11).Text = ""
                    If Limp Then
                        txtAux2(5).Text = ""
                        txtAux2(12).Text = ""
                        For I = 0 To 11
                            txtAux(I).Text = ""
                        Next I
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
    
    Sql = "update cuentas set codigiva = " & DBSet(txtAux(7).Text, "N") & " where codmacta = " & DBSet(txtAux(5).Text, "T")
    Conn.Execute Sql
    
    ActualizarIva = True
    Exit Function
    
eActualizarIva:
    MuestraError Err.Number, "Actualizar Iva", Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim v As Integer
Dim cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'apuntes
    End Select
    ' **************************************************************

    ModificarLinea = False
    If DatosOkLlin(nomframe) Then


        TerminaBloquear
        Conn.BeginTrans
        
        B = True
        If CambiarIva Then B = ActualizarIva
        
        If B And ModificaDesdeFormulario2(Me, 2, nomframe) Then
        
            B = RecalcularTotales
            
            'LOG
            vLog.Insertar 7, vUsu, "Factura : " & Text1(2).Text & Text1(0).Text & " " & Text1(1).Text & " Linea : " & txtAux(4).Text
        
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
            ModificarLinea = True
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
    vWhere = vWhere & "factcli.numserie=" & DBSet(Text1(2).Text, "T") & " and factcli.numfactu=" & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
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
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And Modo = 2
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!Especial, "N") And (Modo <> 0 And Modo <> 5)
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!Especial, "N") And Modo = 2 And vEmpresa.TieneTesoreria
        Me.Toolbar2.Buttons(3).Enabled = DBLet(Rs!Especial, "N") And (Modo = 2 Or Modo = 0)
        Me.Toolbar2.Buttons(4).Enabled = DBLet(Rs!Especial, "N") And (Modo = 2 Or Modo = 0)
        Me.Toolbar2.Buttons(5).Enabled = DBLet(Rs!Especial, "N") And (Modo = 2)
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
        ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        ToolbarAux.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        
        
        
        vUsu.LeerFiltros "ariconta", IdPrograma
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    AntiguoText1 = txtAux(Index).Text
    ConseguirFoco txtAux(Index), Modo
    
    If Index = 11 Then
        If ComprobarCero(txtAux(9).Text) = 0 Then
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
    Dim CodCCost As String
        If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
        
        If txtAux(Index).Text = AntiguoText1 Then
             If Index = 12 And vParam.autocoste Then PonleFoco cmdAceptar
             Exit Sub
        End If
    
        CalcularElIva = True
        Select Case Index
        Case 5
            RC = txtAux(5).Text
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                txtAux(5).Text = RC
                If Modo = 1 Then Exit Sub
                If EstaLaCuentaBloqueada2(RC, CDate(Text1(1).Text)) Then
                    MsgBoxA "Cuenta bloqueada: " & RC, vbExclamation
                    txtAux(5).Text = ""
                Else
                    txtAux2(5).Text = Sql
                    ' traemos el tipo de iva de la cuenta
                    If ModoLineas = 1 Then
                        CodCCost = "ccosteDef"
                        txtAux(7).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtAux(5).Text, "N", CodCCost)
                        IvaCuenta = txtAux(7)
                        If txtAux(7).Text <> "" Then txtAux_LostFocus (7)
                        If vParam.autocoste Then
                            If CodCCost <> "" Then
                                txtAux(12).Text = UCase(CodCCost)
                                Sql = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtAux(12).Text, "T")
                                txtAux2(12).Text = Sql
                            End If
                        End If
                    Else
                        CalcularElIva = False
                    End If
                    
                    RC = ""
                End If
            Else
                If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                    'NO EXISTE LA CUENTA, a�ado que debe de tener permiso de creacion de cuentas
                    If vUsu.PermiteOpcion("ariconta", 201, vbOpcionCrearEliminar) Then
                        Sql = Sql & " �Desea crearla?"
                        If MsgBoxA(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                            CadenaDesdeOtroForm = RC
                            cmdAux(0).Tag = Index
                            Set frmC = New frmColCtas
                            frmC.DatosADevolverBusqueda = "0|1|"
                            frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                            frmC.Show vbModal
                            Set frmC = Nothing
                            If txtAux(5).Text = RC Then
                                Sql = "" 'Para k no los borre
                                ' traemos el tipo de iva de la cuenta
                                txtAux(7).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtAux(5).Text, "N")
                                IvaCuenta = txtAux(7)
                                txtAux_LostFocus (7)
                            End If
                        End If
                    Else
                        MsgBoxA Sql, vbExclamation
                    End If
                Else
                    MsgBoxA Sql, vbExclamation
                End If
                    
                If Sql <> "" Then
                  txtAux(5).Text = ""
                  txtAux2(5).Text = ""
                  RC = "NO"
                End If
            End If
            HabilitarCentroCoste
            If RC <> "" Then PonFoco txtAux(5)
                
            If Modo = 5 And ModoLineas = 1 Then MostrarObservaciones txtAux(Index)
            
        Case 6
            If Not PonerFormatoDecimal(txtAux(Index), 1) Then
                txtAux(Index).Text = ""
            Else
                'Si modificando lienas, no cambia el importe NO recalculo iVA
                If Modo = 5 And ModoLineas = 2 Then
                    If ImporteFormateado(txtAux(Index).Text) = CCur(DBLet(AdoAux(1).Recordset!Baseimpo, "N")) Then CalcularElIva = False
                    
                End If
            End If
            
        Case 7 ' iva
            RC = "porcerec"
            txtAux(8).Text = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", txtAux(7), "N", RC)
            If txtAux(8).Text = "" Then
                MsgBoxA "No existe el Tipo de Iva. Reintroduzca.", vbExclamation
                If txtAux(7).Text <> "" Then txtAux(7) = "": PonFoco txtAux(7)
            Else
                If RC = 0 Then
                    txtAux(9).Text = ""
                Else
                    txtAux(9).Text = RC
                End If
            End If
            
             If Modo = 5 And ModoLineas = 2 Then
                If txtAux(7).Text <> "" Then
                    If Val(txtAux(Index).Text) = Val(DBLet(AdoAux(1).Recordset!codigiva, "N")) Then
                        If Not HaCambiadoTipoIva Then CalcularElIva = False
                    Else
                        HaCambiadoTipoIva = True
                    End If
                End If
            End If
            
            
            
        Case 10, 11
           'LOS IMPORTES
            If PonerFormatoDecimal(txtAux(Index), 1) Then
                If Not vParam.autocoste Then PonleFoco cmdAceptar
            End If
                
        Case 12
'            If txtAux(Index).Text = "" Then Exit Sub
            
            txtAux(12).Text = UCase(txtAux(12).Text)
            Sql = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtAux(12).Text, "T")
            txtAux2(12).Text = ""
            If Sql = "" Then
                MsgBoxA "Concepto NO encontrado: " & txtAux(12).Text, vbExclamation
                txtAux(12).Text = ""
                PonFoco txtAux(12)
                Exit Sub
            Else
                txtAux2(12).Text = Sql
            End If
            
            PonleFoco cmdAceptar
        End Select


        If CalcularElIva Then
            If Index = 5 Or Index = 6 Or Index = 7 Then CalcularIVA
        End If


End Sub

Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtAux(5).Text <> "" Then
            hab = HayKHabilitarCentroCoste(txtAux(5).Text)
        Else
            txtAux(12).Text = ""
        End If
        If hab Then
            txtAux(12).BackColor = &H80000005
        Else
            txtAux(12).BackColor = &H80000018
            txtAux(12).Text = ""
        End If
    End If
    txtAux(12).Enabled = hab
End Sub

'1.-Debe    2.-Haber   3.-Decide en asiento
Private Sub HabilitarImportes(tipoConcepto As Byte)
    Dim bDebe As Boolean
    Dim bHaber As Boolean
    
    'Vamos a hacer .LOCKED =
    Select Case tipoConcepto
    Case 1
        bDebe = False
        bHaber = True
    Case 2
        bDebe = True
        bHaber = False
    Case 3
        bDebe = False
        bHaber = False
    Case Else
        bDebe = True
        bHaber = True
    End Select
    
    txtAux(9).Enabled = Not bDebe
    txtAux(10).Enabled = Not bHaber
    
    If bDebe Then
        txtAux(9).BackColor = &H80000018
    Else
        txtAux(9).BackColor = &H80000005
    End If
    If bHaber Then
        txtAux(10).BackColor = &H80000018
    Else
        txtAux(10).BackColor = &H80000005
    End If
End Sub


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
'        i = Boton
'        SQL = ""
'        If i < 6 Then
'            SQL = "NO"
'        Else
'            If i > 15 Then
'                SQL = "NO"
'            Else
'                'INSERTAR, pero no estamos en edicion lineas
'                If i = 6 And Modo <> 5 Then
'                    SQL = "NO"
'                End If
'            End If
'        End If
'        If SQL <> "" Then
'            msgboxA "Esta modificando el asiento de historico. Finalice primero este proceso.", vbExclamation
'            Exit Sub
'        End If
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
            
            
            frmFacturasCliList.NUmSerie = Text1(2).Text
            frmFacturasCliList.numfactu = Text1(0).Text
            frmFacturasCliList.Fecfactu = Text1(1).Text

            frmFacturasCliList.Show vbModal



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
        
        'Cuando hagamos la contabilizacion especial, que modifica factura pero NO toca el apunte. Aqui debemos guardar en el text1(8) el numero de asiento, para que le modificar no lo ponga a NULL
        'Si no permite modificar el apunte, el total lo recalcul desde los IVAS
        TotalDesdeLineas = False
        If Val(DBLet(Data1.Recordset!no_modifica_apunte, "N")) Then TotalDesdeLineas = True
        B = RecalcularTotalesFactura(TotalDesdeLineas)
        
        If B Then B = ModificaDesdeFormulario2(Me, 1)
        
        If B Then
        End If
  
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


'##### Nuevo para el ambito de fechas
Private Function AmbitoDeFecha(DesbloqueAsiento As Boolean) As Boolean
        AmbitoDeFecha = False
        varFecOk = FechaCorrecta2(CDate(Text1(1).Text))
        If varFecOk > 1 Then
            If varFecOk = 2 Then
                MsgBoxA varTxtFec, vbExclamation
            Else
                MsgBoxA "El asiento pertenece a un ejercicio cerrado.", vbExclamation
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
            txtAux(0).Text = ""
            miI = 3
        Case 3
            txtAux(3).Text = ""
            miI = 0
        Case 4
            txtAux(4).Text = ""
            miI = 1
            
        Case 8
            txtAux(8).Text = ""
            miI = 2
        End Select
        If miI >= 0 Then cmdAux_Click miI
End Sub

Private Sub PonerLineaAnterior(Indice As Integer)
Dim RT As ADODB.Recordset
Dim C As String
On Error GoTo EponerLineaAnterior

    'Si no estamos insertando,modificando lineas
    
    If Modo <> 5 Then Exit Sub
    

    If AdoAux(1).Recordset.EOF Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    
    'Todos los casos menos la ampliacion del concepto
    If Indice <> 5 Then
        Sql = "SELECT "
        Select Case Indice
        Case 0
            C = "codmacta"
            I = 2
        Case 2
            C = "numdocum"
            I = 3
        Case 3
            C = "ctacontr"
            I = 4
        Case 4
            C = "codconce"
            I = 5
        Case 8
            C = "codccost"
            I = -1
        Case Else
            C = ""
        End Select
        If C <> "" Then
            Sql = Sql & C & "  FROM hlinapu"
            Sql = Sql & " WHERE numdiari=" & Data1.Recordset!NumDiari
            Sql = Sql & " AND fechaent='" & Format(Data1.Recordset!FechaEnt, FormatoFecha)
            Sql = Sql & "' AND numasien=" & Data1.Recordset!NumAsien
            If ModoLineas = 2 Then Sql = Sql & " AND linliapu <" & Linliapu
            Sql = Sql & " ORDER BY linliapu DESC"
            Set RT = New ADODB.Recordset
            RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            C = ""
            If Not RT.EOF Then C = DBLet(RT.Fields(0))
            
            'Lo ponemos en txtaux
            If C <> "" Then
                txtAux(Indice).Text = C
                If I >= 0 Then
                    PonFoco txtAux(I)
                End If
            End If
            RT.Close
        End If





    Else
        Sql = "Select hlinliapu,ampconce,nomconce FROM hlinapu,conceptos"
        Sql = Sql & " WHERE conceptos.codconce=hlinapu.codconce AND  numdiari=" & Data1.Recordset!NumDiari
        Sql = Sql & " AND fechaent='" & Format(Data1.Recordset!FechaEnt, FormatoFecha)
        Sql = Sql & "' AND numasien=" & Data1.Recordset!NumAsien
        If ModoLineas = 2 Then Sql = Sql & " AND linliapu <" & Linliapu
           
        Sql = Sql & " ORDER BY linliapu DESC"
        Set RT = New ADODB.Recordset
        RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        C = ""
        If Not RT.EOF Then
            Sql = DBLet(RT.Fields(1))
            C = DBLet(RT.Fields(2))
        End If
        
        'Lo ponemos en txtaux
        If Sql <> "" Then
            If C <> "" Then
                I = InStr(1, Sql, C)
                If I > 0 Then Sql = Trim(Mid(Sql, Len(C) + 1))
            End If
            txtAux(5).Text = txtAux(5).Text & Sql & " "
            txtAux(5).SelStart = Len(txtAux(5).Text)
            PonFoco txtAux(6)
        End If
        RT.Close

    
    End If
    
EponerLineaAnterior:
    If Err.Number <> 0 Then Err.Clear
    Set RT = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerF6()
Dim RsF6 As ADODB.Recordset
Dim C As String

    On Error GoTo EHacerF6
    
    Set RsF6 = New ADODB.Recordset
            
    
    C = "SELECT hlinapu.numasien, hlinapu.linliapu, hlinapu.codmacta, cuentas.nomm  acta,"
    C = C & " hlinapu.numdocum, hlinapu.ctacontr, hlinapu.codconce, conceptos.nomconce as nombreconcepto, hlinapu.ampconce, cuentas_1.nommacta as nomctapar,"
    C = C & " hlinapu.timporteD, hlinapu.timporteH, hlinapu.codccost, ccoste.nomccost as centrocoste,"
    C = C & " hlinapu.numdiari, hlinapu.fechaent"
    C = C & " FROM (((hlinapu LEFT JOIN cuentas AS cuentas_1 ON hlinapu.ctacontr ="
    C = C & " cuentas_1.codmacta) LEFT JOIN ccoste ON hlinapu.codccost = ccoste.codccost)"
    C = C & " INNER JOIN cuentas ON hlinapu.codmacta = cuentas.codmacta) INNER JOIN"
    C = C & " conceptos ON hlinapu.codconce = conceptos.codconce"
    C = C & " WHERE numasien = " & Data1.Recordset!NumAsien
    C = C & " AND numdiari =" & Data1.Recordset!NumDiari
    C = C & " AND fechaent= '" & Format(Data1.Recordset!FechaEnt, FormatoFecha) & "'"
    C = C & " ORDER BY hlinapu.linliapu DESC"
    
    
    
    
    
    RsF6.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RsF6.EOF Then
        C = " numasiento = " & Data1.Recordset!NumAsien & vbCrLf
        C = " fecha= " & Format(Data1.Recordset!FechaEnt, "dd/mm/yyyy")
    
        MsgBoxA "No se ha encontrado las lineas: " & vbCrLf & C, vbExclamation
    Else
        'Ya tengo la ultima linea
        txtAux(4).Text = RsF6!codmacta
        
        txtAux(4).Text = RsF6!codmacta
        txtAux2(4).Text = RsF6!Nommacta
        txtAux(5).Text = DBLet(RsF6!Numdocum, "T")
        txtAux(6).Text = DBLet(RsF6!ctacontr, "T")
        txtAux(7).Text = RsF6!CodConce
        txtAux(8).Text = DBLet(RsF6!Ampconce, "T")
        C = DBLet(RsF6!timported, "T")
        If C <> "" Then
            txtAux(9).Text = Format(C, "0.00")
        Else
            txtAux(9).Text = C
        End If
        C = DBLet(RsF6!timporteH, "T")
        If C <> "" Then
            txtAux(10).Text = Format(C, "0.00")
        Else
            txtAux(10).Text = C
        End If
        txtAux(11).Text = DBLet(RsF6!CodCCost, "T")
        HabilitarImportes 3
        HabilitarCentroCoste
        txtAux2(5).Text = DBLet(RsF6!Nommacta, "T")
        txtAux2(12).Text = DBLet(RsF6!centrocoste, "T")
        
    End If
    RsF6.Close
    Set RsF6 = Nothing
    Exit Sub
EHacerF6:
    MuestraError Err.Number, Err.Description
    Set RsF6 = Nothing
End Sub

Private Function AuxOK() As String
    
    'Cuenta
    If txtAux(4).Text = "" Then
        AuxOK = "Cuenta no puede estar vacia."
        Exit Function
    End If
    
    If Not IsNumeric(txtAux(4).Text) Then
        AuxOK = "Cuenta debe ser num�rica"
        Exit Function
    End If
    
    If txtAux2(4).Text = NO Then
        AuxOK = "La cuenta debe estar dada de alta en el sistema"
        Exit Function
    End If
    
    If Not EsCuentaUltimoNivel(txtAux(4).Text) Then
        AuxOK = "La cuenta no es de �ltimo nivel"
        Exit Function
    End If
    
    
        
    'Codigo de iva
    If txtAux(4).Text = "" Then
        AuxOK = "El c�digo de iva no puede estar vacio"
        Exit Function
    End If
        
    If txtAux(7).Text <> "" Then
        If Not IsNumeric(txtAux(7).Text) Then
            AuxOK = "El c�digo de iva debe de ser num�rico."
            Exit Function
        End If
    End If
    
    'Importe
    If txtAux(6).Text <> "" Then
        If Not EsNumerico(txtAux(6).Text) Then
            AuxOK = "El importe DEBE debe ser num�rico"
            Exit Function
        End If
    End If
    
    
    'cENTRO DE COSTE
    If txtAux(12).Enabled Then
        If txtAux(12).Text = "" Then
            AuxOK = "Centro de coste no puede ser nulo"
            Exit Function
        End If
    End If
    
                                            'Fecha del asiento
    If EstaLaCuentaBloqueada2(txtAux(5).Text, CDate(Text1(1).Text)) Then
        AuxOK = "Cuenta bloqueada: " & txtAux(5).Text
        Exit Function
    End If
    
    AuxOK = ""
End Function


Private Function ActualizarASiento() As Boolean
Dim B As Boolean


End Function

Private Function ComprobarNumeroFactura(Actual As Boolean) As Boolean
Dim cad As String
Dim RT As ADODB.Recordset
        cad = " WHERE numfactu=" & Text1(0).Text
        cad = cad & " and numserie = " & DBSet(Text1(2).Text, "T")
        
        If Actual Then
            I = 0
        Else
            I = 1
        End If
        
        cad = cad & " AND anofactu =" & DBSet(Text1(14).Text, "N")
        
        Set RT = New ADODB.Recordset
        ComprobarNumeroFactura = True
        I = 0
        RT.Open "Select numfactu from factcli" & cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.EOF) Then
                ComprobarNumeroFactura = False
            End If
        End If
        RT.Close
        If ComprobarNumeroFactura Then
            I = 1
            RT.Open "Select numfactu from factcli" & cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
            MsgBoxA cad, vbExclamation
        End If
End Function

Private Function SituarData1(Insertar As Boolean) As Boolean
    Dim Sql As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        Sql = "Select * from factcli WHERE numserie =" & DBSet(Text1(2).Text, "T")
        Sql = Sql & " AND fecfactu=" & DBSet(Text1(1).Text, "F") & " AND numfactu = " & Text1(0).Text
        Data1.RecordSource = Sql
    'End If
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!NUmSerie) = Text1(2).Text Then
                If CStr(.Fields!numfactu) = Text1(0).Text Then
                    If Format(CStr(.Fields!Fecfactu), "dd/mm/yyyy") = Text1(1).Text Then
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
        cad = "select h.numlinea,  h.codigiva, tt.nombriva,  h.baseimpo, h.impoiva, h.imporec from factcli_totales h inner join tiposiva tt on h.codigiva = tt.codigiva  WHERE "
        cad = cad & " numserie=" & DBSet(Data1.Recordset!NUmSerie, "T")
        cad = cad & " and numfactu=" & Data1.Recordset!numfactu
        
        'Julio2021    No deberia linkar por fecha.  Con el a�o SOBRA
        cad = cad & " and fecfactu=" & DBSet(Data1.Recordset!Fecfactu, "F")
        cad = cad & " and anofactu=" & Data1.Recordset!Anofactu
        GroupBy = ""
        BuscaChekc = "numlinea"
        
    End Select
    
    
    'BuscaChekc="" si es la opcion de precios especiales
    cad = cad & " ORDER BY 1"
    
    lw1.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not Rs.EOF
        Set IT = lw1.ListItems.Add()

        IT.Text = Rs!NumLinea
        IT.SubItems(1) = Format(Rs!codigiva, "000")
        IT.SubItems(2) = Rs!nombriva
        IT.SubItems(3) = Format(Rs!Baseimpo, "###,###,##0.00")
        IT.SubItems(4) = Format(Rs!Impoiva, "###,###,##0.00")
        If DBLet(Rs!ImpoRec, "N") <> 0 Then
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
'        msgboxA "No existe el archivo: " & vpaz, vbExclamation
    Else
        Set IT = lw1.ListItems.Add()

        IT.Text = Me.adodc1.Recordset!Orden '"Nuevo " & Contador
        
        IT.SubItems(1) = Me.adodc1.Recordset.Fields(5)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        IT.SubItems(2) = vpaz
        IT.SubItems(3) = Me.adodc1.Recordset.Fields(0)
        
        Set IT = Nothing
    End If
End Sub



Private Sub EliminarImagen()
Dim Sql As String
Dim Mens As String
    
    On Error GoTo eEliminarImagen

    Mens = "Va a proceder a eliminar de la lista correspondiente al asiento. " & vbCrLf & vbCrLf & "� Desea continuar ?" & vbCrLf & vbCrLf
    
    If MsgBoxA(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Sql = "delete from hcabapu_fichdocs where numasien = " & DBSet(Text1(0).Text, "N") & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N") & " and codigo = " & Me.lw1.SelectedItem.SubItems(3)
        Conn.Execute Sql
        FicheroAEliminar = lw1.SelectedItem.SubItems(2)
        CargaDatosLW
        
    End If
    Exit Sub

eEliminarImagen:
    MuestraError Err.Number, "Eliminar im�gen", Err.Description
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
                If MostrarMensaje Then MsgBoxA varTxtFec, vbExclamation
            Else
                If MostrarMensaje Then MsgBoxA "El asiento pertenece a un ejercicio cerrado.", vbExclamation
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
                If MostrarMensaje Then MsgBoxA "Este apunte pertenece a una factura de " & CadenaDesdeOtroForm & " y solo se puede modificar en el registro" & _
                    " de facturas de " & CadenaDesdeOtroForm & ".", vbExclamation
                I = -1
            Exit Function
        Else
            SePuedeModificarAsiento = True
        End If


End Function

Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim J As Long
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I

    'Tipo de factura
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wconce340 ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not Rs.EOF
        Combo1(0).AddItem Rs!Descripcion
        Combo1(0).ItemData(Combo1(0).NewIndex) = Asc(Rs!Codigo)
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Tipo de operacion
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtipopera where codigo <= 3 ORDER BY codigo"
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
    I = 0
    While Not Rs.EOF
        Combo1(3).AddItem Rs!nomintra
        Combo1(3).ItemData(Combo1(3).NewIndex) = Asc(Rs!Codintra)
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    
    'Tipo situacion inmueble
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtipoinmueble ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not Rs.EOF
        Combo1(4).AddItem Rs!Descripcion
        Combo1(4).ItemData(Combo1(4).NewIndex) = Asc(Rs!Codigo)
        I = I + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    'Tipo Documento identifficador
    '0 DEFAULT     1 NIF      2 nifintra   3 Pasapo    7 NO censado'
    Sql = "Defecto|N.I.F.|NIF Intracomunitario|Pasaporte|No censado|"
    I = 0
    While I < 5
        I = I + 1
        Combo1(5).AddItem RecuperaValor(Sql, CInt(I))
        Combo1(5).ItemData(Combo1(5).NewIndex) = RecuperaValor("0|1|2|3|7|", CInt(I))
    Wend
    
    
    
    

End Sub

Private Function ComprobarPeriodo2(Indice As Integer) As Boolean
Dim Cerrado As Boolean
Dim MensajeSII As String
Dim Mostrar As Boolean
Dim ModEspecial As Boolean



    '[Monica]12/09/2016: Si cerrado o liquidado no hacemos nada en cartera
    ModificarCobros = True
    
    
    If vParam.SIITiene Then
            'SI esta presentada...
        If Modo <> 3 And Modo <> 1 Then
            If DBLet(Data1.Recordset!SII_ID, "N") > 0 Then
                If Text1(28).BackColor = vbSiiMofificando Then
                    'Esta modificando la factura YA presentada.
                    'Con lo cual le deja pasar
                    'ok
            
            
                Else
                    'If Val(DBLet(data1.Recordset!sii _status, "N")) > 2 Then
                    If Text1(28).BackColor = vbSiiOK Or Text1(28).BackColor = vbSiiEnProceso Or Text1(28).BackColor = vbSiiAceptadaConErr Then
                        
                        
                        
                            'Si fecha >= fechaini
                            ModEspecial = False
                            If vUsu.Nivel <= 1 Then
                                If Data1.Recordset!Fecfactu >= vParam.fechaini Then
                                     If Val(DBLet(Data1.Recordset!no_modifica_apunte, "N")) = 0 Then ModEspecial = True
                                End If
                            End If
                            
                            If ModEspecial Then
                            
                                'Bloqueamos el registro
                                
                            
                                CadenaDesdeOtroForm = ""
                                Ampliacion = ""
                                Conn.Execute "DELETE from tmpfaclin WHERE codusu = " & vUsu.Codigo
                                With frmFacturaModificar
                                    .Cliente = True
                                    .Anyo = Data1.Recordset!Anofactu
                                    .Codigo = Data1.Recordset!numfactu
                                    .NUmSerie = Data1.Recordset!NUmSerie
                                    .Fecha = Data1.Recordset!Fecfactu
                                    .Show vbModal
                                End With
                                
                                
                                'Si que ha modificado
                                Screen.MousePointer = vbHourglass
                                If CadenaDesdeOtroForm <> "" Or Ampliacion <> "" Then
                                    
                                    If ModificaFacturaSiiPresentada Then
                                        PosicionarData
                                        PonerCampos
                                    End If
                                End If
                                Screen.MousePointer = vbDefault
                            Else
                                MsgBoxA "La factura ya esta presentada en el sistema de SII de la AEAT.", vbExclamation
                                
                            End If
    
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        Exit Function
                    End If
                End If 'modifando-ok
            End If
        End If
        
        
        If Modo > 2 Then
            'If DateDiff("d", CDate(Text1(Indice).Text), Now) > vParam.SIIDiasAviso Then
            If UltimaFechaCorrectaSII(vParam.SIIDiasAviso, Now) > CDate(Text1(Indice).Text) Then
                MensajeSII = "" 'String(70, "*") & vbCrLf
                MensajeSII = MensajeSII & "SII." & vbCrLf & vbCrLf & "Excede del m�ximo dias permitido para comunicar la factura" & vbCrLf & MensajeSII
            End If
        End If
    End If
    'Primero pondremos la fecha a a�o periodo
    I = Year(CDate(Text1(Indice).Text))
    If vParam.periodos = 0 Then
        'Trimestral
        Ancho = ((Month(CDate(Text1(Indice).Text)) - 1) \ 3) + 1
        Else
        Ancho = Month(CDate((Text1(Indice).Text)))
    End If
    Cerrado = False
    If I < vParam.Anofactu Then
        Cerrado = True
    Else
        If I = vParam.Anofactu Then
            'El mismo a�o. Comprobamos los periodos
            If vParam.perfactu >= Ancho Then _
                Cerrado = True
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
            Sql = "La fecha "
            If Indice = 0 Then
                Sql = Sql & "factura"
            Else
                Sql = Sql & "liquidacion"
            End If
            Sql = Sql & " corresponde a un periodo ya liquidado. " & vbCrLf
        Else
            Sql = ""
        End If
        
        If MensajeSII <> "" Then
            MensajeSII = MensajeSII & vbCrLf & vbCrLf
            If Sql <> "" Then Sql = vbCrLf & String(48, "*") & vbCrLf & Sql & vbCrLf & String(48, "*") & vbCrLf
        End If
        Sql = MensajeSII & Sql
        
        
        If vUsu.Nivel = 0 Then
        
            Sql = Sql & vbCrLf & " �Desea continuar igualmente ?"
  
            If MsgBoxA(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then ComprobarPeriodo2 = False
            
        Else
            MsgBoxA Sql, vbExclamation
            
            ComprobarPeriodo2 = False
        
        End If
        
        '[Monica]12/09/2016: no tocar cartera
        ModificarCobros = False
        
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
    
    For I = 15 To 21
        Text1(I).Text = ""
    Next I
    
    If Not Rs.EOF Then
        Sql = ""
        
        If Not IsNull(Rs!forpa) Then
            Sql = Rs!forpa
        Else
            If Modo = 4 Then
                If Cuenta = Data1.Recordset!codmacta Then
                    'Misma cuenta que la que habia
                    'Ha cambiado cli y lo ha vuelto a poner en el mismo modificacion
                    Sql = DBLet(Data1.Recordset!Codforpa, "T")
                End If
            End If
        End If
        If Sql <> "" Then
            Text1(5).Text = DBLet(Sql, "N")
            Text4(5).Text = PonerNombreDeCod(Text1(5), "formapago", "nomforpa", "codforpa", "N")
        End If
        Text1(15).Text = DBLet(Rs!razosoci, "T")
        If Text1(15).Text = "" Then Text1(15).Text = DBLet(Rs!Nommacta, "T")
        Text1(16).Text = DBLet(Rs!dirdatos, "T")
        Text1(17).Text = DBLet(Rs!codposta, "T")
        Text1(18).Text = DBLet(Rs!desPobla, "T")
        Text1(19).Text = DBLet(Rs!desProvi, "T")
        Text1(20).Text = DBLet(Rs!nifdatos, "T")
        Text1(21).Text = DBLet(Rs!codpais, "T")
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
        I = Year(vParam.fechaini)
        If vUsu.FiltroFactCli < 3 Then
            'INicio = actual
            Aux = " anofactu >= " & I
            Else
            Aux = " anofactu >=" & I + 1
        End If
        I = Year(vParam.fechafin)
        If vUsu.FiltroFactCli = 2 Then
            Aux = Aux & " AND anofactu <= " & I
        Else
            Aux = Aux & " AND anofactu <= " & I + 1
        End If
        
    End If  'filtro=0
    AnyadeCadenaFiltro = Aux
End Function



Private Sub CalcularIVA()
Dim J As Integer
Dim Base As Currency
Dim Aux As Currency

    Base = ImporteFormateado(txtAux(6).Text)
    
    'EL iva
    Aux = ImporteFormateado(txtAux(8).Text) / 100
    If Aux = 0 Then
        If txtAux(10).Text = "" Then
            txtAux(10).Text = ""
        Else
            txtAux(10).Text = "0,00"
        End If
    Else
        'txtAux(10).Text = Format(Round((Aux * Base), 2), FormatoImporte)
        txtAux(10).Text = Format(Round2((Aux * Base), 2), FormatoImporte)
    End If
    
    'Recargo
    Aux = ImporteFormateado(txtAux(9).Text) / 100
    If Aux = 0 Then
        txtAux(11).Text = ""
    Else
        'txtAux(11).Text = Format(Round((Aux * Base), 2), FormatoImporte)
        txtAux(11).Text = Format(Round2((Aux * Base), 2), FormatoImporte)
            ' End
    End If

End Sub

Private Function RecalcularTotales() As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim SqlValues As String
Dim I As Long
Dim Rs As ADODB.Recordset

Dim Baseimpo As Currency
Dim Basereten As Currency
Dim Impoiva As Currency
Dim ImpoRec As Currency
Dim Imporeten As Currency
Dim TotalFactura As Currency
Dim ImporAux As Currency
    On Error GoTo eRecalcularTotales

    RecalcularTotales = False

    Sql = "delete from factcli_totales where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    Conn.Execute Sql
    
    SqlInsert = "insert into factcli_totales (numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec) values "
    
    
    'Sumaba los importes de IVAS desde las bases
    'Sql = "select codigiva, porciva, porcrec, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec from factcli_lineas "
    Sql = "select codigiva, porciva, porcrec, sum(baseimpo) baseimpo "
    If vParam.ModificarIvaLineasFraCli Then Sql = Sql & ",sum(impoiva) importeiva,sum(imporec) importerec "
    Sql = Sql & " from factcli_lineas where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    Sql = Sql & " group by 1,2,3"
    Sql = Sql & " order by 1,2,3"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    
    SqlValues = ""
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    ImpoRec = 0
    Imporeten = 0
    
    TotalFactura = 0
    
    While Not Rs.EOF
        Sql = ", (" & DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & "," & DBSet(Text1(14).Text, "N") & ","
        Sql = Sql & DBSet(I, "N") & "," & DBSet(Rs!Baseimpo, "N") & "," & DBSet(Rs!codigiva, "N") & "," & DBSet(Rs!PorcIva, "N") & "," & DBSet(Rs!porcrec, "N") & ","
        If vParam.ModificarIvaLineasFraCli Then
            ImporAux = DBLet(Rs!importeiva, "N")
        Else
            ImporAux = Round2((Rs!Baseimpo * Rs!PorcIva) / 100, 2)
        End If
        Impoiva = Impoiva + ImporAux  ' DBLet(Rs!Imporiva, "N")
        Sql = Sql & DBSet(ImporAux, "N") & ","
        
        If vParam.ModificarIvaLineasFraCli Then
            ImporAux = DBLet(Rs!importerec, "N")
        Else
            ImporAux = DBLet(Rs!porcrec, "N")
            ImporAux = Round2((Rs!Baseimpo * ImporAux) / 100, 2)
        End If
        ImpoRec = ImpoRec + DBLet(ImporAux, "N")
        Sql = Sql & DBSet(ImporAux, "N") & ")"
        
        
        
        
        SqlValues = SqlValues & Sql
        
        Baseimpo = Baseimpo + DBLet(Rs!Baseimpo, "N")
        'Impoiva = Impoiva + DBLet(Rs!Imporiva, "N")
        'ImpoRec = ImpoRec + DBLet(Rs!imporrec, "N")
        
        I = I + 1
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If SqlValues <> "" Then
        'SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        SqlValues = Mid(SqlValues, 2) 'David. Quiamos la primera coma y punto
        Conn.Execute SqlInsert & SqlValues
    End If
    
    
    RecalcularTotales = RecalcularTotalesFactura(False)
    Exit Function
    
eRecalcularTotales:
    MuestraError Err.Number, "Recalcular Totales", Err.Description
End Function


'DesdeA�adirLineaDeTotalesIVA.  Integraciones SAGE
Private Function RecalcularTotalesFactura(DesdeA�adirLineaDeTotalesIVA As Boolean) As Boolean
Dim Sql As String
Dim SqlInsert As String
Dim SqlValues As String
Dim I As Long
Dim Rs As ADODB.Recordset

Dim Baseimpo As Currency
Dim Basereten As Currency
Dim Impoiva As Currency
Dim ImpoRec As Currency
Dim Imporeten As Currency
Dim TotalFactura As Currency
Dim PorcRet As Currency
Dim Aux As Currency
Dim TipoRetencion As Integer
Dim Suplidos As Currency
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
    IvaModificable = False
    
    
    Sql = ""
    SqlInsert = "factcli_lineas"
    If DesdeA�adirLineaDeTotalesIVA Then
    
        
        Sql = "0"
        If Combo1(2).ListIndex > 0 Then Sql = "1"
            
        SqlInsert = "factcli_totales"
    End If
    Sql = "select " & Sql & " aplicret,######.codigiva,tipodiva, sum(baseimpo) baseimpo, porciva,porcrec "
    
    If DesdeA�adirLineaDeTotalesIVA Then
        IvaModificable = True
    Else
        If vParam.ModificarIvaLineasFraCli Then IvaModificable = True
    End If
    
    If IvaModificable Then Sql = Sql & ",sum(impoiva) importeiva,sum(imporec) importerec "

    Sql = Sql & " from ###### "
    Sql = Sql & "  inner join tiposiva on ######.codigiva=tiposiva.codigiva"
    Sql = Sql & " where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    Sql = Sql & " group by 1,2 order by 1,2"
    
    Sql = Replace(Sql, "######", SqlInsert)
    SqlInsert = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
    
        'SUPLIDO
        If Rs!TipoDIva = 4 Then
            Suplidos = Suplidos + Rs!Baseimpo

    
        Else
            Baseimpo = Baseimpo + DBLet(Rs!Baseimpo, "N")
            
            'ANTES
            If IvaModificable Then
                Impoiva = Impoiva + DBLet(Rs!importeiva, "N")
                ImpoRec = ImpoRec + DBLet(Rs!importerec, "N")
            Else
            
                PorcRet = DBLet(Rs!PorcIva, "N")
                Aux = Round2((DBLet(Rs!Baseimpo, "N") * PorcRet) / 100, 2)
                Impoiva = Impoiva + Aux
                
                PorcRet = DBLet(Rs!porcrec, "N")
                If PorcRet > 0 Then
                    Aux = Round2((DBLet(Rs!Baseimpo, "N") * PorcRet) / 100, 2)
                    ImpoRec = ImpoRec + Aux
                End If
            
            End If
            
            
            If Rs!aplicret = 1 Then
                Basereten = Basereten + DBLet(Rs!Baseimpo, "N")
                
                If TipoRetencion = 1 Then
                                    
                    If IvaModificable Then
                        Basereten = Basereten + DBLet(Rs!importeiva, "N")
                    Else
                        PorcRet = DBLet(Rs!PorcIva, "N")
                        Aux = Round2((DBLet(Rs!Baseimpo, "N") * PorcRet) / 100, 2)
                       
                        Basereten = Basereten + Aux
                    End If
                End If
            End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    PorcRet = ImporteFormateado(Text1(7).Text)
    
    If PorcRet = 0 Then Basereten = 0
    
    If PorcRet = 0 Then
        Imporeten = 0
    Else
        Imporeten = Round2((PorcRet * Basereten / 100), 2)
    End If
    
    TotalFactura = Baseimpo + Impoiva + ImpoRec - Imporeten + Suplidos
    
    Text1(9).Text = Format(Baseimpo, FormatoImporte)
    Text1(11).Text = Format(Basereten, FormatoImporte)
    Text1(10).Text = Format(Impoiva, FormatoImporte)
    Text1(12).Text = Format(Imporeten, FormatoImporte)
    Text1(13).Text = Format(TotalFactura, FormatoImporte)
    
    If PorcRet = 0 Then
        Text1(11).Text = ""
        Text1(12).Text = ""
    End If
    
    Sql = "update factcli set "
    Sql = Sql & " totbases = " & DBSet(Baseimpo, "N")
    Sql = Sql & ", totivas = " & DBSet(Impoiva, "N")
    Sql = Sql & ", totrecargo = " & DBSet(ImpoRec, "N")
    Sql = Sql & ", totfaccl = " & DBSet(TotalFactura, "N")
    Sql = Sql & ", totbasesret = " & DBSet(Basereten, "N", "S")
    Sql = Sql & ", trefaccl = " & DBSet(Imporeten, "N", "S")
    Sql = Sql & ", Suplidos = " & DBSet(Suplidos, "N", "S")
    
    
    Sql = Sql & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numfactu= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    Conn.Execute Sql
    
    
    'OCTUB 2017
    'Si ha ha cambiado la fecha factura updateamos
    If Modo = 4 Then
        If Text1(1).Text <> Format(Data1.Recordset!Fecfactu, "dd/mm/yyyy") Then
             Sql = "UPDATE factcli_totales set fecfactu = " & DBSet(Text1(1).Text, "F")
             Sql = Sql & " where numserie= " & DBSet(Text1(2).Text, "T") & " and numfactu= " & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
             Ejecuta Sql
        End If
    End If
    RecalcularTotalesFactura = True
    Exit Function
    
eRecalcularTotalesFactura:
    MuestraError Err.Number, "Recalcular Totales Factura", Err.Description
End Function


Private Function IntegrarFactura(DentroBeginTrans As Boolean) As Boolean
Dim SqlLog As String

    
    
    
    If Not ContabilizaApunte Then
        ContabilizaApunte = True 'Lo dejo por defecto otra vez
        IntegrarFactura = True
        Exit Function
    End If
    
    IntegrarFactura = False
    
    SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " de fecha " & Text1(1).Text
    If Me.AdoAux(1).Recordset.EOF Then
        SqlLog = SqlLog & vbCrLf & "L�nea   : EOF    Vacio"
    Else
        SqlLog = SqlLog & vbCrLf & "L�nea   : " & DBLet(Me.AdoAux(1).Recordset!NumLinea, "N")
        SqlLog = SqlLog & vbCrLf & "Cuenta  : " & DBLet(Me.AdoAux(1).Recordset!codmacta, "T") & " " & DBLet(Me.AdoAux(1).Recordset!Nommacta, "T")
        SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Me.AdoAux(1).Recordset!Baseimpo, "N")
    End If
    
    
    With frmActualizar
        .OpcionActualizar = 6
        'NumAsiento     --> CODIGO FACTURA
        'NumDiari       --> A�O FACTURA
        'NUmSerie       --> SERIE DE LA FACTURA
        'FechaAsiento   --> Fecha factura
        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
        .NumFac = CLng(Text1(0).Text)
        .NumDiari = CInt(Text1(14).Text)
        .NUmSerie = Text1(2).Text
        .FechaAsiento = Text1(1).Text
        .FechaAnterior = FecFactuAnt
        .DentroBeginTrans = DentroBeginTrans
        .SqlLog = SqlLog
        If NumAsien < 0 Then
            
            If Not Text1(8).Enabled Then
                If Text1(8).Text <> "" Then
                    NumAsien = Text1(8).Text
                End If
            End If
            
        End If
        If NumDiario_ <= 0 Then NumDiario_ = vParam.numdiacl
        .DiarioFacturas = NumDiario_
        .NumAsiento = NumAsien
        .Show vbModal
        
        If AlgunAsientoActualizado Then IntegrarFactura = True
        
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
            frmActualizar.NumAsiento = Data1.Recordset!NumAsien
            frmActualizar.FechaAsiento = FecFactuAnt
            frmActualizar.NumDiari = Data1.Recordset!NumDiari
            frmActualizar.FechaAnterior = Data1.Recordset!FechaEnt
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            If AlgunAsientoActualizado Then Desintegrar = True
        End If
End Function


Private Function DesvincularFactura() As Boolean
On Error Resume Next
    Sql = "UPDATE factcli set numasien=NULL, fechaent=NULL, numdiari=NULL"
    Sql = Sql & " WHERE numfactu = " & Data1.Recordset!numfactu
    Sql = Sql & " AND numserie = '" & Data1.Recordset!NUmSerie & "'"
    Sql = Sql & " AND anofactu =" & Data1.Recordset!Anofactu
    NumAsien = Data1.Recordset!NumAsien
    NumDiario_ = Data1.Recordset!NumDiari
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

Public Function HayQueContabilizarDesdePantallaCobros() As Boolean
    HayQueContabilizarDesdePantallaCobros = False
    If Cobrado = 1 Then
        HayQueContabilizarDesdePantallaCobros = True
    
    End If
End Function

Public Function ContabilizarCobros() As Boolean
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
Dim fecefect As Date
    
    On Error GoTo ECon
    
    ContabilizarCobros = False
    
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", CDate(FechaCobro) <= vParam.fechafin, True) = 1 Then Exit Function

    Set FP = New Ctipoformapago
    
    Linea = DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(Text1(5), "N"))
    
    If FP.Leer(Linea) Then
        Set Mc = Nothing
        Set FP = Nothing
    End If
    
    Sql1 = "select * "
    Sql = " from cobros where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    Sql = Sql & " order by numorden"
    
    TotImpo = DevuelveValor("select sum(impvenci) " & Sql)
    
    Sql = Sql1 & Sql
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Inserto cabecera de apunte
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien,feccreacion,usucreacion,desdeaplicacion, obsdiari) VALUES ("
    Sql = Sql & FP.diaricli
    Sql = Sql & ",'" & Format(FechaCobro, FormatoFecha) & "'," & Mc.Contador & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizaci�n Cobro Facturas Cliente ',"
    Sql1 = "Generado desde Facturas de Cliente el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre
    If TotImpo < 0 Then Sql1 = Sql1 & "  (ABONO)"
    Sql1 = DBSet(Sql1, "T")
    Conn.Execute Sql & Sql1 & ")"
    
    Linea = 0
    fecefect = CDate("01/01/2100")
    While Not Rs.EOF
        
        Linea = Linea + 1
        
        'importe
        impo = ImporteFormateado(DBLet(Rs!ImpVenci))
        
        
        If Rs!FecVenci < fecefect Then fecefect = Rs!FecVenci
        
        
        'Inserto en las lineas de apuntes
        Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
        Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfaccl, fecfactu, numorden, tipforpa) VALUES ("
        Sql = Sql & FP.diaricli
        Sql = Sql & ",'" & Format(FechaCobro, FormatoFecha) & "'," & Mc.Contador & ","
        
        
        'numdocum
        Numdocum = Text1(2).Text & Text1(0).Text ' letra de serie y factura
        
        'Concepto y ampliacion del apunte
        Ampliacion = ""
        'CLIENTES
        Debe = False
        If impo < 0 Then
            If Not vParam.abononeg Then Debe = True
        End If
        If Debe Then
            Conce = FP.ampdecli
            LlevaContr = FP.ctrdecli = 1
            ElConcepto = FP.condecli
        Else
            ElConcepto = FP.conhacli
            Conce = FP.amphacli
            LlevaContr = FP.ctrhacli = 1
        End If
               
        'Si el importe es negativo y no permite abonos negativos
        'como ya lo ha cambiado de lado (dbe <-> haber)
        If impo < 0 Then
            If Not vParam.abononeg Then impo = Abs(impo)
        End If
           
        If Conce = 2 Then
           Ampliacion = Ampliacion & DBLet(Rs!FecVenci)  'Fecha vto
        ElseIf Conce = 4 Then
            'Contra partida
            Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaBanco, "T")
        Else
            
           If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
           Ampliacion = Ampliacion & Text1(2).Text & "/" & Text1(0).Text 'RecuperaValor(Vto, 1) & "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
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
        cad = cad & ",'COBROS',0," & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!Fecfactu, "F") & "," & DBSet(Rs!numorden, "N") & "," & DBSet(FP.tipoformapago, "N") & ")"
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
     Debe = True
     If TotImpo < 0 Then
        If Not vParam.abononeg Then
            Debe = False
            TotImpo = Abs(TotImpo)
        End If
    End If
    
        
     
     If Not Debe Then
         Conce = FP.ampdecli
         LlevaContr = FP.ctrdecli = 1
         ElConcepto = FP.condecli
     Else
         ElConcepto = FP.conhacli
         Conce = FP.amphacli
         LlevaContr = FP.ctrhacli = 1
     End If
           
           
    If Conce = 2 Then
       'Ampliacion = Ampliacion & DBLet(Text4(4).Text, "T")  'Fecha vto
       Ampliacion = Ampliacion & "Fec.Vto: " & Format(fecefect, "dd/mm/yyyy") 'Fecha efecto
       
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(Text1(2).Text)
    Else
        
       If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
       Ampliacion = Ampliacion & Text1(2) & "/" & Text1(0).Text
    End If
    
    
    Conce = ElConcepto
    cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    cad = Linea & "," & DBSet(CtaBanco, "T") & ",'" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
    
    If Debe Then
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
    cad = cad & ",'COBROS',0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
    cad = Sql & cad
    Conn.Execute cad
    
    ContabilizarCobros = True

    Set Mc = Nothing
    Set FP = Nothing

    Exit Function
ECon:
    MuestraError Err.Number, "Contabilizar anticipo"
    Set Mc = Nothing
    Set FP = Nothing
End Function

Public Function InsertarCobros(ByRef Mens As String) As Boolean
Dim Sql As String
Dim textCSB As String
Dim CadInsert As String
Dim CadValues As String
Dim Rs As ADODB.Recordset
Dim I As Long

    On Error GoTo eInsertarCobros

    InsertarCobros = False

    Sql = "select * from tmpcobros where codusu = " & DBSet(vUsu.Codigo, "N") & " order by numorden "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    I = 0
    Mens = "Insertando Cobros: " & vbCrLf & vbCrLf
    B = InsertaCobros(Rs, I, Mens)
    
    Set Rs = Nothing
    
    InsertarCobros = B
    Exit Function
    
eInsertarCobros:
    MuestraError Err.Number, "Insertar Cobros", Err.Description
End Function

Public Sub EstableceValoresCobro(CadenaTexto As String)
 If CadenaTexto <> "" Then
        CtaBanco = RecuperaValor(CadenaTexto, 1)
        IBAN = Replace(RecuperaValor(CadenaTexto, 2), " ", "")
        
        Cobrado = RecuperaValor(CadenaTexto, 3)
        FechaCobro = RecuperaValor(CadenaTexto, 4)
        ContinuarCobro = True
    End If

End Sub

Private Function InsertaCobros(ByRef RS1 As ADODB.Recordset, ByRef I As Long, ByRef Mens As String) As Boolean
Dim CadInsert As String
Dim CadValues As String
Dim textCSB As String
Dim Sql As String
        
    On Error GoTo eInsertaCobros
        
    InsertaCobros = False
        
    CadInsert = "insert into cobros (numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci," & _
                "ctabanc1,fecultco,impcobro,emitdocum,recedocu,contdocu," & _
                "text33csb,text41csb,ultimareclamacion,agente,departamento,transfer," & _
                "nomclien,domclien,pobclien,cpclien,proclien,iban,nifclien,codpais,situacion, codusu) values "
    CadValues = ""
    
    While Not RS1.EOF
        I = I + 1
        
        Sql = DBSet(Text1(2).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & "," & DBSet(I, "N") & ","
        Sql = Sql & DBSet(Text1(4).Text, "T") & "," & DBSet(Text1(5).Text, "N") & "," & DBSet(RS1!FecVenci, "F") & "," & DBSet(RS1!ImpVenci, "N") & ","
        Sql = Sql & DBSet(CtaBanco, "T", "S") & ","
        
        If Cobrado Then
            Sql = Sql & DBSet(FechaCobro, "F") & "," & DBSet(RS1!ImpVenci, "N") & ","
        Else
            Sql = Sql & ValorNulo & "," & ValorNulo & ","
        End If
        
        Sql = Sql & "0,0,0,"
        
        textCSB = "Factura " & Trim(Text1(2).Text) & "-" & Text1(0).Text & " de Fecha " & Text1(1).Text
        
        Sql = Sql & DBSet(textCSB, "T") & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Text1(26).Text, "N", "S") & "," & DBSet(Text1(25).Text, "N", "S") & "," & ValorNulo & ","
        Sql = Sql & DBSet(Text1(15).Text, "T", "S") & "," & DBSet(Text1(16).Text, "T", "S") & "," & DBSet(Text1(18).Text, "T", "S") & "," & DBSet(Text1(17).Text, "T", "S") & ","
        Sql = Sql & DBSet(Text1(19).Text, "T", "S") & "," & DBSet(IBAN, "T", "S") & "," & DBSet(Text1(20).Text, "T") & "," & DBSet(Text1(21).Text, "T") & ","
        
        If Cobrado Then
            Sql = Sql & "1"
        Else
            Sql = Sql & "0"
        End If
        
        ' falta el codusu
        Sql = Sql & "," & DBSet(vUsu.Id, "N")
        
        
        CadValues = CadValues & "(" & Sql & "),"
    
        RS1.MoveNext
    Wend

    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute CadInsert & CadValues
    End If

    InsertaCobros = True
    Exit Function

eInsertaCobros:
    Mens = Mens & Err.Description
End Function


Private Function EsFraCliTraspasada() As Boolean
Dim Sql As String

    Sql = "select estraspasada from factcli where numserie = " & DBSet(Text1(2).Text, "T") & " and numfactu = "
    Sql = Sql & DBSet(Text1(0).Text, "N") & " and anofactu = " & DBSet(Text1(14).Text, "N")
    
    EsFraCliTraspasada = (DevuelveValor(Sql) = 1)
    

End Function




Private Function ModificaFacturaSiiPresentada() As Boolean
Dim C As String
On Error GoTo eModificaDesdeFormAux
    ModificaFacturaSiiPresentada = False
        
    Conn.BeginTrans
        
        
    'Borramos de linfact
    '
    If CadenaDesdeOtroForm <> "" Then
        C = ObtenerWhereCP(True)
        Conn.Execute "DELETE FROM factcli_lineas " & C
            
        
        'insertamos  dedesde tmpfaclin
        C = "INSERT INTO factcli_lineas(numserie,numfactu,fecfactu,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost) VALUES "
        C = C & CadenaDesdeOtroForm
        Conn.Execute C
    End If
    
    If Ampliacion <> "" Then
        CtaBanco = Trim(Mid(Ampliacion, 1, 10))
        C = Trim(Mid(Ampliacion, 11, 10))
        C = "UPDATE factcli SET cuereten = " & DBSet(C, "T", "S")
        If CtaBanco <> "" Then C = C & ", codmacta = " & DBSet(CtaBanco, "T")
        Ampliacion = Mid(Ampliacion, 21)
        C = C & " , observa = " & DBSet(Ampliacion, "T", "S")
        C = C & " WHERE numfactu= " & Data1.Recordset!numfactu & " AND numserie =" & DBSet(Data1.Recordset!NUmSerie, "T") & " AND anofactu =" & Data1.Recordset!Anofactu
        Conn.Execute C
        
        CtaBanco = ""
    End If
    
    'Borramos lineas apuntes
    C = Val(DBSet(Data1.Recordset!no_modifica_apunte, "N"))
    If Val(C) = 0 Then
        NumAsien = Data1.Recordset!NumAsien
        NumDiario_ = Data1.Recordset!NumDiari
        FecFactuAnt = Data1.Recordset!Fecfactu
        If NumAsien > 0 Then
            C = " WHERE (numasien=" & NumAsien & " and fechaent = " & DBSet(FecFactuAnt, "F") & " and numdiari = " & DBSet(NumDiario_, "N") & ") "
            Conn.Execute "DELETE FROM hlinapu " & C
            
            IntegrarFactura (True)
            
    
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


Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtaux3(Index), 3
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYImage KeyAscii, 0 ' cta base
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
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
    CadB1 = Replace(CadB1, "factcli.", "")
    If ModoLineas = 1 Then
        CadB = "INSERT INTO factcli_totales set "
        'Para el where
        
        CadB2 = CadB1
        'Numlinea
        CadB2 = DevuelveDesdeBD("max(numlinea)", "factcli_totales", CadB2, 1)
        CadB2 = Val(CadB2) + 1
        
        CadB1 = Replace(CadB1, " and ", ",")
        CadB1 = CadB1 & ", numlinea = " & CadB2
        'impoiva imporec
        CadB1 = CadB1 & ", porciva  = " & DBSet(txtaux3(2).Tag, "N", "N")
        CadB1 = CadB1 & ", porcrec  = " & DBSet(txtaux3(3).Tag, "N", "N")
        CadB1 = CadB1 & ", codigiva  = " & DBSet(txtaux3(0).Text, "N")
        CadB1 = CadB1 & ", fecfactu  = " & DBSet(Text1(1).Text, "F")
        CadB = CadB & CadB1
        
        
        CadB2 = ""
    Else
        CadB2 = " WHERE " & CadB1 & " AND numlinea =" & txtaux3(0).Tag  'para el where
        
        CadB = "UPDATE factcli_totales set porciva=porciva "
    End If
    CadB1 = ""
    For NumRegElim = 2 To 4
        CadB1 = CadB1 & ", " & RecuperaValor("baseimpo|impoiva|imporec|", NumRegElim - 1) & "=" & DBSet(txtaux3(NumRegElim), "N", "N")
    Next
    CadB = CadB & CadB1 & CadB2
    
    If Ejecuta(CadB, False) Then AnyadirModificarIVA = True
    
End Function
