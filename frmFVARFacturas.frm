VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFVARFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Varias"
   ClientHeight    =   10755
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15660
   Icon            =   "frmFVARFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10755
   ScaleWidth      =   15660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Index           =   0
      Left            =   13800
      TabIndex        =   98
      Top             =   270
      Width           =   1785
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   96
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   97
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5340
      TabIndex        =   94
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   95
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3915
      TabIndex        =   92
      Top             =   0
      Width           =   1350
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   225
         TabIndex        =   93
         Top             =   180
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Modificacion Totales"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Carga Masiva"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Envio Masivo"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Contabilizacion"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Retención"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1590
      Left            =   225
      TabIndex        =   75
      Top             =   5310
      Width           =   15210
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
         Index           =   35
         Left            =   5385
         MaxLength       =   20
         TabIndex        =   38
         Tag             =   "RCatas|T|S|||fvarfactura|CatastralREF|||"
         Text            =   "12345678901234567890"
         Top             =   990
         Visible         =   0   'False
         Width           =   2625
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
         ItemData        =   "frmFVARFacturas.frx":000C
         Left            =   10485
         List            =   "frmFVARFacturas.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Tag             =   "Situacion inmueble|N|S|||fvarfactura|CatastralSitu|||"
         Top             =   990
         Visible         =   0   'False
         Width           =   4485
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
         ItemData        =   "frmFVARFacturas.frx":0010
         Left            =   180
         List            =   "frmFVARFacturas.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "Tipo retencion|N|N|||fvarfactura|tiporeten|||"
         Top             =   990
         Width           =   5100
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
         Index           =   26
         Left            =   945
         MaxLength       =   6
         TabIndex        =   34
         Tag             =   "% Ret|N|S|0|100.00|fvarfactura|retfaccl|##0.00|N|"
         Text            =   "99.99"
         Top             =   270
         Width           =   645
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
         Index           =   27
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   35
         Tag             =   "Cta.Contable|T|S|||fvarfactura|cuereten|||"
         Text            =   "1234567890"
         Top             =   270
         Width           =   1350
      End
      Begin VB.TextBox Text2 
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
         Index           =   27
         Left            =   5400
         TabIndex        =   76
         Top             =   270
         Width           =   4785
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
         Index           =   28
         Left            =   12420
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "Importe Retención|N|S|||fvarfactura|trefaccl|#,###,###,##0.00|N|"
         Top             =   270
         Width           =   1635
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
         Height          =   195
         Index           =   26
         Left            =   5385
         TabIndex        =   89
         Top             =   720
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Situación inmueble"
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
         Index           =   25
         Left            =   10485
         TabIndex        =   88
         Top             =   675
         Visible         =   0   'False
         Width           =   1845
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
         Height          =   195
         Left            =   180
         TabIndex        =   87
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "% Ret."
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
         Index           =   12
         Left            =   180
         TabIndex        =   79
         Top             =   270
         Width           =   930
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   3690
         Tag             =   "-1"
         ToolTipText     =   "Buscar Cta Contable"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Contable"
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
         Index           =   17
         Left            =   2250
         TabIndex        =   78
         Top             =   270
         Width           =   1350
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
         Height          =   255
         Index           =   18
         Left            =   10485
         TabIndex        =   77
         Top             =   270
         Width           =   2445
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2550
      Index           =   0
      Left            =   225
      TabIndex        =   53
      Top             =   765
      Width           =   15165
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
         ItemData        =   "frmFVARFacturas.frx":0014
         Left            =   8280
         List            =   "frmFVARFacturas.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   450
         Width           =   4980
      End
      Begin VB.TextBox Text2 
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
         Left            =   765
         MaxLength       =   60
         TabIndex        =   99
         Top             =   450
         Width           =   4200
      End
      Begin VB.TextBox Text2 
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
         Index           =   34
         Left            =   5805
         TabIndex        =   85
         Top             =   2115
         Width           =   2175
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
         Index           =   34
         Left            =   5445
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "Cod.Pais|T|S|||fvarfactura|codpais|||"
         Text            =   "12"
         Top             =   2115
         Width           =   315
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
         Index           =   33
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "Provincia|T|S|||fvarfactura|desprovi|||"
         Text            =   "1234567890"
         Top             =   2115
         Width           =   3240
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
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "Poblacion|T|S|||fvarfactura|despobla|||"
         Text            =   "1234567890"
         Top             =   1710
         Width           =   4725
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
         Index           =   31
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   8
         Tag             =   "C.Postal|T|S|||fvarfactura|codposta|||"
         Text            =   "123456"
         Top             =   1710
         Width           =   900
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
         Index           =   30
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "Direccion|T|S|||fvarfactura|dirdatos|||"
         Text            =   "1234567890"
         Top             =   1305
         Width           =   4320
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
         Left            =   6165
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "Nif|T|S|||fvarfactura|nifdatos|||"
         Text            =   "1234567890"
         Top             =   1305
         Width           =   1800
      End
      Begin VB.TextBox Text2 
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
         Index           =   25
         Left            =   9000
         TabIndex        =   73
         Top             =   1170
         Width           =   4245
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
         Index           =   25
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Forma de Pago|N|N|||fvarfactura|codforpa|000||"
         Top             =   1170
         Width           =   675
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
         Height          =   600
         Index           =   5
         Left            =   8280
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "Observaciones|T|S|||fvarfactura|observac|||"
         Top             =   1890
         Width           =   6630
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   1
         Left            =   14715
         TabIndex        =   13
         Tag             =   "Contabilizada|N|N|0|1|fvarfactura|intconta|||"
         Top             =   450
         Width           =   255
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
         Index           =   4
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Cta.Contable|T|N|||fvarfactura|codmacta|||"
         Text            =   "1234567890"
         Top             =   900
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         Left            =   5040
         TabIndex        =   2
         Tag             =   "Nº de Factura|N|S|0|9999999|fvarfactura|numfactu|0000000|S|"
         Top             =   450
         Width           =   1470
      End
      Begin VB.TextBox Text2 
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
         Index           =   4
         Left            =   3240
         MaxLength       =   60
         TabIndex        =   5
         Tag             =   "Nombre Cuenta|T|S|||fvarfactura|nommacta|||"
         Top             =   900
         Width           =   4740
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
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
         Left            =   180
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "Num.Serie|T|S|||fvarfactura|numserie||S|"
         Top             =   450
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
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
         Left            =   6615
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||fvarfactura|fecfactu|dd/mm/yyyy|S|"
         Top             =   450
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
         Index           =   36
         Left            =   8280
         MaxLength       =   30
         TabIndex        =   102
         Tag             =   "Tipo factura|T|N|||fvarfactura|codconce340|||"
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
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
         Left            =   855
         MaxLength       =   3
         TabIndex        =   103
         Top             =   450
         Width           =   405
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
         Left            =   8295
         TabIndex        =   101
         Top             =   180
         Width           =   1380
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
         Height          =   255
         Index           =   24
         Left            =   4680
         TabIndex        =   86
         Top             =   2160
         Width           =   390
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   5130
         Tag             =   "-1"
         ToolTipText     =   "Buscar País"
         Top             =   2160
         Width           =   240
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
         Height          =   255
         Index           =   23
         Left            =   180
         TabIndex        =   84
         Top             =   2160
         Width           =   990
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
         Height          =   255
         Index           =   22
         Left            =   2205
         TabIndex        =   83
         Top             =   1710
         Width           =   990
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
         Height          =   255
         Index           =   21
         Left            =   180
         TabIndex        =   82
         Top             =   1755
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
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
         Left            =   180
         TabIndex        =   81
         Top             =   1305
         Width           =   945
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
         Height          =   255
         Index           =   19
         Left            =   5715
         TabIndex        =   80
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pago"
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
         Index           =   5
         Left            =   8280
         TabIndex        =   74
         Top             =   900
         Width           =   1170
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   9495
         Tag             =   "-1"
         ToolTipText     =   "Buscar Forma de Pago"
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   9810
         ToolTipText     =   "Zoom descripción"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label Label29 
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
         Left            =   8280
         TabIndex        =   71
         Top             =   1575
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilizada"
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
         Index           =   7
         Left            =   13320
         TabIndex        =   68
         Top             =   450
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
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
         Left            =   5040
         TabIndex        =   57
         Top             =   180
         Width           =   1350
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   765
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Serie"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   7695
         Picture         =   "frmFVARFacturas.frx":0018
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1530
         Tag             =   "-1"
         ToolTipText     =   "Buscar Cta Contable"
         Top             =   900
         Width           =   240
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
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   56
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
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
         Left            =   6615
         TabIndex        =   55
         Top             =   180
         Width           =   1980
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Contable"
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
         TabIndex        =   54
         Top             =   900
         Width           =   1305
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Lineas Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   2985
      Left            =   225
      TabIndex        =   64
      Top             =   6975
      Width           =   15225
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   225
         TabIndex        =   90
         Top             =   270
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   180
            TabIndex        =   91
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
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   10
         Left            =   10140
         MaxLength       =   15
         TabIndex        =   48
         Tag             =   "Precio|N|S|||fvarfactura_lineas|precio|###,##0.0000||"
         Text            =   "precio"
         Top             =   1920
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   9
         Left            =   9330
         MaxLength       =   15
         TabIndex        =   47
         Tag             =   "Cantidad|N|S|||fvarfactura_lineas|cantidad|##,###,##0.00||"
         Text            =   "cantidad"
         Top             =   1920
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Height          =   285
         Index           =   8
         Left            =   5160
         TabIndex        =   72
         Tag             =   "Iva|N|N|0|99|linfact|tipoiva|00||"
         Top             =   1920
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   6
         Left            =   5850
         MaxLength       =   50
         TabIndex        =   46
         Tag             =   "Ampliación|T|S|||fvarfactura_lineas|ampliaci|||"
         Text            =   "Ampliacion"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.CommandButton btnBuscar 
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
         Height          =   300
         Index           =   0
         Left            =   3420
         MaskColor       =   &H00000000&
         TabIndex        =   67
         ToolTipText     =   "Buscar Concepto"
         Top             =   1920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   5
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   45
         Tag             =   "Concepto|N|N|0|999|fvarfactura_lineas|codconce|000||"
         Text            =   "Concep"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   1
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "Serie|T|N|||fvarfactura_lineas|numserie||S|"
         Text            =   "L"
         Top             =   1935
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   4
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   44
         Tag             =   "Número de línea|N|N|1|99|fvarfactura_lineas|numlinea|00|S|"
         Text            =   "li"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   2
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   42
         Tag             =   "Nº Factura|N|N|0|9999999|fvarfactura_lineas|numfactu|0000000|S|"
         Text            =   "Fac"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   3
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   43
         Tag             =   "Fecha Factura|F|N|||fvarfactura_lineas|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecfactu"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
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
         Height          =   290
         Index           =   7
         Left            =   10860
         MaxLength       =   15
         TabIndex        =   49
         Tag             =   "Importe|N|N|||fvarfactura_lineas|importe|##,###,##0.00||"
         Text            =   "Importe"
         Top             =   1920
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
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
         Height          =   285
         Index           =   0
         Left            =   3645
         TabIndex        =   65
         Top             =   1935
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   4560
         Top             =   240
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
         Caption         =   "AdoAux(0)"
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
         Height          =   1905
         Index           =   0
         Left            =   240
         TabIndex        =   66
         Top             =   915
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   3360
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
   Begin VB.Frame FrameTotFactu 
      Caption         =   "Total Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   1890
      Left            =   225
      TabIndex        =   58
      Top             =   3375
      Width           =   15195
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
         Index           =   22
         Left            =   7800
         MaxLength       =   6
         TabIndex        =   31
         Tag             =   "% REC 3|N|S|0|100.00|fvarfactura|porcrec3|##0.00|N|"
         Top             =   1260
         Width           =   645
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
         Index           =   16
         Left            =   7800
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "% REC 2|N|S|0|100.00|fvarfactura|porcrec2|##0.00|N|"
         Top             =   900
         Width           =   645
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
         Left            =   7800
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "% REC 1|N|S|0|100.00|fvarfactura|porcrec1|##0.00|N|"
         Text            =   "99.99"
         Top             =   495
         Width           =   645
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
         Index           =   23
         Left            =   8925
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "Importe REC 3|N|S|||fvarfactura|imporec3|#,###,###,##0.00|N|"
         Top             =   1275
         Width           =   1950
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
         Index           =   17
         Left            =   8925
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Importe REC 2|N|S|||fvarfactura|imporec2|#,###,###,##0.00|N|"
         Top             =   885
         Width           =   1950
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
         Left            =   8925
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Importe Rec 1|N|S|||fvarfactura|imporec1|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1950
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   24
         Left            =   12435
         MaxLength       =   15
         TabIndex        =   33
         Tag             =   "Total Factura|N|S|||fvarfactura|totalfac|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   2505
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
         Left            =   3210
         MaxLength       =   2
         TabIndex        =   16
         Tag             =   "Tipo IVA 1|N|S|0|99|fvarfactura|tipoiva1|00||"
         Text            =   "12"
         Top             =   510
         Width           =   525
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
         Index           =   13
         Left            =   3210
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "Tipo IVA 2|N|S|0|99|fvarfactura|tipoiva2|00||"
         Top             =   885
         Width           =   525
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
         Index           =   19
         Left            =   3210
         MaxLength       =   2
         TabIndex        =   28
         Tag             =   "Tipo IVA 3|N|S|0|99|fvarfactura|tipoiva3|00||"
         Top             =   1275
         Width           =   525
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
         Left            =   4170
         MaxLength       =   6
         TabIndex        =   17
         Tag             =   "% IVA 1|N|S|0|100.00|fvarfactura|porciva1|##0.00|N|"
         Text            =   "99.99"
         Top             =   510
         Width           =   780
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
         Index           =   14
         Left            =   4170
         MaxLength       =   6
         TabIndex        =   23
         Tag             =   "% IVA 2|N|S|0|100.00|fvarfactura|porciva2|##0.00|N|"
         Top             =   885
         Width           =   780
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
         Index           =   20
         Left            =   4170
         MaxLength       =   6
         TabIndex        =   29
         Tag             =   "% IVA 3|N|S|0|100.00|fvarfactura|porciva3|##0.00|N|"
         Top             =   1275
         Width           =   780
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
         Left            =   5400
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Importe IVA 1|N|S|||fvarfactura|impoiva1|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1920
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
         Index           =   15
         Left            =   5400
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "Importe IVA 2|N|S|||fvarfactura|impoiva2|#,###,###,##0.00|N|"
         Top             =   885
         Width           =   1920
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
         Index           =   21
         Left            =   5400
         MaxLength       =   15
         TabIndex        =   30
         Tag             =   "Importe IVA 3|N|S|||fvarfactura|impoiva3|#,###,###,##0.00|N|"
         Top             =   1275
         Width           =   1920
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
         Index           =   6
         Left            =   195
         MaxLength       =   15
         TabIndex        =   15
         Tag             =   "Base IVA 1|N|S|||fvarfactura|baseiva1|#,###,###,##0.00|N|"
         Text            =   "575757575757557"
         Top             =   495
         Width           =   2010
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
         Left            =   195
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Base IVA 2|N|S|||fvarfactura|baseiva2|#,###,###,##0.00|N|"
         Top             =   885
         Width           =   2010
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
         Index           =   18
         Left            =   195
         MaxLength       =   15
         TabIndex        =   27
         Tag             =   "Base IVA 3|N|S|||fvarfactura|baseiva3|#,###,###,##0.00|N|"
         Top             =   1275
         Width           =   2010
      End
      Begin VB.Label Label1 
         Caption         =   "% Rec."
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
         Index           =   8
         Left            =   7830
         TabIndex        =   70
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Recargo"
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
         Left            =   8925
         TabIndex        =   69
         Top             =   270
         Width           =   1860
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   2910
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   2910
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   2910
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   555
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Total Factura"
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
         Left            =   12435
         TabIndex        =   63
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IVA"
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
         Index           =   14
         Left            =   2895
         TabIndex        =   62
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
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
         Index           =   15
         Left            =   4200
         TabIndex        =   61
         Top             =   240
         Width           =   615
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
         Height          =   255
         Index           =   16
         Left            =   5400
         TabIndex        =   60
         Top             =   240
         Width           =   1410
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
         Height          =   255
         Index           =   10
         Left            =   195
         TabIndex        =   59
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   225
      TabIndex        =   51
      Top             =   10035
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   52
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
      Left            =   14385
      TabIndex        =   41
      Top             =   10230
      Width           =   1065
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
      Left            =   13125
      TabIndex        =   40
      Top             =   10230
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4200
      Top             =   7770
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
      Left            =   14355
      TabIndex        =   50
      Top             =   10215
      Visible         =   0   'False
      Width           =   1065
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
      Begin VB.Menu mn_ModTotales 
         Caption         =   "&Mod.Totales"
         Enabled         =   0   'False
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu mnCargaMasiva 
         Caption         =   "&Carga Masiva"
         HelpContextID   =   2
         Shortcut        =   ^C
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFVARFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Private Const IdPrograma = 421


Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public numfactu As Long
Public NUmSerie As String
Public Tipo As Byte ' 0 schfac normal
                    ' 1 schfacr ajena para el Regaixo

Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies
'   6.-  Modificar totales
'***Variables comuns a tots els formularis*****

Dim ModoLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean
Dim ModificarTotales As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim Indice As Integer 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmCont As frmBasico
Attribute frmCont.VB_VarHelpID = -1

Private WithEvents frmCon As frmFVARConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTipIVA As frmBasico2  'Tipos de IVA de la contabilidad
Attribute frmTipIVA.VB_VarHelpID = -1
Private WithEvents frmFpa As frmBasico2 'Formas de Pago de la tesoreria
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmPais As frmBasico2 ' pais de la contabilidad
Attribute frmPais.VB_VarHelpID = -1
Private WithEvents frmFVARFras As frmFVARFacturasPrev ' ayuda de facturas varias
Attribute frmFVARFras.VB_VarHelpID = -1

Dim CtaAnt As String
Dim FormaPagoAnt As String
Dim ModoModificar As Boolean
Dim ModificaImportes As Boolean ' variable que me indica q hay que modificar lineas de la factura de contabilidad
                                ' y cobros en la tesoreria

Dim BdConta As Integer
Dim BdConta1 As Integer

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim TipForpa As String
Dim TipForpaAnt As String

Dim Sql As String


' utilizado para buscar por checks
Private BuscaChekc As String

Dim CadenaBorrado As String

Dim Seguir As Boolean

Dim Mc As Contadores

Dim AntLetraSer As String
Dim ResultadoFechaContaOK As Boolean
Dim MensajeFechaOkConta As String

Private Sub btnBuscar_Click(Index As Integer)
    ' els formularis als que crida son d'una atra BDA
    TerminaBloquear
    
    Select Case Index
        Case 0 'Conceptos
            Set frmCon = New frmFVARConceptos
            frmCon.DatosADevolverBusqueda = "0|1|2|4|"
            frmCon.CodigoActual = txtAux(5).Text
            frmCon.Show vbModal
            Set frmCon = Nothing
            
    End Select
    
    PonFoco txtAux(5)
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim B As Boolean
Dim vTabla As String
Dim CtaClie As String
Dim cad As String

' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency
    
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    ModoModificar = False
    B = True
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                Set Mc = New Contadores
                If Text1(2).Text <> "" Then i = FechaCorrecta2(CDate(Text1(2).Text))
                If Mc.ConseguirContador(Trim(Text1(0).Text), (i = 0), False) = 0 Then
                    'COMPROBAR NUMERO ASIENTO
                    Text1(1).Text = Mc.Contador
                    If InsertarDesdeForm2(Me, 1) Then
                        B = True
                    Else
                        B = False
                    End If
                    
                    If B Then
                        Data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCab(True) & Ordenacion
                        cad = "numserie = " & DBSet(Trim(Text1(0).Text), "T")
                        cad = cad & " and numfactu = " & DBSet(Text1(1).Text, "N")
                        cad = cad & " and fecfactu = " & DBSet(Text1(2).Text, "F")
                        PosicionarData cad
                        PonerModo 2
                        BotonAnyadirLinea 0
                    Else
                        'SI NO INSERTA debemos devolver el contador
                        Mc.DevolverContador Trim(Text1(2).Text), (i = 0), Mc.Contador
                    End If

                    End If
            Else
                ModoLineas = 0
            End If

        Case 4  'MODIFICAR
            If Not DatosOK Then
                ModoLineas = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ModoModificar = True
                Conn.BeginTrans
                
                PorRet = 0
                If Text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(Text1(26).Text))
                If AdoAux(0).Recordset.RecordCount > 0 Then AdoAux(0).Recordset.MoveFirst
                RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet, Combo1(2).ListIndex

                Text1(28).Text = ""
                If ImpRet <> 0 Then Text1(28).Text = Format(ImpRet, "#,###,###,##0.00")
                Text1(24).Text = Format(TotFac, "#,###,###,##0.00")

                If Text1(8).Text = "" Then Text1(8).Text = "0,00"
                If Text1(9).Text = "" Then Text1(9).Text = "0,00"
                
                If CadenaBorrado <> "" Then
                    Conn.Execute CadenaBorrado
                    CadenaBorrado = ""
                    EliminarLinea
                End If
                
                
                If ModificaDesdeFormulario2(Me, 1) Then
                    If Check1(1).Value = 1 Then
                        MsgBox "Los cambios realizados recuerde hacerlos en el Registro de Iva y Tesoreria.", vbExclamation
                    End If
                    TerminaBloquear
                    PosicionarData "numserie = '" & Trim(Text1(0).Text) & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
                End If
            End If
            
        Case 5 'LLINIES
            Select Case ModoLineas
                Case 1 'afegir llinia
                    InsertarLinea
                Case 2 'modificar llinies
                    ModificarLinea
                    PosicionarData "numserie = '" & Trim(Text1(0).Text) & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
                    Screen.MousePointer = vbDefault
                    Exit Sub
            End Select
            
            
        Case 6  'MODIFICAR TOTALES
            If Not DatosOK Then
                ModoLineas = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ModoModificar = True
                Conn.BeginTrans
                
                If ModificaDesdeFormulario2(Me, 1) Then
                    If Check1(1).Value = 1 Then
                        MsgBox "Los cambios realizados recuerde hacerlos en el Registro de Iva y Tesoreria.", vbExclamation
                        
                    End If
                    TerminaBloquear
                    PosicionarData "numserie = '" & Trim(Text1(0).Text) & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
                End If
            End If
            
            
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not B Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        If ModoModificar Then
            Conn.RollbackTrans
            ModoModificar = False
        End If
    Else
        If ModoModificar Then
            Conn.CommitTrans
            ModoModificar = False
        End If
    End If
End Sub

Private Sub Form_Activate()
   
    If PrimeraVez Then PrimeraVez = False
    
    If Combo1(2).ListIndex = 18 Then ReferenciaCatastral True
     Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Sql2 As String

    PrimeraVez = True
    Me.Icon = frmppal.Icon
    
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
        .Buttons(1).Image = 47 '
        .Buttons(2).Image = 41
        .Buttons(3).Image = 41
        .Buttons(4).Image = 47
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
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
   
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next i
    
    LimpiarCampos   'Limpia los campos TextBox
    For i = 0 To DataGridAux.Count - 1 'neteje tots els grids de llinies
        DataGridAux(i).ClearFields
    Next i
    
    '## A mano
    NombreTabla = "fvarfactura"
    Ordenacion = " ORDER BY numserie, numfactu, fecfactu "
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Sql2 = "Select * from " & NombreTabla & " where false "
    Data1.RecordSource = Sql2
    Data1.Refresh
    
    '[Monica]31/07/2019: tipo de retencion, factura y situacion
    CargarCombo
        
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbLightBlue 'numserie
    End If
    
    ModoLineas = 0
    
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, (Modo = 2) 'carregue els datagrids de llinies
    Next i
    
    If NUmSerie <> "" Then
        Text1(0).Text = Trim(NUmSerie)
        Text1(1).Text = numfactu
        PonerModo 1
        cmdAceptar_Click
    End If


End Sub

Private Sub LimpiarCampos()
    On Error Resume Next

    Limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    '[Monica]18/11/2013: cambios por aridoc
    Me.Check1(1).Value = 0
    
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    Me.Combo1(4).ListIndex = -1
    
    
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Integer, NumReg As Byte
Dim B As Boolean
Dim CtaMultiple As Boolean

On Error GoTo EPonerModo
 
    Modo = Kmodo
    BuscaChekc = ""
    
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    B = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DespalzamientoVisible B And (Data1.Recordset.RecordCount > 1)

    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    '---------------------------------------------
    
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
'    BloqueaTXT Me, b
    Check1(1).Enabled = (Modo = 1)
    
'    BloqueaImgBuscar Me, Modo, ModoLineas
       
    B = (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    For i = 0 To Text1.Count - 1
        Text1(i).Locked = B
    Next i
    
       
       
       
    'Bloquear los campos de clave primaria, NO se puede modificar
    B = Not (Modo = 1) 'solo al insertar/buscar estará activo
    For i = 1 To 1
        BloqueaTXT Text1(i), B
        Text1(i).Enabled = Not B
    Next i
    B = (Modo = 4) Or (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    For i = 2 To 2
        BloqueaTXT Text1(i), B
        Text1(i).Enabled = Not B
    Next i
    For i = 0 To 0
        BloqueaTXT Text1(i), B
        Text1(i).Enabled = Not B
    Next i
    
    '[Monica]27/11/2017: el nombre de la cuenta es bloqueado
    B = (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    BloqueaTXT Text2(4), B
    Text2(4).Enabled = Not B
    
    For i = 6 To 24
        BloqueaTXT Text1(i), Not (Modo = 1 Or (Modo = 4 And ModificarTotales))
    Next i
    
    ' el importe de retencion solo se puede consultar
    BloqueaTXT Text1(28), Not (Modo = 1 Or (Modo = 4 And ModificarTotales))
    Text1(28).Enabled = (Modo = 1 Or (Modo = 4 And ModificarTotales))
    
    Combo1(2).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Combo1(4).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4) And Combo1(2).ListIndex = 3
    Text1(35).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4) And Combo1(2).ListIndex = 3
    
    
'    'Los % de IVA siempre bloqueados
'    BloquearTxt text1(8), True
'    BloquearTxt text1(14), True
'    BloquearTxt text1(20), True
'    'Los % de REC siempre bloqueados
'    BloquearTxt text1(10), True
'    BloquearTxt text1(16), True
'    BloquearTxt text1(22), True
    'El total de la factura siempre bloqueado
'    BloquearTxt text1(24), True
    
    '09/02/2007 no dejo modificar la forma de pago
    B = ((Modo = 4) And Me.Check1(1).Value = 1) Or (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    BloqueaTXT Text1(25), B
    
    
    Text1(24).BackColor = &HCAE3FD

    ' **** si n'hi han imagens de buscar en la capçalera *****
    'BloquearImgBuscar Me, Modo, ModoLineas
    For i = 0 To imgBuscar.Count - 1
        imgBuscar(i).Enabled = (Modo = 3) Or (Modo = 1) Or (Modo = 4 And Me.Check1(1).Value = 0)
        imgBuscar(i).visible = (Modo = 3) Or (Modo = 1) Or (Modo = 4 And Me.Check1(1).Value = 0)
    Next i
    'BloquearImgZoom Me, Modo, ModoLineas
    imgZoom(0).Enabled = (Modo = 3) Or (Modo = 1) Or (Modo = 4)
    imgZoom(0).visible = (Modo = 3) Or (Modo = 1) Or (Modo = 4)
    ' ********************************************************

    B = (Modo = 3) Or (Modo = 1)
    Me.imgBuscar(0).Enabled = B
    Me.imgBuscar(0).visible = B
    
    B = (Modo = 3) Or (Modo = 1) Or (Modo = 4 And Me.Check1(1).Value = 0)
    Me.imgBuscar(5).Enabled = B
    Me.imgBuscar(5).visible = B
    
    
    'Imagen Calendario fechas
    B = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    Me.ImgFec(2).Enabled = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
    Me.ImgFec(2).visible = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
                          
    If (Modo < 2) Or (Modo = 3) Then
        For i = 0 To DataGridAux.Count - 1
            CargaGrid i, False
        Next i
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    For i = 0 To DataGridAux.Count - 1
        DataGridAux(i).Enabled = B
    Next i
    
    ' solo podremos tocar el campo de contabilizado si estamos buscando
    Check1(1).Enabled = (Modo = 1)
    
    
    'b = (Modo = 4)
    B = (Modo = 1) Or (Modo = 4 And ModificarTotales)
    FrameTotFactu.Enabled = B
    
    Frame2(0).Enabled = (Modo = 4 And Not ModificarTotales) Or (Modo <> 4)
    
    B = (Modo = 5)
    Me.FrameAux0.Enabled = (Modo = 2) Or (Modo = 5)
    
'    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
'    PonerOpcionesMenu   'Activar opciones de menu según nivel
'                        'de permisos del usuario

    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

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
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") 'And Modo = 2
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!Especial, "N") And False
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!Especial, "N") And (Modo = 2 Or Modo = 0)
        Me.Toolbar2.Buttons(3).Enabled = DBLet(Rs!Especial, "N") And False
        Me.Toolbar2.Buttons(4).Enabled = DBLet(Rs!Especial, "N") And (Modo = 2 Or Modo = 0)
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2)
        ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        ToolbarAux.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        
        vUsu.LeerFiltros "ariconta", IdPrograma
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el nivel de usuario
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean, bAux As Boolean
Dim i As Byte

    '-----  TOOLBAR DE LA CABECERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    B = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(4).Enabled = B
    Me.mnVerTodos.Enabled = B
    'Insertar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnNuevo.Enabled = B
    'modificar totals
    Toolbar1.Buttons(11).Enabled = B
    Me.mnCargaMasiva.Enabled = B
    
    
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And (Check1(1).Value = 0)
    'Modificar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(9).Enabled = B
    Me.mnEliminar.Enabled = B
    'modificar totals
    Toolbar1.Buttons(10).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Imprimir
    'VRS:2.0.1(3)
    Toolbar1.Buttons(13).Enabled = (Modo = 2)
    Me.mnImprimir.Enabled = (Modo = 2)
    
    '-----------  LINEAS
    ' *** MEU: botons de les llínies de cuentas bancarias,
    ' només es poden gastar quan inserte o modifique clients ****
    'b = (Modo = 3 Or Modo = 4)
    B = (Modo = 3 Or (Modo = 4 And Not ModificarTotales) Or Modo = 2) 'And (Check1(1).Value = 0)
    
    ToolbarAux.Buttons(1).Enabled = B
    If B Then bAux = (B And Me.AdoAux(0).Recordset.RecordCount > 0)
    ToolbarAux.Buttons(2).Enabled = bAux
    ToolbarAux.Buttons(3).Enabled = bAux
    
    'Imprimir en pestaña Comisiones de Productos
'    ToolAux(2).Buttons(6).Enabled = (Modo = 2) Or (Modo = 3) Or (Modo = 4) Or (Modo = 5 And ModoLineas = 0)
    ' ************************************************************
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
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
Dim tabla As String
    
    Select Case Index
        Case 0 'Lineas de factura
                tabla = "fvarfactura_lineas"
                Sql = "SELECT numserie,numfactu,fecfactu,numlinea,fvarfactura_lineas.codconce,fvarconceptos.nomconce, fvarfactura_lineas.tipoiva, ampliaci,"
                Sql = Sql & "cantidad, precio, importe"
                Sql = Sql & " FROM fvarfactura_lineas, fvarconceptos "
                Sql = Sql & " WHERE fvarfactura_lineas.codconce = fvarconceptos.codconce "
    
                If Enlaza Then
                    Sql = Sql & " AND " & ObtenerWhereCab(False)
                Else
                    Sql = Sql & " AND false"
                End If
                Sql = Sql & " ORDER BY " & tabla & ".numlinea "
    End Select
    MontaSQLCarga = Sql
End Function

Private Sub frmC_Selec(vFecha As Date)
    'Fecha
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCont_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'nroserie
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
Dim cad As String
    Text1(25).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codconce
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomartic
End Sub

Private Sub frmFVARFras_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "numserie = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "N") & " and fecfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 3), "F")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmPais_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Text1(34).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(34).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
Dim cad As String
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codsecci
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
'    Text1(0).Text = RecuperaValor(CadenaSeleccion, 4) 'numserie
'    Text1(1).Text = RecuperaValor(CadenaSeleccion, 3) 'numfactu
    
    cad = RecuperaValor(CadenaSeleccion, 5)  'numconta
    If cad <> "" Then BdConta = CInt(cad)  'numero de conta
End Sub

Private Sub frmTipIVA_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(Indice)
    Text1(Indice + 1).Text = RecuperaValor(CadenaSeleccion, 3) '% iva
    If Modo <> 1 Then
        Text1(Indice + 3).Text = RecuperaValor(CadenaSeleccion, 4) '% rec
    End If
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim CuentaAnt As String

   'Screen.MousePointer = vbHourglass
    TerminaBloquear
    
    Select Case Index

        Case 0 'Serie
            Set frmCont = New frmBasico
            AyudaContadores frmCont, Text1(Index).Text, "tiporegi REGEXP '^[0-9]+$' = 0"
            Set frmCont = Nothing
            If Sql <> "" Then
                Text1(0).Text = RecuperaValor(Sql, 1)
                Text2(0).Text = RecuperaValor(Sql, 2)
                Text1_LostFocus 0
                PonFoco Text1(2)
            End If
            
        Case 1, 6 'Cuenta Contable
            If Index = 1 Then
                Indice = 4
                CuentaAnt = Text1(4).Text
            Else
                Indice = 27
            End If
            Set frmCtas = New frmColCtas
            frmCtas.DatosADevolverBusqueda = "0|1|2|"
            frmCtas.ConfigurarBalances = 3  'NUEVO
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            If Modo <> 1 And Index = 4 Then
                If CuentaAnt <> Text1(4).Text Then Text1_LostFocus 4
                PonFoco Text1(4)
            End If
        
            If Index = 6 Then PonFoco Text1(27)
            
        Case 5 'forma de pago
            Set frmFpa = New frmBasico2
            AyudaFPago frmFpa
            Set frmFpa = Nothing
            PonFoco Text1(25)
            
        Case 2, 3, 4 'tiposd de IVA (de la contabilidad)
            If Index = 2 Then Let Indice = 7
            If Index = 3 Then Let Indice = 13
            If Index = 4 Then Let Indice = 19
        
            Set frmTipIVA = New frmBasico2
            AyudaTiposIva frmTipIVA
            Set frmTipIVA = Nothing
            
            PonFoco Text1(Indice)
            If Text1(Indice).Text <> "" Then Text1_LostFocus Indice
        
        
        Case 7 'codigo de pais
            Set frmPais = New frmBasico2
            AyudaPais frmPais
            Set frmPais = Nothing
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub imgFec_Click(Index As Integer)
        Indice = 2
        
        Set frmC = New frmCal
        frmC.Fecha = Now
        If Text1(2).Text <> "" Then frmC.Fecha = CDate(Text1(2).Text)
        frmC.Show vbModal
        Set frmC = Nothing
        PonFoco Text1(2)
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 5
        frmZ.pTitulo = "Observaciones de la Factura"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonFoco Text1(Indice)
    End If
End Sub



Private Sub mn_ModTotales_Click()

    'Comprobaciones
    '--------------
    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/09/2006
    ' quitamos el control de no poder modificar ni eliminar si es 0
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    
    ' ### [Monica] 27/09/2006
    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificarTotales
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Me.Check1(1).Value = 0
End Sub

Private Sub mnCargaMasiva_Click()
    BotonCargaMasiva
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
'    'VRS:2.0.1(3): añadido el boton de imprimir
'    cadTitulo = "Reimpresion de Facturas"
'
'    ' ### [Monica] 11/09/2006
'    '****************************
'    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
'    Dim nomDocu As String 'Nombre de Informe rpt de crystal
'
'    indRPT = 1 'Facturas Varias
'
'    '[Monica]26/05/2016: si es materna cogemos otra impresion de facturas varias
'    If EsSeccionMaterna(text1(3).Text) Then indRPT = 4
'
'
'    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
'    'Nombre fichero .rpt a Imprimir
'    frmImprimir.NombreRPT = nomDocu
'    ' he añadido estas dos lineas para que llame al rpt correspondiente
'
'    cadNombreRPT = nomDocu  ' "rFactgas.rpt"
'    cadFormula = ""
'    cadFormula = cadFormula & "({" & NomTabla & ".numserie} = """ & Trim(text1(0).Text) & """) AND ({" & NomTabla & ".numfactu} = " & text1(1).Text & ") and ({" & NomTabla & ".fecfactu} = cdate(""" & text1(2).Text & """)) "
'
'    '23022007 Monica: la separacion de la bonificacion solo la quieren en Alzira
''    If vParamAplic.Cooperativa = 1 Then cadFormula = cadFormula & " and {slhfac.numalbar} <> 'BONIFICA'" ' AND ({ssocio.impfactu}<=1)"
'
'    cadParam = "|pEmpresa=" & vEmpresa.nomempre & "|" '& "|pCodigoISO="11112"|pCodigoRev="01"|
'    LlamarImprimir

Dim frmFVARImp As frmFVARReimpresion

    Set frmFVARImp = New frmFVARReimpresion

    frmFVARImp.txtCodigo(0) = Text1(0).Text
    frmFVARImp.txtCodigo(1) = Text1(0).Text
    frmFVARImp.txtCodigo(2) = Text1(4).Text
    frmFVARImp.txtCodigo(3) = Text1(4).Text
    frmFVARImp.txtCodigo(4) = Text1(2).Text
    frmFVARImp.txtCodigo(5) = Text1(2).Text
    frmFVARImp.txtCodigo(6) = Text1(1).Text
    frmFVARImp.txtCodigo(7) = Text1(1).Text

    frmFVARImp.Show vbModal


End Sub

Private Sub mnModificar_Click()

    'Comprobaciones
    '--------------
    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/09/2006
    ' quitamos el control de no poder modificar ni eliminar si es 0
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    
    ' ### [Monica] 27/09/2006
    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
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


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cad As String
    
    
    Select Case Button.Index
        Case 5  'Buscar
           mnBuscar_Click
        Case 6  'Todos
            mnVerTodos_Click
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            '++monica:12/02/2008
            If CByte(Data1.Recordset!intconta) = 1 Then
               cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
                     "Recuerde modificar la Contabilidad y Tesoreria correspondiente!!!"
               MsgBox cad, vbExclamation
            End If
            '++
            mnModificar_Click
        Case 3  'Borrar
            '++monica:12/02/2008
            If CByte(Data1.Recordset!intconta) = 1 Then
               cad = "No se permite eliminar una Factura Contabilizada!!!"
               MsgBox cad, vbExclamation
            Else
            '++
                mnEliminar_Click
            End If
'        Case 10 'Rectificativa
'            mn_ModTotales_Click
'        Case 11 'Carga Masiva de Facturas
'            mnCargaMasiva_Click
        Case 8 'Imprimir
            mnImprimir_Click
    End Select
End Sub

Private Sub BotonBuscar()
    'Buscar
    Seguir = True
    
    '[Monica]27/11/2017: datos fiscales
    BloquearDatosFiscales False

    If Modo <> 1 Then
        BdConta = 0
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        'LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonFoco Text1(0)
        Text1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 0)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonFoco Text1(0)
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim cWhere As String
Dim cWhere1 As String
    
    Screen.MousePointer = vbHourglass
    
    cWhere = "(1=1)"
    If CadB <> "" Then cWhere = cWhere & " and " & CadB & " "

    Set frmFVARFras = New frmFVARFacturasPrev
    
    frmFVARFras.DatosADevolverBusqueda = "0|1|2|"
    frmFVARFras.cWhere = cWhere
    frmFVARFras.Show vbModal
    
    Set frmFVARFras = Nothing
    

End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
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
'Ver todos
Dim i As Integer

    LimpiarCampos 'Limpia los Text1
    
    For i = 0 To DataGridAux.Count - 1 'Limpias los DataGrid
        CargaGrid i, False
    Next i
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
'Añadir registro en tabla de expedientes individuales: expincab (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
'    LimpiarDataGrids

    '[Monica]27/11/2017: datos fiscales
    BloquearDatosFiscales True

    Seguir = True
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Quan afegixc pose en Fecha
    Text1(2).Text = Format(Now, "dd/mm/yyyy")

    '[Monica]18/03/2020: valor por defecto para el combo para que no haya retencion
    Me.Combo1(2).ListIndex = 0

    Me.Combo1(0).ListIndex = 0
    Text1(36).Text = "0"
    


    'em posicione en el 1r tab
    PonFoco Text1(0)
End Sub

Private Sub BotonModificar()
    Seguir = True

    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModificarTotales = False
    PonerModo 4
    
   
    '[Monica]27/11/2017: datos fiscales
    BloquearDatosFiscales Not EsCuentaMultiple(Text1(4).Text)

    
    
    ' ### [Monica] 27/09/2006
    ' me guardo los valores anteriores de cuenta contable
    CtaAnt = Text1(4).Text
    
    'Quan modifique pose en la F.Modificación la data actual
    PonFoco Text1(4)
End Sub


Private Sub BotonModificarTotales()
    Seguir = True

    'Añadiremos el boton de aceptar y demas objetos para insertar
    ModificarTotales = True
    PonerModo 4
    
    'Quan modifique pose en la F.Modificación la data actual
    PonFoco Text1(4)
End Sub




'Private Sub BotonRectificar()
'
'    Set frmList = New frmListado
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    frmList.CadTag = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & Text2(3).Text & "|" & Format(Check1(1).Value, "0") & "|"
'    frmList.OpcionListado = 12
'    frmList.Show vbModal
'
'End Sub

Private Sub BotonEliminar()
Dim cad As String
Dim NumFacElim As Long 'Numero de la Factura que se ha Eliminado
Dim NumSecElim As String 'Numero de la Seccion que se ha eliminado

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(Data1.Recordset.Fields(1).Value), FormatoCampo(text1(1))) Then Exit Sub

    cad = "¿Seguro que desea eliminar la factura?"
    cad = cad & vbCrLf & "Serie: " & Data1.Recordset!NUmSerie
    cad = cad & vbCrLf & "Nº: " & Format(Data1.Recordset!numfactu, "######0")
    cad = cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields("fecfactu")
    
    'Borramos
    If MsgBoxA(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumSecElim = Data1.Recordset.Fields(0)
        NumFacElim = Data1.Recordset.Fields(2)
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                'Poner los grid sin apuntar a nada
                'LimpiarDataGrids
                PonerModo 0
            End If
            'Devolvemos contador, si no estamos actualizando
            Set Mc = New Contadores
            Mc.DevolverContador CStr(NumSecElim), i = 0, NumFacElim
            Set Mc = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: pone el formato o los campos de la cabecera
    
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, True
    Next i
    
    If Text1(36).Text = "0" Then
        Combo1(0).ListIndex = 0
    Else
        PosicionarCombo Combo1(0), Asc(Text1(36).Text)
    End If
    Text2(0).Text = DevuelveDesdeBD("nomregis", "contadores", "tiporegi", Text1(0), "T")
    
    Text2(25).Text = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", Text1(25).Text, "N")
    Text2(27).Text = ""
    '[Monica]27/11/2017: nombre del pais
    Text2(34).Text = ""
    If Text1(34).Text <> "" Then
        Text2(34).Text = DevuelveDesdeBD("nompais", "paises", "codpais", Text1(34).Text, "T")
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    PonerModoUsuarioGnral Modo, "ariconta"
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim v

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                PonFoco Text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                PonFoco Text1(0)
        
        Case 5 'LINEAS
            Select Case ModoLineas
                Case 1 'afegir llinia
                    ModoLineas = 0
                    DataGridAux(NumTabMto).AllowAddNew = False
'                    SituarTab (NumTabMto)
                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar  'Modificar
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    'If DataGridAux(NumTabMto).Enabled Then DataGridAux(NumTabMto).SetFocus
                    DataGridAux(NumTabMto).Enabled = True
                    DataGridAux(NumTabMto).SetFocus

                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llinies
                    ModoLineas = 0
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        v = AdoAux(NumTabMto).Recordset.Fields(3) 'el 1 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & v)
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            PosicionarData "numserie = '" & Trim(Text1(0).Text) & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
            
    End Select
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Datos As String
Dim Sql As String
Dim UltNiv As Integer

    On Error GoTo EDatosOK




    DatosOK = False
    B = CompForm2(Me, 1) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
    
    '[Monica]20/06/2017: control de fechas que antes no estaba
    If B And Text1(2).Text <> "" Then
        ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text))
        If ResultadoFechaContaOK > 0 Then
            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
            B = False
        End If
    End If
    
    
    'si hay porcentaje de retencion debe de haber cuenta de retencion e
    If B And Text1(26).Text <> "" And Text1(27).Text = "" Then
        If CInt(Text1(26).Text) <> 0 Then
            MsgBox "Si hay porcentaje de retención debe introducir una cuenta contable asociada. Revise.", vbExclamation
            B = False
        End If
    End If
    
    'cuenta contable de retencion
    
    '[Monica]30/11/2017: volvemos a comprobar el nif y si es incorrecto preguntamos si continuar
    If B Then
        If Text1(29).Text <> "" And Not ModificaImportes Then
            If Not Comprobar_NIF(Text1(29).Text) Then
                If MsgBox("¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then B = False
            End If
        End If
    End If
    
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData(cad As String)
'Dim cad As String
Dim Indicador As String
    
  '  cad = ""
    If SituarDataMULTI(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then
            PonerModo 2
        End If
       
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       'Poner los grid sin apuntar a nada
       'LimpiarDataGrids
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar
        
    Conn.BeginTrans
    vWhere = ObtenerWhereCab(True)

    'Eliminar las Lineas de facturas de proveedor
    Conn.Execute "DELETE FROM fvarfactura_lineas " & vWhere
    
    'Eliminar la CABECERA
    Conn.Execute "Delete from " & NombreTabla & vWhere
               
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
    If Index = 4 Then CtaAnt = Text1(4).Text
    If Index = 0 Then AntLetraSer = Text1(0).Text
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim cad As String, Datos As String
Dim Suma As Currency
Dim i As Integer
Dim CtaMultiple As Boolean
Dim Rs As ADODB.Recordset
Dim RC As String
Dim LeerCCuenta As Boolean
Dim ModificandoLineas As Integer

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 1 'Nº factura
            If Text1(Index).Text <> "" Then FormateaCampo Text1(Index)
                        
        Case 2 'Fecha
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
            Sql = ""
            If Not EsFechaOK(Text1(Index)) Then
                MsgBoxA "Fecha incorrecta", vbExclamation
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
                PonFoco Text1(Index)
                Exit Sub
            End If
            
            
            
        Case 0 'Serie
            If Modo = 1 Then Exit Sub
            If IsNumeric(Text1(Index).Text) Then
                MsgBoxA "Debe ser una letra: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                PonFoco Text1(0)
            End If
            Text1(Index).Text = UCase(Text1(Index).Text)
            If Text1(Index).Text = AntLetraSer Then Exit Sub

            Text2(0).Text = DevuelveValor("select nomregis from contadores where tiporegi = " & DBSet(Text1(0).Text, "T") & " and tiporegi REGEXP '^[0-9]+$' = 0")
            If Text2(0).Text = "0" Then
                MsgBoxA "Letra de serie no existe o no es de facturas de cliente. Reintroduzca.", vbExclamation
                Text2(0).Text = ""
                Text1(0).Text = ""
                PonFoco Text1(0)
            Else
                If Modo = 3 Then
                    ' traemos el contador
                    If Text1(0).Text <> AntLetraSer Then
                        If Text1(2).Text <> "" Then i = FechaCorrecta2(CDate(Text1(2).Text))
                        Sql = "select codconce340 from contadores where tiporegi = " & DBSet(Text1(0).Text, "T")
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
            
            
        Case 4, 27 'Cta Contable
'            If Modo = 1 Then Exit Sub
            If Text1(Index).Text = "" Then Exit Sub
'???
                'Cuenta cliente
                RC = Text1(Index).Text
                i = Index
                
                If CuentaCorrectaUltimoNivel(RC, Sql) Then
                    Text1(Index).Text = RC
                    Text2(i).Text = Sql
                    If Text1(2).Text <> "" Then
                        If Modo > 2 Then
                            If EstaLaCuentaBloqueada2(RC, CDate(Text1(2).Text)) Then
                                MsgBoxA "Cuenta bloqueada: " & RC, vbExclamation
                                Text1(Index).Text = ""
                                Text2(i).Text = ""
                                PonFoco Text1(Index)
                                Exit Sub
                            End If
                        End If
                    End If
                    If Index = 4 Then
                        LeerCCuenta = False
                        If Modo = 3 Then
                            If Text1(Index).Text <> CtaAnt Then LeerCCuenta = True
                        Else
                            If Modo = 4 Then
                                If CtaAnt = "" Then
                                    If Text1(Index).Text <> Data1.Recordset!codmacta Then LeerCCuenta = True
                                Else
                                    If Trim(Text1(Index).Text) <> CtaAnt Then LeerCCuenta = True
                                End If
                            End If
                        End If
                        If LeerCCuenta Then
                        
                            CtaMultiple = EsCuentaMultiple(Text1(4).Text)
                            BloquearDatosFiscales Not CtaMultiple
                            If (CtaAnt <> Text1(4).Text And Modo <> 1) Or Not CtaMultiple Then TraerDatosCuenta Text1(4).Text  'Modo <> 4 And
                            If CtaMultiple Then
                                PonFoco Text2(4)
                            Else
                                TraerDatosCuenta Text1(4).Text
                            End If
                            CtaAnt = Text1(Index).Text
                        
                        End If
                    End If
                    RC = ""
                Else
                    
                    If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                            RC = RellenaCodigoCuenta(Text1(Index).Text)
                            Sql = "La cuenta: " & RC & " no existe.       ¿Desea crearla?"
                            If MsgBoxA(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
                                CadenaDesdeOtroForm = RC
                                Set frmCtas = New frmColCtas
                                frmCtas.DatosADevolverBusqueda = "0|1|"
                                frmCtas.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                                frmCtas.Show vbModal
                                Set frmCtas = Nothing
                                If Text1(Index).Text = RC Then Sql = "" 'Para k no los borre
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
                        Text2(i).Text = ""
                        PonFoco Text1(Index)
                    End If
                    
                    
                End If
'???
            
            
            
'            If Text1(Index).Text = "" Then
'                PonFoco Text1(Index)
'            Else
'                CtaMultiple = EsCuentaMultiple(Text1(4).Text)
'                BloquearDatosFiscales Not CtaMultiple
'                If (CtaAnt <> Text1(4).Text And Modo <> 1) Or Not CtaMultiple Then TraerDatosCuenta Text1(4).Text  'Modo <> 4 And
'                If CtaMultiple Then
'                    PonFoco Text2(4)
'                End If
'            End If
        
        '[Monica]27/11/2017: para el caso de que sea una cuenta multiple
        Case 34 ' codigo de pais
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text = "" Then Exit Sub
            
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = DevuelveDesdeBD("nompais", "paises", "codpais", Text1(Index).Text, "T")
                If Text2(Index) = "" Then
                    MsgBox "No existe el País. Reintroduzca.", vbExclamation
                    PonFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 25 'Forma pago
            If Modo = 1 Then Exit Sub
            Text2(25).Text = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", Text1(25).Text, "N")
            If Text2(25).Text = "" And Text1(25).Text <> "" Then
                MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                Text1(25).Text = ""
                Seguir = False
                PonFoco Text1(Index)
            Else
                Seguir = True
            End If
            
        Case 26 'porcentaje de retencion
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 7
            
        Case 8, 10, 14, 16, 20, 22, 24
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 7
            
            
        Case 5 'despues de las observaciones si estamos insertando despues he de ir al campo de retencion
            If Modo = 1 Then Exit Sub
            If Modo = 3 And Seguir Then PonFoco Text1(26)
            
        Case 6, 9, 11, 12, 15, 17, 18, 21, 23    'IMPORTES Base, IVA
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 1
            
        Case 7, 13, 19 'cod. IVA
           If Text1(Index).Text = "" Then
              Text1(Index + 1).Text = ""
           Else
              Text1(Index + 1).Text = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", Text1(Index).Text, "N")
           End If
              
'        Case 27 'cuenta de retencion
'            Text2(Index).Text = ""
'            If Text1(Index).Text = "" Then Exit Sub
'
'            Text2(27) = PonerNombreCuenta(Text1(27), Modo, , BdConta, True)
'            If Text2(Index).Text = "" Then
'                PonFoco Text1(Index)
'            End If
              
        Case 29 ' nif, se valida
            If Text1(29).Text = "" Or Modo = 1 Then Exit Sub
            
            Text1(Index).Text = UCase(Text1(Index).Text)
            Comprobar_NIF Text1(Index).Text
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 5 Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 3 Or Modo = 4 Then
                Select Case Index
                    Case 2: KEYFecha KeyAscii, 2
    '                Case 3: KEYBusqueda KeyAscii, 0
    '                Case 4: KEYBusqueda KeyAscii, 1
    '                Case 5: KEYBusqueda KeyAscii, 2
    '                Case 7: KEYBusqueda KeyAscii, 3
    '                Case 11: KEYBusqueda KeyAscii, 4
    '                Case 15: KEYBusqueda KeyAscii, 5
    '               ' Case 1: KEYFecha KeyAscii, 1
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    Else
        If Text1(Index) = "" Then KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYBusquedaLin(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text2(Index), Modo
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim cad As String, Datos As String
Dim Suma As Currency
Dim i As Integer
Dim CtaMultiple As Boolean


    If Not PerderFocoGnral(Text2(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub






'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim cad As String
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And _
'       Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    '++monica:12/02/2008
     If CByte(Data1.Recordset!intconta) = 1 Then
        cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
              "Recuerde modificar la Contabilidad y Tesoreria correspondiente!!!"
        MsgBox cad, vbExclamation
     End If
    '++
    
    
     Select Case Button.Index
        Case 1
'            TerminaBloquear
            BotonAnyadirLinea Index
        Case 2
'            TerminaBloquear
            BotonModificarLinea Index
        Case 3
'            TerminaBloquear
            BotonEliminarLinea Index
            If Modo = 4 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
        Case 6 'Imprimir
'            BotonImprimirLinea Index
    End Select
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5

'    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    
    If AdoAux(Index).Recordset.RecordCount = 1 Then
        MsgBox "No se puede borrar un única línea de factura, elimine la factura completa", vbExclamation
        PonerModo 2
        Exit Sub
    End If
    
    
    Eliminar = False

    Select Case Index
        Case 0 'lineas de factura
            Sql = "¿Seguro que desea eliminar la línea?"
            Sql = Sql & vbCrLf & "Nº línea: " & DBLet(AdoAux(Index).Recordset!NumLinea)
            Sql = Sql & vbCrLf & "Concepto: " & DBLet(AdoAux(Index).Recordset!CodConce) & "  " & DBLet(AdoAux(Index).Recordset!NomConce)
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                Eliminar = True
                Sql = "DELETE FROM fvarfactura_lineas "
                Sql = Sql & ObtenerWhereCab(True) & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
    End Select

    If Eliminar Then
        TerminaBloquear
        CadenaBorrado = Sql
        '16022007
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click
                ModificaImportes = False
        End If
        
        'antes estaba debajo de situardata
        CargaGrid Index, True
        SituarDataTrasEliminar AdoAux(Index), NumRegElim, True
        
        
        
    End If

    ModoLineas = 0
    PosicionarData "numserie = '" & Trim(Text1(0).Text) & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")

    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer
Dim SumLin As Currency

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    'If ModificaLineas = 2 Then Exit Sub
    ModoLineas = 1 'Ponemos Modo Añadir Linea

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modifcar Cabecera
        cmdAceptar_Click
        'No se ha insertado la cabecera
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5
'    If b Then BloquearText1 Me, 4 'Si viene de Insertar Cabecera no bloquear los Text1


    'Obtener el numero de linea ha insertar
    Select Case Index
        Case 0: vTabla = "fvarfactura_lineas"
    End Select
    'Obtener el sig. nº de linea a insertar
    vWhere = ObtenerWhereCab(False)
    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

    'Situamos el grid al final
    AnyadirLinea DataGridAux(Index), AdoAux(Index)

    anc = DataGridAux(Index).top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
    End If

    LLamaLineas Index, ModoLineas, anc

    Select Case Index
        Case 0 'lineas factura
            txtAux(1).Text = Text1(0).Text 'serie
            txtAux(2).Text = Text1(1).Text 'factura
            txtAux(3).Text = Text1(2).Text 'fecha
            txtAux(4).Text = NumF 'numlinea
'            FormateaCampo txtAux(3)
            For i = 5 To txtAux.Count
                txtAux(i).Text = ""
            Next i
            txtAux2(0).Text = ""

            'desbloquear la linea (se bloquea al añadir)
'            BloquearTxt txtAux(3), False
            PonFoco txtAux(5)
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
     
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
    
    If Modo = 4 Then 'Modificar Cabecera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
    
    NumTabMto = Index
    PonerModo 5
    
    If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
        i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
        DataGridAux(Index).Scroll 0, i
        DataGridAux(Index).Refresh
    End If
      
    anc = DataGridAux(Index).top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
    End If

    Select Case Index
        Case 0 'lineas de factura
            For J = 1 To 5
                txtAux(J).Text = DataGridAux(Index).Columns(J - 1).Text
            Next J
            txtAux2(0).Text = DataGridAux(Index).Columns(5).Text 'DevuelveDesdeBDNew(cPTours, "concefact", "nomconce", "codconce", DataGridAux(Index).Columns(5).Text, "N")
            txtAux(8).Text = DataGridAux(Index).Columns(6).Text 'DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", DataGridAux(Index).Columns(5).Text, "N")
            txtAux(6).Text = DataGridAux(Index).Columns(7).Text    ' ampliacion
            txtAux(7).Text = DataGridAux(Index).Columns(10).Text   ' importe
            txtAux(9).Text = DataGridAux(Index).Columns(8).Text    ' cantidad
            txtAux(10).Text = DataGridAux(Index).Columns(9).Text  ' precio
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    Select Case Index
        Case 0 'lineas de factura
            PonFoco txtAux(5)
    End Select
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    On Error GoTo ELLamaLin

    DeseleccionaGrid DataGridAux(Index)
    
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    Select Case Index
        Case 0 'lineas de factura
            For jj = 5 To 10
                txtAux(jj).top = alto
                txtAux(jj).visible = B
            Next jj
            txtAux(8).visible = B '[Monica]18/03/2020: antes false
            txtAux(8).Enabled = False
            
            txtAux2(0).top = alto
            txtAux2(0).visible = B
            Me.btnBuscar(0).top = alto
            Me.btnBuscar(0).visible = B
    End Select
    
ELLamaLin:
    Err.Clear
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    
        Case 1 'Modificar total factura
            mn_ModTotales_Click
        Case 2 'Carga Masiva de Facturas
            mnCargaMasiva_Click
    
        Case 3 ' envio masivo
            
        Case 4 ' integracion contable
            frmFVARContabFact.Show vbModal
    End Select


End Sub

Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim cad As String
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And _
'       Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    '++monica:12/02/2008
     If CByte(Data1.Recordset!intconta) = 1 Then
        cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
              "Recuerde modificarla en el Registro de Iva y Tesoreria !!!"
        MsgBoxA cad, vbExclamation
     End If
    '++
    
     Select Case Button.Index
        Case 1
            BotonAnyadirLinea 0
        Case 2
            BotonModificarLinea 0
        Case 3
            BotonEliminarLinea 0
            If Modo = 4 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
     End Select

End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            Select Case Index
                Case 5: KEYBusquedaLin KeyAscii, 0
                Case 6: KEYBusquedaLin KeyAscii, 1
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
    txtAux(Index).Text = Trim(txtAux(Index).Text)

    Select Case Index
        Case 5 ' Concepto
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "fvarconceptos", "nomconce", "codconce", "N")
                txtAux(8).Text = PonerNombreDeCod(txtAux(Index), "fvarconceptos", "tipoiva", "codconce", "N")
                PonerFormatoEntero txtAux(8)
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe el Concepto: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBoxA(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCon = New frmFVARConceptos
                        frmCon.DatosADevolverBusqueda = "0|1|"
                        frmCon.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCon.Show vbModal
                        Set frmCon = Nothing
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonFoco txtAux(Index)
                End If
            Else
                txtAux2(0).Text = ""
            End If
        
        Case 6 ' Ampliacion
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 9 ' cantidad
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 3
                txtAux(7).Text = Round2(CCur(txtAux(9).Text) * CCur(ComprobarCero(txtAux(10).Text)), 2)
                PonerFormatoDecimal txtAux(7), 3
            End If
            
        Case 10 ' precio
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 6) Then
                    txtAux(7).Text = Round2(CCur(ComprobarCero(txtAux(9).Text)) * CCur(txtAux(10).Text), 2)
                    PonerFormatoDecimal txtAux(7), 3
                End If
            End If
        
        Case 7 'Importe
           If Not EsNumerico(txtAux(Index).Text) Then
                MsgBox "El Importe debe ser numérico.", vbExclamation
                On Error Resume Next
                txtAux(Index).Text = ""
                PonFoco txtAux(Index)
                Exit Sub
            End If
            'Es numerico
            PonerFormatoDecimal txtAux(Index), 3
            PonerFocoBtn Me.cmdAceptar
    End Select
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ' si vamos a insertar el importe miramos si podemos calcularlo y no entrar en importe
    If Index = 7 And (txtAux(9).Text <> "" Or txtAux(10).Text <> "") And txtAux(Index).Text = "" Then
        txtAux(Index).Text = Round2(ComprobarCero(txtAux(9).Text) * ComprobarCero(txtAux(10).Text), 2)
'        cmdAceptar.SetFocus
        Exit Sub
    End If
    
    ConseguirFocoLin txtAux(Index)
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim B As Boolean
Dim SumLin As Currency
    
    On Error GoTo EDatosOKLlin

    DatosOkLlin = False
        
    B = CompForm2(Me, 2, nomframe) 'Comprobar formato datos ok
    If Not B Then Exit Function
    
' ### [Monica] 29/09/2006
' he quitado la parte de comprobar la suma de lineas
'    'Comprobar que el Importe del total de las lineas suma el total o menos de la factura
'    SumLin = CCur(SumaLineas(txtAux(4).Text))
'
'    'Le añadimos el importe de linea que vamos a insertar
'    SumLin = SumLin + CCur(txtAux(7).Text)
'
'    'comprobamos que no sobrepase el total de la factura
'    If SumLin > CCur(Text1(18).Text) Then
'        MsgBox "La suma del importe de las lineas no puede ser superior al total de la factura.", vbExclamation
'        b = False
'    End If
    
    DatosOkLlin = B
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean

    SepuedeBorrar = False
    If AdoAux(Index).Recordset.EOF Then Exit Function

    SepuedeBorrar = True
End Function

Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim tots As String

    On Error GoTo ECarga

    'b = DataGridAux(Index).Enabled
    'DataGridAux(Index).Enabled = False
    
    tots = MontaSQLCarga(Index, Enlaza)
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'lineas de factura
            'si es visible|control|tipo campo|nombre campo|ancho control|formato campo|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(5)|T|Código|800|;S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Concepto|3100|;S|txtAux(8)|T|T.Iva|650|;"
            tots = tots & "S|txtAux(6)|T|Ampliación|5100|;S|txtAux(9)|T|Cantidad|1150|;S|txtAux(10)|T|Precio|1500|;S|txtAux(7)|T|Importe|1850|;"
            arregla tots, DataGridAux(Index), Me
'           DataGridAux(Index).Columns(6).Alignment = dbgCenter
'           DataGridAux(Index).Columns(9).Alignment = dbgRight
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registro en las tablas de Lineas: provbanc, provdpto
Dim nomframe As String
Dim B As Boolean
Dim v As Integer

' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency

    On Error Resume Next

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            CargaGrid NumTabMto, True
            v = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(4).Name & " =" & v)
            
'            ' ### [Monica] 29/09/2006
            PorRet = 0
            If Text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(Text1(26).Text))
            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet, Combo1(0).ListIndex

            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                Text1(6 + (6 * i)).Text = ""
                Text1(7 + (6 * i)).Text = ""
                Text1(8 + (6 * i)).Text = ""
                Text1(9 + (6 * i)).Text = ""
                Text1(10 + (6 * i)).Text = ""
                Text1(11 + (6 * i)).Text = ""
            Next i
            Text1(26).Text = ""
            Text1(28).Text = ""
            
            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Tipiva(i) <> 0 Then
                    Text1(6 + (6 * i)).Text = Impbas(i)
                    Text1(7 + (6 * i)).Text = Tipiva(i)
                    Text1(8 + (6 * i)).Text = PorIva(i)
                    Text1(9 + (6 * i)).Text = ImpIva(i)
                    If PorRec(i) <> 0 Then Text1(10 + (6 * i)).Text = PorRec(i)
                    If ImpRec(i) <> 0 Then Text1(11 + (6 * i)).Text = ImpRec(i)
                 End If
'12/03/2007
'                 If Impbas(i) <> 0 Then text1(6 + (6 * i)).Text = Impbas(i)
'                 If PorIva(i) <> 0 Then text1(8 + (6 * i)).Text = PorIva(i)
'                 If impiva(i) <> 0 Then text1(9 + (6 * i)).Text = impiva(i)
'                 If PorRec(i) <> 0 Then text1(10 + (6 * i)).Text = PorRec(i)
'                 If ImpRec(i) <> 0 Then text1(11 + (6 * i)).Text = ImpRec(i)

                 'TotFac = Impbas(i) + impiva(i)
            Next i
            If PorRet <> 0 Then Text1(26).Text = PorRet
            If ImpRet <> 0 Then Text1(28).Text = ImpRet
            Text1(24).Text = TotFac

            If Text1(8).Text = "" Then Text1(8).Text = "0,00"
            If Text1(9).Text = "" Then Text1(9).Text = "0,00"
            
            
'++monica: 10/03/2009
            PonerFormatos
'++
            
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click
                ModificaImportes = False
            End If

            LLamaLineas NumTabMto, 0
            
            If B Then BotonAnyadirLinea NumTabMto
        End If
    End If
End Sub

Private Sub ModificarLinea()
'Modifica registro en las tablas de Lineas: provbanc, provdpto
Dim nomframe As String
Dim v As Currency

' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency
    
    'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency


    On Error GoTo EModificarLin

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
'        conn.BeginTrans
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
            
            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva
            
            'BotonModificar
                
            End If
            v = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
            PorRet = 0
            If Text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(Text1(26).Text))
        
            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet, Combo1(0).ListIndex

            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                Text1(6 + (6 * i)).Text = ""
                Text1(7 + (6 * i)).Text = ""
                Text1(8 + (6 * i)).Text = ""
                Text1(9 + (6 * i)).Text = ""
                Text1(10 + (6 * i)).Text = ""
                Text1(11 + (6 * i)).Text = ""
            Next i

            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Impbas(i) <> 0 Then Text1(6 + (6 * i)).Text = Impbas(i)
                 If Tipiva(i) <> 0 Then Text1(7 + (6 * i)).Text = Tipiva(i)
                 If PorIva(i) <> 0 Then Text1(8 + (6 * i)).Text = PorIva(i)
                 If ImpIva(i) <> 0 Then Text1(9 + (6 * i)).Text = ImpIva(i)
                 If PorRec(i) <> 0 Then Text1(10 + (6 * i)).Text = PorRec(i)
                 If ImpRec(i) <> 0 Then Text1(11 + (6 * i)).Text = ImpRec(i)

                 'TotFac = Impbas(i) + impiva(i)
            Next i
            Text1(24).Text = TotFac
            If ImpRet <> 0 Then Text1(28).Text = ImpRet
            
            If Text1(8).Text = "" Then Text1(8).Text = "0,00"
            If Text1(9).Text = "" Then Text1(9).Text = "0,00"
            
'++monica: 10/03/2009
            PonerFormatos
'++
            
            
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
'--monica:10/03/2009
'                PonerCamposForma Me, Me.Data1
                BotonModificar
                cmdAceptar_Click
                ModificaImportes = False
            End If

            LLamaLineas NumTabMto, 0
        End If
    End If
    Exit Sub
    
EModificarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    vWhere = ""
    If conW Then vWhere = " WHERE "
    vWhere = vWhere & " numserie='" & Trim(Text1(0).Text) & "'"
    vWhere = vWhere & " AND numfactu= " & Text1(1).Text & " AND fecfactu= '" & Format(Text1(2).Text, FormatoFecha) & "'"
    ObtenerWhereCab = vWhere
End Function



Private Function SumaLineas(NumLin As String) As String
'Al Insertar o Modificar linea sumamos todas las lineas excepto la que estamos
'Insertando o modificando que su valor sera el del txtaux(4).text
'En el DatosOK de la factura sumamos todas las lineas
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim SumLin As Currency

    SumLin = 0
    Sql = "SELECT SUM(importe) FROM fvarfactura_lineas "
    Sql = Sql & ObtenerWhereCab(True)
    If NumLin <> "" Then Sql = Sql & " AND numlinea<>" & DBSet(txtAux(4).Text, "N") 'numlinea
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'En SumLin tenemos la suma de las lineas ya insertadas
        SumLin = CCur(DBLet(Rs.Fields(0), "N"))
    End If
    Rs.Close
    Set Rs = Nothing
    SumaLineas = CStr(SumLin)
End Function


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


Private Function FacturaModificable(NUmSerie As String, numfactu As String, FecFactu As String, Contabil As String) As Boolean

    FacturaModificable = False
    
    If Contabil = 0 Then
        FacturaModificable = True
    Else
        ' si la factura esta contabilizada tenemos que ver si en la contabilidad esta contabilizada y
        ' si en la tesoreria esta remesada o cobrada en estos casos la factura no puede ser modificada
        If FacturaContabilizada(NUmSerie, numfactu, Year(CDate(FecFactu))) Then
            MsgBox "Factura contabilizada en la Contabilidad, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaRemesada(NUmSerie, numfactu, FecFactu) Then
            MsgBox "Factura Remesada, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaCobrada(NUmSerie, numfactu, FecFactu) Then
            MsgBox "Factura Cobrada, no puede modificarse ni eliminarse."
            Exit Function
        End If
           
        FacturaModificable = True
    End If

End Function



Private Sub ActivarFrameCobros()
Dim Obj As Object

For Each Obj In Me
    If TypeOf Obj Is frame Then
        If Obj.Name = "FrameCobros" Then
            
            
        End If
        
    End If
Next Obj

End Sub


Private Sub EliminarLinea()
Dim nomframe As String
Dim v As Currency
Dim Sql As String

    
 
' variables para el recalculo de iva y totales
    Dim i As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIva(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpRec(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

    'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency


    On Error GoTo EEliminarLin

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select
    

    TerminaBloquear
'        conn.BeginTrans
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then

            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva

            'BotonModificar

            End If
            ModoLineas = 0
'            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True

'            SituarTab (NumTabMto)

' [Monica] 25/01/2010 Daba error cuando elimina linea he quitado el setfocus
'            DataGridAux(NumTabMto).SetFocus

'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)

'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
            PorRet = 0
            If Text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(Text1(26).Text))

            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIva, PorIva, TotFac, ImpRec, PorRec, PorRet, ImpRet, Combo1(0).ListIndex


            '13/02/2007 iniacializo los txt
            For i = 0 To 2
                Text1(6 + (6 * i)).Text = ""
                Text1(7 + (6 * i)).Text = ""
                Text1(8 + (6 * i)).Text = ""
                Text1(9 + (6 * i)).Text = ""
                Text1(10 + (6 * i)).Text = ""
                Text1(11 + (6 * i)).Text = ""
            Next i

            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For i = 0 To 2
                 If Impbas(i) <> 0 Then Text1(6 + (6 * i)).Text = Impbas(i)
                 If Tipiva(i) <> 0 Then Text1(7 + (6 * i)).Text = Tipiva(i)
                 If PorIva(i) <> 0 Then Text1(8 + (6 * i)).Text = PorIva(i)
                 If ImpIva(i) <> 0 Then Text1(9 + (6 * i)).Text = ImpIva(i)
                 If PorRec(i) <> 0 Then Text1(10 + (6 * i)).Text = PorRec(i)
                 If ImpRec(i) <> 0 Then Text1(11 + (6 * i)).Text = ImpRec(i)

                 'TotFac = Impbas(i) + impiva(i)
            Next i
            Text1(24).Text = TotFac
            If ImpRet <> 0 Then Text1(28).Text = ImpRet
            
            If Text1(8).Text = "" Then Text1(8).Text = "0,00"
            If Text1(9).Text = "" Then Text1(9).Text = "0,00"
            
            
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                ModificaImportes = True
'                BotonModificar
'                cmdAceptar_Click
'            End If

'++monica: 10/03/2009
            PonerFormatos
'++
            LLamaLineas NumTabMto, 0
    Exit Sub
    
EEliminarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Linea", Err.Description
End Sub

Private Sub PonerFormatos()
Dim mTag As CTag
Dim i As Integer

    Set mTag = New CTag
    For i = 6 To 24
        mTag.Cargar Text1(i)
        If mTag.Formato <> "" And CStr(Text1(i).Text) <> "" Then
             Text1(i).Text = Format(Text1(i).Text, mTag.Formato)
        End If
    Next i
    Set mTag = Nothing

End Sub

Private Sub BotonCargaMasiva()
    frmFVARCargaMasiva.Show vbModal
End Sub


Private Function EsFechaOKConta(Fecha As Date) As Byte
Dim F2 As Date

    If vParam.fechaini > Fecha Then
        EsFechaOKConta = 1
    Else
        F2 = DateAdd("yyyy", 1, vParam.fechafin)
        If Fecha > F2 Then
            EsFechaOKConta = 2
        Else
            'OK. Dentro de los ejercicios contables
            EsFechaOKConta = 0
        End If
    End If
    '[Monica]20/06/2017: de david
    If EsFechaOKConta = 0 Then
        'Si tiene SII
            If vParam.SIITiene Then
'                If DateDiff("d", Fecha, Now) > vEmpresaFac.SIIDiasAviso Then
                '[Monica]19/02/2018: fines de semana
                If Fecha < UltimaFechaCorrectaSII(vParam.SIIDiasAviso, Now) Then
                    MensajeFechaOkConta = "Fecha fuera de periodo de comunicación SII."
                    'LLEVA SII y han trascurrido los dias
                    If vUsu.Nivel = 0 Then
                        If MsgBox(MensajeFechaOkConta & vbCrLf & "¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then
                            EsFechaOKConta = 4
                        End If
                    Else
                        'NO tienen nivel
                        EsFechaOKConta = 5
                    End If
                End If
            End If
    Else
        MensajeFechaOkConta = "Fuera de ejercicios contables"
    End If

End Function


Private Function EsCuentaMultiple(codmacta As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    EsCuentaMultiple = False

    Sql = "select esctamultiple from cuentas where codmacta = " & DBSet(codmacta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not Rs.EOF Then
        Sql = DBLet(Rs!esctamultiple, "N")
    Else
        Sql = 0
    End If
    EsCuentaMultiple = (Sql = 1)
    Set Rs = Nothing
    
End Function

Private Sub BloquearDatosFiscales(bloqueo As Boolean)
Dim i As Integer
    
    If Modo = 5 Then Exit Sub


    Text2(4).Enabled = Not bloqueo
    Text2(4).Locked = bloqueo
    
    For i = 29 To 34
        Text1(i).Enabled = Not bloqueo
        Text1(i).Locked = bloqueo
    Next i
    
    imgBuscar(7).Enabled = Not bloqueo
    imgBuscar(7).visible = Not bloqueo
    
    
    
    If Text2(4).Enabled Then ' blanco
        Text2(4).BackColor = &H80000005
    Else ' amarillo
        Text2(4).BackColor = &H80000018
    End If
End Sub


Private Sub TraerDatosCuenta(Cuenta As String)
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eTraerDatosCuenta

    Sql = "select * from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not Rs.EOF Then
        Text1(29).Text = DBLet(Rs!nifdatos, "T")
        Text1(30).Text = DBLet(Rs!dirdatos, "T")
        Text1(31).Text = DBLet(Rs!codposta, "T")
        Text1(32).Text = DBLet(Rs!desPobla, "T")
        Text1(33).Text = DBLet(Rs!desProvi, "T")
        Text1(34).Text = DBLet(Rs!codpais, "T")
        If Text1(34).Text <> "" Then Text2(34).Text = DevuelveDesdeBD("nompais", "paises", "codpais", Text1(34).Text, "T")
        
        Text1(25).Text = DBLet(Rs!Forpa, "T")
        If Text1(25).Text <> "" Then Text2(25).Text = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", Text1(25).Text, "N")
            
    End If
    Set Rs = Nothing
    Exit Sub
     
eTraerDatosCuenta:
    MuestraError Err.Number, "Traer Datos Cuenta", Err.Description
End Sub

Private Sub CargarCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim J As Long
Dim i As Long
        
    Combo1(0).Clear
    Combo1(2).Clear
    Combo1(4).Clear

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
'
'    'Tipo de operacion
'    Set Rs = New ADODB.Recordset
'    SQL = "SELECT * FROM usuarios.wtipopera where codigo <= 3 ORDER BY codigo"
'    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not Rs.EOF
'        Combo1(1).AddItem Rs!denominacion
'        Combo1(1).ItemData(Combo1(1).NewIndex) = Rs!Codigo
'        Rs.MoveNext
'    Wend
'    Rs.Close
'    Set Rs = Nothing

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
    
    'Tipo situacion inmueble
  
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM usuarios.wtipoinmueble ORDER BY codigo"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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

Private Sub Combo1_Click(Index As Integer)
    If PrimeraVez Then Exit Sub
    If Index = 2 And (Modo = 3 Or Modo = 4) Then
        If Combo1(Index).ListIndex = 0 Then
            Text1(26).Text = ""
            Text1(27).Text = ""
            Text2(27).Text = ""
            Text1(35).Text = ""
            Combo1(0).ListIndex = -1
        End If
    End If
    
    If Index = 0 And (Modo = 3 Or Modo = 4) Then
        If Combo1(Index).ListIndex = 0 Then
            Text1(26).Text = ""
            Text1(27).Text = ""
            Text2(27).Text = ""
            Text1(28).Text = ""
            Text1(30).Text = ""
            Combo1(4).ListIndex = -1
            Combo1(2).ListIndex = 0
        End If
    End If
    
    If Index = 0 And (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        If Combo1(0).ListIndex = 0 Then
            Text1(36).Text = "0"
        Else
            If Combo1(0).ListIndex <> -1 Then Text1(36).Text = Chr(Combo1(0).ItemData(Combo1(0).ListIndex))
        End If
        
    End If
    
    If Index = 0 And (Modo = 1 Or Modo = 3 Or Modo = 4) Then
        If Combo1(Index).ListIndex = 18 Then
            Text1(35).Enabled = True
            Combo1(4).Enabled = True
        End If
    End If
    If Combo1(0).ListIndex = 18 Then
        ReferenciaCatastral True
    Else
        ReferenciaCatastral False
    End If
    
    
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub ReferenciaCatastral(visible As Boolean)
    Text1(35).visible = visible
    Combo1(4).visible = visible
    Label1(26).visible = visible
    Label1(25).visible = visible
End Sub





