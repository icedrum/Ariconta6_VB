VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBanco 
   Caption         =   "Bancos"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   45
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.CheckBox chkBanco 
      Caption         =   "Cuenta principal transferencias de clientes"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Tag             =   "Cta Transferencia Clientes|N|N|0|1|bancos|ctatransfercli|||"
      Top             =   7680
      Width           =   5685
   End
   Begin VB.Frame FrameDesplazamiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3810
      TabIndex        =   65
      Top             =   180
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   66
         Top             =   180
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   63
      Top             =   180
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   64
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
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Height          =   300
      Left            =   9150
      TabIndex        =   0
      Top             =   300
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   360
      Index           =   23
      Left            =   8190
      MaxLength       =   3
      TabIndex        =   11
      Tag             =   "Cedante|T|S|||bancos|Sufijo3414|||"
      Text            =   "Tex"
      Top             =   3030
      Width           =   1605
   End
   Begin VB.CheckBox chkBanco 
      Caption         =   "Gastos bancarios en pagos separados de apunte banco"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Tag             =   "G.transfer|N|S|||bancos|GastTransDescontad|||"
      Top             =   7320
      Width           =   5925
   End
   Begin VB.Frame FrameAnalitica 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6150
      TabIndex        =   60
      Top             =   6990
      Width           =   5685
   End
   Begin VB.CheckBox chkBanco 
      Caption         =   "Gastos bancarios en cobros separados de apunte banco"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Tag             =   "G.Rem.|N|S|||bancos|GastRemDescontad|||"
      Top             =   6960
      Width           =   5925
   End
   Begin VB.Frame Frame3 
      Caption         =   "Remesas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   55
      Top             =   5940
      Width           =   11715
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   22
         Left            =   7650
         MaxLength       =   15
         TabIndex        =   22
         Tag             =   "Talon dias|N|S|0||bancos|remesadiasmenor|||"
         Text            =   "Text1"
         Top             =   330
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   21
         Left            =   5040
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Talon dias|N|S|0||bancos|remesadiasmayor|||"
         Text            =   "Text1"
         Top             =   330
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   19
         Left            =   1710
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Riesgo|N|S|0||bancos|remesariesgo|#,##0.00||"
         Text            =   "Text1"
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   18
         Left            =   9900
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Riesgo|N|S|0||bancos|remesamaximo|#,##0.00||"
         Text            =   "Text1"
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo menor"
         Height          =   285
         Index           =   22
         Left            =   5820
         TabIndex        =   59
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo mayor"
         Height          =   285
         Index           =   21
         Left            =   3270
         TabIndex        =   58
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Importe riesgo"
         Height          =   285
         Index           =   19
         Left            =   270
         TabIndex        =   57
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Importe m�ximo "
         Height          =   285
         Index           =   18
         Left            =   8250
         TabIndex        =   56
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.Frame FramePagares 
      Caption         =   "Pagar�s"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6060
      TabIndex        =   52
      Top             =   5010
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   17
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Talon dias|N|S|0||bancos|pagaredias|||"
         Text            =   "Text1"
         Top             =   330
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   16
         Left            =   3960
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Riesgo|N|S|0||bancos|pagareriesgo|#,##0.00||"
         Text            =   "Text1"
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo"
         Height          =   315
         Index           =   17
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Importe m�ximo "
         Height          =   255
         Index           =   16
         Left            =   2280
         TabIndex        =   53
         Top             =   360
         Width           =   1770
      End
   End
   Begin VB.Frame FrameTalones 
      Caption         =   "Talones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   49
      Top             =   5010
      Width           =   5745
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   15
         Left            =   4050
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Riesgo|N|S|0||bancos|talonriesgo|#,##0.00||"
         Text            =   "Text1"
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   14
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "Talon dias|N|S|0||bancos|talondias|||"
         Text            =   "Text1"
         Top             =   330
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Importe m�ximo "
         Height          =   285
         Index           =   15
         Left            =   2370
         TabIndex        =   51
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Dias riesgo"
         Height          =   255
         Index           =   14
         Left            =   270
         TabIndex        =   50
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   6030
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "Cta. gastos|T|S|||bancos|ctaefectosdesc|||"
      Text            =   "Text1"
      Top             =   4500
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   13
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "Text2"
      Top             =   4500
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   12
      Left            =   120
      MaxLength       =   30
      TabIndex        =   14
      Tag             =   "Cta. gastos|T|S|||bancos|ctagastostarj|||"
      Text            =   "Text1"
      Top             =   4500
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   12
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   4500
      Width           =   4395
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   10
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   3810
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   10
      Left            =   6030
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "Cta. gastos|T|S|||bancos|ctaingreso|||"
      Text            =   "Text1"
      Top             =   3810
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   360
      Index           =   9
      Left            =   6030
      MaxLength       =   3
      TabIndex        =   10
      Tag             =   "Sufijo OEM|T|S|||bancos|sufijoem|||"
      Text            =   "Tex"
      Top             =   3030
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   2
      Left            =   7470
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Text2"
      Top             =   7320
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   8
      Left            =   6270
      MaxLength       =   4
      TabIndex        =   29
      Tag             =   "Centro Coste|T|S|||bancos|codccost|||"
      Text            =   "Text"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   7
      Left            =   150
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "Contrato Confirming|T|S|||bancos|CaixaConfirming|||"
      Text            =   "Text1"
      Top             =   3030
      Width           =   5685
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6000
      TabIndex        =   38
      Top             =   1110
      Width           =   5715
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   25
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   420
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   24
         Left            =   3765
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   420
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   3
         Left            =   1035
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   420
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   2
         Left            =   120
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   420
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   6
         Left            =   1950
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   420
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   0
         Left            =   2850
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   420
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN"
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   61
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   120
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "Cta. gastos|T|S|||bancos|ctagastos|||"
      Text            =   "Text1"
      Top             =   3810
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   5
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   3810
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Tag             =   "Cta. contable|T|N|||bancos|codmacta||S|"
      Text            =   "0000000000"
      Top             =   1530
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   360
      Index           =   4
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   1530
      Width           =   4425
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   120
      MaxLength       =   40
      TabIndex        =   8
      Tag             =   "Descripcion|T|S|||bancos|descripcion|||"
      Text            =   "Text1"
      Top             =   2370
      Width           =   5715
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   90
      TabIndex        =   31
      Top             =   8040
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   210
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10800
      TabIndex        =   28
      Top             =   8145
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9600
      TabIndex        =   27
      Top             =   8145
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   240
      Top             =   8070
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
      Cancel          =   -1  'True
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10800
      TabIndex        =   30
      Top             =   8160
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   11310
      TabIndex        =   67
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
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   11
      Left            =   2940
      MaxLength       =   10
      TabIndex        =   68
      Tag             =   "idnorma34|T|S|||bancos|idnorma34|||"
      Text            =   "Text1"
      Top             =   3030
      Width           =   1875
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   20
      Left            =   6840
      MaxLength       =   40
      TabIndex        =   69
      Tag             =   "Iban|T|S|||bancos|iban|||"
      Text            =   "Text"
      Top             =   1290
      Width           =   4290
   End
   Begin VB.Label Label1 
      Caption         =   "Sufijo transferencias"
      Height          =   255
      Index           =   25
      Left            =   8190
      TabIndex        =   62
      Top             =   2760
      Width           =   2595
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   13
      Left            =   8580
      ToolTipText     =   "Cta efectos descontados"
      Top             =   4260
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta efectos descontados"
      Height          =   240
      Index           =   13
      Left            =   6030
      TabIndex        =   48
      Top             =   4260
      Width           =   2505
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   12
      Left            =   1980
      ToolTipText     =   "Cuenta tarjeta"
      Top             =   4260
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta gastos tarjeta"
      Height          =   240
      Index           =   12
      Left            =   120
      TabIndex        =   46
      Top             =   4260
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta ingresos"
      Height          =   240
      Index           =   10
      Left            =   6030
      TabIndex        =   44
      Top             =   3570
      Width           =   1230
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   10
      Left            =   7260
      ToolTipText     =   "Cuenta ingresos"
      Top             =   3570
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Sufijo adeudos"
      Height          =   255
      Index           =   9
      Left            =   6030
      TabIndex        =   42
      Top             =   2790
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Centro coste"
      Height          =   240
      Index           =   8
      Left            =   6270
      TabIndex        =   41
      Top             =   7065
      Width           =   1290
   End
   Begin VB.Image imgCC 
      Height          =   240
      Left            =   7620
      Picture         =   "frmBanco.frx":000C
      Top             =   7020
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Contrato Confirming"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   39
      Top             =   2790
      Width           =   2895
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   5
      Left            =   1230
      ToolTipText     =   "Cuenta gastos"
      Top             =   3570
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cta gastos"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   37
      Top             =   3570
      Width           =   1110
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   4
      Left            =   1770
      ToolTipText     =   "Cuenta contable"
      Top             =   1230
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta contable"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   35
      Top             =   1215
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Descripci�n"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   2130
      Width           =   2025
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^F
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
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 207

Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico
Attribute frmCC.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'  Variables comunes a todos los formularios
Private Modo As Byte
Private CadenaConsulta As String
Private CadB As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private DevfrmCCtas As String

Private BuscaChekc As String


Private Sub chkBanco_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkBanco(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkBanco(" & Index & ")|"
    End If
End Sub

Private Sub chkBanco_GotFocus(Index As Integer)
    ConseguirFocoChk Modo
End Sub

Private Sub chkBanco_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub cmdAceptar_Click()
    Dim cad As String
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                PonerModo 0
                lblIndicador.Caption = ""
            End If
        End If
    Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    lblIndicador.Caption = ""
                    If SituarData1 Then
                        PonerModo 2
                    Else
                        LimpiarCampos
                        PonerModo 0
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

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1, 3
    LimpiarCampos
    PonerModo 0
Case 4
    'Modificar
    lblIndicador.Caption = ""
    TerminaBloquear
    PonerModo 2
    PonerCampos
End Select

End Sub


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1() As Boolean
    Dim SQL As String
    On Error GoTo ESituarData1
            'Actualizamos el recordset
            Data1.Refresh
            '#### A mano.
            'El sql para que se situe en el registro en especial es el siguiente
            SQL = " codmacta = " & Text1(4).Text & ""
            Data1.Recordset.Find SQL
            If Data1.Recordset.EOF Then GoTo ESituarData1
            SituarData1 = True
        Exit Function
ESituarData1:
        If Err.Number <> 0 Then Err.Clear
        Limpiar Me
        PonerModo 0
        lblIndicador.Caption = ""
        SituarData1 = False
End Function

Private Sub BotonAnyadir()
    LimpiarCampos
    'A�adiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    SugerirCodigoSiguiente
    '###A mano
    Text1(4).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        '### A mano
        '################################################
        'Si pasamos el control aqui lo ponemos en amarill
        PonFoco Text1(4)
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                PonFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
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
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    'A�adiremos el boton de aceptar y demas objetos para insertar
   ' cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano
    Text1(4).Locked = True
    DespalzamientoVisible False
    Text1(1).SetFocus
End Sub

Private Sub BotonEliminar()

'
    Dim cad As String
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    'Comprobamos si se puede eliminar
    I = 0
    If Not SePuedeEliminar Then I = 1
     
    Set miRsAux = Nothing
    If I = 1 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '### a mano
    cad = "Seguro que desea eliminar de la BD el registro:"
    cad = cad & vbCrLf & "Cta banco: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Decripcion: " & Me.Text2(4).Text
    I = MsgBox(cad, vbQuestion + vbYesNo)
    'Borramos
    If I = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
        Data1.Refresh
        If Data1.Recordset.EOF Then
            'Solo habia un registro
            LimpiarCampos
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
    End If
Error2:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub




Private Sub cmdRegresar_Click()

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If



    
    

    RaiseEvent DatoSeleccionado(CStr(Text1(4).Text & "|" & Text2(4).Text & "|"))
    Unload Me
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Load()
Dim I As Integer


  Me.Icon = frmPpal.Icon


    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    imgCuentas(4).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    imgCuentas(5).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    imgCuentas(10).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    imgCuentas(12).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    imgCuentas(13).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    
    DespalzamientoVisible False


    LimpiarCampos

    'Como son cuentas, como mucho seran
    For I = 4 To 5
        Text1(I).MaxLength = vEmpresa.DigitosUltimoNivel
    Next I
    
    '## A mano
    NombreTabla = "bancos"
    Ordenacion = " ORDER BY codmacta"
        
    PonerOpcionesMenu
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    Data1.RecordSource = "Select * from " & NombreTabla
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
    FrameAnalitica.Visible = Not vParam.autocoste
    If Not vParam.autocoste Then Me.Text1(8).TabIndex = 100
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano

    'Check1.Value = 0
    For kCampo = 0 To 3
        If kCampo <> 2 Then Me.chkBanco(kCampo).Value = 0
    Next
    kCampo = 0
End Sub




Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    CadB = "codmacta = " & RecuperaValor(CadenaSeleccion, 1)
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    'Centro de coste
    Text1(8).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
DevfrmCCtas = CadenaSeleccion
End Sub

Private Sub imgCC_Click()
    'Lanzaremos el vista previa
    Set frmCC = New frmBasico

    AyudaCC frmCC
    
    Set frmCC = Nothing
    
End Sub

Private Sub imgCuentas_Click(Index As Integer)
 Screen.MousePointer = vbHourglass
 Set frmCCtas = New frmColCtas
 DevfrmCCtas = ""
 frmCCtas.DatosADevolverBusqueda = "0"
 frmCCtas.Show vbModal
 Set frmCCtas = Nothing
 If DevfrmCCtas <> "" Then
        Text1(Index).Text = RecuperaValor(DevfrmCCtas, 1)
        Text2(Index).Text = RecuperaValor(DevfrmCCtas, 2)
End If
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
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 4: KEYCta KeyAscii, 4
            Case 5: KEYCta KeyAscii, 5
            Case 10: KEYCta KeyAscii, 10
            Case 12: KEYCta KeyAscii, 12
            Case 13: KEYCta KeyAscii, 13
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYCta(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgCuentas_Click (Indice)
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
    Dim Valor As Currency
    Dim SQL As String
    Dim mTag As CTag
    Dim I As Integer
    Dim SQL2 As String
    
    
    
    If Modo <> 2 Then
        If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    End If
    
    ''Quitamos blancos por los lados
    If Index <> 11 Then Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0, 2, 3, 6, 24, 25
            If Text1(Index).Text = "" Then Exit Sub
            
            If Index = 2 Then
                Text1(Index).Text = UCase(Text1(Index).Text)
            Else
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
            End If
            If Modo = 1 Then Exit Sub
        
            If Index <> 2 Then
                If Not EsNumerico(Text1(Index).Text) Then
                    PonFoco Text1(Index)
                    Exit Sub
                Else
                    Text1(Index).Text = Format(Text1(Index).Text, "0000")
                End If
            
                If Text1(3).Text <> "" And Text1(6).Text <> "" And Text1(0).Text <> "" And Text1(24).Text <> "" And Text1(25).Text <> "" Then
                    ' comprobamos si es correcto
                    SQL = Format(Text1(3).Text, "0000") & Format(Text1(6).Text, "0000") & Format(Text1(0).Text, "0000") & Format(Text1(24).Text, "0000") & Format(Text1(25).Text, "0000")
                End If
            Else
                If Mid(Text1(Index).Text, 1, 2) = "ES" Then
                    If Not IBAN_Correcto(Me.Text1(Index).Text) Then Text1(Index).Text = ""
                End If
            End If
            
            If Text1(2).Text <> "" And Text1(3).Text <> "" And Text1(6).Text <> "" And Text1(0).Text <> "" And Text1(24).Text <> "" And Text1(25).Text <> "" Then
                SQL = Format(Text1(3).Text, "0000") & Format(Text1(6).Text, "0000") & Format(Text1(0).Text, "0000") & Format(Text1(24).Text, "0000") & Format(Text1(25).Text, "0000")
        
                SQL2 = CStr(Mid(Text1(2).Text, 1, 2))
                If DevuelveIBAN2(CStr(SQL2), SQL, SQL) Then
                    If Mid(Text1(2).Text, 3, 2) <> SQL Then
                        MsgBox "Codigo IBAN distinto del calculado [" & SQL2 & SQL & "]", vbExclamation
                    End If
                End If
            End If
            
            Text1(20).Text = Text1(2).Text & Format(ComprobarCero(Text1(3).Text), "0000") & Format(ComprobarCero(Text1(6).Text), "0000") & Format(ComprobarCero(Text1(0).Text), "0000") & Format(ComprobarCero(Text1(24).Text), "0000") & Format(Text1(25).Text, "0000")

             
        Case 20  'IBAN ya no se ve
            
            
        Case 4, 5, 10, 12, 13
            
            If Modo >= 2 Or Modo <= 4 Then
                If Text1(Index).Text = "" Then
                     Text2(Index).Text = SQL
                     Exit Sub
                End If

                DevfrmCCtas = Text1(Index).Text
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                    Text1(Index).Text = DevfrmCCtas
                    Text2(Index).Text = SQL
                Else
                    MsgBox SQL, vbExclamation
                    Text1(Index).Text = ""
                    Text2(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
                DevfrmCCtas = ""
                
            End If
        Case 8
            If Text1(8).Text = "" Then
                Text2(2).Text = ""
                Exit Sub
            End If
            DevfrmCCtas = DevuelveDesdeBD("nomccost", "ccoste", "codccost", Text1(8).Text, "T")
            If DevfrmCCtas = "" Then
                MsgBox "CC no encontrado: " & Text1(8).Text, vbExclamation
                Text1(8).Text = ""
                Exit Sub
            Else
                Text1(8).Text = UCase(Text1(8).Text)
            End If
            Text2(2).Text = DevfrmCCtas
            
        Case 14, 17, 21, 22
            'Dias
            Text1(Index).Text = Trim(Text1(Index).Text)
            If Text1(Index).Text = "" Then Exit Sub
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Campo num�rico: " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
            Else
                Text1(Index).Text = Abs(Val(Text1(Index).Text))
            End If
        Case 15, 16, 18, 19
            'Importe
            Text1(Index).Text = Trim(Text1(Index).Text)
            If Text1(Index).Text = "" Then Exit Sub
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "importe debe ser num�rico", vbExclamation
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            Else
                If InStr(1, Text1(Index).Text, ",") > 0 Then
                    Valor = ImporteFormateado(Text1(Index).Text)
                Else
                    Valor = CCur(TransformaPuntosComas(Text1(Index).Text))
                End If
                Text1(Index).Text = Format(Valor, FormatoImporte)
            End If
                
            
        '....
    End Select
    '---
End Sub

Private Sub HacerBusqueda()
Dim cad As String
Dim CadB As String

CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)

If Text1(2).Text <> "" Then
    If CadB <> "" Then CadB = CadB & " and "
    CadB = CadB & "mid(iban,1,4) = " & DBSet(Text1(2).Text, "T")
End If
If Text1(3).Text <> "" Then
    If CadB <> "" Then CadB = CadB & " and "
    CadB = CadB & "mid(iban,5,4) = " & DBSet(Text1(3).Text, "T")
End If
If Text1(6).Text <> "" Then
    If CadB <> "" Then CadB = CadB & " and "
    CadB = CadB & "mid(iban,9,4) = " & DBSet(Text1(6).Text, "T")
End If
If Text1(0).Text <> "" Then
    If CadB <> "" Then CadB = CadB & " and "
    CadB = CadB & "mid(iban,13,4) = " & DBSet(Text1(0).Text, "T")
End If
If Text1(24).Text <> "" Then
    If CadB <> "" Then CadB = CadB & " and "
    CadB = CadB & "mid(iban,17,4) = " & DBSet(Text1(24).Text, "T")
End If
If Text1(25).Text <> "" Then
    If CadB <> "" Then CadB = CadB & " and "
    CadB = CadB & "mid(iban,21,4) = " & DBSet(Text1(25).Text, "T")
End If



If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)

    Set frmBan = New frmBasico2
    
    AyudaBanco frmBan, , CadB
    
    Set frmBan = Nothing

End Sub



Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
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
    Dim I As Integer
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    PonerCtasIVA
    
    
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(6).Text = ""
    Text1(0).Text = ""
    Text1(24).Text = ""
    Text1(25).Text = ""
    
    Text1(2).ToolTipText = ""
    Text1(3).ToolTipText = ""
    Text1(6).ToolTipText = ""
    Text1(0).ToolTipText = ""
    Text1(24).ToolTipText = ""
    Text1(25).ToolTipText = ""
    
    If Text1(20).Text <> "" Then
        Text1(2).Text = Mid(Text1(20).Text, 1, 4)
        Text1(3).Text = Mid(Text1(20).Text, 5, 4)
        Text1(6).Text = Mid(Text1(20).Text, 9, 4)
        Text1(0).Text = Mid(Text1(20).Text, 13, 4)
        Text1(24).Text = Mid(Text1(20).Text, 17, 4)
        Text1(25).Text = Mid(Text1(20).Text, 21, 4)
        
        Dim CCC As String
        CCC = Text1(2).Text & " " & Text1(3).Text & " " & Text1(6).Text & " " & Mid(Text1(0).Text, 1, 2) & " " & Mid(Text1(0).Text, 3, 2) & Text1(24).Text & Text1(25).Text
        
        Text1(2).ToolTipText = CCC
        Text1(3).ToolTipText = CCC
        Text1(6).ToolTipText = CCC
        Text1(0).ToolTipText = CCC
        Text1(24).ToolTipText = CCC
        Text1(25).ToolTipText = CCC
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim I As Integer
    Dim B As Boolean
    Dim Obj
    
    BuscaChekc = ""
    
    Modo = Kmodo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    B = (Modo = 0 Or Modo = 2)
    
    'chkVistaPrevia.Visible = (Modo = 1)
    
    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B And Me.Data1.Recordset.RecordCount > 1
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    Else
        cmdRegresar.Visible = False
    End If
    
    'Modo insertar o modificar
    B = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.Visible = B Or Modo = 1
    cmdCancelar.Visible = B Or Modo = 1
    mnOpciones.Enabled = Not B
    If cmdCancelar.Visible Then
        cmdCancelar.Cancel = True
        Else
        cmdCancelar.Cancel = False
    End If
    Toolbar1.Buttons(6).Enabled = Not B And vUsu.Nivel < 2
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui a�adiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
    B = (Modo = 2) Or Modo = 0
    For I = 0 To 23
            Text1(I).Locked = B
            If Modo <> 1 Then
                Text1(I).BackColor = vbWhite
            End If
    Next I
    
    For Each Obj In imgCuentas
        Obj.Enabled = Not B
    Next
    Me.imgCC.Enabled = Not B
    
    PonerModoUsuarioGnral Modo, "ariconta"

    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Data1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim SQL As String
Dim RC2 As String

    
    DatosOK = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Comprobamos  si existe
    If Modo = 3 Then
        If ExisteCP(Text1(4)) Then B = False
    End If
    If Not B Then Exit Function
    
    'Si tiene contabilidad analitica EXITGIMOS EL CC
    If vParam.autocoste Then
        If Text1(8).Text = "" Then
            MsgBox "Centro de coste requerido", vbExclamation
            Exit Function
        End If
    End If
    
    'Comprobamos el CCC
    If Text1(2).Text <> "" Then
         SQL = Text1(3).Text & Text1(6).Text & Text1(0).Text & Text1(24).Text & Text1(25).Text
         If Len(SQL) <> 20 Then
             MsgBox "Longitud cuenta bancaria incorrecta", vbExclamation
             Exit Function
         End If

        'Compruebo EL IBAN
        'Meto el CC
        RC2 = SQL
        SQL = ""
        If Me.Text1(2).Text <> "" Then SQL = Mid(Text1(2).Text, 1, 2)

        If DevuelveIBAN2(SQL, RC2, RC2) Then
            If Me.Text1(2).Text = "" Then
                If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(2).Text = RC2
            Else
                If Mid(Text1(2).Text, 3, 2) <> RC2 Then
                    RC2 = "Calculado : " & SQL & RC2
                    RC2 = "Introducido: " & Me.Text1(2).Text & vbCrLf & RC2 & vbCrLf
                    RC2 = "Error en codigo IBAN" & vbCrLf & RC2 & "Continuar?"
                    If MsgBox(RC2, vbQuestion + vbYesNo) = vbNo Then Exit Function
                End If
            End If
        End If
     Else
        Text1(20).Text = ""
     End If
    
    
    
    
    'Si el idNorma34 son espacios en blanco entonces pong "", para que en la BD vaya un NULL
    If Trim(Text1(11).Text) = "" Then Text1(11).Text = ""
    
    If Modo = 3 Or Modo = 4 Then
        SQL = "select count(*) from bancos where codmacta <> " & DBSet(Text1(4).Text, "T") & " and ctatransfercli = 1"
        If TotalRegistros(SQL) <> 0 Then
        ' comprobamos que ya existe un registro marcado, si lo quieren cambiar
            If chkBanco(3).Value = 1 Then
                If MsgBox("Ya existe otro registro marcado como Cuenta de Transferencia Clientes. " & vbCrLf & " � Desea que sea �sta ? " & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    SQL = "update bancos set ctatransfercli = 0 where codmacta <> " & DBSet(Text1(4).Text, "T")
                    Conn.Execute SQL
                Else
                    ' no hacemos nada
                    chkBanco(3).Value = 0
                End If
            End If
        End If
        B = True
    End If
    
    
    
    DatosOK = B
End Function

'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Sub SugerirCodigoSiguiente()
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonAnyadir
        Case 2
            If BLOQUEADesdeFormulario(Me) Then BotonModificar
        Case 3
            BotonEliminar
        Case 5
            BotonBuscar
        Case 6
            BotonVerTodos
        Case 8
            frmBancoList.Show vbModal
        Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerCtasIVA()
On Error GoTo EPonerCtasIVA

    Text1_LostFocus 4
    Text1_LostFocus 5
    Text1_LostFocus 8
    Text1_LostFocus 10
    Text1_LostFocus 12
    Text1_LostFocus 13
Exit Sub
EPonerCtasIVA:
    MuestraError Err.Number, "Poniendo valores ctas.", Err.Description
End Sub



Private Sub PonerFoco(ByRef Text As TextBox)
    On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SePuedeEliminar() As Boolean
Dim B As Boolean
Dim cad As String

    Screen.MousePointer = vbHourglass
    Set miRsAux = New ADODB.Recordset
    
    SePuedeEliminar = False
    
    'Veamos cobros asociados
    cad = "Select count(*) from cobros where (cuentaba = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Banco con cobros pendientes", vbExclamation
        Exit Function
    End If
    
    
    
    cad = "Select count(*) from pagos where (ctabanc1 = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Banco con pagos pendientes", vbExclamation
        Exit Function
    End If
    
    'Remesas
    cad = "Select count(*) from remesas where (codmacta = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Remesas asociadas.", vbExclamation
        Exit Function
    End If
    
    
    cad = "Select count(*) from gastosfijos where (ctaprevista = '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Gasto fijo asociado.", vbExclamation
        Exit Function
    End If
    
    
    
    cad = "Select count(*) from transferencias where (codmacta= '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Transferencia pagos asociada.", vbExclamation
        Exit Function
    End If
    
    'cOMPROBAMOS ai tiene moovimientos en
    'la NORMA 43
    cad = "Select count(*) from norma43 where (codmacta= '" & Data1.Recordset.Fields(0) & "')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        MsgBox "Asociada a norma 43 en la contabilidad.", vbExclamation
        Exit Function
    End If
    
    SePuedeEliminar = True
    Screen.MousePointer = vbDefault
End Function

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim RS As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        Toolbar1.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0)
        
        Toolbar1.Buttons(5).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
