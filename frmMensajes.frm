VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16440
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMensajes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   16440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameElimiaRiesgoTalPag 
      Height          =   6855
      Left            =   1440
      TabIndex        =   218
      Top             =   1560
      Visible         =   0   'False
      Width           =   12375
      Begin VB.ComboBox cboRiesgoYelConfirming 
         Height          =   360
         ItemData        =   "frmMensajes.frx":000C
         Left            =   7320
         List            =   "frmMensajes.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   241
         Top             =   360
         Width           =   4695
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Agrupar apunte"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   221
         Top             =   6060
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   2
         Left            =   2040
         TabIndex        =   220
         Text            =   "Text2"
         Top             =   6000
         Width           =   1635
      End
      Begin VB.CommandButton cmdElimRiesgoTalPag 
         Caption         =   "&Contabilizar"
         Height          =   375
         Index           =   1
         Left            =   8760
         TabIndex        =   222
         Top             =   6000
         Width           =   1695
      End
      Begin VB.CommandButton cmdElimRiesgoTalPag 
         Caption         =   "Salir"
         Height          =   375
         Index           =   0
         Left            =   10560
         TabIndex        =   223
         Top             =   6000
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView13 
         Height          =   4905
         Left            =   120
         TabIndex        =   219
         Top             =   840
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1773
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Anyo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2805
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Banco"
            Object.Width           =   4419
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripcion"
            Object.Width           =   5263
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tipo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Importe"
            Object.Width           =   2723
         EndProperty
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha apunte"
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   225
         Top             =   6060
         Width           =   1365
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1680
         Picture         =   "frmMensajes.frx":0010
         Top             =   6060
         Width           =   240
      End
      Begin VB.Label Label52 
         Caption         =   " x"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   224
         Top             =   360
         Width           =   6825
      End
   End
   Begin VB.Frame frameCtasBalance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   60
      TabIndex        =   39
      Top             =   60
      Visible         =   0   'False
      Width           =   8205
      Begin VB.CheckBox chkResta 
         Caption         =   "Se resta "
         Height          =   255
         Left            =   1650
         TabIndex        =   71
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdCtaBalan 
         Caption         =   "&Cancelar"
         Height          =   435
         Index           =   1
         Left            =   6450
         TabIndex        =   49
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdCtaBalan 
         Caption         =   "Command4"
         Height          =   435
         Index           =   0
         Left            =   5010
         TabIndex        =   48
         Top             =   2640
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Haber"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   47
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Debe"
         Height          =   255
         Index           =   1
         Left            =   5340
         TabIndex        =   46
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   0
         Left            =   3900
         TabIndex        =   45
         Top             =   1920
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.TextBox Text3 
         Height          =   360
         Left            =   1650
         TabIndex        =   43
         Text            =   "Text3"
         Top             =   1860
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1650
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   900
         Width           =   6045
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1290
         Picture         =   "frmMensajes.frx":009B
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label21 
         Caption         =   "Cuenta"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   44
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   420
         TabIndex        =   41
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "MODIFICAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   420
         TabIndex        =   40
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame FrameAmentoDigito 
      Height          =   7575
      Left            =   240
      TabIndex        =   233
      Top             =   240
      Visible         =   0   'False
      Width           =   14055
      Begin VB.ComboBox cboAumento 
         Height          =   360
         Left            =   12120
         Style           =   2  'Dropdown List
         TabIndex        =   240
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdAumentoDigi 
         Caption         =   "&Aceptar"
         Height          =   495
         Index           =   0
         Left            =   10200
         TabIndex        =   237
         Top             =   6840
         Width           =   1695
      End
      Begin VB.CommandButton cmdAumentoDigi 
         Caption         =   "&Cancelar"
         Height          =   495
         Index           =   1
         Left            =   12120
         TabIndex        =   236
         Top             =   6840
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5625
         Index           =   2
         Left            =   360
         TabIndex        =   238
         Top             =   1080
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   9922
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2990
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Resultante"
            Object.Width           =   3572
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   16349
         EndProperty
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Compruebe copia seguridad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   239
         Top             =   6960
         Width           =   4905
      End
      Begin VB.Label Label7 
         Caption         =   "Aumento dígitos contables"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   8
         Left            =   360
         TabIndex        =   234
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejemplo resultante incremento digitos contables"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   235
         Top             =   480
         Width           =   7695
      End
   End
   Begin VB.Frame FrameValidaIBAN 
      Height          =   4455
      Left            =   4080
      TabIndex        =   208
      Top             =   2040
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton cmdValidarIBAN 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   207
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton optIBAN_Intro 
         Caption         =   "Completo"
         Height          =   240
         Index           =   1
         Left            =   5280
         TabIndex        =   199
         Top             =   1020
         Width           =   1575
      End
      Begin VB.OptionButton optIBAN_Intro 
         Caption         =   "Separado"
         Height          =   240
         Index           =   0
         Left            =   3600
         TabIndex        =   198
         Top             =   1020
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame FrameIBANEspaña 
         Height          =   2175
         Left            =   240
         TabIndex        =   210
         Top             =   1440
         Width           =   6615
         Begin VB.TextBox txtIbanES 
            Height          =   375
            Index           =   5
            Left            =   480
            TabIndex        =   205
            Text            =   "Text4"
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtIbanES 
            Height          =   375
            Index           =   4
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   204
            Text            =   "Text4"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtIbanES 
            Height          =   375
            Index           =   3
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   203
            Text            =   "Text4"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtIbanES 
            Height          =   375
            Index           =   2
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   202
            Text            =   "Text4"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtIbanES 
            Height          =   375
            Index           =   1
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   201
            Text            =   "9999"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtIbanES 
            Height          =   375
            Index           =   0
            Left            =   480
            MaxLength       =   4
            TabIndex        =   200
            Text            =   "Text4"
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "IBAN"
            Height          =   240
            Index           =   5
            Left            =   480
            TabIndex        =   217
            Top             =   1200
            Width           =   750
         End
         Begin VB.Label Label13 
            Caption         =   "Cuenta"
            Height          =   240
            Index           =   4
            Left            =   4320
            TabIndex        =   216
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label13 
            Caption         =   "CC"
            Height          =   240
            Index           =   3
            Left            =   3600
            TabIndex        =   215
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label13 
            Caption         =   "Oficina"
            Height          =   240
            Index           =   2
            Left            =   2640
            TabIndex        =   214
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label13 
            Caption         =   "Entidad"
            Height          =   240
            Index           =   1
            Left            =   1560
            TabIndex        =   213
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label13 
            Caption         =   "IBAN"
            Height          =   240
            Index           =   0
            Left            =   480
            TabIndex        =   212
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.ComboBox cboPais 
         Height          =   360
         ItemData        =   "frmMensajes.frx":0A9D
         Left            =   1080
         List            =   "frmMensajes.frx":0AA4
         Style           =   2  'Dropdown List
         TabIndex        =   197
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdValidarIBAN 
         Caption         =   "Validar"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   206
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "País"
         Height          =   255
         Left            =   360
         TabIndex        =   211
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Verificar IBAN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   2280
         TabIndex        =   209
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame FrameIntegAsiento 
      Height          =   3615
      Left            =   1560
      TabIndex        =   180
      Top             =   240
      Visible         =   0   'False
      Width           =   10695
      Begin VB.TextBox txtDiarioNom 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   3840
         TabIndex        =   187
         Text            =   "Text2"
         Top             =   1200
         Width           =   6075
      End
      Begin VB.TextBox txtDiario 
         Height          =   360
         Index           =   0
         Left            =   2760
         TabIndex        =   176
         Text            =   "Text2"
         Top             =   1200
         Width           =   795
      End
      Begin VB.TextBox txtConceptoNom 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   3840
         TabIndex        =   185
         Text            =   "Text2"
         Top             =   2160
         Width           =   6075
      End
      Begin VB.TextBox txtConcepto 
         Height          =   360
         Index           =   1
         Left            =   2760
         TabIndex        =   178
         Text            =   "Text2"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox txtConcepto 
         Height          =   360
         Index           =   0
         Left            =   2760
         TabIndex        =   177
         Text            =   "Text2"
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtConceptoNom 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   3840
         TabIndex        =   182
         Text            =   "Text2"
         Top             =   1680
         Width           =   6075
      End
      Begin VB.CommandButton cmdIntegAsiento 
         Caption         =   "&Integrar"
         Height          =   495
         Index           =   0
         Left            =   6960
         TabIndex        =   179
         Top             =   2640
         Width           =   1365
      End
      Begin VB.CommandButton cmdIntegAsiento 
         Caption         =   "Salir"
         Height          =   495
         Index           =   1
         Left            =   8520
         TabIndex        =   181
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
         Height          =   240
         Index           =   4
         Left            =   480
         TabIndex        =   188
         Top             =   1200
         Width           =   540
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   2280
         Picture         =   "frmMensajes.frx":0AB0
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Concepto haber"
         Height          =   240
         Index           =   3
         Left            =   480
         TabIndex        =   186
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   1
         Left            =   2280
         Picture         =   "frmMensajes.frx":14B2
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Importación externa de apuntes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   5
         Left            =   480
         TabIndex        =   184
         Top             =   360
         Width           =   5295
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   2280
         Picture         =   "frmMensajes.frx":1EB4
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Concepto debe"
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   183
         Top             =   1680
         Width           =   1500
      End
   End
   Begin VB.Frame FrameIntegrLiqNav 
      Height          =   4215
      Left            =   3480
      TabIndex        =   166
      Top             =   2640
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CheckBox chkPagos 
         Caption         =   "Generar pagos"
         Height          =   255
         Left            =   2280
         TabIndex        =   175
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdIntegFraNav 
         Caption         =   "&Integrar"
         Height          =   495
         Index           =   1
         Left            =   5160
         TabIndex        =   169
         Top             =   3240
         Width           =   1365
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   167
         Text            =   "Text2"
         Top             =   1440
         Width           =   1635
      End
      Begin VB.TextBox txtNommacta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   2160
         TabIndex        =   171
         Text            =   "Text2"
         Top             =   2520
         Width           =   6075
      End
      Begin VB.TextBox txtCodmacta 
         Height          =   360
         Index           =   0
         Left            =   480
         TabIndex        =   168
         Text            =   "Text2"
         Top             =   2520
         Width           =   1635
      End
      Begin VB.CommandButton cmdIntegFraNav 
         Caption         =   "Salir"
         Height          =   495
         Index           =   0
         Left            =   6840
         TabIndex        =   170
         Top             =   3240
         Width           =   1365
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1800
         Picture         =   "frmMensajes.frx":28B6
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "F. recepción"
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   174
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label21 
         Caption         =   "Cta retención"
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   173
         Top             =   2250
         Width           =   1350
      End
      Begin VB.Label Label7 
         Caption         =   "Datos integración liquidaciones NAVARRES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   172
         Top             =   360
         Width           =   6720
      End
      Begin VB.Image imgCtas 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmMensajes.frx":2941
         Top             =   2280
         Width           =   240
      End
   End
   Begin VB.Timer tCuadre 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6420
      Top             =   5400
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameeMPRESAS 
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
      Height          =   5415
      Left            =   11490
      TabIndex        =   22
      Top             =   60
      Width           =   5535
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   26
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Regresar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   25
         Top             =   4800
         Width           =   975
      End
      Begin MSComctlLib.ListView lwE 
         Height          =   3615
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "dsdsd"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame frameCalculoSaldos 
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
      Height          =   5535
      Left            =   9090
      TabIndex        =   12
      Top             =   0
      Width           =   6975
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":3343
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":8B35
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":9547
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4800
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   16
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Iniciar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   4800
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Cálculo de saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame frameSaltos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   52
      Top             =   60
      Width           =   9045
      Begin VB.CommandButton cmdCabError 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   7770
         TabIndex        =   59
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdCabError 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   6690
         TabIndex        =   58
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   3255
         Left            =   4440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Text            =   "frmMensajes.frx":9999
         Top             =   900
         Width           =   4365
      End
      Begin VB.TextBox Text5 
         Height          =   3255
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Text            =   "frmMensajes.frx":999F
         Top             =   900
         Width           =   4125
      End
      Begin VB.Label Label22 
         Caption         =   "Salto"
         Height          =   195
         Index           =   10
         Left            =   4440
         TabIndex        =   57
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Repetidos"
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   56
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label24 
         Caption         =   "Asientos Erróneos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   180
         TabIndex        =   53
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame FrameReclamaciones 
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
      Height          =   5535
      Left            =   0
      TabIndex        =   142
      Top             =   0
      Visible         =   0   'False
      Width           =   15915
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   3
         Left            =   8400
         TabIndex        =   143
         Top             =   4800
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":99A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":F197
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":FBA9
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView9 
         Height          =   3735
         Left            =   180
         TabIndex        =   144
         Top             =   840
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label30 
         Caption         =   "Label30"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   210
         TabIndex        =   145
         Top             =   240
         Width           =   9195
      End
   End
   Begin VB.Frame FrameRecibos 
      Height          =   6720
      Left            =   0
      TabIndex        =   151
      Top             =   0
      Width           =   8670
      Begin VB.CommandButton CmdAcepRecibos 
         Caption         =   "Continuar"
         Height          =   375
         Left            =   5160
         TabIndex        =   153
         Top             =   6060
         Width           =   1455
      End
      Begin VB.CommandButton CmdCanRecibos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6750
         TabIndex        =   152
         Top             =   6060
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   4905
         Left            =   225
         TabIndex        =   154
         Top             =   1005
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label32 
         Caption         =   "Recibos con cobros parciales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   155
         Top             =   390
         Width           =   8025
      End
   End
   Begin VB.Frame FrameDescuadre 
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
      Height          =   5535
      Left            =   0
      TabIndex        =   138
      Top             =   30
      Width           =   8865
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   2
         Left            =   7620
         TabIndex        =   140
         Top             =   4800
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":FFFB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":157ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":161FF
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   120
         TabIndex        =   139
         Top             =   4800
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   3735
         Left            =   120
         TabIndex        =   141
         Top             =   840
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2999
         EndProperty
      End
   End
   Begin VB.Frame FrameIconosVisibles 
      Height          =   6720
      Left            =   30
      TabIndex        =   110
      Top             =   0
      Width           =   7050
      Begin VB.CommandButton cmdAcepIconos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   113
         Top             =   6060
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanIconos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   112
         Top             =   6060
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5235
         Left            =   225
         TabIndex        =   111
         Top             =   675
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   9234
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label44 
         Caption         =   "Accesos Directos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   114
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   6480
         Picture         =   "frmMensajes.frx":16651
         Top             =   330
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   6120
         Picture         =   "frmMensajes.frx":1679B
         Top             =   330
         Width           =   240
      End
   End
   Begin VB.Frame FrameInformeBBDD 
      Height          =   6720
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   116
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4905
         Left            =   225
         TabIndex        =   117
         Top             =   1005
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
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
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8370
         TabIndex        =   124
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9240
         TabIndex        =   123
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   122
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5910
         TabIndex        =   121
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejercicio Siguiente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   120
         Top             =   300
         Width           =   3435
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejercicio Actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3210
         TabIndex        =   119
         Top             =   300
         Width           =   3705
      End
      Begin VB.Label Label46 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   270
         TabIndex        =   118
         Top             =   660
         Width           =   2355
      End
   End
   Begin VB.Frame FrameImpPunteo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   120
      TabIndex        =   72
      Top             =   60
      Width           =   6495
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   8
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text9"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   7
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text9"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   6
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "Text9"
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   5
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text9"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   4
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text9"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   3
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Text9"
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   2
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text9"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   1
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Text9"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text9"
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CommandButton cmdPunteo 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4950
         TabIndex        =   73
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
         X1              =   120
         X2              =   6180
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label22 
         Caption         =   "Haber"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   85
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Debe"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   84
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   240
         Index           =   13
         Left            =   5490
         TabIndex        =   77
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Sin puntear"
         Height          =   240
         Index           =   12
         Left            =   2730
         TabIndex        =   76
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label22 
         Caption         =   "Punteada"
         Height          =   195
         Index           =   11
         Left            =   810
         TabIndex        =   75
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label37 
         Caption         =   "Importes punteo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame frameSaldosHco 
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
      Height          =   4875
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.Frame Frame1 
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
         Height          =   495
         Left            =   120
         TabIndex        =   67
         Top             =   3330
         Width           =   5625
         Begin VB.TextBox txtsaldo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   9
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtsaldo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   8
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   68
            Text            =   "Text1"
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label28 
            Caption         =   "SALDO PERIODO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   70
            Top             =   120
            Width           =   1755
         End
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   7
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   2970
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   6
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   2970
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   5
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   4
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   3
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   2
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   1
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   0
         Left            =   4590
         TabIndex        =   1
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Image Image6 
         Height          =   240
         Index           =   0
         Left            =   5820
         Picture         =   "frmMensajes.frx":168E5
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image Image6 
         Height          =   240
         Index           =   1
         Left            =   5820
         Picture         =   "frmMensajes.frx":172E7
         Top             =   2070
         Width           =   240
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5760
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label28 
         Caption         =   "SALDO"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   66
         Top             =   2970
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "TOTALES"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "PENDIENTE"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PUNTEADA"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "HABER"
         Height          =   255
         Left            =   4380
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "DEBE"
         Height          =   255
         Left            =   2580
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Saldos histórico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   150
         TabIndex        =   2
         Top             =   180
         Width           =   5295
      End
   End
   Begin VB.Frame FrameErrorRestore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   107
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   108
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7435
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label29 
         Caption         =   "Cambio caracteres recupera backup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Width           =   4935
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4920
         Picture         =   "frmMensajes.frx":17CE9
         ToolTipText     =   "Quitar seleccion"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4920
         Picture         =   "frmMensajes.frx":17E33
         ToolTipText     =   "Todos"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameBloqueoEmpresas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   60
      TabIndex        =   95
      Top             =   30
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   106
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   105
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   104
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   101
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   9840
         TabIndex        =   99
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   0
         Left            =   210
         TabIndex        =   97
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   96
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   1
         Left            =   6240
         TabIndex        =   98
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloqueadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   10050
         TabIndex        =   103
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label41 
         Caption         =   "Permitidas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   102
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bloqueo de empresas por usuario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   495
         Index           =   2
         Left            =   2880
         TabIndex        =   100
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameTalPagPdtes 
      Height          =   6720
      Left            =   30
      TabIndex        =   156
      Top             =   30
      Visible         =   0   'False
      Width           =   14010
      Begin VB.TextBox txtSuma 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEF7E4&
         Enabled         =   0   'False
         Height          =   360
         Left            =   7950
         TabIndex        =   163
         Top             =   6030
         Width           =   1575
      End
      Begin VB.CommandButton CmdCancelTalPag 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   12300
         TabIndex        =   158
         Top             =   6000
         Width           =   1365
      End
      Begin VB.CommandButton CmdAcepTalPag 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   10710
         TabIndex        =   157
         Top             =   6000
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView12 
         Height          =   4905
         Left            =   225
         TabIndex        =   159
         Top             =   1005
         Width           =   13545
         _ExtentX        =   23892
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   360
         TabIndex        =   161
         Top             =   6000
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Dividir Vencimiento"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSuma 
         Caption         =   "Suma"
         Height          =   255
         Left            =   7260
         TabIndex        =   164
         Top             =   6060
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "boton de dividir vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   810
         TabIndex        =   162
         Top             =   6030
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   6
         Left            =   13440
         Picture         =   "frmMensajes.frx":17F7D
         ToolTipText     =   "Puntear al Debe"
         Top             =   690
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   7
         Left            =   13080
         Picture         =   "frmMensajes.frx":180C7
         ToolTipText     =   "Quitar al Debe"
         Top             =   690
         Width           =   240
      End
      Begin VB.Label Label33 
         Caption         =   "Talones / Pagarés Pendientes "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   160
         Top             =   390
         Width           =   12465
      End
   End
   Begin VB.Frame FrameEliminarRiesgotalPag 
      Height          =   2895
      Left            =   7320
      TabIndex        =   193
      Top             =   2640
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   1
         Left            =   2040
         TabIndex        =   189
         Text            =   "Text2"
         Top             =   1140
         Width           =   1635
      End
      Begin VB.CommandButton cmdEliminarRiesgoT 
         Caption         =   "&Aceptar"
         Height          =   495
         Index           =   0
         Left            =   2880
         TabIndex        =   191
         Top             =   1920
         Width           =   1365
      End
      Begin VB.CommandButton cmdEliminarRiesgoT 
         Caption         =   "Cancelar"
         Height          =   495
         Index           =   1
         Left            =   4560
         TabIndex        =   192
         Top             =   1920
         Width           =   1365
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Agrupar apunte"
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   190
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1680
         Picture         =   "frmMensajes.frx":18211
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha apunte"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   195
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Eliminar riesgo talón / pagaré"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   6
         Left            =   600
         TabIndex        =   194
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame FrameGeneraTransfer 
      Height          =   3495
      Left            =   8040
      TabIndex        =   226
      Top             =   1200
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkVarios 
         Caption         =   "Agrupar vencimientos"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   228
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.TextBox txtFecha 
         Height          =   360
         Index           =   3
         Left            =   2160
         TabIndex        =   227
         Text            =   "Text2"
         Top             =   1320
         Width           =   1635
      End
      Begin VB.CommandButton cmdGeneraTransfer 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   230
         Top             =   2880
         Width           =   1365
      End
      Begin VB.CommandButton cmdGeneraTransfer 
         Caption         =   "Gener&ar"
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   229
         Top             =   2880
         Width           =   1365
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1800
         Picture         =   "frmMensajes.frx":1829C
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha pago"
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   232
         Top             =   1380
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Transfe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   7
         Left            =   360
         TabIndex        =   231
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame FrameAsientoLiquida 
      Height          =   6645
      Left            =   30
      TabIndex        =   133
      Top             =   60
      Visible         =   0   'False
      Width           =   13230
      Begin VB.CommandButton CmdContabilizar 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   9840
         TabIndex        =   137
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   11520
         TabIndex        =   134
         Top             =   6000
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   4905
         Left            =   120
         TabIndex        =   135
         Top             =   840
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label7 
         Caption         =   "Realiza apunte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   11160
         TabIndex        =   165
         Top             =   480
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label54 
         Caption         =   "Asiento Contable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   136
         Top             =   390
         Width           =   3705
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   6720
      Left            =   0
      TabIndex        =   129
      Top             =   30
      Visible         =   0   'False
      Width           =   13410
      Begin VB.CommandButton cmdGenerarVto 
         Caption         =   "Generar"
         Height          =   375
         Left            =   360
         TabIndex        =   196
         Top             =   6120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   12000
         TabIndex        =   130
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   4905
         Left            =   225
         TabIndex        =   131
         Top             =   1005
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label52 
         Caption         =   "Cobros de la factura "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   132
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame FrameBancosRemesas 
      Height          =   6720
      Left            =   0
      TabIndex        =   146
      Top             =   30
      Visible         =   0   'False
      Width           =   8670
      Begin VB.CommandButton CmdCancelBancoRem 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6750
         TabIndex        =   148
         Top             =   6060
         Width           =   1365
      End
      Begin VB.CommandButton CmdAcepBancoRem 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   147
         Top             =   6060
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView10 
         Height          =   4905
         Left            =   225
         TabIndex        =   149
         Top             =   1005
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   7620
         Picture         =   "frmMensajes.frx":18327
         ToolTipText     =   "Quitar al Debe"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   7980
         Picture         =   "frmMensajes.frx":18471
         ToolTipText     =   "Puntear al Debe"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label31 
         Caption         =   "Importe por Banco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   150
         Top             =   390
         Width           =   8025
      End
   End
   Begin VB.Frame FrameShowProcess 
      Height          =   6720
      Left            =   0
      TabIndex        =   125
      Top             =   30
      Width           =   10950
      Begin VB.CommandButton CmdRegresar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   126
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4905
         Left            =   225
         TabIndex        =   127
         Top             =   1005
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label53 
         Caption         =   "Usuarios conectados a"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   128
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame FrameVerObservacionesCuentas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   30
      TabIndex        =   89
      Top             =   30
      Width           =   9375
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   5
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   94
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":185BB
         Top             =   720
         Width           =   7665
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   4
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   93
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":185C1
         Top             =   720
         Width           =   1005
      End
      Begin VB.CommandButton cmdVerObservaciones 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   7920
         TabIndex        =   92
         Top             =   5460
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   3915
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   90
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":185C7
         Top             =   1320
         Width           =   8775
      End
      Begin VB.Label Label22 
         Caption         =   "Descripción cuentas Plan General Contable  2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   19
         Left            =   360
         TabIndex        =   91
         Top             =   360
         Width           =   5010
      End
   End
   Begin VB.Frame frameBalance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   90
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   13455
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   480
         MaxLength       =   10
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CheckBox chkPintar 
         Caption         =   "Escribir si el resultado es negativo"
         Height          =   255
         Left            =   2430
         TabIndex        =   51
         Top             =   3630
         Width           =   3915
      End
      Begin VB.CheckBox chkCero 
         Caption         =   "Poner a CERO si el resultado es negativo"
         Height          =   255
         Left            =   6630
         TabIndex        =   50
         Top             =   3630
         Width           =   4575
      End
      Begin VB.CheckBox chkNegrita 
         Caption         =   "Negrita"
         Height          =   255
         Left            =   11460
         TabIndex        =   38
         Top             =   3630
         Width           =   1035
      End
      Begin VB.CommandButton cmdBalance 
         Caption         =   "Cancelar"
         Height          =   435
         Index           =   1
         Left            =   12000
         TabIndex        =   36
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBalance 
         Caption         =   "Aceptar"
         Height          =   435
         Index           =   0
         Left            =   10680
         TabIndex        =   35
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   480
         MaxLength       =   200
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2760
         Width           =   12675
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   480
         MaxLength       =   100
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1980
         Width           =   12705
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   480
         MaxLength       =   100
         TabIndex        =   28
         Text            =   "WWWWWWWWWWFFFFFFFFFFWWWWWWWWWWFFFFFFFFFFWWWWWWWWWWFFFFFFFFFFWWWWWWWWWWFFFFFFFFFF"
         Top             =   1080
         Width           =   12735
      End
      Begin VB.Label Label15 
         Caption         =   "Código oficial balance"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   60
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "MODIFICAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   510
         TabIndex        =   37
         Top             =   300
         Width           =   4875
      End
      Begin VB.Label Label15 
         Caption         =   "Formula"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   34
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Texto cuentas"
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Nombre"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Label Label11 
      Caption         =   "="
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
      Index           =   2
      Left            =   1800
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   2040
      X2              =   3960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label10 
      Caption         =   "años de vida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label10 
      Caption         =   "Valor adquisición"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TABLAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '1.- Saldos historico
    '2.- Comprobar saldos
    '3.- Mostrar tipos de amortizacion
    '4.- Seleccionar empresas
    
    '5.- Es, como si fuera comprobar saldos , pero se lanza y se cierra autmaticamente
    
    '7.- Nueva linea en configuracion balances
    '8.- Modificar linea balances
    
    
    '9.- Nueva CTA de configuracion balances
    '10- MODIFICAR  "   "             "
        
    '12- Asientos con saltos y/o repetidos
    
    
        
    '16- Traspaso de facturas entre PC's. EXP
    '17-   "                  "           IMPORTAR
    
    
     '18- Importes punteo
     '19- Copiar de un balance a OTRO
     '20- Importar fichero datos 347 externo
     '21- Ver OBSERVACIONES cuentas
     '22- Ver empresas bloquedas
     '23- Menu de ayuda
     '24- Iconos de pantalla principal
     '25- Informe de base de datos
     '26- Show processlist
     
     '27- Cobros de la factura
     '28- Pagos de la factura
     
     '29- Asiento de liquidacion
     
     '30- Asientos descuadrados
     
     '***** TESORERIA *****
     '50- Facturas de Reclamaciones
     
     '51- Facturas remesas
     '52- Bancos remesas
     '53- Recibos con cobros parciales
     '54- Cobros pendientes en recepcion de documentos (talones/pagares)
     '55- Facturas Transferencias de abonos
     '56- Facturas Transferencias de pagos
     
     '57- Facturas de compensaciones
     '58- Facturas de compensaciones proveedor
     '59- Integracon liquidaciones NAvarres
     '60- Integracon asientos. Pedir conceptos
        
     '61. Fecha eliminacion riesgo. Pedir agrupar
     '62  Validar IBAN
    
     '63 Eliminar riesgo talon/pagare
    
     '64 Generacion fichero transferencia
     
     '65 Mensaje Aumento digito contable
     
     '66 'Como el 55,56 Pero para confirming  ( Facturas Transferencias de pagos)
     '67 'Como el 55,56 Pero para PAGO DOMICILIADO ( Facturas Transferencias de pagos)
     
Public Parametros As String
    '1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

'recepcion de talon/pagare
Public Importe As Currency
Public Codigo As String
Public Tipo As String
Public FecCobro As String
Public FecVenci As String
Public Banco As String
Public Referencia As String


Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmPag As frmFacturasProPag
Attribute frmPag.VB_VarHelpID = -1


Private PrimeraVez As Boolean

Dim i As Integer
Dim SQL As String
Dim Rs As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim Ok As Integer

Dim CampoOrden As String
Dim Orden As Boolean


Private Sub cboAumento_Click()
    If PrimeraVez Then Exit Sub
    
    CargaResultadoCuentasIncrementar
End Sub

Private Sub cboRiesgoYelConfirming_Click()
    If PrimeraVez Then Exit Sub
    ponerDatosRiesgoTalPag_Confir
End Sub

Private Sub chkVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, True
End Sub

Private Sub CmdAcepBancoRem_Click()
Dim i As Integer

    CadenaDesdeOtroForm = ""

    For i = 1 To ListView10.ListItems.Count
        If ListView10.ListItems(i).Checked Then
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "'" & Trim(ListView10.ListItems(i).Text) & "',"
        End If
    Next i
        
    If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 1)
    
    Unload Me
End Sub

Private Sub cmdAcepIconos_Click()
Dim i As Integer
Dim SQL As String
Dim CadenaIconos As String


Dim K As Integer
Dim J As Integer
Dim Px As Single
Dim Py As Single


Dim H As Integer
Dim MargenX As Single
Dim MargenY As Single
Dim Ocupado As Boolean
    'Ponemos los que ha desmarcado a cero
    CadenaIconos = ""
    For i = 1 To ListView6.ListItems.Count
        If Not ListView6.ListItems(i).Checked Then
            CadenaIconos = CadenaIconos & ListView6.ListItems(i).Text & ","
        End If
    Next i
    If CadenaIconos <> "" Then
        CadenaIconos = Mid(CadenaIconos, 1, Len(CadenaIconos) - 1)
        SQL = "update menus_usuarios set posx = 0, posy = 0, vericono = 0 where aplicacion = 'ariconta' and codusu = " & vUsu.Id & " and codigo in (" & CadenaIconos & ")"
        Conn.Execute SQL
    End If

    

    For i = 1 To ListView6.ListItems.Count
        If ListView6.ListItems(i).Checked Then
            'SI NO ERA VISIBLE le busco el hueco
            If ListView6.ListItems(i).SubItems(2) = "0" Then
                SQL = ""
                For J = 1 To 8
                    For K = 1 To 5
                        DevuelCoordenadasCuadricula J, K, Px, Py
                        Ocupado = False
                        'Busco hueco
                        For H = 1 To ListView6.ListItems.Count
                            If ListView6.ListItems(H).SubItems(2) = "1" Then
                                MargenX = Abs(Px - CSng(ListView6.ListItems(H).SubItems(3)))
                                MargenY = Abs(Py - CSng(ListView6.ListItems(H).SubItems(4)))
                                
                                If MargenX < 300 And MargenY < 300 Then
                                    'HUECO
                                    Ocupado = True
                                    Exit For
                                End If
                            End If
                        Next H
                        
                        If Not Ocupado Then
                            'OK. Este es. Lo ponemos a true y actualizamos BD
                            ListView6.ListItems(i).SubItems(2) = "1"
                            ListView6.ListItems(i).SubItems(3) = Px
                            ListView6.ListItems(i).SubItems(4) = Py
                            SQL = "update menus_usuarios set posx = " & DBSet(Px, "N")
                            SQL = SQL & ", posy = " & DBSet(Py, "N") & ", vericono = 1 where "
                            SQL = SQL & "aplicacion = 'ariconta' and codusu = " & vUsu.Id
                            SQL = SQL & " and codigo =" & DBSet(ListView6.ListItems(i).Text, "T")
                            Conn.Execute SQL
                           Exit For
                        End If
                    Next K
                    If SQL <> "" Then Exit For
                Next J
            End If
           
        End If
    Next i
    Reorganizar = True
    
    Unload Me
End Sub





Private Sub CmdAcepRecibos_Click()
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub CmdAcepTalPag_Click()

    If Not DatosOK Then Exit Sub

    
    InsertarModificarTalones
    
    Unload Me
End Sub

Private Sub InsertarModificarTalones()
Dim Importe As Currency
Dim importe2 As Currency
Dim TipForpa As String
Dim J As Integer
   
    On Error GoTo EInsertarModificar
    
    CadenaDesdeOtroForm = ""
    
    Conn.BeginTrans
    
    For i = 1 To Me.ListView12.ListItems.Count
        If ListView12.ListItems(i).Checked Then
        
            SQL = "select * from talones_facturas where codigo = " & DBSet(Codigo, "N") & " and " & MontaWhere(i, False)
        
            If TotalRegistrosConsulta(SQL) = 0 Then
                
                SQL = "insert into `talones_facturas` (`codigo`,`numserie`,`numfactu`,`fecfactu`, "
                SQL = SQL & "`numorden`,`importe`,`contabilizado`) VALUES ("
                SQL = SQL & DBSet(Codigo, "N") & "," & DBSet(ListView12.ListItems(i).Text, "T") & "," & DBSet(ListView12.ListItems(i).SubItems(1), "N") & ","
                SQL = SQL & DBSet(ListView12.ListItems(i).SubItems(2), "F") & "," & DBSet(ListView12.ListItems(i).SubItems(4), "N") & ","
                SQL = SQL & DBSet(ListView12.ListItems(i).SubItems(9), "N") & ",0) "

                Conn.Execute SQL
                
                ' ahora en cobros
                ' sacamos el tipforpa
                SQL = MontaWhere(i, False)
                
                SQL = " WHERE cobros.codforpa=formapago.codforpa and " & SQL
                SQL = "select tipforpa from cobros,formapago " & SQL
                
                TipForpa = DevuelveValor(SQL)
                
                'actualizamos cobros
                SQL = "UPDATE cobros SET recedocu=1 "
                SQL = SQL & ", impcobro = coalesce(impcobro,0) + " & DBSet(ListView12.ListItems(i).SubItems(9), "N")
                SQL = SQL & ", fecultco = " & DBSet(FecCobro, "F")
                'Fecha vencimiento tb le pongo la de la recpcion
                SQL = SQL & ", fecvenci = " & DBSet(FecVenci, "F")
                'BANCO LO PONGO EN OBSERVACION
                SQL = SQL & ", observa = " & DBSet(Banco, "T")
                'Si no era forma de pago talon/pagare la pongo
                If Tipo = 0 Then
                    J = vbPagare
                ElseIf Tipo = 2 Then
                    J = vbConfirming
                Else
                    J = vbTalon
                End If
                
                If TipForpa <> J Then
                    'AQUI BUSCARE una forma de pago
                    J = Val(DevuelveDesdeBD("codforpa", "formapago", "tipforpa", CStr(J)))
                    If J > 0 Then SQL = SQL & ", codforpa = " & J
                    
                End If
                            
                Conn.Execute SQL & MontaWhere(i, True)
            Else
                SQL = "UPDATE cobros SET recedocu=1 "
                Importe = 0
                If Trim(ListView12.ListItems(i).SubItems(9)) <> "" Then Importe = ImporteFormateado(ListView12.ListItems(i).SubItems(9))
                Importe = DevuelveValor("select importe from talones_facturas where codigo = " & DBSet(Codigo, "N") & " and " & MontaWhere(i, False)) + Importe

                
                SQL = SQL & ", impcobro =  " & DBSet(Importe, "N")
                SQL = SQL & ", fecultco = " & DBSet(FecCobro, "F")
                'Fecha vencimiento tb le pongo la de la recpcion
                SQL = SQL & ", fecvenci = " & DBSet(FecVenci, "F")
                'BANCO LO PONGO EN OBSERVACION
                SQL = SQL & ", observa = " & DBSet(Banco, "T")
                
                'If TipForpa <> J Then
                '    'AQUI BUSCARE una forma de pago
                '    J = Val(DevuelveDesdeBD("codforpa", "formapago", "tipforpa", CStr(J)))
                '    If J > 0 Then SQL = SQL & ", codforpa = " & J
                '
                'End If
                            
                Conn.Execute SQL & MontaWhere(i, True)
            
            End If
        
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "(" & DBSet(ListView12.ListItems(i).Text, "T") & "," & DBSet(ListView12.ListItems(i).SubItems(1), "N") & "," & DBSet(ListView12.ListItems(i).SubItems(2), "F") & "," & DBSet(ListView12.ListItems(i).SubItems(4), "N") & "),"
            
        Else
            'Obtengo el importe del vto
            Importe = DevuelveValor("select impcobro from cobros " & MontaWhere(i, True))
            importe2 = DevuelveValor("select importe from talones_facturas where codigo = " & DBSet(Codigo, "N") & " and " & MontaWhere(i, False))
            
            If Importe <> importe2 Then
                'TODO EL IMPORTE estaba en la linea. Fecultco a NULL
                J = 1
                Importe = Importe - importe2
            Else
                J = 0
            End If
            
            SQL = "Delete from talones_facturas"
            SQL = SQL & " WHERE codigo =" & DBSet(Codigo, "N")
            SQL = SQL & " AND " & MontaWhere(i, False)
            Conn.Execute SQL
            
            'Updateo en cobros reestableciendo los valores
            
            SQL = "UPDATE cobros SET recedocu=0"
            If J = 0 Then
                SQL = SQL & ", impcobro = NULL, fecultco = NULL"
            Else
                SQL = SQL & ", impcobro = " & TransformaComasPuntos(CStr(Importe))  'NO somos capace sde ver cual fue la utlima fecha de amortizacion
            End If
            SQL = SQL & ", observa = NULL"
            SQL = SQL & " WHERE " & MontaWhere(i, False)
            
            Conn.Execute SQL
        End If
        
        'ponemos la situacion
        SQL = "update cobros set situacion = if(impvenci + coalesce(gastos,0) - coalesce(impcobro,0) = 0,1,0) "
        SQL = SQL & " where " & MontaWhere(i, False)
        Conn.Execute SQL
        
        
    Next i
    Conn.CommitTrans
    Exit Sub
    
EInsertarModificar:
    Conn.RollbackTrans
    MuestraError Err.Number, "Insertar Modificar Talones/Pagarés", Err.Description
End Sub



Private Function MontaWhere(i As Integer, conWhere As Boolean) As String
Dim SQL As String
    
    If conWhere Then
        SQL = " WHERE "
    Else
        SQL = " "
    End If
    SQL = SQL & " numserie = " & DBSet(ListView12.ListItems(i).Text, "T")
    SQL = SQL & " and numfactu = " & DBSet(ListView12.ListItems(i).SubItems(1), "N") & " and fecfactu = " & DBSet(ListView12.ListItems(i).SubItems(2), "F")
    SQL = SQL & " and numorden = " & DBSet(ListView12.ListItems(i).SubItems(4), "N")
    
    MontaWhere = SQL

End Function


Private Function DatosOK() As Boolean
Dim B As Boolean

    DatosOK = False
    
    If CCur(ComprobarCero(txtSuma)) > Importe Then
        If MsgBox("El importe del documento es inferior a la suma de las facturas seleccionadas. " & vbCrLf & "Deberia dividir vencimiento." & vbCrLf & vbCrLf & " ¿ Continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            B = False
        Else
            B = True
        End If
    Else
        If CCur(ComprobarCero(txtSuma)) < Importe Then
            If MsgBox("El importe del documento es superior a la suma de las facturas seleccionadas. " & vbCrLf & vbCrLf & " ¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                B = False
            Else
                B = True
            End If
        Else
            B = True
        End If
    End If
    
    DatosOK = B
    
End Function



Private Sub cmdAumentoDigi_Click(Index As Integer)
    
    If Index = 0 Then
        SQL = Label17(1).Caption & vbCrLf & vbCrLf & "¿Continuar?"
        i = CInt(MsgBox(SQL, vbQuestion + vbYesNoCancel))
        If i <> vbYes Then
            If i = vbNo Then Exit Sub
        Else
            SQL = InputBox("Escriba password de seguridad", "CLAVE")
            If UCase(SQL) <> "ARIADNA" Then
                If SQL <> "" Then MsgBox "Clave incorrecta", vbExclamation
                Exit Sub
            End If
            CadenaDesdeOtroForm = "OK"
        End If
    End If
    Unload Me
End Sub

Private Sub cmdBalance_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
        Unload Me
    Else
        If Text1(0).Text = "" Then
            MsgBox "Primer campo obligatorio", vbExclamation
            Exit Sub
        End If
        If InsertarModificar Then Unload Me
    End If
End Sub

Private Sub cmdBlEmp_Click(Index As Integer)

    Select Case Index
    Case 0, 1
        'Index Me dira que listview
        For Ok = ListView2(Index).ListItems.Count To 1 Step -1
            If ListView2(Index).ListItems(Ok).Selected Then
                i = ListView2(Index).ListItems(Ok).Index
                PasarUnaEmpresaBloqueada Index = 0, i
            End If
        Next Ok
    Case Else
        If Index = 2 Then
            Ok = 0
        Else
            Ok = 1
        End If
        For NumRegElim = ListView2(Ok).ListItems.Count To 1 Step -1
            PasarUnaEmpresaBloqueada Ok = 0, ListView2(Ok).ListItems(NumRegElim).Index
        Next NumRegElim
        Ok = 0
    End Select
End Sub



Private Sub PasarUnaEmpresaBloqueada(ABLoquedas As Boolean, Indice As Integer)
Dim Origen As Integer
Dim Destino As Integer
Dim IT
    If ABLoquedas Then
        Origen = 0
        Destino = 1
        NE = 2
    Else
        Origen = 1
        Destino = 0
        NE = 1 'icono
    End If
    
    SQL = ListView2(Origen).ListItems(Indice).Key
    Set IT = ListView2(Destino).ListItems.Add(, SQL)
    IT.SmallIcon = NE
    IT.Text = ListView2(Origen).ListItems(Indice).Text
    IT.SubItems(1) = ListView2(Origen).ListItems(Indice).SubItems(1)

    'Borramos en origen
    ListView2(Origen).ListItems.Remove Indice
End Sub

Private Sub cmdBloqEmpre_Click(Index As Integer)
    If Index = 0 Then
        SQL = "DELETE FROM usuarios.usuarioempresasariconta WHERE codusu =" & Parametros
        Conn.Execute SQL
        SQL = ""
        For i = 1 To ListView2(1).ListItems.Count
            SQL = SQL & ", (" & Parametros & "," & Val(Mid(ListView2(1).ListItems(i).Key, 2)) & ")"
        Next i
        If SQL <> "" Then
            'Quitmos la primera coma
            SQL = Mid(SQL, 2)
            SQL = "INSERT INTO usuarios.usuarioempresasariconta(codusu,codempre) VALUES " & SQL
            If Not EjecutaSQL(SQL) Then MsgBox "Se han producido errores insertando datos", vbExclamation
        End If
    End If
    Unload Me
End Sub

Private Sub cmdCabError_Click(Index As Integer)
Dim Rs As ADODB.Recordset
Dim J As Long
Dim ii As Long
Dim Anyo As Integer

    If Index = 1 Then
        Unload Me
    Else
        Screen.MousePointer = vbHourglass
        Anyo = 0
        i = 0
        Do
          
            SQL = "select numasien,fechaent from hcabapu where fechaent >= '"
            SQL = SQL & Format(DateAdd("yyyy", Anyo, vParam.fechaini), FormatoFecha)
            SQL = SQL & "' AND fechaent <= '" & Format(DateAdd("yyyy", Anyo, vParam.fechafin), FormatoFecha) & "' ORDER By NumAsien"
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            ii = 0
            While Not Rs.EOF
               J = Rs.Fields(0)
               'Igual
                If J - ii = 0 Then
                    
                    SQL = Format(J, "00000")
                    SQL = SQL & "  -  " & Format(Rs!FechaEnt, "dd/mm/yyyy")
                    Text5.Text = Text5.Text & SQL & vbCrLf
                    i = i + 1
                Else
                    If J - ii > 1 Then
                        If J - ii = 2 Then
                            SQL = Format(J - 1, "00000")
                        Else
                            SQL = "Entre " & Format(ii, "00000") & "  y  " & Format(J, "00000")
                        End If
                        SQL = SQL & " (" & CStr(Year(vParam.fechaini) + Anyo) & ")"
                        Text6.Text = Text6.Text & SQL & vbCrLf
                        i = i + 1
                    End If
                End If
                ii = J
                'Refrescamos
                If i > 50 Then
                    Text5.Refresh
                    Text6.Refresh
                    i = 0
                End If
                
                '
                Rs.MoveNext
            Wend
            Rs.Close
            Anyo = Anyo + 1
        Loop Until Anyo > 1
        Me.Refresh
        Screen.MousePointer = vbDefault
        cmdCabError(0).Enabled = False
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub CmdCancelBancoRem_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub CmdCancelTalPag_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdCanIconos_Click()
    Reorganizar = False
    Unload Me
End Sub

Private Sub CmdCanRecibos_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub CmdContabilizar_Click()
    
    If UCase(ListView7.ColumnHeaders(1).Text) = "EMPR" Then
        If MsgBox("Continuar con el proceso?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub cmdCtaBalan_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
    Else
        If Text3.Text = "" Then
            MsgBox "la cuenta no puede estar en blanco", vbExclamation
            Exit Sub
        End If
        If Not IsNumeric(Text3.Text) Then
            MsgBox "La cuenta debe ser numérica", vbExclamation
            Exit Sub
        End If
        'Esto es el OPTION
        SQL = ""
        For i = 0 To 2
            If Option1(i).Value Then SQL = SQL & Mid(Option1(i).Caption, 1, 1)
        Next i
        If SQL = "" Then
            MsgBox "Seleccione una opción de la cuenta (Saldo - Debe - Haber )"
            Exit Sub
        End If
        
        'RESTA y la resta
        SQL = SQL & "|" & Abs(Me.chkResta.Value)
        CadenaDesdeOtroForm = Text3.Text & "|" & SQL & "|"
    End If
    Unload Me
End Sub

Private Sub cmdEliminarRiesgoT_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        Codigo = ""
        If Me.txtfecha(1).Text = "" Then
            Codigo = "Fecha requerida"
        ElseIf Not IsDate(txtfecha(1).Text) Then
            Codigo = "Fecha incorrecta"
        Else
            i = CInt(FechaCorrecta2(CDate(txtfecha(1).Text)))
            If i > 1 Then
                If i = 2 Then
                    Codigo = varTxtFec
                
                Else
                    If i = 3 Then
                        Codigo = "El ejercicio al que pertenece la fecha: " & txtfecha(1).Text & " está cerrado."
                    Else
                        Codigo = "Ejercicio para: " & txtfecha(1).Text & " todavía no activo"
                    End If
                End If
            End If
        End If
        If Codigo <> "" Then
            MsgBox Codigo, vbExclamation
            Me.txtfecha(1).Text = ""
            PonFoco Me.txtfecha(1)
            Exit Sub
        End If
        
        'LLEGA aqui, todo OK. Continuar?
        If MsgBox("Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        CadenaDesdeOtroForm = Me.txtfecha(1).Text & "|" & Me.chkVarios(0).Value & "|"
    End If
    Unload Me
End Sub

Private Sub cmdElimRiesgoTalPag_Click(Index As Integer)
Dim Byt  As Byte
Dim Forpa As Ctipoformapago
Dim Salir As Boolean

    If Index = 1 Then
        Salir = False
        
        'Alguna seleccionada
        Referencia = ""
        SQL = ""
        For i = 1 To ListView13.ListItems.Count
            If ListView13.ListItems(i).Checked Then SQL = SQL & "X"
        Next
        If SQL = "" Then Referencia = "-Seleccione alguna remesa para eliminar riesgo"
        
                
        'Fecha seleccionada
        Codigo = ""
        i = 2
        If Me.txtfecha(2).Text = "" Then
            Codigo = "-Fecha requerida"
        ElseIf Not IsDate(txtfecha(2).Text) Then
            Codigo = "-Fecha incorrecta"
        Else
            i = CInt(FechaCorrecta2(CDate(txtfecha(2).Text)))
            If i > 1 Then
                If i = 2 Then
                    Codigo = "-" & varTxtFec
                
                Else
                    If i = 3 Then
                        Codigo = "-El ejercicio al que pertenece la fecha: " & txtfecha(2).Text & " está cerrado."
                    Else
                        Codigo = "-Ejercicio para: " & txtfecha(2).Text & " todavía no activo"
                    End If
                End If
            Else
                i = 0
            End If
        End If
        If Referencia <> "" Then Codigo = vbCrLf & Referencia & vbCrLf & Codigo
        
        If Codigo <> "" Then
            MsgBox Codigo, vbExclamation
            If i > 0 Then
                Me.txtfecha(2).Text = ""
                PonFoco Me.txtfecha(2)
            End If
            Exit Sub
        End If
        
        
        If MsgBox("Continuar con el proceso?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        
        Screen.MousePointer = vbHourglass
        Set Forpa = New Ctipoformapago
        Byt = 100
        For i = 1 To ListView13.ListItems.Count
            
            If ListView13.ListItems(i).Checked Then
            
                
                J = vbPagare
                If CByte(ListView13.ListItems(i).Tag) = 3 Then J = vbTalon
                If CByte(ListView13.ListItems(i).Tag) = 127 Then J = vbConfirming
            
                If Forpa.Leer(CInt(J)) = 0 Then
    
            
                    Conn.BeginTrans
                 
                 
                     Byt = RemesasEliminarVtosTalonesPagares(CByte(ListView13.ListItems(i).Tag), ListView13.ListItems(i).Text, ListView13.ListItems(i).SubItems(1), CDate(txtfecha(2).Text), Forpa, Me.chkVarios(1).Value)
                 
                    If Byt < 2 Then
                         Conn.CommitTrans
                         
                    Else
                         TirarAtrasTransaccion
                    End If
                    Exit For
                End If
            End If
         Next
         If Byt <> 100 Then
            If Byt < 2 Then
                'Ha ido bien. Volvemos a cargar
                Screen.MousePointer = vbHourglass
                ponerDatosRiesgoTalPag_Confir
                If ListView13.ListItems.Count = 0 Then Salir = True
            End If
        End If
         Screen.MousePointer = vbDefault
    Else
        Salir = True
    End If
    
    
    
    
    
    If Salir Then
        Unload Me
    Else
    
    End If
End Sub

Private Sub cmdEmpresa_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        SQL = ""
        Parametros = ""
        For i = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(i).Checked Then
                SQL = SQL & Me.lwE.ListItems(i).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next i
        CadenaDesdeOtroForm = Len(Parametros) & "|" & SQL
        'Vemos las conta
        SQL = ""
        For i = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(i).Checked Then
                SQL = SQL & Me.lwE.ListItems(i).Tag & "|"
            End If
        Next i
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    End If
    Unload Me
End Sub


Private Sub cmdGenerarVto_Click()
    
    If Opcion <> 28 Then Exit Sub


    
    SQL = "delete from tmppagos where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute SQL
    SQL = ""
    frmFacturasPro.CargarPagosTemporal RecuperaValor(Codigo, 1), RecuperaValor(Codigo, 2), CCur(RecuperaValor(Codigo, 3))
    Set frmPag = frmFacturasProPag
    ContinuarPago = False
    frmPag.CodigoActual = cmdGenerarVto.Tag
                
    frmPag.Show vbModal
    Set frmPag = Nothing
    If SQL <> "" Then
    
    
        frmFacturasPro.EstablecerValoresSeleccionPago SQL
        If ContinuarPago Then
            frmFacturasPro.InsertarPagos ""
            ContinuarPago = False
            CargaPagosFactura
            If frmFacturasPro.HayQueContabilizarDesdePantallaPagos Then frmFacturasPro.ContabilizarPagos
        End If
        
    End If

End Sub

Private Sub cmdGeneraTransfer_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
        
    Else
        SQL = ""
        If txtfecha(3).Text = "" Then
            SQL = "Fecha cobro debe tener valor"
        Else
            If Not EsFechaOK(txtfecha(3)) Then
                SQL = "Fecha incorrecta"
            Else
                If CDate(txtfecha(3).Text) < CDate(Format(Now, "dd/mm/yyyy")) Then SQL = "Fecha debe ser igual o superior a hoy"
            End If
        End If
            
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            PonFoco txtfecha(3)
            Exit Sub
        End If

        CadenaDesdeOtroForm = txtfecha(3).Text & "|" & Me.chkVarios(2).Value & "|"
        
    End If
    Unload Me
End Sub

Private Sub cmdIntegAsiento_Click(Index As Integer)
    
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
    Else
        For i = 0 To 1
            If txtConcepto(i).Text = "" Or txtConceptoNom(i).Text = "" Then
                MsgBox "Concepto tienen que tener un valor correcto", vbExclamation
                PonFoco txtConcepto(i)
                Exit Sub
            End If
            
            If txtDiario(0).Text = "" Then
                MsgBox "Diario debe tener valor", vbExclamation
                PonFoco txtDiario(0)
                Exit Sub
            End If
            
        Next
    
        'Si llega aqui preguntamos
        If MsgBox("Desea continuar con la importacion del fichero?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        CadenaDesdeOtroForm = txtConcepto(0).Text & "|" & txtConcepto(1).Text & "|" & txtDiario(0).Text & "|"
    End If
    Unload Me
End Sub

Private Sub cmdIntegFraNav_Click(Index As Integer)
    If Index = 1 Then
        If Me.txtCodmacta(0).Text = "" Or Me.txtfecha(0).Text = "" Then
            MsgBox "Campos requeridos", vbExclamation
            Exit Sub
        End If
        
          
          
        i = CInt(FechaCorrecta2(CDate(txtfecha(0).Text), True))
        If i > 2 Then Exit Sub 'fuera ejercicios
        If i = 2 Then
            If vUsu.Nivel > 0 Then Exit Sub
        End If
        
        'desea continuar.....
        If MsgBox("Desea continuar con los datos introducidos?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        
        CadenaDesdeOtroForm = Me.txtCodmacta(0).Text & "|" & Me.txtfecha(0).Text & "|" & Abs(Me.chkPagos.Value) & "|"
        
    End If
    Unload Me
End Sub

Private Sub cmdPunteo_Click()
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub

Private Sub cmdValidarIBAN_Click(Index As Integer)
    
    If Index = 1 Then
        Unload Me
        
    Else
        Errores = ""
        If Me.optIBAN_Intro(0).Value Then
            'Esta introducionedolo separado
            ValidarIBANEspanaSeparado
        Else
            ValidarIBANEspanaJunto
        End If
        If Errores <> "" Then MsgBox Errores, vbExclamation
            
    End If
End Sub

Private Sub cmdVerObservaciones_Click()
    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Command2_Click()
Dim Digitos As Integer
    ListView1.ListItems.Clear
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = vEmpresa.numnivel + 1
    Me.ProgressBar1.visible = True
    Screen.MousePointer = vbHourglass
    Me.ProgressBar1.visible = False
    Screen.MousePointer = vbDefault
End Sub




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        
        

    
        
        
        Select Case Opcion
        Case 5
            Screen.MousePointer = vbHourglass
            Me.tCuadre.Enabled = True
        Case 21
            cargarObservacionesCuenta
        Case 22
            cargaempresasbloquedas
            
        Case 24
            'CargaIconosVisibles
            
        Case 25
            'CargaInformeBBDD
        
        Case 26
            'CargaShowProcessList
        
        Case 27
            CargaCobrosFactura
        Case 28
            CargaPagosFactura
            
        Case 29
            CargarAsiento
        Case 30
            CargarAsientosDescuadrados
        Case 31
            CargarFacturasSinAsientos
        Case 50
            CargarFacturasReclamaciones
        Case 51
            CargarFacturasRemesas
        Case 52
            CargarBancosRemesas
        Case 53
            CargarRecibosConCobrosParciales
        Case 54
            CargarTalonesPagaresPendientes
        Case 55
            CargarFacturasTransfAbonos
        Case 56
            CargarFacturasTransfPagos 0
        Case 57
            CargarFacturasCompensaciones
        Case 58
            CargarFacturasCompensacionesPro
        Case 62
            optIBAN_Intro_Click 0
        Case 63
            ponerDatosRiesgoTalPag_Confir
            
        Case 64
            ValoresTransferencia
        Case 65
            CargaResultadoCuentasIncrementar
        Case 66, 67
            CargarFacturasTransfPagos 0
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And Opcion = 23 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim W As Integer, H As Integer
    
    
    Me.Icon = frmppal.Icon
    Me.tCuadre.Enabled = False
    PrimeraVez = True
    Me.frameSaldosHco.visible = False
    Me.frameCalculoSaldos.visible = False
    Me.FrameeMPRESAS.visible = False
    Me.frameCtasBalance.visible = False
    Me.frameSaltos.visible = False
    frameBalance.visible = False
    Me.FrameImpPunteo.visible = False
    FrameVerObservacionesCuentas.visible = False
    Me.FrameBloqueoEmpresas.visible = False
    Me.FrameIconosVisibles.visible = False
    Me.FrameInformeBBDD.visible = False
    Me.FrameShowProcess.visible = False
    Me.FrameCobros.visible = False
    Me.FrameAsientoLiquida.visible = False
    Me.FrameDescuadre.visible = False
    Me.FrameReclamaciones.visible = False
    Me.FrameBancosRemesas.visible = False
    Me.FrameRecibos.visible = False
    Me.FrameTalPagPdtes.visible = False
    FrameIntegrLiqNav.visible = False
    FrameIntegAsiento.visible = False
    FrameEliminarRiesgotalPag.visible = False
    FrameValidaIBAN.visible = False
    FrameElimiaRiesgoTalPag.visible = False
    FrameGeneraTransfer.visible = False
    FrameAmentoDigito.visible = False
    'YA ESTA DISPONIBLE  PonerFrameVisible
    
    ' botón de dividir vencimiento cuando estamos en talones/pagares pendientes
  
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 44
    End With
    
    
    Select Case Opcion
    Case 1
        Me.Caption = "Cálculo de saldo"
        W = frameSaldosHco.Width
        H = Me.frameSaldosHco.Height
        Me.frameSaldosHco.visible = True
        
        CargaValoresHco
        Command1(0).Cancel = True
    Case 2
        Me.Caption = "Comprobacion saldos"
        W = Me.frameCalculoSaldos.Width
        H = Me.frameCalculoSaldos.Height + 150
        Me.frameCalculoSaldos.visible = True
        Command1(1).Enabled = True
        Command2.Enabled = True
    Case 4
        Me.Caption = "Seleccion"
        W = Me.FrameeMPRESAS.Width
        H = Me.FrameeMPRESAS.Height + 200
        FrameeMPRESAS.top = 0
        FrameeMPRESAS.Left = 0
        Me.FrameeMPRESAS.visible = True
        cargaempresas
    Case 5
        'Lanzar automaticamente la comprobación de saldo
        Me.Caption = "Comprobacion saldos"
        PonerFrameVisible frameCalculoSaldos, H, W
        
        Command1(1).Enabled = False
        Command2.Enabled = False
    Case 7, 8
        Me.Caption = "Lineas configuracion balance"
        W = Me.frameBalance.Width
        H = Me.frameBalance.Height + 300
        Me.frameBalance.visible = True
        PonerCamposBalance
    Case 9, 10
        If Opcion = 9 Then
            Me.cmdCtaBalan(0).Caption = "Insertar"
        Else
            Me.cmdCtaBalan(0).Caption = "Modificar"
        End If
        Me.Caption = "Cuentas configuracion balances"
        W = Me.frameCtasBalance.Width
        H = Me.frameCtasBalance.Height + 300
        frameCtasBalance.visible = True
        PonerCamposCtaBalance
        
    Case 12
        'Saltos y repedtidos
        Me.Caption = "Búsqueda cabeceras asientos incorrectos"
        W = Me.frameSaltos.Width
        H = Me.frameSaltos.Height + 300
        Me.frameSaltos.visible = True
        Me.cmdCabError(0).Enabled = True
        Text5.Text = ""
        Text6.Text = ""
        cmdCabError(1).Cancel = True
        
    Case 18
        Me.FrameImpPunteo.visible = True
        Caption = "Importes"
        For i = 0 To 8
            Me.txtImporteP(i).Text = RecuperaValor(Parametros, i + 1)
        Next i
        W = Me.FrameImpPunteo.Width
        H = Me.FrameImpPunteo.Height + 300
        cmdPunteo.Cancel = True
        
    Case 21
        'obseravaciones cuenta
        FrameVerObservacionesCuentas.visible = True
        Caption = "Observaciones P.G.C."
        W = Me.FrameVerObservacionesCuentas.Width
        H = Me.FrameVerObservacionesCuentas.Height + 300
        
        cmdVerObservaciones.Cancel = True
        
        
    Case 22
        Me.FrameBloqueoEmpresas.visible = True
        Caption = "Bloqueo empresas"
        W = Me.FrameBloqueoEmpresas.Width
        H = Me.FrameBloqueoEmpresas.Height + 300
        'Como cuando venga por esta opcion, viene llamado desde el manteusu
        Me.ListView2(0).SmallIcons = frmMantenusu.ImageList1
        Me.ListView2(1).SmallIcons = frmMantenusu.ImageList1
        Me.cmdBloqEmpre(1).Cancel = True
        
        
    Case 24 ' iconos visbles
        Me.Caption = "Panel de Control"
        Me.FrameIconosVisibles.visible = True
        W = Me.FrameIconosVisibles.Width
        H = Me.FrameIconosVisibles.Height + 300
        
    Case 25 ' informe de base de datos
        Me.Caption = "Información de Base de Datos"
        Me.FrameInformeBBDD.visible = True
        W = Me.FrameInformeBBDD.Width
        H = Me.FrameInformeBBDD.Height + 300
        
        Me.Label47.Caption = "Ejercicio " & vParam.fechaini & " a " & vParam.fechafin
        Me.Label48.Caption = "Ejercicio " & DateAdd("yyyy", 1, vParam.fechaini) & " a " & DateAdd("yyyy", 1, vParam.fechafin)
        
    Case 26 ' show process list
        Me.Caption = "Información de Procesos del Sistema"
        Me.FrameShowProcess.visible = True
        W = Me.FrameShowProcess.Width
        H = Me.FrameShowProcess.Height + 300
        
        Label53.Caption = Label53.Caption & " Ariconta" & vEmpresa.codempre & " (" & vEmpresa.nomempre & ")"
        
    Case 27 ' cobros de facturas
        Me.Caption = "Facturas de Cliente"
        Label52(0).Caption = "Cobros de la Factura " & RecuperaValor(Parametros, 1) & "-" & Format(RecuperaValor(Parametros, 2), "0000000") & " de fecha " & RecuperaValor(Parametros, 3)
        Me.FrameCobros.visible = True
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height + 300
        cmdGenerarVto.visible = False
    Case 28 ' pagos de facturas
        Me.Caption = "Facturas de Proveedor"
        Label52(0).Caption = "Pagos de la Factura " & RecuperaValor(Parametros, 1) & "-" & RecuperaValor(Parametros, 3) & " de fecha " & RecuperaValor(Parametros, 4)
        Me.FrameCobros.visible = True
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height + 300
        cmdGenerarVto.visible = False
        
    Case 29 ' asiento de liquidacion
        Me.Caption = "Asiento de Liquidación"
        Me.FrameAsientoLiquida.visible = True
        W = Me.FrameAsientoLiquida.Width
        H = Me.FrameAsientoLiquida.Height + 300
        
        
    Case 30 ' asientos descuadrados
        Me.Caption = "Asientos descuadrados"
        Me.FrameDescuadre.visible = True
        W = Me.FrameDescuadre.Width
        H = Me.FrameDescuadre.Height + 300
        
    Case 31 ' facturas sin asientos
        Me.Caption = "Facturas sin asiento"
        Me.FrameDescuadre.visible = True
        W = Me.FrameDescuadre.Width
        H = Me.FrameDescuadre.Height + 300
        
    Case 50 ' facturas de reclamaciones
        Me.Caption = "Facturas Reclamadas"
        Me.Label30.Caption = "Reclamación a " & RecuperaValor(Parametros, 2) & " de fecha " & RecuperaValor(Parametros, 3)
        TamayoFrameReclama True
        Me.FrameReclamaciones.visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
        
        Orden = True
        CampoOrden = "fecfactu"
    
    Case 51 ' facturas de remesas
        FrameReclamaciones.top = 0
        FrameReclamaciones.Left = 0
        TamayoFrameReclama False
        Me.Caption = "Facturas de Remesa"
        Me.Label30.Caption = "Remesa " & RecuperaValor(Parametros, 1) & " / " & RecuperaValor(Parametros, 2)
        Me.FrameReclamaciones.visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
    
        Orden = True
        CampoOrden = "fecfactu"
            
    Case 52 ' bancos de remesas
        Me.Caption = "Remesas"
        Me.FrameBancosRemesas.visible = True
        W = Me.FrameBancosRemesas.Width
        H = Me.FrameBancosRemesas.Height + 300
    
    Case 53 ' recibos con cobros parciales
        Me.Caption = "Recibos "
        Me.FrameRecibos.visible = True
        W = Me.FrameRecibos.Width
        H = Me.FrameRecibos.Height + 300
            
    Case 54 ' cobros talones / pagares pendientes
        Me.Caption = "Talones/Pagarés Pendientes"
        Me.FrameTalPagPdtes.visible = True
        W = Me.FrameTalPagPdtes.Width
        H = Me.FrameTalPagPdtes.Height + 300
            
    Case 55 ' facturas de transferencias de abonos
        Me.Caption = "Facturas de Transferencia de Abonos"
        Me.Label30.Caption = "Transferencia " & RecuperaValor(Parametros, 1) & " / " & RecuperaValor(Parametros, 2)
        TamayoFrameReclama False
        Me.FrameReclamaciones.visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
    
        Orden = True
        CampoOrden = "fecfactu"

    Case 56, 66, 67 ' facturas de transferencias de pagos
        If Opcion = 56 Then
            Me.Caption = "Facturas de Transferencia de Pagos"
            Me.Label30.Caption = "Transferencia " & RecuperaValor(Parametros, 1) & " / " & RecuperaValor(Parametros, 2)
        Else
            CampoOrden = IIf(Opcion = 66, "confirming", "pagos domiciliado")
            Me.Caption = "Facturas de " & CampoOrden & "s"
            Me.Label30.Caption = UCase(CampoOrden) & ": " & RecuperaValor(Parametros, 1) & " / " & RecuperaValor(Parametros, 2)
        End If
        TamayoFrameReclama False
        Me.FrameReclamaciones.visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
    
        Orden = True
        CampoOrden = "fecfactu"
    
    Case 57 ' facturas de compensaciones
        Me.Caption = "Facturas de Compensación Cliente"
        Me.Label30.Caption = "Compensación " & RecuperaValor(Parametros, 1) & " de " & RecuperaValor(Parametros, 2)
        TamayoFrameReclama True
        Me.FrameReclamaciones.visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
    
        Orden = True
        CampoOrden = "fecfactu"
    
    Case 58 ' facturas de compensaciones proveedor
        Me.Caption = "Facturas de Compensación Proveedor"
        Me.Label30.Caption = "Compensación " & RecuperaValor(Parametros, 1) & " de " & RecuperaValor(Parametros, 2)
        TamayoFrameReclama True
        Me.FrameReclamaciones.visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
    
        Orden = True
        CampoOrden = "fecfactu"
    
    Case 59
        Limpiar Me
        Me.Caption = "Integra liquidacion"
        PonerFrameVisible FrameIntegrLiqNav, H, W
        cmdIntegFraNav(0).Cancel = True
    
    Case 60
        Limpiar Me
        Me.Caption = "Integra asientos"
        PonerFrameVisible Me.FrameIntegAsiento, H, W
        
        Me.txtDiario(0).Text = vParam.numdiacl
        Me.txtDiarioNom(0).Text = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(0).Text)
        cmdIntegAsiento(1).Cancel = True
        
    Case 61
        Me.Caption = "Riesgo"
        PonerFrameVisible FrameEliminarRiesgotalPag, H, W
        Me.txtfecha(1).Text = Format(CadenaDesdeOtroForm, "dd/mm/yyyy")
        CadenaDesdeOtroForm = ""
        cmdEliminarRiesgoT(1).Cancel = True
    Case 62
        Me.Caption = "IBAN"
        PonerFrameVisible FrameValidaIBAN, H, W
        H = H + 30
        For i = 0 To 5
            Me.txtIbanES(i).Text = ""
        Next
        Me.cboPais.ListIndex = 0
        
    Case 63
        Me.Caption = "Riesgo"
        PonerFrameVisible FrameElimiaRiesgoTalPag, H, W
        Me.txtfecha(2).Text = ""
        
        cboRiesgoYelConfirming.Clear
        i = CInt(Tipo)
        If i <> 2 Then
            cboRiesgoYelConfirming.AddItem "Remesas talón - pagaré"
            cboRiesgoYelConfirming.ItemData(cboRiesgoYelConfirming.NewIndex) = 1
        End If
        If i <> 1 Then
            cboRiesgoYelConfirming.AddItem "Pago confirming vencidos"
            cboRiesgoYelConfirming.ItemData(cboRiesgoYelConfirming.NewIndex) = 2
        End If
        cboRiesgoYelConfirming.ListIndex = 0
        ponerDatosRiesgoTalPag_Confir
        cmdElimRiesgoTalPag(0).Cancel = True
    
    Case 64
        Me.Caption = "Transferencias(Fichero)"
        PonerFrameVisible FrameGeneraTransfer, H, W
        SQL = ""
        Label7(7).Caption = "Transferencias " & IIf(RecuperaValor(Parametros, 3) = "1", "pagos", "abonos")
        
        H = H + 240
    Case 65
        Me.Caption = "Utilidades"
        PonerFrameVisible FrameAmentoDigito, H, W
        Me.cmdAumentoDigi(1).Cancel = True
        CadenaDesdeOtroForm = ""
        H = H + 240
        
        
        cboAumento.AddItem "2*"
        cboAumento.AddItem "40*"
        cboAumento.AddItem "41*"
        cboAumento.AddItem "43*"
        cboAumento.AddItem "57*"
        cboAumento.AddItem "6*"
        cboAumento.AddItem "7*"
        cboAumento.ListIndex = 3
    End Select
    
    Me.Width = W + 120
    Me.Height = H + 180
End Sub

Private Sub PonerFrameVisible(Fr As frame, ByRef H As Integer, ByRef W As Integer)

    Fr.top = 0
    Fr.Left = 0
    H = Fr.Height + 90
    W = Fr.Width + 60
    Fr.visible = True
    
End Sub





Private Sub TamayoFrameReclama(Normal As Boolean)
    If Normal Then
        Me.FrameReclamaciones.Width = 9795
    Else
        Me.FrameReclamaciones.Width = 14915
    End If
    Me.ListView9.Width = Me.FrameReclamaciones.Width - 390
    Me.Command1(3).Left = Me.FrameReclamaciones.Width - 1395
End Sub

Private Sub CargaValoresHco()
'Lo que hace es dado el parametro scamos la cuenta, nomcuenta, saldos
'1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Label6.Caption = RecuperaValor(Parametros, 1)
For i = 0 To 3
    Me.txtsaldo(i).Text = RecuperaValor(Parametros, i + 2)
Next i
CalculaSaldosFinales
End Sub



Private Sub CalculaSaldosFinales()
Dim Importe As Currency
    For i = 0 To 3
        Importe = ImporteFormateado(txtsaldo(i).Text)
        txtsaldo(i).Tag = Importe
    Next i
    txtsaldo(4).Text = ""
    txtsaldo(5).Text = ""
    
    Importe = CCur(txtsaldo(1).Tag) + CCur(txtsaldo(3).Tag)
    txtsaldo(5).Tag = Importe
    txtsaldo(5).Text = Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(0).Tag) + CCur(txtsaldo(2).Tag)
    txtsaldo(4).Tag = Importe
    txtsaldo(4).Text = Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(5).Tag) - CCur(txtsaldo(4).Tag)
    txtsaldo(6).Text = ""
    txtsaldo(7).Text = ""
    If Importe <> 0 Then
        If Importe > 0 Then
            txtsaldo(6).Text = Format(Importe, FormatoImporte)
        Else
            txtsaldo(7).Text = Format(Abs(Importe), FormatoImporte)
        End If
    End If
    
    
    'Ahora veremos si tiene del periodo
    txtsaldo(8).Text = ""
    txtsaldo(9).Text = ""
    SQL = RecuperaValor(Parametros, 6)
    If SQL = "" Then
        NE = 0
        
    Else
        NE = 1
        Importe = CCur(SQL)
        If Importe >= 0 Then
            txtsaldo(8).Text = Format(Importe, FormatoImporte)
        Else
            txtsaldo(9).Text = Format(Abs(Importe), FormatoImporte)
        End If
    End If
    
    Label28(1).visible = (NE = 1)
    txtsaldo(9).visible = (NE = 1)
    txtsaldo(8).visible = (NE = 1)
    
    'Descripcion cuenta
    SQL = Trim(RecuperaValor(Parametros, 7))   'Descripcion cuenta
    If SQL <> "" Then SQL = " - " & SQL
    Label6.Caption = Label6.Caption & SQL
    
    
    
    'NUEVO 14 Febrero... San valentin
    Importe = CCur(txtsaldo(2).Tag) - CCur(txtsaldo(3).Tag)
    Image6(0).ToolTipText = "Saldo punteado: " & Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(0).Tag) - CCur(txtsaldo(1).Tag)
    Image6(1).ToolTipText = "Saldo pendiente: " & Format(Importe, FormatoImporte)
    
    
    
End Sub


Private Sub cargaempresas()
Dim Prohibidas As String
On Error GoTo Ecargaempresas

    VerEmresasProhibidas Prohibidas
    
    SQL = "Select * from usuarios.empresasariconta "
    If vUsu.Codigo > 0 Then SQL = SQL & " WHERE codempre<100 and conta like 'ariconta%'"
    SQL = SQL & " order by codempre"
    Set lwE.SmallIcons = Me.ImageList1
    lwE.ListItems.Clear
    Set Rs = New ADODB.Recordset
    i = -1
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        SQL = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, SQL) = 0 Then
            Set ItmX = lwE.ListItems.Add(, , Rs!nomempre, , 3)
            ItmX.Tag = Rs!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                If CadenaDesdeOtroForm = "" Then
                    ItmX.Checked = True
                    i = ItmX.Index
                End If
            End If
            ItmX.ToolTipText = Rs!CONTA
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    If i > 0 Then Set lwE.SelectedItem = lwE.ListItems(i)
    
    CadenaDesdeOtroForm = ""
    
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
    Set Rs = Nothing
End Sub

Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from usuarios.usuarioempresasariconta WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
          VarProhibidas = VarProhibidas & Rs!codempre & "|"
          Rs.MoveNext
    Wend
    Rs.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
    Set Rs = Nothing
End Sub




Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmF_Selec(vFecha As Date)
    SQL = vFecha
End Sub

Private Sub frmPag_DatoSeleccionado2(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub Image3_Click()
    Set frmC = New frmColCtas
    frmC.ConfigurarBalances = 1
    frmC.DatosADevolverBusqueda = "0|1"
    SQL = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If SQL <> "" Then Text3.Text = RecuperaValor(SQL, 1)
    
End Sub


Private Sub Image6_Click(Index As Integer)
    MsgBox Image6(Index).ToolTipText, vbInformation
End Sub


Private Sub imgCheck_Click(Index As Integer)
    For NE = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(NE).Checked = Index = 1
    Next
    
    Select Case Index
        ' ICONOS VISIBLES EN EL LISTVIEW DEL FRMPPAL
        Case 2 ' marcar todos
            For i = 1 To ListView6.ListItems.Count
                ListView6.ListItems(i).Checked = True
            Next i
        Case 3 ' desmarcar todos
            For i = 1 To ListView6.ListItems.Count
                ListView6.ListItems(i).Checked = False
            Next i
            
        ' bancos remesados
        Case 4 ' marcar todos
            For i = 1 To ListView10.ListItems.Count
                ListView10.ListItems(i).Checked = True
            Next i
        Case 5 ' desmarcar todos
            For i = 1 To ListView10.ListItems.Count
                If ListView10.ListItems(i).Text <> Parametros Then ListView10.ListItems(i).Checked = False
            Next i
            
        ' talones/pagares pendientes
        Case 6 ' marcar todos
            For i = 1 To ListView12.ListItems.Count
                ListView12.ListItems(i).Checked = True
            Next i
            SumaTotales
        Case 7 ' desmarcar todos
            For i = 1 To ListView12.ListItems.Count
                ListView12.ListItems(i).Checked = (ListView12.ListItems(i).Tag = 1)
            Next i
            txtSuma.Text = ""
            
    End Select
        
    
End Sub



Private Sub imgConcepto_Click(Index As Integer)
    SQL = ""
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|"
    frmCon.Show vbModal
    If SQL <> "" Then
        Me.txtConcepto(Index).Text = RecuperaValor(SQL, 1)
        Me.txtConceptoNom(Index).Text = RecuperaValor(SQL, 2)
    End If
    Set frmCon = Nothing
End Sub

Private Sub imgCtas_Click(Index As Integer)
    Set frmC = New frmColCtas
    frmC.ConfigurarBalances = 0
    frmC.DatosADevolverBusqueda = "0|1"
    SQL = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If SQL <> "" Then
        Me.txtCodmacta(Index).Text = RecuperaValor(SQL, 1)
        Me.txtNommacta(Index).Text = RecuperaValor(SQL, 2)
    End If
End Sub

Private Sub imgDiario_Click(Index As Integer)
    SQL = ""
    
    Set frmD = New frmTiposDiario
    frmD.DatosADevolverBusqueda = "0|"
    frmD.Show vbModal
    If SQL <> "" Then
        Me.txtConcepto(Index).Text = RecuperaValor(SQL, 1)
        Me.txtConceptoNom(Index).Text = RecuperaValor(SQL, 2)
    End If
    Set frmCon = Nothing
End Sub

Private Sub imgFec_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Me.txtfecha(Index).Text <> "" Then frmF.Fecha = CDate(txtfecha(Index).Text)
    SQL = ""
    frmF.Show vbModal
    Set frmF = Nothing
    If SQL <> "" Then txtfecha(Index).Text = Format(SQL, "dd/mm/yyyy")
End Sub

Private Sub ListView10_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = Parametros Then
        If Not Item.Checked Then
            MsgBox "El banco por defecto no puede ser desmarcado. ", vbExclamation
            Item.Checked = True
        End If
    End If
End Sub


Private Sub ListView12_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 '   If Item.Tag = 1 Then Item.Checked = True
    SumaTotales
End Sub

Private Sub ListView13_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    For i = 1 To ListView13.ListItems.Count
        If Item.Index <> i Then
            ListView13.ListItems(i).Checked = False
        Else
            txtfecha(2).Text = ListView13.ListItems(i).SubItems(2)
        End If
    Next
End Sub

Private Sub ListView9_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    If Opcion = 51 Then
        Orden = Not Orden
        
        Select Case ColumnHeader
            Case "Serie"
                CampoOrden = "numserie"
            Case "Factura"
                CampoOrden = "numfactu"
            Case "Fecha"
                CampoOrden = "fecfactu"
            Case "Vto"
                CampoOrden = "numorden"
            Case "Fecha Vto"
                CampoOrden = "fecvenci"
            Case "Nombre"
                CampoOrden = "nomclien"
            Case "Importe"
                CampoOrden = "importe"
        End Select
        
        CargarFacturasRemesas
    ElseIf Opcion = 50 Then
        Orden = Not Orden
        
        Select Case ColumnHeader
            Case "Serie"
                CampoOrden = "numserie"
            Case "Factura"
                CampoOrden = "numfactu"
            Case "Fecha"
                CampoOrden = "fecfactu"
            Case "Vto"
                CampoOrden = "numorden"
            Case "Fecha Vto"
                CampoOrden = "fecvenci"
            Case "Importe"
                CampoOrden = "importe"
        End Select
        
        CargarFacturasReclamaciones
    
    End If
    
End Sub

Private Sub optIBAN_Intro_Click(Index As Integer)
Dim B As Boolean
    If Me.cboPais.ListIndex = 0 Then
        For i = 0 To 5
            B = Index = 1
            If i = 5 Then B = Not B
            BloqueaTXT txtIbanES(i), B
        Next i
        PonFoco txtIbanES(IIf(Index = 0, 0, 5))
            
    End If
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub tCuadre_Timer()
    tCuadre.Enabled = False
    Screen.MousePointer = vbHourglass
    Command2_Click
    Me.ListView1.Refresh
    Screen.MousePointer = vbHourglass
    espera 2
    Unload Me
End Sub




Private Sub PonerCamposBalance()
    If Opcion = 7 Then
        Label16.Caption = "NUEVO"
        Label16.ForeColor = &H800000
        Me.chkPintar.Value = 1
        For i = 0 To 3
            Text1(i).Text = ""
        Next i

    Else
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|LibroCD|
        Text1(0).Text = RecuperaValor(Parametros, 7)
        Text1(1).Text = RecuperaValor(Parametros, 8)
        i = Val(RecuperaValor(Parametros, 10))
        If i = 1 Then
            'Tiene cuentas
            Text1(2).Text = ""
            Text1(2).Enabled = False
        Else
            Text1(2).Text = RecuperaValor(Parametros, 9)
        End If
        i = Val(RecuperaValor(Parametros, 11))
        chkNegrita.Value = i
        i = Val(RecuperaValor(Parametros, 12))
        chkCero.Value = i
        i = Val(RecuperaValor(Parametros, 13))
        chkPintar.Value = i
        Text1(3).Text = RecuperaValor(Parametros, 14)
    End If
End Sub



Private Sub PonerCamposCtaBalance()
    'EL grupo se le pasa siempre
    Text2.Text = RecuperaValor(Parametros, 1)
    
    
    If Opcion = 9 Then
        Label19.Caption = "NUEVO"
        Label19.ForeColor = &H800000
        Text3.Text = ""
        Text3.Enabled = True
        chkResta.Value = 0
    Else
        Text3.Enabled = False
        Text3.Text = RecuperaValor(Parametros, 2)
        i = Val(RecuperaValor(Parametros, 3))
        Option1(i).Value = True
        i = Val(RecuperaValor(Parametros, 4))
        chkResta.Value = i
    End If
End Sub




Private Function InsertarModificar() As Boolean
Dim Aux As String

On Error GoTo EInse

    InsertarModificar = False
    
    'Comprobamos el concpeto del libro a CD
     Text1(3).Text = UCase(Trim(Text1(3).Text))
    If Text1(3).Text <> "" Then
        If Not IsNumeric(Text1(3).Text) Then
            MsgBox "El campo 'Concepto Libro CD' debe ser numérico", vbExclamation
            Exit Function
        End If
    End If
    
    'Hay k comprobar, si tiene formula k sea correcta
    Text1(2).Text = UCase(Trim(Text1(2).Text))
    If Text1(2).Text <> "" Then
        SQL = CompruebaFormulaConfigBalan(CInt(RecuperaValor(Parametros, 1)), Text1(2).Text)
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Function
        End If
    End If
    If Opcion = 7 Then
        SQL = "INSERT INTO balances_texto (NumBalan, Pasivo, codigo, padre, "
        SQL = SQL & "Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita,A_Cero,Pintar,LibroCD) VALUES ("
        SQL = SQL & RecuperaValor(Parametros, 1)  'Numero
        SQL = SQL & ",'" & RecuperaValor(Parametros, 2) 'pasivo
        SQL = SQL & "'," & RecuperaValor(Parametros, 3)  'Codigo
        Aux = RecuperaValor(Parametros, 4) 'padre
        If Aux = "" Then
            Aux = ",NULL,"
        Else
            Aux = ",'" & Aux & "',"
        End If
        SQL = SQL & Aux
        SQL = SQL & RecuperaValor(Parametros, 5)
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        SQL = SQL & "," & Aux
        SQL = SQL & ",'" & Text1(0).Text 'Text linea
        SQL = SQL & "','" & Text1(1).Text 'Desc linea
        SQL = SQL & "','" & Text1(2).Text 'Formula
        SQL = SQL & "',0," & chkNegrita.Value
        SQL = SQL & "," & Me.chkCero.Value
        SQL = SQL & "," & Me.chkPintar.Value
        SQL = SQL & ",'" & Text1(3).Text 'Libro CD
        SQL = SQL & "')"
    Else
        'Modificar
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|
        SQL = "UPDATE balances_texto SET "
        SQL = SQL & "deslinea='" & Text1(0).Text & "',"
        SQL = SQL & "texlinea='" & Text1(1).Text & "',"
        SQL = SQL & "formula='" & Text1(2).Text & "',"
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        SQL = SQL & "Tipo =" & Aux & ","
        SQL = SQL & "Negrita = " & chkNegrita.Value
        SQL = SQL & ", A_Cero = " & Me.chkCero.Value
        SQL = SQL & ", Pintar = " & Me.chkPintar.Value
        SQL = SQL & ", LibroCD = '" & Text1(3).Text & "'"
        SQL = SQL & " WHERE numbalan =" & RecuperaValor(Parametros, 1)
        SQL = SQL & " AND Pasivo = '" & RecuperaValor(Parametros, 2)
        SQL = SQL & "' AND codigo = " & RecuperaValor(Parametros, 3)
        
    End If
    Conn.Execute SQL
    InsertarModificar = True
    'Ha insertado
    'Devuelve el texto, el texto auxiliar, y si es formula o no, descripcion cta y concepto oficial
    CadenaDesdeOtroForm = Text1(0).Text & "|" & Text1(1).Text & "|" & Aux & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(3).Text & "|"
    Exit Function
EInse:
    MuestraError Err.Number
End Function





Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub LeerCadenaFicheroTexto()
On Error GoTo ELeerCadenaFicheroTexto
    'Son dos lineas. La primaera indica k campo y la segunda el valor
    Line Input #i, SQL
    Line Input #i, SQL
    Exit Sub
ELeerCadenaFicheroTexto:
    SQL = ""
    Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Function ValorSQL(ByRef C As String) As String
    If C = "" Then
        ValorSQL = "NULL"
    Else
        ValorSQL = "'" & C & "'"
    End If
End Function
Private Function EjecutaSQL2(SQL As String) As Boolean
    EjecutaSQL2 = False
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        AnyadeErrores "SQL: " & SQL, Err.Description
        Err.Clear
    Else
        EjecutaSQL2 = True
    End If
End Function


Private Sub AnyadeErrores(L1 As String, L2 As String)
    NE = NE + 1
    Errores = Errores & "-----------------------------" & vbCrLf
    Errores = Errores & L1 & vbCrLf
    Errores = Errores & L2 & vbCrLf


End Sub

Private Sub ImprimeFichero()
Dim NF As Integer
    On Error GoTo EImprimeFichero
    NF = FreeFile
    Open App.Path & "\errimpdat.txt" For Output As #NF
    Print #NF, Errores
    Close (NF)
    Shell "notepad.exe " & App.Path & "\errimpdat.txt", vbMaximizedFocus
    Exit Sub
EImprimeFichero:
    MsgBox Err.Description & vbCrLf, vbCritical
    Err.Clear
End Sub


Private Function ExisteCuenta(Cta As String) As Boolean
    
    ExisteCuenta = False
    SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Cta, "T")
    If SQL <> "" Then ExisteCuenta = True
    
End Function

Private Sub Text3_LostFocus()
    If IsNumeric(Text3.Text) Then
        If InStr(1, Text3.Text, ".") > 0 Then
            SQL = Trim(Text3.Text)
            Errores = ""
            If CuentaCorrectaUltimoNivel(SQL, Errores) Then
                Text3.Text = SQL
                
            Else
                If SQL <> "" Then Text3.Text = SQL
            End If
        End If
    End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    'dividir vencimientos
    If Me.ListView12.SelectedItem Is Nothing Then Exit Sub
    
    DividirVencimiento

End Sub

Private Sub DividirVencimiento()
Dim Im As Currency

    If ListView12.ListItems.Count = 0 Then Exit Sub
    If ListView12.SelectedItem Is Nothing Then Exit Sub
        
    
    
    'Si esta totalmente cobrado pues no podemos desdoblar ekl vto
    Im = ImporteFormateado(ListView12.SelectedItem.SubItems(9))
    If Im <= 0 Then
        MsgBox "NO puede dividir el vencimiento. Importe totalmente cobrado", vbExclamation
        Exit Sub
    End If
    
       'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
    
    CadenaDesdeOtroForm = "numserie = " & DBSet(ListView12.SelectedItem.Text, "T") & " AND numfactu = " & ListView12.SelectedItem.SubItems(1)
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND fecfactu = " & DBSet(ListView12.SelectedItem.SubItems(2), "F") & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView12.SelectedItem.SubItems(4) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CStr(Im) & "|"
    
    
    'Ok, Ahora pongo los labels
    frmTESCobrosDivVto.Opcion = 27
    frmTESCobrosDivVto.Label4(56).Caption = Label33.Caption
    frmTESCobrosDivVto.txtCodigo(2).Text = ListView12.SelectedItem.SubItems(3)
    
    frmTESCobrosDivVto.Label4(57).Caption = ListView12.SelectedItem.Text & Format(ListView12.SelectedItem.SubItems(1), "0000000") & " / " & ListView12.SelectedItem.SubItems(4) & "      de " & Format(ListView12.SelectedItem.SubItems(2), "dd/mm/yyyy")
    
    'Si ya ha cobrado algo...
    Im = DBLet(ComprobarCero(Trim(ListView12.SelectedItem.SubItems(8))), "N")
    If Im > 0 Then frmTESCobrosDivVto.txtCodigo(1).Text = ListView12.SelectedItem.SubItems(9)
    
    If Text1(0).Text = "" Then
        MsgBox "El cobro no tiene forma de pago. Revise.", vbExclamation
        Exit Sub
    End If
    
    
    frmTESCobrosDivVto.Show vbModal
    
    
    
    
    If CadenaDesdeOtroForm <> "" Then
        CargarTalonesPagaresPendientes
    End If
    
    
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node
    'Si es padre
    If Node.Parent Is Nothing Then
        If Node.Children > 0 Then
            Set N = Node.Child
            Do
                N.Checked = Node.Checked
                Set N = N.Next
            Loop Until N Is Nothing
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------
'
'




Private Sub EncabezadoPieFact(Pie As Boolean, ByVal Text As String, REG As Integer)
    If Pie Then
        Text = "[/" & Text & "]" & REG
    Else
        Text = "[" & Text & "]"
    End If
    Print #NE, Text
End Sub


Private Sub InsertaEnTmpCta()
On Error Resume Next
    
    Conn.Execute "INSERT INTO tmpcierre1 (codusu, cta) VALUES (" & vUsu.Codigo & ",'" & Rs.Fields(0) & "')"
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub EjecutarSQL()
    On Error Resume Next
    
    Conn.Execute SQL
    If Err.Number <> 0 Then
        If Conn.Errors(0).Number = 1062 Then
            Err.Clear
        Else
            'MuestraError Err.Number, Err.Description
        End If
        Err.Clear
    End If
End Sub



Private Sub cargarObservacionesCuenta()
    Set Rs = New ADODB.Recordset
    SQL = "select codmacta,nommacta,obsdatos from cuentas where codmacta = '" & Parametros & "'"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        SQL = Rs!codmacta & "|" & Rs!Nommacta & "|" & DBMemo(Rs!obsdatos) & "|"
    Else
        SQL = "Err|ERROR  LEYENDO DATOS CUENTAS | ****  ERROR ****|"
    End If
    Rs.Close
    Set Rs = Nothing
    For i = 1 To 3
        Text1(i + 3).Text = RecuperaValor(SQL, i)
    Next i
End Sub


Private Sub cargaempresasbloquedas()
Dim IT As ListItem
    On Error GoTo Ecargaempresasbloquedas
    Set Rs = New ADODB.Recordset
    SQL = "select empresasariconta.codempre,nomempre,nomresum,usuarioempresasariconta.codempre bloqueada from usuarios.empresasariconta left join usuarios.usuarioempresasariconta on "
    SQL = SQL & " empresasariconta.codempre = usuarioempresasariconta.codempre And (usuarioempresasariconta.codusu = " & Parametros & " Or codusu Is Null)"
    '[Monica] solo ariconta
    SQL = SQL & " WHERE conta like 'ariconta%' "
    SQL = SQL & " ORDER BY empresasariconta.codempre"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Errores = Format(Rs!codempre, "00000")
        SQL = "C" & Errores
        
        If IsNull(Rs!bloqueada) Then
            'Va al list de la derecha
            Set IT = ListView2(0).ListItems.Add(, SQL)
            IT.SmallIcon = 1
        Else
            Set IT = ListView2(1).ListItems.Add(, SQL)
            IT.SmallIcon = 2
        End If
        IT.Text = Errores
        IT.SubItems(1) = Rs!nomempre
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Errores = ""
    Exit Sub
Ecargaempresasbloquedas:
    MuestraError Err.Number, Err.Description
    Me.cmdBloqEmpre(0).Enabled = False
    Errores = ""
    Set Rs = Nothing
End Sub










'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'Restore desde backup
'
'




Private Function CobroContabilizado(Serie As String, FACTURA As String, Fecha As String, Orden As String) As Boolean
Dim SQL As String

    SQL = "select * from hlinapu where numserie = " & DBSet(Serie, "T") & " and numfaccl = " & DBSet(FACTURA, "N") & " and fecfactu = " & DBSet(Fecha, "F") & " and numorden = " & DBSet(Orden, "N")
    CobroContabilizado = (TotalRegistrosConsulta(SQL) <> 0)

End Function

Private Function PagoContabilizado(Serie As String, Proveedor As String, FACTURA As String, Fecha As String, Orden As String) As Boolean
Dim SQL As String

    SQL = "select * from hlinapu where numserie = " & DBSet(Serie, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfacpr = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(Fecha, "F") & " and numorden = " & DBSet(Orden, "N")
    PagoContabilizado = (TotalRegistrosConsulta(SQL) <> 0)

End Function



Private Sub CargaCobrosFactura()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaCobrosFactura
    
    Set Rs = New ADODB.Recordset
    
    ListView5.ColumnHeaders.Clear
    
    ListView5.ColumnHeaders.Add , , "Ord.", 800.0631
    ListView5.ColumnHeaders.Add , , "Forma de Pago", 3000.2522
    ListView5.ColumnHeaders.Add , , "Fecha Vto", 1450.2522
    ListView5.ColumnHeaders.Add , , "Importe Vto", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "Gastos", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "F.Ult.Cobro", 1450.2522
    ListView5.ColumnHeaders.Add , , "Imp.Pagado", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "Pendiente", 1550.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    ListView5.SmallIcons = frmppal.ImgListComun
    
    Cad = "select numorden, formapago.nomforpa, fecvenci, impvenci, gastos, fecultco, impcobro, (coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) pendiente, cobros.ctabanc1  "
    Cad = Cad & " from (cobros left join formapago on cobros.codforpa = formapago.codforpa) "
    Cad = Cad & " where cobros.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    Cad = Cad & " and cobros.numfactu = " & DBSet(RecuperaValor(Parametros, 2), "N")
    Cad = Cad & " and cobros.fecfactu = " & DBSet(RecuperaValor(Parametros, 3), "F")
    Cad = Cad & " order by numorden "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If CobroContabilizado(RecuperaValor(Parametros, 1), RecuperaValor(Parametros, 2), RecuperaValor(Parametros, 3), DBLet(Rs.Fields(0))) Then IT.SmallIcon = 18
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = Format(DBLet(Rs.Fields(3)), "###,###,##0.00")
        
        
        IT.SubItems(4) = " " & Format(DBLet(Rs.Fields(4)), "###,###,##0.00")
        IT.SubItems(5) = " " & DBLet(Rs.Fields(5))
        IT.SubItems(6) = " " & Format(DBLet(Rs.Fields(6)), "###,###,##0.00")
        IT.SubItems(7) = Format(DBLet(Rs.Fields(7)), "###,###,##0.00")
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargaCobrosFactura:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub



Private Sub CargaPagosFactura()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaPagosFactura
    
    Set Rs = New ADODB.Recordset
    
    ListView5.ColumnHeaders.Clear
    
    ListView5.ColumnHeaders.Add , , "Ord.", 800.0631
    ListView5.ColumnHeaders.Add , , "Forma de Pago", 3000.2522
    ListView5.ColumnHeaders.Add , , "Fecha Vto", 1450.2522
    ListView5.ColumnHeaders.Add , , "Importe Vto", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "F.Ult.Pago", 1450.2522
    ListView5.ColumnHeaders.Add , , "Imp.Pagado", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "Pendiente", 1550.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    ListView5.SmallIcons = frmppal.ImgListComun
    
    Cad = "select numorden, formapago.nomforpa, fecefect, impefect, fecultpa, imppagad, (coalesce(impefect,0)  - coalesce(imppagad,0)) pendiente, pagos.ctabanc1  "
    Cad = Cad & " from (pagos left join formapago on pagos.codforpa = formapago.codforpa) "
    Cad = Cad & " where pagos.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    Cad = Cad & " and pagos.codmacta = " & DBSet(RecuperaValor(Parametros, 2), "T")
    Cad = Cad & " and pagos.numfactu = " & DBSet(RecuperaValor(Parametros, 3), "T")
    Cad = Cad & " and pagos.fecfactu = " & DBSet(RecuperaValor(Parametros, 4), "F")
    Cad = Cad & " order by numorden "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If PagoContabilizado(RecuperaValor(Parametros, 1), RecuperaValor(Parametros, 2), RecuperaValor(Parametros, 3), RecuperaValor(Parametros, 4), DBLet(Rs.Fields(0))) Then IT.SmallIcon = 18

'        If DBLet(RS!NumAsien, "N") <> 0 Then IT.SmallIcon = 18
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = Format(DBLet(Rs.Fields(3)), "###,###,##0.00")
        IT.SubItems(4) = " " & DBLet(Rs.Fields(4))
        IT.SubItems(5) = " " & Format(DBLet(Rs.Fields(5)), "###,###,##0.00")
        IT.SubItems(6) = Format(DBLet(Rs.Fields(6)), "###,###,##0.00")
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    
    
    
    If ListView5.ListItems.Count = 0 Then
        'NO hay pagos.
        'Si la factura NO es trasapasada, permitiré generera el pago desde aqui
        Cad = " select estraspasada,factpro.codmacta,factpro.nommacta, codforpa,ctabanco,iban,totfacpr,FecFactu"
        Cad = Cad & " from (factpro left join cuentas on factpro.codmacta= cuentas.codmacta)  "
        

        Cad = Cad & " where factpro.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
        Cad = Cad & " and factpro.codmacta = " & DBSet(RecuperaValor(Parametros, 2), "T")
        Cad = Cad & " and factpro.numfactu = " & DBSet(RecuperaValor(Parametros, 3), "T")
        Cad = Cad & " and factpro.fecfactu = " & DBSet(RecuperaValor(Parametros, 4), "F")
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cmdGenerarVto.Tag = ""
        Codigo = ""
        If Not Rs.EOF Then
            If DBLet(Rs!estraspasada, "N") = 0 Then
                'OK , pongo en el tag lo que necesito
                Cad = ""
                If DBLet(Rs!CtaBanco, "T") <> "" Then
                    Cad = "bancos.codmacta=cuentas.codmacta and bancos.codmacta = " & DBSet(Rs!CtaBanco, "T") & " AND 1"
                    Cad = DevuelveDesdeBD("if(descripcion is null, nommacta,descripcion)", "bancos,cuentas", Cad, "1")
                End If
                'CtaBanco & "|" & "|" & "|" & "|" & "|" & IBAN & "|" & TipForpa & "|" & NomBanco & "|" & DBLet(Data1.Recordset!totfacpr, "N") & "|"
                cmdGenerarVto.Tag = DBLet(Rs!CtaBanco, "T") & "|||||" & DBLet(Rs!IBAN, "T") & "|0"
                cmdGenerarVto.Tag = cmdGenerarVto.Tag & "|" & Cad & "|" & DBLet(Rs!totfacpr, "N") & "|"
                
                Codigo = DBLet(Rs!Codforpa, "N") & "|" & Rs!FecFactu & "|" & Rs!totfacpr & "|"
            End If
        End If
        Rs.Close
        
        
        cmdGenerarVto.visible = cmdGenerarVto.Tag <> ""
    Else
        cmdGenerarVto.visible = False
    End If
    
    
    
    Set Rs = Nothing
    
    Exit Sub
    
ECargaPagosFactura:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub



Private Sub CargarAsiento()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

Dim MultiEmpresa As Boolean
Dim EmprAnterior As Integer
Dim VerCuadrarApunte As Boolean

    On Error GoTo ECargarAsiento
    
    
    ListView7.ColumnHeaders.Clear
    ListView7.ListItems.Clear
        
    Label7(3).visible = False
    If Parametros = "0" Then
        'Solo generar
        CmdContabilizar.Caption = "Generar"
    Else
        CmdContabilizar.Caption = "Contabilizar"
        Label7(3).visible = True
    End If
    
    Set Rs = New ADODB.Recordset
    
    Cad = "SELECT distinct numasien FROM tmpconext where tmpconext.codusu = " & DBSet(vUsu.Codigo, "N")
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
        Cad = Cad & "X"
        Rs.MoveNext
    Wend
    Rs.Close
    
    MultiEmpresa = Len(Cad) > 1
    
    
    If MultiEmpresa Then ListView7.ColumnHeaders.Add , , "Empr", 800.0631
    ListView7.ColumnHeaders.Add , , "Cuenta", 1600.2522
    ListView7.ColumnHeaders.Add , , "Denominación", 4000.2522
    ListView7.ColumnHeaders.Add , , "Debe", 2050.2522, 1
    ListView7.ColumnHeaders.Add , , "Haber", 2050.2522, 1
    ListView7.ColumnHeaders.Add , , "Saldo", 2050.2522, 1
    
    
    
    
    Pos = DevuelveValor("select max(pos) from tmpconext where codusu = " & DBSet(vUsu.Codigo, "N"))
    
    
    Cad = "select tmpconext.cta, cuentas.nommacta, tmpconext.timported, tmpconext.timporteh, acumtotT, tmpconext.pos, ampconce, numasien "
    Cad = Cad & " from (ariconta" & NumConta & ".tmpconext left join ariconta" & NumConta & ".cuentas on tmpconext.cta = cuentas.codmacta) "
    Cad = Cad & " left join ariconta" & NumConta & ".tmpconextcab on tmpconext.codusu = tmpconextcab.codusu and tmpconext.cta = tmpconextcab.cta"
    Cad = Cad & " where tmpconext.codusu = " & DBSet(vUsu.Codigo, "N")
    Cad = Cad & " order by pos "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    EmprAnterior = -1
    Importe = 0
    While Not Rs.EOF
                    
        Set IT = ListView7.ListItems.Add
        
        If MultiEmpresa Then
            IT.Text = Rs!NumAsien
            IT.SubItems(1) = Rs.Fields(0)
            NE = 2
            
  
            
        Else
            IT.Text = DBLet(Rs.Fields(0))
            NE = 1
        End If
        
        Equipo = DBLet(Rs!Ampconce, "T")
        If Equipo = "" Then Equipo = IT.Text
        IT.SubItems(NE) = Equipo
        NE = NE + 1
        If DBLet(Rs.Fields(2)) <> 0 Then
            IT.SubItems(NE) = Format(DBLet(Rs.Fields(2)), "###,###,##0.00")
        Else
            IT.SubItems(NE) = " "
        End If
        NE = NE + 1
        If DBLet(Rs.Fields(3)) <> 0 Then
            IT.SubItems(NE) = Format(DBLet(Rs.Fields(3)), "###,###,##0.00")
        Else
            IT.SubItems(NE) = " "
        End If
        
        ' si no estamos en la última línea mostramos el saldo de la cuenta
        NE = NE + 1
        IT.SubItems(NE) = " "
        If DBLet(Rs.Fields(5)) <> Pos Then
            If DBLet(Rs.Fields(4)) <> 0 Then
                IT.SubItems(NE) = Format(DBLet(Rs.Fields(4)), "###,###,##0.00")
            Else
                IT.SubItems(NE) = " "
            End If
            IT.ListSubItems(NE).ForeColor = &HAE8859   '&HEED68C
'            IT.ListSubItems(4).Bold = True
            
        End If
        
        
        VerCuadrarApunte = False
        If DBLet(Rs.Fields(5)) = Pos Then
            IT.Bold = True
            IT.ListSubItems(1).Bold = True
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(3).Bold = True
'            IT.ListSubItems(4).Bold = True

        Else
            
            'No es el ultimo registro. Vemos si es muliempresa, para cuadrar el apunte
            If MultiEmpresa Then
                VerCuadrarApunte = True
                Importe = Importe + Rs.Fields(2) - Rs.Fields(3)
                    
            End If
    
        End If
        
        'Siguiente
        Rs.MoveNext
        
        
        'Ver VerCuadrarApunte = True
        If VerCuadrarApunte Then
            If Not Rs.EOF Then
                If EmprAnterior >= 0 And EmprAnterior <> Rs!NumAsien And Importe <> 0 Then
                    'CAMBIO DE EMPRESA. HAY QUE CUADRAR si no es el ultimo, o si la empresa no es donde se esta haciendo la liquidacion
                    If EmprAnterior <> vEmpresa.codempre Then
                        'If DBLet(Rs.Fields(5)) <> Pos Then
                        
                        
                            Set IT = ListView7.ListItems.Add
                            IT.Text = EmprAnterior
                            If Importe < 0 Then
                                Equipo = "ctahpacreedor"
                                IT.SubItems(4) = " "
                                IT.SubItems(3) = Format(Abs(Importe), FormatoImporte)
                                IT.ListSubItems(3).ForeColor = vbRed
                            Else
                                Equipo = "ctahpdeudor"
                                IT.SubItems(3) = " "
                                IT.SubItems(4) = Format(Importe, FormatoImporte)
                                IT.ListSubItems(4).ForeColor = vbRed
                            End If
                            Equipo = DevuelveDesdeBD(Equipo, "ariconta" & EmprAnterior & ".parametros", "1", "1")
                            IT.SubItems(1) = Equipo & " "
                            Equipo = DevuelveDesdeBD("nommacta", "ariconta" & EmprAnterior & ".cuentas", "codmacta", Equipo, "T")
                            If Equipo = "" Then Equipo = "NO encontrada"
                            IT.SubItems(2) = Equipo
                            IT.ListSubItems(2).ForeColor = vbRed
                            IT.SubItems(5) = " "
                        'End If
                    End If
                    Importe = 0
                End If
                EmprAnterior = Rs!NumAsien
            End If
        End If
        
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarAsiento:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub



Private Sub CargarAsientosDescuadrados()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarAsiento
    
    
    ListView8.ColumnHeaders.Clear
    ListView8.ListItems.Clear
    
    ListView8.ColumnHeaders.Add , , "Diario", 1000.2522
    ListView8.ColumnHeaders.Add , , "Asiento", 1500.2522
    ListView8.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView8.ColumnHeaders.Add , , "Debe", 2050.2522, 1
    ListView8.ColumnHeaders.Add , , "Haber", 2050.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    
    Cad = "select tmphistoapu.numdiari, tmphistoapu.numasien, tmphistoapu.fechaent, tmphistoapu.timported, tmphistoapu.timporteh, tmphistoapu.timported - tmphistoapu.timporteh "
    Cad = Cad & " from tmphistoapu "
    Cad = Cad & " where tmphistoapu.codusu = " & DBSet(vUsu.Codigo, "N")
    Cad = Cad & " order by numdiari, numasien, fechaent "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView8.ListItems.Add
        
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        If DBLet(Rs.Fields(3)) <> 0 Then
            IT.SubItems(3) = Format(DBLet(Rs.Fields(3)), "###,###,##0.00")
        Else
            IT.SubItems(3) = ""
        End If
        If DBLet(Rs.Fields(4)) <> 0 Then
            IT.SubItems(4) = Format(DBLet(Rs.Fields(4)), "###,###,##0.00")
        Else
            IT.SubItems(4) = ""
        End If
        
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarAsiento:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub



Private Sub CargarFacturasSinAsientos()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarAsiento
    
    
    ListView8.ColumnHeaders.Clear
    ListView8.ListItems.Clear
    
    
    ListView8.ColumnHeaders.Add , , "Serie", 800.2522
    ListView8.ColumnHeaders.Add , , "Descripción", 2700.2522
    ListView8.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView8.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView8.ColumnHeaders.Add , , "Total", 2000.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    Cad = "select tmpfaclin.numserie, tmpfaclin.nomserie, tmpfaclin.numfac, tmpfaclin.fecha, tmpfaclin.total "
    Cad = Cad & " from tmpfaclin "
    Cad = Cad & " where tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
    Cad = Cad & " order by numserie, numfac, fecha "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView8.ListItems.Add
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = DBLet(Rs.Fields(3))
        If DBLet(Rs.Fields(4)) <> 0 Then
            IT.SubItems(4) = Format(DBLet(Rs.Fields(4)), "###,###,##0.00")
        Else
            IT.SubItems(4) = ""
        End If
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarAsiento:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub




Private Sub CargarFacturasReclamaciones()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmppal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 2000.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 2000.2522
    ListView9.ColumnHeaders.Add , , "Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    Cad = "select numserie, numfactu, fecfactu, numorden, impvenci importe "
    Cad = Cad & " from reclama_facturas "
    Cad = Cad & " where codigo = " & DBSet(RecuperaValor(Parametros, 1), "N")
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
'    Cad = Cad & " order by numserie, numfactu, fecfactu "
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = DBLet(Rs.Fields(3))
        If DBLet(Rs.Fields(4)) <> 0 Then
            IT.SubItems(4) = Format(DBLet(Rs.Fields(4)), "###,###,##0.00")
        Else
            IT.SubItems(4) = ""
        End If
        
        SQL = "select devuelto from cobros where numserie = " & DBSet(Rs.Fields(0), "T") & " and numfactu = " & DBSet(Rs.Fields(1), "N")
        SQL = SQL & " and fecfactu = " & DBSet(Rs.Fields(2), "F") & " and numorden = " & DBSet(Rs.Fields(3), "N")
        
        If DevuelveValor(SQL) = 1 Then
            IT.SmallIcon = 42
        End If
        
        'Siguiente
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarFacturas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub


Private Sub CargarFacturasRemesas()
Dim IT As ListItem
'Dim TotalArray  As Long
'Dim SERVER As String
'Dim EquipoConBD As Boolean
Dim Cad As String
'Dim Equipo As String
'Dim Pos As Long
Dim T1 As Single
Dim rApu As ADODB.Recordset
Dim Col As Collection
Dim Esta As Boolean

Dim FechaContab As Date

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmppal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 1500.2522

    ListView9.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView9.ColumnHeaders.Add , , "Vto", 800.2522
    ListView9.ColumnHeaders.Add , , "Fecha Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Nombre", 5000.2522
    ListView9.ColumnHeaders.Add , , "Gastos", 1000.2522, 1
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    ListView9.ColumnHeaders.Add , , "ID", 0, 1  'Oculto. Numserie|numfac
    
    Me.Refresh
    DoEvents
    Screen.MousePointer = vbHourglass
    Me.Label30.Caption = "Leyendo BD"
    Label30.Refresh
    
    Set Rs = New ADODB.Recordset
    Set rApu = New ADODB.Recordset
    Set Col = New Collection
    
    
    Cad = "select cobros.numserie, cobros.numfactu, cobros.fecfactu, cobros.numorden , cobros.fecvenci,nomclien , cobros.gastos, cobros.impvenci  importe, ' ' devol ,fecultco"
'    Cad = Cad & " ,concat(substring(concat(numserie,'  '),1,3),right(concat('0000000',cobros.numfactu),7),fecvenci,numorden) as localizador"
    Cad = Cad & " from cobros "
    Cad = Cad & " where (cobros.codrem = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and cobros.anyorem = " & DBSet(RecuperaValor(Parametros, 2), "N") & ") "
    
    Cad = Cad & " ORDER BY 1,2,3,4"   'numserie, cobros.numfactu, cobros.fecfactu, cobros.numorden
    
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs.EOF Then Err.Raise 513, , "No se encontro remesa: " & RecuperaValor(Parametros, 1) & " " & RecuperaValor(Parametros, 2)
    FechaContab = "31/12/2200"
    i = 0
    While Not Rs.EOF
        If Not IsNull(Rs!fecultco) Then
            If FechaContab > Rs!fecultco Then FechaContab = Rs!fecultco
        End If
        Cad = Rs.Fields(0) & "|"
        Cad = Cad & Format(Rs.Fields(1), "000000") & "|"
        
        For J = 2 To 5
            Cad = Cad & DBLet(Rs.Fields(J)) & "|"
        Next
        
        'gastos
        Cad = Cad & DBLet(Rs.Fields(6), "N") & "|"
        
        'importe
        Cad = Cad & DBLet(Rs.Fields(7), "N") & "|"
            
        
        Cad = Cad & Rs.Fields(8) & "|"
        
        Col.Add Cad
        Rs.MoveNext
    Wend
    
    
    If FechaContab = "31/12/2200" Then FechaContab = vParam.fechaini
    Me.Label30.Caption = "Leyendo BD(II)  "
    Label30.Refresh
    T1 = Timer
    
    Cad = " select hlinapu.numserie, hlinapu.numfaccl, hlinapu.fecfactu, hlinapu.numorden, cobros.fecvenci, hlinapu.gastodev, coalesce(hlinapu.timporteh,0) - coalesce(hlinapu.timported,0) importe, '*' devol,hlinapu.codmacta"
    Cad = Cad & " from cobros inner join hlinapu on cobros.numserie = hlinapu.numserie and cobros.numfactu = hlinapu.numfaccl and cobros.fecfactu = hlinapu.fecfactu and cobros.numorden = hlinapu.numorden "
    
    Cad = Cad & " where (hlinapu.codrem = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and hlinapu.anyorem = " & DBSet(RecuperaValor(Parametros, 2), "N") & " and hlinapu.esdevolucion = 0) "
    Cad = Cad & " AND hlinapu.fechaent  >=" & DBSet(FechaContab, "F")
    Cad = Cad & " ORDER BY 1,2,3,4"   'numserie, cobros.numfactu, cobros.fecfactu, cobros.numorden
    rApu.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    While Not rApu.EOF
        If Timer - T1 > 0.8 Then
            Me.Label30.visible = Not Me.Label30.visible
            Me.Label30.Caption = "Leyendo BD"
            Screen.MousePointer = vbHourglass
            Label30.Refresh
            T1 = Timer
        End If
        Rs.MoveFirst
        Cad = ""
        Esta = False
        While Cad = ""
        
                If Rs!NUmSerie = rApu!NUmSerie Then
                    If Rs!NumFactu = rApu!numfaccl Then
                        If Rs!FecFactu = rApu!FecFactu Then
                            If Rs!numorden = rApu!numorden Then
                                Cad = "ESTA"
                                Esta = True
                            End If
                        End If
                    End If
                End If
                If Cad = "" Then
                    Rs.MoveNext
                    If Rs.EOF Then Cad = "NO"
                End If
        Wend
        
        If Not Esta Then
            'NO esta. lo añadp
            
            
            Cad = rApu.Fields(0) & "|"
            For J = 1 To 4
                Cad = Cad & DBLet(rApu.Fields(J)) & "|"
            Next
            
            Cad = Cad & DevuelveDesdeBD("nommacta", "cuentas", "apudirec='S' AND codmacta", rApu!codmacta, "T") & "|"
            'gastos
            Cad = Cad & DBLet(rApu.Fields(5), "N") & "|"
            
            'importe
            Cad = Cad & DBLet(rApu.Fields(6), "N") & "|"
                
            
            Cad = Cad & rApu.Fields(7) & "|"
            
            Col.Add Cad
        End If
            
        rApu.MoveNext
    Wend
    rApu.Close
    Rs.Close
    
    Me.Label30.visible = True
    Me.Label30.Caption = "Mostrando datos"
    
    ListView9.SortOrder = lvwAscending
    ListView9.SortKey = 7
    ListView9.Sorted = True
    
    For J = 1 To Col.Count
                    
        Set IT = ListView9.ListItems.Add
        Cad = Col.Item(J)
        IT.Text = RecuperaValor(Cad, 1)
        For i = 1 To 4
            IT.SubItems(i) = RecuperaValor(Cad, i + 1)
        Next
        IT.SubItems(5) = RecuperaValor(Cad, 6)
                        
        'gastos
        SQL = RecuperaValor(Cad, 7)
        If SQL <> 0 Then
            IT.SubItems(6) = Format(CCur(SQL), "###,###,##0.00")
        Else
            IT.SubItems(6) = " "
        End If
        
        'importe
        SQL = RecuperaValor(Cad, 8)
        If SQL <> 0 Then
            IT.SubItems(7) = Format(SQL, "###,###,##0.00")
        Else
            IT.SubItems(7) = " "
        End If
        
        SQL = RecuperaValor(Cad, 9)
        If SQL = "*" Then
            IT.SmallIcon = 42
            IT.ToolTipText = "Devolucion"
        End If
        
        
        'Para la ordenacion
        Cad = IT.Text & Format(IT.SubItems(1), "000000")
        IT.SubItems(8) = Cad
        
        
        
        
        
    Next
    NumRegElim = 0
    
    
    
    
    
ECargarFacturas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
    Set rApu = Nothing
    Set Col = Nothing
    Me.Label30.visible = True
    Me.Label30.Caption = "Remesa " & RecuperaValor(Parametros, 1) & " / " & RecuperaValor(Parametros, 2)
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargarFacturasTransfAbonos()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmppal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView9.ColumnHeaders.Add , , "Vto", 700.2522
    ListView9.ColumnHeaders.Add , , "Fecha Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Nombre", 5000.2522
    ListView9.ColumnHeaders.Add , , "Gastos", 1000.2522, 1
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    Cad = "select cobros.numserie, cobros.numfactu, cobros.fecfactu, cobros.numorden, cobros.fecvenci,cobros.nomclien, cobros.gastos, cobros.impvenci  importe "
    Cad = Cad & " from cobros "
    Cad = Cad & " where (cobros.transfer = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and cobros.anyorem = " & DBSet(RecuperaValor(Parametros, 2), "N") & ") "
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = DBLet(Rs.Fields(3))
        IT.SubItems(4) = DBLet(Rs.Fields(4))
        IT.SubItems(5) = DBLet(Rs.Fields(5))
        
        'gastos
        If DBLet(Rs.Fields(6), "N") <> 0 Then
            IT.SubItems(6) = Format(DBLet(Rs.Fields(6)), "###,###,##0.00")
        Else
            IT.SubItems(6) = " "
        End If
        
        'importe
        If DBLet(Rs.Fields(7), "N") <> 0 Then
            IT.SubItems(7) = Format(DBLet(Rs.Fields(7)), "###,###,##0.00")
        Else
            IT.SubItems(7) = " "
        End If
        
        
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarFacturas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub


Private Sub CargarFacturasTransfPagos(Tipo As Byte)
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmppal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 1350.2522
    ListView9.ColumnHeaders.Add , , "Vto", 700.2522
    ListView9.ColumnHeaders.Add , , "Proveedor", 1500.2522
    ListView9.ColumnHeaders.Add , , "Nombre", 5000.2522
    ListView9.ColumnHeaders.Add , , "Fecha Vto", 1350.2522
    ListView9.ColumnHeaders.Add , , "Importe", 1800.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    Cad = "select pagos.numserie, pagos.numfactu, pagos.fecfactu, pagos.numorden, pagos.codmacta,pagos.nomprove, pagos.fecefect, pagos.impefect  importe "
    Cad = Cad & " from pagos "
    Cad = Cad & " where (pagos.nrodocum = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and pagos.anyodocum = " & DBSet(RecuperaValor(Parametros, 2), "N") & ") "
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = DBLet(Rs.Fields(3))
        IT.SubItems(4) = DBLet(Rs.Fields(4))
        IT.SubItems(5) = DBLet(Rs.Fields(5))
        IT.SubItems(6) = DBLet(Rs.Fields(6))
        
        
        'importe
        If DBLet(Rs.Fields(7), "N") <> 0 Then
            IT.SubItems(7) = Format(DBLet(Rs.Fields(7)), "###,###,##0.00")
        Else
            IT.SubItems(7) = " "
        End If
        
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarFacturas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub

Private Sub CargarBancosRemesas()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarBancosRemesas
    
    Set ListView10.SmallIcons = frmppal.imgListComun16
    
    ListView10.ColumnHeaders.Clear
    ListView10.ListItems.Clear
    
    
    ListView10.ColumnHeaders.Add , , "Banco", 1900.2522
    ListView10.ColumnHeaders.Add , , "Nombre", 3600.2522
    ListView10.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    Cad = "select cta, nomcta, acumperd from tmpcierre1 where codusu = " & vUsu.Codigo
    Cad = Cad & " ORDER BY 1 "
    
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView10.ListItems.Add
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        
        'importe
        If DBLet(Rs.Fields(2), "N") <> 0 Then
            IT.SubItems(2) = Format(DBLet(Rs.Fields(2)), "###,###,##0.00")
        Else
            IT.SubItems(2) = " "
        End If
        
        IT.Checked = True
        
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarBancosRemesas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub


Private Sub CargarRecibosConCobrosParciales()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long
Dim Pagos As Boolean

    On Error GoTo ECargarRecibosConCobrosParciales
    
    Set ListView11.SmallIcons = frmppal.imgListComun16
    
    ListView11.ColumnHeaders.Clear
    ListView11.ListItems.Clear
    
    
    ListView11.ColumnHeaders.Add , , "Serie", 800.2522
    ListView11.ColumnHeaders.Add , , "Factura", 1300.2522
    ListView11.ColumnHeaders.Add , , "Fecha", 1500.2522, 1
    ListView11.ColumnHeaders.Add , , "Vto", 900.2522, 1
    ListView11.ColumnHeaders.Add , , "Importe Vto", 1500.2522, 1
    ListView11.ColumnHeaders.Add , , "Cobrado", 1500.2522, 1
    
    
    Set Rs = New ADODB.Recordset
    
    ' le hemos pasado el select completo de cobros
    Cad = Parametros
    
    Pagos = False
    If InStr(1, UCase(Cad), "PAGOS.") > 0 Then Pagos = True
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView11.ListItems.Add
        
        IT.Text = DBLet(Rs!NUmSerie)
        IT.SubItems(1) = DBLet(Rs!NumFactu)
        IT.SubItems(2) = DBLet(Rs!FecFactu)
        IT.SubItems(3) = DBLet(Rs!numorden)
        
        
        'importe
        IT.SubItems(4) = " "
        If Not Pagos Then
            'COBROS
            If DBLet(Rs!ImpVenci, "N") <> 0 Then IT.SubItems(4) = Format(DBLet(Rs!ImpVenci), "###,###,##0.00")
        Else
            If DBLet(Rs!ImpEfect, "N") <> 0 Then IT.SubItems(4) = Format(DBLet(Rs!ImpEfect), "###,###,##0.00")
        End If
        
        'importe cobrado
        IT.SubItems(5) = " "
        If Not Pagos Then
            'Cobros
            If DBLet(Rs!impcobro, "N") <> 0 Then IT.SubItems(5) = Format(DBLet(Rs!impcobro), "###,###,##0.00")
        Else
            If DBLet(Rs!imppagad, "N") <> 0 Then IT.SubItems(5) = Format(DBLet(Rs!imppagad), "###,###,##0.00")
        End If
        
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarRecibosConCobrosParciales:
    
    MuestraError Err.Number, "Carga Recibos con " & IIf(Pagos, "pagos", "cobros") & " parciales " & Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub




Private Sub CargarTalonesPagaresPendientes()
Dim IT As ListItem
Dim Cad As String
Dim Pendiente As Currency

    On Error GoTo ECargarTalonesPagaresPendientes
    
    Set ListView12.SmallIcons = frmppal.imgListComun16
    
    ListView12.ColumnHeaders.Clear
    ListView12.ListItems.Clear
    
    
    ListView12.ColumnHeaders.Add , , "Serie", 800.2522
    ListView12.ColumnHeaders.Add , , "Factura", 1300.2522
    ListView12.ColumnHeaders.Add , , "Fecha", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Fec.Vto", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Vto", 900.2522, 1
    ListView12.ColumnHeaders.Add , , "Tipo", 900.2522, 1
    
    ListView12.ColumnHeaders.Add , , "Importe", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Gasto", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Cobrado", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Pendiente", 1500.2522, 1
    
    
    Set Rs = New ADODB.Recordset
    
    ' le hemos pasado el select completo de cobros
    Cad = "SELECT cobros.*, formapago.nomforpa, tipofpago.descformapago, tipofpago.siglas, "
    Cad = Cad & " cobros.nomclien nommacta,cobros.codmacta,tipofpago.tipoformapago, 0 aaa "
    Cad = Cad & " FROM ((cobros INNER JOIN formapago ON cobros.codforpa = formapago.codforpa) "
    Cad = Cad & " INNER JOIN tipofpago ON formapago.tipforpa = tipofpago.tipoformapago) "
    
    If Parametros <> "" Then Cad = Cad & " WHERE " & Parametros
    
    Cad = Cad & " Union "
    Cad = Cad & " SELECT cobros.*, formapago.nomforpa, tipofpago.descformapago, tipofpago.siglas, "
    Cad = Cad & " cobros.nomclien nommacta,cobros.codmacta,tipofpago.tipoformapago, 1 aaa "
    Cad = Cad & " FROM ((cobros INNER JOIN formapago ON cobros.codforpa = formapago.codforpa) "
    Cad = Cad & " INNER JOIN tipofpago ON formapago.tipforpa = tipofpago.tipoformapago) "
    Cad = Cad & " where (numserie, numfactu, fecfactu, numorden) in (select numserie, numfactu, fecfactu, numorden from talones_facturas where codigo = " & DBSet(Codigo, "N") & ")"
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView12.ListItems.Add
        
        IT.Text = DBLet(Rs!NUmSerie)
        IT.SubItems(1) = DBLet(Rs!NumFactu)
        IT.SubItems(2) = DBLet(Rs!FecFactu)
        IT.SubItems(3) = DBLet(Rs!FecVenci)
        IT.SubItems(4) = DBLet(Rs!numorden)
        'IT.SubItems(5) = DBLet(Rs!descformapago)
        IT.SubItems(5) = DBLet(Rs!siglas, "T")
        
        
        'importe
        If DBLet(Rs!ImpVenci, "N") <> 0 Then
            IT.SubItems(6) = Format(DBLet(Rs!ImpVenci), "###,###,##0.00")
        Else
            IT.SubItems(6) = " "
        End If
        
        'importe gasto
        If DBLet(Rs!Gastos, "N") <> 0 Then
            IT.SubItems(7) = Format(DBLet(Rs!Gastos), "###,###,##0.00")
        Else
            IT.SubItems(7) = " "
        End If
        
        'importe cobrado
        If DBLet(Rs!impcobro, "N") <> 0 Then
            IT.SubItems(8) = Format(DBLet(Rs!impcobro), "###,###,##0.00")
        Else
            IT.SubItems(8) = " "
        End If
        
        Pendiente = DBLet(Rs!ImpVenci) + DBLet(Rs!Gastos, "N") - DBLet(Rs!impcobro, "N")
        If Pendiente <> 0 Then
            IT.SubItems(9) = Format(Pendiente, "###,###,##0.00")
        Else
            IT.SubItems(9) = " "
        End If
        IT.Tag = 0
        If DBLet(Rs!aaa, "N") = 1 Then
            IT.Checked = True
            IT.Tag = 1
            IT.ToolTipText = "Ya pertenece al documento"
            IT.ForeColor = vbBlue
        End If
        
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarTalonesPagaresPendientes:
    MuestraError Err.Number, "Carga Recibos Talones/Pagarés pendientes" & Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub



Private Function SumaTotales() As Boolean
Dim SQL As String
Dim Suma As Currency
    
    Suma = 0
    txtSuma.Text = ""
    
    For i = 1 To Me.ListView12.ListItems.Count
        If Me.ListView12.ListItems(i).Checked Then
            Suma = Suma + ComprobarCero(Trim(Me.ListView12.ListItems(i).SubItems(9)))
        End If
    Next i
    
    If Suma <> 0 Then txtSuma.Text = Format(Suma, "###,###,##0.00")

End Function


Private Sub CargarFacturasCompensaciones()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmppal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView9.ColumnHeaders.Add , , "Vto", 700.2522
    ListView9.ColumnHeaders.Add , , "Fecha Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Gastos", 1000.2522, 1
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    Cad = "select compensa_facturas.numserie, compensa_facturas.numfactu, compensa_facturas.fecfactu, compensa_facturas.numorden, compensa_facturas.fecvenci, compensa_facturas.gastos, compensa_facturas.impvenci  importe "
    Cad = Cad & " from compensa_facturas "
    Cad = Cad & " where compensa_facturas.codigo = " & DBSet(RecuperaValor(Parametros, 1), "N")
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = DBLet(Rs.Fields(3))
        IT.SubItems(4) = DBLet(Rs.Fields(4))
        
        'gastos
        If DBLet(Rs.Fields(5), "N") <> 0 Then
            IT.SubItems(5) = Format(DBLet(Rs.Fields(5)), "###,###,##0.00")
        Else
            IT.SubItems(5) = " "
        End If
        
        'importe
        If DBLet(Rs.Fields(6), "N") <> 0 Then
            IT.SubItems(6) = Format(DBLet(Rs.Fields(6)), "###,###,##0.00")
        Else
            IT.SubItems(6) = " "
        End If
        
        
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarFacturas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub


Private Sub CargarFacturasCompensacionesPro()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmppal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 1750.2522
    ListView9.ColumnHeaders.Add , , "Vto", 700.2522
    ListView9.ColumnHeaders.Add , , "Fecha Efecto", 1750.2522
    ListView9.ColumnHeaders.Add , , "Importe", 2500.2522, 1
    
    Set Rs = New ADODB.Recordset
    
    Cad = "select compensapro_facturas.numserie, compensapro_facturas.numfactu, compensapro_facturas.fecfactu, compensapro_facturas.numorden, compensapro_facturas.fecefect, compensapro_facturas.impefect  importe "
    Cad = Cad & " from compensapro_facturas "
    Cad = Cad & " where compensapro_facturas.codigo = " & DBSet(RecuperaValor(Parametros, 1), "N")
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
    
    Rs.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(Rs.Fields(0))
        IT.SubItems(1) = DBLet(Rs.Fields(1))
        IT.SubItems(2) = DBLet(Rs.Fields(2))
        IT.SubItems(3) = DBLet(Rs.Fields(3))
        IT.SubItems(4) = DBLet(Rs.Fields(4))
        
        
        'importe
        If DBLet(Rs.Fields(5), "N") <> 0 Then
            IT.SubItems(5) = Format(DBLet(Rs.Fields(5)), "###,###,##0.00")
        Else
            IT.SubItems(5) = " "
        End If
        
        
        Rs.MoveNext
    Wend
    NumRegElim = 0
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ECargarFacturas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set Rs = Nothing
End Sub



Private Sub txtCodmacta_GotFocus(Index As Integer)
    ConseguirFoco txtCodmacta(Index), 3
End Sub

Private Sub txtCodmacta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtCodmacta_LostFocus(Index As Integer)
    SQL = Trim(txtCodmacta(Index).Text)
    Errores = ""
    If CuentaCorrectaUltimoNivel(SQL, Errores) Then
        Me.txtCodmacta(Index).Text = SQL
        Me.txtNommacta(Index).Text = Errores
    Else
        If Me.txtCodmacta(Index).Text <> "" Then MsgBox Errores, vbExclamation
        Me.txtCodmacta(Index).Text = ""
        Me.txtNommacta(Index).Text = ""
        PonFoco txtCodmacta(Index)
    End If
End Sub


Private Sub txtConcepto_GotFocus(Index As Integer)
    ConseguirFoco txtConcepto(Index), 3
End Sub

Private Sub txtConcepto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Or KeyCode = 187 Then
        KeyCode = 0
        txtConcepto(Index).Text = ""
        
        imgConcepto_Click Index
        
    End If
End Sub

Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
    If txtConcepto(Index).Text = "+" Then
        txtConcepto(Index).Text = ""
        Exit Sub
    End If
    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    SQL = ""
    If txtConcepto(Index).Text <> "" Then
        If PonerFormatoEntero(txtConcepto(Index)) Then
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index).Text)
            If SQL = "" Then MsgBox "No existe el concepto", vbExclamation
        End If
    End If
    If SQL = "" Then
        If txtConcepto(Index).Text <> "" Then
            txtConcepto(Index).Text = ""
            PonFoco txtConcepto(Index)
        End If
    End If
    Me.txtConceptoNom(Index).Text = SQL
End Sub

Private Sub txtDiario_GotFocus(Index As Integer)
     ConseguirFoco txtDiario(Index), 3
End Sub

Private Sub txtDiario_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Or KeyCode = 187 Then
        KeyCode = 0
        txtDiario(Index).Text = ""
        
        imgDiario_Click Index
        
    End If
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
   If txtDiario(Index).Text = "+" Then
        txtDiario(Index).Text = ""
        Exit Sub
    End If
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    SQL = ""
    If txtDiario(Index).Text <> "" Then
        If PonerFormatoEntero(txtDiario(Index)) Then
            SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text)
            If SQL = "" Then MsgBox "No existe el diario", vbExclamation
        End If
    End If
    If SQL = "" Then
        If txtDiario(Index).Text <> "" Then
            txtDiario(Index).Text = ""
            PonFoco txtDiario(Index)
        End If
    End If
    Me.txtDiarioNom(Index).Text = SQL
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtfecha(Index), 3
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    '
    If Not PonerFormatoFecha(txtfecha(Index)) Then txtfecha(Index).Text = ""
End Sub

Private Sub txtIbanES_GotFocus(Index As Integer)
    ConseguirFoco txtIbanES(Index), 3
End Sub

Private Sub txtIbanES_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, False
End Sub

Private Sub txtIbanES_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        'Permite letras y numeros
        If txtIbanES(Index).Text <> "" Then
            txtIbanES(Index).Text = UCase(txtIbanES(Index).Text)
            If Mid(txtIbanES(Index).Text, 1, 2) <> "ES" Then
                MsgBox "Comienza por ES", vbExclamation
                txtIbanES(Index).Text = "ES"
                PonFoco txtIbanES(Index)
            End If
            txtIbanES(5).Text = ""
        End If
    Case 1 To 4
        If txtIbanES(Index).Text <> "" Then
            If Not IsNumeric(txtIbanES(Index).Text) Then
                MsgBox "Campo numérico", vbExclamation
                PonFoco txtIbanES(Index)
            Else
                Me.txtIbanES(Index).Text = Right("0000000000" & txtIbanES(Index).Text, IIf(Index = 4, 10, IIf(Index = 3, 2, 4)))
            End If
        End If
        txtIbanES(5).Text = ""
        If Index = 4 Then PonleFoco Me.cmdValidarIBAN(0)
    Case 5
        If Me.optIBAN_Intro(1).Value Then
            For i = 0 To 4
                txtIbanES(i).Text = ""
            Next i
        End If
    End Select
End Sub


Private Sub ValidarIBANEspanaSeparado()

    'Validamos el CC
    SQL = txtIbanES(1).Text & txtIbanES(2).Text & txtIbanES(3).Text & txtIbanES(4).Text
    If Len(SQL) <> 20 Then
        Errores = "Longitud CCC debe ser 20"
        Exit Sub
    End If
    
    SQL = txtIbanES(1).Text & txtIbanES(2).Text & txtIbanES(4).Text
    SQL = CodigoDeControl(SQL)
    If SQL <> txtIbanES(3).Text Then
        Errores = "Digitos de control  calcluado distinto del introducido : " & SQL & " --- " & txtIbanES(3).Text
        Exit Sub
    End If
    
    'Vamos bien. Vamos a por el IBAN
    SQL = txtIbanES(1).Text & txtIbanES(2).Text & txtIbanES(3).Text & txtIbanES(4).Text
    If Not DevuelveIBAN2("ES", SQL, Errores) Then
        Errores = "ERROR"
        Exit Sub
    End If
    If txtIbanES(0).Text = "" Then txtIbanES(0).Text = "ES" & Errores
    If "ES" & Errores <> txtIbanES(0).Text Then
        Errores = "IBAN calculado distinto del introducido : ES" & Errores & " --- " & txtIbanES(0).Text
        Exit Sub
    End If
    
    'LLegados aqui, iban correcto
    SQL = txtIbanES(0).Text & txtIbanES(1).Text & txtIbanES(2).Text & txtIbanES(3).Text & txtIbanES(4).Text
    txtIbanES(5).Text = DevuelveIBANSeparado(SQL)
    Errores = ""
End Sub

Private Sub ValidarIBANEspanaJunto()
    txtIbanES(5).Text = Replace(txtIbanES(5).Text, " ", "")
    SQL = txtIbanES(5).Text
    If Len(SQL) <> 24 Then
        Errores = "Longitud IBAN debe ser 24"
        Exit Sub
    End If
    If Mid(SQL, 1, 2) <> "ES" Then
        Errores = "Debe empezar por ES"
        Exit Sub
    End If
    SQL = Mid(SQL, 3)
    If Not IsNumeric(SQL) Then
        Errores = "Deben ser numeros"
        Exit Sub
    End If
    
    
    'CC
    SQL = Mid(txtIbanES(5).Text, 5, 4) & Mid(txtIbanES(5).Text, 9, 4) & Mid(txtIbanES(5).Text, 15, 10)
    SQL = CodigoDeControl(SQL)
    If SQL <> Mid(txtIbanES(5).Text, 13, 2) Then
        Errores = "Digitos de control  calculado distinto del introducido : " & SQL & " --- " & Mid(txtIbanES(5).Text, 13, 2)
        Exit Sub
    End If

    
    
    SQL = Mid(txtIbanES(5).Text, 5)
    If Not DevuelveIBAN2("ES", SQL, Errores) Then
        Errores = "ERROR"
        Exit Sub
    End If
    If Errores <> Mid(txtIbanES(5).Text, 3, 2) Then
        Errores = "IBAN calculado distinto del introducido : ES" & Errores & " --- " & Mid(txtIbanES(5).Text, 1, 4)
        Exit Sub
    End If
    
    txtIbanES(0).Text = Mid(txtIbanES(5).Text, 1, 4)
    txtIbanES(1).Text = Mid(txtIbanES(5).Text, 5, 4)
    txtIbanES(2).Text = Mid(txtIbanES(5).Text, 9, 4)
    txtIbanES(3).Text = Mid(txtIbanES(5).Text, 13, 2)
    txtIbanES(4).Text = Mid(txtIbanES(5).Text, 15, 10)
    Errores = ""
End Sub


Private Sub ponerDatosRiesgoTalPag_Confir()
Dim TalonPagare As Boolean
    
    TalonPagare = cboRiesgoYelConfirming.ItemData(cboRiesgoYelConfirming.ListIndex) = 1
    SQL = "Confirming vencidos"
    If TalonPagare Then SQL = "Cancelación remesas cuentas puente"
    
    Label52(1).Caption = SQL
    Label52(1).Refresh
    cmdElimRiesgoTalPag(1).visible = TalonPagare
    lblFecha(2).visible = TalonPagare
    ImgFec(2).visible = TalonPagare
    txtfecha(2).visible = TalonPagare
    chkVarios(1).visible = TalonPagare
    If PrimeraVez Then Exit Sub
    ListView13.Checkboxes = TalonPagare
    CargaLW_RemesasRiesgo TalonPagare
End Sub



Private Sub CargaLW_RemesasRiesgo(TalonPagare As Boolean)
Dim IT
    Set Rs = New ADODB.Recordset
    Me.ListView13.ListItems.Clear
    If TalonPagare Then
    
        SQL = QueRemesasMostrarEliminarRiesgoTalPag
        If SQL = "" Then Exit Sub
        SQL = Mid(SQL, 2)
        
        SQL = " IN (" & SQL & ") ORDER BY fecremesa"
        SQL = " from remesas,cuentas where remesas.codmacta=cuentas.codmacta and  (codigo,anyo)" & SQL
        SQL = " if(tiporem=2,'Pagare','Talon'),if(situacion='Y','Elimindos','Abonado'),Descripcion,tiporem " & SQL
        SQL = "select codigo,anyo,fecremesa,situacion,nommacta,importe," & SQL
        
        
    
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Set IT = ListView13.ListItems.Add
            
            IT.Text = DBLet(Rs.Fields(0))
            IT.SubItems(1) = DBLet(Rs.Fields(1))
            IT.SubItems(2) = DBLet(Rs.Fields(2))
            IT.SubItems(3) = DBLet(Rs.Fields(4))
            IT.SubItems(4) = DBLet(Rs.Fields(8))
            IT.SubItems(5) = DBLet(Rs.Fields(6))
            IT.SubItems(6) = Format(DBLet(Rs.Fields(5), "N"), "#,##0.00")
            IT.Tag = Rs!Tiporem
                
        
            Rs.MoveNext
        Wend
        Rs.Close
        
        
    Else
    
        'Veremos si hay riesgo CONFIRMING
        SQL = "select codigo,anyo,fecha,situacion,nommacta,importe,"
        SQL = SQL & " 'Confirming',if(situacion='Y','Elimindos','Abonado'),transferencias.Descripcion,127 subtipo  "
        
        SQL = SQL & " FROM transferencias, bancos,cuentas WHERE transferencias.codmacta=bancos.codmacta "
        SQL = SQL & " AND transferencias.codmacta=cuentas.codmacta and "
        SQL = SQL & " transferencias.tipotrans = 0 and transferencias.subtipo = 2 and situacion>='Q' and situacion <'Z' "
        
        If vParamT.ConfirmingCtaPuente Then
            'Nada
        Else
            SQL = SQL & " and ctaconfirming<>'' "  'Solo si lleva cta confirming
            SQL = SQL & "  AND (codigo,anyo) IN (Select distinct nrodocum,anyodocum  from pagos, bancos  where bancos.codmacta =pagos.ctabanc1"
            SQL = SQL & "  and ctaconfirming<>'' and DATE_ADD(now() , INTERVAL coalesce(confirmingriesgo,0) DAY) >fecefect AND situacion=0 and situdocum>='Q' and situdocum <'Z' )"
            
        End If
        
        
        
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Set IT = ListView13.ListItems.Add
            
            IT.Text = DBLet(Rs.Fields(0))
            IT.SubItems(1) = DBLet(Rs.Fields(1))
            IT.SubItems(2) = DBLet(Rs.Fields(2))
            IT.SubItems(3) = DBLet(Rs.Fields(4))
            IT.SubItems(4) = DBLet(Rs.Fields(8))
            IT.SubItems(5) = DBLet(Rs.Fields(6))
            IT.SubItems(6) = Format(DBLet(Rs.Fields(5), "N"), "#,##0.00")
            IT.Tag = Rs!SubTipo
                
        
            Rs.MoveNext
        Wend
        Rs.Close
        
    End If
        
        
    
    
    
    
    
    
    
    
    Set Rs = Nothing
    
    
End Sub

Private Sub ValoresTransferencia()
    Set Rs = New ADODB.Recordset
    
    SQL = IIf(RecuperaValor(Parametros, 3) = "1", "0", "1") '1. Es el 0 en la BD y cero es el 1 en la BD . NOOOOOOOOOOOOOOO!!!!
    SQL = " AND tipotrans = " & SQL
    SQL = " AND anyo = " & DBSet(RecuperaValor(Parametros, 2), "N") & SQL
    SQL = " where codigo = " & DBSet(RecuperaValor(Parametros, 1), "N") & SQL
    
    SQL = "Select * from transferencias" & SQL
    
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        Me.cmdGeneraTransfer(0).Enabled = False
        txtfecha(3).Text = ""
    Else
        'fecha,LlevaAgrupados
        txtfecha(3).Text = Format(Rs!Fecha, "dd/mm/yyyy")
        chkVarios(2).Value = DBLet(Rs!LlevaAgrupados, "N")
    End If
    Rs.Close
    
    Set Rs = Nothing
End Sub




Private Sub CargaResultadoCuentasIncrementar()
    On Error GoTo eCargaResultadoCuentasIncrementar
    
    i = RecuperaValor(Parametros, 1)    'Cuantos ceros
    J = RecuperaValor(Parametros, 2) - 1  'Posicion es donde va. Luego qitaremos lo de antes, y pondremos los ceros
    Errores = String(i, "0")
    SQL = "Select codmacta,nommacta from cuentas WHERE apudirec='S' and codmacta like '" & Replace(Me.cboAumento.Text, "*", "%")
    SQL = SQL & "'  "
    ListView2(2).ListItems.Clear
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not Rs.EOF
        NumRegElim = NumRegElim + 1
        ListView2(2).ListItems.Add , , CStr(Rs!codmacta)
        'Update en codmacta
        SQL = Mid(Rs!codmacta, 1, J) & Errores & Mid(Rs!codmacta, J + 1)
        
        ListView2(2).ListItems(NumRegElim).SubItems(1) = SQL
        
        ListView2(2).ListItems(NumRegElim).SubItems(2) = Rs!Nommacta
        Rs.MoveNext
    Wend
    Rs.Close
    
eCargaResultadoCuentasIncrementar:
    If Err.Number <> 0 Then Err.Clear
    Set Rs = Nothing
End Sub
