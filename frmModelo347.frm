VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModelo347 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMasivo 
      Caption         =   "Enviar"
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
      Left            =   8760
      TabIndex        =   46
      Top             =   6210
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame FrameEnvioMasivo 
      Height          =   6015
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CheckBox chkAmazonCopia 
         Caption         =   "Copia destinatario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   48
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5175
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NIF"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6244
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "eMail"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Importe"
            Object.Width           =   2893
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   3480
         Picture         =   "frmModelo347.frx":0000
         ToolTipText     =   "Puntear al Debe"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   3120
         Picture         =   "frmModelo347.frx":014A
         ToolTipText     =   "Quitar al Debe"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Máximo numero de mensajes 40"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   7680
         TabIndex        =   44
         Top             =   360
         Width           =   3420
      End
      Begin VB.Label Label3 
         Caption         =   "Envios email disponibles"
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
         Height          =   240
         Index           =   6
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   2670
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
      Height          =   3285
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   6915
      Begin VB.TextBox Text347 
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
         Left            =   1560
         TabIndex        =   40
         Top             =   2400
         Width           =   1425
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   810
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "N.I.F"
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
         Index           =   2
         Left            =   480
         TabIndex        =   39
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Factura"
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
         Left            =   480
         TabIndex        =   27
         Top             =   510
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
         Left            =   480
         TabIndex        =   26
         Top             =   870
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
         Left            =   480
         TabIndex        =   25
         Top             =   1230
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1170
         Picture         =   "frmModelo347.frx":0294
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1170
         Picture         =   "frmModelo347.frx":031F
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label lblCuentas 
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
         Left            =   2520
         TabIndex        =   21
         Top             =   5190
         Width           =   4095
      End
      Begin VB.Label lblCuentas 
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
         Left            =   2520
         TabIndex        =   20
         Top             =   4800
         Width           =   4095
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
      Height          =   6015
      Left            =   7050
      TabIndex        =   22
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox chkMasivo 
         Caption         =   "Email masivo"
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
         Left            =   1920
         TabIndex        =   47
         Top             =   4320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame FrameSeccion 
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   180
         TabIndex        =   35
         Top             =   1170
         Width           =   4215
         Begin MSComctlLib.ListView ListView1 
            Height          =   2085
            Index           =   1
            Left            =   60
            TabIndex        =   36
            Top             =   360
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   3678
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
            Index           =   0
            Left            =   3360
            Picture         =   "frmModelo347.frx":03AA
            ToolTipText     =   "Quitar al Debe"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   3750
            Picture         =   "frmModelo347.frx":04F4
            ToolTipText     =   "Puntear al Debe"
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Empresas"
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
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "frmModelo347.frx":063E
         Left            =   1920
         List            =   "frmModelo347.frx":0648
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3840
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         TabIndex        =   29
         Top             =   5160
         Width           =   4245
         Begin VB.OptionButton OptProv 
            Caption         =   "Fecha recepción"
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
            TabIndex        =   32
            Top             =   360
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton OptProv 
            Caption         =   "Fecha factura"
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
            Left            =   2370
            TabIndex        =   31
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedores"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.TextBox Text347 
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
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   4680
         Width           =   1275
      End
      Begin VB.TextBox Text347 
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
         Left            =   210
         TabIndex        =   2
         Top             =   600
         Width           =   4065
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3840
         TabIndex        =   23
         Top             =   150
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
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   3720
         TabIndex        =   38
         Top             =   3840
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda carta"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Informe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   3960
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Importe Límite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   4680
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Responsable"
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
         Left            =   210
         TabIndex        =   24
         Top             =   360
         Width           =   1260
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
      TabIndex        =   7
      Top             =   6210
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
      TabIndex        =   5
      Top             =   6210
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
      Left            =   90
      TabIndex        =   6
      Top             =   6210
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
      Left            =   60
      TabIndex        =   8
      Top             =   3360
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
         TabIndex        =   19
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   18
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "A.E.A.T."
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   60
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Label24"
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
      Index           =   30
      Left            =   1710
      TabIndex        =   34
      Top             =   6120
      Width           =   6585
   End
   Begin VB.Label Label2 
      Caption         =   "Label24"
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
      Index           =   31
      Left            =   1710
      TabIndex        =   33
      Top             =   6420
      Width           =   6825
   End
End
Attribute VB_Name = "frmModelo347"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 410


'Febrero 2021
' Hemos creado un arimailCDO
'       Envia a traves de Aamazon WEB Service
'       Para ello, si esta el programa en la carpeta haremos el envio por el programa.
'       En caso contrario seguira haciendolo via OUTLOOK
'       OUTLOOK.   Como estaba hasta ahora
'       arimailCDO Graba en la tabla wenvioemail
'           Una vez grabaado los registros, llamamos al exe
'               arimailCDO odbc|codus|
Private TieneArimaiCDO As Boolean
Private EslaPrimeraEmpresa As Boolean  'Cuando SOLO hay una empresa, no hace falta que comprueba si existe ya el NIF de una empresa anterior

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


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmCar As frmCartas
Attribute frmCar.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim Rs As Recordset

Dim I As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim Tablas As String

Dim Importe As Currency

Dim UltimoPeriodoLiquidacion As Boolean
Dim C2 As String



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
Dim B As Boolean
Dim B2 As Boolean

Dim Rs As ADODB.Recordset

Dim indRPT As String
Dim nomDocu As String


    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    Screen.MousePointer = vbHourglass
    
    'Modificacion de 26 Marzo 2007
    '------------------------------------
    'Hay una tabla auxiliar donde se guardan datos externos de 347.
    'Cuando voy a imprimir los datos pedire si de una y/o de la otra
    
    Sql = "DELETE FROM tmp347tot where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    Sql = "DELETE FROM tmp347trimestral where codusu = " & vUsu.Codigo
    Conn.Execute Sql
        
    
    Set miRsAux = New ADODB.Recordset
    
    'El de siempre
    B = ComprobarCuentas347
    Label2(30).Caption = ""
    Label2(31).Caption = ""
    If Not B Then Exit Sub
    
    
    'Cobros efectivo
    'Updatearemos a cero los metalicos que no llegen al minimo
    Set miRsAux = New ADODB.Recordset
    Sql = "Select ImporteMaxEfec340 from parametros "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = DBLet(miRsAux!ImporteMaxEfec340, "N")
    miRsAux.Close
    If Val(Sql) > 0 Then
        Sql = "UPDATE tmp347trimestral set metalico=0  WHERE codusu = " & vUsu.Codigo & " AND metalico < " & TransformaComasPuntos(Sql)
         Conn.Execute Sql
    End If
     
     
    'Ahora borramos todas las entrdas k no superan el importe limite
    Label2(31).Caption = "Comprobar importes"
    Label2(31).Refresh
    Importe = ImporteFormateado(Text347(1).Text)
    Sql = "Delete from tmp347tot where codusu = " & vUsu.Codigo & " AND Importe  <" & TransformaComasPuntos(CStr(Importe))
    Conn.Execute Sql
    
    
    'Comprobare si hay datos
    'Comprobamos si hay datos
    Sql = "Select count(*) FROM tmp347tot where codusu = " & vUsu.Codigo
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            CONT = miRsAux.Fields(0)
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    Screen.MousePointer = vbDefault
    Label2(31).Caption = ""
    Label2(30).Caption = ""
    If CONT = 0 Then
        MsgBox "No se ha devuelto ningun dato", vbExclamation
        Exit Sub
    End If
    
    'Precomprobacion de NIFs
    If Not ComprobarNifs347 Then Exit Sub
    
    
    Label2(31).Caption = ""
    Label2(30).Caption = ""
    DoEvent2
    Screen.MousePointer = vbDefault
    

    
    If B Then
        If optTipoSal(1).Value Then
            'Si es impresion y el numero de registros es superior a 25 no
            'puede imprimirse
            CONT = 0
            Sql = ""
                

            'Modelo de haciend a
            B2 = Modelo347(Year(CDate(txtFecha(1).Text)))
            
            If B2 Then
                'CopiarFicheroASalida False, txtTipoSalida(1).Text
                CopiarFicheroHaciend2 True
            End If
        
        Else
            If optTipoSal(2).Value Or optTipoSal(3).Value Then
                ExportarPDF = True 'generaremos el pdf
            Else
                ExportarPDF = False
            End If
            SoloImprimir = False
            If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
            Select Case Combo5.ListIndex
            Case 0
                'La carta
                cad = "¿ Desea imprimir también los proveedores ?"
                If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then
               
                    
                    If ExportarPDF And Me.chkMasivo.Value = 1 Then
                        'Hay que borrar de temp los asc 49
                        cad = "DELETE from tmp347tot WHERE codusu =" & vUsu.Codigo & " AND cliprov= " & Asc(1)
                        Conn.Execute cad
                        
                    Else
                        cad = " {tmp347tot.cliprov} = " & Asc(0) & " AND "
                        cadFormula = cadFormula & cad
                    End If
                    
                End If
                cad = ""
                cadFormula = cadFormula & "{tmp347tot.codusu} = " & vUsu.Codigo
                cadFormula = cadFormula & " and {cartas.codcarta} = 999 "
                
                cadParam = cadParam & "Responsable=""" & Me.Text347(0).Text & """|"
                numParam = numParam + 1
                indRPT = "0410-01"
                
                If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
                
                cadNomRPT = nomDocu '"Carta.rpt"
                
                If ExportarPDF And Me.chkMasivo.Value Then
                    'No genero el rpt
                    
                Else
                    ImprimeGeneral
                End If
            Case Else
            
                'LISTADO
                '-----------------------------------------------------------------
                cadFormula = ""
                If Me.Text347(2).Text <> "" Then cadFormula = "NIF: " & Text347(2).Text & "       "
                cadParam = cadParam & "Fechas= """ & cadFormula & "Desde " & txtFecha(0).Text & "      hasta  " & txtFecha(1).Text & """|"
                numParam = numParam + 1
                    
            
                cadFormula = "{tmp347tot.codusu} = " & vUsu.Codigo

                indRPT = "0410-00"
                
                If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
                
                cadNomRPT = nomDocu '"Carta.rpt"
                
                ImprimeGeneral
            
                
            End Select
            
            If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
            If optTipoSal(3).Value Then EnviarPorEmail  ' LanzaProgramaAbrirOutlook 21
                
            If Not optTipoSal(3).Value Then
                If SoloImprimir Or ExportarPDF Then Unload Me
            End If
            Screen.MousePointer = vbDefault
            
            
            
        End If
    End If
    
    
    
    
    
    
End Sub

Private Sub cmdCancelar_Click()
    If FrameEnvioMasivo.visible Then
        PonerFrameEnvioMasivo False
        
    Else
        Unload Me
    End If
End Sub



Private Sub cmdMasivo_Click()
Dim MaximoNumeroEmails As Integer
Dim Archivo As String

    
    MaximoNumeroEmails = 40
    If TieneArimaiCDO Then MaximoNumeroEmails = 10000

    'Varias cosas
    ' Primero ver que hay seleccionados y no exceden de 40
    Sql = ""
    For I = 1 To Me.ListView2.ListItems.Count
        If ListView2.ListItems(I).Checked Then Sql = Sql & "X"
    Next I
    
    If Len(Sql) = 0 Then
        Sql = "Seleccione algun envio"
    Else
        If Len(Sql) > MaximoNumeroEmails Then
            Sql = "Por problemas de reglas de envio no debe enviar los " & MaximoNumeroEmails & " emails al mismo tiempo"
        Else
            IndCodigo = Len(Sql)
            Sql = ""
        End If
    End If
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Va a enviar " & IndCodigo & " correo" & IIf(IndCodigo > 1, "s", "") & " por email.  ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    If TieneArimaiCDO Then
        If Dir(App.Path & "\flagEnv.txt", vbArchive) <> "" Then
            MsgBox "Imposible lanzar programa envio." & vbCrLf & "Ultimo envio no finalizado (FLAGenv)", vbExclamation
            Exit Sub
        End If
        
    End If
    
    
    
    
    
    
    'Los primero. Generar TODOS los pdfs
    InicializarVbles True
    ExportarPDF = True
    If Not PonerParamRPT("0410-01", cadNomRPT) Then Exit Sub
        
    
    cadParam = cadParam & "Responsable=""" & Me.Text347(0).Text & """|"
    numParam = numParam + 1
    
   
    
    
    If Not EliminaDocum Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Label2(31).Caption = "Prepara carpeta temporal"
    Label2(31).Refresh
    C2 = "" 'Para saber si ha generado alguno
    If PreparaCarpeta Then
            
        
        
        
        'De atras a delante
        For I = Me.ListView2.ListItems.Count To 1 Step -1
            If ListView2.ListItems(I).Checked Then
            
                
                cadFormula = ListView2.ListItems(I).Tag
                cadFormula = cadFormula & " and {cartas.codcarta} = 999 "
                
                Label2(31).Caption = "Generando PDF " & ListView2.ListItems(I).Text
                Label2(31).Refresh
                ImprimeGeneral
                
                K = 0
                Do
                    K = K + 1
                    cadFormula = Dir(App.Path & "\docum.pdf", vbArchive)
                    If cadFormula <> "" Then
                        K = 100
                    Else
                        espera 0.5
                    End If
                Loop Until K > 3
                
                'Si esta generado el docum.pdf lo copiamos
                If K = 100 Then
                    If Not CopiaEnTemporal Then
                        Sql = "Error copiando en carpeta temporal     "
                                
                    Else
                                
                                
                                
                        'Ok ya esta generado el RPT. AHora updateo en la BD dirdatos y pongo el email
                        cad = ListView2.ListItems(I).Tag
                        cad = Replace(cad, "{", "")
                        cad = Replace(cad, "}", "")
                        
                        'El nombre lleva \  con lo cual, hay que "escaparlos para MYSQL
                        Sql = Replace(Sql, "\", "?")
                        Sql = Replace(Sql, "?", "\\")
                        
                        Sql = " SET despobla = " & DBSet(Sql, "T") & " ,dirdatos = " & DBSet(ListView2.ListItems(I).SubItems(2), "T") & ", PAIS='XXXXXXXX' "
                        cad = "UPDATE tmp347tot " & Sql & " WHERE " & cad
                        Conn.Execute cad
                    
                        'QUito el item
                        ListView2.ListItems.Remove ListView2.ListItems(I).Index
                        C2 = "OK"
                    End If
                Else
                    Sql = "Error generando PDF"
                End If
                
                If K < 100 Then
                    ListView2.ListItems(I).Bold = True
                    ListView2.ListItems(I).ForeColor = vbRed
                    ListView2.ListItems(I).ToolTipText = Sql
                End If
            End If
        Next I
    
    
    
    End If
    
    If C2 = "OK" Then
        'Alguno se ha generado
        'XXXXXXXX
            
        Sql = "Recuerde: En el archivo adjunto le enviamos información de su interés."
        Dim CopiaRemite As Boolean
        CopiaRemite = Me.chkAmazonCopia.Value = 1
        CopiaRemite = False
        LanzaProgramaAbrirOutlookMasivo 2, Sql, TieneArimaiCDO, CopiaRemite, Label2(31)
    
    End If
    Screen.MousePointer = vbDefault
    Label2(31).Caption = ""
    Label2(31).Refresh
End Sub

Private Function PreparaCarpeta() As Boolean

    PreparaCarpeta = False
        
    If Dir(vParam.PathFicherosInteg & "\*.pdf", vbArchive) <> "" Then
        Kill vParam.PathFicherosInteg & "\*.pdf"
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Preparando carpeta temporal"
        Else
            PreparaCarpeta = True
        End If

    Else
        PreparaCarpeta = True
    End If
End Function

Private Function EliminaDocum() As Boolean

    EliminaDocum = False
    If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then
        Kill App.Path & "\docum.pdf"
        If Err.Number <> 0 Then
            MuestraError Err.Number, "Eliminando docum temporal"
        Else
            EliminaDocum = True
        End If
    Else
        EliminaDocum = True
    End If
End Function


'Copia EN temporal
Private Function CopiaEnTemporal() As Boolean
Dim C As String
    CopiaEnTemporal = False
    
    If InStr(ListView2.ListItems(I).Tag, "=48") Then
        Sql = ""
    Else
        Sql = "_"
    End If
    Sql = ListView2.ListItems(I).Text & Sql & ".pdf"
    
    If vParam.PathFicherosInteg = "" Then
        C = App.Path & "\temp\" & Sql
    Else
        C = vParam.PathFicherosInteg & "\" & Sql
    End If
    FileCopy App.Path & "\docum.pdf", C
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminando docum temporal"
    Else
        CopiaEnTemporal = True
    End If
        
End Function





Private Sub Combo5_Click()
    VisibleMasivo
End Sub

Private Sub Combo5_Validate(Cancel As Boolean)
   ' optTipoSal(3).Enabled = (Combo5.ListIndex = 1)
   ' If Not optTipoSal(3).Enabled Then optTipoSal(0).Value = True
    
    Me.Toolbar1.Buttons(1).Enabled = (Combo5.ListIndex = 0)
    
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
    Me.Caption = "Modelo 347"

    ' boton al mto de cartas
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 4
    End With
     
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
     
    txtFecha(0).Text = "01/01/" & Year(vParam.fechaini)
    txtFecha(1).Text = "31/12/" & Year(vParam.fechaini)
    Text347(1).Text = Format(vParam.limimpcl, FormatoImporte)
    Text347(0).Text = DevuelveDesdeBD("responsable", "paramtesor", "1", "1")
    Label2(30).Caption = ""
    Label2(31).Caption = ""
    
    Combo5.ListIndex = 1
    Toolbar1.Buttons(1).Enabled = False
     
    CargarListView 1
    
    FrameSeccion.Enabled = vParam.EsMultiseccion
    
    optTipoSal(3).Enabled = (Combo5.ListIndex = 1)
    
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    PonerDatosFicheroSalida
    
    
    TieneArimaiCDO = False
    Sql = DevuelveDesdeBD("emailAWS", "parametros", "1", "1")
    If Sql <> "" Then
        'If Dir(App.Path & "\CreamailCDO.exe", vbArchive) <> "" Then
           TieneArimaiCDO = True
            Label3(7).Caption = "Envio por Amazon Web Service"
        'End If
        chkAmazonCopia.visible = False
    End If
    
End Sub

Private Sub PonerDatosFicheroSalida()
    
    txtTipoSalida(1).Text = App.Path & "\Exportar\Mod347.txt"

End Sub


Private Sub PonerFrameEnvioMasivo(visible As Boolean)
    
    cmdMasivo.visible = visible
    cmdAccion(0).visible = Not visible
    cmdAccion(1).visible = Not visible
    If visible Then
        Me.FrameEnvioMasivo.top = 0
        FrameEnvioMasivo.Left = 60
    Else
        ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 3
    End If
    FrameEnvioMasivo.visible = visible
    
    
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
Dim I As Integer
Dim B As Boolean
    Screen.MousePointer = vbHourglass
    
    If Index < 2 Then
    
            For I = 1 To ListView1(1).ListItems.Count
                ListView1(1).ListItems(I).Checked = Index = 1
            Next I
    Else
            For I = 1 To ListView2.ListItems.Count
                B = False
                If Index = 3 Then
                    If Trim(ListView2.ListItems(I).SubItems(2)) <> "" Then B = True
                End If
                Me.ListView2.ListItems(I).Checked = B
            Next I
    End If
    
    Screen.MousePointer = vbDefault


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


Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ListView2.SortKey = ColumnHeader.Index - 1 Then
        'misma columna
        If ListView2.SortOrder = lvwAscending Then
            ListView2.SortOrder = lvwDescending
        Else
            ListView2.SortOrder = lvwAscending
        End If
    Else
        ListView2.SortKey = ColumnHeader.Index - 1
        ListView2.SortOrder = lvwAscending
    End If
End Sub

Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Trim(Item.SubItems(2)) = "" Then Item.Checked = False
    
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
    
    VisibleMasivo
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Set frmCar = New frmCartas
            
            frmCar.CodigoActual = 999
            frmCar.Desde347 = True
            frmCar.Show vbModal
    
            Set frmCar = Nothing
    
    End Select

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtPag2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
End Sub


Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    Sql = ""
    
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0410-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "FacturasCliFecha.rpt"

    cadParam = cadParam & "pFecha=""" & txtFecha(2).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "Empresas= """
    For I = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            cadParam = cadParam & Me.ListView1(1).ListItems(I).SubItems(1) & "  "
        End If
    Next I
    cadParam = Trim(cadParam)
    
    cadParam = cadParam & """|"
    
    
    cadFormula = "{tmp347.codusu}=" & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 21
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Function CargarTemporal() As Boolean
Dim Sql As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    Sql = "delete from tmpfaclin where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "insert into tmpfaclin (codusu, codigo, numserie, nomserie, numfac, fecha, cta, cliente, nif, imponible, impiva, total, retencion,"
    Sql = Sql & " recargo, tipoopera, tipoformapago) "
    Sql = Sql & " select distinct " & vUsu.Codigo & ",0, factcli.numserie, contadores.nomregis, factcli.numfactu, factcli.fecfactu, factcli.codmacta, "
    Sql = Sql & " factcli.nommacta, factcli.nifdatos, factcli.totbases, factcli.totivas, factcli.totfaccl, factcli.trefaccl, "
    Sql = Sql & " factcli.totrecargo, tipofpago.descformapago , aa.denominacion"
    Sql = Sql & " from " & tabla
    Sql = Sql & " where " & cadselect
    
    Conn.Execute Sql
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Resumen", Err.Description
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
    
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        MsgBox "Introduce las fechas de consulta.", vbExclamation
        Exit Function
    End If

    If Not ComprobarFechas(0, 1) Then Exit Function
    
    
    If Year(CDate(txtFecha(0).Text)) <> Year(CDate(txtFecha(1).Text)) Then
        MsgBox "Esta abarcando dos años. Se considera el año: " & Year(CDate(txtFecha(1).Text)), vbExclamation
    End If
    If Combo5.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de informe.", vbExclamation
        Exit Function
    End If
    
    
    If Combo5.ListIndex = 0 And Text347(0).Text = "" Then
        MsgBox "Escriba el nombre del responsable.", vbExclamation
        Exit Function
    End If
            
    
    If Combo5.ListIndex = 2 Then 'antes 3
        'Enero 2012
        'Tiene que ser una año exacto
        If Month(CDate(txtFecha(0).Text)) <> 1 Or Month(CDate(txtFecha(0).Text)) <> 1 Then
            MsgBox "Año natural. Enero diciembre", vbExclamation
            Exit Function
        End If
        If Month(CDate(txtFecha(1).Text)) <> 12 Or Day(CDate(txtFecha(1).Text)) <> 31 Then
            MsgBox "Año natural. Hasta 31 diciembre", vbExclamation
            Exit Function
        End If
        
    End If
    
    If Text347(1).Text = "" Then
        MsgBox "Importe limite en blanco", vbExclamation
        Exit Function
    End If
    
    If optTipoSal(1).Value And Text347(2).Text <> "" Then
        MsgBox "No puede indicar un NIF generaando el modelo de la AEAT", vbExclamation
        Exit Function
    End If
    
    If EmpresasSeleccionadas = 0 Then
        MsgBox "Seleccione una empresa", vbExclamation
        Exit Function
    End If
    
    
    '++ comprobamos que todas las facturas tienen nif asignado
    DatosOK = ComprobarNifFacturas
    
    If DatosOK Then DatosOK = ComprobarCPostalFacturas
    
       
End Function


Private Function ComprobarNifFacturas() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim CadResul As String

    ComprobarNifFacturas = False

    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            
            'facturas de clientes
            Sql = "select distinct factcli.codmacta from ariconta" & Me.ListView1(1).ListItems(I).Text & ".factcli, ariconta" & Me.ListView1(1).ListItems(I).Text & ".cuentas where "
            Sql = Sql & " cuentas.codmacta=factcli.codmacta and model347=1 "
            Sql = Sql & " AND fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            Sql = Sql & " AND fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            Sql = Sql & " and (factcli.nifdatos is null or factcli.nifdatos = '')"
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            CadResul = ""
            
            While Not Rs.EOF
                CadResul = CadResul & DBLet(Rs!codmacta, "T") & ","
                Rs.MoveNext
            Wend
            
            If CadResul <> "" Then
                CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
                CadResul = Me.ListView1(1).ListItems(I).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de cliente sin nif de las cuentas: " & vbCrLf & vbCrLf & CadResul
                                
                MsgBox CadResul, vbExclamation
                
                Set Rs = Nothing
                Exit Function
            End If
            Set Rs = Nothing
        
            If OptProv(0).Value Then
                cad = "fecharec"
            Else
                If OptProv(1).Value Then
                    cad = "fecfactu"
                End If
            End If
            
            ' facturas de proveedores
            Sql = "SELECT distinct factpro.codmacta from ariconta" & Me.ListView1(1).ListItems(I).Text & ".factpro, ariconta" & Me.ListView1(1).ListItems(I).Text & ".cuentas  where "
            Sql = Sql & " cuentas.codmacta=factpro.codmacta and model347=1 "
            Sql = Sql & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            Sql = Sql & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            Sql = Sql & " and (factpro.nifdatos is null or factpro.nifdatos = '')"
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs.EOF
                CadResul = CadResul & DBLet(Rs!codmacta, "T") & ","
                Rs.MoveNext
            Wend
            If CadResul <> "" Then
                CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
                CadResul = Me.ListView1(1).ListItems(I).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de proveedor sin nif de las cuentas: " & vbCrLf & vbCrLf & CadResul
                                
                MsgBox CadResul, vbExclamation
'
'                Set Rs = Nothing
'                Exit Function
            End If
            Set Rs = Nothing
        
        
        End If
    Next I
    
    ComprobarNifFacturas = True
  
    

End Function


Private Function ComprobarCPostalFacturas() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim CadResul As String

    ComprobarCPostalFacturas = False

    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            
            'facturas de clientes
'            SQL = "select distinct factcli.codmacta from ariconta" & Me.ListView1(1).ListItems(i).Text & ".factcli, ariconta" & Me.ListView1(1).ListItems(i).Text & ".cuentas where "
'            SQL = SQL & " cuentas.codmacta=factcli.codmacta and model347=1 "
'            SQL = SQL & " AND fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
'            SQL = SQL & " AND fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
'            SQL = SQL & " and (factcli.codpobla is null or factcli.codpobla = '')"
            
            'Set Rs = New ADODB.Recordset
            'Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
           ' CadResul = ""
           '
           ' While Not Rs.EOF
           '     CadResul = CadResul & DBLet(Rs!codmacta, "T") & ","
           '     Rs.MoveNext
           ' Wend
            
            'If CadResul <> "" Then
            '    CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
            '    CadResul = Me.ListView1(1).ListItems(i).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de cliente sin código postal de las cuentas: " & vbCrLf & vbCrLf & CadResul
            '
            '    MsgBox CadResul, vbExclamation
            '
            '    Set Rs = Nothing
            '    If MsgBox("¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
            'End If
            'Set Rs = Nothing
        
            If OptProv(0).Value Then
                cad = "fecharec"
            Else
                If OptProv(1).Value Then
                    cad = "fecfactu"
                End If
            End If
            
            ' facturas de proveedores
            'SQL = "SELECT distinct factpro.codmacta from ariconta" & Me.ListView1(1).ListItems(i).Text & ".factpro, ariconta" & Me.ListView1(1).ListItems(i).Text & ".cuentas  where "
            'SQL = SQL & " cuentas.codmacta=factpro.codmacta and model347=1 "
            'SQL = SQL & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            'SQL = SQL & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            'SQL = SQL & " and (factpro.codpobla is null or factpro.codpobla = '')"
           '
           ' Set Rs = New ADODB.Recordset
           ' Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
           '
           ' While Not Rs.EOF
           '     CadResul = CadResul & DBLet(Rs!codmacta, "T") & ","
           '     Rs.MoveNext
           ' Wend
           ' If CadResul <> "" Then
           '     CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
           '     CadResul = Me.ListView1(1).ListItems(i).SubItems(1) & vbCrLf & vbCrLf & "Hay facturas de proveedor sin código postal de las cuentas: " & vbCrLf & vbCrLf & CadResul
           '
           '     MsgBox CadResul, vbExclamation
           '
           '     Set Rs = Nothing
           '     If MsgBox("¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
           ' End If
           ' Set Rs = Nothing
       '
        
        End If
    Next I
    
    ComprobarCPostalFacturas = True
  
    

End Function





Private Function EmpresasSeleccionadas() As Integer
Dim Sql As String
Dim I As Integer
Dim NSel As Integer

    NSel = 0
    For I = 1 To ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then NSel = NSel + 1
    Next I
    EmpresasSeleccionadas = NSel

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

    ListView1(Index).ColumnHeaders.Add , , "Código", 600
    ListView1(Index).ColumnHeaders.Add , , "Descripción", 3200
    
    Sql = "SELECT codempre, nomempre, conta "
    Sql = Sql & " FROM usuarios.empresasariconta "
    
    If Not vParam.EsMultiseccion Then
        Sql = Sql & " where conta = " & DBSet(Conn.DefaultDatabase, "T")
    Else
        Sql = Sql & " where mid(conta,1,8) = 'ariconta'"
    End If
    Sql = Sql & " ORDER BY codempre "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        If vParam.EsMultiseccion Then
            If EsMultiseccion(DBLet(Rs!CONTA)) Then
                Set ItmX = ListView1(Index).ListItems.Add
                
                If DBLet(Rs!CONTA) = Conn.DefaultDatabase Then ItmX.Checked = True
                ItmX.Text = Rs.Fields(0).Value
                ItmX.SubItems(1) = Rs.Fields(1).Value
            End If
        Else
            Set ItmX = ListView1(Index).ListItems.Add
            
            ItmX.Checked = True
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Rs.Fields(1).Value
        End If
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Empresas.", Err.Description
    End If
End Sub

Private Function ComprobarCuentas347() As Boolean
Dim I As Integer
Dim I1 As Currency
Dim I2 As Currency
Dim i3 As Currency
Dim i4 As Currency
Dim I5 As Currency
Dim PAIS As String
Dim cadena2 As String
Dim SoloUnaEmpresa347 As Boolean
Dim EsLaPrimera As Byte '0 empieza proc   1- Esta es la primera     2.Hay mas de una
    
    ComprobarCuentas347 = False
    
    'Esto sera para las inserciones de despues
    Tablas = "INSERT INTO tmp347tot (codusu, cliprov, nif, importe, razosoci, dirdatos, codposta, despobla,Provincia,pais) "
    
        
    Sql = ""
    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then Sql = Sql & "X"
    Next
    SoloUnaEmpresa347 = Len(Sql) = 1  'Para no hacer inserts que no sirvan ,y retrasen el proceso
    cadena2 = ""
    EslaPrimeraEmpresa = False
    EsLaPrimera = 0
    
    For I = 1 To Me.ListView1(1).ListItems.Count
        If Me.ListView1(1).ListItems(I).Checked Then
            If EsLaPrimera = 0 Then
                EsLaPrimera = 1
                EslaPrimeraEmpresa = True
            Else
                EsLaPrimera = 2
                EslaPrimeraEmpresa = False
            End If
                
        
            Label2(30).Caption = Me.ListView1(1).ListItems(I).SubItems(1)
            Label2(31).Caption = "Comprobar Cuentas"
            Me.Refresh
            If Not ComprobarCuentas347_DOS("ariconta" & Me.ListView1(1).ListItems(I).Text, Me.ListView1(1).ListItems(I).SubItems(1), SoloUnaEmpresa347) Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            
        
           'Iremos NIF POR NIF
           
              Label2(31).Caption = "Insertando datos tmp(I)"
              Label2(31).Refresh
              Sql = "SELECT  cliprov,nif, sum(importe) as suma, razosoci,dirdatos,codposta,"
              Sql = Sql & "despobla,provincia,pais from ariconta" & Me.ListView1(1).ListItems(I).Text & ".tmp347 where codusu=" & vUsu.Codigo
              Sql = Sql & " group by cliprov, nif"
              
              Set Rs = New ADODB.Recordset
              Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
              
              While Not Rs.EOF

              
                   Label2(31).Caption = Rs!NIF & " (Ins.)"
                   Label2(31).Refresh
                   If ExisteEntrada Then
                        Importe = Importe + Rs!Suma
                        'SQL = "UPDATE tmp347tot SET importe=importe + " & TransformaComasPuntos(CStr(Rs!Suma))
                        Sql = "UPDATE tmp347tot SET importe= " & TransformaComasPuntos(CStr(Importe))
                        Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & Rs!cliprov
                        Sql = Sql & " AND nif = '" & Rs!NIF & "';"
                        Conn.Execute Sql
                   Else
                        'Nuevo para lo de las agencias de viajes
                        Sql = "," & Rs!cliprov & ",'" & Rs!NIF & "'," & TransformaComasPuntos(CStr(Rs!Suma))
                        Sql = Sql & ",'" & DevNombreSQL(DBLet(Rs!razosoci, "T")) & "','" & DevNombreSQL(DBLet(Rs!dirdatos)) & "','" & DBLet(Rs!codposta, "T") & "','"
                        Sql = Sql & DevNombreSQL(DBLet(Rs!desPobla, "T")) & "','" & DevNombreSQL(DBLet(Rs!provincia, "T"))
                        If DBLet(Rs!PAIS, "T") = "" Then
                            PAIS = "ES"
                        Else
                            PAIS = Rs!PAIS
                        End If
                        Sql = Sql & "','" & DevNombreSQL(DBLet(PAIS, "T")) & "')"
                        cadena2 = cadena2 & ", (" & vUsu.Codigo & Sql
                        If Len(cadena2) > 80000 Then
                            Sql = ""
                            Label2(31).Caption = "Actualiza BD"
                            Label2(31).Refresh
                            Sql = Mid(cadena2, 2)
                            Sql = Tablas & " VALUES " & Sql
                            Conn.Execute Sql
                            cadena2 = ""
                        End If
                   End If
                   
                   Rs.MoveNext
              Wend
              Rs.Close
              If Len(cadena2) > 0 Then
              
                    Label2(31).Caption = "Actualiza BD(2)"
                    Label2(31).Refresh
                    Sql = Mid(cadena2, 2)
                    Sql = Tablas & " VALUES " & Sql
                    Ejecuta Sql
                    
                    cadena2 = ""
              End If
              
              
              
              
              'trimestral
              Label2(31).Caption = "Insertando datos trimest(II)"
              Label2(31).Refresh
              cadena2 = ""
              Sql = "SELECT  tmp347trimestre.cliprov,tmp347.nif,tmp347.codposta, sum(trim1) as t1, sum(trim2) as t2,"
              Sql = Sql & " sum(trim3) as t3, sum(trim4) as t4,sum(metalico) as metalico"
              Sql = Sql & " from ariconta" & Me.ListView1(1).ListItems(I).Text & ".tmp347,ariconta" & Me.ListView1(1).ListItems(I).Text & ".tmp347trimestre where tmp347.codusu=" & vUsu.Codigo
              Sql = Sql & " and tmp347.codusu=tmp347trimestre.codusu"
              Sql = Sql & " and tmp347.cliprov=tmp347trimestre.cliprov"
              Sql = Sql & " and tmp347.cta=tmp347trimestre.cta "
              Sql = Sql & " group by tmp347.cliprov,tmp347.nif"
              
              
              'AHORA
              Sql = "SELECT  tmp347trimestre.cliprov,tmp347trimestre.nif, sum(trim1) as t1, sum(trim2) as t2,"
              Sql = Sql & " sum(trim3) as t3, sum(trim4) as t4,sum(metalico) as metalico"
              Sql = Sql & " from ariconta" & Me.ListView1(1).ListItems(I).Text & ".tmp347trimestre where tmp347trimestre.codusu=" & vUsu.Codigo
              Sql = Sql & " group by tmp347trimestre.cliprov,tmp347trimestre.nif"
              
              Set Rs = New ADODB.Recordset
              Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
              cadena2 = ""
              While Not Rs.EOF
              
           
              
                   Label2(31).Caption = Rs!NIF
                   Label2(31).Refresh
                   If ExisteEntradaTrimestral(I1, I2, i3, i4, I5) Then
                        I1 = I1 + Rs!T1
                        I2 = I2 + Rs!t2
                        i3 = i3 + Rs!T3
                        i4 = i4 + Rs!T4
                        I5 = I5 + Rs!metalico
                        Sql = "UPDATE tmp347trimestral SET "
                        
                        
                        Sql = Sql & " trim1=" & TransformaComasPuntos(CStr(I1))
                        Sql = Sql & ", trim2=" & TransformaComasPuntos(CStr(I2))
                        Sql = Sql & ", trim3=" & TransformaComasPuntos(CStr(i3))
                        Sql = Sql & ", trim4=" & TransformaComasPuntos(CStr(i4))
                        Sql = Sql & ", metalico=" & TransformaComasPuntos(CStr(I5))
                        Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " AND cliprov =" & Rs!cliprov
                        Sql = Sql & " AND nif = '" & Rs!NIF & "';"
                        Conn.Execute Sql
                   Else
                        
                        
                        Sql = ", (" & vUsu.Codigo & "," & Rs!cliprov & ",'" & Rs!NIF & "',"
                        Sql = Sql & TransformaComasPuntos(CStr(Rs!T1)) & "," & TransformaComasPuntos(CStr(Rs!t2)) & ","
                        Sql = Sql & TransformaComasPuntos(CStr(Rs!T3)) & "," & TransformaComasPuntos(CStr(Rs!T4))
                        Sql = Sql & ",0," '& DBSet(Rs!codposta, "T") & ","
                        Sql = Sql & TransformaComasPuntos(CStr(Rs!metalico)) & ")"
                        cadena2 = cadena2 & Sql
                    
                        If Len(cadena2) > 80000 Then
                            Label2(31).Caption = "Actualiza BD(3)"
                            Label2(31).Refresh
                            Sql = "insert into tmp347trimestral (`codusu`,`cliprov`,`nif`,`trim1`,`trim2`"
                            Sql = Sql & ",`trim3`,`trim4`,`codposta`,metalico) VALUES "
                            cadena2 = Mid(cadena2, 2)
                            Sql = Sql & cadena2
                            Conn.Execute Sql
                            cadena2 = ""
                        End If
                   End If
                   
                   Rs.MoveNext
              Wend
              Rs.Close
              If Len(cadena2) > 0 Then
                    Label2(31).Caption = "Actualiza BD(4)"
                    Label2(31).Refresh
                    Sql = "insert into tmp347trimestral (`codusu`,`cliprov`,`nif`,`trim1`,`trim2`"
                    Sql = Sql & ",`trim3`,`trim4`,`codposta`,metalico) VALUES "
                    cadena2 = Mid(cadena2, 2)
                    Sql = Sql & cadena2
                    Conn.Execute Sql
                    cadena2 = ""
                End If
              
              
              
              
              espera 0.5
         End If
    Next I
    ComprobarCuentas347 = True
    
End Function



Private Sub CopiarFicheroHaciend2(Modelo347 As Boolean)
    On Error GoTo ECopiarFichero347
   
    Sql = ""
    If txtTipoSalida(1).Text <> "" Then cd1.FileName = txtTipoSalida(1).Text
    cd1.CancelError = True
    cd1.ShowSave
    If Modelo347 Then
        Sql = App.Path & "\347.txt"
    Else
        Sql = App.Path & "\mod349.txt"
    End If
    If cd1.FileTitle <> "" Then
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El fichero ya existe. ¿Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Sql = ""
        End If
        If Sql <> "" Then
            FileCopy Sql, cd1.FileName
            MsgBox Space(20) & "Copia efectuada correctamente" & Space(20), vbInformation
        End If
    End If
    Exit Sub
ECopiarFichero347:
    If Err.Number <> 32755 Then MuestraError Err.Number, "Copiar fichero 347"
    
End Sub

Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If txtFecha(Indice1).Text <> "" And txtFecha(Indice2).Text <> "" Then
        If CDate(txtFecha(Indice1).Text) > CDate(txtFecha(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function

Private Function ComprobarCuentas347_DOS(Contabilidad As String, Empresa As String, SoloUnaEMpresa As Boolean) As Boolean
Dim Sql2 As String
Dim SqlTot As String
Dim Rs As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim I1 As Currency
Dim I2 As Currency
Dim i3 As Currency
Dim Trimestre(3) As Currency
Dim Impor As Currency
Dim Tri As Byte
Dim VectorFacturas As String
Dim NIF_En_PROCESO As String


'Acelerar proceso
Dim NuevoSelect As String
Dim ImporteLimite As Currency
Dim CodmactaTratando As String
Dim CIFTratando As String
Dim FinBucle As Boolean
Dim ComprobarInsercion As Boolean
Dim CadenaCifs As String
Dim RestoValoresInsert As String
Dim KK As Integer
Dim IvaSuplidos As String

'Solucion de "mergencia"
Dim VectorFacturasISP As String

Dim IvasSuplidos As String

Dim CadenaInsercion As String
Dim CadenaInsercionCob As String


On Error GoTo EComprobarCuentas347
    
    ComprobarCuentas347_DOS = False
    
    'Solo una empresa
    If SoloUnaEMpresa Then ImporteLimite = ImporteFormateado(Text347(1).Text)
    
    
    'Vaciamos
    Sql = "DELETE FROM " & Contabilidad & ".tmp347 where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    Sql = "DELETE FROM " & Contabilidad & ".tmp347trimestre where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    Set Rs = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    
    
    IvasSuplidos = ""
    Sql = "Select codigiva FROM " & Contabilidad & ".tiposiva where tipodiva= 4" '4: suplidos
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        While Not Rs.EOF
            IvasSuplidos = IvasSuplidos & ", " & Rs!codigiva
            Rs.MoveNext
        Wend
        IvasSuplidos = Mid(IvasSuplidos, 2)
        IvasSuplidos = "(" & IvasSuplidos & ")"
    End If
    Rs.Close
    
    'Para lo nuevo. Iremos codmacta a codmacta
    Sql = " Select factcli.codmacta,trim(factcli.nifdatos) nifdatos,factcli.dirdatos,coalesce(factcli.codpobla,0) codpobla,factcli.nommacta,factcli.despobla,factcli.desprovi,factcli.codpais from "
    Sql = Sql & Contabilidad & ".factcli, " & Contabilidad & ".cuentas  where "
    Sql = Sql & " cuentas.codmacta=factcli.codmacta and model347=1 "
    Sql = Sql & " AND fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    Sql = Sql & " AND fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    
    ' Para debug
    'SQL = SQL & " AND factcli.nifdatos IN ('X19455039Y','X19844591F','X20164371H','24367501J','724367501J')"
    
    If Text347(2).Text <> "" Then
        Sql = Sql & " AND factcli.nifdatos IN ("
        If InStr(1, Text347(2).Text, ",") = 0 Then
            Sql = Sql & DBSet(Text347(2).Text, "T")
        Else
            Sql = Sql & Text347(2).Text
        End If
        Sql = Sql & " )"
    End If
    Sql = Sql & " group by  factcli.codmacta,factcli.nifdatos "
    Sql = Sql & " ORDER  by  factcli.nifdatos,factcli.codmacta "
    Label2(31).Caption = "Leyendo cta/CIF"
    Label2(31).Refresh
    Rs.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    CadenaInsercion = ""
    CadenaInsercionCob = ""
    While Not Rs.EOF
        
        
        Tri = 0
        CadenaCifs = ""
        RestoValoresInsert = ""
        FinBucle = False
        Label2(31).Caption = "CLI "
        
        While Not FinBucle
            If Rs.EOF Then
                FinBucle = True
            Else
                
'                Do
'                    NuevoSelect = Mid(Rs!nifdatos, 1, 2)
'                    Rs.MoveNext
'                Loop Until Val(NuevoSelect) > 48
'
            
                NuevoSelect = Rs!nifdatos
                'If Rs.Fields(1) = "20342507L      " Then S top
            
                CadenaCifs = CadenaCifs & ", " & UCase(DBSet(Rs.Fields(1).Value, "T"))
                
                Sql = UCase(Rs!nifdatos) & "@#€" & DBLet(Rs!CodPobla, "T") & "|" & DBLet(Rs!Nommacta, "T") & "|" & DBLet(Rs!dirdatos, "T") & "|" & DBLet(Rs!desPobla, "T") & "|" & DBLet(Rs!desProvi, "T") & "|" & DBLet(Rs!codpais, "T") & "|€€##"
                RestoValoresInsert = RestoValoresInsert & Sql
                Tri = Tri + 1
                
                Rs.MoveNext
                
                If Tri > 5 Then
                    If Not Rs.EOF Then
                        
                        If NuevoSelect = UCase(Rs!nifdatos) Then
                            FinBucle = False
                        Else
                            FinBucle = True
                        End If
                        
                    Else
                        FinBucle = True
                    End If
                Else
                    If Rs.EOF Then FinBucle = True
                End If
                
            End If
        Wend
        If CadenaCifs <> "" Then CadenaCifs = Mid(CadenaCifs, 2)
        Label2(31).Caption = Mid(CadenaCifs, 3, 9) & ".."
        Label2(31).Refresh
        
        
        NuevoSelect = "Select codmacta,nifdatos,(month(c.fecfactu)-1) div 3  trimestre ,sum(coalesce(baseimpo,0)) base, sum(coalesce(impoiva,0)) iva, sum(coalesce(imporec,0)) recargo"
          
        NuevoSelect = NuevoSelect & " from " & Contabilidad & ".factcli c ," & Contabilidad & ".factcli_totales t"
        NuevoSelect = NuevoSelect & " WHERE  C.NUmSerie = T.NUmSerie AND c.numfactu  = t.numfactu"
        NuevoSelect = NuevoSelect & " AND c.fecfactu  = t.fecfactu AND c.anofactu = t.anofactu"
        NuevoSelect = NuevoSelect & " AND c.fecfactu >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        NuevoSelect = NuevoSelect & " AND c.fecfactu <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        
        If IvasSuplidos <> "" Then NuevoSelect = NuevoSelect & " AND t.codigiva not in " & IvasSuplidos
        
        NuevoSelect = NuevoSelect & " AND c.nifdatos IN (" & CadenaCifs & ")"
        NuevoSelect = NuevoSelect & " GROUP BY 1,2,3"
        RT.Open NuevoSelect, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
        
        If Not RT.EOF Then
            CIFTratando = ""
            CodmactaTratando = ""
            FinBucle = False
        
            
        

            
            Do
                
                If RT.EOF Then
                    ComprobarInsercion = True
                    FinBucle = True
                Else
                    
                    'If RT!nifdatos = "NIF" Then S top
            
                    
                    If CIFTratando <> UCase(RT!nifdatos) Then
                        ComprobarInsercion = True
                    Else
                        ComprobarInsercion = False
                    End If
                End If
                    
                If ComprobarInsercion Then
                    'El importe final es la suma de las bases mas los ivas
                    If CIFTratando <> "" Then
                            
                        I1 = I1 + I2  'importe total
                        
                        CadenaCifs = UCase(CIFTratando) & "@#€"
                        KK = InStr(1, RestoValoresInsert, CadenaCifs)
                        If KK = 0 Then Err.Raise 513, , "Imposible localizar NIF posicion 1 " & CIFTratando
                        
                        CadenaCifs = Mid(RestoValoresInsert, Len(CadenaCifs) + KK)
                        KK = InStr(1, CadenaCifs, "€€##")
                        If KK = 0 Then Err.Raise 513, , "Imposible localizar NIF posicion 2" & CIFTratando
                        CadenaCifs = Mid(CadenaCifs, 1, KK - 1)
                    
                    
                        
                        Sql = ", (" & vUsu.Codigo & "," & Asc("0") & ",'" & CodmactaTratando & "','"
                        Sql = Sql & CIFTratando & "'," & DBSet(RecuperaValor(CadenaCifs, 1), "T") & "," & TransformaComasPuntos(CStr(I1))
                        Sql = Sql & "," & DBSet(RecuperaValor(CadenaCifs, 2), "T") & "," & DBSet(RecuperaValor(CadenaCifs, 3), "T") & "," & DBSet(RecuperaValor(CadenaCifs, 4), "T") & "," & DBSet(RecuperaValor(CadenaCifs, 5), "T") & "," & DBSet(RecuperaValor(CadenaCifs, 6), "T") & ")"
                        CadenaInsercion = CadenaInsercion & Sql
                                               
                         
                        
                        'El del trimestre
                        'esta bajo
                        Sql = ", (" & vUsu.Codigo & "," & Asc("0") & ",'" & CodmactaTratando & "'," & DBSet(CIFTratando, "T") & "," & DBSet(RecuperaValor(CadenaCifs, 1), "T")
                        Sql = Sql & "," & TransformaComasPuntos(CStr(Trimestre(0))) & "," & TransformaComasPuntos(CStr(Trimestre(1)))
                        Sql = Sql & "," & TransformaComasPuntos(CStr(Trimestre(2))) & "," & TransformaComasPuntos(CStr(Trimestre(3))) & ")"
                        CadenaInsercionCob = CadenaInsercionCob & Sql
                    
                    
                        If Len(CadenaInsercion) > 50000 Then
                            Label2(31).Caption = "Insert cta-nif BD"
                            Label2(31).Refresh
                            'Este es la insercion
                            Sql = "INSERT INTO " & Contabilidad & ".tmp347 (codusu, cliprov, cta, nif, codposta, importe, razosoci, dirdatos, despobla, provincia, pais )  VALUES "
                            CadenaInsercion = Mid(CadenaInsercion, 2)
                            Sql = Sql & CadenaInsercion
                            Conn.Execute Sql
                            CadenaInsercion = ""
                            
                            Sql = "insert into " & Contabilidad & ".`tmp347trimestre` (`codusu`,`cliprov`,`cta`,`nif`,`codposta`,`trim1`,`trim2`,`trim3`,`trim4`) VALUES "
                            CadenaInsercionCob = Mid(CadenaInsercionCob, 2)
                            Sql = Sql & CadenaInsercionCob
                            Conn.Execute Sql
                            CadenaInsercionCob = ""
                          End If
                    End If
                    If Not FinBucle Then
                        'If InStr(1, RT!nifdatos, "20405409Q") > 0 Then S top
                        CIFTratando = Trim(UCase(RT!nifdatos))
                        CodmactaTratando = RT!codmacta
                    End If
                    Trimestre(0) = 0: Trimestre(1) = 0: Trimestre(2) = 0: Trimestre(3) = 0
        
                    I1 = 0
                    I2 = 0
                End If
            
                If Not RT.EOF Then
                    I1 = I1 + DBLet(RT!Base, "N")
                    I2 = I2 + DBLet(RT!IVA, "N") + DBLet(RT!recargo, "N")
                    Impor = DBLet(RT!Base, "N") + DBLet(RT!IVA, "N") + DBLet(RT!recargo, "N")
                
                
                
                    'El trimestre
                    
                    'Tri = QueTrimestre(RT!fecliqcl)
                    Tri = RT!Trimestre
                    'Tri = Tri - 1
                    
                    Trimestre(Tri) = Trimestre(Tri) + Impor
                    RT.MoveNext
                End If
            
            Loop Until FinBucle
            
            
        End If
        RT.Close
        
    Wend
    Rs.Close
    
    If Len(CadenaInsercion) > 0 Then
        Label2(31).Caption = "Insert cta-nif BD"
        Label2(31).Refresh
        'Este es la insercion
        Sql = "INSERT INTO " & Contabilidad & ".tmp347 (codusu, cliprov, cta, nif, codposta, importe, razosoci, dirdatos, despobla, provincia, pais )  VALUES "
        CadenaInsercion = Mid(CadenaInsercion, 2)
        Sql = Sql & CadenaInsercion
        Conn.Execute Sql
        
        
        Sql = "insert into " & Contabilidad & ".`tmp347trimestre` (`codusu`,`cliprov`,`cta`,`nif`,`codposta`,`trim1`,`trim2`,`trim3`,`trim4`) VALUES "
        CadenaInsercionCob = Mid(CadenaInsercionCob, 2)
        Sql = Sql & CadenaInsercionCob
        Conn.Execute Sql
    End If
    
    
    
   
    
    Label2(31).Caption = "Comprobando datos facturas proveedor"
    DoEvent2
    espera 0.2
    
     If OptProv(0).Value Then
        cad = "fecharec"
    Else
        cad = "fecfactu"
        
    End If
    
    
    Sql = "SELECT factpro.codmacta,coalesce(factpro.nifdatos,'ERROR') nifdatos, factpro.codpobla, factpro.dirdatos, factpro.nommacta,factpro.despobla,factpro.desprovi,"
    Sql = Sql & " factpro.codpais from " & Contabilidad & ".factpro," & Contabilidad & ".cuentas  where "
    Sql = Sql & Contabilidad & ".cuentas.codmacta=" & Contabilidad & ".factpro.codmacta and model347=1 "
    Sql = Sql & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
    Sql = Sql & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
    'Para debug
    'SQL = SQL & " AND factpro.nifdatos IN ('19455039Y','19844591F','20164371H','24367501J','724367501J')"
    If Text347(2).Text <> "" Then
        Sql = Sql & " AND factpro.nifdatos IN ("
        If InStr(1, Text347(2).Text, ",") = 0 Then
            Sql = Sql & DBSet(Text347(2).Text, "T")
        Else
            Sql = Sql & Text347(2).Text
        End If
        Sql = Sql & " )"
    End If
    
    Sql = Sql & " group by factpro.codmacta, factpro.nifdatos "
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Label2(31).Caption = "PRO " & Rs!nifdatos & " " & Rs!codmacta
        Label2(31).Refresh
        DoEvent2
        'SQL = "Select factpro.*," & cad & " fecha from " & Contabilidad & ".factpro factpro where codmacta = '" & Rs.Fields(0) & "' AND "
        Sql = "Select factpro.*," & cad & " fecha from " & Contabilidad & ".factpro factpro where "
        Sql = Sql & " nifdatos = " & DBSet(Rs!nifdatos, "T")
        Sql = Sql & " AND codmacta = " & DBSet(Rs!codmacta, "T")
        Sql = Sql & " AND " & cad & " >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
        Sql = Sql & " AND " & cad & " <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
        RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I1 = 0
        I2 = 0
        VectorFacturas = ""
        VectorFacturasISP = ""
        Trimestre(0) = 0: Trimestre(1) = 0: Trimestre(2) = 0: Trimestre(3) = 0
        While Not RT.EOF
            If RT!CodOpera = 4 Then
                'St op
                VectorFacturasISP = VectorFacturasISP & ", (" & DBSet(RT!NUmSerie, "T") & "," & RT!Numregis & "," & RT!Anofactu & ")"
            Else
                VectorFacturas = VectorFacturas & ", (" & DBSet(RT!NUmSerie, "T") & "," & RT!Numregis & "," & RT!Anofactu & ")"
            End If
            RT.MoveNext
        Wend
        RT.Close
        
        
        'FACTURAS NORMALES
        If VectorFacturas <> "" Then
            Impor = 0
            VectorFacturas = Mid(VectorFacturas, 2)
            
            If OptProv(0).Value Then
                SqlTot = "select (month(factpro_totales.fecharec)-1) div 3  trimestre ,sum(coalesce(baseimpo,0)) base, sum(coalesce(impoiva,0)) iva, sum(coalesce(imporec,0)) recargo "
            Else
                SqlTot = "select (month(fecfactu)-1) div 3  trimestre ,sum(coalesce(baseimpo,0)) base, sum(coalesce(impoiva,0)) iva, sum(coalesce(imporec,0)) recargo "
            End If
            SqlTot = SqlTot & " from " & Contabilidad & ".factpro_totales  "
            If Not OptProv(0).Value Then
                SqlTot = SqlTot & " INNER JOIN " & Contabilidad & ".factpro  "
                SqlTot = SqlTot & " ON factpro_totales.numserie=factpro.numserie and factpro_totales.numregis=factpro.numregis"
                SqlTot = SqlTot & " and factpro_totales.anofactu=factpro.anofactu "
            End If
            SqlTot = SqlTot & " WHERE  (factpro_totales.numserie,factpro_totales.numregis,factpro_totales.anofactu) IN (" & VectorFacturas & ") GROUP BY 1"
            If IvasSuplidos <> "" Then NuevoSelect = NuevoSelect & " AND factpro_totales.codigiva not in " & IvasSuplidos
            
            RT.Open SqlTot, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                I1 = I1 + DBLet(RT!Base, "N")
                
                    I2 = I2 + DBLet(RT!IVA, "N") + DBLet(RT!recargo, "N")
                    Impor = DBLet(RT!Base, "N") + DBLet(RT!IVA, "N") + DBLet(RT!recargo, "N")
                
                
            
            
                'El trimestre
                'Tri = QueTrimestre(RT!Fecha)
                Tri = RT!Trimestre
                'Tri = Tri - 1
                Trimestre(Tri) = Trimestre(Tri) + Impor
                
            
                RT.MoveNext
            Wend
            RT.Close
        End If 'VectorFacturas
        
        'INVERSION DEL SUJETO PASIVO. EL IVA no lo lleva la factura. Nos lo autoimputamos, con lo cual NO lo tenemos qe sumar
        If VectorFacturasISP <> "" Then
            Impor = 0
            VectorFacturasISP = Mid(VectorFacturasISP, 2)
            SqlTot = "select (month(fecharec)-1) div 3  trimestre ,sum(coalesce(baseimpo,0)) base, sum(coalesce(impoiva,0)) iva, sum(coalesce(imporec,0)) recargo "
            SqlTot = SqlTot & " from " & Contabilidad & ".factpro_totales  WHERE "
            'SqlTot = SqlTot & " numserie = " & DBSet(RT!NUmSerie, "T")
            'SqlTot = SqlTot & " and numregis = " & DBSet(RT!Numregis, "N")
            'SqlTot = SqlTot & " and anofactu = " & DBSet(RT!anofactu, "N")
            SqlTot = SqlTot & " (numserie,numregis,anofactu) IN (" & VectorFacturasISP & ") GROUP BY 1"
            
            
            RT.Open SqlTot, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                    I1 = I1 + DBLet(RT!Base, "N")
                
                    I2 = I2 + 0  'NO CARGAMOS EL IVA I2 + DBLet(RT!IVA, "N") + DBLet(RT!recargo, "N")
                    Impor = DBLet(RT!Base, "N") '+ DBLet(RT!IVA, "N") + DBLet(RT!recargo, "N")
                
                
            
            
                'El trimestre
                Tri = RT!Trimestre
                Trimestre(Tri) = Trimestre(Tri) + Impor
                
                RT.MoveNext
            Wend
            RT.Close
        End If 'VectorFacturas
        
        'El importe final es la suma de las bases mas los ivas
        I1 = I1 + I2
        Sql = "INSERT INTO " & Contabilidad & ".tmp347 (codusu, cliprov, cta, nif, codposta, importe, razosoci, dirdatos, despobla, provincia, pais)  "
        'SQL = SQL & " VALUES (" & vUsu.Codigo & ",1,'" & RS!Codmacta & "','"
        Sql = Sql & " VALUES (" & vUsu.Codigo & "," & Asc("1") & ",'" & Rs!codmacta & "','" & DBLet(Rs!nifdatos) & "',"
        If IsNull(Rs!CodPobla) Then
            Sql = Sql & "'00000'"
        Else
            Sql = Sql & DBSet(Rs!CodPobla, "T")
        End If
        Sql = Sql & "," & TransformaComasPuntos(CStr(I1))
        Sql = Sql & "," & DBSet(Rs!Nommacta, "T") & "," & DBSet(Rs!dirdatos, "T") & "," & DBSet(Rs!desPobla, "T") & "," & DBSet(Rs!desProvi, "T") & "," & DBSet(Rs!codpais, "T") & ")"
        Conn.Execute Sql
        
        
        'El del trimestre
        Sql = "insert into " & Contabilidad & ".`tmp347trimestre` (`codusu`,`cliprov`,`cta`,`nif`,`codposta`,`trim1`,`trim2`,`trim3`,`trim4`)"
        Sql = Sql & " VALUES (" & vUsu.Codigo & "," & Asc("1") & ",'" & Rs!codmacta & "'," & DBSet(Rs!nifdatos, "T") & "," & DBSet(IIf(IsNull(Rs!CodPobla), "0000", Rs!CodPobla), "T")
        Sql = Sql & "," & TransformaComasPuntos(CStr(Trimestre(0))) & "," & TransformaComasPuntos(CStr(Trimestre(1)))
        Sql = Sql & "," & TransformaComasPuntos(CStr(Trimestre(2))) & "," & TransformaComasPuntos(CStr(Trimestre(3))) & ")"
        Conn.Execute Sql
        
        
        Rs.MoveNext
        
    Wend
    Rs.Close
    
    ' CObros en metalico superiores a 6000
    Label2(31).Caption = "Cobros metalico"
    Label2(31).Refresh
    DoEvent2
    Sql = "Select ImporteMaxEfec340 from " & Contabilidad & ".parametros "
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO pues ser eof
    I1 = DBLet(Rs!ImporteMaxEfec340, "N")
    Rs.Close
    If I1 > 0 Then
        'SI que lleva control de cobros en efectivo
        'Veremos si hay conceptos de efectivo
        Sql = "Select codconce from " & Contabilidad & ".conceptos where EsEfectivo340 = 1"
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        While Not Rs.EOF
            Sql = Sql & ", " & Rs!CodConce
            Rs.MoveNext
        Wend
        Rs.Close
        Sql2 = "" 'Errores en Datos en efectivo sin ventas
        If Sql <> "" Then
            Sql = Mid(Sql, 2) 'quit la coma
            
            cad = "Select * from " & Contabilidad & ".tmp347trimestre WHERE codusu = " & vUsu.Codigo & " ORDER BY cta"
            RT.Open cad, Conn, adOpenKeyset, adCmdText
            
            'HABER -DEBE"
            cad = "Select hlinapu.codmacta,sum(if(timporteh is null,0,timporteh))-sum(if(timported is null,0,timported)) importe, cuentas.nifdatos, cuentas.codposta"
            cad = cad & " from " & Contabilidad & ".hlinapu,cuentas WHERE hlinapu.codmacta =cuentas.codmacta "
            cad = cad & " AND model347=1 AND fechaent >='" & Format(txtFecha(0).Text, FormatoFecha) & "'"
            cad = cad & " AND fechaent <='" & Format(txtFecha(1).Text, FormatoFecha) & "'"
            cad = cad & " AND codconce IN (" & Sql & ")"
            cad = cad & " group by 1 order by 1"

            Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs.EOF
                Label2(31).Caption = Rs!codmacta
                Label2(31).Refresh
        
                If Rs!Importe <> 0 Then
                    Sql = "cta  = '" & Rs!codmacta & "'"
                    RT.Find Sql, , adSearchForward, 1
                    
                    If RT.EOF Then
                        Sql2 = Sql2 & Rs!codmacta & " (" & Rs!Importe & ") " & vbCrLf
                    Else
                        Sql = "UPDATE " & Contabilidad & ".tmp347trimestre SET metalico = " & TransformaComasPuntos(CStr(Rs!Importe))
                        Sql = Sql & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & RT!Cta & "'"
                        '++
                        Sql = Sql & " and nif = " & DBSet(Rs!nifdatos, "T")
                        Conn.Execute Sql
                    End If
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            RT.Close
            
            If Sql2 <> "" Then
                Sql2 = "Cobros en efectivo sin asociar a ventas" & vbCrLf & Sql2
                MsgBox Sql2, vbExclamation
            End If
        End If
    End If
    
    Set RT = Nothing
    RC = ""
    cad = ""
    Sql2 = ""
    'Comprobaremos k el nif no es nulo, ni el codppos de las cuentas a tratar
    Sql = "Select cta from " & Contabilidad & ".tmp347 where (nif is null or nif = '') and codusu = " & vUsu.Codigo
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not Rs.EOF
        I = I + 1
        cad = cad & Rs.Fields(0) & "       "
        If I = 3 Then
            cad = cad & vbCrLf
            I = 0
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    If cad <> "" Then
        RC = "Cuentas con NIF sin valor: " & vbCrLf & vbCrLf & cad
        cad = ""
    End If
    
    'Comprobamos el codpos
    Sql = "Select cta,razosoci,codposta from " & Contabilidad & ".tmp347 where codusu = " & vUsu.Codigo
    Sql = Sql & " AND (codposta is null or codposta='')"

    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not Rs.EOF
        I = I + 1
        cad = cad & Rs.Fields(0) & "       "
        If I = 3 Then
            cad = cad & vbCrLf
            I = 0
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    If cad <> "" Then
        If RC <> "" Then RC = RC & vbCrLf & vbCrLf & vbCrLf
        RC = RC & "Cuentas con codigo postal sin valor: " & vbCrLf & vbCrLf & cad
    End If
    
    If RC <> "" Then
        RC = "Empresa: " & Empresa & vbCrLf & vbCrLf & RC & vbCrLf & " Desea continuar igualmente?"
        If MsgBox(RC, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    
    Set Rs = Nothing
    
    ComprobarCuentas347_DOS = True
    Exit Function
EComprobarCuentas347:

    RC = CStr(Sql)
    If Len(RC) > 400 Then RC = Mid(RC, 1, 400)
    cad = Err.Description
    MuestraError Err.Number, "Comprobar Cuentas 347" & vbCrLf & RC & vbCrLf & Err.Description
    CreaFichErrores347
    
    
End Function

Private Sub CreaFichErrores347()
    On Error Resume Next
    Dim NFF As Integer
    NFF = FreeFile
    Open App.Path & "\Err347.txt" For Output As #NFF
    Print #NFF, Err.Description
    Print #NFF, ""
    Print #NFF, cad
    Print #NFF, ""
    Print #NFF, Sql
    Close #NFF
    Err.Clear
    
End Sub

Private Function ExisteEntrada() As Boolean
Dim B As Boolean
    
    If EslaPrimeraEmpresa Then
        ExisteEntrada = False
        Exit Function
    End If
        
    Sql = "Select importe from tmp347tot  where codusu = " & vUsu.Codigo & " and cliprov =" & Rs!cliprov & " AND nif ='" & Rs!NIF & "'"
    'SQL = SQL & " and codposta = " & DBSet(Rs!codposta, "T") & ";"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        ExisteEntrada = True
        Importe = miRsAux!Importe
    Else
        ExisteEntrada = False
    End If
    miRsAux.Close
End Function

Private Function ExisteEntradaTrimestral(ByRef I1 As Currency, ByRef I2 As Currency, ByRef i3 As Currency, ByRef i4 As Currency, ByRef I5 As Currency) As Boolean
    
    If EslaPrimeraEmpresa Then
        ExisteEntradaTrimestral = False
        I1 = 0: I2 = 0: i3 = 0: i4 = 0: I5 = 0
        Exit Function
    End If

    'SQL = "Select trim1,trim2,trim3,trim4,metalico from tmp347trimestral  where codusu = " & vUsu.Codigo & " and cliprov =" & Rs!cliprov & " AND nif ='" & Rs!NIF & "' and codposta = " & DBSet(Rs!codposta, "T") & ";"
    Sql = "Select trim1,trim2,trim3,trim4,metalico from tmp347trimestral  where codusu = " & vUsu.Codigo & " and cliprov =" & Rs!cliprov & " AND nif ='" & Rs!NIF & "';"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        ExisteEntradaTrimestral = True
        I1 = miRsAux!trim1
        I2 = miRsAux!trim2
        i3 = miRsAux!trim3
        i4 = miRsAux!trim4
        I5 = DBLet(miRsAux!metalico, "N")
    Else
        ExisteEntradaTrimestral = False
        I1 = 0: I2 = 0: i3 = 0: i4 = 0: I5 = 0
    End If
    miRsAux.Close
End Function

'Dada una fecha me da el trimestre
Private Function QueTrimestre(Fecha As Date) As Byte
Dim C As Byte
    
        C = Month(Fecha)
        If C < 4 Then
            QueTrimestre = 1
        ElseIf C < 7 Then
            QueTrimestre = 2
        ElseIf C < 10 Then
            QueTrimestre = 3
        Else
            QueTrimestre = 4
        End If
    
End Function

Private Sub VisibleMasivo()
    
    Me.chkMasivo.visible = False

    If Me.optTipoSal(3).Value Then
        If Combo5.ListIndex = 0 Then chkMasivo.visible = True
    End If
    

End Sub


Private Sub EnviarPorEmail()
Dim B As Boolean
Dim Email As String
        
    If Me.chkMasivo.Value = 1 Then
    
        ListView2.ListItems.Clear
    
        'Frame de envio email
        B = CargaDatosEnvioMasivo
        
        If B Then
            PonerFrameEnvioMasivo True
            Me.Refresh
            Screen.MousePointer = vbHourglass
            PonEmails
            
            
            
            
        End If
        Set miRsAux = Nothing
        Label2(31).Caption = ""
    Else
        
        
        If Text347(2).Text <> "" Then Email = DevuelveEmail(Me.Text347(2).Text, False)
        LanzaProgramaAbrirOutlook 21, Email
    End If
    
End Sub




Private Function CargaDatosEnvioMasivo() As Boolean
Dim IT As ListItem
    CargaDatosEnvioMasivo = False
    
'    'Guardare en tmp347tot.direcion  el email
'    Label2(31).Caption = "Leyendo emails"
'    Label2(31).Refresh
'    Sql = "DELETE FROM tmp347tot where codusu=" & vUsu.Codigo & " and not nif  in ("
'    Sql = Sql & "select nifdatos from cuentas where nifdatos<>'' and apudirec='S' and codmacta like '4%' and codmacta <'47' and maidatos<>'')"
'    Conn.Execute Sql
    
    
    'Los que quedan tienen nif
    Label2(31).Caption = "Cargando desde cuentas"
    Label2(31).Refresh
    
    Set miRsAux = New ADODB.Recordset
    
    Sql = "Select * from tmp347tot where codusu=" & vUsu.Codigo & " ORDER BY nif,cliprov"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Ningun registro con email disponible para su envio", vbExclamation
        miRsAux.Close
    End If
        
        
    While Not miRsAux.EOF
        Label2(31).Caption = "Cargando desde cuentas"
        Label2(31).Refresh
        Set IT = ListView2.ListItems.Add()
        IT.Checked = True
        IT.Text = miRsAux!NIF
        IT.SubItems(1) = miRsAux!razosoci
        IT.SubItems(2) = " " 'email
        IT.SubItems(3) = Right(Space(15) & Format(miRsAux!Importe, FormatoImporte), 15)
        
        
        Sql = "{tmp347tot.codusu} =" & vUsu.Codigo & " AND  {tmp347tot.cliprov} =" & miRsAux!cliprov & " AND {tmp347tot.nif} = " & DBSet(miRsAux!NIF, "F")
        IT.Tag = CStr(Sql)
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    CargaDatosEnvioMasivo = True
End Function


Private Sub PonEmails()

    
    'De tres en tres
    RC = "N"
    
    I = 0
    Do
        I = I + 1
        Label2(31).Caption = "Leyendo email " & ListView2.ListItems(I).Text
        Label2(31).Refresh
        cad = "Select maidatos FROM cuentas where apudirec='S' and maidatos<>'' and nifdatos=" & DBSet(ListView2.ListItems(I).Text, "T")
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        cad = " "
        If Not miRsAux.EOF Then If Not IsNull(miRsAux.Fields(0)) Then cad = miRsAux.Fields(0)
        miRsAux.Close
        
       
        
        ListView2.ListItems(I).SubItems(2) = cad
        If cad = " " Then
            ListView2.ListItems(I).ForeColor = vbRed
            ListView2.ListItems(I).Checked = False
        End If
        If I >= ListView2.ListItems.Count Then RC = ""
    Loop Until RC = ""
End Sub




'Genera uno a uno los rpts
Private Sub GenerarRpt()

    'En I tenemos el indice del lw2

                With frmVisReport
                    .Informe = App.Path & "\Informes\"
                    If ExportarPDF Then
                        'PDF
                        .Informe = .Informe & cadPDFrpt
                    Else
                        'IMPRIMIR
                        .Informe = .Informe & cadNomRPT
                    End If
                    .FormulaSeleccion = "{tmpentrefechas.codusu}=" & vUsu.Codigo & " AND {tmpentrefechas.nomconam}= """ & Rs.Fields(0) & """"
                    .SoloImprimir = False
                    .OtrosParametros = cadParam
                    .NumeroParametros = numParam
                    .ConSubInforme = True
            
                    .NumCopias2 = 1
            
                    .SoloImprimir = SoloImprimir
                    .ExportarPDF = True
                    .MostrarTree = vMostrarTree
                    .Show vbModal
                    'HaPulsadoImprimir = .EstaImpreso
                 
                 End With



End Sub


