VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAsientosHco 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   17805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCreacion 
      Height          =   2145
      Left            =   10170
      TabIndex        =   75
      Top             =   870
      Visible         =   0   'False
      Width           =   6885
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
         Index           =   4
         Left            =   1680
         MaxLength       =   19
         TabIndex        =   81
         Tag             =   "Fecha entrada|FH|S|||hcabapu|feccreacion|dd/mm/yyyy hh:mm:ss||"
         Text            =   "1234567890"
         Top             =   360
         Width           =   2460
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   83
         Tag             =   "Desde Aplicacion|T|S|||hcabapu|desdeaplicacion|||"
         Top             =   1410
         Width           =   4905
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
         Index           =   5
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   82
         Tag             =   "Usuario creacion|T|S|||hcabapu|usucreacion|||"
         Top             =   900
         Width           =   4905
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmAsientosHco.frx":0000
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Datos de Creaci�n"
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
         Left            =   3060
         TabIndex        =   79
         Top             =   210
         Width           =   3480
      End
      Begin VB.Label Label7 
         Caption         =   "Aplicaci�n "
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
         Left            =   330
         TabIndex        =   78
         Top             =   1470
         Width           =   1620
      End
      Begin VB.Label Label6 
         Caption         =   "Usuario "
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
         Left            =   330
         TabIndex        =   77
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label Label5 
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
         Left            =   330
         TabIndex        =   76
         Top             =   450
         Width           =   1620
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3930
      TabIndex        =   73
      Top             =   90
      Width           =   1485
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   74
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Asientos Descuadrados"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Asientos con n�meros incorrectos"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   10170
      TabIndex        =   70
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
         ItemData        =   "frmAsientosHco.frx":008B
         Left            =   90
         List            =   "frmAsientosHco.frx":0098
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   210
         Width           =   2235
      End
   End
   Begin VB.Frame FrameAux2 
      Height          =   2145
      Left            =   10170
      TabIndex        =   58
      Top             =   840
      Width           =   6885
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
         TabIndex        =   68
         Tag             =   "Documento|T|N|||hcabapu_fichdocs|docum|||"
         Text            =   "docum"
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
         TabIndex        =   67
         Tag             =   "Campo|T|N|||hcabapu_fichdocs|campo||S|"
         Text            =   "campo"
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
         TabIndex        =   66
         Tag             =   "Descripcion|T|N|||hcabapu_fichdocs|descripfich||N|"
         Text            =   "descripcion"
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
         TabIndex        =   65
         Tag             =   "Orden|N|N|0||hcabapu_fichdocs|orden||S|"
         Text            =   "Orden"
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
         TabIndex        =   64
         Tag             =   "numero diario|N|N|0||hcabapu_fichdocs|numdiari||S|"
         Text            =   "Codigo"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   63
         Tag             =   "numero diario|N|N|0||hcabapu_fichdocs|numdiari||S|"
         Text            =   "Diario"
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
         Tag             =   "Fecha entrada|F|N|||hcabapu_fichdocs|fechaent|dd/mm/yyyy|S|"
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
         Index           =   0
         Left            =   330
         TabIndex        =   61
         Tag             =   "N� asiento|N|S|0||hcabapu_fichdocs|numasien||S|"
         Text            =   "Asiento"
         Top             =   1620
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   90
         TabIndex        =   59
         Top             =   120
         Width           =   2625
         Begin MSComctlLib.Toolbar ToolbarAux0 
            Height          =   330
            Left            =   210
            TabIndex        =   60
            Top             =   0
            Width           =   2235
            _ExtentX        =   3942
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
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
            EndProperty
         End
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
         Left            =   90
         TabIndex        =   69
         Top             =   510
         Width           =   6555
         _ExtentX        =   11562
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
         Alignment       =   1  'Right Justify
         Caption         =   "Documentos Asociados"
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
         Left            =   3120
         TabIndex        =   72
         Top             =   210
         Width           =   3480
      End
   End
   Begin VB.Frame frameextras 
      Enabled         =   0   'False
      Height          =   915
      Left            =   240
      TabIndex        =   35
      Top             =   9300
      Width           =   14265
      Begin VB.TextBox Text3 
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
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text3"
         Top             =   450
         Width           =   4605
      End
      Begin VB.TextBox Text3 
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
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text3"
         Top             =   450
         Width           =   4245
      End
      Begin VB.TextBox Text3 
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text3"
         Top             =   450
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Centro Coste"
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
         Index           =   3
         Left            =   9360
         TabIndex        =   41
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
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
         Left            =   4950
         TabIndex        =   40
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Cta. Contrapartida"
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
         Index           =   5
         Left            =   360
         TabIndex        =   39
         Top             =   180
         Width           =   2295
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
      TabIndex        =   33
      Top             =   270
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5550
      TabIndex        =   31
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   32
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
      TabIndex        =   28
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   30
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
      Height          =   2160
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   870
      Width           =   9810
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
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text4"
         Top             =   510
         Width           =   4785
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
         Height          =   855
         Index           =   3
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "Observaciones|T|S|||hcabapu|obsdiari|||"
         Top             =   1200
         Width           =   9375
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
         Index           =   2
         Left            =   3360
         TabIndex        =   2
         Tag             =   "numero diario|N|N|0||hcabapu|numdiari||S|"
         Text            =   "1234567890"
         Top             =   510
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEF7E4&
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
         Left            =   225
         TabIndex        =   0
         Tag             =   "N� asiento|N|S|0||hcabapu|numasien||S|"
         Top             =   510
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
         Index           =   1
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Fecha entrada|F|N|||hcabapu|fechaent|dd/mm/yyyy|S|"
         Top             =   510
         Width           =   1245
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   1740
         Top             =   900
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   2700
         Picture         =   "frmAsientosHco.frx":00CF
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   4470
         Top             =   210
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
         Left            =   210
         TabIndex        =   11
         Top             =   930
         Width           =   1515
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Inicio"
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
         Left            =   1770
         TabIndex        =   10
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "N� Asiento"
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
         TabIndex        =   8
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Diario"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   6150
      Left            =   225
      TabIndex        =   12
      Top             =   3045
      Width           =   17130
      Begin VB.Frame Frame2 
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
         Height          =   825
         Index           =   1
         Left            =   9960
         TabIndex        =   51
         Top             =   -90
         Width           =   6885
         Begin VB.TextBox Text2 
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
            Left            =   810
            TabIndex        =   54
            Text            =   "Text2"
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox Text2 
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
            Left            =   4230
            TabIndex        =   53
            Text            =   "Text2"
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox Text2 
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
            Left            =   2580
            TabIndex        =   52
            Text            =   "Text2"
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label Label1 
            Caption         =   "SALDO"
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
            Left            =   810
            TabIndex        =   57
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "HABER"
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
            Left            =   4230
            TabIndex        =   56
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "DEBE"
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
            Left            =   2580
            TabIndex        =   55
            Top             =   150
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdSaldoHco 
         Height          =   495
         Index           =   1
         Left            =   3510
         Picture         =   "frmAsientosHco.frx":015A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Extractos"
         Top             =   30
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdSaldoHco 
         Height          =   495
         Index           =   0
         Left            =   2910
         Picture         =   "frmAsientosHco.frx":69AC
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Saldos en historico"
         Top             =   30
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   48
         Tag             =   "Linea|N|N|0||hlinapu|linliapu||S|"
         Text            =   "linea"
         Top             =   2880
         Visible         =   0   'False
         Width           =   345
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
         Index           =   3
         Left            =   14190
         TabIndex        =   47
         ToolTipText     =   "Buscar centro coste"
         Top             =   2910
         Visible         =   0   'False
         Width           =   195
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
         Index           =   11
         Left            =   13380
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "CC|T|S|||hlinapu|codccost|||"
         Text            =   "CC"
         Top             =   2910
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
         TabIndex        =   24
         Tag             =   "Imp.Haber|N|S|||hlinapu|timporteH|##,###,##0.00||"
         Text            =   "Imp.Haber"
         Top             =   2880
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
         TabIndex        =   23
         Tag             =   "Imp.Debe|N|S|||hlinapu|timporteD|##,###,##0.00||"
         Text            =   "Imp.Debe"
         Top             =   2880
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
         TabIndex        =   22
         Tag             =   "Ampliaci�n|T|S|||hlinapu|ampconce|||"
         Text            =   "ampliacion"
         Top             =   2880
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
         Left            =   9870
         TabIndex        =   46
         ToolTipText     =   "Buscar concepto"
         Top             =   2880
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
         TabIndex        =   20
         Tag             =   "Contrapartida|T|S|||hlinapu|ctacontr|||"
         Text            =   "contrapartida"
         Top             =   2880
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
         Left            =   9060
         TabIndex        =   45
         ToolTipText     =   "Buscar cuenta"
         Top             =   2880
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
         TabIndex        =   21
         Tag             =   "Concepto|T|N|||hlinapu|codconce|||"
         Text            =   "concepto"
         Top             =   2880
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   60
         TabIndex        =   43
         Top             =   0
         Width           =   2625
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   180
            TabIndex        =   44
            Top             =   150
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Saldos"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Extractos"
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Asiento Predefinido"
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
         Left            =   7590
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Documento|T|S|||hlinapu|numdocum|||"
         Text            =   "documento"
         Top             =   2880
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
         TabIndex        =   18
         Tag             =   "Cuenta|T|N|||hlinapu|codmacta|||"
         Text            =   "cta"
         Top             =   2880
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
         TabIndex        =   16
         Tag             =   "numero diario|N|N|0||hlinapu|numdiari||S|"
         Text            =   "diario"
         Top             =   2880
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
         TabIndex        =   15
         Tag             =   "N� asiento|N|S|0||hlinapu|numasien||S|"
         Text            =   "Asiento"
         Top             =   2865
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
         TabIndex        =   17
         Tag             =   "Fecha entrada|F|N|||hlinapu|fechaent|dd/mm/yyyy|S|"
         Text            =   "fecha"
         Top             =   2865
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
         Left            =   4020
         TabIndex        =   14
         ToolTipText     =   "Buscar cuenta"
         Top             =   2910
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
         Index           =   4
         Left            =   4260
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   13
         Text            =   "Nombre cuenta"
         Top             =   2910
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
         Height          =   5310
         Index           =   1
         Left            =   45
         TabIndex        =   26
         Top             =   780
         Width           =   16770
         _ExtentX        =   29580
         _ExtentY        =   9366
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
      Left            =   255
      TabIndex        =   4
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
         TabIndex        =   5
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
      Left            =   15930
      TabIndex        =   29
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
      Left            =   14640
      TabIndex        =   27
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
      Left            =   16620
      TabIndex        =   34
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
      Left            =   15930
      TabIndex        =   9
      Top             =   10350
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarCreacion 
      Height          =   390
      Left            =   16020
      TabIndex        =   80
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
            Object.ToolTipText     =   "Datos de Creaci�n"
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmAsientosHco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Public ASIENTO As String  'Con pipes numdiari|fechanormal|numasien
Public vLinapu As Integer
Public SoloImprimir As Boolean

Public DesdeNorma43 As Byte  'La uno y la 2 son validas
Public Datos As String  'Tendra, empipado, numero asiento  y demas

Private Const NO = "No encontrado"

Private Const IdPrograma = 301

Private WithEvents frmAsi As frmAsientosHcoPrev 'frmBasico2
Attribute frmAsi.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDi As frmTiposDiario
Attribute frmDi.VB_VarHelpID = -1
Private WithEvents frmPre As frmAsiPre
Attribute frmPre.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1


Private WithEvents frmUtil As frmUtilidades
Attribute frmUtil.VB_VarHelpID = -1


Dim AntiguoText1 As String
Private CadenaAmpliacion As String
Private Sql As String

Private LlevaContraPartida As Boolean

Dim PosicionGrid As Integer

Dim Linliapu As Long
Dim FicheroAEliminar As String
Dim IndCodigo As Integer

Dim CtaAnt As String
Dim DebeAnt As String
Dim HaberAnt As String



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

' VARIABLES DOCUMENTOS ASOCIADOS
Dim IT As ListItem
Dim Contador As Integer
Dim Fichero As String
Dim TipoDocu As Byte

Private Const CarpetaIMG = "temp" 'ImgFicFT2

Dim cadFiltro As String
Dim I As Integer

Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    If Modo = 0 Then Exit Sub
    HacerBusqueda2
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
                Set Mc = New Contadores
                I = FechaCorrecta2(CDate(Text1(1).Text))
                If Mc.ConseguirContador("0", (I = 0), False) = 0 Then
                    cmdCancelar.Caption = "Cancelar"
                    'COMPROBAR NUMERO ASIENTO
                    Text1(0).Text = Mc.Contador
                    If ComprobarNumeroAsiento((I = 0)) Then
            
                        B = InsertarDesdeForm2(Me, 1)
                    Else
                        B = False
                    End If
                    
                    If B Then
                        Data1.RecordSource = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PosicionarData
                        PonerCampos
                        BotonAnyadirLinea 1, True
                    Else
                        'SI NO INSERTA debemos devolver el contador
                        Mc.DevolverContador "0", (I = 0), Mc.Contador
                    End If
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                '-----------------------------------------
                'Hacemos modificar
                'PreparaBloquear
                Limp = Modificar
                'TerminaBloquear
                If Limp Then
                    'MsgBox "El registro ha sido modificado", vbInformation
                    If SituarData1(False) Then
                        lblIndicador.Caption = ""
                        PonerModo 2
                    Else
                        PonerModo 0
                    End If
'                    DesBloqAsien   'Desbloqueamos el asiento
                    TerminaBloquear
                    
                    AsientoConExtModificado = 1
                Else
                    PonerCampos
                End If
            Else
                ModoLineas = 0
            End If
        
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    InsertarLinea
                Case 2 'modificar ll�nies
                    ModificarLinea
                    PosicionarData
            End Select
            
            AsientoConExtModificado = 1
    End Select
    
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "numasien= " & DBSet(Text1(0).Text, "N") & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0
            cmdAux(0).Tag = 0
            LlamaContraPar
            If txtAux(4).Text <> "" Then
                PonFoco txtAux(5)
            Else
                PonFoco txtAux(4)
            End If
        Case 1 'Cta contrapartida
            cmdAux(0).Tag = 1
            LlamaContraPar
            txtAux(5).SetFocus
        Case 2 'Conceptos
            Set frmCon = New frmConceptos
            frmCon.DatosADevolverBusqueda = "0|"
            frmCon.Show vbModal
            Set frmCon = Nothing
        Case 3 ' centro de coste
            If txtAux(11).Enabled Then
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

Private Sub cmdSaldoHco_Click(Index As Integer)
Dim Cta As String
    If Modo = 5 And ModoLineas > 0 Then
        If txtAux(4).Text = "" Then
            MsgBox "Seleccione una cuenta", vbExclamation
            Exit Sub
        End If
        Sql = txtAux(4).Text
        Cta = txtAux2(4).Text
    Else
        If AdoAux(1).Recordset.EOF Then
            MsgBox "Ning�n registro activo.", vbExclamation
            Exit Sub
        End If
        Sql = AdoAux(1).Recordset!codmacta
        Cta = DBLet(AdoAux(1).Recordset!Nommacta)
    End If
    If Index = 0 Then
        SaldoHistorico Sql, "", Cta, False
    Else
        If VieneDeDesactualizar Then
            MsgBox "Acaba de desactualizar asientos. No puede hacer consulta desde aqui.", vbExclamation
        Else
            Screen.MousePointer = vbHourglass
            frmConExtr.EjerciciosCerrados = False
            frmConExtr.Cuenta = Sql
            frmConExtr.Show vbModal
        End If
    End If

End Sub


Private Sub Form_Activate()
'    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    
    If PrimeraVez Then
        B = False
        If ASIENTO <> "" Then
            B = True
            Modo = 2
            Sql = "Select * from hcabapu "
            Sql = Sql & " WHERE numasien = " & RecuperaValor(ASIENTO, 3)
            Sql = Sql & " AND numdiari =" & RecuperaValor(ASIENTO, 1)
            Sql = Sql & " AND fechaent= '" & Format(RecuperaValor(ASIENTO, 2), FormatoFecha) & "'"
            CadenaConsulta = Sql
            PonerCadenaBusqueda
            'BOTON lineas
            
            cboFiltro.ListIndex = 0
            
        Else
            FijarDiarioPorDefecto
            Modo = 0
            CadenaConsulta = "Select * from " & NombreTabla & " WHERE numasien = -1"
            Data1.RecordSource = CadenaConsulta
            Data1.Refresh
            
            cboFiltro.ListIndex = vUsu.FiltroAsientos
            
        End If
        
        CargarSqlFiltro
        
        PonerModo CInt(Modo)
        VieneDeDesactualizar = B
        If Modo <> 2 Then
            
            If ASIENTO <> "" Then
                MsgBox "Proceso de sistema. Stop. Frm_Activate"
            End If
        Else

        End If
        If ASIENTO <> "" Then
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
            cadFiltro = "hcabapu.fechaent >= " & DBSet(vParam.fechaini, "F")
        
        Case 2 ' ejercicio actual
            cadFiltro = "hcabapu.fechaent between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
        
        Case 3 ' ejercicio siguiente
            cadFiltro = "hcabapu.fechaent > " & DBSet(vParam.fechafin, "F")
    
    End Select
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    
    vUsu.ActualizarFiltro "ariconta", IdPrograma, Me.cboFiltro.ListIndex
    
End Sub

Private Sub Form_Load()
Dim I As Integer

    Me.Icon = frmPpal.Icon

    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    CadAncho = False
    ActualizandoAsiento = False
    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 42
        .Buttons(2).Image = 47
    End With

    ' Botonera Principal
    With Me.ToolbarCreacion
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 48
    End With


    Me.ToolbarCreacion.Visible = (vUsu.Nivel = 0)
    Me.ToolbarCreacion.Enabled = (vUsu.Nivel = 0)

    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
   
    With Me.ToolbarAux
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 45
        .Buttons(6).Image = 30
        .Buttons(7).Image = 32
    End With
    
    With Me.ToolbarAux0
        .HotImageList = frmPpal.imgListComun_OM16
        .DisabledImageList = frmPpal.imgListComun_BN16
        .ImageList = frmPpal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    For I = 1 To 2
        imgppal(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    CargaFiltros
    
    If vParam.autocoste Then
        Text2(0).Left = 2580
        Text2(1).Left = 4230
        Text2(2).Left = 810
        Text2(0).Width = 1665
        Text2(1).Width = 1665
        Text2(2).Width = 1665
        Label1(3).Left = 4230
        Label1(4).Left = 810
    Else
        Text2(0).Left = 2580
        Text2(1).Left = 4580
        Text2(2).Left = 475
        Text2(0).Width = 2000
        Text2(1).Width = 2000
        Text2(2).Width = 2000
        Label1(3).Left = 4580
        Label1(4).Left = 475
        
        Label2(3).Visible = False
        Text3(3).Visible = False
        Me.frameextras.Width = 9660
    End If
    
    Caption = "Introducci�n de Asientos"
    
    NumTabMto = 1
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han ll�nies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "hcabapu"
    Ordenacion = " ORDER BY numasien"
    '************************************************
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    '***** canviar el nom de la PK de la cap�alera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numasien='-1'"
    Data1.Refresh
       
    
    ModoLineas = 0
    DiarioPorDefecto = ""
       
    CargarColumnas
    
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    'Maxima longitud cuentas
    txtAux(3).MaxLength = vEmpresa.DigitosUltimoNivel
    txtAux(5).MaxLength = vEmpresa.DigitosUltimoNivel
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

    Columnas = "C�digo|Nombre|Documento|Id|Tipo|"
    Ancho = "1000|5450|0|0|0|"
    'vwColumnRight =1  left=0   center=2
    Alinea = "0|0|0|0|0|"
    'Formatos
    Formato = "|||||"
    Ncol = 5

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
'    Me.chkAbonos(0).Value = 0
    
    lw1.ListItems.Clear

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
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
    DespalzamientoVisible B And (Data1.Recordset.RecordCount > 1)
    Toolbar1.Buttons(8).Enabled = B
    
    B = (Modo = 2) Or Modo = 0
    
    For I = 0 To Text1.Count - 1
        Text1(I).Locked = B
        If Modo <> 1 Then
            Text1(I).BackColor = vbWhite
        End If
    Next I
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
    
    If ASIENTO <> "" Then
        cmdRegresar.Visible = Not B
    End If
       
    PonerOpcionesMenuGeneral Me
    PonerModoUsuarioGnral Modo, "ariconta"
    
    B = (Modo < 5)
    chkVistaPrevia.Visible = B
    frameextras.Visible = True 'B
    
    Text1(0).Enabled = (Modo = 1)
    
    
    B = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
            
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 1, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    
    DataGridAux(1).Enabled = B
        
    'lineas de asiento
    Dim anc As Single
    anc = DataGridAux(1).Top
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
    
    'lineas de documentos
    B = (Modo = 5) And (NumTabMto = 0) And (ModoLineas <> 3)
    For I = 0 To txtaux3.Count - 1
        If (I >= 0 And I <= 3) Or (I >= 6 And I <= 7) Then
            txtaux3(I).Enabled = False
            txtaux3(I).Visible = False
        Else
            txtaux3(I).Enabled = B
            txtaux3(I).Visible = B
        End If
    Next I
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).BackColor = vbWhite
    Next I
    For I = 0 To txtaux3.Count - 1
        txtaux3(I).BackColor = vbWhite
    Next I
    imgppal(2).Enabled = (Data1.Recordset.RecordCount <> 0)
    
    FrameCreacion.Enabled = (Modo = 1)
    
    
    PonerModoUsuarioGnral Modo, "ariconta"

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
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
        Case 0
            tabla = "hcabapu_fichdocs"
            Sql = "SELECT hcabapu_fichdocs.codigo, hcabapu_fichdocs.campo, hcabapu_fichdocs.numdiari, hcabapu_fichdocs.fechaent, hcabapu_fichdocs.numasien, hcabapu_fichdocs.descripfich, hcabapu_fichdocs.orden, hcabapu_fichdocs.docum"
            Sql = Sql & " FROM " & tabla
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "hcabapu", "hcabapu_fichdocs")
            Else
                Sql = Sql & " WHERE hcabapu_fichdocs.codigo = -1"
            End If
            Sql = Sql & " ORDER BY orden"
            
       
       
       Case 1 ' lineas de asiento
            tabla = "hlinapu"
            Sql = "SELECT hlinapu.numasien, hlinapu.fechaent, hlinapu.numdiari, hlinapu.linliapu, hlinapu.codmacta, cuentas.nommacta, hlinapu.numdocum, hlinapu.ctacontr,"
            Sql = Sql & " hlinapu.codconce, hlinapu.ampconce, hlinapu.timporteD, hlinapu.timporteH, hlinapu.codccost, cuentas_1.nommacta as nommactactr, conceptos.nomconce, ccoste.nomccost, hlinapu.idcontab "
            Sql = Sql & " FROM (((hlinapu LEFT JOIN cuentas AS cuentas_1 ON hlinapu.ctacontr = "
            Sql = Sql & " cuentas_1.codmacta) LEFT JOIN ccoste ON hlinapu.codccost = ccoste.codccost)            "
            Sql = Sql & " INNER JOIN cuentas ON hlinapu.codmacta = cuentas.codmacta) "
            Sql = Sql & " INNER JOIN conceptos ON hlinapu.codconce = conceptos.codconce "
            If Enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "hcabapu", "hlinapu")
            Else
                Sql = Sql & " WHERE hlinapu.linliapu = '-1'"
            End If
            Sql = Sql & " ORDER BY 1,2,3,4"
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function

Private Sub frmAsi_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "numasien = " & RecuperaValor(CadenaSeleccion, 2) & " and fechaent = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " and numdiari = " & RecuperaValor(CadenaSeleccion, 3)
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmAsiP_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "numaspre = " & RecuperaValor(CadenaSeleccion, 1)
        
        
        ' Llamamos a un formulario para introducir los importes e importarlo al asiento
        frmAsiLinAdd.TotalLineas = RecuperaValor(CadenaSeleccion, 1)
        frmAsiLinAdd.Show vbModal
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.cmdAux(0).Tag + 2)
    txtAux(Indice).Text = Format(vFecha, "dd/mm/yyyy")
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
            If cmdAux(0).Tag = "0" Then txtAux(4).Text = ""
            Exit Sub
        End If
    End If
    If cmdAux(0).Tag = 0 Then
        'Cuenta normal
        txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 2)
        
        'Habilitaremos el ccoste
        HabilitarCentroCoste
        
    Else
        'contrapartida
        txtAux(6).Text = RecuperaValor(CadenaSeleccion, 1)
        Text3(5).Text = RecuperaValor(CadenaSeleccion, 2)
    End If

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    'Centro de coste
    txtAux(11).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(3).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
Dim RC As Byte
    'Concepto
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1)
    Text3(4).Text = RecuperaValor(CadenaSeleccion, 2)
    txtAux(8).Text = RecuperaValor(CadenaSeleccion, 2) & " "
    'Habilitamos importes
    RC = CByte(Val(RecuperaValor(CadenaSeleccion, 3)))
    HabilitarImportes RC
End Sub


Private Sub frmDi_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text4.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    If IndCodigo = 0 Then
        Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
    Else
        Text1(6).Text = Format(vFecha, "dd/mm/yyyy")
    End If
End Sub

Private Sub frmUtil_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion = "" Then
        ' no hacemos nada
    Else
        cboFiltro.ListIndex = 0
        
        Sql = "Select * from hcabapu "
        Sql = Sql & " WHERE numasien = " & RecuperaValor(CadenaSeleccion, 1)
        Sql = Sql & " AND numdiari =" & RecuperaValor(CadenaSeleccion, 3)
        Sql = Sql & " AND fechaent= '" & Format(RecuperaValor(CadenaSeleccion, 2), FormatoFecha) & "'"
        
        CadenaConsulta = Sql
        PonerCadenaBusqueda
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    
    If (Modo = 2 Or Modo = 5 Or Modo = 0) And (Index <> 2) Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        IndCodigo = 0
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(1).Text <> "" Then frmF.Fecha = CDate(Text1(1).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco Text1(1)
        
    Case 1
        'Tipos diario
        Set frmDi = New frmTiposDiario
        frmDi.DatosADevolverBusqueda = "0"
        frmDi.Show vbModal
        Set frmDi = Nothing
        PonFoco Text1(2)
        
    Case 2
        ' observaciones
        Screen.MousePointer = vbDefault
        
        Indice = 3
        
        Set frmZ = New frmZoom
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
        frmZ.Caption = "Observaciones Asientos"
        frmZ.Show vbModal
        Set frmZ = Nothing
    
    Case 3 ' fecha de creacion
        IndCodigo = 3
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Text1(6).Text <> "" Then frmF.Fecha = CDate(Text1(6).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco Text1(6)
    
        
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub lw1_DblClick()
    ImprimirImagen
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
        PonFoco Text1(0) ' <===
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
    
    HacerBusqueda2
    
End Sub

Private Sub HacerBusqueda2()

    CargarSqlFiltro
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia
    ElseIf CadB <> "" Or CadB1 <> "" Or cadFiltro <> "" Then
        CadenaConsulta = "select distinct hcabapu.* from " & NombreTabla & " LEFT JOIN hlinapu ON hcabapu.numdiari = hlinapu.numdiari and hcabapu.fechaent = hlinapu.fechaent and hcabapu.numasien = hlinapu.numasien "
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
    
    CargaDatosLW

End Sub



Private Sub MandaBusquedaPrevia()
Dim cWhere As String
Dim cWhere1 As String
    
    cWhere1 = ""
    
    cWhere = "(numdiari, fechaent, numasien) in (select hcabapu.numdiari,hcabapu.fechaent,hcabapu.numasien from "
    cWhere = cWhere & "hcabapu INNER JOIN hlinapu ON hcabapu.numdiari = hlinapu.numdiari and hcabapu.fechaent = hlinapu.fechaent and hcabapu.numasien = hlinapu.numasien "
    cWhere = cWhere & " WHERE (1=1) "
    
    If CadB <> "" Then cWhere1 = cWhere1 & " and " & CadB & " "
    If CadB1 <> "" Then cWhere1 = cWhere1 & " and " & CadB1 & " "
    If cadFiltro <> "" Then cWhere1 = cWhere1 & " and " & cadFiltro & " "
    
    If Trim(cWhere1) <> "and (1=1)" Then
        cWhere = cWhere & cWhere1 & ")"
    Else
        cWhere = ""
    End If
    
     Set frmAsi = New frmAsientosHcoPrev
     
     frmAsi.DatosADevolverBusqueda = "0|1|2|"
     frmAsi.cWhere = cWhere
     frmAsi.Show vbModal
     
     Set frmAsi = Nothing
     
        
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer
    
    Unload Me
    
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
    PonerModo 3
    
    If DiarioPorDefecto <> "" Then
        Text1(2).Text = RecuperaValor(DiarioPorDefecto, 1)
        Text4.Text = RecuperaValor(DiarioPorDefecto, 2)
    End If
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    Text1(4).Text = Format(Now, "dd/mm/yyyy hh:mm:ss")
    Text1(5).Text = vUsu.Login
    Text1(6).Text = "ARICONTA 6: Introducci�n de Asientos"
    PonFoco Text1(1)
    ' ***********************************************************
    
End Sub


Private Sub BotonModificar()

    If Not Me.AdoAux(1).Recordset.EOF Then
        If Not SePuedeModificarAsiento(True) Then Exit Sub
    End If

    PonerModo 4

    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonFoco Text1(1)
    ' *********************************************************
End Sub


Private Sub BotonEliminar(EliminarDesdeActualizar As Boolean)
    Dim I As Integer
    Dim Mc As Contadores
    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not Me.AdoAux(1).Recordset.EOF Then
        If Not SePuedeModificarAsiento(True) Then Exit Sub
    End If
    
    
     'Comprobamos que la fecha es de ejerccio actual
    If Not AmbitoDeFecha(True) Then Exit Sub
       
    
    If Not EliminarDesdeActualizar Then
'        If BloqAsien Then Exit Sub  'Bloqueamos el asiento, para ver si no esta bloqueado por otro
        '### a mano
        Sql = "Cabecera de apuntes." & vbCrLf
        Sql = Sql & "-----------------------------" & vbCrLf & vbCrLf
        Sql = Sql & "Va a eliminar el asiento:"
        Sql = Sql & vbCrLf & "N� Asiento   :   " & Data1.Recordset.Fields(2)
        Sql = Sql & vbCrLf & "Fecha        :   " & CStr(Data1.Recordset.Fields(1))
        Sql = Sql & vbCrLf & "Diario           :   " & Text1(2).Text & " - " & Text4.Text & vbCrLf & vbCrLf
        Sql = Sql & "      �Desea continuar ? "
        I = MsgBox(Sql, vbQuestion + vbYesNoCancel)
        'Borramos
        If I <> vbYes Then
'            DesBloqAsien
            Exit Sub
        End If
        'Hay que eliminar
        On Error GoTo Error2
        If Not Eliminar Then Exit Sub
    End If
    
    'Devolvemos contador, si no estamos actualizando
    If Not ActualizandoAsiento Then
        I = FechaCorrecta2(CDate(Data1.Recordset.Fields(1)))
        Set Mc = New Contadores
        NumRegElim = Data1.Recordset.Fields(2)
        Mc.DevolverContador "0", I = 0, NumRegElim
        Set Mc = Nothing
    End If
    
    NumRegElim = Data1.Recordset.AbsolutePosition
    DataGridAux(1).Enabled = False
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid 1, False
        PonerModo 0
        Else
            If NumRegElim > Data1.Recordset.RecordCount Then
                Data1.Recordset.MoveLast
            Else
                Data1.Recordset.MoveFirst
                Data1.Recordset.Move NumRegElim - 1
            End If
            PonerCampos
            DataGridAux(1).Enabled = True
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Not EliminarDesdeActualizar Then
        Else
           If VieneDeDesactualizar Then
                PulsadoSalir = True
                Unload Me
           End If
        End If
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub


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
    
    Text4.Text = ""
    If Text1(2).Text <> "" Then Text4.Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", Text1(2).Text, "N")

    CargaDatosLW

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
End Sub


Private Sub cmdCancelar_Click()
Dim I As Integer
Dim v
Dim Mc As Contadores
    Select Case Modo
        Case 1, 3 'B�squeda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la cap�alera ***
                PonFoco Text1(0)
                ' ***************************************************

        Case 4  'Modificar
                lblIndicador.Caption = ""
                TerminaBloquear
                PonerModo 2
                PonerCampos
                PonFoco Text1(0)
                
        Case 5 'LL�NIES
            TerminaBloquear
            LlevaContraPartida = False
        
            If ModoLineas = 1 Then 'INSERTAR
                ModoLineas = 0
                DataGridAux(1).AllowAddNew = False
                If Not AdoAux(1).Recordset.EOF Then
                    AdoAux(1).Recordset.MoveFirst
                End If
            End If
            ModoLineas = 0
            LLamaLineas 1, 0, 0
            PonerModo 2
            DataGridAux(1).Enabled = True
            If Not Data1.Recordset.EOF Then _
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
            'Habilitar las opciones correctas del menu segun Modo
            DataGridAux(1).Enabled = True
            PonerFocoGrid DataGridAux(1)
        
        
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Sql As String
Dim cad As String

    On Error GoTo EDatosOK

    DatosOK = False
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
            MsgBox varTxtFec, vbExclamation
        Else
            MsgBox "La fecha no pertenece al ejercicio actual ni al siguiente", vbExclamation
        End If
        B = False

    End If
    
    DatosOK = B

EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(numasien=" & Trim(Text1(0).Text) & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N") & ") "
    
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
Dim Sql As String
Dim SqlAux As String
Dim Rs As ADODB.Recordset

    On Error GoTo FinEliminar

    Conn.BeginTrans
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE (numasien=" & Trim(Text1(0).Text) & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N") & ") "
        ' ***********************************************************************
        
        
    'El LOG
    Sql = "N� Asiento : " & Data1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Fecha      : " & CStr(Data1.Recordset.Fields(1))
    Sql = Sql & vbCrLf & "Diario     : " & Text1(2).Text & " - " & Text4.Text & vbCrLf & vbCrLf
    Sql = Sql & vbCrLf & RellenaABlancos("Cuenta", True, 10) & " " & RellenaABlancos("Debe", False, 14) & " " & RellenaABlancos("Haber", False, 14) & " "
    Sql = Sql & vbCrLf & String(40, "-") & vbCrLf
    
    
    SqlAux = "select * from hlinapu where numasien = " & DBSet(Data1.Recordset.Fields(2), "N")
    SqlAux = SqlAux & " and fechaent = " & DBSet(Data1.Recordset.Fields(1), "F")
    SqlAux = SqlAux & " and numdiari = " & DBSet(Text1(2).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlAux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Sql = Sql & RellenaABlancos(DBLet(Rs!codmacta, "T"), True, 10) & " " & RellenaABlancos(Format(DBLet(Rs!timported, "N"), "###,###,##0.00"), False, 14) & " " & RellenaABlancos(Format(DBLet(Rs!timporteH, "N"), "###,###,##0.00"), False, 14) & vbCrLf
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    vLog.Insertar 2, vUsu, Sql
        
        
        
        
        
    Conn.Execute "DELETE FROM hlinapu " & vWhere
    
    Conn.Execute "DELETE FROM hcabapu_fichdocs " & vWhere

'    ' *******************************
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
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim RC As Byte

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    If Text1(Index).Text = "" Then Exit Sub
    
    If Modo = 5 Then Exit Sub
    
    Select Case Index
        Case 1 'fecha de entrada
            Sql = ""
            If Not EsFechaOK(Text1(1)) Then
                MsgBox "Fecha incorrecta. (dd/mm/yyyy)", vbExclamation
                Sql = "mal"
            Else
                RC = FechaCorrecta2(CDate(Text1(1).Text))
                Sql = ""
                If RC > 1 Then
                    If RC = 2 Then
                        Sql = varTxtFec
                    
                    Else
                        If RC = 3 Then
                            Sql = "El ejercicio al que pertenece la fecha: " & Text1(Index).Text & " est� cerrado."
                        Else
                            Sql = "Ejercicio para: " & Text1(Index).Text & " todav�a no activo"
                        End If
                    End If
                    MsgBox Sql, vbExclamation
                 Else
                    'Fecha correcta. Si tiene valor DiarioPorDefecto entonces NO paso por ese campo
                    'Y me voy directamente al siguiente
                    If DiarioPorDefecto <> "" Then PonFoco Text1(2)
                 End If
            End If
            
        Case 2 'diario
            If Not IsNumeric(Text1(2).Text) Then
                MsgBox "Tipo de diario no es num�rico: " & Text1(2).Text, vbExclamation
                Text1(2).Text = ""
                Text4.Text = ""
                PonFoco Text1(2)
                Exit Sub
            End If
             Sql = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text1(2).Text, "N")
             If Sql = "" Then
                    Sql = "Diario no encontrado: " & Text1(2).Text
                    Text1(2).Text = ""
                    Text4.Text = ""
                    MsgBox Sql, vbExclamation
                    PonFoco Text1(2)
            End If
            Text1(2).Text = Val(Text1(2))
            Text4.Text = Sql
        
        Case 6 ' fecha de creacion
            PonerFormatoFecha Text1(6)
        
    End Select
End Sub

'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 3 Then
        If KeyAscii = teclaBuscar Then
            Select Case Index
                Case 1:  KEYBusqueda KeyAscii, 0
                Case 0:  KEYBusqueda KeyAscii, 1
                Case 2:  KEYBusqueda KeyAscii, 2
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

'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index, True
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1 'Asientos descuadrados
            Screen.MousePointer = vbHourglass
            
            Set frmUtil = New frmUtilidades
            
            frmUtil.Opcion = 1
            frmUtil.Show vbModal
            
            Set frmUtil = Nothing
            
        Case 2 'Asientos con numeros descuadrados
            Screen.MousePointer = vbHourglass
            frmMensajes.Opcion = 12
            frmMensajes.Show vbModal
        
        
    End Select

End Sub

Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim LINASI As Long


    'Fuerzo que se vean las lineas
    
    Select Case Button.Index
        Case 1
            'A�ADIR linea factura
            If Not Me.AdoAux(1).Recordset.EOF Then
                If Not SePuedeModificarAsiento(True) Then Exit Sub
            End If
            BotonAnyadirLinea 1, True
        Case 2
            'MODIFICAR linea factura
            If Not SePuedeModificarAsiento(True) Then Exit Sub
            BotonModificarLinea 1
        Case 3
            'ELIMINAR linea factura
            If Not SePuedeModificarAsiento(True) Then Exit Sub
            BotonEliminarLinea 1
        Case 5
            cmdSaldoHco_Click (0)
        Case 6
            cmdSaldoHco_Click (1)
        Case 7 ' asiento predefinido
            If Not Me.AdoAux(1).Recordset.EOF Then
                If Not SePuedeModificarAsiento(True) Then Exit Sub
            End If
            
            ' Llamamos a un formulario para introducir los importes e importarlo al asiento
            NumAsiPre = ""
            Ampliacion = ""
            
            frmAsiLinAdd.TotalLineas = 0
            frmAsiLinAdd.Show vbModal
            
            
            'Si tienen algun registro tendremos
            If CadenaDesdeOtroForm <> "" Then
                Set miRsAux = New ADODB.Recordset
                
                Sql = " SELECT max(linliapu) FROM hlinapu WHERE hlinapu.numdiari= " & Data1.Recordset!NumDiari
                Sql = Sql & " AND hlinapu.fechaent= " & DBSet(Data1.Recordset!FechaEnt, "F")
                Sql = Sql & " AND hlinapu.numasien=" & Data1.Recordset!NumAsien & ";"
                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                LINASI = 0
                If Not miRsAux.EOF Then LINASI = DBLet(miRsAux.Fields(0), "N")
                miRsAux.Close
                
                Sql = "SELECT cta,nomdocum,tmpconext.timported, tmpconext.timporteh,pos ,ccost, ctacontr, codconce, numdocum, asipre_lineas.ampconce FROM tmpconext, asipre_lineas where codusu =" & vUsu.Codigo
                Sql = Sql & " and asipre_lineas.numaspre = " & DBSet(NumAsiPre, "N") & " and asipre_lineas.linlapre = tmpconext.pos ORDER BY pos"
                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Sql = ""
                While Not miRsAux.EOF
                    LINASI = LINASI + 1
                    Sql = Sql & ", (" & Data1.Recordset!NumDiari & "," & DBSet(Data1.Recordset!FechaEnt, "F") & "," & Data1.Recordset!NumAsien
                    Sql = Sql & "," & LINASI & ",'" & miRsAux!Cta & "'," & DBSet(miRsAux!Numdocum, "T") & "," & DBSet(miRsAux!codconce, "N")
                    Sql = Sql & "," & DBSet(miRsAux!Ampconce & " " & Ampliacion, "T") & "," & DBSet(miRsAux!timported, "N", "S") & "," & DBSet(miRsAux!timporteH, "N", "S")
                    Sql = Sql & "," & DBSet(miRsAux!ctacontr, "T")
                    If vParam.autocoste Then
                        Sql = Sql & "," & DBSet(miRsAux!CCost, "T") & ")"
                    Else
                        Sql = Sql & ",null)"
                    End If
                    
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                Set miRsAux = Nothing
                
                
                If Sql <> "" Then
                    Sql = Mid(Sql, 2)
                    Sql = "INSERT INTO hlinapu (numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,ctacontr,codccost) VALUES " & Sql
                    Conn.Execute Sql
                    CargaGrid 1, True
                End If
                
                
            End If
            
            Exit Sub
            
            
            
    End Select

End Sub

Private Sub ToolbarAux0_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Fuerzo que se vean las lineas
    
    Select Case Button.Index
        Case 1 ' insertar
            InsertarDesdeFichero
            
            CargaDatosLW
            
        Case 3 ' eliminar
            EliminarImagen
        
    End Select

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarCreacion_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    
        Case 1 'Informacion de creacion
            FrameCreacion.Visible = Not (FrameCreacion.Visible)
            FrameCreacion.Enabled = (FrameCreacion.Visible) And (Modo = 1)
            
        
        
    End Select


End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim SqlAux As String
Dim Rs As ADODB.Recordset

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
            Sql = "�Seguro que desea eliminar la l�nea del asiento?"
            Sql = Sql & vbCrLf & "C�digo: " & AdoAux(Index).Recordset!NumAsien & " - " & AdoAux(Index).Recordset!FechaEnt & " - " & AdoAux(Index).Recordset!NumDiari & " - " & AdoAux(Index).Recordset!Linliapu
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                
                'El LOG
                Sql = "N� Asiento : " & Data1.Recordset.Fields(2)
                Sql = Sql & vbCrLf & "Fecha      : " & CStr(Data1.Recordset.Fields(1))
                Sql = Sql & vbCrLf & "Diario     : " & Text1(2).Text & " - " & Text4.Text
                Sql = Sql & vbCrLf & "L�nea      : " & DBSet(AdoAux(Index).Recordset!Linliapu, "N") & vbCrLf & vbCrLf
                Sql = Sql & vbCrLf & RellenaABlancos("Cuenta", True, 10) & " " & RellenaABlancos("Debe", False, 14) & " " & RellenaABlancos("Haber", False, 14) & " "
                Sql = Sql & vbCrLf & String(40, "-") & vbCrLf
                
                
                SqlAux = "select * from hlinapu where numasien = " & DBSet(AdoAux(Index).Recordset!NumAsien, "N")
                SqlAux = SqlAux & " and fechaent = " & DBSet(AdoAux(Index).Recordset!FechaEnt, "F")
                SqlAux = SqlAux & " and numdiari = " & DBSet(AdoAux(Index).Recordset!NumDiari, "N")
                SqlAux = SqlAux & " and linliapu = " & DBSet(AdoAux(Index).Recordset!Linliapu, "N")
                 
                Set Rs = New ADODB.Recordset
                Rs.Open SqlAux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    Sql = Sql & RellenaABlancos(DBLet(Rs!codmacta, "T"), True, 10) & " " & RellenaABlancos(Format(DBLet(Rs!timported, "N"), "###,###,##0.00"), False, 14) & " " & RellenaABlancos(Format(DBLet(Rs!timporteH, "N"), "###,###,##0.00"), False, 14) & vbCrLf
                    Rs.MoveNext
                Wend
                Set Rs = Nothing
                
                vLog.Insertar 4, vUsu, Sql
                
                
                
                Sql = "DELETE FROM hlinapu "
                Sql = Sql & Replace(vWhere, "hcabapu", "hlinapu") & " and linliapu = " & DBLet(AdoAux(Index).Recordset!Linliapu, "N")
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute Sql
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
        Case 1: vTabla = "hlinapu"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 1   'hlinapu
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = ""
            NumF = SugerirCodigoSiguienteStr(vTabla, "linliapu", Replace(vWhere, "hcabapu", "hlinapu"))
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)

            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 230
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'lineas de asiento
                    If Limpia Then
                        For I = 0 To txtAux.Count - 1
                            txtAux(I).Text = ""
                        Next I
                    End If
                    txtAux(0).Text = Text1(0).Text 'asiento
                    txtAux(1).Text = Text1(1).Text 'fecha
                    txtAux(2).Text = Text1(2).Text 'diario
                    
                    txtAux(3).Text = Format(NumF, "000") 'linea contador
                    If Limpia Then
                        txtAux2(4).Text = ""
                        Text3(3).Text = ""
                        Text3(4).Text = ""
                        Text3(5).Text = ""
                    End If
                    
                    If Limpia Then
                        PonFoco txtAux(4)
                    Else
                        PonFoco txtAux(5)
                    End If
            End Select

    End Select
End Sub


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

    Select Case Index
        Case 0, 1 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If

            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

    End Select

    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 'asientos
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux(3).Text = DataGridAux(Index).Columns(3).Text
            
            txtAux(4).Text = DataGridAux(Index).Columns(4).Text 'cuenta
            txtAux2(4).Text = DataGridAux(Index).Columns(5).Text 'denominacion
            txtAux(5).Text = DataGridAux(Index).Columns(6).Text 'documento
            txtAux(6).Text = DataGridAux(Index).Columns(7).Text 'contrapartida
            txtAux(7).Text = DataGridAux(Index).Columns(8).Text 'concepto
            txtAux(8).Text = DataGridAux(Index).Columns(9).Text 'ampliacion
            txtAux(9).Text = DataGridAux(Index).Columns(10).Text 'importe al debe
            txtAux(10).Text = DataGridAux(Index).Columns(11).Text 'importe al haber
            txtAux(11).Text = DataGridAux(Index).Columns(12).Text 'centro de coste
            
            CtaAnt = txtAux(4).Text
            DebeAnt = txtAux(9).Text
            HaberAnt = txtAux(10).Text
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    
    PonFoco txtAux(4)
    
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
        Case 1 'lineas de asiento
            For jj = 4 To txtAux.Count - 1
                txtAux(jj).Visible = B
                txtAux(jj).Top = alto
            Next jj
            
            txtAux2(4).Visible = B
            txtAux2(4).Top = alto

            For jj = 0 To cmdAux.Count - 1
                cmdAux(jj).Visible = B
                cmdAux(jj).Top = txtAux(4).Top
                cmdAux(jj).Height = txtAux(4).Height
            Next jj
            
            If Not vParam.autocoste Then
                cmdAux(3).Visible = False
                cmdAux(3).Enabled = False
                txtAux(11).Visible = False
                txtAux(11).Enabled = False
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
        If txtAux(4).Text = "" Then
            MsgBox "Cuenta no puede estar vacia.", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(4)
            Exit Function
        End If
        
        If Not IsNumeric(txtAux(4).Text) Then
            MsgBox "Cuenta debe ser numrica", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(4)
            Exit Function
        End If
        
        If txtAux(4).Text = NO Then
            MsgBox "La cuenta debe estar dada de alta en el sistema", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(4)
            Exit Function
        End If
        
        If Not EsCuentaUltimoNivel(txtAux(4).Text) Then
            MsgBox "La cuenta no es de �ltimo nivel", vbExclamation
            DatosOkLlin = False
            PonFoco txtAux(4)
            Exit Function
        End If
        
        'Centro de coste
        If txtAux(11).Visible Then
            If txtAux(11).Enabled Then
                If txtAux(11).Text = "" Then
                    MsgBox "Centro de coste no puede ser nulo", vbExclamation
                    PonFoco txtAux(11)
                    Exit Function
                End If
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
Dim I As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 1 'APUNTES
                If DataGridAux(Index).Columns.Count > 2 Then
                    Text3(5).Text = DBLet(AdoAux(1).Recordset!nommactactr, "T")
                    Text3(4).Text = DBLet(AdoAux(1).Recordset!NomConce, "T")
                    Text3(3).Text = DBLet(AdoAux(1).Recordset!nomccost, "T")
                End If
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub


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
            
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(4)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(4)|T|Denominaci�n|3995|;"
            tots = tots & "S|txtaux(5)|T|Documento|1905|;S|txtaux(6)|T|Contrapartida|1425|;S|cmdAux(1)|B|||;"
            tots = tots & "S|txtaux(7)|T|Cto|465|;S|cmdAux(2)|B|||;S|txtaux(8)|T|Ampliaci�n|3000|;"
            If vParam.autocoste Then
                tots = tots & "S|txtaux(9)|T|Debe|1654|;S|txtaux(10)|T|Haber|1654|;S|txtaux(11)|T|CC|710|;S|cmdAux(3)|B|||;"
            Else
                tots = tots & "S|txtaux(9)|T|Debe|1989|;S|txtaux(10)|T|Haber|1989|;N||||0|;"
            End If
            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgLeft
            DataGridAux(Index).Columns(6).Alignment = dbgLeft
            
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (Enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For I = 0 To 4
                    txtAux(I).Text = ""
                Next I
                txtAux2(4).Text = ""
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
    ObtenerSumas
    
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
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    
                    DataGridAux(1).AllowAddNew = False
                    
                    If Not AdoAux(1).Recordset.EOF Then PosicionGrid = DataGridAux(1).FirstRow
                    CargaGrid 1, True
                    Limp = True

                    'Estabamos insertando insertando lineas
                    'Si ha puesto contrapartida borramos
                    If txtAux(6).Text <> "" Then
                        If EstaLaCuentaBloqueada(txtAux(6).Text, CDate(Text1(1).Text)) Then
                            LlevaContraPartida = False
                        Else
                            If LlevaContraPartida Then
                                'Ya lleva la contra partida, luego no hacemos na
                                LlevaContraPartida = False
                            Else
                                cad = "Generar asiento de la contrapartida?"
                                If MsgBox(cad, vbQuestion + vbYesNoCancel) = vbYes Then
                                    FijarContraPartida
                                    Limp = False
                                    LlevaContraPartida = True
                                End If
                            End If
                        End If
                    Else
                        LlevaContraPartida = False
                    End If
                    txtAux(11).Text = ""
                    Text3(3).Text = ""
                    If Limp Then
                        For I = 3 To 5
                            Text3(I).Text = ""
                        Next I
                        For I = 0 To 11
                            txtAux(I).Text = ""
                        Next I
                    End If
                    ModoLineas = 0
                    If B Then
                            BotonAnyadirLinea NumTabMto, Not LlevaContraPartida
                    End If
            End Select
           
            SituarTab (NumTabMto)
        End If
    End If
End Sub

Private Function CadCambios() As String
Dim Sql As String

    Sql = ""
    
    If CtaAnt <> txtAux(4).Text Then Sql = Sql & RellenaABlancos("Cuenta", True, 10) & " " & RellenaABlancos(CtaAnt, False, 14) & " " & RellenaABlancos(txtAux(4).Text, False, 14) & vbCrLf
    If DebeAnt <> txtAux(9).Text Then Sql = Sql & RellenaABlancos("Debe", True, 10) & " " & RellenaABlancos(DebeAnt, False, 14) & " " & RellenaABlancos(txtAux(9).Text, False, 14) & vbCrLf
    If HaberAnt <> txtAux(10).Text Then Sql = Sql & RellenaABlancos("Haber", True, 10) & " " & RellenaABlancos(HaberAnt, False, 14) & " " & RellenaABlancos(txtAux(10).Text, False, 14) & vbCrLf

    CadCambios = Sql
    
End Function


Private Sub ModificarLinea()
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim v As Integer
Dim cad As String
Dim Sql As String
Dim Sql2 As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'apuntes
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
        
            Sql2 = CadCambios
            If Sql2 <> "" Then
                Sql = "N� Asiento : " & Data1.Recordset.Fields(2)
                Sql = Sql & vbCrLf & "Fecha      : " & CStr(Data1.Recordset.Fields(1))
                Sql = Sql & vbCrLf & "Diario     : " & Text1(2).Text & " - " & Text4.Text
                Sql = Sql & vbCrLf & "L�nea      : " & DBSet(AdoAux(1).Recordset!Linliapu, "N") & vbCrLf & vbCrLf
                Sql = Sql & vbCrLf & RellenaABlancos("Campo", True, 10) & " " & RellenaABlancos("Valor anterior", False, 14) & " " & RellenaABlancos("Valor actual", False, 14) & " "
                Sql = Sql & vbCrLf & String(40, "-") & vbCrLf
                
                Sql = Sql & Sql2
        
                vLog.Insertar 3, vUsu, Sql
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
            
        End If
    End If
        
End Sub




Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & "hcabapu.numdiari=" & DBSet(Text1(2).Text, "N") & " and hcabapu.fechaent=" & DBSet(Text1(1).Text, "F") & " and numasien = " & DBSet(Text1(0).Text, "N")
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
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0 And Not SoloImprimir
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2) And DesdeNorma43 = 0 And Not SoloImprimir
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2) And DesdeNorma43 = 0 And Not SoloImprimir
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0 And Not SoloImprimir
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0 And Not SoloImprimir
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N")
        
        Me.Toolbar2.Buttons(1).Enabled = DBLet(Rs!especial, "N") And DesdeNorma43 = 0 And Not SoloImprimir
        Me.Toolbar2.Buttons(2).Enabled = DBLet(Rs!especial, "N") And DesdeNorma43 = 0 And Not SoloImprimir
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2) And Not SoloImprimir
        ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 2 And Me.AdoAux(1).Recordset.RecordCount > 0) And Not SoloImprimir
        ToolbarAux.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.AdoAux(1).Recordset.RecordCount > 0) And Not SoloImprimir
        ToolbarAux.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And ((Modo = 2 And Me.AdoAux(1).Recordset.RecordCount > 0) Or (Modo = 5)) And DesdeNorma43 = 0 And Not SoloImprimir
        ToolbarAux.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And ((Modo = 2 And Me.AdoAux(1).Recordset.RecordCount > 0) Or (Modo = 5)) And DesdeNorma43 = 0 And Not SoloImprimir
        ToolbarAux.Buttons(7).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2) And DesdeNorma43 = 0 And Not SoloImprimir
        
        ToolbarAux0.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0) And DesdeNorma43 = 0 And Not SoloImprimir
        ToolbarAux0.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.Data1.Recordset.RecordCount > 0) And DesdeNorma43 = 0 And Not SoloImprimir
        
        
        vUsu.LeerFiltros "ariconta", IdPrograma
        
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    AntiguoText1 = txtAux(Index).Text
    ConseguirFoco txtAux(Index), Modo
    
    If Index = 8 Then txtAux(Index).SelStart = Len(txtAux(Index).Text)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        'Esto sera k hemos pulsado el ENTER
        txtAux_LostFocus Index
        cmdAceptar_Click
    Else
        If KeyCode = 113 Then
            'Esto sera k pedimos la calculadora
            PideCalculadora
        Else
            'Ha pulsado F5. Ponemos linea anterior
            Select Case KeyCode
            Case 116
                PonerLineaAnterior (Index)
                
            Case 117
                'F6
                'Si es el primer campo , y ha pulsado f6
                'cogera la linea de arriba y la pondra en los txtaux
                If Not AdoAux(1).Recordset Is Nothing Then
                    If Not AdoAux(1).Recordset.EOF Then
                        Screen.MousePointer = vbHourglass
                        HacerF6
                        Screen.MousePointer = vbDefault
                    End If
                End If
                
            Case Else
                If (Shift And vbCtrlMask) > 0 Then
                    If UCase(Chr(KeyCode)) = "B" Then
                        'OK. Ha pulsado Control + B
                        '----------------------------------------------------
                        '----------------------------------------------------
                        '
                        ' Dependiendo de index lanzaremos una opcion uotra
                        '
                        '----------------------------------------------------
                        
                        'De momento solo para el 5. Cliente
                        Select Case Index
                        Case 4
                            txtAux(4).Text = ""
                            Image1_Click 1
                        Case 8
                            txtAux(8).Text = ""
                            Image1_Click 2
                        End Select
                     End If
                End If
            End Select
        End If
    End If
End Sub

'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 4:  KEYImage KeyAscii, 3
            Case 6:  KEYImage KeyAscii, 0
            Case 7:  KEYImage KeyAscii, 1
            Case 11:  KEYImage KeyAscii, 2
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub
'++


Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    'Cta contrapartida
    cmdAux(0).Tag = 1
    LlamaContraPar
    PonFoco txtAux(6)
Case 1
    Set frmCon = New frmConceptos
    frmCon.DatosADevolverBusqueda = "0|"
    frmCon.Show vbModal
    Set frmCon = Nothing
Case 2
    If txtAux(11).Enabled Then
        Set frmCC = New frmCCCentroCoste
        frmCC.DatosADevolverBusqueda = "0|1|"
        frmCC.Show vbModal
        Set frmCC = Nothing
    End If
Case 3
    'Como si hubeiran pulsado sobre el cmd +
    cmdAux(0).Tag = 0
    LlamaContraPar
    PonFoco txtAux(6)
End Select
End Sub


Private Sub txtaux_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo <> 1 Then
        If KeyCode = 107 Or KeyCode = 187 Then
                KeyCode = 0
                LanzaPantalla Index
        End If
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Importe As Currency
        
        If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
        
        If txtAux(Index).Text = AntiguoText1 Then
             Exit Sub
        End If
    
        'Comun a todos
        If txtAux(Index).Text = "" Then
            Select Case Index
            Case 4
                HabilitarCentroCoste
                txtAux(1).Text = ""
            Case 6
                Text3(5).Text = ""
            Case 9
                HabilitarImportes 0
            End Select
            Exit Sub
        End If
        
        Select Case Index
        Case 4
            RC = txtAux(4).Text
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                txtAux(4).Text = RC
                If Modo = 1 Then Exit Sub
                If EstaLaCuentaBloqueada(RC, CDate(Text1(1).Text)) Then
                    MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                    txtAux(4).Text = ""
                Else
                    txtAux2(4).Text = Sql
                    RC = ""
                End If
            Else
                If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                    If vUsu.PermiteOpcion("ariconta", 201, vbOpcionCrearEliminar) Then
                        'NO EXISTE LA CUENTA
                        Sql = Sql & " �Desea crearla?"
                        If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
                            CadenaDesdeOtroForm = RC
                            cmdAux(0).Tag = Index
                            Set frmC = New frmColCtas
                            frmC.DatosADevolverBusqueda = "0|1|"
                            frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                            frmC.Show vbModal
                            Set frmC = Nothing
                            If txtAux(4).Text = RC Then Sql = "" 'Para k no los borre
                        End If
                    Else
                        MsgBox Sql, vbExclamation
                    End If
                Else
                    MsgBox Sql, vbExclamation
                End If
                    
                If Sql <> "" Then
                  txtAux(4).Text = ""
                  txtAux2(4).Text = ""
                  RC = "NO"
                End If
            End If
            HabilitarCentroCoste
            If RC <> "" Then PonFoco txtAux(4)
                
            If Modo = 5 And ModoLineas = 1 Then MostrarObservaciones txtAux(Index)
            
        Case 6
        
            'Contrapartida
        
            RC = txtAux(6).Text
            If CuentaCorrectaUltimoNivel(RC, Sql) Then
                txtAux(6).Text = RC
                Text3(5).Text = Sql
            Else
            
                If InStr(1, Sql, "No existe la cuenta :") > 0 Then
                    'NO EXISTE LA CUENTA
                    Sql = Sql & " �Desea crearla?"
                    If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
                        CadenaDesdeOtroForm = RC
                        cmdAux(0).Tag = Index
                        Set frmC = New frmColCtas
                        frmC.DatosADevolverBusqueda = "0|1|"
                        frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                        frmC.Show vbModal
                        Set frmC = Nothing
                        If txtAux(6).Text = RC Then Sql = "" 'Para k no los borre
                    End If
                Else
                    MsgBox Sql, vbExclamation
                End If
                If Sql <> "" Then
                    txtAux(6).Text = ""
                    Text3(5).Text = ""
                    PonFoco txtAux(6)
                End If
            End If
            
        Case 7
             If Not IsNumeric(txtAux(7).Text) Then
                    MsgBox "El concepto debe de ser num�rico", vbExclamation
                    PonFoco txtAux(7)
                    Exit Sub
                End If
                
                If Val(txtAux(7).Text) >= 900 Then
                    If vUsu.Nivel > 1 Then
                        MsgBox "Los conceptos superiores a 900 se los reserva la aplicaci�n.", vbExclamation
                        Text3(4).Text = ""
                        txtAux(7).Text = ""
                        PonFoco txtAux(7)
                        Exit Sub
                    Else
                        If Me.Tag = "" Then
                            MsgBox "Los conceptos superiores a 900 se los reserva la aplicaci�n.", vbExclamation
                            Me.Tag = "0"
                        End If
                    End If
                End If
                
                
                
                CadenaAmpliacion = ""
                If Text3(4).Text <> "" Then
                    'Tenia concepto anterior
                    If InStr(1, txtAux(8).Text, Text3(4).Text) > 0 Then CadenaAmpliacion = Trim(Mid(txtAux(8).Text, Len(Text3(4).Text) + 1))
                End If
                
                RC = "tipoconce"
                Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtAux(7).Text, "N", RC)
                If Sql = "" And RC = "tipoconce" Then
                    MsgBox "Concepto NO encontrado: " & txtAux(7).Text, vbExclamation
                    txtAux(7).Text = ""
                    RC = "0"
                End If
                HabilitarImportes CByte(Val(RC))
                Text3(4).Text = Sql
                txtAux(8).Text = Sql
                If txtAux(8).Text <> "" Then txtAux(8).Text = txtAux(8).Text & " "
                txtAux(8).Text = txtAux(8).Text & CadenaAmpliacion
                If RC = "0" Then PonFoco txtAux(7)
                
        Case 9, 10
                'LOS IMPORTES
                If Not EsNumerico(txtAux(Index).Text) Then
                    MsgBox "Importes deben ser num�ricos.", vbExclamation
                    On Error Resume Next
                    txtAux(Index).Text = ""
                    PonFoco txtAux(Index)
                    Exit Sub
                End If
                
                
                'Es numerico
                Sql = TransformaPuntosComas(txtAux(Index).Text)
                If CadenaCurrency(Sql, Importe) Then
                    txtAux(Index).Text = Format(Importe, "0.00")
                    'Ponemos el otro campo a ""
                    If Index = 9 Then
                        txtAux(10).Text = ""
                    Else
                        txtAux(9).Text = ""
                    End If
                End If
                
                
                
        Case 11
                txtAux(11).Text = UCase(txtAux(11).Text)
                Sql = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtAux(11).Text, "T")
                If Sql = "" Then
                    MsgBox "Concepto NO encontrado: " & txtAux(11).Text, vbExclamation
                    txtAux(11).Text = ""
                    PonFoco txtAux(11)
                End If
                Text3(3).Text = Sql
                
        End Select
End Sub

Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtAux(4).Text <> "" Then
            hab = HayKHabilitarCentroCoste(txtAux(4).Text)
        Else
            txtAux(11).Text = ""
        End If
        If hab Then
            txtAux(11).BackColor = &H80000005
            Else
            txtAux(11).BackColor = &H80000018
            txtAux(11).Text = ""
        End If
    End If
    txtAux(11).Enabled = hab
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
            frmAsientosHcoList.NumAsien = Text1(0).Text
            frmAsientosHcoList.NumDiari = Text1(2).Text
            frmAsientosHcoList.FechaEnt = Text1(1).Text
            
            frmAsientosHcoList.Show vbModal

    End Select
End Sub


Private Function Modificar() As Boolean
Dim B1 As Boolean
Dim VC As Contadores

    On Error GoTo EModificar
         
        Modificar = False
     
        '-----------------------------------------------
        ' ABRIL 2006
        '
        ' Si cambia de ejercicio le ofertaremos un nuevo numero de ASIENTO
        '
        B1 = False
        If Data1.Recordset!FechaEnt <> CDate(Text1(1).Text) Then
            'HAN CAMBIADO DE FECHA
            
            
            Sql = ""
            'Estabamos(pasado) en ejercicio actual
            If Data1.Recordset!FechaEnt <= vParam.fechafin Then Sql = "A"
                
                
            B1 = False 'Hay que preguntar cambio de contador. De momento NO
            If CDate(Text1(1).Text) <= vParam.fechafin Then
                'La nueva fecha es del actual
                'Si la otra era del siguiente hay que preguntar
                If Sql = "" Then B1 = True
            Else
                If Sql <> "" Then B1 = True
            End If
            
            If B1 Then
                Sql = "Ha cambiado de ejercicios la fecha del asiento." & vbCrLf & " �Desea obtener nuevo numero de asiento?"
                Sql = MsgBox(Sql, vbQuestion + vbYesNoCancel)
                If CByte(Sql) = vbCancel Then Exit Function
                
                If CByte(Sql) = vbNo Then B1 = False
                
            End If
        End If
        Set VC = New Contadores
        If B1 Then
            'Obtengo nuevo contador
            If VC.ConseguirContador("0", (CDate(Text1(1).Text) <= vParam.fechafin), False) > 0 Then Exit Function
        Else
            VC.Contador = Data1.Recordset!NumAsien
        End If
                    
                    
        Conn.BeginTrans
        'Comun
        
        Conn.Execute "set foreign_key_checks = 0"
        
        
        Sql = " WHERE  numdiari=" & Data1.Recordset!NumDiari
        Sql = Sql & " AND fechaent='" & Format(Data1.Recordset!FechaEnt, FormatoFecha)
        Sql = Sql & "' AND numasien=" & Data1.Recordset!NumAsien
        
        'BLoqueamos
        Conn.Execute "Select * from hcabapu " & Sql & " FOR UPDATE"
        
        'A�adimos tb el nunmero de asiento
        Sql = " numasien = " & VC.Contador & " , numdiari= " & Text1(2).Text & " , fechaent ='" & Format(Text1(1).Text, FormatoFecha) & "'" & Sql
        
        
       'Las lineas de apuntes
        Conn.Execute "UPDATE hlinapu SET " & Sql
      
        
        'Modificamos la cabecera
        If Text1(3).Text = "" Then
            Sql = "obsdiari = NULL," & Sql
        Else
            Sql = "Obsdiari ='" & DevNombreSQL(Text1(3).Text) & "'," & Sql
        End If

        Conn.Execute "UPDATE hcabapu SET " & Sql
        
        ' tema del log
        If Data1.Recordset!FechaEnt <> CDate(Text1(1).Text) Then
            Sql = "N� Asiento : " & Data1.Recordset.Fields(2)
            Sql = Sql & vbCrLf & "Fecha      : " & CStr(Data1.Recordset.Fields(1))
            Sql = Sql & vbCrLf & "Diario     : " & Text1(2).Text & " - " & Text4.Text & vbCrLf & vbCrLf
            
            Sql = Sql & vbCrLf & "Nueva Fecha: " & Text1(1).Text
            
            vLog.Insertar 1, vUsu, Sql
        
        End If
  
  
  
  
  
  
EModificar:
        Conn.Execute "set foreign_key_checks = 1"
        If Err.Number <> 0 Then
            MuestraError Err.Number
            Conn.RollbackTrans
            Modificar = False
            B1 = False
        Else
            Conn.CommitTrans
            Modificar = True
        End If
        
        'Si habia que devolver contador
        If B1 Then
            Text1(0).Text = VC.Contador
            Set VC = Nothing
            Set VC = New Contadores
            VC.DevolverContador "0", (Data1.Recordset!FechaEnt <= vParam.fechafin), Data1.Recordset!NumAsien
            
        End If
        Set VC = Nothing
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
    
End Function


Private Sub ObtenerSumas()
    Dim Deb As Currency
    Dim hab As Currency
    Dim Rs As ADODB.Recordset
    
    Text2(0).Text = "": Text2(1).Text = "": Text2(2).Text = ""
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If AdoAux(1).Recordset Is Nothing Then Exit Sub
    
    If AdoAux(1).Recordset.EOF Then Exit Sub
    
    
    Set Rs = New ADODB.Recordset
    Sql = "SELECT Sum(hlinapu.timporteD) AS SumaDetimporteD, Sum(hlinapu.timporteH) AS SumaDetimporteH"
    Sql = Sql & " ,hlinapu.numdiari,hlinapu.fechaent,hlinapu.numasien"
    Sql = Sql & " From hlinapu GROUP BY hlinapu.numdiari, hlinapu.fechaent, hlinapu.numasien "
    Sql = Sql & " HAVING (((hlinapu.numdiari)=" & Data1.Recordset!NumDiari
    Sql = Sql & ") AND ((hlinapu.fechaent)='" & Format(Data1.Recordset!FechaEnt, FormatoFecha)
    Sql = Sql & "') AND ((hlinapu.numasien)=" & Data1.Recordset!NumAsien
    Sql = Sql & "));"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Deb = 0
    hab = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Deb = Rs.Fields(0)
        If Not IsNull(Rs.Fields(1)) Then hab = Rs.Fields(1)
    End If
    Rs.Close
    Set Rs = Nothing
    Text2(0).Text = Format(Deb, FormatoImporte): Text2(1).Text = Format(hab, FormatoImporte)
    'Metemos en DEB el total
    Deb = Deb - hab
    If Deb < 0 Then
        Text2(2).ForeColor = vbRed
        Else
        Text2(2).ForeColor = vbBlack
    End If
    If Deb <> 0 Then Text2(2).Text = Format(Deb, FormatoImporte)
End Sub

Private Sub PideCalculadora()
On Error GoTo EPideCalculadora
    Shell App.Path & "\arical.exe", vbNormalFocus
    Exit Sub
EPideCalculadora:
    Err.Clear
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
        If miI >= 0 Then Image1_Click miI
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
    If Indice <> 8 Then
        Sql = "SELECT "
        Select Case Indice
        Case 4
            C = "codmacta"
            I = 5
        Case 5
            C = "numdocum"
            I = 6
        Case 6
            C = "ctacontr"
            I = 7
        Case 7
            C = "codconce"
            I = 8
        Case 11
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
        Sql = "Select linliapu,ampconce,nomconce FROM hlinapu,conceptos"
        Sql = Sql & " WHERE conceptos.codconce=hlinapu.codconce AND  numdiari=" & Data1.Recordset!NumDiari
        Sql = Sql & " AND fechaent='" & Format(Data1.Recordset!FechaEnt, FormatoFecha)
        Sql = Sql & "' AND numasien=" & Data1.Recordset!NumAsien
        If ModoLineas = 2 Then Sql = Sql & " AND hlinapu <" & Linliapu
           
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
            txtAux(8).Text = txtAux(8).Text & Sql & " "
            txtAux(8).SelStart = Len(txtAux(8).Text)
            PonFoco txtAux(9)
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
            
    
    C = "SELECT hlinapu.numasien, hlinapu.linliapu, hlinapu.codmacta, cuentas.nommacta,"
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
    
        MsgBox "No se ha encontrado las lineas: " & vbCrLf & C, vbExclamation
    Else
        'Ya tengo la ultima linea
        txtAux(4).Text = RsF6!codmacta
        
        txtAux(4).Text = RsF6!codmacta
        txtAux2(4).Text = RsF6!Nommacta
        txtAux(5).Text = DBLet(RsF6!Numdocum, "T")
        txtAux(6).Text = DBLet(RsF6!ctacontr, "T")
        txtAux(7).Text = RsF6!codconce
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
        txtAux(11).Text = DBLet(RsF6!codccost, "T")
        HabilitarImportes 3
        HabilitarCentroCoste
        Text3(5).Text = DBLet(RsF6!nomctapar, "T")
        Text3(4).Text = DBLet(RsF6!nombreconcepto, "T")
        Text3(3).Text = DBLet(RsF6!centrocoste, "T")
        
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
    
    
    'Contrapartida
    If txtAux(6).Text <> "" Then
        If Not IsNumeric(txtAux(6).Text) Then
            AuxOK = "Cuenta contrapartida debe ser num�rica"
            Exit Function
        End If
        If Text3(5).Text = NO Then
            AuxOK = "La cta. contrapartida no esta dada de alta en el sistema."
            Exit Function
        End If
        If Not EsCuentaUltimoNivel(txtAux(6).Text) Then
            AuxOK = "La cuenta contrapartida no es de �ltimo nivel"
            Exit Function
        End If
    End If
        
    'Concepto
    If txtAux(4).Text = "" Then
        AuxOK = "El concepto no puede estar vacio"
        Exit Function
    End If
        
    If txtAux(7).Text <> "" Then
        If Not IsNumeric(txtAux(7).Text) Then
            AuxOK = "El concepto debe de ser num�rico."
            Exit Function
        End If
    End If
    
    'Importe
    If txtAux(9).Text <> "" Then
        If Not EsNumerico(txtAux(9).Text) Then
            AuxOK = "El importe DEBE debe ser num�rico"
            Exit Function
        End If
    End If
    
    If txtAux(10).Text <> "" Then
        If Not EsNumerico(txtAux(10).Text) Then
            AuxOK = "El importe HABER debe ser num�rico"
            Exit Function
        End If
    End If
    
    If Not (txtAux(9).Text = "" Xor txtAux(10).Text = "") Then
        AuxOK = "Solo el debe, o solo el haber, tiene que tener valor"
        Exit Function
    End If
    
    
    'cENTRO DE COSTE
    If txtAux(11).Enabled Then
        If txtAux(11).Text = "" Then
            AuxOK = "Centro de coste no puede ser nulo"
            Exit Function
        End If
    End If
    
                                            'Fecha del asiento
    If EstaLaCuentaBloqueada(txtAux(4).Text, CDate(Text1(1).Text)) Then
        AuxOK = "Cuenta bloqueada: " & txtAux(4).Text
        Exit Function
    End If
    
    'Si lleva contrapartida
    If txtAux(6).Text <> "" Then
        If EstaLaCuentaBloqueada(txtAux(6).Text, CDate(Text1(1).Text)) Then
            AuxOK = "Cuenta contrapartida bloqueada: " & txtAux(6).Text
            Exit Function
        End If
    End If
    AuxOK = ""
End Function



Private Function ComprobarNumeroAsiento(Actual As Boolean) As Boolean
Dim cad As String
Dim RT As ADODB.Recordset
        cad = " WHERE numasien=" & Text1(0).Text
        If Actual Then
            I = 0
        Else
            I = 1
        End If
        cad = cad & " AND fechaent >='" & Format(DateAdd("yyyy", I, vParam.fechaini), FormatoFecha)
        cad = cad & "' AND fechaent <='" & Format(DateAdd("yyyy", I, vParam.fechafin), FormatoFecha) & "'"
        Set RT = New ADODB.Recordset
        ComprobarNumeroAsiento = True
        I = 0
        RT.Open "Select numasien from hlinapu" & cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.EOF) Then
                ComprobarNumeroAsiento = False
            End If
        End If
        RT.Close
        If ComprobarNumeroAsiento Then
            I = 1
            RT.Open "Select numasien from hlinapu" & cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RT.EOF Then
                If Not IsNull(RT.EOF) Then
                    ComprobarNumeroAsiento = False
                End If
            End If
            RT.Close
        End If
        Set RT = Nothing
        If Not ComprobarNumeroAsiento Then
            cad = "Verifique los contadores. Ya exsite el asiento; " & Text1(0).Text & vbCrLf
            If I = 0 Then
                cad = cad & " en la introducci�n de apuntes"
            Else
                cad = cad & " en el hist�rico."
            End If
            MsgBox cad, vbExclamation
        End If
End Function

Private Function SituarData1(Insertar As Boolean) As Boolean
    Dim Sql As String
    
    On Error GoTo ESituarData1
    
    
    'Si es insertar, lo que hace es simplemente volver a poner el el recordset
    'este unico registro
    'If Insertar Then
        Sql = "Select * from hcabapu WHERE numasien =" & Text1(0).Text
        Sql = Sql & " AND fechaent='" & Format(Text1(1).Text, FormatoFecha) & "' AND numdiari = " & Text1(2).Text
        Data1.RecordSource = Sql
    'End If
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!NumAsien) = Text1(0).Text Then
                If CStr(.Fields!NumDiari) = Text1(2).Text Then
                    If Format(CStr(.Fields!FechaEnt), "dd/mm/yyyy") = Text1(1).Text Then
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

Private Sub FijarContraPartida()
    Dim cad As String
    'Hay contrapartida
    'Reasignamos campos de cuentas
    cad = txtAux(4).Text
    txtAux(4).Text = txtAux(6).Text
    txtAux(6).Text = cad
    HabilitarCentroCoste
    cad = txtAux2(4).Text
    txtAux2(4).Text = Text3(5).Text
    Text3(5).Text = cad
    
    'Los importes
    HabilitarImportes 3
    cad = txtAux(9).Text
    txtAux(9).Text = txtAux(10).Text
    txtAux(10).Text = cad
End Sub

'********************************************************
'
' FUNCIONES CORRESPONDIENTES A LA INSERCION DE DOCUMENTOS
'
'********************************************************
Private Function InsertarDesdeFichero() As Boolean
Dim CADENA As String
Dim Carpeta As String
Dim Aux As String
Dim J As Integer
Dim C As String
Dim Rs As ADODB.Recordset
Dim L As Long


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
        MsgBox "No se permite insertar ficheros de tama�o superior a 1 M", vbExclamation
        InsertarDesdeFichero = False
        Exit Function
    End If
    
    
'    '******* Cambiamos cursor
    Screen.MousePointer = vbHourglass

    J = InStr(1, cd1.FileName, Chr(0))
    CADENA = cd1.FileName
    TipoDocu = 0
    If InStr(1, cd1.FileName, "pdf") <> 0 Then TipoDocu = 1
    Fichero = CADENA
        
            
    Screen.MousePointer = vbDefault
    
    txtaux3(4).Text = CCur(DevuelveValor("select max(orden) from hcabapu_fichdocs where numasien = " & DBSet(Text1(0), "N") & " and fechaent = " & DBSet(Text1(1).Text, "F") & " and numdiari = " & DBSet(Text1(2).Text, "N")) + 1)
    txtaux3(5).Text = Dir(CADENA)
    
    C = "Select max(codigo) from hcabapu_fichdocs"
    Set Rs = New ADODB.Recordset
    Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then L = Rs.Fields(0)
    End If
    L = L + 1
    Rs.Close
    
    ' es nuevo
    C = "insert into hcabapu_fichdocs (codigo, numasien, fechaent, numdiari, orden, docum) values"
    C = C & " (" & DBSet(L, "N") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Text1(1).Text, "F") & "," & DBSet(Text1(2).Text, "N") & "," & DBSet(txtaux3(4).Text, "N") & "," & DBSet(txtaux3(5).Text, "T") & ")"
    Conn.Execute C
    
    espera 0.2
    
    'Abro parar guardar el binary
    C = "Select * from hcabapu_fichdocs where codigo =" & L '& " and codsocio = " & DBSet(RecuperaValor(vDatos, 1), "N")
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = C
    adodc1.Refresh
'
    If adodc1.Recordset.EOF Then
        'MAAAAAAAAAAAAL

    Else
        'Guardar
        GuardarBinary adodc1.Recordset!Campo, Fichero
        adodc1.Recordset.Update
    End If
    
End Function



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
    
    If Modo <> 2 Then Exit Sub
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 5 ' imagenes
        cad = "select h.orden, h.campo, h.codigo, h.docum from hcabapu_fichdocs h WHERE "
        cad = cad & " numasien=" & Data1.Recordset!NumAsien
        cad = cad & " and fechaent=" & DBSet(Data1.Recordset!FechaEnt, "F")
        cad = cad & " and numdiari=" & Data1.Recordset!NumDiari
        GroupBy = ""
        BuscaChekc = "orden"
        
    End Select
    
    
    'BuscaChekc="" si es la opcion de precios especiales
    cad = cad & " ORDER BY 1"
    
    lw1.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not Rs.EOF
        Set IT = lw1.ListItems.Add

        IT.Text = Rs!Orden '"Nuevo " & Contador

        IT.SubItems(1) = Rs!DOCUM  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        IT.SubItems(2) = App.Path & "\" & CarpetaIMG & "\" & Rs!DOCUM
        IT.SubItems(3) = Rs!Codigo

        Set IT = Nothing

        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    ProcesarCarpetaImagenes
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
End Sub

Private Sub CargarArchivos()
Dim C As String
Dim L As Long


    ProcesarCarpetaImagenes


    C = "Select * from hcabapu_fichdocs where numasien=" & DBSet(Text1(0).Text, "N")
    C = C & " and fechaent = " & DBSet(Text1(1).Text, "F")
    C = C & " and numdiari = " & DBSet(Text1(2).Text, "N")
    C = C & " and codigo = " & DBSet(lw1.SelectedItem.SubItems(3), "N")

    adodc1.ConnectionString = Conn
    adodc1.RecordSource = C
    adodc1.Refresh

    If adodc1.Recordset.EOF Then
        'NO HAY NINGUNA
    Else
        'LEEMOS LA IMAGEN
        L = adodc1.Recordset!Codigo
        C = App.Path & "\" & CarpetaIMG & "\" & L
        If DBLet(adodc1.Recordset!DOCUM) <> "0" Then
            C = App.Path & "\" & CarpetaIMG & "\" & adodc1.Recordset!DOCUM
        End If
        LeerBinary adodc1.Recordset!Campo, C
    End If

End Sub



Private Sub ProcesarCarpetaImagenes()
Dim C As String
Dim MiNombre As String

    On Error GoTo EProcesarCarpetaImagenes
    
    C = App.Path & "\" & CarpetaIMG
    If Dir(C, vbDirectory) = "" Then
        MkDir C
    Else
        On Error Resume Next
        If Dir(C & "\*.*", vbArchive) <> "" Then 'Kill c & "\*.*"
            MiNombre = Dir(C & "\*.*")   ' Recupera la primera entrada.
            Do While MiNombre <> ""   ' Inicia el bucle.
               ' Ignora el directorio actual y el que lo abarca.
               If MiNombre <> "." And MiNombre <> ".." Then
                    Kill C & "\" & MiNombre
               End If
               MiNombre = Dir   ' Obtiene siguiente entrada.
            Loop
        End If
        On Error GoTo EProcesarCarpetaImagenes
    
    End If
    
    Exit Sub
EProcesarCarpetaImagenes:
    MuestraError Err.Number, "ProcesarCarpetaImagenes"
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

        IT.Text = Me.adodc1.Recordset!Orden '"Nuevo " & Contador
        
        IT.SubItems(1) = Me.adodc1.Recordset.Fields(5)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        IT.SubItems(2) = vpaz
        IT.SubItems(3) = Me.adodc1.Recordset.Fields(0)
        
        Set IT = Nothing
    End If
End Sub


Private Sub ImprimirImagen()
Dim NFic As Long
Dim vAdobe As String

   CargarArchivos
   
   Call ShellExecute(Me.hWnd, "Open", Me.lw1.SelectedItem.SubItems(2), "", "", 1)
   
End Sub


Private Sub EliminarImagen()
Dim Sql As String
Dim Mens As String
    
    On Error GoTo eEliminarImagen

    Mens = "Va a proceder a eliminar de la lista correspondiente al asiento. " & vbCrLf & vbCrLf & "� Desea continuar ?" & vbCrLf & vbCrLf
    
    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
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
    


Private Sub CargarArchivosOLD()
Dim C As String
Dim L As Long
Dim Rs As ADODB.Recordset
Dim nFile As Long


    ProcesarCarpetaImagenes

    C = "Select * from hcabapu_fichdocs where numasien=" & DBSet(Text1(0).Text, "N")
    C = C & " and fechaent = " & DBSet(Text1(1).Text, "F")
    C = C & " and numdiari = " & DBSet(Text1(2).Text, "N")
    C = C & " ORDER BY orden"

    adodc1.ConnectionString = Conn
    adodc1.RecordSource = C
    adodc1.Refresh

    If adodc1.Recordset.EOF Then
        'NO HAY NINGUNA
    Else
        'LEEMOS LAS IMAGENES
        While Not adodc1.Recordset.EOF
            L = adodc1.Recordset!Codigo
            C = App.Path & "\" & CarpetaIMG & "\" & L
            If DBLet(adodc1.Recordset!DOCUM) <> "0" Then
                C = App.Path & "\" & CarpetaIMG & "\" & adodc1.Recordset!DOCUM
            End If
            If Dir(C) <> "" Then
                AnyadirAlListview C, True
            Else
                If LeerBinary(adodc1.Recordset!Campo, C) Then
                    AnyadirAlListview C, True
                End If
            End If

            adodc1.Recordset.MoveNext
        Wend
    
    End If

End Sub


Private Function SePuedeModificarAsiento(MostrarMensaje As Boolean) As Boolean
Dim CadFac As String

        CadFac = ""
        
        SePuedeModificarAsiento = False
      
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
                I = -1

            Exit Function
        Else
            SePuedeModificarAsiento = True
        End If


End Function

