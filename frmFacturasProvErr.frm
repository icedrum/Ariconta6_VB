VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmFacturProvErr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro facturas proveedores con ERRORES"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmFacturasProvErr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   10680
      TabIndex        =   77
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   8160
      TabIndex        =   29
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
      TabIndex        =   27
      Top             =   7200
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10680
      TabIndex        =   24
      Top             =   7800
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   3720
      TabIndex        =   26
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
      TabIndex        =   36
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
      TabIndex        =   28
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   34
      Top             =   7770
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   32
      Top             =   7680
      Width           =   2655
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9480
      TabIndex        =   23
      Top             =   7800
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacturasProvErr.frx":030A
      Height          =   2655
      Left            =   1680
      TabIndex        =   35
      Top             =   4920
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   37
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
            Object.ToolTipText     =   "Recuperar factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Verificar"
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
         Left            =   9960
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame framecabeceras 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      TabIndex        =   39
      Top             =   480
      Width           =   11895
      Begin VB.CheckBox Check2 
         Caption         =   "No deducible"
         Height          =   255
         Left            =   7080
         TabIndex        =   80
         Tag             =   "Extranjero|N|S|||cabfactprove|nodeducible|||"
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Extranjero"
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Tag             =   "Extranjero|N|S|||cabfactprove|extranje|||"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   3720
         TabIndex        =   79
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   29
         Left            =   3780
         TabIndex        =   6
         Tag             =   "Fecha liquidacion|F|N|||cabfactprove|fecliqpr|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   28
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Nº factura|T|N|||cabfactprove|numfacpr|||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   27
         Left            =   120
         TabIndex        =   73
         Tag             =   "año factura|N|S|||cabfactprove|totfacpr||N|"
         Text            =   "Text1"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   26
         Left            =   120
         TabIndex        =   72
         Tag             =   "año factura|N|S|||cabfactprove|anofacpr||S|"
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
         TabIndex        =   5
         Tag             =   "Observaciones(Concepto)|T|S|||cabfactprove|confacpr|||"
         Text            =   "DDDDDDDDDDDDDDD"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   4800
         TabIndex        =   3
         Tag             =   "Fecha factura|F|N|||cabfactprove|fecfacpr|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Tag             =   "Fecha recepcion|F|N|||cabfactprove|fecrecpr|||"
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
         Tag             =   "Nº registro|N|S|0||cabfactprove|numregis||S|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   10800
         TabIndex        =   54
         Tag             =   "Numero serie|N|S|||cabfactprove|numasien|||"
         Text            =   "9999999999"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   7560
         TabIndex        =   53
         Text            =   "Text4"
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   6240
         TabIndex        =   4
         Tag             =   "Cuenta cliente|T|N|||cabfactprove|codmacta|||"
         Text            =   "0000000000"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   6
         Left            =   1680
         TabIndex        =   8
         Tag             =   "Base imponible 1|N|N|||cabfactprove|ba1facpr|#,###,###,##0.00||"
         Top             =   1755
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   3240
         TabIndex        =   9
         Tag             =   "Tipo IVA 1|N|S|0|100|cabfactprove|tp1facpr|||"
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
         TabIndex        =   52
         Tag             =   "Porcentaje IVA 1|N|S|||cabfactprove|pi1facpr|#0.00||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   9
         Left            =   6960
         TabIndex        =   10
         Tag             =   "Importe IVA 1|N|S|||cabfactprove|ti1facpr|#,###,##0.00||"
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
         TabIndex        =   51
         Tag             =   "Porcentaje recargo 1|N|S|||cabfactprove|pr1facpr|#0.00||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   11
         Left            =   9240
         TabIndex        =   11
         Tag             =   "Importe recargo 1|N|S|||cabfactprove|tr1facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   1755
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   12
         Left            =   1680
         TabIndex        =   12
         Tag             =   "Base imponible 2|N|S|||cabfactprove|ba2facpr|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   3240
         TabIndex        =   13
         Tag             =   "Tipo IVA 2|N|S|0|100|cabfactprove|tp2facpr|||"
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
         TabIndex        =   50
         Tag             =   "Porcentaje IVA 2|N|S|||cabfactprove|pi2facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   15
         Left            =   6960
         TabIndex        =   14
         Tag             =   "Importe IVA 2|N|S|||cabfactprove|ti2facpr|#,###,##0.00||"
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
         TabIndex        =   15
         Tag             =   "Porcentaje recargo 2|N|S|||cabfactprove|pr2facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   17
         Left            =   9240
         TabIndex        =   49
         Tag             =   "Importe recargo 2|N|S|||cabfactprove|tr2facpr|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   18
         Left            =   1680
         TabIndex        =   16
         Tag             =   "Base imponible 3|N|S|||cabfactprove|ba3facpr|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   19
         Left            =   3240
         TabIndex        =   17
         Tag             =   "Tipo IVA 3|N|S|0|100|cabfactprove|tp3facpr|||"
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
         TabIndex        =   48
         Tag             =   "Porcentaje IVA 3|N|S|||cabfactprove|pi3facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   21
         Left            =   6960
         TabIndex        =   18
         Tag             =   "Importe IVA 3|N|S|||cabfactprove|ti3facpr|#,###,##0.00||"
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
         TabIndex        =   47
         Tag             =   "Porcentaje recargo 3|N|S|||cabfactprove|pr3facpr|#0.00||"
         Text            =   "Text1"
         Top             =   2805
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   23
         Left            =   9240
         TabIndex        =   19
         Tag             =   "Importe recargo 3|N|S|||cabfactprove|tr3facpr|#,###,##0.00||"
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   25
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Text            =   "Text4"
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   2520
         TabIndex        =   21
         Tag             =   "Cuenta retencion|T|S|||cabfactprove|cuereten|||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   24
         Left            =   1680
         TabIndex        =   20
         Tag             =   "Porcentaje retencion|N|S|||cabfactprove|retfacpr|#0.00||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   22
         Tag             =   "Cuenta retencion|N|S|||cabfactprove|trefacpr|#,##0.00||"
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
         TabIndex        =   40
         Text            =   "123.123.123.123,11"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   7
         Left            =   4740
         Picture         =   "frmFacturasProvErr.frx":031F
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F.Liquidacion"
         Height          =   195
         Index           =   4
         Left            =   3780
         TabIndex        =   78
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Nº factura"
         Height          =   195
         Index           =   3
         Left            =   3120
         TabIndex        =   76
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Recepción"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   75
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   71
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   " Fecha"
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   70
         Top             =   120
         Width           =   495
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   1080
         Picture         =   "frmFacturasProvErr.frx":0421
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   5280
         Picture         =   "frmFacturasProvErr.frx":04AC
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº registro"
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   69
         Top             =   120
         Width           =   735
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   7080
         Picture         =   "frmFacturasProvErr.frx":0537
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   7
         Left            =   6240
         TabIndex        =   68
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento"
         Height          =   195
         Index           =   8
         Left            =   10800
         TabIndex        =   67
         Top             =   0
         Width           =   975
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   3705
         Picture         =   "frmFacturasProvErr.frx":0F39
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   3705
         Picture         =   "frmFacturasProvErr.frx":193B
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   3705
         Picture         =   "frmFacturasProvErr.frx":233D
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
         TabIndex        =   66
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   3795
         Width           =   1455
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   3675
         Picture         =   "frmFacturasProvErr.frx":2D3F
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
         ForeColor       =   &H000040C0&
         Height          =   360
         Index           =   2
         Left            =   8640
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   3600
         Width           =   570
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas proveedores  erroneas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2760
      TabIndex        =   74
      Top             =   7740
      Width           =   6735
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
End
Attribute VB_Name = "frmFacturProvErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
Dim I As Integer
Dim ancho As Integer


Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
'Para pasar de lineas a cabeceras
Dim Linfac As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean
Dim RS As Recordset
Dim Aux As Currency
Dim Base As Currency
Dim AUX2 As Currency
Dim SumaLinea As Currency
Dim AntiguoText1 As String


Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Function ActualizaFactura() As Boolean
Dim B As Boolean
On Error GoTo EActualiza
ActualizaFactura = False

B = ModificaDesdeFormularioClaves(Me, SQL)
If Not B Then Exit Function

'Las lineas
If Not Adodc1.Recordset.EOF Then
    SQL = "UPDATE linfactprove SET numregis =" & Text1(2).Text
    SQL = SQL & " ,anofacpr = " & Text1(26).Text
    SQL = SQL & " WHERE numregis=" & Data1.Recordset!NumRegis
    SQL = SQL & " AND anofacpr =" & Data1.Recordset!anofacpr
    Conn.Execute SQL
End If

ActualizaFactura = True
Exit Function
EActualiza:
    MuestraError Err.Number, "Modificando claves factura"
End Function

Private Sub cmdAceptar_Click()
    Dim Cad As String
    Dim I As Integer
    Dim RC As Boolean
    Dim Contador As Long
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            I = FechaCorrecta2(CDate(Text1(1).Text))
            If I > 1 Then
                If I = 2 Then
                    MsgBox varTxtFec, vbExclamation
                Else
                    MsgBox "La fecha factura no pertenece al ejercicio actural ni al siguiente.", vbExclamation
                End If
                Exit Sub
            End If
            Text1(2).Text = Siguiente
                        
            '-----------------------------------------
            'Hacemos insertar
                If InsertarDesdeForm(Me) Then
                    Data1.Refresh
                    If SituarData1 Then
                        PonerModo 5
                        'Haremos como si pulsamo el boton de insertar nuevas lineas
                        'Ponemos el importe en AUX
                        Aux = ImporteFormateado(Text2(4).Text)
                        cmdCancelar.Caption = "Cabecera"
                        ModificandoLineas = 0
                        AnyadirLinea True, False
                    Else
                        SQL = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                        SQL = SQL & vbCrLf & " CLAVE: FrmFacturas. cmdAceptar. SituarData1"
                        MsgBox SQL, vbCritical
                    End If
                End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hay que comprobar si ha modificado, o no la clave de la factura
                I = 1
                If Data1.Recordset!NumRegis = Text1(2).Text Then
                        If Data1.Recordset!anofacpr = Text1(26).Text Then
                            I = 0
                            'NO HA MODIFICADO NADA
                    End If
                End If
            
                'Hacemos MODIFICAR
                If I <> 0 Then
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
                    If SituarData1 Then
                        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                        PonerModo 2
                    Else
                        PonerModo 0
                    End If
                    'MsgBox "El registro ha sido modificado", vbInformation
                    
                    
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
                    If Aux <> 0 Then txtaux(0).SetFocus
                    
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
    If txtaux(2).Text <> "" Then txtaux(4).SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
    Case 1, 3
        LimpiarCampos
        PonerModo 0
        'Contador de facturas
    Case 4

        lblIndicador.Caption = ""
        PonerModo 2
        PonerCampos

    Case 5
        CamposAux False, 0, False

        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            PonerModo 2
        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
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
Private Function SituarData1() As Boolean
    Dim SQL As String
    
    On Error GoTo ESituarData1
    
    Data1.Refresh
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!NumRegis) = Text1(2).Text Then
                If CStr(.Fields!anofacpr) = Text1(26).Text Then
                        SituarData1 = True
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
End Function


Private Function IntegrarFactura() As Boolean
Dim Mc As Contadores
Dim FechaBien As Byte
Dim EsDeAnal As Boolean
Dim EjercicioActualoSiguiente As Byte
IntegrarFactura = False


If vParam.Constructoras Then
    'Fecha liquidacion
    EjercicioActualoSiguiente = FechaCorrecta2(CDate(Text1(29).Text))
    If EjercicioActualoSiguiente > 1 Then
        If EjercicioActualoSiguiente = 2 Then
            MsgBox varTxtFec, vbExclamation
            Exit Function
        Else
            If EjercicioActualoSiguiente = 3 Then
                SQL = " cerrado"
            Else
                SQL = " no abierto"
            End If
        End If
        SQL = "Fecha incorrecta. La fecha LIQUIDACION pertenece a un ejercicio "
        SQL = SQL & vbCrLf & vbCrLf & "¿Desea continuar igualmente?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
End If

EjercicioActualoSiguiente = FechaCorrecta2(CDate(Text1(1).Text))
If EjercicioActualoSiguiente > 1 Then
    If EjercicioActualoSiguiente = 2 Then
        SQL = varTxtFec
    Else
        If EjercicioActualoSiguiente = 3 Then
            SQL = " cerrado"
        Else
            SQL = " no abierto"
        End If
        SQL = "Fecha incorrecta. La fecha pertenece a un ejercicio " & SQL
    End If
    MsgBox SQL, vbExclamation
    Exit Function
End If

If Text1(2).Text = "" Then Exit Function
'Comprobamos k existen las cuenas, de IVA y demas
SQL = ""
If Text4(0).Text = "" Then SQL = SQL & ".- Cliente incorrecto" & vbCrLf
If Text1(7).Text <> "" And Text4(1).Text = "" Then SQL = SQL & ".- IVA1 incorrecto" & vbCrLf
If Text1(13).Text <> "" And Text4(2).Text = "" Then SQL = SQL & ".- IVA2 incorrecto" & vbCrLf
If Text1(19).Text <> "" And Text4(3).Text = "" Then SQL = SQL & ".- IVA3 incorrecto" & vbCrLf
If Text1(3).Text <> "" And Text4(4).Text = "" Then SQL = SQL & ".- Cta retencion incorrectaq" & vbCrLf
If Text1(7).Text = "" Then SQL = SQL & ".- IVA1 requerido" & vbCrLf
If SQL <> "" Then
    MsgBox "Error en encabezado de factura: " & vbCrLf & vbCrLf & SQL, vbExclamation
    Exit Function
End If



'Las lineas de factura
SQL = ""
If Adodc1.Recordset.EOF Then
    SQL = "No hay lineas para la factura" & vbCrLf
Else
    Adodc1.Recordset.MoveFirst
    While Not Adodc1.Recordset.EOF
        If IsNull(Adodc1.Recordset.Fields(1)) Then SQL = SQL & "Cuenta de bases incorrecta: " & Adodc1.Recordset.Fields(0) & vbCrLf
        'If Not IsNull(Adodc1.Recordset.Fields(2)) And IsNull(Adodc1.Recordset.Fields(3)) Then SQL = SQL & "Centro de coste incorrecto: " & Adodc1.Recordset.Fields(3) & vbCrLf
        Adodc1.Recordset.MoveNext
    Wend
    Adodc1.Recordset.MoveFirst
End If
If SQL <> "" Then
    MsgBox SQL, vbExclamation
    Exit Function
End If


If vParam.autocoste Then
    Adodc1.Recordset.MoveFirst
    FechaBien = 0
    While FechaBien = 0
        'Para cada cuenta compruebo si es de analitica
        SQL = ""
        If DBLet(Adodc1.Recordset!codccost, "T") <> "" Then SQL = "1"
        If DBLet(Adodc1.Recordset!nomccost, "T") <> "" Then SQL = SQL & "2"
        If Len(SQL) = 1 Then
            'Por algun motivo, esta el CC, pero no el nombre, o al reves
            MsgBox "Error en centro de coste(" & SQL & ")- Cuenta: " & Adodc1.Recordset!codtbase, vbExclamation
            Exit Function
        End If
        
        EsDeAnal = False
        
        
        SQL = Mid(Adodc1.Recordset!codtbase, 1, 1)
        If SQL = vParam.grupogto Then
            EsDeAnal = True
        Else
            If SQL = vParam.grupovta Then
                EsDeAnal = True
            Else
                If SQL = vParam.grupoord Then EsDeAnal = True
            End If
        End If
        
        If Not EsDeAnal Then
            If vParam.Subgrupo1 <> "" Then
                SQL = Mid(Adodc1.Recordset!codtbase, 1, 3)
                If vParam.Subgrupo1 = SQL Then EsDeAnal = True
            End If
        End If
        
        SQL = DBLet(Adodc1.Recordset!codccost, "T")
        If EsDeAnal Then
            'Es de analitica
            If SQL = "" Then
                MsgBox "Centro de coste requerido", vbExclamation
                FechaBien = 1
            End If
        Else
            If SQL <> "" Then
                MsgBox "No debe poner CC para la cuenta: " & Adodc1.Recordset!codtbase, vbExclamation
                FechaBien = 1
            End If
        End If
        
        'Siguiente
        If FechaBien = 0 Then
            Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF Then FechaBien = 2
        End If
    Wend
    Adodc1.Recordset.MoveFirst
    
    If FechaBien <> 2 Then Exit Function
    
    
    If vParam.CuentasBloqueadas <> "" Then
    
    
        'Primero compruebo las cuentas de cabecera
        If EstaLaCuentaBloqueada(Data1.Recordset!Codmacta, Data1.Recordset!fecrecpr) Then
            MsgBox "Cuenta bloqueada: " & Data1.Recordset!Codmacta, vbExclamation
            Exit Function
        End If
        SQL = DBLet(Data1.Recordset!cuereten, "T")
        If SQL <> "" Then
            If EstaLaCuentaBloqueada(SQL, Data1.Recordset!fecrecpr) Then
                MsgBox "Cuenta bloqueada: " & SQL, vbExclamation
                Exit Function
            End If
        End If
    
    
        SQL = ""
        Adodc1.Recordset.MoveFirst
        While Not Adodc1.Recordset.EOF
            If EstaLaCuentaBloqueada(Adodc1.Recordset!codtbase, Data1.Recordset!fecrecpr) Then
                If SQL = "" Then SQL = "Cuentas bloqueadas: " & vbCrLf
                SQL = SQL & "        - " & Adodc1.Recordset!codtbase & vbCrLf
            End If
            Adodc1.Recordset.MoveNext
        Wend
        Adodc1.Recordset.MoveFirst
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Function
        End If
    End If
    
End If
    








SQL = ""
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
    AntiguoText1 = ""
    
    Set Mc = New Contadores

    If Mc.ConseguirContador("1", EjercicioActualoSiguiente = 0, False) = 1 Then
        AntiguoText1 = "Error al conseguir el contador."
        I = 1
    Else
        Linfac = Mc.Contador
        I = 0
    End If
    
    'Ahora compruebo k la factura no existe
    If I = 0 Then
        Set RS = New Recordset
        SQL = "Select numregis from cabfactprov where numregis=" & Linfac & " AND anofacpr=" & Data1.Recordset!anofacpr
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            I = 1
            AntiguoText1 = "Ya existe la factura devuelta por el contador"
        End If
        RS.Close
        Set RS = Nothing
    End If
    
    If I = 0 Then
        'Si llega aqui
        'Si llega aqui es k podemos pasarla a facturas correctas
        Conn.BeginTrans
        If ActualizarRegistros Then
            Conn.CommitTrans
        Else
            I = 1
            Conn.RollbackTrans
        End If
    End If
    
    'Borramos la linea
    If I = 0 Then
        
        If Not BorrarFactura(Linfac) Then
            MsgBox "Error: Elimine la factura errornea manualmente", vbExclamation
        Else
            MsgBox "Traspaso realizado con éxito", vbExclamation
            CargaGrid False
            IntegrarFactura = True
        End If
    Else
        'Intentamos devolver contador
        Linfac = FechaCorrecta2(CDate(Text1(1).Text))
        Mc.DevolverContador "1", Linfac = 0, Mc.Contador
        MsgBox "Se han producido errores durante el traspaso.", vbExclamation
    End If
    Set Mc = Nothing
End Function



Private Sub BotonAnyadir()
    LimpiarCampos
    Check1.Value = 0 'Intracomunitaria

    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    '###A mano
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(1).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        cmdAceptar.Caption = "Aceptar"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(1).SetFocus
        Text1(1).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                Text1(kCampo).SetFocus
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    SQL = ""
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    DataGrid1.Enabled = False
    CargaGrid False
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & SQL & Ordenacion
        PonerCadenaBusqueda
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

       
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    'Escondemos el navegador y ponemos insertando
    'Como el campo 1 es clave primaria, NO se puede modificar
    '### A mano

    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    DataGrid1.Enabled = False
        
    
    'Comprobamos que no esta actualizada ya
    SQL = ""
    SQL = SQL & vbCrLf & vbCrLf & "Va usted a eliminar la factura :" & vbCrLf
    SQL = SQL & "Numero : " & Data1.Recordset!NumRegis & vbCrLf
    SQL = SQL & "Fecha recepcion : " & Format(Data1.Recordset!fecrecpr, "dd/mm/yyyy") & vbCrLf
    SQL = SQL & "Proveedor : " & Data1.Recordset!Codmacta & " - " & Text4(0).Text & vbCrLf
    SQL = SQL & vbCrLf & "          ¿Desea continuar ?" & vbCrLf
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    NumRegElim = Data1.Recordset.AbsolutePosition
    If Data1.Recordset.RecordCount = NumRegElim Then
        NumRegElim = NumRegElim - 2
    Else
        NumRegElim = NumRegElim - 1
    End If
    Screen.MousePointer = vbHourglass
     'La borrara desde este mismo form
    If BorrarFactura(Data1.Recordset!NumRegis) Then
        AlgunAsientoActualizado = True
    Else
        AlgunAsientoActualizado = False
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
            NumRegElim = NumRegElim
            If NumRegElim > -1 Then Data1.Recordset.Move NumRegElim
            PonerCampos
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function BorrarFactura(NumRegis As Long) As Boolean
    
    On Error GoTo EBorrar
    SQL = " WHERE numregis = " & NumRegis
    SQL = SQL & " AND anofacpr= " & Data1.Recordset!anofacpr
    'Las lineas
    AntiguoText1 = "DELETE from linfactprove " & SQL
    Conn.Execute AntiguoText1
    'La factura
    AntiguoText1 = "DELETE from cabfactprove " & SQL
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
Dim I As Integer
Dim J As Integer
Dim Aux As String

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
    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    SQL = ""
    
    
    If vParam.Constructoras Then
       ancho = FrameTapa.Left + FrameTapa.Width + 100
    Else
        ancho = FrameTapa.Left
    End If
    Check1.Left = ancho
    Check2.Left = Check1.Left + Check1.Width + 120
    Text1(29).Enabled = vParam.Constructoras
    FrameTapa.Visible = Not vParam.Constructoras
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 18
        .Buttons(13).Image = 20
        .Buttons(14).Image = 15
        .Buttons(16).Image = 6
        .Buttons(17).Image = 7
        .Buttons(18).Image = 8
        .Buttons(19).Image = 9
    End With
    
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
        Me.Width = 12000
        Me.Height = Screen.Height
    End If
    Me.Height = 9000
    'Los campos auxiliares
    CamposAux False, 0, True
    
    
    '## A mano
    NombreTabla = "cabfactprove"
    Ordenacion = " ORDER BY fecfacpr"
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
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
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
End Sub


Private Sub Form_Resize()
If Me.WindowState <> 0 Then Exit Sub
If Me.Width < 11610 Then Me.Width = 11610
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(26), CadenaDevuelta, 1)
        CadB = Aux
        
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 2)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " "
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas
Select Case cmdAux(0).Tag
Case 2
    'Cuenta normal
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text4(0).Text = RecuperaValor(CadenaSeleccion, 2)
Case 6
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1)
    Text4(4).Text = RecuperaValor(CadenaSeleccion, 2)
Case 100
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Select
End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
'Centro de coste
txtaux(2).Text = RecuperaValor(CadenaSeleccion, 1)
txtaux(3).Text = RecuperaValor(CadenaSeleccion, 2)
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
    I = CInt(Aux - 2)
    Text1(((I) * 6) + 1).Text = RecuperaValor(CadenaSeleccion, 1)
    If PonerValoresIva(I) Then
        CalcularIVA I
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
        
        If Index = 0 Then
            Linfac = 0
        Else
            If Index = 1 Then
                Linfac = 1
            Else
                Linfac = 29
            End If
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
    Case 3, 4, 5
        Aux = Index
        Set frmI = New frmIVA
        frmI.DatosADevolverBusqueda = "0|1|"
        frmI.Show vbModal
        Set frmI = Nothing
    End Select
    Screen.MousePointer = vbDefault
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
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
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
Else
    If Modo = 1 Then
        If Index = 5 Or Index = 3 Then PierdeFoco3 Index
    End If
End If
End Sub


'Para cuando piede foco y estamos insertando o modificando
Private Sub PierdeFoco3(Indice As Integer)
Dim RC As String
Dim Correcto As Boolean
Dim Valor As Currency
Dim L As Integer
Dim J As Integer
    Text1(Indice).Text = Trim(Text1(Indice).Text)
    If Text1(Indice).Text = "" Then
        'Hemos puesto a blancos el campo, luego quitaremos
        'los valores asociados a el
        If Text1(Indice) = AntiguoText1 Then Exit Sub
        Select Case Indice
        Case 1
            'Ponemos a blanco tb el año de factura
            Text1(26).Text = ""
        
        Case 3
            Text4(4).Text = ""
        Case 5
            Text4(0).Text = ""
            
       Case 6 To 23
               'AUQI AQUI AQUI
            If Indice < 12 Then
                'PRIMERA LINEA
                L = 1
                'Numero de campo k ocupa
                I = Indice - 6
            End If
            
            'Ponemos IVA
            If I = 1 Then
                'Ha puesto a blanco el IVA. Borarmos el resto de campos
                J = (L * 6) + 5
                Text4(L).Text = ""
                For J = Indice To J
                    Text1(J).Text = ""
                Next J
            End If
            'Ha cambiado la base o el iva. Luego hay k recalcular valores
            If I < 2 Then CalcularIVA CInt(L)
            TotalesRecargo
            TotalesIVA
            TotalFactura
        Case 24
            Text2(3).Text = ""
            TotalFactura
        End Select
    Else
        With Text1(Indice)
           Select Case Indice
           Case 1
                If Not EsFechaOK(Text1(Indice)) Then
                    MsgBox "Fecha incorrecta: " & .Text, vbExclamation
                    .Text = ""
                    Text1(26).Text = ""
                    .SetFocus
                    Exit Sub
                End If
                
                'Hay que comprobar que las fechas estan
                'en los ejercicios y si
                '       0 .- Año actual
                '       1 .- Siguiente
                '       2 .- Ambito  !!!!!
                '       3 .- Anterior al inicio
                '       4 .- Posterior al fin
                ModificandoLineas = FechaCorrecta2(CDate(.Text))
                If ModificandoLineas > 1 Then
                    If ModificandoLineas = 2 Then
                        RC = varTxtFec
                    Else
                        If ModificandoLineas = 2 Then
                            RC = "ya esta cerrado"
                        Else
                            RC = " todavia no ha sido abierto"
                        End If
                        RC = "La fecha pertenece a un ejercicio que " & RC
                    End If
                    MsgBox RC, vbExclamation
                    .Text = ""
                    Text1(26).Text = ""
                    .SetFocus
                    Exit Sub
                End If
                
                
                .Text = Format(.Text, "dd/mm/yyyy")
                Text1(26).Text = Year(CDate(.Text))
                'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
                'que el periodo de liquidacion del IVA no ha pasado.
                
                'ATENCION###
                'If Not ComprobarPeriodo Then .SetFocus
                
                
                
                
                
                
            Case 0, 29
                If Not EsFechaOK(Text1(Indice)) Then
                    MsgBox "Fecha incorrecta: " & .Text
                    .Text = ""
                    .SetFocus
                    Exit Sub
                End If
                .Text = Format(.Text, "dd/mm/yyyy")
            Case 3, 5
                'Cuenta cliente
                If AntiguoText1 = .Text Then Exit Sub
                RC = .Text
                If Indice = 3 Then
                    I = 4
                    Else
                    I = 0
                End If
                If CuentaCorrectaUltimoNivel(RC, SQL) Then
                    .Text = RC
                    Text4(I).Text = SQL
                    RC = ""
                Else
                    MsgBox SQL, vbExclamation
                    .Text = ""
                    Text4(I).Text = ""
                    .SetFocus
                End If
                
            Case 7, 13, 19  'TIpos de iva
                I = ((Indice - 1) / 6)
                'If Not IsNumeric(.Text) Then
                If Not EsNumerico(.Text) Then
                    MsgBox "Tipo de iva " & I & " incorrecto:  " & .Text
                    .Text = ""
                    Text4(I).Text = ""
                    .SetFocus
                    Exit Sub
                End If
                If .Text = AntiguoText1 Then Exit Sub
                If PonerValoresIva(I) Then
                    CalcularIVA I
                    TotalesRecargo
                    TotalesIVA
                    TotalFactura
                End If
            Case 6, 12, 18
                'BASES IMPONIBLES
                Correcto = True
                I = ((Indice) / 6)
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
                CalcularIVA I
                TotalesRecargo
                TotalesIVA
                TotalesBase
                TotalFactura
                If Not Correcto Then .SetFocus
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
                        Text2(3).Text = Format(Base, "#,###,##0.00")
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
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
        Else
            'Se muestran en el mismo form
            If CadB <> "" Then
                If SQL <> "" Then SQL = " AND (" & SQL & ")"
                CadB = CadB & SQL
                CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(1), 20, "Recepción: ")
        Cad = Cad & ParaGrid(Text1(26), 10, "Año: ")
        Cad = Cad & ParaGrid(Text1(2), 20, "Nº registro")
        Cad = Cad & ParaGrid(Text1(28), 30, "Nº factura")
        Cad = Cad & ParaGrid(Text1(0), 20, "Fecha fac:")
    
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

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla facturas proveedores.", vbInformation
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
    
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True
    
    'En SQL almacenamos el importe
    Base = Data1.Recordset!totfacpr
'    If Not IsNull(Data1.Recordset!trefaccl) Then
'        Base = Base + Data1.Recordset!trefaccl
'    End If
    SQL = Base
    'Cargamos datos extras
    TotalesBase
    TotalesIVA
    TotalesRecargo
    TotalFactura
    If SQL <> CStr(Aux) Then
        
        MsgBox "Importe factura distinto Importe calculado: " & SQL & " - " & CStr(Aux), vbExclamation
    End If
    
    'Cliente
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(5).Text, "T")
    Text4(0).Text = SQL
    
    'IVAS
    For I = 1 To 3
        kCampo = (I * 6) + 1
        If Text1(kCampo).Text <> "" Then
            SQL = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", Text1(kCampo).Text, "N")
        Else
            SQL = ""
        End If
        Text4(I).Text = SQL
    Next I
    
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
    If Modo = 1 Then
        'Reestablecer colores
        For I = 0 To Text1.Count - 1
            Text1(I).BackColor = vbWhite
        Next I
        Text1(2).Enabled = False
        Text1(2).BackColor = &HFEF7E4
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

        
    B = B Or (Modo = 5)
    DataGrid1.Enabled = B
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    cmdAceptar.Visible = B Or Modo = 1
    'PRueba###
    
    Me.framecabeceras.Enabled = B Or Modo = 1
    

    '
    B = B Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    mnOpcionesAsiPre.Enabled = Not B
   
   

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
End Sub


Private Function DatosOk() As Boolean
    'Dim RS As ADODB.Recordset
    Dim B As Boolean
    
    
    
    'Si no es constructoras igualamos los campos fecfac y fecliquidacion
    If Not vParam.Constructoras Then Text1(29).Text = Text1(1).Text
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
   
    'No puede tener % de retencion sin cuenta de retencion
    If ((Text1(24).Text = "") Xor (Text1(3).Text = "")) Then
        MsgBox "No hay porcentaje de rentencion sin cuenta de retencion", vbExclamation
        B = False
        Exit Function
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

    
    DatosOk = B
End Function



Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        BotonBuscar
    Case 2
        BotonVerTodos
    Case 6
        If Modo <> 5 Then
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
            BotonEliminar
        Else
            'ELIMINAR linea factura
            EliminarLineaFactura
        End If
    Case 10
   
        'Nuevo Modo
        PonerModo 5
        'Fuerzo que se vean las lineas
        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
    Case 11
        If Data1.Recordset.EOF Then Exit Sub
    
        SQL = "Seguro que desea corregir la factura" & vbCrLf
        SQL = SQL & "Numero: " & Data1.Recordset!NumRegis & vbCrLf
        SQL = SQL & "Fecha : " & Data1.Recordset!fecfacpr & "?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
        Screen.MousePointer = vbHourglass
        'Actualizar
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Data1.Recordset.RecordCount = NumRegElim Then
            NumRegElim = NumRegElim - 2
        Else
            NumRegElim = NumRegElim - 1
        End If
        If IntegrarFactura Then
            espera 0.3
            Data1.Refresh
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.Move NumRegElim
                PonerCampos
                PonerModo 2
            Else
                LimpiarCampos
                CargaGrid False
                PonerModo 0
            End If
        End If
        Screen.MousePointer = vbDefault
    Case 13
        Screen.MousePointer = vbHourglass
        HazVerificacion
        Screen.MousePointer = vbDefault
    Case 14
        mnSalir_Click
    Case 16 To 19
        Desplazamiento (Button.Index - 16)
    Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    For I = 16 To 19
        Toolbar1.Buttons(I).Visible = bol
    Next I
End Sub



Private Sub CargaGrid2(Enlaza As Boolean)
    Dim anc As Single
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    Adodc1.ConnectionString = Conn
    Adodc1.RecordSource = MontaSQLCarga(Enlaza)
    Adodc1.CursorType = adOpenDynamic
    Adodc1.LockType = adLockPessimistic
    Adodc1.Refresh
    
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
        txtaux(0).Left = anc + 330
        txtaux(0).Width = DataGrid1.Columns(0).Width - 15
        
        'El boton para CTA
        cmdAux(0).Left = anc + DataGrid1.Columns(1).Left
                
        txtaux(1).Left = cmdAux(0).Left + cmdAux(0).Width
        txtaux(1).Width = DataGrid1.Columns(1).Width - cmdAux(0).Width - 30
        
        If vParam.autocoste Then
            txtaux(2).Left = anc + DataGrid1.Columns(2).Left + 30
            txtaux(2).Width = DataGrid1.Columns(2).Width - 20
        
            cmdAux(1).Left = anc + DataGrid1.Columns(3).Left
            
            txtaux(3).Left = cmdAux(1).Left + cmdAux(1).Width
            txtaux(3).Width = DataGrid1.Columns(3).Width - cmdAux(0).Width - 30
        End If
           
        txtaux(4).Left = anc + DataGrid1.Columns(4).Left + 30
        txtaux(4).Width = DataGrid1.Columns(4).Width - 30
        
        
        If vParam.autocoste Then
            cmdAux(1).Visible = False
        
        End If
        CadAncho = True
    End If
        
    For I = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(I).AllowSizing = False
    Next I
   
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
    
    SQL = "SELECT linfactprove.codtbase, cuentas.nommacta, linfactprove.codccost, cabccost.nomccost, linfactprove.impbaspr, linfactprove.numlinea"
    SQL = SQL & " FROM (cabccost RIGHT JOIN linfactprove ON cabccost.codccost = linfactprove.codccost) LEFT JOIN cuentas ON linfactprove.codtbase = cuentas.codmacta WHERE "
    If Enlaza Then
        SQL = SQL & " numregis = " & Data1.Recordset!NumRegis
        SQL = SQL & " AND anofacpr= " & Data1.Recordset!anofacpr
        Else
        SQL = SQL & " anofacpr = -1"
    End If
    SQL = SQL & " ORDER BY linfactprove.numlinea"
    MontaSQLCarga = SQL
End Function

Private Sub AnyadirLinea(Limpiar As Boolean, DesdeBoton As Boolean)
    Dim anc As Single
    
    If ModificandoLineas = 2 Then Exit Sub
    'Obtenemos la siguiente numero de factura
    Linfac = ObtenerSigueinteNumeroLinea   'Fijamos en aux el importe que queda
    If Aux = 0 Then
        anc = 0
        If DesdeBoton Then
            If MsgBox("Importes exactos. ¿Continuar?", vbQuestion + vbYesNo) = vbYes Then anc = 1
        End If
        If anc = 0 Then
            LLamaLineas anc, 0, True
            cmdCancelar.Caption = "Cabecera"
            Exit Sub
        End If
    End If
    
    
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        Adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If
    LLamaLineas anc, 1, Limpiar
    'Ponemos el importe
    
    txtaux(4).Text = Aux
    HabilitarCentroCoste
    'Ponemos el foco
    txtaux(0).SetFocus
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 15
    End If

    'Asignar campos
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtaux(1).Text = DataGrid1.Columns(1).Text
    If vParam.autocoste Then
        txtaux(2).Text = DataGrid1.Columns(2).Text
        txtaux(3).Text = DataGrid1.Columns(3).Text
    End If
    txtaux(4).Text = Adodc1.Recordset!impbaspr

    LLamaLineas anc, 2, False
    HabilitarCentroCoste
    txtaux(0).SetFocus
End Sub

Private Sub EliminarLineaFactura()
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub
    If Adodc1.Recordset.EOF Then Exit Sub
    If ModificandoLineas <> 0 Then Exit Sub
    SQL = "Lineas de factura." & vbCrLf & vbCrLf
    SQL = SQL & "Va a eliminar la linea: " & vbCrLf
    SQL = SQL & Adodc1.Recordset.Fields(0) & " - " & Adodc1.Recordset.Fields(1) & ": " & Adodc1.Recordset.Fields(4)
    SQL = SQL & vbCrLf & vbCrLf & "     Desea continuar? "
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        SQL = "Delete from linfactprove"
        SQL = SQL & " WHERE numlinea = " & Adodc1.Recordset!NumLinea
        SQL = SQL & " AND anofacpr=" & Data1.Recordset!anofacpr
        SQL = SQL & " AND numregis = " & Data1.Recordset!NumRegis & ";"
        Conn.Execute SQL
        CargaGrid (Not Data1.Recordset.EOF)
    End If
End Sub


'Ademas de obtener el siguiente nº de linea, tb obtiene la suma de
'las lineas de factura, Y fijamos lo que falta en aux para luego ofertarlo

Private Function ObtenerSigueinteNumeroLinea() As Long
    Dim RS As ADODB.Recordset
    Dim I As Long
    
    Set RS = New ADODB.Recordset
    
    SQL = " WHERE linfactprove.numregis= " & Data1.Recordset!NumRegis
    SQL = SQL & " AND linfactprove.anofacpr=" & Data1.Recordset!anofacpr & ";"
    RS.Open "SELECT Max(numlinea) FROM linfactprove" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then I = RS.Fields(0)
    End If
    RS.Close

    'La suma
    SumaLinea = 0
    If I > 0 Then
        RS.Open "SELECT sum(impbaspr) FROM linfactprove" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then SumaLinea = RS.Fields(0)
        End If
        RS.Close
    End If
    Set RS = Nothing
    
    'Lo que falta lo fijamos en aux. El importe es de la bASE IMPONIBLE si fuera del total seria Text2(4).Text
    Aux = ImporteFormateado(Text2(0).Text)
    Aux = Aux - SumaLinea
    ObtenerSigueinteNumeroLinea = I + 1
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
        txtaux(3).Visible = Visible
        txtaux(2).Visible = Visible
        cmdAux(1).Top = Altura
    Else
        txtaux(3).Visible = False
        txtaux(2).Visible = False
        txtaux(3).Text = ""
        txtaux(2).Text = ""
        cmdAux(1).Visible = False
    End If
    For I = 0 To txtaux.Count - 1
        If I < 2 Or I > 3 Then txtaux(I).Visible = Visible
        txtaux(I).Top = Altura
    Next I

    If Limpiar Then
        For I = 0 To txtaux.Count - 1
            txtaux(I).Text = ""
        Next I
    End If
    
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
With txtaux(Index)
    If Index <> 5 Then
         .SelStart = 0
        .SelLength = Len(.Text)
    Else
        .SelStart = Len(.Text)
    End If
End With

End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    Dim RC As String
    Dim Sng As Double
        
        If ModificandoLineas = 0 Then Exit Sub
        
        'Comprobaremos ciertos valores
        txtaux(Index).Text = Trim(txtaux(Index).Text)
    
        'Comun a todos
        If txtaux(Index).Text = "" Then
            Select Case Index
            Case 0
                txtaux(1).Text = ""
                HabilitarCentroCoste
            Case 2
                txtaux(3).Text = ""
            End Select
            Exit Sub
        End If
        
        Select Case Index
        Case 0
            'Cta
            
            RC = txtaux(0).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(0).Text = RC
                txtaux(1).Text = SQL
                RC = ""
            Else
                MsgBox SQL, vbExclamation
                txtaux(0).Text = ""
                txtaux(1).Text = ""
                RC = "NO"
            End If
            HabilitarCentroCoste
            If RC <> "" Then
                txtaux(0).SetFocus
            Else
                If txtaux(2).Visible And txtaux(2).Enabled Then
                    txtaux(2).SetFocus
                Else
                    txtaux(4).SetFocus
                End If
            End If
        Case 2
            txtaux(2).Text = UCase(txtaux(2).Text)
            RC = "idsubcos"
            SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtaux(2).Text, "T", RC)
            If SQL = "" Then
                MsgBox "Centro de coste no encontrado: " & txtaux(2).Text, vbExclamation
                txtaux(2).Text = ""
                txtaux(2).SetFocus
            End If
            txtaux(3).Text = SQL
            If SQL <> "" Then txtaux(4).SetFocus
        Case 4
            If Not EsNumerico(txtaux(4).Text) Then
               ' MsgBox "Importe incorrecto: " & txtaux(4).Text, vbExclamation
                txtaux(4).Text = ""
                txtaux(4).SetFocus
            Else
                txtaux(4).Text = TransformaPuntosComas(txtaux(4).Text)
                cmdAceptar.SetFocus
            End If
            
        End Select
End Sub


Private Function AuxOK() As String
    
    'Cuenta
    If txtaux(0).Text = "" Then
        AuxOK = "Cuenta no puede estar vacia."
        Exit Function
    End If
    If Len(txtaux(0).Text) <> vEmpresa.DigitosUltimoNivel Then
        AuxOK = "Longitud cuenta incorrecta"
        Exit Function
    End If
    If EstaLaCuentaBloqueada(txtaux(0).Text, CDate(Text1(1).Text)) Then
        AuxOK = "Cuenta bloqueada: " & txtaux(0).Text
        Exit Function
    End If
    'Importe
    If txtaux(4).Text = "" Then
        AuxOK = "El importe no puede estar vacio"
        Exit Function
    End If
        
    If txtaux(4).Text <> "" Then
        If Not IsNumeric(txtaux(4).Text) Then
            AuxOK = "El importe debe de ser numérico."
            Exit Function
        End If
    End If
    
    'cENTRO DE COSTE
    If txtaux(2).Visible Then
        If txtaux(2).Enabled Then
            If txtaux(2).Text = "" Then
                AuxOK = "Centro de coste no puede ser nulo"
                Exit Function
            End If
        End If
    End If
    
    
    'Vemos los importes
    '--------------------------
    'Total factura en AUX
    Aux = ImporteFormateado(Text2(4).Text)
    
    'Si tiene retencion
    AUX2 = 0
    If Text2(3).Text <> "" Then AUX2 = ImporteFormateado(Text2(3).Text)
    Aux = Aux + AUX2
    
    'El iVA
    AUX2 = 0
    If Text2(1).Text <> "" Then AUX2 = ImporteFormateado(Text2(1).Text)
    Aux = Aux - AUX2
    
    
    
    
    'Importe linea en aux2
    AUX2 = CCur(txtaux(4).Text)
    
    'Suma de las lineas mas lo que acabamos de poner
    AUX2 = AUX2 + SumaLinea
    
    'Si estamos insertando entonces la suma de lineas + aux2 no debe ser superior a toal fac
    If ModificandoLineas = 2 Then
        'Si estasmos insertando no hacemos nada puesto que el importe son las sumas directamente
       'Estamos modificando, hay que quitarle el importe que tenia
        AUX2 = AUX2 - Adodc1.Recordset!impbaspr
    End If
    If Aux > 0 Then
        If AUX2 > Aux Then
               ' AuxOK = "El importe excede de la base"
               ' Exit Function
        End If
    Else
    
        If AUX2 < Aux Then
                'AuxOK = "El importe excede de la base"
                'Exit Function
        End If
    End If
    
    AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        SQL = "INSERT INTO linfactprove (numregis, anofacpr, numlinea, codtbase, impbaspr, codccost) VALUES ("
        ''R', 11, 2003, 1, '6000001', 1500, 'TIEN')
        SQL = SQL & Data1.Recordset!NumRegis & ","
        SQL = SQL & Data1.Recordset!anofacpr & "," & Linfac & ",'"
        'Cuenta
        SQL = SQL & txtaux(0).Text & "',"
        'Importe
        SQL = SQL & TransformaComasPuntos(txtaux(4).Text) & ","
   
        'Centro coste
        If txtaux(2).Text = "" Then
          SQL = SQL & ValorNulo
          Else
          SQL = SQL & "'" & txtaux(2).Text & "'"
        End If
        SQL = SQL & ")"
        
    Else
    
        'MODIFICAR
        'UPDATE linasipre SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
        '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
        SQL = "UPDATE linfactprove SET "
        
        SQL = SQL & " codtbase = '" & txtaux(0).Text & "',"
        SQL = SQL & " impbaspr = "
        SQL = SQL & TransformaComasPuntos(txtaux(4).Text) & ","
        
        'Centro coste
        If txtaux(2).Text = "" Then
          SQL = SQL & " codccost = " & ValorNulo
          Else
          SQL = SQL & " codccost = '" & txtaux(2).Text & "'"
        End If
    
        SQL = SQL & " WHERE numregis= " & Data1.Recordset!NumRegis
        SQL = SQL & " AND anofacpr=" & Data1.Recordset!anofacpr
        SQL = SQL & " AND numlinea =" & Adodc1.Recordset!NumLinea & ";"

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
DataGrid1.Enabled = False
CargaGrid2 Enlaza
DataGrid1.Enabled = B
End Sub

Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub


Private Function PonerValoresIva(numero As Integer) As Boolean
Dim J As Integer
J = ((numero - 1) * 6) + 7
Set RS = New ADODB.Recordset
RS.Open "Select nombriva,porceiva,porcerec from tiposiva where codigiva =" & Text1(J).Text, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
If RS.EOF Then
    MsgBox "Tipo de IVA incorrecto: " & Text1(J).Text, vbExclamation
    Text1(J).Text = ""
    Text4(numero).Text = ""
    PonerValoresIva = False
Else
    PonerValoresIva = True
    
    Text4(numero).Text = RS.Fields(0)
    Text1(J + 1).Text = Format(RS.Fields(1), "#0.00")
    Aux = DBLet(RS.Fields(2))
    If Aux = 0 Then
        Text1(J + 3).Text = ""
        Else
        Text1(J + 3).Text = Format(Aux, "#0.00")
    End If
End If
RS.Close
End Function



Private Sub CalcularIVA(numero As Integer)
Dim J As Integer


J = ((numero - 1) * 6) + 6
Base = ImporteFormateado(Text1(J).Text)

'EL iva
Aux = ImporteFormateado(Text1(J + 2).Text) / 100
If Aux = 0 Then
    If Text1(J + 2).Text = "" Then
        Text1(J + 3).Text = ""
    Else
        Text1(J + 3).Text = "0,00"
    End If
Else
    Text1(J + 3).Text = Format(Round((Aux * Base), 2), FormatoImporte)
End If

'Recargo
Aux = ImporteFormateado(Text1(J + 4).Text) / 100
If Aux = 0 Then
    Text1(J + 5).Text = ""
Else
    Text1(J + 5).Text = Format(Round((Aux * Base), 2), FormatoImporte)
End If

End Sub


Private Sub TotalesBase()
    'Base imponible
    Aux = 0
    For I = 1 To 3
        If Text1(I * 6).Text <> "" Then
            Base = ImporteFormateado(Text1(I * 6).Text)
            Aux = Aux + Base
        End If
    Next I
    If Aux = 0 Then
        Text2(0).Text = ""
    Else
        Text2(0).Text = Format(Aux, FormatoImporte)
    End If
End Sub


Private Sub TotalesIVA()
    'Total IVA
    Aux = 0
    For I = 1 To 3
        ancho = (I * 6) + 3
        If Text1(ancho).Text <> "" Then
            Base = ImporteFormateado(Text1(ancho).Text)
            Aux = Aux + Base
        End If
    Next I
    If Aux = 0 Then
        Text2(1).Text = ""
    Else
        Text2(1).Text = Format(Aux, FormatoImporte)
    End If
End Sub

Private Sub TotalesRecargo()
    'RECARGO
    Aux = 0
    For I = 1 To 3
        ancho = (I * 6) + 5
        If Text1(ancho).Text <> "" Then
            Base = ImporteFormateado(Text1(ancho).Text)
            Aux = Aux + Base
        End If
    Next I
    If Aux = 0 Then
        Text2(2).Text = ""
    Else
        Text2(2).Text = Format(Aux, FormatoImporte)
    End If
End Sub

Private Sub TotalFactura()
    'El total
    Aux = 0
    ' Base + iva + recargao   -  retencion
    For I = 0 To 2
        If Text2(I).Text = "" Then
   
        Else
            Base = ImporteFormateado(Text2(I).Text)
            Aux = Aux + Base
        End If
    Next I
    If Text2(3).Text = "" Then
        
    Else
        Base = ImporteFormateado(Text2(3).Text)
        Aux = Aux - Base
    End If
    
    If Aux = 0 Then
        Text2(4).Text = ""
    Else
        Text2(4).Text = Format(Aux, FormatoImporte)
    End If
    Text1(27).Text = TransformaComasPuntos(CStr(Aux))
End Sub


'Comprobara si el periodo esta liquidado o no.
'Si la fecha pertenece a un periodo liquidado entonces
'mostraremos un mensaje para preguntar si desea continuar o no
Private Function ComprobarPeriodo() As Boolean
Dim Cerrado As Boolean
'Primero pondremos la fecha a año periodo
I = Year(CDate(Text1(0).Text))
If vParam.periodos Then
    'Trimestral
    ancho = (CDate(Text1(0).Text) Mod 3) + 1
    Else
    ancho = Month(CDate((Text1(0).Text)))
End If
Cerrado = False
If I < vParam.anofactu Then
    Cerrado = True
Else
    If I = vParam.anofactu Then
        'El mismo año. Comprobamos los periodos
        If vParam.perfactu >= ancho Then _
            Cerrado = True
    End If
End If
ComprobarPeriodo = True
If Cerrado Then
    SQL = "La fecha corresponde a un periodo ya liquidado. " & vbCrLf & " ¿Desea continuar igualmente ?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then ComprobarPeriodo = False
End If
End Function




Private Sub HabilitarCentroCoste()
Dim hab As Boolean

    hab = False
    If vParam.autocoste Then
        If txtaux(0).Text <> "" Then
            hab = HayKHabilitarCentroCoste(txtaux(0).Text)
        Else
            txtaux(2).Text = ""
            txtaux(3).Text = ""
        End If
        If hab Then
            txtaux(2).BackColor = &H80000005
            Else
            txtaux(2).BackColor = &H80000018
        End If
    End If
    txtaux(2).Enabled = hab
    cmdAux(1).Enabled = hab
    Me.Refresh
End Sub



Private Function Siguiente() As Long
Dim RS As Recordset
    Set RS = New ADODB.Recordset
    SQL = "Select max(numregis) from cabfactprove"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Siguiente = 0
    If Not RS.EOF Then Siguiente = DBLet(RS.Fields(0), "N")
    RS.Close
    Set RS = Nothing
    Siguiente = Siguiente + 1
End Function



Private Function ActualizarRegistros() As Boolean
On Error GoTo EACt
    ActualizarRegistros = False
    'UPDATEAMOS AL CONTADOR devuelto
        SQL = "UPDATE cabfactprove set numregis=" & Linfac & " WHERE numregis=" & Data1.Recordset!NumRegis
        SQL = SQL & " AND anofacpr = " & Data1.Recordset!anofacpr
        Conn.Execute SQL
        
        
        SQL = "UPDATE linfactprove set numregis=" & Linfac & " WHERE numregis=" & Data1.Recordset!NumRegis
        SQL = SQL & " AND anofacpr = " & Data1.Recordset!anofacpr
        Conn.Execute SQL
        
    
        SQL = "INSERT INTO cabfactprov SELECT * from cabfactprove WHERE "
        SQL = SQL & " numregis = " & Linfac
        SQL = SQL & " AND anofacpr= " & Data1.Recordset!anofacpr
        Conn.Execute SQL
    
        SQL = "INSERT INTO linfactprov SELECT * from linfactprove WHERE "
        SQL = SQL & " numregis = " & Linfac
        SQL = SQL & " AND anofacpr= " & Data1.Recordset!anofacpr
        Conn.Execute SQL
    

    ActualizarRegistros = True
    Exit Function
EACt:
    MuestraError Err.Number, "Actualizar registros"
End Function









Private Sub HazVerificacion()
Dim VC As String
Dim RT As ADODB.Recordset
   
    Set RT = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    AntiguoText1 = ""
    RS.Open "Select numregis,anofacpr from linfactprove group by numregis,anofacpr ", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        VC = "Select numregis from cabfactprove where anofacpr=" & RS!anofacpr
        VC = VC & " AND numregis = " & RS!NumRegis
        RT.Open VC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RT.EOF Then AntiguoText1 = AntiguoText1 & Format(RS!NumRegis, "00000000") & "      " & RS!anofaccl & vbCrLf
        RT.Close
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Set RT = Nothing
    Screen.MousePointer = vbDefault
    If AntiguoText1 <> "" Then
        VC = " Las siguientes lineas no corresponden a ningun encabezado de factura de proveedores erronea." & vbCrLf & vbCrLf & _
            "   Codigo     Año  " & vbCrLf & "----------------------------" & vbCrLf & AntiguoText1
        MsgBox VC, vbExclamation
    Else
        MsgBox "Comprobación finalizada", vbInformation
    End If
    
End Sub


