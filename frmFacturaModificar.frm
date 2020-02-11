VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacturaModificar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar factura SII"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1200
      Index           =   1
      Left            =   11160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Tag             =   "Cuenta Retencion|T|S|||factpro|cuereten|||"
      Text            =   "frmFacturaModificar.frx":0000
      Top             =   1080
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
      Index           =   0
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text4"
      Top             =   2760
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
      Height          =   360
      Index           =   0
      Left            =   9360
      TabIndex        =   32
      Tag             =   "Cuenta Retencion|T|S|||factpro|cuereten|||"
      Text            =   "1234567890"
      Top             =   2760
      Width           =   1575
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
      Left            =   14760
      TabIndex        =   12
      Top             =   8160
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   16080
      TabIndex        =   13
      Top             =   8160
      Width           =   1035
   End
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   120
      TabIndex        =   17
      Top             =   3480
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
         Index           =   5
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   24
         Text            =   "Nombre cuenta"
         Top             =   2190
         Visible         =   0   'False
         Width           =   3285
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
         TabIndex        =   23
         ToolTipText     =   "Buscar cuenta"
         Top             =   2190
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
         Index           =   1
         Left            =   840
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Nº factura|N|N|0||factcli_lineas|numfactu|000000|S|"
         Text            =   "factura"
         Top             =   2145
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   0
         Tag             =   "Nº Serie|T|S|||factcli_lineas|numserie||S|"
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
         Index           =   2
         Left            =   2220
         TabIndex        =   2
         Tag             =   "Fecha|F|N|||factcli_lineas|fecfactu|dd/mm/yyyy||"
         Text            =   "fecha"
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
         TabIndex        =   4
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
         Index           =   5
         Left            =   4050
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "Cuenta|T|N|||factcli_lineas|codmacta|||"
         Text            =   "Cta Base"
         Top             =   2160
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   60
         TabIndex        =   21
         Top             =   0
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   180
            TabIndex        =   22
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
         TabIndex        =   7
         Tag             =   "Codigo Iva|N|N|||factcli_lineas|codigiva|000||"
         Text            =   "Iva"
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
         TabIndex        =   20
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
         Index           =   6
         Left            =   8370
         MaxLength       =   10
         TabIndex        =   6
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
         Index           =   2
         Left            =   15420
         TabIndex        =   19
         ToolTipText     =   "Buscar concepto"
         Top             =   2130
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
         Index           =   8
         Left            =   10140
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "% Iva|N|S|||factcli_lineas|porciva|##0.00||"
         Text            =   "%iva"
         Top             =   2160
         Visible         =   0   'False
         Width           =   795
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
         TabIndex        =   14
         Tag             =   "% Recargo|N|S|||factcli_lineas|porcrec|##0.00||"
         Text            =   "%rec"
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
         Index           =   10
         Left            =   12090
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Importe Iva|N|S|||factcli_lineas|impoiva|###,###,##0.00||"
         Text            =   "ImpIva"
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
         Index           =   12
         Left            =   14520
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "CC|T|S|||factcli_lineas|codccost|||"
         Text            =   "CC"
         Top             =   2160
         Visible         =   0   'False
         Width           =   885
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
         TabIndex        =   3
         Tag             =   "Año factura|N|N|||factcli_lineas|anofactu||S|"
         Text            =   "año"
         Top             =   2160
         Visible         =   0   'False
         Width           =   345
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
         TabIndex        =   9
         Tag             =   "Importe Rec|N|S|||factcli_lineas|imporec|###,###,##0.00||"
         Text            =   "ImpRec"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chkAux 
         BackColor       =   &H80000005&
         Height          =   255
         Index           =   0
         Left            =   14250
         TabIndex        =   10
         Tag             =   "Aplica Retencion|N|N|0|1|factcli_lineas|aplicret|||"
         Top             =   2190
         Visible         =   0   'False
         Width           =   285
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
         TabIndex        =   18
         Text            =   "Nombre cuenta"
         Top             =   2160
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
         TabIndex        =   25
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
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   17295
      Begin VB.Label Label1 
         Caption         =   "Desglose IVAs"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
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
         TabIndex        =   29
         Top             =   240
         Width           =   14760
      End
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   1785
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   1320
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   3149
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
   Begin MSComctlLib.ListView lw1 
      Height          =   1545
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   7200
      Visible         =   0   'False
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
   Begin VB.Image imgppal 
      Height          =   240
      Index           =   0
      Left            =   11280
      Top             =   2400
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta retencion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   9360
      TabIndex        =   31
      Top             =   2400
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   9360
      TabIndex        =   30
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Desglose IVAs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   960
      Width           =   1560
   End
End
Attribute VB_Name = "frmFacturaModificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cliente As Boolean
Public NUmSerie As String
Public Codigo As Long
Public Anyo As Integer
Public Fecha As Date    'fecfactu    fecharec

Private WithEvents frmTIva As frmBasico2
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmC2 As frmColCtas
Attribute frmC2.VB_VarHelpID = -1



Dim SQL As String
Dim AntiguoText1 As String

Dim Modo2 As Byte
Dim ModoLineas As Byte

Dim HaCambiadoLineas As Byte



Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, Modo2, False
End Sub

Private Sub cmdAceptar_Click()
Dim B As Boolean
Dim Importe As Currency
Dim importe2 As Currency
Dim CambiaCabecera As Boolean
Dim TipoRetencion As Integer
    Ampliacion = ""

    If Modo2 = 5 Then
            B = False
            'lineas
            Select Case ModoLineas
                Case 1 'afegir llínia
                    If InsertarLinea Then B = True
                    
                Case 2 'modificar llínies
            
                    If ModificarLinea Then B = True
                    
            End Select
            If B Then
                VariableCambios True
                Me.DataGridAux(1).AllowAddNew = False
                LLamaLineas 1, 0
                Modo2 = 2
                CargaGrid 1, True
                RecalcularTotales
            End If
        Else
            CambiaCabecera = False
            SQL = ""
            If Text1(0).visible Then
                If Text1(0).Text <> Text1(0).Tag Then
                    'Ha cambiado lo que habia en la cuenta de retencion
                    If Trim(Text1(0).Text) = "" Then
                        SQL = "Debe indicar cuenta retencion"
                    Else
                        If Text4(0).Text = "" Then SQL = "Error en cuenta de rentencion"
                    End If
                    
                    If SQL = "" Then CambiaCabecera = True
                        
                    
                End If
                
            End If
            If Text1(1).Text <> Text1(1).Tag Then CambiaCabecera = True
            If SQL <> "" Then
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
            
            '                               los 10 primeros codmacta
            If CambiaCabecera Then Ampliacion = Mid(Text1(0).Text & Space(10), 1, 10) & Text1(1).Text
            
            
                    
            'VALIDAR LOS CAMBIOS
            If HaCambiadoLineas = 1 Then
                'Primera comprobacion.
                'Que los IVAS son iguales
                SQL = ""
                If lw1(0).ListItems.Count <> lw1(1).ListItems.Count Then
                    SQL = "No coinciden los IVAS a modificar "
                Else
                    'Tienen el mismo numero de IVA
                    For NumRegElim = 1 To lw1(0).ListItems.Count
                        AntiguoText1 = ""
                        For ModoLineas = 1 To lw1(1).ListItems(1).ListSubItems.Count
                            
                            If Trim(lw1(1).ListItems(NumRegElim).ListSubItems(ModoLineas).Text) <> Trim(lw1(0).ListItems(NumRegElim).ListSubItems(ModoLineas).Text) Then
                                AntiguoText1 = AntiguoText1 & lw1(1).ColumnHeaders(ModoLineas + 1).Text & ":  " & lw1(0).ListItems(NumRegElim).ListSubItems(ModoLineas).Text & "   /    " & lw1(1).ListItems(NumRegElim).ListSubItems(ModoLineas).Text & vbCrLf
                            End If
                        Next
                        If AntiguoText1 <> "" Then SQL = SQL & vbCrLf & "IVA: " & NumRegElim & vbCrLf & AntiguoText1
                    Next
                    AntiguoText1 = ""
                End If
                ModoLineas = 0
                
                If SQL <> "" Then
                    MsgBox SQL, vbExclamation
                    Exit Sub
                End If
                
                
                
                'Retencion. Veremos si la factura lleva retencion
                
                
                Set miRsAux = New ADODB.Recordset
                SQL = "SELECT tiporeten  ,totbasesret FROM "
                SQL = SQL & IIf(Cliente, "factcli", "factpro") & " WHERE numserie = '" & NUmSerie & "' AND "
                SQL = SQL & IIf(Cliente, "numfactu", "numregis") & "  = " & Codigo & " AND anofactu =" & Anyo
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                'NO PUEDE SER EOF
                SQL = ""
                If miRsAux!tiporeten > 0 Then
                    Importe = DBLet(miRsAux!totbasesret, "N")
                    NumRegElim = miRsAux!tiporeten
                    TipoRetencion = DevuelveValor("select tipo from usuarios.wtiporeten where codigo = " & NumRegElim)
                    SQL = "TIENE"
                End If
                miRsAux.Close
                
                If SQL <> "" Then
                    SQL = "Select sum(imponible) bases,sum(impiva) iva,sum(coalesce(retencion,0)) rec from tmpfaclin"
                    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND tipoopera =1"
                    importe2 = 0
                    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not miRsAux.EOF Then
                        importe2 = DBLet(miRsAux!Bases, "N")
                        If TipoRetencion = 0 Then
                            'Sobre bases
                            'Ya esta sumado
                        Else
                            importe2 = importe2 + DBLet(miRsAux!IVA, "N") + DBLet(miRsAux!rec, "N")
                        End If
                    End If
                    
                    If importe2 <> Importe Then
                        SQL = "Importe retencion no coincide con la factura original." & vbCrLf & "Factura: " & Format(Importe, FormatoImporte) & vbCrLf
                        SQL = SQL & "Modificada: " & Format(importe2, FormatoImporte) & vbCrLf
                        MsgBox SQL, vbExclamation
                    Else
                        SQL = ""
                    End If
                    miRsAux.Close
                End If
               
                Set miRsAux = Nothing
                If SQL <> "" Then Exit Sub
                
                
                'PREGUNTAMOS
                ModoLineas = MsgBox("¿Desea modificar la factura ?", vbQuestion + vbYesNoCancel)
                If ModoLineas = vbCancel Then Exit Sub
                
                If ModoLineas = vbYes Then
                    SQL = "Select * from tmpfaclin where codusu =" & vUsu.Codigo & " ORDER BY numfac"
                    Set miRsAux = New ADODB.Recordset
                    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    SQL = ""
                    While Not miRsAux.EOF
                        'factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost)
                        'factcli_lineas(numserie,numfactu,fecfactu,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost)
                        SQL = SQL & ", (" & DBSet(NUmSerie, "T") & "," & Codigo & "," & DBSet(Fecha, "F") & "," & Anyo & ","
                        SQL = SQL & Val(miRsAux!NumFac) & "," & DBSet(miRsAux!Cta, "T") & "," & DBSet(miRsAux!Imponible, "N")
                        SQL = SQL & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux!IVA, "N") & "," & DBSet(miRsAux!porcrec, "N")
                        SQL = SQL & "," & DBSet(miRsAux!ImpIva, "N") & "," & DBSet(miRsAux!recargo, "N") & "," & Val(miRsAux!tipoopera)
                        SQL = SQL & "," & DBSet(miRsAux!NumFactura, "T", "S") & ")"
                        miRsAux.MoveNext
                    Wend
                    miRsAux.Close
                    Set miRsAux = Nothing
                    CadenaDesdeOtroForm = Mid(SQL, 2)
                Else
                    CadenaDesdeOtroForm = ""
                End If
            Else
                If CambiaCabecera Then
                    ModoLineas = MsgBox("¿Desea modificar la factura ?", vbQuestion + vbYesNoCancel)
                    If ModoLineas = vbCancel Then Exit Sub
                    
                    If ModoLineas = vbNo Then Ampliacion = ""
                    
                End If
            
                CadenaDesdeOtroForm = ""
                
            End If
            Unload Me
            
        End If
    

End Sub

Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 ' cuenta base
            cmdAux(0).Tag = 0
            LlamaContraPar
            If txtaux(4).Text <> "" Then
                PonFoco txtaux(5)
            Else
                PonFoco txtaux(4)
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
'    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub



Private Sub cmdCancelar_Click()
    If Modo2 = 2 Then
        'CANCELAMOS TODOS LOS CAMBIOS
        If HaCambiadoLineas > 0 Then
            If MsgBox("¿Desea salir sin realizar las modificaciones?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        CadenaDesdeOtroForm = ""
        Unload Me
        
    Else
        Me.DataGridAux(1).AllowAddNew = False
        LLamaLineas 1, 0
        Modo2 = 2
        CargaGrid 1, True
    End If
End Sub

Private Sub Form_Load()
        
    
    Me.Icon = frmppal.Icon
    Me.Caption = "Modificar factura " & IIf(Cliente, "cliente", "proveedor")
    ' Botonera Principal 2
    With Me.ToolbarAux
         .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5

    End With
    imgppal(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    CargarColumnas 0
    CargarColumnas 1
    CargaDatosLW2 1   'Este es el IVA ORIGINAL
    Modo2 = 2
    HaCambiadoLineas = 0
    SQL = "INSERT INTO tmpfaclin (codusu,numserie,codigo,Fecha,numfac,cta,Imponible,tipoiva,IVA,porcrec,ImpIVA,recargo,numfactura,tipoopera)"
    
    
    If Cliente Then
        SQL = SQL & " select " & vUsu.Codigo & ", Numserie , NumFactu, FecFactu, NumLinea, codmacta, Baseimpo, codigiva, porciva, porcrec,"
        SQL = SQL & "  Impoiva, ImpoRec,  codccost,aplicret From factcli_lineas "
        Label1(0).Caption = "Factura cliente. " & NUmSerie & Format(Codigo, "00000") & " de " & Format(Fecha, "dd/mm/yyyy")
        Label1(0).ForeColor = &H8000&
        
    Else
        SQL = SQL & " select " & vUsu.Codigo & ",Numserie , Numregis, fecharec,  NumLinea, codmacta, Baseimpo, codigiva, porciva, porcrec, "
        SQL = SQL & " Impoiva, ImpoRec,  codccost,aplicret From factpro_lineas "
        Label1(0).Caption = "Factura proveedor. Nº Registro" & IIf(NUmSerie = 1, "", NUmSerie) & " " & Format(Codigo, "00000") & " de " & Format(Fecha, "dd/mm/yyyy")
        Label1(0).ForeColor = &H80&
    End If
    
    SQL = SQL & " WHERE numserie = " & DBSet(NUmSerie, "T") & " AND anofactu = " & Anyo & " AND "
    SQL = SQL & IIf(Cliente, "numfactu", "numregis") & " = " & Codigo
    
    Conn.Execute SQL
    
    CargaGrid 1, True
    
    
    'Observaciones y cuenta retencion... su tuviera
    AntiguoText1 = "concat(coalesce(cuereten,''),'#',coalesce(observa,''))"
    SQL = " numserie = " & DBSet(NUmSerie, "T") & " AND anofactu = " & Anyo & " AND "
    SQL = SQL & IIf(Cliente, "numfactu", "numregis") & " = " & Codigo
    
    AntiguoText1 = DevuelveDesdeBD(AntiguoText1, IIf(Cliente, "factcli", "factpro"), SQL, "1")
    NumRegElim = InStr(1, AntiguoText1, "#")
    SQL = ""
    If NumRegElim > 0 Then
        SQL = Mid(AntiguoText1, 1, NumRegElim - 1)
        AntiguoText1 = Mid(AntiguoText1, NumRegElim + 1)
    Else
        AntiguoText1 = ""
    End If
    Text1(0).Text = SQL
    Text1(1).Text = AntiguoText1
    
    Text1(0).visible = SQL <> ""
    Label1(3).visible = SQL <> ""
    imgppal(0).visible = SQL <> ""
    Text4(0).visible = SQL <> ""
    Text1(0).Tag = SQL
    Text1(1).Tag = AntiguoText1
    If SQL <> "" Then
        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", SQL, "T")
        Text4(0).Text = SQL
    End If


    
End Sub





Private Sub imgppal_Click(Index As Integer)
    Set frmC2 = New frmColCtas
    SQL = ""
    cmdAux(0).Tag = 2
    frmC2.DatosADevolverBusqueda = "0|1|"
    frmC2.ConfigurarBalances = 3  'NUEVO
    frmC2.Show vbModal
    Set frmC2 = Nothing
    If SQL <> "" Then
        If SQL <> Text1(0).Text Then
            Text1(0).Text = RecuperaValor(SQL, 1)
            Text4(0).Text = RecuperaValor(SQL, 2)
        End If
    End If
    imgppal(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Index = 0 Then ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
        If Index = 0 Then
                AntiguoText1 = Text1(Index).Text
                If CuentaCorrectaUltimoNivel(AntiguoText1, SQL) Then
                    Text1(Index).Text = AntiguoText1
                    Text4(Index).Text = SQL
                    If Text1(Index).Text <> "" Then
                        If EstaLaCuentaBloqueada(AntiguoText1, Fecha) Then
                            MsgBox "Cuenta bloqueada: " & AntiguoText1, vbExclamation
                            Text1(Index).Text = ""
                            Text4(Index).Text = ""
                            PonFoco Text1(Index)
                            Exit Sub
                        End If
                    End If
                Else
                    Text1(0).Text = ""
                    Text4(0).Text = ""
                End If
        End If
        
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    AntiguoText1 = txtaux(Index).Text
    ConseguirFoco txtaux(Index), Modo2
    
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
        KEYpressGnral KeyAscii, Modo2, False
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
        If Not PerderFocoGnral(txtaux(Index), Modo2) Then Exit Sub
        
        If txtaux(Index).Text = AntiguoText1 Then
             If Index = 12 And vParam.autocoste Then cmdAceptar.SetFocus
             Exit Sub
        End If
    
        CalcularElIva = True
        Select Case Index
        Case 5
            RC = txtaux(5).Text
            If CuentaCorrectaUltimoNivel(RC, SQL) Then
                txtaux(5).Text = RC
    
                If EstaLaCuentaBloqueada(RC, Fecha) Then
                    MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                    txtaux(5).Text = ""
                Else
                    txtAux2(5).Text = SQL
                    ' traemos el tipo de iva de la cuenta
                    If ModoLineas = 1 Then
                        txtaux(7).Text = DevuelveDesdeBD("codigiva", "cuentas", "codmacta", txtaux(5).Text, "N")
                       
                        If txtaux(7).Text <> "" Then txtAux_LostFocus (7)
                    Else
                        CalcularElIva = False
                    End If
                    
                    RC = ""
                End If
            Else
                If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                    'NO EXISTE LA CUENTA, añado que debe de tener permiso de creacion de cuentas
                                            MsgBox SQL, vbExclamation
                Else
                    MsgBox SQL, vbExclamation
                End If
                    
                If SQL <> "" Then
                  txtaux(5).Text = ""
                  txtAux2(5).Text = ""
                  RC = "NO"
                End If
            End If
            HabilitarCentroCoste
            If RC <> "" Then PonFoco txtaux(5)
                
            If Modo2 = 5 And ModoLineas = 1 Then MostrarObservaciones txtaux(Index)
            
        Case 6
            If Not PonerFormatoDecimal(txtaux(Index), 1) Then
                txtaux(Index).Text = ""
            Else
                'Si modificando lienas, no cambia el importe NO recalculo iVA
                If Modo2 = 5 And ModoLineas = 2 Then
                    If ImporteFormateado(txtaux(Index).Text) = CCur(DBLet(AdoAux(1).Recordset!Imponible, "N")) Then CalcularElIva = False
                    
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
            
             If Modo2 = 5 And ModoLineas = 2 Then
                If txtaux(7).Text <> "" Then
                    If Val(txtaux(Index).Text) = Val(DBLet(AdoAux(1).Recordset!TipoIva, "N")) Then CalcularElIva = False
                End If
            End If
            
            
            
        Case 10, 11
           'LOS IMPORTES
            If PonerFormatoDecimal(txtaux(Index), 1) Then
                If Not vParam.autocoste Then cmdAceptar.SetFocus
            End If
                
        Case 12
'            If txtAux(Index).Text = "" Then Exit Sub
            
            txtaux(12).Text = UCase(txtaux(12).Text)
            SQL = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtaux(12).Text, "T")
            txtAux2(12).Text = ""
            If SQL = "" Then
                MsgBox "Concepto NO encontrado: " & txtaux(12).Text, vbExclamation
                txtaux(12).Text = ""
                PonFoco txtaux(12)
                Exit Sub
            Else
                txtAux2(12).Text = SQL
            End If
            
            cmdAceptar.SetFocus
        End Select


        If CalcularElIva Then
            If Index = 5 Or Index = 6 Or Index = 7 Then CalcularIVA
        End If


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
    
    txtaux(9).Enabled = Not bDebe
    txtaux(10).Enabled = Not bHaber
    
    If bDebe Then
        txtaux(9).BackColor = &H80000018
    Else
        txtaux(9).BackColor = &H80000005
    End If
    If bHaber Then
        txtaux(10).BackColor = &H80000018
    Else
        txtaux(10).BackColor = &H80000005
    End If
End Sub




Private Sub LlamaContraPar()
    Set frmC2 = New frmColCtas
    frmC2.DatosADevolverBusqueda = "0|1|"
    frmC2.ConfigurarBalances = 3
    frmC2.FILTRO = IIf(Cliente, 6, 5)
    frmC2.Show vbModal
    Set frmC2 = Nothing
    
End Sub



Private Sub CalcularIVA()
Dim J As Integer
Dim Base As Currency
Dim Aux As Currency

    Base = ImporteFormateado(txtaux(6).Text)
    
    'EL iva
    Aux = ImporteFormateado(txtaux(8).Text) / 100
    If Aux = 0 Then
        If txtaux(10).Text = "" Then
            txtaux(10).Text = ""
        Else
            txtaux(10).Text = "0,00"
        End If
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



Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
    
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
            
            
    End Select


End Sub





Private Sub BotonAnyadirLinea(Index As Integer, Limpia As Boolean)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer

    ModoLineas = 1 'Posem Modo Afegir Llínia


    Modo2 = 5

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vTabla = "tmpfaclin"
    End Select
    ' ********************************************************

    vWhere = "" 'ObtenerWhereCab(False)

    Select Case Index
         Case 1   'hlinapu
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = ""
            NumF = SugerirCodigoSiguienteStr(vTabla, "numfac", "codusu = " & vUsu.Codigo)
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
                    
                    'Para que no de errrores
                    txtaux(0).Text = "2"
                    txtaux(1).Text = "2"
                    txtaux(2).Text = Fecha
                    txtaux(3).Text = "2"
                    txtaux(4).Text = NumF
                    
                    
                    If Limpia Then
                        txtAux2(5).Text = ""
                        txtAux2(12).Text = ""
                    End If
                    
                    
                    
                    
                   
                    
                    If Limpia Then
                        PonFoco txtaux(5)
                    Else
                        PonFoco txtaux(5)
                    End If
            
                    ' traemos la cuenta de contrapartida habitual
                    PonFoco txtaux(5)

                    
                    If txtaux(5).Text <> "" Then
                        If EstaLaCuentaBloqueada(txtaux(5).Text, CDate(txtaux(5).Text)) Then
                            txtaux(5).Text = ""
                        Else
                            
                                txtAux_LostFocus (5)
                                PonFoco txtaux(7)
                                txtAux_LostFocus (7)
                                PonFoco txtaux(6)
                            
                        End If
                        
                    End If
            
            End Select

    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub


    ModoLineas = 2 'Modificar llínia


    Modo2 = 5

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
            txtaux(9).Text = DataGridAux(Index).Columns(10).Text '%recargo
            txtaux(10).Text = DataGridAux(Index).Columns(11).Text 'importe iva
            txtaux(11).Text = DataGridAux(Index).Columns(12).Text 'importe recargo eq
            
            txtaux(12).Text = DataGridAux(Index).Columns(13).Text 'centro de coste
            txtAux2(12).Text = DataGridAux(Index).Columns(14).Text 'nombre centro de coste
            
            If DBLet(Me.AdoAux(1).Recordset!ret, "T") = "*" Then
                chkAux(0).Value = 1 ' DataGridAux(Index).Columns(14).Text 'aplica retencion
            Else
                chkAux(0).Value = 0
            End If
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    HabilitarCentroCoste
    PonFoco txtaux(5)
    
    ' ***************************************************************************************
End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim SqlLog As String

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    



    If AdoAux(Index).Recordset.EOF Then Exit Sub


    Eliminar = False
   
    vWhere = "" ' ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'linea de asiento
            SQL = "¿Seguro que desea eliminar la línea de la factura?"
            SQL = SQL & vbCrLf & AdoAux(Index).Recordset!NumFac & ": " & AdoAux(Index).Recordset!Nommacta & "     Base " & Format(AdoAux(Index).Recordset!Imponible, FormatoImporte) & "€"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                VariableCambios True
                Eliminar = True
                'Delete
                SQL = "DELETE from tmpfaclin where codusu = " & vUsu.Codigo & " AND NumFac =" & AdoAux(Index).Recordset!NumFac
              
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute SQL
        espera 0.5
        RecalcularTotales
        
        '**** parte de contabilizacion de la factura
        TerminaBloquear
        


        
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If

    End If
    
    ModoLineas = 0
    
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub VariableCambios(Lineas As Boolean)
    '  HaCambiadoLineas2  0. no cambio   1-Lineas iva   2. datos cabecera
    If HaCambiadoLineas = 0 Then
        If Lineas Then
            HaCambiadoLineas = 1
        Else
            HaCambiadoLineas = 3
        End If
    Else
        If Lineas Then
            If HaCambiadoLineas = 2 Then
                HaCambiadoLineas = 3
            Else
                HaCambiadoLineas = 1
            End If
        Else
            If HaCambiadoLineas = 1 Then
                HaCambiadoLineas = 3
            Else
                HaCambiadoLineas = 2
            End If
        End If
    End If
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    If B Then
        cmdCancelar.Caption = "Cancelar"
        cmdAceptar.Caption = "&Aceptar"
    Else
        cmdCancelar.Caption = "Salir"
        cmdAceptar.Caption = "&Modificar"
    End If
    cmdCancelar.Cancel = True
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
            
            'Porcentaje de IVA
            BloqueaTXT txtaux(8), True
            BloqueaTXT txtaux(9), True
            
            'Los campos importes de IVA siempre bloqueados excepto que el parametro indique lo contrario
            If B Then
                If Cliente Then
                    If Not vParam.ModificarIvaLineasFraCli Then B = True
                Else
                    
                End If
                
            Else
                
                B = False
            End If
            B = Not B
            BloqueaTXT txtaux(10), B
            BloqueaTXT txtaux(11), B
        
    End Select
    
    
        
    
End Sub

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



Private Sub RecalcularTotales()

End Sub



Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = "select numserie ,codigo,'" & Format(Fecha, FormatoFecha) & "' fecfactu," & Anyo & " anofac,Numfac,cta,nommacta,imponible,tipoiva,IVA,porcrec,impiva,recargo ,"
    tots = tots & "  codccost,nomccost,if(tipoopera=1,'*','') Ret "
    tots = tots & " from tmpfaclin left join  ccoste ON codusu =" & vUsu.Codigo & "  AND numfactura = ccoste.codccost"
    tots = tots & " inner join cuentas on codusu =" & vUsu.Codigo & " and cta=codmacta"
    tots = tots & " WHERE codusu = " & vUsu.Codigo & " ORDER BY numfac"

    B = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = Conn
    AdoAux(Index).RecordSource = tots
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 350
    
    'If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    'End If

    For i = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(i).AllowSizing = False
    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, True
    
    
    Select Case Index
         
        Case 1 'lineas de asiento
            
            If vParam.autocoste Then
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(5)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(5)|T|Denominación|3995|;"
                tots = tots & "S|txtaux(6)|T|Importe|1905|;S|txtaux(7)|T|Iva|625|;S|cmdAux(1)|B|||;S|txtaux(8)|T|%Iva|765|;"
                tots = tots & "S|txtaux(9)|T|%Rec|765|;S|txtaux(10)|T|Importe Iva|1554|;S|txtaux(11)|T|Importe Rec|1554|;"
                tots = tots & "S|txtaux(12)|T|CC|710|;S|cmdAux(2)|B|||;S|txtAux2(12)|T|Nombre|2470|;S|chkAux(0)|T|Ret|470|;"
            Else
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(5)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(5)|T|Denominación|5695|;"
                tots = tots & "S|txtaux(6)|T|Importe|2405|;S|txtaux(7)|T|Iva|625|;S|cmdAux(1)|B|||;S|txtaux(8)|T|%Iva|855|;"
                tots = tots & "S|txtaux(9)|T|%Rec|855|;S|txtaux(10)|T|Importe Iva|1954|;S|txtaux(11)|T|Importe Rec|1954|;"
                tots = tots & "N||||0|;N||||0|;S|chkAux(0)|CB|Ret|400|;"
            End If
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgLeft
            DataGridAux(Index).Columns(6).Alignment = dbgLeft
            DataGridAux(Index).Columns(14).Alignment = dbgCenter
            
            B = (Modo2 = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

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
      
  
    ' **********************************************************
      
    'Obtenemos las sumas
'    ObtenerSumas
    If Enlaza Then CargaDatosLW2 0
    

      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub



' 0. Origen,   2 Modificado
Private Sub CargaDatosLW2(QueLW As Integer)
Dim cad As String
Dim Rs As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String
Dim C As String


    On Error GoTo ECargaDatosLW
    
    
    If QueLW = 1 Then
        cad = "select h.numlinea,  h.codigiva, h.baseimpo, h.impoiva, h.imporec from "
        cad = cad & IIf(Cliente, "factcli_totales", "factpro_totales") & " h WHERE "
        cad = cad & " numserie=" & DBSet(NUmSerie, "T") & " and "
        cad = cad & IIf(Cliente, "numfactu", "numregis") & " = "
        cad = cad & Codigo & " and anofactu=" & Anyo
        cad = cad & " ORDER BY 1"
        GroupBy = ""
  
    Else
        
        
        cad = "select tipoiva codigiva, IVA, porcrec, sum(Imponible) baseimpo ,sum(coalesce(ImpIVA,0)) impoiva,sum(coalesce(recargo,0)) imporec"
        cad = cad & " from tmpfaclin where codusu = " & vUsu.Codigo
        cad = cad & " group by 1"
        cad = cad & " order by 1"
    
        
    End If
    
    lw1(QueLW).ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = "0"
    
    While Not Rs.EOF
        Set It = lw1(QueLW).ListItems.Add
        cad = Val(cad) + 1
        It.Text = cad
        It.SubItems(1) = Format(Rs!codigiva, "000")
        It.SubItems(2) = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", Rs!codigiva)
        It.SubItems(3) = Format(Rs!Baseimpo, "###,###,##0.00")
        It.SubItems(4) = Format(Rs!Impoiva, "###,###,##0.00")
        If DBLet(Rs!ImpoRec) <> 0 Then
            It.SubItems(5) = Format(Rs!ImpoRec, "###,###,##0.00")
        Else
            It.SubItems(5) = " "
        End If
        
        Set It = Nothing

        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
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




Private Sub frmC2_DatoSeleccionado(CadenaSeleccion As String)
Dim vFe As String


    If cmdAux(0).Tag = 0 Then
        'Cuenta normal
        txtaux(5).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
        
        'Habilitaremos el ccoste
        HabilitarCentroCoste
        
    ElseIf cmdAux(0).Tag = 2 Then
        'cuenta retencion
        SQL = CadenaSeleccion
    Else
        'contrapartida
        txtaux(6).Text = RecuperaValor(CadenaSeleccion, 1)
    End If

End Sub

Private Sub frmCC_DatoSeleccionado(CadenaSeleccion As String)
    'Centro de coste
    txtaux(12).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux2(12).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub CargarColumnas(QueLW As Byte)
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

    lw1(QueLW).Tag = "5|" & Ncol & "|"
    
    lw1(QueLW).ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1(QueLW).ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim


End Sub





Private Function InsertarLinea() As Boolean
    InsertarLinea = False
    If DatosOkLlin("FrameAux1") Then
        SQL = "INSERT INTO tmpfaclin(codusu,numfac,numserie,codigo,Fecha,cta,Imponible,IVA,ImpIVA,tipoiva,porcrec,recargo,tipoopera,numfactura) VALUES ("
        '                                  numlinea                serie                         fac/reg             frec,fecfac
        SQL = SQL & vUsu.Codigo & "," & txtaux(4).Text & "," & DBSet(NUmSerie, "T") & "," & Codigo & "," & DBSet(Fecha, "F") & ","
        '                   cta                             impo                              iva
        SQL = SQL & DBSet(txtaux(5).Text, "T") & "," & DBSet(txtaux(6).Text, "N") & "," & DBSet(txtaux(8).Text, "N") & ","
        '                   importeiva                   tipoiva                            porrec
        SQL = SQL & DBSet(txtaux(10).Text, "N") & "," & DBSet(txtaux(7).Text, "N") & "," & DBSet(txtaux(9).Text, "T", "S") & ","
        '               recargo                               tiene ret                     codccost
        SQL = SQL & DBSet(txtaux(11).Text, "N", "S") & "," & Abs(Me.chkAux(0).Value) & "," & DBSet(txtaux(12).Text, "T", "S") & ")"
        
        
        If Ejecuta(SQL) Then InsertarLinea = True
        
        
    End If
End Function


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Importe As Currency

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And (Modo2 = 5 And ModoLineas = 1) Then  'insertar
    
    End If
    
    If B And Modo2 = 5 Then ' tanto si insertamos como si modificamos en lineas
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
        
        If txtaux(5).Text = "NO" Then
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
        
        
    End If
    
    
    
    
    
    'Como puede modificar los IVA, hay que comprobar
    If B And vParam.ModificarIvaLineasFraCli Then
        
        Importe = ImporteFormateado(txtaux(8).Text) / 100
        Importe = ImporteFormateado(txtaux(6).Text) * Importe
        
        
        
        If Abs(Importe - ImporteFormateado(txtaux(10).Text)) >= 0.1 Then
            Mens = "Iva calculado: " & Format(Importe, FormatoImporte) & vbCrLf
            Mens = Mens & "Iva introducido: " & txtaux(10).Text & vbCrLf
            Mens = "DIFERENCIAS EN IVA" & vbCrLf & vbCrLf & Mens & vbCrLf & "¿Desea continuar igualmente?"
            
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
                    Mens = "DIFERENCIAS EN RECARGO EQUIVALENCIA" & vbCrLf & vbCrLf & Mens & vbCrLf & "¿Desea continuar igualmente?"
                    
                    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then B = False
                End If
                
            End If
        End If
    End If
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function




Private Function ModificarLinea() As Boolean
    ModificarLinea = False
    Screen.MousePointer = vbHourglass
    If DatosOkLlin(FrameAux1) Then
        If UpdateLinea Then ModificarLinea = True
    End If
    Screen.MousePointer = vbDefault
End Function


Private Function UpdateLinea() As Boolean
    On Error GoTo eUpdateLinea
    UpdateLinea = False
    
    SQL = "UPDATE tmpfaclin SET "
   
    '                   cta                             impo                              iva
    SQL = SQL & "cta = " & DBSet(txtaux(5).Text, "T") & ", Imponible = " & DBSet(txtaux(6).Text, "N") & ","
    SQL = SQL & "IVA = " & DBSet(txtaux(8).Text, "N") & ", ImpIVA = "
    '                   importeiva                   tipoiva                            porrec
    SQL = SQL & DBSet(txtaux(10).Text, "N") & ", tipoiva = " & DBSet(txtaux(7).Text, "N") & ", porcrec=" & DBSet(txtaux(9).Text, "T", "S") & ", recargo ="
    '               recargo                               tiene ret                     codccost
    SQL = SQL & DBSet(txtaux(11).Text, "N", "S") & ", tipoopera= " & Abs(Me.chkAux(0).Value) & ",numfactura =" & DBSet(txtaux(12).Text, "T", "S")
        
    
    
    
    
    SQL = SQL & " WHERE codusu  = " & vUsu.Codigo
    SQL = SQL & " AND numfac =" & Me.AdoAux(1).Recordset!NumFac
    
    Conn.Execute SQL
    UpdateLinea = True
    Exit Function
eUpdateLinea:
    MuestraError Err.Number, Err.Description

End Function
