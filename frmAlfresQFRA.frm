VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmAlfresQFRA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepcion facturas proveedor"
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11130
   ScaleWidth      =   17310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   40
      Top             =   7680
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Cuentas base"
      TabPicture(0)   =   "frmAlfresQFRA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAux1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos extendidos"
      TabPicture(1)   =   "frmAlfresQFRA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lw1(0)"
      Tab(1).Control(1)=   "Text1(14)"
      Tab(1).Control(2)=   "Text1(13)"
      Tab(1).Control(3)=   "Text1(12)"
      Tab(1).Control(4)=   "Combo1(2)"
      Tab(1).Control(5)=   "Text4(11)"
      Tab(1).Control(6)=   "Text1(11)"
      Tab(1).Control(7)=   "Text1(10)"
      Tab(1).Control(8)=   "Text1(15)"
      Tab(1).Control(9)=   "Text1(16)"
      Tab(1).Control(10)=   "Label6(2)"
      Tab(1).Control(11)=   "Label6(1)"
      Tab(1).Control(12)=   "Label1(12)"
      Tab(1).Control(13)=   "Label6(0)"
      Tab(1).Control(14)=   "Image1(6)"
      Tab(1).Control(15)=   "Label1(11)"
      Tab(1).Control(16)=   "Label1(10)"
      Tab(1).ControlCount=   17
      Begin MSComctlLib.ListView lw1 
         Height          =   1785
         Index           =   0
         Left            =   -67680
         TabIndex        =   62
         Top             =   840
         Width           =   9435
         _ExtentX        =   16642
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
         Left            =   -67080
         TabIndex        =   66
         Tag             =   "Porcentaje Retencion|N|S|||factpro|retfacpr|##0.00||"
         Text            =   "impovia"
         Top             =   1560
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
         Index           =   13
         Left            =   -74760
         TabIndex        =   65
         Tag             =   "Porcentaje Retencion|N|S|||factpro|retfacpr|##0.00||"
         Text            =   "basereten"
         Top             =   2280
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
         Index           =   12
         Left            =   -67080
         TabIndex        =   64
         Tag             =   "Porcentaje Retencion|N|S|||factpro|retfacpr|##0.00||"
         Text            =   "baseimpo"
         Top             =   840
         Width           =   1350
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
         ItemData        =   "frmAlfresQFRA.frx":0038
         Left            =   -74760
         List            =   "frmAlfresQFRA.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Tag             =   "Tipo retencion|N|N|||factpro|tiporeten|||"
         Top             =   750
         Width           =   4560
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
         Index           =   11
         Left            =   -73290
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text4"
         Top             =   1470
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
         Index           =   11
         Left            =   -74730
         TabIndex        =   56
         Tag             =   "Cuenta Retencion|T|S|||factpro|cuereten|||"
         Text            =   "1234567890"
         Top             =   1470
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
         Index           =   10
         Left            =   -69840
         TabIndex        =   55
         Tag             =   "Porcentaje Retencion|N|S|||factpro|retfacpr|##0.00||"
         Text            =   "1234567890"
         Top             =   750
         Width           =   1350
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   2340
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   16695
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   9
            Tag             =   "Cuenta|T|N|||factcli_lineas|codmacta|||"
            Text            =   "Cta Base"
            Top             =   2160
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Frame FrameToolAux 
            Height          =   555
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   1605
            Begin MSComctlLib.Toolbar ToolbarAux 
               Height          =   330
               Left            =   180
               TabIndex        =   47
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
            TabIndex        =   11
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
            TabIndex        =   45
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
            TabIndex        =   10
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
            TabIndex        =   44
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   17
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
            TabIndex        =   43
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   42
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
            Height          =   2520
            Index           =   1
            Left            =   0
            TabIndex        =   54
            Top             =   600
            Width           =   16770
            _ExtentX        =   29580
            _ExtentY        =   4445
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
         Left            =   -73320
         TabIndex        =   67
         Tag             =   "Porcentaje Retencion|N|S|||factpro|retfacpr|##0.00||"
         Text            =   "imporeten"
         Top             =   2280
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
         Index           =   16
         Left            =   -65280
         TabIndex        =   70
         Tag             =   "Porcentaje Retencion|N|S|||factpro|retfacpr|##0.00||"
         Text            =   "recargo"
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label Label6 
         Caption         =   "Importe ret."
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
         Left            =   -73320
         TabIndex        =   69
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Base reten."
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
         Left            =   -74760
         TabIndex        =   68
         Top             =   2040
         Width           =   1380
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
         Index           =   12
         Left            =   -67680
         TabIndex        =   63
         Top             =   480
         Width           =   1560
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   61
         Top             =   480
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   6
         Left            =   -72840
         Top             =   1200
         Width           =   240
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
         Index           =   11
         Left            =   -74730
         TabIndex        =   60
         Top             =   1200
         Width           =   1815
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
         Index           =   10
         Left            =   -69840
         TabIndex        =   59
         Top             =   480
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5295
      Left            =   11280
      TabIndex        =   30
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9340
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº Albaran"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Importe"
         Object.Width           =   2187
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      TabIndex        =   18
      Top             =   10560
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15960
      TabIndex        =   19
      Top             =   10560
      Width           =   1155
   End
   Begin VB.TextBox txtNomFich 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   480
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   120
      Width           =   7695
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      DragMode        =   1  'Automatic
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   10935
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   6000
      Width           =   17055
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   15120
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   15120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   38
         Text            =   "frmAlfresQFRA.frx":003C
         Top             =   1140
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   10920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "frmAlfresQFRA.frx":0042
         Top             =   1140
         Width           =   4095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   9720
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1140
         Width           =   1095
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
         ItemData        =   "frmAlfresQFRA.frx":0048
         Left            =   6960
         List            =   "frmAlfresQFRA.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Tipo operación|N|N|||factpro|codopera|||"
         Top             =   1140
         Width           =   2610
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
         ItemData        =   "frmAlfresQFRA.frx":004C
         Left            =   240
         List            =   "frmAlfresQFRA.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1140
         Width           =   6330
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   13560
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   11880
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "frmAlfresQFRA.frx":0050
         Top             =   420
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   9
         Left            =   15120
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   5
         Left            =   16440
         Picture         =   "frmAlfresQFRA.frx":0056
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   11280
         Top             =   840
         Width           =   240
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
         Index           =   0
         Left            =   6990
         TabIndex        =   35
         Top             =   840
         Width           =   2385
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
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   14760
         Picture         =   "frmAlfresQFRA.frx":00E1
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   0
         Left            =   13560
         TabIndex        =   31
         Top             =   120
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   13200
         Picture         =   "frmAlfresQFRA.frx":016C
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total factura"
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
         Index           =   7
         Left            =   15120
         TabIndex        =   28
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   6
         Left            =   11880
         TabIndex        =   27
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   4
         Left            =   4080
         TabIndex        =   26
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Index           =   2
         Left            =   9600
         TabIndex        =   25
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   24
         Top             =   120
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   1680
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Forma de pago"
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
         Left            =   9720
         TabIndex        =   37
         Top             =   840
         Width           =   1920
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   120
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Albaranes"
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
      Index           =   8
      Left            =   16080
      TabIndex        =   33
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Albaranes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   11280
      TabIndex        =   32
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmAlfresQFRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CarpetaAlbaranes As String
Public CarpetaDestino As String

'
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTIva As frmBasico2
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmCC As frmBasico
Attribute frmCC.VB_VarHelpID = -1
Private WithEvents frmFPag As frmBasico2
Attribute frmFPag.VB_VarHelpID = -1


Dim Modo2 As Byte
Dim ModoLineas As Byte

Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntiguoText1 As String

Dim PrimVez As Boolean



Private Sub cmdAceptar_Click()
Dim B As Boolean
    If Modo2 <> 5 Then
        If AceptarFactura Then
           
            Unload Me
        End If
    Else
        If ModoLineas = 1 Then
            B = InsertarLinea
        Else
            B = ModificarLinea
        End If
        If B Then
            Me.DataGridAux(1).AllowAddNew = False
            LLamaLineas 1, 0
                        
            
            
            PonerModo 2
            CargaGrid 1, True
            
        End If
    End If
End Sub



Private Function AceptarFactura() As Boolean
Dim B As Boolean
Dim FicheroFinal As String
Dim Aux As String

On Error GoTo eCmdAceptar_Click
    
    AceptarFactura = False
    Msg = ""
    If Me.txtNomFich.Text = "" Then Msg = "- FICHERO" & vbCrLf
    If Me.txtNomFich.Tag = "" Then Msg = "- FICHERO(2)" & vbCrLf
    For i = 0 To 9
        If Text1(i).Text = "" Then Msg = Msg & "-" & RecuperaValor("Cuenta|NIF|Nombre|NºFactura|Fecha factura|Fecha recepcion|Importe|Cod. formapago|Forma de pago|Fecha liquidacion|", i + 1) & vbCrLf
    Next i
    
    If Me.AdoAux(1).Recordset.EOF Then Msg = Msg & "-Ninguna cuenta base añadida"
    If Msg <> "" Then Msg = "Campos obligatorios" & vbCrLf & Msg & vbCrLf
    
    
    Aux = Msg
    If Text1(5).Text <> "" And Text1(9).Text <> "" Then
        Msg = ""
        If Not ComprobarPeriodo(True, 0) Then
            If Msg = "" Then Msg = "Error comprobando peridod"
        
        End If
        If Msg <> "" Then
            Aux = Aux & vbCrLf & Msg
            Msg = ""
        End If
        
        Msg = Aux  'Dejo el error como estaba
        
       If Text1(9).Text <> "" Then
            If EstaLaCuentaBloqueada(Text1(0).Text, CDate(Text1(9).Text)) Then Msg = Msg & "- Cuenta proveedor BLOQUEADA"
        End If
        
    End If
    
    Aux = ""
    If Me.Combo1(2).ListIndex > 0 Then
        'Tiene retencion. Que indique cuenta y %
        If Me.Text1(11).Text = "" Xor Text4(11).Text = "" Then Aux = "M"
        If Me.Text1(10).Text = "" Then Aux = "M"
        If Aux <> "" Then Msg = Msg & vbCrLf & vbCrLf & " Si lleva retencion indique resto de datos retencion"
    Else
        
        If Me.Text1(11).Text <> "" Or Text4(11).Text <> "" Then Aux = "M"
        If Me.Text1(10).Text <> "" Then Aux = "M"
        If Aux <> "" Then Msg = Msg & vbCrLf & vbCrLf & " NO lleva retencion . No debe indicar resto de campos retencion"
    End If
    
    If Me.Combo1(2).ListIndex = 0 And Text1(15).Text <> "" Then Msg = "Sin tipo de retencion, pero importe calculado:" & Text1(15).Text
    
    
    If Msg <> "" Then
        MsgBox Msg, vbExclamation
        Exit Function
    End If
    
    
    'Comprobaremos las lineas
    AdoAux(1).Recordset.MoveFirst
    While Not AdoAux(1).Recordset.EOF
        If EstaLaCuentaBloqueada(AdoAux(1).Recordset!Cta, CDate(Text1(9).Text)) Then Msg = Msg & "    " & AdoAux(1).Recordset!Cta & " -" & AdoAux(1).Recordset!Nommacta
        AdoAux(1).Recordset.MoveNext
    Wend
    AdoAux(1).Recordset.MoveFirst
    If Msg <> "" Then
        MsgBox "Cuentas bloqueadas: " & vbCrLf & vbCrLf & Msg, vbExclamation
        Exit Function
    End If
    
    
    
    
    NombreArchivoDestino False
    Aux = Msg
    
    Msg = ""
    
    If Dir(App.Path & "\pdftk.exe", vbArchive) = "" Then Msg = Msg & "No existe el programa concatenar PDFs " & vbCrLf
    
    If Dir(Me.txtNomFich.Tag, vbArchive) = "" Then Msg = Msg & "No existe el PDF origen: " & txtNomFich.Tag & vbCrLf
    
    If Dir(CarpetaDestino & "\" & Aux, vbArchive) <> "" Then Msg = Msg & "YA existe el PDF en el destino: " & CarpetaDestino & "\" & Aux & vbCrLf
    
    If Msg <> "" Then
        MsgBox Msg, vbExclamation
        Exit Function
    End If
        
        
    If Not ComprobarArchivosSeleccionados Then Exit Function
    
    
    'No puede exitir esa factura e       en proveedor año
    Msg = "codmacta = " & DBSet(Text1(0).Text, "T") & " AND numserie =1 AND numfactu =" & DBSet(Text1(3).Text, "T") & " AND year(fecfactu)"
    Msg = DevuelveDesdeBD("fecharec", "factpro", Msg, CStr(Year(CDate(Text1(4).Text))))
    If Msg <> "" Then
        MsgBox "Ya existe la factura en el registro de IVA", vbExclamation
        Exit Function
    End If
        
    'Ultimas comprobaciones
        
            
        
        
        
        
    If MsgBox("¿Desea insertar la factura?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    CadenaDesdeOtroForm = ""
    'Trae los albaranes seleccionados
    If ComprobarArchivosSeleccionados Then
        
        B = ConcatenaYSubeFichero
        If B Then
            Conn.BeginTrans
            B = InsertaFactura
           
            
            If B Then
                Conn.CommitTrans
                BorraAlbaranes
                AceptarFactura = True
            Else
                
                Conn.RollbackTrans
            End If
        End If
    End If
    Exit Function
eCmdAceptar_Click:
    MuestraError Err.Number, Err.Description
End Function

Private Sub cmdAux_Click(Index As Integer)
   TerminaBloquear
   Msg = ""
    Select Case Index
        Case 0 ' cuenta base
            cmdAux(0).Tag = 0
            SQL = ""
            LlamaContraPar
            If SQL <> "" Then
               txtaux(5).Text = RecuperaValor(SQL, 1)
               txtAux2(5).Text = RecuperaValor(SQL, 1)
            
                PonFoco txtaux(5)
            End If
        Case 1 'tipo de iva
            cmdAux(0).Tag = 1
            
            Set frmTIva = New frmBasico2
            AyudaTiposIva frmTIva
            Set frmTIva = Nothing
            If Msg <> "" Then
                txtaux(7).Text = Msg
                txtAux_LostFocus 7
                'PonFoco txtaux(7)
            End If
        Case 2 'cento de coste
            If txtaux(12).Enabled Then
                Set frmCC = New frmBasico
                AyudaCC frmCC
                Set frmCC = Nothing
            End If

    End Select
End Sub

Private Sub cmdCancelar_Click()

    If Modo2 <> 5 Then

    
        If Me.txtNomFich.Tag <> "" Then
            If MsgBox("Desea cancelar el proceso?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        Unload Me
    
    
    Else
        
        Me.DataGridAux(1).AllowAddNew = False
        LLamaLineas 1, 0
        PonerModo 2
        Modo2 = 2
        CargaGrid 1, True
    End If
End Sub

Private Sub Combo1_Click(Index As Integer)
        If PrimVez Then Exit Sub
    If Index > 0 Then CalculaTotales
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        If Dir(App.Path & "\pdftk.exe", vbArchive) = "" Then MsgBox "No existe el programa concatenar PDFs ", vbExclamation
    
    End If
End Sub

Private Sub Form_Load()
    'Me.AcroPDF1.DragIcon = vbNoDrop
    'Screen.MousePointer = vbHourglass
'    AcroPDF1.LoadFile "C:\Users\David\Downloads\borrame.pdf"
'    AcroPDF1.setZoom 100
    PrimVez = True
    Me.Icon = frmppal.Icon
    Image1(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image1(1).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image1(4).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Image1(6).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    With Me.ToolbarAux
         .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5

    End With
    Label1(8).Caption = ""
    Label1(8).Tag = 0
    
    BloqueaTXT Text1(9), Not vParam.IvaEnFechaPago
    Me.Image1(5).visible = vParam.IvaEnFechaPago
    Text1(9).Enabled = vParam.IvaEnFechaPago
    BloqueaTXT Text1(13), True
    BloqueaTXT Text1(15), True
   
    
    
    
    
    CargarColumnas
    Limpiar Me
    Ejecuta "DELETE FROM tmpfaclin where codusu = " & vUsu.Codigo
    CargaGrid 1, True
    CargarCombo
    PonerTxtTo "", ""
    
    
    Text1(5).Text = Format(Now, "dd/mm/yyyy")
    Text1(9).Text = Text1(5).Text
End Sub


Private Sub PonerTxtTo(NomFichero As String, Corto As String)

    If NomFichero = "" Then
        txtNomFich.Text = "Arraste y suelte aqui el archivo pdf o haga click en la lupa para buscarlo"
        txtNomFich.FontItalic = True
        txtNomFich.ForeColor = &H808080
        txtNomFich.Tag = ""
    Else
        txtNomFich.Text = Corto
        txtNomFich.FontItalic = False
        txtNomFich.ForeColor = vbBlack
        txtNomFich.Tag = NomFichero
    End If
End Sub

Private Sub Form_Resize()
'Dim H
'
'
'    If Me.WindowState = vbMinimized Then Exit Sub
'
'
'    H = Me.Width - Frame1.Width - 600  '400 es el minimo
'
'    If H < 0 Then
'        Me.Width = Frame1.Width + 600
'        H = 400
'    End If
'    Me.Frame1.Left = Me.Width - Me.Frame1.Width - 240
'    Me.AcroPDF1.Width = H
'    Me.cmdCancelar.Left = Me.Width - 420 - Me.cmdCancelar.Width
'    Me.cmdAceptar.Left = cmdCancelar.Left - cmdCancelar.Width - 240
'
'    H = Me.Height - 400 - 8000   '8000 es el minimo
'
'    If H < 0 Then Me.Height = 8000
'
'
'    Me.AcroPDF1.Height = Me.Height - 420 - 400
'    Me.cmdCancelar.top = Me.Height - 640 - Me.cmdCancelar.Height
'    cmdAceptar.top = cmdCancelar.top
'
'
'
End Sub



Private Sub frmC_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmFPag_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub AbirCD1()
        frmppal.cd1.FileName = ""
        frmppal.cd1.Filter = "*.pdf|*.pdf"
        frmppal.cd1.InitDir = "c:\" 'PathSalida
        frmppal.cd1.FilterIndex = 1
        frmppal.cd1.ShowOpen
        If frmppal.cd1.FileName <> "" Then
            
            If UCase(Right(frmppal.cd1.FileTitle, 4)) <> ".PDF" Then
                MsgBox "Solo acepta archivos '.pdf'", vbExclamation
                Exit Sub
            End If
            Msg = "  " & CStr(frmppal.cd1.FileName) & "  "
            Msg = Trim(Msg)
            SQL = "    " & CStr(frmppal.cd1.FileTitle) & "  "
            SQL = Trim(SQL)
            frmppal.cd1.FileName = ""
        End If
        Err.Clear
End Sub


Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
    
    Msg = RecuperaValor(CadenaSeleccion, 1)
    
End Sub

Private Sub Image1_Click(Index As Integer)
    
    Select Case Index
    Case 0
        
            Msg = ""
            AbirCD1
            If Msg = "" Then Exit Sub
           ' Msg = "C:\Users\David\Downloads\PagoWachina.pdf"
           ' SQL = "PagoWachina.pdf"

            CargarPDF "", ""
            Me.txtNomFich.Text = "Leyendo fichero ..."
            Me.AcroPDF1.visible = False
            Screen.MousePointer = vbHourglass
            Me.Refresh
            CargarPDF Msg, SQL
            Me.AcroPDF1.visible = True
       
        
    Case 1, 6
        SQL = ""
        Set frmCtas = New frmColCtas
        frmCtas.DatosADevolverBusqueda = "0|1|2|"
        frmCtas.ConfigurarBalances = 3  'NUEVO
        frmCtas.Show vbModal
        Set frmCtas = Nothing
        
        If SQL <> "" Then
            i = IIf(Index = 1, 0, 11)
            Text1(i).Text = RecuperaValor(SQL, 1)
            Text1_LostFocus CInt(i)
        End If
        
    Case 2, 3, 5
        
        
        
        SQL = Now
        i = IIf(Index = 5, 9, Index + 2)
        If Text1(i).Text <> "" Then
            If IsDate(Text1(i).Text) Then SQL = CDate(Text1(i).Text)
        End If
        AbrirFecha
        If SQL <> "" Then
            Text1(i).Text = SQL
            If Index = 3 And Not vParam.IvaEnFechaPago Then Text1(9).Text = SQL
                
        End If
    Case 4
        SQL = ""
        Set frmFPag = New frmBasico2
        AyudaFPago frmFPag
        Set frmFPag = Nothing
        If SQL <> "" Then
            Text1(7).Text = RecuperaValor(SQL, 1)
            Text1(8).Text = RecuperaValor(SQL, 2)
        End If
    End Select
    Me.Refresh
End Sub

Private Sub AbrirFecha()
    
    Set frmC = New frmCal
    frmC.Fecha = SQL
    SQL = ""
    frmC.Show vbModal
    Set frmC = Nothing
    
End Sub


Private Sub imgppal_Click(Index As Integer)

End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Msg = CarpetaAlbaranes & "\" & ListView1.SelectedItem.Tag
    
    If Dir(Msg, vbArchive) = "" Then
        MsgBox "No existe el fichero: " & Msg, vbExclamation
        Exit Sub
    End If
    
    Call ShellExecute(Me.hwnd, "Open", Msg, "", "", 1)
    
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim TotalImp As Currency
    TotalImp = 0
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then TotalImp = TotalImp + ImporteFormateado(ListView1.ListItems(i).SubItems(2))
    Next i
    Label1(8).Tag = TotalImp
    Label1(8).Caption = Format(TotalImp, FormatoImporte) & " €"
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub


Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Impor As Currency

  Select Case Button.Index
        Case 1
            'AÑADIR linea factura
            BotonAnyadirLinea 1, True
            
            If Me.AdoAux(1).Recordset.EOF Then
                If Text1(0).Text <> "" Then
                    'Primera
                    SQL = "cuentas.codcontrhab=cuentas2.codmacta AND cuentas.codmacta"
                    SQL = DevuelveDesdeBD("concat(cuentas2.codmacta,'|',cuentas2.nommacta,'|')", "cuentas,cuentas as cuentas2", SQL, Text1(0).Text, "T")
                    If SQL <> "||" Then
                        txtaux(5).Text = RecuperaValor(SQL, 1)
                        txtAux2(5).Text = RecuperaValor(SQL, 2)
                        PonFoco txtaux(6)
                    End If
                    
                    
                    'Si hay seleeciondos albaranes, pintamos el importe
                    If Me.Label1(8).Tag <> 0 Then
                        
                        txtaux(6).Text = Format(Label1(8).Tag, FormatoImporte)
                    End If
                End If
            Else
                'NO es la primera linea
                If Me.Label1(8).Tag <> 0 Then
                    Impor = ImporteFormateado(Text1(12).Text)
                    Impor = Label1(8).Tag - Impor
                    If Impor > 0 Then txtaux(6).Text = Format(Impor, FormatoImporte)
                End If
            End If
            
        Case 2
            'MODIFICAR linea factura
            BotonModificarLinea 1
        Case 3
            'ELIMINAR linea factura
            BotonEliminarLinea 1
            CalculaTotales
            
    End Select


End Sub

Private Sub txtNomFich_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cad As String
Dim Nombre As String
    On Error GoTo eT

    
    
    cad = ""
    If Data.Files Is Nothing Then
        cad = "N"
    Else
        If Data.Files.Count <> 1 Then cad = "N"
    End If
    If cad <> "" Then
        MsgBox "Solo se puede cargar un fichero", vbExclamation
        Exit Sub
    End If
    
    cad = UCase(Right(Data.Files.Item(1), 4))
    If cad <> ".PDF" Then
        MsgBox "Solo acepta archivos '.pdf'", vbExclamation
        Exit Sub
    End If
    
    cad = Data.Files.Item(1)
    If Dir(cad, vbArchive) = "" Then
        MsgBox "No es una ruta válidad", vbExclamation
        Exit Sub
    End If
    
    NumRegElim = InStrRev(cad, "\")
    If NumRegElim = 0 Then
        MsgBox "Imposible encontrar path archivo", vbExclamation
        Exit Sub
    End If
    
    Nombre = Mid(cad, NumRegElim + 1)
    
    
    CargarPDF cad, Nombre
        
    Exit Sub
eT:
        MsgBox Err.Description, vbExclamation
        Err.Clear
End Sub



Private Sub CargarPDF(Archivo As String, Nombre As String)

On Error GoTo EC
    Screen.MousePointer = vbHourglass
    AcroPDF1.LoadFile Archivo
    AcroPDF1.setZoom 70
    PonerTxtTo Archivo, Nombre
    
    
EC:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        PonerTxtTo "", ""
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text1_GotFocus(Index As Integer)
  
    ConseguirFoco Text1(Index), 3
  
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim Limpi As Boolean
Dim Id As String



    If Not PerderFocoGnral(Text1(Index), 3) Then Exit Sub
    
   

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
   
    
    Select Case Index
        
        Case 0, 1
                
                Limpi = True
                If Text1(Index).Text <> "" Then
                    Id = ""
                    If Index = 0 Then
                        If Not IsNumeric(Text1(Index).Text) Then
                            MsgBox "Cuenta debe ser numerica", vbExclamation
                        Else
                            Id = RellenaCodigoCuenta(Text1(Index).Text)
                        End If
                    Else
                        Id = Text1(Index).Text
                        If Index = 1 And Text1(0).Text <> "" Then Exit Sub
                    End If
                    If Id <> "" Then
                        'Proveedor distinto
                        
                        Limpi = False
                        PonerNombreCuentaNIF Index = 0, Id
                                                
                        'If Text1(2).Text <> "" Then
                        PonFoco Text1(3)
                        
                        If Text1(0).Text <> Text1(0).Tag Then
                                Me.ListView1.ListItems.Clear
                                Label1(8).Tag = 0
                                Label1(8).Caption = ""
                                Set miRsAux = New ADODB.Recordset
                                Msg = "Select * from factproalbaranes where codmacta =" & DBSet(Text1(0).Text, "T") & " ORDER BY fechaalb,numalbar"
                                miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                i = 0
                                
                                While Not miRsAux.EOF
                                    i = i + 1
                                    ListView1.ListItems.Add , "C" & miRsAux!Id
                                    ListView1.ListItems(i).Text = miRsAux!numalbar
                                    ListView1.ListItems(i).SubItems(1) = Format(miRsAux!fechaalb, "dd/mm/yyyy")
                                    ListView1.ListItems(i).SubItems(2) = Format(miRsAux!BIMponible, FormatoImporte)
                                    ListView1.ListItems(i).Tag = miRsAux!Nombre
                        
                                    miRsAux.MoveNext
                                Wend
                                miRsAux.Close
                                Set miRsAux = Nothing

                        End If
                        
                    End If
                End If
                If Limpi Then
                    Text1(0).Text = ""
                    Text1(1).Text = ""
                    Text1(2).Text = ""
                    ListView1.ListItems.Clear
                    Label1(8).Tag = 0
                    Label1(8).Caption = ""
                End If
        
        
        Case 3
            '*?"|
            Limpi = False
            For i = 1 To 4
                If InStr(1, Text1(3).Text, Mid("*?""<>|", i, 1)) > 0 Then Limpi = True
            Next i
            If Limpi Then
                MsgBox "El numero de factura no pude contaener los siguientes caracteres:  *?""<>| ", vbExclamation
                Text1(3).Text = ""
            End If
        Case 11
            
          
            
            
            
            If Text1(Index).Text = "" Then
                Text4(Index).Text = ""
            Else
                Limpi = False
                Id = Text1(Index).Text
                If CuentaCorrectaUltimoNivel(Id, SQL) Then
                    Text1(Index).Text = Id
                    Text4(Index).Text = SQL
                    
                    If Text1(9).Text <> "" Then
                        If EstaLaCuentaBloqueada(Id, CDate(Text1(9).Text)) Then
                            MsgBoxA "Cuenta bloqueada: " & Id, vbExclamation
                            Limpi = True
                            PonFoco Text1(Index)
                        End If
                    End If
                Else
                    Limpi = True
                End If
                If Limpi Then
                    Text1(Index).Text = ""
                    Text4(Index).Text = ""
                    PonFoco Text1(Index)
                End If
            End If
        
                
        Case 6 '
        
            If Not PonerFormatoDecimal(Text1(Index), 1) Then Text1(Index).Text = ""
        Case 10
            If Not PonerFormatoDecimal(Text1(Index), 4) Then
                Text1(Index).Text = ""
            End If
            CalculaTotales
        Case 4, 5, 9 ' fecha de factura
            If Text1(Index).Text <> "" Then
                If Not EsFechaOK(Text1(Index)) Then
                    MsgBox "Fecha incorrecta", vbExclamation
                    Text1(Index).Text = ""
                    PonFoco Text1(Index)
                    Exit Sub
                End If
                
                If Text1(Index).Text <> "" And Index <> 4 Then
                    'Fecha dentro de ejercicios
                    ComprobarPeriodo False, Index
                    
                    If Index = 9 And Text1(0).Text <> "" Then
                        If EstaLaCuentaBloqueada(Text1(0).Text, CDate(Text1(9).Text)) Then
                            MsgBox "Cuenta bloqueada", vbExclamation
                            Text1(9).Text = ""
                        End If
                    End If
                End If
                
            End If
            If Index = 5 And Not vParam.IvaEnFechaPago Then
                Text1(9).Text = Text1(5).Text
                PonleFoco Me.Combo1(0)
            End If
        Case 7
            Id = ""
            Msg = ""
            If Text1(Index).Text <> "" Then
                If Not IsNumeric(Text1(Index).Text) Then
                    Msg = "forma de pago debe ser numerico"
                Else
                    Id = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", Text1(Index).Text)
                    If Id = "" Then Msg = "No existe la forma de pago"
                End If
                If Id = "" Then
                    MsgBox Msg, vbExclamation
                    Text1(Index).Text = ""
                    PonFoco Text1(Index)
                End If
            End If
            Text1(8).Text = Id
    End Select
End Sub

Private Sub PonerNombreCuentaNIF(Cuenta As Boolean, Id As String)

    Set Rs = New ADODB.Recordset
    
    Msg = "Select codmacta,nommacta,nifdatos,codforpa,nomforpa from cuentas left join formapago on cuentas.forpa=formapago.codforpa"
    Msg = Msg & " WHERE " & IIf(Cuenta, "codmacta", "nifdatos") & " = " & DBSet(Id, "T")
    Rs.Open Msg, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Rs.EOF Then
        Msg = ""
        K = -1
    Else
        
        K = 0
        While Not Rs.EOF
            K = K + 1
            Rs.MoveNext
        Wend
        Rs.MoveFirst
        If K > 1 Then
            MsgBox "Mas de una cuenta contable  para este dato: " & Id, vbExclamation
            Msg = ""
        Else
            
            Text1(0).Text = Rs!codmacta
            Text1(1).Text = DBLet(Rs!nifdatos, "T")
            Text1(2).Text = DBLet(Rs!Nommacta, "T")
            Text1(7).Text = DBLet(Rs!Codforpa, "T")
            Text1(8).Text = DBLet(Rs!nomforpa, "T")
        End If
    End If
    If Msg = "" Then
        If K < 0 Then MsgBox "No existe ninguna cuenta vinculada al " & IIf(Cuenta, "codigo", "NIF") & " " & Id, vbExclamation
        Text1(0).Text = ""
        Text1(1).Text = ""
        Text1(7).Text = ""
        Text1(8).Text = ""
    End If
    'POnemos albaranes
    ListView1.ListItems.Clear
    Label1(8).Tag = 0
    Label1(8).Caption = ""
End Sub




Private Sub CargaGrid(Index As Integer, Enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = "select numserie ,codigo,'" & Format(Now, FormatoFecha) & "' fecfactu," & Year(Now) & " anofac,Numfac,cta,nommacta,imponible,tipoiva,IVA,porcrec,impiva,recargo ,"
    tots = tots & " if(tipoopera=1,'*','') Ret, numfactura,codccost,nomccost "
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
                tots = tots & "N||||0|;S|txtaux(12)|T|CC|710|;S|cmdAux(2)|B|||;S|txtAux2(12)|T|Nombre|2470|;"
            Else
                tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;S|txtaux(5)|T|Cuenta|1405|;S|cmdAux(0)|B|||;S|txtAux2(5)|T|Denominación|5695|;"
                tots = tots & "S|txtaux(6)|T|Importe|2405|;S|txtaux(7)|T|Iva|625|;S|cmdAux(1)|B|||;S|txtaux(8)|T|%Iva|855|;"
                tots = tots & "S|txtaux(9)|T|%Rec|855|;S|txtaux(10)|T|Importe Iva|1954|;S|txtaux(11)|T|Importe Rec|1954|;"
                tots = tots & "S|chkAux(0)|CB|Ret|400|;N||||0|;N||||0|;N||||0|;"
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
      
  
    If Not PrimVez Then CalculaTotales

      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Sub PonerModo(vModo As Integer)
    Modo2 = vModo
    
    Me.Frame1.Enabled = vModo <> 5
    SQL = "&Aceptar"
    If Modo2 = 5 Then
        If ModoLineas = 1 Then
            SQL = "&Aceptar"
        Else
            SQL = "&Modificar"
        End If
    Else
        SQL = "&Generar"
    End If
    Me.cmdAceptar.Caption = SQL
End Sub


Private Sub BotonAnyadirLinea(Index As Integer, Limpia As Boolean)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer

    ModoLineas = 1 'Posem Modo Afegir Llínia

    PonerModo 5
    

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
                    
                    
                    txtaux(0).Text = "1"
                    txtaux(1).Text = "1"
                    txtaux(2).Text = Now
                    txtaux(3).Text = "1"
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
            If DataGridAux(Index).Columns(13).Text = "*" Then
                chkAux(0).Value = 1 ' DataGridAux(Index).Columns(14).Text 'aplica retencion
            Else
                chkAux(0).Value = 0
            End If
            txtaux(12).Text = DataGridAux(Index).Columns(15).Text 'centro de coste
            txtAux2(12).Text = DataGridAux(Index).Columns(16).Text 'nombre centro de coste
            
            
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    
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
        'RecalcularTotales
        
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
            
                
            B = False
        
            B = Not B
            BloqueaTXT txtaux(10), B
            BloqueaTXT txtaux(11), B
        
    End Select
    
    
        
    
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


Private Sub TxtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    
                If EstaLaCuentaBloqueada(RC, Now) Then
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
            Msg = ""
            If txtaux(Index).Text <> "" Then
                If EsNumerico(txtaux(Index).Text) Then
                    Msg = "O"
                Else
                    txtaux(Index).Text = ""
                    PonFoco txtaux(7)
                End If
            End If
            If Msg <> "" Then
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











Private Function InsertarLinea() As Boolean
    InsertarLinea = False
    If DatosOkLlin("FrameAux1") Then
        SQL = "INSERT INTO tmpfaclin(codusu,numfac,numserie,codigo,Fecha,cta,Imponible,IVA,ImpIVA,tipoiva,porcrec,recargo,tipoopera,numfactura) VALUES ("
        '                                  numlinea                serie                         fac/reg             frec,fecfac
        SQL = SQL & vUsu.Codigo & "," & txtaux(4).Text & "," & DBSet(1, "T") & ",1," & DBSet(Now, "F") & ","
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
        
        If Text1(9).Text <> "" Then
            If EstaLaCuentaBloqueada(txtaux(5).Text, CDate(Text1(9).Text)) Then
                MsgBox "- Cuenta  BLOQUEADA", vbExclamation
                PonFoco txtaux(5)
                Exit Function
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




Private Sub CargarCombo()
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
    Combo1(0).ListIndex = 0
    
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
    Combo1(1).ListIndex = 0
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
    Combo1(2).ListIndex = 0
    
    
'
'    'Tipo de intracomunitaria
'    Set Rs = New ADODB.Recordset
'    SQL = "SELECT * FROM usuarios.wtipointra ORDER BY codintra"
'    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    i = 0
'    While Not Rs.EOF
'        Combo1(3).AddItem Rs!nomintra
'        Combo1(3).ItemData(Combo1(3).NewIndex) = Asc(Rs!Codintra)
'        i = i + 1
'        Rs.MoveNext
'    Wend
'    Rs.Close
    
    
    
'
'    SQL = "SELECT * FROM usuarios.wtipoinmueble ORDER BY codigo"
'    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    i = 0
'    While Not Rs.EOF
'        Combo1(4).AddItem Rs!Descripcion
'        Combo1(4).ItemData(Combo1(4).NewIndex) = Asc(Rs!Codigo)
'        i = i + 1
'        Rs.MoveNext
'    Wend
'    Rs.Close
'
    
    
    
    Set Rs = Nothing

End Sub

Private Function InsertaFactura() As Boolean
Dim SQL As String
Dim SqlInsert As String
Dim SqlValues As String
Dim i As Long

Dim Mc As Contadores

    On Error GoTo eInsertaFactura
    InsertaFactura = False
    
    
    Set Mc = New Contadores
    
    
    If Not Mc.ConseguirContador("1", CDate(Text1(5).Text) <= vParam.fechafin, True) = 0 Then Err.Raise 513, "Error consiguiendo contador facturas"
    
    
    
    'Cebecera de factura
    SqlInsert = "INSERT INTO factpro (numserie,numregis,fecharec,numfactu,fecfactu,codconce340,codopera,codmacta,anofactu,codforpa,observa,totbases,totbasesret,totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten,fecliqpr,nommacta,dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,estraspasada)"
    
    'numserie,numregis,fecharec,numfactu,fecfactu,
    SqlValues = Mc.TipoContador & "," & Mc.Contador & "," & DBSet(Text1(5).Text, "F") & "," & DBSet(Text1(3).Text, "T") & "," & DBSet(Text1(4).Text, "F") & ","
    
    'codconce340,codopera,codmacta,anofactu,
    SqlValues = SqlValues & Me.Combo1(0).ItemData(Combo1(0).ListIndex) & "," & Me.Combo1(1).ItemData(Combo1(1).ListIndex) & ","
    SqlValues = SqlValues & DBSet(Text1(0).Text, "T") & "," & Year(CDate(Text1(5).Text)) & ","
    
    'codforpa,observa
    SQL = "Recepcion facturas Ariconta. " & vUsu.Login & " " & Now
    SqlValues = SqlValues & DBSet(Text1(7).Text, "T") & "," & DBSet(SQL, "T") & ","
    
    'totbases,totbasesret,totivas,totrecargo,totfacpr,
    SqlValues = SqlValues & DBSet(Text1(12).Text, "N") & "," & DBSet(Text1(13).Text, "N") & "," & DBSet(Text1(14).Text, "N") & ","
    SqlValues = SqlValues & DBSet(Text1(16).Text, "N") & "," & DBSet(Text1(6).Text, "N") & ","
    
    
    'retfacpr,trefacpr,cuereten,tiporeten,
    If Me.Combo1(2).ListIndex > 0 Then
        'Tiene retencion. Que indique cuenta y %
        'If Me.Text1(11).Text = "" Xor Text4(11).Text = "" Then Aux = "M"
        'If Me.Text1(10).Text = "" Then Aux = "M"
        SqlValues = SqlValues & DBSet(Text1(10).Text, "N") & "," & DBSet(Text1(15).Text, "N", "S") & ","
        SqlValues = SqlValues & DBSet(Text1(11).Text, "T") & "," & Combo1(2).ItemData(Combo1(2).ListIndex)
    Else
        SqlValues = SqlValues & "null,null,null,0"
    End If
    'fecliqpr,nommacta,dirdatos,codpobla,despobla,desprovi,nifdatos,codpais,estraspasada)"
    SqlValues = SqlValues & "," & DBSet(Text1(9).Text, "F")
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select nommacta,dirdatos,codposta,despobla,desprovi,nifdatos,codpais from cuentas where codmacta =" & DBSet(Text1(0).Text, "T"), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    For i = 0 To 6
        SqlValues = SqlValues & "," & DBSet(miRsAux.Fields(i), "T", "N")
    Next i
    SqlValues = SqlValues & ",0)"
    miRsAux.Close
    
    SqlValues = SqlInsert & " VALUES (" & SqlValues
    Conn.Execute SqlValues
    
    'TOTALEES
    
    SqlInsert = "insert into factpro_totales (numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec) values "
    
    
    SQL = "select tipoiva,  sum(imponible) baseimpo, sum(coalesce(impiva,0)) imporiva, sum(coalesce(recargo,0)) imporrec, sum(if(tipoopera=1,imponible,0))  ParaRet"
    SQL = SQL & " ,IVA,porcrec"
    SQL = SQL & " from tmpfaclin  where codusu =" & vUsu.Codigo & " group by 1 order by 1"
    
    
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    SqlValues = ""
    While Not miRsAux.EOF
        i = i + 1
        'numserie,numregis,fecharec,anofactu,numlinea,
        SqlValues = SqlValues & ", (" & Mc.TipoContador & "," & Mc.Contador & "," & DBSet(Text1(5).Text, "F") & "," & Year(CDate(Text1(5).Text)) & "," & i
        
        'baseimpo,codigiva,porciva,porcrec,impoiva,imporec
        SqlValues = SqlValues & "," & DBSet(miRsAux!Baseimpo, "N", "N") & "," & DBSet(miRsAux!TipoIva, "N", "N")
        SqlValues = SqlValues & "," & DBSet(miRsAux!IVA, "N", "N") & "," & DBSet(miRsAux!porcrec, "N", "S")
        SqlValues = SqlValues & "," & DBSet(miRsAux!Imporiva, "N", "N") & "," & DBSet(miRsAux!imporrec, "N", "S") & ")"
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    SqlValues = Mid(SqlValues, 2) 'fuera primera coma
    SqlValues = SqlInsert & SqlValues
    Conn.Execute SqlValues
    
    
    
    'LINEAS
    'insertamos  dedesde tmpfaclin
    SqlInsert = "INSERT INTO factpro_lineas(numserie,numregis,fecharec,anofactu,numlinea,codmacta,baseimpo,codigiva,porciva,porcrec,impoiva,imporec,aplicret,codccost) VALUES "
   
    SQL = "select * from tmpfaclin  where codusu =" & vUsu.Codigo & " ORDER by numfac "

    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    SqlValues = ""
    While Not miRsAux.EOF
        i = i + 1
        'numserie,numregis,fecharec,anofactu,numlinea,
        SqlValues = SqlValues & ", (" & Mc.TipoContador & "," & Mc.Contador & "," & DBSet(Text1(5).Text, "F") & "," & Year(CDate(Text1(5).Text)) & "," & i
        
        'codmacta,baseimpo,codigiva,porciva
        SqlValues = SqlValues & "," & DBSet(miRsAux!Cta, "T", "N") & "," & DBSet(miRsAux!Imponible, "N", "N")
        SqlValues = SqlValues & "," & DBSet(miRsAux!TipoIva, "N", "S") & "," & DBSet(miRsAux!IVA, "N", "N")
        ',porcrec,impoiva,imporec,aplicret,codccost
        SqlValues = SqlValues & "," & DBSet(miRsAux!porcrec, "N", "S") & "," & DBSet(miRsAux!ImpIva, "N", "N")
        SqlValues = SqlValues & "," & DBSet(miRsAux!recargo, "N", "N") & "," & DBSet(miRsAux!tipoopera, "N", "N")
        SqlValues = SqlValues & ",NULL)"
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    SqlValues = Mid(SqlValues, 2) 'fuera primera coma
    SqlValues = SqlInsert & SqlValues
    Conn.Execute SqlValues
    
    
    
     
    SQL = DevuelveDesdeBD("max(codigo)", "factpro_fichdocs", "1", "1")
    NumRegElim = Val(SQL) + 1
    'Insertamos el documento que acabamos de concatenar
    'factpro_fichdocs(codigo,numserie,numregis,anofactu,numfactu,orden,fechacrea,usucrea,docum,campo)
    
    
    NombreArchivoDestino False
    
    SQL = Msg
    'codigo,numserie,numregis,anofactu
    SqlValues = NumRegElim & "," & Mc.TipoContador & "," & Mc.Contador & "," & Year(CDate((Text1(5).Text))) & ","
    'numfactu,orden,fechacrea,usucrea,docum,campo
    SqlValues = SqlValues & DBSet(Text1(3).Text, "T") & ",1," & DBSet(Now, "FH") & "," & vUsu.Id & "," & DBSet(SQL, "T") & ")"
    
    SqlValues = "INSERT INTO factpro_fichdocs(codigo,numserie,numregis,anofactu,numfactu,orden,fechacrea,usucrea,docum) values (" & SqlValues
    Conn.Execute SqlValues
    
    SqlInsert = CarpetaDestino & "\" & SQL
    
    
    espera 0.2
    
    'Abro parar guardar el binary
    SqlValues = "Select * from factpro_fichdocs where codigo =" & NumRegElim
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = SqlValues
    adodc1.Refresh
'
    NombreArchivoDestino True

    If adodc1.Recordset.EOF Then
        'MAAAAAAAAAAAAL

    Else
        'Guardar
        GuardarBinary adodc1.Recordset!Campo, Msg
        adodc1.Recordset.Update
        
        adodc1.RecordSource = "Select * from factpro_fichdocs where false"
        adodc1.Refresh
    End If
    
    
    
    NombreArchivoDestino True
    FileCopy Msg, SqlInsert
    
    
    
    'LLegados a este punto. TODO PERFECTO
    CadenaDesdeOtroForm = Mc.TipoContador & "|" & Mc.Contador & "|" & Year(CDate(Text1(5).Text)) & "|"
    InsertaFactura = True
    
    
eInsertaFactura:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function


Private Sub BorraAlbaranes()

On Error GoTo eBorraAlbaranes

    
    For i = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            '\\PCDAVID\Programas\zBorrame\FraPendientes\00003 Seccion Almazar
            
            SQL = CarpetaAlbaranes & "\" & ListView1.ListItems(i).Tag
            ElminarAlbaran
            
            
            
            SQL = "DELETE FROM factproalbaranes WHERE id = " & Mid(ListView1.ListItems(i).Key, 2)
            Ejecuta SQL
        End If
    Next i
    
    If vParam.EliminaPdfOriginal Then
    
        Kill txtNomFich.Tag
    
    End If
    
eBorraAlbaranes:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    AntiguoText1 = ""
End Sub


Private Function ElminarAlbaran()
Dim Vez As Byte

    On Error Resume Next
    For Vez = 1 To 2
        Kill SQL
        If Err.Number <> 0 Then
            If Err.Number = 70 And Vez = 1 Then
                MsgBox "Albaran abierto. " & vbCrLf & SQL & vbCrLf & vbCrLf & "Ciérrelo y pulse aceptar", vbExclamation
            Else
                MsgBox "Imposible borrar " & SQL, vbExclamation
            End If
            Err.Clear
        Else
            Vez = 3
        End If
    Next Vez
End Function


Private Function ComprobarArchivosSeleccionados() As Boolean
    On Error GoTo eTraerArchivosSeleccionados
    ComprobarArchivosSeleccionados = False
    
    
    SQL = ""
    Msg = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            'Traigo el fichero
           
            If ListView1.ListItems(i).Tag = "" Then
                SQL = SQL & "    - " & ListView1.ListItems(i) & " ->INCORRE"
            Else
                Msg = CarpetaAlbaranes & "\" & ListView1.ListItems(i).Tag
                If Dir(Msg, vbArchive) = "" Then SQL = SQL & "    - " & ListView1.ListItems(i).Tag & " ->NO EXISTE"
                
            End If
        End If
    Next
    If SQL <> "" Then
        Msg = "Error en albaranes seleccionados:" & vbCrLf & vbCrLf & SQL
        MsgBox Msg, vbExclamation
    
    Else
        'El archvivo que voy a crear
        NombreArchivoDestino True
        If Dir(Msg, vbArchive) <> "" Then Kill Msg
        ComprobarArchivosSeleccionados = True
    End If
    Exit Function
eTraerArchivosSeleccionados:
    MuestraError Err.Number, Err.Description
End Function


Private Sub NombreArchivoDestino(DesdeTemp As Boolean)
    Msg = Text1(3).Text
    'No pude contener los siguientes caracteres
    'los reemplazamos
    For K = 1 To 9
        '/\:*?"<>|
        Msg = Replace(Msg, Mid("/\:*?""<>|", K, 1), "")
    Next K
    Msg = Text1(0).Text & "_" & Mid(Replace(Text1(2).Text, " ", ""), 1, 13) & "_" & Msg & ".pdf"
    If DesdeTemp Then Msg = App.Path & "\Temp\" & Msg
    
End Sub

Private Function ConcatenaYSubeFichero() As Boolean
Dim C As String
Dim AlgunoSeleccionado As Boolean
    On Error GoTo eConcatenaYSubeFichero
    ConcatenaYSubeFichero = False
    
    
    
   
    
    
    
    SQL = " """ & txtNomFich.Tag & """"
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            'Traigo el fichero
            Msg = " """ & CarpetaAlbaranes & "\" & ListView1.ListItems(i).Tag & """"
            SQL = SQL & Msg

        End If
    Next
    
    NombreArchivoDestino True
    C = """" & App.Path & "\pdftk.exe"" " & SQL & " cat output """ & Msg & """ verbose"
    Shell C, vbMaximizedFocus
    
    i = 0
    Do
        espera 1
        If Dir(Msg, vbArchive) <> "" Then
            If FileLen(Msg) > 0 Then
                i = 100
            Else
                i = i + 1
            End If
        Else
            i = i + 1
            DoEvents
        End If
    Loop Until i > 30
    
    If i = 31 Then
        MsgBox "Error generando PDF: " & Msg, vbExclamation
    Else
        ConcatenaYSubeFichero = True
    End If
        
            
eConcatenaYSubeFichero:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function
    





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

    lw1(0).Tag = "5|" & Ncol & "|"
    
    lw1(0).ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1(0).ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim


End Sub





'TOTALES
Private Function CalculaTotales() As Boolean
Dim IT

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

Dim TipoRetencion As Integer

    On Error GoTo eRecalcularTotalesFactura

    CalculaTotales = False


    lw1(0).ListItems.Clear

    TipoRetencion = 0
    If Combo1(2).ListIndex > 0 Then TipoRetencion = DevuelveValor("select tipo from usuarios.wtiporeten where codigo = " & DBSet(Combo1(2).ListIndex, "N"))
    
    Baseimpo = 0
    Basereten = 0
    Impoiva = 0
    Imporeten = 0
    ImpoRec = 0
    TotalFactura = 0
    
    SQL = "select tipoiva,  sum(imponible) baseimpo, sum(coalesce(impiva,0)) imporiva, sum(coalesce(recargo,0)) imporrec, sum(if(tipoopera=1,imponible,0))  ParaRet"
    SQL = SQL & "  from tmpfaclin  where codusu = " & vUsu.Codigo
    SQL = SQL & " group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    While Not Rs.EOF
        
        Set IT = lw1(0).ListItems.Add
        J = J + 1
        IT.Text = J
        IT.SubItems(1) = Format(Rs!TipoIva, "000")
        IT.SubItems(2) = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", Rs!TipoIva)
        IT.SubItems(3) = Format(Rs!Baseimpo, "###,###,##0.00")
        IT.SubItems(4) = Format(Rs!Imporiva, "###,###,##0.00")
        If DBLet(Rs!imporrec) <> 0 Then
            IT.SubItems(5) = Format(Rs!imporrec, "###,###,##0.00")
        Else
            IT.SubItems(5) = " "
        End If
        
        Set IT = Nothing
        
        
        Baseimpo = Baseimpo + DBLet(Rs!Baseimpo, "N")
        
        ' en el caso de inversion sujeto pasivo o intracomunitarias no se suma a totales los ivas y recargos
        If Combo1(1).ListIndex = 1 Or Combo1(1).ListIndex = 4 Then
        
        Else
            Impoiva = Impoiva + DBLet(Rs!Imporiva, "N")
            ImpoRec = ImpoRec + DBLet(Rs!imporrec, "N")
        End If
    
        If Combo1(2).ListIndex > 0 Then
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
    
    PorcRet = ImporteFormateado(Text1(10).Text)
    
    If PorcRet = 0 Then Basereten = 0
   
    
    If PorcRet = 0 Then
        Imporeten = 0
        Basereten = 0
    Else
        Imporeten = Round2((PorcRet * Basereten / 100), 2)
    End If
    
    TotalFactura = Baseimpo + Impoiva + ImpoRec - Imporeten
    
    Text1(12).Text = Format(Baseimpo, FormatoImporte)
    Text1(13).Text = Format(Basereten, FormatoImporte)
    Text1(14).Text = Format(Impoiva, FormatoImporte)
    Text1(15).Text = ""
    If Imporeten <> 0 Then Text1(15).Text = Format(Imporeten, FormatoImporte)
    Text1(16).Text = Format(ImpoRec, FormatoImporte)
    Text1(6).Text = Format(TotalFactura, FormatoImporte)
    
    If PorcRet = 0 Then
        Text1(11).Text = ""
        Text4(11).Text = ""
        Text1(10).Text = ""
    End If
    
    
    
    
    
    
    CalculaTotales = True
    Exit Function
    
eRecalcularTotalesFactura:
    MuestraError Err.Number, "Recalcular Totales Factura", Err.Description
End Function




Private Sub LlamaContraPar()
    Set frmCtas = New frmColCtas
    frmCtas.DatosADevolverBusqueda = "0|1"
    frmCtas.ConfigurarBalances = 3
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub





Private Function ComprobarPeriodo(EsEnDatosOK As Boolean, Indice As Integer) As Boolean
Dim F As Date
Dim m As Byte
Dim Cerrado As Boolean

    On Error GoTo eComprobarPeriodo
    Msg = ""
    ComprobarPeriodo = False
    If EsEnDatosOK Then Indice = 5
    F = CDate(Text1(Indice).Text)
    
    
    If F < vParam.fechaini Then
        Msg = "Ejercicio cerrado"
    Else
        If F > DateAdd("yyyy", 1, vParam.fechafin) Then
            Msg = "Ejercicio sin abrir"
        Else
            i = Year(F)
            If vParam.periodos = 0 Then
                'Trimestral
                m = ((Month(F) - 1) \ 3) + 1
            Else
                m = Month(F)
            End If
            Cerrado = False
        
            If i < vParam.anofactu Then
                Cerrado = True
            Else
                If i = vParam.anofactu Then
                    'El mismo año. Comprobamos los periodos
                    If vParam.perfactu >= m Then Cerrado = True
                End If
            End If
    
            If Cerrado Then Msg = "Pertenece a un periodo de IVA liquidado"
        End If
    End If
        
    If vParam.SIITiene Then
        
        
        Cerrado = False
        If EsEnDatosOK Then
            F = CDate(Text1(5).Text)
            If vParam.SII_Periodo_DesdeLiq Then F = CDate(Text1(9).Text)
            Cerrado = True
        Else
        
            F = CDate(Text1(Indice).Text)
            If vParam.SII_Periodo_DesdeLiq Then
                If Indice = 9 Then Cerrado = True
            Else
                If Indice = 5 Then Cerrado = True
            End If
        End If
        
        
        If Cerrado Then
        
            If UltimaFechaCorrectaSII(vParam.SIIDiasAviso, Now) > F Then
                SQL = String(70, "*") & vbCrLf
                SQL = SQL & "SII.  Excede del maximo dias permitido para comunicar la factura" & vbCrLf & vbCrLf & SQL
                Msg = Msg & vbCrLf & SQL
            End If
    
        End If
    End If
    
    If Msg <> "" Then
        If Not EsEnDatosOK Then MsgBox Msg, vbExclamation
        If EsEnDatosOK Then Exit Function
    End If
    ComprobarPeriodo = True
    
    Exit Function
eComprobarPeriodo:
    MuestraError Err.Number, Err.Description
End Function


