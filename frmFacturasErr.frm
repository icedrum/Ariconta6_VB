VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacturasErr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro facturas clientes con ERRORES"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmFacturasErr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10620
      TabIndex        =   74
      Top             =   360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   60
      Top             =   5160
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
   Begin VB.Frame framecabeceras 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      TabIndex        =   37
      Top             =   600
      Width           =   11895
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   3780
         TabIndex        =   76
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   28
         Left            =   3840
         TabIndex        =   6
         Tag             =   "Fecha liquidacion|F|N|||cabfact|fecliqcl|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   900
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   27
         Left            =   120
         TabIndex        =   72
         Tag             =   "año factura|N|S|||cabfacte|totfaccl||N|"
         Text            =   "Text1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   26
         Left            =   60
         TabIndex        =   71
         Tag             =   "año factura|N|S|||cabfacte|anofaccl||S|"
         Text            =   "Text1"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   25
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "Observaciones(Concepto)|T|S|||cabfacte|confaccl|||"
         Text            =   "DDDDDDDDDDDDDDD"
         Top             =   900
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Fecha factura|F|N|||cabfacte|fecfaccl|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   1
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "Nº de serie|T|N|||cabfacte|numserie||S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   2
         Tag             =   "Código factura|N|N|0||cabfacte|codfaccl||S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   10800
         TabIndex        =   52
         Tag             =   "Numero serie|N|S|||cabfacte|numasien|||"
         Text            =   "9999999999"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   5220
         TabIndex        =   51
         Text            =   "Text4"
         Top             =   240
         Width           =   3195
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   4080
         TabIndex        =   3
         Tag             =   "Cuenta cliente|T|N|||cabfacte|codmacta|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Intracomunitaria"
         Height          =   255
         Left            =   8820
         TabIndex        =   4
         Tag             =   "Extranjera|N|S|||cabfacte|intracom|||"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   6
         Left            =   1680
         TabIndex        =   7
         Tag             =   "Base imponible 1|N|N|||cabfacte|ba1faccl|#,###,###,##0.00||"
         Top             =   1635
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   3240
         TabIndex        =   8
         Tag             =   "Tipo IVA 1|N|N|0|100|cabfacte|tp1faccl|||"
         Text            =   "Text1"
         Top             =   1635
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   8
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   50
         Tag             =   "Porcentaje IVA 1|N|S|||cabfacte|pi1faccl|#0.00||"
         Text            =   "Text1"
         Top             =   1635
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   9
         Left            =   6960
         TabIndex        =   9
         Tag             =   "Importe IVA 1|N|S|||cabfacte|ti1faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   10
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   49
         Tag             =   "Porcentaje recargo 1|N|S|||cabfacte|pr1faccl|#0.00||"
         Text            =   "Text1"
         Top             =   1635
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   11
         Left            =   9240
         TabIndex        =   10
         Tag             =   "Importe recargo 1|N|S|||cabfacte|tr1faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   12
         Left            =   1680
         TabIndex        =   11
         Tag             =   "Base imponible 2|N|S|||cabfacte|ba2faccl|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   3240
         TabIndex        =   12
         Tag             =   "Tipo IVA 2|N|S|0|100|cabfacte|tp2faccl|||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   14
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   48
         Tag             =   "Porcentaje IVA 2|N|S|||cabfacte|pi2faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   15
         Left            =   6960
         TabIndex        =   13
         Tag             =   "Importe IVA 2|N|S|||cabfacte|ti2faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   16
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "Porcentaje recargo 2|N|S|||cabfacte|pr2faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   17
         Left            =   9240
         TabIndex        =   47
         Tag             =   "Importe recargo 2|N|S|||cabfacte|tr2faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   18
         Left            =   1680
         TabIndex        =   15
         Tag             =   "Base imponible 3|N|S|||cabfacte|ba3faccl|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2685
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   19
         Left            =   3240
         TabIndex        =   16
         Tag             =   "Tipo IVA 3|N|S|0|100|cabfacte|tp3faccl|||"
         Text            =   "Text1"
         Top             =   2685
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   20
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   46
         Tag             =   "Porcentaje IVA 3|N|S|||cabfacte|pi3faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2685
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   21
         Left            =   6960
         TabIndex        =   17
         Tag             =   "Importe IVA 3|N|S|||cabfacte|ti3faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2685
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   22
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   45
         Tag             =   "Porcentaje recargo 3|N|S|||cabfacte|pr3faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2685
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   23
         Left            =   9240
         TabIndex        =   18
         Tag             =   "Importe recargo 3|N|S|||cabfacte|tr3faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2685
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   42
         Text            =   "Text4"
         Top             =   1635
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   3960
         TabIndex        =   41
         Text            =   "Text4"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   3960
         TabIndex        =   40
         Text            =   "Text4"
         Top             =   2700
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3960
         TabIndex        =   39
         Text            =   "Text4"
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   2520
         TabIndex        =   20
         Tag             =   "Cuenta retencion|T|S|||cabfacte|cuereten|||"
         Text            =   "Text1"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   24
         Left            =   1680
         TabIndex        =   19
         Tag             =   "Porcentaje retencion|N|S|||cabfacte|retfaccl|#0.00||"
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
         TabIndex        =   21
         Tag             =   "Cuenta retencion|N|S|||cabfacte|trefaccl|#,##0.00||"
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
         TabIndex        =   38
         Text            =   "123.123.123.123,11"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "F. Liquidacion"
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   75
         Top             =   660
         Width           =   1035
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   7
         Left            =   4860
         Picture         =   "frmFacturasErr.frx":030A
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   70
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   " Fecha"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   0
         Width           =   495
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmFacturasErr.frx":040C
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmFacturasErr.frx":0497
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Serie"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   68
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Factura"
         Height          =   195
         Index           =   5
         Left            =   2280
         TabIndex        =   67
         Top             =   0
         Width           =   735
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   4740
         Picture         =   "frmFacturasErr.frx":0E99
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   66
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento"
         Height          =   195
         Index           =   8
         Left            =   10800
         TabIndex        =   65
         Top             =   0
         Width           =   975
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   3705
         Picture         =   "frmFacturasErr.frx":189B
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   3705
         Picture         =   "frmFacturasErr.frx":229D
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   3705
         Picture         =   "frmFacturasErr.frx":2C9F
         Top             =   2760
         Width           =   240
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   10440
         Y1              =   3045
         Y2              =   3045
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
         TabIndex        =   64
         Top             =   1395
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
         ForeColor       =   &H00FF8080&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   1320
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
         TabIndex        =   62
         Top             =   1395
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
         TabIndex        =   61
         Top             =   1395
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
         TabIndex        =   60
         Top             =   1395
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
         TabIndex        =   59
         Top             =   1395
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
         TabIndex        =   58
         Top             =   1395
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
         ForeColor       =   &H00FF8080&
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   3795
         Width           =   1455
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   3675
         Picture         =   "frmFacturasErr.frx":36A1
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
         ForeColor       =   &H00FF8080&
         Height          =   360
         Index           =   2
         Left            =   8640
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   3600
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdAux 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   315
      Index           =   1
      Left            =   8160
      TabIndex        =   28
      Top             =   7200
      Width           =   195
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   60
      Top             =   7080
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
      TabIndex        =   26
      Top             =   7200
      Width           =   195
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10620
      TabIndex        =   23
      Top             =   7860
      Width           =   1035
   End
   Begin VB.TextBox txtaux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   3720
      TabIndex        =   25
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
      TabIndex        =   34
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
      TabIndex        =   27
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
      TabIndex        =   30
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
      TabIndex        =   29
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
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   31
      Top             =   7680
      Width           =   2235
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9420
      TabIndex        =   22
      Top             =   7860
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacturasErr.frx":40A3
      Height          =   2595
      Left            =   1680
      TabIndex        =   33
      Top             =   4980
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4577
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
      TabIndex        =   35
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
            Enabled         =   0   'False
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
         Left            =   9720
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas clientes erroneas"
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
      Left            =   2880
      TabIndex        =   73
      Top             =   7740
      Width           =   6195
   End
   Begin VB.Menu mnOpcionesAsiPre 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFacturasErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Const NO = "No encontrado"
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
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private SQL As String
Dim I As Integer
Dim ancho As Integer


'para cuando modifica factura, y vuelve a integrar para forzar el numero de asiento
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
    SQL = "UPDATE linfacte SET numserie='" & Text1(1).Text & "'"
    SQL = SQL & " ,codfaccl = " & Text1(2).Text
    SQL = SQL & " ,anofaccl = " & Text1(26).Text
    SQL = SQL & " WHERE numserie='" & Data1.Recordset!NUmSerie
    SQL = SQL & "' AND codfaccl= " & Data1.Recordset!codfaccl
    SQL = SQL & " AND anofaccl=" & Data1.Recordset!anofaccl
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
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
                If InsertarDesdeForm(Me) Then
                    If SituarData1 Then
                        EnlazaADOs
                        PonerModo 5
                        'Haremos como si pulsamo el boton de insertar nuevas lineas
                        'Ponemos el importe en AUX
                        Aux = ImporteFormateado(Text2(4).Text)
                        cmdCancelar.Caption = "Cabecera"
                        ModificandoLineas = 0
                        'Bloqueamos pa' k nadie entre
                        BloqueaRegistroForm Me
                        AnyadirLinea True, False
                    Else
                        SQL = "Error situando los datos. Llame a soporte técnico." & vbCrLf
                        SQL = SQL & vbCrLf & " CLAVE: FrmFacturas. cmdAceptar. SituarData1"
                        MsgBox SQL, vbCritical
                        Exit Sub
                    End If
                End If
        End If
    Case 4
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hay que comprobar si ha modificado, o no la clave de la factura
                I = 1
                If Data1.Recordset!NUmSerie = Text1(1).Text Then
                    If Data1.Recordset!codfaccl = Text1(2).Text Then
                        If Data1.Recordset!anofaccl = Text1(26).Text Then
                            I = 0
                            'NO HA MODIFICADO NADA
                        End If
                    End If
                End If
            
                'Hacemos MODIFICAR
                If I <> 0 Then
                    RC = False
                    'Modificar claves
                    SQL = " numserie='" & Data1.Recordset!NUmSerie
                    SQL = SQL & "' AND codfaccl= " & Data1.Recordset!codfaccl
                    SQL = SQL & " AND anofaccl=" & Data1.Recordset!anofaccl
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
        'Contador de facturas
        If Modo = 3 Then
            'Intentetamos devolver el contador
            If Text1(0).Text <> "" Then
                'I = FechaCorrecta(CDate(Text1(0).Text))
                'Mc.DevolverContador Mc.TipoContador, I = 0, Mc.Contador
            End If
        End If
        LimpiarCampos
        PonerModo 0
        'Set Mc = Nothing
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
            AntiguoText1 = ""
            If Adodc1.Recordset.EOF Then
                AntiguoText1 = "La factura no tiene lineas."
            Else
                'Comprobamos que el total de factura es el de suma
               ObtenerSigueinteNumeroLinea
               If Aux <> 0 Then AntiguoText1 = "El importe de lineas no suma el importe facturas: " & Format(Aux, FormatoImporte)
            End If
            If AntiguoText1 <> "" Then
                AntiguoText1 = AntiguoText1 & vbCrLf & "¿Continuar?"
                If MsgBox(AntiguoText1, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
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
            If CStr(.Fields!NUmSerie) = Text1(1).Text Then
                If CStr(.Fields!anofaccl) = Text1(26).Text Then
                    If CStr(.Fields!codfaccl) = Text1(2).Text Then
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


Private Function IntegrarFactura() As Boolean
Dim RF As ADODB.Recordset
Dim FechaBien As Byte
Dim MCContador As Long
Dim EsDeAnal As Boolean


IntegrarFactura = False


'Comprobamos k existen las cuenas, de IVA y demas
IntegrarFactura = False
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
        If Not IsNull(Adodc1.Recordset.Fields(2)) And IsNull(Adodc1.Recordset.Fields(3)) Then SQL = SQL & "Centro de coste incorrecta: " & Adodc1.Recordset.Fields(3) & vbCrLf
        Adodc1.Recordset.MoveNext
    Wend
    Adodc1.Recordset.MoveFirst
End If
If SQL <> "" Then
    MsgBox SQL, vbExclamation
    Exit Function
End If



'Compruebo k si es analitica, k las bases tienen analitica
'Hay lineas
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
    
    
End If
    


SQL = ""
'Fecha BIEN
FechaBien = FechaCorrecta2(CDate(Adodc2.Recordset!fecfaccl))
If FechaBien > 1 Then
    If FechaBien = 2 Then
        SQL = varTxtFec
    Else
        SQL = "Fecha no pertenece al ejercicio actual ni al siguiente"
    End If
    MsgBox SQL, vbExclamation
    Exit Function
End If

'Comprobamos k la serie existe
SQL = DevuelveDesdeBD("nomregis", "contadores", "tiporegi", Data1.Recordset!NUmSerie, "T")
If SQL = "" Then
    MsgBox "Numero de serie incorrecto", vbExclamation
    Exit Function
End If



'Primero comprobamos que esta cuadrada
If IsNull(Adodc2.Recordset!totfaccl) Then
    MsgBox "La factura no tiene importes", vbExclamation
    Exit Function
End If
'Sumamos las bases
Base = 0
If Not IsNull(Adodc2.Recordset!ba1faccl) Then Base = Base + Adodc2.Recordset!ba1faccl
If Not IsNull(Adodc2.Recordset!ba2faccl) Then Base = Base + Adodc2.Recordset!ba2faccl
If Not IsNull(Adodc2.Recordset!ba3faccl) Then Base = Base + Adodc2.Recordset!ba3faccl
AUX2 = Base 'Sumatorio imponibles

'Le sumamos los IVAS
If Not IsNull(Adodc2.Recordset!ti1faccl) Then Base = Base + Adodc2.Recordset!ti1faccl
If Not IsNull(Adodc2.Recordset!ti2faccl) Then Base = Base + Adodc2.Recordset!ti2faccl
If Not IsNull(Adodc2.Recordset!ti3faccl) Then Base = Base + Adodc2.Recordset!ti3faccl

'Los recargos
If Not IsNull(Adodc2.Recordset!tr1faccl) Then Base = Base + Adodc2.Recordset!tr1faccl
If Not IsNull(Adodc2.Recordset!tr2faccl) Then Base = Base + Adodc2.Recordset!tr2faccl
If Not IsNull(Adodc2.Recordset!tr3faccl) Then Base = Base + Adodc2.Recordset!tr3faccl

'La retencion( es en negativo)
If Not IsNull(Adodc2.Recordset!trefaccl) Then Base = Base - Adodc2.Recordset!trefaccl

If Base <> Adodc2.Recordset!totfaccl Then
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





    'Si tiene cuentas bloqueadas
    If vParam.CuentasBloqueadas <> "" Then
    
        'Primero compruebo las cuentas de cabecera
        If EstaLaCuentaBloqueada(Adodc2.Recordset!codmacta, Adodc2.Recordset!fecfaccl) Then
            MsgBox "Cuenta bloqueada: " & Adodc2.Recordset!codprove, vbExclamation
            Exit Function
        End If
        SQL = DBLet(Adodc2.Recordset!cuereten, "T")
        If SQL <> "" Then
            If EstaLaCuentaBloqueada(SQL, Adodc2.Recordset!fecfaccl) Then
                MsgBox "Cuenta bloqueada: " & SQL, vbExclamation
                Exit Function
            End If
        End If
        'Luego las lineas
        SQL = ""
        Adodc1.Recordset.MoveFirst
        While Not Adodc1.Recordset.EOF
            If EstaLaCuentaBloqueada(Adodc1.Recordset!codtbase, Adodc2.Recordset!fecfaccl) Then
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


    
'    'Obtenemos el contador
'    If Mc.ConseguirContador(Data1.Recordset!NUmSerie, FechaBien = 0, False) = 1 Then
'        MsgBox "Error consiguiendo contador.", vbExclamation
'        Exit Function
'    End If
'
    
    'Comprobamos k no existe la factura
    
    MCContador = Data1.Recordset!codfaccl
    
'Si llega aqui es k podemos integrar
    Conn.BeginTrans
        I = 0
        On Error Resume Next
        
        'Primero UPDATEAMOS el numero de factura en la cabecera
        If I = 0 Then
            SQL = "UPDATE cabfacte set codfaccl=" & MCContador
            SQL = SQL & " WHERE numserie = '" & Data1.Recordset!NUmSerie & "'"
            SQL = SQL & " AND codfaccl = " & Data1.Recordset!codfaccl
            SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
            
            Conn.Execute SQL
            If Err.Number <> 0 Then
                I = 1
                Err.Clear
            End If
        End If
            
        'Actualizamos  el numero de factura en las lineas
        If I = 0 Then
            SQL = "UPDATE linfacte set codfaccl=" & MCContador
            SQL = SQL & " WHERE numserie = '" & Data1.Recordset!NUmSerie & "'"
            SQL = SQL & " AND codfaccl = " & Data1.Recordset!codfaccl
            SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
            
            Conn.Execute SQL
            If Err.Number <> 0 Then
                I = 1
                Err.Clear
            End If
        End If
    
        If I = 0 Then
            SQL = "INSERT INTO cabfact SELECT * from cabfacte WHERE "
            SQL = SQL & " numserie = '" & Data1.Recordset!NUmSerie & "'"
            SQL = SQL & " AND codfaccl = " & MCContador
            SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
            
            Conn.Execute SQL
            If Err.Number <> 0 Then
                I = 1
                Err.Clear
            End If
        End If
        
        If I = 0 Then
            SQL = "INSERT INTO linfact SELECT * from linfacte WHERE "
            SQL = SQL & " numserie = '" & Data1.Recordset!NUmSerie & "'"
            SQL = SQL & " AND codfaccl = " & MCContador
            SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
            Conn.Execute SQL
            If Err.Number <> 0 Then
                I = 1
                Err.Clear
            End If
        End If
        If I = 0 Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
        End If
        On Error GoTo 0
        'Borramos la linea
        If I = 0 Then
            IntegrarFactura = True
            If Not BorrarFactura(MCContador) Then
                MsgBox "Error: Elimine la factura errornea manualmente", vbExclamation
            Else
                IntegrarFactura = True
                MsgBox "Traspaso realizado con éxito", vbExclamation
            End If
        Else
            'Devolvemos contador
           ' Mc.DevolverContador Data1.Recordset!NUmSerie, FechaBien = 0, Mc.Contador
        End If
        'Set Mc = Nothing
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
    Text1(0).Text = Format(Now, "dd/mm/yyyy")
    Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        cmdCancelar.Caption = "Cancelar"
        cmdAceptar.Caption = "Aceptar"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(0).SetFocus
        Text1(0).BackColor = vbYellow
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
        MandaBusquedaPrevia SQL
    Else
        If SQL <> "" Then SQL = " WHERE " & SQL
        CadenaConsulta = "Select numserie,codfaccl,anofaccl from cabfacte " & SQL & Ordenacion
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
    If Adodc2.Recordset.EOF Then Exit Sub
    
    'Comprobamos la fecha pertenece al ejercicio
    varFecOk = FechaCorrecta2(CDate(Text1(0).Text))
    If varFecOk >= 2 Then
        If varFecOk = 2 Then
            MsgBox varTxtFec
        Else
            MsgBox "La factura pertenece a un ejercicio cerrado.", vbExclamation
        End If
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    PonerModo 4
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    
End Sub

Private Sub BotonEliminar()
    Dim I As Integer

    'Ciertas comprobaciones
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    If Adodc2.Recordset.EOF Then Exit Sub
    DataGrid1.Enabled = False
    
    'Bloqueamos
    
    
    'Comprobamos que no esta actualizada ya
    SQL = ""
    If Not IsNull(Adodc2.Recordset!Numasien) Then
        SQL = "Esta factura ya esta actualizada. "
    End If
    
    SQL = SQL & vbCrLf & vbCrLf & "Va usted a eliminar la factura errónea :" & vbCrLf
    SQL = SQL & "Numero : " & Adodc2.Recordset!codfaccl & vbCrLf
    SQL = SQL & "Fecha  : " & Adodc2.Recordset!fecfaccl & vbCrLf
    SQL = SQL & "Cliente : " & Adodc2.Recordset!codmacta & " - " & Text4(0).Text & vbCrLf
    SQL = SQL & vbCrLf & "          ¿Desea continuar ?" & vbCrLf
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    NumRegElim = Data1.Recordset.AbsolutePosition
    Screen.MousePointer = vbHourglass
    'Lo hara en actualizar
    
        'La borrara desde este mismo form
        If BorrarFactura(0) Then
            AlgunAsientoActualizado = True
        Else
            AlgunAsientoActualizado = False
        End If
    
    If Not AlgunAsientoActualizado Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    NumRegElim = Data1.Recordset.AbsolutePosition
    If Data1.Recordset.RecordCount = NumRegElim Then
        NumRegElim = NumRegElim - 2
    Else
        NumRegElim = NumRegElim - 1
    End If
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        CargaGrid False
        PonerModo 0
        Else
            If NumRegElim > 0 Then Data1.Recordset.Move NumRegElim
            PonerCampos
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function BorrarFactura(numero As Long) As Boolean
    
    On Error GoTo EBorrar
    SQL = " WHERE numserie = '" & Data1.Recordset!NUmSerie & "'"
    If numero > 0 Then
        SQL = SQL & " AND codfaccl = " & numero
    Else
        SQL = SQL & " AND codfaccl = " & Data1.Recordset!codfaccl
    End If
    SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
    'Las lineas
    AntiguoText1 = "DELETE from linfacte " & SQL
    Conn.Execute AntiguoText1
    'La factura
    AntiguoText1 = "DELETE from cabfacte " & SQL
    Conn.Execute AntiguoText1
EBorrar:
    If Err.Number = 0 Then
        BorrarFactura = True
    Else
        MuestraError Err.Number, "Eliminar factura"
        BorrarFactura = False
    End If
End Function


Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False

        PonerModo CInt(Modo)
        CargaGrid (Modo = 2)
        If Modo <> 2 Then
            CadenaConsulta = "Select * from cabfacte " & Ordenacion
            Data1.RecordSource = CadenaConsulta
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    SQL = ""
    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    
    'Si mostramos fecha liquidacion o no
    FrameTapa.Visible = Not vParam.Constructoras
    Text1(28).Enabled = vParam.Constructoras
    
    
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
    Ordenacion = " ORDER BY numserie,fecfaccl"
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
    Data1.RecordSource = "Select numserie,codfaccl,anofaccl from Cabfacte WHERE numserie ='David'"
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
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
        CadB = Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 2)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        
        Aux = ValorDevueltoFormGrid(Text1(26), CadenaDevuelta, 3)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & Aux
        'Se muestran en el mismo form
        CadenaConsulta = "select * from cabfacte WHERE " & CadB & " "
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
'    If Text1(0).Text = "" Then
'         MsgBox "No hay fecha seleccionada ", vbExclamation
'         Exit Sub
'    End If
'    SQL = RecuperaValor(CadenaSeleccion, 1)
'    b = CDate(Text1(0).Text) <= vParam.fechafin
'    If Mc Is Nothing Then Set Mc = New Contadores
'    If Mc.ConseguirContador(SQL, b, False) = 0 Then
'        Text1(1).Text = SQL
'        Text1(2).Text = Mc.Contador
'    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(CInt(cmdAux(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
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
    Case 0
        Set frmCo = New frmContadores
        frmCo.DatosADevolverBusqueda = "0|"
        frmCo.Show vbModal
        Set frmCo = Nothing
        If Text1(1).Text <> "" Then Text1(3).SetFocus
    Case 1, 7
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Index = 1 Then
            I = 0
        Else
            I = 28
        End If
        cmdAux(0).Tag = I
        If Text1(I).Text <> "" Then frmF.Fecha = CDate(Text1(I).Text)
        frmF.Show vbModal
        Set frmF = Nothing
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
Dim L As Long
    Text1(Indice).Text = Trim(Text1(Indice).Text)
    If Text1(Indice).Text = "" Then
        'Hemos puesto a blancos el campo, luego quitaremos
        'los valores asociados a el
        If Text1(Indice) = AntiguoText1 Then Exit Sub
        Select Case Indice
        Case 0
            'Ponemos a blanco tb el año de factura
            Text1(26).Text = ""
        Case 1, 28
            'Ha puesto a blanco la serie de las facturas
            'por lo tanto habra que mirar si es el ultimo
            If Text1(0).Text <> "" Then
                Correcto = CDate(Text1(0).Text) <= vParam.fechafin
                If Text1(2).Text <> "" Then
                    Linfac = Val(Text1(2).Text)
                    'Mc.DevolverContador AntiguoText1, Correcto, Linfac
                End If
            End If
        Case 6, 12, 18, 9, 15, 21, 11, 17, 23
            
            'Los importes
            CalcularIVA I
            TotalesRecargo
            TotalesIVA
            TotalFactura
        Case 3
            Text4(4).Text = ""
        Case 5
            Text4(0).Text = ""
        Case 7
            Text4(1).Text = ""
        Case 13
            Text4(2).Text = ""
        Case 19
            Text4(3).Text = ""
        Case 24
            Text2(3).Text = ""
            TotalFactura
        End Select
    Else
        With Text1(Indice)
           Select Case Indice
           Case 0, 28
                If Not EsFechaOK(Text1(Indice)) Then
                    MsgBox "Fecha incorrecta: " & .Text, vbExclamation
                    .Text = ""
                    If Indice = 0 Then Text1(26).Text = ""
                    Text1(Indice).SetFocus
                    Exit Sub
                End If
                
                'Hay que comprobar que las fechas estan
                'en los ejercicios y si
                '       0 .- Año actual
                '       1 .- Siguiente
                '       2 .- Ambito
                '       3 .- Anterior al inicio
                '       4 .- Posterior al fin
                ModificandoLineas = FechaCorrecta2(CDate(.Text))
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
                    MsgBox RC, vbExclamation
                    .Text = ""
                    If Indice = 0 Then Text1(26).Text = ""
                    Text1(Indice).SetFocus
                    Exit Sub
                End If
                
                
                .Text = Format(.Text, "dd/mm/yyyy")
                If Indice = 0 Then Text1(26).Text = Year(CDate(.Text))
                
                'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
                'que el periodo de liquidacion del IVA no ha pasado.
                If Not ComprobarPeriodo(Indice) Then Text1(Indice).SetFocus
                
                
                
                

                
                
                
                
            Case 1
                 If IsNumeric(.Text) Then
                    MsgBox "Debe ser una letra: " & .Text, vbExclamation
                    .Text = ""
                    .SetFocus
                End If
                .Text = UCase(.Text)
                If .Text = AntiguoText1 Then Exit Sub
                'letra distinta
                'ASignaremos contador, si la feha esta puesta
                If Text1(0).Text <> "" Then
                    Correcto = CDate(Text1(0).Text) <= vParam.fechafin
                    If Text1(2).Text <> "" Then
                        L = Val(Text1(2).Text)
                    Else
                        L = 0
                    End If
'                    If Mc.ConseguirContador(.Text, Correcto, False) = 0 Then
'                        Text1(2).Text = Mc.Contador
'                    Else
'                        MsgBox "La letra no es de contadores: " & .Text, vbExclamation
'                        .Text = ""
'                        Text1(2).Text = ""
'                        .SetFocus
'                    End If
                End If

            Case 2
                If Not IsNumeric(.Text) Then
                    MsgBox "El numero de factura no es correcto: " & .Text, vbExclamation
                    .Text = ""
                    .SetFocus
                End If
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
                                            'menos si no estamos buscando, k dejaremos
                    If Modo = 1 Then
                        SQL = ""
                    Else
                        MsgBox SQL, vbExclamation
                        .Text = ""
                        Text4(I).Text = ""
                        .SetFocus
                    End If
                End If
                
            Case 7, 13, 19  'TIpos de iva
                I = ((Indice - 1) / 6)
                If Not IsNumeric(.Text) Then
                
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
                        Text2(3).Text = Format(Base, FormatoImporte)
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
                CadenaConsulta = "select numserie,codfaccl,anofaccl from cabfacte WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(1), 10, "Serie: ")
        Cad = Cad & ParaGrid(Text1(2), 20, "Nº Fac.")
        Cad = Cad & ParaGrid(Text1(26), 10, "Año")
        Cad = Cad & ParaGrid(Text1(0), 30)
        Cad = Cad & ParaGrid(Text1(6), 30)
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.VCampos = Cad
            frmB.vTabla = "cabfacte"
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|"
            frmB.vTitulo = "Facturas"
            frmB.vSelElem = 4
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
        MsgBox "No hay ningún registro en la tabla Facturas clientes.", vbInformation
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


Private Function EnlazaADOs() As String
Dim SQL As String
    SQL = "Select * from cabfacte where codfaccl=" & Data1.Recordset!codfaccl
    SQL = SQL & " AND anofaccl = " & Data1.Recordset!anofaccl
    SQL = SQL & " AND numserie= '" & Data1.Recordset!NUmSerie & "'"
    Adodc2.ConnectionString = Conn
    Adodc2.RecordSource = SQL
    Adodc2.Refresh
End Function



Private Sub PonerCampos()
    Dim mTag As CTag
    Dim SQL As String
    If Data1.Recordset.EOF Then Exit Sub
    
    EnlazaADOs
    
    PonerCamposForma Me, Adodc2
    
    'Cargamos el LINEAS
    DataGrid1.Enabled = False
    CargaGrid True
    If Modo = 2 Then DataGrid1.Enabled = True
    
    'En SQL almacenamos el importe
    Base = Adodc2.Recordset!totfaccl
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
    mnNuevo.Enabled = Not B
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
'    Dim RS As ADODB.Recordset
    Dim B As Boolean
    
     'Si no es constructoras igualamos los campos fecfac y fecliquidacion
    If Not vParam.Constructoras Then Text1(28).Text = Text1(0).Text
    
    B = CompForm(Me)
    
    If Not B Then Exit Function
    
   
    'No puede tener % de retencion sin cuenta de retencion
    If ((Text1(24).Text = "") Xor (Text1(3).Text = "")) Then
        MsgBox "No hay porcentaje de rentencion sin cuenta de retencion", vbExclamation
        B = False
        Exit Function
    End If
    
    
    'Compruebo si hay cuentas bloqueadas
    If vParam.CuentasBloqueadas <> "" Then
        If EstaLaCuentaBloqueada(Text1(5).Text, CDate(Text1(0).Text)) Then
            MsgBox "Cuenta bloqueada: " & Text1(5).Text, vbExclamation
            B = False
            Exit Function
        End If
        If Text1(3).Text <> "" Then
            If EstaLaCuentaBloqueada(Text1(3).Text, CDate(Text1(0).Text)) Then
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
        If Adodc2.Recordset.EOF Then Exit Sub
        
        
        'Comprobamos k no existe el numero de factura
        If ExisteFactura Then
            MsgBox "Ya existe una  factura con ese número", vbExclamation
            Exit Sub
        End If
        SQL = "Seguro que desea corregir la factura" & vbCrLf
        SQL = SQL & "Numero: " & Data1.Recordset!NUmSerie & " - " & Data1.Recordset!codfaccl & vbCrLf
        SQL = SQL & "Fecha : " & Adodc2.Recordset!fecfaccl & "?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Screen.MousePointer = vbHourglass
            NumRegElim = Data1.Recordset.AbsolutePosition
            If Data1.Recordset.RecordCount = NumRegElim Then
                NumRegElim = NumRegElim - 2
            Else
                NumRegElim = NumRegElim - 1
            End If
            
            If IntegrarFactura Then
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

        'Verificar
        HazVerificacion
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
    
    SQL = "SELECT linfacte.codtbase, cuentas.nommacta, linfacte.codccost, cabccost.nomccost, linfacte.impbascl, linfacte.numlinea"
    SQL = SQL & " FROM (cabccost RIGHT JOIN linfacte ON cabccost.codccost = linfacte.codccost) LEFT JOIN cuentas ON linfacte.codtbase = cuentas.codmacta WHERE "
    If Enlaza Then
        SQL = SQL & " numserie = '" & Data1.Recordset!NUmSerie & "'"
        SQL = SQL & " AND codfaccl = " & Data1.Recordset!codfaccl
        SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
        Else
        SQL = SQL & " numserie = 'david'"
    End If
    SQL = SQL & " ORDER BY linfacte.numlinea"
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
            If MsgBox(" Importes exactos. ¿Continuar?", vbQuestion + vbYesNo) = vbYes Then anc = 1
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
    txtaux(4).Text = Adodc1.Recordset!impbascl

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
        SQL = "Delete from linfacte"
        SQL = SQL & " WHERE numlinea = " & Adodc1.Recordset!NumLinea
        SQL = SQL & " AND anofaccl=" & Data1.Recordset!anofaccl
        SQL = SQL & " AND numserie='" & Data1.Recordset!NUmSerie
        SQL = SQL & "' AND codfaccl=" & Data1.Recordset!codfaccl & ";"
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
    
    SQL = " WHERE linfacte.numserie= '" & Data1.Recordset!NUmSerie & "'"
    SQL = SQL & " AND linfacte.codfaccl= " & Data1.Recordset!codfaccl
    SQL = SQL & " AND linfacte.anofaccl=" & Data1.Recordset!anofaccl & ";"
    RS.Open "SELECT Max(numlinea) FROM linfacte" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then I = RS.Fields(0)
    End If
    RS.Close

    'La suma
    SumaLinea = 0
    If I > 0 Then
        RS.Open "SELECT sum(impbascl) FROM linfacte" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
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
                'MsgBox "Importe incorrecto: " & txtaux(4).Text, vbExclamation
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
    If EstaLaCuentaBloqueada(txtaux(0).Text, CDate(Text1(0).Text)) Then
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
        AUX2 = AUX2 - Adodc1.Recordset!impbascl
    End If
    If AUX2 > 0 Then
        If AUX2 > Aux Then
'                AuxOK = "El importe excede del total de factura"
'                Exit Function
        End If
    Else
        If AUX2 < Aux Then
'                AuxOK = "El importe excede del total de factura"
'                Exit Function
        End If
    End If
    AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        SQL = "INSERT INTO linfacte (numserie, codfaccl, anofaccl, numlinea, codtbase, impbascl, codccost) VALUES ('"
        ''R', 11, 2003, 1, '6000001', 1500, 'TIEN')
        SQL = SQL & Data1.Recordset!NUmSerie & "',"
        SQL = SQL & Data1.Recordset!codfaccl & ","
        SQL = SQL & Data1.Recordset!anofaccl & "," & Linfac & ",'"
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
        SQL = "UPDATE linfacte SET "
        
        SQL = SQL & " codtbase = '" & txtaux(0).Text & "',"
        SQL = SQL & " impbascl = "
        SQL = SQL & TransformaComasPuntos(txtaux(4).Text) & ","
        
        'Centro coste
        If txtaux(2).Text = "" Then
          SQL = SQL & " codccost = " & ValorNulo
          Else
          SQL = SQL & " codccost = '" & txtaux(2).Text & "'"
        End If
    
        SQL = SQL & " WHERE numserie='" & Data1.Recordset!NUmSerie
        SQL = SQL & "' AND codfaccl= " & Data1.Recordset!codfaccl
        SQL = SQL & " AND anofaccl=" & Data1.Recordset!anofaccl
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
Private Function ComprobarPeriodo(Indice As Integer) As Boolean
Dim Cerrado As Boolean
'Primero pondremos la fecha a año periodo
I = Year(CDate(Text1(Indice).Text))
If vParam.periodos Then
    'Trimestral
    ancho = (CDate(Text1(Indice).Text) Mod 3) + 1
    Else
    ancho = Month(CDate((Text1(Indice).Text)))
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
    SQL = "La fecha "
    If Indice = 0 Then
        SQL = SQL & "factura"
    Else
        SQL = SQL & "liquidacion"
    End If
    SQL = SQL & "corresponde a un periodo ya liquidado. " & vbCrLf & " ¿Desea continuar igualmente ?"
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








Private Sub HazVerificacion()
Dim VC As String
Dim RT As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Set RT = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    AntiguoText1 = ""
    
    
    
    SQL = "SELECT cabfacte.numserie, contadores.nomregis, cabfacte.codfaccl"
    SQL = SQL & " FROM cabfacte LEFT JOIN contadores ON cabfacte.numserie = contadores.tiporegi"
    SQL = SQL & " WHERE (((contadores.nomregis) Is Null));"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RS.EOF
        SQL = SQL & RS.Fields(0) & " - " & RS.Fields(2) & vbCrLf
        RS.MoveNext
    Wend
    RS.Close
    If SQL <> "" Then SQL = "Serie incorrecta" & vbCrLf & "----------------------" & vbCrLf & SQL
    
    
    RS.Open "Select numserie,codfaccl,anofaccl from linfacte group by numserie,codfaccl,anofaccl ", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = 1
        VC = "Select codfaccl from cabfacte where anofaccl=" & RS!anofaccl
        VC = VC & " AND codfaccl = " & RS!codfaccl
        VC = VC & " AND numserie ='" & RS!NUmSerie & "'"
        RT.Open VC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RT.EOF Then AntiguoText1 = AntiguoText1 & Format(RS!codfaccl, "00000000") & "    " & RS!NUmSerie & "      " & RS!anofaccl & vbCrLf
        RT.Close
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Set RT = Nothing
    Screen.MousePointer = vbDefault
    If AntiguoText1 <> "" Then
        AntiguoText1 = " Las siguientes lineas no corresponden a ningun encabezado de factura erronea." & vbCrLf & vbCrLf & _
            "   Codigo    Serie   Año  " & vbCrLf & "----------------------------" & vbCrLf & AntiguoText1
    End If
    If SQL <> "" Then AntiguoText1 = SQL & vbCrLf & AntiguoText1

        
    If AntiguoText1 <> "" Then
        MsgBox AntiguoText1, vbExclamation
    Else
        MsgBox "Comprobación finalizada", vbInformation
    End If
    
End Sub




Private Function ExisteFactura() As Boolean

    Set RS = New ADODB.Recordset
    ExisteFactura = True
    SQL = "Select * from cabfact where Numserie='" & Data1.Recordset!NUmSerie
    SQL = SQL & "' AND codfaccl = " & Data1.Recordset!codfaccl
    SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        
        
        'Podrimaos comprobar las bases, pero no lo hacemos
        ExisteFactura = False
    End If
    RS.Close
    Set RS = Nothing
End Function
