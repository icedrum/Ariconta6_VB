VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacturas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmFacturas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   210
      Top             =   5310
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Height          =   4215
      Left            =   0
      TabIndex        =   39
      Top             =   600
      Width           =   11895
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   5520
         TabIndex        =   76
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   28
         Left            =   5640
         TabIndex        =   6
         Tag             =   "Fecha liquidacion|F|N|||cabfact|fecliqcl|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   27
         Left            =   120
         TabIndex        =   73
         Tag             =   "total factura|N|S|||cabfact|totfaccl||N|"
         Text            =   "Text1"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   26
         Left            =   120
         TabIndex        =   72
         Tag             =   "año factura|N|S|||cabfact|anofaccl||S|"
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   25
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Observaciones(Concepto)|T|S|||cabfact|confaccl|||"
         Text            =   "DDDDDDDDDDDDDDD"
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Fecha factura|F|N|||cabfact|fecfaccl|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   1
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "Nº de serie|T|N|||cabfact|numserie||S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   2160
         TabIndex        =   2
         Tag             =   "Código factura|N|N|0||cabfact|codfaccl||S|"
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   10560
         TabIndex        =   53
         Tag             =   "Numero serie|N|S|||cabfact|numasien|||"
         Text            =   "9999999999"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4800
         TabIndex        =   52
         Text            =   "Text4"
         Top             =   240
         Width           =   3795
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   3600
         TabIndex        =   3
         Tag             =   "Cuenta cliente|T|N|||cabfact|codmacta|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Intracomunitaria"
         Height          =   255
         Left            =   8880
         TabIndex        =   4
         Tag             =   "Extranjera|N|S|||cabfact|intracom|||"
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   6
         Left            =   1680
         TabIndex        =   7
         Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
         Top             =   1620
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   3240
         TabIndex        =   8
         Tag             =   "Tipo IVA 1|N|N|0|100|cabfact|tp1faccl|||"
         Text            =   "Text1"
         Top             =   1620
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   8
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   51
         Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
         Text            =   "Text1"
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   9
         Left            =   6960
         TabIndex        =   9
         Tag             =   "Importe IVA 1|N|N|||cabfact|ti1faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   10
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   50
         Tag             =   "Porcentaje recargo 1|N|S|||cabfact|pr1faccl|#0.00||"
         Text            =   "Text1"
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   11
         Left            =   9240
         TabIndex        =   10
         Tag             =   "Importe recargo 1|N|S|||cabfact|tr1faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   12
         Left            =   1680
         TabIndex        =   11
         Tag             =   "Base imponible 2|N|S|||cabfact|ba2faccl|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2100
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   3240
         TabIndex        =   12
         Tag             =   "Tipo IVA 2|N|S|0|100|cabfact|tp2faccl|||"
         Text            =   "Text1"
         Top             =   2100
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   14
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   49
         Tag             =   "Porcentaje IVA 2|N|S|||cabfact|pi2faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2100
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   15
         Left            =   6960
         TabIndex        =   13
         Tag             =   "Importe IVA 2|N|S|||cabfact|ti2faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2100
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   16
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "Porcentaje recargo 2|N|S|||cabfact|pr2faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2100
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   17
         Left            =   9240
         TabIndex        =   48
         Tag             =   "Importe recargo 2|N|S|||cabfact|tr2faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2100
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   18
         Left            =   1680
         TabIndex        =   15
         Tag             =   "Base imponible 3|N|S|||cabfact|ba3faccl|#,###,###,##0.00||"
         Text            =   "Text1"
         Top             =   2565
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   19
         Left            =   3240
         TabIndex        =   16
         Tag             =   "Tipo IVA 3|N|S|0|100|cabfact|tp3faccl|||"
         Text            =   "Text1"
         Top             =   2565
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   20
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   47
         Tag             =   "Porcentaje IVA 3|N|S|||cabfact|pi3faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2565
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   21
         Left            =   6960
         TabIndex        =   17
         Tag             =   "Importe IVA 3|N|S|||cabfact|ti3faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2565
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   22
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   46
         Tag             =   "Porcentaje recargo 3|N|S|||cabfact|pr3faccl|#0.00||"
         Text            =   "Text1"
         Top             =   2565
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   23
         Left            =   9240
         TabIndex        =   18
         Tag             =   "Importe recargo 3|N|S|||cabfact|tr3faccl|#,###,##0.00||"
         Text            =   "Text1"
         Top             =   2565
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text2"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   3000
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
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   43
         Text            =   "Text4"
         Top             =   1620
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   3960
         TabIndex        =   42
         Text            =   "Text4"
         Top             =   2100
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   3960
         TabIndex        =   41
         Text            =   "Text4"
         Top             =   2565
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3960
         TabIndex        =   40
         Text            =   "Text4"
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   2520
         TabIndex        =   20
         Tag             =   "Cuenta retencion|T|S|||cabfact|cuereten|||"
         Text            =   "Text1"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   24
         Left            =   1680
         TabIndex        =   19
         Tag             =   "Porcentaje retencion|N|S|||cabfact|retfaccl|#0.00||"
         Text            =   "Text1"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   21
         Tag             =   "Cuenta retencion|N|S|||cabfact|trefaccl|#,##0.00||"
         Text            =   "Text2"
         Top             =   3720
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
         TabIndex        =   22
         Text            =   "123.123.123.123,11"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   7
         Left            =   6660
         Picture         =   "frmFacturas.frx":000C
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Liquidacion"
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   75
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   71
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   " Fecha"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   0
         Width           =   495
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmFacturas.frx":010E
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   2160
         Picture         =   "frmFacturas.frx":0199
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Serie"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   69
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Factura"
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   68
         Top             =   0
         Width           =   735
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   2
         Left            =   4200
         Picture         =   "frmFacturas.frx":0B9B
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   7
         Left            =   3600
         TabIndex        =   67
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento"
         Height          =   195
         Index           =   8
         Left            =   10560
         TabIndex        =   66
         Top             =   0
         Width           =   975
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   3
         Left            =   3705
         Picture         =   "frmFacturas.frx":159D
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   4
         Left            =   3705
         Picture         =   "frmFacturas.frx":1F9F
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   5
         Left            =   3705
         Picture         =   "frmFacturas.frx":29A1
         Top             =   2640
         Width           =   240
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   10440
         Y1              =   2925
         Y2              =   2925
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
         TabIndex        =   65
         Top             =   1380
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   64
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
         TabIndex        =   63
         Top             =   1380
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
         TabIndex        =   62
         Top             =   1380
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
         TabIndex        =   61
         Top             =   1380
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
         TabIndex        =   60
         Top             =   1380
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
         TabIndex        =   59
         Top             =   1380
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   3675
         Width           =   1455
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   6
         Left            =   3675
         Picture         =   "frmFacturas.frx":33A3
         Top             =   3750
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Index           =   2
         Left            =   8640
         TabIndex        =   57
         Top             =   3720
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
         TabIndex        =   56
         Top             =   3480
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
         TabIndex        =   55
         Top             =   3480
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
         TabIndex        =   54
         Top             =   3480
         Width           =   570
      End
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
      Left            =   8160
      Top             =   7380
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Top             =   7440
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
      Top             =   6720
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
      Top             =   7440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   32
      Top             =   7320
      Width           =   3495
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9480
      TabIndex        =   23
      Top             =   7440
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFacturas.frx":3DA5
      Height          =   2355
      Left            =   1680
      TabIndex        =   35
      Top             =   4860
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4154
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
            Object.ToolTipText     =   "Contabilizar factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
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
         Left            =   9480
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar ToolAux 
      Height          =   390
      Index           =   0
      Left            =   450
      TabIndex        =   77
      Top             =   4860
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
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
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "C L I E N T E S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4440
      TabIndex        =   74
      Top             =   7320
      Width           =   3495
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
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro ejercicio"
      Begin VB.Menu mnActuralySiguiente 
         Caption         =   "Actual y siguiente"
      End
      Begin VB.Menu mnActual 
         Caption         =   "Actual"
      End
      Begin VB.Menu mnSiguiente 
         Caption         =   "Siguiente"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSinfiltro 
         Caption         =   "Sin filtro"
      End
   End
End
Attribute VB_Name = "frmFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const IdPrograma = 401
Public FACTURA As String  'Con pipes numdiari|fechanormal|numasien


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
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean
Private SQL As String
Dim i As Integer
Dim ancho As Integer


'para cuando modifica factura, y vuelve a integrar para forzar el numero de asiento
Dim Numasien2 As Long
Dim NumDiario As Integer
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas



'Para pasar de lineas a cabeceras
Dim Linfac As Long
Private ModificandoLineas As Byte
'0.- A la espera 1.- Insertar   2.- Modificar

Dim PrimeraVez As Boolean
Dim PulsadoSalir As Boolean
Dim Rs As Recordset
Dim AUx As Currency
Dim Base As Currency
Dim AUX2 As Currency
Dim SumaLinea As Currency
Dim AntiguoText1 As String


'Para los contadores
Private Mc As Contadores
Private FILTRO As Byte
Private NuevaFactura As Boolean

'Por si esta en un periodo liquidado, que pueda modificar CONCEPTO , cuentas,
Private ModificaFacturaPeriodoLiquidado As Boolean

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then cmdAceptar_Click
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Function ActualizaFactura() As Boolean
Dim B As Boolean
On Error GoTo EActualiza
ActualizaFactura = False

B = ModificaDesdeFormularioClaves(Me, SQL)
If Not B Then Exit Function

'Las lineas
If Not Adodc1.Recordset.EOF Then
    SQL = "UPDATE linfact SET numserie='" & Text1(1).Text & "'"
    SQL = SQL & " ,codfaccl = " & Text1(2).Text
    SQL = SQL & " ,anofaccl = " & Text1(26).Text
    SQL = SQL & " WHERE numserie='" & Data1.Recordset!NumSerie
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
    Dim i As Integer
    Dim RC As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
                If InsertarDesdeForm(Me) Then
                    
                    'LOG
                    vLog.Insertar 4, vUsu, Text1(1).Text & Text1(2).Text & " " & Text1(0).Text
                    
                    If SituarData1(1) Then
                        PonerModo 5
                        'Haremos como si pulsamo el boton de insertar nuevas lineas
                        'Ponemos el importe en AUX
                        AUx = ImporteFormateado(Text2(4).Text)
                        
                        '--
                        'cmdCancelar.Caption = "Cabecera"
                        
                        
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
            If DatosOK Then
                '-----------------------------------------
                'Hay que comprobar si ha modificado, o no la clave de la factura
                i = 1
                If Data1.Recordset!NumSerie = Text1(1).Text Then
                    If Data1.Recordset!codfaccl = Text1(2).Text Then
                        If Data1.Recordset!anofaccl = Text1(26).Text Then
                            i = 0
                            'NO HA MODIFICADO NADA
                        End If
                    End If
                End If
            
                'Hacemos MODIFICAR
                If i <> 0 Then
                    MsgBox "No se puede cambiar campos clave  de la factura.", vbExclamation
                    RC = False
                Else
                    RC = ModificaDesdeFormulario(Me)
                End If
                    
                If RC Then
                    DesBloqueaRegistroForm Me.Text1(1)
                    
                    
                    'LOG
                    vLog.Insertar 5, vUsu, Text1(1).Text & Text1(2).Text & " " & Text1(0).Text
                    'Creo que no hace falta volver a situar el datagrid
                    'If SituarData1(0) Then
                    If True Then
                        SituarADO2
                        lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
                        
                        If Numasien2 > 0 Then
                            If IntegrarFactura Then
                                Text1(4).Text = Numasien2
                                Numasien2 = -1
                                NumDiario = 0
                            End If
                        End If
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
'--
'                cmdAceptar.Visible = False
'--
                
                
                DataGrid1.AllowAddNew = False
                DataGrid1.ReBind
                CargaGrid True
                If ModificandoLineas = 1 Then
                    'Estabamos insertando insertando lineas
                    AnyadirLinea True, False
                    If AUx <> 0 Then PonerFoco txtaux(0)
                Else
                    ModificandoLineas = 0
                    CamposAux False, 0, False
'--
'                    cmdCancelar.Caption = "Cabecera"
'--
'++
                    cmdCancelar_Click
'++
                
                
                
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
    If txtaux(2).Text <> "" Then PonerFoco txtaux(4)
End If
End Sub

Private Sub cmdCancelar_Click()

    On Error Resume Next
    Select Case Modo
    Case 1, 3
        'Contador de facturas
        If Modo = 3 Then
            'Intentetamos devolver el contador
            If Text1(0).Text <> "" Then
                i = FechaCorrecta2(CDate(Text1(0).Text))
                Mc.DevolverContador Mc.TipoContador, i = 0, Mc.Contador
            End If
        End If
        LimpiarCampos
        PonerModo 0
        Set Mc = Nothing
    Case 4
        lblIndicador.Caption = ""
       
        
        Modo = 2   'Para que el lostfocus NO haga nada
        If Numasien2 > 0 Then
            'Ha cancelado. Tendre que situar los campos correctamente
            'Es decir. Anofacl
            Text1(0).Text = Adodc2.Recordset!fecfaccl
            Text1(2).Text = Adodc2.Recordset!codfaccl
            Text1(26).Text = Data1.Recordset!anofaccl
            If Not IntegrarFactura Then
                Modo = 4 'lo pongo por si acaso
                Exit Sub
            End If
        End If
        PonerCampos
        Modo = 4  'Reestablezco el modo para que vuelva a hahacer ponercampos
        DesBloqueaRegistroForm Me.Text1(1)
        PonerModo 2
        'Contador de facturas
        Set Mc = Nothing
        Screen.MousePointer = vbDefault
    Case 5
        CamposAux False, 0, False

        BotonCabecera

'
'        'Si esta insertando/modificando lineas haremos unas cosas u otras
'        DataGrid1.Enabled = True
'        If ModificandoLineas = 0 Then
'            'NUEVO
'            AntiguoText1 = ""
'            If Adodc1.Recordset.EOF Then
'                AntiguoText1 = "La factura no tiene lineas. ¿SEGURO que desea salir?"
'                If MsgBox(AntiguoText1, vbQuestion + vbYesNoCancel) = vbYes Then
'                    AntiguoText1 = ""
'                Else
'                    'Para k no muestre el siguiente punto de error
'                    AntiguoText1 = "###"
'                End If
'            Else
'                'Comprobamos que el total de factura es el de suma
'               ObtenerSigueinteNumeroLinea
'               If AUx <> 0 Then AntiguoText1 = "El importe de lineas no suma el importe facturas: " & Format(AUx, FormatoImporte)
'            End If
'            If AntiguoText1 <> "" Then
'                If AntiguoText1 <> "###" Then MsgBox AntiguoText1, vbExclamation
'                Exit Sub
'            End If
'            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
'            If Adodc2.Recordset Is Nothing Then
'                CargaGrid True
'                If Not SituarData1(0) Then
'                    PonerModo 0
'                Else
'
'                    PonerCampos
'                    PonerModo 2
'                    NuevaFactura = True
'                End If
'            Else
'                EnlazaADOs
'                PonerModo 2
'            End If
'
'            If Modo = 2 Then
'                If Numasien2 > 0 Then
'                    If IntegrarFactura Then
'                        Text1(4).Text = Numasien2
'                        Numasien2 = 0
'                        NumDiario = 0
'                    End If
'                End If
'
'                'Nuevo enero 2009
'                'Ademas del vto, que ya lo hacia, ahora contabilizara la factura
'                If NuevaFactura Then
'                    Screen.MousePointer = vbHourglass
'                    espera 0.2
'                    'Contabilizar la factura autmaticamente
'                    If vParam.ContabilizaFactura Then HacerToolbar1 11, True
'
'                    'Esto ya lo hacia. Insertar vto
'                    If vEmpresa.TieneTesoreria Then HacerToolbar1 12, False
'
'                End If
'            End If
'            DesBloqueaRegistroForm Me.Text1(1)
'            Screen.MousePointer = vbDefault
'        Else
'            If ModificandoLineas = 1 Then
'                 DataGrid1.AllowAddNew = False
'                 If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
'                 DataGrid1.Refresh
'            End If
''--
''            cmdAceptar.Visible = False
''            cmdCancelar.Caption = "Cabeceras"
''--
'
'            ModificandoLineas = 0
'
'        End If
    End Select
End Sub

'PRUEBA   20 DICIEMBRE 2005
'--------------------------
' Despues de insertar, el situardata, refresca TOOOOOODO
' Para evitarlo haremos que carge solo la factura que ha creado
' Ademas. Cuando venga de insertar cargaremos el ADO con solo
' esa factura
'
'
'       Opcion INSERTAR:  0- NADA
'                         1- Insertando


' Cuando modificamos el data1 se mueve de lugar, luego volvemos
' ponerlo en el sitio
' Para ello con find y un SQL lo hacemos
' Buscamos por el codigo, que estara en un text u  otro
' Normalmente el text(0)
Private Function SituarData1(OpcionInsertar As Byte) As Boolean
    On Error GoTo ESituarData1
    
    
    'ESTA PREPARADO. FALTA DESCOMENTAR
    '
    If OpcionInsertar <> 1 Then
        If FILTRO > 1 Then
            'Por si acaso pone la de una año u la de otro
            CadenaConsulta = "Select numserie,codfaccl,anofaccl from cabfact where fecfaccl>='" & Format(vParam.fechaini, FormatoFecha) & "' " & Ordenacion
            Data1.RecordSource = CadenaConsulta
        End If
    Else
        'INSERTANDO
        CadenaConsulta = "Select numserie,codfaccl,anofaccl from cabfact where"
        CadenaConsulta = CadenaConsulta & " NUmSerie  = '" & Text1(1).Text
        CadenaConsulta = CadenaConsulta & "' AND anofaccl = " & Text1(26).Text
        CadenaConsulta = CadenaConsulta & " AND codfaccl = " & Text1(2).Text
        Data1.RecordSource = CadenaConsulta
    End If
    espera 0.2
    Data1.Refresh
    'Debug.Print CadenaConsulta
    With Data1.Recordset
        If .EOF Then Exit Function
        .MoveLast
        .MoveFirst
        While Not Data1.Recordset.EOF
            If CStr(.Fields!NumSerie) = Text1(1).Text Then
                If CStr(.Fields!anofaccl) = Text1(26).Text Then
                    If CStr(.Fields!codfaccl) = Val(Text1(2).Text) Then
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
IntegrarFactura = False
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


'Esta "cuadrado"
With frmActualizar
    .OpcionActualizar = 6
    'NumAsiento     --> CODIGO FACTURA
    'NumDiari       --> AÑO FACTURA
    'NUmSerie       --> SERIE DE LA FACTURA
    'FechaAsiento   --> Fecha factura
    .NumFac = CLng(Text1(2).Text)
    .NumDiari = CInt(Text1(26).Text)
    .NumSerie = Text1(1).Text
    .FechaAsiento = Text1(0).Text
    If Numasien2 < 0 Then
        
        If Not Text1(4).Enabled Then
            If Text1(4).Text <> "" Then
                Numasien2 = Text1(4).Text
            End If
        End If
        
    End If
    If NumDiario <= 0 Then NumDiario = vParam.numdiacl
    .DiarioFacturas = NumDiario
    .NumAsiento = Numasien2
    .Show vbModal
    If AlgunAsientoActualizado Then IntegrarFactura = True
    Screen.MousePointer = vbHourglass
    Me.Refresh
End With
If IntegrarFactura Then
    'Data1.Refresh   'Lo he quitado el 20 de diciembre de 2005
    If Modo <> 4 Then
        If Not SituarData1(0) Then
            If TieneRegistros Then
                Data1.Recordset.MoveFirst
                EnlazaADOs
            End If
        End If
    Else
        'MODIFICAR:
        '.Refresco el adodc2
        EnlazaADOs
    End If
    
End If
End Function


'    Data1.Refresh
'    If Not SituarData1(0) Then
'       'If Not Data1.Recordset.EOF Then
'       If TieneRegistros Then Data1.Recordset.MoveFirst
'        Else
'            LimpiarCampos
'            PonerModo 0
'        End If
'    End If
'End If
'End Function

Private Function TieneRegistros() As Boolean
    On Error Resume Next
    TieneRegistros = False
    If Data1.Recordset.RecordCount > 0 Then TieneRegistros = True
End Function



Private Sub BotonAnyadir()
    LimpiarCampos
    
    'MODIFICACIONES 23 DICIEMBRE 2005
    '--------------------------------
    'Situamos el recordset
        SQL = AnyadeCadenaFiltro
        If SQL <> "" Then SQL = " WHERE " & SQL
        CadenaConsulta = "Select numserie,codfaccl,anofaccl from cabfact " & SQL & Ordenacion
        
    CadenaConsulta = "Select numserie,codfaccl,anofaccl from cabfact where numserie='1'"
        
    PonerCadenaBusqueda True

    

    Check1.Value = 0 'Intracomunitaria
    
    'Contador de facturas
    Set Mc = New Contadores
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "&Aceptar"
    PonerModo 3
    NuevaFactura = True
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    CargaGrid False
    
    'Escondemos el navegador y ponemos insertando
    DespalzamientoVisible False
    lblIndicador.Caption = "INSERTANDO"
    
    '###A mano
    Text1(0).Text = Format(Now, "dd/mm/yyyy")
    PonerFoco Text1(0)
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        lblIndicador.Caption = "Búsqueda"
        cmdCancelar.Caption = "Cancelar"
        cmdAceptar.Caption = "&Aceptar"
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        '### A mano
        '------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(kCampo).Text = ""
                Text1(kCampo).BackColor = vbYellow
                PonerFoco Text1(kCampo)
            End If
    End If
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    'Ponemos el grid lineasfacturas enlazando a ningun sitio
    DataGrid1.Enabled = False
    CargaGrid False
    SQL = AnyadeCadenaFiltro
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia SQL
    Else
        If SQL <> "" Then SQL = " WHERE " & SQL
        CadenaConsulta = "Select numserie,codfaccl,anofaccl from cabfact " & SQL & Ordenacion
        PonerCadenaBusqueda False
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
Screen.MousePointer = vbHourglass
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
Me.Refresh
espera 0.2
Screen.MousePointer = vbDefault
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
            SQL = varTxtFec
        Else
            SQL = "La factura pertenece a un ejercicio cerrado."
        End If
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    If Not ComprobarPeriodo2(28) Then Exit Sub

    
    NumDiario = 0
    'Comprobamos que no esta actualizada ya
    If Not IsNull(Adodc2.Recordset!NumAsien) Then
        Numasien2 = Adodc2.Recordset!NumAsien
        If Numasien2 = 0 Then
            MsgBox "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
            Exit Sub
        End If
            
        SQL = "Esta factura ya esta contabilizada. Desea desactualizar para poder modificarla?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Numasien2 = Adodc2.Recordset!NumAsien
        NumDiario = Adodc2.Recordset!NumDiari
    Else
        Numasien2 = -1
    End If
        
        
    'Si viene a esta factura buscando por un campo k no sea clave entonces no le dejo seguir
    If InStr(1, Data1.Recordset.Source, "numasien") Then
        MsgBox "Busque la factura por su numero de factura", vbExclamation
        Numasien2 = -1
        Exit Sub
    End If
    
    'Llegados aqui bloqueamos desde form
    If Not BloqueaRegistroForm(Me) Then Exit Sub
    If Numasien2 >= 0 Then
        'Tengo desintegrar la factura del hco
        If Not Desintegrar Then
            DesBloqueaRegistroForm Me.Text1(1)
            Exit Sub
        End If
        Text1(4).Text = ""
    End If
    
    If Mc Is Nothing Then Set Mc = New Contadores
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Modificar"
    cmdCancelar.Caption = "Cancelar"
    PonerModo 4
    DespalzamientoVisible False
    lblIndicador.Caption = "Modificar"
    PonerFoco Text1(5)
End Sub

Private Sub BotonEliminar()
    Dim i As Long
    Dim Fec As Date
    Dim Mc As Contadores

    'Ciertas comprobaciones
    If Data1.Recordset Is Nothing Then Exit Sub
    If Data1.Recordset.EOF Then Exit Sub
    If Adodc2.Recordset.EOF Then Exit Sub
    DataGrid1.Enabled = False
    
    'Comprobamos si esta liquidado
    If Not ComprobarPeriodo2(28) Then Exit Sub
    
    'Comprobamos que no esta actualizada ya
    SQL = ""
    If Not IsNull(Adodc2.Recordset!NumAsien) Then
        SQL = "Esta factura ya esta contabilizada. "
    End If
    
    SQL = SQL & vbCrLf & vbCrLf & "Va usted a eliminar la factura :" & vbCrLf
    SQL = SQL & "Numero : " & Adodc2.Recordset!codfaccl & vbCrLf
    SQL = SQL & "Fecha  : " & Adodc2.Recordset!fecfaccl & vbCrLf
    SQL = SQL & "Cliente : " & Adodc2.Recordset!codmacta & " - " & Text4(0).Text & vbCrLf
    SQL = SQL & vbCrLf & "          ¿Desea continuar ?" & vbCrLf
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    NumRegElim = Data1.Recordset.AbsolutePosition
    Screen.MousePointer = vbHourglass
    'Lo hara en actualizar
    i = 0
    If Not IsNull(Adodc2.Recordset!NumAsien) Then i = Adodc2.Recordset!NumAsien
    If i > 0 Then
        
            'Memorizamos el numero de asiento y la fechaent para ver si devolvemos el contador
            'de asientos
            i = Adodc2.Recordset!NumAsien
            Fec = Adodc2.Recordset!FechaEnt
        
            'La borrara desde actualizar
            AlgunAsientoActualizado = False
            With frmActualizar
                .OpcionActualizar = 7
                .NumAsiento = Adodc2.Recordset!NumAsien
                .NumFac = Adodc2.Recordset!codfaccl
                .FechaAsiento = Adodc2.Recordset!fecfaccl
                .NumSerie = Adodc2.Recordset!NumSerie & "|" & Adodc2.Recordset!anofaccl & "|"
                .NumDiari = Adodc2.Recordset!NumDiari
                .Show vbModal
            End With
            Set Mc = New Contadores
            Mc.DevolverContador "0", Fec <= vParam.fechafin, i
            Set Mc = Nothing
        
    Else
        'La borrara desde este mismo form
        Conn.BeginTrans
        
        i = Adodc2.Recordset!codfaccl
        Fec = Adodc2.Recordset!fecfaccl
        If BorrarFactura Then
            'LOG
            vLog.Insertar 6, vUsu, CStr(DBLet(Adodc2.Recordset!NumSerie)) & Format(i, "000000")
            
        
            AlgunAsientoActualizado = True
            Conn.CommitTrans
            Set Mc = New Contadores
            Mc.DevolverContador CStr(DBLet(Adodc2.Recordset!NumSerie)), (Fec <= vParam.fechafin), i
            Set Mc = Nothing
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
        CargaGrid False
        PonerModo 0
        Else
            Data1.Recordset.MoveFirst
            NumRegElim = NumRegElim - 1
            If NumRegElim > 1 Then
                For i = 1 To NumRegElim - 1
                    Data1.Recordset.MoveNext
                Next i
            End If
            PonerCampos
    End If

Error2:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            Data1.Recordset.CancelUpdate
        End If
End Sub


Private Function BorrarFactura() As Boolean
    
    On Error GoTo EBorrar
    SQL = " WHERE numserie = '" & Data1.Recordset!NumSerie & "'"
    SQL = SQL & " AND codfaccl = " & Data1.Recordset!codfaccl
    SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
    'Las lineas
    AntiguoText1 = "DELETE from linfact " & SQL
    Conn.Execute AntiguoText1
    'La factura
    AntiguoText1 = "DELETE from cabfact " & SQL
    Conn.Execute AntiguoText1
    
    ComprobarContador Data1.Recordset!NumSerie, CDate(Text1(0).Text), Data1.Recordset!codfaccl
    
EBorrar:
    If Err.Number = 0 Then
        BorrarFactura = True
    Else
        MuestraError Err.Number, "Eliminar factura"
        BorrarFactura = False
    End If
End Function

Private Sub cmdRegresar_Click()
'Dim Cad As String
'Dim I As Integer
'Dim J As Integer
'Dim Aux As String

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

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++

Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False

        PonerModo CInt(Modo)
        CargaGrid (Modo = 2)
        If Modo <> 2 Then
            'CadenaConsulta = "Select * from cabfact " & Ordenacion
           ' Data1.RecordSource = CadenaConsulta
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    LimpiarCampos
    PrimeraVez = True
    PulsadoSalir = False
    
    LeerFiltro True
    PonerFiltro FILTRO
    Label4.Tag = ""
    
    
    Caption = "Registro facturas clientes (" & vEmpresa.nomresum & ")"
    
    
    'Si mostramos fecha liquidacion o no
    FrameTapa.Visible = Not vParam.Constructoras
    Text1(28).Enabled = vParam.Constructoras
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 17
        .Buttons(13).Image = 16
        .Buttons(14).Image = 15
        .Buttons(16).Image = 6
        .Buttons(17).Image = 7
        .Buttons(18).Image = 8
        .Buttons(19).Image = 9
        
        
        
        'Si tiene tesoreria entonces
        If vEmpresa.TieneTesoreria Then
            .Buttons(12).Style = tbrDefault
            .Buttons(12).ToolTipText = "Generar vencimientos"
            .Buttons(12).Image = 25
        Else
            .Buttons(12).Style = tbrSeparator
        End If
        
        
    End With
    
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For kCampo = 0 To 0
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next kCampo
   ' ***********************************
    
        
    
    
    If Screen.Width > 12000 Then
        Top = 400
        Left = 400
    Else
        Top = 0
        Left = 0
  '      Me.Width = 12000
  '      Me.Height = Screen.Height
    End If
    Me.Height = 8610
    'Los campos auxiliares
    CamposAux False, 0, True
    
    
    '## A mano
    Ordenacion = " ORDER BY numserie,fecfaccl,codfaccl"
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
'    Data1.UserName = vUsu.Login
'    Data1.password = vUsu.Passwd
'    Adodc1.password = vUsu.Passwd
'    Adodc1.UserName = vUsu.Login
    
    '++
    CargaGrid False
    '++
    
    PonerOpcionesMenu
    
    'Maxima longitud cuentas
    txtaux(0).MaxLength = vEmpresa.DigitosUltimoNivel
    'Bloqueo de tabla, cursor type
    Data1.CursorType = adOpenDynamic
    Data1.LockType = adLockPessimistic
    Data1.RecordSource = "Select numserie,codfaccl,anofaccl from Cabfact WHERE numserie ='David'"
    Data1.Refresh
    CadAncho = False
    Modo = 0
End Sub



Private Sub LimpiarCampos()
    Limpiar Me   'Metodo general
    lblIndicador.Caption = ""
    NuevaFactura = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Modo > 2 Then
        If Not PulsadoSalir Then
            Cancel = 1
            Exit Sub
        End If
    End If
    LeerFiltro False
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    
 
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim AUx As String
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        AUx = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
        CadB = AUx
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        AUx = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 2)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & AUx
        
        AUx = ValorDevueltoFormGrid(Text1(26), CadenaDevuelta, 3)
        If CadB <> "" Then CadB = CadB & " AND "
        CadB = CadB & AUx
        'Se muestran en el mismo form
        CadenaConsulta = "select * from cabfact WHERE " & CadB & " "
        PonerCadenaBusqueda False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas
    SQL = RecuperaValor(CadenaSeleccion, 3)
    If SQL <> "" Then
        'Cuenta bloqueada
        If Text1(0).Text <> "" Then 'Hay fecha
            SQL = RecuperaValor(CadenaSeleccion, 1)
            If EstaLaCuentaBloqueada(SQL, CDate(Text1(0).Text)) Then
                MsgBox "Cuenta bloqueada: " & SQL, vbExclamation
                Exit Sub
            End If
        End If
    End If

    Select Case cmdAux(0).Tag
    Case 2, 5
        'Cuenta normal
        Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(0).Text = RecuperaValor(CadenaSeleccion, 2)
    Case 3, 6
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
'[Monica]--
'    If Text1(0).Text = "" Then
'         MsgBox "No hay fecha seleccionada ", vbExclamation
'         Exit Sub
'    End If
    SQL = RecuperaValor(CadenaSeleccion, 1)
'[Monica]--
'    B = CDate(Text1(0).Text) <= vParam.fechafin
    If Mc Is Nothing Then Set Mc = New Contadores
    If Mc.ConseguirContador(SQL, B, False) = 0 Then
        Text1(1).Text = SQL
'[Monica]--
'        Text1(2).Text = Mc.Contador
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text1(CInt(cmdAux(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmI_DatoSeleccionado(CadenaSeleccion As String)
    'Solo me interesa el codigo
    i = CInt(AUx - 2)
    Text1(((i) * 6) + 1).Text = RecuperaValor(CadenaSeleccion, 1)
    If PonerValoresIva(i) Then
        CalcularIVA i
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
        If Text1(1).Text <> "" Then PonerFoco Text1(3)
    Case 1, 7
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If Index = 1 Then
            i = 0
        Else
            i = 28
        End If
        cmdAux(0).Tag = i
        If Text1(i).Text <> "" Then frmF.Fecha = CDate(Text1(i).Text)
        frmF.Show vbModal
        Set frmF = Nothing
    Case 2, 6, 100
        cmdAux(0).Tag = Index
        'Cliente y cta retencion
        Set frmC = New frmColCtas
        frmC.DatosADevolverBusqueda = "0|1|"
        frmC.ConfigurarBalances = 3  'NUEVO
        frmC.Show vbModal
        Set frmC = Nothing
        'Lo vuelvo a posicionar ande toca
        If Index = 100 Then txtAux_LostFocus 0
        
    Case 3, 4, 5
        AUx = Index
        Set frmI = New frmIVA
        frmI.DatosADevolverBusqueda = "0|1|"
        frmI.Show vbModal
        Set frmI = Nothing
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Label4_DblClick()
    If Label4.Tag <> "" Then
        If Text1(4).Text = "" Then
            Label4.Tag = InputBox("NA:")
            If Val(Label4.Tag) > 0 Then Text1(4).Text = Val(Label4.Tag)
        End If
        Label4.Tag = ""
    End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Shift = 1 Then
            Label4.Tag = "OK"
        End If
    End If
End Sub

Private Sub mnActual_Click()
    PonerFiltro 2
End Sub

Private Sub mnActuralySiguiente_Click()
    PonerFiltro 1
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
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnSiguiente_Click()
    PonerFiltro 3
End Sub

Private Sub mnSinFiltro_Click()
    PonerFiltro 0
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
    Else
        If (Shift And vbCtrlMask) > 0 Then
            If UCase(Chr(KeyCode)) = "B" Then
                'OK. Ha pulsado Control + B
                LanzaPantalla Index
            End If
        End If
            
    End If
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
        Select Case Index
        Case 5
            Text1(5).Text = ""
            miI = 2
        
        Case 3
            Text1(3).Text = ""
            miI = 6
        Case 7, 13, 19
            Text1(Index).Text = ""
            If Index = 7 Then
                miI = 3
            Else
                If Index = 13 Then
                    miI = 4
                Else
                    miI = 5
                End If
            End If
        End Select
        imgppal_Click miI
End Sub


'++
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYBusqueda KeyAscii, 1 'fecha
            Case 1:  KEYBusqueda KeyAscii, 0 'contador
            Case 3:  KEYBusqueda KeyAscii, 6 'cta
            Case 5:  KEYBusqueda KeyAscii, 2 'cta
            Case 7:  KEYBusqueda KeyAscii, 3 'iva
            Case 13:  KEYBusqueda KeyAscii, 4 'iva
            Case 19:  KEYBusqueda KeyAscii, 5 'iva
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgppal_Click (indice)
End Sub
'++





Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Modo <> 1 Then
        If KeyCode = 107 Or KeyCode = 187 Then
                KeyCode = 0
                LanzaPantalla Index
        End If
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
    
    
    'Ahora, si no ha pueto Base2 lo pasamos a retencion
    'o si no pone retencion lo pasamos a boton aceptar
    If Index = 12 Then
        If Text1(12).Text = "" Then PonerFoco Text1(24)
    Else
        If Index = 24 Then
            If Text1(24).Text = "" Then PonerFoco cmdAceptar
        End If
    End If
    
    
Else
    If Modo = 1 Then
        If Index = 5 Or Index = 3 Then PierdeFoco3 Index
    End If
End If
End Sub


'Para cuando piede foco y estamos insertando o modificando
Private Sub PierdeFoco3(indice As Integer)
Dim RC As String
Dim Correcto As Boolean
Dim Valor As Currency
Dim L As Long
Dim i As Integer
Dim J As Integer

    Text1(indice).Text = Trim(Text1(indice).Text)
    If Text1(indice).Text = "" Then
        'Hemos puesto a blancos el campo, luego quitaremos
        'los valores asociados a el
        If Text1(indice) = AntiguoText1 Then Exit Sub
        Select Case indice
        Case 0
            'Ponemos a blanco tb el año de factura
            Text1(26).Text = ""
        Case 1
            'Ha puesto a blanco la serie de las facturas
            'por lo tanto habra que mirar si es el ultimo
            If Text1(0).Text <> "" Then
                Correcto = CDate(Text1(0).Text) <= vParam.fechafin
                If Text1(2).Text <> "" Then
                    Linfac = Val(Text1(2).Text)
                    Mc.DevolverContador AntiguoText1, Correcto, Linfac
                End If
            End If
        Case 6 To 23
            
            If indice < 12 Then
                'PRIMERA LINEA
                L = 1
                'Numero de campo k ocupa
                i = indice - 6
            Else
                J = indice - 6
                L = (J \ 6) + 1
                i = indice - (L * 6)
                
            End If
            
            'Ponemos IVA
            If i = 1 Then
                'Ha puesto a blanco el IVA. Borarmos el resto de campos
                J = (L * 6) + 5
                Text4(L).Text = ""
                For J = indice To J
                    Text1(J).Text = ""
                Next J
            End If
            'Ha cambiado la base o el iva. Luego hay k recalcular valores
            If i < 2 Then CalcularIVA CInt(L)
            TotalesRecargo
            TotalesIVA
            TotalesBase
            TotalFactura
        Case 3
            Text4(4).Text = ""
        Case 5
            Text4(0).Text = ""
        Case 24
            Text2(3).Text = ""
            TotalFactura
        End Select
    Else
        With Text1(indice)
           Select Case indice
           Case 0, 28
                If Not EsFechaOK(Text1(indice)) Then
                    MsgBox "Fecha incorrecta: " & .Text, vbExclamation
                    .Text = ""
                    If indice = 0 Then Text1(26).Text = ""
                    PonerFoco Text1(indice)
                    Exit Sub
                End If
                
                'NUEVO!!!!
                'Hay que comprobar que las fechas estan
                'en los ejercicios y si
                '       0 .- Año actual
                '       1 .- Siguiente
                '       2 .- Ambito                !!!!  NUEVO !!!
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
                    If indice = 0 Then Text1(26).Text = ""
                    PonerFoco Text1(indice)
                    Exit Sub
                End If
                
                
                .Text = Format(.Text, "dd/mm/yyyy")
                If indice = 0 Then Text1(26).Text = Year(CDate(.Text))
                
                'Si que pertenece a ejerccios en curso. Por lo tanto comprobaremos
                'que el periodo de liquidacion del IVA no ha pasado.
                i = 0
                If vParam.Constructoras Then
                    If indice = 28 Then i = 1
                Else
                    If indice = 0 Then i = 1
                End If
                If i > 0 Then
                    If Not ComprobarPeriodo2(indice) Then PonerFoco Text1(indice)
                End If
                
            Case 1
                 If IsNumeric(.Text) Then
                    MsgBox "Debe ser una letra: " & .Text, vbExclamation
                    .Text = ""
                    PonerFoco Text1(1)
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
                    If Mc.ConseguirContador(.Text, Correcto, False) = 0 Then
                        Text1(2).Text = Mc.Contador
                    Else
                        MsgBox "La letra no es de contadores: " & .Text, vbExclamation
                        .Text = ""
                        Text1(2).Text = ""
                        PonerFoco Text1(1)
                    End If
                End If

            Case 2
                If Not IsNumeric(.Text) Then
                    MsgBox "El numero de factura no es correcto: " & .Text, vbExclamation
                    .Text = ""
                    PonerFoco Text1(2)
                End If
            Case 3, 5
                'Cuenta cliente
                If AntiguoText1 = .Text Then Exit Sub
                RC = .Text
                If indice = 3 Then
                    i = 4
                    Else
                    i = 0
                End If
                
                If CuentaCorrectaUltimoNivel(RC, SQL) Then
                    .Text = RC
                    Text4(i).Text = SQL
                    If Text1(0).Text <> "" Then
                        If Modo > 2 Then
                            If EstaLaCuentaBloqueada(RC, CDate(Text1(0).Text)) Then
                                MsgBox "Cuenta bloqueada: " & RC, vbExclamation
                                .Text = ""
                                Text4(i).Text = ""
                            End If
                        End If
                    End If
                    RC = ""
                Else
                    
                    If InStr(1, SQL, "No existe la cuenta :") > 0 Then
                        'NO EXISTE LA CUENTA
                            RC = RellenaCodigoCuenta(Text1(indice).Text)
                            SQL = "La cuenta: " & RC & " no existe. ¿Desea crearla?"
                            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                                CadenaDesdeOtroForm = RC
                                cmdAux(0).Tag = indice
                                Set frmC = New frmColCtas
                                frmC.DatosADevolverBusqueda = "0|1|"
                                frmC.ConfigurarBalances = 4   ' .- Nueva opcion de insertar cuenta
                                frmC.Show vbModal
                                Set frmC = Nothing
                                If Text1(5).Text = RC Then SQL = "" 'Para k no los borre
                            End If
                    Else
                        'Cualquier otro error
                        'menos si no estamos buscando, k dejaremos
                        If Modo = 1 Then
                            SQL = ""
                        Else
                            MsgBox SQL, vbExclamation
                        End If
                    End If
                    
                    If SQL <> "" Then
                        .Text = ""
                        Text4(i).Text = ""
                        PonerFoco Text1(indice)
                    End If
                    
                    
                End If
                
            Case 7, 13, 19  'TIpos de iva
                i = ((indice - 1) \ 6)
                If Not IsNumeric(.Text) Then
                
                    MsgBox "Tipo de iva " & i & " incorrecto:  " & .Text
                    .Text = ""
                    Text4(i).Text = ""
                    PonerFoco Text1(indice)
                    Exit Sub
                End If
                If .Text = AntiguoText1 Then Exit Sub
                If PonerValoresIva(i) Then
                    CalcularIVA i
                    TotalesRecargo
                    TotalesIVA
                    TotalesBase
                    TotalFactura
                End If
            Case 6, 12, 18
                'BASES IMPONIBLES
                Correcto = True
                i = ((indice) \ 6)
                
                If Not EsNumerico(.Text) Then
                    MsgBox "Importe debe de ser numérico: " & .Text, vbExclamation
                    .Text = ""
                    Correcto = False
                Else
                    If InStr(1, .Text, ",") > 0 Then
                        Valor = ImporteFormateado(.Text)
                    Else
                        Valor = CCur(TransformaPuntosComas(.Text))
                    End If
                    .Text = Format(Valor, FormatoImporte)
                    If .Text = AntiguoText1 Then Exit Sub
                End If
                'Recalculamos iva
                CalcularIVA i
                TotalesRecargo
                TotalesIVA
                TotalesBase
                TotalFactura
                If Not Correcto Then PonerFoco Text1(indice)
            Case 9, 15, 21
                Correcto = True
                i = ((indice - 3) \ 6)
                
                If Not EsNumerico(.Text) Then
                    MsgBox "Importe debe de ser numérico: " & .Text, vbExclamation
                    .Text = ""
                    Correcto = False
                Else
                    If InStr(1, .Text, ",") > 0 Then
                        Valor = ImporteFormateado(.Text)
                    Else
                        Valor = CCur(TransformaPuntosComas(.Text))
                    End If
                    .Text = Format(Valor, FormatoImporte)
                    If .Text = AntiguoText1 Then Exit Sub
                End If
                TotalesRecargo
                TotalesIVA
                TotalesBase
                TotalFactura
                If Not Correcto Then PonerFoco Text1(indice)
            Case 11, 17, 23
                Correcto = True
                i = ((indice - 5) \ 6)
                If Not EsNumerico(.Text) Then
                    MsgBox "Importe debe de ser numérico: " & .Text, vbExclamation
                    .Text = ""
                    Correcto = False
                Else
                    If InStr(1, .Text, ",") > 0 Then
                        Valor = ImporteFormateado(.Text)
                    Else
                        Valor = CCur(TransformaPuntosComas(.Text))
                    End If
                    .Text = Format(Valor, FormatoImporte)
                    If .Text = AntiguoText1 Then Exit Sub
                End If
                TotalesRecargo
                TotalesIVA
                TotalesBase
                TotalFactura
                If Not Correcto Then PonerFoco Text1(indice)
                
                
            Case 24
                If Not EsNumerico(.Text) Then
                    MsgBox "% de recargo debe de ser numérico: " & .Text, vbExclamation
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
    
    SQL = AnyadeCadenaFiltro
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
        Else
            'Se muestran en el mismo form
            If CadB <> "" Then
                If SQL <> "" Then SQL = " AND (" & SQL & ")"
                CadB = CadB & SQL
                CadenaConsulta = "select numserie,codfaccl,anofaccl from cabfact WHERE " & CadB & " " & Ordenacion
                PonerCadenaBusqueda False
            End If
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
        Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(1), 10, "Serie: ")
        Cad = Cad & ParaGrid(Text1(2), 17, "Nº Fac.")
        Cad = Cad & ParaGrid(Text1(26), 10, "Año")
        Cad = Cad & ParaGrid(Text1(0), 18)
        Cad = Cad & ParaGrid(Text1(27), 20)
        Cad = Cad & ParaGrid(Text1(5), 25)
        
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = "cabfact"
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
               PonerFoco Text1(kCampo)
            End If
        End If
End Sub

Private Sub PonerCadenaBusqueda(Insertando As Boolean)
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Insertando Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Data1.Recordset.EOF Then
        MsgBox "No hay ningún registro en la tabla Facturas clientes con estos parámetros.", vbInformation
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


Private Function EnlazaADOs() As Boolean
 On Error GoTo EEnalaza
 
    EnlazaADOs = False

   
    SQL = "Select * from cabfact where codfaccl=" & Data1.Recordset!codfaccl
    SQL = SQL & " AND anofaccl = " & Data1.Recordset!anofaccl
    SQL = SQL & " AND numserie= '" & Data1.Recordset!NumSerie & "'"
    Adodc2.ConnectionString = Conn
    Adodc2.RecordSource = SQL
    Adodc2.Refresh
    EnlazaADOs = True
    Exit Function
EEnalaza:
    MuestraError Err.Number, "Enlazando datos con BD's"
    SituarData1 (0)
End Function



Private Sub PonerCampos()
    Dim mTag As CTag
    Dim SQL As String
    
    On Error GoTo EPonerCampos
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not EnlazaADOs Then Exit Sub
    
    PonerCamposForma Me, Adodc2
    
    'Por si modifica factura
    Numasien2 = -1
    NumDiario = 0
    NuevaFactura = False
    
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
    If SQL <> CStr(AUx) Then
        MsgBox "Importe factura distinto Importe calculado: " & SQL & " - " & CStr(AUx), vbExclamation
    End If
    
    'Cliente
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(5).Text, "T")
    Text4(0).Text = SQL
    
    'IVAS
    For i = 1 To 3
        kCampo = (i * 6) + 1
        If Text1(kCampo).Text <> "" Then
            SQL = DevuelveDesdeBD("nombriva", "tiposiva", "codigiva", Text1(kCampo).Text, "N")
        Else
            SQL = ""
        End If
        Text4(i).Text = SQL
    Next i
    
    'Retencion
    If Text1(3).Text <> "" Then
        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Text1(3).Text, "T")
    Else
        SQL = ""
    End If
    Text4(4).Text = SQL
        
        
    If Modo = 2 Then lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    '++
    PonerOpcionesMenu
    '++
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner datos" & vbCrLf & Err.Description
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
Private Sub PonerModo(Kmodo As Integer)
    Dim B As Boolean
    
    If Modo = 1 Or Modo = 4 Then 'ntiguos modo 1 o modificar
        'Reestablecer colores
        For i = 0 To Text1.Count - 1
            Text1(i).BackColor = vbWhite
            Text1(i).Enabled = True
        Next i
        For i = 3 To 5
            imgppal(i).Enabled = True
        Next i
        imgppal(0).Enabled = True
    End If
    Text1(4).Enabled = (Kmodo = 1)
        
'--
'    If Modo = 5 And Kmodo <> 5 Then
'        'El modo antigu era modificando las lineas
'        'Luego hay que reestablecer los dibujitos y los TIPS
'        '-- insertar
'        Toolbar1.Buttons(6).Image = 3
'        Toolbar1.Buttons(6).ToolTipText = "Nuevo factura"
'        '-- Modificar
'        Toolbar1.Buttons(7).Image = 4
'        Toolbar1.Buttons(7).ToolTipText = "Modificar factura"
'        '-- eliminar
'        Toolbar1.Buttons(8).Image = 5
'        Toolbar1.Buttons(8).ToolTipText = "Eliminar factura"
'    End If
'--
    
    'ASIGNAR MODO
    Modo = Kmodo
'++
    PonerIndicador lblIndicador, Modo, ModificandoLineas
'++

'--
'    If Modo = 5 Then
'        'Ponemos nuevos dibujitos y tal y tal
'        'Luego hay que reestablecer los dibujitos y los TIPS
'        '-- insertar
'        Toolbar1.Buttons(6).Image = 12
'        Toolbar1.Buttons(6).ToolTipText = "Nueva linea factura"
'        '-- Modificar
'        Toolbar1.Buttons(7).Image = 13
'        Toolbar1.Buttons(7).ToolTipText = "Modificar linea factura"
'        '-- eliminar
'        Toolbar1.Buttons(8).Image = 14
'        Toolbar1.Buttons(8).ToolTipText = "Eliminar linea factura"
'    End If
'--

    B = (Modo < 5)
    chkVistaPrevia.Visible = B

    'Modo 2. Hay datos y estamos visualizandolos
    B = (Kmodo = 2)
    DespalzamientoVisible B
    Toolbar1.Buttons(10).Enabled = B  'Lineas factur
    Toolbar1.Buttons(11).Enabled = B

'++
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.Visible = B
    cmdAceptar.Visible = B
'++


'--
'    B = B Or (Modo = 5)
'    DataGrid1.Enabled = B
'--
    
    
    'Modo insertar o modificar
    B = (Modo = 3) Or (Modo = 4) '-->Luego not b sera kmodo<3
    Toolbar1.Buttons(6).Enabled = Not B
    mnNuevo.Enabled = Not B
'--
'    cmdAceptar.Visible = B Or Modo = 1
'--
    
    Me.framecabeceras.Enabled = B Or Modo = 1
    'Si es modiifcar y de periodo CERRADO
    If Modo = 4 Then
        If ModificaFacturaPeriodoLiquidado Then HabilitarTXTCabecerasAlModificar True
    End If
    

    '
    B = B Or (Modo = 5)
    Toolbar1.Buttons(1).Enabled = Not B
    Toolbar1.Buttons(2).Enabled = Not B
    mnOpcionesAsiPre.Enabled = Not B
   
    'El boton de vto sera enable si
    If vEmpresa.TieneTesoreria Then
        Toolbar1.Buttons(12).Enabled = (Modo = 2) And vUsu.Nivel < 3
    End If
    'El text
    B = (Modo = 2) '--Or (Modo = 5)
    Toolbar1.Buttons(7).Enabled = B
    mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(8).Enabled = B
    mnEliminar.Enabled = B
    
    'Busquedas por importe. Ponemos el tag momentaneamente
    If Modo = 1 Then
        Text2(4).Tag = "Importe|N|N|||cabfact|totfaccl|#,##0.00||"
    Else
        Text2(4).Tag = ""
    End If
    Text2(4).Locked = Not (Modo = 1)
   
   
    If Kmodo = 0 Then lblIndicador.Caption = ""
    
    '### A mano
    'Aqui añadiremos controles para datos especificos. Esto es, si hay imagenes en el form
    ' o cualquier objeto que dependiendo en el modo en el que esteos se visualizaran o no
    ' Bloqueamos los campos de texto y demas controles en funcion
    ' del modo en el que estamos.
    ' Es decir, si estamos en modo busqueda, insercion o modificacion estaran enables
    ' si no  disable. la variable b nos devuelve esas opciones
'--
'    B = Modo > 2 Or Modo = 1
'    cmdCancelar.Visible = B
'--
    'Detalles
    'DataGrid1.Enabled = Modo = 5
    
    
    
    '--PonerOpcionesMenuGeneral Me
    '++
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    '++
End Sub


Private Function DatosOK() As Boolean
'    Dim RS As ADODB.Recordset
    Dim B As Boolean
    Dim ii As Integer
    
    'Si no es constructoras igualamos los campos fecfac y fecliquidacion
    If Not vParam.Constructoras Then Text1(28).Text = Text1(0).Text
    
    B = CompForm(Me)
    If Not B Then Exit Function
    
   
   'No pude tener Base imponible sin IVA
   If ((Text1(6).Text = "") Xor (Text1(7).Text = "")) Then
        B = False
   Else
        'Tampoco puede tener datos del IVA sin el tipo De IVA
        If Text1(7).Text = "" Then
            'Ningun campo puede estar puesto
            If ((Text1(9).Text <> "") Or (Text1(10).Text <> "") Or (Text1(11).Text <> "")) Then
                MsgBox "Datos de IVA (1) sin poner el tipo", vbExclamation
                Exit Function
            End If
                
        End If
   End If
   If ((Text1(12).Text = "") Xor (Text1(13).Text = "")) Then
        B = False
    Else
    
            If Text1(13).Text = "" Then
                'Ningun campo puede estar puesto
                If ((Text1(15).Text <> "") Or (Text1(16).Text <> "") Or (Text1(17).Text <> "")) Then
                    MsgBox "Datos de IVA (2) sin poner el tipo", vbExclamation
                    Exit Function
                End If
            End If
    End If
    
   If ((Text1(18).Text = "") Xor (Text1(19).Text = "")) Then
        B = False
    Else
            If Text1(19).Text = "" Then
                'Ningun campo puede estar puesto
                If ((Text1(21).Text <> "") Or (Text1(22).Text <> "") Or (Text1(23).Text <> "")) Then
                    MsgBox "Datos de IVA (3) sin poner el tipo", vbExclamation
                    Exit Function
                End If
                    
            End If
    End If
    
   If Not B Then
        MsgBox "No puede tener base imponible sin IVA.", vbExclamation
        Exit Function
    End If
   
    'No puede tener % de retencion sin cuenta de retencion
    If ((Text1(24).Text = "") Xor (Text1(3).Text = "")) Then
        MsgBox "No hay porcentaje de rentención sin cuenta de retención", vbExclamation
        B = False
        Exit Function
    End If
    
    'Compruebo si hay fechas bloqueadas
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
    
    
    'Ahora. Si estamos modificando, y el año factura NO es el mismo, entonces
    'la estamos liando, y para evitar lios, NO dejo este tipo de modificacion
    If Modo = 4 Then
        If CDate(Text1(0).Text) <> Adodc2.Recordset!fecfaccl Then
            'HAN CAMBIADO LA FECHA. Veremos si dejo
            If Year(CDate(Text1(0).Text)) <> Adodc2.Recordset!anofaccl Then
                MsgBox "No puede cambiar de año la factura. ", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    
    DatosOK = B
End Function



Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolbar1 Button.Index, False
End Sub

Private Sub HacerToolbar1(Indi As Integer, ContabilizarDespuesInsertar As Boolean)
Dim N As Long
    Select Case Indi
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
            'Modificar lineas
            If FechaCorrecta2(CDate(Text1(0).Text)) >= 2 Then
                MsgBox "La factura pertenece a un ejercicio cerrado o fuera de ambito.", vbExclamation
                Exit Sub
            End If
            
            
            If Not ComprobarPeriodo2(28) Then Exit Sub
            
            If BloqueaRegistroForm(Me) Then BotonEliminar
            
        Else
            'ELIMINAR linea factura
            EliminarLineaFactura
        End If
    Case 10
        If FechaCorrecta2(CDate(Text1(0).Text)) >= 2 Then
            MsgBox "La factura pertenece a un ejercicio cerrado o fuera de ambito.", vbExclamation
            Exit Sub
        End If
    
        If Numasien2 > 0 Then
            'Viene, de modificar cabecera factura. Tenemos k volvera enlazar
            CargaGrid True
            espera 0.1
        End If
    
    
        If Not ComprobarPeriodo2(28) Then Exit Sub
    
        If Not IsNull(Adodc2.Recordset!NumAsien) Then
            N = Adodc2.Recordset!NumAsien
            If N = 0 Then
                MsgBox "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
                Exit Sub
            End If
            SQL = "Esta factura ya esta contabilizada. Desea desactualizar para poder modificarla?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            Numasien2 = Adodc2.Recordset!NumAsien
            NumDiario = Adodc2.Recordset!NumDiari
            'Tengo desintegre la factura del hco
            If Not Desintegrar Then Exit Sub
            ObtenerSigueinteNumeroLinea
            Text1(4).Text = ""
        End If
            
        
   
        'If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
        If Not BloqueaRegistroForm(Me) Then Exit Sub
        
        'Fuerzo que se vean las lineas
        cmdCancelar.Caption = "Cabecera"
        lblIndicador.Caption = "Lineas detalle"
        
        
        PonerModo 5
        ModificandoLineas = 0
        'Si tiene numasien es k kiere modificar algo, luego se lo sugiero
        If Numasien2 > 0 Then
            If Adodc1.Recordset.RecordCount = 1 Then ModificarLinea
        End If
        
        
    Case 11
        If Data1.Recordset.EOF Then Exit Sub
        If Adodc2.Recordset.EOF Then Exit Sub
        If Not IsNull(Adodc2.Recordset!NumAsien) Then
            MsgBox "La factura ya esta contabilizada.", vbExclamation
            Exit Sub
        End If
        
        
        If FechaCorrecta2(CDate(Text1(0).Text)) >= 2 Then
            MsgBox "No se puede contabilizar con esta fecha.", vbExclamation
            Exit Sub
        End If
        
        
        If FacturaContabilizada(Text1(1).Text, Text1(2).Text, Text1(26).Text) Then
            MsgBox "Factura ya contabilizada(Step: 2). ", vbExclamation
            PonerCampos
            Exit Sub
        End If
        
        If Not ContabilizarDespuesInsertar Then
            'Si ha pusado el boton entoces hago la pregunta
            SQL = "Va a contabilizar la factura" & vbCrLf & vbCrLf & "Numero:  " & _
                Data1.Recordset!NumSerie & "-" & Data1.Recordset!codfaccl & "       Fecha: " & Adodc2.Recordset!fecfaccl
            SQL = SQL & vbCrLf & vbCrLf & "     ¿Desea continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            'Bloqueo el registro
            If Not BloqueaRegistroForm(Me) Then Exit Sub
        Else
            'Viene de insertar la factura. NO tengo que hacer la pregunta
            
        End If
        
        Screen.MousePointer = vbHourglass
        
        'Actualizar
        If IntegrarFactura Then
        
            If Data1.Recordset.EOF Then
                LimpiarCampos
                CargaGrid False
                PonerModo 0
            Else
                PonerCampos
                PonerModo 2
            End If
        End If
    
        If Not ContabilizarDespuesInsertar Then
            'Ha pulsado la camara de fotos
            TerminaBloquear
            DesBloqueaRegistroForm Me.Text1(1)
        End If
        Screen.MousePointer = vbDefault
        
    Case 12
        'GEnerar VTOS
        If Data1.Recordset.EOF Then Exit Sub
        If Adodc2.Recordset.EOF Then Exit Sub
        
        varFecOk = FechaCorrecta2(CDate(Text1(0).Text))
        SQL = ""
        If varFecOk >= 2 Then
            If varFecOk = 2 Then
                SQL = varTxtFec
            Else
                SQL = "Fecha factura fuera de ejercicios."
            End If
            SQL = SQL & vbCrLf & "¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        SQL = ""
        'Llamo al formualiro de generar VTOS
        frmVto.opcion = 0
        With Adodc2.Recordset
            frmVto.Importe = !totfaccl
            'frmVto.Datos = "A|2|2005|16/05/05|4300002|INTERVANESTRUJEN|"
            frmVto.Datos = !NumSerie & "|" & !codfaccl & "|" & !anofaccl & "|" & !fecfaccl & "|" & !codmacta & "|" & Text4(0).Text & "|"
        End With
        frmVto.Show vbModal
        
    Case 13
        frmListado.opcion = 8
        frmListado.Show vbModal
    Case 14
        If Modo = 4 Or Modo = 5 Then If MsgBox("Esta editando el registro. Seguro que desea salir?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        mnSalir_Click
    Case 16 To 19
        Desplazamiento (Indi - 16)
    Case Else
    
    End Select
   
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    For i = 16 To 19
        Toolbar1.Buttons(i).Visible = bol
    Next i
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
        
    For i = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(i).AllowSizing = False
    Next i
   
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
    
    SQL = "SELECT linfact.codtbase, cuentas.nommacta, linfact.codccost, ccoste.nomccost, linfact.impbascl, linfact.numlinea"
    SQL = SQL & " FROM (ccoste RIGHT JOIN linfact ON ccoste.codccost = linfact.codccost) INNER JOIN cuentas ON linfact.codtbase = cuentas.codmacta WHERE "
    If Enlaza Then
        SQL = SQL & " numserie = '" & Data1.Recordset!NumSerie & "'"
        SQL = SQL & " AND codfaccl = " & Data1.Recordset!codfaccl
        SQL = SQL & " AND anofaccl= " & Data1.Recordset!anofaccl
        Else
        SQL = SQL & " numserie = 'david'"
    End If
    SQL = SQL & " ORDER BY linfact.numlinea"
    MontaSQLCarga = SQL
End Function

Private Sub AnyadirLinea(Limpiar As Boolean, DesdeBoton As Boolean)
    Dim anc As Single
    Dim Preg As String
    
    If ModificandoLineas = 2 Then Exit Sub
    
    '++
    ModificandoLineas = 1
    PonerModo 5
    '++
    
    'Obtenemos la siguiente numero de factura
    Linfac = ObtenerSigueinteNumeroLinea   'Fijamos en aux el importe que queda
    If AUx = 0 Then
        If DesdeBoton Then
            Preg = "La suma de las lineas coincide con el importe factura. ¿Seguro que desea insertar mas lineas?"
            If MsgBox(Preg, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Else
            LLamaLineas anc, 0, True
'--
'            cmdCancelar.Caption = "Cabecera"
'++
            cmdCancelar_Click
'++

            Exit Sub
        End If
    End If
    
    
   'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveLast
        DataGrid1.HoldFields
        'DataGrid1.Row = DataGrid1.Row + 1
    End If
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 220
        Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row + 1) + 15
    End If
    LLamaLineas anc, 1, Limpiar
    'Ponemos el importe
    
    txtaux(4).Text = AUx
    HabilitarCentroCoste
    'Ponemos el foco
    PonerFoco txtaux(0)
    
End Sub

Private Sub ModificarLinea()
Dim Cad As String
Dim anc As Single
    If Adodc1.Recordset.EOF Then Exit Sub
    If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If ModificandoLineas <> 0 Then Exit Sub
    
    
    Me.lblIndicador.Caption = "MODIFICAR"
     
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
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
    PonerFoco txtaux(0)
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
        SQL = "Delete from linfact"
        SQL = SQL & " WHERE numlinea = " & Adodc1.Recordset!NumLinea
        SQL = SQL & " AND anofaccl=" & Data1.Recordset!anofaccl
        SQL = SQL & " AND numserie='" & Data1.Recordset!NumSerie
        SQL = SQL & "' AND codfaccl=" & Data1.Recordset!codfaccl & ";"
        Conn.Execute SQL
        
        vLog.Insertar 5, vUsu, "Lin_e: " & Data1.Recordset!NumSerie & Format(Data1.Recordset!codfaccl, "000000") & "  nº:" & Adodc1.Recordset!NumLinea
        
        
        CargaGrid (Not Data1.Recordset.EOF)
        
    End If
    '++
    cmdCancelar_Click
    '++
End Sub


'Ademas de obtener el siguiente nº de linea, tb obtiene la suma de
'las lineas de factura, Y fijamos lo que falta en aux para luego ofertarlo

Private Function ObtenerSigueinteNumeroLinea() As Long
    Dim Rs As ADODB.Recordset
    Dim i As Long
    
    Set Rs = New ADODB.Recordset
    
    SQL = " WHERE linfact.numserie= '" & Data1.Recordset!NumSerie & "'"
    SQL = SQL & " AND linfact.codfaccl= " & Data1.Recordset!codfaccl
    SQL = SQL & " AND linfact.anofaccl=" & Data1.Recordset!anofaccl & ";"
    Rs.Open "SELECT Max(numlinea) FROM linfact" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    i = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then i = Rs.Fields(0)
    End If
    Rs.Close

    'La suma
    SumaLinea = 0
    If i > 0 Then
        Rs.Open "SELECT sum(impbascl) FROM linfact" & SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then SumaLinea = Rs.Fields(0)
        End If
        Rs.Close
    End If
    Set Rs = Nothing
    
    'Lo que falta lo fijamos en aux. El importe es de la bASE IMPONIBLE si fuera del total seria Text2(4).Text
    AUx = ImporteFormateado(Text2(0).Text)
    AUx = AUx - SumaLinea
    ObtenerSigueinteNumeroLinea = i + 1
End Function




Private Sub LLamaLineas(alto As Single, xModo As Byte, Limpiar As Boolean)
    Dim B As Boolean
    DeseleccionaGrid
'--
'    cmdCancelar.Caption = "Cancelar"
    ModificandoLineas = xModo
    B = (xModo = 0)
'--
'    cmdAceptar.Visible = Not B
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
    For i = 0 To txtaux.Count - 1
        If i < 2 Or i > 3 Then txtaux(i).Visible = Visible
        txtaux(i).Top = Altura
    Next i

    If Limpiar Then
        For i = 0 To txtaux.Count - 1
            txtaux(i).Text = ""
        Next i
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

'++
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYBusqueda KeyAscii, 100 'cuenta
            Case 2:  KEYBusquedaAux KeyAscii, 1 'contador
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYBusquedaAux(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    cmdAux_Click (indice)
End Sub
'++




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
                PonerFoco txtaux(0)
            Else
                If txtaux(2).Visible And txtaux(2).Enabled Then
                    PonerFoco txtaux(2)
                Else
                    PonerFoco txtaux(4)
                End If
            End If
        Case 2
            txtaux(2).Text = UCase(txtaux(2).Text)
            RC = "idsubcos"
            SQL = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtaux(2).Text, "T", RC)
            If SQL = "" Then
                MsgBox "Centro de coste no encontrado: " & txtaux(2).Text, vbExclamation
                txtaux(2).Text = ""
                PonerFoco txtaux(2)
            End If
            txtaux(3).Text = SQL
            If SQL <> "" Then PonerFoco txtaux(4)
        Case 4
            If Not EsNumerico(txtaux(4).Text) Then
                'MsgBox "Importe incorrecto: " & txtaux(4).Text, vbExclamation
                txtaux(4).Text = ""
                PonerFoco txtaux(4)
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
    AUx = ImporteFormateado(Text2(4).Text)
    
    'Si tiene retencion
    AUX2 = 0
    If Text2(3).Text <> "" Then AUX2 = ImporteFormateado(Text2(3).Text)
    AUx = AUx + AUX2
    
    'El iVA
    AUX2 = 0
    If Text2(1).Text <> "" Then AUX2 = ImporteFormateado(Text2(1).Text)
    AUx = AUx - AUX2
    
    'La retencion
    AUX2 = 0
    If Text2(2).Text <> "" Then AUX2 = ImporteFormateado(Text2(2).Text)
    AUx = AUx - AUX2
    
    
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
    
    'Facturas positivas
'    If Aux > 0 Then
'        If AUX2 > Aux Then
'                AuxOK = "El importe excede del total de factura"
'                Exit Function
'        End If
'    Else
'        If AUX2 < Aux Then
'                AuxOK = "El importe excede del total de factura"
'                Exit Function
'        End If
'    End If
    AuxOK = ""
End Function





Private Function InsertarModificar() As Boolean
    
    On Error GoTo EInsertarModificar
    InsertarModificar = False
    
    If ModificandoLineas = 1 Then
        SQL = "INSERT INTO linfact (numserie, codfaccl, anofaccl, numlinea, codtbase, impbascl, codccost) VALUES ('"
        ''R', 11, 2003, 1, '6000001', 1500, 'TIEN')
        SQL = SQL & Data1.Recordset!NumSerie & "',"
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
        'UPDATE asipre_lineas SET numdocum= '3' WHERE numaspre=1 AND linlapre=1
        '(codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab)
        SQL = "UPDATE linfact SET "
        
        SQL = SQL & " codtbase = '" & txtaux(0).Text & "',"
        SQL = SQL & " impbascl = "
        SQL = SQL & TransformaComasPuntos(txtaux(4).Text) & ","
        
        'Centro coste
        If txtaux(2).Text = "" Then
          SQL = SQL & " codccost = " & ValorNulo
          Else
          SQL = SQL & " codccost = '" & txtaux(2).Text & "'"
        End If
    
        SQL = SQL & " WHERE numserie='" & Data1.Recordset!NumSerie
        SQL = SQL & "' AND codfaccl= " & Data1.Recordset!codfaccl
        SQL = SQL & " AND anofaccl=" & Data1.Recordset!anofaccl
        SQL = SQL & " AND numlinea =" & Adodc1.Recordset!NumLinea & ";"


        
        'LOG
        vLog.Insertar 5, vUsu, "Lin: " & Data1.Recordset!NumSerie & Format(Data1.Recordset!codfaccl, "000000") & "  nº:" & Adodc1.Recordset!NumLinea
        

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
Dim B As Boolean, bAux As Boolean
Dim i As Integer

    PonerOpcionesMenuGeneral Me

    '++
    B = (Modo = 2)
    For i = 1 To 1
        ToolAux(0).Buttons(1).Enabled = B
        
        If B Then
            Select Case i
              Case 1
                bAux = (B And Me.Adodc1.Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(0).Buttons(2).Enabled = bAux
        ToolAux(0).Buttons(3).Enabled = bAux
    Next i
    '++
End Sub


Private Function PonerValoresIva(numero As Integer) As Boolean
Dim J As Integer
J = ((numero - 1) * 6) + 7
Set Rs = New ADODB.Recordset
Rs.Open "Select nombriva,porceiva,porcerec from tiposiva where codigiva =" & Text1(J).Text, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
If Rs.EOF Then
    MsgBox "Tipo de IVA incorrecto: " & Text1(J).Text, vbExclamation
    Text1(J).Text = ""
    Text4(numero).Text = ""
    PonerValoresIva = False
Else
    PonerValoresIva = True
    
    Text4(numero).Text = Rs.Fields(0)
    Text1(J + 1).Text = Format(Rs.Fields(1), "#0.00")
    AUx = DBLet(Rs.Fields(2), "N")
    If AUx = 0 Then
        Text1(J + 3).Text = ""
        Else
        Text1(J + 3).Text = Format(AUx, "#0.00")
    End If
End If
Rs.Close
End Function



Private Sub CalcularIVA(numero As Integer)
Dim J As Integer


J = ((numero - 1) * 6) + 6
Base = ImporteFormateado(Text1(J).Text)

'EL iva
AUx = ImporteFormateado(Text1(J + 2).Text) / 100
If AUx = 0 Then
    If Text1(J + 2).Text = "" Then
        Text1(J + 3).Text = ""
    Else
        Text1(J + 3).Text = "0,00"
    End If
Else
    Text1(J + 3).Text = Format(Round((AUx * Base), 2), FormatoImporte)
End If

'Recargo
AUx = ImporteFormateado(Text1(J + 4).Text) / 100
If AUx = 0 Then
    Text1(J + 5).Text = ""
Else
    Text1(J + 5).Text = Format(Round((AUx * Base), 2), FormatoImporte)
End If

End Sub


Private Sub TotalesBase()
    'Base imponible
    AUx = 0
    For i = 1 To 3
        If Text1(i * 6).Text <> "" Then
            Base = ImporteFormateado(Text1(i * 6).Text)
            AUx = AUx + Base
        End If
    Next i
    If AUx = 0 Then
        Text2(0).Text = ""
    Else
        Text2(0).Text = Format(AUx, FormatoImporte)
    End If
End Sub


Private Sub TotalesIVA()
On Error GoTo et
    'Total IVA
    AUx = 0
    For i = 1 To 3
        ancho = (i * 6) + 3
        If Text1(ancho).Text <> "" Then
            Base = ImporteFormateado(Text1(ancho).Text)
            AUx = AUx + Base
        End If
    Next i
    If AUx = 0 Then
        Text2(1).Text = ""
    Else
        Text2(1).Text = Format(AUx, FormatoImporte)
    End If
    
    Exit Sub
et:
    MuestraError Err.Number, "Calculando total IVA"
    Text2(1).Text = ""
End Sub

Private Sub TotalesRecargo()
    'RECARGO
    AUx = 0
    For i = 1 To 3
        ancho = (i * 6) + 5
        If Text1(ancho).Text <> "" Then
            Base = ImporteFormateado(Text1(ancho).Text)
            AUx = AUx + Base
        End If
    Next i
    If AUx = 0 Then
        Text2(2).Text = ""
    Else
        Text2(2).Text = Format(AUx, FormatoImporte)
    End If
End Sub

Private Sub TotalFactura()
    'El total
    AUx = 0
    ' Base + iva + recargao   -  retencion
    For i = 0 To 2
        If Text2(i).Text = "" Then
   
        Else
            Base = ImporteFormateado(Text2(i).Text)
            AUx = AUx + Base
        End If
    Next i
    If Text2(3).Text = "" Then
        
    Else
        Base = ImporteFormateado(Text2(3).Text)
        AUx = AUx - Base
    End If
    
    If AUx = 0 Then
        Text2(4).Text = ""
    Else
        Text2(4).Text = Format(AUx, FormatoImporte)
    End If
    Text1(27).Text = TransformaComasPuntos(CStr(AUx))
End Sub


'Comprobara si el periodo esta liquidado o no.
'Si la fecha pertenece a un periodo liquidado entonces
'mostraremos un mensaje para preguntar si desea continuar o no
Private Function ComprobarPeriodo2(indice As Integer) As Boolean
Dim Cerrado As Boolean

    'Primero pondremos la fecha a año periodo
    i = Year(CDate(Text1(indice).Text))
    If vParam.periodos = 0 Then
        'Trimestral
        ancho = ((Month(CDate(Text1(indice).Text)) - 1) \ 3) + 1
        Else
        ancho = Month(CDate((Text1(indice).Text)))
    End If
    Cerrado = False
    If i < vParam.anofactu Then
        Cerrado = True
    Else
        If i = vParam.anofactu Then
            'El mismo año. Comprobamos los periodos
            If vParam.perfactu >= ancho Then _
                Cerrado = True
        End If
    End If
    ComprobarPeriodo2 = True
    ModificaFacturaPeriodoLiquidado = False
    If Cerrado Then
        ModificaFacturaPeriodoLiquidado = True
        SQL = "La fecha "
        If indice = 0 Then
            SQL = SQL & "factura"
        Else
            SQL = SQL & "liquidacion"
        End If
        SQL = SQL & " corresponde a un periodo ya liquidado. " & vbCrLf
        SQL = SQL & vbCrLf & " ¿Desea continuar igualmente ?"
  
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then ComprobarPeriodo2 = False
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




Private Function Desintegrar() As Boolean
        Desintegrar = False
        'Primero hay que desvincular la factura de la tabla de hco
        If DesvincularFactura Then
            frmActualizar.OpcionActualizar = 2  'Desactualizar para eliminar
            frmActualizar.NumAsiento = Adodc2.Recordset!NumAsien
            frmActualizar.FechaAsiento = Adodc2.Recordset!FechaEnt
            frmActualizar.NumDiari = Adodc2.Recordset!NumDiari
            AlgunAsientoActualizado = False
            frmActualizar.Show vbModal
            If AlgunAsientoActualizado Then Desintegrar = True
        End If
End Function


Private Function DesvincularFactura() As Boolean
On Error Resume Next
    SQL = "UPDATE cabfact set numasien=NULL, fechaent=NULL, numdiari=NULL"
    SQL = SQL & " WHERE codfaccl = " & Adodc2.Recordset!codfaccl
    SQL = SQL & " AND numserie = '" & Adodc2.Recordset!NumSerie & "'"
    SQL = SQL & " AND anofaccl =" & Adodc2.Recordset!anofaccl
    Numasien2 = Adodc2.Recordset!NumAsien
    NumDiario = Adodc2.Recordset!NumDiari
    Conn.Execute SQL
    If Err.Number <> 0 Then
        DesvincularFactura = False
        MuestraError Err.Number, "Desvincular factura"
    Else
        DesvincularFactura = True
    End If
End Function



Private Sub LeerFiltro(Leer As Boolean)
    SQL = App.Path & "\filfac.dat"
    If Leer Then
        FILTRO = 0
        If Dir(SQL) <> "" Then
            AbrirFicheroFiltro True
            If IsNumeric(SQL) Then FILTRO = CByte(SQL)
        End If
    Else
        AbrirFicheroFiltro False
    End If
End Sub


Private Sub AbrirFicheroFiltro(Leer As Boolean)
On Error GoTo eAbrir
    i = FreeFile
    If Leer Then
        Open SQL For Input As #i
        SQL = "0"
        Line Input #i, SQL
    Else
        Open SQL For Output As #i
        Print #i, FILTRO
    End If
    Close #i
    Exit Sub
eAbrir:
    Err.Clear
End Sub


Private Sub PonerFiltro(NumFilt As Byte)
    FILTRO = NumFilt
    Me.mnActual.Checked = (NumFilt = 2)
    Me.mnActuralySiguiente.Checked = (NumFilt = 1)
    Me.mnSiguiente.Checked = (NumFilt = 3)
    Me.mnSinFiltro.Checked = (NumFilt = 0)
End Sub


Private Function AnyadeCadenaFiltro() As String
Dim AUx As String

    AUx = ""
    If FILTRO <> 0 Then
        '-------------------------------- INICIO
        i = Year(vParam.fechaini)
        If FILTRO < 3 Then
            'INicio = actual
            AUx = " anofaccl >= " & i
            Else
            AUx = " anofaccl >=" & i + 1
        End If
        i = Year(vParam.fechafin)
        If FILTRO = 2 Then
            AUx = AUx & " AND anofaccl <= " & i
        Else
            AUx = AUx & " AND anofaccl <= " & i + 1
        End If
        
    End If  'filtro=0
    AnyadeCadenaFiltro = AUx
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




Private Function SituarADO2() As Boolean
Dim L As Long

On Error GoTo ESituar2
    SituarADO2 = False
    Adodc2.Refresh
    SituarADO2 = True
ESituar2:
    Err.Clear
End Function



Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub HabilitarTXTCabecerasAlModificar(Preparando As Boolean)
Dim J As Integer

    'Si el usuario no tiene permisos le bloqueamos muchas cosas
    If vUsu.Nivel = 2 Then
        For J = 0 To 25
            'Solo dejamos enabled fecha, codclien, concepto
            'cta retencion.  Index: 0,5,25,3
            If Preparando Then
                If Not (J = 0 Or J = 5 Or J = 25 Or J = 3) Then Text1(J).Enabled = False
            Else
                Text1(J).Enabled = True
            End If
        Next J
        
        If Preparando Then
            imgppal(0).Enabled = False
            For J = 3 To 5
                imgppal(J).Enabled = False
            Next J
        End If
        
    End If
End Sub


'++ SUBMENU DE LINEAS


Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

'    If BloqueaRegistro("cabfact", "numserie= '" & Trim(Data1.Recordset!NUmSerie) & "' and codfaccl= " & Data1.Recordset!codfaccl & " and anofaccl = " & Data1.Recordset!anofaccl) Then
'    If BLOQUEADesdeFormulario(Me, data1, 1) Then
    If BotonLineas Then
        Select Case Button.Index
            Case 1
                'AÑADIR linea factura
                AnyadirLinea True, True
            
            Case 2
                'MODIFICAR linea factura
                ModificarLinea
            
            Case 3
                'ELIMINAR linea factura
                EliminarLineaFactura
            
            Case Else
        End Select
    End If
'    End If

End Sub

'++ convierto el boton 10 de lineas en una fucntion antes de ejecutar el toolaux
Private Function BotonLineas() As Boolean
Dim N As Long
     

        If FechaCorrecta2(CDate(Text1(0).Text)) >= 2 Then
            MsgBox "La factura pertenece a un ejercicio cerrado o fuera de ambito.", vbExclamation
            BotonLineas = False
            Exit Function
        End If
    
        If Numasien2 > 0 Then
            'Viene, de modificar cabecera factura. Tenemos k volvera enlazar
            CargaGrid True
            espera 0.1
        End If
    
    
        If Not ComprobarPeriodo2(28) Then
            BotonLineas = False
            Exit Function
        End If
    
        If Not IsNull(Adodc2.Recordset!NumAsien) Then
            N = Adodc2.Recordset!NumAsien
            If N = 0 Then
                MsgBox "Contabilizacion de facturas especial. No puede modificarse", vbExclamation
                BotonLineas = False
                Exit Function
            End If
            SQL = "Esta factura ya esta contabilizada. Desea desactualizar para poder modificarla?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
            Numasien2 = Adodc2.Recordset!NumAsien
            NumDiario = Adodc2.Recordset!NumDiari
            'Tengo desintegre la factura del hco
            If Not Desintegrar Then
                BotonLineas = False
                Exit Function
            End If
            ObtenerSigueinteNumeroLinea
            Text1(4).Text = ""
        End If
        
   
        'If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
        '++
        DesBloqueaRegistroForm Me.Text1(1)
        '++
        
        
        If Not BloqueaRegistroForm(Me) Then
            BotonLineas = False
            Exit Function
        End If
        
        'Fuerzo que se vean las lineas
'        cmdCancelar.Caption = "Cabecera"
'        lblIndicador.Caption = "Lineas detalle"
        
        
        PonerModo 5
        ModificandoLineas = 0
        'Si tiene numasien es k kiere modificar algo, luego se lo sugiero
        If Numasien2 > 0 Then
            If Adodc1.Recordset.RecordCount = 1 Then ModificarLinea
        End If
        BotonLineas = True

End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codfaccl=" & Data1.Recordset!codfaccl
    Cad = Cad & " AND anofaccl = " & Data1.Recordset!anofaccl
    Cad = Cad & " AND numserie= '" & Data1.Recordset!NumSerie & "'" & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
        If ModificandoLineas <> 0 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub


Private Sub BotonCabecera()
        
        'Si esta insertando/modificando lineas haremos unas cosas u otras
        DataGrid1.Enabled = True
        If ModificandoLineas = 0 Then
            'NUEVO
            AntiguoText1 = ""
            If Adodc1.Recordset.EOF Then
                AntiguoText1 = "La factura no tiene lineas. ¿SEGURO que desea salir?"
                If MsgBox(AntiguoText1, vbQuestion + vbYesNoCancel) = vbYes Then
                    AntiguoText1 = ""
                Else
                    'Para k no muestre el siguiente punto de error
                    AntiguoText1 = "###"
                End If
            Else
                'Comprobamos que el total de factura es el de suma
               ObtenerSigueinteNumeroLinea
               If AUx <> 0 Then AntiguoText1 = "El importe de lineas no suma el importe facturas: " & Format(AUx, FormatoImporte)
            End If
            If AntiguoText1 <> "" Then
                If AntiguoText1 <> "###" Then MsgBox AntiguoText1, vbExclamation
'--
'                Exit Sub
'--
            End If
            lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            If Adodc2.Recordset Is Nothing Then
                CargaGrid True
                If Not SituarData1(0) Then
                    PonerModo 0
                Else
                    
                    PonerCampos
                    PonerModo 2
                    NuevaFactura = True
                End If
            Else
                EnlazaADOs
                PonerModo 2
            End If
            
            If Modo = 2 Then
                If Numasien2 > 0 Then
                    If IntegrarFactura Then
                        Text1(4).Text = Numasien2
                        Numasien2 = 0
                        NumDiario = 0
                    End If
                End If
                
                'Nuevo enero 2009
                'Ademas del vto, que ya lo hacia, ahora contabilizara la factura
                If NuevaFactura Then
                    Screen.MousePointer = vbHourglass
                    espera 0.2
                    'Contabilizar la factura autmaticamente
                    If vParam.ContabilizaFactura Then HacerToolbar1 11, True
                
                    'Esto ya lo hacia. Insertar vto
                    If vEmpresa.TieneTesoreria Then HacerToolbar1 12, False
                    
                End If
            End If
            DesBloqueaRegistroForm Me.Text1(1)
            Screen.MousePointer = vbDefault
        Else
            If ModificandoLineas = 1 Then
                 DataGrid1.AllowAddNew = False
                 If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
                 DataGrid1.Refresh
            End If
'--
'            cmdAceptar.Visible = False
'            cmdCancelar.Caption = "Cabeceras"
'--
           
            ModificandoLineas = 0

        End If
        
End Sub
