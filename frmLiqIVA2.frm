VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLiqIVA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación del IVA"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   Icon            =   "frmLiqIVA2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   38
      Left            =   7680
      TabIndex        =   73
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   37
      Left            =   7680
      TabIndex        =   71
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   4446
      Width           =   1395
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8190
      TabIndex        =   68
      Top             =   6780
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmLiqIVA2.frx":030A
      Left            =   5400
      List            =   "frmLiqIVA2.frx":0317
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9360
      TabIndex        =   65
      Top             =   6780
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   36
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   62
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   6240
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   35
      Left            =   7680
      TabIndex        =   54
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   4092
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   34
      Left            =   9090
      Locked          =   -1  'True
      TabIndex        =   51
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   2400
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   33
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   50
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   32
      Left            =   9120
      TabIndex        =   48
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   3738
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   31
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   47
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   30
      Left            =   9120
      TabIndex        =   45
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   3384
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   29
      Left            =   9120
      TabIndex        =   43
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   3030
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   27
      Left            =   9090
      TabIndex        =   41
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1800
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   28
      Left            =   9090
      TabIndex        =   39
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1470
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   26
      Left            =   9090
      TabIndex        =   37
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1140
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   25
      Left            =   7560
      TabIndex        =   36
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1140
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   24
      Left            =   9090
      TabIndex        =   34
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   810
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   23
      Left            =   7560
      TabIndex        =   33
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   810
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   22
      Left            =   9090
      TabIndex        =   31
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   480
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   21
      Left            =   7560
      TabIndex        =   30
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   480
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   20
      Left            =   2700
      TabIndex        =   27
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4155
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   19
      Left            =   1860
      TabIndex        =   26
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   4155
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   18
      Left            =   180
      TabIndex        =   25
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   4155
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   17
      Left            =   2700
      TabIndex        =   24
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   3705
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   16
      Left            =   1860
      TabIndex        =   23
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   3705
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   15
      Left            =   180
      TabIndex        =   22
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   14
      Left            =   2700
      TabIndex        =   21
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   3270
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   13
      Left            =   1860
      TabIndex        =   20
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   3270
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   12
      Left            =   180
      TabIndex        =   19
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   3270
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   11
      Left            =   2700
      TabIndex        =   15
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   5370
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   10
      Left            =   1860
      TabIndex        =   14
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   5370
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   9
      Left            =   180
      TabIndex        =   13
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   5370
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   8
      Left            =   2700
      TabIndex        =   12
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   2070
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   7
      Left            =   1860
      TabIndex        =   11
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   2070
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   6
      Left            =   180
      TabIndex        =   10
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   2070
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   5
      Left            =   2700
      TabIndex        =   9
      Tag             =   "vv"
      Text            =   "Text1"
      Top             =   1620
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   4
      Left            =   1860
      TabIndex        =   8
      Tag             =   "vv"
      Text            =   "Text1"
      Top             =   1620
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Tag             =   "vv"
      Top             =   1620
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   2
      Left            =   2700
      TabIndex        =   2
      Tag             =   "vv"
      Text            =   "Text1"
      Top             =   1185
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   1
      Left            =   1860
      TabIndex        =   1
      Tag             =   "vv"
      Text            =   "Text1"
      Top             =   1185
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "vv"
      Top             =   1185
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Entregas intracomunitarias"
      Height          =   195
      Index           =   10
      Left            =   4560
      TabIndex        =   74
      Top             =   4870
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "Exportaciones y operaciones asimiladas"
      Height          =   195
      Index           =   9
      Left            =   4560
      TabIndex        =   72
      Top             =   4514
      Width           =   3105
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   120
      TabIndex        =   70
      Top             =   6600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Destino: "
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
      Index           =   17
      Left            =   4200
      TabIndex        =   69
      Top             =   6900
      Width           =   975
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
      Index           =   16
      Left            =   7470
      TabIndex        =   67
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   15
      Left            =   7200
      TabIndex        =   64
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   14
      Left            =   810
      TabIndex        =   63
      Top             =   6090
      Width           =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   4680
      X2              =   10530
      Y1              =   6150
      Y2              =   6150
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
      Index           =   13
      Left            =   180
      TabIndex        =   61
      Top             =   5100
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "%Recar."
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
      Index           =   12
      Left            =   1860
      TabIndex        =   60
      Top             =   5100
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Cuota"
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
      Left            =   2700
      TabIndex        =   59
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "ADQUISICIONES INTRACOMUNITARIAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   11
      Left            =   180
      TabIndex        =   58
      Top             =   4830
      Width           =   3510
   End
   Begin VB.Label Label3 
      Caption         =   "RECARGO EQUIVALENCIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   10
      Left            =   180
      TabIndex        =   57
      Top             =   2730
      Width           =   2325
   End
   Begin VB.Label Label3 
      Caption         =   "REGIMEN GENERAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   56
      Top             =   630
      Width           =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      BorderWidth     =   2
      Index           =   2
      X1              =   6240
      X2              =   10560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label1 
      Caption         =   "Entregas intracomunitarias"
      Height          =   195
      Index           =   8
      Left            =   4530
      TabIndex        =   55
      Top             =   4158
      Width           =   3105
   End
   Begin VB.Label Label3 
      Caption         =   "Cuota"
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
      Index           =   8
      Left            =   9600
      TabIndex        =   53
      Top             =   240
      Width           =   510
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
      Index           =   7
      Left            =   7560
      TabIndex        =   52
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cuotas a compensar de  periodos anteriores"
      Height          =   195
      Index           =   7
      Left            =   4530
      TabIndex        =   49
      Top             =   3802
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "Atribuible a la admon del estado"
      Height          =   195
      Index           =   6
      Left            =   4530
      TabIndex        =   46
      Top             =   3446
      Width           =   2250
   End
   Begin VB.Label Label1 
      Caption         =   "Diferencia"
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
      Index           =   5
      Left            =   4530
      TabIndex        =   44
      Top             =   3090
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Liquidación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Index           =   2
      Left            =   4530
      TabIndex        =   42
      Top             =   2700
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   7560
      X2              =   10530
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "Regularización inversiones"
      Height          =   195
      Index           =   4
      Left            =   4440
      TabIndex        =   40
      Top             =   1920
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Compensacion régimen especial A.G. y P."
      Height          =   195
      Index           =   3
      Left            =   4410
      TabIndex        =   38
      Top             =   1560
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Adquisiciones intracomunitarias"
      Height          =   195
      Index           =   2
      Left            =   4410
      TabIndex        =   35
      Top             =   1200
      Width           =   2385
   End
   Begin VB.Label Label1 
      Caption         =   "En importaciones"
      Height          =   285
      Index           =   1
      Left            =   4410
      TabIndex        =   32
      Top             =   870
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "En operaciones interiores"
      Height          =   195
      Index           =   0
      Left            =   4410
      TabIndex        =   29
      Top             =   540
      Width           =   2955
   End
   Begin VB.Label Label2 
      Caption         =   "I.V.A  deducible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   1
      Left            =   4410
      TabIndex        =   28
      Top             =   90
      Width           =   2445
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   1800
      X2              =   3960
      Y1              =   5910
      Y2              =   5910
   End
   Begin VB.Label Label3 
      Caption         =   "Cuota"
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
      Index           =   5
      Left            =   2700
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "%Recar."
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
      Left            =   1860
      TabIndex        =   17
      Top             =   3000
      Width           =   735
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
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Cuota"
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
      Index           =   4
      Left            =   3150
      TabIndex        =   6
      Top             =   855
      Width           =   510
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
      Left            =   1860
      TabIndex        =   5
      Top             =   855
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "I.V.A  devengado"
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
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   2445
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
      Left            =   180
      TabIndex        =   3
      Top             =   855
      Width           =   1455
   End
End
Attribute VB_Name = "frmLiqIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'

Public Periodo As String 'PeriodoINI|PeriodFIN|año|Detallado 0 o 1
Public Modelo As Byte
    '0.- 300
    '1.- 320
    '2.- 330
    '3.- 332
Dim PrimeraVez As Boolean
Dim cad As String
Dim I As Integer
Dim Importe As Currency

Dim Text1Ant As String



Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    'Imprimir
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione una forma de impresión", vbExclamation
        Exit Sub
    End If
    
    
    Select Case Combo1.ItemData(Combo1.ListIndex)
    Case 3, 4
            'Generamos la cadena con los importes a mostrar
            cad = ""
            GeneraCadenaImportes

            'Ahora enviamos a generar fichero IVA
            If GenerarFicheroIVA(cad, CCur(Text1(36).Text), Now, Periodo) Then
                'Si es de tipo 1, es decir impresion automatica de datos
                If Combo1.ItemData(Combo1.ListIndex) = 3 Then
                    ImpresionModeloOficial
                Else
                    GuardarComo
                End If
            End If
    Case Else
        
         If vParam.periodos = 1 Then
            Text1Ant = "MENSUAL"
         Else
            Text1Ant = "TRIMESTRAL"
         End If
         cad = "CampoSeleccion= """ & Text1Ant & """|"
         If vParam.periodos = 0 Then
            Text1Ant = "Trimestre "
         Else
            Text1Ant = "MES "
         End If
         'Inicio
         cad = cad & "TInicio= """ & Text1Ant & "inicio""|"
         cad = cad & "Inicio= """ & RecuperaValor(Periodo, 1) & """|"
         'Fin
         cad = cad & "TFin= """ & Text1Ant & "fin""|"
         cad = cad & "Fin= """ & RecuperaValor(Periodo, 2) & """|"
         'Anyo
         cad = cad & "Anyo= """ & RecuperaValor(Periodo, 3) & """|"
         If Combo1.ListIndex = 0 Then
            I = 30
         Else
            I = 63
         End If
         With frmImprimir
                .OtrosParametros = cad
                .NumeroParametros = 6
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = I
                .Show vbModal
         End With
    End Select
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        'Cargamos los datos del IVA
        PonDatosIVA
        'Vemos las adquisiciones intracomunitrias
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Limpiar Me
    
    cad = RecuperaValor(Periodo, 4)
'    I = InStr(1, Periodo, "|" & cad & "|")
'    If I > 0 Then Periodo = Mid(Periodo, 1, I) 'Sin -1 para conservar el pipe
    
    'Modelo
    Label4.Caption = "Modelo "
    Select Case Modelo
    Case 0
        Label4.Caption = Label4.Caption & "300"
    Case 1
        Label4.Caption = Label4.Caption & "320"
    Case 2
        Label4.Caption = Label4.Caption & "330"
    Case 3
        Label4.Caption = Label4.Caption & "332"
    Case Else
        Label4.Caption = "Error pasando modelo"
    End Select
    Combo1.Clear
    Combo1.AddItem "Borrador"
    Combo1.ItemData(Combo1.NewIndex) = 1
    If Val(cad) > 0 Then
        Combo1.AddItem "Informe detallado"
        Combo1.ItemData(Combo1.NewIndex) = 2
        I = 1
    Else
        I = 0
    End If
    'Solo el modelo 300 tiene impresion hacienda
    If Modelo = 0 Then
        Combo1.AddItem "Modelo hacienda"
        Combo1.ItemData(Combo1.NewIndex) = 3
    End If
    Combo1.AddItem "Internet"
    Combo1.ItemData(Combo1.NewIndex) = 4
    Combo1.ListIndex = I
End Sub

''Del objeto miRSAUX, el field vcampo le pondra el valor
'Private Function PonerValorImporte(Vcampo As Integer) As String
'Dim Valor As Currency
'    If IsNull(miRsAux.Fields(Vcampo)) Then
'        Valor = 0
'    Else
'        Valor = miRsAux.Fields(Vcampo)
'    End If
'    If Valor = 0 Then
'        PonerValorImporte = ""
'    Else
'        PonerValorImporte = Format(Valor, FormatoImporte)
'    End If
'End Function

Private Sub PonDatosIVA()
Dim Bases As Currency
Dim Cuotas As Currency
Dim IVAS_NO As String
    'IVA DEVENGADO
    
    Set miRsAux = New ADODB.Recordset
    
    IVAS_NO = BuscarIvasADescartar
    'En clientes, el IVA 0 tampoco entra  metemos el 0%
    IVAS_NO = IVAS_NO & "0|"
    
    cad = "select iva,sum(bases),sum(ivas) from Usuarios.zliquidaiva where "
    cad = cad & " codusu=" & vUsu.Codigo & " and cliente= 1 group by iva order by iva ASC"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 1
    While Not miRsAux.EOF
        cad = "|" & miRsAux!iva & "|"
        'Si el IVA no se debe poner... NO lo ponemos
        If InStr(1, IVAS_NO, cad) = 0 Then
            PonerElIva I
            I = I + 1
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
            
    
    'Cargamos INTRACOM. Habra k cargar datos en temporales
    
    
    
    'IVA DEDUCIBLE
    Bases = 0: Cuotas = 0
    IVAS_NO = BuscarIvasADescartar
    
    cad = "select * from Usuarios.zliquidaiva where codusu=" & vUsu.Codigo
    cad = cad & " and cliente= 0"
    miRsAux.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = "|" & miRsAux!iva & "|"
        'Si el IVA no se debe poner... NO lo ponemos
        If InStr(1, IVAS_NO, cad) = 0 Then
            Bases = Bases + miRsAux!Bases
            Cuotas = Cuotas + miRsAux!ivas
        End If
        miRsAux.MoveNext
    Wend
    
    
    
    'Ponemos el IVA
    If Bases = 0 Then
        Text1(21).Text = ""
    Else
        Text1(21).Text = Format(Bases, FormatoImporte)
    End If
    If Cuotas = 0 Then
        Text1(22).Text = ""
    Else
        Text1(22).Text = Format(Cuotas, FormatoImporte)
    End If
    
    
    'Ponemos, si tiene, el REA
    miRsAux.MoveFirst
    Cuotas = 0
    While Not miRsAux.EOF
        cad = "|" & miRsAux!iva & "|"
        If InStr(1, IVAS_NO, cad) <> 0 Then
            Cuotas = miRsAux!ivas
        End If
        miRsAux.MoveNext
    Wend
    If Cuotas = 0 Then
        Text1(28).Text = ""
    Else
        Text1(28).Text = Format(Cuotas, FormatoImporte)
    End If
    
    
    
    
    'Cerramos El recordset
    miRsAux.Close
    
    
    'Sumamos el IVA devengado
    SumaDevengado
    
    'Sumamos el deducible
    SumaDeducible
    
    'Diferencia y demas
    Liquidacion
End Sub


Private Sub PonerElIva(numero As Integer)
Dim J As Integer
Dim H As Integer
Dim Vcampo As Integer
    If miRsAux.EOF Then Exit Sub
    J = (I - 1) * 3
'    Text1(J).Text = PonerValorImporte(1)
'    Text1(J + 1).Text = PonerValorImporte(0)
'    Text1(J + 2).Text = PonerValorImporte(2)
    
    For H = 0 To 2
        Select Case H
        Case 0
            Vcampo = 1
        Case 1
            Vcampo = 0
        Case Else
            Vcampo = 2
        End Select
        If IsNull(miRsAux.Fields(Vcampo)) Then
            Text1(J + H).Text = ""
        Else
            Text1(J + H).Text = Format(miRsAux.Fields(Vcampo), FormatoImporte)
        End If
    Next H
    

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        Text1Ant = .Text
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub Liquidacion()
Dim Porcentaje As Currency

    Importe = 0
    Importe = CCur(Text1(33).Tag) - CCur(Text1(34).Tag)
    Text1(29).Text = Format(Importe, FormatoImporte)
    Text1(29).Tag = Importe
    
    'Atribuible al estado
    Porcentaje = 100
    If Text1(31).Text <> "" Then Porcentaje = CCur(Text1(31).Text)
    Importe = (Importe * Porcentaje) / 100
    Text1(30).Text = Format(Importe, FormatoImporte)
    
    
    'Periodos anteriores
    Importe = Importe + ImporteFormateado(Text1(32).Text)
    
    'Entregas intracomunitarias
    Importe = Importe + ImporteFormateado(Text1(35).Text)

    Text1(36).Tag = Importe
    Text1(36).Text = Format(Abs(Importe), FormatoImporte)
    
    
    'Si es negativo es devolver
    If Importe < 0 Then
        'A devolver
        Label3(16).Caption = "A DEVOLVER"
    Else
        Label3(16).Caption = "A INGRESAR"
    End If
End Sub


Private Sub SumaDevengado()

    Importe = 0
    Importe = Importe + ImporteFormateado(Text1(2).Text)
    Importe = Importe + ImporteFormateado(Text1(5).Text)
    Importe = Importe + ImporteFormateado(Text1(8).Text)
    Importe = Importe + ImporteFormateado(Text1(11).Text)
    Importe = Importe + ImporteFormateado(Text1(14).Text)
    Importe = Importe + ImporteFormateado(Text1(17).Text)
    Importe = Importe + ImporteFormateado(Text1(20).Text)
    Text1(33).Text = Format(Importe, FormatoImporte)
    Text1(33).Tag = Importe
End Sub


Private Sub SumaDeducible()
    
    Importe = 0
    Importe = Importe + ImporteFormateado(Text1(22).Text)
    Importe = Importe + ImporteFormateado(Text1(24).Text)
    Importe = Importe + ImporteFormateado(Text1(26).Text)
    Importe = Importe + ImporteFormateado(Text1(28).Text)
    Importe = Importe + ImporteFormateado(Text1(27).Text)
    
    Text1(34).Text = Format(Importe, FormatoImporte)
    Text1(34).Tag = Importe
End Sub

Private Sub Text1_LostFocus(Index As Integer)
With Text1(Index)
    If .Text = Text1Ant Then Exit Sub
    
    .Text = Trim(.Text)
    If .Text <> "" Then
        If Not IsNumeric(.Text) Then
            MsgBox "El campo debe ser numérico: " & .Text, vbExclamation
            .Text = ""
            .SetFocus
            Exit Sub
        End If
        
        
        'FORMATEAMOS EL NUMERO
        If InStr(1, .Text, ",") Then
            cad = .Text
            'Quitamos los puntos
            Do
                I = InStr(1, cad, ".")
                If I > 0 Then cad = Mid(cad, 1, I - 1) & Mid(cad, I + 1)
            Loop Until I = 0
            .Text = cad
        Else
            'Lo formateamos
            .Text = TransformaPuntosComas(.Text)
        End If
        .Text = Format(.Text, FormatoImporte)
        
    
    End If
    
    Select Case Index
    Case 2, 5, 8, 11, 14, 17, 20
        SumaDevengado
        Liquidacion
    Case 22, 24, 26, 27, 28
        SumaDeducible
        Liquidacion
    Case 31, 35
        Liquidacion
    End Select
End With

End Sub


'Cojera los importes y los formateara para los programitas de hacineda
Private Sub GeneraCadenaImportes()
Dim Impor As Currency


    'En devuelveimporte
    ' Tipo 0:   11 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales

    For I = 1 To 3
        'Hacemos los IVAS devengados
        DevuelveImporte ((I - 1) * 3), 0
        DevuelveImporte ((I - 1) * 3) + 1, 1
        DevuelveImporte ((I - 1) * 3) + 2, 0
    Next I

    
    'Los recargos
    For I = 4 To 6
        DevuelveImporte ((I - 1) * 3), 0
        DevuelveImporte ((I - 1) * 3) + 1, 2
        DevuelveImporte ((I - 1) * 3) + 2, 0
    Next I
    
    'Adquisiciones intra
    DevuelveImporte 9, 0  'base
    DevuelveImporte 9, 0  'cuota
    
    'total
    DevuelveImporte 33, 0
    
    'Campo en el fichero presentacion telematica: 42
    
    '------------------------------------------------------------------------
    '------------------------------------------------------------------------
    'DEDUCIBLE
    DevuelveImporte 22, 0  'operaciones interiores
    DevuelveImporte 24, 0  'importaciones
    DevuelveImporte 26, 0  'adqisiciones intracom
    DevuelveImporte 28, 0  'Regimen especial
    DevuelveImporte 27, 0  'Regularizacion inversiones
    
    'total a deducir
    DevuelveImporte 34, 0  'cuota
    
    
    'Diferencia
    DevuelveImporte 29, 0  'base
    
    'Atribuible a la admon del estado
    DevuelveImporte 31, 3  '%
    DevuelveImporte 30, 0  'base
    
    'A compensar
    DevuelveImporte 32, 0  'base
    
    'Entregas intracomunitarias
    DevuelveImporte 35, 0  'base
    
    'Diputacion foral
    cad = cad & "0000000000000"
    
    'Total
    DevuelveImporte 36, 0  'base

    
       
End Sub


'Desde un text box
Private Sub DevuelveImporte(Indice As Integer, Tipo As Byte)
Dim J As Integer
Dim Aux As String
Dim Resul As String

    Resul = ""
    If Text1(Indice).Text = "" Then
        Importe = 0
        Aux = "0"
    Else
        Aux = Text1(Indice).Text
        Do
            J = InStr(1, Aux, ".")
            If J > 0 Then Aux = Mid(Aux, 1, J - 1) & Mid(Aux, J + 1)
        Loop Until J = 0
        Importe = CCur(Aux)
        If Importe < 0 Then
            Aux = ""
            Resul = "N"
            Importe = Abs(Importe)
        Else
            Aux = "0"
        End If
        Importe = Importe * 100
        Importe = Int(Importe)
    End If
    
    'Tipo sera la mascara.
    ' Tipo 0:   11 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales
    Select Case Tipo
    Case 1
        Aux = Aux & "000"
    Case 2
        Aux = Aux & "00"
    Case 3
        Aux = Aux & "0000"
    Case Else
        Aux = Aux & "000000000000"
    End Select
    
    cad = cad & Resul & Format(Importe, Aux)
        
End Sub




'-------------------------------------------------------------
'-------------------------------------------------------------
Private Sub GuardarComo()

    On Error GoTo EGuardarComo
    
    'Una vez generado el archvio
    'App.path & "\Hacienda\mod300\miIVA.txt"
    
    cd1.ShowSave
    cad = cd1.FileName
    If cad <> "" Then
        FileCopy App.path & "\Hacienda\mod300\miIVA.txt", cad
    End If
    Exit Sub
EGuardarComo:
    MuestraError Err.Number
End Sub




'-------------------------------------------------------------------
'-------------------------------------------------------------------
'Vamos a ver para los IVA's si hay alguno que se debe quitar.
'  El cero(0%) seguro
'  Ademas quitaremos si hay algun REA
'
'
Private Function BuscarIvasADescartar() As String

    cad = "Select * from tiposiva where tipodiva=3"   'Tipo de IVA: REA
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "|"
    While Not miRsAux.EOF
        cad = cad & miRsAux!porceiva & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    espera 0.1
    BuscarIvasADescartar = cad
End Function
