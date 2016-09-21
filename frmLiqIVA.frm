VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLiqIVA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación del IVA."
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   Icon            =   "frmLiqIVA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmLiqIVA.frx":030A
      Left            =   5400
      List            =   "frmLiqIVA.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   6360
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8190
      TabIndex        =   67
      Top             =   6300
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9360
      TabIndex        =   65
      Top             =   6300
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
      Top             =   5760
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   35
      Left            =   7560
      TabIndex        =   54
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   5220
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
      Top             =   3060
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
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   32
      Left            =   9000
      TabIndex        =   48
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   4770
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   31
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   47
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   30
      Left            =   9000
      TabIndex        =   45
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   4320
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   29
      Left            =   9000
      TabIndex        =   43
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   3870
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
      Top             =   2520
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
      Top             =   2070
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
      Top             =   1620
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
      Top             =   1620
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
      Top             =   1170
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
      Top             =   1170
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
      Top             =   720
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
      Top             =   720
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   20
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   27
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4035
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   19
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   26
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   4035
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   18
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   25
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   4035
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   17
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   24
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   3585
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   16
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   23
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   3585
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   15
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   22
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   14
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   21
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   3150
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   13
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   20
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   3150
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   12
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   19
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   3150
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   11
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   15
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   5130
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   10
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   14
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   5130
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   9
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   13
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   5130
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   8
      Left            =   2700
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "vv"
      Top             =   1185
      Width           =   1455
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
      TabIndex        =   69
      Top             =   6240
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
      TabIndex        =   68
      Top             =   6420
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
      TabIndex        =   66
      Top             =   5760
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
      Top             =   3150
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
      Top             =   5850
      Width           =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   7470
      X2              =   10530
      Y1              =   5670
      Y2              =   5670
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
      Top             =   4860
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
      Top             =   4860
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
      Top             =   4860
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
      Top             =   4590
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
      Top             =   2610
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
      X1              =   6120
      X2              =   10440
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Entregas intracomunitarias"
      Height          =   195
      Index           =   8
      Left            =   4410
      TabIndex        =   55
      Top             =   5280
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
      Left            =   9630
      TabIndex        =   53
      Top             =   450
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
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cuotas a compensar de  periodos anteriores"
      Height          =   195
      Index           =   7
      Left            =   4410
      TabIndex        =   49
      Top             =   4830
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "Atribuible a la admon del estado"
      Height          =   195
      Index           =   6
      Left            =   4410
      TabIndex        =   46
      Top             =   4380
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
      Left            =   4410
      TabIndex        =   44
      Top             =   3930
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
      Left            =   4410
      TabIndex        =   42
      Top             =   3420
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   7560
      X2              =   10530
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Regularización inversiones"
      Height          =   195
      Index           =   4
      Left            =   4410
      TabIndex        =   40
      Top             =   2580
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Compensacion régimen especial A.G. y P."
      Height          =   195
      Index           =   3
      Left            =   4410
      TabIndex        =   38
      Top             =   2130
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Adquisiciones intracomunitarias"
      Height          =   195
      Index           =   2
      Left            =   4410
      TabIndex        =   35
      Top             =   1680
      Width           =   2385
   End
   Begin VB.Label Label1 
      Caption         =   "En importaciones"
      Height          =   285
      Index           =   1
      Left            =   4410
      TabIndex        =   32
      Top             =   1230
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "En operaciones interiores"
      Height          =   195
      Index           =   0
      Left            =   4410
      TabIndex        =   29
      Top             =   780
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
      Y1              =   5670
      Y2              =   5670
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
      Top             =   2880
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
      Top             =   2880
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
      Top             =   2880
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
Public Consolidado As String 'Si viene a "" no hay mas que su proipa empresa
Public FechaIMP As String
Public Modelo As Byte
    '0.- 300
    '1.- 320
    '2.- 330
    '3.- 332
Dim PrimeraVez As Boolean
Dim Cad As String
Dim I As Integer
Dim Importe As Currency

Dim Text1Ant As String



Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Dim A_Compensar As Boolean
    'Imprimir
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione una forma de impresión", vbExclamation
        Exit Sub
    End If
    
    
    Select Case Combo1.ItemData(Combo1.ListIndex)
    Case 3, 4
    
            'Si el importe es negativo tendriamos que preguntar si quiere realizar
            'compensacion o ingreso/devolucion
            If CCur(CCur(Text1(36).Tag)) < 0 Then
                Cad = "Importe a devolver / compensar." & vbCrLf & vbCrLf & _
                    "¿ Desea que sea a compensar ?"
                I = MsgBox(Cad, vbQuestion + vbYesNoCancel)
                If I = vbCancel Then Exit Sub
                
                A_Compensar = (I = vbYes)
            End If
            'Generamos la cadena con los importes a mostrar
            Cad = ""
            GeneraCadenaImportes
 
            'Ahora enviamos a generar fichero IVA. EL TAG
            If GenerarFicheroIVA(Cad, CCur(Text1(36).Tag), CDate(FechaIMP), Periodo, A_Compensar) Then
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
         Cad = "CampoSeleccion= """ & Text1Ant & """|"
         If vParam.periodos = 0 Then
            Text1Ant = "Trimestre "
         Else
            Text1Ant = "MES "
         End If
         'Inicio
         Cad = Cad & "TInicio= """ & Text1Ant & "inicio""|"
         Cad = Cad & "Inicio= """ & RecuperaValor(Periodo, 1) & """|"
         'Fin
         Cad = Cad & "TFin= """ & Text1Ant & "fin""|"
         Cad = Cad & "Fin= """ & RecuperaValor(Periodo, 2) & """|"
         'Anyo
         Cad = Cad & "Anyo= """ & RecuperaValor(Periodo, 3) & """|"
         'Empresas
         If Consolidado = "" Then
            Cad = Cad & "Empresas= """ & Consolidado & """|"
         Else
            Cad = Cad & "Empresas= " & Consolidado & "|"
         End If
         If Combo1.ListIndex = 0 Then
            I = 30
         Else
            I = 63
         End If
         Cad = Cad & "FechaIMP= """ & FechaIMP & """|"
                 
         'Nuevo del 29 Diciembre 2004
         ' Mi santo, por cierto
         'Es cuotas compensables de otros periodos
         If Text1(32).Text = "" Then
            Cad = Cad & "Compensable= 0|"
         Else
            Cad = Cad & "Compensable= " & TransformaComasPuntos(CStr(ImporteFormateado(Text1(32).Text))) & "|"
         End If
         
         With frmImprimir
                .OtrosParametros = Cad
                Cad = vEmpresa.nomempre
                If Consolidado <> "" Then vEmpresa.nomempre = "CONSOLIDADO"
                .NumeroParametros = 9
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = I
                .Show vbModal
                vEmpresa.nomempre = Cad
         End With
    End Select
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        'Cargamos los datos del IVA
        PonDatosIVA
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim Detallado As Byte
    PrimeraVez = True
    Limpiar Me
    
    Cad = RecuperaValor(Periodo, 4)
    If Cad = "1" Then
        Detallado = 1
    Else
        Detallado = 0
    End If
    
    'Ahora vamos a ir generando periodo
    Cad = ""
    For I = 1 To 3
        Cad = Cad & RecuperaValor(Periodo, I) & "|"
    Next I
    Periodo = Cad
    'Informe detallado
    Cad = Detallado
    
    
    'Fecha
    If FechaIMP = "" Then FechaIMP = Format(Now, "dd/mm/yyyy")
    
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
    If Val(Cad) > 0 Then
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

'Del objeto miRSAUX, el field vcampo le pondra el valor
Private Function PonerValorImporte(Vcampo As Integer) As String
Dim Valor As Currency
    If IsNull(miRsAux.Fields(Vcampo)) Then
        Valor = 0
    Else
        Valor = miRsAux.Fields(Vcampo)
    End If
    If Valor = 0 Then
        PonerValorImporte = ""
    Else
        PonerValorImporte = Format(Valor, FormatoImporte)
    End If
End Function

Private Sub PonDatosIVA()

    '       cliprov     0- Facturas clientes
    '                   1- RECARGO EQUIVALENCIA !!nuevo
    '                   2- Facturas proveedores
    '                   3- libre                !!nuevo
    '                   4- IVAS no deducible
    '                   5- Facturas NO DEDUCIBLES



    'IVA DEVENGADO
    Set miRsAux = New ADODB.Recordset
    Cad = "select iva,sum(bases),sum(ivas) from Usuarios.zliquidaiva where "
    Cad = Cad & " codusu=" & vUsu.Codigo & " and cliente= 0 group by iva order by iva ASC"
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    I = 1
    While Not miRsAux.EOF
        If miRsAux.Fields(0) <> 0 Then
            If I < 4 Then
                PonerElIva I
                I = I + 1
            End If
            
        Else
            'SI EL IVA ES EL 0, pongo la base en ENTREGAS intracom
            Text1(35).Text = Format(miRsAux.Fields(1), FormatoImporte)
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
            
    
    'IVA DEDUCIBLE
    Cad = "select sum(bases),sum(ivas) from Usuarios.zliquidaiva where codusu=" & vUsu.Codigo
    Cad = Cad & " and cliente= 2"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Text1(21).Text = PonerValorImporte(0)
        Text1(22).Text = PonerValorImporte(1)
    End If
    miRsAux.Close
    
    
    'PONGO RECARGO DE EQUIVA EN CLIENTES
    '-----------------------------------
    Cad = "select iva,sum(bases),sum(ivas) from Usuarios.zliquidaiva where "
    Cad = Cad & " codusu=" & vUsu.Codigo & " and cliente= 1 group by iva order by iva ASC"
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    I = 5
    While Not miRsAux.EOF
        If miRsAux.Fields(0) <> 0 Then
            If I < 7 Then
                PonerElIva I
                I = I + 1
                
            Else
                MsgBox "Mas recargos de equivalencia que los tres permitidos", vbExclamation
            End If
            
        Else
            'SI EL IVA ES EL 0, pongo la base en ENTREGAS intracom
            Text1(35).Text = Format(miRsAux.Fields(1), FormatoImporte)
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    'Ahora
    PonerIntraComYCampo
    
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
    Importe = Importe - ImporteFormateado(Text1(32).Text)
    
    'Entregas intracomunitarias    13 Junio 05
'    Importe = Importe + ImporteFormateado(Text1(35).Text)


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
            Cad = .Text
            'Quitamos los puntos
            Do
                I = InStr(1, Cad, ".")
                If I > 0 Then Cad = Mid(Cad, 1, I - 1) & Mid(Cad, I + 1)
            Loop Until I = 0
            .Text = Cad
        Else
            'Lo formateamos
            .Text = TransformaPuntosComas(.Text)
        End If
        .Text = Format(.Text, FormatoImporte)
        If Index = 32 Then
            If CCur(.Text) < 0 Then
                MsgBox "Cuotas a compensar se escriben en positivo", vbExclamation
                .Text = Abs(CCur(.Text))
                .SetFocus
                Exit Sub
            End If
        End If
    End If
    
    Select Case Index
    Case 2, 5, 8, 11, 14, 17, 20
        SumaDevengado
        Liquidacion
    Case 22, 24, 26, 27, 28
        SumaDeducible
        Liquidacion
    Case 31, 32, 35
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
    For I = 0 To 2
        DevuelveImporte ((3 * I) + 12), 0
        DevuelveImporte (I * 3) + 13, 2
        DevuelveImporte ((I * 3)) + 14, 0
    Next I
    
    'Adquisiciones intra
    DevuelveImporte 9, 0  'base
    DevuelveImporte 11, 0  'cuota
    
    'total cuota devengada
    DevuelveImporte 33, 0
    
    
    
    '------------------------------------------------------------------------
    '------------------------------------------------------------------------
    'DEDUCIBLE
    DevuelveImporte 22, 0  'operaciones interiores
    DevuelveImporte 24, 0  'importaciones
    DevuelveImporte 26, 0  'adqisiciones intracom
    DevuelveImporte 28, 0  'Regimen especial
    DevuelveImporte 27, 0  'Regularizacion inversiones
    
    'Nuevo campo en 2008
    Cad = Cad & "0000000000000"
    
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
    Cad = Cad & "0000000000000"
    
    'Total
    DevuelveImporte 36, 0  'base

    'Nuevo campo en 2008.
    Cad = Cad & "0000000000000"
    
       
End Sub

'
'Desde un text box
Private Sub DevuelveImporte(Indice As Integer, Tipo As Byte)
Dim J As Integer
Dim Aux As String
Dim Resul As String
Dim Valor As String
    
    Valor = Text1(Indice).Text
    If Modelo = 0 Then
        If Indice = 36 Then
            'Si es el importe, miro el tag
            Valor = Text1(Indice).Tag
        End If
    End If
    
    
    Resul = ""
    If Valor = "" Then
        Importe = 0
        Aux = "0"
    Else
        Aux = Valor
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
    
    Cad = Cad & Resul & Format(Importe, Aux)
        
End Sub




'-------------------------------------------------------------
'-------------------------------------------------------------
Private Sub GuardarComo()

    On Error GoTo EGuardarComo
    
    'Una vez generado el archvio
    'App.path & "\Hacienda\mod300\miIVA.txt"
    If DirectorioEAT Then cd1.InitDir = "C:\AEAT"
    
    cd1.ShowSave
    Cad = cd1.FileName
    If Cad <> "" Then
        FileCopy App.path & "\Hacienda\mod300\miIVA.txt", Cad
    End If
    Exit Sub
EGuardarComo:
    MuestraError Err.Number
End Sub


Private Sub PonerIntraComYCampo()


    On Error GoTo EPonerIntraComYCampo
    
    Set miRsAux = New ADODB.Recordset
    
    'Cojo el REA
    Cad = "select sum(acumperd) as C1 ,sum(acumperh) as C2 from tmpctaexplotacioncierre "
    Cad = Cad & " where codusu =" & vUsu.Codigo & " and cta like 'a%'"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux!C2) Then
            Text1(28).Text = Format(miRsAux!C2, FormatoImporte)
            'EL IVA
            Importe = ImporteFormateado(Text1(22).Text)
            Importe = Importe - miRsAux!C2
            Text1(22).Text = Format(Importe, FormatoImporte)
            
            'IMPONIBLE
            Importe = ImporteFormateado(Text1(21).Text)
            Importe = Importe - miRsAux!C1
            Text1(21).Text = Format(Importe, FormatoImporte)
            'Imp1 = miRsAux!C1
        End If
    End If
    miRsAux.Close
    
    
    
    'El dificil, las intracom de clientes
    Cad = "select cta,sum(acumperd) as C1 ,sum(acumperh) as C2 from tmpctaexplotacioncierre "
    Cad = Cad & " where codusu =" & vUsu.Codigo & " and cta like 'c%'"
    
    'QUITAMOS LOS DEL 0%
    Cad = Cad & " and not (cta like '%0000')"
    
    Cad = Cad & "  group by cta"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        While Not miRsAux.EOF
            Cad = miRsAux!Cta
            Cad = Right(Cad, 4)  'Los 4 ultimos es el % de iva * 100
            Importe = Val(Cad) / 100
            
            Cad = Format(Importe, FormatoImporte)
            I = 0
            If Cad = Text1(1).Text Then
                I = 1
            Else
                If Cad = Text1(4).Text Then
                    I = 4
                Else
                    If Cad = Text1(7).Text Then I = 7
                End If
            End If
            
            If I = 0 Then
            
                'Si el % el el 0% entonces siginifica que son ventas al extranjero,
                'NO a la CEE, con lo caul no doy mensaje de error alguno
                If Val(Cad) = 0 Then
                
                    MsgBox "Error situan datos facturas , % <>0", vbExclamation
                
                
                Else
            
            
                    Cad = "%: " & Cad & vbCrLf & miRsAux!Cta & vbCrLf & miRsAux!C1 & " - " & miRsAux!C2 & vbCrLf
                    Cad = "Error en tipo IVA para facturas con INTRACOM: " & vbCrLf & vbCrLf & Cad
                    Cad = Cad & vbCrLf & vbCrLf & "Consulte soporte técnico"
                    MsgBox Cad, vbExclamation
                    miRsAux.Close
                    Set miRsAux = Nothing
                    Exit Sub
                End If
                
                
            Else
            
                'Ya sabemos que tipo de IVA es
                Importe = ImporteFormateado(Text1(I - 1).Text)
                Importe = Importe - miRsAux!C1
                Text1(I - 1).Text = Format(Importe, FormatoImporte)
                
                'INTRACOM
                Importe = ImporteFormateado(Text1(9).Text)
                Importe = Importe + miRsAux!C1
                Text1(9).Text = Format(Importe, FormatoImporte)
                
                Importe = ImporteFormateado(Text1(I + 1).Text)
                Importe = Importe - miRsAux!C2
                Text1(I + 1).Text = Format(Importe, FormatoImporte)
                
                'INTRACOM
                Importe = ImporteFormateado(Text1(11).Text)
                Importe = Importe + miRsAux!C2
                Text1(11).Text = Format(Importe, FormatoImporte)

            End If

            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    
        
        
'
'    'VENTAS al 0%, es decir o al extranjero o al intracom
'    cad = "select cta,sum(acumperd) as C1 ,sum(acumperh) as C2 from tmpctaexplotacioncierre "
'    cad = cad & " where codusu =" & vUsu.Codigo & " and cta like 'c%0000'"
'    cad = cad & "  group by cta"
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not miRsAux.EOF Then
'
'
'    End If
'    miRsAux.Close
    
    
    
    
    'INTRACOM proveedores
    Cad = "select sum(acumperd) as C1 ,sum(acumperh) as C2 from tmpctaexplotacioncierre "
    Cad = Cad & " where codusu =" & vUsu.Codigo & " and cta like 'p%'"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux!C2) Then
            Text1(26).Text = Format(miRsAux!C2, FormatoImporte)
            'EL IVA
            Importe = ImporteFormateado(Text1(22).Text)
            Importe = Importe - miRsAux!C2
            Text1(22).Text = Format(Importe, FormatoImporte)
            
            'IMPONIBLE
            Importe = ImporteFormateado(Text1(21).Text)
            Importe = Importe - miRsAux!C1
            Text1(21).Text = Format(Importe, FormatoImporte)
            
            Text1(25).Text = Format(miRsAux!C1, FormatoImporte)
        End If
    End If
    miRsAux.Close
    
    
    
    
    Set miRsAux = Nothing
    Exit Sub
EPonerIntraComYCampo:
    MuestraError Err.Number, "Poner IntraCom y Campo"
    Set miRsAux = Nothing

End Sub
