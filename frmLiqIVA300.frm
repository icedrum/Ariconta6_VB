VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLiqIVA2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación del IVA. "
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   Icon            =   "frmLiqIVA300.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   44
      Left            =   7560
      TabIndex        =   84
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   6360
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   43
      Left            =   7560
      TabIndex        =   27
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4620
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   42
      Left            =   6000
      TabIndex        =   26
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4620
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   40
      Left            =   6000
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   3240
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   41
      Left            =   7560
      TabIndex        =   23
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   3240
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   38
      Left            =   6000
      TabIndex        =   18
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1875
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   39
      Left            =   7560
      TabIndex        =   19
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1875
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   37
      Left            =   11400
      TabIndex        =   36
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   4800
      Width           =   1395
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   10470
      TabIndex        =   38
      Top             =   7740
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmLiqIVA300.frx":000C
      Left            =   7200
      List            =   "frmLiqIVA300.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   11880
      TabIndex        =   39
      Top             =   7740
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   36
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   37
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   7117
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   35
      Left            =   11400
      TabIndex        =   35
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   3960
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   34
      Left            =   7530
      Locked          =   -1  'True
      TabIndex        =   30
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   7117
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   33
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   59
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   7117
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   32
      Left            =   11400
      TabIndex        =   34
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   3240
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   31
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   32
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   30
      Left            =   11400
      TabIndex        =   33
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   2160
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   29
      Left            =   11400
      TabIndex        =   31
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "123.156.253.23"
      Top             =   1320
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   27
      Left            =   7560
      TabIndex        =   29
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   5880
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   28
      Left            =   7560
      TabIndex        =   28
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   5400
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   26
      Left            =   7530
      TabIndex        =   25
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4200
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   25
      Left            =   6000
      TabIndex        =   24
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4200
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   24
      Left            =   7560
      TabIndex        =   21
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   2850
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   23
      Left            =   6000
      TabIndex        =   20
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   2850
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   22
      Left            =   7560
      TabIndex        =   17
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1320
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   21
      Left            =   6000
      TabIndex        =   16
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   1320
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   20
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   51
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4635
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   19
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   50
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   4635
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   18
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   49
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   4635
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   17
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   48
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   4185
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   16
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   47
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   4185
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   15
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   46
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   14
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   45
      Tag             =   "Importe IVA 1|N|S|||cabfact|ti1faccl|#,###,##0.00||"
      Text            =   "Text1"
      Top             =   3750
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   13
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   44
      Tag             =   "Porcentaje IVA 1|N|S|||cabfact|pi1faccl|#0.00||"
      Text            =   "Text1"
      Top             =   3750
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   12
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   43
      Tag             =   "Base imponible 1|N|N|||cabfact|ba1faccl|#,###,###,##0.00||"
      Top             =   3750
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
      Top             =   6090
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
      Top             =   6090
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
      Top             =   6090
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
      Top             =   2550
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
      Top             =   2550
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
      Top             =   2550
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
      Top             =   2100
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
      Top             =   2100
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
      Top             =   2100
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
      Top             =   1665
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
      Top             =   1665
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
      Top             =   1665
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Liquidación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Index           =   2
      Left            =   9240
      TabIndex        =   86
      Top             =   210
      Width           =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Reg. por aplicacion de % def. de prorrata"
      Height          =   195
      Index           =   13
      Left            =   4200
      TabIndex        =   85
      Top             =   6360
      Width           =   3285
   End
   Begin VB.Label Label1 
      Caption         =   "Bienes corrientes"
      Height          =   195
      Index           =   12
      Left            =   4440
      TabIndex        =   83
      Top             =   4320
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Bien de inversion"
      Height          =   195
      Index           =   2
      Left            =   4440
      TabIndex        =   82
      Top             =   4800
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   20
      Left            =   4200
      TabIndex        =   81
      Top             =   3840
      Width           =   3510
   End
   Begin VB.Label Label1 
      Caption         =   "Bienes corrientes"
      Height          =   195
      Index           =   11
      Left            =   4440
      TabIndex        =   80
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Bien de inversion"
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   79
      Top             =   3360
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "IMPORTACIONES"
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
      Index           =   19
      Left            =   4200
      TabIndex        =   78
      Top             =   2400
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "OPERACIONES INTERIORES"
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
      Index           =   18
      Left            =   4200
      TabIndex        =   77
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Bien de inversion"
      Height          =   195
      Index           =   10
      Left            =   4440
      TabIndex        =   76
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Exportaciones y operaciones asimiladas"
      Height          =   195
      Index           =   9
      Left            =   9360
      TabIndex        =   75
      Top             =   4560
      Width           =   3105
   End
   Begin VB.Label Label4 
      Caption         =   "Modelo 330"
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
      TabIndex        =   74
      Top             =   7680
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
      Left            =   6240
      TabIndex        =   73
      Top             =   7860
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
      Left            =   9870
      TabIndex        =   72
      Top             =   7140
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
      Left            =   5640
      TabIndex        =   70
      Top             =   7170
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
      TabIndex        =   69
      Top             =   7170
      Width           =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   9240
      X2              =   12930
      Y1              =   6960
      Y2              =   6960
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
      TabIndex        =   68
      Top             =   5820
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
      TabIndex        =   67
      Top             =   5820
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
      TabIndex        =   66
      Top             =   5820
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
      TabIndex        =   65
      Top             =   5550
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
      TabIndex        =   64
      Top             =   3120
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
      TabIndex        =   63
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Entregas intracomunitarias"
      Height          =   195
      Index           =   8
      Left            =   9330
      TabIndex        =   62
      Top             =   3840
      Width           =   2025
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
      Left            =   8160
      TabIndex        =   61
      Top             =   1080
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
      Left            =   6120
      TabIndex        =   60
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cuotas a compensar de  periodos anteriores"
      Height          =   195
      Index           =   7
      Left            =   9330
      TabIndex        =   58
      Top             =   3000
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "Atribuible a la admon del estado"
      Height          =   195
      Index           =   6
      Left            =   9330
      TabIndex        =   57
      Top             =   1920
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
      Left            =   9360
      TabIndex        =   56
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   4200
      X2              =   8970
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label1 
      Caption         =   "Regularización inversiones"
      Height          =   195
      Index           =   4
      Left            =   4200
      TabIndex        =   55
      Top             =   5880
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Compensacion régimen especial A.G. y P."
      Height          =   195
      Index           =   3
      Left            =   4200
      TabIndex        =   54
      Top             =   5400
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Bienes corrientes"
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   53
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "I.V.A  deducible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Index           =   1
      Left            =   4200
      TabIndex        =   52
      Top             =   210
      Width           =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   3960
      Y1              =   6960
      Y2              =   6960
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
      TabIndex        =   42
      Top             =   3480
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
      TabIndex        =   41
      Top             =   3480
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
      TabIndex        =   40
      Top             =   3480
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
      Top             =   1335
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
      Top             =   1335
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "I.V.A  devengado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   210
      Width           =   2970
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
      Top             =   1335
      Width           =   1455
   End
End
Attribute VB_Name = "frmLiqIVA2"
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
    '4.- 303    Nuevo Enero 2009
Dim PrimeraVez As Boolean
Dim Cad As String
Dim i As Integer
Dim Importe As Currency

Dim Text1Ant As String

Dim ValoresIVA_SinEstarEnPAntalla As Boolean

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Dim Es_A_Compensar As Byte
    'Imprimir
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione una forma de impresión", vbExclamation
        Exit Sub
    Else
        
'            Cad = UCase(Combo1.List(Combo1.ListIndex))
'            If Cad = "INTERNET" Then
'                MsgBox "La agencia tributaria no contempla el envio del fichero de datos", vbExclamation
'                Exit Sub
'            End If
        
    End If
    
    
    Select Case Combo1.ItemData(Combo1.ListIndex)
    Case 3, 4
            If ValoresIVA_SinEstarEnPAntalla Then
                Cad = "Hay valores de iva devengado que no salen en el formulario. Importe liquidacion incorrecto. "
                MsgBox Cad, vbExclamation
                Exit Sub
            End If
    
    
            'Si el importe es negativo tendriamos que preguntar si quiere realizar
            'compensacion o ingreso/devolucion
            If CCur(CCur(Text1(36).Tag)) < 0 Then
                'NEGATIVO
                Cad = "Importe a devolver / compensar." & vbCrLf & vbCrLf & _
                    "¿ Desea que sea a compensar ?"
                i = MsgBox(Cad, vbQuestion + vbYesNoCancel)
                If i = vbCancel Then Exit Sub
                Es_A_Compensar = 0
                If i = vbYes Then Es_A_Compensar = 1
                
            Else
                Cad = "Ingreso por cta banco?" & vbCrLf & vbCrLf
                '
                i = MsgBox(Cad, vbQuestion + vbYesNoCancel)
                If i = vbCancel Then Exit Sub
                Es_A_Compensar = 2
                If i = vbYes Then Es_A_Compensar = 3
            End If
    
    
    
            'Generamos la cadena con los importes a mostrar
            Cad = ""
            GeneraCadenaImportes
 
            'Ahora enviamos a generar fichero IVA
            If GenerarFicheroIVA_303(Cad, CCur(Text1(36).Tag), CDate(FechaIMP), Periodo, Es_A_Compensar) Then
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
            i = 30
         Else
            i = 63
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
                .opcion = i
                .Show vbModal
                vEmpresa.nomempre = Cad
         End With
    End Select
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresIVA_SinEstarEnPAntalla = False
        'Cargamos los datos del IVA
        PonDatosIVA
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim Detallado As Byte

    Me.Icon = frmPpal.Icon

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
    For i = 1 To 3
        Cad = Cad & RecuperaValor(Periodo, i) & "|"
    Next i
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
    Case 4
        Label4.Caption = Label4.Caption & "303"
        
        
    Case Else
        Label4.Caption = "Error pasando modelo"
    End Select
    Combo1.Clear
    Combo1.AddItem "Borrador"
    Combo1.ItemData(Combo1.NewIndex) = 1
    If Val(Cad) > 0 Then
        Combo1.AddItem "Informe detallado"
        Combo1.ItemData(Combo1.NewIndex) = 2
        i = 1
    Else
        i = 0
    End If
    
    'Solo el modelo 300 tiene impresion hacienda
    If Modelo = 0 Then
        Combo1.AddItem "Modelo hacienda"
        Combo1.ItemData(Combo1.NewIndex) = 3
    End If
    Combo1.AddItem "Internet"
    Combo1.ItemData(Combo1.NewIndex) = 4
    Combo1.ListIndex = i
End Sub

'Del objeto miRSAUX, el field vcampo le pondra el valor
Private Function PonerValorImporte(vCampo As Integer) As String
Dim Valor As Currency
    If IsNull(miRsAux.Fields(vCampo)) Then
        Valor = 0
    Else
        Valor = miRsAux.Fields(vCampo)
    End If
    If Valor = 0 Then
        PonerValorImporte = ""
    Else
        PonerValorImporte = Format(Valor, FormatoImporte)
    End If
End Function

Private Sub PonDatosIVA()
    'IVA DEVENGADO
    
    Set miRsAux = New ADODB.Recordset
    Cad = "select iva,sum(bases),sum(ivas) from Usuarios.zliquidaiva where "
    Cad = Cad & " codusu=" & vUsu.Codigo & " and cliente= 0 group by iva order by iva ASC"
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    i = 1
    Cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(0) <> 0 Then
            If i < 4 Then
                PonerElIva i
                i = i + 1
            Else
                'HAY MAS DE TRES TIPOS DE IVA
                Cad = Cad & "IVA " & miRsAux.Fields(0) & "%      " & miRsAux.Fields(1) & "       " & miRsAux.Fields(2) & vbCrLf
            End If
        Else
    
            'SI EL IVA ES EL 0, pongo la base en ENTREGAS intracom
            Text1(37).Text = Format(miRsAux.Fields(1), FormatoImporte)
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Cad <> "" Then
        Cad = "Mas de tres tipos de ivas" & vbCrLf & "Los siguientes ivas no salen en el fomulario" & vbCrLf & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        ValoresIVA_SinEstarEnPAntalla = True
    End If
        
            
     'PONGO RECARGO DE EQUIVA EN CLIENTES
    '-----------------------------------
    
    Cad = "select iva,sum(bases),sum(ivas) from Usuarios.zliquidaiva where "
    Cad = Cad & " codusu=" & vUsu.Codigo & " and cliente= 1 group by iva order by iva ASC"
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    i = 5
    While Not miRsAux.EOF
        If miRsAux.Fields(0) <> 0 Then
            If i < 8 Then
                PonerElIva i
                i = i + 1
                
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
           
            
            
            
            
            
            
            
            
            
            
    
    'IVA DEDUCIBLE
    Cad = "select sum(bases),sum(ivas) from Usuarios.zliquidaiva where codusu=" & vUsu.Codigo
    Cad = Cad & " and cliente= 2"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Text1(21).Text = PonerValorImporte(0)
        Text1(22).Text = PonerValorImporte(1)
    End If
    miRsAux.Close
    
    
    'IVA deducible en BIEN DE inversion
    Cad = "select sum(bases),sum(ivas) from Usuarios.zliquidaiva where codusu=" & vUsu.Codigo
    Cad = Cad & " and cliente= 6"  'Bien de inversion
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Text1(38).Text = PonerValorImporte(0)
        Text1(39).Text = PonerValorImporte(1)
    End If
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
Dim vCampo As Integer
    If miRsAux.EOF Then Exit Sub
    
    
    
    J = (i - 1) * 3
'    Text1(J).Text = PonerValorImporte(1)
'    Text1(J + 1).Text = PonerValorImporte(0)
'    Text1(J + 2).Text = PonerValorImporte(2)
    
    For H = 0 To 2
        Select Case H
        Case 0
            vCampo = 1
        Case 1
            vCampo = 0
        Case Else
            vCampo = 2
        End Select
        If IsNull(miRsAux.Fields(vCampo)) Then
            Text1(J + H).Text = ""
        Else
            Text1(J + H).Text = Format(miRsAux.Fields(vCampo), FormatoImporte)
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
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
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
     'Dedu por bien de inversion
    Importe = Importe + ImporteFormateado(Text1(39).Text)
    Importe = Importe + ImporteFormateado(Text1(41).Text)
    Importe = Importe + ImporteFormateado(Text1(43).Text)
    
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
                i = InStr(1, Cad, ".")
                If i > 0 Then Cad = Mid(Cad, 1, i - 1) & Mid(Cad, i + 1)
            Loop Until i = 0
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
    Case 22, 24, 26, 27, 28, 39, 41, 43
        SumaDeducible
        Liquidacion
    Case 31, 32, 35
        Liquidacion
    End Select
End With

End Sub


'Cojera los importes y los formateara para los programitas de hacineda
Private Sub GeneraCadenaImportes()



    'En devuelveimporte
    ' Tipo 0:   11 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales

    For i = 1 To 3
        'Hacemos los IVAS devengados
        DevuelveImporte ((i - 1) * 3), 0
        DevuelveImporte ((i - 1) * 3) + 1, 3
        DevuelveImporte ((i - 1) * 3) + 2, 0
    Next i

    
    'Los recargos
    For i = 0 To 2
        DevuelveImporte ((3 * i) + 12), 0
        DevuelveImporte (i * 3) + 13, 3
        DevuelveImporte ((i * 3)) + 14, 0
    Next i
    
    'Adquisiciones intra
    DevuelveImporte 9, 0  'base
    DevuelveImporte 11, 0  'cuota
    
    'total
    DevuelveImporte 33, 0
    

    
    '------------------------------------------------------------------------
    '------------------------------------------------------------------------
    'DEDUCIBLE
    'operaciones interiores
    DevuelveImporte 21, 0
    DevuelveImporte 22, 0
    
    'operaciones interiores BIENES INVERSION
    DevuelveImporte 38, 0
    DevuelveImporte 39, 0
    
    'importaciones
    DevuelveImporte 23, 0
    DevuelveImporte 24, 0
    
    'importaciones BIEN INVERSION
    DevuelveImporte 40, 0
    DevuelveImporte 41, 0
    
    
    'adqisiciones intracom
    DevuelveImporte 25, 0
    DevuelveImporte 26, 0
    
    'adqisiciones intracom BIEN INVERSION
    DevuelveImporte 42, 0
    DevuelveImporte 43, 0
    
    DevuelveImporte 28, 0  'Regimen especial
    DevuelveImporte 27, 0  'Regularizacion inversiones
    DevuelveImporte 44, 0  'Regularizacion por aplicacion del porcentaje def de prorrata
    
    
    'total a deducir
    DevuelveImporte 34, 0  'cuota
    
    
    'Diferencia
    DevuelveImporte 29, 0  'base
    
    'Atribuible a la admon del estado
    DevuelveImporte 31, 3  '%
    DevuelveImporte 30, 0  'base
    
    'A compensar de otros periodos
    DevuelveImporte 32, 0  'base
    
    'Entregas intracomunitarias
    DevuelveImporte 35, 0  'base
    
    'Exportaciones y asimiladas
    DevuelveImporte 37, 0  'base
    
    
    'DE estos dos NO hay text
    '---------------------
    'Op no sujetas o con conversion del sujeto pasivo
    Cad = Cad & String(17, "0")
    'Diputacion foral
    Cad = Cad & String(17, "0")
    
    'Total
    
    Text1Ant = Text1(36).Text      'por si acaso es negativo
    Text1(36).Text = Format(CCur(Text1(36).Tag), FormatoImporte)
    DevuelveImporte 36, 0  'base
    Text1(36).Text = Text1Ant
    Text1Ant = ""
    
       
End Sub


'Desde un text box
Private Sub DevuelveImporte(indice As Integer, Tipo As Byte)
Dim J As Integer
Dim AUx As String
Dim Resul As String

    Resul = ""
    If Text1(indice).Text = "" Then
        Importe = 0
        AUx = "0"
    Else
        AUx = Text1(indice).Text
        Do
            J = InStr(1, AUx, ".")
            If J > 0 Then AUx = Mid(AUx, 1, J - 1) & Mid(AUx, J + 1)
        Loop Until J = 0
        Importe = CCur(AUx)
        If Importe < 0 Then
            AUx = ""
            Resul = "N"
            Importe = Abs(Importe)
        Else
            AUx = "0"
        End If
        Importe = Importe * 100
        Importe = Int(Importe)
    End If
    
    'Tipo sera la mascara.
    ' si Modelo<>303
        ' Tipo 0:   11 enteros y 2 decimales
        'Else
        ' Tipo 0:   15 enteros y 2 decimales
    ' Tipo 1:   2 ente y 2 decimales
    ' Tipo 2:   1 entero y 2 decimales
    ' tipo 3:   3 enetero y dos decimales
    Select Case Tipo
    Case 1
        AUx = AUx & "000"
    Case 2
        AUx = AUx & "00"
    Case 3
        AUx = AUx & "0000"
    Case Else
        If Modelo = 4 Then
            AUx = AUx & String(16, "0")  '15 enteros 2 decima  17-1
        Else
            'Aux = Aux & "000000000000"
            AUx = AUx & String(10, "0")   '11 enteros 2 decimales  13-1
        End If
    End Select
    
    Cad = Cad & Resul & Format(Importe, AUx)
        
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
        FileCopy App.Path & "\Hacienda\mod300\miIVA.txt", Cad
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
            i = 0
            If Cad = Text1(1).Text Then
                i = 1
            Else
                If Cad = Text1(4).Text Then
                    i = 4
                Else
                    If Cad = Text1(7).Text Then i = 7
                End If
            End If
            
            If i = 0 Then
            
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
                Importe = ImporteFormateado(Text1(i - 1).Text)
                Importe = Importe - miRsAux!C1
                Text1(i - 1).Text = Format(Importe, FormatoImporte)
                
                'INTRACOM
                Importe = ImporteFormateado(Text1(9).Text)
                Importe = Importe + miRsAux!C1
                Text1(9).Text = Format(Importe, FormatoImporte)
                
                Importe = ImporteFormateado(Text1(i + 1).Text)
                Importe = Importe - miRsAux!C2
                Text1(i + 1).Text = Format(Importe, FormatoImporte)
                
                'INTRACOM
                Importe = ImporteFormateado(Text1(11).Text)
                Importe = Importe + miRsAux!C2
                Text1(11).Text = Format(Importe, FormatoImporte)

            End If

            miRsAux.MoveNext
        Wend
    End If
    miRsAux.Close
    
        
        
        
    'VENTAS al 0%, es decir o al extranjero o al intracom
    Cad = "select sum(acumperd) as C1 ,sum(acumperh) as C2 from tmpctaexplotacioncierre "
    Cad = Cad & " where codusu =" & vUsu.Codigo & " and cta like 'c%0000'"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        
        If Not IsNull(miRsAux!C1) Then
            Text1(35).Text = Format(miRsAux!C1, FormatoImporte)
            Importe = ImporteFormateado(Text1(37).Text)
            Importe = Importe - miRsAux!C1
            If Importe < 0 Then
                'MsgBox "Se ha producido un error leyendo los datos VENTAS 0% intracom. Importe menor que 0", vbExclamation
            End If
                Text1(37).Text = Format(Importe, FormatoImporte)
        End If
    End If
    miRsAux.Close
    
    
    
    
    'INTRACOM proveedores. Dos tipos. Normales y bien inversion
    'intracom NORMAL
    Cad = "select sum(acumperd) as C1 ,sum(acumperh) as C2 from tmpctaexplotacioncierre "
    Cad = Cad & " where codusu =" & vUsu.Codigo & " and cta like 'p%' "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux!C2) Then
            Text1(26).Text = Format(miRsAux!C2, FormatoImporte)
            'EL IVA en BI
            Importe = ImporteFormateado(Text1(22).Text)
            Importe = Importe - miRsAux!C2
            Text1(22).Text = Format(Importe, FormatoImporte)
            
            'IMPONIBLE en BI
            Importe = ImporteFormateado(Text1(21).Text)
            Importe = Importe - miRsAux!C1
            Text1(21).Text = Format(Importe, FormatoImporte)
            
            Text1(25).Text = Format(miRsAux!C1, FormatoImporte)
        End If
    End If
    miRsAux.Close
    'Intracom BIEN INVERSION
    Cad = "select sum(acumperd) as C1 ,sum(acumperh) as C2 from tmpctaexplotacioncierre "
    Cad = Cad & " where codusu =" & vUsu.Codigo & " and cta like 'q%'"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux!C2) Then
            
            Text1(43).Text = Format(miRsAux!C2, FormatoImporte)
            'EL IVA
            Importe = ImporteFormateado(Text1(39).Text)
            Importe = Importe - miRsAux!C2
            Text1(39).Text = Format(Importe, FormatoImporte)
            
            'IMPONIBLE
            Importe = ImporteFormateado(Text1(38).Text)
            Importe = Importe - miRsAux!C1
            Text1(38).Text = Format(Importe, FormatoImporte)
            
            Text1(42).Text = Format(miRsAux!C1, FormatoImporte)
        End If
    End If
    miRsAux.Close
    
    
    
    
    Set miRsAux = Nothing
    Exit Sub
EPonerIntraComYCampo:
    MuestraError Err.Number, "Poner IntraCom y Campo"
    Set miRsAux = Nothing

End Sub
