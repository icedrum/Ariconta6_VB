VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportarNavProceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar fichero Consum"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAcciones 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtCD1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6015
   End
   Begin VB.CommandButton cmdAcciones 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image ImgCd1 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmImportarNavProceso.frx":0000
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Importar factura CONSUM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label6 
      Caption         =   "Fichero"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmImportarNavProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcciones_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = txtCD1(0).Text
        If CadenaDesdeOtroForm = "" Then
            MsgBox "Seleccione el fichero para importar", vbExclamation
            Exit Sub
        End If
    Else
        CadenaDesdeOtroForm = ""
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
End Sub

Private Sub ImgCd1_Click(Index As Integer)


    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.CancelError = False
    If Index = 0 Then
        cd1.Filter = "DAT (*.dat)|*.dat|RTF (*.rtf)|*.rtf"
        cd1.FilterIndex = 0
    End If
    cd1.ShowOpen
    
    Screen.MousePointer = vbDefault
    If cd1.FileName = "" Then Exit Sub
    
    txtCD1(Index).Text = cd1.FileName
    
End Sub
