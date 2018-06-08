VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTesDividVtoResult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dividir vencimientos"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3720
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdDivVto 
      Caption         =   "Aceptar"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Vencimientos a generar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   18
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmTesDividVtoResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Vtos As String
Public TIeneGastos As Boolean



Private Sub cmdCancelar_Click()
     Unload Me
End Sub

Private Sub cmdDivVto_Click()
    Ampliacion = "OK"
    Unload Me
End Sub

Private Sub Form_Load()
Dim N As Integer
    Me.Icon = frmppal.Icon
    
    
    
    N = 0
    Do
        i = InStr(1, Vtos, "|")
        If i = 0 Then
            Vtos = ""
        Else
            Msg = Mid(Vtos, 1, i - 1)
            Vtos = Mid(Vtos, i + 1)
            N = N + 1
            ListView1.ListItems.Add , , Mid(Msg, 1, 10)
            
            Msg = Mid(Msg, 11)
            ListView1.ListItems(N).SubItems(1) = Msg
            
            If N = 1 And TIeneGastos Then
                ListView1.ListItems(N).Bold = True
                ListView1.ListItems(N).ToolTipText = "Tiene gastos"
            End If
        End If
    Loop Until Vtos = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TIeneGastos = False
End Sub
