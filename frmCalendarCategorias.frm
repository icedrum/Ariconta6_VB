VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalendarCategorias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categorias eventos calendario"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descripcion"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmCalendarCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemModificando As Integer

Private Sub Form_Load()
Dim pLabel As CalendarEventLabel
Dim nLabelID As Long

    Me.Icon = frmppal.Icon

    
    'Cargamos categorias
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from usuarios.calendaretiquetas ORDER BY id", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    While Not miRsAux.EOF
        
            nLabelID = miRsAux!Id
            NumRegElim = NumRegElim + 1
            ListView1.ListItems.Add , "K" & Format(nLabelID, "00"), miRsAux!TEXTO
            'ListView1.ListItems(NumRegElim).SubItems(1) = " "
            Set pLabel = frmInbox.CalendarControl.DataProvider.LabelList.Find(nLabelID)
            If Not pLabel Is Nothing Then
                'ctrlColor.BackColor = pLabel.Color
                ListView1.ListItems(NumRegElim).ForeColor = pLabel.Color
            End If
    
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    If ItemModificando <= 1 Then
        Cancel = 1
    Else
        
        If NewString = "" Then
            Cancel = 1
        Else
            Msg = "UPDATE usuarios.calendaretiquetas SET texto =" & DBSet(NewString, "T") & " WHERE id =" & ItemModificando
            If Not Ejecuta(Msg) Then
                Cancel = 1
            Else
                Cancel = 0
            End If
        End If
    End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    On Error Resume Next
    
    ItemModificando = -1
    ItemModificando = Mid(ListView1.SelectedItem.Key, 2)
    If Err.Number <> 0 Then Err.Clear
    
End Sub

