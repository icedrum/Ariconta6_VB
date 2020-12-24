VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesbloqueoMYSQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos bloqueados"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKill 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ListView lw1 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7646
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1729
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Inicio trx"
         Object.Width           =   4260
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usuario"
         Object.Width           =   3096
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Host"
         Object.Width           =   4179
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "BD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Info"
         Object.Width           =   3120
      EndProperty
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Image imgMysql 
      Height          =   255
      Index           =   0
      Left            =   2640
      Tag             =   "Leer estado"
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Consultar estado  BD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2325
   End
End
Attribute VB_Name = "frmDesbloqueoMYSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cad As String
Dim NF As Integer

Private Sub cmdKill_Click()

    If Me.lw1.ListItems.Count = 0 Then Exit Sub

    If Me.lw1.SelectedItem Is Nothing Then Exit Sub
    
    Cad = vbCrLf & lw1.ColumnHeaders(1).Text & ": " & lw1.SelectedItem.Text
    For i = 1 To Me.lw1.SelectedItem.ListSubItems.Count
        Cad = Cad & vbCrLf & lw1.ColumnHeaders(i + 1).Text & ": " & lw1.SelectedItem.ListSubItems(i)
    Next
    Msg = "Va a eliminar el proceso:" & vbCrLf & vbCrLf & Cad
    If MsgBox(Msg, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
'''''    Cad = ""
'''''    For i = 1 To lw1.ColumnHeaders.Count
'''''        Cad = Cad & vbCrLf & lw1.ColumnHeaders(i).Text & ": " & lw1.ColumnHeaders(i).Width
'''''    Next
'''''
'''''    Stop
'''''    Exit Sub
    
    If lw1.SelectedItem.Tag = 0 Then
        Msg = "Kill " & lw1.SelectedItem.Text
    Else
        Msg = lw1.SelectedItem.SubItems(4)
        Msg = "DElete from " & Msg & ".zbloqueos WHERE codusu= " & RecuperaValor(lw1.SelectedItem.SubItems(6), 2)
        Msg = Msg & " AND tabla = " & DBSet(RecuperaValor(lw1.SelectedItem.SubItems(6), 1), "T")
    End If
    
    If Ejecuta(Msg, False) Then
        If lw1.SelectedItem.Tag = 1 Then Cad = "zbloqueos" & vbCrLf & Cad
        Cad = "[KILL]" & vbCrLf & Cad
        vLog.Insertar 32, vUsu, Cad
        
        imgMysql_Click 0
        
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = 1 Then
        Me.Tag = 0
        imgMysql_Click 0
        
    End If
End Sub

Private Sub Form_Load()
    Me.Tag = 1
    Me.Icon = frmppal.Icon
    Me.imgMysql(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
End Sub


Private Sub imgMysql_Click(Index As Integer)

    
    Screen.MousePointer = vbHourglass
    
    If Index = 0 Then
            
        Label1(1).Caption = " leyendo BD "
        Label1(1).Refresh
        cmdKill.Enabled = False
        lw1.ListItems.Clear
        Set miRsAux = New ADODB.Recordset
        
        LeerInnodb
        
        Set miRsAux = Nothing
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub LeerInnodb()
Dim IT As ListItem
Dim idZbloq As Long
Dim ColBDs As Collection
Dim CuantasAplicaciones As Integer
Dim N As Integer
Dim K As Integer
Dim TablaBloqueos As String


On Error GoTo eLeerInnodb
    
    Msg = "SELECT ps.id, trx.trx_started,ps.user,ps.host,ps.DB,ps.command,ps.time,ps.state,ps.info,trx.trx_id"
    Msg = Msg & " FROM INFORMATION_SCHEMA.INNODB_TRX trx"
    Msg = Msg & " JOIN INFORMATION_SCHEMA.PROCESSLIST ps ON trx.trx_mysql_thread_id = ps.id"
    Msg = Msg & " WHERE trx.trx_started < CURRENT_TIMESTAMP - INTERVAL 60 SECOND"
    Msg = Msg & " AND ps.user != 'system_user'"
    
    miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add(, "K" & miRsAux!trx_id)
        IT.Text = miRsAux!Id
        For J = 1 To 5
            IT.SubItems(J) = DBLet(miRsAux.Fields(J), "T")
        Next
        IT.Tag = 0 '0 eliminar con KILL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Bloqueos producios desde ARICONTA/ARIGES
    CuantasAplicaciones = 4  'ariconta -  ariges  -arigasol  - ariagro
    
    For N = 1 To CuantasAplicaciones
       
       'ariges    Pass arigasol  Conta
       
        TablaBloqueos = "zbloqueos"  'si hay que personarlizarla, en su IF
        If N = 4 Then
            'ariagro
            Msg = "Select ariagro as queBD from usuarios.empresasariagro  "  'por si acaso
            Label1(1).Caption = "ariagro"
            
        ElseIf N = 2 Then
            'Ariges
            Msg = "Select ariges as queBD from usuarios.empresasariges  where codempre>0"  'por si acaso
            Label1(1).Caption = "ariges"
        ElseIf N = 3 Then
            'arigaso
            Msg = "Select arigasol as queBD from usuarios.empresasarigasol  "  'por si acaso
            Label1(1).Caption = "arigasol"
        
        Else
            Msg = "Select conta as queBD from usuarios.empresasariconta "
            Label1(1).Caption = "ariconta"
        End If
        Label1(1).Caption = "Leyendo tabla bloqueos " & Label1(1).Caption
        Label1(1).Refresh
        
        
        CargaEmpresasApliacionAriadna Msg, ColBDs
        
        
        For K = 1 To ColBDs.Count
        
            Label1(1).Caption = "Leyendo tabla bloqueos " & ColBDs.Item(K)
            Label1(1).Refresh
        
        
            Msg = "Select * from " & ColBDs.Item(K) & ".zbloqueos "
            AbrirRs Msg
            
            While Not miRsAux.EOF
                
                i = lw1.ListItems.Count + 1
                Label1(1).Caption = miRsAux!tabla & " " & i
                Label1(1).Refresh
                Set IT = lw1.ListItems.Add(, "B" & Format(i, "0000"))
                IT.Text = i
                
                IT.SubItems(1) = " - "
                        
                
                idZbloq = miRsAux!CodUsu \ 1000
                i = miRsAux!CodUsu Mod 1000
                
                Msg = DevuelveDesdeBD("login", "usuarios.usuarios", "codusu", CStr(i))
                If Msg = "" Then Msg = "N/D"
                IT.SubItems(2) = Msg
                
                Msg = DevuelveDesdeBD("nompc", "usuarios.pcs", "codpc", CStr(idZbloq))
                If Msg = "" Then Msg = "N/D"
                IT.SubItems(3) = Msg
                
                IT.SubItems(4) = ColBDs.Item(K)
                IT.SubItems(5) = "Bloqueado"
                IT.SubItems(6) = DBLet(miRsAux!tabla, "T") & "|" & miRsAux!CodUsu & "|"
                IT.Tag = 1 '0 eliminar con delete from table
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        Next K
    Next N
    Label1(1).Caption = "" 'Todo bien
    
    If vUsu.Nivel = 0 Then
        If Me.lw1.ListItems.Count > 0 Then cmdKill.Enabled = True
    End If
    
    Exit Sub
eLeerInnodb:
    Label1(1).Caption = "ERROR   ***************" & vbCrLf & Err.Description
    MsgBox Label1(1).Caption, vbExclamation
    Err.Clear
    Set miRsAux = Nothing
End Sub


'Llvara el codempre
Private Sub CargaEmpresasApliacionAriadna(NombreTabla_En_Sql As String, ByRef ColEmpreas As Collection)
Dim rsEmpresas As ADODB.Recordset
Dim Cad As String

    On Error GoTo eCargaEmpresasApliacionAriadna
    Set ColEmpreas = New Collection
    Set rsEmpresas = New ADODB.Recordset
    
    
    rsEmpresas.Open NombreTabla_En_Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not rsEmpresas.EOF
        Cad = rsEmpresas!queBD
        
        ColEmpreas.Add Cad
    
        rsEmpresas.MoveNext
    Wend
    rsEmpresas.Close
    
    Set rsEmpresas = Nothing
    Exit Sub
eCargaEmpresasApliacionAriadna:
    Err.Clear
    Conn.Errors.Clear
End Sub


Private Sub AbrirRs(sql As String)
    On Error Resume Next
    
    miRsAux.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Err.Number <> 0 Then
        Err.Clear
        Conn.Errors.Clear
        sql = "Select * from zbloqueos WHERE false" 'eof
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE DAR ERROR
    End If
End Sub

