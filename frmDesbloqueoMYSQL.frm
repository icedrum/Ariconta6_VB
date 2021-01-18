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
Dim cad As String
Dim NF As Integer

Private Sub cmdKill_Click()

    If Me.lw1.ListItems.Count = 0 Then Exit Sub

    If Me.lw1.SelectedItem Is Nothing Then Exit Sub
    
    
    If InStr(1, Me.lw1.SelectedItem.Key, "KPN") > 0 Then
        MsgBox "Informacion de proceso incompleta. No se puede eliminar", vbExclamation
        Exit Sub
    End If
    
    cad = vbCrLf & lw1.ColumnHeaders(1).Text & ": " & lw1.SelectedItem.Text
    For I = 1 To Me.lw1.SelectedItem.ListSubItems.Count
        cad = cad & vbCrLf & lw1.ColumnHeaders(I + 1).Text & ": " & lw1.SelectedItem.ListSubItems(I)
    Next
    Msg = "Va a eliminar el proceso:" & vbCrLf & vbCrLf & cad
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
        If lw1.SelectedItem.Tag = 1 Then cad = "zbloqueos" & vbCrLf & cad
        cad = "[KILL]" & vbCrLf & cad
        vLog.Insertar 32, vUsu, cad
        
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
Dim m As Integer
Dim TablaBloqueos As String
Dim VersionAntiguaMysql As Boolean
Dim ColProcessList As Collection
Dim DelShowProcess As Long


On Error GoTo eLeerInnodb
    
    VersionAntiguaMysql = False
    Msg = "Select * from INFORMATION_SCHEMA.INNODB_TRX trx "
    If Not Ejecuta(Msg, True) Then VersionAntiguaMysql = True
    
    
    
    If VersionAntiguaMysql Then
    
        'Processlist
        Label1(1).Caption = "show processlist"
        Label1(1).Refresh
        Set ColProcessList = New Collection
        Msg = "SHOW PROCESSLIST"
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Msg = miRsAux.Fields(0) & "|" & miRsAux.Fields(2) & "|" & DBLet(miRsAux.Fields(3), "T") & "|" & miRsAux.Fields(5) & "|"
            ColProcessList.Add Msg
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
        Msg = "SHOW INNODB STATUS"
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Msg = ""
        If Not miRsAux.EOF Then
            If Not IsNull(miRsAux.Fields(0)) Then Msg = miRsAux.Fields(0)
        End If
        miRsAux.Close
        'CreaCAdena
        If Msg <> "" Then
            Label1(1).Caption = "show innodb"
            'Si no tiene saltos de linea, se los pongo
            If InStr(1, Msg, vbCrLf) = 0 Then Msg = Replace(Msg, vbLf, vbCrLf)
                
            N = 1
            TablaBloqueos = "------------" & vbCrLf & "TRANSACTIONS"
            K = InStr(N, Msg, TablaBloqueos)
            If K > 0 Then
                TablaBloqueos = "--------" & vbCrLf & "FILE I/O"
                N = InStr(K, Msg, TablaBloqueos)
            End If
            
            Label1(1).Caption = "Buscando tramo " & K & " - " & N
            
            
            If K = 0 Or N = 0 Then
                
                
                
                frmZoom.pValor = Msg
                frmZoom.pModo = 3
                frmZoom.Caption = "Show innodb "
                frmZoom.Show vbModal
                
                
            Else
                
                Msg = Mid(Msg, K, N - K)  'las transacciones ---TRANSACTION
                
                Set ColBDs = New Collection
                
                'Aqui ahora vamos a buscar los lockst ruc
                'lock struct(s), linea a linea
                N = 1
                Do
                    Label1(1).Caption = "lock strut: " & N
                    K = InStr(N, Msg, "lock struct(s),")
                    If K = 0 Then
                        'FIN
                        N = 0
                    Else
                        
                        m = InStr(K, Msg, "---TRANSAC")
                        If m > 0 Then
                            'Aun hay mas
                            TablaBloqueos = Mid(Msg, K, m - K - 1)
                            N = m + 1
                        Else
                            TablaBloqueos = Mid(Msg, K)
                            N = 0
                        End If
                        
                        ColBDs.Add TablaBloqueos
                    End If
                Loop Until N = 0
                
                Label1(1).Caption = "Bloqueados " & ColBDs.Count
                
                For N = 1 To ColBDs.Count
                    Msg = ColBDs.Item(N)
                    TablaBloqueos = "MySQL thread id "
                    K = InStr(1, Msg, TablaBloqueos)
                    
                    If K > 0 Then
                        Label1(1).Caption = "Bloqueados " & ColBDs.Count & " //" & Msg
                        Msg = Mid(Msg, K + Len(TablaBloqueos))
                        K = InStr(2, Msg, ",")
                        If K > 0 Then
                            TablaBloqueos = Val(Mid(Msg, 1, K - 1))
                            idZbloq = Val(TablaBloqueos)
                            
                            DelShowProcess = 0
                            For NumRegElim = 1 To ColProcessList.Count
                                TablaBloqueos = RecuperaValor(ColProcessList.Item(NumRegElim), 1)
                                If idZbloq = CLng(TablaBloqueos) Then
                                    'Este es
                                    DelShowProcess = NumRegElim
                                    Exit For
                                End If
                            Next
                            
                            If DelShowProcess = 0 Then
                                'No he encontrado el process list . RARO. pero inserto sin mas
                                K = 1 'para que inserte el nodo
                                
                            Else
                                TablaBloqueos = RecuperaValor(ColProcessList.Item(DelShowProcess), 4)
                                NumRegElim = Val(TablaBloqueos)
                                K = IIf(NumRegElim > 70, 1, 0) 'Si han ppasdo mas de 60 segundos
                            End If
                            
                            If K > 0 Then
                            
                                Set IT = lw1.ListItems.Add(, "K" & idZbloq)
                                IT.Tag = 0
                                IT.Text = idZbloq
                                For K = 1 To 5
                                    IT.SubItems(K) = " "
                                Next
                                IT.SubItems(2) = "root"
                                
                                K = InStr(1, Msg, vbCrLf)
                                If K = 0 Then K = InStr(1, Msg, "root")
                                
                                If K = 0 Then
                                    'No disponible
                                    
                                Else
                                    Msg = Mid(Msg, 1, K - 1)
                                    IT.SubItems(6) = CStr(Msg)
                                    Msg = Trim(Replace(Msg, "root", ""))
                                    
                                    'Por IP
                                    K = InStrRev(Msg, " ")
                                    If K > 0 Then
                                        IT.SubItems(3) = Trim(Mid(Msg, K))
                                    Else
                                        IT.SubItems(3) = "N/D"
                                        K = 1
                                    End If
                                End If
                                    
                                If DelShowProcess > 0 Then
                                    Msg = DateAdd("s", -NumRegElim, Now)
                                    IT.SubItems(1) = Msg
                                    IT.SubItems(3) = RecuperaValor(ColProcessList.Item(DelShowProcess), 2)
                                    IT.SubItems(4) = RecuperaValor(ColProcessList.Item(DelShowProcess), 3)
                                    IT.SubItems(5) = "N/D"
                                End If
                            Else
                                K = 1 'para que no lo inserte aqui abajo
                            End If
                            
                        Else
                            K = 0
                        End If
                    End If
                    If K = 0 Then
                    
                        'No lo ubico. Ponemos texto lo que podamos
                        Set IT = lw1.ListItems.Add(, "KPN" & N)  '-->No podra hacer kill
                        For K = 1 To 5
                             IT.SubItems(K) = "N/D"
                        Next
                        IT.SubItems(6) = Msg
                    End If
                Next
                Set ColBDs = Nothing
                TablaBloqueos = ""
            End If
        
        End If  'msg<>''
    Else
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
    End If
    
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
                
                I = lw1.ListItems.Count + 1
                Label1(1).Caption = miRsAux!tabla & " " & I
                Label1(1).Refresh
                Set IT = lw1.ListItems.Add(, "B" & Format(I, "0000"))
                IT.Text = I
                
                IT.SubItems(1) = " - "
                        
                
                idZbloq = miRsAux!CodUsu \ 1000
                I = miRsAux!CodUsu Mod 1000
                
                Msg = DevuelveDesdeBD("login", "usuarios.usuarios", "codusu", CStr(I))
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
    Label1(1).Caption = "ERROR   ***************" & Label1(1).Caption & vbCrLf & Err.Description
    MsgBox Label1(1).Caption, vbExclamation
    Err.Clear
    Set miRsAux = Nothing
End Sub


'Llvara el codempre
Private Sub CargaEmpresasApliacionAriadna(NombreTabla_En_Sql As String, ByRef ColEmpreas As Collection)
Dim rsEmpresas As ADODB.Recordset
Dim cad As String

    On Error GoTo eCargaEmpresasApliacionAriadna
    Set ColEmpreas = New Collection
    Set rsEmpresas = New ADODB.Recordset
    
    
    rsEmpresas.Open NombreTabla_En_Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not rsEmpresas.EOF
        cad = rsEmpresas!queBD
        
        ColEmpreas.Add cad
    
        rsEmpresas.MoveNext
    Wend
    rsEmpresas.Close
    
    Set rsEmpresas = Nothing
    Exit Sub
eCargaEmpresasApliacionAriadna:
    Err.Clear
    Conn.Errors.Clear
End Sub


Private Sub AbrirRs(Sql As String)
    On Error Resume Next
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Err.Number <> 0 Then
        Err.Clear
        Conn.Errors.Clear
        Sql = "Select * from zbloqueos WHERE false" 'eof
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE DAR ERROR
    End If
End Sub

Private Sub lw1_DblClick()
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub
    
    MsgBox lw1.SelectedItem.Key & vbCrLf & lw1.SelectedItem.SubItems(6), vbInformation
End Sub





Private Sub CreaCAdena()

'
'
Msg = ""
Msg = Msg & vbCrLf & "====================================="
Msg = Msg & vbCrLf & "----------"
Msg = Msg & vbCrLf & "SEMAPHORES"
Msg = Msg & vbCrLf & "----------"
Msg = Msg & vbCrLf & "OS WAIT ARRAY INFO: reservation count 483, signal count 477"
Msg = Msg & vbCrLf & "Mutex spin waits 71, rounds 321, OS waits 1"
Msg = Msg & vbCrLf & "RW-shared spins 678, rounds 18518, OS waits 474"
Msg = Msg & vbCrLf & "RW-excl spins 11, rounds 10, OS waits 0"
Msg = Msg & vbCrLf & "Spin rounds per wait: 4.52 mutex, 27.31 RW-shared, 0.91 RW-excl"
Msg = Msg & vbCrLf & "------------"
Msg = Msg & vbCrLf & "TRANSACTIONS"
Msg = Msg & vbCrLf & "------------"
Msg = Msg & vbCrLf & "Trx id counter 7BD2135"
Msg = Msg & vbCrLf & "Purge done for trx's n:o < 7BD1E9C undo n:o < 0"
Msg = Msg & vbCrLf & "History list length 905"
Msg = Msg & vbCrLf & "LIST OF TRANSACTIONS FOR EACH SESSION:"
Msg = Msg & vbCrLf & "---TRANSACTION 0, not started"
Msg = Msg & vbCrLf & "MySQL thread id 1, OS thread handle 0x22c0, query id 3146 PCDAVID 192.100.100.231 root"
Msg = Msg & vbCrLf & "show engine innodb  status"
Msg = Msg & vbCrLf & "---TRANSACTION 7BD2069, not started"
Msg = Msg & vbCrLf & "MySQL thread id 8, OS thread handle 0x2644, query id 2936 localhost 127.0.0.1 root"
Msg = Msg & vbCrLf & "---TRANSACTION 7BD212C, ACTIVE 536 sec"
Msg = Msg & vbCrLf & "5 lock struct(s), heap size 1248, 3 row lock(s)"
Msg = Msg & vbCrLf & "MySQL thread id 33, OS thread handle 0x1eb8, query id 3124 localhost 127.0.0.1 root"
Msg = Msg & vbCrLf & "Trx read view will not see trx with id >= 7BD212D, sees < 7BD212D"
Msg = Msg & vbCrLf & "--------"
Msg = Msg & vbCrLf & "FILE I/O"
Msg = Msg & vbCrLf & "--------"
Msg = Msg & vbCrLf & ""

End Sub

