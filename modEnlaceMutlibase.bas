Attribute VB_Name = "modEnlaceMutlibase"
Private Cnn As Connection

Private cad As String
Private Rs As ADODB.Recordset


Private Function AbreConexionMultibase() As Boolean
    
    On Error GoTo EAbreConexionMultibase
    AbreConexionMultibase = False
    
    
    cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & vParam.EnlazaCtasMultibase
    Set Cnn = New Connection
    
    Cnn.Open cad
    
    AbreConexionMultibase = True
    Exit Function
EAbreConexionMultibase:
    Sql = "Abrir conexión multibase" & vbCrLf & vbCrLf
    Sql = Sql & "ODBC: " & vParam.EnlazaCtasMultibase & vbCrLf
    Sql = Sql & Err.Description
    Sql = Sql & vbCrLf & vbCrLf & vbCrLf
    Sql = Sql & "¿Intentar enlazar con " & vParam.EnlazaCtasMultibase & " durante esta sesion?"
    
    If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then vParam.EnlazaCtasMultibase = ""
    
    
    
    Set Cnn = Nothing
    cad = ""
    
End Function


Public Sub HacerEnlaceMultibase(Opcion As Byte, Datos As String)
    DoEvent2
    espera 0.2
    If Not AbreConexionMultibase Then Exit Sub
    
    
    
    
    Select Case Opcion
    Case 0
        'INSERTAR CUENTA EN MULTIBASE
        InsertaCuentaMultibase Datos
    
    Case 1
        UpdateaCuentaMultibase Datos
    
    End Select
    
    Cnn.Close
    Set Cnn = Nothing
End Sub


Private Sub InsertaCuentaMultibase(Datos As String)
Dim Existe As Boolean

    On Error GoTo EInsertaCuentaMultibase
    
    Set Rs = New ADODB.Recordset
    cad = "Select * from smacta where codmacta ='" & RecuperaValor(Datos, 1) & "'"
    
    Rs.Open cad, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Existe = False
    If Not Rs.EOF Then Existe = True
    Rs.Close
    Set Rs = Nothing
    
    
    If Existe Then Exit Sub
        
    'NO EXISTE. La creo
    cad = "INSERT INTO smacta (codmacta,nommacta,apudirec) VALUES (""" & RecuperaValor(Datos, 1)
    cad = cad & """,""" & RecuperaValor(Datos, 2) & """,""S"")"
    Cnn.Execute cad
    Exit Sub
EInsertaCuentaMultibase:
    MsgBox "Creando cta. multibase" & vbCrLf & Err.Description, vbExclamation
End Sub



Private Sub UpdateaCuentaMultibase(Datos As String)
    On Error GoTo EUpdateaCuentaMultibase
    
    InsertaCuentaMultibase Datos
    
    'NO EXISTE. La creo
    cad = "UPDATE smacta SET nommacta = """ & RecuperaValor(Datos, 2) & """ WHERE"
    cad = cad & " codmacta= """ & RecuperaValor(Datos, 1) & """"
    
    Cnn.Execute cad
    Exit Sub
EUpdateaCuentaMultibase:
    MsgBox "Update: cta. multibase" & vbCrLf & Err.Description, vbExclamation
End Sub
