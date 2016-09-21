Attribute VB_Name = "modEnlaceMutlibase"
Private Cnn As Connection

Private Cad As String
Private RS As ADODB.Recordset


Private Function AbreConexionMultibase() As Boolean
    
    On Error GoTo EAbreConexionMultibase
    AbreConexionMultibase = False
    
    
    Cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & vParam.EnlazaCtasMultibase
    Set Cnn = New Connection
    
    Cnn.Open Cad
    
    AbreConexionMultibase = True
    Exit Function
EAbreConexionMultibase:
    SQL = "Abrir conexión multibase" & vbCrLf & vbCrLf
    SQL = SQL & "ODBC: " & vParam.EnlazaCtasMultibase & vbCrLf
    SQL = SQL & Err.Description
    SQL = SQL & vbCrLf & vbCrLf & vbCrLf
    SQL = SQL & "¿Intentar enlazar con " & vParam.EnlazaCtasMultibase & " durante esta sesion?"
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then vParam.EnlazaCtasMultibase = ""
    
    
    
    Set Cnn = Nothing
    Cad = ""
    
End Function


Public Sub HacerEnlaceMultibase(Opcion As Byte, Datos As String)
    DoEvents
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
    
    Set RS = New ADODB.Recordset
    Cad = "Select * from smacta where codmacta ='" & RecuperaValor(Datos, 1) & "'"
    
    RS.Open Cad, Cnn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Existe = False
    If Not RS.EOF Then Existe = True
    RS.Close
    Set RS = Nothing
    
    
    If Existe Then Exit Sub
        
    'NO EXISTE. La creo
    Cad = "INSERT INTO smacta (codmacta,nommacta,apudirec) VALUES (""" & RecuperaValor(Datos, 1)
    Cad = Cad & """,""" & RecuperaValor(Datos, 2) & """,""S"")"
    Cnn.Execute Cad
    Exit Sub
EInsertaCuentaMultibase:
    MsgBox "Creando cta. multibase" & vbCrLf & Err.Description, vbExclamation
End Sub



Private Sub UpdateaCuentaMultibase(Datos As String)
    On Error GoTo EUpdateaCuentaMultibase
    
    InsertaCuentaMultibase Datos
    
    'NO EXISTE. La creo
    Cad = "UPDATE smacta SET nommacta = """ & RecuperaValor(Datos, 2) & """ WHERE"
    Cad = Cad & " codmacta= """ & RecuperaValor(Datos, 1) & """"
    
    Cnn.Execute Cad
    Exit Sub
EUpdateaCuentaMultibase:
    MsgBox "Update: cta. multibase" & vbCrLf & Err.Description, vbExclamation
End Sub
