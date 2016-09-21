Attribute VB_Name = "libBusADO"
Option Explicit




'-----------------------------------------------------------
'----------------------------------------------------------
'   Para los formualrios intod. astos
                    '   fac client
                    '   fac prov
                    
                    
                    

                        
                        




'Borra los datos del tmp
Public Function BorraTmp(Tabla As Byte)
    On Error Resume Next
    Conn.Execute "Delete from tmpwbusca" & Tabla & " where codusu = " & vUsu.Codigo
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function


'0  Asientos
'1  Clientes
'2  Proveedores
Public Function InsertaTmp(ByRef vSQL As String, Tabla As Byte)
Dim SQL As String

    
    Select Case Tabla
    Case 0
        'ASientos
        SQL = "INSERT INTO tmpwBusca0(codusu,tabla,long1,long2"
        SQL = SQL & ",fechaent)"
        SQL = SQL & " SELECT " & vUsu.Codigo & ",0,numasien,numdiari,fechaent From cabapu "
        'El SQL empezar a partir del where , inclusive
        SQL = SQL & vSQL
    
    Case 1
        'Faccli
        SQL = "INSERT INTO tmpwBusca1(codusu,codfaccl,anofaccl,numserie)"
        SQL = SQL & " SELECT " & vUsu.Codigo & ",codfaccl,anofaccl,numserie From cabfact "
        'El SQL empezar a partir del where , inclusive
        SQL = SQL & vSQL
        SQL = SQL & " ORDER BY fecfaccl ASC ,codfaccl ASC"
    Case 2
        'Proveedores
        SQL = "INSERT INTO tmpwBusca2(codusu,numregis,anofacpr)"
        SQL = SQL & " SELECT " & vUsu.Codigo & ",numregis,anofacpr From cabfactprov "
        'El SQL empezar a partir del where , inclusive
        SQL = SQL & vSQL
        SQL = SQL & " ORDER BY fecrecpr ASC ,numregis ASC"
    
    End Select
    
    
    
    
    Conn.Execute SQL
End Function


Public Function InsertaValoresAsientos(NumAsi As String, fechaent As String, NumDiari As String) As Boolean

Dim SQL As String
    On Error Resume Next
    InsertaValoresAsientos = False
    SQL = "INSERT INTO tmpwBusca0(codusu,tabla,long1,long2,fechaent) VALUES (" & vUsu.Codigo & ",0,"
    SQL = SQL & NumAsi & "," & NumDiari & ",'" & Format(fechaent, FormatoFecha) & "')"
    Conn.Execute SQL
    If Err.Number = 0 Then
        InsertaValoresAsientos = True
    Else
        InsertaValoresAsientos = False
        'Mostraremos error
        MuestraError Err.Number, "InsertaValoresAsientos"
    End If
End Function


Public Function EliminaValoresAsientos(NumAsi As String, fechaent As String, NumDiari As String) As Boolean

Dim SQL As String
    On Error Resume Next
    EliminaValoresAsientos = False
    SQL = "DELETE FROM tmpwBusca0 WHERE codusu = " & vUsu.Codigo & " AND long1 = "
    SQL = SQL & NumAsi & " AND long2 = " & NumDiari & " AND Fechaent = '" & Format(fechaent, FormatoFecha) & "'"
    Conn.Execute SQL
    If Err.Number = 0 Then
        EliminaValoresAsientos = True
    Else
        EliminaValoresAsientos = False
        'Mostraremos error
        MuestraError Err.Number, "EliminaValoresAsientos"
    End If
End Function


'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Public Function EliminaValoresFACCLI(Numserie As String, codfaccl As String, anofaccl As String) As Boolean

Dim SQL As String
    On Error Resume Next
    EliminaValoresFACCLI = False
    SQL = "DELETE FROM tmpwBusca1 WHERE codusu = " & vUsu.Codigo & " AND numserie = '"
    SQL = SQL & Numserie & "' AND codfaccl = " & codfaccl & " AND anofaccl = " & anofaccl
    Conn.Execute SQL
    If Err.Number = 0 Then
        EliminaValoresFACCLI = True
    Else
        EliminaValoresFACCLI = False
        'Mostraremos error
        MuestraError Err.Number, "EliminaValoresFactura"
    End If
End Function



Public Function InsertaValoresFACCLI(Numserie As String, Codfac As String, anofac As String) As Boolean

Dim SQL As String
    On Error Resume Next
    SQL = "INSERT INTO tmpwBusca1(codusu,codfaccl,anofaccl,numserie) VALUES (" & vUsu.Codigo & ","
    SQL = SQL & Codfac & "," & anofac & ",'" & Numserie & "')"
    Conn.Execute SQL
    If Err.Number = 0 Then
        InsertaValoresFACCLI = True
    Else
        InsertaValoresFACCLI = False
        'Mostraremos error
        MuestraError Err.Number, "InsertaValoresAsientos"
    End If
End Function

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'         P  R  O  V  E  E  D  O  R  E  S
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
Public Function EliminaValoresFACPRO(numregis As String, anofacpr As String) As Boolean

Dim SQL As String
    On Error Resume Next
    SQL = "DELETE FROM tmpwBusca2 WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " AND numregis = " & numregis & " AND anofacpr = " & anofacpr
    Conn.Execute SQL
    If Err.Number = 0 Then
        EliminaValoresFACPRO = True
    Else
        EliminaValoresFACPRO = False
        'Mostraremos error
        MuestraError Err.Number, "EliminaValoresFactura"
    End If
End Function



Public Function InsertaValoresFACPRO(numregis As String, anofac As String) As Boolean

Dim SQL As String
    On Error Resume Next
    SQL = "INSERT INTO tmpwBusca2(codusu,numregis,anofacpr) VALUES (" & vUsu.Codigo & ","
    SQL = SQL & numregis & "," & anofac & ")"
    Conn.Execute SQL
    If Err.Number = 0 Then
        InsertaValoresFACPRO = True
    Else
        InsertaValoresFACPRO = False
        'Mostraremos error
        MuestraError Err.Number, "InsertaValoresAsientos"
    End If
End Function



'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------



Public Function CargaADOBUS(ByRef AD As Adodc)
    Set AD.Recordset = Nothing
    AD.RecordSource = "Select * from tmpwBusca0 where codusu = " & vUsu.Codigo & " ORDER BY fechaent,long1"
    AD.ConnectionString = Conn
    AD.Refresh
End Function



'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------


Public Sub BorrarTmpWBusca()
    On Error GoTo EBorrarTmpWBusca
    
    Conn.Execute "DELETE FROM tmpwBusca0"
    Conn.Execute "DELETE FROM tmpwBusca1"
    Conn.Execute "DELETE FROM tmpwBusca2"
    Exit Sub
EBorrarTmpWBusca:
    Err.Clear
    
    
End Sub
