Attribute VB_Name = "modTesoreria"
Option Explicit

Public Function CargarCobrosTemporal(Forpa As String, FecFactu As String, TotalFac As Currency) As Boolean
Dim Sql As String
Dim CadValues As String
Dim rsVenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarCobros

    CargarCobrosTemporal = False

    Sql = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    If Not rsVenci.EOF Then
        If rsVenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            i = 1
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
            '===
            
            'IMPORTE del Vencimiento
            If rsVenci!numerove = 1 Then
                ImpVenci = TotalFac
            Else
                ImpVenci = Round(TotalFac / rsVenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * rsVenci!numerove <> TotalFac Then
                    ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci.Fields(0).Value), 2)
                End If
            End If
            CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For i = 2 To rsVenci!numerove
                FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round(TotalFac / rsVenci.Fields(0).Value, 2)
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next i
        End If
    End If
    
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        Sql = "INSERT INTO tmpcobros (codusu, numorden, fecvenci, impvenci)"
        Sql = Sql & " VALUES " & Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute Sql
    End If
    
    CargarCobrosTemporal = True
    Exit Function

eCargarCobros:
    MuestraError Err.Number, "Cargar Cobros en Temporal", Err.Description
End Function


Public Function BancoPropio() As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    BancoPropio = ""

    Sql = "select codmacta from bancos "
    
    If TotalRegistrosConsulta(Sql) = 1 Then
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then BancoPropio = DBLet(Rs!codmacta, "T")
        Set Rs = Nothing
    End If

End Function

