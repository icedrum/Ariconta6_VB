Attribute VB_Name = "modTesoreria"
Option Explicit

Public Function CargarCobrosTemporal(Forpa As String, FecFactu As String, TotalFac As Currency) As Boolean
Dim SQL As String
Dim CadValues As String
Dim Rsvenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarCobros

    CargarCobrosTemporal = False

    SQL = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set Rsvenci = New ADODB.Recordset
    Rsvenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    If Not Rsvenci.EOF Then
        If Rsvenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            i = 1
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(Rsvenci!primerve, "N"), FecVenci)
            '===
            
            'IMPORTE del Vencimiento
            If Rsvenci!numerove = 1 Then
                ImpVenci = TotalFac
            Else
                ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * Rsvenci!numerove <> TotalFac Then
                    ImpVenci = Round2(ImpVenci + (TotalFac - ImpVenci * Rsvenci.Fields(0).Value), 2)
                End If
            End If
            CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For i = 2 To Rsvenci!numerove
                FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next i
        End If
    End If
    
    Set Rsvenci = Nothing
    
    If CadValues <> "" Then
        SQL = "INSERT INTO tmpcobros (codusu, numorden, fecvenci, impvenci)"
        SQL = SQL & " VALUES " & Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute SQL
    End If
    
    CargarCobrosTemporal = True
    Exit Function

eCargarCobros:
    MuestraError Err.Number, "Cargar Cobros en Temporal", Err.Description
End Function


'Cargara sobre un collection los cobros.
'Cada linea el SQL
'       insert into cobros(numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,numorden,fecvenci,impvenci)
'    Para ello enviaremos TODO el sql menos y numorden fecvenci e impvenci
Public Function CargarCobrosSobreCollectionConSQLInsert(ByRef ColCobros As Collection, Forpa As String, FecFactu As String, TotalFac As Currency, PartFijaSQL As String) As Boolean
Dim SQL As String
Dim Rsvenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarCobros

    CargarCobrosSobreCollectionConSQLInsert = False

    Set ColCobros = New Collection
    
    SQL = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set Rsvenci = New ADODB.Recordset
    Rsvenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    
    If Not Rsvenci.EOF Then
        If Rsvenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(Rsvenci!primerve, "N"), FecVenci)
            '===
            
            'IMPORTE del Vencimiento
            If Rsvenci!numerove = 1 Then
                ImpVenci = TotalFac
            Else
                ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * Rsvenci!numerove <> TotalFac Then
                    ImpVenci = Round2(ImpVenci + (TotalFac - ImpVenci * Rsvenci.Fields(0).Value), 2)
                End If
            End If
            'CadValues = "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            ColCobros.Add PartFijaSQL & "1," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & ")"
            
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For i = 2 To Rsvenci!numerove
                FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
                
                'CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
                ColCobros.Add PartFijaSQL & i & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & ")"
            Next i
        End If
    End If
    
    Set Rsvenci = Nothing
    
    
    
    CargarCobrosSobreCollectionConSQLInsert = True
    Exit Function

eCargarCobros:
    MuestraError Err.Number, "Cargar Cobros auxiliar", Err.Description
End Function











Public Function BancoPropio() As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    BancoPropio = ""

    SQL = "select codmacta from bancos "
    
    If TotalRegistrosConsulta(SQL) = 1 Then
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then BancoPropio = DBLet(Rs!codmacta, "T")
        Set Rs = Nothing
    End If

End Function








Public Function HayQueMostrarEliminarRiesgoTalPag() As Boolean
Dim SQL As String
Dim Col As Collection
    
    On Error GoTo eHayQueMostrarEliminarRiesgoTalPag
    HayQueMostrarEliminarRiesgoTalPag = False
    Set miRsAux = New ADODB.Recordset
    SQL = "Select codigo,anyo,codmacta,tiporem  from remesas where  tiporem > 1  AND (situacion ='Q' or situacion ='Y') ORDER BY codmacta,1,2 "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    Set Col = New Collection
    Msg = ""
    While Not miRsAux.EOF
        If Msg <> miRsAux!codmacta Then
            '
            '           tiporem|dias                Resto. Remesas
            If Msg <> "" Then Col.Add SQL
            
            
            SQL = "concat( pagaredias,'|',talondias,'|')"
            Msg = DevuelveDesdeBD(SQL, "bancos", "codmacta", miRsAux!codmacta, "T")
            If Msg = "" Then Err.Raise 513, "No existe banco?" & miRsAux!codmacta
            If miRsAux!tiporem = 2 Then
                SQL = RecuperaValor(Msg, 1)
            Else
                SQL = RecuperaValor(Msg, 2)
            End If
            If SQL = "" Then SQL = "0"
            Msg = Val(SQL) + 1
            SQL = Format(Msg, "000")
            Msg = miRsAux!codmacta
            
        End If
        
        SQL = SQL & ", (" & miRsAux!Codigo & "," & miRsAux!Anyo & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If SQL <> "" Then Col.Add SQL
        
        
    For i = 1 To Col.Count
        SQL = Col.Item(i)
        J = Val(Mid(SQL, 1, 3))
        SQL = Mid(SQL, 5)
        
        Msg = "select count(*)"
        Msg = Msg & " from cobros where (codrem,anyorem) in ("
        Msg = Msg & SQL
        Msg = Msg & ") and date_add(fecvenci, interval " & J & " day) <now()"
        Msg = Msg & " order by fecvenci"
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            If DBLet(miRsAux.Fields(0), "N") > 0 Then HayQueMostrarEliminarRiesgoTalPag = True
        End If
        miRsAux.Close
        If HayQueMostrarEliminarRiesgoTalPag Then Exit For
    Next i
        
eHayQueMostrarEliminarRiesgoTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set Col = Nothing
    Msg = ""
End Function



Public Function QueRemesasMostrarEliminarRiesgoTalPag() As String
Dim SQL As String
Dim Col As Collection
    
    On Error GoTo eHayQueMostrarEliminarRiesgoTalPag
    QueRemesasMostrarEliminarRiesgoTalPag = ""
    Set miRsAux = New ADODB.Recordset
    SQL = "Select codigo,anyo,codmacta,tiporem  from remesas where  tiporem > 1  AND (situacion ='Q' or situacion ='Y') ORDER BY codmacta,1,2 "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    Set Col = New Collection
    Msg = ""
    While Not miRsAux.EOF
        If Msg <> miRsAux!codmacta Then
            '
            '           tiporem|dias                Resto. Remesas
            If Msg <> "" Then Col.Add SQL
            
            
            SQL = "concat( pagaredias,'|',talondias,'|')"
            Msg = DevuelveDesdeBD(SQL, "bancos", "codmacta", miRsAux!codmacta, "T")
            If Msg = "" Then Err.Raise 513, "No existe banco?" & miRsAux!codmacta
            If miRsAux!tiporem = 2 Then
                SQL = Format(RecuperaValor(Msg, 1), "000")
            Else
                SQL = Format(RecuperaValor(Msg, 2), "000")
            End If
            Msg = miRsAux!codmacta
            
        End If
        
        SQL = SQL & ", (" & miRsAux!Codigo & "," & miRsAux!Anyo & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If SQL <> "" Then Col.Add SQL
        
        
    For i = 1 To Col.Count
        SQL = Col.Item(i)
        J = Val(Mid(SQL, 1, 3))
        SQL = Mid(SQL, 5)
        
        Msg = "select distinct codrem,anyorem "
        Msg = Msg & " from cobros where (codrem,anyorem) in ("
        Msg = Msg & SQL
        Msg = Msg & ") and date_add(fecvenci, interval " & J & " day) <now()"
        Msg = Msg & " order by fecvenci"
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            QueRemesasMostrarEliminarRiesgoTalPag = QueRemesasMostrarEliminarRiesgoTalPag & ", (" & miRsAux!CodRem & "," & miRsAux!AnyoRem & ")"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
    Next i
        
eHayQueMostrarEliminarRiesgoTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set Col = Nothing
    Msg = ""
End Function


