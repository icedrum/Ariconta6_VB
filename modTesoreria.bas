Attribute VB_Name = "modTesoreria"
Option Explicit

Public Function CargarCobrosTemporal(Forpa As String, FecFactu As String, TotalFac As Currency) As Boolean
Dim Sql As String
Dim CadValues As String
Dim Rsvenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarCobros

    CargarCobrosTemporal = False

    Sql = "SELECT numerove, primerve, restoven FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set Rsvenci = New ADODB.Recordset
    Rsvenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    If Not Rsvenci.EOF Then
        If Rsvenci.Fields(0).Value > 0 Then
            '-------- Primer Vencimiento
            I = 1
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
            CadValues = "(" & vUsu.Codigo & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            
            'Resto Vencimientos
            '--------------------------------------------------------------------
            For I = 2 To Rsvenci!numerove
                FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
                
                CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next I
        End If
    End If
    
    Set Rsvenci = Nothing
    
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


'Cargara sobre un collection los cobros.
'Cada linea el SQL
'       insert into cobros(numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,numorden,fecvenci,impvenci)
'    Para ello enviaremos TODO el sql menos y numorden fecvenci e impvenci
Public Function CargarCobrosSobreCollectionConSQLInsert(ByRef ColCobros As Collection, Forpa As String, FecFactu As String, TotalFac As Currency, FechaVto As String, PartFijaSQL As String) As Boolean
Dim Sql As String
Dim Rsvenci As ADODB.Recordset
Dim FecVenci As String
Dim ImpVenci As Currency

    On Error GoTo eCargarCobros

    CargarCobrosSobreCollectionConSQLInsert = False

    Set ColCobros = New Collection
    If FechaVto <> "" Then
        If CDate(FechaVto) < "01/01/2020" Then FechaVto = ""
    End If
    If FechaVto <> "" Then
            'Ha puesto una fecha de vto
        Sql = "SELECT 1 numerove, 0 primerve,1  restoven "
    Else
        Sql = "SELECT numerove, primerve, restoven "
    End If
    Sql = Sql & " FROM formapago WHERE codforpa=" & DBSet(Forpa, "N")
    
    Set Rsvenci = New ADODB.Recordset
    Rsvenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

    
    If Not Rsvenci.EOF Then
    
        
            '-------- Primer Vencimiento
        If FechaVto <> "" Then
            FecVenci = CDate(FechaVto)
        
        Else
            'FECHA VTO
            FecVenci = CDate(FecFactu)
            FecVenci = DateAdd("d", DBLet(Rsvenci!primerve, "N"), FecVenci)
            '===
        End If
        
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
        For I = 2 To Rsvenci!numerove
            FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                
            'IMPORTE Resto de Vendimientos
            ImpVenci = Round2(TotalFac / Rsvenci.Fields(0).Value, 2)
            
            'CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(i, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            ColCobros.Add PartFijaSQL & I & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & ")"
        Next I
        
    End If
    
    Set Rsvenci = Nothing
    
    
    
    CargarCobrosSobreCollectionConSQLInsert = True
    Exit Function

eCargarCobros:
    MuestraError Err.Number, "Cargar Cobros auxiliar", Err.Description
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








Public Function HayQueMostrarEliminarRiesgoTalPag() As Boolean
Dim Sql As String
Dim Col As Collection
    
    On Error GoTo eHayQueMostrarEliminarRiesgoTalPag
    HayQueMostrarEliminarRiesgoTalPag = False
    Set miRsAux = New ADODB.Recordset
    Sql = "Select codigo,anyo,codmacta,tiporem  from remesas where  tiporem > 1  AND (situacion ='Q' or situacion ='Y') ORDER BY codmacta,1,2 "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""
    Set Col = New Collection
    Msg = ""
    While Not miRsAux.EOF
        If Msg <> miRsAux!codmacta Then
            '
            '           tiporem|dias                Resto. Remesas
            If Msg <> "" Then Col.Add Sql
            
            
            Sql = "concat( pagaredias,'|',talondias,'|')"
            Msg = DevuelveDesdeBD(Sql, "bancos", "codmacta", miRsAux!codmacta, "T")
            If Msg = "" Then Err.Raise 513, "No existe banco?" & miRsAux!codmacta
            If miRsAux!Tiporem = 2 Then
                Sql = RecuperaValor(Msg, 1)
            Else
                Sql = RecuperaValor(Msg, 2)
            End If
            If Sql = "" Then Sql = "0"
            Msg = Val(Sql) + 1
            Sql = Format(Msg, "000")
            Msg = miRsAux!codmacta
            
        End If
        
        Sql = Sql & ", (" & miRsAux!Codigo & "," & miRsAux!Anyo & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Sql <> "" Then Col.Add Sql
        
        
    For I = 1 To Col.Count
        Sql = Col.Item(I)
        J = Val(Mid(Sql, 1, 3))
        Sql = Mid(Sql, 5)
        
        Msg = "select count(*)"
        Msg = Msg & " from cobros where (codrem,anyorem) in ("
        Msg = Msg & Sql
        Msg = Msg & ") and date_add(fecvenci, interval " & J & " day) <now()"
        Msg = Msg & " order by fecvenci"
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            If DBLet(miRsAux.Fields(0), "N") > 0 Then HayQueMostrarEliminarRiesgoTalPag = True
        End If
        miRsAux.Close
        If HayQueMostrarEliminarRiesgoTalPag Then Exit For
    Next I
        
eHayQueMostrarEliminarRiesgoTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set Col = Nothing
    Msg = ""
End Function



Public Function QueRemesasMostrarEliminarRiesgoTalPag2(SoloEfectos As Boolean) As String
Dim Sql As String
Dim Col As Collection
    
    On Error GoTo eHayQueMostrarEliminarRiesgoTalPag
    QueRemesasMostrarEliminarRiesgoTalPag2 = ""
    Set miRsAux = New ADODB.Recordset
    Sql = ">"
    If SoloEfectos Then Sql = "="
    Sql = "Select codigo,anyo,codmacta,tiporem  from remesas where  tiporem " & Sql
    
    Sql = Sql & " 1  "
    Sql = Sql & " AND lcase(descripcion)<>'traspasada' AND (situacion ='Q' or situacion ='Y') ORDER BY codmacta,1,2 "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""
    Set Col = New Collection
    Msg = ""
    While Not miRsAux.EOF
        If Msg <> miRsAux!codmacta Then
            '
            '           tiporem|dias                Resto. Remesas
            If Msg <> "" Then Col.Add Sql
            
            If SoloEfectos Then
                Sql = "'0|0|'"
            Else
                Sql = "concat( pagaredias,'|',talondias,'|')"
            End If
            Msg = DevuelveDesdeBD(Sql, "bancos", "codmacta", miRsAux!codmacta, "T")
            
            
            
            
            If Msg = "" Then Msg = "0|0|"
            
            
            If Msg = "" Then Err.Raise 513, , "No existe banco?" & miRsAux!codmacta
            
            If miRsAux!Tiporem = 2 Then
                Sql = Format(RecuperaValor(Msg, 1), "000")
            Else
                Sql = Format(RecuperaValor(Msg, 2), "000")
            End If
            Msg = miRsAux!codmacta
            
        End If
        
        Sql = Sql & ", (" & miRsAux!Codigo & "," & miRsAux!Anyo & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Sql <> "" Then Col.Add Sql
        
        
    For I = 1 To Col.Count
        Sql = Col.Item(I)
        J = Val(Mid(Sql, 1, 3))
        Sql = Mid(Sql, 5)
        
        Msg = "select distinct codrem,anyorem "
        Msg = Msg & " from cobros where (codrem,anyorem) in ("
        Msg = Msg & Sql
        Msg = Msg & ") and date_add(fecvenci, interval " & J & " day) <=now()"
        Msg = Msg & " order by fecvenci"
        miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            QueRemesasMostrarEliminarRiesgoTalPag2 = QueRemesasMostrarEliminarRiesgoTalPag2 & ", (" & miRsAux!Codrem & "," & miRsAux!Anyorem & ")"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
    Next I
        
eHayQueMostrarEliminarRiesgoTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set Col = Nothing
    Msg = ""
End Function







