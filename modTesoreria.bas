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

