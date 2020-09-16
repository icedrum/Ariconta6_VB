Attribute VB_Name = "ModAvnics"
Option Explicit

Private BaseImp As Currency
Private IvaImp As Currency

Private CCoste As String


Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency


Public Function InsertarCabAsientoDia(Diario As String, Asiento As String, Fecha As String, Obs As String, caderr As String, bd As Byte) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Nulo2 As String
Dim Nulo3 As String

    On Error GoTo eInsertar
       
    cad = Format(Diario, "00") & ", " & DBSet(Fecha, "F") & "," & Format(Asiento, "000000") & ","
    cad = cad & DBSet(Obs, "T")
    cad = cad & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'AVNICS'"
    cad = "(" & cad & ")"

    'Insertar en la contabilidad
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) "
    Sql = Sql & " VALUES " & cad
    
    
    Conn.Execute Sql
    

eInsertar:
    If Err.Number <> 0 Then
        InsertarCabAsientoDia = False
        caderr = Err.Description
    Else
        InsertarCabAsientoDia = True
    End If
End Function


Public Function InsertarLinAsientoDia(cad As String, caderr As String, bd As Byte) As Boolean
' el Tipo me indica desde donde viene la llamada
' tipo = 0 srecau.codmacta
' tipo = 1 scaalb.codmacta

Dim Rs As ADODB.Recordset
Dim Aux As String
Dim Sql As String
Dim i As Byte
Dim totimp As Currency, ImpLinea As Currency

    On Error GoTo EInLinea

    Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, "
    Sql = Sql & " ampconce, timporteD, timporteH, codccost, ctacontr, idcontab, punteada) "
    Sql = Sql & " VALUES " & cad
    
    Conn.Execute Sql
    

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinAsientoDia = False
        caderr = Err.Description
    Else
        InsertarLinAsientoDia = True
    End If
End Function

Public Function ActualizarMovimientos(cadwhere As String, caderr As String) As Boolean
'Poner el movimiento como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE movim SET intconta=1 "
    Sql = Sql & " WHERE " & cadwhere

    Conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarMovimientos = False
        caderr = Err.Description
    Else
        ActualizarMovimientos = True
    End If
End Function



Public Function DesBloqueoManual(cadTABLA As String) As Boolean
Dim Sql As String
'Solo me interesa la tabla
On Error Resume Next

        Sql = "DELETE FROM zbloqueos WHERE codusu=" & vUsu.Codigo & " and tabla='" & cadTABLA & "'"
        Conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
End Function



Public Function ComprobarCtaContable(cadTABLA As String, Opcion As Byte, Optional cadwhere As String, Optional bd As Byte, Optional Tipo As Byte) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContable = False
    
    Sql = "SELECT codmacta FROM cuentas "
    Sql = Sql & " WHERE apudirec='S'"
    If cadG <> "" Then Sql = Sql & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, Conn, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                Sql = "SELECT DISTINCT avnic.codmacta, avnic.codavnic  "
                Sql = Sql & " FROM avnic, movim  "
                Sql = Sql & " where " & cadwhere & " and avnic.codavnic = movim.codavnic and avnic.anoejerc = movim.anoejerc "
        End If
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not Rs.EOF 'And b
            Sql = "codmacta= " & DBLet(Rs.Fields(0).Value, "T") 'DBSet(RS.Fields(0).Value, "T") '& " and apudirec='S' "
            enc = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", Rs.Fields(0).Value, "T")
                 
            If enc = "" Then
                B = False 'no encontrado
                If Opcion = 1 Then
                        Sql = Rs!codmacta & " del Código Avnic " & Format(Rs!codavnic, "0000000")
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                End If
            End If
                
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not B Then
            ComprobarCtaContable = False
        Else
            ComprobarCtaContable = True
        End If
    Else
        ComprobarCtaContable = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function

' ### [Monica] 07/05/2007
Public Function InsertarEnTesoreriaNew(Fechamov As String, FecVenci As String, codavnic As String, anoejerc As Integer, Codmacta2 As String, Concepto As String, forpa As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rsx As ADODB.Recordset
Dim Sql As String, text1csb As String, text2csb As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Rs3 As ADODB.Recordset
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim cadvalues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci1 As Date
Dim ImpVenci As Single
Dim i As Byte
Dim CodmacBPr As String
Dim cadWHERE2 As String
Dim DigConta As String

Dim vvIban As String


    On Error GoTo EInsertarTesoreriaNew

    B = False
    InsertarEnTesoreriaNew = False
    CadValues = ""
    cadvalues2 = ""

    Sql = "select * from movim where fechamov = " & DBSet(Fechamov, "F") & " and codavnic = " & DBSet(codavnic, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
    
        text1csb = "'Nro:" & Format(codavnic, "0000000") & " " & Format(Fechamov, "dd/mm/yy")
        text1csb = text1csb & " de " & DBSet(Rs!timporte, "N") & "'"
        text2csb = Concepto
        
              
        Sql4 = "select codmacta, codbanco, codsucur, digcontr, cuentaba, iban, nombrper nommacta,nomcalle dirdatos, "
        Sql4 = Sql4 & " poblacio despobla,codposta,provinci desprovi,nifperso nifdatos from avnic "
        Sql4 = Sql4 & " where codavnic = " & codavnic & " and anoejerc = " & DBSet(anoejerc, "N")
        
        Set Rs4 = New ADODB.Recordset
        Rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs4.EOF Then
            DigConta = DBLet(Rs4!digcontr, "T")
            If DBLet(Rs4!digcontr, "T") = "**" Then DigConta = "00"
        
            CadValuesAux2 = "("
            CadValuesAux2 = CadValuesAux2 & DBSet("1", "T") & ","
            CadValuesAux2 = CadValuesAux2 & DBSet(Rs4!codmacta, "T") & "," & DBSet(codavnic, "N") & ", " & DBSet(Fechamov, "F") & ", 1,"
            
            cadvalues2 = CadValuesAux2 & DBSet(forpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rs!timporte, "N") & ","
            
            vvIban = DBLet(Rs4!IBAN, "T") & Format(Rs4!codbanco, "0000") & Format(Rs4!codsucur, "0000") & Format(DigConta, "00") & Right("0000000000" & DBLet(Rs4!cuentaba, "T"), 10)
        
            cadvalues2 = cadvalues2 & DBSet(Codmacta2, "T") & "," & text1csb & "," & DBSet(text2csb, "T") & ","
            cadvalues2 = cadvalues2 & DBSet(vvIban, "T", "S") & ", "
            cadvalues2 = cadvalues2 & DBSet(Rs4!Nommacta, "T", "S") & "," & DBSet(Rs4!dirdatos, "T", "S") & "," & DBSet(Rs4!desPobla, "T", "S") & ","
            cadvalues2 = cadvalues2 & DBSet(Rs4!codposta, "T", "S") & "," & DBSet(Rs4!desProvi, "T", "S") & "," & DBSet(Rs4!nifdatos, "T", "S") & ",'ES')"
            
            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
            Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
            
            
            Sql = Sql & " VALUES " & cadvalues2
            Conn.Execute Sql
        End If

    End If

    B = True

EInsertarTesoreriaNew:
    If Err.Number <> 0 Then B = False
    InsertarEnTesoreriaNew = B
End Function


Private Sub InsertarError(Cadena As String)
Dim Sql As String

    Sql = "insert into tmperrcomprob values ('" & Cadena & "')"
    Conn.Execute Sql

End Sub


