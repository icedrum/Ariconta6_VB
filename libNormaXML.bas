Attribute VB_Name = "libNormaXML"
Option Explicit



'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'
'
'
' SEPA en XML
'
'
'
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////


Dim NFic As Integer   'Para no tener que pasarlo a todas las funciones

Private Function XML(CADENA As String) As String
Dim i As Integer
Dim Aux As String
Dim Le As String
Dim C As Integer
    'Carácter no permitido en XML  Representación ASCII
    '& (ampersand)          &amp;
    '< (menor que)          &lt;
    ' > (mayor que)         &gt;
    '“ (dobles comillas)    &quot;
    '' (apóstrofe)          &apos;
    
    'La ISO recomienda trabajar con los carcateres:
    'a b c d e f g h i j k l m n o p q r s t u v w x y z
    'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z
    '0 1 2 3 4 5 6 7 8 9
    '/ - ? : ( ) . , ' +
    'Espacio
    Aux = ""
    For i = 1 To Len(CADENA)
        Le = Mid(CADENA, i, 1)
        C = Asc(Le)
        
        
        Select Case C
        Case 40 To 57
            'Caracteres permitidos y numeros
            
        Case 65 To 90
            'Letras mayusculas
            
        Case 97 To 122
            'Letras minusculas
            
        Case 32
            'espacio en balanco
            
        Case Else
            Le = " "
        End Select
        Aux = Aux & Le
    Next
    XML = Aux
End Function


Public Function GeneraFicheroNorma34SEPA_XML(CIF As String, Fecha As Date, CuentaPropia2 As String, NumeroTransferencia As Long, Pagos As Boolean, ConceptoTr As String, Anyo As String, IdFich As String, AgrupaVtos As Boolean) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim Cad As String
Dim Aux As String
Dim SufijoOEM As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean

Dim CuentasAgrupadas As String
Dim RepeticionBucle As Byte
Dim nR As Byte
Dim cLineas As Collection
Dim LineaInsecionSumatorios As Byte
    On Error GoTo EGen3
    GeneraFicheroNorma34SEPA_XML = False
    
'    NFic = -1
    
    
    'Cargamos la cuenta
    Cad = "Select * from bancos where codmacta='" & CuentaPropia2 & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        Cad = ""
    Else
        If IsNull(miRsAux!IBAN) Then
            Cad = ""
        Else
            SufijoOEM = "000" ''Sufijo3414
            Cad = miRsAux!IBAN
            If DBLet(miRsAux!Sufijo3414, "T") <> "" Then SufijoOEM = Right("000" & miRsAux!Sufijo3414, 3)
            CuentaPropia2 = Cad
        End If
        
        
    End If
    miRsAux.Close
  
    If Cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    
    Set cLineas = New Collection
    
    
    
    cLineas.Add "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    cLineas.Add "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"">"
    cLineas.Add "<CstmrCdtTrfInitn>"
    cLineas.Add "   <GrpHdr>"
    Cad = "TRAN" & IIf(Pagos, "PAG", "ABO") & Format(NumeroTransferencia, "000000") & "F" & Format(Now, "yyyymmddThhnnss")
    IdFich = Cad
    cLineas.Add "      <MsgId>" & Cad & "</MsgId>"
    cLineas.Add "      <CreDtTm>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</CreDtTm>"
    LineaInsecionSumatorios = 6
    
    'cLineas.Add "      <NbOfTxs>" & RecuperaValor(Aux, 1) & "</NbOfTxs>"
    'cLineas.Add "      <CtrlSum>" & TransformaComasPuntos(RecuperaValor(Aux, 2)) & "</CtrlSum>"
    
     
    cLineas.Add "      <InitgPty>"
    cLineas.Add "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    cLineas.Add "         <Id>"
    Cad = Mid(CIF, 1, 1)
    
    EsPersonaJuridica2 = Not IsNumeric(Cad)

    Cad = "PrvtId"
    If EsPersonaJuridica2 Then Cad = "OrgId"
    
    cLineas.Add "           <" & Cad & ">"
    cLineas.Add "               <Othr>"
    cLineas.Add "                  <Id>" & CIF & SufijoOEM & "</Id>"
    cLineas.Add "               </Othr>"
    cLineas.Add "           </" & Cad & ">"
    
    cLineas.Add "         </Id>"
    cLineas.Add "      </InitgPty>"
    cLineas.Add "   </GrpHdr>"

    cLineas.Add "   <PmtInf>"
    
    cLineas.Add "      <PmtInfId>" & Format(Now, "yyyymmddhhnnss") & CIF & "</PmtInfId>"
    cLineas.Add "      <PmtMtd>TRF</PmtMtd>"
    cLineas.Add "      <ReqdExctnDt>" & Format(Fecha, "yyyy-mm-dd") & "</ReqdExctnDt>"
    cLineas.Add "      <Dbtr>"
    
     'Nombre
    miRsAux.Open "Select siglasvia ,direccion ,numero ,codpobla,pobempre,provempre,provincia from empresa2"
    Cad = Cad & FrmtStr(vEmpresa.nomempre, 70)
    If miRsAux.EOF Then Err.Raise 513, , "Error obteniendo datos empresa(empresa2)"
    
    cLineas.Add "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    cLineas.Add "         <PstlAdr>"
    cLineas.Add "            <Ctry>ES</Ctry>"

    Cad = DBLet(miRsAux!siglasvia, "T") & " " & miRsAux!Direccion & " " & DBLet(miRsAux!numero, "T") & " "
    Cad = Cad & Trim(DBLet(miRsAux!CodPobla, "T") & " " & miRsAux!pobempre) & " "
    Cad = Cad & DBLet(miRsAux!provincia, "T")
    miRsAux.Close
    cLineas.Add "            <AdrLine>" & XML(Trim(Cad)) & "</AdrLine>"
    
    cLineas.Add "         </PstlAdr>"
    cLineas.Add "         <Id>"
    
    Aux = "PrvtId"
    If EsPersonaJuridica2 Then Aux = "OrgId"
   
    
    cLineas.Add "            <" & Aux & ">"
    
    cLineas.Add "               <Othr>"
    cLineas.Add "                  <Id>" & CIF & SufijoOEM & "</Id>"
    cLineas.Add "               </Othr>"
    cLineas.Add "            </" & Aux & ">"
    cLineas.Add "         </Id>"
    cLineas.Add "    </Dbtr>"
    
    
    cLineas.Add "    <DbtrAcct>"
    cLineas.Add "       <Id>"
    cLineas.Add "          <IBAN>" & Trim(CuentaPropia2) & "</IBAN>"
    cLineas.Add "       </Id>"
    cLineas.Add "       <Ccy>EUR</Ccy>"
    cLineas.Add "    </DbtrAcct>"
    cLineas.Add "    <DbtrAgt>"
    cLineas.Add "       <FinInstnId>"
    
    Cad = Mid(CuentaPropia2, 5, 4)
    Cad = DevuelveDesdeBD("bic", "bics", "entidad", Cad)
    cLineas.Add "          <BIC>" & Trim(Cad) & "</BIC>"
    cLineas.Add "       </FinInstnId>"
    cLineas.Add "    </DbtrAgt>"
    
    
    
    

    
    'J = numerototalregistro
    'ImpEfe = total remesa
    
    RepeticionBucle = 1
    If AgrupaVtos Then
    
    
        CuentasAgrupadas = ""
    
    
        Cad = "select codmacta,count(*) from "
        If Pagos Then
            Cad = Cad & " pagos where nrodocum = " & NumeroTransferencia
            Cad = Cad & " and anyodocum = " & Anyo
        Else
            Cad = Cad & " cobros where transfer = " & NumeroTransferencia
            Cad = Cad & " and anyorem = " & Anyo
        End If
                
        Cad = Cad & " group by codmacta having count(*) >1"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            CuentasAgrupadas = CuentasAgrupadas & ", '" & miRsAux!codmacta & "'"
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
    
        If CuentasAgrupadas = "" Then Err.Raise 513, , "NO hay vencimientos para agrupar"
        CuentasAgrupadas = Mid(CuentasAgrupadas, 2)
        RepeticionBucle = 2
    End If
    
    
    
    'enero18. uNA VEZ PARA LOS NORMALES Y TRA para los agrupadaos
    'Para ello abrimos la tabla tmpNorma34
    Regs = 0
    For nR = 1 To RepeticionBucle
        If Pagos Then
    
            Cad = "Select mid(pagos.iban,5,4) as entidad,mid(pagos.iban,9,4) as oficina,mid(pagos.iban,15,10) cuentaba,mid(pagos.iban,13,2) as CC,pagos.iban, "
            Cad = Cad & "nomprove nommacta,domprove dirdatos,cpprove codposta,pobprove despobla,pagos.codmacta,codpais"
            If nR = 1 Then
                Cad = Cad & ",impefect,0 Gastos,imppagad,numorden,text1csb,text2csb"
            Else
                Cad = Cad & ",sum(impefect) impefect,0 Gastos,sum(imppagad) imppagad, count(*) numorden,"
                Cad = Cad & "GROUP_CONCAT( numfactu separator ',') text1csb,"
                Cad = Cad & " concat('Num. Vtos:' , count(*)) text2csb"
            End If
            Cad = Cad & ",proprove desprovi,NUmSerie,numfactu,fecfactu,bic,nifprove nifdatos from pagos"
            
            Cad = Cad & " left join bics on mid(pagos.iban,5,4)=bics.entidad "
            Cad = Cad & " WHERE nrodocum =" & NumeroTransferencia & " and anyodocum = " & DBSet(Anyo, "N")
        Else
            'ABONOS
            
            Cad = "Select mid(cobros.iban,5,4) as entidad,mid(cobros.iban,9,4) as oficina,mid(cobros.iban,15,10) cuentaba,mid(cobros.iban,13,2) as CC,cobros.iban"
            Cad = Cad & ",nomclien nommacta,domclien dirdatos,cpclien codposta,pobclien despobla,cobros.codmacta,codpais,proclien desprovi"
            Cad = Cad & " ,NUmSerie,numfactu,fecfactu,bic,nifclien nifdatos,"
            If nR = 1 Then
                Cad = Cad & "impvenci,gastos,impcobro,numorden,text33csb,text41csb"
            Else
                Cad = Cad & "sum(impvenci) impvenci,sum(coalesce(Gastos,0)) Gastos,sum(impcobro) impcobro, count(*) numorden,"
                Cad = Cad & "GROUP_CONCAT( numfactu separator ',') text33csb,"
                Cad = Cad & " concat('Num. Vtos:' , count(*)) text41csb"
            End If
            Cad = Cad & " from cobros LEFT JOIN bics on mid(cobros.iban,5,4)=bics.entidad "
            Cad = Cad & " WHERE transfer =" & NumeroTransferencia & " and anyorem = " & DBSet(Anyo, "N")
        End If
        
        If AgrupaVtos Then
            Cad = Cad & " AND "
            If nR = 1 Then Cad = Cad & " NOT "
            Cad = Cad & "codmacta IN  (" & CuentasAgrupadas & ")"
            If nR = 2 Then Cad = Cad & " GROUP BY codmacta"
        End If
        miRsAux.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
        While Not miRsAux.EOF
            cLineas.Add "   <CdtTrfTxInf>"
            cLineas.Add "      <PmtId>"
            
             
            If nR = 2 Then
                'AGRUPADO
                
                Aux = Right(miRsAux!codmacta, 3) & "GR" & Format(Now, "yymmddhhnnss") & "R" & Format(miRsAux!numorden, "000")
            Else
                If Pagos Then
                    'numfactu fecfactu numorden
                    Aux = FrmtStr(miRsAux!NumFactu, 10)
                    Aux = Aux & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
                
                Else
                    'fecfaccl
                    Aux = FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!NumFactu, "00000000")
                    Aux = Aux & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
                End If
            End If
            cLineas.Add "         <EndToEndId>" & Aux & "</EndToEndId>"
            cLineas.Add "      </PmtId>"
            cLineas.Add "      <PmtTpInf>"
            'Enero 2018.
                'Esto NO ex correcto.
                ' Lo que hay que llevar es impvenci,cobro ya que los pagos -cobros no hay parcial. Simplemente ha que
                
            If Pagos Then
                'Im = DBLet(miRsAux!imppagad, "N")
                Im = 0
                Im = miRsAux!ImpEfect - Im
                Aux = miRsAux!codmacta
    
            Else
                'Im = Abs(miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N"))
                Im = Abs(miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N"))
                Aux = miRsAux!codmacta
            End If
            
            'Persona fisica o juridica
            Cad = Mid(miRsAux!nifdatos, 1, 1)
            EsPersonaJuridica2 = Not IsNumeric(Cad)
            'Como da problemas Cajamar, siempre ponemos Perosna juridica. Veremos
            EsPersonaJuridica2 = True
            
            
            Importe = Importe + Im
            Regs = Regs + 1
            
            cLineas.Add "          <SvcLvl><Cd>SEPA</Cd></SvcLvl>"
            'cLineas.Add  "          <LclInstrm><Cd>SDCL</Cd></LclInstrm>"
            If ConceptoTr = "1" Then
                Aux = "SALA"
            ElseIf ConceptoTr = "0" Then
                Aux = "PENS"
            Else
                Aux = "TRAD"
            End If
            cLineas.Add "          <CtgyPurp><Cd>" & Aux & "</Cd></CtgyPurp>"
            cLineas.Add "       </PmtTpInf>"
            cLineas.Add "       <Amt>"
            Cad = Format(Im, "#.00")
            cLineas.Add "          <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(Cad) & "</InstdAmt>"
            cLineas.Add "       </Amt>"
            cLineas.Add "       <CdtrAgt>"
            cLineas.Add "          <FinInstnId>"
            Cad = DBLet(miRsAux!BIC, "T")
            If Cad = "" Then Err.Raise 513, , "No existe BIC: " & miRsAux!Nommacta & vbCrLf & "Entidad: " & miRsAux!Entidad
            cLineas.Add "             <BIC>" & DBLet(miRsAux!BIC, "T") & "</BIC>"
            cLineas.Add "          </FinInstnId>"
            cLineas.Add "       </CdtrAgt>"
            cLineas.Add "       <Cdtr>"
            cLineas.Add "          <Nm>" & XML(miRsAux!Nommacta) & "</Nm>"
            
            
            'Como cajamar da problemas, lo quitamos para todos
            'cLineas.Add  "          <PstlAdr>"
            '
            'Cad = "ES"
            'If Not IsNull(miRsAux!PAIS) Then Cad = Mid(miRsAux!PAIS, 1, 2)
            'cLineas.Add  "              <Ctry>" & Cad & "</Ctry>"
            '
            'If Not IsNull(miRsAux!dirdatos) Then cLineas.Add  "              <AdrLine>" & XML(miRsAux!dirdatos) & "</AdrLine>"
            'Cad = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
            'If Cad <> "" Then cLineas.Add  "              <AdrLine>" & Cad & "</AdrLine>"
            'If Not IsNull(miRsAux!desprovi) Then cLineas.Add  "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
            'cLineas.Add  "           </PstlAdr>"
            
            
            
            cLineas.Add "           <Id>"
            Aux = "PrvtId"
            If EsPersonaJuridica2 Then Aux = "OrgId"
          
            cLineas.Add "               <" & Aux & ">"
            cLineas.Add "                  <Othr>"
            cLineas.Add "                     <Id>" & miRsAux!nifdatos & "</Id>"
            'Da problemas.... con Cajamar
            'cLineas.Add  "                     <Issr>NIF</Issr>"
            cLineas.Add "                  </Othr>"
            cLineas.Add "               </" & Aux & ">"
            cLineas.Add "           </Id>"
            cLineas.Add "        </Cdtr>"
            cLineas.Add "        <CdtrAcct>"
            cLineas.Add "           <Id>"
            cLineas.Add "              <IBAN>" & IBAN_Destino & "</IBAN>"
            cLineas.Add "           </Id>"
            cLineas.Add "        </CdtrAcct>"
            cLineas.Add "      <Purp>"
            
            If ConceptoTr = "1" Then
                Aux = "SALA"
            ElseIf ConceptoTr = "0" Then
                Aux = "PENS"
            Else
                Aux = "TRAD"
            End If
            
            cLineas.Add "         <Cd>" & Aux & "</Cd>"
            cLineas.Add "      </Purp>"
            cLineas.Add "      <RmtInf>"
            'cLineas.Add  "      <Ustrd>ESTE ES EL CONCEPTO, POR TANTO NO SE SI SERA EL TEXTO QUE IRA DONDE TIENE QUE IR, O EN OTRO LADAO... A SABER. LO QUE ESTA CLARO ES QUE VA.</Ustrd>
            
            If nR = 2 Then
                'AGRUPADO
                If Pagos Then
                    ''`text1csb` `text2csb`
                    K = Len(miRsAux!text2csb)
                    K = 140 - K - 1
                    
                    Aux = "Fras: " & DBLet(miRsAux!text1csb, "T")
                    If Len(Aux) > K Then Aux = Mid(Aux, 1, K - 4) & "..."
                    Aux = Aux & " " & miRsAux!text2csb
                Else
                    K = Len(miRsAux!text41csb)
                    K = 140 - K - 1
                    Aux = "Fras: " & DBLet(miRsAux!text33csb, "T")
                    If Len(Aux) > K Then Aux = Mid(Aux, 1, K - 4) & "..."
                    Aux = Aux & " " & miRsAux!text41csb
                   
                End If
                
            Else
                If Pagos Then
                    ''`text1csb` `text2csb`
                    Aux = DBLet(miRsAux!text1csb, "T") & " " & DBLet(miRsAux!text2csb, "T")
                Else
                    '`text33csb` `text41csb`
                    Aux = DBLet(miRsAux!text33csb, "T") & " " & DBLet(miRsAux!text41csb, "T")
                End If
            End If
            If Trim(Aux) = "" Then Aux = miRsAux!Nommacta
            cLineas.Add "         <Ustrd>" & XML(Trim(Aux)) & "</Ustrd>"
            cLineas.Add "      </RmtInf>"
            cLineas.Add "   </CdtTrfTxInf>"
     
           
        
                
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    Next nR 'Repeticion bucle
    cLineas.Add "   </PmtInf>"
    cLineas.Add "</CstmrCdtTrfInitn></Document>"
    
    
    
    
    
    
    
    
    
    NFic = FreeFile
    CerrarFichero NFic
    Open App.Path & "\norma34.txt" For Output As #NFic
    
    
    
    
    For J = 1 To LineaInsecionSumatorios
        Print #NFic, cLineas.Item(J)
    Next J
    
    'TOTALES
    Print #NFic, "      <NbOfTxs>" & Regs & "</NbOfTxs>"
    Aux = Format(Importe, "###0.00")
    Print #NFic, "      <CtrlSum>" & TransformaComasPuntos(Aux) & "</CtrlSum>"
        
    
    For J = LineaInsecionSumatorios + 1 To cLineas.Count
        Print #NFic, cLineas.Item(J)
    Next J
    
    Close #NFic
    
    
    
    Set miRsAux = Nothing
    
    
    
    NFic = -1
    
    If Regs > 0 Then GeneraFicheroNorma34SEPA_XML = True
    Exit Function
    
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
'    If NFic > 0 Then Close (NFic)
    CerrarFichero NFic
    
End Function





Private Sub CerrarFichero(nFile As Integer)

    On Error Resume Next
    
    Close #nFile
    
    Err.Clear

End Sub






Private Function IBAN_Destino() As String
    IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 5, 4) ' Código de entidad receptora
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 9, 4) ' Código de oficina receptora
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 13, 2) ' Dígitos de control
    IBAN_Destino = IBAN_Destino & Mid(DBLet(miRsAux!IBAN, "T"), 15, 10) ' Código de cuenta
End Function












'Devolucion SEPA
'
Public Sub ProcesaFicheroDevolucionSEPA_XML(Fichero As String, ByRef Remesa As String)
Dim aux2 As String  'Para buscar los vencimientos
Dim FinLecturaLineas As Boolean

Dim ErroresVto As String

Dim posicion As Long
Dim L2 As Long
Dim SQL As String
Dim ContenidoFichero As String
Dim NF As Integer
Dim CadenaComprobacionDevueltos As String  'cuantos|importe|


    On Error GoTo eProcesaCabeceraFicheroDevolucionSEPA_XML
    Remesa = ""
    
    
    
   

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
    Set miRsAux = New ADODB.Recordset
    
    'Vamos a obtener el ID de la remesa  enviada
    ' Buscaremos la linea
    'Idententificacion propia  Ejemplo: <MsgId>PRE2015093012481641020RE10000802015</MsgId>  de donde RE mesa, 1 tipo 000080 Nº   ano;2015
    posicion = PosicionEnFichero(1, ContenidoFichero, "<CstmrPmtStsRpt>")
    
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlMsgId>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlMsgId>")
    
    aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
    aux2 = Mid(aux2, InStr(10, aux2, "RE") + 3) 'QUTIAMO EL RE y ye tipo RE1(ejemp)
    
    'Los 4 ultimos son año
    Remesa = Mid(aux2, 1, 6) & "|" & Mid(aux2, 7, 4) & "|"
    
    
    'Voy a buscar el numero total de vencimientos que devuelven y el importe total(comproabacion ultima
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlPmtInfAndSts>")
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlNbOfTxs>")
    '<OrgnlNbOfTxs>1</OrgnlNbOfTxs>
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlNbOfTxs>")
    CadenaComprobacionDevueltos = Mid(ContenidoFichero, posicion, L2 - posicion)
    
    '<OrgnlCtrlSum>5180.98</OrgnlCtrlSum>
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlCtrlSum>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlCtrlSum>")
    CadenaComprobacionDevueltos = CadenaComprobacionDevueltos & Mid(ContenidoFichero, posicion, L2 - posicion)
            
    
    
    'Primera comprobacion. Existe la remesa obtenida
    
    
    'Vamos con los vtos  4300106840T  0001180220150925001

    Do
        posicion = InStr(posicion, ContenidoFichero, "<TxInfAndSts>")
        If posicion > 0 Then
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlEndToEndId>")
            aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
            
            'Id del recibo devuleto. Ejemplo
            '4300106840T  0001180220150925001
            ' asi es como se monta el el generador de remesa
            '           FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            '           Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            
            SQL = "Select codrem,anyorem,siturem from cobros where fecfactu='" & Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2)
            SQL = SQL & "' AND numserie = '" & Trim(Mid(aux2, 11, 3)) & "' AND numfactu = " & Val(Mid(aux2, 14, 8)) & " AND numorden=" & Mid(aux2, 30, 3)

            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = Mid(SQL, InStr(1, UCase(SQL), " WHERE ") + 7)
            SQL = Replace(SQL, "fecfactu", "F.Fac:")
            SQL = Replace(SQL, "numserie", "Serie:")
            SQL = Replace(SQL, "numfactu", "NºFac:")
            SQL = Replace(SQL, "numorden", "Ord:")
            SQL = Replace(SQL, "AND", ""): SQL = Replace(SQL, "=", "")
            SQL = "Vto no encontrado: " & Mid(SQL, InStr(1, UCase(SQL), " WHERE ") + 7)
            If Not miRsAux.EOF Then
                If IsNull(miRsAux!CodRem) Then
                    SQL = "Vencimiento sin Remesa: " & aux2
                Else
        
                    SQL = ""
                End If
            End If
            miRsAux.Close
            
            If SQL <> "" Then ErroresVto = ErroresVto & vbCrLf & SQL
            
            
            posicion = InStr(posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            
        
        Else
           posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
        
    Loop Until posicion > Len(ContenidoFichero)
    

    If ErroresVto <> "" Then
        MsgBox ErroresVto, vbExclamation
        Remesa = ""
    Else
        


    
        'En aux2 tendre codrem|anñorem|
        aux2 = RecuperaValor(Remesa, 1) & " AND anyo = " & RecuperaValor(Remesa, 2)
        aux2 = "Select situacion from remesas where codigo = " & aux2
        miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            aux2 = "-No se encuentra remesa"
            
        Else
            'Si que esta.
            'Situacion
            If CStr(miRsAux!Situacion) <> "Q" And CStr(miRsAux!Situacion) <> "Y" Then
                aux2 = "- Situacion incorrecta : " & miRsAux!Situacion
            Else
                aux2 = "" 'TODO OK
            End If
        End If

        If aux2 <> "" Then
            aux2 = aux2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
            aux2 = Replace(aux2, " AND ", " ")
            aux2 = Replace(aux2, "anyo", "año")
            ErroresVto = ErroresVto & vbCrLf & aux2
        End If
        miRsAux.Close

    
    


        If ErroresVto <> "" Then
            aux2 = "Error remesas " & vbCrLf & String(30, "=") & ErroresVto
            MsgBox aux2, vbExclamation

            'Pongo REMESA=""
            Remesa = "" 'para que no continue el preoceso de devolucion
        End If

    End If
    Set miRsAux = Nothing
    Exit Sub
eProcesaCabeceraFicheroDevolucionSEPA_XML:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion SEPA XML" & Err.Description
    Set miRsAux = Nothing
End Sub

'Si no se encuentra lo que busco saltara un error
Private Function PosicionEnFichero(ByVal Inicio As Long, ContenidoDelFichero As String, QueBusco As String) As Long
    PosicionEnFichero = InStr(Inicio, ContenidoDelFichero, QueBusco)
    If PosicionEnFichero = 0 Then
        Err.Raise 513, , "No se encuentra cadena: " & QueBusco
    Else
        If InStr(1, QueBusco, "</") Then
            'PosicionEnFichero = PosicionEnFichero - Len(QueBusco)
        Else
            PosicionEnFichero = PosicionEnFichero + Len(QueBusco)
        End If
    End If
        
End Function


'XML
Public Sub ProcesaLineasFicheroDevolucionXML(Fichero As String, ByRef Listado As Collection)
Dim NF As Integer
Dim ContenidoFichero As String
Dim posicion As Long
Dim L2 As Long
Dim aux2 As String

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
   
    posicion = 1
    Do
        posicion = InStr(posicion, ContenidoFichero, "<TxInfAndSts>")
        If posicion > 0 Then
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlEndToEndId>")
            aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
            
            'Id del recibo devuleto. Ejemplo
            '4300106840T  0001180220150925001
            ' asi es como se monta el el generador de remesa
            '           FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            '           Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            
            'Vamos a guardar en el col la linea en formato antiguo SEPA y asi no toco el programa
            'M  0330047820131201001   430000061
            aux2 = Mid(aux2, 11, 23) & "   " & Mid(aux2, 1, 10)
            Listado.Add aux2
            posicion = InStr(posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            
        
        Else
           posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
    Loop Until posicion > Len(ContenidoFichero)
    
End Sub


Public Sub LeerLineaDevolucionSEPA_XML(Fichero As String, ByRef Remesa As String, ByRef lwCobros As ListView)
Dim aux2 As String  'Para buscar los vencimientos
Dim AUX3 As String
Dim FinLecturaLineas As Boolean

Dim ErroresVto As String

Dim posicion As Long
Dim L2 As Long
Dim SQL As String
Dim ContenidoFichero As String
Dim NF As Integer
Dim CadenaComprobacionDevueltos As String  'cuantos|importe|

Dim VtoEncontrado As Boolean
Dim DatosXMLVto As String
Dim Itm As ListItem
Dim Rs As ADODB.Recordset

Dim RegistroErroneo As Boolean
Dim RemesasNoContabilizadas As String
Dim TipoPopular As Boolean
Dim EtiquetaBuscar As String

    On Error GoTo eLeerLineaDevolucionSEPA_XML
    Remesa = ""
    
   

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
    
    
    
    
    'Comprobacion 1
    'El NIF del fichero enviado es el de la empresa
    'Lo busco acotandolo por etiquetas XML
    posicion = PosicionEnFichero(1, ContenidoFichero, "<OrgnlPmtInfAndSts>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlPmtInfAndSts>")
    If posicion > 0 And L2 > 0 Then
    
    
        'Mayo 2018. Devoluciones tipo CAIXA POPULAR
        'Llevan una cadena dentro del GrpHdr /GrpHdr
        ' InitgPty
        aux2 = "NIFMAL"
        TipoPopular = True
        If Not DevolucionTipoPopular(ContenidoFichero, aux2) Then
            TipoPopular = False
            
            'DEVOLUCION NORMAL.
            'LO que haciamos antes de mayo. NO se ha tocado nada
            
            '
            aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
            posicion = PosicionEnFichero(1, aux2, "<StsRsnInf>")
            L2 = PosicionEnFichero(posicion, aux2, "</StsRsnInf>")
            
            If posicion > 0 And L2 > 0 Then
                aux2 = Mid(aux2, posicion, L2 - posicion)
                posicion = PosicionEnFichero(1, aux2, "<Id>ES")   'de momento todos los clientes seran de españa
                L2 = PosicionEnFichero(posicion, aux2, "</Id>")
        
                aux2 = Mid(aux2, posicion, L2 - posicion)
            Else
                aux2 = "NIFMAL"
                
            End If
        End If
        
        If aux2 = "NIFMAL" Then Err.Raise 513, , "NIF empresa no encontrado en el fichero "
        
        If Len(aux2) > 5 Then
        
            
        
            SQL = DevuelveDesdeBD("nifempre", "empresa2", "1", "1")
            
            
            
            If InStr(1, aux2, SQL) = 0 Then

                Err.Raise 513, , "NIF empresa del fichero no coincide con el de la empresa en Ariconta"
            End If

        End If
    End If
    
    Set miRsAux = New ADODB.Recordset
    
    'Vamos a obtener el ID de la remesa  enviada
    ' Buscaremos la linea
    'Idententificacion propia  Ejemplo: <MsgId>PRE2015093012481641020RE10000802015</MsgId>  de donde RE mesa, 1 tipo 000080 Nº   ano;2015
    posicion = PosicionEnFichero(1, ContenidoFichero, "<CstmrPmtStsRpt>")
    
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlMsgId>")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, "</OrgnlMsgId>")
    
    aux2 = Mid(ContenidoFichero, posicion, L2 - posicion)
    aux2 = Mid(aux2, InStr(10, aux2, "RE") + 3) 'QUTIAMO EL RE y ye tipo RE1(ejemp)
    
    'Los 4 ultimos son año
    Remesa = Mid(aux2, 1, 6) & "|" & Mid(aux2, 7, 4) & "|"
    
    
    'Voy a buscar el numero total de vencimientos que devuelven y el importe total(comproabacion ultima
    
    
    posicion = PosicionEnFichero(posicion, ContenidoFichero, "<OrgnlPmtInfAndSts>")
    'NORMAL:<OrgnlNbOfTxs>1</OrgnlNbOfTxs>        POLUPAR:<DtldNbOfTxs>1</DtldNbOfTxs>
    EtiquetaBuscar = IIf(TipoPopular, "<DtldNbOfTxs>", "<OrgnlNbOfTxs>")
    posicion = PosicionEnFichero(posicion, ContenidoFichero, EtiquetaBuscar)
    EtiquetaBuscar = Replace(EtiquetaBuscar, "<", "</")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, EtiquetaBuscar)
    CadenaComprobacionDevueltos = Mid(ContenidoFichero, posicion, L2 - posicion) & "|"
    
    'NORMAL:<OrgnlCtrlSum>5180.98</OrgnlCtrlSum>     POPULAR:<DtldCtrlSum>2891.15</DtldCtrlSum>
    EtiquetaBuscar = IIf(TipoPopular, "<DtldCtrlSum>", "<OrgnlCtrlSum>")
    posicion = PosicionEnFichero(posicion, ContenidoFichero, EtiquetaBuscar)
    EtiquetaBuscar = Replace(EtiquetaBuscar, "<", "</")
    L2 = PosicionEnFichero(posicion, ContenidoFichero, EtiquetaBuscar)
    CadenaComprobacionDevueltos = CadenaComprobacionDevueltos & Mid(ContenidoFichero, posicion, L2 - posicion) & "|"
            
    'Primera comprobacion. Existe la remesa obtenida
    
    
    'Vamos con los vtos  4300106840T  0001180220150925001
    
    Dim jj As Long
    jj = 1
    Set Rs = New ADODB.Recordset
    RemesasNoContabilizadas = ""
    
    RemesasNoContabilizadas = "codigo =" & RecuperaValor(Remesa, 1) & " AND anyo =" & RecuperaValor(Remesa, 2) & " AND 1"
    RemesasNoContabilizadas = DevuelveDesdeBD("situacion", "remesas", RemesasNoContabilizadas, "1")
    If RemesasNoContabilizadas < "Q" Then
        If RemesasNoContabilizadas = "" Then
            RemesasNoContabilizadas = "no encontrada "
        Else
            RemesasNoContabilizadas = "incorrecta. " & RemesasNoContabilizadas
        End If
        RemesasNoContabilizadas = "Situación remesa " & RemesasNoContabilizadas & ".  " & Replace(Remesa, "|", " ")
        MsgBox RemesasNoContabilizadas, vbExclamation
        Exit Sub
        
    End If
    
    Do
        posicion = InStr(posicion, ContenidoFichero, "<TxInfAndSts>")
        If posicion > 0 Then
            L2 = PosicionEnFichero(posicion, ContenidoFichero, "</TxInfAndSts>")
            DatosXMLVto = Mid(ContenidoFichero, posicion, L2 - posicion)
            
            ContenidoFichero = Mid(ContenidoFichero, L2 + 14)
            
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            posicion = PosicionEnFichero(1, DatosXMLVto, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</OrgnlEndToEndId>")
            aux2 = Mid(DatosXMLVto, posicion, L2 - posicion)
            
            
            
            
            
            
            
            Set Itm = lwCobros.ListItems.Add(, "C" & jj)
            Itm.Text = Trim(Mid(aux2, 11, 3))  'miRsAux!NUmSerie
            
            Itm.SubItems(1) = Mid(aux2, 14, 8) ' numfactu
            Itm.SubItems(2) = Mid(aux2, 30, 3) ' miRsAux!numorden
            Itm.SubItems(3) = Mid(aux2, 1, 10) 'miRsAux!codmacta
            Itm.Tag = Format(Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2), "dd/mm/yyyy")
            
            Itm.SubItems(8) = RecuperaValor(Remesa, 1) ' remesa
            Itm.SubItems(9) = RecuperaValor(Remesa, 2) ' año de remesa
            Itm.SubItems(10) = DevuelveValor("select codmacta from remesas where codigo = " & RecuperaValor(Remesa, 1) & " and anyo = " & RecuperaValor(Remesa, 2))
            
            SQL = "select * from cobros where "
            SQL = SQL & " numserie = " & DBSet(Trim(Mid(aux2, 11, 3)), "T") & " and numfactu = " & DBSet(Val(Mid(aux2, 14, 8)), "N")
            SQL = SQL & " and fecfactu = '" & Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2) & "'"
            
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            VtoEncontrado = False
            If Not Rs.EOF Then
                Itm.SubItems(4) = DBLet(Rs!nomclien, "T")    'miRsAux!nomclien
                If Rs!Devuelto = 1 Then
                    Itm.Bold = True
                    Itm.ForeColor = vbRed
                End If
                VtoEncontrado = True
            Else
                Itm.SubItems(4) = " "    'miRsAux!nomclien    'AVISAR A MONICA--> Si no pones espacio en blanco cuando lo selecciona sale raro
            End If
            
            posicion = PosicionEnFichero(1, DatosXMLVto, "<InstdAmt Ccy=""EUR"">")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</InstdAmt>")
            AUX3 = Mid(DatosXMLVto, posicion, L2 - posicion)
            
            If posicion > 0 Then
            
            
                AUX3 = TransformaPuntosComas(AUX3)
                Itm.SubItems(5) = Format(AUX3, FormatoImporte)
                If VtoEncontrado Then
                    'El importe deberia coincidir. Si no lo marcariamos como error

                    
                    Dim ImporteRemesado As Currency

                    SQL = "select impcobro FROM cobros where "
                    SQL = SQL & " numserie = " & DBSet(Trim(Mid(aux2, 11, 3)), "T") & " and numfactu = " & DBSet(Val(Mid(aux2, 14, 8)), "N")
                    SQL = SQL & " and fecfactu = '" & Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2) & "' "
                    
                    ImporteRemesado = DevuelveValor(SQL)
                    
                    If ImporteRemesado <> AUX3 Then
                    
                        MsgBox "La factura " & DBSet(Trim(Mid(aux2, 11, 3)), "T") & "-" & DBSet(Val(Mid(aux2, 14, 8)), "N") & " de fecha " & Mid(aux2, 28, 2) & "/" & Mid(aux2, 26, 2) & "/" & Mid(aux2, 22, 4) & " es de " & ImporteRemesado & " euros", vbExclamation
                    
                    Else
                        
                    End If
                End If
            Else
                Itm.SubItems(5) = " "
            End If
           
           
            'Motivo devolucion   EJEMPLO
            '<Rsn>
            '   <Cd>AM04</Cd>
            '</Rsn>
            posicion = PosicionEnFichero(1, DatosXMLVto, "<Rsn>")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</Rsn>")
            aux2 = Mid(DatosXMLVto, posicion, L2 - posicion)
            
            posicion = PosicionEnFichero(1, DatosXMLVto, "<Cd>")
            L2 = PosicionEnFichero(posicion, DatosXMLVto, "</Cd>")
            If posicion > 0 And L2 > 0 Then
                aux2 = Mid(DatosXMLVto, posicion, L2 - posicion)
                
                aux2 = DevuelveDesdeBD("concat(codigo,' - ', descripcion)", "usuarios.wdevolucion", "codigo", aux2, "T")
                
                If aux2 = "" Then aux2 = " "
           
            Else
                'MOTIVO no encontrado
                'Ver por que
                'Ver que poner
                aux2 = " "
                
                
            End If
            Itm.SubItems(11) = aux2
           
           
            If Not VtoEncontrado Then
                Itm.ForeColor = vbRed
'                Itm.Ghosted = True
                For posicion = 1 To Itm.ListSubItems.Count
                    Debug.Print lwCobros.ColumnHeaders(posicion).Text & ":" & Itm.ListSubItems(posicion).Text
                    Itm.ListSubItems(posicion).ForeColor = vbRed
                Next
                
            Else
                Itm.Checked = True
            End If
            
            'posicion = InStr(posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            posicion = 1
            jj = jj + 1 'numero de item
            Rs.Close
        Else
           posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
        
    Loop Until posicion > Len(ContenidoFichero)
    
    
    Exit Sub
eLeerLineaDevolucionSEPA_XML:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion SEPA XML" & vbCrLf, Err.Description
    Set miRsAux = Nothing
    Set Rs = New ADODB.Recordset
           
End Sub










Public Function GrabarDisketteNorma19_SEPA_XML(NomFichero As String, Remesa_ As String, FecPre As String, TipoReferenciaCliente As Byte, Sufijo As String, FechaCobro As String, SEPA_EmpresasGraboNIF As Boolean, Norma19_15 As Boolean, DatosBanco As String, NifEmpresa As String, esAnticipoCredito As Boolean, ByRef IdGrabadoEnFichero As String, AgruparVtos As Boolean) As Boolean
    Dim ValorEnOpcionales As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    
    
    Dim SQL As String
    Dim ImpEf As Currency
    Dim TotalRem As Currency
    '
    Dim IdDeudor As String
    Dim Cuenta As String
    Dim Fecha2 As Date
    Dim FinFecha As Boolean

    
    Dim EsPersonaJuridica As Boolean
    Dim J As Long
    
    
    Dim RepeticionBucle As Byte   'Si lleva agrupacion serán dos veces. 1los normales 2 Los agrupados
    Dim rp As Byte
    Dim CuentasAgrupadas As String
    Dim cLineas As Collection
    
    On Error GoTo Err_Remesa19sepa
    
    

   
    
    
    
    'J = numerototalregistro
    'ImpEfe = total remesa
    
    RepeticionBucle = 1
    If AgruparVtos Then
    
    
        CuentasAgrupadas = ""
    
        SQL = "select codmacta,count(*) from cobros where "
        SQL = SQL & " codrem = " & RecuperaValor(Remesa_, 1)
        SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa_, 2)
        SQL = SQL & " group by codmacta having count(*) >1"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            
            CuentasAgrupadas = CuentasAgrupadas & ", '" & miRsAux!codmacta & "'"
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
    
        If CuentasAgrupadas = "" Then Err.Raise 513, , "NO hay vencimientos para agrupar"
        CuentasAgrupadas = Mid(CuentasAgrupadas, 2)
        RepeticionBucle = 2
    End If
    
    
    
    Conn.Execute "DELETE FROM tmpcobros2 where codusu =" & vUsu.Codigo
    
    For rp = 1 To RepeticionBucle
    
    
        
        
        SQL = "insert into tmpcobros2(codusu,numserie,numfactu,fecfactu,numorden,codmacta,codrem,anyorem,fecvenci,impvenci,"
        SQL = SQL & " text33csb,text41csb,gastos,iban,nomclien,nifclien,domclien,cpclien,pobclien,proclien,codpais,referencia) "
        SQL = SQL & " SELECT " & vUsu.Codigo & ","
        If rp = 1 Then
            SQL = SQL & "numserie,numfactu,fecfactu,numorden,cobros.codmacta,codrem,anyorem,"
            If FechaCobro = "" Then
                SQL = SQL & " fecvenci"
            Else
                SQL = SQL & "'" & Format(FechaCobro, FormatoFecha) & "'"
            End If
            SQL = SQL & " ,impvenci,text33csb,text41csb,cobros.gastos,"
        Else
            SQL = SQL & " 'GRP' numserie,substring(codmacta,4),max(fecfactu),1,cobros.codmacta,codrem,anyorem,"
            If FechaCobro = "" Then
                SQL = SQL & " max(fecvenci)"
            Else
                SQL = SQL & "'" & Format(FechaCobro, FormatoFecha) & "'"
            End If
            SQL = SQL & " as fecvenci ,sum(impvenci),GROUP_CONCAT( concat(numserie,numfactu) separator ' ') ,"
            SQL = SQL & " concat('Numero Vencimientos : ' , count(*)),sum(coalesce(cobros.gastos,0)),"
            
        
        End If
        
        
        SQL = SQL & " cobros.iban,cobros.nomclien,cobros.nifclien,cobros.domclien,cobros.cpclien,cobros.pobclien,"
        SQL = SQL & " cobros.proclien,cobros.codpais,cobros.referencia from cobros WHERE "
        SQL = SQL & " codrem = " & RecuperaValor(Remesa_, 1)
        SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa_, 2)

        If AgruparVtos Then
            SQL = SQL & " AND "
            If rp = 1 Then SQL = SQL & " NOT "
            SQL = SQL & " cobros.codmacta IN (" & CuentasAgrupadas & ")"
            
        End If
            
        If rp = 2 Then SQL = SQL & " group by cobros.codmacta "

        Conn.Execute SQL
    
    Next rp
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
        SQL = "select  numserie,numfactu,fecfactu,numorden,tmpcobros2.codmacta,codrem,anyorem,"
        SQL = SQL & " fecvenci,impvenci,text33csb,text41csb,tmpcobros2.gastos,tmpcobros2.iban,"
        SQL = SQL & "tmpcobros2.nomclien,tmpcobros2.nifclien,tmpcobros2.domclien,tmpcobros2.cpclien,tmpcobros2.pobclien,"
        SQL = SQL & " tmpcobros2.proclien,tmpcobros2.codpais,bics.bic,"
        SQL = SQL & "tmpcobros2.referencia,cuentas.SEPA_Refere,cuentas.SEPA_FecFirma  from tmpcobros2"
        SQL = SQL & "  left join bics on mid(tmpcobros2.iban,5,4)=bics.entidad inner join cuentas on "
        SQL = SQL & " tmpcobros2.codmacta = cuentas.codmacta WHERE "
        'SQL = SQL & " codrem = " & RecuperaValor(Remesa_, 1)
        'SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa_, 2)
        SQL = SQL & " codusu = " & vUsu.Codigo
        'sepa
        SQL = SQL & " order by  fecvenci,nifdatos,tmpcobros2.codmacta"
        
        
        
        
        
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
        
        
        
        
        
        Set cLineas = New Collection
        
        
        J = 0
        TotalRem = 0
        SQL = ""
        If Not miRsAux.EOF Then
            
                
                Fecha2 = "01/01/1900"
                FinFecha = False
                While Not miRsAux.EOF
                
                    'Informacion del PAGO.
                    ' Se imprime una vez cada FECHA
                    If Fecha2 <> miRsAux!FecVenci Then
                            
                            
                            
                            
                            
                            
                            If Fecha2 > CDate("01/02/1900") Then cLineas.Add "</PmtInf>"
                            Fecha2 = miRsAux!FecVenci
                            
                            
                            'Previo envio vtos
                           cLineas.Add "<PmtInf>"
    
                            'SQL = "RE" & miRsAux!Tiporem & Format(miRsAux!CodRem, "000000") & Format(miRsAux!AnyoRem, "0000") & " " & Format(Fecha2, "dd/mm/yyyy")
                            SQL = "RE" & Format(miRsAux!CodRem, "00000") & Format(miRsAux!AnyoRem, "0000") & " " & Format(FecPre, "dd/mm/yy") & NifEmpresa
                            
                            cLineas.Add "   <PmtInfId>" & SQL & "</PmtInfId>"
                            cLineas.Add "   <PmtMtd>DD</PmtMtd>"             'DirectDebit
                            cLineas.Add "   <BtchBookg>false</BtchBookg>"    'True: un apunte por cada recib   False: Por el total
                            cLineas.Add "   <PmtTpInf>"
                            cLineas.Add "      <SvcLvl>"
                            cLineas.Add "          <Cd>SEPA</Cd>"
                            cLineas.Add "      </SvcLvl>"
                            cLineas.Add "      <LclInstrm>"
                            cLineas.Add "         <Cd>CORE</Cd>"   'CORE o COR1(YA NO VA EL COR1)
                            cLineas.Add "      </LclInstrm>"
                            cLineas.Add "      <SeqTp>RCUR</SeqTp>"
                            cLineas.Add "      <CtgyPurp>"
                            cLineas.Add "         <Cd>TRAD</Cd>"
                            cLineas.Add "      </CtgyPurp>"
                            cLineas.Add "   </PmtTpInf>"
                            'cLineas.Add "   <ReqdColltnDt>" & Format(FecCobro, "yyyy-mm-dd") & "</ReqdColltnDt>"
                            cLineas.Add "   <ReqdColltnDt>" & Format(Fecha2, "yyyy-mm-dd") & "</ReqdColltnDt>"
                            cLineas.Add "   <Cdtr>"
                            cLineas.Add "      <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
                            cLineas.Add "      <PstlAdr>"
                            cLineas.Add "          <Ctry>ES</Ctry>"
                            
                            Dim RsDirec As ADODB.Recordset
                            Dim SqlDirec As String
                            Dim Direccion As String
                            
                            Direccion = ""
                            
                            SqlDirec = "select direccion, numero, escalera, piso, puerta from empresa2"
                            Set RsDirec = New ADODB.Recordset
                            RsDirec.Open SqlDirec, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If Not RsDirec.EOF Then
                                Direccion = DBLet(RsDirec!Direccion) & " " & DBLet(RsDirec!numero) & " " & DBLet(RsDirec!escalera) & " " & DBLet(RsDirec!piso) & " " & DBLet(RsDirec!puerta)
                            End If
                            Set RsDirec = Nothing
                            
                            SQL = Direccion
                            If SQL <> "" Then cLineas.Add "          <AdrLine>" & XML(SQL) & "</AdrLine>"
                            cLineas.Add "      </PstlAdr>"
                            cLineas.Add "   </Cdtr>"
                            cLineas.Add "   <CdtrAcct>"
                            cLineas.Add "      <Id>"
                            'IBAN
    
                            cLineas.Add "         <IBAN>" & DatosBanco & "</IBAN>"
                            cLineas.Add "      </Id>"
                            cLineas.Add "   </CdtrAcct>"
                            cLineas.Add "   <CdtrAgt>"
                            cLineas.Add "      <FinInstnId>"
                            SQL = Mid(DatosBanco, 5, 4)
                            SQL = DevuelveDesdeBD("bic", "bics", "entidad", SQL)
                            cLineas.Add "         <BIC>" & Trim(SQL) & "</BIC>"
                            cLineas.Add "      </FinInstnId>"
                            cLineas.Add "   </CdtrAgt>"
                            
                            cLineas.Add "   <CdtrSchmeId>"
                            cLineas.Add "       <Id>"
                            cLineas.Add "          <PrvtId>"
                            cLineas.Add "             <Othr>"
                            
                            SQL = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
                            SQL = CadenaTextoMod97(SQL)
                            'Si no es dos digitos es un mensaje de error
                            If Len(SQL) <> 2 Then Err.Raise 513, , SQL
                            SQL = "ES" & SQL & Sufijo & NifEmpresa
                            cLineas.Add "                 <Id>" & SQL & "</Id>"
                            cLineas.Add "                 <SchmeNm><Prtry>SEPA</Prtry></SchmeNm>"
                            cLineas.Add "             </Othr>"
                            cLineas.Add "          </PrvtId>"
                            cLineas.Add "       </Id>"
                            cLineas.Add "   </CdtrSchmeId>"
                    End If 'de cambio de fecha
                    
                
                
                
                
                    'Tipo identificador deudor.  Persona fisica (2) o juridica (1)
                    SQL = Mid(miRsAux!nifclien, 1, 1)
                    EsPersonaJuridica = Not IsNumeric(SQL)
                    
                    
                    
                
                
                    cLineas.Add "   <DrctDbtTxInf>"
                    cLineas.Add "      <PmtId>"
                    
                    'Referencia del adeudo
                    SQL = FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!NumFactu, "00000000")
                    SQL = SQL & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
                    SQL = FrmtStr(SQL, 35)
                    cLineas.Add "          <EndToEndId>" & SQL & "</EndToEndId>"
                    cLineas.Add "      </PmtId>"
                    
                    
                    ImpEf = DBLet(miRsAux!Gastos, "N")
                    ImpEf = miRsAux!ImpVenci + ImpEf
                    TotalRem = TotalRem + ImpEf
                    J = J + 1
                    SQL = TransformaComasPuntos(Format(ImpEf, "####0.00"))
                    cLineas.Add "      <InstdAmt Ccy=""EUR"">" & SQL & "</InstdAmt>"
                    cLineas.Add "      <DrctDbtTx>"
                    cLineas.Add "         <MndtRltdInf>"
                    
                    'Si la cuenta tiene ORDEN de mandato, coge este
                    SQL = DBLet(miRsAux!SEPA_Refere, "T")
                    If SQL = "" Then
                        Select Case TipoReferenciaCliente
                        Case 0
                            'Marzo 2017   Si es IBAN, es tooodo el iban
                            SQL = miRsAux!IBAN
                            
                        Case 1
                            'NIF
                            SQL = DBLet(miRsAux!nifclien, "T")
                            
                        Case 2
                            'Referencia en el VTO. No es Nula
                            SQL = DBLet(miRsAux!Referencia, "T")
                            
                        End Select
                    End If
                    cLineas.Add "            <MndtId>" & SQL & "</MndtId>"   'Orden de mandato
                    
                    'Si tiene fecha firma de mandato
                    SQL = "2009-10-31"
                    If Not IsNull(miRsAux!SEPA_FecFirma) Then SQL = Format(miRsAux!SEPA_FecFirma, "yyyy-mm-dd")
                    cLineas.Add "            <DtOfSgntr>" & SQL & "</DtOfSgntr>"
                    
                    cLineas.Add "         </MndtRltdInf>"
                    cLineas.Add "      </DrctDbtTx>"
                    cLineas.Add "      <DbtrAgt>"
                    cLineas.Add "         <FinInstnId>"
                    SQL = FrmtStr(DBLet(miRsAux!BIC, "T"), 11)
                    cLineas.Add "            <BIC>" & SQL & "</BIC>"
                    cLineas.Add "         </FinInstnId>"
                    cLineas.Add "      </DbtrAgt>"
                    cLineas.Add "      <Dbtr>"
                    
                    cLineas.Add "         <Nm>" & XML(miRsAux!nomclien) & "</Nm>"
                    cLineas.Add "         <PstlAdr>"
                    
                    SQL = "ES"
                    If Not IsNull(miRsAux!codpais) Then SQL = Mid(miRsAux!codpais, 1, 2)
                    cLineas.Add "            <Ctry>" & SQL & "</Ctry>"
                    
                    
                    If Not IsNull(miRsAux!domclien) Then cLineas.Add "              <AdrLine>" & XML(miRsAux!domclien) & "</AdrLine>"
                    
                    SQL = ""
                    'SQL = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
                    'If SQL <> "" Then cLineas.Add "              <AdrLine>" & SQL & "</AdrLine>"If Not IsNull(miRsAux!desprovi) Then cLineas.Add "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
                    If DBLet(miRsAux!pobclien, "T") = DBLet(miRsAux!proclien, "N") Then
                        SQL = Trim(DBLet(miRsAux!cpclien, "T") & "   " & DBLet(miRsAux!pobclien, "T"))
                    
                    Else
                        SQL = Trim(DBLet(miRsAux!pobclien, "T") & "   " & DBLet(miRsAux!cpclien, "T"))
                        If Not IsNull(miRsAux!proclien) Then SQL = SQL & "     " & miRsAux!proclien
                    End If
                    If SQL <> "" Then cLineas.Add "              <AdrLine>" & XML(Mid(SQL, 1, 70)) & "</AdrLine>"
                    
                    
                    
                    cLineas.Add "         </PstlAdr>"
                    cLineas.Add "         <Id>"
                    cLineas.Add "            <PrvtId>"
                    cLineas.Add "               <Othr>"
                    
                    
                    'Opcion nueva: 3   Quiere el campo referencia de cobros
    '??             SQL = DBLet(miRsAux!SEPA_Refere, "T")
    '??             If SQL = "" Then
                       Select Case TipoReferenciaCliente
                       Case 0
                           'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                           SQL = Mid(DBLet(miRsAux!IBAN), 13, 2) ' Dígitos de control
                           SQL = SQL & Mid(DBLet(miRsAux!IBAN), 15, 10) ' Código de cuenta
                       Case 1
                           'NIF
                           SQL = DBLet(miRsAux!nifclien, "T")
                    
                       Case 2
                           'Referencia en el VTO. No es Nula
                           SQL = DBLet(miRsAux!Referencia, "T")
                       
                       End Select
    '??             End If
                    
                    cLineas.Add "                   <Id>" & SQL & "</Id>"
                    If TipoReferenciaCliente = 1 Then cLineas.Add "                   <Issr>NIF</Issr>"
                    cLineas.Add "               </Othr>"
                    cLineas.Add "            </PrvtId>"
                    cLineas.Add "         </Id>"
                    cLineas.Add "      </Dbtr>"
                    cLineas.Add "      <DbtrAcct>"
                    cLineas.Add "         <Id>"
                    
                    SQL = IBAN_Destino   'Hay que poner TRUE aunque sea cobro
                    cLineas.Add "            <IBAN>" & SQL & "</IBAN>"
                    cLineas.Add "         </Id>"
                    cLineas.Add "      </DbtrAcct>"
                    cLineas.Add "      <Purp>"
                    cLineas.Add "         <Cd>TRAD</Cd>"
                    cLineas.Add "      </Purp>"
                    cLineas.Add "      <RmtInf>"
                    
                    SQL = Trim(DBLet(miRsAux!text33csb, "T") & " " & FrmtStr(DBLet(miRsAux!text41csb, "T"), 60))
                    If SQL = "" Then SQL = miRsAux!nomclien
                    cLineas.Add "         <Ustrd>" & XML(SQL) & "</Ustrd>"
                    cLineas.Add "      </RmtInf>"
                    cLineas.Add "   </DrctDbtTxInf>"
            
                
                
                'Siguiente
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            
                  
                
        End If  'De EOF
    
   



    '-- Abrir el fichero a enviar
    NFic = FreeFile()
    Open NomFichero For Output As #NFic
    'El encabezado del fichero
            'Encabezado
    Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.008.001.02"">"
    Print #NFic, "<CstmrDrctDbtInitn>"
                
    Print #NFic, "<GrpHdr>"
    
    If esAnticipoCredito Then
        SQL = "FSDD"
    Else
        SQL = "PRE"
    End If
    
    SQL = SQL & Format(Now, "yyyymmddhhnnss")
    
    'Los milisegundos
    SQL = SQL & Format((Timer - Int(Timer)) * 10000, "0000") & "0"
    'Idententificacion propia
    '   tiporem,codrem,anyorem
    'SQL = SQL & "RE" & miRsAux!Tiporem & Format(miRsAux!CodRem, "000000") & Format(miRsAux!AnyoRem, "0000") Antes En18    3l 3 es SEPA
    SQL = SQL & "RE" & "3" & Format(RecuperaValor(Remesa_, 1), "000000") & Format(RecuperaValor(Remesa_, 2), "0000")
            
    IdGrabadoEnFichero = SQL 'Es lo grabare en la remesa
    Print #NFic, "<MsgId>" & SQL & "</MsgId>"
    
    SQL = Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss")   '<CreDtTm>2015-09-10T16:26:56</CreDtTm>
    Print #NFic, "   <CreDtTm>" & SQL & "</CreDtTm>"
    
    'Control sumatorio y numero de registro
    'LO hemos calculado arriba
    'Lo tenemos en impefec y j
    Print #NFic, "   <NbOfTxs>" & J & "</NbOfTxs>"
    SQL = Format(TotalRem, "###0.00")
    SQL = Replace(SQL, ",", ".")
    Print #NFic, "   <CtrlSum>" & SQL & "</CtrlSum>"
    
    
    'Empezamos datos
    Print #NFic, "   <InitgPty>"
    Print #NFic, "     <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    Print #NFic, "     <Id>"
                
    'Tipo identificador deudor.  Persona fisica (2) o juridica (1)
    SQL = Mid(NifEmpresa, 1, 1)
    EsPersonaJuridica = Not IsNumeric(SQL)
    If EsPersonaJuridica Then
        Print #NFic, "        <OrgId>"
    Else
        Print #NFic, "        <PrvtId>"
    End If
    
    SQL = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
    SQL = CadenaTextoMod97(SQL)
    'Si no es dos digitos es un mensaje de error
    If Len(SQL) <> 2 Then Err.Raise 513, , SQL
    SQL = "ES" & SQL & Sufijo & NifEmpresa
    Print #NFic, "           <Othr>"
    Print #NFic, "              <Id>" & SQL & "</Id>"   'Ejemplo: ES3100024348588Y
    Print #NFic, "           </Othr>"
    
    If EsPersonaJuridica Then
        Print #NFic, "        </OrgId>"
    Else
        Print #NFic, "        </PrvtId>"
    End If
    
    
    Print #NFic, "      </Id>"
    Print #NFic, "   </InitgPty>"
    Print #NFic, "</GrpHdr>"


    For J = 1 To cLineas.Count
        Print #NFic, cLineas.Item(J)
    Next


          
    Print #NFic, "</PmtInf>"
    Print #NFic, "</CstmrDrctDbtInitn></Document>"



    Close #NFic
    
    GrabarDisketteNorma19_SEPA_XML = True
Err_Remesa19sepa:
    If Err.Number <> 0 Then
        MsgBox "Err: " & Err.Number & vbCrLf & _
            Err.Description, vbCritical, "Grabación del diskette de Remesa SEPA"
    End If
    Ejecuta "DELETE FROM tmpcobros2 where codusu =" & vUsu.Codigo, True
    CerrarFichero NFic
End Function


Private Function DevolucionTipoPopular(TextoFichero As String, NIF As String) As Boolean
Dim Cad As String
Dim N As Integer
On Error GoTo eDevolucionTipoPopular
    DevolucionTipoPopular = False
    N = InStr(1, TextoFichero, "<GrpHdr>")
    If N > 0 Then
        N = InStr(N + 5, TextoFichero, "</GrpHdr>")
        If N > 0 Then
            Cad = Mid(TextoFichero, 1, N)
            N = InStr(1, Cad, "<InitgPty>")
            If N > 0 Then
                Cad = Mid(Cad, N + 3)
                'Ejemplo: <Id>ES74000B98734098</Id>
                N = InStr(1, Cad, "<Id>ES")
                Cad = Mid(Cad, N + 11)
                N = InStr(1, Cad, "</Id>")
                If N > 0 Then
                    NIF = Mid(Cad, 1, N - 1)
                    DevolucionTipoPopular = True
                End If
                
            End If
        End If
    End If
    Exit Function
eDevolucionTipoPopular:
    MuestraError Err.Number, Err.Description
End Function
