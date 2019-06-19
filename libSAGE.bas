Attribute VB_Name = "libSAGE"
Option Explicit

Dim NumAsien As Long
Dim NumDeFactu As Long
Dim SerieDeFactu  As String
Dim Nivel As Integer


Dim CadenaInsert As String

Public Function ProcesaFicheroClientesSAGE(Fichero As String, ByRef Lb As Label, PirmeraLineaEncabezados As Boolean, Escliente As Boolean) As Byte
Dim NF As Integer
Dim Ok As Boolean
Dim Linea As String
Dim Seguir As Boolean
Dim RA As ADODB.Recordset
Dim PrimLinea As Boolean


On Error GoTo eProcesaFicheroClientes

    ProcesaFicheroClientesSAGE = 2 'NADA no procesa nada


    Lb.Caption = "Leyendo csv"
    Lb.Refresh
    
    'Preparamos tabla de insercion para ver cuantas facturas o si hay errores...
    Conn.Execute "DELETE FROM tmpintefrafracli WHERE codusu = " & vUsu.Codigo
    
    'Apuntes
    Conn.Execute "DELETE FROM tmpintegrapu WHERE codusu = " & vUsu.Codigo
        
    'Llevara los insertrs a ejecutar
    Conn.Execute "DELETE FROM tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
        
    'Las cuentas del fichero
    Conn.Execute "DELETE FROM tmpcuentas WHERE codusu = " & vUsu.Codigo
    
    
    
    CadenaInsert = ""
    NF = FreeFile
    Ok = False
    Open Fichero For Input As #NF
    Seguir = Not EOF(NF)
    NumAsien = -1
    PrimLinea = True
    Msg = ""
    While Seguir
        Line Input #NF, Linea
         
        
        If PrimLinea Then
            J = InStr(1, Linea, ";")
            If J > 0 Then
                SerieDeFactu = Trim(Mid(Linea, 1, J - 1))
                If PirmeraLineaEncabezados Then
                   If IsNumeric(SerieDeFactu) Then Msg = "La primera linea NO parece ser de encabezados. "
                Else
                    If Not IsNumeric(SerieDeFactu) Then Msg = "La primera linea parece ser de encabezados. "
                End If
            End If
            
            If Msg <> "" Then
                Ok = False
                Msg = Msg & vbCrLf & vbCrLf & Mid(Linea, 1, 50) & "..."
                Msg = Msg & vbCrLf & vbCrLf & "¿Continuar?"
                If MsgBox(Msg, vbQuestion + vbYesNoCancel) = vbYes Then Msg = ""
            Else
                If PirmeraLineaEncabezados Then
                    Msg = "N" 'para no procesar la linea
                    Ok = True
                End If
            End If
            SerieDeFactu = ""
            PrimLinea = False
        End If
        If Msg = "" Then Ok = ProcesarLineaAsiento(Linea, Escliente)
        If Not Ok Then
            Seguir = False
        Else
            Seguir = Not EOF(NF)
            Msg = ""
        End If
    Wend
    Close (NF)

    If Ok Then
        Lb.Caption = "Fichero de cuentas"
        Lb.Refresh
        
        'Procesamos el fichero de cuentas
        J = InStrRev(Fichero, "\")
        If J = 0 Then
            MsgBox "Imposible localizar fichero cuentas. Falta \.", vbCritical
        Else
            Msg = Mid(Fichero, 1, J)
            Msg = Msg & "XSUBCTA.csv"
        
            If Dir(Msg, vbArchive) = "" Then
                MsgBox "Imposible lozalizar fichero de cuentas: " & Msg, vbExclamation
            
            Else
                NF = FreeFile
                Open Msg For Input As #NF
                Seguir = Not EOF(NF)
                NumAsien = -1
                PrimLinea = True
                Msg = ""
                While Seguir
                    Line Input #NF, Linea
                 
                
                    ProcesarLineaCuentasContables Linea
                    Seguir = Not EOF(NF)
                Wend
                Close (NF)
                
                If Msg <> "" Then
                    Msg = Mid(Msg, 2)
                    Msg = "INSERT IGNORE INTO tmpcuentas(codusu,codmacta,nommacta,nifdatos,razosoci,dirdatos,codposta,despobla,desprovi) VALUES " & Msg
                    Ejecuta Msg
                Else
                    MsgBox "Ningun dato en fichero cuentas contables", vbExclamation
                End If
                    
            End If
        End If
    End If
    
    If Ok Then
        Lb.Caption = "Comprobando apunte"
        Lb.Refresh
        
        If CadenaInsert <> "" Then
            CadenaInsert = Mid(CadenaInsert, 2)
            SerieDeFactu = DevuelevInsertInttmpAputes
            CadenaInsert = SerieDeFactu & CadenaInsert
            Conn.Execute CadenaInsert
        End If
        espera 0.5
        Conn.Execute "update tmpintegrapu set numdocum='' where numdocum is null and codusu=" & vUsu.Codigo
        
        Conn.Execute "update tmpintegrapu set timporteh=0 where timporteh is null and timported is null and codmacta like '477%'  and codusu=" & vUsu.Codigo
        Conn.Execute "update tmpintegrapu set timported=0 where timporteh is null and timported is null  and codusu=" & vUsu.Codigo
                
                
                        
                
                
                
                
                
        Linea = "Select * from tiposiva"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Linea, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
                
        'Si llega a aqui, vamos a generar las facturas
        Set RA = New ADODB.Recordset
        NumRegElim = 0
        
        If Not Escliente Then
            Linea = "FRAPRO"
        Else
            Linea = "FRACLI"
        End If
        
        Linea = "select numasien,fechaent,1 numdiari from tmpintegrapu where idcontab='" & Linea & "' AND codusu=" & vUsu.Codigo & " GROUP by  numasien,fechaent"
        RA.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Seguir = True
        
        While Seguir
            'Es una factura
            Lb.Caption = "Leyendo facturas fich."
            Lb.Refresh
                
    
            
            'Prepararemos las facturas
            If Escliente Then
                Ok = CrearFacturaClientes(RA!NumAsien, RA!FechaEnt, RA!NumDiari)
            Else
                Ok = CrearFacturaProveedores(RA!NumAsien, RA!FechaEnt, RA!NumDiari)
            End If
            If Not Ok Then
                Seguir = False
            Else
                RA.MoveNext
                If RA.EOF Then Seguir = False
            End If
        Wend
        RA.Close
                
        'Si ha ido todo bien, haremos un par de comprobaciones.
        'Cuentas de cliente/proveedor DEBEN existir
        'Cuanteas de los pauntes tambien.
        'Una salvedad, si la de cli/pro NO existe debe buscarla en el fichero adjunto xsubcta.csv
        Lb.Caption = "Comprobando cuentas contables"
        Lb.Refresh
        
        If Ok Then Ok = ComprobarCuentasContables
        
        
        
        'PEqueña comprobacion
        'para el año de la factura, NO existen ya en contabilidad
        Lb.Caption = "Comprobando facturas"
        Lb.Refresh
        If Ok Then
            If ComprobarNumerosDeFactura(Escliente) Then
               ProcesaFicheroClientesSAGE = 0  'TODO BIEN
            Else
                 ProcesaFicheroClientesSAGE = 1  'Duplicados
            End If
        Else
            ProcesaFicheroClientesSAGE = 2
        End If
    Else
        
        
        
        
    End If
    
    
    
eProcesaFicheroClientes:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
       
    End If
    Set miRsAux = Nothing
    Set RA = Nothing
End Function

Private Function DevuelevInsertInttmpAputes() As String
        
        DevuelevInsertInttmpAputes = "INSERT INTO tmpintegrapu(codusu,numdiari,fechaent,numasien,codconce,linliapu,codmacta,ctacontr,ampconce,"
        DevuelevInsertInttmpAputes = DevuelevInsertInttmpAputes & "timporteD,timporteH,codccost,numdocum,idcontab,numfaccl,numserie,baseimpo,numfacpr"
        DevuelevInsertInttmpAputes = DevuelevInsertInttmpAputes & ",fecfactu,nif,nommacta,ticketIni,ticketFin) VALUES "
        
        
End Function

Private Function ProcesarLineaAsiento(Linea As String, Clientes As Boolean) As Boolean
Dim numero As Long
Dim cad As String
Dim F As Date
Dim Aux As String
Dim aux2 As String
Dim IVA As Currency
Dim Cta As String
Dim Orden As Integer 'Por si hay mas de una linea de 477. Para saber el importe
Dim Baseimpo As String
Dim EsFactura As Byte '0 NO 1 cli 2 pro
Dim strArray() As String


    On Error GoTo EProcesarLineaAsientO
    
    ProcesarLineaAsiento = False
    
    Linea = Replace(Linea, """", "")
    strArray = Split(Linea, ";")
            
            
    If UBound(strArray) < 100 Then
        Aux = "Campos en fichero: " & UBound(strArray) & "       Campos para procesar: 100"
        MsgBox Aux, vbExclamation
        Exit Function
    End If
    
    
            
    'Numasiento
    aux2 = strArray(0)
    numero = Val(aux2)
   ' If numero = 215 Then Stop
    'If numero = 216 Then Stop
    
    
        
        
        aux2 = Trim(strArray(1))
        If InStr(1, aux2, "/") > 0 Then
            
            J = InStr(1, aux2, " ")
            If J > 0 Then aux2 = Trim(Mid(aux2, 1, J))
            
            F = CDate(aux2)
            
        Else
            F = CDate(Mid(aux2, 7, 2) & "/" & Mid(aux2, 5, 2) & "/" & Mid(aux2, 1, 4))
        End If
            
    If NumAsien <> numero Then Nivel = 1
        
    NumAsien = numero
        
   


    'tmpintegrapu(codusu,numdiari,fechaent,numasien,codconce,linliapu,codmacta,ctacontr,ampconce,timporteD,timporteH,codccost,
    'numdocum,idcontab,numfaccl,numserie,baseimpo,numfacpr,fecfactu,nif,nommacta)
    cad = IIf(Clientes, vParam.conceacl, vParam.conceapr)
    cad = "(" & vUsu.Codigo & ",1,'" & Format(F, "yyyy-mm-dd") & "'," & NumAsien & "," & cad & "," & Nivel
    
    'Cta contable
    Aux = Trim(strArray(2))
    cad = cad & ",'" & Aux & "'"
    Cta = Aux
    
    'Contrpartida si tiene
    Aux = Trim(strArray(3))
    If Aux <> "" Then
        If Not IsNumeric(Aux) Then Aux = ""
    End If
    If Aux = "" Then
        aux2 = "NULL"
    Else
        aux2 = "'" & Aux & "'"
    End If
    cad = cad & "," & aux2

    'Ampliacion concepto
    aux2 = strArray(5)
    Aux = Trim(DevNombreSQL(aux2))
    cad = cad & ",'" & Aux & "'"

    'importe Debe
    aux2 = Trim(strArray(27))
    aux2 = Replace(aux2, ",", ".")
    If aux2 = "" Then aux2 = "0.00"
    If aux2 = "0.00" Then
        Aux = "NULL"
    Else
        Aux = Trim(aux2)
    End If
    cad = cad & "," & Aux


    'importe Debe
     aux2 = Trim(strArray(28))
    aux2 = Replace(aux2, ",", ".")
    If aux2 = "" Then aux2 = "0.00"
    If aux2 = "0.00" Then
        Aux = "NULL"
    Else
        Aux = aux2
    End If
    cad = cad & "," & Aux


    'IVA
    '------------------------------
    IVA = 0
    aux2 = Trim(strArray(9))
    aux2 = Replace(aux2, ",", ".")
    If aux2 = "" Then aux2 = "0.00"
    IVA = CCur(Trim(aux2))
    
    
    
        
    'Centro de coste
    Aux = "NULL"
    aux2 = Trim(strArray(12))
    If Trim(aux2) <> "" Then
        Aux = Trim(strArray(13))
        aux2 = aux2 & Aux
        Aux = "'" & aux2 & "'"
    End If

    cad = cad & "," & Aux & ","
    
    'Numdocum
    Aux = "NULL"
    aux2 = Trim(strArray(11))
    If aux2 <> "" Then
        If aux2 <> "0" Then Aux = "'" & aux2 & "'"

    End If
    cad = cad & Aux & ","
    
    Dim Raiz As String
    
    'Febreero2019
    NumDeFactu = 0
     aux2 = Trim(strArray(7))
    If Val(Trim(aux2)) > 0 Then NumDeFactu = Val(Trim(aux2))
    SerieDeFactu = ""
    aux2 = Trim(strArray(22))
    If Trim(aux2) <> "" Then SerieDeFactu = aux2
    'If NumDeFactu <> 0 Then Stop
    
    Raiz = Mid(Trim(Trim(strArray(2))), 1, 3)
'    If iva = 0 Then
    If Raiz <> "472" And Raiz <> "477" Then
        cad = cad & "'CONTAB'"
        Baseimpo = "null"
        EsFactura = 0
    Else
        If Clientes Then
            cad = cad & "'FRACLI'"
        Else
            cad = cad & "'FRAPRO'"
        End If
        'Ahora creamos la linea para la insercion de la base imponible
        aux2 = Trim((strArray(29)))
        aux2 = Replace(aux2, ",", ".")
        If aux2 = "" Then aux2 = "0.00"
        
        Baseimpo = aux2
        
        
        
        
        If Raiz = "477" Then
            EsFactura = 1
            If SerieDeFactu = "" Then SerieDeFactu = "P"  'NO ha escificado numero serie
        
        Else
            'proveedor
            EsFactura = 2
            SerieDeFactu = "1"
        End If
        
        
    End If
    
    If EsFactura > 0 Then
         cad = cad & "," & NumDeFactu & ",'" & SerieDeFactu & "'"
    Else
        cad = cad & ",null,null"
    End If
    
    cad = cad & "," & Baseimpo
    
    
    'numfacpr   num factura provee
    If EsFactura = 2 Then
        
        aux2 = Trim((strArray(71)))
        aux2 = Right(aux2, 10)
        If aux2 = "" Then Err.Raise 513, , "Numero factura no encontrado (Columna BT)"
        cad = cad & "," & DBSet(aux2, "T")
        
        aux2 = Trim((strArray(1)))   'FALTA### Ver donde esta la fecha
        aux2 = Right(aux2, 8)
        If aux2 <> "" Then
            F = CDate(Mid(aux2, 7, 2) & "/" & Mid(aux2, 5, 2) & "/" & Mid(aux2, 1, 4))
            cad = cad & "," & DBSet(F, "F")
        Else
            cad = cad & ",null"
        End If
    Else
        cad = cad & ",null,null"
    End If
    
    
    'coddevol bancotalonpag
    '  nifdatos nommacta
    If EsFactura = 0 Then
        cad = cad & ",null,null,null,null"
        
    Else
   
        aux2 = Trim((strArray(61)))
        cad = cad & "," & DBSet(aux2, "T", "N")

        aux2 = Trim((strArray(62)))
        cad = cad & "," & DBSet(aux2, "T", "N")
        
        aux2 = Trim((strArray(120)))
        If aux2 = "" Then
            cad = cad & ",null,null"
        Else
            'Ticket inicial y FINAL
        
            cad = cad & ",'" & aux2
            aux2 = Trim((strArray(121)))
            If Trim(aux2) = "" Then Err.Raise 513, , "NO esta indicado el ticket final: " & SerieDeFactu & NumDeFactu
            cad = cad & "','" & aux2 & "'"
        End If
    End If

    cad = cad & ")"


    CadenaInsert = CadenaInsert & ", " & cad


    Nivel = Nivel + 1

    If Len(CadenaInsert) > 1200 Then
        CadenaInsert = Mid(CadenaInsert, 2)
        cad = DevuelevInsertInttmpAputes
        cad = cad & CadenaInsert
        Conn.Execute cad
        CadenaInsert = ""
    End If

    
    ProcesarLineaAsiento = True
    Exit Function
EProcesarLineaAsientO:
    MuestraError Err.Number, Err.Description
End Function












Private Function ProcesarLineaCuenta(Linea As String) As Boolean
Dim cad As String
Dim Aux As String
Dim Nommacta As String
Dim codmacta As String
Dim aux2 As String
Dim T1 As Boolean
Dim SQL As String

    On Error GoTo EProcesarLinea

    ProcesarLineaCuenta = False
    cad = ""

    'Codmacta
    Aux = Mid(Linea, 1, 12)
    codmacta = Trim(Aux)
    cad = cad & "'" & Trim(Aux) & "'"

    If Nivel < 0 Then
        'Es la primera vez
        Nivel = NivelCuenta(codmacta)
    End If
    

    'Nommacta
    aux2 = Mid(Linea, 13, 40)
    Aux = Trim(DevNombreSQL(aux2))
    Nommacta = Aux
    Aux = "'" & Aux & "'"
    cad = cad & "," & Aux
    

    'NIF
    aux2 = Mid(Linea, 53, 15)
    Aux = Trim(DevNombreSQL(aux2))
    If Aux = "" Then
        Aux = "NULL"
        T1 = False
    Else
        T1 = True
        Aux = "'" & Aux & "'"
    End If
    cad = cad & "," & Aux
    
    
    'Direccion
    aux2 = Mid(Linea, 68, 35)
    aux2 = Mid(aux2, 1, 30)
    Aux = Trim(DevNombreSQL(aux2))
    If Aux = "" Then
        Aux = "NULL"
    Else
        Aux = "'" & Aux & "'"
    End If
    cad = cad & "," & Aux
    
    
    'Poblacion
    aux2 = Mid(Linea, 103, 25)
    Aux = Trim(DevNombreSQL(aux2))
    If Aux = "" Then
        Aux = "NULL"
    Else
        Aux = "'" & Aux & "'"
    End If
    cad = cad & "," & Aux
    
    
    'Provincia
    aux2 = Mid(Linea, 128, 20)
    Aux = Trim(DevNombreSQL(aux2))
    If Aux = "" Then
        Aux = "NULL"
    Else
        Aux = "'" & Aux & "'"
    End If
    cad = cad & "," & Aux


    'Cod pos
    aux2 = Mid(Linea, 148, 5)
    Aux = Trim(DevNombreSQL(aux2))
    If Aux = "" Then
        Aux = "NULL"
    Else
        Aux = "'" & Aux & "'"
    End If
    cad = cad & "," & Aux
    

    'Si tiene NIF ponemos 347 a 1
    'y razosoci le ponemos la misma que el cliente
    If T1 Then
        cad = cad & ",1,'" & Nommacta & "','ES'"
    Else
            cad = cad & ",0,NULL,NULL"
    End If
    
    'maidatos,webdatos
    cad = cad & ",NULL,NULL)"
    
    

    'Montamos el SQL
    cad = SQL & cad
    Conn.Execute cad


    ProcesarLineaCuenta = True
    Exit Function
EProcesarLinea:
    Aux = "Error procesando linea: " & vbCrLf & Linea & vbCrLf & vbCrLf
    Aux = Aux & Err.Description
    MsgBox Aux, vbExclamation
End Function






Private Function CrearFacturaClientes(NA As Long, FechaEnt As Date, NumDiari As Integer) As Boolean
Dim i  As Integer
Dim cad As String
Dim SQL As String

Dim FinBucle As Boolean
Dim HayQueInsertarFactura As Boolean
Dim ReestableceValores As Boolean
Dim N As Long

Dim NumLiena2 As Integer
Dim InsTotales As String

Dim NumLienaB As Integer
Dim InsBases As String

Dim INsFactura As String
Dim InsrListadoFacturasFichero As String


Dim TotalFac As Currency
Dim Tbases As Currency
Dim TotIVA As Currency
Dim Suplidos As Currency
Dim TotalAprox As Currency

Dim Serie As String
Dim PrimeraContrapartida As Boolean
Dim ImportAuxiliar As Currency
Dim Rs As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim NIF As String
Dim Cta As String

Dim CambiaSerieFactura As Boolean


Dim IVA_Suplidos As Integer
Dim CadenaInserPorsiSuplidos As String
Dim RaizRetencion As String
Dim ImporReten As Currency
Dim CtaReten As String



    On Error GoTo Salida
    Set Rs = New ADODB.Recordset
    Set R2 = New ADODB.Recordset

    cad = "select tmpintegrapu.* ,  if(substring(codmacta,1,3)='477',0,1) orden "
    cad = cad & " from tmpintegrapu where numasien=" & NA & " AND numdiari=" & NumDiari & " AND codusu =" & vUsu.Codigo
    cad = cad & " And fechaent='" & Format(FechaEnt, FormatoFecha) & "'  ORDER BY  orden,numserie,linliapu"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = "select codmacta,linliapu, timported,timporteh,numdocum "
    cad = cad & " from tmpintegrapu where numasien=" & NA & " AND numdiari=" & NumDiari & " AND codusu =" & vUsu.Codigo
    cad = cad & " And fechaent='" & Format(FechaEnt, FormatoFecha) & "' AND substring(codmacta,1,3)<>'477'  ORDER BY  codmacta"
    R2.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    
    'RaizRetencion ="473"
    RaizRetencion = "473"

    
    
    'Vamos a fijar el total factura
    NumDeFactu = -1
        
        
        
    IVA_Suplidos = -1
    FinBucle = False
    
    While Not FinBucle
        HayQueInsertarFactura = False
        ReestableceValores = False
        If Rs.EOF Then
            HayQueInsertarFactura = True
        Else
            N = IIf(IsNull(Rs!numfaccl), 0, 1)
            If N > 0 Then
                CambiaSerieFactura = False
                If NumDeFactu <> Rs!numfaccl Then
                    CambiaSerieFactura = True
                Else
                    SQL = DBLet(Rs!NUmSerie, "T")
                    If Serie <> SQL Then CambiaSerieFactura = True
                End If
                If CambiaSerieFactura Then
                    If NumDeFactu > 0 Then HayQueInsertarFactura = True
                    ReestableceValores = True
                End If
            End If
        End If
        
            
        If HayQueInsertarFactura Then
            
            'If ImporReten > 0 Then Stop
            TotalFac = TotalFac - ImporReten
            If TotalAprox <> 0 Then Suplidos = TotalAprox - TotalFac
            If Suplidos <> 0 Then
                
                If IVA_Suplidos < 0 Then
                    'Tenemos que obtener un IVA de suplidos
                    cad = DevuelveDesdeBD("codigiva", "tiposiva", "tipodiva", "4") 'SUPLIDOS=tipodigva 4
                    If cad = "" Then Err.Raise 513, , "No se ha configurado un IVA de suplidos"
                    IVA_Suplidos = Val(cad)
                   
                End If
                    
                'INSERTAMOS LA LINEACad = ", ('" & Rs!NUmSerie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt) & "," & NumLiena2
               
                cad = CadenaInserPorsiSuplidos & "," & NumLiena2
                cad = cad & "," & IVA_Suplidos & "," & DBSet(Suplidos, "N")
                cad = cad & ",0,NULL,0,null)"
                InsTotales = InsTotales & cad
                 NumLiena2 = NumLiena2 + 1
                
                'Stop
                TotalFac = TotalAprox
            End If
            INsFactura = Replace(INsFactura, "#BASES#", TransformaComasPuntos(CStr(Tbases)))
            
            INsFactura = Replace(INsFactura, "#totivas#", TransformaComasPuntos(CStr(TotIVA)))
            INsFactura = Replace(INsFactura, "#totrecargo#", "null")
            INsFactura = Replace(INsFactura, "#totfaccl#", TransformaComasPuntos(CStr(TotalFac)))
            'retencion
            If CtaReten <> "" Then
                INsFactura = Replace(INsFactura, "#retfaccl#", TransformaComasPuntos(CStr(19)))
                INsFactura = Replace(INsFactura, "#trretfac#", TransformaComasPuntos(CStr(ImporReten)))
                INsFactura = Replace(INsFactura, "#cuereten#", DBSet(CtaReten, "T"))
                INsFactura = Replace(INsFactura, "#tiporeten#", "3")
                INsFactura = Replace(INsFactura, "#BASESRET#", TransformaComasPuntos(CStr(Tbases)))
            Else
                
                
                INsFactura = Replace(INsFactura, "#retfaccl#", "null")
                INsFactura = Replace(INsFactura, "#trretfac#", "null")
                INsFactura = Replace(INsFactura, "#cuereten#", "null")
                INsFactura = Replace(INsFactura, "#tiporeten#", "0")
            End If
            INsFactura = Replace(INsFactura, "#BASESRET#", "null")
            If Suplidos = 0 Then
                cad = "null"
            Else
                cad = TransformaComasPuntos(CStr(Suplidos))
            End If
            INsFactura = Replace(INsFactura, "#suplidos#", cad)
            
            'If Serie <> "I" Then InsertaEnTmpInsertrs INsFactura  ######### borrar si esta todo ok (hay unas mas abajo)
            InsertaEnTmpInsertrs INsFactura
            
            If InsTotales <> "" Then
                InsTotales = Mid(InsTotales, 2)
                cad = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,codigiva,baseimpo,porciva,porcrec,impoiva,imporec"
                cad = cad & ") VALUES " & InsTotales
                'If Serie <> "I" Then InsertaEnTmpInsertrs cad
                InsertaEnTmpInsertrs cad
            End If
            
            If InsBases <> "" Then
                InsBases = Mid(InsBases, 2)
                cad = "INSERT INTO factcli_lineas (numserie, numfactu, fecfactu,anofactu, numlinea, codmacta, baseimpo, codccost) VALUES "
                cad = cad & InsBases
                'If Serie <> "I" Then InsertaEnTmpInsertrs cad
                InsertaEnTmpInsertrs cad
            End If
             
             
            'Para la pantalla que indca cuantoas vamos a integrarar y cuales
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#BASES#", TransformaComasPuntos(CStr(Tbases)))
            
              'retencion
            cad = "null"
            If CtaReten <> "" Then cad = TransformaComasPuntos(CStr(ImporReten))
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#BASESRET#", cad)
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#totivas#", TransformaComasPuntos(CStr(TotIVA)))
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#totrecargo#", "null")
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#totfaccl#", TransformaComasPuntos(CStr(TotalFac)))
          

            If Suplidos = 0 Then
                cad = "null"
            Else
                cad = TransformaComasPuntos(CStr(Suplidos))
            End If
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#suplidos#", cad)
            cad = "INSERT INTO tmpintefrafracli (codusu ,codigo ,serie ,factura ,fecha  ,cta_cli ,iban  ,impventa  ,impret ,impiva ,imprecargo,CalculoImponible ,totalfactura ,txtcsb )"
            cad = cad & " VALUES " & Mid(InsrListadoFacturasFichero, 2)
            'If Serie <> "I" Then Conn.Execute cad
            Conn.Execute cad
            
            '-------------
            'Seguimos con los datos, si hay
            If Not Rs.EOF Then
                ReestableceValores = True
            Else
                ReestableceValores = False
                FinBucle = True
            End If
        End If
        If ReestableceValores Then
                
                InsBases = ""
                NumDeFactu = Rs!numfaccl
                InsTotales = ""
                NumLienaB = 1
                TotalFac = 0
                TotIVA = 0
                Suplidos = 0
                ImporReten = 0
                CtaReten = ""
                INsFactura = ""
                Tbases = 0
                Serie = Rs!NUmSerie
                TotalAprox = 0
                NumLiena2 = 0
                
                CadenaInserPorsiSuplidos = ""
                'Vemos si lleva suplidos. Obtenemos el total factura
                If DBLet(Rs!ctacontr, "T") <> "" Then
                    R2.MoveFirst
                    
                    Do
                        R2.Find "codmacta = " & DBSet(Rs!ctacontr, "T"), , adSearchForward
                        If R2.EOF Then
                            NumLiena2 = 1
                            
                        Else
                            If InStr(1, R2!Numdocum, NumDeFactu) > 0 Then
                                TotalAprox = R2!timported
                                NumLiena2 = 1
                            Else
                                
                                R2.MoveNext
                            End If
                        End If
                    Loop Until NumLiena2 = 1
                End If
                
                
                
                PrimeraContrapartida = True
                
                
                'Veremos si lleva RETENCION
                R2.MoveFirst
                NumLiena2 = 0
                Do
                    InsBases = "codmacta like  '" & RaizRetencion & "%'"
                    R2.Find InsBases, , adSearchForward
                    If R2.EOF Then
                        NumLiena2 = 1
                        
                    Else
                        If InStr(1, R2!Numdocum, NumDeFactu) > 0 Then
                            ImporReten = DBLet(R2!timported, "N") + DBLet(R2!timporteH, "N")
                            CtaReten = R2!codmacta
                            NumLiena2 = 1
                        Else
                            
                            R2.MoveNext
                        End If
                    End If
                Loop Until NumLiena2 = 1
                
                
                
                
                NumLiena2 = 1
                InsBases = ""
        End If
        
        
        'Separamos datos del apunte
        If Not FinBucle Then
            If Mid(Rs!codmacta, 1, 3) = "477" Then
                'numserie,numfactu,fecfactu,anofactu,numlinea,codigiva,baseimpo,porciva,porcrec,impoiva,imporec
                
                cad = ", ('" & Rs!NUmSerie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt) & "," & NumLiena2
                i = DevuelveTipoIva(Rs!codmacta, False)
                cad = cad & "," & i & "," & DBSet(Rs!Baseimpo, "N")
                cad = cad & "," & DBSet(miRsAux!porceiva, "N") & ",NULL," & DBSet(Rs!timporteH, "N") & ",null"
                
                TotIVA = TotIVA + Rs!timporteH
                TotalFac = TotalFac + Rs!Baseimpo + Rs!timporteH
                Tbases = Tbases + Rs!Baseimpo
                InsTotales = InsTotales & cad & ")"
                NumLiena2 = NumLiena2 + 1
                
                'Por si hubieran suplidos
                
                CadenaInserPorsiSuplidos = ", ('" & Rs!NUmSerie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt)
                
            
                'La contrpartida es el cliente
                    'CLIENTE. Esta es el numero de factura
                If PrimeraContrapartida Then
                    PrimeraContrapartida = False
                    '", numdiari ,fechaent, numasien, fecliqcl, codconce340,codopera,no_modifica_apunte) VALUES  "
                    
                    
                    If IsNull(Rs!ctacontr) Then
                        Cta = "4300000000" 'Habra que personalizar
                    Else
                        Cta = Rs!ctacontr
                    End If
                    
                    
                    SQL = "Generada por traspaso contaplus"
                    If DBLet(Rs!ticketini, "T") <> "" Then SQL = SQL & "    Tickets: " & DBLet(Rs!ticketini, "T") & " - " & DBLet(Rs!ticketfin, "T")
                    
                    cad = DBSet(Cta, "T") & "," & DBSet(SQL, "T")
                    cad = "'" & Serie & "'," & NumDeFactu & "," & Year(Rs!FechaEnt) & "," & DBSet(Rs!FechaEnt, "F") & "," & cad
                    cad = cad & ",  #BASES# , #BASESRET# ,#totivas# ,#totrecargo# , #totfaccl# "
                    cad = cad & ",  #retfaccl# , #trretfac# ,#cuereten# ,#tiporeten# , #suplidos#"
                    cad = cad & ", " & Rs!NumDiari & "," & DBSet(Rs!FechaEnt, "F") & "," & Rs!NumAsien & "," & DBSet(Rs!FechaEnt, "F")
                    
                    'Concepto 340
                    SQL = "'0'"
                    If DBLet(Rs!ticketini, "T") <> "" Then SQL = "'B'"   'Resumen factura tiket
                    cad = cad & " ," & SQL
                    
                    'codopera,no_modifica_apunte
                    cad = cad & " ,0 ,1"
                    
                    'nommacta, nifdatos   , en apuntes: bancotalonpag coddevol
                    SQL = DBLet(Rs!Nommacta, "T")
                    NIF = ""
                    If SQL = "" Then
                        NIF = "nifdatos"
                        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T", NIF)
                        If SQL = "" Then
                            MsgBox "Nombre cuenta vacia: " & Cta, vbExclamation
                            SQL = "VACIO"
                        End If
                    End If
                    cad = cad & "," & DBSet(SQL, "T") & ","
                    
                    SQL = DBLet(Rs!NIF, "T")
                    If SQL = "" Then SQL = NIF
                    If SQL = "" Then
                        SQL = "null"
                    Else
                        SQL = DBSet(SQL, "T")
                    End If
                    cad = cad & SQL
                    
                    
                    If DBLet(Rs!ticketini, "T") <> "" Then
                        SQL = DBSet(Rs!ticketini, "T") & "," & DBSet(Rs!ticketfin, "T")
                    Else
                        SQL = "NULL,NULL"
                    End If
                    
                    cad = cad & "," & SQL & ")"
                    SQL = FijarCadenaInsercionSQL(True)
                    INsFactura = SQL & cad
                    
                    'para el listado de facturas que vamops a insertar
                    'tmpintefrafracli
                    '                                             nommac   base    ret    iva    recar          suplidos     total        nºAsiento
                    'codusu codigo serie factura fecha  ctaventas iban  impventa  impret impiva imprecargo CalculoImponible totalfactura    txtcsb
                    InsrListadoFacturasFichero = ", (" & vUsu.Codigo & "," & NumRegElim
                    InsrListadoFacturasFichero = InsrListadoFacturasFichero & ",'" & Serie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F")
                    InsrListadoFacturasFichero = InsrListadoFacturasFichero & ",'" & Cta & "',null,"
                    '                                                           base      ret           iva      recar          suplidos     total
                    InsrListadoFacturasFichero = InsrListadoFacturasFichero & " #BASES# , #BASESRET# ,#totivas# ,#totrecargo# , #suplidos# , #totfaccl# , " & NA & ")"
                Else
                    'Stop
                End If
            Else
                'BASES
                'Sql = "INSERT INTO factcli_lineas (numserie, numfactu, anofactu, numlinea, codmacta, baseimpo, codccost) VALUES ("

                
                If Val(Mid(Rs!codmacta, 1, 2)) > 48 Then
                    cad = ", ('" & Serie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt) & "," & NumLienaB
                    If IsNull(Rs!timporteH) Then
                        ImportAuxiliar = -Rs!timported
                    Else
                        ImportAuxiliar = Rs!timporteH
                    End If
                    cad = cad & ",'" & Rs!codmacta & "'," & DBSet(ImportAuxiliar, "N") & ",'" & Rs!codccost & "')"
                    NumLienaB = NumLienaB + 1
                    InsBases = InsBases & cad
                End If
            End If
        End If
        If Not FinBucle Then Rs.MoveNext
    Wend
    
    
    Rs.Close
    CrearFacturaClientes = True 'Si llega aqui ha ido bien
Salida:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Inserta Factura" & vbCrLf & Err.Description
        Err.Clear
    End If
    Set Rs = Nothing
    Set R2 = Nothing
End Function




Private Function DevuelveTipoIva(codmacta As String, Soportado As Boolean) As Integer
Dim Aux As String
    
    DevuelveTipoIva = -1
    
    If Soportado Then
        Aux = "cuentaso"
    Else
        Aux = "cuentare"
    End If
    Aux = Aux & " = '" & codmacta & "'"
    
    miRsAux.Find Aux, , adSearchForward, 1
    
    If Not miRsAux.EOF Then DevuelveTipoIva = miRsAux!codigiva
    
    
    

    
    'ESTO ESTA MAL
    If DevuelveTipoIva < 0 Then Err.Raise 513, , "IVA no encontrado"
   
End Function



'Insertamos en la tabla tmptesoreriacomun . Cad registro llevará el insert into a realizar
Private Sub InsertaEnTmpInsertrs(CADENA As String)
Dim cad As String
    NumRegElim = NumRegElim + 1
    cad = Replace(CADENA, "'", "·")
    cad = "INSERT INTO tmptesoreriacomun(codusu,codigo,Texto) VALUES (" & vUsu.Codigo & "," & NumRegElim & ",'" & cad & "')"
    Conn.Execute cad
End Sub

Private Function FijarCadenaInsercionSQL(Clientes As Boolean) As String
Dim SQL As String

    'Clientes
    If Clientes Then
        SQL = "INSERT INTO factcli (numserie, numfactu, anofactu, fecfactu, codmacta, observa "
        SQL = SQL & ",totbases, totbasesret, totivas, totrecargo, totfaccl "
        SQL = SQL & ",retfaccl, trefaccl, cuereten, tiporeten,suplidos "
        SQL = SQL & ",numdiari,fechaent, numasien, fecliqcl, codconce340,codopera,no_modifica_apunte,nommacta,nifdatos,FraResumenIni ,FraResumenFin) VALUES ( "
        
    Else
    
        'Proveedores
        SQL = "INSERT INTO factpro (numserie,numregis, anofactu ,fecfactu, fecharec, numfactu, codmacta, observa "
        
        SQL = SQL & ",totbases, totbasesret, totivas, totrecargo, totfacpr "
        SQL = SQL & ",retfacpr, trefacpr, cuereten,tiporeten,suplidos "
        SQL = SQL & ",numdiari, fechaent, numasien, fecliqpr, codconce340, estraspasada,no_modifica_apunte,nommacta ,nifdatos) VALUES ("
        
    End If
    FijarCadenaInsercionSQL = SQL
    
    
End Function



Private Function ComprobarCuentasContables() As Boolean
Dim Rs As ADODB.Recordset
Dim C As String
Dim CtasACrear As String
Dim RSubgr As ADODB.Recordset

    ComprobarCuentasContables = False
    'Facil.
    '
    Set RSubgr = New ADODB.Recordset
    Set Rs = New ADODB.Recordset
    CtasACrear = ""
    C = "select  substring(codmacta,1,4) raiz from tmpintegrapu where codusu = " & vUsu.Codigo & "  group by 1"
    RSubgr.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RSubgr.EOF
    
            C = "select distinct codmacta from tmpintegrapu where codusu = " & vUsu.Codigo & " AND codmacta like '" & RSubgr!Raiz & "%'"
            C = C & " and not codmacta IN "
            C = C & " (select codmacta from cuentas where apudirec='S' AND codmacta like '" & RSubgr!Raiz & "%')"
            
            Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs.EOF
                CtasACrear = CtasACrear & ", '" & Rs!codmacta & "'"
                Rs.MoveNext
            Wend
            Rs.Close
         RSubgr.MoveNext
    Wend
    RSubgr.Close
    
    
    C = "DELETE from tmpcuentas where codusu =" & vUsu.Codigo
    If CtasACrear <> "" Then
        CtasACrear = Mid(CtasACrear, 2)
        CtasACrear = "(" & CtasACrear & ")"
        
        'Hay cuentas que no existen. Comprobamos que estan el tmpcuentas (deberia), y las INSERTAMOS
        C = C & " AND not codmacta in " & CtasACrear
        
    End If
    Conn.Execute C
    
    
    
    'Resto de cuentas del APUNTE.
    
    Msg = ""
    C = "select tmpintegrapu.codmacta,cuentas.codmacta CtaEnCuentas from tmpintegrapu left join cuentas on tmpintegrapu.codmacta = cuentas.codmacta where codusu=" & vUsu.Codigo
    If CtasACrear <> "" Then C = C & " AND not tmpintegrapu.codmacta in " & CtasACrear
    Set Rs = New ADODB.Recordset
    Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        If IsNull(Rs!CtaEnCuentas) Then Msg = Msg & "   " & Rs!codmacta
    
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    C = "select tmpintegrapu.ctacontr,cuentas.codmacta CtaEnCuentas from tmpintegrapu left join cuentas on tmpintegrapu.ctacontr = cuentas.codmacta where ctacontr<>'' and codusu =" & vUsu.Codigo
    If CtasACrear <> "" Then C = C & " AND not tmpintegrapu.ctacontr in " & CtasACrear
   
    Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        If IsNull(Rs!CtaEnCuentas) Then Msg = Msg & "   " & Rs!codmacta
    
        Rs.MoveNext
    Wend
    Rs.Close
    

    If Msg <> "" Then
        MsgBox "Existen cuentas en el cichero que no estan en contabilidad" & vbCrLf & Msg, vbExclamation
        
    Else
        ComprobarCuentasContables = True
    End If
End Function



'tmpcuentas(codusu,codmacta,nommacta,nifdatos,razosoci,dirdatos,codposta,despobla,desprovi)
Private Sub ProcesarLineaCuentasContables(Linea As String)
Dim strArray() As String
Dim J As Integer

    On Error GoTo eProcesarLineaCuentasContables
    
      strArray = Split(Linea, ";")
            
      strArray(0) = Replace(strArray(0), """", "")
      If Len(strArray(0)) <> vEmpresa.DigitosUltimoNivel Then Exit Sub
      If Not IsNumeric(strArray(0)) Then Exit Sub
            
      For J = 1 To UBound(strArray)
            strArray(J) = Trim(Replace(strArray(J), """", ""))
            If strArray(J) <> "" Then strArray(J) = Replace(strArray(J), """", "")
      Next
          
        SerieDeFactu = vUsu.Codigo & "," & DBSet(Trim(strArray(0)), "T") & "," & DBSet(Trim(strArray(1)), "T")
        SerieDeFactu = SerieDeFactu & "," & DBSet(Trim(strArray(2)), "T") & "," & DBSet(Trim(strArray(1)), "T") & "," & DBSet(Trim(strArray(3)), "T")
        SerieDeFactu = SerieDeFactu & "," & DBSet(Trim(strArray(6)), "T") & "," & DBSet(Trim(strArray(4)), "T") & "," & DBSet(Trim(strArray(5)), "T") & ")"
        Msg = Msg & ", (" & SerieDeFactu
        
      
eProcesarLineaCuentasContables:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        
    End If
End Sub



Private Function ComprobarNumerosDeFactura(Cliente As Boolean) As Boolean
    
    NumAsien = 0
    ComprobarNumerosDeFactura = False
    SerieDeFactu = DevuelveDesdeBD("min(fecha)", "tmpintefrafracli", "codusu", vUsu.Codigo)
    
    If Cliente Then
        Msg = "select * from tmpintefrafracli where codusu =" & vUsu.Codigo
        Msg = Msg & " AND (serie,factura,year(fecha)) IN ("
        Msg = Msg & " select numserie,numfactu,anofactu from factcli where fecfactu>=" & DBSet(SerieDeFactu, "F") & ")"
    Else
       Msg = "select * from tmpintefrafracli where codusu =" & vUsu.Codigo
        Msg = Msg & " AND (serie,factura,year(fecha)) IN ("
        Msg = Msg & " select numserie,numregis,anofactu from factpro where fecharec>=" & DBSet(SerieDeFactu, "F") & ")"
    End If
    
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SerieDeFactu = ""
    While Not miRsAux.EOF
        
        NumAsien = NumAsien + 1
        
        'select codigo,texto1,texto2,observa1 from tmptesoreriacomun
        Msg = ""
        If Cliente Then Msg = DBLet(miRsAux!Serie, "T")
        Msg = "(" & vUsu.Codigo & "," & NumAsien & "," & DBSet(Msg & Format(miRsAux!FACTURA, "00000"), "T") & ",'" & Format(miRsAux!Fecha, "dd/mm/yyyy") & "','YA existe factura')"
        SerieDeFactu = SerieDeFactu & ", " & Msg
    
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    Msg = "select min(fecha) minima,max(fecha) maxima from tmpintefrafracli where  codusu=" & vUsu.Codigo
    miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'no puede ser eof
    Msg = ""
    If miRsAux!minima < vParam.fechaini Then
        Msg = "Anterior incio ejercicio: " & miRsAux!minima
    Else
        If miRsAux!minima < vParam.FechaActiva Then
            
            Msg = "Menor fecha avtiva: " & miRsAux!minima
        Else
            If miRsAux!minima <= UltimoDiaPeriodoLiquidado Then
                'FACTURAS CLIENTE. Obliado comprobar
                Msg = "Menor que ultimo periodo liquidado: " & miRsAux!minima



            End If
        End If
    End If
    
    If miRsAux!maxima > DateAdd("yyyy", 1, vParam.fechafin) Then Msg = "Mayor fecha ejercicios: " & miRsAux!minima
    miRsAux.Close
    If Msg <> "" Then
        NumAsien = NumAsien + 1
        'codigo,texto1,texto2,observa1 from tmptesoreriacomun
        Msg = "(" & vUsu.Codigo & "," & NumAsien & ",'ERROR',' ','" & Msg & "')"
        SerieDeFactu = SerieDeFactu & ", " & Msg
    End If
    
    
    
    
    Msg = "SELECT  distinct codccost FROM tmpintegrapu WHERE codusu =" & vUsu.Codigo & " and codccost<>'' AND  NOT codccost IN (select codccost from ccoste)"
    miRsAux.Open Msg, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        
        NumAsien = NumAsien + 1
        
        'select codigo,texto1,texto2,observa1 from tmptesoreriacomun
        Msg = "(" & vUsu.Codigo & "," & NumAsien & "," & DBSet(miRsAux!codccost, "T") & ",'','No existe centro de coste')"
        SerieDeFactu = SerieDeFactu & ", " & Msg
    
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    
    If SerieDeFactu <> "" Then
        Conn.Execute "DELETE from tmptesoreriacomun WHERE codusu =" & vUsu.Codigo
        espera 0.5
        SerieDeFactu = Mid(SerieDeFactu, 2)
        Msg = "INSERT INTO  tmptesoreriacomun(codusu,codigo,texto1,texto2,observa1) VALUES " & SerieDeFactu
        Conn.Execute Msg
    Else
        ComprobarNumerosDeFactura = True
    End If


End Function








Private Function CrearFacturaProveedores(NA As Long, FechaEnt As Date, NumDiari As Integer) As Boolean
Dim i  As Integer
Dim cad As String
Dim SQL As String

Dim FinBucle As Boolean
Dim HayQueInsertarFactura As Boolean
Dim ReestableceValores As Boolean
Dim N As Long

Dim NumLiena2 As Integer
Dim InsTotales As String

Dim NumLienaB As Integer
Dim InsBases As String

Dim INsFactura As String
Dim InsrListadoFacturasFichero As String


Dim TotalFac As Currency
Dim Tbases As Currency
Dim TotIVA As Currency
Dim Suplidos As Currency
Dim TotalAprox As Currency

Dim Serie As String
Dim PrimeraContrapartida As Boolean
Dim ImportAuxiliar As Currency
Dim Rs As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim NIF As String
Dim Cta As String

Dim CambiaSerieFactura As Boolean


Dim IVA_Suplidos As Integer
Dim CadenaInserPorsiSuplidos As String



    On Error GoTo Salida
    Set Rs = New ADODB.Recordset
    Set R2 = New ADODB.Recordset

    cad = "select tmpintegrapu.* ,  if(substring(codmacta,1,3)='472',0,1) orden "
    cad = cad & " from tmpintegrapu where numasien=" & NA & " AND numdiari=" & NumDiari & " AND codusu =" & vUsu.Codigo
    cad = cad & " And fechaent='" & Format(FechaEnt, FormatoFecha) & "'  ORDER BY  orden,numserie,linliapu"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = "select codmacta,linliapu, timported,timporteh,numdocum "
    cad = cad & " from tmpintegrapu where numasien=" & NA & " AND numdiari=" & NumDiari & " AND codusu =" & vUsu.Codigo
    cad = cad & " And fechaent='" & Format(FechaEnt, FormatoFecha) & "' AND substring(codmacta,1,3)<>'472'  ORDER BY  codmacta"
    R2.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    'Vamos a fijar el total factura
    NumDeFactu = -1
        
    IVA_Suplidos = -1
    FinBucle = False
    
    While Not FinBucle
        HayQueInsertarFactura = False
        ReestableceValores = False
        If Rs.EOF Then
            HayQueInsertarFactura = True
        Else
            N = IIf(IsNull(Rs!numfaccl), 0, 1)
            If N > 0 Then
                CambiaSerieFactura = False
                If NumDeFactu <> Rs!numfaccl Then
                    CambiaSerieFactura = True
                Else
                    SQL = DBLet(Rs!NUmSerie, "T")
                    If Serie <> SQL Then CambiaSerieFactura = True
                End If
                If CambiaSerieFactura Then
                    If NumDeFactu > 0 Then HayQueInsertarFactura = True
                    ReestableceValores = True
                End If
            End If
        End If
        
            
        If HayQueInsertarFactura Then
            
            Suplidos = TotalAprox - TotalFac
            If Suplidos <> 0 Then
                
                If IVA_Suplidos < 0 Then
                    'Tenemos que obtener un IVA de suplidos
                    cad = DevuelveDesdeBD("codigiva", "tiposiva", "tipodiva", "4") 'SUPLIDOS=tipodigva 4
                    If cad = "" Then Err.Raise 513, , "No se ha configurado un IVA de suplidos"
                    IVA_Suplidos = Val(cad)
                   
                End If
                    
                'INSERTAMOS LA LINEACad = ", ('" & Rs!NUmSerie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt) & "," & NumLiena2
               
                cad = CadenaInserPorsiSuplidos & "," & NumLiena2
                cad = cad & "," & IVA_Suplidos & "," & DBSet(Suplidos, "N")
                cad = cad & ",0,NULL,0,null)"
                InsTotales = InsTotales & cad
                 NumLiena2 = NumLiena2 + 1
                
                'Stop
                TotalFac = TotalAprox
            End If
            INsFactura = Replace(INsFactura, "#BASES#", TransformaComasPuntos(CStr(Tbases)))
            INsFactura = Replace(INsFactura, "#BASESRET#", "null")
            INsFactura = Replace(INsFactura, "#totivas#", TransformaComasPuntos(CStr(TotIVA)))
            INsFactura = Replace(INsFactura, "#totrecargo#", "null")
            INsFactura = Replace(INsFactura, "#totfaccl#", TransformaComasPuntos(CStr(TotalFac)))
            'retencion
            INsFactura = Replace(INsFactura, "#retfaccl#", "null")
            INsFactura = Replace(INsFactura, "#trretfac#", "null")
            INsFactura = Replace(INsFactura, "#cuereten#", "null")
            INsFactura = Replace(INsFactura, "#tiporeten#", "0")
            If Suplidos = 0 Then
                cad = "null"
            Else
                cad = TransformaComasPuntos(CStr(Suplidos))
            End If
            INsFactura = Replace(INsFactura, "#suplidos#", cad)
            
            'If Serie <> "I" Then InsertaEnTmpInsertrs INsFactura
            InsertaEnTmpInsertrs INsFactura
            
            If InsTotales <> "" Then
                InsTotales = Mid(InsTotales, 2)
                cad = "INSERT INTO factpro_totales(numserie,numregis,fecharec,anofactu,numlinea,codigiva,baseimpo,porciva,porcrec,impoiva,imporec"
                cad = cad & ") VALUES " & InsTotales
                'If Serie <> "I" Then InsertaEnTmpInsertrs cad
                InsertaEnTmpInsertrs cad
            End If
            
            If InsBases <> "" Then
                InsBases = Mid(InsBases, 2)
                cad = "INSERT INTO factpro_lineas (numserie, numregis, fecharec,anofactu, numlinea, codmacta, baseimpo, codccost) VALUES "
                cad = cad & InsBases
                'If Serie <> "I" Then InsertaEnTmpInsertrs cad
                InsertaEnTmpInsertrs cad
            End If
             
             
            'Para la pantalla que indca cuantoas vamos a integrarar y cuales
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#BASES#", TransformaComasPuntos(CStr(Tbases)))
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#BASESRET#", "null")
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#totivas#", TransformaComasPuntos(CStr(TotIVA)))
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#totrecargo#", "null")
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#totfaccl#", TransformaComasPuntos(CStr(TotalFac)))
            'retencion
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#retfaccl#", "null")
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#trretfac#", "null")
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#cuereten#", "null")
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#tiporeten#", "0")
            If Suplidos = 0 Then
                cad = "null"
            Else
                cad = TransformaComasPuntos(CStr(Suplidos))
            End If
            InsrListadoFacturasFichero = Replace(InsrListadoFacturasFichero, "#suplidos#", cad)
            cad = "INSERT INTO tmpintefrafracli (codusu ,codigo ,serie ,factura ,fecha  ,cta_cli ,iban  ,impventa  ,impret ,impiva ,imprecargo,CalculoImponible ,totalfactura ,txtcsb )"
            cad = cad & " VALUES " & Mid(InsrListadoFacturasFichero, 2)
            'If Serie <> "I" Then Conn.Execute cad
            Conn.Execute cad
            
            '-------------
            'Seguimos con los datos, si hay
            If Not Rs.EOF Then
                ReestableceValores = True
            Else
                ReestableceValores = False
                FinBucle = True
            End If
        End If
        If ReestableceValores Then
                
                InsBases = ""
                NumDeFactu = Rs!numfaccl
                InsTotales = ""
                NumLienaB = 1
                TotalFac = 0
                TotIVA = 0
                Suplidos = 0
                
                INsFactura = ""
                Tbases = 0
                Serie = Rs!NUmSerie
                TotalAprox = 0
                NumLiena2 = 0
                
                CadenaInserPorsiSuplidos = ""
                'Vemos si lleva suplidos. Obtenemos el total factura
                If DBLet(Rs!ctacontr, "T") <> "" Then
                    R2.MoveFirst
                    
                    Do
                        R2.Find "codmacta = " & DBSet(Rs!ctacontr, "T"), , adSearchForward
                        If R2.EOF Then
                            NumLiena2 = 1
                            
                        Else
                            If InStr(1, R2!Numdocum, NumDeFactu) > 0 Then
                                TotalAprox = R2!timporteH
                                NumLiena2 = 1
                            Else
                                
                                R2.MoveNext
                            End If
                        End If
                    Loop Until NumLiena2 = 1
                End If
                InsBases = ""
                NumLiena2 = 1
                PrimeraContrapartida = True
        End If
        
        
        'Separamos datos del apunte
        If Not FinBucle Then
            If Mid(Rs!codmacta, 1, 3) = "472" Then
                'INSERT INTO factpro_lineas (numserie, numregis, fecharec,anofactu, numlinea, codmacta, baseimpo, codccost)
                
                cad = ", ('" & Rs!NUmSerie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt) & "," & NumLiena2
                i = DevuelveTipoIva(Rs!codmacta, True)
                cad = cad & "," & i & "," & DBSet(Rs!Baseimpo, "N")
                cad = cad & "," & DBSet(miRsAux!porceiva, "N") & ",NULL," & DBSet(Rs!timported, "N") & ",null"
                
                TotIVA = TotIVA + Rs!timported
                TotalFac = TotalFac + Rs!Baseimpo + Rs!timported
                Tbases = Tbases + Rs!Baseimpo
                InsTotales = InsTotales & cad & ")"
                NumLiena2 = NumLiena2 + 1
                
                'Por si hubieran suplidos
                
                CadenaInserPorsiSuplidos = ", ('" & Rs!NUmSerie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt)
                
            
                'La contrpartida es el cliente
                    'CLIENTE. Esta es el numero de factura
                If PrimeraContrapartida Then
                    PrimeraContrapartida = False
                    '", numdiari ,fechaent, numasien, fecliqcl, codconce340,codopera,no_modifica_apunte) VALUES  "
                    
                    
                    If IsNull(Rs!ctacontr) Then
                        Cta = "4100000000" 'Habra que personalizar
                    Else
                        Cta = Rs!ctacontr
                    End If
                    cad = DBSet(Cta, "T") & ",'Generada por traspaso contaplus'"
                    ' fecharec numfactu, codmacta,.observa
                    cad = DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!NumFacpr, "T") & "," & cad
                    
                    'numserie,numregis, anofactu ,fecfactu, & cad
                    cad = "'" & Serie & "'," & NumDeFactu & "," & Year(Rs!FechaEnt) & "," & DBSet(Rs!FecFactu, "F") & "," & cad
                    cad = cad & ",  #BASES# , #BASESRET# ,#totivas# ,#totrecargo# , #totfaccl# "
                    cad = cad & ",  #retfaccl# , #trretfac# ,#cuereten# ,#tiporeten# , #suplidos#"
                    cad = cad & ", " & Rs!NumDiari & "," & DBSet(Rs!FechaEnt, "F") & "," & Rs!NumAsien & "," & DBSet(Rs!FechaEnt, "F")
                    cad = cad & ",0 ,0 ,1"
                    
                    'nommacta, nifdatos   , en apuntes: bancotalonpag coddevol
                    SQL = DBLet(Rs!Nommacta, "T")
                    NIF = ""
                    If SQL = "" Then
                        NIF = "nifdatos"
                        SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T", NIF)
                        If SQL = "" Then
                            MsgBox "Nombre cuenta vacia: " & Cta, vbExclamation
                            SQL = "VACIO"
                        End If
                    End If
                    cad = cad & "," & DBSet(SQL, "T") & ","
                    
                    SQL = DBLet(Rs!NIF, "T")
                    If SQL = "" Then SQL = NIF
                    If SQL = "" Then
                        SQL = "null"
                    Else
                        SQL = DBSet(SQL, "T")
                    End If
                    cad = cad & SQL & ")"
                    SQL = FijarCadenaInsercionSQL(False)
                    INsFactura = SQL & cad
                    
                    'para el listado de facturas que vamops a insertar
                    'tmpintefrafracli
                    '                                             nommac   base    ret    iva    recar          suplidos     total        nºAsiento
                    'codusu codigo serie factura fecha  ctaventas iban  impventa  impret impiva imprecargo CalculoImponible totalfactura    txtcsb
                    InsrListadoFacturasFichero = ", (" & vUsu.Codigo & "," & NumRegElim
                    InsrListadoFacturasFichero = InsrListadoFacturasFichero & ",'" & Serie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F")
                    InsrListadoFacturasFichero = InsrListadoFacturasFichero & ",'" & Cta & "',null,"
                    '                                                           base      ret           iva      recar          suplidos     total
                    InsrListadoFacturasFichero = InsrListadoFacturasFichero & " #BASES# , #BASESRET# ,#totivas# ,#totrecargo# , #suplidos# , #totfaccl# , " & NA & ")"
                Else
                    'Stop
                End If
            Else
                'BASES
                'Sql = "INSERT INTO factcli_lineas (numserie, numfactu, anofactu, numlinea, codmacta, baseimpo, codccost) VALUES ("
                If Val(Mid(Rs!codmacta, 1, 3)) > 419 Then
                    cad = ", ('" & Serie & "'," & NumDeFactu & "," & DBSet(Rs!FechaEnt, "F") & "," & Year(Rs!FechaEnt) & "," & NumLienaB
                    If IsNull(Rs!timported) Then
                        ImportAuxiliar = -Rs!timporteH
                    Else
                        ImportAuxiliar = Rs!timported
                    End If
                    cad = cad & ",'" & Rs!codmacta & "'," & DBSet(ImportAuxiliar, "N") & ",'" & Rs!codccost & "')"
                    NumLienaB = NumLienaB + 1
                    InsBases = InsBases & cad
                End If
            End If
        End If
        If Not FinBucle Then Rs.MoveNext
    Wend
    
    
    Rs.Close
    CrearFacturaProveedores = True 'Si llega aqui ha ido bien
Salida:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Inserta Factura" & vbCrLf & Err.Description
        Err.Clear
    End If
    Set Rs = Nothing
    Set R2 = Nothing
End Function


