Attribute VB_Name = "Contabilizar"
Option Explicit


        'Se ha añadido un concepto mas a la ampliacion  26 Abril 2007
        '------------------------------------------------------------
        ' De momento lo resolveremos con un simple
        '    devuelvedesdebd.   Si se realentiza mucho deberiamos obtener un recodset
        '    con las titlos de las contrapartidas si el tipo de ampliacion es el 4



Public Sub InsertaTmpActualizar(NumAsien, NumDiari, FechaEnt)
Dim SQL As String
        SQL = "INSERT INTO tmpactualizar (numdiari, fechaent, numasien, codusu) VALUES ("
        SQL = SQL & NumDiari & ",'" & Format(FechaEnt, FormatoFecha) & "'," & NumAsien
        SQL = SQL & "," & vUsu.Codigo & ")"
        Conn.Execute SQL
End Sub


'TipoRemesa:
'           0. Efecto
'           1. Pagare
'           2. Talon
'
' El abono(CONTABILIZACION) de la remesa sera la 572 contra 5208(del banco)
'Y punto, como mucho los gastos si quiero contabilizarlis
Public Function ContabilizarRecordsetRemesa(TipoRemesa As Byte, Norma19 As Boolean, Codigo As Integer, Anyo As Integer, CtaBanco As String, FechaAbono As Date, GastosBancarios As Currency, CtaPuenteRemesa As String) As Boolean
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaParametros As String
Dim Cuenta As String
Dim CuentaPuente As Boolean

'Dim ImporteTalonPagare As Currency    'beneficiosPerdidasTalon: por si hay diferencias entre vtos y total talon
Dim ImpoAux As Currency
Dim VaAlHaber As Boolean
Dim Aux As String
Dim GastosGeneralesRemesasDescontadosDelImporte As Boolean
Dim LCta As Integer
'Noviembre 2009.
'Paramero nuevo
'Contabiliza contra cuenta efectos comerciales decontados
'Son DOS apuntes en el abono
Dim LlevaCtaEfectosComDescontados As Boolean
Dim CtaEfectosComDescontados As String

Dim Obs As String

    On Error GoTo ECon
    ContabilizarRecordsetRemesa = False

    
    GastosGeneralesRemesasDescontadosDelImporte = False
    Cuenta = "GastRemDescontad" 'gastos tramtiaacion remesa descontados importe
    CtaParametros = DevuelveDesdeBD("ctaefectosdesc", "bancos", "codmacta", RecuperaValor(CtaBanco, 1), "T", Cuenta)
    GastosGeneralesRemesasDescontadosDelImporte = Cuenta = "1"
    If GastosGeneralesRemesasDescontadosDelImporte Then
        'Si no tiene gastos generales pongo esto a false tb
        If GastosBancarios = 0 Then GastosGeneralesRemesasDescontadosDelImporte = False
    End If
    Cuenta = ""
    LlevaCtaEfectosComDescontados = False   'Solo sera para efectos bancarios. Tipo FONTENAS
    
    'La forma de pago
    Set vCP = New Ctipoformapago
    If TipoRemesa = 1 Then
        Linea = vbTipoPagoRemesa
        Cuenta = "Efectos"
        'CuentaPuente = CtaPuenteRemesa <> ""
        'CtaParametros = CtaPuenteRemesa
    ElseIf TipoRemesa = 2 Then
        Linea = vbPagare
        Cuenta = "Pagarés"
        'CtaParametros = "pagarecta"
        CuentaPuente = vParamT.PagaresCtaPuente
        
    ElseIf TipoRemesa = 5 Then
        Linea = vbConfirming
        Cuenta = "Confirming"
        CuentaPuente = vParamT.ConfirmingCtaPuente
    Else
        Linea = vbTalon
        Cuenta = "Talones"
        'CtaParametros = "taloncta"
        CuentaPuente = vParamT.TalonesCtaPuente
    End If
    
    
    
    If CuentaPuente Then
        If CtaParametros = "" Then
            MsgBox "Mal configurado el banco. Falta configurar cuenta efectos descontados del banco: " & Cuenta, vbExclamation
            Exit Function
        End If
    End If
            
            
    
    
    
            
    'Si llevamos las dos cuentas de efectos descontados, la de cancelacion YA las combrpobo en el proceso de cancelacion
    'ahora tenemos que comprobar la de efectos descontados pendientes de cobro
    If LlevaCtaEfectosComDescontados Then
        Set Rs = New ADODB.Recordset
        LCta = Len(CtaEfectosComDescontados)
        If LCta < vEmpresa.DigitosUltimoNivel Then
        
            Conn.Execute "DELETE from tmpcierre1 where codusu = " & vUsu.Codigo
                
            Ampliacion = ",CONCAT('" & CtaEfectosComDescontados & "',SUBSTRING(codmacta," & LCta + 1 & ")" & ")"
            ''SQL = "Select " & vUsu.Codigo & Ampliacion & " from scarecepdoc where codigo=" & IdRecepcion
                
            SQL = "Select " & vUsu.Codigo & Ampliacion & " from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
            SQL = SQL & " GROUP BY codmacta"
            'INSERT
            SQL = "INSERT INTO tmpcierre1(codusu,cta) " & SQL
            Conn.Execute SQL
            
            'Ahora monto el select para ver que cuentas 430 no tienen la 4310
            SQL = "Select cta,codmacta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
            SQL = SQL & " HAVING codmacta is null"
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            Linea = 0
            While Not Rs.EOF
                Linea = Linea + 1
                SQL = SQL & Rs!Cta & "     "
                If Linea = 5 Then
                    SQL = SQL & vbCrLf
                    Linea = 0
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            
            If SQL <> "" Then
                
                AmpRemesa = "Abono remesa"
                
                SQL = "Cuentas " & AmpRemesa & ".  No existen las cuentas: " & vbCrLf & String(90, "-") & vbCrLf & SQL
                SQL = SQL & vbCrLf & "¿Desea crearlas?"
                Linea = 1
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                    'Ha dicho que si desea crearlas
                    
                    Ampliacion = "CONCAT('" & CtaEfectosComDescontados & "',SUBSTRING(codmacta," & LCta + 1 & ")) "
                    
                    'SQL = "Select codmacta," & Ampliacion & " from scarecepdoc where codigo=" & IdRecepcion
                    SQL = "Select codmacta," & Ampliacion & " from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
                    SQL = SQL & " and " & Ampliacion & " in "
                    SQL = SQL & "(Select cta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
                    SQL = SQL & " AND codmacta is null)"
                    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not Rs.EOF
                    
                         SQL = "INSERT IGNORE INTO  cuentas(codmacta,nommacta ,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) SELECT '"
                                    ' CUenta puente
                         SQL = SQL & Rs.Fields(1) & "',nommacta ,'S',0,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from cuentas where codmacta = '"
                                    'Cuenta en la scbro (codmacta)
                         SQL = SQL & Rs.Fields(0) & "'"
                         Conn.Execute SQL
                         Rs.MoveNext
                         
                    Wend
                    Rs.Close
                    Linea = 0
                End If
                If Linea = 1 Then GoTo ECon
            End If
            
        Else
            'Cancela contra UNA unica cuenta todos los vencimientos
            SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", CtaEfectosComDescontados, "T")
            If SQL = "" Then
                MsgBox "No existe la cuenta efectos comerciales descontados : " & CtaEfectosComDescontados, vbExclamation
                GoTo ECon
            End If
        End If
        Set Rs = Nothing
    End If  'de comprobar cta efectos comerciales descontados
            
            
    If vCP.Leer(Linea) = 1 Then GoTo ECon
    
    
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Abono remesa: " & Codigo & " / " & Anyo & "   " & Cuenta & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "',"
    
    Obs = Codigo & " / " & Anyo & "   " & Cuenta
    
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Abono remesa: " & Obs & "');"
    If Not Ejecuta(SQL) Then Exit Function
    
    
    Linea = 1
    Importe = 0
    Gastos = 0
    Set Rs = New ADODB.Recordset
    
    
    
    
    'La ampliacion para el banco
    AmpRemesa = ""
    SQL = "Select * from remesas WHERE codigo=" & Codigo & " AND anyo = " & Anyo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'NO puede ser EOF
    
    
    Importe = Rs!Importe

    
    If Not IsNull(Rs!Descripcion) Then AmpRemesa = Rs!Descripcion
    
    
    If AmpRemesa = "" Then
        AmpRemesa = " Remesa: " & Codigo & "/" & Anyo
    Else
        AmpRemesa = " " & AmpRemesa
    End If
    
    Rs.Close
    
    'AHORA Febrero 2009
    '572 contra  5208  Efectos descontados
    '-------------------------------------
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, "
    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"


    Gastos = 0
    If CuentaPuente Then
        
        'DOS LINEAS POR APUNTE, banco contra efectos descontados
        'A no ser que sea TAL/PAG y pueden haber beneficios o perdidas por diferencias de importes
        SQL = SQL & CtaParametros & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli
    
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        Ampliacion = Ampliacion & " RE. " & Codigo & "/" & Anyo
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
    
    
        SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,"
    
        If vCP.ctrhacli = 1 Then
            If CuentaPuente And Not LlevaCtaEfectosComDescontados Then
                SQL = SQL & "'" & RecuperaValor(CtaBanco, 1) & "',"
            Else
                'NO lleva cuenta puente
                'Directamente contra el cliente
                If Not LlevaCtaEfectosComDescontados Then
                    SQL = SQL & "'" & Rs!codmacta & "',"
                Else
                    SQL = SQL & "NULL,"
                End If
            End If
        Else
            SQL = SQL & "NULL,"
        End If
        SQL = SQL & "'COBROS',0,"
        
       
        Aux = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
        Rs.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        'los datos de la factura (solo en el apunte del cliente)
        Dim TipForpa As Byte
        TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Rs!Codforpa, "N")
        
        SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & "," & DBSet(TipForpa, "N") & ","
        SQL = SQL & TipoRemesa & "," & Codigo & "," & Anyo & ")"

        

        If Not Ejecuta(SQL) Then Exit Function
  
        Linea = Linea + 1
    
    
    
       'Lleva cta efectos comerciales descontados
        If LlevaCtaEfectosComDescontados Then
            'AQUI
            'Para cada efecto cancela la 5208 contra las CtaEfectosComDescontados(4311x)
 
            
            Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            
            
            SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
            Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not Rs.EOF
        
                'Banco contra cliente
                'La linea del banco
                SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
                SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
                SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
                SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
            
                'Cuenta
                SQL = SQL & CtaEfectosComDescontados
                If LCta <> vEmpresa.DigitosUltimoNivel Then SQL = SQL & Mid(Rs!codmacta, LCta + 1)
                
                SQL = SQL & "','" & Format(Rs!NumFactu, "000000000") & "'," & vCP.conhacli
            
            
                
                Ampliacion = Aux & " "
            
                                'Neuvo dato para la ampliacion en la contabilizacion
                Select Case vCP.amphacli
                Case 2
                   Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
                Case 4
                    'Contrapartida BANCO
                    Cuenta = RecuperaValor(CtaBanco, 1)
                    Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                    Ampliacion = Ampliacion & AmpRemesa
                Case 6
                
                Case Else
                   If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                   Ampliacion = Ampliacion & Rs!NUmSerie & "/" & Rs!NumFactu
                End Select
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
                
                
                ' debe timporteH, codccost, ctacontr, idcontab, punteada
                'Importe
                SQL = SQL & TransformaComasPuntos(Rs!ImpVenci) & ",NULL,NULL,"
            
                If vCP.ctrdecli = 1 Then
                    SQL = SQL & "'" & CtaParametros & "',"
                Else
                    SQL = SQL & "NULL,"
                End If
                SQL = SQL & "'COBROS',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & "," & ValorNulo & ValorNulo & "," & ValorNulo & ")"
                '###FALTA1
                
                
                If Not Ejecuta(SQL) Then Exit Function
                
                Linea = Linea + 1
                Rs.MoveNext
            Wend
            Rs.Close
            
        End If   'de lleva cta de efectos comerciales descontados
        
        
    Else
        
        
        
        Importe = 0
        SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
        
            'Banco contra cliente
            'La linea del banco
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
        
            'Cuenta
            SQL = SQL & Rs!codmacta & "','" & Rs!NUmSerie & Format(Rs!NumFactu, "0000000") & "'," & vCP.conhacli
    
            
            
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
            Ampliacion = Ampliacion & " "
                   
            'Neuvo dato para la ampliacion en la contabilizacion
            Select Case vCP.amphacli
            Case 2
               Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
            Case 4
                'Contrapartida BANCO
                Cuenta = RecuperaValor(CtaBanco, 1)
                Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                Ampliacion = Ampliacion & AmpRemesa
            Case 6
                
                Ampliacion = DBLet(Rs!nomclien, "T")
                If Ampliacion = "" Then Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!codmacta, "T")
                
                MiVariableAuxiliar = Rs!NUmSerie & Format(Rs!NumFactu, "0000000")
                Ampliacion = Mid(Ampliacion, 1, 34 - Len(MiVariableAuxiliar))
                Ampliacion = Ampliacion & " " & MiVariableAuxiliar
                
                
                
                
                
                
            Case Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFactu, "0000000")
            End Select
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            Importe = Importe + Rs!ImpVenci
                
            Gastos = Gastos + DBLet(Rs!Gastos, "N")
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe
            SQL = SQL & "NULL," & TransformaComasPuntos(Rs!ImpVenci) & ",NULL,"
        
            If vCP.ctrdecli = 1 Then
                SQL = SQL & "'" & RecuperaValor(CtaBanco, 1) & "',"
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'COBROS',0,"
            
            'los datos de la factura (solo en el apunte del cliente)
            TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Rs!Codforpa, "N")
            
            SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & "," & DBSet(TipForpa, "N") & ","
            SQL = SQL & TipoRemesa & "," & Codigo & "," & Anyo & ")"
            
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
            Rs.MoveNext
        
        Wend
        Rs.Close
            
    End If
    
    'La linea del banco
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","

    
    'Gastos de los recibos.
    'Si tiene alguno de los efectos remesados gastos
    If Gastos > 0 Then
        Linea = Linea + 1
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        Ampliacion = "RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli & ",'" & Ampliacion & " " & Codigo & "/" & Anyo & "'"



        Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 4) & "','" & Ampliacion & ",NULL,"
        Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Gastos)) & ","

        If RecuperaValor(CtaBanco, 3) = "" Then
            Ampliacion = Ampliacion & "NULL"
        Else
            Ampliacion = Ampliacion & "'" & RecuperaValor(CtaBanco, 3) & "'"
        End If
        
        Ampliacion = Ampliacion & ",NULL,'COBROS',0)"

        Ampliacion = SQL & Ampliacion
        If Not Ejecuta(Ampliacion) Then Exit Function
        Linea = Linea + 1
    End If
    
  
    'AGOSTO 2009
    'Importe final banco
    'Y desglose en TAL/PAG de los beneficios o perdidas si es que tuviera
    
    ImpoAux = Importe + Gastos
    
    'NOV 2009
    'Gastos tramitacion descontados del importe
    If GastosGeneralesRemesasDescontadosDelImporte And GastosBancarios > 0 Then
        ImpoAux = ImpoAux - GastosBancarios
        'Para que la linea salga al final del asiento, juego con numlinea
        Linea = Linea + 1
    End If
    
    Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    Ampliacion = Ampliacion & AmpRemesa
    Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 1) & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli & ",'" & Ampliacion & "',"
    Ampliacion = Ampliacion & TransformaComasPuntos(CStr(ImpoAux)) & ",NULL,NULL,"
    
    If vCP.ctrdecli = 0 Then
        Ampliacion = Ampliacion & "NULL"
    Else
        If CuentaPuente Then
            If Not LlevaCtaEfectosComDescontados Then
                Ampliacion = Ampliacion & "'" & CtaParametros & "'"
            Else
                Ampliacion = Ampliacion & "NULL"
            End If
        Else
            Ampliacion = Ampliacion & "NULL"
        End If
       
    End If
    Ampliacion = Ampliacion & ",'COBROS',0)"


    Ampliacion = SQL & Ampliacion
    If Not Ejecuta(Ampliacion) Then Exit Function
    
    'Juego con la linea
    
    'Gastos bancarios derivados de la tramitacion de la remesa.
    'Metemos dos lineas mas. Podriamos meter una si en el importe anterior le restamos los gastos bancarios
    If GastosBancarios > 0 Then
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,tiporem,codrem,anyorem) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
        
        
        
        'imporeted timporteH, codccost, ctacontr, idcontab, punteada) "
        If GastosGeneralesRemesasDescontadosDelImporte Then
            'He jugado con el orden para k la linea anterior salga la ultima
            Linea = Linea - 1
        Else
            Linea = Linea + 1
        End If
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
        Ampliacion = Ampliacion & " Gastos Remesa:" & Codigo & " / " & Anyo
        Ampliacion = DevNombreSQL(Ampliacion)
    
        ' numdocum, codconce, ampconce
        Ampliacion = "'RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli & ",'" & Ampliacion & "',"
        Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 2) & "'," & Ampliacion
        
        Ampliacion = Ampliacion & DBSet(GastosBancarios, "N") & ",NULL,"
        'CENTRO DE COSTE
        If vParam.autocoste Then
            Ampliacion = Ampliacion & "'" & RecuperaValor(CtaBanco, 3) & "'"
        Else
            Ampliacion = Ampliacion & "NULL"
        End If
        Ampliacion = Ampliacion & ",'" & RecuperaValor(CtaBanco, 1) & "','COBROS',0," & TipoRemesa & "," & DBSet(Codigo, "N") & "," & DBSet(Anyo, "N") & ")"
        Ampliacion = SQL & Ampliacion
        
        If Not Ejecuta(Ampliacion) Then Exit Function
        
        If Not GastosGeneralesRemesasDescontadosDelImporte Then
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
            
            
            
            Linea = Linea + 1
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
            Ampliacion = Ampliacion & " Gastos Remesa: " & Codigo & " / " & Anyo
            Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 1) & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
            Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(GastosBancarios)) & ",NULL,'"
            Ampliacion = Ampliacion & RecuperaValor(CtaBanco, 2) & "','COBROS',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then Exit Function
        End If
            
        If GastosGeneralesRemesasDescontadosDelImporte Then Linea = Linea + 2
    End If
    
    
    'Noviembre 2009
    '-------------------------------------------
    'Efectos. Si lleva cta puente, y lleva la segunda cuenta puente
    If LlevaCtaEfectosComDescontados Then
    
        'Crearemos n x 2 lineas de apunte de los efectos remesados
        'siendo
        '       CtaEfectosComDescontados        contra   CtaParametros (431x)
        '            y el aseinto de contrapartida
    
        Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        CtaEfectosComDescontados = DevuelveDesdeBD("RemesaCancelacion", "paramtesor", "codigo", "1")
        LCta = Len(CtaEfectosComDescontados)
        If LCta = 0 Then
            MsgBox "Deberia tener valor el paremtro de cta puente", vbCritical
            LCta = Val(Rs!davidadavi) 'QUE GENERE UN ERROR
        End If
        
        CtaParametros = RecuperaValor(CtaBanco, 1) 'Cuenta del banco para la contrpartida
        Linea = Linea + 1
        SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
        
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
        
            'Cuenta
            SQL = SQL & CtaEfectosComDescontados
            If LCta <> vEmpresa.DigitosUltimoNivel Then SQL = SQL & Mid(Rs!codmacta, LCta + 1)
            
            SQL = SQL & "','" & Rs!NUmSerie & Format(Rs!NumFactu, "0000000") & "'," & vCP.conhacli
        
        
            
            Ampliacion = Aux & " "
        
                            'Neuvo dato para la ampliacion en la contabilizacion
            Select Case vCP.amphacli
            Case 2
               Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
            Case 4
                'Contrapartida BANCO
                Cuenta = RecuperaValor(CtaBanco, 1)
                Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                Ampliacion = Ampliacion & AmpRemesa
            Case Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFactu, "0000000")
            End Select
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe
            SQL = SQL & "NULL," & TransformaComasPuntos(Rs!ImpVenci) & ",NULL,"
        
            If vCP.ctrdecli = 1 Then
                SQL = SQL & "'" & CtaParametros & "',"
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'COBROS',0)"
            
            If Not Ejecuta(SQL) Then Exit Function
            Linea = Linea + 1
            
            Rs.MoveNext
        Wend
        Rs.Close
    
    End If
    

    'AHora actualizamos los efectos.
    SQL = "UPDATE cobros SET"
    SQL = SQL & " siturem= 'Q'"
    SQL = SQL & ", situacion = 1 "
'    SQL = SQL & ", ctabanc2= '" & RecuperaValor(CtaBanco, 1) & "'"
'    SQL = SQL & ", contdocu= 1"   'contdocu indica k se ha contabilizado
    SQL = SQL & " WHERE codrem=" & Codigo
    SQL = SQL & " and anyorem=" & Anyo
'++ la he añadido yo, antes no estaba
    SQL = SQL & " and tiporem = " & TipoRemesa
    
    Conn.Execute SQL

    Dim MaxLin As Integer

    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaAbono
    
    'Todo OK
    ContabilizarRecordsetRemesa = True
    
ECon:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set Rs = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function


'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'   DEVOLUCION DE REGISTROS
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------


    'OK. Ya tengo grabada la temporal con los recibos que devuelvo. Ahora
    'hare:
    '       - generar un asiento con los datos k devuelvo
    '       - marcar los cobros como devueltos, añadirle el gasto y insertar en la
    '           tabla de hco de devueltos
    
    'La variable remesa traera todos los valores
    
    '21 Octubre 2011
    'Desdoblaremos el procedimiento de deolucion
    'de talones-pagares frente a efectos
Public Function RealizarDevolucionRemesa(FechaDevolucion As Date, ContabilizoGastoBanco As Boolean, CtaBenBancarios As String, Remesa As String, DatosContabilizacionDevolucion As String) As Boolean
Dim C As String
    
    C = RecuperaValor(Remesa, 10)
    
    CtaBenBancarios = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", RecuperaValor(Remesa, 3), "T")
    If CtaBenBancarios = "" Then
        CtaBenBancarios = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
    End If
    
    
    If C = "1" Then
        RealizarDevolucionRemesa = RealizarDevolucionRemesaEfectos(FechaDevolucion, ContabilizoGastoBanco, CtaBenBancarios, Remesa, DatosContabilizacionDevolucion)
    Else
        RealizarDevolucionRemesa = RealizarDevolucionRemesaTalPag(FechaDevolucion, ContabilizoGastoBanco, CtaBenBancarios, Remesa, DatosContabilizacionDevolucion)
    End If
    
End Function


Public Function RealizarDevolucionRemesaEfectos(FechaDevolucion As Date, ContabilizoGastoBanco As Boolean, CtaBenBancarios As String, Remesa As String, DatosContabilizacionDevolucion As String) As Boolean

'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim Amp11 As String
Dim DescRemesa As String
Dim CuentaPuente2 As Boolean
Dim TipoRemesa As Byte
Dim SubCtaPte2 As String
'Dim AgrupaApunteBanco As Boolean
Dim GastoDevolucion As Currency
Dim DescuentaImporteDevolucion As Boolean
Dim GastoVto As Currency
Dim Gastos As Currency  'de cada recibo/vto
Dim Aux As String
Dim Importeauxiliar As Currency
Dim CtaBancoGastos As String
Dim CCBanco As String
Dim Agrupa431x As Boolean
Dim Agrupa4311x As Boolean   'Segunad cuenta de cancelacion TIPO fontenas
Dim CtaEfectosComDescontados As String   '   tipo FONTENAS
Dim LinApu As String

    On Error GoTo ECon
    RealizarDevolucionRemesaEfectos = False
    
   
    'La forma de pago
    Set vCP = New Ctipoformapago
    Set Rs = New ADODB.Recordset
    
    
    'Leo la descipcion de la remesa si alguna de las ampliaciones me la solicita
    
    SubCtaPte2 = "RemesaCancelacion"
    SQL = DevuelveDesdeBD("contaefecpte", "paramtesor", "1", "1", , SubCtaPte2)
    CuentaPuente2 = False
    If SQL <> "" Then
        If Val(SQL) <> 0 Then CuentaPuente2 = True
        
    End If
    
    DescRemesa = ""
    Aux = RecuperaValor(Remesa, 8)
    If Aux <> "" Then
        'OK viene de fichero
        Aux = RecuperaValor(Remesa, 9)
        'Vuelvo a susitiuri los # por |
        Aux = Replace(Aux, "#", "|")
        SQL = ""
        For Linea = 1 To Len(Aux)
            If Mid(Aux, Linea, 1) = "·" Then SQL = SQL & "X"
        Next
        
        If Len(SQL) > 1 Then
            'Tienen mas de una remesa
            SQL = ""
            While Aux <> ""
                Linea = InStr(1, Aux, "·")
                If Linea = 0 Then
                    Aux = ""
                Else
                    SQL = SQL & ",    " & Format(RecuperaValor(Mid(Aux, 1, Linea - 1), 1), "000") & "/" & RecuperaValor(Mid(Aux, 1, Linea - 1), 2) & ""
                    Aux = Mid(Aux, Linea + 1)
                End If
            
            Wend
            Aux = RecuperaValor(Remesa, 8)
            SQL = "Devolución remesas: " & Trim(Mid(SQL, 2))
            DescRemesa = SQL & vbCrLf & "Fichero: " & Aux
        End If
        
    End If

    
    
    DescRemesa = RecuperaValor(Remesa, 9)
    TipoRemesa = RecuperaValor(Remesa, 10)
    
    
    If TipoRemesa = 1 Then
        Linea = vbTipoPagoRemesa

    
        SQL = "ctaefectcomerciales"
    Else
    
       Err.Raise 513, , "Error en codigo. Parametro tipo remesa incorrecto. "
    
        If TipoRemesa = 2 Then
            Linea = vbPagare
            SQL = "pagarecta"
            
        ElseIf TipoRemesa = 5 Then
            Linea = vbConfirming
            SQL = "confirmingcta"
        Else
            Linea = vbTalon
            SQL = "taloncta"
        End If

    End If
    
    If vCP.Leer(Linea) = 1 Then GoTo ECon


    'Los parametros de contbilizacion se le pasan en el frame de pedida de datos
    'Ahora se los asignaremos a la formma de pago
    vCP.condecli = RecuperaValor(DatosContabilizacionDevolucion, 1)
    vCP.ampdecli = RecuperaValor(DatosContabilizacionDevolucion, 2)
    vCP.conhacli = RecuperaValor(DatosContabilizacionDevolucion, 1) '3)
    vCP.amphacli = RecuperaValor(DatosContabilizacionDevolucion, 2) '4)
    SQL = RecuperaValor(DatosContabilizacionDevolucion, 5)  'agupa o no
    Agrupa431x = SQL = "1"
    If Len(SubCtaPte2) <> vEmpresa.DigitosUltimoNivel Then Agrupa431x = False
    
    
    SQL = RecuperaValor(Remesa, 7)
    GastoDevolucion = TextoAimporte(SQL)
    DescuentaImporteDevolucion = False
    If GastoDevolucion > 0 Then
        CtaBancoGastos = "CtaIngresos"
        SQL = RecuperaValor(Remesa, 3)
        SQL = DevuelveDesdeBD("GastRemDescontad", "bancos", "codmacta", SQL, "T")
        If SQL = "1" Then DescuentaImporteDevolucion = True
    End If
    
    'Datos del banco
    SQL = RecuperaValor(Remesa, 3)
    SQL = "Select * from bancos where codmacta ='" & SQL & "'"
    CCBanco = ""
    CtaBancoGastos = ""
    CtaEfectosComDescontados = ""
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        SQL = "No se ha encontrado banco: " & vbCrLf & SQL
        Err.Raise 516, SQL
    End If
    CCBanco = DBLet(Rs!CodCcost, "T")
    CtaBancoGastos = DBLet(Rs!ctagastos, "T")
    If Not vParam.autocoste Then CCBanco = ""  'NO lleva analitica
    CtaEfectosComDescontados = DBLet(Rs!ctaefectosdesc, "T")
    Rs.Close
    
    
    If TipoRemesa = 1 Then
        If CtaEfectosComDescontados = "" Then CtaEfectosComDescontados = DevuelveDesdeBD("ctaefectcomerciales", "paramtesor", "codigo", "1")
    Else
        CtaEfectosComDescontados = ""
    End If
    Agrupa4311x = False 'La de fontenas
    If Agrupa431x Then
        Err.Raise 513, , "Error en codigo. Parametro tipo remesa incorrecto. "
        'QUIERE AGRUPAR. Veremos is por la longitud de las puentes, puede agrupar
        Agrupa4311x = True
        If Len(SubCtaPte2) <> vEmpresa.DigitosUltimoNivel Then Agrupa431x = False 'NO puede agrupar
        If Len(CtaEfectosComDescontados) <> vEmpresa.DigitosUltimoNivel Then Agrupa4311x = False 'NO puede agrupar
        
    End If
    
    'EMPEZAMOS
    'Borramos tmpactualizar
    SQL = "DELETE FROM tmpactualizar where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    'Cargaremos los registros a devolver que estaran en la tabla temporal
    'para codusu
    SQL = "Select * from tmpfaclin where codusu=" & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "EOF.  NO se han cargado datos devolucion", vbExclamation
        Rs.Close
        GoTo ECon
    End If

    Set Mc = New Contadores


    If Mc.ConseguirContador("0", FechaDevolucion <= vParam.fechafin, True) = 1 Then GoTo ECon


    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ",'"
    
    'Ahora esta en desc remesa
    DescRemesa = DescRemesa & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:nn") & " por " & vUsu.Nombre
    SQL = SQL & DevNombreSQL(DescRemesa) & "',"
    'SQL = SQL & "'Devolucion remesa: " & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2)
    'SQL = SQL & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "')"
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Devolución efectos')"

    
    If Not Ejecuta(SQL) Then GoTo ECon




    Linea = 1
    Importe = 0

    If vCP.ampdecli = 3 Then
        Amp11 = DescRemesa
    Else
        Amp11 = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    End If
    
    'Lo meto en una VAR
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, "
    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem,esdevolucion) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
    LinApu = SQL
    
    While Not Rs.EOF

        'Lineas de apuntes .
         SQL = LinApu & Linea & ",'"
         SQL = SQL & Rs!Cta
         SQL = SQL & "','" & Rs!NUmSerie & Format(Rs!NumFac, "0000000") & "'," & vCP.condecli

        Ampliacion = Amp11 & " "
    
        If vCP.ampdecli = 3 Then
            'NUEVA forma de ampliacion
            'No hacemos nada pq amp11 ya lleva lo solicitado
            
        Else
            If vCP.ampdecli = 4 Then
                'COntrapartida
                Ampliacion = Ampliacion & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!Cta, "T")
                
            Else
                If vCP.ampdecli = 2 Then
                   Ampliacion = Ampliacion & Format(Rs!Fecha, "dd/mm/yyyy")
                Else
                
                    If vCP.ampdecli = 6 Then
                        
                        Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!Cta, "T")
                        
                        MiVariableAuxiliar = Rs!NUmSerie & Format(Rs!NumFac, "0000000")
                        Ampliacion = Mid(Ampliacion, 1, 34 - Len(MiVariableAuxiliar))
                        Ampliacion = Ampliacion & " " & MiVariableAuxiliar
                
                    Else
                        If vCP.ampdecli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                        'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                        Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFac, "0000000") ' & "/" & RS!NumFac
                    End If
                End If
            End If
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 35)) & "',"

        Importe = Importe + Rs!Imponible


        GastoVto = 0
        Aux = " numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
        Aux = Aux & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden"
        Aux = DevuelveDesdeBD("gastos", "cobros", Aux, CStr(Rs!NIF), "N")
        
        If Aux <> "" Then GastoVto = CCur(Aux)
        Gastos = Gastos + GastoVto

        ' timporteH, codccost, ctacontr, idcontab, punteada
        Importeauxiliar = Rs!Imponible - GastoVto
        SQL = SQL & TransformaComasPuntos(CCur(Importeauxiliar)) & ",NULL,NULL,"

        If vCP.ctrdecli = 1 Then
            SQL = SQL & "'" & Rs!Cliente & "',"
        Else
            SQL = SQL & "NULL,"
        End If
        SQL = SQL & "'COBROS',0,"
        
        '%%%%% aqui van todos los datos de la devolucion en la linea de cuenta
        SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFac, "N") & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Rs!NIF, "N") & ","
            
         '-------------------------------------------------------------------------------------
         'Ahora
         '-------------------------------------------------------------------------------------
         'Lo pongo en uno
             'Actualizamos el registro. Ponemos la marca de devuelto. Y aumentamos el importe de gastos
         'Si es que hay
         Dim SqlCobro As String
         Dim RsCobro As ADODB.Recordset
         Dim ImporteNue As Currency
         
         SqlCobro = "select tipforpa, tiporem, codrem, anyorem, gastos from cobros inner join formapago on cobros.codforpa = formapago.codforpa "
         SqlCobro = SqlCobro & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
         SqlCobro = SqlCobro & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
         
         Set RsCobro = New ADODB.Recordset
         RsCobro.Open SqlCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         If Not RsCobro.EOF Then
         
'    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem) "
            SQL = SQL & DBSet(RsCobro!TipForpa, "N") & "," & DBSet(FechaDevolucion, "F") & "," & DBSet(Rs!CtaBase, "T", "S") & ","
            SQL = SQL & DBSet(Rs!ImpIva, "N") & "," & DBSet(RsCobro!Tiporem, "N") & "," & DBSet(RsCobro!Codrem, "N") & "," & DBSet(RsCobro!Anyorem, "N") & ",1)"
              
         
            Ampliacion = "UPDATE cobros SET "
            Ampliacion = Ampliacion & " Devuelto = 1, situacion = 0   "
            ImporteNue = Rs!Total - Rs!Imponible '- Rs!impiva
            
            ImporteNue = DBLet(RsCobro!Gastos, "N")
            If DBLet(Rs!ImpIva, "N") > 0 Then
            
                If ImporteNue = 0 Then
                    Ampliacion = Ampliacion & " , Gastos = " & TransformaComasPuntos(CStr(Rs!ImpIva))
                Else
                    Ampliacion = Ampliacion & " , Gastos = Gastos + " & TransformaComasPuntos(CStr(Rs!ImpIva))
                End If
            End If
            Ampliacion = Ampliacion & " ,impcobro=NULL,codrem= NULL, anyorem = NULL, siturem = NULL,tiporem=NULL,fecultco=NULL,recedocu=0"
            Ampliacion = Ampliacion & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
            Ampliacion = Ampliacion & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
            
            Ejecuta Ampliacion
             
         End If
         Set RsCobro = Nothing

        '%%%%% hasta aqui
        

        If Not Ejecuta(SQL) Then GoTo ECon

        Linea = Linea + 1
        
        
        
        'Gasto.
        ' Si tiene y no agrupa
        '-------------------------------------------------------
        If GastoVto > 0 And Not Agrupa4311x And Not Agrupa431x Then
            'Err.Raise 513, , "Error en codigo. Parametro tipo remesa incorrecto. "
           'Lineas de apuntes .
            SQL = LinApu & Linea & ",'"
    
    
            SQL = SQL & CtaBancoGastos & "','" & Rs!NUmSerie & Format(Rs!NumFac, "0000000") & "'," & vCP.condecli
            SQL = SQL & ",'Gastos vto.'"
            
            
            'Importe al debe
            SQL = SQL & "," & TransformaComasPuntos(CStr(GastoVto)) & ",NULL,"
            
            If CCBanco <> "" Then
                SQL = SQL & "'" & DevNombreSQL(CCBanco) & "'"
            Else
                SQL = SQL & "NULL"
            End If
                
            'Contra partida
            'Si no lleva cuenta puente contabiliza los gastos
            Aux = "NULL"
           
            SQL = SQL & "," & Aux & ",'COBROS',0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
        
        
        End If
        
        
        
        If CuentaPuente2 Then
            
            i = Len(SubCtaPte2)
            If i = vEmpresa.DigitosUltimoNivel Then
                SQL = SubCtaPte2
            Else
                SQL = SubCtaPte2 & Mid(Rs!Cta, i + 1)
            End If
            SQL = LinApu & Linea & ",'" & SQL
            SQL = SQL & "','" & Rs!NUmSerie & Format(Rs!NumFac, "0000000") & "'," & vCP.condecli
    
            Ampliacion = Amp11 & " "
        
            'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
            Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFac, "0000000") ' & "/" & RS!NumFac
            
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
    
            
    
    
    
            Importeauxiliar = Rs!Imponible - GastoVto
            SQL = SQL & "NULL," & TransformaComasPuntos(CCur(Importeauxiliar)) & ",NULL,"
    
            
            SQL = SQL & "'" & CtaEfectosComDescontados & "',"
            SQL = SQL & "'COBROS',0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
            If Not Ejecuta(SQL) Then Exit Function
        
            Linea = Linea + 1
    
            SQL = LinApu & Linea & ",'" & CtaEfectosComDescontados
            SQL = SQL & "','" & Rs!NUmSerie & Format(Rs!NumFac, "0000000") & "'," & vCP.condecli
    
            Ampliacion = Amp11 & " "
        
            'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
            Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFac, "0000000") ' & "/" & RS!NumFac
            
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
    
            
    
    
    
            
            SQL = SQL & TransformaComasPuntos(CCur(Importeauxiliar)) & ",NULL,NULL,"
            
            If i = vEmpresa.DigitosUltimoNivel Then
                Ampliacion = SubCtaPte2
            Else
                Ampliacion = SubCtaPte2 & Mid(Rs!Cta, i + 1)
            End If
            
            SQL = SQL & "'" & Ampliacion & "',"
            SQL = SQL & "'COBROS',0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
            If Not Ejecuta(SQL) Then Exit Function
        
            Linea = Linea + 1
        
        
        End If
        Rs.MoveNext
    Wend
    
    
    Rs.MoveFirst



    'Linea de los gastos de cada RECIBO
    'Gastos de los recibos.
    'Si tiene alguno de los efectos remesados gastos
    If Gastos > 0 And (Agrupa4311x Or Agrupa431x) Then
        
        Err.Raise 513, , "Error en codigo. Parametro tipo remesa incorrecto. "
        
        
        If CtaBancoGastos = "" Then CtaBancoGastos = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1")
        
        Aux = "RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2)
        
        SQL = LinApu & Linea & ",'"
        SQL = SQL & CtaBancoGastos & "','" & Aux & "'," & vCP.condecli
        SQL = SQL & ",'Gastos vtos. " & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2) '"
        
        
        'Importe al debe
        SQL = SQL & "'," & TransformaComasPuntos(CStr(Gastos)) & ",NULL,"
        
        If CCBanco <> "" Then
            SQL = SQL & "'" & DevNombreSQL(CCBanco) & "'"
        Else
            SQL = SQL & "NULL"
        End If
            
        'Contra partida
        SQL = SQL & ",NULL,'COBROS',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
        
        
        If Not Ejecuta(SQL) Then Exit Function
        
        Linea = Linea + 1
    
    End If

    'La linea del banco
    '*********************************************************************
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","

    'Nuevo tipo ampliacion
    If vCP.amphacli = 3 Then
        Ampliacion = DescRemesa
    Else
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
    End If
    
    Ampliacion = Ampliacion & " Dev.rem:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
    
    Amp11 = Rs!Cliente  'cta banco

    'Lleva gasto pero lo descontamos de aqui
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Importe = Importe + GastoDevolucion
        'Para que la linea salga al fina
        Linea = Linea + 2
    End If
    Ampliacion = Linea & ",'" & Amp11 & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
    Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,"
  '  If CuentaPuente2 Then
  '      St op
  '      Ampliacion = Ampliacion & "'" & SubCtaPte2 & "'"
  '  Else
   
        Ampliacion = Ampliacion & "NULL"
  '  End If
    Ampliacion = Ampliacion & ",'COBROS',0)"
    Ampliacion = SQL & Ampliacion
    If Not Ejecuta(Ampliacion) Then GoTo ECon
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Linea = Linea - 2
        'Dejo el importe como estaba
        Importe = Importe - GastoDevolucion
    Else
        Linea = Linea + 1
    End If
    
    
    'SI hay que contabilizar los gastos de devolucion
    If ContabilizoGastoBanco Then
        
         If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
         Else
            Linea = Linea + 1
         End If
         SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
         SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
         SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
         SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"

         'Cuenta
         SQL = SQL & CtaBenBancarios & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.condecli
         'SQL = SQL & Rs!Cta & "','REM" & Format(Rs!numfac, "000000000") & "'," & vCP.condecli
        

        If vCP.ampdecli = 3 Then
            Ampliacion = DescRemesa
        Else
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"


        ' timporteH, codccost, ctacontr, idcontab, punteada
        'Importe.  Va al debe
        SQL = SQL & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,"
        
        'Centro de coste.
        '--------------------------
        Amp11 = "NULL"
        If vParam.autocoste Then
            Amp11 = DevuelveDesdeBD("codccost", "bancos", "codmacta", Rs!Cliente, "T")
            Amp11 = "'" & Amp11 & "'"
        End If
        SQL = SQL & Amp11 & ","

        
        SQL = SQL & "'" & Rs!Cliente & "',"
        SQL = SQL & "'COBROS',0)"

        If Not Ejecuta(SQL) Then GoTo ECon

        
        
    
        'El total del banco..
        
        'La linea del banco
        'Rs.MoveFirst
        'Si no agrupa dto importe
        If Not DescuentaImporteDevolucion Then
            Linea = Linea + 1
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
        
            
            If vCP.amphacli = 3 Then
                Ampliacion = DescRemesa
            Else
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
                Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
            End If
            
            Ampliacion = Linea & ",'" & Rs!Cliente & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
            'Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,'" & CtaBenBancarios & "','CONTAB',0)"
            Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,'" & CtaBenBancarios & "','COBROS',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then GoTo ECon
            
        End If
      
    
    End If

    'Ya tenemos generado el apunte de devolucion
    'Insertamos para su actualziacion
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaDevolucion
    
    
    RealizarDevolucionRemesaEfectos = True
ECon:
    If Err.Number <> 0 Then
        
        Amp11 = "Devolución remesa: " & Remesa & vbCrLf
        If Not Mc Is Nothing Then Amp11 = Amp11 & "MC.cont: " & Mc.Contador & vbCrLf
        Amp11 = Amp11 & Err.Description
        MuestraError Err.Number, Amp11
        
    End If
    Set Rs = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function


'*************************************************************************************
Public Function RealizarDevolucionRemesaTalPag(FechaDevolucion As Date, ContabilizoGastoBanco As Boolean, CtaBenBancarios As String, Remesa As String, DatosContabilizacionDevolucion As String) As Boolean

'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim Amp11 As String
Dim DescRemesa As String
Dim CuentaPuente As Boolean
Dim TipoRemesa2 As Byte
Dim SubCtaPte As String
'Dim AgrupaApunteBanco As Boolean
Dim GastoDevolucion As Currency
Dim DescuentaImporteDevolucion As Boolean
Dim GastoVto As Currency
Dim Gastos As Currency  'de cada recibo/vto
Dim Aux As String
Dim Importeauxiliar As Currency
Dim CtaBancoGastos As String
Dim CCBanco As String
Dim CtaEfectosComDescontados As String   '   tipo FONTENAS
Dim LinApu As String

Dim Obs As String

    On Error GoTo ECon
    RealizarDevolucionRemesaTalPag = False
    
    
    'La forma de pago
    Set vCP = New Ctipoformapago
    
    
    'Leo la descipcion de la remesa si alguna de las ampliaciones me la solicita
    SQL = "Select descripcion,tiporem from remesas where codigo =" & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyo =" & RecuperaValor(Remesa, 2)
    
    DescRemesa = "Remesa: " & RecuperaValor(Remesa, 1) & " / " & RecuperaValor(Remesa, 2)
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TipoRemesa2 = Rs!Tiporem
    If Not IsNull(Rs.Fields(0)) Then DescRemesa = DevNombreSQL(Rs.Fields(0))
    Rs.Close
    
    CuentaPuente = False
    
    
    If TipoRemesa2 = 2 Then
        Linea = vbPagare
        SQL = "pagarecta"
        CuentaPuente = vParamT.PagaresCtaPuente
        
    ElseIf TipoRemesa2 = 5 Then
        Linea = vbConfirming
        SQL = "confirmingcta"
        CuentaPuente = vParamT.ConfirmingCtaPuente
    
    Else
        Linea = vbTalon
        SQL = "taloncta"
        CuentaPuente = vParamT.TalonesCtaPuente
    End If

    If CuentaPuente Then
     
        SubCtaPte = DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1")
             
        If SubCtaPte = "" Then
            MsgBox "Falta por configurar en parametros", vbExclamation
            Exit Function
           
        End If
    End If

    
    If vCP.Leer(Linea) = 1 Then GoTo ECon


    'Los parametros de contbilizacion se le pasan en el frame de pedida de datos
    'Ahora se los asignaremos a la formma de pago
    vCP.condecli = RecuperaValor(DatosContabilizacionDevolucion, 1)
    vCP.ampdecli = RecuperaValor(DatosContabilizacionDevolucion, 2)
    vCP.conhacli = RecuperaValor(DatosContabilizacionDevolucion, 1)
    vCP.amphacli = RecuperaValor(DatosContabilizacionDevolucion, 2)
    
    
    
    
    SQL = RecuperaValor(Remesa, 7)
    GastoDevolucion = TextoAimporte(SQL)
    DescuentaImporteDevolucion = False
    If GastoDevolucion > 0 Then
        CtaBancoGastos = "CtaIngresos"
        SQL = RecuperaValor(Remesa, 3)
        SQL = DevuelveDesdeBD("GastRemDescontad", "bancos", "codmacta", SQL, "T")
        If SQL = "1" Then DescuentaImporteDevolucion = True
    End If
    
    'Datos del banco
    SQL = RecuperaValor(Remesa, 3)
    SQL = "Select * from bancos where codmacta ='" & SQL & "'"
    CCBanco = ""
    CtaBancoGastos = ""
    CtaEfectosComDescontados = ""
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        SQL = "No se ha encontrado banco: " & vbCrLf & SQL
        Err.Raise 516, SQL
    End If
    CCBanco = DBLet(Rs!CodCcost, "T")
    CtaBancoGastos = DBLet(Rs!ctagastos, "T")
    If Not vParam.autocoste Then CCBanco = ""  'NO lleva analitica
    Rs.Close
    

    CtaEfectosComDescontados = ""


    
    'EMPEZAMOS
    'Borramos tmpactualizar
    SQL = "DELETE FROM tmpactualizar where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    'Cargaremos los registros a devolver que estaran en la tabla temporal
    'para codusu
    SQL = "Select * from tmpfaclin where codusu=" & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "EOF.  NO se han cargado datos devolucion", vbExclamation
        Rs.Close
        GoTo ECon
    End If

    Set Mc = New Contadores


    If Mc.ConseguirContador("0", FechaDevolucion <= vParam.fechafin, True) = 1 Then GoTo ECon


    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Devolucion remesa(T/P): " & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2)
    SQL = SQL & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "',"
    
    
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Devolución remesa(T/P)" & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2) & "')"
    
    
    If Not Ejecuta(SQL) Then GoTo ECon


    Linea = 1
    Importe = 0

    If vCP.ampdecli = 3 Then
        Amp11 = DescRemesa
    Else
        Amp11 = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    End If
    
    'Lo meto en una VAR
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,  numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem,esdevolucion) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
    LinApu = SQL
    
    While Not Rs.EOF

        'Lineas de apuntes .
         SQL = LinApu & Linea & ",'"
         SQL = SQL & Rs!Cta
         SQL = SQL & "','" & Format(Rs!NumFac, "0000000") & "'," & vCP.condecli

        Ampliacion = Amp11 & " "
    
        If vCP.ampdecli = 3 Then
            'NUEVA forma de ampliacion
            'No hacemos nada pq amp11 ya lleva lo solicitado
            
        Else
            If vCP.ampdecli = 4 Then
                'COntrapartida
                Ampliacion = Ampliacion & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!Cta, "T")
                
            Else
                If vCP.ampdecli = 2 Then
                   Ampliacion = Ampliacion & Format(Rs!Fecha, "dd/mm/yyyy")
                Else
                   If vCP.ampdecli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                   'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                   Ampliacion = Ampliacion & Rs!IVA & "/" & Rs!NumFac
                   
                End If
            End If
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"

        Importe = Importe + Rs!Imponible


        GastoVto = 0
        Aux = " numserie='" & Rs!IVA & "' AND numfactu=" & Rs!NumFac
        Aux = Aux & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden"
        Aux = DevuelveDesdeBD("gastos", "cobros", Aux, CStr(Rs!NIF), "N")
        
        If Aux <> "" Then GastoVto = CCur(Aux)
        Gastos = Gastos + GastoVto

        ' timporteH, codccost, ctacontr, idcontab, punteada
        Importeauxiliar = Rs!Imponible - GastoVto
        SQL = SQL & TransformaComasPuntos(CCur(Importeauxiliar)) & ",NULL,NULL,"

        If vCP.ctrdecli = 1 Then
            If CuentaPuente Then
                If Len(SubCtaPte) = vEmpresa.DigitosUltimoNivel Then
                    SQL = SQL & "'" & SubCtaPte & "',"
                Else
                    SQL = SQL & "'" & SubCtaPte & Mid(Rs!Cta, Len(SubCtaPte) + 1) & "',"
                End If
            Else
                SQL = SQL & "'" & Rs!Cliente & "',"
            End If
        Else
            SQL = SQL & "NULL,"
        End If
        SQL = SQL & "'COBROS',0,"
        
        '%%%%% aqui van todos los datos de la devolucion en la linea de cuenta
        SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFac, "N") & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Rs!NIF, "N") & ","

         '-------------------------------------------------------------------------------------
         'Ahora
         '-------------------------------------------------------------------------------------
         'Lo pongo en uno
             'Actualizamos el registro. Ponemos la marca de devuelto. Y aumentamos el importe de gastos
         'Si es que hay
         Dim SqlCobro As String
         Dim RsCobro As ADODB.Recordset
         Dim ImporteNue As Currency
         
         SqlCobro = "select tipforpa, tiporem, codrem, anyorem, gastos from cobros inner join formapago on cobros.codforpa = formapago.codforpa "
         SqlCobro = SqlCobro & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
         SqlCobro = SqlCobro & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
         
         Set RsCobro = New ADODB.Recordset
         RsCobro.Open SqlCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         If Not RsCobro.EOF Then
         
'    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem) "
            SQL = SQL & DBSet(RsCobro!TipForpa, "N") & "," & DBSet(FechaDevolucion, "F") & "," & DBSet(Rs!CtaBase, "T", "S") & ","
            SQL = SQL & DBSet(Rs!ImpIva, "N") & "," & DBSet(RsCobro!Tiporem, "N") & "," & DBSet(RsCobro!Codrem, "N") & "," & DBSet(RsCobro!Anyorem, "N") & ",1)"
              
         
            Ampliacion = "UPDATE cobros SET "
            Ampliacion = Ampliacion & " Devuelto = 1, situacion = 0   "
            ImporteNue = Rs!Total - Rs!Imponible '- Rs!impiva
            
            ImporteNue = DBLet(RsCobro!Gastos, "N")
            If DBLet(Rs!ImpIva, "N") > 0 Then
            
                If ImporteNue = 0 Then
                    Ampliacion = Ampliacion & " , Gastos = " & TransformaComasPuntos(CStr(Rs!ImpIva))
                Else
                    Ampliacion = Ampliacion & " , Gastos = Gastos + " & TransformaComasPuntos(CStr(Rs!ImpIva))
                End If
            End If
            Ampliacion = Ampliacion & " ,impcobro=NULL,codrem= NULL, anyorem = NULL, siturem = NULL,tiporem=NULL,fecultco=NULL,recedocu=0"
            Ampliacion = Ampliacion & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
            Ampliacion = Ampliacion & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
            
            If Not Ejecuta(Ampliacion) Then GoTo ECon
             
         End If
         Set RsCobro = Nothing

        '%%%%% hasta aqui

        If Not Ejecuta(SQL) Then GoTo ECon

        Linea = Linea + 1
        
 
        'Lineas de apuntes del GASTO del vto en curso
        SQL = LinApu & Linea & ",'"


        SQL = SQL & CtaBancoGastos & "','" & Format(Rs!NumFac, "000000000") & "'," & vCP.condecli
        SQL = SQL & ",'Gastos vto.'"
        
        
        'Importe al debe
        SQL = SQL & "," & TransformaComasPuntos(CStr(GastoVto)) & ",NULL,"
        
        If CCBanco <> "" Then
            SQL = SQL & "'" & DevNombreSQL(CCBanco) & "'"
        Else
            SQL = SQL & "NULL"
        End If
            
        'Contra partida
        'Si no lleva cuenta puente contabiliza los gastos
        Aux = "NULL"
       
        SQL = SQL & "," & Aux & ",'COBROS',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
        If GastoVto <> 0 Then
            If Not Ejecuta(SQL) Then Exit Function
        
            Linea = Linea + 1
        End If

        
        'Si tiene cuenta puente cancelo la puente tb
        If CuentaPuente Then
                
            'Si lleva cta efectos comerciales descontados, tipo fontenas, NO HACE este contrapunte
            If CtaEfectosComDescontados = "" Then
                'Lineas de apuntes .
                 SQL = LinApu & Linea & ",'"
              
                 If Len(SubCtaPte) = vEmpresa.DigitosUltimoNivel Then
                     SQL = SQL & SubCtaPte
                 Else
                     SQL = SQL & SubCtaPte & Mid(Rs!Cta, Len(SubCtaPte) + 1)
                 End If
                 SQL = SQL & "','" & Format(Rs!NumFac, "0000000") & "'," & vCP.conhacli
    
                
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli) & " "
            
                If vCP.amphacli = 3 Then
                    'NUEVA forma de ampliacion
                    'No hacemos nada pq amp11 ya lleva lo solicitado
                    
                Else
                    If vCP.amphacli = 4 Then
                        'COntrapartida
                        Ampliacion = Ampliacion & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!Cta, "T")
                        
                    Else
                        If vCP.amphacli = 2 Then
                           Ampliacion = Ampliacion & Format(Rs!Fecha, "dd/mm/yyyy")
                        Else
                           If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                           'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                           Ampliacion = Ampliacion & Rs!IVA & "/" & Rs!NumFac
                           
                        End If
                    End If
                End If
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',NULL,"
        
                SQL = SQL & TransformaComasPuntos(Rs!Imponible) & ",NULL,"
        
                If vCP.ctrhacli = 1 Then
                    SQL = SQL & "'" & Rs!Cta & "'"
                Else
                    SQL = SQL & "NULL"
                End If
                SQL = SQL & ",'COBROS',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
            
                            
                If Not Ejecuta(SQL) Then GoTo ECon
                Linea = Linea + 1
            End If 'de eefctosdescontados=""
        End If 'de ctapte
            
        Rs.MoveNext
    Wend
    
    
    Rs.MoveFirst









    If CuentaPuente Then
        SubCtaPte = Rs!Cliente
        SubCtaPte = DevuelveDesdeBD("ctaefectosdesc", "bancos", "codmacta", SubCtaPte, "T")
        If SubCtaPte = "" Then
            MsgBox "Cuenta efectos descontados erronea. Revisar apunte " & Mc.Contador, vbExclamation
            SubCtaPte = Rs!Cliente
        End If
    End If
    
    'La linea del banco
    '*********************************************************************
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","

    'Nuevo tipo ampliacion
    If vCP.amphacli = 3 Then
        Ampliacion = DescRemesa
    Else
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
    End If
    
    Ampliacion = Ampliacion & " Dev.rem:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
    
    Amp11 = Rs!Cliente  'cta banco

    'Lleva gasto pero lo descontamos de aqui
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Importe = Importe + GastoDevolucion
        'Para que la linea salga al fina
        Linea = Linea + 2
    End If
    Ampliacion = Linea & ",'" & Amp11 & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
    Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,"
    If CuentaPuente Then
        Ampliacion = Ampliacion & "'" & SubCtaPte & "'"
    Else
        'Nulo
        Ampliacion = Ampliacion & "NULL"
    End If
    Ampliacion = Ampliacion & ",'COBROS',0)"
    Ampliacion = SQL & Ampliacion
    If Not Ejecuta(Ampliacion) Then GoTo ECon
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Linea = Linea - 2
        'Dejo el importe como estaba
        Importe = Importe - GastoDevolucion
    Else
        Linea = Linea + 1
    End If
    If CuentaPuente Then
        'EL ANTERIOR contrapuenteado
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
    
        'Nuevo tipo ampliacion
        If vCP.ampdecli = 3 Then
            Ampliacion = DescRemesa
        Else
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
        End If
        
        Ampliacion = Ampliacion & " Dev.rem:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
        
        
        Amp11 = SubCtaPte  'cta efectos dtos
        
        Ampliacion = Linea & ",'" & Amp11 & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.condecli & ",'" & Ampliacion & "',"
        Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL,"
        'Cta efectos descontados
        Ampliacion = Ampliacion & "'" & Rs!Cliente & "'"

        Ampliacion = Ampliacion & ",'COBROS',0)"
        Ampliacion = SQL & Ampliacion
        If Not Ejecuta(Ampliacion) Then GoTo ECon
        Linea = Linea + 1
  
    End If
    
    
    'SI hay que contabilizar los gastos de devolucion
    If ContabilizoGastoBanco Then
        
             If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
             
             Else
                Linea = Linea + 1
             End If
             SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
             SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
             SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
    
             'Cuenta
             SQL = SQL & CtaBenBancarios & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.condecli
             'SQL = SQL & Rs!Cta & "','REM" & Format(Rs!numfac, "000000000") & "'," & vCP.condecli
            
    
            If vCP.ampdecli = 3 Then
                Ampliacion = DescRemesa
            Else
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
                Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
            End If
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
    
    
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe.  Va al debe
            SQL = SQL & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,"
            
            'Centro de coste.
            '--------------------------
            Amp11 = "NULL"
            If vParam.autocoste Then
                Amp11 = DevuelveDesdeBD("codccost", "bancos", "codmacta", Rs!Cliente, "T")
                Amp11 = "'" & Amp11 & "'"
            End If
            SQL = SQL & Amp11 & ","
    
            
            SQL = SQL & "'" & Rs!Cliente & "',"
            SQL = SQL & "'COBROS',0)"
    
            If Not Ejecuta(SQL) Then GoTo ECon
    
            
            
  
            'Si no agrupa dto importe
            If Not DescuentaImporteDevolucion Then
                Linea = Linea + 1
                SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
                SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
                SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
                SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
            
                
                If vCP.amphacli = 3 Then
                    Ampliacion = DescRemesa
                Else
                    Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
                    Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
                End If
                
                Ampliacion = Linea & ",'" & Rs!Cliente & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
                'Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,'" & CtaBenBancarios & "','CONTAB',0)"
                Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,'" & CtaBenBancarios & "','COBROS',0)"
                Ampliacion = SQL & Ampliacion
                If Not Ejecuta(Ampliacion) Then GoTo ECon
                
            End If
      
    
    End If

    'Ya tenemos generado el apunte de devolucion
    'Insertamos para su actualziacion
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaDevolucion
    
    
    
    'Cerramos RS
    Rs.Close
    Set miRsAux = Nothing
    
    RealizarDevolucionRemesaTalPag = True
ECon:
    If Err.Number <> 0 Then
        
        Amp11 = "Devolución remesa: " & Remesa & vbCrLf
        If Not Mc Is Nothing Then Amp11 = Amp11 & "MC.cont: " & Mc.Contador & vbCrLf
        Amp11 = Amp11 & Err.Description
        MuestraError Err.Number, Amp11
        
    End If
    Set Rs = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function








'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'
'   COMPENSACIONES.
'       Contabilizara las compensaciones. Es decir. Desde el FORM de las compensaciones
'       le mandara el conjunto de cobros, el de pagos
'       cta bancaria
'
'       Y generara un UNICO apunte eliminando todos los cobros y pagos seleccionados
'       excepto si la compensacion se efectua sobre un determinado VTO
'       que sera updateado
'       Si AumentaElImporteDelVto significa que el vto aumenta ;)
Public Function ContabilizarCompensaciones(ByRef ColCobros As Collection, ByRef ColPagos As Collection, ByVal DatosAdicionales As String, AumentaElImporteDelVto As Boolean) As Boolean
Dim SQL As String
Dim Mc As Contadores
Dim CadenaSQL As String
Dim FechaContab As Date
Dim i As Integer
Dim Obs As String

Dim SqlNue As String
Dim RsNue As ADODB.Recordset



    On Error GoTo EEContabilizarCompensaciones

    ContabilizarCompensaciones = False
        
    
    'Fecha contabilizacion
    FechaContab = RecuperaValor(DatosAdicionales, 4)
    
    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Conn.Execute SQL


    Conn.BeginTrans    'TRANSACCION
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", FechaContab <= vParam.fechafin, True) = 1 Then GoTo EEContabilizarCompensaciones
        
        
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & CInt(RecuperaValor(DatosAdicionales, 3)) & ",'" & Format(FechaContab, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Compensa: " & DevNombreSQL(RecuperaValor(DatosAdicionales, 7)) & vbCrLf
    If AumentaElImporteDelVto Then SQL = SQL & "Aumento VTO" & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:nn") & " por " & vUsu.Nombre & "',"
    
    Obs = "ARICONTA 6: Compensa: " & RecuperaValor(DatosAdicionales, 7)
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Obs, "T") & ");"
    Conn.Execute SQL

    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, RecuperaValor(DatosAdicionales, 3), FechaContab
    
    

    'Añadimos las facturas de clientes
    'Lineas de apuntes .
    CadenaSQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    CadenaSQL = CadenaSQL & "codconce, numdocum, ampconce , "
    'Toda esta linea viene juntita
    CadenaSQL = CadenaSQL & "codmacta, timporteD,timporteH,"
    'Numdocum viene con otro valor
    CadenaSQL = CadenaSQL & " ctacontr, codccost, idcontab, punteada, "
    CadenaSQL = CadenaSQL & " numserie, numfaccl, numfacpr, fecfactu, numorden, tipforpa) "
    CadenaSQL = CadenaSQL & "VALUES (" & RecuperaValor(DatosAdicionales, 3) & ",'" & Format(FechaContab, FormatoFecha) & "'," & Mc.Contador & ","
    

    NumRegElim = 1
    'Los cobros
    For i = 1 To ColCobros.Count
        
        
        SQL = NumRegElim & "," & RecuperaValor(ColCobros.Item(i), 1) & "NULL,'COBROS',0,"
        
        'parte donde indicamos en el apunte que se ha cobrado
        SqlNue = "select * from cobros " & RecuperaValor(ColCobros.Item(i), 3)
        Set RsNue = New ADODB.Recordset
        RsNue.Open SqlNue, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsNue.EOF Then
            SQL = SQL & DBSet(RsNue!NUmSerie, "T") & ","
            SQL = SQL & DBSet(RsNue!NumFactu, "N") & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & DBSet(RsNue!FecFactu, "F") & ","
            SQL = SQL & DBSet(RsNue!numorden, "N") & ","
            SQL = SQL & DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(RsNue!Codforpa, "N")) & ")"
            
            
        Else
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ")"
        End If
        
        
        
        
        
        Set RsNue = Nothing
        
        Conn.Execute CadenaSQL & SQL
        
        
        NumRegElim = NumRegElim + 1
        'Borro el cobro pago
        SQL = RecuperaValor(ColCobros.Item(i), 2)
        If Mid(SQL, 1, 6) = "UPDATE" Then
            'UPDATEAMOS
            Conn.Execute SQL
        Else
            ' ATENCION !!!!!
            ' ya no borramos hemos de darlo como cobrado
'            Conn.Execute "DELETE FROM cobros " & Sql
'
'            'Borramos de efectos devueltos... por si acaso sefecdev
'            Ejecuta "DELETE FROM sefecdev " & Sql
            SqlNue = "update cobros set fecultco = " & DBSet(FechaContab, "F") & ", impcobro = coalesce(impcobro,0) + impvenci + coalesce(gastos,0), situacion = 1 "
            SqlNue = SqlNue & SQL

            Ejecuta SqlNue
        End If
    
    Next i

    
    
    'Los pagos
    For i = 1 To ColPagos.Count
        SQL = NumRegElim & "," & RecuperaValor(ColPagos.Item(i), 1) & "NULL,'PAGOS',0,"
        
        'parte donde indicamos en el apunte que se ha pagado
        SqlNue = "select * from pagos " & RecuperaValor(ColPagos.Item(i), 3)
        Set RsNue = New ADODB.Recordset
        RsNue.Open SqlNue, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsNue.EOF Then
            SQL = SQL & DBSet(RsNue!NUmSerie, "T") & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & DBSet(RsNue!NumFactu, "T") & ","
            SQL = SQL & DBSet(RsNue!FecFactu, "F") & ","
            SQL = SQL & DBSet(RsNue!numorden, "N") & ","
            SQL = SQL & DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(RsNue!Codforpa, "N")) & ")"
        Else
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ")"
        End If
        Set RsNue = Nothing
        
        
        Conn.Execute CadenaSQL & SQL
        NumRegElim = NumRegElim + 1
        'Borro el  pago   La linea del banco va aqui dentro, con lo cual
        'Si tengo que comprobar si es la linea del banco o no para borrar
        SQL = RecuperaValor(ColPagos.Item(i), 2)
        If SQL <> "" Then
            If Mid(SQL, 1, 6) = "UPDATE" Then
                'UPDATEAMOS
                Conn.Execute SQL
            Else
                ' ATENCION !!!!!
                ' ya no borramos hemos de darlo como pagado
'                Conn.Execute "DELETE FROM pagos " & Sql
            
                SqlNue = "update pagos set fecultpa = " & DBSet(FechaContab, "F") & ",imppagad = coalesce(imppagad,0) + impefect, situacion = 1 "
                SqlNue = SqlNue & SQL
    
                Ejecuta SqlNue
            End If

        End If
    Next i

    Conn.CommitTrans   'TODO HA IDO BIEN
    

        
    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Ejecuta SQL
    
    ContabilizarCompensaciones = True
    Exit Function
EEContabilizarCompensaciones:
    If Err.Number <> 0 Then MuestraError Err.Number
    Conn.RollbackTrans
    
End Function





'----------------------------------------------------------------------------------------------------
' NORMA 32,58, Febrero 2009: TOoooodas las remesas
' ================================================
'
'
'Mod Nov 2009


'*********************************************************************************
'
'   TALONES / PAGARES
'
'*********************************************************************************
'*********************************************************************************
'
'
'   LaOpcion:   0. Cancelar cliente
'
'   Mayo 2009.  Cambio.  La cancelacion la realiza en la recepcion de documentos
'
'DiarioConcepto:Llevara el diario y los conceptos al debe y al haber. NO cojera los de la stipforpa, si no de una window anterior
'              El cuarto pipe que lleva es si agrrupa en la cuenta puente
'                   es decir, en lugar de 43.1 contra 431.1
'                                         43.2 contra 431.1
'                           hacemos   43.1 y 43.2   contra la suma en 431.1
' Septiembre 2009
'           El quinto y sexto pipe llevaran, si necesario, cta dodne poner el benefic po perd del talon y si requiere cc

'### Noviembre 2014
' Si es contra una unica cuenta puente de talon / pagare, entonces para
' el concepto del esta pondremos:
'       o la contrapartida(nomacta)
'       o como esta, el numero de talon pagare


' 0 Pagare  1talon  2 Confirming
Public Function RemesasCancelacionTALONPAGARE(TalonTipoDoc As Byte, IdRecepcion As Integer, FechaAbono As Date, DiarioConcepto As String) As Boolean
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
'Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacion As String
Dim CuentaUnica As Boolean
Dim LCta As Integer
Dim Importeauxiliar As Currency
Dim AgrupaVtosPuente As Boolean
Dim CadenaAgrupaVtoPuente As String
Dim aux2 As String
Dim RequiereCCDiferencia As Boolean

Dim Obs As String
Dim TipForpa As Byte

    On Error GoTo ERemesa_CancelacionCliente2
    RemesasCancelacionTALONPAGARE = False
    

    If TalonTipoDoc = 1 Then
        'Sobre talones
        Cuenta = "taloncta"
    ElseIf TalonTipoDoc = 2 Then
        Cuenta = "confirmingcta"
    Else
        Cuenta = "pagarecta"
    End If
    RaizCuentasCancelacion = DevuelveDesdeBD(Cuenta, "paramtesor", "codigo", "1", "N")
    If RaizCuentasCancelacion = "" Then
        MsgBox "Error grave en configuración de  parámetros de tesorería. Falta cuenta cancelación", vbExclamation
        Exit Function
    End If
    
    LCta = Len(RaizCuentasCancelacion)
    CuentaUnica = LCta = vEmpresa.DigitosUltimoNivel
    
    
    'Comprobacion.  Para todos los efectos de la 43.... se cancelan con la 4310....
    '
    'Tendre que ver que existen estas cuentas
    Set Rs = New ADODB.Recordset
    
    SQL = "DELETE FROM tmpcierre1 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
        
    If Not CuentaUnica Then
        'Cancela contra subcuentas de cliente
        

        Ampliacion = ",CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")" & ")"
            
        SQL = "Select " & vUsu.Codigo & Ampliacion & " from talones where codigo=" & IdRecepcion
        SQL = SQL & " GROUP BY codmacta"
        'INSERT
        SQL = "INSERT INTO tmpcierre1(codusu,cta) " & SQL
        Conn.Execute SQL
        
        'Ahora monto el select para ver que cuentas 430 no tienen la 4310
        SQL = "Select cta,codmacta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " HAVING codmacta is null"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        Linea = 0
        While Not Rs.EOF
            Linea = Linea + 1
            SQL = SQL & Rs!Cta & "     "
            If Linea = 5 Then
                SQL = SQL & vbCrLf
                Linea = 0
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        
        If SQL <> "" Then
            
            AmpRemesa = "CANCELACION remesa"
            
            SQL = "Cuentas " & AmpRemesa & ".  No existen las cuentas: " & vbCrLf & String(90, "-") & vbCrLf & SQL
            SQL = SQL & vbCrLf & "¿Desea crearlas?"
            Linea = 1
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                'Ha dicho que si desea crearlas
                
                Ampliacion = "CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")) "
            
                SQL = "Select codmacta," & Ampliacion & " from talones where codigo=" & IdRecepcion
                SQL = SQL & " and " & Ampliacion & " in "
                SQL = SQL & "(Select cta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
                SQL = SQL & " AND codmacta is null)"
                Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                
                     SQL = "INSERT  IGNORE INTO cuentas(codmacta,nommacta ,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) SELECT '"
                                ' CUenta puente
                     SQL = SQL & Rs.Fields(1) & "',nommacta ,'S',0,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from cuentas where codmacta = '"
                                'Cuenta en la scbro (codmacta)
                     SQL = SQL & Rs.Fields(0) & "'"
                     Conn.Execute SQL
                     Rs.MoveNext
                     
                Wend
                Rs.Close
                Linea = 0
            End If
            If Linea = 1 Then GoTo ERemesa_CancelacionCliente2
        End If
        
    Else
        'Cancela contra UNA unica cuenta todos los vencimientos
        SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", RaizCuentasCancelacion, "T")
        If SQL = "" Then
            MsgBox "No existe la cuenta puente: " & RaizCuentasCancelacion, vbExclamation
            GoTo ERemesa_CancelacionCliente2
        End If
    End If

    
    'La forma de pago
    Set vCP = New Ctipoformapago
    If TalonTipoDoc = 1 Then
        SQL = CStr(vbTalon)
        Ampliacion = "Talones"
    ElseIf TalonTipoDoc = 2 Then
        SQL = CStr(vbConfirming)
        Ampliacion = "Confirming"
    Else
        SQL = CStr(vbPagare)
        Ampliacion = "Pagarés"
    End If
    If vCP.Leer(CInt(SQL)) = 1 Then GoTo ERemesa_CancelacionCliente2
    'Ahora fijo los valores que me ha dado el
    vCP.diaricli = RecuperaValor(DiarioConcepto, 1)
    vCP.condecli = RecuperaValor(DiarioConcepto, 2)
    vCP.conhacli = RecuperaValor(DiarioConcepto, 3)
    AgrupaVtosPuente = RecuperaValor(DiarioConcepto, 4) = 1
 '   AgrupaVtosPuente = AgrupaVtosPuente 'And CuentaUnica
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Cancelacion cliente"

    SQL = SQL & " NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & "',"
    
    Obs = "ARICONTA 6: Cancelacion cliente NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Obs, "T") & ") ;"
    
    
    
    If Not Ejecuta(SQL) Then Exit Function
    
    
    
    
    Linea = 1
    Importe = 0
    'Gastos = 0
    
    vCP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)  'DEBE
    vCP.CadenaAuxiliar = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)   'DEBE
    
    
    SQL = "select cobros.*,l.importe,l.codigo,c.numeroref reftalonpag,c.banco  from (talones c inner join talones_facturas l on c.codigo = l.codigo) left join  cobros  on l.numserie=cobros.numserie and"
    SQL = SQL & " l.numfactu=cobros.numfactu and   l.fecfactu=cobros.fecfactu and l.numorden=cobros.numorden"
    SQL = SQL & " WHERE c.codigo= " & IdRecepcion
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Trozo comun
    AmpRemesa = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
    AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada, "
    AmpRemesa = AmpRemesa & " numserie, numfaccl, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
    AmpRemesa = AmpRemesa & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
    
    CadenaAgrupaVtoPuente = ""

    While Not Rs.EOF
        Ampliacion = Rs!NUmSerie   'SI DA ERROR ES QUE NO EXISTE mediante el left join
        
        
        
        'Neuvo dato para la ampliacion en la contabilizacion
        Ampliacion = " "
        Select Case vCP.amphacli
        Case 2, 4
            'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
            Ampliacion = DBLet(Rs!reftalonpag, "T")
            If Ampliacion = "" Then Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
        Case 5
            Ampliacion = DBLet(Rs!reftalonpag, "T")
            If Ampliacion = "" Then
                Ampliacion = Rs!NUmSerie & "/" & Rs!NumFactu
            Else
                Ampliacion = "Doc: " & Ampliacion
            End If
        Case Else
           If vCP.amphacli = 1 Then Ampliacion = vCP.siglas & " "
           Ampliacion = Ampliacion & Rs!NUmSerie & "/" & Rs!NumFactu
        End Select
        If Mid(Ampliacion, 1, 1) <> " " Then Ampliacion = " " & Ampliacion
        
        'Cancelacion
        If CuentaUnica Then
            Cuenta = RaizCuentasCancelacion
        Else
            Cuenta = RaizCuentasCancelacion & Mid(Rs!codmacta, LCta + 1)
            
        End If
        CtaCancelacion = Rs!codmacta
    
        
        
        
        'Si dice que agrupamos vto entonces NO
        If AgrupaVtosPuente Then
            If CadenaAgrupaVtoPuente = "" Then
                'Para insertarlo al final del proceso
                'Antes de ejecutar el sql(al final) substituiremos, la cadena
                ' la cadena ### por el importe total
                
                SQL = "1,'" & Cuenta & "','Nº" & Format(IdRecepcion, "0000000") & "'," & vCP.condecli
                
                
                'Noviembre 2014
                'si pone contrapartida, pondre la nommacta
                aux2 = ""
                If vCP.ampdecli = 4 Then aux2 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaCancelacion, "T")
                
                If aux2 = "" Then aux2 = Mid(vCP.descformapago & " " & DBLet(Rs!reftalonpag, "T"), 1, 30)
                
                SQL = SQL & ",'" & DevNombreSQL(aux2) & "',"
                aux2 = ""
                'Importe al DEBE.
                SQL = SQL & "###,NULL,NULL,"
                'Contra partida
                SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
                
                CadenaAgrupaVtoPuente = AmpRemesa & SQL
            End If
        End If
            
            
        
        
        'Crearemos el apnte y la contrapartida
        ' Es decir
        '   4310  contra 430
        '   430  contr   4310
        'Lineas de apuntes .
        
         
        'Cuenta
        SQL = Linea & ",'" & Cuenta & "','" & Format(Rs!NumFactu, "000000000") & "'," & vCP.condecli
        
        
        'Noviembre 2014
        'Noviembre 2014
        'si pone contrapartida, pondre la nommacta
        aux2 = ""
        If vCP.ampdecli = 4 Then aux2 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaCancelacion, "T")
        If aux2 = "" Then aux2 = Mid(vCP.descformapago & Ampliacion, 1, 30)
        SQL = SQL & ",'" & DevNombreSQL(aux2) & "',"
        
        
        
        
        Importe = Importe + Rs!Importe
        'Gastos = Gastos + DBLet(Rs!Gastos, "N")
        
        
        'Importe VA alhaber del cliente, al debe de la cancelacion
        SQL = SQL & TransformaComasPuntos(Rs!Importe) & ",NULL,NULL,"
    
        'Contra partida
        SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
        
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
        
        
        
        
        SQL = AmpRemesa & SQL
        If Not AgrupaVtosPuente Then
            If Not Ejecuta(SQL) Then Exit Function
        End If
        Linea = Linea + 1 'Siempre suma mas uno
        
        
        'La contrapartida
        SQL = Linea & ",'" & CtaCancelacion & "','" & Format(Rs!NumFactu, "000000000") & "'," & vCP.conhacli
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & Ampliacion, 1, 30)) & "',"
        
        
        '
        SQL = SQL & "NULL," & TransformaComasPuntos(Rs!Importe) & ",NULL,"
    
        'Contra partida
        SQL = SQL & "'" & Cuenta & "','CONTAB',0,"
        
        TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Rs!Codforpa, "N")

        
        SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & ","
        SQL = SQL & DBSet(TipForpa, "T") & "," & DBSet(Rs!reftalonpag, "T") & "," & DBSet(Rs!Banco, "T") & ") "
        
        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
            
        '
        Rs.MoveNext
    Wend
    Rs.Close



    
    SQL = "Select importe,codmacta,numeroref from talones where codigo = " & IdRecepcion
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then Err.Raise 513, , "No se ha encontrado documento: " & IdRecepcion
    Importeauxiliar = Rs!Importe
    Cuenta = Rs!codmacta
    Ampliacion = DevNombreSQL(Rs!numeroref)
    Rs.Close


    If Importe <> Importeauxiliar Then
    
        'numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce,timporteD, timporteH, codccost, ctacontr, idcontab, punteada,  numserie, numfaccl, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag
    
    
        CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
        If CtaCancelacion = "" Then Err.Raise 513, , "Cuenta beneficios/pérdidas  NO espeficicada"
        
        'Hemos llegado a aqui.
        'Veremos si hay diferencia entre la suma de importe y el importe del documento.
        '
        Importeauxiliar = Importeauxiliar - Importe
        If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
        
        SQL = Linea & ",'" & CtaCancelacion & "','Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.condecli
        
        'Ampliacion
        If TalonTipoDoc = 1 Then
            aux2 = " Tal nº: " & Ampliacion
        ElseIf TalonTipoDoc = 2 Then
            aux2 = " Confi. nº: " & Ampliacion
        Else
            aux2 = " Pag. nº: " & Ampliacion
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & aux2, 1, 30)) & "',"

        
        If Importeauxiliar < 0 Then
            'NEgativo. Va en positivo al otro lado
            SQL = SQL & TransformaComasPuntos(Abs(Importeauxiliar)) & ",NULL,"
        Else
            SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Importeauxiliar)) & ","
        End If
                
        'Centro de coste
        RequiereCCDiferencia = False
        If vParam.autocoste Then
            aux2 = Mid(CtaCancelacion, 1, 1)
            If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
        End If
        If RequiereCCDiferencia Then
            CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
            If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
            CtaCancelacion = "'" & CtaCancelacion & "'"
        Else
             CtaCancelacion = "NULL"
        End If
        SQL = SQL & CtaCancelacion
        

        
        
        SQL = SQL & "," & Cuenta & ",'CONTAB',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
        
        
        
        
        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
        
        
        If AgrupaVtosPuente Then
            'Modificamos el importe final por si esta agrupando vencimientos
            Importe = Importeauxiliar + Importe
        Else
                'creamos la contrapartida para que  cuadre el asiento
                
                If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
    
                SQL = Linea & "," & Cuenta & ",'Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.conhacli
                
                If TalonTipoDoc = 1 Then
                    aux2 = " Tal nº: " & Ampliacion
                ElseIf TalonTipoDoc = 2 Then
                    aux2 = " Confi. nº: " & Ampliacion
                Else
                    aux2 = " Pag. nº: " & Ampliacion
                End If
                SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & aux2, 1, 30)) & "',"
                
                If Importeauxiliar > 0 Then
                    'NEgativo. Va en positivo al otro lado
                    SQL = SQL & TransformaComasPuntos(CStr(Importeauxiliar)) & ",NULL,"
                Else
                    SQL = SQL & "NULL," & TransformaComasPuntos(Abs(Importeauxiliar)) & ","
                End If
                        
                'Centro de coste
                RequiereCCDiferencia = False
                If vParam.autocoste Then
                    aux2 = Mid(Cuenta, 2, 1)  'pq lleva una comilla
                    If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
                End If
                If RequiereCCDiferencia Then
                    CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
                    If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
                    CtaCancelacion = "'" & CtaCancelacion & "'"
                Else
                     CtaCancelacion = "NULL"
                End If
                
                
                SQL = SQL & CtaCancelacion
                
                'Contra partida
                CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
                SQL = SQL & ",'" & CtaCancelacion & "','CONTAB',0,"
                
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
        
                SQL = AmpRemesa & SQL
                
                If Not Ejecuta(SQL) Then Exit Function
                Linea = Linea + 1
            End If
                
    End If
    
    If AgrupaVtosPuente Then
        'Tenmos que reemplazar
        'en CadenaAgrupaVtoPuente    ###:importe
        SQL = TransformaComasPuntos(CStr(Importe))
        SQL = Replace(CadenaAgrupaVtoPuente, "###", SQL)
        Conn.Execute SQL
    End If


    AmpRemesa = "F"    ' cancelada
    
    SQL = "UPDATE talones SET contabilizada = 1"
    SQL = SQL & " WHERE codigo = " & IdRecepcion
    
    Conn.Execute SQL

    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaAbono
    
    'Todo OK
    RemesasCancelacionTALONPAGARE = True
    
    
ERemesa_CancelacionCliente2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
   
    End If
    Set Rs = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function





















'*********************************************************************************
'*********************************************************************************
'   Eliminar TALON PAGARE contabilizado (contra ctas puente)
'
'
'DiarioConcepto:Llevara el diario y los conceptos al debe y al haber. NO cojera los de la stipforpa, si no de una window anterior
'              El cuarto pipe que lleva es si agrrupa en la cuenta puente
'                   es decir, en lugar de 43.1 contra 431.1
'                                         43.2 contra 431.1
'                           hacemos   43.1 y 43.2   contra la suma en 431.1
' Septiembre 2009
'           El quinto y sexto pipe llevaran, si necesario, cta dodne poner el benefic po perd del talon y si requiere cc
' Talones   0 Pagare   1 Talon   2 Confirming
Public Function EliminarCancelacionTALONPAGARE(Talones2 As Byte, IdRecepcion As Integer, FechaAbono As Date, DiarioConcepto As String) As Boolean

'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
'Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacion As String
Dim CuentaUnica As Boolean
Dim LCta As Integer
Dim Importeauxiliar As Currency
Dim AgrupaVtosPuente As Boolean
Dim CadenaAgrupaVtoPuente As String
Dim aux2 As String
Dim RequiereCCDiferencia As Boolean

Dim Obs As String
Dim TipForpa As String


    On Error GoTo ERemesa_CancelacionCliente3
    EliminarCancelacionTALONPAGARE = False
    

    If Talones2 = 1 Then
        'Sobre talones
        Cuenta = "taloncta"
    ElseIf Talones2 = 2 Then
        Cuenta = "confirmingcta"
    Else
        Cuenta = "pagarecta"
    End If
    RaizCuentasCancelacion = DevuelveDesdeBD(Cuenta, "paramtesor", "codigo", "1", "N")
    If RaizCuentasCancelacion = "" Then
        MsgBox "Error grave en configuración de  parámetros de tesorería. Falta cuenta cancelación", vbExclamation
        Exit Function
    End If
    
    LCta = Len(RaizCuentasCancelacion)
    CuentaUnica = LCta = vEmpresa.DigitosUltimoNivel
    
    
    'Comprobacion.  Para todos los efectos de la 43.... se cancelan con la 4310....
    '
    'Tendre que ver que existen estas cuentas
    Set Rs = New ADODB.Recordset
    
    SQL = "DELETE FROM tmpcierre1 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
        
    If Not CuentaUnica Then
        'Cancela contra subcuentas de cliente
        

        Ampliacion = ",CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")" & ")"
            
        SQL = "Select " & vUsu.Codigo & Ampliacion & " from talones where codigo=" & IdRecepcion
        SQL = SQL & " GROUP BY codmacta"
        'INSERT
        SQL = "INSERT INTO tmpcierre1(codusu,cta) " & SQL
        Conn.Execute SQL
        
        'Ahora monto el select para ver que cuentas 430 no tienen la 4310
        SQL = "Select cta,codmacta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " HAVING codmacta is null"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        Linea = 0
        While Not Rs.EOF
            Linea = Linea + 1
            SQL = SQL & Rs!Cta & "     "
            If Linea = 5 Then
                SQL = SQL & vbCrLf
                Linea = 0
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        
        If SQL <> "" Then
            
            AmpRemesa = "CANCELACION contab"
            
            SQL = "Cuentas " & AmpRemesa & ".  No existen las cuentas: " & vbCrLf & String(90, "-") & vbCrLf & SQL
            SQL = SQL & vbCrLf & "¿Desea crearlas?"
            Linea = 1
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                'Ha dicho que si desea crearlas
                
                Ampliacion = "CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")) "
            
                SQL = "Select codmacta," & Ampliacion & " from talones where codigo=" & IdRecepcion
                SQL = SQL & " and " & Ampliacion & " in "
                SQL = SQL & "(Select cta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
                SQL = SQL & " AND codmacta is null)"
                Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                
                     SQL = "INSERT  IGNORE INTO cuentas(codmacta,nommacta ,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) SELECT '"
                                ' CUenta puente
                     SQL = SQL & Rs.Fields(1) & "',nommacta ,'S',0,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from cuentas where codmacta = '"
                                'Cuenta en la scbro (codmacta)
                     SQL = SQL & Rs.Fields(0) & "'"
                     Conn.Execute SQL
                     Rs.MoveNext
                     
                Wend
                Rs.Close
                Linea = 0
            End If
            If Linea = 1 Then GoTo ERemesa_CancelacionCliente3
        End If
        
    Else
        'Cancela contra UNA unica cuenta todos los vencimientos
        SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", RaizCuentasCancelacion, "T")
        If SQL = "" Then
            MsgBox "No existe la cuenta puente: " & RaizCuentasCancelacion, vbExclamation
            GoTo ERemesa_CancelacionCliente3
        End If
    End If

    
    'La forma de pago
    Set vCP = New Ctipoformapago
    If Talones2 = 1 Then
        SQL = CStr(vbTalon)
        Ampliacion = "Talones"
    ElseIf Talones2 = 2 Then
        SQL = CStr(vbConfirming)
        Ampliacion = "Confirming"
    Else
        SQL = CStr(vbPagare)
        Ampliacion = "Pagarés"
    End If
    If vCP.Leer(CInt(SQL)) = 1 Then GoTo ERemesa_CancelacionCliente3
    'Ahora fijo los valores que me ha dado el
    vCP.diaricli = RecuperaValor(DiarioConcepto, 1)
    'En la contabilizacion
    'vCP.condecli = RecuperaValor(DiarioConcepto, 2)
    'vCP.conhacli = RecuperaValor(DiarioConcepto, 3)
    'En la eliminacion
    vCP.conhacli = RecuperaValor(DiarioConcepto, 2)
    vCP.condecli = RecuperaValor(DiarioConcepto, 3)
    AgrupaVtosPuente = RecuperaValor(DiarioConcepto, 4) = 1
 
 
 
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion ) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Eliminar recepcion documentos contabilizada(cancelada )"

    SQL = SQL & " NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & "Generado el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & "',"
    
    Obs = "ARICONTA 6: Eliminar recepción documentos contabilizada: " & vbCrLf & " NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Obs, "T") & ");"
    
    If Not Ejecuta(SQL) Then Exit Function
    
    Linea = 1
    Importe = 0
    'Gastos = 0
    
    vCP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)  'DEBE
    vCP.CadenaAuxiliar = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)   'DEBE
    
    
    SQL = "select cobros.*,l.importe,l.codigo, c.numeroref reftalonpag, c.banco from  (talones c inner join  talones_facturas l on c.codigo = l.codigo)  left join  cobros  on l.numserie=cobros.numserie and"
    SQL = SQL & " l.numfactu=cobros.numfactu and   l.fecfactu=cobros.fecfactu and l.numorden=cobros.numorden"
    SQL = SQL & " WHERE l.codigo= " & IdRecepcion
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Trozo comun
    AmpRemesa = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
    AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada, "
    AmpRemesa = AmpRemesa & " numserie, numfaccl, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
    AmpRemesa = AmpRemesa & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
    
    CadenaAgrupaVtoPuente = ""

    While Not Rs.EOF
        Ampliacion = Rs!NUmSerie   'SI DA ERROR ES QUE NO EXISTE mediante el left join
        
        
        'Neuvo dato para la ampliacion en la contabilizacion
        Ampliacion = " "
        Select Case vCP.ampdecli
        Case 2, 4
            'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
            Ampliacion = DBLet(Rs!reftalonpag, "T")
            If Ampliacion = "" Then Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
        Case 5
            Ampliacion = DBLet(Rs!reftalonpag, "T")
            If Ampliacion = "" Then
                Ampliacion = Rs!NUmSerie & "/" & Rs!NumFactu
            Else
                Ampliacion = "Doc: " & Ampliacion
            End If
        Case Else
           If vCP.ampdecli = 1 Then Ampliacion = vCP.siglas & " "
           Ampliacion = Ampliacion & Rs!NUmSerie & "/" & Rs!NumFactu
        End Select
        If Mid(Ampliacion, 1, 1) <> " " Then Ampliacion = " " & Ampliacion
        
        'Cancelacion
        If CuentaUnica Then
            Cuenta = RaizCuentasCancelacion
        Else
            Cuenta = RaizCuentasCancelacion & Mid(Rs!codmacta, LCta + 1)
            
        End If
        CtaCancelacion = Rs!codmacta
    
        
        'Si dice que agrupamos vto entonces NO
        If AgrupaVtosPuente Then
            If CadenaAgrupaVtoPuente = "" Then
                'Para insertarlo al final del proceso
                'Antes de ejecutar el sql(al final) substituiremos, la cadena
                ' la cadena ### por el importe total
                
                SQL = "1,'" & Cuenta & "','Nº" & Format(IdRecepcion, "0000000") & "'," & vCP.condecli
                
                SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & " " & DBLet(Rs!reftalonpag, "T"), 1, 30)) & "',"
                'Importe al HABER.
                SQL = SQL & "NULL,###,NULL,"
                'Contra partida
                SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
                
                CadenaAgrupaVtoPuente = AmpRemesa & SQL
            End If
        End If
        
        
        'Crearemos el apnte y la contrapartida
        ' Es decir
        '   4310  contra 430
        '   430  contr   4310
        'Lineas de apuntes .
        
         
        'Cuenta
        SQL = Linea & ",'" & Cuenta & "','" & Format(Rs!NumFactu, "000000000") & "'," & vCP.condecli
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & Ampliacion, 1, 30)) & "',"
        
        
        
        Importe = Importe + Rs!Importe
        'Gastos = Gastos + DBLet(Rs!Gastos, "N")
        
        
        'Importe VA alhaber del cliente, al debe de la cancelacion
        SQL = SQL & "NULL," & TransformaComasPuntos(Rs!Importe) & ",NULL,"
    
        'Contra partida
        SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        
        
        
        SQL = AmpRemesa & SQL
        If Not AgrupaVtosPuente Then
            If Not Ejecuta(SQL) Then Exit Function
        End If
        Linea = Linea + 1 'Siempre suma mas uno
        
        
        'La contrapartida
        SQL = Linea & ",'" & CtaCancelacion & "','" & Format(Rs!NumFactu, "000000000") & "'," & vCP.conhacli
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & Ampliacion, 1, 30)) & "',"
        
        
        '
        SQL = SQL & TransformaComasPuntos(Rs!Importe) & ",NULL,NULL,"
    
        'Contra partida
        SQL = SQL & "'" & Cuenta & "','CONTAB',0,"
        TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Rs!Codforpa, "N")
        
        SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & ","
        SQL = SQL & DBSet(TipForpa, "T") & "," & DBSet(Rs!reftalonpag, "T") & "," & DBSet(Rs!Banco, "T") & ")"

        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
            
        
        Rs.MoveNext
    Wend
    Rs.Close


    
    SQL = "Select importe,codmacta,numeroref from talones where codigo = " & IdRecepcion
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then Err.Raise 513, , "No se ha encontrado documento: " & IdRecepcion
    Importeauxiliar = Rs!Importe
    Cuenta = Rs!codmacta
    Ampliacion = DevNombreSQL(Rs!numeroref)
    Rs.Close


    If Importe <> Importeauxiliar Then
    
        CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
        If CtaCancelacion = "" Then Err.Raise 513, , "Cuenta beneficios/pérdidas  NO espeficicada"
        
        'Hemos llegado a aqui.
        'Veremos si hay diferencia entre la suma de importe y el importe del documento.
        '
        Importeauxiliar = Importeauxiliar - Importe
        If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
        
        SQL = Linea & ",'" & CtaCancelacion & "','Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.conhacli
        
        'Ampliacion
        If Talones2 = 1 Then
            aux2 = " Tal nº: " & Ampliacion
        ElseIf Talones2 = 2 Then
            aux2 = " Conf. nº: " & Ampliacion
        Else
            aux2 = " Pag. nº: " & Ampliacion
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & aux2, 1, 30)) & "',"

        
        If Importeauxiliar < 0 Then
            'NEgativo. Va en positivo al otro lado
            SQL = SQL & "NULL," & TransformaComasPuntos(Abs(Importeauxiliar)) & ","
        Else
            SQL = SQL & TransformaComasPuntos(CStr(Importeauxiliar)) & ",NULL,"
        End If
                
        'Centro de coste
        RequiereCCDiferencia = False
        If vParam.autocoste Then
            aux2 = Mid(CtaCancelacion, 1, 1)
            If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
        End If
        If RequiereCCDiferencia Then
            CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
            If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
            CtaCancelacion = "'" & CtaCancelacion & "'"
        Else
             CtaCancelacion = "NULL"
        End If
        SQL = SQL & CtaCancelacion
        
        'Contra partida
        If CuentaUnica Then
            Cuenta = "'" & RaizCuentasCancelacion & "'"
        Else
            Cuenta = "NULL"
        End If
        
        
        SQL = SQL & "," & Cuenta & ",'CONTAB',0,"
        
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        
        
        
        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
        
        
        If AgrupaVtosPuente Then
            'Modificamos el importe final por si esta agrupando vencimientos
            Importe = Importeauxiliar + Importe
        Else
                'creamos la contrapartida para que  cuadre el asiento
            
                If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
                
                SQL = Linea & "," & Cuenta & ",'Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.conhacli
                
                 If Talones2 = 1 Then
                    aux2 = " Tal nº: " & Ampliacion
                ElseIf Talones2 = 2 Then
                    aux2 = " Conf. nº: " & Ampliacion
                Else
                    aux2 = " Pag. nº: " & Ampliacion
                End If
                SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & aux2, 1, 30)) & "',"
        
                
                If Importeauxiliar > 0 Then
                    'NEgativo. Va en positivo al otro lado
                    SQL = SQL & TransformaComasPuntos(CStr(Importeauxiliar)) & ",NULL,"
                Else
                    SQL = SQL & "NULL," & TransformaComasPuntos(Abs(Importeauxiliar)) & ","
                End If
                        
                'Centro de coste
                RequiereCCDiferencia = False
                If vParam.autocoste Then
                    aux2 = Mid(Cuenta, 2, 1)  'pq lleva una comilla
                    If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
                End If
                If RequiereCCDiferencia Then
                    CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
                    If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
                    CtaCancelacion = "'" & CtaCancelacion & "'"
                Else
                     CtaCancelacion = "NULL"
                End If
                SQL = SQL & CtaCancelacion
                
                'Contra partida
                CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
                SQL = SQL & ",'" & CtaCancelacion & "','CONTAB',0,"
                
                '###FALTA1
                
                SQL = AmpRemesa & SQL
                
                If Not Ejecuta(SQL) Then Exit Function
                Linea = Linea + 1
            End If
                
    End If
    
    If AgrupaVtosPuente Then
        'Tenmos que reemplazar
        'en CadenaAgrupaVtoPuente    ###:importe
        SQL = TransformaComasPuntos(CStr(Importe))
        SQL = Replace(CadenaAgrupaVtoPuente, "###", SQL)
        Conn.Execute SQL
    End If


    AmpRemesa = "F"    ' cancelada
    




    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaAbono
    
    'Todo OK
    EliminarCancelacionTALONPAGARE = True
    
    
ERemesa_CancelacionCliente3:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set Rs = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function







'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'
'
'   Contabilizacion especial N19.
'   Genera tantos apuntes como fechas vto haya que sera la fecha del asie
'
'
'
'
'   Solo Recibo bancario, norma 19, si ctas puente
'
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------


'  CuentaPuente  "":NO    len utleve: UNA UNICA   2 Raiz
Public Function ContabNorma19PorFechaVto(Codigo As Integer, Anyo As Integer, CtaBanco As String, CuentaPuente As String, LaFechaSiPuente As Date) As Boolean
Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim AmpRemesa As String
'Dim CtaParametros As String
'Dim Cuenta As String
'
'
Dim ImpoAux As Currency
Dim aux2 As String
Dim AUX3 As String  'por si lleva cuenta puente por cuenta
Dim LaCt As Byte

Dim ColFechas As Collection  'Cada una de las fechas de vencimiento de la remesa
Dim NF As Integer
Dim FecAsto As Date

    On Error GoTo ECon
    
    ContabNorma19PorFechaVto = False

    'La forma de pago
    Set vCP = New Ctipoformapago
    If vCP.Leer(vbTipoPagoRemesa) = 1 Then GoTo ECon
    
    Set Rs = New ADODB.Recordset
    Set ColFechas = New Collection
    
    
    
    SQL = "fecvenci"
    If CuentaPuente <> "" Then SQL = "   STR_TO_DATE('" & LaFechaSiPuente & "' , '%d/%m/%Y') "
    
    SQL = "Select " & SQL & " from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo & " GROUP BY 1 ORDER By 1"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        SQL = Rs.Fields(0)
        ColFechas.Add SQL
        Rs.MoveNext
    Wend
    Rs.Close
    If ColFechas.Count = 0 Then Err.Raise 513, "No hay vencimientos(n19)"
    
    
    For NF = 1 To ColFechas.Count
        FecAsto = CDate(ColFechas.Item(NF))
        
        Set Mc = New Contadores
    
    
        If Mc.ConseguirContador("0", FecAsto <= vParam.fechafin, True) = 1 Then Exit Function
    
    
        'Insertamos la cabera
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion,desdeaplicacion) VALUES ("
        SQL = SQL & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador
        SQL = SQL & ", '"
        SQL = SQL & "Abono remesa: " & Codigo & " / " & Anyo & "       N19" & vbCrLf
        SQL = SQL & "Proceso: " & NF & " / " & ColFechas.Count & vbCrLf & "',"
        'SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "');"
        SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Abono remesa')"
        If Not Ejecuta(SQL) Then Exit Function
        
        Linea = 1
        Importe = 0
        Gastos = 0
        
        'La ampliacion para el banco
        AmpRemesa = ""
        SQL = "Select * from remesas WHERE codigo=" & Codigo & " AND anyo = " & Anyo

        Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'NO puede ser EOF
        
        
        If Not IsNull(Rs!Descripcion) Then AmpRemesa = Rs!Descripcion
        
        
        If AmpRemesa = "" Then
            AmpRemesa = " Remesa: " & Codigo & "/" & Anyo
        Else
            AmpRemesa = " " & AmpRemesa
        End If
        
        Rs.Close
        
        'AHORA Febrero 2009
        '572 contra  5208  Efectos descontados
        '-------------------------------------
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
    
        Gastos = 0
        
        Importe = 0
        SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
        'y por vencimiento
        'si  no va contra puente ABRIL 2019
        If CuentaPuente = "" Then SQL = SQL & " AND fecvenci = '" & Format(FecAsto, FormatoFecha) & "'"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not Rs.EOF
            'Banco contra cliente. Si lleva pte, puente
            'La linea del banco
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
        
            'Cuenta
            SQL = SQL & Rs!codmacta & "','" & Rs!NUmSerie & Format(Rs!NumFactu, "0000000") & "'," & vCP.conhacli
            
            
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
            Ampliacion = Ampliacion & " "
                   
            'Neuvo dato para la ampliacion en la contabilizacion
            Select Case vCP.amphacli
            Case 2
               Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
            Case 4
                'Contrapartida BANCO
                Cuenta = RecuperaValor(CtaBanco, 1)
                Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                Ampliacion = Ampliacion & AmpRemesa
            Case Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFactu, "0000000")
            End Select
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            Importe = Importe + Rs!ImpVenci
                
            Gastos = Gastos + DBLet(Rs!Gastos, "N")
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe
            SQL = SQL & "NULL," & TransformaComasPuntos(Rs!ImpVenci) & ",NULL,"
        
            
            If vCP.ctrdecli = 1 Then
            
            
                If CuentaPuente = "" Then
                    aux2 = RecuperaValor(CtaBanco, 1)
                Else
                    
                    If Len(CuentaPuente) = vEmpresa.DigitosUltimoNivel Then
                        aux2 = CuentaPuente
                    Else
                        LaCt = Len(CuentaPuente) + 1
                        
                        aux2 = CuentaPuente & Mid(Rs!codmacta, LaCt)
                    End If
                End If
                
                SQL = SQL & "'" & aux2 & "',"
                    
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'COBROS',0,"
            
            'los datos de la factura (solo en el apunte del cliente)
            Dim TipForpa As Byte
            TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", Rs!Codforpa, "N")
            
            SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & DBSet(Rs!NumFactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!numorden, "N") & "," & DBSet(TipForpa, "N") & ","
            SQL = SQL & "1," & DBSet(Codigo, "N") & "," & DBSet(Anyo, "N") & ")"
                
            
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
            
            
            
            If CuentaPuente <> "" Then
                If Len(CuentaPuente) <> vEmpresa.DigitosUltimoNivel Then
                    'Cuenta puente
            
            
                    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
                    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
                    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
                    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
                
                    'Cuenta
                    SQL = SQL & aux2 & "','" & Rs!NUmSerie & Format(Rs!NumFactu, "0000000") & "'," & vCP.conhacli
                    
                    
                    Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
                    Ampliacion = Ampliacion & " "
                           
                    
                   If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                   Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFactu, "0000000")
                
                    SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
                    
                    
                    
                    ' timporteH, codccost, ctacontr, idcontab, punteada
                    'Importe
                    SQL = SQL & TransformaComasPuntos(Rs!ImpVenci) & ",NULL,NULL,"
                
                    
                  
                        
                    SQL = SQL & "'" & Rs!codmacta & "',"
                    
                    SQL = SQL & "'COBROS',0,"
                    
                    
                    
                    SQL = SQL & "null,null,null,null,null,null,null,null)"
                        
                    
                    If Not Ejecuta(SQL) Then Exit Function
                    
                    Linea = Linea + 1
            
                End If
            End If
            
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        'La linea del banco

            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador & ","
        
            
            'Gastos de los recibos.
            'Si tiene alguno de los efectos remesados gastos
            If Gastos > 0 Then
                Linea = Linea + 1
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
                Ampliacion = "RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli & ",'" & Ampliacion & " " & Codigo & "/" & Anyo & "'"
        
        
        
                Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 2) & "','" & Ampliacion & ",NULL,"
                Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Gastos)) & ","
        
              
                Ampliacion = Ampliacion & "NULL"
               
                Ampliacion = Ampliacion & ",NULL,'COBROS',0)"
                Ampliacion = SQL & Ampliacion
                If Not Ejecuta(Ampliacion) Then Exit Function
                Linea = Linea + 1
            End If
            
          
           
        
            ImpoAux = Importe + Gastos
        
        
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & AmpRemesa
            
            aux2 = RecuperaValor(CtaBanco, 1)
            
            Ampliacion = Linea & ",'" & aux2 & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli & ",'" & Ampliacion & "',"
            Ampliacion = Ampliacion & TransformaComasPuntos(CStr(ImpoAux)) & ",NULL,NULL,"
            
            If vCP.ctrdecli = 0 Then
                Ampliacion = Ampliacion & "NULL"
            Else
        
                Ampliacion = Ampliacion & "NULL"
        
            End If
            Ampliacion = Ampliacion & ",'COBROS',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then Exit Function
            
            
            
            
            
            If CuentaPuente <> "" Then
            
                    Linea = Linea + 1
            
                    Ampliacion = "Abono remesa  " 'DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
                    Ampliacion = Ampliacion & AmpRemesa
                
                    aux2 = RecuperaValor(CtaBanco, 1)
                    aux2 = DevuelveDesdeBD("ctaefectosdesc", "bancos", "codmacta", aux2)
                    If aux2 = "" Then Err.Raise 513, , "Cta efectos descontados, vacia"
                
                     Ampliacion = Linea & ",'" & aux2 & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli & ",'" & Ampliacion & "',"
                    Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,'"
                
                    aux2 = RecuperaValor(CtaBanco, 1)
                                 
                     Ampliacion = Ampliacion & aux2
                   Ampliacion = Ampliacion & "','COBROS',0)"
                  Ampliacion = SQL & Ampliacion
                 If Not Ejecuta(Ampliacion) Then Exit Function
                
            
            End If
            
            
            
            
        'Insertamos para pasar a hco
        InsertaTmpActualizar Mc.Contador, vCP.diaricli, FecAsto
        
        
        
        
        'Estamos recorriendo por fechas
        Set Mc = Nothing
   Next NF
        
        
    'AHora actualizamos los efectos.
    SQL = "UPDATE cobros SET"
    SQL = SQL & " siturem= 'Q'"
    SQL = SQL & ", situacion = 1 "
    SQL = SQL & " WHERE codrem=" & Codigo
    SQL = SQL & " and anyorem=" & Anyo
    Conn.Execute SQL

    'Todo OK
    ContabNorma19PorFechaVto = True
ECon:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set Rs = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
    Set ColFechas = Nothing
End Function















'TALONES PAGARES
Public Function RemesasEliminarVtosTalonesPagares(TipoRemesa As Byte, Codigo As Integer, Anyo As Integer, FechaAbono As Date, ByRef FP As Ctipoformapago, AgrupaCancelacion_ As Boolean) As Byte
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacion As String
Dim CuentaUnica As Boolean
Dim LCta As Integer
Dim Dias1 As Integer 'Si no llega al limite
Dim J As Long
Dim F As Date
Dim CuentaPuente As Boolean
Dim BorrarEfecto As Boolean
Dim ImporteVto As Currency
Dim EstaAgrupandoVtos As Boolean
Dim CadenaAgrupacion As String
Dim ParaLineasDocumentosRecibidos As String
Dim vId As Integer
Dim ImporteDocumento As Currency
Dim SumasImportesDocumentos As Currency
Dim EfectosEliminados As Boolean   'Realmente no se borran. Updateados
Dim EfectosEliminados_ As Integer   'Realmente no se borran. Updateados
Dim EliminaEnRecepcionDocumentos As String

Dim ParaElLog As String


    On Error GoTo ERemesa_Elivto
    RemesasEliminarVtosTalonesPagares = 2
    
    
    
    
    CuentaPuente = False
    EstaAgrupandoVtos = False
    If TipoRemesa = 3 Then
        'Sobre talones
        Cuenta = "taloncta"
        CuentaPuente = vParamT.TalonesCtaPuente
       
    ElseIf TipoRemesa = 2 Then
        CuentaPuente = vParamT.PagaresCtaPuente
        Cuenta = "pagarecta"
    
    ElseIf TipoRemesa = 5 Then
        CuentaPuente = vParamT.ConfirmingCtaPuente
        Cuenta = "pagarecta"
    Else
        'Efectos. Viene de cancelacion
        Cuenta = "RemesaCancelacion"
        CuentaPuente = True
        
    End If
    
    
    
    
    
    If CuentaPuente Then
        RaizCuentasCancelacion = DevuelveDesdeBD(Cuenta, "paramtesor", "codigo", "1", "N")
        If RaizCuentasCancelacion = "" Then
            MsgBox "Error grave en configuracion de  parametros de tesoreria. Falta cuenta cancelacion", vbExclamation
            Exit Function
        End If
        
        LCta = Len(RaizCuentasCancelacion)
        CuentaUnica = LCta = vEmpresa.DigitosUltimoNivel
      
        EstaAgrupandoVtos = AgrupaCancelacion_
       
    End If
            
    Set Rs = New ADODB.Recordset
    
    EliminaEnRecepcionDocumentos = ""
    'Datos bancos. Importe maximo para dias 1, dias2 si no llega
    If TipoRemesa = 3 Then
        'Sobre talones
        Cuenta = "talondias"
    ElseIf TipoRemesa = 1 Then
        Cuenta = "0 "   'recibos, fectos
    ElseIf TipoRemesa = 4 Then
        Cuenta = "0 "   '
    Else
        Cuenta = "pagaredias"
        
    End If
        
    SQL = "select ctaefectosdesc," & Cuenta & " from remesas r,bancos b where r.codmacta=b.codmacta and codigo=" & Codigo & " AND anyo = " & Anyo
    CtaCancelacion = ""
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    If Not Rs.EOF Then
        
        If Not IsNull(Rs.Fields(0)) Then
            CtaCancelacion = Rs.Fields(0)
        Else
            If CuentaPuente Then SQL = "Cuenta efectos descontados"
        End If
        
        If IsNull(Rs.Fields(1)) Then
            SQL = SQL & "Dias eliminacion"
        Else
            Dias1 = Rs.Fields(1)
        End If

    End If
    Rs.Close
    If SQL <> "" Then
        MsgBox "Falta configurar: " & SQL, vbExclamation
        GoTo ERemesa_Elivto
    End If


    
  
    'Si es talon pagare el importe lo coje de las lineas de vto, NO de la scobro
    If TipoRemesa <> 4 Then
      
        SQL = "select cobros.*,l.importe,l.numserie vto,codigo from  talones_facturas l left join  cobros "
        SQL = SQL & " on l.numserie=cobros.numserie and l.numfactu=cobros.numfactu and"
        SQL = SQL & " l.fecfactu=cobros.fecfactu and l.numorden=cobros.numorden"
        SQL = SQL & " where codrem=" & Codigo & " AND anyorem = " & Anyo
        SQL = SQL & " AND siturem<>'Z'     "
        SQL = SQL & " ORDER BY codigo" 'Pazra ir comprobando documento por documento si
        
    
    Else
        SQL = "select cobros.* from    cobros "
        SQL = SQL & " where codrem=" & Codigo & " AND anyorem = " & Anyo
        SQL = SQL & " AND siturem<>'Z'     "
        SQL = SQL & " ORDER BY fecvenci" 'Pazra ir comprobando documento por documento si
        
    End If
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "Ningun vencimiento seleccionado", vbExclamation
        Set Rs = Nothing
        Exit Function
        SQL = "UPDATE remesas SET situacion= 'Z'"
        SQL = SQL & " WHERE codigo=" & Codigo & " and anyo=" & Anyo
        Conn.Execute SQL
        Rs.Close
        RemesasEliminarVtosTalonesPagares = 1
        Exit Function
    End If
    
    'La forma de pago
    If CuentaPuente Then
            
            If TipoRemesa = 3 Then
                'SQL = CStr(vbTalon)
                Ampliacion = "Talones"
            ElseIf TipoRemesa = 2 Then
                'SQL = CStr(vbPagare)
                Ampliacion = "Pagarés"
            Else
                'SQL = CStr(vbTipoPagoRemesa)
                Ampliacion = "Efectos"
            End If
            
            
            
            Set Mc = New Contadores
            
            
            If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then GoTo ERemesa_Elivto
            
        
        
            'Insertamos la cabera
            SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien,  obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
            SQL = SQL & FP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
            SQL = SQL & ", 'Eliminacion riesgo. Remesa: " & Codigo & " / " & Anyo & "   " & Ampliacion & vbCrLf
            SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & "',"
            SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Eliminar riesgo remesas" & " ');"
            If Not Ejecuta(SQL) Then Exit Function
        
    
        
    
            Linea = 1
            Importe = 0

    
            FP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.conhacli)
            FP.CadenaAuxiliar = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.condecli)
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Si no lleva cuenta puente, NO contbiliza nada
    If CuentaPuente Then
        AmpRemesa = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
        AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada) "
        AmpRemesa = AmpRemesa & "VALUES (" & FP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
    End If
    CadenaAgrupacion = ""
    NumRegElim = 0 'Total registro
    SumasImportesDocumentos = 0
    EliminaEnRecepcionDocumentos = "|"
    EfectosEliminados_ = 0
    ParaElLog = ""
    While Not Rs.EOF
        J = DateDiff("d", Rs!FecVenci, Now)
        J = J - Dias1
        NumRegElim = NumRegElim + 1
        If J >= 0 Then
            
            'AHora me guardo, si procede, el id, asi luego vere si puedo eliminar
            'en recepcion de documentos
            If TipoRemesa <> vbTipoPagoRemesa Then
                SQL = "|" & Rs!Codrem & "|"
                If InStr(1, EliminaEnRecepcionDocumentos, SQL) = 0 Then EliminaEnRecepcionDocumentos = EliminaEnRecepcionDocumentos & Rs!Codigo & "|"
            End If
    
            
            If CuentaPuente Then
            
                If TipoRemesa <> vbTipoPagoRemesa Then
                    If vId <> Rs!Codigo Then
                        'Ha cambiado de documento
                        If vId > 0 Then
                            'Conseguimos importe documento
                            Cuenta = "codmacta"
                            Ampliacion = DevuelveDesdeBD("importe", "talones", "codigo", CStr(vId), "N", Cuenta)
                            ImporteVto = CCur(Ampliacion)
                            'Comprobamos con el importe parcial.
                            If ImporteVto <> ImporteDocumento Then
                                'Ha habido difernecias
    
                            
                                'Y si no agrupamos
                                If EstaAgrupandoVtos Then
                                    'Metemos uno ajustando importe
                            
                            
                                End If 'de agrupando
                            End If '<>importe
                        End If 'ID>0
                    End If
                    'Inicializamos valores
                    vId = Rs!Codigo
                    SumasImportesDocumentos = SumasImportesDocumentos + ImporteDocumento
                    ImporteDocumento = 0
                    
                     
                End If 'rs!id <>ID
                    
                'Han pasado mas dias de los que poner en paraemtros. Podremos borrar el efecto
                'Ampliacion
                Ampliacion = FP.descformapago & " "
                   
                   
                If EstaAgrupandoVtos Then
                   If CadenaAgrupacion = "" Then
                        'Creo el la linea para insertar
 
                        If Not CuentaUnica Then
                            Cuenta = RaizCuentasCancelacion
                            Cuenta = Cuenta & Mid(Rs!codmacta, LCta + 1)
                        Else
                            Cuenta = RaizCuentasCancelacion
                        End If
                        CadenaAgrupacion = Linea & ",'" & Cuenta & "','RE" & Format(Codigo, "0000") & Anyo & "'," & FP.conhacli
                        CadenaAgrupacion = CadenaAgrupacion & ",'" & DevNombreSQL(Mid(Ampliacion & "Rem: " & Codigo & "-" & Anyo, 1, 30)) & "',NULL,@@@@@@"
                
                        CadenaAgrupacion = CadenaAgrupacion & ",NULL,'" & CtaCancelacion & "','CONTAB',0)"
                        
                        
                        
                        
                        
                        
                        
                        'Luego reemplazare @@@@@@ por el importe total
                    End If

                End If
                
                'Neuvo dato para la ampliacion en la contabilizacion
                Select Case FP.amphacli
                Case 2, 4
                    'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
                    Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
                    
                Case Else
                   If FP.amphacli = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
                   Ampliacion = Ampliacion & Rs!NUmSerie & "/" & Rs!NumFactu
                End Select
                
                
                Cuenta = RaizCuentasCancelacion
                If Not CuentaUnica Then Cuenta = Cuenta & Mid(Rs!codmacta, LCta + 1)
                
            
            
             
                'Cuenta
                SQL = Linea & ",'" & Cuenta & "','" & Format(Rs!NumFactu, "000000000") & "'," & FP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',NULL,"
                If TipoRemesa = vbTipoPagoRemesa Then
                    ImporteVto = Rs!ImpVenci
                    
                Else
                    'Talones pagares, podria ser que no pagara todo.
                    'El rs esta con left join sobre slirecdoc
                    'Si es NULL vto significa que esta en scobro y no en slirecdoc. Algo esta mal.
                    If IsNull(Rs!Vto) Then
                        If MsgBox("Vencimiento no existe en rececpcion documentos. ¿Continuar?", vbYesNo) = vbNo Then
                            'esto generara un error
                            SQL = Rs!Vto
                        Else
                            ImporteVto = Rs!impcobro
                        End If
                    Else
                        ImporteVto = Rs!Importe
                    End If
                End If
                ParaElLog = ParaElLog & vbCrLf & Format(Rs!Codrem, "00000") & "  " & Format(Rs!NumFactu, "000000000") & "-" & Rs!numorden & "  " & ImporteVto
                Importe = Importe + ImporteVto
                ImporteDocumento = ImporteDocumento + ImporteVto
                
                'Importe VA alhaber de
                SQL = SQL & TransformaComasPuntos(CStr(ImporteVto)) & ",NULL,"
            
                'Contra partida
                SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0)"
                SQL = AmpRemesa & SQL
                If Not EstaAgrupandoVtos Then
                    If Not Ejecuta(SQL) Then Exit Function
                End If
            Else
               ' ImporteVto = Rs!Importe
                'falta falta falta .. Estaba lo de arriba
                ImporteVto = Rs!ImpVenci
            End If
            Linea = Linea + 1
            

            
            'Me cargo el efecto y si tuviera devoluciones
            'Para talones/pagares podria darse el caso que NO todo el importe es
            'es el que ha sido pagado. Entonces procederemos de otra forma
            BorrarEfecto = True
            If TipoRemesa <> 1 Then
                'TALONES REMESAS
                
                'Compruebo que el total impcobrado es el total +gastos
                
                If Rs!ImpVenci + DBLet(Rs!Gastos, "N") > ImporteVto Then BorrarEfecto = False
                    
                    
            End If
            'Comun a borrar y updatear
            SQL = " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu = " & Rs!NumFactu
            SQL = SQL & " AND fecfactu = '" & Format(Rs!FecFactu, FormatoFecha) & "' AND numorden = " & Rs!numorden
            If BorrarEfecto Then
                'YA NO SE BORRAN LOS EFECTOS
                Conn.Execute "UPDATE cobros set  siturem='Z' " & SQL
                EfectosEliminados_ = EfectosEliminados_ + 1
                
                
            Else
            
                'Este caso no lo he vivido todavia en la contabilidad nueva
                Err.Raise 513, , "Eliminar vtos talon -pagare.  Vencimientos no cobrados en su totalidad. Avise soporte tecnico"
            
            
                'Se trata de actualizar
                'Vamos a quitar la marca de remesado. En importe pondremos el importe que habia y en el campo observaciones
                'indicaremos que ya ha sido remesado y por cuanto importe
                
                'Campo obs..servaciones guardara los datos antiguos
                Ampliacion = "Vto: " & Format(Rs!ImpVenci, FormatoImporte)
                If DBLet(Rs!Gastos, "N") > 0 Then Ampliacion = Ampliacion & "/ Gastos " & Format(Rs!Gastos, FormatoImporte)
                'Fecha del ultimo cobro
                Ampliacion = Ampliacion & "   Ultimo cobro: " & Format(Rs!fecultco, "dd/mm/yyyy") & "  " & Format(Rs!impcobro, FormatoImporte)
                'Lo meto en la observacion
                Ampliacion = "obs = '" & Ampliacion & "',"
                
                
                'Agosto 2009
                'Como el vto esta en slirecepdoc NO hace falta cponer esto
                'Ampliacion = Ampliacion & "fecultco = NULL, impcobro=NULL,"
                Ampliacion = Ampliacion & "codrem=NULL,Tiporem=NULL,Anyorem=NULL,siturem=NULL"
                
                
                'Los gastos los pondre a null
                Ampliacion = Ampliacion & ",gastos=NULL"
                'ImporteVto = Rs!impvenci + DBLet(Rs!Gastos, "N") - Rs!impcobro
                'Ampliacion = Ampliacion & ",impvenci = " & TransformaComasPuntos(CStr(ImporteVto))
                
                'Raferencia talon/pagare tb
                Ampliacion = Ampliacion & ",reftalonpag=NULL,recedocu=0"
                
                ParaLineasDocumentosRecibidos = SQL
                
                SQL = "UPDATE scobro SET " & Ampliacion & SQL
                Ampliacion = ""
                Conn.Execute SQL
                
                
                        
                
                
                
                
                'Realmente para tipo=1 NO deberia llegar aquin
                If TipoRemesa <> 1 Then

         
                        SQL = "UPDATE slirecepdoc SET numserie="" """ & ParaLineasDocumentosRecibidos
                        SQL = Replace(SQL, "numorden", "numvenci")
                        SQL = Replace(SQL, "codfaccl", "numfaccl")
                        If Not Ejecuta(SQL) Then
                            SQL = "Error actualizando tabla lineas documentos recibidos"
                            MsgBox SQL, vbExclamation
                        End If
                End If
                
            End If  'Borrar efecto
        End If  'de             j>0
            
        Rs.MoveNext
    Wend
    Rs.Close


    If ParaElLog <> "" Then
        ParaElLog = "[Eliminar riesgo Tal/Pag]" & vbCrLf & "Vtos: " & EfectosEliminados_ & vbCrLf & vbCrLf & ParaElLog & vbCrLf & vbCrLf & Rs.Source
        vLog.Insertar 32, vUsu, ParaElLog
        SQL = ""
    End If
    
    'Comprobamos que el importe del talon es el correcto
        If CuentaPuente And vId > 0 Then
            'Conseguimos importe documento
            Cuenta = "codmacta"
            Ampliacion = DevuelveDesdeBD("importe", "talones", "codigo", CStr(vId), "N", Cuenta)
            ImporteVto = CCur(Ampliacion)
            
            If Not CuentaUnica Then
                Cuenta = RaizCuentasCancelacion & Mid(Cuenta, LCta + 1)
            Else
                Cuenta = RaizCuentasCancelacion
            End If
            'Comprobamos con el importe parcial.
            If ImporteVto <> ImporteDocumento And False Then  'ABRIL 2020. Pongo el FALSe, pq no tiene sentido lo que hacia.
                ImporteVto = ImporteDocumento - ImporteVto
                
                
                Importe = Importe - ImporteVto
                'Ha habido difernecias
                'Y si no agrupamos
                If EstaAgrupandoVtos Then
                    'ya hemos cambiado el importe para los dos apuntes que
                    'quedan abajo uno ajustando importe
                    
                Else
                    'Creo una linea de ap
                    SQL = Linea & ",'" & Cuenta & "','" & Format(vId, "000000000") & "',"
                    
                    
                    If ImporteVto > 0 Then
                        'al debe o al haber
                        SQL = SQL & FP.condecli & ",'" & DevNombreSQL(FP.CadenaAuxiliar & " Elim. " & vId) & "'," & TransformaComasPuntos(CStr(ImporteVto)) & ",NULL,"
                    Else
                        SQL = SQL & FP.conhacli & ",'" & DevNombreSQL(FP.descformapago & " Elim." & vId) & "',NULL," & TransformaComasPuntos(Abs(ImporteVto)) & ","
                    End If
                    'Contra partida
                    SQL = SQL & "NULL,'" & CtaCancelacion & "','CONTAB',0)"
                    SQL = AmpRemesa & SQL
                    Ejecuta SQL
                    Linea = Linea + 1
                
                    
                End If 'EstaAgrupandoVtos
            End If  'ImporteVto <> ImporteDocumento
        End If  ' vId > 0

    
    
        If EstaAgrupandoVtos Then
            If CadenaAgrupacion <> "" Then
                'OK inserto el total
                Ampliacion = TransformaComasPuntos(CStr(Importe))
                SQL = Replace(CadenaAgrupacion, "@@@@@@", Ampliacion)
                Conn.Execute AmpRemesa & SQL
                Linea = 2 'La uno es
            End If
        End If
    
        If Linea > 1 Then
            'Hago el contrapunte
            If CuentaPuente Then
                Ampliacion = FP.descformapago & " Re: " & Codigo & " - " & Anyo
                SQL = "RE" & Format(Codigo, "0000") & Format(Anyo, "0000")
                SQL = Linea & ",'" & CtaCancelacion & "','" & SQL & "'," & FP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
                'Importe al DEBE
                SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL,"
        
                'Contra partida
                If CuentaUnica Then
                    Cuenta = "'" & RaizCuentasCancelacion & "'"
                Else
                    If Len(Cuenta) = vEmpresa.DigitosUltimoNivel And IsNumeric(Cuenta) Then
                        'Dejo cuenta como esta
                        Cuenta = "'" & Cuenta & "'"
                    Else
                        Cuenta = "NULL"
                    End If
                End If
                SQL = SQL & Cuenta & ",'CONTAB',0)"
                SQL = AmpRemesa & SQL
            
            
  
                Conn.Execute SQL
            Else
                Linea = Linea + 1 'Para que despues no de distinto el numero de efectos eliminados
            End If
        Else
            
            
            'Muestro un mesaje diciendo que Ningun vto ha sido eliminado. No deberia pasar pero por si acaso
            'compruebo que tenga vtos
            'If NumRegElim > 0 Then MsgBox "No se ha podido eliminar ningun vencimiento de la remesa " & Codigo & " / " & Anyo, vbInformation
            RemesasEliminarVtosTalonesPagares = 0
            
            If CuentaPuente Then
                SQL = "DELETE FROM hcabapu  WHERE numdiari =" & FP.diaricli
                SQL = SQL & " and fechaent = '" & Format(FechaAbono, FormatoFecha) & "' and numasien = " & Mc.Contador
                Conn.Execute SQL
            End If
            
                        
          

            
        End If


    'Si la hemos borrado toda, o no....
    Linea = Linea - 1 'Empieza en uno, luego el total vtos eliminados es linea-1
                      'En numregelim tengo los vtos totales de la remesa
                      'Si queda alguno o no, haremos unas cosas u otras
    If Linea > 0 Then
        If NumRegElim <> EfectosEliminados_ Then
            AmpRemesa = "Y"
        
           
        Else
            AmpRemesa = "Z"  'TOdos eliminados
        End If
        SQL = "UPDATE remesas SET"
        SQL = SQL & " situacion= '" & AmpRemesa
        SQL = SQL & "' WHERE codigo=" & Codigo
        SQL = SQL & " and anyo=" & Anyo
        Conn.Execute SQL
        
    End If
   ' If CuentaPuente Then InsertaTmpActualizar Mc.Contador, FP.diaricli, FechaAbono
    
    'Todo OK
    If NumRegElim > 0 Then
        RemesasEliminarVtosTalonesPagares = 1
        'Para que no actualice el apunte , ya que no se ha creado
        If Not CuentaPuente Then RemesasEliminarVtosTalonesPagares = 0
    End If
    
ERemesa_Elivto:
    If Err.Number <> 0 Then
        
        MuestraError Err.Number, Err.Description
        RemesasEliminarVtosTalonesPagares = 2
    End If
    Set Rs = Nothing
    Set Mc = Nothing

End Function


































'********************* ********************* ********************* *********************
'********************* ********************* ********************* *********************
'********************* ********************* ********************* *********************
'
'               CANCELACION CUENTA CAARTERA. Pondra la situacion en H
'
'********************* ********************* ********************* *********************
'********************* ********************* ********************* *********************
'********************* ********************* ********************* *********************

' Cancela   Debe     Haber
'           4311     4310
'
Public Function CancearCuentaPuenteContraCartera(Fecha As Date, Codrem As Long, Anyorem As Integer, TipoRemesa As Byte) As Boolean

    CancearCuentaPuenteContraCartera = False
    
    Conn.BeginTrans
    
    
    If CancelarCuentaPuenteContraCarteraApunte(Fecha, Codrem, Anyorem, TipoRemesa) Then
        CancearCuentaPuenteContraCartera = True
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
    
    
    
    
    
End Function




Private Function CancelarCuentaPuenteContraCarteraApunte(Fecha As Date, Codigo As Long, Anyo As Integer, TipoRemesa As Byte) As Boolean

    
    

'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaParametros  As String
Dim Cuenta As String
Dim CuentaPuente As Boolean

'Dim ImporteTalonPagare As Currency    'beneficiosPerdidasTalon: por si hay diferencias entre vtos y total talon
Dim Aux As String

Dim LCta As Integer

Dim Obs As String
Dim ParBD As String
    On Error GoTo eCancearCuentaPuenteContraCarteraApunte
    CancelarCuentaPuenteContraCarteraApunte = False




    'La forma de pago
    CuentaPuente = True
    Set vCP = New Ctipoformapago
    If TipoRemesa = 2 Then
        Linea = vbPagare
        Cuenta = "pagarecta"
        
        
        
        
    ElseIf TipoRemesa = 5 Then
        Linea = vbConfirming
       Err.Raise 513, , "Error leyendo datos cta. puente confirming"
        
    Else
        Linea = vbTalon
        Cuenta = "taloncta"
        
    End If
    
    Cuenta = DevuelveDesdeBD(Cuenta, "paramtesor", "1", "1")
    
    
    LCta = Len(Cuenta)
    
    
     'Estamos canceland 4311(debe ) contra 43100   (haber
     'INSERT IGNORE INTO
     If LCta <> vEmpresa.DigitosUltimoNivel Then
        SQL = "concat('4311',substring(cobros.codmacta,5)) , nommacta,'S'"
        
        
        SQL = "Select " & SQL & " from cobros,cuentas WHERE cobros.codmacta=cuentas.codmacta and codrem=" & Codigo & " AND anyorem = " & Anyo
        SQL = "INSERT IGNORE INTO cuentas(codmacta,nommacta,apudirec) " & SQL
        Conn.Execute SQL
        
        SQL = "concat('4311',substring(cobros.codmacta,5)) , nommacta,'S'"
        
        
        SQL = "Select " & SQL & " from cobros,cuentas WHERE cobros.codmacta=cuentas.codmacta and codrem=" & Codigo & " AND anyorem = " & Anyo
        SQL = "INSERT IGNORE INTO cuentas(codmacta,nommacta,apudirec) " & SQL
        Conn.Execute SQL
        
        
    End If
            
    If vCP.Leer(Linea) = 1 Then Err.Raise 513, , "Leyendo formas de pago"
    
    
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", Fecha <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(Fecha, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Cancelar cartera efectos cta puente remesa: " & Codigo & " / " & Anyo & "   " & Cuenta & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "',"
    
    Obs = Codigo & " / " & Anyo & "   " & Cuenta
    
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Canc. efect.carter: " & Obs & "');"
    
    Conn.Execute SQL
    
    
    Linea = 1
    Importe = 0
    
    Set Rs = New ADODB.Recordset
    
    
    
    
    'La ampliacion para el banco
    AmpRemesa = ""
    SQL = "Select * from remesas WHERE codigo=" & Codigo & " AND anyo = " & Anyo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'NO puede ser EOF
    
    Importe = Rs!Importe

    If Not IsNull(Rs!Descripcion) Then AmpRemesa = Rs!Descripcion
    
    
    If AmpRemesa = "" Then
        AmpRemesa = " Remesa: " & Codigo & "/" & Anyo
    Else
        AmpRemesa = " " & AmpRemesa
    End If
    
    Rs.Close
    
    
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(Fecha, FormatoFecha) & "'," & Mc.Contador & ","
    ParBD = SQL
    
    
    Importe = 0
    SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
    
        SQL = ParBD & Linea & ","
    
        'Cuenta
        If LCta = vEmpresa.DigitosUltimoNivel Then
            CtaParametros = Cuenta
            Aux = "4310" & Mid(Cuenta, 5)
            Err.Raise 513, , "Falta cuenta . Ver codigo fuente"
        Else
            CtaParametros = "4311" & Mid(Rs!codmacta, 5)
            Aux = "4310" & Mid(Rs!codmacta, 5)
        End If
        SQL = SQL & DBSet(CtaParametros, "T")
        
        SQL = SQL & ",'" & Rs!NUmSerie & Format(Rs!NumFactu, "0000000") & "'," & vCP.conhacli
        
        
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        Ampliacion = Ampliacion & " vto: "
               
        'Neuvo dato para la ampliacion en la contabilizacion
        
        Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
        
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 40)) & "',"
        
            

        'de timporteH, codccost, ctacontr, idcontab, punteada
        'Importe
        SQL = SQL & "" & TransformaComasPuntos(Rs!ImpVenci) & ",NULL,NULL,"
    

        SQL = SQL & "'" & Aux & "',"

        SQL = SQL & "'COBROS',0,"
        SQL = SQL & "null,null ,null,null,null,null,null,null)"
        
        Conn.Execute SQL
        
        Linea = Linea + 1
        
        SQL = ParBD & Linea & ","
    
        SQL = SQL & DBSet(Aux, "T")
        
        SQL = SQL & ",'" & Rs!NUmSerie & Format(Rs!NumFactu, "0000000") & "'," & vCP.conhacli
        
        
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        Ampliacion = Ampliacion & " vto: "
               
        'Neuvo dato para la ampliacion en la contabilizacion
        
        Ampliacion = Ampliacion & Format(Rs!FecVenci, "dd/mm/yyyy")
        
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 40)) & "',"
        
            

        ' timporteH, codccost, ctacontr, idcontab, punteada
        'Importe
        SQL = SQL & "NULL," & TransformaComasPuntos(Rs!ImpVenci) & ",NULL,"
    

        SQL = SQL & "'" & CtaParametros & "',"

        SQL = SQL & "'COBROS',0,"
        SQL = SQL & "null,null ,null,null,null,null,null,null)"
        
        Conn.Execute SQL
        
        
        Linea = Linea + 1
        
        Rs.MoveNext
    
    Wend
    Rs.Close
        
    
    
    
    

    'AHora actualizamos los efectos.
    SQL = "UPDATE cobros SET"
    SQL = SQL & " siturem= 'H'"
    SQL = SQL & ", situacion = 1 "

    SQL = SQL & " WHERE codrem=" & Codigo
    SQL = SQL & " and anyorem=" & Anyo
    SQL = SQL & " and tiporem = " & TipoRemesa
    
    
    
    Conn.Execute SQL
    
    SQL = "UPDATE remesas SET "
    SQL = SQL & " situacion= 'H'"
    SQL = SQL & " WHERE codigo=" & Codigo
    SQL = SQL & " and anyo=" & Anyo
    SQL = SQL & " and tiporem = " & TipoRemesa
    
    Conn.Execute SQL

    

    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, Fecha
    
    'Todo OK
    CancelarCuentaPuenteContraCarteraApunte = True



    
eCancearCuentaPuenteContraCarteraApunte:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set Rs = Nothing
End Function







'-----------------------------------
Public Function RealizarLaDevolucionTransferencia(EsCobro As Boolean, FechaDevolucion As Date, ContabilizoGastoBanco As Boolean, CtaBenBancarios As String, Remesa As String, DatosContabilizacionDevolucion As String) As Boolean

'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Rs As ADODB.Recordset
Dim Amp11 As String
Dim DescRemesa As String
'Dim AgrupaApunteBanco As Boolean
Dim GastoDevolucion As Currency
Dim DescuentaImporteDevolucion As Boolean
Dim GastoVto As Currency
Dim Gastos As Currency  'de cada recibo/vto
Dim Aux As String
Dim Importeauxiliar As Currency
Dim CtaBancoGastos As String
Dim CCBanco As String
Dim LinApu As String

Dim TipoAmpli As Integer


    On Error GoTo ECon
    RealizarLaDevolucionTransferencia = False
    
   
    'La forma de pago
    Set vCP = New Ctipoformapago
    Set Rs = New ADODB.Recordset
    
    
    DescRemesa = ""
    Aux = RecuperaValor(Remesa, 8)
    If Aux <> "" Then
        'OK viene de fichero
        Aux = RecuperaValor(Remesa, 9)
        'Vuelvo a susitiuri los # por |
        Aux = Replace(Aux, "#", "|")
        SQL = ""
        For Linea = 1 To Len(Aux)
            If Mid(Aux, Linea, 1) = "·" Then SQL = SQL & "X"
        Next
        
        If Len(SQL) > 1 Then
            'Tienen mas de una remesa
            SQL = ""
            While Aux <> ""
                Linea = InStr(1, Aux, "·")
                If Linea = 0 Then
                    Aux = ""
                Else
                    SQL = SQL & ",    " & Format(RecuperaValor(Mid(Aux, 1, Linea - 1), 1), "000") & "/" & RecuperaValor(Mid(Aux, 1, Linea - 1), 2) & ""
                    Aux = Mid(Aux, Linea + 1)
                End If
            
            Wend
            Aux = RecuperaValor(Remesa, 8)
            SQL = "Devolución remesas: " & Trim(Mid(SQL, 2))
            DescRemesa = SQL & vbCrLf & "Fichero: " & Aux
        End If
        
    End If

    
    
    DescRemesa = RecuperaValor(Remesa, 9)

   
    
    
    If vCP.Leer(vbTransferencia) = 1 Then GoTo ECon


    'Los parametros de contbilizacion se le pasan en el frame de pedida de datos
    'Ahora se los asignaremos a la formma de pago
    vCP.condecli = RecuperaValor(DatosContabilizacionDevolucion, 1)
    vCP.ampdecli = RecuperaValor(DatosContabilizacionDevolucion, 2)
    vCP.conhacli = RecuperaValor(DatosContabilizacionDevolucion, 1) '3)
    vCP.amphacli = RecuperaValor(DatosContabilizacionDevolucion, 2) '4)
    SQL = RecuperaValor(DatosContabilizacionDevolucion, 5)  'agupa o no
  
    
    SQL = RecuperaValor(Remesa, 7)
    GastoDevolucion = TextoAimporte(SQL)
    DescuentaImporteDevolucion = False
    If GastoDevolucion > 0 Then
        CtaBancoGastos = "CtaIngresos"
        SQL = RecuperaValor(Remesa, 3)
        SQL = DevuelveDesdeBD("GastRemDescontad", "bancos", "codmacta", SQL, "T")
        If SQL = "1" Then DescuentaImporteDevolucion = True
    End If
    
    'Datos del banco
    SQL = RecuperaValor(Remesa, 3)
    SQL = "Select * from bancos where codmacta ='" & SQL & "'"
    CCBanco = ""
    CtaBancoGastos = ""

    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        SQL = "No se ha encontrado banco: " & vbCrLf & SQL
        Err.Raise 516, SQL
    End If
    CCBanco = DBLet(Rs!CodCcost, "T")
    CtaBancoGastos = DBLet(Rs!ctagastos, "T")
    If Not vParam.autocoste Then CCBanco = ""  'NO lleva analitica
    Rs.Close
    
  
    
    'EMPEZAMOS
    'Borramos tmpactualizar
    SQL = "DELETE FROM tmpactualizar where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    'Cargaremos los registros a devolver que estaran en la tabla temporal
    'para codusu
    SQL = "Select * from tmpfaclin where codusu=" & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "EOF.  NO se han cargado datos devolucion", vbExclamation
        Rs.Close
        GoTo ECon
    End If

    Set Mc = New Contadores


    If Mc.ConseguirContador("0", FechaDevolucion <= vParam.fechafin, True) = 1 Then GoTo ECon


    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ",'"
    
    'Ahora esta en desc remesa
    DescRemesa = DescRemesa & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:nn") & " por " & vUsu.Nombre
    SQL = SQL & DevNombreSQL(DescRemesa) & "',"
    'SQL = SQL & "'Devolucion remesa: " & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2)
    'SQL = SQL & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "')"
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Devolución transferencia')"

    
    If Not Ejecuta(SQL) Then GoTo ECon


    Linea = 1
    Importe = 0

    If vCP.ampdecli = 3 Then
        Amp11 = DescRemesa
    Else
        Amp11 = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    End If
    
    'Lo meto en una VAR
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, "
    SQL = SQL & " numserie," & IIf(EsCobro, "numfaccl", "numfacpr")
    SQL = SQL & ",fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem,esdevolucion) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
    LinApu = SQL
    
    While Not Rs.EOF

        'Lineas de apuntes .
        SQL = LinApu & Linea & ",'"
        SQL = SQL & Rs!Cta
        TipoAmpli = vCP.ampdecli
        If EsCobro Then
            SQL = SQL & "','" & Rs!NUmSerie & Format(Rs!NumFac, "0000000") & "'," & vCP.condecli
        Else
            SQL = SQL & "'," & DBSet(Rs!nomserie, "T") & "," & vCP.conhacli
        End If
        Ampliacion = Amp11 & " "
    
        Select Case TipoAmpli
        Case 3
            'NUEVA forma de ampliacion
            'No hacemos nada pq amp11 ya lleva lo solicitado
            
        Case 4
                'COntrapartida
                Ampliacion = Ampliacion & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!Cta, "T")
                
        Case 6
                Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Rs!Cta, "T")
                
               
                If EsCobro Then
                    MiVariableAuxiliar = Rs!NUmSerie & Format(Rs!NumFac, "0000000")
        
                Else
                    MiVariableAuxiliar = Rs!nomserie
                End If
                Ampliacion = Mid(Ampliacion, 1, 34 - Len(MiVariableAuxiliar))
                Ampliacion = Ampliacion & " " & MiVariableAuxiliar
                
        Case Else
                If TipoAmpli = 2 Then
                   Ampliacion = Ampliacion & Format(Rs!Fecha, "dd/mm/yyyy")
                Else
                   If TipoAmpli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                   
                   If EsCobro Then
                        Ampliacion = Ampliacion & Rs!NUmSerie & Format(Rs!NumFac, "0000000")
                   Else
                        Ampliacion = Ampliacion & Rs!nomserie
                   End If
                End If
           
        End Select
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 35)) & "',"

        Importe = Importe + Rs!Imponible

    
        GastoVto = 0
        If EsCobro Then
            Aux = " numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
            Aux = Aux & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden"
            Aux = DevuelveDesdeBD("gastos", "cobros", Aux, CStr(Rs!NIF), "N")
            
            If Aux <> "" Then GastoVto = CCur(Aux)
        End If
        
        
        Gastos = Gastos + GastoVto

        ' timporteH, codccost, ctacontr, idcontab, punteada
        Importeauxiliar = Rs!Imponible - GastoVto
        
        SQL = SQL & "NULL," & TransformaComasPuntos(CCur(Importeauxiliar)) & ",NULL,"
        
        'Contrapartida
        SQL = SQL & "'" & Rs!Cliente & "',"
        
        If EsCobro Then
            SQL = SQL & "'COBROS',0,"
        Else
            SQL = SQL & "'PAGOS',0,"
        End If
        '%%%%% aqui van todos los datos de la devolucion en la linea de cuenta
        SQL = SQL & DBSet(Rs!NUmSerie, "T") & "," & IIf(EsCobro, DBSet(Rs!NumFac, "N"), DBSet(Rs!nomserie, "T"))
        SQL = SQL & "," & DBSet(Rs!Fecha, "F") & "," & DBSet(Rs!NIF, "N") & ","
            
         '-------------------------------------------------------------------------------------
         'Ahora
         '-------------------------------------------------------------------------------------
         'Lo pongo en uno
             'Actualizamos el registro. Ponemos la marca de devuelto. Y aumentamos el importe de gastos
         'Si es que hay
         Dim SqlCobro As String
         Dim RsCobro As ADODB.Recordset
         Dim ImporteNue As Currency
         
         
         Set RsCobro = New ADODB.Recordset
         
         If EsCobro Then
                 SqlCobro = "select tipforpa, tiporem, codrem, anyorem, gastos from cobros inner join formapago on cobros.codforpa = formapago.codforpa "
                 SqlCobro = SqlCobro & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
                 SqlCobro = SqlCobro & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
                 
                 Set RsCobro = New ADODB.Recordset
                 RsCobro.Open SqlCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                 If Not RsCobro.EOF Then
                 
        
                    SQL = SQL & DBSet(RsCobro!TipForpa, "N") & "," & DBSet(FechaDevolucion, "F") & "," & DBSet(Rs!CtaBase, "T", "S") & ","
                    SQL = SQL & DBSet(Rs!ImpIva, "N") & "," & DBSet(RsCobro!Tiporem, "N") & "," & DBSet(RsCobro!Codrem, "N") & "," & DBSet(RsCobro!Anyorem, "N") & ",1)"
                      
                 
                    Ampliacion = "UPDATE cobros SET "
                    Ampliacion = Ampliacion & " Devuelto = 1, situacion = 0   "
                    
                    
                    ImporteNue = DBLet(RsCobro!Gastos, "N")
                    If DBLet(Rs!ImpIva, "N") > 0 Then
                    
                        If ImporteNue = 0 Then
                            Ampliacion = Ampliacion & " , Gastos = " & TransformaComasPuntos(CStr(Rs!ImpIva))
                        Else
                            Ampliacion = Ampliacion & " , Gastos = Gastos + " & TransformaComasPuntos(CStr(Rs!ImpIva))
                        End If
                    End If
                    Ampliacion = Ampliacion & " ,impcobro=NULL,codrem= NULL, anyorem = NULL, siturem = NULL,tiporem=NULL,fecultco=NULL,recedocu=0,transfer=null"
                    Ampliacion = Ampliacion & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & Rs!NumFac
                    Ampliacion = Ampliacion & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
                    
                    Ejecuta Ampliacion
                End If
         Else
                     
                    'PAGOS
                   SqlCobro = "select tipforpa, null tiporem,nrodocum codrem, anyodocum anyorem,0 gastos from pagos inner join formapago on pagos.codforpa = formapago.codforpa "
                   SqlCobro = SqlCobro & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & DBSet(Rs!nomserie, "T")
                   SqlCobro = SqlCobro & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
                   SqlCobro = SqlCobro & " AND codmacta=" & DBSet(Rs!Cta, "T")
                   
        
                    RsCobro.Open SqlCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not RsCobro.EOF Then
                    
           
                       SQL = SQL & DBSet(RsCobro!TipForpa, "N") & "," & DBSet(FechaDevolucion, "F") & "," & DBSet(Rs!CtaBase, "T", "S") & ","
                       SQL = SQL & DBSet(Rs!ImpIva, "N") & ",null," & DBSet(RsCobro!Codrem, "N") & "," & DBSet(RsCobro!Anyorem, "N") & ",1)"
                         
                    
                       Ampliacion = "UPDATE pagos SET "
                       Ampliacion = Ampliacion & "  situacion = 0,"
                       Ampliacion = Ampliacion & " fecultpa =null, imppagad=null,nrodocum= NULL, anyodocum = NULL, situdocum = NULL"
                       Ampliacion = Ampliacion & " WHERE numserie='" & Rs!NUmSerie & "' AND numfactu=" & DBSet(Rs!nomserie, "T")
                       Ampliacion = Ampliacion & " AND fecfactu='" & Format(Rs!Fecha, FormatoFecha) & "' AND numorden=" & Rs!NIF
                       Ampliacion = Ampliacion & " AND codmacta=" & DBSet(Rs!Cta, "T")
                       Ejecuta Ampliacion
                    End If
                
         End If
         Set RsCobro = Nothing

        '%%%%% hasta aqui
        

        If Not Ejecuta(SQL) Then GoTo ECon

        Linea = Linea + 1
        
        
        
        'Gasto.
        ' Si tiene y no agrupa
        '-------------------------------------------------------
        If GastoVto > 0 Then
            'Err.Raise 513, , "Error en codigo. Parametro tipo remesa incorrecto. "
           'Lineas de apuntes .
            SQL = LinApu & Linea & ",'"
    
    
            SQL = SQL & CtaBancoGastos & "','" & Rs!NUmSerie & Format(Rs!NumFac, "0000000") & "'," & vCP.condecli
            SQL = SQL & ",'Gastos vto.'"
            
            
            'Importe al debe
            SQL = SQL & "," & TransformaComasPuntos(CStr(GastoVto)) & ",NULL,"
            
            If CCBanco <> "" Then
                SQL = SQL & "'" & DevNombreSQL(CCBanco) & "'"
            Else
                SQL = SQL & "NULL"
            End If
                
            'Contra partida
            'Si no lleva cuenta puente contabiliza los gastos
            Aux = "NULL"
           
            SQL = SQL & "," & Aux & ",'COBROS',0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
        
        
        End If
        
        
        
        
        Rs.MoveNext
    Wend
    
    
    Rs.MoveFirst



    'La linea del banco
    '*********************************************************************
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","

    'Nuevo tipo ampliacion
    TipoAmpli = vCP.ampdecli
    If TipoAmpli = 3 Then
        Ampliacion = DescRemesa
    Else
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    End If
    
    
    Aux = Trim(RecuperaValor(Remesa, 2))
    If Aux = "" Then
        Aux = RecuperaValor(Remesa, 5) 'devuelve una cuebnta
        TipoAmpli = InStr(1, Aux, ":")
        If TipoAmpli > 0 Then Aux = Mid(Aux, TipoAmpli + 1)
    Else
        Aux = Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
    End If
    Ampliacion = Ampliacion & " Dev.rem:" & Aux
    
    Amp11 = Rs!Cliente  'cta banco

    'Lleva gasto pero lo descontamos de aqui
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Importe = Importe + GastoDevolucion
        'Para que la linea salga al fina
        Linea = Linea + 2
    End If
    Aux = Trim(RecuperaValor(Remesa, 2))
    If Aux = "" Then
        Aux = "DEV. TRANSFERENCIA"
    Else
        Aux = "TRANS" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2)
    End If
    Ampliacion = Linea & ",'" & Amp11 & "'," & DBSet(Aux, "T") & "," & vCP.condecli & ",'" & Ampliacion & "',"
   
    Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Importe)) & ",NULL"
   
    Ampliacion = Ampliacion & ",NULL,NULL"
    Ampliacion = Ampliacion & ",'" & IIf(EsCobro, "COBROS", "PAGOS") & "',0)"
    Ampliacion = SQL & Ampliacion
    If Not Ejecuta(Ampliacion) Then GoTo ECon
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Linea = Linea - 2
        'Dejo el importe como estaba
        Importe = Importe - GastoDevolucion
    Else
        Linea = Linea + 1
    End If
    
    
    'SI hay que contabilizar los gastos de devolucion
    If ContabilizoGastoBanco Then
        
         If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
         Else
            Linea = Linea + 1
         End If
         SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
         SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
         SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
         SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"

         'Cuenta
         SQL = SQL & CtaBenBancarios & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.condecli
         'SQL = SQL & Rs!Cta & "','REM" & Format(Rs!numfac, "000000000") & "'," & vCP.condecli
        

        If vCP.ampdecli = 3 Then
            Ampliacion = DescRemesa
        Else
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"


        ' timporteH, codccost, ctacontr, idcontab, punteada
        'Importe.  Va al debe
        SQL = SQL & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,"
        
        'Centro de coste.
        '--------------------------
        Amp11 = "NULL"
        If vParam.autocoste Then
            Amp11 = DevuelveDesdeBD("codccost", "bancos", "codmacta", Rs!Cliente, "T")
            Amp11 = "'" & Amp11 & "'"
        End If
        SQL = SQL & Amp11 & ","

        
        SQL = SQL & "'" & Rs!Cliente & "',"
        SQL = SQL & "'COBROS',0)"

        If Not Ejecuta(SQL) Then GoTo ECon

        
        
    
        'El total del banco..
        
        'La linea del banco
        'Rs.MoveFirst
        'Si no agrupa dto importe
        If Not DescuentaImporteDevolucion Then
            Linea = Linea + 1
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
        
            
            If vCP.amphacli = 3 Then
                Ampliacion = DescRemesa
            Else
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
                Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
            End If
            
            Ampliacion = Linea & ",'" & Rs!Cliente & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
            'Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,'" & CtaBenBancarios & "','CONTAB',0)"
            Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,'" & CtaBenBancarios & "','COBROS',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then GoTo ECon
            
        End If
      
    
    End If

    'Ya tenemos generado el apunte de devolucion
    'Insertamos para su actualziacion
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaDevolucion
    
    
    RealizarLaDevolucionTransferencia = True
ECon:
    If Err.Number <> 0 Then
        
        Amp11 = "Devolución transferemcoa: " & Remesa & vbCrLf
        If Not Mc Is Nothing Then Amp11 = Amp11 & "MC.cont: " & Mc.Contador & vbCrLf
        Amp11 = Amp11 & Err.Description
        MuestraError Err.Number, Amp11
        
    End If
    Set Rs = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function

