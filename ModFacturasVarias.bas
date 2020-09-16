Attribute VB_Name = "ModFacturasVarias"
'===================================================================================
'CONTABILIZAR FACTURAS:
'Modulo para el traspaso de registros de cabecera y lineas de tablas de FACTURACION
'A las tablas de FACTURACION de Contabilidad
'====================================================================================

Private BaseImp As Currency
Private IvaImp As Currency

Private CCoste As String


Private vTipoIva(2) As Currency
Private vPorcIva(2) As Currency
Private vPorcRec(2) As Currency
Private vBaseIva(2) As Currency
Private vImpIva(2) As Currency
Private vImpRec(2) As Currency




Public Sub RecalculoBasesIvaFactura(ByRef Rs As ADODB.Recordset, ByRef Imptot As Variant, ByRef Tipiva As Variant, ByRef Impbas As Variant, ByRef ImpIva As Variant, ByRef PorIva As Variant, ByRef TotFac As Currency, ByRef ImpRec As Variant, ByRef PorRec As Variant, ByRef PorRet As Variant, ByRef ImpRet As Variant, ByRef Tipo As Integer)

    Dim i As Integer
    Dim Sql As String
    Dim Baseimpo As Dictionary
    Dim CodIVA As Integer

    Set Baseimpo = New Dictionary

    ' inicializamos los importes de los totales de la cabecera
    TotFac = 0
    totimp = 0
    Base = 0
    ImpRet = 0
    For i = 0 To 2
         Tipiva(i) = 0
         Imptot(i) = 0
         Impbas(i) = 0
         ImpIva(i) = 0
         PorIva(i) = 0
         PorRec(i) = 0
         ImpRec(i) = 0
    Next i

    ' recorremos todas las lineas de la factura
    If Not Rs.EOF Then Rs.MoveFirst
    While Not Rs.EOF
        CodIVA = DBLet(Rs!TipoIva, "N") ' DevuelveDesdeBDNewFac("tiposiva", "codigiva", "sartic", "codartic", DBLet(RS!codartic), "N")
        Baseimpo(Val(CodIVA)) = DBLet(Baseimpo(Val(CodIVA)), "N") + DBLet(Rs!Importe, "N")

        Rs.MoveNext
    Wend

    For i = 0 To Baseimpo.Count - 1
        If i <= 2 Then
            Tipiva(i) = Baseimpo.keys(i)
            Impbas(i) = Baseimpo.Items(i)
 
            PorIva(i) = DevuelveDesdeBD("porceiva", "tiposiva", "codigiva", CStr(Tipiva(i)), "N")
            PorRec(i) = DevuelveDesdeBD("porcerec", "tiposiva", "codigiva", CStr(Tipiva(i)), "N")
            ImpIva(i) = DBLet(Round2(Impbas(i) * PorIva(i) / 100, 2), "N")
            ImpRec(i) = DBLet(Round2(Impbas(i) * PorRec(i) / 100, 2), "N")
            Imptot(i) = Impbas(i) + ImpIva(i) + ImpRec(i)
            TotFac = TotFac + Imptot(i)
 
'antes el iva estaba incluido
'            PorIva(i) = DevuelveDesdeBDNewFac(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
'            Impbas(i) = Round2(Imptot(i) / (1 + (PorIva(i) / 100)), 2)
'            impiva(i) = Imptot(i) - Impbas(i)
'            TotFac = TotFac + Imptot(i)
        
        
        End If
    Next i
    'si hay retencion la calculamos
    If PorRet <> 0 Then
        Base = 0
        For i = 0 To Baseimpo.Count - 1
            Base = Base + Impbas(i)
        Next i
        '[Monica]06/08/2020: no lo calculaba nunca sobre base + iva
        If Tipo = 2 Then
            For i = 0 To Baseimpo.Count - 1
                Base = Base + ImpIva(i)
            Next i
        End If
        ImpRet = Round2(Base * PorRet / 100, 2)
        TotFac = TotFac - ImpRet
    Else
        ImpRet = 0
    End If
End Sub

' ### [Monica] 27/09/2006
Public Function FacturaContabilizada(NUmSerie As String, numfactu As String, Anofactu As String) As Boolean
Dim Sql As String
Dim NumAsi As Currency

    FacturaContabilizada = False
    Sql = ""
    Sql = DevuelveDesdeBDNew(cConta, "factcli", "numserie", Trim(NUmSerie), "T", , "numfactu", numfactu, "N", "anofactu", Anofactu, "N")
    
    If Sql = "" Then Exit Function
    
    NumAsi = DBLet(Sql, "N")
    
    If NumAsi <> 0 Then FacturaContabilizada = True

End Function

' ### [Monica] 27/09/2006
Public Function FacturaRemesada(NUmSerie As String, numfactu As String, FecFactu As String) As Boolean
Dim Sql As String
Dim NumRem As Currency

    FacturaRemesada = False
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cConta, "cobros", "codrem", "numserie", Trim(NUmSerie), "T", , "numfactu", numfactu, "N", "fecfactu", FecFactu, "F")
    
    If Sql = "" Then Exit Function
    
    NumRem = DBLet(Sql, "N")
    
    If NumRem <> 0 Then FacturaRemesada = True
    
End Function

' ### [Monica] 27/09/2006
Public Function FacturaCobrada(NUmSerie As String, numfactu As String, FecFactu As String) As Boolean
Dim Sql As String
Dim ImpCob As Currency

    FacturaCobrada = False
    Sql = ""
    Sql = DevuelveDesdeBDNew(cConta, "cobros", "impcobro", "numserie", Trim(NUmSerie), "T", , "numfactu", numfactu, "N", "fecfactu", FecFactu, "F")
    If Sql = "" Then Exit Function
    ImpCob = DBLet(Sql, "N")
    
    If ImpCob <> 0 Then FacturaCobrada = True
    
End Function

Public Sub BorrarTMPFacturas()
On Error Resume Next

    Conn.Execute " DROP TABLE IF EXISTS tmpfactu;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function CrearTMPFacturas(cadTABLA As String, cadwhere As String, Optional facturas As Boolean, Optional Telefono As Boolean) As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
' facturas indica si viene de facturas varias o de telefonia
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPFacturas = False
    
    Sql = "CREATE TEMPORARY TABLE tmpfactu ( "
    Sql = Sql & "numserie char(3) NOT NULL default '',"
    Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00') "
    Conn.Execute Sql
     
    Sql = "SELECT numserie, numfactu, fecfactu"
    Sql = Sql & " FROM " & cadTABLA
    Sql = Sql & " WHERE " & cadwhere
    Sql = " INSERT INTO tmpfactu " & Sql
    Conn.Execute Sql

    CrearTMPFacturas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPFacturas = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpfactu;"
        Conn.Execute Sql
    End If
End Function

Public Sub BorrarTMPErrFact()
On Error Resume Next
    Conn.Execute " DROP TABLE IF EXISTS tmperrfac;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BorrarTMPErrComprob()
On Error Resume Next
    Conn.Execute " DROP TABLE IF EXISTS tmperrcomprob;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function CrearTMPErrComprob() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrComprob = False
    
    Sql = "CREATE TEMPORARY TABLE tmperrcomprob ( "
    Sql = Sql & "error varchar(100) NULL )"
    Conn.Execute Sql
     
    CrearTMPErrComprob = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrComprob = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmperrcomprob;"
        Conn.Execute Sql
    End If
End Function


Public Function ComprobarNumFacturasFacContaNueva(cadWConta) As Boolean
'Comprobar que no exista ya en la contabilidad un nº de factura para la fecha que
'vamos a contabilizar
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean

    On Error GoTo ECompFactuFac

    ComprobarNumFacturasFacContaNueva = False
    
    Sql = "SELECT numserie,numfactu,anofactu FROM factcli "
    Sql = Sql & " WHERE " & cadWConta
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos las distintas facturas que vamos a facturar
        Sql = "SELECT DISTINCT tmpfactu.numserie,tmpfactu.numfactu,tmpfactu.fecfactu "
        Sql = Sql & " FROM tmpfactu "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        B = True
        While Not Rs.EOF 'And b
            Sql = ""
            Sql = DevuelveDesdeBDNew(cConta, "factcli", "numfactu", "numfactu", Rs!numfactu, "N", , "numserie", Trim(Rs!NUmSerie), "T", "anofactu", Year(Rs!FecFactu), "N")
            If Sql <> "" Then
                B = False
                Sql = "          Nº Fac.: " & Format(Rs!numfactu, "0000000") & vbCrLf
                Sql = Sql & "          Fecha: " & Rs!FecFactu
                
                Sql = "Ya existe la factura: " & vbCrLf & Sql
                InsertarError Sql
            
            End If
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
        
        If Not B Then
            Sql = "Ya existe la factura: " & vbCrLf & Sql
            Sql = "Comprobando Nº Facturas en Contabilidad...       " & vbCrLf & vbCrLf & Sql
            
            'MsgBox sql, vbExclamation
            ComprobarNumFacturasFacContaNueva = False
        Else
            ComprobarNumFacturasFacContaNueva = True
        End If
    Else
        ComprobarNumFacturasFacContaNueva = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompFactuFac:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Nº Facturas Varias", Err.Description
    End If
End Function

Public Function ComprobarCtaContableFac(Opcion As Byte, Optional cadwhere As String) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim cadG As String
Dim enc As String
    
    On Error GoTo ECompCta

    ComprobarCtaContableFac = False
    
    Sql = "SELECT codmacta FROM cuentas "
    Sql = Sql & " WHERE apudirec='S'"
    If cadG <> "" Then Sql = Sql & cadG
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, Conn, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        If Opcion = 1 Then
                Sql = "SELECT DISTINCT fvarfactura.codmacta, fvarfactura.numfactu   "
                Sql = Sql & " FROM fvarfactura  "
                Sql = Sql & " where " & cadwhere
        ElseIf Opcion = 2 Then
                Sql = "SELECT distinct fvarconceptos.codmacta, fvarconceptos.codconce "
                Sql = Sql & " from fvarconceptos, fvarfactura_lineas, fvarfactura "
                Sql = Sql & " where " & cadwhere & " and fvarconceptos.codconce = fvarfactura_lineas.codconce"
                Sql = Sql & " and fvarfactura.numserie = fvarfactura_lineas.numserie "
                Sql = Sql & " and fvarfactura.numfactu = fvarfactura_lineas.numfactu "
                Sql = Sql & " and fvarfactura.fecfactu = fvarfactura_lineas.fecfactu "
        ElseIf Opcion = 3 Then
                'si hay analitica comprobar que todas las cuentas
                'empiezan por el digito que hay en conta.parametros.grupovta
                cadG = vParam.GrupoVta
        
                Sql = "SELECT distinct fvarconceptos.codconce "
                Sql = Sql & ", fvarconceptos.codmacta"
                Sql = Sql & " from ((fvarfactura_lineas "
                Sql = Sql & " INNER JOIN tmpfactu ON fvarfactura_lineas.numserie=tmpfactu.numserie AND fvarfactura_lineas.numfactu=tmpfactu.numfactu AND fvarfactura_lineas.fecfactu=tmpfactu.fecfactu) "
                Sql = Sql & " INNER JOIN fvarconceptos on fvarfactura_lineas.codconce = fvarconceptos.codconce) "
                Sql = Sql & " where fvarconceptos.codmacta "
                If cadG <> "" Then
                     Sql = Sql & " AND not (fvarconceptos.codmacta like '" & cadG & "%') "
                End If
        ElseIf Opcion = 4 Then
            B = True
            enc = ""
            enc = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", cadwhere, "T")
            If enc = "" Then
                B = False
                Sql = "No existe la cta contable de banco" & cadwhere
                InsertarError Sql
            End If
        ElseIf Opcion = 8 Then
            Sql = "SELECT DISTINCT fvarfactura.cuereten   "
            Sql = Sql & " FROM fvarfactura  "
            Sql = Sql & " where " & cadwhere
        End If
        If Opcion <> 4 Then
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            B = True
            While Not Rs.EOF 'And b
                If Opcion = 3 Then
                    Sql = Rs!codmacta
                    Sql = "La cuenta " & Sql & " del concepto " & Rs!CodConce & " no es del grupo correcto."
                    InsertarError Sql
                Else
                    Sql = "codmacta= " & DBLet(Rs.Fields(0).Value, "T") '& " and apudirec='S' "
                End If
                
                enc = ""
                If Opcion <> 8 Then
                    enc = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Rs.Fields(0).Value, "T")
                Else
                    If DBLet(Rs.Fields(0).Value, "T") <> "" Then
                        enc = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Rs.Fields(0).Value, "T")
                    End If
                End If
                     
                If enc = "" Then
                    If Opcion <> 8 Then B = False 'no encontrado
                    If Opcion = 1 Then
                            Sql = Rs!codmacta & " de la Factura " & Format(Rs!numfactu, "0000000")
                            Sql = "No existe la cta contable " & Sql
                            InsertarError Sql
                    End If
                    If Opcion = 2 Then
                        Sql = Rs!codmacta & " del Concepto " & Rs!CodConce
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 4 Then
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 5 Or Opcion = 6 Or Opcion = 7 Then
                        Sql = "No existe la cta contable " & Sql
                        InsertarError Sql
                    End If
                    If Opcion = 8 Then
                        Sql = DBLet(Rs!cuereten, "T")
                        If Sql <> "" Then
                            B = False
                            Sql = "No existe la cta contable de retención: " & Sql
                            InsertarError Sql
                        End If
                    End If
                End If
                    
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        End If
        If Not B Then
            ComprobarCtaContableFac = False
        Else
            ComprobarCtaContableFac = True
        End If
    Else
        ComprobarCtaContableFac = True
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompCta:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function

Public Function ComprobarTiposIVA() As Boolean
'Comprobar que todos los Tipos de IVA de las distintas facturas (scafac.codigiva1, codigiv2,codigiv3)
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSconta As ADODB.Recordset
Dim B As Boolean
Dim i As Byte

    On Error GoTo ECompIVA

    ComprobarTiposIVA = False
    
    Sql = "SELECT distinct codigiva FROM tiposiva "
    
    Set RSconta = New ADODB.Recordset
    RSconta.Open Sql, Conn, adOpenStatic, adLockPessimistic, adCmdText

    If Not RSconta.EOF Then
        'Seleccionamos los distintos tipos de IVA de las facturas a Contabilizar
        For i = 1 To 3
            Sql = "SELECT DISTINCT fvarfactura.tipoiva" & i
            Sql = Sql & " FROM fvarfactura "
            Sql = Sql & " INNER JOIN tmpfactu ON fvarfactura.numserie=tmpfactu.numserie AND fvarfactura.numfactu=tmpfactu.numfactu AND fvarfactura.fecfactu=tmpfactu.fecfactu "
            Sql = Sql & " WHERE not isnull(tipoiva" & i & ")"
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            B = True
            While Not Rs.EOF 'And b
                If Rs.Fields(0) <> 0 Then ' añadido pq en arigasol sino tiene tipo de iva pone ceros
                    Sql = "codigiva= " & DBSet(Rs.Fields(0), "N")
                    RSconta.MoveFirst
                    RSconta.Find (Sql), , adSearchForward
                    If RSconta.EOF Then
                        B = False 'no encontrado
                        Sql = "No existe el " & Sql
                        Sql = "Tipo de IVA: " & Rs.Fields(0)
                        InsertarError Sql
                    End If
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            Set Rs = Nothing
        
            If Not B Then
                ComprobarTiposIVA = False
                Exit For
            Else
                ComprobarTiposIVA = True
            End If
        Next i
    End If
    RSconta.Close
    Set RSconta = Nothing
    
ECompIVA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipo de IVA.", Err.Description
    End If
End Function


Public Function ComprobarCCoste() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim B As Boolean


    On Error GoTo ECCoste

    ComprobarCCoste = False
            
    Sql = "SELECT distinct fvarconceptos.codconce "
    Sql = Sql & ", fvarconceptos.codccost"
    Sql = Sql & " from ((fvarfactura_lineas "
    Sql = Sql & " INNER JOIN tmpfactu ON fvarfactura_lineas.numserie=tmpfactu.numserie AND fvarfactura_lineas.numfactu=tmpfactu.numfactu AND fvarfactura_lineas.fecfactu=tmpfactu.fecfactu) "
    Sql = Sql & " INNER JOIN fvarconceptos on fvarfactura_lineas.codconce = fvarconceptos.codconce) "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    B = True
    
    While Not Rs.EOF
       'comprobar que el Centro de Coste existe en la Contabilidad
       If DBLet(Rs.Fields(1).Value, "T") <> "" Then
            '[Monica]02/07/2019: faltaba cuando era con contabilidad nueva. CORREGIDO
            Sql = DevuelveDesdeBD("codccost", "ccoste", "codccost", Rs.Fields(1).Value, "T")
            If Sql = "" Then
                B = False
                Sql = "No existe el centro de coste: " & DBLet(Rs.Fields(1).Value, "T")
                Sql = Sql & " del concepto: " & DBLet(Rs.Fields(0).Value, "N")
                InsertarError Sql
            End If
       Else
            B = False
            Sql = "El concepto: " & DBLet(Rs.Fields(0).Value, "N")
            Sql = Sql & " no tiene centro de coste asociado. "
            InsertarError Sql
       End If
       Rs.MoveNext
    Wend
    
    ComprobarCCoste = B
    Set Rs = Nothing
    
ECCoste:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Centros de Coste", Err.Description
    End If
End Function

Public Function CrearTMPErrFact(cadTABLA As String) As Boolean
'Crea una temporal donde insertara la clave primaria de las
'facturas erroneas al facturar
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPErrFact = False
    
    Sql = "CREATE TEMPORARY TABLE tmperrfac ( "
    Sql = Sql & "numserie char(3) NOT NULL default '',"
    Sql = Sql & "numfactu mediumint(7) unsigned NOT NULL default '0',"
    Sql = Sql & "fecfactu date NOT NULL default '0000-00-00', "
    Sql = Sql & "error varchar(400) NULL )"
    Conn.Execute Sql
     
    CrearTMPErrFact = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErrFact = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmperrfac;"
        Conn.Execute Sql
    End If
End Function


Public Function PasarFacturaFac(cadwhere As String, FecVenci As String, CtaBanco As String, CodCCost As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariagroutil.cabfact --> conta.cabfact
' ariagroutil.linfact --> conta.linfact
'Actualizar la tabla ariagroutil.cabfact.inconta=1 para indicar que ya esta contabilizada

Dim B As Boolean
Dim cadMen As String
Dim Sql As String
Dim codsoc As Long
Dim Rs As ADODB.Recordset

Dim Rsx As ADODB.Recordset
Dim Sql2 As String
Dim codfor As Integer
Dim TipForpa As String
Dim PorIva As String

    On Error GoTo EContab

    Conn.BeginTrans
     
    'Insertar en la conta Cabecera Factura
    B = InsertarCabFactFac(cadwhere, cadMen, vContaFra)
    cadMen = "Insertando Cab. Factura: " & cadMen
    
    ' insertar en tesoreria
    If B Then
        Sql = "select * from fvarfactura where " & cadwhere
        Set Rsx = New ADODB.Recordset
        Rsx.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        B = InsertarEnTesoreriaNewFac(Rsx, FecVenci, CtaBanco, "")
        cadMen = "Insertando en Tesoreria: " & cadMen
        Set Rsx = Nothing
    End If

    If B Then
        'Insertar lineas de Factura en la Conta
        B = InsertarLinFactFacContaNueva("fvarfactura_lineas", Replace(cadwhere, "fvarfactura", "fvarfactura_lineas"), cadMen, CodCCost)
        cadMen = "Insertando Lin. Factura: " & cadMen

        If B Then
            vContaFra.AnyadeElError vContaFra.IntegraLaFacturaCliente(vContaFra.NumeroFactura, vContaFra.Anofac, vContaFra.Serie)
        
        
            'Poner intconta=1 en ariagroutil.cabfact
            B = ActualizarCabFact("fvarfactura", cadwhere, cadMen)
            cadMen = "Actualizando Factura: " & cadMen
        End If
    End If
    
    If Not B Then
        Sql = "Insert into tmperrfac(codtipom,numfactu,fecfactu,error) "
        Sql = Sql & " Select *," & DBSet(cadMen, "T") & " as error From tmpfactu "
        Sql = Sql & " WHERE " & Replace(cadwhere, "cabfact", "tmpfactu")
        Conn.Execute Sql
    End If
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Contabilizando Factura", Err.Description
    End If
    If B Then
        Conn.CommitTrans
        PasarFacturaFac = True
    Else
        Conn.RollbackTrans
        PasarFacturaFac = False
    End If
End Function

Private Function InsertarCabFactFac(cadwhere As String, caderr As String, ByRef vContaFra As cContabilizarFacturas) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim SqlDatos As String
Dim RsDatos As ADODB.Recordset
Dim Sql2 As String
Dim CadenaInsertFaclin2 As String

    On Error GoTo eInsertar
    
    Sql = " SELECT numserie,numfactu,fecfactu,codmacta, year(fecfactu) as anofaccl,"
    Sql = Sql & "baseiva1, baseiva2, baseiva3, impoiva1, impoiva2, impoiva3, imporec1,"
    Sql = Sql & "imporec2, imporec3, totalfac, tipoiva1, tipoiva2, tipoiva3, porciva1,"
    Sql = Sql & "porciva2, porciva3, porcrec1, porcrec2, porcrec3, totalfac, retfaccl, "
    Sql = Sql & "trefaccl, cuereten, codforpa "
    '[Monica]27/11/2017: hemos insertado todos los datos fiscales en la tabla
    Sql = Sql & ", nommacta, dirdatos, codposta, despobla, desprovi, nifdatos, codpais "
    '[Monica]31/07/2019: seleccionamos tiporeten, la referencia catastral y la situacion catastral
    Sql = Sql & ", tiporeten, CatastralREF, CatastralSITU "
    Sql = Sql & " FROM fvarfactura "
    Sql = Sql & " WHERE " & cadwhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    If Not Rs.EOF Then
        'guardamos estos valores para utilizarlos cuando insertemos las lineas de la factura
        vContaFra.NumeroFactura = DBLet(Rs!numfactu)
        vContaFra.Serie = DBLet(Rs!NUmSerie)
        vContaFra.Anofac = DBLet(Rs!anofaccl)
        
        Sql = ""
        Sql = DBSet(Trim(Rs!NUmSerie), "T") & "," & DBSet(Rs!numfactu, "N") & "," & DBSet(Rs!FecFactu, "F") & "," & DBSet(Rs!codmacta, "T") & "," & Year(Rs!FecFactu) & ",'FACTURACION',"
        
        BaseImp = Rs!baseiva1 + CCur(DBLet(Rs!baseiva2, "N")) + CCur(DBLet(Rs!baseiva3, "N"))
        IvaImp = DBLet(Rs!impoiva1, "N") + DBLet(Rs!impoiva2, "N") + DBLet(Rs!impoiva3, "N")
        
        '[Monica]31/07/2019: si el tipo de retencion es de arrendamiento el codconce340 = R
        If DBLet(Rs!tiporeten, "N") = 3 Then
            Sql = Sql & "'R',"
        Else
            ' esto es lo que estaba
            Sql = Sql & "'0',"
        End If
        
        Sql = Sql & "0," & DBSet(Rs!Codforpa, "N") & "," & DBSet(BaseImp, "N") & "," & ValorNulo & "," & DBSet(IvaImp, "N") & ","
        Sql = Sql & ValorNulo & "," & DBSet(Rs!TotalFac, "N") & "," & DBSet(Rs!retfaccl, "N") & "," & DBSet(Rs!trefaccl, "N") & "," & DBSet(Rs!cuereten, "T")
        
        Sql = Sql & "," & DBSet(Rs!tiporeten, "N")
        Sql = Sql & "," & DBSet(Rs!FecFactu, "F") & ","
        
        Sql = Sql & DBSet(Rs!Nommacta, "T", "S") & "," & DBSet(Rs!dirdatos, "T", "S") & "," & DBSet(Rs!codposta, "T", "S") & ","
        Sql = Sql & DBSet(Rs!desPobla, "T", "S") & "," & DBSet(Rs!desProvi, "T", "S") & "," & DBSet(Rs!nifdatos, "T", "S") & "," & DBSet(Rs!codpais, "T") & ",1"
        
        '[Monica]31/07/2019: metemos los campos correspondientes a arrendamiento
        Sql = Sql & "," & DBSet(Rs!CatastralREF, "T", "S") & "," & DBSet(Rs!CatastralSitu, "N", "S")
        
        Sql = "(" & Sql & ")"
        
        Sql2 = "INSERT INTO factcli (numserie,numfactu,fecfactu,codmacta,anofactu,observa,codconce340,codopera,codforpa,totbases,totbasesret,totivas,"
        Sql2 = Sql2 & "totrecargo,totfaccl, retfaccl,trefaccl,cuereten,tiporeten,fecliqcl,nommacta,dirdatos,codpobla,despobla, desprovi,nifdatos,"
        Sql2 = Sql2 & "codpais,codagente,"
        '[Monica]31/07/2019: metemos la referencia catastral y la situacion catastral
        Sql2 = Sql2 & "CatastralREF, CatastralSITU) "
        Sql2 = Sql2 & " VALUES " & Sql
        
        Conn.Execute Sql2
        
        CadenaInsertFaclin2 = ""
        
        'numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)
        'IVA 1, siempre existe
        Sql2 = "'" & Rs!NUmSerie & "'," & Rs!numfactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
        Sql2 = Sql2 & "1," & DBSet(Rs!baseiva1, "N") & "," & Rs!tipoiva1 & "," & DBSet(Rs!porciva1, "N") & ","
        Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva1, "N") & "," & ValorNulo
        CadenaInsertFaclin2 = CadenaInsertFaclin2 & "(" & Sql2 & ")"
        
        'para las lineas
        vTipoIva(0) = Rs!tipoiva1
        vPorcIva(0) = Rs!porciva1
        vPorcRec(0) = DBLet(Rs!porcrec1, "N")
        vImpIva(0) = Rs!impoiva1
        vImpRec(0) = DBLet(Rs!imporec1, "N")
        vBaseIva(0) = Rs!baseiva1
        
        vTipoIva(1) = 0: vTipoIva(2) = 0
        
        If Not IsNull(Rs!PorcIva2) Then
            Sql2 = "'" & Rs!NUmSerie & "'," & Rs!numfactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            Sql2 = Sql2 & "2," & DBSet(Rs!baseiva2, "N") & "," & Rs!tipoiva2 & "," & DBSet(Rs!PorcIva2, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva2, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(1) = Rs!tipoiva2
            vPorcIva(1) = Rs!PorcIva2
            vPorcRec(1) = DBLet(Rs!porcrec2, "N")
            vImpIva(1) = Rs!impoiva2
            vImpRec(1) = DBLet(Rs!imporec2, "N")
            vBaseIva(1) = Rs!baseiva2
        End If
        If Not IsNull(Rs!PorcIva3) Then
            Sql2 = "'" & Rs!NUmSerie & "'," & Rs!numfactu & "," & DBSet(Rs!FecFactu, "F") & "," & Year(Rs!FecFactu) & ","
            Sql2 = Sql2 & "3," & DBSet(Rs!baseiva3, "N") & "," & Rs!tipoiva3 & "," & DBSet(Rs!PorcIva3, "N") & ","
            Sql2 = Sql2 & ValorNulo & "," & DBSet(Rs!impoiva3, "N") & "," & ValorNulo
            CadenaInsertFaclin2 = CadenaInsertFaclin2 & " , (" & Sql2 & ")"
            vTipoIva(2) = Rs!tipoiva3
            vPorcIva(2) = Rs!PorcIva3
            vPorcRec(2) = DBLet(Rs!porcrec3, "N")
            vImpIva(2) = Rs!impoiva3
            vImpRec(2) = DBLet(Rs!imporec3, "N")
            vBaseIva(2) = Rs!baseiva3
        End If

        Sql = "INSERT INTO factcli_totales(numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,"
        Sql = Sql & "porciva,porcrec,impoiva,imporec) VALUES " & CadenaInsertFaclin2
        Conn.Execute Sql
        
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
eInsertar:
    If Err.Number <> 0 Then
        InsertarCabFactFac = False
        caderr = Err.Description
    Else
        InsertarCabFactFac = True
    End If
End Function



Public Function InsertarEnTesoreriaNewFac(ByRef Rsx As ADODB.Recordset, FecVenci As String, CtaBan As String, MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
Dim B As Boolean
Dim Sql As String, text33csb As String, text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim cadvalues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String

    On Error GoTo EInsertarTesoreriaNewFac

    B = False
    InsertarEnTesoreriaNewFac = False
    CadValues = ""
    cadvalues2 = ""


'[Monica]28/11/2017: para el caso de contabilidad nueva, metemos los datos que hemos dado en la factura
    Sql4 = "select * "
    Sql4 = Sql4 & " from cuentas where codmacta = " & Rsx!codmacta

    Set Rs4 = New ADODB.Recordset

    Rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not Rs4.EOF Then

        text33csb = "'Factura:" & DBLet(Trim(Rsx!NUmSerie), "T") & "-" & DBLet(Rsx!numfactu, "T") & " " & Format(DBLet(Rsx!FecFactu, "F"), "dd/mm/yy") & "'"
        text41csb = "de " & DBSet(Rsx!TotalFac, "N")
              
        CadValuesAux2 = "(" & DBSet(Trim(Rsx!NUmSerie), "T") & "," & DBSet(Rsx!numfactu, "N") & "," & DBSet(Rsx!FecFactu, "F") & ", 1," & DBSet(Rsx!codmacta, "T") & ","
        cadvalues2 = CadValuesAux2 & DBSet(Rsx!Codforpa, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(Rsx!TotalFac, "N") & "," & ValorNulo & ","
        cadvalues2 = cadvalues2 & DBSet(CtaBan, "T") & "," & ValorNulo & "," & ValorNulo & ","
        cadvalues2 = cadvalues2 & text33csb & "," & DBSet(text41csb, "T") & ",1," & DBSet(Rs4!IBAN, "T", "S") & ","
        '[Monica]28/11/2017: antes era del rs4, ahora lo tenemos en la factura
        cadvalues2 = cadvalues2 & DBSet(Rsx!Nommacta, "T", "S") & "," & DBSet(Rsx!dirdatos, "T", "S") & "," & DBSet(Rsx!desPobla, "T", "S") & ","
        cadvalues2 = cadvalues2 & DBSet(Rsx!codposta, "T", "S") & "," & DBSet(Rsx!desProvi, "T", "S") & "," & DBSet(Rsx!nifdatos, "T", "S") & "," & DBSet(Rsx!codpais, "T") & ")"
       
        'Insertamos en la tabla cobros de la CONTA
        Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, gastos, "
        Sql = Sql & "ctabanc1, fecultco, impcobro, "
        Sql = Sql & " text33csb, text41csb, agente, iban, "
        Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais"
        Sql = Sql & ") "
        
        Sql = Sql & " VALUES " & cadvalues2
        Conn.Execute Sql
    End If

    B = True

EInsertarTesoreriaNewFac:
    If Err.Number <> 0 Then B = False
    InsertarEnTesoreriaNewFac = B
End Function

Private Function InsertarLinFactFacContaNueva(cadTABLA As String, cadwhere As String, caderr As String, CCoste As String) As Boolean
'cadWHere: selecciona un registro de scafac
'codtipom=x and numfactu=y and fecfactu=z
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cad As String, Aux As String
Dim i As Long
Dim totimp As Currency, ImpLinea As Currency
Dim CodIVA As String
Dim IVA As String
Dim vIva As Currency

Dim NumeroIVA As Byte
Dim K As Integer
Dim HayQueAjustar As Boolean
Dim ImpImva As Currency
Dim ImpRec As Currency

    
    On Error GoTo EInLinea


    Sql = " SELECT numserie, numfactu, fecfactu, fvarconceptos.codmacta, fvarconceptos.codccost, " & cadTABLA & ".tipoiva , sum(importe) from " & cadTABLA
    Sql = Sql & ", fvarconceptos where " & cadwhere
    Sql = Sql & " and fvarconceptos.codconce = " & cadTABLA & ".codconce"
    Sql = Sql & " GROUP BY 1,2,3,4,5,6 "
    '[mONICA]04/02/2020: FALTA EL ORDEN
    Sql = Sql & " ORDER BY 6,4,5,1,2,3 "
    
    Set Rs = New ADODB.Recordset
    '[Monica]04/02/2020: cambio del tipo de recordset. CORREGIDO
    Rs.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenForwardOnly, adLockPessimistic, adCmdText

    cad = ""
    i = 1
    totimp = 0
    While Not Rs.EOF
        'concatenamos linea para insertar en la tabla de conta.linfact
        Sql = ""
        Sql = "'" & Trim(Rs!NUmSerie) & "'," & Rs!numfactu & "," & Year(Rs!FecFactu) & "," & i & ","
        
        'dependiendo del colectivo del socio cogemos la cta contable cliente o socio del articulo
        Sql = Sql & DBSet(Rs!codmacta, "T") & ","
        
        
        ImpLinea = DBLet(Rs.Fields(6).Value, "N")
        'Vemos que tipo de IVA es en el vector de importes
        NumeroIVA = 127
        For K = 0 To 2
            If Rs!TipoIva = vTipoIva(K) Then
                NumeroIVA = K
                Exit For
            End If
        Next
        If NumeroIVA > 100 Then Err.Raise 513, "Error obteniendo IVA: " & Rs!codigiva
        
        
        
        If DBLet(Rs!CodCCost, "T") = "" Then
            Sql = Sql & ValorNulo
        Else
            '[Monica]05/07/2012: comprobamos aqui que si hay analitica y tiene centro de coste, la cuenta debe comenzar por
            '                    los digitos que indican la contabilidad
            Dim GrupoGto As String
            Dim GrupoVta As String
            Dim GrupoOrd As String
            
            GrupoGto = vParam.GrupoGto 'DevuelveDesdeBDNewFac("parametros", "grupogto", "", "", "T")
            GrupoVta = vParam.GrupoVta 'DevuelveDesdeBDNewFac("parametros", "grupovta", "", "", "T")
            GrupoOrd = vParam.GrupoOrd 'DevuelveDesdeBDNewFac("parametros", "grupoord", "", "", "T")
             
            If vParam.autocoste And (Mid(Trim(Rs!codmacta), 1, 1) = GrupoGto Or Mid(Trim(Rs!codmacta), 1, 1) = GrupoVta Or Mid(Trim(Rs!codmacta), 1, 1) = GrupoOrd) Then
                Sql = Sql & DBSet(Rs!CodCCost, "T")
            Else
                Sql = Sql & ValorNulo
            End If
        End If
        
        vBaseIva(NumeroIVA) = vBaseIva(NumeroIVA) - ImpLinea   'Para ajustar el importe y que no haya descuadre
        HayQueAjustar = False
        If vBaseIva(NumeroIVA) <> 0 Then
            'falta importe.
            'Puede ser que hayan mas lineas, o haya descuadre. Como esta ordenado por tipo de iva
            Rs.MoveNext
            If Rs.EOF Then
                'No hay mas lineas
                'Hay que ajustar SI o SI
                HayQueAjustar = True
            Else
                'Si que hay mas lineas.
                'Son del mismo tipo de IVA
                If Rs!TipoIva <> vTipoIva(NumeroIVA) Then
                    'NO es el mismo tipo de IVA
                    'Hay que ajustar
                    HayQueAjustar = True
                End If
            End If
            Rs.MovePrevious
        End If

        Sql = Sql & "," & DBSet(Rs!FecFactu, "F") & "," & vTipoIva(NumeroIVA) & "," & DBSet(vPorcIva(NumeroIVA), "N") & "," & DBSet(vPorcRec(NumeroIVA), "N", "S") & ","

        If HayQueAjustar Then
            MsgBox "hay que ajustar"
        Else

        End If

        
        'Caluclo el importe de IVA y el de recargo de equivalencia
        ImpImva = vPorcIva(NumeroIVA) / 100
        ImpImva = Round2(ImpLinea * ImpImva, 2)
        If vPorcRec(NumeroIVA) = 0 Then
            ImpRec = 0
        Else
            ImpRec = vPorcRec(NumeroIVA) / 100
            ImpRec = Round2(ImpLinea * ImpRec, 2)
        End If
        vImpIva(NumeroIVA) = vImpIva(NumeroIVA) - ImpImva
        vImpRec(NumeroIVA) = vImpRec(NumeroIVA) - ImpRec
        
        
        ' baseimpo , impoiva, imporec
        Sql = Sql & DBSet(ImpLinea, "N") & "," & DBSet(ImpImva, "N") & "," & DBSet(ImpRec, "N", "S")
        
        cad = cad & "(" & Sql & ")" & ","
        
        i = i + 1
        Rs.MoveNext
        
    Wend
    
    Rs.Close
    Set Rs = Nothing

    'Insertar en la contabilidad
    If cad <> "" Then
        cad = Mid(cad, 1, Len(cad) - 1) 'quitar la ult. coma
        Sql = "INSERT INTO factcli_lineas (numserie,numfactu,anofactu,numlinea,codmacta,codccost,fecfactu,codigiva,porciva,porcrec,baseimpo,impoiva,imporec) "
        Sql = Sql & " VALUES " & cad
        Conn.Execute Sql
    End If

EInLinea:
    If Err.Number <> 0 Then
        InsertarLinFactFacContaNueva = False
        caderr = Err.Description
    Else
        InsertarLinFactFacContaNueva = True
    End If
End Function


Private Sub InsertarError(Cadena As String)
Dim Sql As String

    Sql = "insert into tmperrcomprob values ('" & Cadena & "')"
    Conn.Execute Sql

End Sub

Public Function EntreFechas(FIni As String, FechaComp As String, FFin As String) As Boolean
Dim B As Boolean
    B = False
    If FIni <> "" And FFin <> "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) And EsFechaIgualPosterior(FechaComp, FFin, False) Then
            B = True
        End If
    ElseIf FIni = "" And FFin <> "" Then
        If EsFechaIgualPosterior(FechaComp, FFin, False) Then
            B = True
        End If
    ElseIf FIni <> "" And FFin = "" Then
        If EsFechaIgualPosterior(FIni, FechaComp, False) Then
            B = True
        End If
    End If
    EntreFechas = B
End Function


Private Function ActualizarCabFact(cadTABLA As String, cadwhere As String, caderr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE " & cadTABLA & " SET intconta=1 "
    Sql = Sql & " WHERE " & cadwhere

    Conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCabFact = False
        caderr = Err.Description
    Else
        ActualizarCabFact = True
    End If
End Function

Public Function EsFechaIgualPosterior(FIni As String, FFin As String, MError As Boolean, Optional Men As String) As Boolean
'Comprueba que la Fecha Fin es igual o posterior a la Fecha de Inicio
'Si se pasa un cadena Men, se muestra esta como Mensaje de Error
'OUT -> true: Ffin >= Fini
On Error Resume Next

'    EsFechaIgualPosterior = True
    If Trim(FIni) <> "" And Trim(FFin) <> "" Then
        If CDate(FIni) > CDate(FFin) Then
            EsFechaIgualPosterior = False
            If MError Then
                If Men <> "" Then
                    MsgBox Men, vbInformation
                Else
                    MsgBox "La Fecha Fin debe ser igual o posterior a la Fecha Inicio", vbInformation
                End If
            End If
        Else
            EsFechaIgualPosterior = True
        End If
    Else
        EsFechaIgualPosterior = True
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


