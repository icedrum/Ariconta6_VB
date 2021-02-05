Attribute VB_Name = "libSII"
Option Explicit





'********************************************************
'  0 No tiene     1 Clientes     2 Proveedores   3 Ambos
''
'
Private Function TieneFacturasPendientesSubirSII() As Byte
Dim cad As String
Dim F As Date
Dim Aux As String
Dim RN As ADODB.Recordset
Dim C2 As String

Dim FIncio As Date

    TieneFacturasPendientesSubirSII = 0   'No tiene
    
    If vUsu.Nivel > 0 Then Exit Function
    If Not vParam.SIITiene Then Exit Function
    
 
    
    F = DateAdd("d", -1, Now)  'Han pasado los x Dias en parametros
    Set RN = New ADODB.Recordset
    
    FIncio = vParam.SIIFechaInicio
    If vParam.fechaini > FIncio Then FIncio = vParam.fechaini
        
    Aux = "0"
    
    'Dividimos el proceso en ir a FACTCLI, y a FACTPRO con NULL
    ' y luego ir a buscar erroers
    C2 = "select count(*) From factcli  WHERE "
    If vParam.SII_Periodo_DesdeLiq Then
        C2 = C2 & " fecliqcl >=" & DBSet(FIncio, "F")
        C2 = C2 & " AND fecliqcl <= " & DBSet(F, "F")
    Else
        C2 = C2 & " fecfactu >=" & DBSet(FIncio, "F")
        C2 = C2 & " AND fecfactu <= " & DBSet(F, "F")
    End If
    
    'Enero 2012. Sii_estado
    ' SII_estado 0.  Pendiente grabar aswi  1.  Insertada en aswii- pdte subir por daemon
    '            2.  Subida con errores     3.  Aceptada con errores
    '            8. Modificando factura presentada     9. Todo OK.
    
    'C2 = C2 & " AND sii_id is null"
    C2 = C2 & " AND sii_estado <9 "  'pendientes de subir o con error , o MODIFICANDOSE (8)
    
    RN.Open C2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RN.EOF Then
        If DBLet(RN.Fields(0), "N") > 0 Then Aux = "1"
    End If
    RN.Close
    If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = 1
    
    
    
    If TieneFacturasPendientesSubirSII = 0 Then
        C2 = "Select count(*) From factpro WHERE "
        If vParam.SII_Periodo_DesdeLiq Then
            C2 = C2 & " fecliqpr >=" & DBSet(FIncio, "F")
            C2 = C2 & " AND fecliqpr <= " & DBSet(F, "F")
        Else
            If vParam.SII_ProvDesdeFechaRecepcion Then
                C2 = C2 & " fecharec >=" & DBSet(FIncio, "F")
                C2 = C2 & " AND fecharec <= " & DBSet(F, "F")
            
            Else
                'Enero 2020
                C2 = C2 & " DATE(fecregcontable) >=" & DBSet(FIncio, "F")
                C2 = C2 & " AND DATE(fecregcontable) <= " & DBSet(F, "F")
            End If
        End If
        'Enero 2012. Sii_estado
        ' SII_estado 0.  Pendiente grabar aswi  1.  Insertada en aswii- pdte subir por daemon
        '            2.  Subida con errores     3.  Aceptada con errores
        '            8. Modificando factura presentada     9. Todo OK.
        
        'C2 = C2 & " AND sii_id is null"
        C2 = C2 & " AND sii_estado <9 "  'pendientes de subir o con error
        RN.Open C2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RN.EOF Then
            If DBLet(RN.Fields(0), "N") > 0 Then Aux = "1"
        End If
        RN.Close
        If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = 2
    
    End If
        
        
        
    '******************************    ****************************** ******************************
    ' YA NO linkamos con aswwi
    ' lo que habia aqui lo he copiado fuera de la funcion.
    'Si queremos verlo:   **Antiguo link**
    Set RN = Nothing
    
End Function

'******************************    ****************************** ******************************
'
'                   **Antiguo link**
'
'
'''''''
'''''''
'''''''    If TieneFacturasPendientesSubirSII = 0 Then
'''''''
'''''''        'incio fecha sii
'''''''        C2 = "select count(*) From factcli  left join aswsii.envio_facturas_emitidas"
'''''''        C2 = C2 & " on factcli.SII_ID = envio_facturas_emitidas.IDEnvioFacturasEmitidas where "
'''''''        If vParam.SII_Periodo_DesdeLiq Then
'''''''            C2 = C2 & " fecliqcl >=" & DBSet(FIncio, "F")
'''''''            C2 = C2 & " AND fecliqcl <= " & DBSet(F, "F")
'''''''        Else
'''''''            C2 = C2 & " fecfactu >=" & DBSet(FIncio, "F")
'''''''            C2 = C2 & " AND fecfactu <= " & DBSet(F, "F")
'''''''        End If
'''''''
'''''''        'Noviembre 2020
'''''''        C2 = C2 & " AND REG_FE_FA_IDFA_FechaExpedicionFacturaEmisor>=" & DBSet(vParam.fechaini, "F")
'''''''
'''''''
'''''''
'''''''        C2 = C2 & " and (csv is null or resultado='AceptadoConErrores')"
'''''''
'''''''
'''''''        RN.Open C2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''''''        If Not RN.EOF Then
'''''''            If DBLet(RN.Fields(0), "N") > 0 Then Aux = "1"
'''''''        End If
'''''''        RN.Close
'''''''        If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = 1
'''''''
'''''''
'''''''    End If
'''''''
'''''''
'''''''    If TieneFacturasPendientesSubirSII = 0 Then
'''''''
'''''''        C2 = "Select count(*) From factpro left join aswsii.envio_facturas_recibidas"
'''''''        C2 = C2 & " on factpro.SII_ID = envio_facturas_recibidas.IDEnvioFacturasRecibidas WHERE "
'''''''        If vParam.SII_Periodo_DesdeLiq Then
'''''''            C2 = C2 & " fecliqpr >=" & DBSet(FIncio, "F")
'''''''            C2 = C2 & " AND fecliqpr <= " & DBSet(F, "F")
'''''''        Else
'''''''            If vParam.SII_ProvDesdeFechaRecepcion Then
'''''''                C2 = C2 & " fecharec >=" & DBSet(FIncio, "F")
'''''''                C2 = C2 & " AND fecharec <= " & DBSet(F, "F")
'''''''
'''''''            Else
'''''''                'Enero 2020
'''''''                C2 = C2 & " DATE(fecregcontable) >=" & DBSet(FIncio, "F")
'''''''                C2 = C2 & " AND DATE(fecregcontable) <= " & DBSet(F, "F")
'''''''            End If
'''''''        End If
'''''''
'''''''        'Noviembre 2020
'''''''        C2 = C2 & " AND REG_FR_FechaRegContable>=" & DBSet(vParam.fechaini, "F")
'''''''
'''''''
'''''''
'''''''        C2 = C2 & " and (csv is null or resultado='AceptadoConErrores')"
'''''''        RN.Open C2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''''''        If Not RN.EOF Then
'''''''            If DBLet(RN.Fields(0), "N") > 0 Then Aux = "1"
'''''''        End If
'''''''        RN.Close
'''''''        If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = 2
'''''''    End If
'''''''





Private Sub LblIndica(ByRef LL As Label, TEXTO As String)
    If Not LL Is Nothing Then
        LL.Caption = TEXTO
        LL.Refresh
    End If
End Sub




'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'
'   Vamos a grabar en las tablas de la BD: aswsii
' Esta funcion retornara el SQL para dada una factura, insertarla en envio_facturas_emitidas
'
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
' si lleva numeroSII_ID_paraModificar : significa que estamos MODIFICANDO el registro.
' es mucho mas comodo poner REPLACE INTO
Public Function Sii_FraCLI(Serie As String, NumFac As Long, Anofac As Integer, IDEnvioFacturasEmitidas As Long, ByRef SQL_Insert As String, EsModificando As Boolean) As Boolean
Dim Sql As String
Dim RN As ADODB.Recordset
Dim Clave As String
Dim Aux As String
Dim rIVAS As ADODB.Recordset
Dim NumIVas As Integer
Dim CadenaIVAS As String
Dim LlevaIVAs As Boolean
Dim H As Integer
Dim C1 As String
Dim C2 As String
Dim c3 As String
Dim BloqueIVA As Byte
Dim FechaPeriodo2 As Date
Dim NumFactura As String
Dim FacturaResumenTicket As Boolean
Dim ImporteIvaCero As Currency
Dim LlevaIvasCero As Boolean
Dim B As Boolean
Dim GrabaTotalFactura_ As Boolean  'ene21

    On Error GoTo eSii_FraCLI
    Sii_FraCLI = False
    
    Sql = "Select factcli.*,Sii_SoloNUmeroFra from factcli left join contadores on factcli.numserie=contadores.tiporegi"
    Sql = Sql & " where factcli.numserie =" & DBSet(Serie, "T") & " AND factcli.numfactu =" & NumFac & " AND factcli.anofactu =" & Anofac
    Set RN = New ADODB.Recordset
    RN.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""

'#1
    'IDEnvioFacturasEmitidas,Origen,FechaHoraCreacion,EnvioInmediato,          'Enviada,Resultado: NO los pongo en el insert
    Sql = IDEnvioFacturasEmitidas & ",'ARICONTA'," & DBSet(Now, "FH") & ",1,"

'#2
    FechaPeriodo2 = RN!FecFactu
    If vParam.SII_Periodo_DesdeLiq Then FechaPeriodo2 = RN!fecliqcl
    
    'CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, REG_PI_Ejercicio, REG_PI_Periodo
    Sql = Sql & "'" & vParam.SII_Version & "'," & DBSet(vEmpresa.NombreEmpresaOficial, "T") & ",NULL," & DBSet(vEmpresa.NIF, "T") & ",'"
    
    
    If EsModificando Then
        Sql = Sql & "A1"
    Else
        'Lo que habia
        Sql = Sql & "A0"
    End If
    Sql = Sql & "','" & Year(FechaPeriodo2) & "','" & Format(Month(FechaPeriodo2), "00") & "',"
    
'#3
    'REG_IDF_IDEF_NIF,REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura
    Sql = Sql & DBSet(vEmpresa.NIF, "T") & ","

    FacturaResumenTicket = False
    If RN!codconce340 = "B" Then  'asiento resumen de factura (tickets agrupados indicando desde hasta
        FacturaResumenTicket = True
        'If DBLet(RN!FraResumenIni, "T") <> "" Then
        '
        '    FacturaResumenTicket = True
       '
       ' End If
    End If
    
    
    If FacturaResumenTicket Then
        'INCIO de las factiras de tickets agrupadas
        If DBLet(RN!FraResumenIni, "T") = "" Then
            NumFactura = RN!NUmSerie & Format(RN!numfactu, "0000000")
        Else
            NumFactura = RN!FraResumenIni
        End If
    
    Else
        'FACTURAS "NORMALES"
        NumFactura = RN!NUmSerie
        If DBLet(RN!Sii_SoloNUmeroFra, "N") = 1 Then NumFactura = ""
        NumFactura = NumFactura & Format(RN!numfactu, "0000000")
    End If
    
    'REG_IDF_NumSerieFacturaEmisor
    Sql = Sql & DBSet(NumFactura, "T") & ","
    
    
    'Si son de tickets agrupados deberiamos poner primera y ultima.
    'REG_IDF_NumSerieFacturaEmisorResumenFin
    If FacturaResumenTicket Then
        
        Aux = DBLet(RN!FraResumenFin, "T")
        
        'Si no hay nada, dejo lo que haciamos antes
        If Aux = "" Then Aux = "FTI" & Format(RN!numfactu, "0000000")
        Sql = Sql & DBSet(Aux, "T")
    Else
        Sql = Sql & "null"
    End If
    'REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura
    Sql = Sql & "," & DBSet(FechaPeriodo2, "F") & ","
    
    
    
    '#3.1
    ',REG_FE_TipoRectificativa,REG_FE_IR_BaseRectificada,REG_FE_IR_CuotaRectificada,REG_FE_IR_CuotaRecargoRectificado,
    Clave = DevuelveTipoFacturaEmitida(RN)   'Ver hoja. Hay tipos:    f1 factura   f2 tiket    r1 rectificativas
    Aux = ""
    Sql = Sql & DBSet(Clave, "T") & ","
  
    If Clave = "R1" Then
        Aux = "I"  'factura rectificativa por DIFERENCIAS
        Sql = Sql & DBSet(Aux, "T", "S") & ","
        'Opcionales. Numafac retificada
        Sql = Sql & "null,null,null,"

    Else
        'Los cuatro campos de la rectificativa a NULL
        Sql = Sql & "null,null,null,null,"

    End If
    
'#4

    GrabaTotalFactura_ = True
    If DBLet(RN!trefaccl, "N") <> 0 Then GrabaTotalFactura_ = False
    If DBLet(RN!Suplidos, "N") <> 0 Then GrabaTotalFactura_ = False

    'REG_FE_ClaveRegimenEspecialOTrascendencia,REG_FE_ImporteTotal,REG_FE_BaseImponibleACoste,REG_FE_DescripcionOperacion
    Clave = DevuelveClaveTranscendenciaEmitida(RN)
    Sql = Sql & DBSet(Clave, "T") & ","
    Sql = Sql & IIf(GrabaTotalFactura_, DBSet(RN!totfaccl, "N"), "NULL")
    Sql = Sql & ",NULL,"
    
    If FacturaResumenTicket Then
        Aux = "Factura " & RN!NUmSerie & RN!numfactu
    Else
    
        If vParam.TipoIntegracionSeleccionable = 1 Then
            Aux = "VENTAS"
        Else
            Aux = "Factura " & RN!NUmSerie & RN!numfactu
        End If
    End If
    Sql = Sql & DBSet(Aux, "T") & ","

'#4.1
    'REG_FE_DI_DT_ReferenciaCatastral,REG_FE_DI_DT_ReferenciaCatastral
    If RN!codconce340 = "R" Then
        'ARRENDAMIENTO
        Aux = DBLet(RN!CatastralREF, "T")
        If Aux = "" Then
            Sql = Sql & "NULL,NULL,"
        Else
            Aux = DBLet(RN!CatastralSitu, "N")
            If Val(Aux) = "0" Then
                Aux = "1"
            Else
                If Val(Aux) < 49 Or Val(Aux) > 52 Then
                    Aux = "1"
                Else
                    Aux = Val(Aux) - 48
                End If
            End If
            Sql = Sql & DBSet(RN!CatastralREF, "T") & "," & Aux & ","
        End If
    Else
        Sql = Sql & "NULL,NULL,"
    End If
    
    
    
    
'#5
    'REG_FE_EmitidaPorTercero,REG_FE_CNT_NombreRazon,REG_FE_CNT_NIF,REG_FE_CNT_IDOtro_CodigoPais,REG_FE_CNT_IDOtro_IDType,REG_FE_CNT_IDOtro_ID,
    If RN!codconce340 = "J" Or RN!codconce340 = "B" Then
        Aux = "null"
    Else
        Aux = DBSet(DBLet(RN!Nommacta, "T"), "T")
    End If

    
    Sql = Sql & "NULL," & Aux & ","
    
    'NIF. Para las intracoms el NIF debe llevar las letras
    BloqueIVA = 0 'NORMAL
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOMUNITARIAS    EXTRANJERO
        'PAIS doc
            
        C2 = DBSet(DBLet(RN!codpais, "T"), "T", "S")
        If RN!CodOpera = 1 Then
            Aux = DBLet(RN!nifdatos, "T")   'DBLet(RN!codPAIS, "T") & DBLet(RN!nifdatos, "T")
            C1 = "'02'"
        Else
            Aux = DBLet(RN!nifdatos, "T")
            C1 = "'03'"
        End If
        Sql = Sql & "''" & "," & C2 & "," & C1 & "," & DBSet(Aux, "T", "N") & ","
        BloqueIVA = 1 'Intracom y Exportacion
    Else
        'EL NIF
        'NO hacemos nada  AUX y c1 ya teiene los valores que toca
        C1 = "null"
        C2 = "null"
        c3 = "NULL"
        If RN!codconce340 = "J" Or RN!codconce340 = "B" Then
            'TICKETS NO `presentmaos NIFS
            Aux = "null"
        Else
        
            'Factura normal, pero a un "extranjero o lo que sea
            Aux = DBLet(RN!codpais, "T")
            If Aux = "" Then Aux = "ES"
        
            Aux = "ES"  'De momento Fuerzo un ES
        
            If Aux = "ES" Then
                'LO que estaba, no toco nada
                Aux = DBLet(RN!nifdatos, "T")
                Aux = DBSet(Aux, "T", "S")
                
                
            Else
                'Enero 2020
                'Si el pais, en codpais, NO es españa, pero es una factura normal normal
                'REG_FE_CNT_NIF,REG_FE_CNT_IDOtro_CodigoPais,REG_FE_CNT_IDOtro_IDType,REG_FE_CNT_IDOtro_ID
                C2 = DBSet(Aux, "T")
                C1 = "'02'"
                Aux = "''"
                c3 = DBSet(RN!nifdatos, "T")
                
            End If
        End If
        
        Sql = Sql & Aux & "," & C2 & "," & C1 & "," & c3 & ","
    End If
    
   
        
    
'6#
    'EXENTA
    
    'Modificacion SII 01/10/2019
    ' Los ivas a cero, aunque no sean exportacione-intracomunitaria, se reflejan en esta casilla. Con causa de exencion E1
    ImporteIvaCero = 0
    LlevaIvasCero = False
    
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOM y exportacion
        LlevaIVAs = False
        If RN!CodOpera = 1 Then
            Aux = "'E5'," 'intra
        Else
            Aux = "'E2',"  'export
        End If
        
        Aux = Aux & DBSet(RN!TotBases, "N") & ",null"
    Else
        LlevaIVAs = True
        'Aux = "NULL,NULL,'S1'"
        Aux = "#@CAUSA#,#@IMPOR#,#MotExen#"   '--> despues de ver los ivas, si alguno es cero replace esto, si no, replace por NULL
    End If
    Sql = Sql & Aux
    
    RN.Close
    
'7#
    'Bloque desglose IVAS hasta 6 ivas. Cambia el numerito ...DT1   DT2..
    CadenaIVAS = ""
    NumIVas = 0
    If LlevaIVAs Then
        
        Aux = "Select * from factcli_totales  left join tiposiva on  factcli_totales.codigiva=tiposiva.codigiva where numserie =" & DBSet(Serie, "T") & " AND numfactu =" & NumFac & " AND anofactu =" & Anofac
        RN.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
        'TipoImpositivo,BaseImponible,CuotaRepercutida,TipoREquivalencia,CuotaREquivalencia,"
        While Not RN.EOF
            If RN!TipoDIva <> 4 Then
            
                If RN!PorcIva = 0 Then
                    B = False
                    
                    LlevaIvasCero = True
                    ImporteIvaCero = ImporteIvaCero + RN!Baseimpo
                Else
                    B = True
                End If
                If B Then
                    Aux = "," & DBSet(RN!PorcIva, "N") & "," & DBSet(RN!Baseimpo, "N") & "," & DBSet(RN!Impoiva, "N") & ","
                    If IsNull(RN!porcrec) Then
                        Aux = Aux & "NULL,NULL"
                    Else
                        Aux = Aux & DBSet(RN!porcrec, "N") & "," & DBSet(RN!ImpoRec, "N")
                    End If
                    CadenaIVAS = CadenaIVAS & Aux
                    NumIVas = NumIVas + 1
                End If
            End If
            RN.MoveNext
        Wend
        RN.Close
        
        
        If LlevaIvasCero Then
            Sql = Replace(Sql, "#@CAUSA#", "'E1'")
            Sql = Replace(Sql, "#@IMPOR#", DBSet(ImporteIvaCero, "N"))
            If NumIVas > 0 Then
                'AParte del exteno lleva otro mas
                Sql = Replace(Sql, "#MotExen#", "'S1'")
            Else
                Sql = Replace(Sql, "#MotExen#", "NULL")
            End If
        
        Else
            'Aux = "NULL,NULL,'S1'"
            Sql = Replace(Sql, "#@CAUSA#", "NULL")
            Sql = Replace(Sql, "#@IMPOR#", "NULL")
            Sql = Replace(Sql, "#MotExen#", "'S1'")
        End If
    End If
    
    For H = NumIVas + 1 To 6
        If BloqueIVA = 0 Then
            CadenaIVAS = CadenaIVAS & ",NULL,NULL,NULL,NULL,NULL"
        Else
            'En los IVAS de intracom/exportacion NO llevamos REcargo de equivalencia. Ni % ni cuota
            CadenaIVAS = CadenaIVAS & ",NULL,NULL,NULL"
        End If
    Next
    Sql = Sql & CadenaIVAS
    
    
    'Montamos el SQL
    SQL_Insert = Sii_FraCLI_SQL(BloqueIVA, EsModificando) & ") VALUES (" & Sql & ")"
    
    Sii_FraCLI = True
    
eSii_FraCLI:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RN = Nothing
End Function


'  0.- Facturas normales                ->  REG_FE_TD_DF_SU
'  1.- Intracomunitarias // Extranjera  ->  REG_FE_TD_DTE_SU
' Si modifica hara un REPLACE INTO
Private Function Sii_FraCLI_SQL(BloquesIVA As Byte, EsModificando As Boolean) As String
Dim cad As String
Dim H As Integer
    
    If EsModificando Then
        Sii_FraCLI_SQL = "REPLACE INTO"
    Else
        Sii_FraCLI_SQL = "INSERT  INTO"
    End If
    
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & " aswsii.envio_facturas_emitidas("
    '#1
    'IDEnvioFacturasEmitidas,Origen,FechaHoraCreacion,EnvioInmediato,Enviada,Resultado,
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & "IDEnvioFacturasEmitidas,Origen,FechaHoraCreacion,EnvioInmediato,"       'Enviada,Resultado,"
    
    '#2
    'CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, REG_PI_Ejercicio, REG_PI_Periodo
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & "CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, CAB_TipoComunicacion, REG_PI_Ejercicio, REG_PI_Periodo,"
    
    '#3
    'REG_IDF_IDEF_NIF,REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & "REG_IDF_IDEF_NIF,REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura,"
    
    '#3.1
    'REG_FE_TipoRectificativa,REG_FE_IR_BaseRectificada,REG_FE_IR_CuotaRectificada,REG_FE_IR_CuotaRecargoRectificado,
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & "REG_FE_TipoRectificativa,REG_FE_IR_BaseRectificada,REG_FE_IR_CuotaRectificada,REG_FE_IR_CuotaRecargoRectificado,"
    
    '#4
    'REG_FE_ClaveRegimenEspecialOTrascendencia,REG_FE_ImporteTotal,REG_FE_BaseImponibleACoste,REG_FE_DescripcionOperacion
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & "REG_FE_ClaveRegimenEspecialOTrascendencia,REG_FE_ImporteTotal,REG_FE_BaseImponibleACoste,REG_FE_DescripcionOperacion,"
    
    '#4.1
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & "REG_FE_DI_DT_ReferenciaCatastral,REG_FE_DI_DT_SituacionInmueble,"
    
    
    
    '#5
    Sii_FraCLI_SQL = Sii_FraCLI_SQL & "REG_FE_EmitidaPorTercero,REG_FE_CNT_NombreRazon,REG_FE_CNT_NIF,REG_FE_CNT_IDOtro_CodigoPais,REG_FE_CNT_IDOtro_IDType,REG_FE_CNT_IDOtro_ID,"
    
    '
    '6#  BLOQUE  facturas normales IVA
    '
    If BloquesIVA = 0 Then
        Sii_FraCLI_SQL = Sii_FraCLI_SQL & "REG_FE_TD_DF_SU_EX_CausaExencion , REG_FE_TD_DF_SU_EX_BaseImponible, REG_FE_TD_DF_SU_NEX_TipoNoExenta"
        
                    'REG_FE_TD_DF_SU_NEX_DI_DT1_TipoImpositivo,REG_FE_TD_DF_SU_NEX_DI_DT1_BaseImponible,REG_FE_TD_DF_SU_NEX_DI_DT1_CuotaRepercutida,REG_FE_TD_DF_SU_NEX_DI_DT1_TipoREquivalencia,REG_FE_TD_DF_SU_NEX_DI_DT1_CuotaREquivalencia,
        For H = 1 To 6
            cad = ",REG_FE_TD_DF_SU_NEX_DI_DT" & H & "_TipoImpositivo,REG_FE_TD_DF_SU_NEX_DI_DT" & H & "_BaseImponible,REG_FE_TD_DF_SU_NEX_DI_DT" & H & "_CuotaRepercutida,REG_FE_TD_DF_SU_NEX_DI_DT" & H & "_TipoREquivalencia,REG_FE_TD_DF_SU_NEX_DI_DT" & H & "_CuotaREquivalencia"
            Sii_FraCLI_SQL = Sii_FraCLI_SQL & cad
        Next
        
    Else
        'Facturas intracomunitarias e exportaciones
        Sii_FraCLI_SQL = Sii_FraCLI_SQL & "REG_FE_TD_DTE_SU_EX_CausaExencion , REG_FE_TD_DTE_SU_EX_BaseImponible, REG_FE_TD_DTE_SU_NEX_TipoNoExenta"
        
                    'REG_FE_TD_DTE_SU 1_TipoImpositivo, _BaseImponible,R _CuotaRepercutida, TipoREquivalencia, _CuotaREquivalencia,
        For H = 1 To 6
            cad = ",REG_FE_TD_DTE_SU_NEX_DI_DT" & H & "_TipoImpositivo,REG_FE_TD_DTE_SU_NEX_DI_DT" & H & "_BaseImponible,REG_FE_TD_DTE_SU_NEX_DI_DT" & H & "_CuotaRepercutida"
            Sii_FraCLI_SQL = Sii_FraCLI_SQL & cad
        Next
    
    
    End If
    
    
End Function



Private Function DevuelveTipoFacturaEmitida(ByRef R As ADODB.Recordset) As String
    
    If R!codconce340 = "D" Then
        'Rectificativa
        DevuelveTipoFacturaEmitida = "R1"

    ElseIf R!codconce340 = "J" Then
        DevuelveTipoFacturaEmitida = "F2"
        
    ElseIf R!codconce340 = "B" Then
        DevuelveTipoFacturaEmitida = "F4"
    Else
        'NORMAL
        DevuelveTipoFacturaEmitida = "F1"
    End If
    
End Function
Private Function DevuelveClaveTranscendenciaEmitida(ByRef R As ADODB.Recordset) As String
    'Valores de codopera
    '0   "GENERAL"
    '1   "INTRACOMUNITARIA"
    '2   "EXPORT. - IMPORT."
    '3   "INTERIOR EXENTA"
    '4   "INV. SUJETO PASIVO"
    '5   "R.E.A."
    'Si es exportaciobn
    If R!CodOpera = "2" Then
        DevuelveClaveTranscendenciaEmitida = "02"
    Else
    
        'Si es operaciones de ARRENDAMIENTO
        'EMITIDAS 11: S/REF C/RET 13 C/REF C/S/RET
        If R!codconce340 = "R" Then
        
            If DBLet(R!cuereten, "T") <> "" Then
                DevuelveClaveTranscendenciaEmitida = "13"
            Else
                DevuelveClaveTranscendenciaEmitida = "11"
            End If
        Else
            DevuelveClaveTranscendenciaEmitida = "01"
        End If
    End If

    
End Function


Private Function DevuelveClaveTranscendenciaRecibida(ByRef R As ADODB.Recordset) As String
    'Valores de codopera
    '0   "GENERAL"
    '1   "INTRACOMUNITARIA"
    '2   "EXPORT. - IMPORT."
    '3   "INTERIOR EXENTA"
    '4   "INV. SUJETO PASIVO"
    '5   "R.E.A."
    'Si es exportaciobn
    If R!CodOpera = "5" Then
        DevuelveClaveTranscendenciaRecibida = "02"
    ElseIf R!CodOpera = "1" Then
        DevuelveClaveTranscendenciaRecibida = "09"
    Else
    
    
        DevuelveClaveTranscendenciaRecibida = "01"
        
        
        'Normal
        'Si es operaciones de ARRENDAMIENTO
        'EMITIDAS 11: S/REF C/RET 13 C/REF C/S/RET
'        If R!codconce340 = "R" Then
'            If DBLet(R!CatastralREF, "T") = "" Then
'                DevuelveClaveTranscendenciaEmitida = "11"
'            Else
'                DevuelveClaveTranscendenciaEmitida = "13"
'            End If
'        Else
        
        
    End If

    
End Function

































'****************************************************************************
'****************************************************************************
'
' RECIBIDAS
'
'****************************************************************************
'****************************************************************************
Public Function Sii_FraPRO(Serie As String, Numregis As Long, Anofac As Integer, IDEnvioFacturasRecibidas As Long, ByRef SQL_Insert As String, EsModificando As Boolean) As Boolean
Dim Sql As String
Dim RN As ADODB.Recordset
Dim Clave As String
Dim Aux As String
Dim rIVAS As ADODB.Recordset
Dim NumIVas As Integer
Dim CadenaIVAS As String
Dim H As Integer
Dim C1 As String
Dim C2 As String
Dim TotalDecucible As Currency
Dim CodOpera As Byte
Dim InversionSujetoPasivo As Boolean
Dim FechaPeriodo2 As Date
Dim NoDeducible As Boolean  '2019 Septiembre

Dim GrabaTotalFactura As Boolean
    'Total factura es cun campo "opcional".
    'Modificacion de Ene-2021 SII hace una comprobacion de Totalbases + IVAS.
    ' Porblema: las facturas de retencion que no son ARRENDAMIENTOS, (rea con retencion) esa suma no es correcta, con lo cual da error

    On Error GoTo eSii_FraCLI
    Sii_FraPRO = False
    
    Sql = "Select * from factpro where numserie =" & DBSet(Serie, "T") & " AND numregis =" & Numregis & " AND anofactu =" & Anofac
    Set RN = New ADODB.Recordset
    RN.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""

'#1
    'IDEnvioFacturasEmitidas,Origen,FechaHoraCreacion,EnvioInmediato,          'Enviada,Resultado: NO los pongo en el insert
    Sql = IDEnvioFacturasRecibidas & ",'ARICONTA'," & DBSet(Now, "FH") & ",1,"

'#2
    'CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, REG_PI_Ejercicio, REG_PI_Periodo
    FechaPeriodo2 = RN!fecharec
    If vParam.SII_Periodo_DesdeLiq Then FechaPeriodo2 = RN!fecliqpr
    
    Sql = Sql & "'" & vParam.SII_Version & "'," & DBSet(vEmpresa.NombreEmpresaOficial, "T") & ",NULL," & DBSet(vEmpresa.NIF, "T") & ",'"
     If EsModificando Then
        Sql = Sql & "A1"
    Else
        'Lo que habia
        Sql = Sql & "A0"
    End If
    Sql = Sql & "'," & Year(FechaPeriodo2) & "," & "'" & Format(Month(FechaPeriodo2), "00") & "',"
'#3
    'REG_IDF_IDEF_NIF,REG_IDF_IDEF_IDOtro_CodigoPais,REG_IDF_IDEF_IDOtro_IDType,REG_IDF_IDEF_IDOtro_ID
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOMUNITARIAS    EXTRANJERO
        'PAIS doc
        C2 = DBSet(DBLet(RN!codpais, "T"), "T", "S")
        If RN!CodOpera = 1 Then
            Aux = DBLet(RN!nifdatos, "T")  'DBLet(RN!codPAIS, "T") & DBLet(RN!nifdatos, "T")
            C1 = "'02'"
        Else
            Aux = DBLet(RN!nifdatos, "T")
            C1 = "'03'"
        End If
        Sql = Sql & "''" & "," & C2 & "," & C1 & "," & DBSet(Aux, "T", "N") & ","
    Else
    
        'Abril 2020   DUA
        If RN!CodOpera = 6 Then
            
            C1 = "null"
            Aux = DBLet(vEmpresa.NIF, "T")   'DUAS presentamos como NIF el de la empresa
            C2 = "null"
        Else
            'EL NIF
            'NO hacemos nada  AUX y c1 ya teiene los valores que toca
            C1 = "null"
            Aux = DBLet(RN!nifdatos, "T")
            C2 = "null"
        End If
        Sql = Sql & DBSet(Aux, "T", "N") & "," & C2 & "," & C1 & ",NULL,"
    End If
    
    
'#4
    'REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura,REG_FE_TipoRectificativa
    'Si son de tickets agrupados deberiamos poner primera y ultima. De momento null
    Sql = Sql & DBSet(RN!numfactu, "T") & "," & "NULL," & DBSet(RN!FecFactu, "F") & ","
    Clave = DevuelveTipoFacturaRecibida(RN)
    Aux = ""
    If Clave = "R1" Then Aux = "I"  'factura rectificativa por diferencias
    Sql = Sql & DBSet(Clave, "T") & "," & DBSet(Aux, "T", "S") & ","
    
    
    
    
'#4.1   No implmentado en ASWSII
'    If RN!codconce340 = "R" Then
'        'ARRENDAMIENTO
'        Aux = DBLet(RN!CatastralREF, "T")
'        If Aux = "" Then
'            SQL = SQL & "NULL,NULL,"
'        Else
'            Aux = DBLet(RN!CatastralSitu, "N")
'            If Val(Aux) = "0" Then
'                Aux = "1"
 '           Else
 '               If Val(Aux) < 49 Or Val(Aux) > 52 Then
 '                   Aux = "1"
 '               Else
 '                   Aux = Val(Aux) - 48
 '               End If
 '           End If
 '           SQL = SQL & DBSet(RN!CatastralREF, "T") & "," & Aux & ","
 '       End If
 '   Else
 '       SQL = SQL & "NULL,NULL,"
 '   End If
    
    
    GrabaTotalFactura = True
    If DBLet(RN!trefacpr, "N") <> 0 Then GrabaTotalFactura = False
    If DBLet(RN!Suplidos, "N") <> 0 Then GrabaTotalFactura = False
    
'#5
     
    'REG_FE_ClaveRegimenEspecialOTrascendencia,REG_FE_ImporteTotal,REG_FE_BaseImponibleACoste,REG_FE_DescripcionOperacion
    Clave = DevuelveClaveTranscendenciaRecibida(RN)
    Sql = Sql & DBSet(Clave, "T") & ","
    Sql = Sql & IIf(GrabaTotalFactura, DBSet(RN!totfacpr, "N"), "NULL")
    Sql = Sql & ",NULL," 'REG_FE_DescripcionOperacion
    
    If vParam.TipoIntegracionSeleccionable = 1 Then
        Aux = "numserie =" & DBSet(RN!NUmSerie, "T") & " AND numregis =" & RN!Numregis & " AND anofactu "
        Aux = DevuelveDesdeBD("codmacta", "factpro_lineas", Aux, CStr(RN!Anofactu) & " ORDER by numlinea", "N")
        If Aux = "" Then
            Aux = "6000"
        Else
            Aux = Mid(Aux, 1, 4)
        End If
        If Val(Aux) > 60000 Then
            Aux = "GASTOS"
        Else
            Aux = "COMPRAS"
        End If
        
        Sql = Sql & "'" & Aux & "',"
        
    Else
        'SQL = SQL & "'Factura" & IIf(RN!NUmSerie = 1, "", " ser: " & RN!NUmSerie) & " " & RN!NumFactu & "',"
        Sql = Sql & "'Factura" & RN!numfactu & "',"
    End If
    
    
    
'#6
    'REG_FE_EmitidaPorTercero,REG_FE_CNT_NombreRazon,REG_FE_CNT_NIF,REG_FE_CNT_IDOtro_CodigoPais,REG_FE_CNT_IDOtro_IDType,REG_FE_CNT_IDOtro_ID,
    Aux = DBLet(RN!Nommacta, "T")
    'If RN!CodOpera = 6 Then Aux = vEmpresa.NombreEmpresaOficial
    Sql = Sql & DBSet(Aux, "T") & ","
    
    'NIF. Para las intracoms el NIF debe llevar las letras
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOMUNITARIAS    EXTRANJERO
        'PAIS doc
        C2 = DBSet(DBLet(RN!codpais, "T"), "T", "S")
        If RN!CodOpera = 1 Then
            Aux = DBLet(RN!nifdatos, "T")   ' DBLet(RN!codPAIS, "T") & DBLet(RN!nifdatos, "T")
            C1 = "'02'"
        Else
            Aux = DBLet(RN!nifdatos, "T")
            C1 = "'03'"
        End If
        Sql = Sql & "''" & "," & C2 & "," & C1 & "," & DBSet(Aux, "T", "N") & ","
    Else
        'EL NIF
        'NO hacemos nada  AUX y c1 ya teiene los valores que toca
        
        If RN!CodOpera = 6 Then
            'DUAs
            C1 = "null"
            Aux = DBLet(vEmpresa.NIF, "T")
            C2 = "null"
        Else
            C1 = "null"
            Aux = DBLet(RN!nifdatos, "T")
            C2 = "null"
        End If
        Sql = Sql & DBSet(Aux, "T", "N") & "," & C2 & "," & C1 & ",NULL,"
    End If
    
    
    '#7  REG_FR_FechaOperacion  REG_FR_FechaRegContable  REG_FR_CuotaDeducible
    CodOpera = RN!CodOpera
    InversionSujetoPasivo = False
    If CodOpera = 4 Then InversionSujetoPasivo = True
        
        
    'Enero 2020
        'Se añade campo fecregcontable . Los de SII desde liquidacion lo dejamos como esta
    If vParam.SII_ProvDesdeFechaRecepcion Then
        FechaPeriodo2 = RN!fecharec
    Else
        FechaPeriodo2 = RN!fecregcontable
    End If
    If vParam.SII_Periodo_DesdeLiq Then FechaPeriodo2 = RN!fecliqpr
    
    Sql = Sql & DBSet(RN!FecFactu, "F") & "," & DBSet(FechaPeriodo2, "F") & ",#@#@#@$$$$"   'Sumaremos el total de cuotas deducibles y luego haremos un replace

    
    
    
    RN.Close
    
    TotalDecucible = 0
    
    
    
'#8 Inversion sujeto apsivo   ***ISP^^^^
    'hasta 6 ivas. Cambia el numerito ...DT1   DT2..   REG_FR_DF_ISP_DI_DT6_CuotaREquivalencia
    
    CadenaIVAS = ""
    NumIVas = 0
    If InversionSujetoPasivo Then
        
        Aux = "Select * from factpro_totales where numserie =" & DBSet(Serie, "T") & " AND numregis =" & Numregis & " AND anofactu =" & Anofac
        RN.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        'TipoImpositivo,BaseImponible,CuotaRepercutida,TipoREquivalencia,CuotaREquivalencia,"
        While Not RN.EOF
            
            Aux = "," & DBSet(RN!PorcIva, "N") & "," & DBSet(RN!Baseimpo, "N") & "," & DBSet(RN!Impoiva, "N") & ","
            If IsNull(RN!porcrec) Then
                Aux = Aux & "NULL,NULL"
            Else
                Aux = Aux & DBSet(RN!porcrec, "N") & "," & DBSet(RN!ImpoRec, "N")
            End If
            CadenaIVAS = CadenaIVAS & Aux
            NumIVas = NumIVas + 1
            
            TotalDecucible = TotalDecucible + RN!Impoiva + DBLet(RN!ImpoRec, "N")
            
            RN.MoveNext
        Wend
        RN.Close
    End If
    
    For H = NumIVas + 1 To 6
        CadenaIVAS = CadenaIVAS & ",NULL,NULL,NULL,NULL,NULL"
    Next
    Sql = Sql & CadenaIVAS
    
    

    
'#9
    'hasta 6 ivas. Cambia el numerito ...DT1   DT2..  REG_FR_DF_DGI_DI_DT1_TipoImpositivo
    
    CadenaIVAS = ""
    NumIVas = 0
    If Not InversionSujetoPasivo Then
        
        Aux = "Select * from factpro_totales  left join tiposiva on  factpro_totales.codigiva=tiposiva.codigiva  where numserie =" & DBSet(Serie, "T") & " AND numregis =" & Numregis & " AND anofactu =" & Anofac
        RN.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
        'TipoImpositivo,BaseImponible,CuotaRepercutida,TipoREquivalencia,CuotaREquivalencia,"

        
        While Not RN.EOF
            Aux = ""
            NoDeducible = False
            If CodOpera = 5 Then
                'Si el tipo de IVA es REA
                Aux = ",null," & DBSet(RN!Baseimpo, "N") & ",null,null,null,"
                '% REA impor REA
                Aux = Aux & DBSet(RN!PorcIva, "N") & "," & DBSet(RN!Impoiva, "N")
            
                
            Else
                If RN!TipoDIva <> 4 Then
                    
                    
                
                    Aux = "," & DBSet(RN!PorcIva, "N") & "," & DBSet(RN!Baseimpo, "N") & "," & DBSet(RN!Impoiva, "N") & ","
                    If IsNull(RN!porcrec) Then
                        Aux = Aux & "NULL,NULL"
                    Else
                        Aux = Aux & DBSet(RN!porcrec, "N") & "," & DBSet(RN!ImpoRec, "N")
                    End If
                    Aux = Aux & ",NULL,NULL"             'REA A null
                    
                    
                    If RN!TipoDIva = 3 Then NoDeducible = True
                    
                    
                End If
            End If
            If Aux <> "" Then
                CadenaIVAS = CadenaIVAS & Aux
                NumIVas = NumIVas + 1
                If NoDeducible Then
                    'Este importe NO suma al deducible
                Else
                    'Normal
                    TotalDecucible = TotalDecucible + RN!Impoiva + DBLet(RN!ImpoRec, "N")
                End If
            End If
            RN.MoveNext
        
        Wend
        RN.Close
    End If
    
    For H = NumIVas + 1 To 6
        CadenaIVAS = CadenaIVAS & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL"
    Next
    Sql = Sql & CadenaIVAS
    
    'Total deducciones
    Sql = Replace(Sql, "#@#@#@$$$$", DBSet(TotalDecucible, "N"))
    
    
    
    'Montamos el SQL
    SQL_Insert = Sii_FraPRO_SQL(EsModificando) & ") VALUES (" & Sql & ")"
    
    Sii_FraPRO = True
    
eSii_FraCLI:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RN = Nothing
End Function


Private Function Sii_FraPRO_SQL(EsModificando As Boolean) As String
Dim cad As String
Dim H As Integer

    If EsModificando Then
        Sii_FraPRO_SQL = "REPLACE INTO"
    Else
        Sii_FraPRO_SQL = "INSERT  INTO"
    End If


    Sii_FraPRO_SQL = Sii_FraPRO_SQL & " aswsii.envio_facturas_recibidas("
    '#1
    'IDEnvioFacturasEmitidas,Origen,FechaHoraCreacion,EnvioInmediato,Enviada,Resultado,
    Sii_FraPRO_SQL = Sii_FraPRO_SQL & "IDEnvioFacturasRecibidas,Origen,FechaHoraCreacion,EnvioInmediato,"       'Enviada,Resultado,"
    
    '#2
    'CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, REG_PI_Ejercicio, REG_PI_Periodo
    Sii_FraPRO_SQL = Sii_FraPRO_SQL & "CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, CAB_TipoComunicacion, REG_PI_Ejercicio, REG_PI_Periodo,"
    
    '#3
    'REG_IDF_IDEF_NIF,  REG_IDF_IDEF_IDOtro_CodigoPais REG_IDF_IDEF_IDOtro_IDType REG_IDF_IDEF_IDOtro_ID
    Sii_FraPRO_SQL = Sii_FraPRO_SQL & "REG_IDF_IDEF_NIF , REG_IDF_IDEF_IDOtro_CodigoPais , REG_IDF_IDEF_IDOtro_IDType , REG_IDF_IDEF_IDOtro_ID,"
    
    
    
    '#4
    'REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura,REG_FE_TipoRectificativa
    Sii_FraPRO_SQL = Sii_FraPRO_SQL & "REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FR_TipoFactura,REG_FR_TipoRectificativa,"
    
    '#4.1 NO Implementado en ASWSII
    'Sii_FraPRO_SQL = Sii_FraPRO_SQL & "REG_FE_DI_DT_ReferenciaCatastral,REG_FE_DI_DT_SituacionInmueble,"
    
    
    '#5
    'REG_FE_ClaveRegimenEspecialOTrascendencia,REG_FE_ImporteTotal,REG_FE_BaseImponibleACoste,REG_FE_DescripcionOperacion
    Sii_FraPRO_SQL = Sii_FraPRO_SQL & "REG_FR_ClaveRegimenEspecialOTrascendencia,REG_FR_ImporteTotal,REG_FR_BaseImponibleACoste,REG_FR_DescripcionOperacion,"
    
    '#6 Coincide con el #3
    Sii_FraPRO_SQL = Sii_FraPRO_SQL & "REG_FR_CNT_NombreRazon,REG_FR_CNT_NIF,REG_FR_CNT_IDOtro_CodigoPais,REG_FR_CNT_IDOtro_IDType,REG_FR_CNT_IDOtro_ID,"
    
    '#7
    Sii_FraPRO_SQL = Sii_FraPRO_SQL & " REG_FR_FechaOperacion , REG_FR_FechaRegContable , REG_FR_CuotaDeducible"
    
    
    '#8
    
    'ISP
    For H = 1 To 6
        'REG_FR_DF_ISP_DI_DT1_TipoImpositivo
        'REG_FR_DF_ISP_DI_DT1_TipoImpositivo  REG_FR_DF_ISP_DI_DT1_BaseImponible
        'REG_FR_DF_ISP_DI_DT1_CuotaSoportada  REG_FR_DF_ISP_DI_DT1_TipoREquivalencia REG_FR_DF_ISP_DI_DT1_CuotaREquivalencia
        cad = ",REG_FR_DF_ISP_DI_DT" & H & "_TipoImpositivo,REG_FR_DF_ISP_DI_DT" & H & "_BaseImponible,REG_FR_DF_ISP_DI_DT" & H & "_CuotaSoportada,REG_FR_DF_ISP_DI_DT" & H & "_TipoREquivalencia,"
        cad = cad & "REG_FR_DF_ISP_DI_DT" & H & "_CuotaREquivalencia"
        Sii_FraPRO_SQL = Sii_FraPRO_SQL & cad
    Next
    
    
    'Resto
    'REG_FR_DF_DGI_DI_DT1_TipoImpositivo REG_FR_DF_DGI_DI_DT1_BaseImponible REG_FR_DF_DGI_DI_DT1_CuotaSoportada
    'REG_FR_DF_DGI_DI_DT1_TipoREquivalencia REG_FR_DF_DGI_DI_DT1_CuotaREquivalencia REG_FR_DF_DGI_DI_DT1_PorcentCompensacionREAGYP REG_FR_DF_DGI_DI_DT1_ImporteCompensacionREAGYP
    For H = 1 To 6
        'REG_FR_DF_DGI_DI_DT1
        '_TipoImpositivo _BaseImponible _DT1_CuotaSoportada _TipoREquivalencia
        '_CuotaREquivalencia _PorcentCompensacionREAGYP _ImporteCompensacionREAGYP
        cad = ",REG_FR_DF_DGI_DI_DT" & H & "_TipoImpositivo,REG_FR_DF_DGI_DI_DT" & H & "_BaseImponible"
        cad = cad & ",REG_FR_DF_DGI_DI_DT" & H & "_CuotaSoportada,REG_FR_DF_DGI_DI_DT" & H & "_TipoREquivalencia"
        cad = cad & ",REG_FR_DF_DGI_DI_DT" & H & "_CuotaREquivalencia,REG_FR_DF_DGI_DI_DT" & H & "_PorcentCompensacionREAGYP"
        cad = cad & ",REG_FR_DF_DGI_DI_DT" & H & "_ImporteCompensacionREAGYP "
        Sii_FraPRO_SQL = Sii_FraPRO_SQL & cad
    Next

    
End Function


Private Function DevuelveTipoFacturaRecibida(ByRef R As ADODB.Recordset) As String
    
    
    'Nuevo.
    'Abril 2020
    'Si es codopera = REA  o codopera = DUA T
    If R!CodOpera = 5 Or R!CodOpera = 6 Then
        If R!CodOpera = 5 Then
            DevuelveTipoFacturaRecibida = "F6"
        Else
            DevuelveTipoFacturaRecibida = "F5"
        End If
        Exit Function
    End If
    
    
    
    If R!codconce340 = "D" Then
        'Rectificativa
        DevuelveTipoFacturaRecibida = "R1"

    ElseIf R!codconce340 = "J" Or R!codconce340 = "B" Then
            DevuelveTipoFacturaRecibida = "F2"
    
    Else
        'NORMAL
        DevuelveTipoFacturaRecibida = "F1"
    End If
    
End Function




'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************
'
'  Sistema de avisos de mensajes. Para que no este dando los mensajes a todas horas
'
'
'********************************************************************************
'********************************************************************************

'
'
'Private Sub ComprobarTablaFechas()
'    On Error Resume Next
'
'    Conn.Execute "Select * from usuarios.wavisoscontabilizacion where false"
'    If Err.Number <> 0 Then
'        Err.Clear
'        CrearTableTablasFechas
'    End If
'
'
'
'End Sub
'
'Private Sub CrearTableTablasFechas()
'Dim cad As String
'
'    cad = "CREATE TABLE usuarios.wavisoscontabilizacion ("
'    cad = cad & "login varchar(20) NOT NULL DEFAULT '0',"
'    cad = cad & "aplicacion tinyint(4) NOT NULL DEFAULT '0',"
'    cad = cad & "codempre smallint(1) unsigned NOT NULL DEFAULT '0',"
'    cad = cad & "ultaviso datetime DEFAULT NULL,"
'    cad = cad & "PRIMARY KEY (`login`,`aplicacion`,`codempre`)"
'    cad = cad & ") ENGINE=MyISAM ;"
'
'
'    Ejecuta cad
'End Sub


Public Function DarAvisoPendientesSII() As Byte
Dim cad As String
Dim FecUltAviso As Date
Dim Horas As Long
Dim VerSiDamosAviso As Boolean
Dim Mensaje As String
Dim Resul As Byte


    Mensaje = ""
    
    Resul = TieneFacturasPendientesSubirSII()
    If Resul > 0 Then
        '
        'cad = "replace into usuarios.wavisoscontabilizacion(`login`,`aplicacion`,`codempre`,`ultaviso`) values ("
        'cad = cad & DBSet(vUsu.Login, "T") & ",'2'," & vEmpresa.codempre & "," & DBSet(Now, "FH") & ")"
        'Ejecuta cad

        DarAvisoPendientesSII = Resul
    End If


End Function




Public Function UltimaFechaCorrectaSII(DiasAVisoSII As Integer, FechaPresentacion As Date) As Date
Dim DiaSemanaPresen As Integer
Dim DiaSemanaUltimoDiaPresentar As Integer
Dim F As Date

Dim Resta As Integer

    If DiasAVisoSII > 5 Then
        
        UltimaFechaCorrectaSII = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
        

    Else
        DiaSemanaPresen = WeekDay(FechaPresentacion, vbMonday)
       
                
                
        If DiaSemanaPresen >= 6 Then
            'Si presento el sabado o el domingo tengo mas dias ,  1 o dos
            If DiaSemanaPresen = 6 Then
                Resta = DiasAVisoSII
            Else
                Resta = DiasAVisoSII + 1
            End If
        Else
            F = DateAdd("d", -DiasAVisoSII, FechaPresentacion)
            DiaSemanaUltimoDiaPresentar = WeekDay(F, vbMonday)
            
            If DiaSemanaUltimoDiaPresentar > DiaSemanaPresen Then
                Resta = DiasAVisoSII + 2
            
            Else
                'Directamente la resta son 4
                Resta = DiasAVisoSII
            End If
        End If
        UltimaFechaCorrectaSII = DateAdd("d", -Resta, FechaPresentacion)
    End If

    UltimaFechaCorrectaSII = Format(UltimaFechaCorrectaSII, "dd/mm/yyyy")

End Function
'************** RUTINA COPMPROBACION
'   Dim fin As Boolean
'    fin = False
'
'    Dim F As Date
'    Dim F2 As Date
'    Dim Cad As String
'    Dim c2 As String
'    Dim I As Integer
'
'    Do
'        Cad = ""
'        For I = 1 To 28
'            F = CDate(Format(I, "00") & "/02/2018")
'
'            F2 = UltimaFechaCorrectaSII(3, F)
'
'
'            c2 = F & "  " & Weekday(F, vbMonday) & " --> "
'            c2 = c2 & F2 & "  " & Weekday(F2, vbMonday)
'            Cad = Cad & c2 & vbCrLf
'        Next
'
'        MsgBox Cad, vbExclamation
'
'
'
'
'    Loop Until fin

