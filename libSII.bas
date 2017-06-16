Attribute VB_Name = "libSII"
Option Explicit




'********************************************************
'  0 No tiene     1 Clientes     2 Proveedores   3 Ambos
Private Function TieneFacturasPendientesSubirSII() As Byte
Dim cad As String
Dim F As Date
Dim Aux As String

    TieneFacturasPendientesSubirSII = 0   'No tiene
    
    If vUsu.Nivel > 0 Then Exit Function
    If Not vParam.SIITiene Then Exit Function
    
    
    'Primer SQl. Ver si ha facturas que ha sido subida, pero no sabemos el resultado
    ComprobarResultadoEnviadasASWII 0, Nothing
        
    
    
    F = DateAdd("d", vParam.SIIDiasAviso, Now)  'Han pasado los x Dias en parametros
    
    'incio fecha sii
    cad = " fecfactu >= " & DBSet(vParam.SIIFechaInicio, "T")
    cad = cad & " AND fecfactu <= " & DBSet(F, "F")
    cad = cad & " AND ( coalesce(SII_ID,0) =0  or "
    cad = cad & " (SII_ID>0 and SII_status<3 )) "
    
    Aux = DevuelveDesdeBD("count(*)", "factcli", cad & " AND 1", " 1 ", "N")
    If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = TieneFacturasPendientesSubirSII + 1
    
    Aux = DevuelveDesdeBD("count(*)", "factpro", cad & " AND 1", " 1 ", "N")
    If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = TieneFacturasPendientesSubirSII + 2

    If TieneFacturasPendientesSubirSII = 0 Then
        
        'Veremos si tiene pendientes
        
        Aux = "10"
        Aux = DevuelveDesdeBD("count(*)", "aswsii.envio_facturas_emitidas", "enviada", "0", "N")
        If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = TieneFacturasPendientesSubirSII + 2
    
        Aux = DevuelveDesdeBD("count(*)", "aswsii.envio_facturas_recibidas", "enviada", "0", "N")
        If Val(Aux) > 0 Then TieneFacturasPendientesSubirSII = TieneFacturasPendientesSubirSII + 2

    End If


End Function

'Las facturas las grabamos en aswii. Comprobaremos que han sido subidas. O bien correctamente, o bien incorrectamente
'   Opcion: 0 Todo   1CLIENTES    2 PROVEE
Public Sub ComprobarResultadoEnviadasASWII(Opcion As Byte, L As Label)

    If Not L Is Nothing Then
        L.Caption = "Leyendo registro ASWSII"
        L.Refresh
    End If
    
    If Opcion <> 2 Then ComprobarResultadoTipoFra True, L
    If Opcion <> 1 Then ComprobarResultadoTipoFra False, L
    
    
    
End Sub



Private Sub ComprobarResultadoTipoFra(Emitidas As Boolean, ByRef L As Label)
Dim cad As String
Dim R As ADODB.Recordset
Dim NumFacturas As String
Dim C As Integer

    
    
    LblIndica L, "Leyendo BD " & IIf(Emitidas, "Emitidas", "Recibidas")
    
    If Emitidas Then
        cad = "select * from factcli where SII_ID >0 and SII_status =0"
    Else
        cad = "select * from factpro where SII_ID >0 and SII_status =0"
    End If
    Set R = New ADODB.Recordset
    R.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = 0
    cad = ""
    While Not R.EOF
        C = C + 1
        cad = cad & ", " & R!SII_ID
        
        If C > 10 Then
            LblIndica L, "Actualizando BD"
            
            ComprobarEnASWII_EstadoFacturas Emitidas, cad
            C = 0
            cad = ""
            
            LblIndica L, "Leyendo registros"
        End If
        R.MoveNext
    Wend
    R.Close
    
    If C > 0 Then
        LblIndica L, "Actualizando 2"
        ComprobarEnASWII_EstadoFacturas Emitidas, cad
    End If
    LblIndica L, ""
End Sub






Private Sub LblIndica(ByRef LL As Label, TEXTO As String)
    If Not LL Is Nothing Then
        LL.Caption = TEXTO
        LL.Refresh
    End If
End Sub

Private Sub ComprobarEnASWII_EstadoFacturas(Emitidas As Boolean, CadenaConFacturas As String)
Dim R2 As ADODB.Recordset
Dim Aux As String

    On Error GoTo eComprobarEnASWII_EstadoFacturas
    Set R2 = New ADODB.Recordset
    If Mid(CadenaConFacturas, 1, 1) = "," Then CadenaConFacturas = Mid(CadenaConFacturas, 2)
    If Emitidas Then
        Aux = "Select  IDEnvioFacturasEmitidas, resultado from aswsii.envio_facturas_emitidas WHERE enviada=1 and IDEnvioFacturasEmitidas IN (" & CadenaConFacturas & ")"
    Else
        Aux = "Select  IDEnvioFacturasRecibidas, resultado from aswsii.envio_facturas_recibidas WHERE enviada=1 and IDEnvioFacturasRecibidas IN (" & CadenaConFacturas & ")"
    End If
    R2.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not R2.EOF
        
        If UCase(R2!Resultado) = "ERROR" Then
            Aux = "1"
        ElseIf UCase(R2!Resultado) = "INCORRECTO" Then
            Aux = "2"
        ElseIf UCase(R2!Resultado) = "ACEPTADOCONERRORES" Then
            Aux = "3"
        Else
            Aux = "4"
        End If
        
        
        
        If Emitidas Then
            Aux = "UPDATE factcli SET SII_status =" & Aux
            Aux = Aux & " WHERE SII_ID = " & R2!IDEnvioFacturasEmitidas
        Else
            Aux = "UPDATE factpro SET SII_status =" & Aux
            Aux = Aux & " WHERE SII_ID = " & R2!IDEnvioFacturasRecibidas
        End If
        
        
        Conn.Execute Aux
        
        R2.MoveNext
    Wend
    R2.Close
    
eComprobarEnASWII_EstadoFacturas:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set R2 = Nothing
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
Public Function Sii_FraCLI(Serie As String, NumFac As Long, Anofac As Integer, IDEnvioFacturasEmitidas As Long, ByRef SQL_Insert As String) As Boolean
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
Dim BloqueIVA As Byte


    On Error GoTo eSii_FraCLI
    Sii_FraCLI = False
    
    Sql = "Select * from factcli where numserie =" & DBSet(Serie, "T") & " AND numfactu =" & NumFac & " AND anofactu =" & Anofac
    Set RN = New ADODB.Recordset
    RN.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""

'#1
    'IDEnvioFacturasEmitidas,Origen,FechaHoraCreacion,EnvioInmediato,          'Enviada,Resultado: NO los pongo en el insert
    Sql = IDEnvioFacturasEmitidas & ",'ARICONTA'," & DBSet(Now, "FH") & ",1,"

'#2
    'CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, REG_PI_Ejercicio, REG_PI_Periodo
    'FALTA
    'Sql = Sql & "'0.7'," & DBSet(vEmpresa.nomempre, "T") & ",NULL," & DBSet(vEmpresa.NIF, "T") & ",'A0'," & Year(RN!FecFactu) & ","
    '
    Sql = Sql & "'0.7'," & DBSet("Ariadna Software SL", "T") & ",NULL," & DBSet("B96470190", "T") & ",'A0'," & Year(RN!FecFactu) & ","
    
    
    
    
    
    Sql = Sql & "'" & Format(Month(RN!FecFactu), "00") & "',"
    
'#3
    'REG_IDF_IDEF_NIF,REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura
    'Sql = Sql & DBSet(RN!nifdatos, "T") & "," & DBSet(RN!NUmSerie & Format(RN!NumFactu, "0000000"), "T") & ","
    Sql = Sql & DBSet("B96470190", "T") & "," & DBSet(RN!NUmSerie & Format(RN!NumFactu, "0000000"), "T") & ","
    
    
    
    'Si son de tickets agrupados deberiamos poner primera y ultima. De momento null
    Sql = Sql & "NULL," & DBSet(RN!FecFactu, "F") & ","
    
    '#3.1
    ',REG_FE_TipoRectificativa,REG_FE_IR_BaseRectificada,REG_FE_IR_CuotaRectificada,REG_FE_IR_CuotaRecargoRectificado,
        
    Clave = DevuelveTipoFacturaEmitida(RN)   'Ver hoja. Hay tipos:    f1 factura   f2 tiket    r1 rectificativas
    Aux = ""
    Sql = Sql & DBSet(Clave, "T") & ","
  
    If Clave = "R1" Then
        Aux = "I"  'factura rectificativa por sustitcion
        Sql = Sql & "," & DBSet(Aux, "T", "S") & ","
        'Deberiamos obtener los importes
        Sql = Sql & "12,12,12,"
    Else
        'Los cuatro campos de la rectificativa a NULL
        Sql = Sql & "null,null,null,null,"

    End If
    
'#4
    'REG_FE_ClaveRegimenEspecialOTrascendencia,REG_FE_ImporteTotal,REG_FE_BaseImponibleACoste,REG_FE_DescripcionOperacion
    Clave = DevuelveClaveTranscendenciaEmitida(RN)
    Sql = Sql & DBSet(Clave, "T") & "," & DBSet(RN!totfaccl, "N") & ",NULL,'Factura " & RN!NUmSerie & RN!NumFactu & "',"
    
    
'#5
    'REG_FE_EmitidaPorTercero,REG_FE_CNT_NombreRazon,REG_FE_CNT_NIF,REG_FE_CNT_IDOtro_CodigoPais,REG_FE_CNT_IDOtro_IDType,REG_FE_CNT_IDOtro_ID,
    Aux = DBLet(RN!Nommacta, "T")

    
    Sql = Sql & "NULL," & DBSet(Aux, "T") & ","
    
    'NIF. Para las intracoms el NIF debe llevar las letras
    BloqueIVA = 0 'NORMAL
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOMUNITARIAS    EXTRANJERO
        'PAIS doc
            
        C2 = DBSet(DBLet(RN!codPAIS, "T"), "T", "S")
        If RN!CodOpera = 1 Then
            Aux = DBLet(RN!codPAIS, "T") & DBLet(RN!nifdatos, "T")
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
        Aux = DBLet(RN!nifdatos, "T")
        C2 = "null"
        Sql = Sql & DBSet(Aux, "T", "N") & "," & C2 & "," & C1 & ",NULL,"
    End If
    
   
        
    
'6#
    'EXENTA
    
    
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOM y exportacion
        LlevaIVAs = False
        If RN!CodOpera = 1 Then
            Aux = "'E5'," 'intra
        Else
            Aux = "'E2',"  'export
        End If
        
        Aux = Aux & DBSet(RN!totbases, "N") & ",null"
    Else
        LlevaIVAs = True
        Aux = "NULL,NULL,'S1'"
    End If
    Sql = Sql & Aux
    
    RN.Close
    
'7#
    'Bloque desglose IVAS hasta 6 ivas. Cambia el numerito ...DT1   DT2..
    CadenaIVAS = ""
    NumIVas = 0
    If LlevaIVAs Then
        
        Aux = "Select * from factcli_totales where numserie =" & DBSet(Serie, "T") & " AND numfactu =" & NumFac & " AND anofactu =" & Anofac
        RN.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
        'TipoImpositivo,BaseImponible,CuotaRepercutida,TipoREquivalencia,CuotaREquivalencia,"
        While Not RN.EOF
            
            Aux = "," & DBSet(RN!porciva, "N") & "," & DBSet(RN!Baseimpo, "N") & "," & DBSet(RN!Impoiva, "N") & ","
            If IsNull(RN!porcrec) Then
                Aux = Aux & "NULL,NULL"
            Else
                Aux = Aux & DBSet(RN!porcrec, "N") & "," & DBSet(RN!ImpoRec, "N")
            End If
            CadenaIVAS = CadenaIVAS & Aux
            NumIVas = NumIVas + 1
            RN.MoveNext
        Wend
        RN.Close
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
    SQL_Insert = Sii_FraCLI_SQL(BloqueIVA) & ") VALUES (" & Sql & ")"
    
    Sii_FraCLI = True
    
eSii_FraCLI:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RN = Nothing
End Function


'  0.- Facturas normales                ->  REG_FE_TD_DF_SU
'  1.- Intracomunitarias // Extranjera  ->  REG_FE_TD_DTE_SU
Private Function Sii_FraCLI_SQL(BloquesIVA As Byte) As String
Dim cad As String
Dim H As Integer

    Sii_FraCLI_SQL = "INSERT INTO aswsii.envio_facturas_emitidas("
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

    ElseIf R!codconce340 = "J" Or R!codconce340 = "B" Then
            DevuelveTipoFacturaEmitida = "F2"
    
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
        DevuelveClaveTranscendenciaEmitida = "01"
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
    End If

    
End Function

































'****************************************************************************
'****************************************************************************
'
' RECIBIDAS
'
'****************************************************************************
'****************************************************************************
Public Function Sii_FraPRO(Serie As String, numregis As Long, Anofac As Integer, IDEnvioFacturasRecibidas As Long, ByRef SQL_Insert As String) As Boolean
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



    On Error GoTo eSii_FraCLI
    Sii_FraPRO = False
    
    Sql = "Select * from factpro where numserie =" & DBSet(Serie, "T") & " AND numregis =" & numregis & " AND anofactu =" & Anofac
    Set RN = New ADODB.Recordset
    RN.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Sql = ""

'#1
    'IDEnvioFacturasEmitidas,Origen,FechaHoraCreacion,EnvioInmediato,          'Enviada,Resultado: NO los pongo en el insert
    Sql = IDEnvioFacturasRecibidas & ",'ARICONTA'," & DBSet(Now, "FH") & ",1,"

'#2
    'CAB_IDVersionSii , CAB_Titular_NombreRazon, CAB_Titular_NIFRepresentante, CAB_Titular_NIF, REG_PI_Ejercicio, REG_PI_Periodo
    'FALTA
    'Sql = Sql & "'0.7'," & DBSet(vEmpresa.nomempre, "T") & ",NULL," & DBSet(vEmpresa.NIF, "T") & ",'A0'," & Year(RN!fecharec) & "," & "'" & Format(Month(RN!fecharec), "00") & "',"
    '
    Sql = Sql & "'0.7'," & DBSet("Ariadna Software SL", "T") & ",NULL," & DBSet("B96470190", "T") & ",'A0'," & Year(RN!fecharec) & "," & "'" & Format(Month(RN!fecharec), "00") & "',"
    
'#3
    'REG_IDF_IDEF_NIF,REG_IDF_IDEF_IDOtro_CodigoPais,REG_IDF_IDEF_IDOtro_IDType,REG_IDF_IDEF_IDOtro_ID
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOMUNITARIAS    EXTRANJERO
        'PAIS doc
        C2 = DBSet(DBLet(RN!codPAIS, "T"), "T", "S")
        If RN!CodOpera = 1 Then
            Aux = DBLet(RN!codPAIS, "T") & DBLet(RN!nifdatos, "T")
            C1 = "'02'"
        Else
            Aux = DBLet(RN!nifdatos, "T")
            C1 = "'03'"
        End If
        Sql = Sql & "''" & "," & C2 & "," & C1 & "," & DBSet(Aux, "T", "N") & ","
    Else
        'EL NIF
        'NO hacemos nada  AUX y c1 ya teiene los valores que toca
        C1 = "null"
        Aux = DBLet(RN!nifdatos, "T")
        C2 = "null"
        Sql = Sql & DBSet(Aux, "T", "N") & "," & C2 & "," & C1 & ",NULL,"
    End If
    
    
'#4
    'REG_IDF_NumSerieFacturaEmisor,REG_IDF_NumSerieFacturaEmisorResumenFin,REG_IDF_FechaExpedicionFacturaEmisor,REG_FE_TipoFactura,REG_FE_TipoRectificativa
    'Si son de tickets agrupados deberiamos poner primera y ultima. De momento null
    Sql = Sql & DBSet(RN!NumFactu, "T") & "," & "NULL," & DBSet(RN!FecFactu, "F") & ","
    Clave = DevuelveTipoFacturaRecibida(RN)
    Aux = ""
    If Clave = "D" Then Aux = "S"  'factura rectificativa por sustitcion
    Sql = Sql & DBSet(Clave, "T") & "," & DBSet(Aux, "T", "S") & ","
    
    
'#5
     
    'REG_FE_ClaveRegimenEspecialOTrascendencia,REG_FE_ImporteTotal,REG_FE_BaseImponibleACoste,REG_FE_DescripcionOperacion
    Clave = DevuelveClaveTranscendenciaRecibida(RN)
    Sql = Sql & DBSet(Clave, "T") & "," & DBSet(RN!totfacpr, "N") & ",NULL,'Factura" & IIf(RN!NUmSerie = 1, "", " ser: " & RN!NUmSerie) & " " & RN!NumFactu & "',"
    
    
'#6
    'REG_FE_EmitidaPorTercero,REG_FE_CNT_NombreRazon,REG_FE_CNT_NIF,REG_FE_CNT_IDOtro_CodigoPais,REG_FE_CNT_IDOtro_IDType,REG_FE_CNT_IDOtro_ID,
    Aux = DBLet(RN!Nommacta, "T")

    
    Sql = Sql & DBSet(Aux, "T") & ","
    
    'NIF. Para las intracoms el NIF debe llevar las letras
    If RN!CodOpera = 1 Or RN!CodOpera = 2 Then
        'INTRACOMUNITARIAS    EXTRANJERO
        'PAIS doc
        C2 = DBSet(DBLet(RN!codPAIS, "T"), "T", "S")
        If RN!CodOpera = 1 Then
            Aux = DBLet(RN!codPAIS, "T") & DBLet(RN!nifdatos, "T")
            C1 = "'02'"
        Else
            Aux = DBLet(RN!nifdatos, "T")
            C1 = "'03'"
        End If
        Sql = Sql & "''" & "," & C2 & "," & C1 & "," & DBSet(Aux, "T", "N") & ","
    Else
        'EL NIF
        'NO hacemos nada  AUX y c1 ya teiene los valores que toca
        C1 = "null"
        Aux = DBLet(RN!nifdatos, "T")
        C2 = "null"
        Sql = Sql & DBSet(Aux, "T", "N") & "," & C2 & "," & C1 & ",NULL,"
    End If
    
    
    '#7  REG_FR_FechaOperacion  REG_FR_FechaRegContable  REG_FR_CuotaDeducible
    CodOpera = RN!CodOpera
    InversionSujetoPasivo = False
    If CodOpera = 4 Then InversionSujetoPasivo = True
        
    Sql = Sql & DBSet(RN!FecFactu, "F") & "," & DBSet(RN!fecharec, "F") & ",#@#@#@$$$$"   'Sumaremos el total de cuotas deducibles y luego haremos un replace

    
    
    
    RN.Close
    
    TotalDecucible = 0
    
    
    
'#8 Inversion sujeto apsivo   ***ISP^^^^
    'hasta 6 ivas. Cambia el numerito ...DT1   DT2..   REG_FR_DF_ISP_DI_DT6_CuotaREquivalencia
    
    CadenaIVAS = ""
    NumIVas = 0
    If InversionSujetoPasivo Then
        
        Aux = "Select * from factpro_totales where numserie =" & DBSet(Serie, "T") & " AND numregis =" & numregis & " AND anofactu =" & Anofac
        RN.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        'TipoImpositivo,BaseImponible,CuotaRepercutida,TipoREquivalencia,CuotaREquivalencia,"
        While Not RN.EOF
            
            Aux = "," & DBSet(RN!porciva, "N") & "," & DBSet(RN!Baseimpo, "N") & "," & DBSet(RN!Impoiva, "N") & ","
            If IsNull(RN!porcrec) Then
                Aux = Aux & "NULL,NULL"
            Else
                Aux = Aux & DBSet(RN!porcrec, "N") & "," & DBSet(RN!ImpoRec, "N")
            End If
            CadenaIVAS = CadenaIVAS & Aux
            NumIVas = NumIVas + 1
            
            TotalDecucible = RN!Impoiva + DBLet(RN!ImpoRec, "N")
            
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
        
        Aux = "Select * from factpro_totales where numserie =" & DBSet(Serie, "T") & " AND numregis =" & numregis & " AND anofactu =" & Anofac
        RN.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
        'TipoImpositivo,BaseImponible,CuotaRepercutida,TipoREquivalencia,CuotaREquivalencia,"

        
        While Not RN.EOF
        
            If CodOpera = 5 Then
                'Si el tipo de IVA es REA
                Aux = ",null," & DBSet(RN!Baseimpo, "N") & ",null,null,null,"
                '% REA impor REA
                Aux = Aux & DBSet(RN!porciva, "N") & "," & DBSet(RN!Impoiva, "N")
            
                
            Else
                Aux = "," & DBSet(RN!porciva, "N") & "," & DBSet(RN!Baseimpo, "N") & "," & DBSet(RN!Impoiva, "N") & ","
                If IsNull(RN!porcrec) Then
                    Aux = Aux & "NULL,NULL"
                Else
                    Aux = Aux & DBSet(RN!porcrec, "N") & "," & DBSet(RN!ImpoRec, "N")
                End If
                Aux = Aux & ",NULL,NULL"             'REA A null
            End If
            
            CadenaIVAS = CadenaIVAS & Aux
            NumIVas = NumIVas + 1
            TotalDecucible = RN!Impoiva + DBLet(RN!ImpoRec, "N")
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
    SQL_Insert = Sii_FraPRO_SQL & ") VALUES (" & Sql & ")"
    
    Sii_FraPRO = True
    
eSii_FraCLI:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RN = Nothing
End Function


Private Function Sii_FraPRO_SQL() As String
Dim cad As String
Dim H As Integer

    Sii_FraPRO_SQL = "INSERT INTO aswsii.envio_facturas_recibidas("
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
'********************************************************************************
Public Sub ComprobarEstrcuturaAvisosSII()
    
    'Solo para nuevas contabilidad
    If Not vParam.SIITiene Then Exit Sub
    
    'Solo usuarios con nivel 0-1
    If vUsu.Nivel > 1 Then Exit Sub
    
    ComprobarTablaFechas
    
    
    
End Sub


Private Sub ComprobarTablaFechas()
    On Error Resume Next
    
    Conn.Execute "Select * from usuarios.wavisoscontabilizacion where false"
    If Err.Number <> 0 Then
        Err.Clear
        CrearTableTablasFechas
    End If
    
    
    
End Sub

Private Sub CrearTableTablasFechas()
Dim cad As String
    
    cad = "CREATE TABLE usuarios.wavisoscontabilizacion ("
    cad = cad & "login varchar(20) NOT NULL DEFAULT '0',"
    cad = cad & "aplicacion tinyint(4) NOT NULL DEFAULT '0',"
    cad = cad & "codempre smallint(1) unsigned NOT NULL DEFAULT '0',"
    cad = cad & "ultaviso datetime DEFAULT NULL,"
    cad = cad & "PRIMARY KEY (`login`,`aplicacion`,`codempre`)"
    cad = cad & ") ENGINE=MyISAM ;"
    
    
    Ejecuta cad
End Sub


Public Function DarAvisoPendientesSII() As Byte
Dim cad As String
Dim FecUltAviso As Date
Dim Horas As Long
Dim VerSiDamosAviso As Boolean
Dim Mensaje As String
Dim Resul As Byte


    '      1: ariges   2:Ariconta6
    DarAvisoPendientesSII = 0
    cad = "aplicacion = 2 AND codempre = " & vEmpresa.codempre & " AND login "
    cad = DevuelveDesdeBD("ultaviso", "usuarios.wavisoscontabilizacion", cad, vUsu.Login, "T")
    If cad = "" Then
       FecUltAviso = DateAdd("yyyy", -1, Now)
    Else
        FecUltAviso = CDate(cad)
    End If
    
    VerSiDamosAviso = False
    If Year(FecUltAviso) - Year(Now) > 1 Then
        VerSiDamosAviso = True
    Else
        'Si hay mas de un dia de diferencia
        Horas = DateDiff("d", FecUltAviso, Now)
        If Horas > 0 Then
            VerSiDamosAviso = True
        Else
            
            Horas = DateDiff("h", FecUltAviso, Now)
            If Horas > 4 Then VerSiDamosAviso = True
        End If
    End If
    
    If Not VerSiDamosAviso Then Exit Function
    Mensaje = ""
    
    Resul = TieneFacturasPendientesSubirSII
    If Resul > 0 Then
        '
        cad = "replace into usuarios.wavisoscontabilizacion(`login`,`aplicacion`,`codempre`,`ultaviso`) values ("
        cad = cad & DBSet(vUsu.Login, "T") & ",'2'," & vEmpresa.codempre & "," & DBSet(Now, "FH") & ")"
        Ejecuta cad

        DarAvisoPendientesSII = Resul
    End If


End Function


