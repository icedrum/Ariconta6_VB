Attribute VB_Name = "libAsegurados"
Option Explicit





Public Function Asegurados_HayAvisos(DeFalta As Boolean) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String

    Asegurados_HayAvisos = False
    If DeFalta Then
        MontaSQLAvisosFalta SQL
    Else
        MontaSQLAvisoSiniestro SQL
    End If
    SQL = "Select count(*) " & SQL
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then Asegurados_HayAvisos = True
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
    
End Function



Public Sub MontaSQLAvisosSeguros(Faltas As Boolean, SQ As String)
    If Faltas Then
        MontaSQLAvisosFalta SQ
    Else
        MontaSQLAvisoSiniestro SQ
    End If
End Sub

Private Sub MontaSQLAvisosFalta(ByRef SQ As String)

    



    SQ = " FROM cobros,cuentas WHERE cobros.codmacta = cuentas.codmacta AND cuentas.numpoliz<>'' and situacionjuri=0"
    'Fecha factura > que cuando implantamos el sistema de asegurados

    'No es desde inicio aseguradoras, sino desde inicio SOCIO asegurado
    SQ = SQ & " AND fecfactu >= '" & Format(vParamT.FechaIniSeg, FormatoFecha) & "'"
    SQ = SQ & " AND fecfactu >= fecconce "
    
    'Que no esten avisados ya
    SQ = SQ & " AND feccomunica is null and fecprorroga is null and fecsiniestro is null"
    'Fecha vto entre los datos de parametros
    'FALTA###   Pongo el de abajo pq si no no me salen datos
    'SQ = SQ & " DATEDIFF(curdate(),fecvenci) between " & vParam.DiasMaxAvisoD & " AND " & vParam.DiasMaxAvisoH
    
    
    
    If vParamT.FechaSeguroEsFra Then
        'ALZIRA
        SQ = SQ & " AND DATEDIFF(curdate(),fecfactu) >= " & vParamT.DiasMaxAvisoD
    Else
        'HERBELCA
        SQ = SQ & " AND DATEDIFF(curdate(),fecvenci) >= " & vParamT.DiasMaxAvisoD
    End If
    
    'Mayo 2012
    SQ = SQ & " AND if(fecbajcre is null,True,fecfactu<=fecbajcre)"
    
    'Solo importes positivos, que esten pendientes de pago
    SQ = SQ & " AND impvenci>0   AND impvenci+coalesce(gastos,0)>coalesce(impcobro,0)"
    
End Sub


Private Sub MontaSQLAvisoSiniestro(ByRef SQ As String)
Dim F As Date
    SQ = " FROM cobros,cuentas WHERE cobros.codmacta = cuentas.codmacta AND cuentas.numpoliz<>'' and situacionjuri=0"
    SQ = SQ & " AND fecfactu >= fecconce "
    SQ = SQ & " AND if(fecbajcre is null,True,fecfactu<=fecbajcre)"
    SQ = SQ & " AND fecsiniestro is null  AND impvenci+coalesce(gastos,0)>coalesce(impcobro,0)"
    
    'Que ya ha sido comunicado o prorrogado
    SQ = SQ & " AND (not feccomunica is null or not fecprorroga is null)"
    
    'Segun fecacmunica o prroroga
    SQ = SQ & " and  DATEDIFF(curdate(),if(feccomunica is null,fecprorroga,feccomunica )) >=if(feccomunica is null," & vParamT.DiasAvisoDesdeProrroga & "," & vParamT.DiasMaxSiniestroD & ")"
    
 
    
    
End Sub
