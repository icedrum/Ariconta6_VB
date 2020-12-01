Attribute VB_Name = "BaseDato"
Option Explicit

Private Sql As String

Dim ImpD As Currency
Dim ImpH As Currency
Dim RT As ADODB.Recordset


Dim d As String
Dim H As String
'Para los balances
Dim M1 As Integer   ' años y kmeses para el balance
Dim M2 As Integer

Dim M22 As Long

Dim M3 As Integer
Dim A1 As Integer
Dim A2 As Integer
Dim A3 As Integer
Dim vCta As String
Dim vDig As Byte
Dim ImAcD As Currency  'importes
Dim ImAcH As Currency
Dim ImPerD As Currency  'importes
Dim ImPerH As Currency
Dim ImCierrD As Currency  'importes
Dim ImCierrH As Currency
Dim Contabilidad As Integer
Dim Aux As String
Dim vFecha1 As Date
Dim vFecha2 As Date
Dim VFecha3 As Date
Dim Codigo As String
Dim EjerciciosCerrados As Boolean
Dim NumAsiento As Integer
Dim Nulo1 As Boolean
Dim Nulo2 As Boolean

Dim FIniPeriodo As Date
Dim FFinPeriodo As Date

Dim VarConsolidado(2) As String

Dim EsBalancePerdidas_y_ganancias As Boolean

'Para la precarga de datos del balance de sumas y saldos
Dim RsBalPerGan As ADODB.Recordset
Dim RsApertura As ADODB.Recordset


'--------------------------------------------------------------------
'--------------------------------------------------------------------
Private Function ImporteASQL(ByRef Importe As Currency) As String
ImporteASQL = ","
If Importe = 0 Then
    ImporteASQL = ImporteASQL & "NULL"
Else
    ImporteASQL = ImporteASQL & TransformaComasPuntos(CStr(Importe))
End If
End Function



'--------------------------------------------------------------------
'--------------------------------------------------------------------
' El dos sera para k pinte el 0. Ya en el informe lo trataremos.
' Con esta opcion se simplifica bastante la opcion de totales
Private Function ImporteASQL2(ByRef Importe As Currency) As String
    ImporteASQL2 = "," & TransformaComasPuntos(CStr(Importe))
End Function



'--------------------------------------------------------------------
'--------------------------------------------------------------------



Public Sub CommitConexion()
    On Error Resume Next
    Conn.Execute "Commit"
    If Err.Number <> 0 Then Err.Clear
End Sub




'EsClave: Si es clave, en los char NO forzaremos los *
Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, Cadena As String, ByRef DevSQL As String, EsClave As Boolean) As Byte
Dim cad As String
Dim Aux As String
Dim Ch As String
Dim Fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    I = CararacteresCorrectos(Cadena, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo numerico
        cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
        If Not IsNumeric(cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = Campo & " >= " & cad & " AND " & Campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If Cadena = ">>" Or Cadena = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    I = 1
                    cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        Ch = Mid(Cadena, I, 1)
                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                            cad = cad & Ch
                            Else
                                Aux = Mid(Cadena, I)
                                Fin = True
                        End If
                        I = I + 1
                        If I > Len(Cadena) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If cad = "" Then cad = " = "
                    DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(Cadena, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo de fechas
        cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
        If Not EsFechaOKString(cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        cad = Format(cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & cad & "' AND " & Campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If Cadena = ">>" Or Cadena = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    Ch = Mid(Cadena, I, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        cad = cad & Ch
                        Else
                            Aux = Mid(Cadena, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(Cadena) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If cad = "" Then cad = " = "
                DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
    
    
    
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(Cadena, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If Cadena = ">>" Or Cadena = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo numerico

        cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
        
        'Intervalo correcto
        'Construimos la cadena
        cad = DevNombreSQL(cad)
        Aux = DevNombreSQL(Aux)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & cad & "' AND " & Campo & " <= '" & Aux & "'"
    
    
    Else
    
        'Cambiamos el * por % puesto que en ADO es el caraacter para like
        I = 1
        Aux = Cadena
        
        '++
        If Len(Aux) <> 0 Then
            If InStr(1, Aux, "*") = 0 Then
                If Not EsClave Then Aux = "*" & Aux & "*"
            End If
        End If
        '++
        
        
        While I <> 0
            I = InStr(1, Aux, "*")
            If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
        Wend
        'Cambiamos el ? por la _ pue es su omonimo
        I = 1
        While I <> 0
            I = InStr(1, Aux, "?")
            If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
        Wend
        cad = Mid(Cadena, 1, 2)
        If cad = "<>" Then
            Aux = Mid(Cadena, 3)
            DevSQL = Campo & " not LIKE '" & Aux & "'"
            Else
            DevSQL = Campo & " LIKE '" & Aux & "'"
        End If
    End If


    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, Cadena, "<>")
    If I = 0 Then
        'IGUAL A valor
        cad = " = "
        Else
            'Distinto a valor
        cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, Cadena, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = Campo & " " & cad & " " & Aux
   
   
Case "FH"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(Cadena, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo de fechas
        cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
        If Not EsFechaOKString(cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        cad = Format(cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & cad & " 00:00:00' AND " & Campo & " <= '" & Aux & " 23:59:59'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If Cadena = ">>" Or Cadena = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    Ch = Mid(Cadena, I, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        cad = cad & Ch
                        Else
                            Aux = Mid(Cadena, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(Cadena) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If cad = "" Then
                    cad = " = "
                    Campo = " DATE(" & Campo & ")"
                End If
                DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
    
   
   
   
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
Dim I As Integer
Dim Ch As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "a" To "z"
            Case "è", "é", "í" 'Añade Laura: 16/03/06
            Case "A" To "Z"
            Case "0" To "9"
            'QUITAR#### o no.
            'Modificacion hecha 26-OCT-2006.  Es para que meta la coma como caracter en la busqueda
            Case "*", "%", "?", "_", "\", "/", ":", ".", " ", "-", "," ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "#", "@", "$"
            Case "<", ">"
            Case "Ñ", "ñ"
            Case Else
                Error = True
                Exit For
                
        End Select
    Next I
Case "F"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vCad)
        Ch = Mid(vCad, I, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function
















'-------------------------------------------------------------------

Public Function CargaDatosConExt(ByRef Cuenta As String, fec1 As Date, fec2 As Date, ByRef vSql As String, ByRef DescCuenta As String, Optional DesdeCCoste As Boolean) As Byte
Dim ACUM As Double  'Acumulado anterior

On Error GoTo ECargaDatosConExt
CargaDatosConExt = 1

'Insertamos en los campos de cabecera de cuentas
NombreSQL DescCuenta
Sql = Cuenta & "    -    " & DescCuenta
Sql = "INSERT INTO tmpconextcab (codusu,cta,fechini,fechfin,cuenta) VALUES (" & vUsu.Codigo & ", '" & Cuenta & "','" & Format(fec1, "dd/mm/yyyy") & "','" & Format(fec2, "dd/mm/yyyy") & "','" & Sql & "')"
Conn.Execute Sql


''los totatales
'Dim T1, cad
'cad = "Cuenta: " & DescCuenta & vbCrLf
'T1 = Timer


If Not CargaAcumuladosTotales(Cuenta, DesdeCCoste) Then Exit Function
'cad = cad & "Acum Total:" & Format(Timer - T1, "0.000") & vbCrLf
'T1 = Timer

'Los caumulados anteriores
If Not CargaAcumuladosAnteriores(Cuenta, fec1, ACUM, DesdeCCoste) Then Exit Function
'cad = cad & "Anterior:   " & Format(Timer - T1, "0.000") & vbCrLf
'T1 = Timer

'GENERAMOS LA TBLA TEMPORAL
If DesdeCCoste Then
    If Not CargaTablaTemporalConExtCC(Cuenta, vSql, ACUM) Then Exit Function
Else
    If Not CargaTablaTemporalConExt(Cuenta, vSql, ACUM) Then Exit Function
End If


'cad = cad & "Tabla:    " & Format(Timer - T1, "0.000") & vbCrLf
'MsgBox cad


CargaDatosConExt = 0
Exit Function
ECargaDatosConExt:
    CargaDatosConExt = 2
    MuestraError Err.Number, "Cargando datos temporales. Cta: " & Cuenta, Err.Description
End Function



Private Function CargaAcumuladosTotales(ByRef Cta As String, Optional DesdeCCoste As Boolean) As Boolean
    CargaAcumuladosTotales = False
    If DesdeCCoste Then
        Sql = "SELECT Sum(perD) AS SumaDetimporteD, Sum(perH) AS SumaDetimporteH"
        Sql = Sql & " from tmplinccexplo where codccost='" & Cta & "'"
        Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N")
    Else
        Sql = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
        Sql = Sql & " from hlinapu where codmacta='" & Cta & "'"
    End If
    Sql = Sql & " AND fechaent >=  '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
        Else
        ImpD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
    End If
    RT.Close
    Set RT = Nothing
    Sql = "UPDATE tmpconextcab SET acumtotD= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    Sql = Sql & ", acumtotH= " & TransformaComasPuntos(CStr(ImpH)) 'Format(ImpH, "#,###,##0.00")
    ImpD = ImpD - ImpH
    Sql = Sql & ", acumtotT= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    Sql = Sql & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute Sql
    CargaAcumuladosTotales = True
End Function


Private Function CargaAcumuladosAnteriores(ByRef Cta As String, ByRef FI As Date, ByRef ACUM As Double, Optional DesdeCCoste As Boolean) As Boolean
Dim F1 As Date

    CargaAcumuladosAnteriores = False
    If DesdeCCoste Then
        Sql = "SELECT Sum(perD) AS SumaDetimporteD, Sum(perH) AS SumaDetimporteH"
        Sql = Sql & " from tmplinccexplo where codccost='" & Cta & "'"
        Sql = Sql & " and codusu = " & DBSet(vUsu.Codigo, "N")
    Else
        Sql = "SELECT Sum(timporteD) AS SumaDetimporteD, Sum(timporteH) AS SumaDetimporteH"
        Sql = Sql & " from hlinapu where codmacta='" & Cta & "'"
    End If
    F1 = vParam.fechaini

    Do
        If FI < F1 Then F1 = DateAdd("yyyy", -1, F1)
    Loop Until F1 <= FI
    'SQL = SQL & " AND fechaent >=  '" & Format(vParam.fechaini, FormatoFecha) & "'"
    Sql = Sql & " AND fechaent >=  '" & Format(F1, FormatoFecha) & "'"
    Sql = Sql & " AND fechaent <  '" & Format(FI, FormatoFecha) & "'"
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
    Else
        ImpD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
    End If
    RT.Close
    ACUM = ImpD - ImpH
    Sql = "UPDATE tmpconextcab SET acumantD= " & TransformaComasPuntos(CStr(ImpD))
    Sql = Sql & ", acumantH= " & TransformaComasPuntos(CStr(ImpH))
    Sql = Sql & ", acumantT= " & TransformaComasPuntos(CStr(ACUM))
    Sql = Sql & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute Sql
    Set RT = Nothing
    CargaAcumuladosAnteriores = True
End Function



Private Function CargaTablaTemporalConExt(Cta As String, vSele As String, ByRef ACUM As Double) As Boolean
Dim Aux As Currency
Dim ImporteD As String
Dim ImporteH As String
Dim Contador As Long
Dim RC As String

Dim Inserts As String  'Octubre 2013. Iba muy lento


On Error GoTo Etmpconext


'TIEMPOS
'Dim T1, Cadenita
'T1 = Timer
'Cadenita = "Cuenta: " & Cta & vbCrLf

CargaTablaTemporalConExt = False

'Conn.Execute "Delete from tmpconext where codusu =" & vUsu.Codigo
Set RT = New ADODB.Recordset
Sql = "Select * from hlinapu where codmacta='" & Cta & "'"
Sql = Sql & " AND " & vSele & " ORDER BY fechaent,numasien,linliapu"  'NO ESTABA linliapu, NO ME LO PUEDO CREER
RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

'Cadenita = Cadenita & "Select: " & Format(Timer - T1, "0.0000") & vbCrLf
'T1 = Timer

Sql = "INSERT INTO tmpconext (codusu, POS,numdiari, fechaent, numasien, linliapu, timporteD, timporteH, saldo, Punteada,nomdocum,ampconce,cta,contra,ccost) VALUES "
'ImpD = 0 ASI LLEVAMOS EL ACUMULADO
'ImpH = 0
Contador = 0
Inserts = ""
While Not RT.EOF
    Contador = Contador + 1
    If Not IsNull(RT!timported) Then
        Aux = DBLet(RT!timported, "N")
        ImpD = ImpD + Aux
        ImporteD = TransformaComasPuntos(RT!timported)
        ImporteH = "Null"
    Else
        Aux = DBLet(RT!timporteH, "N")
        ImporteD = "Null"
        ImporteH = TransformaComasPuntos(RT!timporteH)
        ImpH = ImpH + Aux
        Aux = -1 * Aux
    End If
    ACUM = ACUM + Aux
    
    'Insertar
    RC = vUsu.Codigo & "," & Contador & "," & RT!NumDiari & ",'" & Format(RT!FechaEnt, FormatoFecha) & "'," & RT!NumAsien & "," & RT!Linliapu & ","
    RC = RC & ImporteD & "," & ImporteH
    If RT!punteada <> 0 Then
        ImporteD = "SI"
        Else
        ImporteD = ""
    End If
    RC = RC & "," & TransformaComasPuntos(CStr(ACUM)) & ",'" & ImporteD & "',"
    RC = RC & DBSet(RT!Numdocum, "T") & "," & DBSet(RT!Ampconce, "T") & ",'" & Cta & "',"
'    If IsNull(RT!ctacontr) Then
'        RC = RC & "NULL"
'    Else
'        RC = RC & "'" & RT!ctacontr & "'"
'    End If
    RC = RC & DBSet(RT!ctacontr, "T")
    RC = RC & ","
    If IsNull(RT!CodCCost) Then
        RC = RC & "NULL"
    Else
        RC = RC & "'" & RT!CodCCost & "'"
    End If
    RC = RC & ")"
    
    
    'octubre 2013
    Inserts = Inserts & ", (" & RC
    If (Contador Mod 150) = 0 Then
        
        Inserts = Mid(Inserts, 2)
        Conn.Execute Sql & Inserts
        Inserts = ""
    End If
    'Sig
    RT.MoveNext
Wend
RT.Close

If Inserts <> "" Then
    Inserts = Mid(Inserts, 2)
    Conn.Execute Sql & Inserts
End If


'Cadenita = Cadenita & "Recorrer: " & Format(Timer - T1, "0.0000") & vbCrLf
'T1 = Timer

    Sql = "UPDATE tmpconextcab SET acumperD= " & TransformaComasPuntos(CStr(ImpD))
    Sql = Sql & ", acumperH= " & TransformaComasPuntos(CStr(ImpH))
    ImpD = ImpD - ImpH
    Sql = Sql & ", acumperT= " & TransformaComasPuntos(CStr(ImpD))
    Sql = Sql & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute Sql

    CargaTablaTemporalConExt = True



Exit Function
Etmpconext:
    MuestraError Err.Number, "Generando datos saldos" & vbCrLf & Err.Description
    Set RT = Nothing
End Function


Private Function CargaTablaTemporalConExtCC(Cta As String, vSele As String, ByRef ACUM As Double) As Boolean
Dim Aux As Currency
Dim ImporteD As String
Dim ImporteH As String
Dim Contador As Long
Dim RC As String

Dim Inserts As String  'Octubre 2013. Iba muy lento


On Error GoTo Etmpconext

    'TIEMPOS
    'Dim T1, Cadenita
    'T1 = Timer
    'Cadenita = "Cuenta: " & Cta & vbCrLf
    
    CargaTablaTemporalConExtCC = False
    
    'Conn.Execute "Delete from tmpconext where codusu =" & vUsu.Codigo
    Set RT = New ADODB.Recordset
    Sql = "Select * from tmplinccexplo where codccost='" & Cta & "'"
    Sql = Sql & " and codusu = " & vUsu.Codigo
    Sql = Sql & " AND " & vSele & " ORDER BY fechaent,numasien"
    RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Cadenita = Cadenita & "Select: " & Format(Timer - T1, "0.0000") & vbCrLf
    'T1 = Timer
    
    Sql = "INSERT INTO tmpconext (codusu, POS,numdiari, fechaent, numasien, linliapu, timporteD, timporteH, saldo, Punteada,nomdocum,ampconce,cta,contra,ccost) VALUES "
    'ImpD = 0 ASI LLEVAMOS EL ACUMULADO
    'ImpH = 0
    Contador = 0
    Inserts = ""
    While Not RT.EOF
        Contador = Contador + 1
        If Not IsNull(RT!perd) Then
            Aux = DBLet(RT!perd, "N")
            ImpD = ImpD + Aux
            ImporteD = TransformaComasPuntos(RT!perd)
            ImporteH = "Null"
        Else
            Aux = DBLet(RT!perh, "N")
            ImporteD = "Null"
            ImporteH = TransformaComasPuntos(RT!perh)
            ImpH = ImpH + Aux
            Aux = -1 * Aux
        End If
        ACUM = ACUM + Aux
        
        'Insertar
        RC = vUsu.Codigo & "," & Contador & "," & RT!NumDiari & ",'" & Format(RT!FechaEnt, FormatoFecha) & "'," & RT!NumAsien & "," & RT!LinApu & ","
        RC = RC & ImporteD & "," & ImporteH
        If RT!punteada <> 0 Then
            ImporteD = "SI"
            Else
            ImporteD = ""
        End If
        RC = RC & "," & TransformaComasPuntos(CStr(ACUM)) & ",'" & ImporteD & "',"
        RC = RC & DBSet(RT!DOCUM, "T") & "," & DBSet(RT!Ampconce, "T") & ",'" & Cta & "',"
    '    If IsNull(RT!ctacontr) Then
    '        RC = RC & "NULL"
    '    Else
    '        RC = RC & "'" & RT!ctacontr & "'"
    '    End If
        RC = RC & DBSet(RT!codmacta, "T")
        RC = RC & ","
        If IsNull(RT!CodCCost) Then
            RC = RC & "NULL"
        Else
            RC = RC & "'" & RT!CodCCost & "'"
        End If
        RC = RC & ")"
        
        
        'octubre 2013
        Inserts = Inserts & ", (" & RC
        If (Contador Mod 150) = 0 Then
            
            Inserts = Mid(Inserts, 2)
            Conn.Execute Sql & Inserts
            Inserts = ""
        End If
        'Sig
        RT.MoveNext
    Wend
    RT.Close
    
    If Inserts <> "" Then
        Inserts = Mid(Inserts, 2)
        Conn.Execute Sql & Inserts
    End If
    
    
    'Cadenita = Cadenita & "Recorrer: " & Format(Timer - T1, "0.0000") & vbCrLf
    'T1 = Timer
    
    Sql = "UPDATE tmpconextcab SET acumperD= " & TransformaComasPuntos(CStr(ImpD))
    Sql = Sql & ", acumperH= " & TransformaComasPuntos(CStr(ImpH))
    ImpD = ImpD - ImpH
    Sql = Sql & ", acumperT= " & TransformaComasPuntos(CStr(ImpD))
    Sql = Sql & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute Sql
    
    CargaTablaTemporalConExtCC = True
        
    'Cadenita = Cadenita & "Actualizar: " & Format(Timer - T1, "0.0000") & vbCrLf
    'MsgBox Cadenita
    
    
    
    Exit Function
Etmpconext:
    MuestraError Err.Number, "Generando datos saldos"
    Set RT = Nothing
End Function





Private Function HacerRepartoSubcentrosCoste() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim ImporteTot As Currency
Dim ImporteLinea As Currency
Dim UltSubCC As String
Dim NRegs As Long

    On Error GoTo eHacerRepartoSubcentrosCoste

    HacerRepartoSubcentrosCoste = False
    
    ' hacemos el desdoble
    Sql = "select * from tmpconext where codusu = " & DBSet(vUsu.Codigo, "N") & " and cta in (select ccoste.codccost from ccoste inner join ccoste_lineas on ccoste.codccost = ccoste_lineas.codccost) "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

'    Nregs = TotalRegistrosConsulta(SQL)
'
'    If Nregs <> 0 Then
'        pb2.Visible = True
'        CargarProgres pb2, Nregs
'    End If


    While Not Rs.EOF
'        IncrementarProgres pb2, 1
        
        Sql2 = "select ccoste.codccost, subccost, porccost from ccoste inner join ccoste_lineas on ccoste.codccost = ccoste_lineas.codccost where ccoste.codccost =  " & DBSet(Rs!Cta, "T")

        ImporteTot = 0
        UltSubCC = ""

        Set Rs2 = New ADODB.Recordset

        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            
            Sql = "INSERT INTO tmpconext (codusu, POS,numdiari, fechaent, numasien, linliapu, timporteD, timporteH, saldo, Punteada, nomdocum, ampconce, cta, contra, ccost, desdoblado) VALUES ("
            Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!Pos, "N") & "," & DBSet(Rs!NumDiari, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!NumAsien, "N") & "," & DBSet(Rs!Linliapu, "N") & ","
            
            
            If DBLet(Rs!timported, "N") <> 0 Then
                ImporteLinea = Round(DBLet(Rs!timported, "N") * DBLet(Rs2!porccost, "N") / 100, 2)
                Sql = Sql & DBSet(ImporteLinea, "N") & ",0,0,"
            Else
                ImporteLinea = Round(DBLet(Rs!timporteH, "N") * DBLet(Rs2!porccost, "N") / 100, 2)
                Sql = Sql & "0," & DBSet(ImporteLinea, "N") & ",0,"
            End If

            Sql = Sql & DBSet(Rs!punteada, "N") & "," & DBSet(Rs!nomdocum, "T") & "," & DBSet(Rs!Ampconce, "T") & "," & DBSet(Rs2!subccost, "T") & "," & DBSet(Rs!contra, "T") & "," & DBSet(Rs!CCost, "T") & ",1)"

            Conn.Execute Sql

            ImporteTot = ImporteTot + ImporteLinea

            UltSubCC = Rs2!subccost

            Rs2.MoveNext
        Wend

        If DBLet(Rs!timported, "N") <> 0 Then
            If ImporteTot <> DBLet(Rs!timported, "N") Then
                Sql = "update tmpconext set timported = timported + (" & DBSet(Round(DBLet(Rs!timported, "N") - ImporteTot, 2), "N") & ")"
                Sql = Sql & " where codusu = " & vUsu.Codigo
                Sql = Sql & " and cta = " & DBSet(UltSubCC, "T")
                Sql = Sql & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                Sql = Sql & " and numdiari = " & DBSet(Rs!NumDiari, "N")
                Sql = Sql & " and numasien = " & DBSet(Rs!NumAsien, "N")
                Sql = Sql & " and desdoblado = 1"

                Conn.Execute Sql
            End If
        Else
            If ImporteTot <> DBLet(Rs!timporteH, "N") Then
                Sql = "update tmpconext set timporteh = timporteh + (" & DBSet(Round(DBLet(Rs!timporteH, "N") - ImporteTot, 2), "N") & ")"
                Sql = Sql & " where codusu = " & vUsu.Codigo
                Sql = Sql & " and cta = " & DBSet(UltSubCC, "T")
                Sql = Sql & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                Sql = Sql & " and numdiari = " & DBSet(Rs!NumDiari, "N")
                Sql = Sql & " and numasien = " & DBSet(Rs!NumAsien, "N")
                Sql = Sql & " and desdoblado = 1"

                Conn.Execute Sql
            End If
        End If

        Sql = "delete from tmpconext where codusu = " & vUsu.Codigo
        Sql = Sql & " and cta = " & DBSet(Rs!Cta, "T")
        Sql = Sql & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
        Sql = Sql & " and numdiari = " & DBSet(Rs!NumDiari, "N")
        Sql = Sql & " and numasien = " & DBSet(Rs!NumAsien, "N")
        Sql = Sql & " and desdoblado = 0"

        Conn.Execute Sql

        Set Rs2 = Nothing


        Rs.MoveNext
    Wend

    Set Rs = Nothing

    HacerRepartoSubcentrosCoste = True
'    pb2.Visible = False
    Exit Function
    
eHacerRepartoSubcentrosCoste:
    MuestraError Err.Number, "Reparto Subcentros de Coste", Err.Description
'    pb2.Visible = False
End Function






'--------------------------------------------------------
'  BALANCE DE SUMAS Y SALDOS  ***NUEVO***  31/01/2017
'--------------------------------------------------------
' Viene ya datos "CARGADOS"
' Codmacta: que estamos procesando
' Es decir, vendran en colImportes vendra una serie de string
'
'       201701|123.45|2400.01|
'       201703|12.9|1430.04|
'       201704|1133.5|445.21|
'       201705|60.45|2.62|
Public Sub CargaBalanceNuevaContabilidad(ByRef Cta As String, NomCuenta As String, ConApertura As Boolean, ByRef FechaInicioPeriodo As Date, ByRef FechaFinPeriodo As Date, F_Ini As Date, F_Fin As Date, EjerciCerrados As Boolean, QuitarCierre As Byte, vContabili As Integer, DesdeBalancesConfigurados As Boolean, Resetea6_7 As Boolean, RecordSetPrecargado As Boolean, ByRef ColImportes As Collection)


'FechInicioEsMesInicio ->  QUiere decir que si el mes incio que he puesto coincide
'                          con la fecha incio entoces,  no calcularemos anteriores
'                          y si ademas desglosamos la apertura, se la restaremos a
'                          los moviemientos del periodo, NO al anterior
'
'
'
'  QUitarCierre  :  0.- NO
'                   1.- Ambos
'                   2.- Solo perdidas y ganancias
'                   3.- Cierre
'   RecordSetPrecargado ....
'                   Si precargamos el RS significa que antes de lanzar este proceso cargamos un RS
'                   con los valores de la apertura (Y O CIERRE)
'
' En ejercicios siguientes, las cuentas 6  7 e cogen desde inicio de ejercicio(siguiente)

Dim miSQL As String
Dim ActualD As Currency
Dim ActualH As Currency
Dim NuloAC As Boolean   'Del actual
Dim NuloAC1 As Boolean  'Si huberia del siguiente
Dim NuloPer As Boolean
Dim NuloAper As Boolean
Dim CalcularImporteAnterior As Boolean
    
Dim N As Integer
Dim idPer As Long
Dim Limite1 As Long
Dim Limite2 As Long
Dim AUX3 As String
Dim MenorFecha As Boolean
Dim Be As Boolean

    M1 = Month(FechaInicioPeriodo)
    M2 = Month(FechaFinPeriodo)
    A1 = Year(FechaInicioPeriodo)
    A2 = Year(FechaFinPeriodo)
    vCta = Cta
    vDig = Len(Cta)
    
    'Agosto2014
    FIniPeriodo = FechaInicioPeriodo
    FFinPeriodo = FechaFinPeriodo

    Contabilidad = vContabili
    NombreSQL NomCuenta
    
    NuloAper = True
    miSQL = "INSERT INTO tmpbalancesumas (codusu,"
    miSQL = miSQL & "cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES "
    miSQL = miSQL & " (" & vUsu.Codigo
    miSQL = miSQL & ",'" & vCta & "','" & NomCuenta & "',"
    
        
        
        
    NuloAper = True
    If ConApertura Then
        ImpD = 0:  ImpH = 0
        If RsApertura Is Nothing Then ObtenerAperturaBalance EjerciCerrados, F_Ini, F_Fin, NuloAper
        
        AUX3 = "codmacta = '" & vCta & "'"
        RsApertura.Find AUX3, , adSearchForward, 1
        NuloAper = True
        If Not RsApertura.EOF Then
            NuloAper = False
            ImpD = DBLet(RsApertura!SumaDetimporteD, "N")
            ImpH = DBLet(RsApertura!SumaDetimporteH, "N")

        End If
    Else
        ImpD = 0
        ImpH = 0
    End If
    'Para la cadena de insercion
    'Modificacion 1 Junio 2004. -> Ver A_versiones

    d = TransformaComasPuntos(CStr(ImpD))
    H = TransformaComasPuntos(CStr(ImpH))
    

    miSQL = miSQL & d & "," & H & ","
    
    '----------------------------
    'Calcula Acumulados Anteriores

    
    'Si es el ejercicio siguiente, es decir NO estamos en cerrados
    'Vemos todos los saldos
    ActualD = 0: ActualH = 0
    NuloAC = True
    CalcularImporteAnterior = False
    If Not EjerciCerrados Then
        If vParam.fechafin < FechaInicioPeriodo Then
            Be = Resetea6_7
           
            
            If DesdeBalancesConfigurados Then Be = False
            If Be Then
                If Mid(Cta, 1, 1) = vParam.GrupoGto Or Mid(Cta, 1, 1) = vParam.GrupoVta Then
                    CalcularImporteAnterior = True
                Else
                    If vParam.GrupoOrd <> "" Then
                        If Mid(Cta, 1, 1) = vParam.GrupoOrd Then
                            CalcularImporteAnterior = True
                            If vParam.Automocion <> "" Then
                                If Mid(Cta, 1, Len(vParam.Automocion)) = vParam.Automocion Then CalcularImporteAnterior = False
                            End If
                        End If
                    End If
                End If
                
            Else
                CalcularImporteAnterior = False
            End If
        End If
     End If
    
'
'    'La fecha incio y periodo estan en   FIniPeriodo FFinPeriodo
'    '*************************************************************

    If CalcularImporteAnterior Then
        
        
        
        AUX3 = Format(F_Ini, "yyyymm")
        Limite1 = Val(AUX3)
        AUX3 = DateAdd("yyyy", 1, F_Ini)
        AUX3 = DateAdd("d", -1, CDate(AUX3))
    
        Limite2 = Format(AUX3, "yyyymm")
        
        For N = 1 To ColImportes.Count
            idPer = Val(RecuperaValor(ColImportes(N), 1))
            If idPer >= Limite1 And idPer <= Limite2 Then
               ActualD = ActualD + CCur(RecuperaValor(ColImportes(N), 2))
               ActualH = ActualH + CCur(RecuperaValor(ColImportes(N), 3))
            Else
                Exit For
            End If

        Next
      

    End If
        
    'Las variabled de acumaldo hay k reestablecerlas
    ImAcD = 0: ImAcH = 0
    NuloAC1 = True
   
        
        
    If FIniPeriodo > F_Ini Then
            
        AUX3 = Format(F_Ini, "yyyymm")
        Limite1 = Val(AUX3)
        Limite2 = Format(FIniPeriodo, "yyyymm")
        For N = 1 To ColImportes.Count
            idPer = Val(RecuperaValor(ColImportes(N), 1))
    
            If idPer >= Limite1 And idPer < Limite2 Then
               ImAcD = ImAcD + CCur(RecuperaValor(ColImportes(N), 2))
               ImAcH = ImAcH + CCur(RecuperaValor(ColImportes(N), 3))
            End If
        Next
     
    
        'CalculaAcumuladosAnterioresBalance EjerciCerrados, F_Ini, False, NuloAC1
        ImAcD = ImAcD - ImpD
        ImAcH = ImAcH - ImpH
    End If
    NuloAC = NuloAC1 And NuloAC
    ImAcD = ImAcD - ActualD
    ImAcH = ImAcH - ActualH

    d = TransformaComasPuntos(CStr(ImAcD))
    H = TransformaComasPuntos(CStr(ImAcH))

    
    
    miSQL = miSQL & d & "," & H & ","
    
    
    'Calcula moviemientos periodo
    'MoviemientosPeridoBalance EjerciCerrados, F_Ini, F_Fin, NuloPer
    ' SELECT que hace el sub de arriba
    ' Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH from hlinapu where
    'mid(codmacta,1,9)='100000001' AND fechaent between '2015-05-01' AND '2015-09-30'
    '  FIniPeriodo  FFinPeriodo
    ImPerD = 0: ImPerH = 0
    Limite1 = Format(FIniPeriodo, "yyyymm")
    Limite2 = Format(FFinPeriodo, "yyyymm")
    For N = 1 To ColImportes.Count
        idPer = Val(RecuperaValor(ColImportes(N), 1))
        If idPer >= Limite1 Then
            If idPer <= Limite2 Then
                ImPerD = ImPerD + CCur(RecuperaValor(ColImportes(N), 2))
                ImPerH = ImPerH + CCur(RecuperaValor(ColImportes(N), 3))
            Else
                Exit For
            End If
        End If
    Next
    

    If F_Ini = FechaInicioPeriodo Then
        'Le restamos los movimientos del desglose apertura
        ImPerD = ImPerD - ImpD
        ImPerH = ImPerH - ImpH
    End If
    
    '--------------------------------------------------------
    'Nuevo: 19 de Mayo de 2003
    'Ahora, le restamos, si  asi lo pide, y si se puede, el perdidads y ganacias y cierre
    'Meteremos los valores en imacd imach
    If QuitarCierre > 0 Then
        'Modificacion 24 Noviembre
        
        If RecordSetPrecargado Then
            'Esta es la mod.
            'Tendre un RS ya cargado con los valores, y el lo que antes era un RS.open
            'ahoa sera un RS.find
            BuscarValorEnPrecargado vCta
        Else
        
        
            ObtenerPerdidasyGanancias EjerciCerrados, F_Ini, F_Fin, QuitarCierre  'El 1 significa los dos pyg   y cierre
        End If
            ImPerD = ImPerD - ImCierrD
            ImPerH = ImPerH - ImCierrH
    End If
    
    
    d = TransformaComasPuntos(CStr(ImPerD))
    H = TransformaComasPuntos(CStr(ImPerH))
    miSQL = miSQL & d & "," & H & ","

    If ImpD = 0 And ImAcD = 0 And ImpH = 0 And ImAcH = 0 Then
        If ImPerD = 0 And ImPerH = 0 Then
            NuloPer = NuloPer And NuloAC And NuloAper
            If NuloPer Then Exit Sub
        End If
    End If
    'El saldo sera
    'Apertura k esta en impd
    ' anterior que esta en imacd
    ' periodo que esta en imperd
    ImpD = ImpD + ImAcD + ImPerD
    ImpH = ImpH + ImAcH + ImPerH
    
    
    'Si estamos en balnces configurados entonces no necesito insertar en la BD
    'Lo  Unico k kiero son los valores imd y imph
    If DesdeBalancesConfigurados Then Exit Sub
    
    
    
    'Si vengo para mostarar el balance de sumas y slados entocnes sigo y luego imprimire
    If ImpD >= ImpH Then
        ImpD = ImpD - ImpH
        miSQL = miSQL & TransformaComasPuntos(CStr(ImpD)) & ",NULL)"
    Else
        ImpH = ImpH - ImpD
        miSQL = miSQL & "NULL," & TransformaComasPuntos(CStr(ImpH)) & ")"
    End If
    
    
    Conn.Execute miSQL
End Sub



'AGOSTO 2014
'DAVID
'Desde hlinapu
Private Sub CalculaAcumuladosAnterioresBalance(EjeCerrado As Boolean, ByRef fec1 As Date, EsSiguiente As Boolean, ByRef NulAcum As Boolean)


    Sql = "SELECT Sum(coalesce(timported,0)) AS SumaDetimporteD, Sum(coalesce(timporteh,0)) AS SumaDetimporteH"
    Sql = Sql & " from "
    If Contabilidad >= 0 Then Sql = Sql & " ariconta" & Contabilidad & "."
    Sql = Sql & "hlinapu"
    If EjeCerrado Then Sql = Sql & "1"
    Sql = Sql & " where mid(codmacta,1," & vDig & ")='" & vCta & "'  AND "
    
    
    If Not EsSiguiente Then
        'NORMAL ----------------
        'Desde la fecha de incio correspondiente
        Aux = " fechaent>=" & DBSet(fec1, "F") & " AND fechaent <" & DBSet(FIniPeriodo, "F")
    Else
        'Saldos para ejercicios siguiente
        'Para k acumule el saldo desde fecha inicio actual
        Aux = " fechaent>=" & DBSet(vParam.fechaini, "F") & " AND fechaent <" & DBSet(FIniPeriodo, "F")
    End If
    Sql = Sql & Aux
    Nulo1 = True
    Nulo2 = True
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImAcD = 0
    Else
        ImAcD = RT.Fields(0)
        Nulo1 = False
    End If
    If IsNull(RT.Fields(1)) Then
        ImAcH = 0
    Else
        ImAcH = RT.Fields(1)
        Nulo2 = False
    End If
    NulAcum = Nulo1 And Nulo2
    RT.Close
End Sub

'DAVID
'AGOSTO 2014
'Sobre hlinapu
Private Sub MoviemientosPeridoBalance(Cerrado As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, ByRef NuloPerio As Boolean)
Dim DiaFin As Integer



    

    Sql = "SELECT Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    'Modificacion para las cuentas k tienen movimientos positivos y negativos y
    Sql = Sql & " from "
    If Contabilidad >= 0 Then Sql = Sql & " ariconta" & Contabilidad & "."
    Sql = Sql & "hlinapu"
    If Cerrado Then Sql = Sql & "1"
    

    Sql = Sql & " where  mid(codmacta,1," & vDig & ")='" & vCta & "' AND fechaent between " & DBSet(FIniPeriodo, "F") & " AND " & DBSet(FFinPeriodo, "F")
    
    Nulo1 = False
    Nulo2 = False
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImPerD = 0
        Nulo1 = True
    Else
        ImPerD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImPerH = 0
        Nulo2 = True
    Else
        ImPerH = RT.Fields(1)
    End If
    NuloPerio = Nulo1 And Nulo2
    RT.Close
End Sub








Private Sub ObtenerAperturaBalance(EjerCerrados As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, ByRef NulAper As Boolean)
Dim Aux As String

    'El movimietno de apertura se clacula mirando el asiento de apertura (codigo
    'concepto 970)
    If EsCuentaUltimoNivel(vCta) Then
        Aux = "codmacta"
    Else
        Aux = " substring(codmacta,1," & Len(vCta) & ") codmacta"
    End If
    Sql = "SELECT " & Aux & ",Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH  "
    

    
    Sql = Sql & " from "
    If Contabilidad >= 0 Then Sql = Sql & " ariconta" & Contabilidad & "."
    Sql = Sql & "hlinapu"
    If EjerCerrados Then Sql = Sql & "1"
    Sql = Sql & " WHERE fechaent >= '"
    
    If fec1 >= vParam.fechaini Then
        Sql = Sql & Format(vParam.fechaini, FormatoFecha)
    Else
        Sql = Sql & Format(fec1, FormatoFecha)
    End If
    Sql = Sql & "' and fechaent <='" & Format(fec2, FormatoFecha) & "'"
    Sql = Sql & " AND codconce= 970" '970 es el asiento de apertura
    Sql = Sql & " group by 1"
    Sql = Sql & " ORDER BY codmacta"
    Set RsApertura = New ADODB.Recordset
    RsApertura.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
  
   
End Sub




Private Sub ObtenerApertura(EjerCerrados As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, ByRef NulAper As Boolean)
Dim Aux As String

    'El movimietno de apertura se clacula mirando el asiento de apertura (codigo
    'concepto 970)
    Sql = "SELECT Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    If EsCuentaUltimoNivel(vCta) Then
        Aux = vCta
    Else
        Aux = vCta & "%"
    End If
    
    Sql = Sql & " from "
    If Contabilidad >= 0 Then Sql = Sql & " ariconta" & Contabilidad & "."
    Sql = Sql & "hlinapu"
    If EjerCerrados Then Sql = Sql & "1"
    Sql = Sql & " where codmacta like '" & Aux & "'"
    Sql = Sql & " and fechaent >='" & Format(fec1, FormatoFecha) & "'"
    Sql = Sql & " and fechaent <='" & Format(fec2, FormatoFecha) & "'"
    Sql = Sql & " AND codconce= 970" '970 es el asiento de apertura
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Nulo1 = True
    Nulo2 = True
    If IsNull(RT.Fields(0)) Then
        ImpD = 0
    Else
        ImpD = RT.Fields(0)
        Nulo1 = False
    End If
    If IsNull(RT.Fields(1)) Then
        ImpH = 0
    Else
        ImpH = RT.Fields(1)
        Nulo2 = False
    End If
    NulAper = Nulo1 And Nulo2
    RT.Close
End Sub




Public Function AgrupacionCtasBalance(Codigo As String, Nommacta As String) As Boolean
Dim C As Integer
On Error GoTo EAgrupacionCtasBalance

    AgrupacionCtasBalance = False
    ImAcD = 0
    ImAcH = 0
    ImPerD = 0
    ImPerH = 0
    ImCierrD = 0
    ImCierrH = 0
    ImpD = 0
    ImpH = 0
    vCta = Mid(Codigo & "__________", 1, vEmpresa.DigitosUltimoNivel)
    
    Sql = "Select * from tmpbalancesumas where codusu =" & vUsu.Codigo
    Sql = Sql & " AND cta like '" & vCta & "'"
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = 0
    While Not RT.EOF
        'Apertura
        ImAcD = ImAcD + DBLet(RT.Fields(3), "N")
        ImAcH = ImAcH + DBLet(RT.Fields(4), "N")
        'anterior
        ImPerD = ImPerD + DBLet(RT.Fields(5), "N")
        ImPerH = ImPerH + DBLet(RT.Fields(6), "N")
        'periodo
        ImCierrD = ImCierrD + DBLet(RT.Fields(7), "N")
        ImCierrH = ImCierrH + DBLet(RT.Fields(8), "N")
        'Total
        ImpD = ImpD + DBLet(RT.Fields(9), "N")
        ImpH = ImpH + DBLet(RT.Fields(10), "N")
        
        RT.MoveNext
        C = C + 1
    Wend
    RT.Close
    If C = 0 Then
        AgrupacionCtasBalance = True
        Exit Function
    End If
    
    'Acumulamos saldo en uno de los lados
    If ImpD > ImpH Then
        ImpD = ImpD - ImpH
        ImpH = 0
    Else
        ImpH = ImpH - ImpD
        ImpD = 0
    End If
    
    
    'Borramos las entradas
    Sql = "DELETE from tmpbalancesumas where codusu =" & vUsu.Codigo
    Sql = Sql & " AND cta like '" & vCta & "'"
    Conn.Execute Sql
    Conn.Execute "commit"
    espera 0.5
    
    Sql = "INSERT INTO tmpbalancesumas (codusu, cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) VALUES (" & vUsu.Codigo
    Aux = Mid(Codigo & "**********", 1, vEmpresa.DigitosUltimoNivel)
    Sql = Sql & ",'" & Aux & "','" & Mid("AGRUP- " & Nommacta, 1, 30) & "'"
    Sql = Sql & ImporteASQL(ImAcD) & ImporteASQL(ImAcH) & ImporteASQL(ImPerD) & ImporteASQL(ImPerH)
    Sql = Sql & ImporteASQL(ImCierrD) & ImporteASQL(ImCierrH) & ImporteASQL(ImpD) & ImporteASQL(ImpH) & ")"
    Conn.Execute Sql
    AgrupacionCtasBalance = True
    Exit Function
EAgrupacionCtasBalance:
    MuestraError Err.Number, "Agrupacion Ctas Balance"
End Function










' MAYO 2004.  Vamos a poder separar dependiendo del tipo de llamada
'         0.- No llegara hasta aqui
'         1.- Los dos  pérdidas/ganancias   y   cierre
'         2.- Perdidas y GAnancias
'         3.- Solo Cierre
Private Sub ObtenerPerdidasyGanancias(EjerCerrados As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, OpcionBusqueda As Byte)
Dim Aux As String

    'Perdidas y ganancias: 960
    'Cierre             : 980
    
    Sql = "SELECT Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    If EsCuentaUltimoNivel(vCta) Then
        Aux = vCta
    Else
        Aux = vCta & "%"
    End If
    Sql = Sql & " from "
    If Contabilidad >= 0 Then Sql = Sql & " ariconta" & Contabilidad & "."
    Sql = Sql & "hlinapu"
    If EjerCerrados Then Sql = Sql & "1"
    Sql = Sql & " where codmacta like '" & Aux & "'"
    Sql = Sql & " and fechaent >='" & Format(fec1, FormatoFecha) & "'"
    Sql = Sql & " and fechaent <='" & Format(fec2, FormatoFecha) & "'"
    
    '  960 P y G
    '  970 es el asiento de apertura
    '  980 Cierre
    Aux = ""
    If OpcionBusqueda < 3 Then Aux = "codconce= 960"
    If OpcionBusqueda <> 2 Then
        If Aux <> "" Then Aux = Aux & " OR "
        Aux = Aux & "codconce= 980"
    End If
    Aux = " AND (" & Aux & ")"
    Sql = Sql & Aux
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImCierrD = 0
    Else
        ImCierrD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImCierrH = 0
    Else
        ImCierrH = RT.Fields(1)
    End If
    RT.Close
End Sub


'---------------------------------------------------------
'Precarga de los datos del balance
'
Public Sub PrecargaPerdidasyGanancias(EjerCerrados As Boolean, ByRef fec1 As Date, ByRef fec2 As Date, OpcionBusqueda As Byte)
Dim Aux As String

    'Perdidas y ganancias: 960
    'Cierre             : 980
    
    Sql = "SELECT codmacta,Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    Sql = Sql & " from "
    Sql = Sql & "hlinapu"
    If EjerCerrados Then Sql = Sql & "1"
    Sql = Sql & " where fechaent >='" & Format(fec1, FormatoFecha) & "'"
    Sql = Sql & " and fechaent <='" & Format(fec2, FormatoFecha) & "'"
    
    '  960 P y G
    '  970 es el asiento de apertura
    '  980 Cierre
    Aux = ""
    If OpcionBusqueda < 3 Then Aux = "codconce= 960"
    If OpcionBusqueda <> 2 Then
        If Aux <> "" Then Aux = Aux & " OR "
        Aux = Aux & "codconce= 980"
    End If
    Aux = " AND (" & Aux & ")"
    Sql = Sql & Aux & " GROUP BY codmacta"
        
        
    Set RsBalPerGan = New ADODB.Recordset
    RsBalPerGan.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
End Sub


Public Sub PrecargaApertura()

    Sql = "SELECT codmacta,Sum(timported) AS SumaDetimporteD, Sum(timporteh) AS SumaDetimporteH"
    Sql = Sql & " from hlinapu"
    Sql = Sql & " where fechaent ='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Sql = Sql & " AND codconce= 970 GROUP BY codmacta"
        
        
    Set RsBalPerGan = New ADODB.Recordset
    RsBalPerGan.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
End Sub



Public Sub CerrarPrecargaPerdidasyGanancias()
    RsBalPerGan.Close
    Set RsBalPerGan = Nothing
End Sub

Public Sub CerrarLeerApertura()
    On Error Resume Next
    RsApertura.Close
    Set RsApertura = Nothing
    Err.Clear
End Sub


Public Sub BuscarValorEnPrecargado(ByRef codmacta As String)


    RsBalPerGan.Find "codmacta = '" & codmacta & "'", , adSearchForward, 1
    If RsBalPerGan.EOF Or RsBalPerGan.BOF Then
        ImCierrD = 0
        ImCierrH = 0
    Else
        If IsNull(RsBalPerGan.Fields(1)) Then
            ImCierrD = 0
        Else
            ImCierrD = RsBalPerGan.Fields(1)
        End If
        If IsNull(RsBalPerGan.Fields(2)) Then
            ImCierrH = 0
        Else
            ImCierrH = RsBalPerGan.Fields(2)
        End If
    End If
End Sub






'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'
' Cuentas de explotacion
'
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'--->  OPCION  0.- Con anterior y movimientos     1.- Solo SALDO
'El ctaSQL es para no tener que copiar el SQL de insertar


'''''''''''
''''''''''''Calculamos los importes de los cierres para obtener la consulta sin ellos
'''''''''''Private Sub CalcularImporteCierreCtaExplotacion(Cerrados As Boolean, ByRef fFin As Date)
'''''''''''
'''''''''''    SQL = "Select SUM(timporteD),sum(timporteH) from "
'''''''''''    If Contabilidad > 0 Then SQL = SQL & "conta" & Contabilidad & "."
'''''''''''    SQL = SQL & "hlinapu"
'''''''''''    If Cerrados Then SQL = SQL & "1"
'''''''''''    SQL = SQL & " WHERE codmacta  like '" & vCta
'''''''''''    If Len(vCta) <> vEmpresa.DigitosUltimoNivel Then SQL = SQL & "%"
'''''''''''    SQL = SQL & "' and codconce ="
'''''''''''    d = Mid(vCta, 1, 1)
'''''''''''    If d = vParam.grupogto Or d = vParam.grupovta Or d = vParam.grupoord Then
'''''''''''        SQL = SQL & "960" 'perdidas y ganacias
'''''''''''    Else
'''''''''''        SQL = SQL & "980" ' cierre
'''''''''''    End If
'''''''''''    SQL = SQL & " AND fechaent = '" & Format(fFin, FormatoFecha) & "'"
'''''''''''    SQL = SQL & ";"
'''''''''''
'''''''''''
'''''''''''    RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'''''''''''    If IsNull(RT.Fields(0)) Then
'''''''''''        ImCierrD = 0
'''''''''''    Else
'''''''''''        ImCierrD = RT.Fields(0)
'''''''''''    End If
'''''''''''    If IsNull(RT.Fields(1)) Then
'''''''''''        ImCierrH = 0
'''''''''''    Else
'''''''''''        ImCierrH = RT.Fields(1)
'''''''''''    End If
'''''''''''
'''''''''''    RT.Close
'''''''''''End Sub







Private Function DevuelveIVANODEDUCIBLE(ByRef RD As ADODB.Recordset, Tipo As String) As String
    RD.Find "codigiva = " & Tipo, , adSearchForward, 1
    DevuelveIVANODEDUCIBLE = ",''"
    If Not RD.EOF Then
        If RD!TipoDIva = 4 Then DevuelveIVANODEDUCIBLE = ",'ND'"
    End If
    
End Function







'------------------------------------------------------------------------

Public Function ImporteBalancePresupuestario(ByRef vSql As String) As Currency

ImPerH = 0
Set RT = New ADODB.Recordset
RT.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RT.EOF Then
    If Not IsNull(RT.Fields(0)) Then
        ImpD = RT.Fields(0)
    Else
        ImpD = 0
    End If
    If Not IsNull(RT.Fields(1)) Then
        ImpH = RT.Fields(1)
    Else
        ImpH = 0
    End If
    ImPerH = ImpD - ImpH
End If
RT.Close
ImporteBalancePresupuestario = ImPerH
End Function


'Desde donde:
'       0.- Listado simulacin
'       1.- Venta / baja de elmento
Public Function HazSimulacion(ByRef vSql As String, Fecha As Date, DesdeDonde As Byte, ByRef Lb As Label) As Boolean
Dim FechaCalculoVentaBaja As Date
Dim I2 As Integer
On Error GoTo EHazSimulacion
    
    HazSimulacion = False
    'Obtenemos la ultmia fecha de amortizacion
    Set RT = New ADODB.Recordset
    RT.Open "Select * from paramamort where codigo = 1", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        MsgBox "Error leyendo parámetros.", vbExclamation
        RT.Close
        Exit Function
    End If
    
    'Ademas, en M2 pondremos el tipo de amortizacion, el valor por el k habra que dividir + adelante
    Select Case Val(RT!tipoamor)
    Case 2
        'Semestral
        M2 = 2
        I2 = 6 'dato auxiliar
    Case 3
        'Trimestral
        M2 = 4
        I2 = 3   'ENERO2018   Ponia un 4!!!!!!
    Case 4
        'mensual
        M2 = 12
        I2 = 1
    Case Else
        'Anual
        M2 = 1
        I2 = 12
    End Select
    
    
    If RT.Fields(3) > CDate(Fecha) Then
        MsgBox "Fecha última amortizacion mayor que fecha operacion.", vbExclamation
        RT.Close
        Exit Function
    End If
    'En m1 almacenamos los dias del diferencia
    vFecha1 = CDate(Fecha)  'La nueva fechamo
    vFecha2 = RT.Fields(3)  'Ultmfechaamort
    
    
    
    If DesdeDonde = 1 Then
        FechaCalculoVentaBaja = DateAdd("m", I2, vFecha2)
    '++
    Else
        M1 = DiasMes(Month(vFecha2), Year(vFecha2))
        If M1 = Day(vFecha2) Then
            'Ultmo dia mes
            M1 = 1
        Else
            M1 = 0
        End If
        vFecha1 = DateAdd("m", I2, vFecha2)
        
        If M1 = 1 Then
            'ULTIMO DIA MES
            M1 = DiasMes(Month(vFecha1), Year(vFecha1))
            vFecha1 = CDate(M1 & Format(vFecha1, "/mm/yyyy"))
            
        End If
    End If
    
    
    RT.Close
    If Not Lb Is Nothing Then
            Lb.Caption = "Leyendo registros"
            Lb.Refresh
    End If
    
    Sql = "Delete from tmpsimulainmo where codusu=" & vUsu.Codigo
    Conn.Execute Sql
    
    
    
    'Obtenemos el recordset
    Sql = "select codinmov,fechaadq,valoradq,anovidas,amortacu,inmovele.conconam,"
    Sql = Sql & " tipoamor,coeficie,valorres,fecventa,coefimaxi,nominmov,nomconam from inmovcon,inmovele where inmovele.conconam=inmovcon.codconam"
    Sql = Sql & " AND fecventa is null AND impventa is null AND situacio<>4"
    'Junio 2005
    '-------------
    ' Indicamos que la fecha adq sea menor que la fecha simulacion
    Sql = Sql & " AND fechaadq<='" & Format(vFecha1, FormatoFecha) & "'"
    If vSql <> "" Then Sql = Sql & " AND " & vSql
    
    RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        If DesdeDonde = 0 Then
            MsgBox "Ningún  elemento de inmovilizado con estos valores.", vbExclamation
        Else
            'Estamos vendiendo un elto, o de baja
            'Signifaca que NO hay que amortizar
            HazSimulacion = True
        End If
        RT.Close
        Exit Function
    End If

    
    Sql = "INSERT INTO tmpsimulainmo (codusu, codigo, conconam, nomconam, codinmov,"
    Sql = Sql & "nominmov, fechaadq, valoradq, amortacu, totalamor) VALUES (" & vUsu.Codigo & ","
    M1 = 1
    
    'Dias totales
    If DesdeDonde = 1 Then
        A1 = DateDiff("d", vFecha2, FechaCalculoVentaBaja)
    Else
        A1 = DateDiff("d", vFecha2, vFecha1)
    End If
    If A1 <= 0 Then
        MsgBox "Diferencia entre fechas amortización es <=0", vbExclamation
        Exit Function
    End If
    
    While Not RT.EOF
            
        If Not Lb Is Nothing Then
            Lb.Caption = RT!Codinmov & " " & DBLet(RT!nominmov, "T")
            Lb.Refresh
        End If
        ObtenAmortizacionAnualSimulacion 'En IMPERD esta almacenada
        'Vemos los dias del period a aplicar el valor
        CalcularDiasAplicablesSimulacion DesdeDonde = 1, Fecha
    
        If ImPerD > 0 Then
                    
            'Vemos, en funcion de los dias
           
            ImCierrD = Round(ImPerD * (A2 / A1), 2)
            'ImPerD = Round(ImPerD, 2)
            'ImCierrD = ImPerD * ImCierrD
        
            'Calcualmos los valores
            'ImPerH = Round(ImPerD * A2 / A1, 2)
            ImPerH = ImCierrD
            
            'Ahora, si lo k ahy k amortizar es mayor de lo que queda, entonces amortizamos
            'solo lo k queda
            ImPerD = Round(RT!valoradq - RT!amortacu, 2)
            If ImPerH > ImPerD Then ImPerH = ImPerD

            d = TransformaComasPuntos(CStr(ImPerH))
            'Insertamos
            Aux = Sql & M1 & "," & RT!conconam & ",'" & RT!nomconam & "'," & RT!Codinmov
            Aux = Aux & ",'" & DevNombreSQL(RT!nominmov) & "','" & Format(RT!fechaadq, "dd/mm/yyyy") & "',"
            H = TransformaComasPuntos(CStr(RT!valoradq))
            Aux = Aux & H & ","
            H = TransformaComasPuntos(CStr(RT!amortacu))
            Aux = Aux & H & "," & d & ")"
            Conn.Execute Aux
        End If
        'Siguiente elemento
        RT.MoveNext
        M1 = M1 + 1
    Wend
    HazSimulacion = True
    Exit Function
EHazSimulacion:
    MuestraError Err.Number, Err.Description
    Set RT = Nothing
End Function

'////////////////////////////////////
' A partir del RT , recordset k tiene los datos,
' pondremos en ImAcD el valor de la amortizacion anual
' en el segundo paso pondremos la apliacble( en funcion de los dias transcurridos
Private Sub ObtenAmortizacionAnual()
    Select Case RT!tipoamor
    Case 2
        'Lineal
        ImPerD = (RT!valoradq - RT!valorres) / RT!anovidas
    Case 3
        'Degresiva
        ImPerD = (RT!valoradq - RT!amortacu) * (RT!coefimaxi / 100)
    Case 4
        'Porcentual
        ImPerD = (RT!valoradq * RT!coeficie) / 100
    Case Else
        'Tablas
        ImPerD = RT!valoradq / RT!anovidas
    End Select
    ImPerD = Round(ImPerD, 2)  'Redondeando
    'Aplicamos al period mensual, trimestr...
    ImPerD = Round(ImPerD / M2, 2)
End Sub



'///////////////////////////////////////
'
' Esto es, si son 60 dias pero solo hay que aplicar 20 entonces
Private Sub CalcularDiasAplicables()
Dim DiasCalculados As Boolean
    
    
    DiasCalculados = False
    If RT!fechaadq > vFecha2 Then
        If RT!fechaadq > vFecha1 Then
            'Ha comprado incluso despues de
            'la fecha de amortizacion
            A2 = 0
        Else
            'Modificado 15 Octubre 2008 al igual que calculardiasaplicablessimulacion
            A2 = DateDiff("d", RT!fechaadq, vFecha1) + 1
            If A2 > A1 Then A2 = A1
        End If
        DiasCalculados = True
    Else
        A2 = A1
    End If
    If Not IsNull(RT!fecventa) Then
    
        'Una vez comprobadom que vaya todo. Hay que "vovler a hacer " esta funcion.
    
        Err.Raise 513, , "Fecha venta no es nula. llame a soporte tecnico. ", vbCritical
        If RT!fechaadq > vFecha2 Then
             A2 = DateDiff("d", vFecha2, RT!fecventa)
        Else
            If RT!fecventa > vFecha2 Then
                'Se ha vendido despues del la ultima amortizacion
                A2 = DateDiff("d", vFecha2, RT!fecventa)
            End If
        End If
        If RT!fecventa < vFecha2 Then
            'Se ha vendido despues del la ultima amortizacion
            A2 = 0
        End If
    
    Else
        If Not DiasCalculados Then A2 = DateDiff("d", vFecha2, vFecha1)
        
        If A2 > A1 Then
            MsgBox "Error calculando datos. Dias mayor que el maximo del periodo. (DiasAplicables)", vbExclamation
            A2 = A1
        End If
        
    End If
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'      SIMULACION
'----------------------------------------------------------------
'----------------------------------------------------------------
'nuevo a 29 Marzo 2005
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub ObtenAmortizacionAnualSimulacion()
    Select Case RT!tipoamor
    Case 2
        'Lineal
        ImPerD = (RT!valoradq - RT!valorres) / RT!anovidas
    Case 3
        'Degresiva
        ImPerD = (RT!valoradq - RT!amortacu) * (RT!coefimaxi / 100)
    Case 4
        'Porcentual
        ImPerD = (RT!valoradq * RT!coeficie) / 100
    Case Else
        'Tablas
        ImPerD = RT!valoradq / RT!anovidas
    End Select
    
    'En imperd tengo la simulacion ANUAL
    
    ImPerD = Round(ImPerD, 2)  'Redondeando
    'Aplicamos al period mensual, trimestr...
    ImPerD = Round(ImPerD / M2, 2)
End Sub



Private Sub CalcularDiasAplicablesSimulacion(EsEnVentabaja As Boolean, FechaVtaBaja As Date)

    A2 = -1
    If RT!fechaadq > vFecha2 Then
        If RT!fechaadq > vFecha1 Then
            'Ha comprado incluso despues de
            'la fecha de amortizacion
            A2 = 0
        Else
            'Se ha comprado despues de la ultima amortizacion
            'Nuevo 15 Octubre 2008
            'A2 = DateDiff("d", RT!fechaadq, vFecha1)
            A2 = DateDiff("d", RT!fechaadq, vFecha1) + 1  'Ya que el primer dia tb se utiliza
            If A2 > A1 Then A2 = A1 'Esto no deberia pasar NUNCA, pero mas vale prevenir que curar
        End If
    Else
        A2 = A1 + DateDiff("d", vFecha1, FechaVtaBaja)
    End If
    
    If EsEnVentabaja Then
        'Aqui veremos cuantos dias hay que aplicar la amortizacion
        If A2 < 0 Then A2 = DateDiff("d", vFecha2, FechaVtaBaja)
    Else
        If A2 < 0 Then
            If Not IsNull(RT!fecventa) Then
                If RT!fecventa >= vFecha2 Then
                    'Se ha vendido despues del la ultima amortizacion
                    A2 = DateDiff("d", vFecha2, RT!fecventa)
                Else
                    A2 = 0
                End If
            End If
        End If
    End If
    
    If A2 = -1 Then Err.Raise 513, , "Calculando dias aplicables"
    
End Sub



'///////////////////////////////////////////////////////////////////////////
'
'   CALCULO AMORTIZACION
'
Public Function CalculaAmortizacion(Codinmov As Long, Fecha As Date, DivMes As Integer, UltimaAmort As Date, ParametrosContabiliza As String, mContador As Long, ByRef NumLinea As Integer, EsVentaBaja As Boolean) As Boolean
Dim Rs As Recordset
Dim NomConce As String

On Error GoTo ECalculaAmortizacion
    CalculaAmortizacion = False

    
    
    Set RT = New ADODB.Recordset
    Aux = "select inmovele.*,inmovcon.coefimaxi from inmovcon,inmovele where inmovele.conconam=inmovcon.codconam"
    Aux = Aux & " AND codinmov = " & Codinmov
    RT.Open Aux, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    vFecha1 = Fecha
    vFecha2 = UltimaAmort
    M2 = DivMes
  
    ObtenAmortizacionAnual 'En IMPERD esta almacenada
    'Vemos los dias del period a aplicar el valor
    If EsVentaBaja Then
        'En i metermos los meses a  sumar a la fecha
        If DivMes = 1 Then
            'Amortizacion anual
            A1 = 12 'le suma
        ElseIf DivMes = 12 Then
            'MENSUAL
            A1 = 1
        ElseIf DivMes = 4 Then
            'TRIMESTRAL
            A1 = 3
        Else
            'Semestral
            A1 = 6
        End If
        'EL sumamos los I meses a la ultima fecha de amortizacion
        A1 = DateDiff("d", vFecha2, DateAdd("m", A1, vFecha2))
    Else
        A1 = DateDiff("d", vFecha2, vFecha1)
    End If
    CalcularDiasAplicables
    
    
    
    'Ya tenesmo en A1 los dias totales y en a2 los aplicables
    If A1 = 0 Then ImPerD = 0
    ImPerH = 0
    If ImPerD > 0 Then
        'Calcualmos los valores
        ImPerH = Round(ImPerD * (A2 / A1), 2)
            
        'Ahora, si lo k ahy k amortizar es mayor de lo que queda, entonces amortizamos
        'solo lo k queda
        ImPerD = Round(RT!valoradq - RT!amortacu, 2)

        If ImPerH > ImPerD Then ImPerH = ImPerD
            
        
        'Calculo el % de amortizacion
        ImpD = Round((ImPerH / RT!valoradq) * 100, 2)
        
        
        'Metemos en hco inmovilizado
        '--------------------------
        Sql = "INSERT INTO inmovele_his (codinmov, fechainm, imporinm, porcinm) VALUES ("
        Sql = Sql & Codinmov & ",'" & Format(vFecha1, FormatoFecha) & "',"
        H = TransformaComasPuntos(CStr(ImPerH))
        d = TransformaComasPuntos(CStr(ImpD))
        Sql = Sql & H & ","
        Sql = Sql & d & ")"
        Conn.Execute Sql

        'ParametrosContabiliza :=>  contabiliza|debe|haber|diario
        If RecuperaValor(ParametrosContabiliza, 1) = "1" Then
            'Contabilizamos insertando en diario de apuntes
                'Insertamos las lineas
                'Este trozo es comun para las del debe y las del haber
                Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce,"
                Sql = Sql & "timporteD, timporteH, codccost, ctacontr, idcontab, punteada) VALUES ("
                Sql = Sql & RecuperaValor(ParametrosContabiliza, 4) & ",'"
                Sql = Sql & Format(vFecha1, FormatoFecha)
                Sql = Sql & "'," & mContador & ","
                
                NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & RecuperaValor(ParametrosContabiliza, 2))
                'amortizacion acumulada -->Haber
                Aux = NumLinea & ",'" & RT!codmact3 & "','" & Format(Codinmov, "000000") & "',"
                Aux = Aux & RecuperaValor(ParametrosContabiliza, 2)   'Concepto DEBE
                '[Monica]15/09/2015: añadido el nombre de concepto que no estaba en la ampliacion
                Aux = Aux & ",'" & DevNombreSQL(NomConce) & " " & DevNombreSQL(RT!nominmov)
                Aux = Aux & "',NULL," & H    'H tiene el importe del inmovilizado
                'El Centro de coste es 0
                Aux = Aux & ",NULL"
                If RT!repartos = 0 Then
                    vCta = "'" & RT!codmact2 & "'"
                Else
                    vCta = "NULL"
                End If
                Aux = Aux & "," & vCta & ",'CONTAI',0)"
                Conn.Execute Sql & Aux
                NumLinea = NumLinea + 1
             
                'Cta gastos --> Debe
                If RT!repartos = 0 Then
                    
                    Aux = NumLinea & ",'" & RT!codmact2 & "','" & Format(Codinmov, "000000") & "',"
                    Aux = Aux & RecuperaValor(ParametrosContabiliza, 3)   'Concepto HABER
                    
                    '[Monica]15/09/2015: añadido el nombre de concepto que no estaba en la ampliacion
                    NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & RecuperaValor(ParametrosContabiliza, 3))
                    Aux = Aux & ",'" & DevNombreSQL(NomConce) & " " & DevNombreSQL(RT!nominmov) & "'," & H & ",NULL"       'H tiene el importe del inmovilizado
                    If vParam.autocoste Then
                        If IsNull(RT!CodCCost) Then
                            Aux = Aux & ",NULL"
                        Else
                            Aux = Aux & ",'" & RT!CodCCost & "'"
                        End If
                    Else
                        'No lleva centro de coste
                        Aux = Aux & ",NULL"
                    End If
                    Aux = Aux & ",'" & RT!codmact3 & "','CONTAI',0)"
                    Conn.Execute Sql & Aux
                    
                Else
                    'Si k tiene reparto
                    Set Rs = New ADODB.Recordset
                    Rs.Open "Select * from inmovele_rep where codinmov =" & Codinmov, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    ImAcD = 0 'Tendre el sumatorio de los repartos
                    While Not Rs.EOF
                        'Calculamos el importe por centual y lo metemos en IMACH
                        ImAcH = Round(((ImPerH * Rs!porcenta) / 100), 2)
                        If vParam.autocoste Then
                            If IsNull(Rs!CodCCost) Then
                                vCta = "NULL"
                            Else
                                vCta = "'" & Rs!CodCCost & "'"
                            End If
                        Else
                            vCta = "NULL"
                        End If
                        Aux = NumLinea & ",'" & Rs!Codmacta2 & "','" & Format(Codinmov, "000000") & "',"
                        
                        Aux = Aux & RecuperaValor(ParametrosContabiliza, 3)   'Concepto HABER
                        '[Monica]15/09/2015: añadido el nombre de concepto que no estaba en la ampliacion
                        NomConce = DevuelveValor("select nomconce from conceptos where codconce = " & RecuperaValor(ParametrosContabiliza, 3))
                        Aux = Aux & ",'" & DevNombreSQL(NomConce) & " " & DevNombreSQL(RT!nominmov)
                        
                        'Avanzamos al siguiente
                        Rs.MoveNext
                        If Rs.EOF Then
                            'Es la ultima linea. Compruebo k el sumatorio de las lineas sea le total
                            ImAcH = ImPerH - ImAcD
                        Else
                            ImAcD = ImAcD + ImAcH
                            NumLinea = NumLinea + 1
                        End If
                        H = TransformaComasPuntos(CStr(ImAcH))
                        'Aux = Aux & "',NULL," & H & "," 'H tiene el importe del inmovilizado
                        Aux = Aux & "'," & H & ",NULL," 'H tiene el importe del inmovilizado
                        Aux = Aux & vCta   'ccoste
                        Aux = Aux & ",'" & RT!codmact3 & "','CONTAI',0)"
                        Conn.Execute Sql & Aux
                    Wend
                    Rs.Close
                    Set Rs = Nothing
                End If
        End If
    End If
    
    
    
    'ACtualizamos eltos. inmovilizado
    'En imperh tengo lo k voy a amortizar
    'En imperd tengo la nueva amortizacon acumulada
    ImPerD = RT!amortacu + ImPerH
    H = TransformaComasPuntos(CStr(ImPerD))
    Sql = "UPDATE inmovele set amortacu=" & H
    If ImPerD = RT!valoradq Then
        'Totalmente amortizado
        Sql = Sql & ", situacio= 4"
    End If
    Sql = Sql & " WHERE codinmov=" & Codinmov
    Conn.Execute Sql
    
    CalculaAmortizacion = True
    Exit Function
ECalculaAmortizacion:
    MuestraError Err.Number, "Calcula Amortizacion" & vbCrLf & Err.Description
    Set RT = Nothing
End Function


Public Function ObtenerparametrosAmortizacion(ByRef DivMes As Integer, ByRef UltmAmort As Date, ByRef RestParametros As String) As Boolean

    Set RT = New ADODB.Recordset
    RT.Open "Select * from paramamort where codigo =1", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        MsgBox "Error leyendo parámetros.", vbExclamation
        RT.Close
        ObtenerparametrosAmortizacion = False
        Exit Function
    End If
    
    'Ademas, en M2 pondremos el tipo de amortizacion, el valor por el k habra que dividir + adelante
    M1 = RT!tipoamor
    Select Case M1
    Case 2
        'Semestral
        M2 = 2
    Case 3
        'Trimestral
        M2 = 4
    Case 4
        'mensual
        M2 = 12
    Case Else
        'Anual
        M2 = 1
    End Select
    DivMes = M2  'En cuantos trozos dividimos el año
    UltmAmort = RT.Fields(3)  'Ultmfechaamort
    RestParametros = RT!intcont & "|"
    If (RT!intcont = 1) Then RestParametros = RestParametros & RT!condebes & "|" & RT!conhaber & "|" & RT!NumDiari & "|"
    RestParametros = RestParametros & RT!Preimpreso & "|"
    RT.Close
    ObtenerparametrosAmortizacion = True
    Set RT = Nothing
End Function






Private Sub CargarIvasATratar(IvaClientes As Boolean)
    Sql = "Delete from tmpliqiva where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    On Error Resume Next  'Por k si da fallo es k ya estaba introducido
    
    If IvaClientes Then
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    End If
    
    Set RT = New ADODB.Recordset
    
    If IvaClientes Then
        Sql = "select porciva,codigiva from " & vCta & ".factcli_totales INNER JOIN " & vCta & ".factcli On factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu WHERE " & Aux & " group by porciva"
    Else
        Sql = "select porciva,codigiva from " & vCta & ".factpro_totales INNER JOIN " & vCta & ".factpro On factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu WHERE " & Aux & " group by porciva"
    End If

    RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RT.EOF
        If Not IsNull(RT.Fields(0)) Then
            d = TransformaComasPuntos(CStr(RT.Fields(0)))
            Sql = "INSERT INTO tmpliqiva (codusu, iva) VALUES (" & vUsu.Codigo & "," & d & ")"
            Conn.Execute Sql
            If Err.Number <> 0 Then Err.Clear
        End If
        RT.MoveNext
    Wend
    RT.Close


    Set RT = Nothing
    On Error GoTo 0
End Sub




Private Function CargarRecargosATratar(IvaClientes As Boolean) As Boolean
    CargarRecargosATratar = False
    
    
    Sql = "Delete from tmpliqiva where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    On Error Resume Next  'Por k si da fallo es k ya estaba introducido
    
    If IvaClientes Then
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    End If
    
    Set RT = New ADODB.Recordset
'    For M1 = 1 To 3
        If IvaClientes Then
            Sql = "select porcrec, imporec from " & vCta & ".factcli_totales INNER JOIN factcli ON factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
            Sql = Sql & " WHERE " & Aux & " group by porcrec"
        Else
            Sql = "select porcrec, imporec from " & vCta & ".factpro_totales INNER JOIN factpro ON factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
            Sql = Sql & " WHERE " & Aux & " group by porcrec"
        End If
        
        RT.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RT.EOF
            If Not IsNull(RT.Fields(0)) Then
                d = TransformaComasPuntos(CStr(RT.Fields(0)))
                Sql = "INSERT INTO tmpliqiva (codusu, iva) VALUES (" & vUsu.Codigo & "," & d & ")"
                Conn.Execute Sql
                If Err.Number <> 0 Then Err.Clear
                CargarRecargosATratar = True
            End If
            RT.MoveNext
        Wend
        RT.Close
'    Next M1
    Set RT = Nothing
    On Error GoTo 0
End Function





Private Sub TotalIva(Porcentaje As String, Clientes As Byte, SoloElDeDucible As Boolean)

    Set RT = New ADODB.Recordset
        
    If Clientes = 0 Then
        'Clientes
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        'Proveedores DEDUCIBLE
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    End If



    
    'Comprobaremos para los tres tipos de iva
    'En el futuro podremos desglosar por tipo de iva, empresa y demas
    d = TransformaComasPuntos(CStr(Porcentaje))
'    For A1 = 1 To 3
        If Clientes = 0 Then
            Sql = "select baseimpo, impoiva from " & vCta & ".factcli_totales INNER JOIN " & vCta & ".factcli ON factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
            Sql = Sql & "," & vCta & ".tiposiva where factcli_totales.codigiva = tiposiva.codigiva and tipodiva"
            If SoloElDeDucible Then
                Sql = Sql & "="
            Else
                Sql = Sql & "<>"
            End If
            Sql = Sql & "4 AND  factcli_totales.porciva="
                
        Else
            Sql = "select baseimpo, impoiva from " & vCta & ".factpro_totales INNER JOIN " & vCta & ".factpro ON factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
            Sql = Sql & "," & vCta & ".tiposiva where factpro_totales.codigiva = tiposiva.codigiva and tipodiva"
            If SoloElDeDucible Then
                Sql = Sql & "="
            Else
                Sql = Sql & "<>"
            End If
            Sql = Sql & "4 AND  factpro_totales.porciva="
            
        End If
        Sql = Sql & d
        Sql = Sql & " AND " & Aux
        
        RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpD = ImpD + RT.Fields(0)
            If Not IsNull(RT.Fields(1)) Then ImpH = ImpH + RT.Fields(1)
        End If
        RT.Close
'    Next A1
    Set RT = Nothing
End Sub




'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'TOTAL RETENCION    nuevo                    26 JULIO 2005
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Private Sub TotalRetencion(Porcentaje As String, Clientes As Byte, SoloElDeDucible As Boolean)

    Set RT = New ADODB.Recordset
        
    If Clientes = 0 Then
        'Clientes
        Aux = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        'Proveedores DEDUCIBLE
        Aux = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
'        AUx = AUx & " AND nodeducible = "
'
'        'Modificacion 31 Enero
'        'PROVEEDORES NO deducibles
'        If Clientes = 1 Then
'            AUx = AUx & "0"
'        Else
'            If Clientes = 2 Then
'
'                If Not SoloElDeDucible Then
'                    AUx = AUx & "1"
'                Else
'                    'Facturas con tipo IVA NO deducible
'                    AUx = AUx & "0"
'                End If
'            End If
'        End If
    End If



    
    'Comprobaremos para los tres tipos de iva
    'En el futuro podremos desglosar por tipo de iva, empresa y demas
    d = TransformaComasPuntos(CStr(Porcentaje))
'    For A1 = 1 To 3
        If Clientes = 0 Then
            Sql = "select sum(baseimpo),sum(imporec) from " & vCta & ".factcli_totales INNER JOIN " & vCta & ".factcli ON factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
            'MODIFICACION 16 MAYO 2005
            ' IVA NO DEDUCIBLE
            Sql = Sql & "," & vCta & ".tiposiva WHERE tiposiva.codigiva = factcli_totales.codigiva and tipodiva"
            If SoloElDeDucible Then
                Sql = Sql & "="
            Else
                Sql = Sql & "<>"
            End If
            Sql = Sql & "4 AND  porcrec="
                
        Else
        
            Sql = "select sum(baseimpo),sum(imporec) from " & vCta & ".factpro_totales INNER JOIN " & vCta & ".factpro ON factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
            'MODIFICACION 16 MAYO 2005
            ' IVA NO DEDUCIBLE
            
            Sql = Sql & "," & vCta & ".tiposiva WHERE tiposiva.codigiva = factpro_totales.codigiva and tipodiva"
            If SoloElDeDucible Then
                Sql = Sql & "="
            Else
                Sql = Sql & "<>"
            End If
            Sql = Sql & "4 AND  porcrec="
            
        End If
        Sql = Sql & d
        Sql = Sql & " AND " & Aux
        
        RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpD = ImpD + RT.Fields(0)
            If Not IsNull(RT.Fields(1)) Then ImpH = ImpH + RT.Fields(1)
        End If
        RT.Close
'    Next A1
    Set RT = Nothing
End Sub







'0.-Cli
'1.- Recargo equivalencia
'2.- PRoveeed
'3.- Recargo equivalencia PRO
'5.- Prov. NO deducible

'me falta detallar
'para los recrgos de equivalencia para ello segunsea 0,1 o no deducible sera clientes
'y para 2,3 y no deducbile sera proveedores
'PERO, el 1 sera sobre los RECARGOS DE EQUIVALENCIA
Private Sub GeneraIVADetallado(Clientes As Byte)
Dim C As String
Dim Insertar As Boolean

    

    If Clientes < 2 Then
        Aux = "cl"
    Else
        Aux = "pr"
    End If
    
    Set RT = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    
    
    'Generamos el SQL para la insercion
    Sql = "Select * from  tmpimpbalance WHERE codusu=" & vUsu.Codigo
    'MAYO 2005
    'Los valores de abajo los pondremos a mano
    'SQL = SQL & " AND pasivo=" & Clientes
    'SQL = SQL & " AND codigo ="
    Sql = Sql & " AND pasivo="
    
    d = "INSERT INTO tmpimpbalance (codusu, Pasivo,codigo,importe1, importe2,descripcion,linea ) VALUES ("
    d = d & vUsu.Codigo & ","
    
    If Clientes < 2 Then
        C = " fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Else
        C = " fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    End If
    
'    If Clientes > 1 Then
'
'
'        If Clientes = 8 Then
'            'ISP
'            C = C & " AND extranje =3"
'        Else
'            C = C & " AND extranje < 1"
'
'        End If
'
'        C = C & " AND nodeducible = "
'        If Clientes <> 4 Then
'            C = C & "0"
'        Else
'            C = C & "1"
'        End If
'
'
'
'    End If
    
    
    
    '
    'Para las tres bases
    '------------------
'    For M1 = 1 To 3
        '1 y 3 son los RECARGOS DE EQUIVALENCIA
        If Clientes <> 1 And Clientes <> 3 Then
            '-----
            'IVAS
            '-----
'            Codigo = "SELECT Sum(ba" & M1 & "fac" & AUx & ") AS Sumab, Sum(ti" & M1 & "fac" & AUx & ") AS SumaT, tp" & M1 & "fac" & AUx
'            'mayo2005
'            'Codigo = Codigo & " From " & vCta & ".cabfact"
'            Codigo = Codigo & ",tipodiva From " & vCta & ".tiposiva," & vCta & ".cabfact"
            
            Codigo = "SELECT Sum(baseimpo) AS Sumab, Sum(impoiva) AS SumaT, factcli_totales.codigiva, tipodiva "
            'mayo2005
            'Codigo = Codigo & " From " & vCta & ".cabfact"
            Codigo = Codigo & " From " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
            
        
            
        
        Else
            '--------------------
            'RECARGO EQUIVALENCIA
            '--------------------
'            Codigo = "SELECT Sum(ba" & M1 & "fac" & AUx & ") AS Sumab, Sum(tr" & M1 & "fac" & AUx & ") AS SumaT, tp" & M1 & "fac" & AUx
'            'mayo2005
'            Codigo = Codigo & ",tipodiva From " & vCta & ".tiposiva," & vCta & ".cabfact"
            Codigo = "SELECT Sum(baseimpo) AS Sumab, Sum(imporec) AS SumaT, factcli_totales.codigiva, tipodiva"
            'mayo2005
            Codigo = Codigo & " From " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
            
        
        End If
        If Clientes > 1 Then Codigo = Replace(Codigo, "factcli", "factpro") 'Codigo = Codigo & "prov"
        
        Codigo = Codigo & " WHERE " & C
        Codigo = Codigo & " AND "
        Codigo = Codigo & vCta & ".tiposiva.codigiva = " & vCta & ".factcli_totales"
        

        Codigo = Codigo & ".codigiva" '".tp" & M1 & "fac" & AUx
        
        
        If Clientes > 1 Then Codigo = Replace(Codigo, "factcli", "factpro") 'Codigo & "prov"
        
        
        If Clientes > 1 Then
            Codigo = Codigo & " and " & vCta & ".factpro_totales.numserie = " & vCta & ".factpro.numserie "
            Codigo = Codigo & " and " & vCta & ".factpro_totales.numregis = " & vCta & ".factpro.numregis "
            Codigo = Codigo & " and " & vCta & ".factpro_totales.anofactu = " & vCta & ".factpro.anofactu "
        Else
            Codigo = Codigo & " and " & vCta & ".factcli_totales.numserie = " & vCta & ".factcli.numserie "
            Codigo = Codigo & " and " & vCta & ".factcli_totales.numfactu = " & vCta & ".factcli.numfactu "
            Codigo = Codigo & " and " & vCta & ".factcli_totales.anofactu = " & vCta & ".factcli.anofactu "
        
        End If
        
        
        'TEngo que separar las facturas deducibles de las no deducibles, en IVA
        'SOLO para proveedores
        If Clientes >= 2 Then
            If Clientes = 5 Then
                Codigo = Codigo & " and tipodiva = 4"
            Else
                Codigo = Codigo & " and tipodiva <> 4"
            End If
        End If
        

        
        Codigo = Codigo & " GROUP BY codigiva"
        RT.Open Codigo, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
    
        While Not RT.EOF
            'Para cada tipo, para la empresa esta
            If Not IsNull(RT.Fields(2)) Then
                A2 = RT.Fields(2)
                If IsNull(RT!sumab) Then
                    ImPerD = 0
                Else
                    ImPerD = RT!sumab
                End If
                If IsNull(RT!sumat) Then
                    ImPerH = 0
                Else
                    ImPerH = RT!sumat
                End If
                
                'Si es retencion, el importe tendremos que comprobar que es superior a 0
                Insertar = True
                If Clientes = 1 Or Clientes = 3 Then
                      Insertar = Not (ImPerH = 0)
                End If
                                            'EL DEDUCIBLE
                If Insertar Then InsertaIVADetallado Clientes
            End If
            RT.MoveNext
        Wend
        RT.Close
 '   Next M1

    Set miRsAux = Nothing
    Set RT = Nothing
End Sub


'No le pasamos parametros pq las variables k va a utilizar son globales
Private Sub InsertaIVADetallado(Clientes2 As Byte)

    H = Sql & Clientes2 & " AND Codigo ="
    miRsAux.Open H & A2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ImpD = 0
    ImpH = 0
    If miRsAux.EOF Then
        'No esta insertado
        M2 = 0
    Else
        'Ya esta insertado
        If Not IsNull(miRsAux!Importe1) Then ImpD = miRsAux!Importe1
        If Not IsNull(miRsAux!importe2) Then ImpH = miRsAux!importe2
        M2 = 1
    End If
    miRsAux.Close
    
    ImpD = ImpD + ImPerD
    ImpH = ImpH + ImPerH
    
    
    'Cargamos sobre H
    If M2 = 0 Then
        'Nuevo
        
        'Ponemos el texto del iva
        H = "Select nombriva,porceiva,tipodiva,porcerec FROM " & vCta & ".tiposiva  WHERE codigiva =" & A2
        miRsAux.Open H, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            'Nombre IVA
            H = miRsAux!nombriva & "','"
            If Clientes2 = 1 Or Clientes2 = 3 Then
                'RECARGO EQUIVALENCIA
                H = H & Format(miRsAux!porcerec, "#0.00")
            Else
                H = H & Format(miRsAux!porceiva, "#0.00")
            End If
        Else
            H = "','"
        End If
        
        miRsAux.Close
        H = ",'" & H & "')"
        'SQL
        
            
        H = d & Clientes2 & "," & A2 & "," & TransformaComasPuntos(CStr(ImpD)) & "," & TransformaComasPuntos(CStr(ImpH)) & H
    Else
        'Modificar
        H = "UPDATE tmpimpbalance SET importe1=" & TransformaComasPuntos(CStr(ImpD))
        H = H & ",importe2 =" & TransformaComasPuntos(CStr(ImpH))
        H = H & " WHERE codusu=" & vUsu.Codigo & " AND Pasivo = " & Abs(Clientes2)
        H = H & " AND codigo =" & A2
    End If
    Conn.Execute H
End Sub








Private Sub IvaDeducibleBienInversion(Porcentaje As String)
Dim R As ADODB.Recordset

    ImPerD = 0: ImPerH = 0
    Set R = New ADODB.Recordset
        Codigo = "SELECT Sum(baseimpo) AS Sumab, Sum(impoiva) AS SumaT"
        Codigo = Codigo & " From " & vCta & ".tiposiva tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
        Codigo = Codigo & " WHERE factpro_totales.codigiva = tiposiva.codigiva and tipodiva=2"  'Bien inversion
        Codigo = Codigo & " AND fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Codigo = Codigo & " AND porciva = " & TransformaComasPuntos(Porcentaje)
        Codigo = Codigo & " AND factpro_totales.numserie = factpro.numserie "
        Codigo = Codigo & " AND factpro_totales.numregis = factpro.numregis "
        Codigo = Codigo & " AND factpro_totales.anofactu = factpro.anofactu "

        R.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not R.EOF Then
            If Not IsNull(R!sumab) Then
                ImPerD = ImPerD + R!sumab
                ImPerH = ImPerH + DBLet(R!sumat, "N")
            End If
        End If
        R.Close
'    Next M1
    Set R = Nothing
    
End Sub

Private Function CargaAcumuladosTotalesCerrados(ByRef Cta As String) As Boolean
    CargaAcumuladosTotalesCerrados = False
    ImpD = 0
    ImpH = 0
    Sql = "UPDATE tmpconextcab SET acumtotD= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    Sql = Sql & ", acumtotH= " & TransformaComasPuntos(CStr(ImpH)) 'Format(ImpH, "#,###,##0.00")
    ImpD = ImpD - ImpH
    Sql = Sql & ", acumtotT= " & TransformaComasPuntos(CStr(ImpD)) 'Format(ImpD, "#,###,##0.00")
    Sql = Sql & " WHERE codusu=" & vUsu.Codigo & " AND cta='" & Cta & "'"
    Conn.Execute Sql
    CargaAcumuladosTotalesCerrados = True
End Function

Public Function BloqueoManual(bloquear As Boolean, tabla As String, Clave As String) As Boolean
    If bloquear Then
        Sql = "INSERT INTO zbloqueos (codusu, tabla, clave) VALUES (" & vUsu.Codigo
        Sql = Sql & ",'" & UCase(tabla) & "','" & UCase(Clave) & "')"
    Else
        Sql = "DELETE FROM zbloqueos where codusu = " & vUsu.Codigo & " AND tabla ='"
        Sql = Sql & tabla & "'"
    End If
    On Error Resume Next
    Conn.Execute Sql
    If Err.Number <> 0 Then
        Err.Clear
        BloqueoManual = False
    Else
        BloqueoManual = True
    End If
End Function





'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'   Cuenta explotacion por centro de coste
'
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Cadena fechas tendra, enpipado, mesinicio,añoinicio, mespedido,anopedido,mesfin,anofin
'

    
    












'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'           L I B R O        R E S U M E N
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'
'
'
'
'
'   Para los asiento k haya k quitar saldos estos los guardaremos en la tabla
'
'


Public Sub FijaValoresLibroResumen(FIni As Date, FFin As Date, Nivel As Integer, EjerciciosCerr As Boolean, NumAsiento As String)
    Sql = "INSERT INTO tmpdirioresum (codusu, clave, fecha, asiento, cuenta, titulo, concepto, debe, haber) VALUES (" & vUsu.Codigo & ","
    vFecha1 = FIni
    vFecha2 = FFin
    
    'Septiembre 2020
    A3 = Nivel
    If Nivel = 10 Then A3 = vEmpresa.DigitosUltimoNivel
    
    
    
    EjerciciosCerrados = EjerciciosCerr
    
    'Numero de asiento
    A2 = 1
    If NumAsiento <> "" Then
        If IsNumeric(NumAsiento) Then A2 = CInt(NumAsiento)
    End If
        
        
    'M2 sera el contador para cada registro
    M22 = 122222
End Sub



Public Function ProcesaLibroResumen(Mes As Long, Ano As Integer, I1 As Currency, I2 As Currency)
Dim Opcion As Byte
    ' 0.- Mes normal
    ' 1.- Mes con apertura
    ' 2.- Mes con cierre
Dim TienekKitar As Boolean

    A1 = Ano
    M1 = CInt(Mes)
  

    
    Set RT = New ADODB.Recordset
    'Comprobamos si tiene el mes de apertura de ejercicio
    Opcion = 0
    If M1 = Month(vFecha1) Then
        Opcion = 1
    Else
        If (M1 = Month(vFecha2)) Then Opcion = 2
    End If
        
        
    TienekKitar = False
    If Opcion = 1 Then
        TienekKitar = True
        NumAsiento = A2
        GeneraAperturaResumen 0
    Else
        If Opcion = 2 Then
            TienekKitar = True
            NumAsiento = A2 + 1
            GeneraAperturaResumen 1
            GeneraAperturaResumen 2
            GeneraAperturaResumen 3
            NumAsiento = A2
        Else
            NumAsiento = A2
        End If
    End If
    
    'hacemos el mes
    If I1 <> 0 Or I2 <> 0 Then
        'Insertaremos el acumulado que nos han indicado
        vCta = CStr(DiasMes(CByte(M1), A1))
        vCta = vCta & "/" & M1 & "/" & A1
        VFecha3 = CDate(vCta)
        vCta = "'','ACUMULADO ANTERIOR'"
        Codigo = ""
        
        ImpD = I1
        ImpH = I2
        InsertaParaListadoDiarioResum
    End If
    HacerMes TienekKitar
    A2 = NumAsiento
    'Si tiene fin hacer fin
    Set RT = Nothing
End Function




Private Sub HacerMes(HayKRestarSaldos As Boolean)
Dim RTT As Recordset
Dim F1 As Date
Dim F2 As Date

   F2 = CDate(Format(DiasMes(CByte(M1), A1), "00") & "/" & Format(M1, "00") & "/" & Format(A1, "0000"))
   F1 = CDate("01/" & Format(M1, "00") & "/" & Format(A1, "0000"))

   Aux = "select sum(coalesce(timported,0)),sum(coalesce(timporteh,0)),mid(hlinapu.codmacta,1," & A3 & "),cuentas.nommacta from hlinapu"
   Aux = Aux & ",cuentas where "
   Aux = Aux & "mid(hlinapu.codmacta,1," & A3 & ")=cuentas.codmacta  "
   Aux = Aux & " and fechaent between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
   'Aux = Aux & " group by hlinapu.codmacta order by hlinapu.codmacta"  quito estoAbril2017
   Aux = Aux & " group by 3 order by 3"
   
   Set RTT = New ADODB.Recordset
   RTT.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   If RTT.EOF Then
        RTT.Close
        Exit Sub
    End If
    
    
   d = CStr(DiasMes(CByte(M1), A1))
   d = d & "/" & M1 & "/" & A1
   VFecha3 = CDate(d)
   'para cada valor insertaremos en la tabla
   While Not RTT.EOF
        d = RTT.Fields(2)
        If HayKRestarSaldos Then
            FijaImporteResta (d)
        Else
            ImAcD = 0
            ImAcH = 0
        End If
        
        ImpD = 0
        ImpH = 0
        If Not IsNull(RTT.Fields(0)) Then ImpD = RTT.Fields(0)
        If Not IsNull(RTT.Fields(1)) Then ImpH = RTT.Fields(1)
   
        ImpD = ImpD - ImAcD
        ImpH = ImpH - ImAcH
        
        
        vCta = "'" & RTT.Fields(2) & "','" & DevNombreSQL(RTT.Fields(3)) & "'"
        Codigo = ""
        InsertaParaListadoDiarioResum
                
        
        RTT.MoveNext
    Wend
    RTT.Close
    Set RTT = Nothing
    NumAsiento = NumAsiento + 1
End Sub






Private Sub FijaImporteResta(ByRef KCuenta As String)
Dim Au As String


    Au = "Select Debe, Haber from tmpdiarresum WHERE codusu =" & vUsu.Codigo & " AND codmacta ='" & KCuenta & "';"
    ImAcD = 0
    ImAcH = 0
    RT.Open Au, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImAcD = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImAcH = RT.Fields(1)
    End If
    RT.Close
    
End Sub






'Opcion
'   0.- Apertura
'   1.- PyG
'   2.- Cierre
'   3.- pyG y cierre
Private Sub GeneraAperturaResumen(Opcion As Byte)
Dim Rs As Recordset

    Conn.Execute "DELETE from tmpdiarresum where codusu =" & vUsu.Codigo
    
    Aux = "SELECT "
    If A3 <> vEmpresa.DigitosUltimoNivel Then Aux = Aux & "substring(codmacta,1," & A3 & ")  "
    Aux = Aux & "codmacta , " & vUsu.Codigo & " codusu , sum(coalesce(timported,0)) ,sum(coalesce(timporteh,0))"
    
    'Aux = "Select codmacta from hlinapu"
    Aux = Aux & " from hlinapu"
    If EjerciciosCerrados Then Aux = Aux & "1"
    Aux = Aux & " WHERE "
    Select Case Opcion
    Case 0
        'Apertura. El primero
        Aux = Aux & "codconce = 970"
    Case 1
        'Py G
        'Eje: Diciembre:    234
        '     Py G:         235
        '     Cierre:       236
        Aux = Aux & "codconce = 960"
    Case 2
        Aux = Aux & "codconce = 980"
    Case 3
        Aux = Aux & "(codconce = 960 or codconce = 980)"
        'Este no insertara
    End Select
    'Fechas
    Aux = Aux & " AND fechaent >='" & Format(vFecha1, FormatoFecha)
    Aux = Aux & "' AND fechaent <='" & Format(vFecha2, FormatoFecha)
    Aux = Aux & "' GROUP BY 1 "
    
    
    Aux = "INSERT INTO tmpdiarresum(codmacta,codusu,debe,haber) " & Aux
    Conn.Execute Aux
    
    
    '------------->> ANTES SEPTIEMBRE
       
''''                            Set Rs = New ADODB.Recordset
''''                            Rs.Open Aux, Conn, adOpenForwardOnly, adCmdText
''''                            While Not Rs.EOF
''''                                vCta = Rs.Fields(0)
''''                                vCta = Mid(vCta, 1, A3)
''''                                Insertatmpdiarresum
''''                                Rs.MoveNext
''''                            Wend
''''                            Rs.Close
''''
''''
''''                            'Ya tenemos en tmpdiarresum
''''                            'las subcuentas del diario
''''                            Codigo = "Select * from tmpdiarresum where codusu =" & vUsu.Codigo
''''                            Rs.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''                            While Not Rs.EOF
''''                                vCta = Rs.Fields(1)
''''                                'para cada cuenta, obtienes los importes
''''                                CalcularImporteCierreAperturaPyG Opcion
''''                                If ImCierrD <> 0 Or ImCierrH <> 0 Then
''''                                    d = TransformaComasPuntos(CStr(ImCierrD))
''''                                    H = TransformaComasPuntos(CStr(ImCierrH))
''''                                    InsertarTMP
''''                                Else
''''                                    Conn.Execute "DELETE from tmpdiarresum where codusu =" & vUsu.Codigo & " AND codmacta='" & vCta & "';"
''''                                End If
''''                                Rs.MoveNext
''''                            Wend
''''                            Rs.Close
''''
    
    'Ahora si la opcion es 3 no seguimos. Esto de abajo es para insertar
    If Opcion = 3 Then Exit Sub
    
    'Volvemos abri el temporal del diario resumen
    Set Rs = New ADODB.Recordset
    Codigo = "select tmpdiarresum.*, nommacta from tmpdiarresum,cuentas where tmpdiarresum.codmacta = cuentas.codmacta and codusu =" & vUsu.Codigo
    Codigo = Codigo & " order by codmacta"
    Rs.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Select Case Opcion
    Case 0
        VFecha3 = vFecha1
        Codigo = "APERTURA"
    Case 1
        Codigo = "PERDIDAS Y GANACIAS"
        VFecha3 = vFecha2
    Case 2
        Codigo = "CIERRE"
        VFecha3 = vFecha2
    End Select
    Codigo = Codigo & " AL "
    
        
    If Not Rs.EOF Then
        While Not Rs.EOF
            ImpD = Rs.Fields(2)
            ImpH = Rs.Fields(3)
            vCta = "'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs.Fields(4)) & "'"
            InsertaParaListadoDiarioResum
            Rs.MoveNext
        Wend
        'Aumentamos el contador
        NumAsiento = NumAsiento + 1
    End If
    Rs.Close
    Set Rs = Nothing
End Sub

'Septiembre 2020
'MAL PARIDO
'''Private Sub GeneraAperturaResumen(Opcion As Byte)
'''Dim Rs As Recordset
'''
'''    Conn.Execute "DELETE from tmpdiarresum where codusu =" & vUsu.Codigo
'''
'''    Aux = "Select codmacta from hlinapu"
'''    If EjerciciosCerrados Then Aux = Aux & "1"
'''    Aux = Aux & " WHERE "
'''    Select Case Opcion
'''    Case 0
'''        'Apertura. El primero
'''        Aux = Aux & "codconce = 970"
'''    Case 1
'''        'Py G
'''        'Eje: Diciembre:    234
'''        '     Py G:         235
'''        '     Cierre:       236
'''        Aux = Aux & "codconce = 960"
'''    Case 2
'''        Aux = Aux & "codconce = 980"
'''    Case 3
'''        Aux = Aux & "(codconce = 960 or codconce = 980)"
'''        'Este no insertara
'''    End Select
'''    'Fechas
'''    Aux = Aux & " AND fechaent >='" & Format(vFecha1, FormatoFecha)
'''    Aux = Aux & "' AND fechaent <='" & Format(vFecha2, FormatoFecha)
'''    Aux = Aux & "' GROUP BY codmacta"
'''
'''    Set Rs = New ADODB.Recordset
'''    Rs.Open Aux, Conn, adOpenForwardOnly, adCmdText
'''    While Not Rs.EOF
'''        vCta = Rs.Fields(0)
'''        vCta = Mid(vCta, 1, A3)
'''        Insertatmpdiarresum
'''        Rs.MoveNext
'''    Wend
'''    Rs.Close
'''
'''
'''    'Ya tenemos en tmpdiarresum
'''    'las subcuentas del diario
'''    Codigo = "Select * from tmpdiarresum where codusu =" & vUsu.Codigo
'''    Rs.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''    While Not Rs.EOF
'''        vCta = Rs.Fields(1)
'''        'para cada cuenta, obtienes los importes
'''        CalcularImporteCierreAperturaPyG Opcion
'''        If ImCierrD <> 0 Or ImCierrH <> 0 Then
'''            d = TransformaComasPuntos(CStr(ImCierrD))
'''            H = TransformaComasPuntos(CStr(ImCierrH))
'''            InsertarTMP
'''        Else
'''            Conn.Execute "DELETE from tmpdiarresum where codusu =" & vUsu.Codigo & " AND codmacta='" & vCta & "';"
'''        End If
'''        Rs.MoveNext
'''    Wend
'''    Rs.Close
'''
'''
'''    'Ahora si la opcion es 3 no seguimos. Esto de abajo es para insertar
'''    If Opcion = 3 Then Exit Sub
'''
'''    'Volvemos abri el temporal del diario resumen
'''    Codigo = "select tmpdiarresum.*, nommacta from tmpdiarresum,cuentas where tmpdiarresum.codmacta = cuentas.codmacta and codusu =" & vUsu.Codigo
'''    Codigo = Codigo & " order by codmacta"
'''    Rs.Open Codigo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''    Select Case Opcion
'''    Case 0
'''        VFecha3 = vFecha1
'''        Codigo = "APERTURA"
'''    Case 1
'''        Codigo = "PERDIDAS Y GANACIAS"
'''        VFecha3 = vFecha2
'''    Case 2
'''        Codigo = "CIERRE"
'''        VFecha3 = vFecha2
'''    End Select
'''    Codigo = Codigo & " AL "
'''
'''
'''    If Not Rs.EOF Then
'''        While Not Rs.EOF
'''            ImpD = Rs.Fields(2)
'''            ImpH = Rs.Fields(3)
'''            vCta = "'" & Rs.Fields(1) & "','" & DevNombreSQL(Rs.Fields(4)) & "'"
'''            InsertaParaListadoDiarioResum
'''            Rs.MoveNext
'''        Wend
'''        'Aumentamos el contador
'''        NumAsiento = NumAsiento + 1
'''    End If
'''    Rs.Close
'''    Set Rs = Nothing
'''End Sub







Private Sub InsertarTMP()
    Aux = "UPDATE tmpdiarresum set debe = " & d & ", Haber = " & H
    Aux = Aux & " WHERE codmacta = '" & vCta & "' and codusu = " & vUsu.Codigo
    Conn.Execute Aux
End Sub

Private Sub Insertatmpdiarresum()
On Error Resume Next
Conn.Execute "INSERT INTO tmpdiarresum (codusu, codmacta) VALUES (" & vUsu.Codigo & ",'" & vCta & "')"
If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub InsertaParaListadoDiarioResum()
If ImpD <> 0 Then
    d = TransformaComasPuntos(CStr(ImpD))
    H = "NULL"
    InsertaParaListadoDiarioResum4 "DEBE"
End If
If ImpH <> 0 Then
    H = TransformaComasPuntos(CStr(ImpH))
    d = "NULL"
    InsertaParaListadoDiarioResum4 "HABER"
End If



End Sub

Private Sub InsertaParaListadoDiarioResum4(DebeHaber As String)
Dim C As String

' clave, fecha, asiento, cuenta, titulo, concepto, debe, haber)
C = Sql & M22 & ",'" & Format(VFecha3, "dd/mm/yyyy") & "'," & NumAsiento & "," & vCta
C = C & ",'" & Codigo & DebeHaber & "'," & d & "," & H & ")"

Conn.Execute C
M22 = M22 + 1
End Sub



Private Sub CalcularImporteCierreAperturaPyG(Opcion As Byte)

    Aux = "Select SUM(timporteD),sum(timporteH) from hlinapu"
    If EjerciciosCerrados Then Aux = Aux & "1"
    Aux = Aux & " WHERE codmacta  like '" & vCta & "%"
    Aux = Aux & "' and "
    Select Case Opcion
    Case 0
        'Apertura. El primero
        Aux = Aux & "codconce = 970"
    Case 1
        Aux = Aux & "codconce = 960"
    Case 2
        Aux = Aux & "codconce = 980"
    Case 3
        Aux = Aux & "(codconce = 960 or codconce = 980)"
        'Este no insertara
    End Select
    Aux = Aux & " AND fechaent >= '" & Format(vFecha1, FormatoFecha) & "'"
    Aux = Aux & " AND fechaent <= '" & Format(vFecha2, FormatoFecha) & "'"
    Aux = Aux & ";"

    
    RT.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If IsNull(RT.Fields(0)) Then
        ImCierrD = 0
    Else
        ImCierrD = RT.Fields(0)
    End If
    If IsNull(RT.Fields(1)) Then
        ImCierrH = 0
    Else
        ImCierrH = RT.Fields(1)
    End If
    RT.Close
End Sub





'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'           Detalles cuenta por centro de coste
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'
'
'
'
'
'
'
'
'

'En nombres iran empipaditos nommacta y nomccost
Public Sub Cta_por_CC(ByRef vCuenta As String, vCCos As String, Nombres As String)
Dim miSQL As String
Dim cad As String
    vCta = vCuenta
    Codigo = vCCos
    Set RT = New ADODB.Recordset
    If vFecha1 > VFecha3 Then
        'Calculamos anteriores
        ImAcD = 0
        ImAcH = 0
        CalculaAnterioresCtaPorCC
    Else
        ImAcD = 0
        ImAcH = 0
    End If
    'En impcierrD llevare el saldo
    ImCierrD = ImAcD - ImAcH
    
    'Importes totales
    ImPerD = ImAcD
    ImPerH = ImAcH
    
'    miSQL = "Select * from hlinapu"
'    If EjerciciosCerrados Then miSQL = miSQL & "1"
'    miSQL = miSQL & " WHERE codmacta ='" & vCta & "'"
'    miSQL = miSQL & " AND codccost ='" & Codigo & "'"
'    miSQL = miSQL & " AND fechaent >='" & Format(vFecha1, FormatoFecha) & "'"
'    miSQL = miSQL & " AND fechaent <='" & Format(vFecha2, FormatoFecha) & "'"
'    miSQL = miSQL & " ORDER BY fechaent, linliapu"

'SELECT Tabla1.num, Tabla1.cta, cuentas1.nommacta
'FROM Tabla1 LEFT JOIN cuentas1 ON Tabla1.cta = cuentas1.codmacta;

    miSQL = "Select *, cuentas.nommacta from hlinapu"
    If EjerciciosCerrados Then miSQL = miSQL & "1"
    miSQL = miSQL & " LEFT JOIN cuentas ON hlinapu"
    If EjerciciosCerrados Then miSQL = miSQL & "1"
    miSQL = miSQL & ".ctacontr = cuentas.codmacta WHERE hlinapu"
    If EjerciciosCerrados Then miSQL = miSQL & "1"
    miSQL = miSQL & ".codmacta ='" & vCta & "'"
    miSQL = miSQL & " AND codccost ='" & Codigo & "'"
    miSQL = miSQL & " AND fechaent >='" & Format(vFecha1, FormatoFecha) & "'"
    miSQL = miSQL & " AND fechaent <='" & Format(vFecha2, FormatoFecha) & "'"
    miSQL = miSQL & " ORDER BY fechaent, linliapu"



    RT.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    A1 = 0
    miSQL = Sql & "'" & Codigo & "','" & vCta & "',"
    While Not RT.EOF
        A1 = A1 + 1
        cad = A1 & ",'" & DevNombreSQL(DBLet(RT!Numdocum)) & "','"
        cad = cad & Format(RT!FechaEnt, FormatoFecha) & "','" & DevNombreSQL(DBLet(RT!Ampconce)) & "',"
        If IsNull(RT!timported) Then
            ImpD = 0
            d = "NULL"
        Else
            ImpD = RT!timported
            d = TransformaComasPuntos(CStr(RT!timported))
        End If
            
        'importe HABER
        If IsNull(RT!timporteH) Then
            ImpH = 0
            H = "NULL"
        Else
            ImpH = RT!timporteH
            H = TransformaComasPuntos(CStr(RT!timporteH))
        End If
        cad = cad & d & "," & H & ","
        
        ImPerD = ImPerD + ImpD
        ImPerH = ImPerH + ImpH
        
        'Saldo
        ImCierrH = ImpD - ImpH
        ImCierrD = ImCierrD + ImCierrH
        H = TransformaComasPuntos(CStr(ImCierrD))
        cad = cad & H
        
        
        'Ctra partida
        If IsNull(RT!ctacontr) Then
            cad = cad & ",'',''"
        Else
            cad = cad & ",'" & RT!ctacontr & "','" & DevNombreSQL(DBLet(RT!Nommacta)) & "'"
        End If
        
        
        cad = cad & ")"
        'Ejecutamos
        Conn.Execute miSQL & cad
    
        'Sig
        RT.MoveNext
    Wend
    RT.Close
    
    
    
    'La cabecera
    '->INSERT INTO zcabccexplo (codusu, codccost, codmacta, nommacta, nomccost,
    '-->acumD, acumH, acumS, totD, totH, totS) VALUES (
    cad = "'" & Codigo & "','" & vCta & "','" & DevNombreSQL(RecuperaValor(Nombres, 1))
    cad = cad & "','" & DevNombreSQL(RecuperaValor(Nombres, 2)) & "',"
    
    '-----------------------------------------
    'Acumulado anterior
    If ImAcD = 0 And ImAcH = 0 Then
        cad = cad & """"""
    Else
        cad = cad & """S"""
    End If
    cad = cad & ","
    If ImAcH = 0 Then
        H = "NULL"
    Else
        H = TransformaComasPuntos(CStr(ImAcH))
    End If
    'Acumulado anterior
    If ImAcD = 0 Then
        d = "NULL"
    Else
        d = TransformaComasPuntos(CStr(ImAcD))
    End If
    cad = cad & d & "," & H
    
    
        
    
    
    'SALDO anterior
    ImAcD = ImAcD - ImAcH
    If ImAcD = 0 Then
        H = "NULL"
    Else
        H = TransformaComasPuntos(CStr(ImAcD))
    End If
    cad = cad & "," & H

    '--------------------- TOTALES
    d = TransformaComasPuntos(CStr(ImPerD))
    H = TransformaComasPuntos(CStr(ImPerH))
    cad = cad & "," & d & "," & H
    
    'Saldo final
    d = TransformaComasPuntos(CStr(ImCierrD))
    cad = cad & "," & d & ")"
    
    Conn.Execute Aux & cad
    
End Sub


Private Sub CalculaAnterioresCtaPorCC()
Dim C As String

    C = "Select sum(timported),sum(timporteh) from hlinapu"
    If EjerciciosCerrados Then C = C & "1"
    C = C & " WHERE codmacta ='" & vCta & "'"
    C = C & " AND codccost ='" & Codigo & "'"
    C = C & " AND fechaent >='" & Format(VFecha3, FormatoFecha) & "'"
    C = C & " AND fechaent <='" & Format(vFecha1, FormatoFecha) & "'"
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT.Fields(0)) Then ImAcD = RT.Fields(0)
        If Not IsNull(RT.Fields(1)) Then ImAcH = RT.Fields(1)
    End If
    RT.Close
End Sub














'-----------------------------------------------------------------------
' Desde aqui generaremos el encabezado para las cartas.
'
'       Tipo 0: Cartas para el 347 del IVA
'       Tipo 1: Factura venta inmovilizado
'
'
Public Function CargaEncabezadoCarta(Opcion As Byte, Optional ByRef contacto As String)
    
    'El contatacto par el futuro
    Codigo = DevNombreSQL(contacto)
    
    'Borramos el anterior
    Sql = "DELETE FROm tmpTesoreria2 WHERE codusu = " & vUsu.Codigo
    Conn.Execute Sql

    'Cadena insert
    Sql = "INSERT INTO tmpTesoreria2 (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos,"
    Sql = Sql & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, Asunto, Referencia, contacto) VALUES ("
    Sql = Sql & vUsu.Codigo
    
    'Los datos de la empresa son comunes
    'El resto de sql lo montamos en H
    H = ""
    MontaDatosEmpresa
    Sql = Sql & H
        
    'Por si da fallo, o para el inmovilizado
    H = ""
    If Opcion = 0 Then
        'Para el 347. Cogeremos los datos de un achivo
        Monta347
    Else
        MontaFacturaVenta
    End If
    Sql = Sql & H & ")"
    Conn.Execute Sql
End Function


Private Sub MontaDatosEmpresa()
    Set RT = New ADODB.Recordset
    RT.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    If RT.EOF Then
        MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
        H = ",'','','','','',''"  '6 campos
    Else
        H = ",'" & DBLet(RT!nifempre) & "','" & vEmpresa.NombreEmpresaOficial & "','"
        d = DBLet(RT!siglasvia) & " " & DBLet(RT!Direccion) & "  " & DBLet(RT!numero) & ", " & DBLet(RT!puerta)
        H = H & d & "','" & DBLet(RT!codpos) & "','" & DBLet(RT!Poblacion) & "','" & DBLet(RT!provincia) & "'"
    End If
    RT.Close
    Set RT = Nothing
End Sub


Private Sub Monta347()
Dim Fin As Boolean
On Error GoTo Emon347
    M1 = FreeFile
    'Archivo con los datos
    vCta = App.Path & "\txt347.dat"
    If Dir(vCta) = "" Then
        H = ",'','','','','','','','',''"   '8 pares
        H = H & ",'" & Codigo & "'"
        Exit Sub
    End If
    
    Open vCta For Input As #M1
    M2 = 0
    d = ""
    While Not Fin
        M2 = M2 + 1
        If M2 <= 9 Then
            'Las lineas van por pares, y hay 8 pares
            Line Input #M1, vCta
            Line Input #M1, vCta
            d = d & ",'" & vCta & "'"
            Fin = EOF(M1)
        Else
            Fin = True
        End If
    Wend
    
    Close #M1
    
    If M2 = 9 Then
        d = d & ",'" & Codigo & "'"
        H = d
    End If
    Exit Sub
Emon347:
    MuestraError Err.Number, "Fichero datos para el 347"
End Sub


Private Sub MontaFacturaVenta()

    'SABEMOS K en CadenaDesdeOtroForm estan los valores a guardar
    For A1 = 1 To 6
        H = H & ",'" & RecuperaValor(CadenaDesdeOtroForm, A1) & "'"
    Next A1
    
    For A1 = 7 To 10
        H = H & ",NULL"
    Next A1
    
    
End Sub













'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'
'
'       Comprobar Formula de Configuracion de balabnce
'
'
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
Public Function CompruebaFormulaConfigBalan(NumBalan As Integer, Formula As String) As String

    CompruebaFormulaConfigBalan = ""

    NumAsiento = NumBalan
    
    
    'Kitamos todos los esapacios en blanco
    Formula = Trim(Formula)
    Do
        A1 = InStr(1, Formula, " ")
        If A1 > 0 Then Formula = Mid(Formula, 1, A1 - 1) & Mid(Formula, A1 + 1)
    Loop Until A1 = 0
    
    
    
    'Comprobamos k los caracteres son correctos
    M2 = 1 'Bien
    For M1 = 1 To Len(Formula)
        d = Mid(Formula, M1, 1)
        Select Case d
        Case "0" To "9"
            'Son los numeros
            
        Case "+", "-"
            'El mas y el menos
            
        Case "A", "B"
        
        Case Else
            M2 = 0
            Exit For
        End Select
    Next M1
    
    If M2 = 0 Then
        CompruebaFormulaConfigBalan = "Caracteres incorrectos"
        Exit Function
    End If

    'para cada campo de la formula buscamos "+" o "-"
    M3 = 1
    Set RT = New ADODB.Recordset
    Do
        M1 = 0
        A1 = InStr(1, Formula, "-")
        A2 = InStr(1, Formula, "+")
        If A1 <> 0 Or A2 <> 0 Then
            If A1 = 0 Then A1 = 32000
            If A2 = 0 Then A2 = 32000
            
            If A1 > A2 Then
                M1 = A2
            Else
                M1 = A1
            End If
        Else
            If Formula <> "" Then
                M1 = Len(Formula) + 1
            End If
        End If
        
        If M1 > 0 Then
            d = Mid(Formula, 1, 1)     'activo pasivo, A o B
            H = Mid(Formula, 2, M1 - 2) 'Codigo
            
            If Not ExisteCodigoBalance Then
                CompruebaFormulaConfigBalan = "No existe codigo balance:  " & d & H
                Set RT = Nothing
                Exit Function
            End If
            
            Formula = Mid(Formula, M1 + 1)
        End If
        Loop Until M1 = 0
        
End Function



Private Function ExisteCodigoBalance() As Boolean
On Error GoTo EExisteCodigoBalance
    ExisteCodigoBalance = False
    Sql = "Select * from balances_texto where numbalan=" & NumAsiento & " AND pasivo='" & d
    Sql = Sql & "' AND Codigo=" & H
    RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        If Not IsNull(RT!Codigo) Then
            ExisteCodigoBalance = True
        End If
    End If
    RT.Close
    Exit Function
EExisteCodigoBalance:
    MuestraError Err.Number
End Function




'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'           Impresion de balances configurables
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------


'--> Se puede mejorar puesto k algunas tablas temporales se cargan con datos k luego no
'    son utilizados
'

'Imprime el listado para que vena las cuentas que entran dentro de k punto etc etc
Public Function GeneraDatosBalanConfigImpresion(NumBalan As Integer, ImportesDeSoloUnMes As Boolean)

        Sql = "Delete from "
        Sql = Sql & "tmpimpbalance where codusu = " & vUsu.Codigo
        Conn.Execute Sql
        Conn.Execute "DELETE FROM tmpimpbalan WHERE codusu = " & vUsu.Codigo
        
        Set RT = New ADODB.Recordset
        M3 = 1 'Sera el orden de insercion
        
        'Sera el numero de balance
        NumAsiento = NumBalan
        
        
        'Vemos si es de perdidas y ganacias
        Sql = DevuelveDesdeBD("perdidas", "balances", "numbalan", CStr(NumBalan), "N")
        EsBalancePerdidas_y_ganancias = (Val(Sql) = 1)
        
        
        
        'Vamos a utilizar la temporal de balances donde dejara los valores
        Sql = "DELETE from tmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute Sql
        
        Contabilidad = -1

       
        CargaArbol 0, 0, -1, "", Month(vParam.fechaini), Year(vParam.fechaini), Month(vParam.fechafin), Year(vParam.fechafin), "", True, Nothing, False, ImportesDeSoloUnMes
 
 
        Sql = "Select * from "
        If Contabilidad > 0 Then Sql = Sql & "ariconta" & Contabilidad & "."
        Sql = Sql & "balances_texto where numbalan=" & NumBalan & " AND padre"
        Aux = "INSERT INTO tmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,LibroCD,QueCuentas) VALUES (" & vUsu.Codigo
        Codigo = "Select importe1,importe2,quecuentas from "
        If Contabilidad > 0 Then Codigo = Codigo & "ariconta" & Contabilidad & "."
        Codigo = Codigo & "tmpimpbalance where codusu=" & vUsu.Codigo & " AND pasivo='"
        M1 = 1
 
 
        CargaArbolImpresion -1, "", 1, False, False
End Function


Public Function GeneraDatosBalanceConfigurable(NumBalan As Integer, Mes1 As Integer, Anyo1 As Integer, Mes2 As Integer, Anyo2 As Integer, LibroCD As Boolean, vContabilidad As String, Optional PB As ProgressBar, Optional ImportesDeSoloUnMes As Boolean)
Dim QuitarUno As Boolean
Dim EsPyGNoAbreviado As Boolean
Dim AuxPyG As String

    If vContabilidad = "-1" Then vContabilidad = "-1|"
    
    Set RT = New ADODB.Recordset
    
    
    'Metemos en las variable varconsolidado
    '       (0): fechas hco  . Posiciones fijas dd/mm/yyyy|
    '        1:  quitar 1     0|  o 1|
    '        2:  quitar 2 """"
            
    d = vContabilidad
    VarConsolidado(0) = "": VarConsolidado(1) = "": VarConsolidado(2) = ""
    
    
    
    M2 = 1 'PRUEBA
    
    While d <> ""
            'Vemos cual es
            M1 = InStr(1, d, "|")
            A1 = CInt(Mid(d, 1, M1 - 1))
            d = Mid(d, M1 + 1)
            
            Contabilidad = A1
            
            
            
            
            'Ete trozo es por si tuvieramos hcabapu1 y halinapu1
            VFecha3 = CDate("01/12/1900")

            
            
            VarConsolidado(0) = VarConsolidado(0) & VFecha3 & "|"
            

            '------------------------------------------------
            'Si hay k quitar saldos para cada ano
            'Comprobamos si hay que quitar el pyg y el cierre
            QuitarUno = False
            EjerciciosCerrados = (CDate("15/" & Mes1 & "/" & Anyo1) < VFecha3)
            'Si el mes contiene el cierre, entonces adelante
            If Month(vParam.fechafin) = Mes1 Then
                'Si estamos en ejerccicios cerrados seguro que hay asiento de cierre y p y g
                If EjerciciosCerrados Then
                    QuitarUno = True
                Else
                'Si no lo comprobamos. Concepto=960 y 980
                    QuitarUno = HayAsientoCierreBalances(CByte(Mes1), Anyo1)
                End If
            End If
            VarConsolidado(1) = VarConsolidado(1) & Abs(QuitarUno) & "|"
    
    
            'Si hay k quitar saldos para el segundo
            'Comprobamos si hay que quitar el pyg y el cierre
            QuitarUno = False
            If Mes2 > 0 Then
                EjerciciosCerrados = (CDate("15/" & Mes2 & "/" & Anyo2) < VFecha3)
                'Si el mes contiene el cierre, entonces adelante
                If Month(vParam.fechafin) = Mes2 Then
                    'Si estamos en ejerccicios cerrados seguro que hay asiento de cierre y p y g
                    If EjerciciosCerrados Then
                        QuitarUno = True
                    Else
                    'Si no lo comprobamos. Concepto=960 y 980
                        QuitarUno = HayAsientoCierreBalances(CByte(Mes2), Anyo2)
                    End If
                End If
            End If
            VarConsolidado(2) = VarConsolidado(2) & Abs(QuitarUno) & "|"
    
            
        Wend
    
    
    
    
    
    

        
        'Borramos las temporales
        Sql = "Delete from "
        Sql = Sql & "tmpimpbalance where codusu = " & vUsu.Codigo
        Conn.Execute Sql
        Conn.Execute "DELETE FROM tmpimpbalan WHERE codusu = " & vUsu.Codigo
        
        Set RT = New ADODB.Recordset
        M3 = 1 'Sera el orden de insercion
        
      
        
       
        'Sera el numero de balance
        NumAsiento = NumBalan
        
        
        'Vemos si es de perdidas y ganacias
        Sql = DevuelveDesdeBD("perdidas", "balances", "numbalan", CStr(NumBalan), "N")
        EsBalancePerdidas_y_ganancias = (Val(Sql) = 1)
        
        
        
        'Vamos a utilizar la temporal de balances donde dejara los valores
        Sql = "DELETE from tmpbalancesumas where codusu= " & vUsu.Codigo
        Conn.Execute Sql
        
        Contabilidad = -1
        
        
        
        PB.Tag = -1
        
        CargaArbol 0, 0, -1, "", Mes1, Anyo1, Mes2, Anyo2, vContabilidad, False, PB, EsBalancePerdidas_y_ganancias, ImportesDeSoloUnMes
        
    
            
        If Not PB Is Nothing Then PB.visible = False
        
        'Cuando termina de cargar el arbol vamos calculando las sumas
        Sql = "SELECT * FROM "
        'Al ponerle Conta?.   lo k damos a entender es k lee la configuracion de su PROIPA sperdi
        If Contabilidad > 0 Then Sql = Sql & "ariconta" & Contabilidad & "."
        Sql = Sql & "balances_texto where numbalan=" & NumBalan & " AND tipo = 1"
        Sql = Sql & " ORDER BY orden"
        
        'Modificacion 12 Febrero 2004
        '----------------------------
        '  A igual numero de orden, ordena por creacion entonces da la casualidad de que
        ' muestra hace primero el BV del pasivo k el AiV
        Sql = Sql & ",Pasivo"
        
        
        RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RT.EOF
        
            
            
        
            CalculaSuma DBLet(RT!Formula), Val(RT!A_Cero) <> 0
            d = TransformaComasPuntos(CStr(ImpD))
            'UPDATEAMOS
            
            
            
            
            Sql = "UPDATE "
            If Contabilidad > 0 Then Sql = Sql & "ariconta" & Contabilidad & "."
            Sql = Sql & "tmpimpbalance SET importe1 =" & d
            If M2 > 0 Then
                H = TransformaComasPuntos(CStr(ImpH))
                Sql = Sql & ",importe2 = " & H
            End If
            Sql = Sql & " where codusu = " & vUsu.Codigo
            Sql = Sql & " AND Pasivo='" & RT!Pasivo & "' AND Codigo=" & RT!Codigo
            Conn.Execute Sql
            RT.MoveNext
        Wend
        RT.Close
    
    
    'QUITAR. Ya ha calculado los saldos todos juntitos
    '
    'If Contabilidad < 0 Then
        'Una vez todos los importes y demas vamos a generar los datos en impresion
        'Lo unico a tener en cuenta es k en las formulas si es menor k 0 no se imprime
        '-----------------------------------------------------------------------------
        
        Sql = "Select * from "
        If Contabilidad > 0 Then Sql = Sql & "ariconta" & Contabilidad & "."
        Sql = Sql & "balances_texto where numbalan=" & NumBalan & " AND padre"
        Aux = "INSERT INTO tmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,LibroCD,QueCuentas) VALUES (" & vUsu.Codigo
        Codigo = "Select importe1,importe2,quecuentas from "
        If Contabilidad > 0 Then Codigo = Codigo & "ariconta" & Contabilidad & "."
        Codigo = Codigo & "tmpimpbalance where codusu=" & vUsu.Codigo & " AND pasivo='"
        M1 = 1
        
        
        
        
        'Para que no imprima en la primera columna lo que seria la suma de "los hijos"
        EsPyGNoAbreviado = False
        If EsBalancePerdidas_y_ganancias Then

            AuxPyG = DevuelveDesdeBD("nombalan", "balances", "numbalan", CStr(NumBalan), "N")
            AuxPyG = UCase(AuxPyG)
            If InStr(1, AuxPyG, "ABREV") = 0 Then EsPyGNoAbreviado = True
        End If
        
        
        CargaArbolImpresion -1, "", 1, LibroCD, EsPyGNoAbreviado
        
    'End If
    
    Set RT = Nothing
End Function

Private Sub CargaArbol(ByRef vImporte As Currency, ByRef vimporte2 As Currency, Padre As Integer, Pasivo As String, ByRef Mes1 As Integer, ByRef Anyo1 As Integer, ByRef Mes2 As Integer, ByRef Anyo2 As Integer, ByRef Contabilidades As String, EsListado As Boolean, ByRef PrB_ As ProgressBar, EsPerdidasyGanancias As Boolean, ElImporteDeSoloUnMes As Boolean)
Dim Rs As ADODB.Recordset
Dim nodImporte As Currency
Dim MiAux As String
Dim OtroImporte As Currency
Dim OtroImporte2 As Currency
Dim QueCuentas As String
Dim NRegs As Long

'Nuevo PGC.  Puede ser que UN nodo raiz sea la SUMA , con lo cual pasan dos cosas:
'   .- 1: Puede que tenga nodos colgando, que habra que calcular
'           el importe que habra que pintarle sera LA de la formula
'   .- 2:  Un nodo raiz No es una formula.
'           el importe que pinmtara sera el de la suma
Dim Tipo As Integer

    If Padre < 0 Then
        MiAux = " is null" 'NODO RAIZ
    Else
        MiAux = " = " & Padre & " AND Pasivo = '" & Pasivo & "'"
    End If
    
    
    MiAux = "balances_texto where numbalan=" & NumAsiento & " AND padre" & MiAux
    MiAux = "Select * from " & MiAux
    
    If Not PrB_ Is Nothing Then
        
        If Padre < 0 Then
            
        
            NRegs = TotalRegistrosConsulta("Select codigo from balances_texto WHERE numbalan =" & NumAsiento)
            NRegs = NRegs + 2
            PrB_.Value = 0
            PrB_.Max = NRegs
            PrB_.Tag = NRegs
            PrB_.visible = True
        
            
    
       End If
    End If

            
    

    Set Rs = New ADODB.Recordset
    Rs.Open MiAux & " ORDER By Orden", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
    
         If Not PrB_ Is Nothing Then
            If PrB_.Value < PrB_.Max Then
                PrB_.Value = PrB_.Value + 1
            Else
                PrB_.Value = PrB_.Value - 1
            End If
            PrB_.Refresh
            DoEvents
        End If
       
    
        OtroImporte = 0
        OtroImporte2 = 0
        'If RS!Tipo = 1 Then St op
        If vParam.NuevoPlanContable Then
            Tipo = 0  'ASi seguro que entro
        Else
            Tipo = Rs!Tipo
        End If
        If Tipo = 0 Then
        
           
        
            If Rs!tienenctas = 1 Then
                QueCuentas = ""
                If EsListado Then
                    'LISTADITO
                    
                    QueCuentas = PonerCuentasBalances(Rs!Pasivo, Rs!Codigo)
                Else
                    'IMPORTES
                    OtroImporte = CalculaImporteCtas_(Rs!Pasivo, Rs!Codigo, Mes1, Anyo1, True, Contabilidades, EsPerdidasyGanancias, ElImporteDeSoloUnMes)
                
                    'Debug.Print Rs!Pasivo & Rs!Codigo & " " & OtroImporte
                
                
                    If Mes2 > 0 Then OtroImporte2 = CalculaImporteCtas_(Rs!Pasivo, Rs!Codigo, Mes2, Anyo2, False, Contabilidades, EsPerdidasyGanancias, ElImporteDeSoloUnMes)
                End If
            Else
                CargaArbol OtroImporte, OtroImporte2, Rs!Codigo, Rs!Pasivo, Mes1, Anyo1, Mes2, Anyo2, Contabilidades, EsListado, PrB_, EsPerdidasyGanancias, ElImporteDeSoloUnMes
                QueCuentas = ""
                
            End If
        Else
            QueCuentas = ""
        End If
        
        
        vImporte = vImporte + OtroImporte
        vimporte2 = vimporte2 + OtroImporte2
        
        

        'Insertamos la linea
        'en aux hay
        'codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,orden) VALUES (1
        MiAux = "'" & Rs!Pasivo & "'," & Rs!Codigo & ",'" & Rs!texlinea & "','"
        d = TransformaComasPuntos(CStr(OtroImporte))
        MiAux = MiAux & Rs!deslinea & "'," & d & ","
        d = TransformaComasPuntos(CStr(OtroImporte2))
        MiAux = MiAux & d & ",0," & M3 & ",'" & QueCuentas & "')"
        Aux = "INSERT INTO "
        'If Contabilidad > 0 Then Aux = Aux & "Conta" & Contabilidad & "."
        Aux = Aux & "tmpimpbalance (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,orden,quecuentas) VALUES (" & vUsu.Codigo & ","
        MiAux = Aux & MiAux
        Conn.Execute MiAux
    
        M3 = M3 + 1
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    
End Sub



Private Function CalculaImporteCtas_(Pasivo As String, Codigo As Integer, ByRef mess1 As Integer, ByRef anyos1 As Integer, Año1_o2 As Boolean, ByRef Contabilidades As String, EsPerdidasyGanancias As Boolean, ImporteSoloUnMes As Boolean) As Currency
Dim RT As ADODB.Recordset
Dim X As Integer
Dim Y As Integer
Dim vI1 As Currency
Dim vI2 As Currency
Dim QuitarUno As Boolean
Dim Contador As Integer
Dim ContaX As String


        If Contabilidades = "" Then
            ContaX = "-1|"
        Else
            ContaX = Contabilidades
        End If
        Set RT = New ADODB.Recordset
        CalculaImporteCtas_ = 0
        vI2 = 0
        Contador = 0
        'para cada contbilida
        While ContaX <> ""
            
            'Vemos cual ContaX
            X = InStr(1, ContaX, "|")
            Y = CInt(Mid(ContaX, 1, X - 1))
            ContaX = Mid(ContaX, X + 1)
            
            Contabilidad = Y
            
            'Fecha del utlimo en hco
            VFecha3 = CDate(Mid(VarConsolidado(0), (11 * Contador) + 1, 10))
                
            'Quitar1
            If Año1_o2 Then
                QuitarUno = (Mid(VarConsolidado(1), (Contador * 2) + 1, 1) = 1)
            Else
                QuitarUno = Mid(VarConsolidado(2), (Contador * 2) + 1, 1)
            End If
            
            EjerciciosCerrados = (CDate("15/" & mess1 & "/" & anyos1) < VFecha3)
            vI1 = CalculaImporteCtas1Contabilidad(Pasivo, Codigo, mess1, anyos1, QuitarUno, EsPerdidasyGanancias, ImporteSoloUnMes)
        

        
            vI2 = vI2 + vI1

            Contador = Contador + 1
    Wend
    Set RT = Nothing
    Contabilidad = -1
    CalculaImporteCtas_ = vI2
End Function






Private Function CalculaImporteCtas1Contabilidad(Pasivo As String, Codigo As Integer, ByRef mess1 As Integer, ByRef anyos1 As Integer, QuitarSaldos As Boolean, EsPerdidasyGanancias As Boolean, SoloDatosUnMes As Boolean) As Currency
Dim RC As ADODB.Recordset
Dim F1 As Date
Dim F2 As Date
Dim I1 As Currency
Dim B1 As Byte

Dim ColImportes As Collection
Dim RN As ADODB.Recordset
Dim cad As String
Dim CalculoPyG As Boolean

    Set RC = New ADODB.Recordset
        
    'Vamos a calcular el importe para cada cuenta, para cada contbiliadad
        
    vCta = "SELECT * from "
    vCta = vCta & "balances_ctas WHERE pasivo ='" & Pasivo & "' AND codigo = " & Codigo & " AND numbalan = " & NumAsiento
    RC.Open vCta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Primer mes
    If Not SoloDatosUnMes Then
    
        
        If mess1 < Month(vParam.fechaini) Then
            
                'Lo que hacia
                F1 = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & anyos1 - 1)
                
                'Esto lo añado el 15/07/2020
                If F1 > vParam.fechafin Then
                    If Not EsPerdidasyGanancias Then F1 = vParam.fechaini
                End If
    
        Else
        
    
                'Lo que hacia
                F1 = CDate(Day(vParam.fechaini) & "/" & Month(vParam.fechaini) & "/" & anyos1)
                'CUANDO EL BALANCE SE pide sobre el año "siguiente" los saldos se cogen desde el inicio de ejercicio ACTUAL, menos en perdidas y ganancias
                If F1 > vParam.fechafin Then
                    If Not EsPerdidasyGanancias Then F1 = vParam.fechaini
                End If
        
        End If
    Else
        F1 = "01/" & Format(mess1, "00") & "/" & anyos1   'mes slicitado
    End If
    
    'Debug.Assert Not (F1 <> "01/10/2018")
    F2 = CDate(DiasMes(CInt(mess1), anyos1) & "/" & mess1 & "/" & anyos1)
    I1 = 0
    EjerciciosCerrados = F1 < VFecha3
    
    B1 = 0
    If Not SoloDatosUnMes Then
        'Lo que hacia
        If QuitarSaldos Then B1 = 1             'Ambos, pyg y cierr
    End If
    
    Set RN = New ADODB.Recordset
    While Not RC.EOF
                                                                                                            
                                                                                                            
                                                           
                                                                                                            
                                                                                                            
                                                                                                            
            cad = "SELECT substring(line.codmacta,1," & Len(RC!codmacta)
            cad = cad & ") as codmacta,'' nommacta ,   year(fechaent) anyo,month(fechaent) mes,"
            cad = cad & " sum(coalesce(timported,0)) debe, sum(coalesce(timporteh,0)) haber FROM "
            If Contabilidad >= 0 Then cad = cad & "ariconta" & Contabilidad & "."
            cad = cad & "hlinapu"
            cad = cad & " as line "
            cad = cad & " WHERE codmacta like '" & Trim(RC!codmacta) & "%'"
            cad = cad & " AND fechaent between " & DBSet(F1, "F") & " AND " & DBSet(F2, "F")
            cad = cad & " GROUP BY 1,anyo,mes "
            cad = cad & " ORDER By 1 ,anyo,mes"

            Set ColImportes = Nothing
            Set ColImportes = New Collection
            RN.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RN.EOF
                cad = RN!Anyo & Format(RN!Mes, "00") & "|" & RN!Debe & "|" & RN!Haber & "|"
                ColImportes.Add cad
                RN.MoveNext
            Wend
            RN.Close
            If ColImportes.Count = 0 Then
                'CEROS
                ImpH = 0
                ImpD = 0
                ImCierrH = 0
                ImCierrD = 0
            Else
                'Antes en resetea 6y7 habia un true
                
                CargaBalanceNuevaContabilidad RC!codmacta, "", False, F1, F2, F1, F2, EjerciciosCerrados, B1, Contabilidad, True, True, False, ColImportes
            End If

            
            
            
         'FALTA:   en ejercicicos cerrados tambien debe quitar 6 y 7 si lo solicita, pero hasta el mes selecionado

         If RC!codmacta = Mid(vParam.ctaperga, 1, 3) Then Stop
         
        'Balance de situacion. Tratar cta perdidas y ganancias
        If Not EsBalancePerdidas_y_ganancias And RC!codmacta = Mid(vParam.ctaperga, 1, 3) Then
            
                ImCierrH = 0
                ImCierrD = 0
                If Len(vCta) > 5 Then vCta = RC!codmacta   'Esta situacion se da porque NO ha tenido pyg ni movimientos
                
                
                CalculoPyG = False
                
                If F1 < vParam.fechaini Then
                    If Month(F2) = Month(vParam.fechafin) Then CalculoPyG = True
                End If
                
                If CalculoPyG Then
                    'Ejercicios cerrados
                    ObtenerPerdidasyGanancias EjerciciosCerrados, F1, F2, 3    'SOLO LE QUITO EL CIERRE
                Else
                
                    If Saldo6y7en129 Then
                         'Obtenemos las pyg del ejercicio
                         OntenerPyGActual F1, F2
                         
                    End If
                End If
                


                ImpH = ImpH - ImCierrH
                ImpD = ImpD - ImCierrD
                
                
                
        End If

        
        'NUEVO NUEVO
        '-----------
        
        'If RC!Codmacta = "640" Then St op
        If vParam.NuevoPlanContable Then
        
            If EsBalancePerdidas_y_ganancias Then
                'If Mid(RC!Codmacta, 1, 1) = "6" Then
                '    ImpH = ImpD - ImpH
                'Else
                    ImpH = ImpH - ImpD
                'End If
            Else
                ImpH = ImpD - ImpH   'Como estaba
                If Pasivo = "B" Then ImpH = -1 * ImpH
            End If
        Else
            ImpH = ImpD - ImpH
            If Pasivo = "B" Then ImpH = -1 * ImpH
        End If


        If RC!TipSaldo <> "S" Then
            'Debug.Print RC!codmacta
            'St op
        End If
        
        Select Case RC!TipSaldo
        Case "D"
            'Y la cuenta es de haber pongo a 0
            If ImpH < 0 Then ImpH = 0
        Case "H"
           If ImpH < 0 Then ImpH = 0
        Case "S"

 
        End Select

        
        If RC!Resta = 1 Then ImpH = ImpH * -1
        I1 = I1 + ImpH
        
        
        'Siguiente
        RC.MoveNext
    Wend
    RC.Close
    CalculaImporteCtas1Contabilidad = I1
    Set RC = Nothing
End Function






'Obteiene la diferencia de la 6y7 para pintarla en el balnace de situacion

Private Sub OntenerPyGActual(ByRef fec1 As Date, ByRef fec2 As Date)
Dim cad As String
Dim RT As ADODB.Recordset
    Set RT = New ADODB.Recordset
    cad = "select sum(if(timported is null,0,timported)),sum(if(timporteh is null,0,timporteh)  ) FROM "
    If Contabilidad >= 0 Then cad = cad & " ariconta" & Contabilidad & "."
    cad = cad & "hlinapu where substring(codmacta,1,1) IN ('6','7') and fechaent >=" & DBSet(fec1, "F") & " and fechaent <=" & DBSet(fec2, "F")
    RT.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        ImCierrH = DBLet(RT.Fields(0), "N")
        ImCierrD = DBLet(RT.Fields(1), "N")
    End If
    

End Sub

Private Function PonerCuentasBalances(Pasivo As String, Codigo As Integer) As String
Dim RC As ADODB.Recordset


    Set RC = New ADODB.Recordset
        
    'Vamos a calcular el importe para cada cuenta, para cada contbiliadad
        
    vCta = "SELECT * from "
   ' If Contabilidad > 0 Then vCta = vCta & "Conta" & Contabilidad & "."
    vCta = vCta & "balances_ctas WHERE pasivo ='" & Pasivo & "' AND codigo = " & Codigo & " AND numbalan = " & NumAsiento
    RC.Open vCta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
  
    vCta = ""
    While Not RC.EOF
        If vCta <> "" Then vCta = vCta & ","
        If RC!Resta = 1 Then vCta = vCta & "-"
        vCta = vCta & RC!codmacta & " "

        'Siguiente
        RC.MoveNext
    Wend
    RC.Close
    PonerCuentasBalances = CStr(vCta)
    Set RC = Nothing
End Function





Private Sub CalculaSuma(CadenaSuma As String, A_Cero As Boolean)
Dim RA As ADODB.Recordset
    CadenaSuma = Trim(CadenaSuma)
    ImpD = 0
    ImpH = 0
    If CadenaSuma = "" Then Exit Sub
        
    
    'Quitamos todos los blancos
    Do
        A1 = InStr(1, CadenaSuma, " ")
        If A1 > 0 Then CadenaSuma = Mid(CadenaSuma, 1, A1 - 1) & Mid(CadenaSuma, A1 + 1)
    Loop Until A1 = 0
                  
    Aux = Mid(CadenaSuma, 1, 1)
    If Aux <> "+" And Aux <> "-" Then CadenaSuma = "+" & CadenaSuma
    
    Set RA = New ADODB.Recordset
    'Dejamos medio montado el sql
    ' "INSERT INTO tmpimpbalance (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita,orden) VALUES (" & vUsu.Codigo & ","
    Sql = "Select importe1,importe2 from "
    If Contabilidad > 0 Then Sql = Sql & "ariconta" & Contabilidad & "."
    Sql = Sql & "tmpimpbalance where codusu =" & vUsu.Codigo

    
    'Iremos deglosando cadenasuma
    ImAcD = 0
    ImAcH = 0
    Do
        'Empezamos en dos, pq lo primero es siempre un mas o un menos
        A1 = InStr(2, CadenaSuma, "+")
        A2 = InStr(2, CadenaSuma, "-")
        If A1 = 0 And A2 = 0 Then
            'Ya no hay mas para procesar
            A3 = 0
            Aux = CadenaSuma
            CadenaSuma = ""
            M1 = 1
        Else
            If A1 = 0 Then A1 = 32000
            If A2 = 0 Then A2 = 32000
            If A1 > A2 Then
                A3 = A2
            Else
                A3 = A1
            End If
            Aux = Mid(CadenaSuma, 1, A3 - 1)
            CadenaSuma = Mid(CadenaSuma, A3)
        End If
        
        'El signo
        If Mid(Aux, 1, 1) = "-" Then
            M1 = -1
        Else
            M1 = 1
        End If
        
        'La letra del pasivo / activo
        vCta = " AND Pasivo = '" & Mid(Aux, 2, 1)
        
        'El codigo del campo
        d = Mid(Aux, 3)
        Codigo = "-2"
        If d <> "" Then
            If IsNumeric(d) Then Codigo = d
        End If
        
        'SQL para la BD
        vCta = vCta & "' AND codigo =" & d
        
        RA.Open Sql & vCta, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ImCierrD = 0
        ImCierrH = 0
        If Not RA.EOF Then
            If Not IsNull(RA.Fields(0)) Then ImCierrD = RA.Fields(0)
            If Not IsNull(RA.Fields(1)) Then ImCierrH = RA.Fields(1)
        End If
        RA.Close
        ImAcD = ImAcD + (M1 * ImCierrD)
        ImAcH = ImAcH + (M1 * ImCierrH)
    Loop Until CadenaSuma = ""
    'Ponemos a cero si asi lo dice la funcion
    If A_Cero Then
        If ImAcD < 0 Then ImAcD = 0
        If ImAcH < 0 Then ImAcH = 0
    End If
    ImpD = ImAcD
    ImpH = ImAcH
    Set RA = Nothing
    
End Sub



Private Sub CargaArbolImpresion(Padre As Integer, Pasivo As String, Nivel As Byte, vLibroCD As Boolean, EsBalancePyGNOabreviado As Boolean)
Dim Rs As ADODB.Recordset
Dim MiAux As String
Dim TieneHijos As Boolean
Dim QueCuentas As String

    
    If Padre < 0 Then
        MiAux = " is null" 'NODO RAIZ
    Else
        MiAux = " = " & Padre & " AND Pasivo = '" & Pasivo & "'"
    End If
    MiAux = Sql & MiAux & " ORDER By Pasivo, Orden"
  
    Set Rs = New ADODB.Recordset
    Rs.Open MiAux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
         
        
        If vParam.NuevoPlanContable Then
            'Nueva contabilidad
            TieneHijos = True   'Existe la posibilidad (de hecho lo hace, que teniendo hijos sea una FORMULA
        
        
        
        Else
            'Antiguo plan
            TieneHijos = False
            If Rs!Tipo = 0 Then
                If Rs!tienenctas = 0 Then TieneHijos = True
            End If
        End If
        
        'Obtenmos el importe
        MiAux = Codigo & Rs!Pasivo & "' AND codigo =" & Rs!Codigo
        RT.Open MiAux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ImpD = 0: ImpH = 0: QueCuentas = ""
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImpD = RT.Fields(0)
            If Not IsNull(RT.Fields(1)) Then ImpH = RT.Fields(1)
            QueCuentas = DBLet(RT.Fields(2))
        End If
        RT.Close
        
        If vParam.NuevoPlanContable And Padre < 0 And Rs!Tipo = 0 Then
            'Marzo 2011
            'Para los pyG NO abreviados
            If EsBalancePyGNOabreviado Then
            'If RS!NumBalan = 2 Then
                'FALTA### comprobar que es para todos.    De momento para PyG normal
                'Aqui no pintaremos el resultado de la suma de subnodos si el padre es null
                ' y NO es una formula
                ImpD = 0: ImpH = 0:
            End If
        End If
        
        
        
        'Si no se pinta si el resultado es negativo entonces entonces
        If Rs!Pintar = 0 Then   'PINTAR: SI siempre NO. Ngativos no
            If ImpD < 0 Then ImpD = 0
            If ImpH < 0 Then ImpH = 0
        End If
        
        
            'Insertamos la linea
            'AUX tiene:
            'INSERT INTO usuari.ztmpimpbalan (codusu, Pasivo, codigo, descripcion, linea, importe1, importe2, negrita) VALUES (" & vUsu.Codigo
            MiAux = Aux & ",'" & Rs!Pasivo & "'," & M1 & ",'" & DBLet(Rs!texlinea) & "','"
            'El sangrado para el texto
            MiAux = MiAux & Space((Val(Nivel) - 1) * 4)  'SANGRIA TABULADO PARRAFO
            MiAux = MiAux & Rs!deslinea & "'" & ImporteASQL(ImpD) & ImporteASQL(ImpH)
            MiAux = MiAux & "," & Rs!negrita & ","
            If vLibroCD Then
                MiAux = MiAux & "'" & DBLet(Rs!LibroCD) & "'"
            Else
                MiAux = MiAux & "NULL"
            End If
            MiAux = MiAux & ",'" & QueCuentas & "')"
            'MiAux = MiAux & ")"
            Conn.Execute MiAux
    
    
        M1 = M1 + 1
        
        'Ahora,si tiene hijos cargamos el subarbol
        
        If TieneHijos Then CargaArbolImpresion Rs!Codigo, Rs!Pasivo, Nivel + 1, vLibroCD, EsBalancePyGNOabreviado
        
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub




Private Function HayAsientoCierreBalances(Mes As Byte, Anyo As Integer) As Boolean
Dim C As String
Dim Rs As Recordset
    HayAsientoCierreBalances = False
    'C = "01/" & CStr(Me.cmbFecha(1).ListIndex + 1) & "/" & txtAno(1).Text
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierreBalances = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            Set Rs = New ADODB.Recordset
            C = "Select count(*) from "
            If Contabilidad > 0 Then C = C & "ariconta" & Contabilidad & "."
            C = C & "hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If Not IsNull(Rs.Fields(0)) Then
                    If Rs.Fields(0) > 0 Then HayAsientoCierreBalances = True
                End If
            End If
            Rs.Close
            Set Rs = Nothing
        End If
    End If
End Function










'--------------------------------------------------------
'Listado evolucion de saldos, menusal

Public Sub FijarValoresEvolucionMensualSaldos(fec1 As Date, fec2 As Date)
    vFecha1 = fec1
    vFecha2 = fec2
    Set RT = Nothing
    Aux = "INSERT INTO tmpconext (codusu, cta,  Pos, fechaent, timporteD, timporteH, saldo) VALUES (" & vUsu.Codigo & ",'"
    Contabilidad = -1
    A3 = Year(vFecha2)
    
End Sub

Public Function DatosEvolucionMensualSaldos(ByRef Cuenta As String, ByRef DescCuenta As String, vSql As String, MostrarTodosMeses As Boolean, EsEnHlinapu1 As Boolean, QuitarCierre As Boolean, Optional FechaInicio As Date, Optional Tipo As Integer, Optional Acumular As Boolean) As Byte
Dim NuloApertura As Boolean
Dim HacerAñoAnterior As Boolean
Dim importeCierreD As Currency
Dim importeCierreH As Currency

    vCta = Cuenta
    ObtenerApertura False, vFecha1, vFecha2, NuloApertura
        
    If QuitarCierre Then
        ObtenerPerdidasyGanancias False, vFecha1, vFecha2, 1
        importeCierreD = ImCierrD 'los guardo aqui, pq luego estas variables las reutilizo
        importeCierreH = ImCierrH
    End If
    Sql = "INSERT INTO tmpconextcab (codusu, cuenta, cta, acumantD, acumantH, acumantT ) VALUES ("
    Sql = Sql & vUsu.Codigo & ",'" & DevNombreSQL(DescCuenta) & "','" & vCta & "',"
    If NuloApertura Then
        Sql = Sql & "0,0,0)"
        ImAcD = 0: ImAcH = 0: ImPerD = 0
    Else
        ImAcD = ImpD
        ImAcH = ImpH
        Sql = Sql & TransformaComasPuntos(CStr(ImpD)) & "," & TransformaComasPuntos(CStr(ImpH)) & ","
        ImPerD = ImpD - ImpH
        Sql = Sql & TransformaComasPuntos(CStr(ImPerD)) & ")"
        
    End If
    Conn.Execute Sql
    
    
    ' Nueva tabla para el informe apaisado de meses
    Dim K As Integer
    Dim K1 As Integer
    Dim Anyo As Integer
    Dim Mes As Integer

    Sql = "insert into tmpevolsal (codusu, codmacta, nommacta, apertura, mes1, mes2, mes3, mes4, mes5, mes6, mes7, mes8, mes9, mes10, mes11, mes12, "
    Sql = Sql & "importemes1, importemes2, importemes3, importemes4, importemes5, importemes6, importemes7, importemes8, importemes9, importemes10, importemes11, importemes12) values ("
    Sql = Sql & vUsu.Codigo & "," & DBSet(vCta, "T") & ",'" & DescCuenta & "',"

    Select Case Tipo
        Case 0
            Sql = Sql & DBSet(ImAcD, "N") & ","
        Case 1
            Sql = Sql & DBSet(ImAcH, "N") & ","
        Case 2
            Sql = Sql & DBSet(ImPerD, "N") & ","
    End Select


    If Year(vParam.fechaini) = Year(vParam.fechafin) Then ' año natural
        For K = 1 To 12
            Sql = Sql & Format(Year(FechaInicio), "0000") & Format(K, "00") & ","
        Next K
    Else
        Anyo = Year(FechaInicio)
        For K = 1 To 12
            K1 = Month(FechaInicio) - 1 + K
            If K1 > 12 Then
                K1 = K1 - 12
                Anyo = Year(FechaInicio) + 1
            End If
            'Mes = Format(Anyo, "0000") & Format(K1, "00")
            Sql = Sql & DBSet(Format(Anyo, "0000") & Format(K1, "00"), "N") & ","
        Next K
        
    End If
    Sql = Sql & "0,0,0,0,0,0,0,0,0,0,0,0)"

    Conn.Execute Sql

    
    Sql = "Select year(fechaent) anopsald, month(fechaent) mespsald, sum(coalesce(timported,0)) impmesde, sum(coalesce(timporteh,0)) impmesha from hlinapu where "
    If EsCuentaUltimoNivel(vCta) Then
        Sql = Sql & "codmacta = '" & vCta & "'"
    Else
        Sql = Sql & "codmacta like '" & vCta & "%' "
    
    End If
    'Las fechas
    Sql = Sql & " AND fechaent between " & DBSet(vFecha1, "F") & " AND " & DBSet(vFecha2, "F")
    Sql = Sql & " GROUP BY year(fechaent), month(fechaent) "
    Sql = Sql & " ORDER by year(fechaent),month(fechaent) "
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    M1 = Month(vFecha1)
    A1 = Year(vFecha1)
    ImCierrD = 0: ImCierrH = 0
    NumAsiento = 0
    While Not RT.EOF
    
        A2 = RT!anopsald
        M2 = RT!mespsald
                    
        If MostrarTodosMeses Then
            If M2 <> M1 Then
                ImpD = 0
                ImpH = 0
                  
            
                If Year(vFecha1) = Year(vFecha2) Then
                    'Se ha saltado algun(os) mes(es)
                    'Los rellenaremos ?????
        
                    
                        
                    For M3 = M1 To M2 - 1
                        
                        VFecha3 = CDate("01/" & M3 & "/" & Year(vFecha1))
                        InsertaLineaEvolucion Tipo, Acumular
                    Next M3
                    M1 = M2
                                
                
                
                Else
                
                    If A1 = A2 Then
                                       
                                         'El ultimo mes en meterse fue el anterior
                        For M3 = M1 To M2 - 1
                            VFecha3 = CDate("01/" & M3 & "/" & A1)
                            InsertaLineaEvolucion Tipo, Acumular
                        Next M3
                        
                        
                    Else
                        For M3 = M1 To 12
                            VFecha3 = CDate("01/" & M3 & "/" & A1)
                            InsertaLineaEvolucion Tipo, Acumular
                        Next M3
                        
                        A1 = A2 'cambiamos año
                        For M3 = 1 To M2 - 1
                            VFecha3 = CDate("01/" & M3 & "/" & A1)
                            InsertaLineaEvolucion Tipo, Acumular
                        Next M3
                        
                    End If
                    M1 = M2
                
                End If
            End If
        End If
        ImpD = RT!impmesde
        ImpH = RT!impmesha
        
        VFecha3 = CDate("01/" & RT!mespsald & "/" & RT!anopsald)
        If VFecha3 = vFecha1 Then
            'Hay que quitar los saldos de apertura
            ImpD = ImpD - ImAcD
            ImpH = ImpH - ImAcH
            If ImpD <> 0 Or ImpH <> 0 Then NumAsiento = 1
        Else
            If QuitarCierre Then
                If Format(VFecha3, "mmyyyy") = Format(vFecha2, "mmyyyy") Then
                    'Cierre
                    ImpD = ImpD - importeCierreD
                    ImpH = ImpH - importeCierreH
                    If ImpD <> 0 Or ImpH <> 0 Then NumAsiento = 1
                Else
                    NumAsiento = 1
                End If
            Else
                NumAsiento = 1
            End If
        End If
        
        ImCierrD = ImCierrD + ImpD
        ImCierrH = ImCierrH + ImpH
        
        
        If MostrarTodosMeses Then
            M3 = 1
        Else
            If ImpD = 0 And ImpH = 0 Then
                M3 = 0
            Else
                M3 = 1
            End If
        End If
        
        If M3 = 1 Then InsertaLineaEvolucion Tipo, Acumular
        
        M1 = M1 + 1
        If Year(vFecha1) <> Year(vFecha2) Then   'Si años partidos
            If M1 > 12 Then
                M1 = 1   'Ponemos el mes a 1 otra vez
                A1 = Year(vFecha2)
            End If
        End If
                
        
                
        RT.MoveNext
    Wend
    RT.Close
  
    If NumAsiento = 0 Then
        Sql = " WHERE codusu =" & vUsu.Codigo & " AND cta='" & vCta & "'"
        Conn.Execute "DELETE FROM tmpconext" & Sql
        Conn.Execute "DELETE FROM tmpconextcab" & Sql
        Conn.Execute "DELETE FROM tmpevolsal where codusu = " & vUsu.Codigo & " and codmacta = " & DBSet(vCta, "N")
        Exit Function
    End If
    
    
    If MostrarTodosMeses Then
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
        
        
        
            If M1 <= 12 Then
                'Se ha saltado algun(os) mes(es)
                'Los rellenaremos ?????
    
                ImpD = 0
                ImpH = 0
            
                'Año natural
                For M3 = M1 To 12
                    VFecha3 = CDate("01/" & M3 & "/" & A2)
                    InsertaLineaEvolucion Tipo, Acumular
                Next M3
            End If
        Else
        
            'Años tipo cooperativas
            
            ImpD = 0
            ImpH = 0
            
            'Veremos donde se ha quedado, si en la mitad del año primero o en el segundo
            If Year(vFecha2) <> A1 Then
            
                'Se ha quedado en la primera parte de los años
                'Si el mes donde se ha quedado es el ultimo
             
                    For M3 = M1 To 12
                        VFecha3 = CDate("01/" & M3 & "/" & A2)
                        InsertaLineaEvolucion Tipo, Acumular
                    Next M3
                    M1 = 1
                    M2 = Month(vFecha2)
            Else
                
                M2 = Month(vFecha2)
            End If
            
            'OK Hay que rellenar
            If M1 <= M2 Then
                'Rellenamos primero el año1
                
                For M3 = M1 To M2
                    VFecha3 = CDate("01/" & M3 & "/" & Year(vFecha2))
                    InsertaLineaEvolucion Tipo, Acumular
                Next M3

            End If
            
            

        End If
    End If
    
    'Updateo el total
    'ImPerD = ImPerD + (ImCierrD - ImCierrH)
    ImCierrD = ImCierrD + ImAcD
    ImCierrH = ImCierrH + ImAcH
    Sql = "UPDATE tmpconextcab SET acumtotD=" & TransformaComasPuntos(CStr(ImCierrD))
    Sql = Sql & " , acumtotH=" & TransformaComasPuntos(CStr(ImCierrH))
    Sql = Sql & " , acumtotT=" & TransformaComasPuntos(CStr(ImPerD))
    Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " AND cta='" & vCta & "'"
    Conn.Execute Sql
End Function


Private Sub InsertaLineaEvolucion(Tipo As Integer, Acumulado As Boolean)
Dim Importe As Currency
Dim vAux As Long

    'Aux = "INSERT INTO ztmpconext (codusu, cta,  Pos,numdiari, fechaent,
    'timporteD, timporteH, saldo) VALUES ("
    Codigo = Aux & vCta & "'," & 1 & ",'" & Format(VFecha3, FormatoFecha) & "',"
    Codigo = Codigo & TransformaComasPuntos(CStr(ImpD)) & "," & TransformaComasPuntos(CStr(ImpH)) & ","
    ImPerD = ImPerD + (ImpD - ImpH)
    Codigo = Codigo & TransformaComasPuntos(CStr(ImPerD)) & ")"
    Conn.Execute Codigo
    
    
    
    
    Dim Sql2 As String
    Dim Rs As ADODB.Recordset
    
    Sql2 = "select mes1, mes2, mes3, mes4, mes5, mes6, mes7, mes8, mes9, mes10, mes11, mes12 from tmpevolsal where codusu = " & vUsu.Codigo
    Sql2 = Sql2 & " and codmacta = " & DBSet(vCta, "T")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Select Case Tipo
        Case 0
            Importe = ImpD
        Case 1
            Importe = ImpH
        Case 2
            Importe = ImPerD 'ImpD - ImpH
    End Select
    
    If Not Acumulado Then ImPerD = 0
    
    
    If Not Rs.EOF Then
        Sql2 = "update tmpevolsal set "
        vAux = Format(Year(VFecha3), "0000") & Format(Month(VFecha3), "00")
        Select Case vAux
            Case DBLet(Rs!Mes1, "N")
                Sql2 = Sql2 & " importemes1 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes2, "N")
                Sql2 = Sql2 & " importemes2 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes3, "N")
                Sql2 = Sql2 & " importemes3 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes4, "N")
                Sql2 = Sql2 & " importemes4 = " & DBSet(Importe, "N")
            Case DBLet(Rs!mes5, "N")
                Sql2 = Sql2 & " importemes5 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes6, "N")
                Sql2 = Sql2 & " importemes6 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes7, "N")
                Sql2 = Sql2 & " importemes7 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes8, "N")
                Sql2 = Sql2 & " importemes8 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes9, "N")
                Sql2 = Sql2 & " importemes9 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes10, "N")
                Sql2 = Sql2 & " importemes10 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes11, "N")
                Sql2 = Sql2 & " importemes11 = " & DBSet(Importe, "N")
            Case DBLet(Rs!Mes12, "N")
                Sql2 = Sql2 & " importemes12 = " & DBSet(Importe, "N")
        End Select

        Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
        Sql2 = Sql2 & " and codmacta = " & DBSet(vCta, "T")

        Conn.Execute Sql2
    End If
    Set Rs = Nothing
    
    
End Sub


'-------------------------------------------------------------
Public Function BorrarCuenta(Cuenta As String, ByRef L1 As Label) As String
On Error GoTo Salida
Dim Sql As String
Dim Rs As ADODB.Recordset

'Con ls tablas declarads sin el ON DELETE , no dejara borrar
BorrarCuenta = "Error procesando datos"
Set Rs = New ADODB.Recordset



'Nuevo 15 Noviembre 2006.
'Comprobare a mano hlinapu tanto en cta como en contrapr
'Lo hare com TieneDatosSQLCount que utiliza el count
L1.Caption = "Historicos"
L1.Refresh
Sql = "SELECT count(*) from hlinapu where codmacta = '" & Cuenta & "'"
If TieneDatosSQLCount(Rs, Sql, 0) Then
    BorrarCuenta = "Cuenta en historico de apuntes"
    GoTo Salida
End If


Sql = "SELECT count(*) from hlinapu where ctacontr ='" & Cuenta & "'"
If TieneDatosSQLCount(Rs, Sql, 0) Then
    BorrarCuenta = "Contrapartida en historico de apuntes"
    GoTo Salida
End If


'lineas de apuntes, contrapartidads   -->1
Sql = "Select * from asipre_lineas where ctacontr ='" & Cuenta & "'"
If TieneDatosSQL(Rs, Sql) Then
    BorrarCuenta = "Contrapartida en asientos predefinidos"
    GoTo Salida
End If

' cuenta de contrapartida habitual   -->1
Sql = "Select * from cuentas where codcontrhab ='" & Cuenta & "'"
If TieneDatosSQL(Rs, Sql) Then
    BorrarCuenta = "Cuenta Contrapartida habitual en Cuentas"
    GoTo Salida
End If


'Cerrados
'----------------------------------




L1.Caption = "Otras tablas"
L1.Refresh

'-->3
'Otras tablas
'Reparto de gastos para inmovilizado
Sql = "Select codmacta2 from inmovele_rep where codmacta2='" & Cuenta & "'"
If TieneDatosSQL(Rs, Sql) Then
    BorrarCuenta = "Reparto de gastos para inmovilizado"
    GoTo Salida
End If

'-->4
Sql = "Select * from presupuestos where codmacta ='" & Cuenta & "'"
If TieneDatosSQL(Rs, Sql) Then
    BorrarCuenta = "Presupuestos"
    GoTo Salida
End If




'-->5    Referencias a ctas desde eltos de inmovilizado
Sql = "select codinmov from inmovele where codmact1='" & Cuenta & "'"
Sql = Sql & " or codmact2='" & Cuenta & "'"
Sql = Sql & " or codmact3='" & Cuenta & "'"
Sql = Sql & " or codprove='" & Cuenta & "'"
If TieneDatosSQL(Rs, Sql) Then
    BorrarCuenta = "Elementos de inmovilizado"
    GoTo Salida
End If





'-->6    Referencias a ctas desde eltos de inmovilizado
Sql = "select codiva from paramamort where codiva='" & Cuenta & "'"
If TieneDatosSQL(Rs, Sql) Then
    BorrarCuenta = "IVA en elmentos de inmovilizado"
    GoTo Salida
End If

'Cta bancaria
    Sql = "select codmacta from bancos where codmacta='" & Cuenta & "'"
    If TieneDatosSQL(Rs, Sql) Then
        BorrarCuenta = "Asociado a cuenta bancaria."
        GoTo Salida
    End If
    
DoEvents
If vEmpresa.TieneTesoreria Then
    L1.Caption = "Tesoreria"
    L1.Refresh
    
    'Habra k buscar en las tablas de tesoreria, k no esten enlazadas
    ' con FOREING KEY
'    SQL = "Select codmacta from scaja where codmacta = '" & Cuenta & "'"
'    If TieneDatosSQL(RS, SQL) Then
'        BorrarCuenta = "TESORERIA: Cuenta de caja."
'        GoTo Salida
'    End If

        
'    SQL = "select Ctacaja from susucaja where ctacaja='" & Cuenta & "'"
'    If TieneDatosSQL(Rs, SQL) Then
'        BorrarCuenta = "TESORERIA: Usuarios - caja."
'        GoTo Salida
'    End If

    Sql = "select codmacta from Departamentos where codmacta='" & Cuenta & "'"
    If TieneDatosSQL(Rs, Sql) Then
        BorrarCuenta = "TESORERIA: Departamentos."
        GoTo Salida
    End If
            
    Sql = "select codmacta from cobros where codmacta='" & Cuenta & "'"
    If TieneDatosSQL(Rs, Sql) Then
        BorrarCuenta = "TESORERIA: Cobros."
        GoTo Salida
    End If
    
    Sql = "select codmacta  from pagos where codmacta ='" & Cuenta & "'"
    If TieneDatosSQL(Rs, Sql) Then
        BorrarCuenta = "TESORERIA: Pagos."
        GoTo Salida
    End If
    
    
    Sql = "select ctaingreso from bancos where ctaingreso='" & Cuenta & "'"
    If TieneDatosSQL(Rs, Sql) Then
        BorrarCuenta = "TESORERIA: Pagos. ctaingreso"
        GoTo Salida
    End If
    
        
            
    'Contrapartida de gastosfijos
    Sql = "select contrapar from gastosfijos where contrapar='" & Cuenta & "'"
    If TieneDatosSQL(Rs, Sql) Then
        BorrarCuenta = "TESORERIA: Gastos fijos. contrapar"
        GoTo Salida
    End If
    
    
    Sql = "select ctagastos from bancos where ctagastos='" & Cuenta & "'"
    If TieneDatosSQL(Rs, Sql) Then
        BorrarCuenta = "TESORERIA: Cuenta bancaria. Cta gastos."
        GoTo Salida
    End If
    
End If




'SI kkega aqui es k ha ido bien
BorrarCuenta = ""
Salida:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar ctas." & Err.Description
    Set Rs = Nothing
 
End Function


'le pasamos el SQL y vemos si tiene algun dato
Private Function TieneDatosSQL(ByRef Rs As ADODB.Recordset, vSql As String) As Boolean
    TieneDatosSQL = False
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then TieneDatosSQL = True
    Rs.Close

End Function


Private Function TieneDatosSQLCount(ByRef Rs As ADODB.Recordset, vSql As String, IndexdelCount As Integer) As Boolean
    TieneDatosSQLCount = False
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(IndexdelCount)) Then If Rs.Fields(IndexdelCount) > 0 Then TieneDatosSQLCount = True
    End If
        
    Rs.Close

End Function

'-----------------------------------------------------------------------
'
'   I N F O R M E S        C R I S T A L
Public Function DevNombreInformeCrystal(QueInforme As Integer) As String

    DevNombreInformeCrystal = DevuelveDesdeBD("informe", "scryst", "codigo", CStr(QueInforme), "N")
    If DevNombreInformeCrystal = "" Then
        MsgBox "Opcion NO encontrada: " & QueInforme, vbExclamation
        DevNombreInformeCrystal = "ERROR"
    End If

End Function






'------------------------------------------------------------------
' BALANCE INICIO EJERCICIO
'
'   Es un balance que tiene todos los saldos de las cuentas a fecha
' inicio de ejerecicio. Es decir, tiene la apertura mas todos los apuntes que
' se hayan introducido con esa fecha
'
' Niveles:  Sera un string con los niveles del balance
Public Function CargaBalanceInicioEjercicio(Niveles As String, FechaInicioEjercicioSolicitado) As Boolean


On Error GoTo ECargaBalanceInicioEjercicio
    CargaBalanceInicioEjercicio = False
    
    ImCierrH = 0: ImCierrD = 0: ImPerD = 0: ImPerH = 0
    vCta = "INSERT INTO tmpbalancesumas (codusu,cta, nomcta, aperturaD, aperturaH, acumAntD, acumAntH, acumPerD, acumPerH, TotalD, TotalH) "
    Sql = "select " & vUsu.Codigo & ",hlinapu.codmacta,nommacta,sum(coalesce(timported,0)) debe,"
    Sql = Sql & " sum(coalesce(timporteH,0)) haber from hlinapu,cuentas "
    Sql = Sql & " where cuentas.codmacta = hlinapu.codmacta and "
    
    Sql = Sql & " fechaent=" & DBSet(FechaInicioEjercicioSolicitado, "F")
    Sql = Sql & " and cuentas.codmacta < '6' "
    'Apunte de apertura
    Sql = Sql & " AND codconce =970"
    Sql = Sql & " group by 1,2 order by 2"
    
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not RT.EOF
        ImpD = DBLet(RT!Debe, "N")
        ImpH = DBLet(RT!Haber, "N")
        'Apertura. La carga sobre los valores Imcierrd y H
        'BuscarValorEnPrecargado RT!codmacta
        'ImPerD = ImpD - ImCierrD  'Para obtener los valores del periodo reales
        'ImPerH = ImpH - ImCierrH
        'Finalmente INsert mostraremos
        Sql = Sql & ",(" & vUsu.Codigo & ",'" & RT!codmacta & "','" & DevNombreSQL(RT!Nommacta) & "',"
        Sql = Sql & TransformaComasPuntos(CStr(ImCierrD)) & "," & TransformaComasPuntos(CStr(ImCierrH))
        'ANterior
        Sql = Sql & ",0,0,"
        'Periodo
        Sql = Sql & TransformaComasPuntos(CStr(ImPerD)) & "," & TransformaComasPuntos(CStr(ImPerH)) & ","
        'Total
        If ImpD >= ImpH Then
            ImpD = ImpD - ImpH
            Sql = Sql & TransformaComasPuntos(CStr(ImpD)) & ",0)"
        Else
            ImpH = ImpH - ImpD
            Sql = Sql & "0," & TransformaComasPuntos(CStr(ImpH)) & ")"
        End If
        RT.MoveNext
        
        If Len(Sql) > 100000 Then
            Sql = Mid(Sql, 2) 'kito la primera coma
            Sql = vCta & " VALUES " & Sql
            Conn.Execute Sql
            Sql = ""
        End If
    Wend
    RT.Close
    
    
    If Sql <> "" Then
        Sql = Mid(Sql, 2) 'kito la primera coma
        Sql = vCta & " VALUES " & Sql
        Conn.Execute Sql
        Sql = ""
    End If




    'Ya estan cargados a ultimo nivel. AHora cogere y segun los niveles se vean o no
    'Hare un insert into group by
    
    For M1 = 1 To 9
        If Mid(Niveles, M1, 1) = "1" Then

            '---------------------------------------------------------------
                    
            Sql = "select " & vUsu.Codigo & ",substring(cta,1," & M1 & "), nomcta, sum(aperturaD), sum(aperturaH), sum(acumAntD), sum(acumAntH),"
            Sql = Sql & "sum(acumPerD), sum(acumPerH), sum(TotalD), sum(TotalH) from tmpbalancesumas where codusu = " & vUsu.Codigo & " and cta like '" & String(vEmpresa.DigitosUltimoNivel, "_") & "' GROUP by 1,2"
            Sql = vCta & Sql
            Conn.Execute Sql
            
            'Updateo las nommactas
            Sql = "Select codmacta,nommacta from cuentas where codmacta like '" & String(M1, "_") & "'"
            RsBalPerGan.Close
            RsBalPerGan.Open Sql, Conn, adOpenKeyset, adLockPessimistic, adCmdText
            Sql = "Select cta from tmpbalancesumas where codusu = " & vUsu.Codigo & " and cta like '" & String(M1, "_") & "' GROUP BY 1"
            RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                RsBalPerGan.Find "codmacta = '" & RT!Cta & "'", , adSearchForward, 1
                If RsBalPerGan.EOF Then
                    Sql = "###"
                Else
                    Sql = DevNombreSQL(RsBalPerGan!Nommacta)
                End If
                Sql = "UPDATE tmpbalancesumas set nomcta = '" & Sql & "' WHERE cta = '" & RT!Cta & "' and codusu = " & vUsu.Codigo
                Conn.Execute Sql
                RT.MoveNext
            Wend
            RT.Close
        End If
    Next
    'Si no quiere a ultimo nivel me cargo a ultimo nivel
    If Mid(Niveles, 10, 1) = "0" Then
        Sql = "DELETE FROM tmpbalancesumas WHERE codusu = " & vUsu.Codigo & " AND cta like '" & String(vEmpresa.DigitosUltimoNivel, "_") & "'"
        Conn.Execute Sql
    End If
        
    
    CargaBalanceInicioEjercicio = True
ECargaBalanceInicioEjercicio:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RT = Nothing
End Function


'Atencion                     -<***************************
' Si tocams algi aqui, mirar en la funcion mod: libIVA.bas  LiquidacionIVAFinAnyo
Public Function LiquidacionIVANew(Periodo As Byte, Anyo As Integer, Empresa As Integer, Detallado As Boolean) As Boolean
Dim RIVA As Recordset
Dim TieneDeducibles As Boolean    'Para ahorrar tiempo
Dim HayRecargoEquivalencia As Boolean  'Para ahorrar tiempo tb
Dim IvasBienInversion As String 'Para saber si hemos comprado bien de inversion

    '       cliente     0- Facturas clientes
    '                   1- RECARGO EQUIVALENCIA
    '                   10- Intracomunitarias
    '                   12- Sujeto pasivo
    '                   14- Entregas intracomunitarias (no deducibles)
    '                   16- Exportaciones y operaciones asimiladas
    '                   2- Facturas proveedores
    '                   30- Proveedores bien de inversion
    '                   32- iva de importacion de bienes corrientes
    '                   36- iva intracomunitario de bienes corrientes
    '                   38- iva intracomunitario de bien de inversion
    '                   42- iva regimen especial agrario
    '                   61- Operaciones no sujetas o con inversión del sujeto pasivo que originan el derecho a deducción  (IVA 0% en ventas conISP)
    '                   77- DUA

    On Error GoTo eLiquidacionIVANew

    LiquidacionIVANew = False

    If vParam.periodos = 1 Then
        'Esamos en mensual
        If Periodo > 12 Then
            MsgBox "Error en el periodo a tratar.", vbExclamation
            Exit Function
        End If
        vFecha1 = CDate("01/" & Periodo & "/" & Anyo)
        M1 = DiasMes(Periodo, Anyo)
        vFecha2 = CDate(M1 & "/" & Periodo & "/" & Anyo)
        
    Else
        'IVA TRIMESTRAL
        If Periodo > 4 Then
            MsgBox "Error en el periodo a tratar.", vbExclamation
            Exit Function
        End If
        M2 = ((Periodo - 1) * 3) + 1
        vFecha1 = CDate("01/" & M2 & "/" & Anyo)
        M2 = ((Periodo - 1) * 3) + 3
        M1 = DiasMes(CByte(M2), Anyo)
        vFecha2 = CDate(M1 & "/" & M2 & "/" & Anyo)
    End If
    
    
    vCta = "ariconta" & Empresa
    
    'Para la cadena de busqueda
    LiquidacionIVANew = False
    

    '-----------------------------------------------
    '-----------------------------------------------
    '-----------------------------------------------
    'CLIENTES
    '-----------------------------------------------
    ' iva REGIMEN GENERAL
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente )"
        
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,0"
    Sql = Sql & " ,sum(baseimpo),sum(impoiva), 0"
    Sql = Sql & ", " & Empresa & "," & Periodo & "," & Anyo & ",0 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
    Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factcli.codopera = 0 " ' tipo de operacion general
    Sql = Sql & " and tipodiva in (0,1) " 'solo iva e igic
    Sql = Sql & " and factcli_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    
    'Junio2019
    'Rectificativas SEPRADAS
    If vParam.RectificativasSeparadas303 Then Sql = Sql & " and factcli.codconce340<>'D'"
    
    Sql = Sql & " group by 1,2,3"
    Conn.Execute Sql
    
    
    'Junio2019
    'Rectificativas SEPRADAS
    If vParam.RectificativasSeparadas303 Then
            Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente )"
    
            'GRABAMOS EN IVA un 100. En el report sabremos que son facturas normales, sin recargo equivalencia peeeero, rectificativas
            Sql = Sql & " select " & vUsu.Codigo & ",100 porciva,0"
            Sql = Sql & " ,sum(baseimpo),sum(impoiva), 0"
            Sql = Sql & ", " & Empresa & "," & Periodo & "," & Anyo & ",0 "
            Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
            Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
            Sql = Sql & " and factcli.codopera = 0 " ' tipo de operacion general
            Sql = Sql & " and tipodiva in (0,1) " 'solo iva e igic
            Sql = Sql & " and factcli_totales.codigiva = tiposiva.codigiva "
            Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
            Sql = Sql & " and factcli.codconce340='D'"
            Sql = Sql & " group by 1,2,3"
            Conn.Execute Sql
    End If
        
    
    
    
    ' recargo de equivalencia
    ' La cuot a de IVA ya la hemos sumado arriba. Ahora no la volvemos a poner
    Sql = "insert into tmpliquidaiva(codusu,iva,bases,ivas,codempre,periodo,ano,cliente,porcrec)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,sum(baseimpo),sum(coalesce(imporec,0)),"
    Sql = Sql & Empresa & "," & Periodo & "," & Anyo & ",1 "
    Sql = Sql & " ,coalesce(porcrec,0)"
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
    Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva in (0,1) " 'solo iva e igic
    Sql = Sql & " and factcli.codopera = 0 " ' tipo de operacion general
    Sql = Sql & " and factcli_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    Sql = Sql & " and coalesce(porcerec,0)>0"
     'Junio2019
    'Rectificativas SEPRADAS
    If vParam.RectificativasSeparadas303 Then Sql = Sql & " and factcli.codconce340<>'D'"
    Sql = Sql & " group by 1,2"
    Conn.Execute Sql
    
    If vParam.RectificativasSeparadas303 Then
            Sql = "insert into tmpliquidaiva(codusu,iva,bases,ivas,codempre,periodo,ano,cliente,porcrec)"
    
            'GRABAMOS EN IVA un 100. En el report sabremos que son facturas normales, sin recargo equivalencia peeeero, rectificativas
            Sql = Sql & " select " & vUsu.Codigo & ",101 porciva,sum(baseimpo),sum(coalesce(imporec,0)),"
            Sql = Sql & Empresa & "," & Periodo & "," & Anyo & ",1 "
            Sql = Sql & " ,coalesce(porcrec,0)"
            Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
            Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
            Sql = Sql & " and tipodiva in (0,1) " 'solo iva e igic
            Sql = Sql & " and factcli.codopera = 0 " ' tipo de operacion general
            Sql = Sql & " and factcli_totales.codigiva = tiposiva.codigiva "
            Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
            Sql = Sql & " and coalesce(porcerec,0)>0"
             'Junio2019
            Sql = Sql & " and factcli.codconce340='D'"
            Sql = Sql & " group by 1,2"
            Conn.Execute Sql
    End If
    
    
    
    
    
    ' intracomunitarias
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",10 "
    Sql = Sql & " from " & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factpro.codopera = 1 " ' tipo de operacion intracomunitaria
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    ' inversion sujeto pasivo
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva),sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",12 "
    Sql = Sql & " from " & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factpro.codopera = 4 " ' tipo de operacion inversion sujeto pasivo
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    ' entregas intracomunitarias
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",14 "
    Sql = Sql & " from " & vCta & ".factcli_totales," & vCta & ".factcli"
    Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factcli.codopera = 1 " ' tipo de operacion intracomunitaria
    Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    
    ' exportaciones y operaciones asimiladas
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec, sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",16 "
    Sql = Sql & " from " & vCta & ".factcli_totales," & vCta & ".factcli"
    Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
        
    Sql = Sql & " and factcli.codopera = 2 " ' tipo de operacion exportacion / importacion
    
    Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    
    
      
    
    
    
    ' iva REGIMEN GENERAL
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente )"
        
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,0"
    Sql = Sql & " ,sum(baseimpo),sum(impoiva), 0"
    Sql = Sql & ", " & Empresa & "," & Periodo & "," & Anyo & ", 61 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factcli_totales," & vCta & ".factcli"
    Sql = Sql & " where fecliqcl >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqcl <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factcli.codopera = 3 "
    Sql = Sql & " and factcli_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factcli_totales.numserie = factcli.numserie and factcli_totales.numfactu = factcli.numfactu and factcli_totales.anofactu = factcli.anofactu "
    
    Sql = Sql & " group by 1,2,3"
    Conn.Execute Sql
    
    
    
    
    
    
    '-----------------------------------------------
    '-----------------------------------------------
    '-----------------------------------------------
    '           PROVEEDORES
    '-----------------------------------------------
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,coalesce(porcrec,0),sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",2 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factpro.codopera = 0 " ' tipo de operacion general
    'Marzo 2019
    'SQL = SQL & " and not tipodiva in (2) " ' no sean de bienes de inversion
    'septiembre 2019
    'SQL = SQL & " and not tipodiva in (2,4) " ' no sean de bienes de inversion NI Suplidos
    Sql = Sql & " and not tipodiva in (2,3,4) " ' no sean de bienes de inversion NI Suplidos , NI no deducible   - SE LEVA a
    
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    
    If vParam.RectificativasSeparadas303 Then Sql = Sql & " and factpro.codconce340<>'D'"
    
    If vParam.ExcluirBasesIvaCeroRecibidas303 Then Sql = Sql & " AND porceiva>0"
    
    
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    
    
    If vParam.RectificativasSeparadas303 Then
    
        Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
        
        Sql = Sql & " select " & vUsu.Codigo & ",100 porciva,coalesce(porcrec,0),sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",40 "
        Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
        Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Sql = Sql & " and factpro.codopera = 0 " ' tipo de operacion general
        Sql = Sql & " and not tipodiva in (2,3,4) " ' no sean de bienes de inversion NI Suplidos NI no deducble
        Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
        Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
        Sql = Sql & " and factpro.codconce340='D'"
        If vParam.ExcluirBasesIvaCeroRecibidas303 Then Sql = Sql & " AND porceiva>0"
        Sql = Sql & " group by 1,2,3"
                        
        Conn.Execute Sql
        
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    ' bienes de inversion
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",30 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva = 2 " 'solo bienes de inversion y no de importacion / exportacion
    Sql = Sql & " and factpro.codopera = 0 " ' tipo de operacion general
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    
    ' iva de importacion de bienes corrientes
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",32 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva <> 2 " ' no tipo de iva de bien de inversion
    Sql = Sql & " and factpro.codopera = 2 " ' tipo facturas de importacion
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    
    ' iva de importacion de bienes de inversion
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",34 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva = 2 " ' no tipo de iva de bien de inversion
    Sql = Sql & " and factpro.codopera = 2 " ' tipo facturas de importacion
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    
    
    
    ' iva intracomunitaria normales
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",36 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and not tipodiva in (2) " ' tipo de iva distinto de BI
    Sql = Sql & " and factpro.codopera = 1 " ' tipo intracomunitaria
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    ' iva intracomunitaria bien de inversion
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",38 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and tipodiva = 2 " ' tipo de iva de BI
    Sql = Sql & " and factpro.codopera = 1 " ' tipo intracomunitaria
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
                    
    Conn.Execute Sql
    
    
    ' compensaciones regimen especial agrario
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",42 "
    Sql = Sql & " from " & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factpro.codopera = 5 " ' factura de REA
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2"
    Conn.Execute Sql
                    
                    
    ' DUA
    '2020 Julio SIEMPRE ENTRA
    'If vParam.InscritoDeclarDUA Then
        Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
        
        Sql = Sql & " select " & vUsu.Codigo & ",porciva,porcrec,sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",77 "
        Sql = Sql & " from " & vCta & ".factpro_totales," & vCta & ".factpro"
        Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
        Sql = Sql & " and factpro.codopera = 6 " ' factura de DUA
        Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
        Sql = Sql & " group by 1,2"
        Conn.Execute Sql
    'End If
        
                    
    'NO DEDUCIBLE EN CONMPRAS
    Sql = "insert into tmpliquidaiva(codusu,iva,porcrec,bases,ivas,imporec,codempre,periodo,ano,cliente)"
    Sql = Sql & " select " & vUsu.Codigo & ",porciva,coalesce(porcrec,0),sum(baseimpo),sum(impoiva), sum(coalesce(imporec,0))," & Empresa & "," & Periodo & "," & Anyo & ",199 "
    Sql = Sql & " from " & vCta & ".tiposiva," & vCta & ".factpro_totales," & vCta & ".factpro"
    Sql = Sql & " where fecliqpr >= '" & Format(vFecha1, FormatoFecha) & "'  AND fecliqpr <= '" & Format(vFecha2, FormatoFecha) & "'"
    Sql = Sql & " and factpro.codopera = 0 " ' tipo de operacion general
    Sql = Sql & " and tipodiva = 3 "   'NO deducible
    Sql = Sql & " and factpro_totales.codigiva = tiposiva.codigiva "
    Sql = Sql & " and factpro_totales.numserie = factpro.numserie and factpro_totales.numregis = factpro.numregis and factpro_totales.anofactu = factpro.anofactu "
    Sql = Sql & " group by 1,2,3"
    Conn.Execute Sql
    
    
    
    LiquidacionIVANew = True
eLiquidacionIVANew:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, Err.Description
        
    
End Function

'******************************************************
'**************  TESORERIA
'******************************************************
Public Function DevuelveLaCtaBanco(ByRef Cta As String) As String
Dim Rs As ADODB.Recordset
    
    DevuelveLaCtaBanco = "|"
    Set Rs = New ADODB.Recordset
    Rs.Open "Select iban from cuentas where codmacta ='" & Cta & "'", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs!IBAN, "T") <> "" Then DevuelveLaCtaBanco = UCase(DBLet(Rs!IBAN, "T")) & "|"
    End If
        
    Rs.Close
    Set Rs = Nothing
End Function

Public Function CuentaBloqeada(Cuenta As String, Fecha As Date, MostrarMensaje As Boolean) As Boolean
Dim Rs As ADODB.Recordset

    On Error GoTo ECtaB
    CuentaBloqeada = False
    Set Rs = New ADODB.Recordset
    Rs.Open "Select fecbloq from cuentas where codmacta = '" & Cuenta & "'", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs!FecBloq) Then
            If Rs!FecBloq <= Fecha Then
                CuentaBloqeada = True
                If MostrarMensaje Then _
                    MsgBox "Cuenta bloqueada: " & Cuenta & " -  Fecha: " & Format(Rs!FecBloq, "dd/mm/yyyy"), vbExclamation
            End If
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ECtaB:
    MuestraError Err.Number
End Function



