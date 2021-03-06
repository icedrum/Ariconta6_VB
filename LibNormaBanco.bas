Attribute VB_Name = "LinNormasBanco"
Option Explicit

    Dim NF As Integer
    Dim Registro As String
    Dim Sql As String


    

    Dim AuxD As String
    Private NumeroTransferencia As Integer


Public Function FrmtStr(Campo As String, Longitud As Integer) As String
    FrmtStr = Mid(Trim(Campo) & Space(Longitud), 1, Longitud)
End Function



'Formatea SIN decimales, a la derecha, rellenando a ceros
'dejando la primera posicion vacia, o con el signo
Private Function FrmtCurren(Importe As Currency, ByVal TotalPosiciones As Integer) As String

    TotalPosiciones = TotalPosiciones - 1
    FrmtCurren = Format(Importe, FormatoImporte)
    FrmtCurren = Replace(FrmtCurren, ".", "")
    FrmtCurren = Replace(FrmtCurren, ",", "")
    FrmtCurren = IIf(Importe < 0, "-", " ") & Right(String(TotalPosiciones, "0") & FrmtCurren, TotalPosiciones)
End Function


'DATOSEXTRA  :
' 1: SUFIJOEM
' 2: TEXTO ORDENANTE
' Nuevo parametro:  Si el banco emite o no  (BancoEmiteDocumento)

'MODIFICACION 20 JUNIO 2012
'------------------------------
'  Si llevamos: vParam.Norma19xFechaVto presentara un fichero con varios ordenantes
' ENE 2014.
'  SEPA. Campo 17. Identifacador deudor. Si grabo BIC o CIF para las EMPRESAS. Particulares siempre NIF

'OCT 2015
'   Si lleva F.Cobro significa que van todos a esa fecha. Si es "" es que es fec vencimientos
' si graba en la etique de zmsg> la referencia larga ,, o corta
Public Function GrabarDisketteNorma19(NomFichero As String, Remesa As String, FecPre As String, DatosExtra As String, TipoReferenciaCliente As Byte, FecCobro2 As String, BancoEmiteDocumento As Boolean, SepaEmpresasGraboNIF As Boolean, N19_15 As Boolean, FormatoXML As Boolean, esAnticipoCredito As Boolean, IdGrabadoEnFichero As String, AgruparVtos As Boolean) As Boolean

    
    If vParamT.NuevasNormasSEPA Then
        GrabarDisketteNorma19 = GrabarFicheroNorma19SEPA(NomFichero, Remesa, FecPre, TipoReferenciaCliente, RecuperaValor(DatosExtra, 1), FecCobro2, SepaEmpresasGraboNIF, N19_15, FormatoXML, esAnticipoCredito, IdGrabadoEnFichero, AgruparVtos)
    Else
        MsgBox "Error. NO SEPA", vbCritical
    End If
End Function














Private Function HayKImprimirOpcionales() As Boolean
Dim i As Integer
Dim C As String

    On Error GoTo EImprimirOpcionales
    HayKImprimirOpcionales = False
    
    'Compruebo los cuatro primeros
    i = 0

    If Not IsNull(miRsAux.Fields!text41csb) Then i = i + 1
    If Not IsNull(miRsAux.Fields!text42csb) Then i = i + 1
    If Not IsNull(miRsAux.Fields!text43csb) Then i = i + 1
        
    If i > 0 Then HayKImprimirOpcionales = True
        
    

    

    Exit Function
EImprimirOpcionales:
    Err.Clear



End Function




Private Function ImprimeOpcionales(N19 As Boolean, Valores As String, Registro As Integer, ByRef ValorEnOpcionalesVar As Boolean) As String
Dim C As String
Dim J As Integer
Dim N As Integer
    ImprimeOpcionales = ""
    ValorEnOpcionalesVar = False
    If N19 Then
        ImprimeOpcionales = "56" & CStr(80 + Registro)
    End If
    ImprimeOpcionales = ImprimeOpcionales & Valores
    N = 0
    For J = 1 To 3
        C = "text" & (Registro + 3) & CStr(J) & "csb"
        C = DBLet(miRsAux.Fields(C), "T")
        If C <> "" Then N = N + 1
        C = FrmtStr(C, 40)
        ImprimeOpcionales = ImprimeOpcionales & C
    Next J
    ImprimeOpcionales = Mid(ImprimeOpcionales & Space(60), 1, 162)
    ValorEnOpcionalesVar = N > 0
End Function





Private Function comprobarCuentasBancariasRecibos_NIF(Remesa As String) As Boolean
Dim CC As String
Dim NifsVacios As String

On Error GoTo EcomprobarCuentasBancariasRecibos

    comprobarCuentasBancariasRecibos_NIF = False

    Sql = "select * from cobros where codrem = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2)
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    NifsVacios = ""
    NF = 0
    While Not miRsAux.EOF

        If DBLet(miRsAux!IBAN, "T") = "" Or Len(DBLet(miRsAux!IBAN, "T")) <> 24 Then
            Sql = ""
        Else
            Sql = "D"
        End If

    
        If Sql = "" Then
             Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!numfactu & "-" & miRsAux!numorden
             If NF < 2 Then
                Registro = Registro & "         "
                NF = NF + 1
             Else
                Registro = Registro & vbCrLf
                NF = 0
            End If
    
        End If
    
        If DBLet(miRsAux!nifclien, "T") = "" Then
            NifsVacios = NifsVacios & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!numfactu & "-" & miRsAux!numorden & vbCrLf
        End If
        
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NifsVacios <> "" Then NifsVacios = "NIFs vacios: " & vbCrLf & vbCrLf & NifsVacios
    
    If Registro <> "" Then
        Sql = "Los siguientes vencimientos no tienen la cuenta bancaria con todos los datos." & vbCrLf & Registro
        MsgBox Sql, vbExclamation
        If NifsVacios <> "" Then MsgBox NifsVacios, vbExclamation

        Exit Function
    End If
    
    If NifsVacios <> "" Then
        MsgBox NifsVacios, vbExclamation
        Exit Function
    End If
    
    'Si llega aqui es que todos tienen DATOS
    Sql = "select iban from cobros where codrem = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2)
    Sql = Sql & " GROUP BY iban "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                Sql = Mid(miRsAux!IBAN, 5, 4) ' C�digo de entidad receptora
                Sql = Sql & Mid(miRsAux!IBAN, 9, 4) ' C�digo de oficina receptora
                
                Sql = Sql & Mid(miRsAux!IBAN, 15, 10) ' C�digo de cuenta
                
                CC = Mid(miRsAux!IBAN, 13, 2) ' D�gitos de control
                
                'Este lo mando.
                Sql = CodigoDeControl(Sql)
                If Sql <> CC Then
                    
                    Sql = " - " & Mid(miRsAux!IBAN, 13, 2) & "- " & Mid(miRsAux!IBAN, 15, 10) & " --> CC. correcto:" & Sql
                    Sql = Mid(miRsAux!IBAN, 5, 4) & " - " & Mid(miRsAux!IBAN, 9, 4) & Sql
                    Registro = Registro & Sql & vbCrLf
                End If
                miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Registro <> "" Then
        Sql = "Las siguientes cuentas no son correctas.:" & vbCrLf & Registro
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    
    If vParamT.NuevasNormasSEPA Then
        'Si continuar y esta bien, veremos si todas los bancos tienen BIC asociado
        Registro = ""
        Sql = "select mid(cobros.iban,5,4) codbanco,bics.entidad from cobros left join bics on mid(cobros.iban,5,4)=bics.entidad WHERE "
        Sql = Sql & " codrem = " & RecuperaValor(Remesa, 1)
        Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2) & " group by 1"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Registro = ""
        While Not miRsAux.EOF
            If IsNull(miRsAux!Entidad) Then Registro = Registro & "/    " & miRsAux!codbanco & "    "
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Registro <> "" Then
            Registro = Mid(Registro, 2) & vbCrLf & vbCrLf & "�Continuar?"
            Sql = "Las siguientes bancos no tiene BIC asocidado:" & vbCrLf & vbCrLf & Registro
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        
        
    End If
    
    
    
    comprobarCuentasBancariasRecibos_NIF = True
    Exit Function
EcomprobarCuentasBancariasRecibos:
    MuestraError Err.Number, "comprobar Cuentas Bancarias Recibos / NIFs"
End Function

'La norma 19 acepta como identificador del "cliente" el campo referencia en la BD
'Con lo cual comporbaremos que no esta en blanco
Private Function ComprobarCampoReferenciaRemesaNorma19(Remesa As String) As Boolean
    ComprobarCampoReferenciaRemesaNorma19 = False
    Sql = "select codmacta,NUmSerie,numfactu,numorden,referencia from cobros where codrem = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2) & " ORDER BY codmacta"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    Sql = ""
    NF = 0
    While Not miRsAux.EOF
        If DBLet(miRsAux!Referencia, "T") = "" Then
            Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!numfactu & "-" & miRsAux!numorden & vbCrLf
            NF = NF + 1
        Else
            If Len(miRsAux!Referencia) > 12 Then Sql = Sql & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!numfactu & "-" & miRsAux!numorden & "(" & miRsAux!Referencia & ")" & vbCrLf
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If NF > 0 Then
        Registro = "Referencias vacias: " & NF & vbCrLf & vbCrLf & Registro
        MsgBox Registro, vbExclamation
    Else
        If Sql <> "" Then
            Registro = "Longitud referencia incorrecta: " & vbCrLf & vbCrLf & Sql
            Registro = Registro & vbCrLf & "�Continuar?"
            If MsgBox(Registro, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        ComprobarCampoReferenciaRemesaNorma19 = True
    End If
End Function


'Modificacion noviembre 2012
'El fichero(en alzira) viene en formato WRI
'es decir el salto de linea no es el mismo. Por lo tanto
' input nf,cad  solo le UN registro con toda la informacion
' Preprocesaremos el fichero.
'  0.- Abrir
'  1.- Leer linea y apuntar a siguiente
'  2.- Preguntar si es ultima linea
'  3.- Cerrar coolee0ction
Private Sub ProcesoFicheroDevolucion(OptProces As Byte, ByRef LinFichero As Collection)
Dim B As Boolean
    'No pongo on error Que salte en el SUB ProcesaCabeceraFicheroDevolucion

    Select Case OptProces
    Case 0
        'Abrir el fichero y cargar el objeto COLLECTION
        NF = FreeFile
        Open Registro For Input As #NF
        Line Input #NF, Registro
        Set LinFichero = New Collection
        
        
        'Veremos que tipo de fichero es Normal. Ni lleva saltos de linea ni lleva vbcr ni vblf
        B = InStr(1, Registro, vbCrLf) > 0
        If B > 0 Then
            Sql = vbCrLf 'separaremos por este
        Else
            B = InStr(1, Registro, vbCr) > 0
            If B Then
                Sql = vbCr
            Else
                B = InStr(1, Registro, vbLf)
                If B Then Sql = vbLf
            End If
        End If
        
        If Not B Then
            'Normal.
            LinFichero.Add Registro
            While Not EOF(NF)
                Line Input #NF, Registro
                LinFichero.Add Registro
            Wend
        Else
            'El fichero NO va separado correctamente(tipo alzira nuevo WRI)
            Do
                NumRegElim = InStr(1, Registro, Sql)
                If NumRegElim = 0 Then
                    'NO DEBERIA PASAR
                    MsgBox "Preproceso fichero banco. Numregelim=0.  Avise soporte tecnico", vbExclamation
                Else

                    LinFichero.Add Mid(Registro, 1, NumRegElim - 1)
                    NumRegElim = NumRegElim + Len(Sql)
                    Registro = Mid(Registro, NumRegElim)  'quito el separador
                End If
                    
            Loop Until Registro = ""
        
        End If
        Close #NF
        NF = 1 'Puntero a la linea en question
        
    Case 1
        'Recorrer el COLLECTION
        'Damos la linea y movemos a la siguiente
        If NF <= LinFichero.Count Then
            Registro = LinFichero(NF)
            NF = NF + 1
        Else
            Err.Raise 513, "Sobrepasaod vector"
        End If
    Case 2
        'reutilizamos variables
        If NF > LinFichero.Count Then
            Registro = "Si"
        Else
            Registro = ""
        End If
    Case 4
        'Cerrar
        Set LinFichero = Nothing
    End Select

End Sub


'---------------------------------------------------------------------
'  DEVOLUCION FICHERO

Public Sub ProcesaCabeceraFicheroDevolucion(Fichero As String, ByRef Remesa As String)
Dim aux2 As String  'Para buscar los vencimientos
Dim FinLecturaLineas As Boolean
Dim TodoOk As Boolean
Dim ErroresVto As String
Dim Cuantos As Integer
Dim Bien As Integer
Dim LinDelFichero As Collection
Dim EsFormatoAntiguoDevolucion As Boolean

    On Error GoTo EDevRemesa
    Remesa = ""
    
    EsFormatoAntiguoDevolucion = Dir(App.Path & "\DevRecAnt.dat") <> ""
    
    
    'ANTES nov 2012
    '
    'nf = FreeFile
    'Open Fichero For Input As #nf
    Registro = Fichero 'para no pasr mas variables al proceso
    ProcesoFicheroDevolucion 0, LinDelFichero 'abrir el fichero y volcarlo sobre un Collection
    
    'Proceso la primera linea. A veriguare a que norma pertenece
    ' y hallare la remesa
    'Line Input #nf, Registro
    ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
    
    'Comproamos ciertas cosas
    Sql = "Linea 1 vacia"
    If Registro <> "" Then
        
        'NIF
        Sql = Mid(Registro, 5, 9)
        
        'Tiene valor
        If Len(Registro) <> 162 Then
            Sql = "Longitud linea incorrecta(162)"
        Else
            'Noviembre 2012
            'en lugar de 5190 comprobamos que sea 519
            If Mid(Registro, 1, 3) <> "519" Then
                Sql = "Cadena control incorrecta(519)"
            Else
                Sql = ""
            End If
        End If
    End If
    
    If Sql = "" Then
    
        'Segunda LINEA.
        'Line Input #nf, Registro
        ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
        
        Sql = "Linea 2 vacia"
        If Registro <> "" Then
            
            'NIF
            Sql = Mid(Registro, 5, 9)
            
            
            'Tiene valor
            If Len(Registro) <> 162 Then
                Sql = "Longitud linea incorrecta(162)"
            Else
                'En lugar de 5390 comprobamos por 539
                If Mid(Registro, 1, 3) <> "539" Then
                    Sql = "Cadena control incorrecta(539)"
                Else
                    
                    Sql = "Falta linea 569"
                    Remesa = ""
                    Do
                        ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                        
                        If Registro <> "" Then
                            Sql = "FIN LINEAS. No se ha encontrado linea: 569"
                            Remesa = "NO"
                        Else
                            'Line Input #nf, Registro
                            ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            
                            'BUsco la linea:
                            '5690
                            If Registro <> "" Then
                                'Nov 2012   En lugar de 5690 comprobamos 569
                                If Mid(Registro, 1, 3) = "569" Then
                                    Sql = ""
                                    Remesa = "NO"
                                End If
                            End If
                        End If
                        
                    Loop Until Remesa <> ""
                    Remesa = ""
                    
                    If Sql = "" Then
                        'VAMOS BIEN. Veremos si a partir de los datos del recibo nos dan la remesa
                        'Para ello bucaremos en registro, la cadena que contiene los datos
                        'del vencimiento
                        'Registro=
                        '5690B97230080000970000100066COSTURATEX,  S.L.                       007207779700001000660000022516311205A020574911Fac
                        '5690F46024196009242820002250DAVID MONTAGUD CARRASCO                 318871052428200022500000010187                FRA 2731591 GASOLINERA ALZICOOP         1

                        Set miRsAux = New ADODB.Recordset
                        ErroresVto = ""
                        FinLecturaLineas = False
                        Cuantos = 0
                        Bien = 0
                        Do
                            
                            If Mid(Registro, 1, 3) = "569" Then
                                'Los vtos vienen en estas lineas
                                Cuantos = Cuantos + 1
                                Registro = Mid(Registro, 99, 17)
                                Sql = "Select codrem,anyorem,siturem from cobros where fecfactu='20" & Mid(Registro, 5, 2) & "-" & Mid(Registro, 3, 2) & "-" & Mid(Registro, 1, 2)
                                aux2 = Sql
                                
                                'Problemas en alzira
                                'If Not IsNumeric(Mid(Registro, 17, 1)) Then
                                'Sept 2013
                                If Not EsFormatoAntiguoDevolucion Then
                                    Sql = Sql & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 7)) & " AND numorden=" & Mid(Registro, 16, 1)
                                    'Problema en herbelca. El numero de vto NO viene con la factura
                                    aux2 = aux2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 8))
                                    
                                Else
                                    'El vencimiento si que es el 17
                                    Sql = Sql & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 7)) & " AND numorden=" & Mid(Registro, 17, 1)
                                    aux2 = aux2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 8))
                                    
                                End If
                                
                                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                TodoOk = False
                                Sql = "Vencimiento no encontrado: " & Registro
                                If Not miRsAux.EOF Then
                                    If IsNull(miRsAux!Codrem) Then
                                        Sql = "Vencimiento sin Remesa: " & Registro
                                    Else
                                        Sql = miRsAux!Codrem & "|" & miRsAux!Anyorem & "|�"
                                        
                                        If InStr(1, Remesa, Sql) = 0 Then Remesa = Remesa & Sql
                                        Sql = ""
                                        TodoOk = True
                                    End If
                                End If
                                miRsAux.Close
                                
                                
                                If Not TodoOk Then
                                    'Los busco sin Numorden
                                    miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                    If Not miRsAux.EOF Then
                                        If IsNull(miRsAux!Codrem) Then
                                            Sql = "Vencimiento sin Remesa: " & Registro
                                        Else
                                            Sql = miRsAux!Codrem & "|" & miRsAux!Anyorem & "|�"
                                            
                                            If InStr(1, Remesa, Sql) = 0 Then Remesa = Remesa & Sql
                                            Sql = ""
                                            TodoOk = True
                                        End If
                                    End If
                                    miRsAux.Close
                                
                                End If
                                
                                
                                
                                If Sql <> "" Then
                                    ErroresVto = ErroresVto & vbCrLf & Sql
                                Else
                                    Bien = Bien + 1
                                End If
                            Else
                                'La linea no empieza por 569
                                'veremos los totales
                                
                                If Mid(Registro, 1, 3) = "599" Then
                                    'TOTAL TOTAL
                                    Sql = Mid(Registro, 105, 10)
                                    If Val(Sql) <> Cuantos Then ErroresVto = "Fichero: " & Sql & "   Leidos" & Cuantos & vbCrLf & ErroresVto & vbCrLf & Sql
                                End If
                            End If
                            
                            'Siguiente linea
                            ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                            
                            If Registro <> "" Then
                                FinLecturaLineas = True
                            Else
                                'Line Input #nf, Registro
                                ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            End If
                            
                        Loop Until FinLecturaLineas
                        
                        If Cuantos <> Bien Then ErroresVto = ErroresVto & vbCrLf & "Total: " & Cuantos & "   Correctos:" & Bien
                        
                        Sql = ErroresVto
                        Set miRsAux = Nothing
                    
                    End If
                End If  'Control SEGUNDA LINEA
        
        
            End If
        End If
    
    End If  'DE SEGUNDA LINEA
    
    ProcesoFicheroDevolucion 3, LinDelFichero
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
    Else
        'Remesa = Mid(Registro, 1, 4) & "|" & Mid(Registro, 5) & "|"
        
        
        'Ahora comprobaremos que para cada remesa  veremos si existe y si la situacion es la contabilizadxa
        Sql = Remesa
        Registro = "" 'Cadena de error de situacion remesas
        Set miRsAux = New ADODB.Recordset
        Do
            Cuantos = InStr(1, Sql, "�")
            If Cuantos = 0 Then
                Sql = ""
            Else
                aux2 = Mid(Sql, 1, Cuantos - 1)
                Sql = Mid(Sql, Cuantos + 1)
                
                
                'En aux2 tendre codrem|an�orem|
                aux2 = RecuperaValor(aux2, 1) & " AND anyo = " & RecuperaValor(aux2, 2)
                aux2 = "Select situacion from remesas where codigo = " & aux2
                miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    aux2 = "-No se encuentra remesa"
                Else
                    'Si que esta.
                    'Situacion
                    If CStr(miRsAux!Situacion) <> "Q" Then
                        aux2 = "- Situacion incorrecta : " & miRsAux!Situacion
                    Else
                        aux2 = "" 'TODO OK
                    End If
                End If
            
                If aux2 <> "" Then
                    aux2 = aux2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
                    aux2 = Replace(aux2, " AND ", " ")
                    aux2 = Replace(aux2, "anyo", "a�o")
                    Registro = Registro & vbCrLf & aux2
                End If
                miRsAux.Close
            End If
        Loop Until Sql = ""
        Set miRsAux = Nothing
        
        
        If Registro <> "" Then
            Registro = "Error remesas " & vbCrLf & String(30, "=") & Registro
            MsgBox Registro, vbExclamation
            
            'Pongo REMESA=""
            Remesa = "" 'para que no continue el preoceso de devolucion
        End If
        
    End If
    
    Exit Sub
EDevRemesa:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion"
End Sub




Public Sub ProcesaLineasFicheroDevolucion(Fichero As String, ByRef Listado As Collection, ByRef EsSepa As Boolean)
Dim Registro As String
Dim SumaComprobacion As Currency
Dim impo As Currency
Dim Fin As Boolean
Dim B As Boolean
Dim Aux As String
Dim C2 As String
Dim bol As Boolean

    On Error GoTo EDevRemesa1
  
    
    
    

    NF = FreeFile
    Open Fichero For Input As #NF
    
    'Las dos primeras son el encabezado.
    ' Noviembre 2012. Hay que comprobar que si vienen todo en una linea o NO
    Line Input #NF, Registro
    
    
    B = InStr(1, Registro, vbCrLf) > 0
    If B > 0 Then
        Aux = vbCrLf 'separaremos por este
    Else
        B = InStr(1, Registro, vbCr) > 0
        If B Then
            Aux = vbCr
        Else
            B = InStr(1, Registro, vbLf)
            If B Then Aux = vbLf
        End If
    End If
    
    EsSepa = False
    If Mid(Registro, 1, 4) = "2119" Then EsSepa = True
        
    
    
    If B Then
        'TRAE TODO en una unica linea. Separaremos por el vbcr o vbcrlf
        Do
                NumRegElim = InStr(1, Registro, Aux)
                If NumRegElim = 0 Then
                    
                Else

                    Sql = Mid(Registro, 1, NumRegElim - 1)
                    NumRegElim = NumRegElim + Len(Aux)
                    Registro = Mid(Registro, NumRegElim)  'quito el separador
                    
                    
                   
                    
                    
                    If EsSepa Then
                        C2 = Mid(Sql, 1, 2)
                        If C2 = "23" Then
                            impo = Val(Mid(Sql, 89, 11)) / 100
                            SumaComprobacion = SumaComprobacion + impo
                            
                            'Cuestion 2
                            'Datos identifictivos del vencimiento
                            Sql = Mid(Sql, 21, 35)
                            Listado.Add Sql
                            Sql = ""
                        Else
                            If C2 = "99" Then 'antes 5990
                                Fin = True
                                impo = Val(Mid(Sql, 3, 17)) / 100
                            Else
                                Sql = ""
                            End If
                        End If
                    Else
                        C2 = Mid(Sql, 1, 3)
                        If C2 = "569" Then
                            impo = Val(Mid(Sql, 89, 10)) / 100
                            SumaComprobacion = SumaComprobacion + impo
                            
                            'Cuestion 2
                            'Datos identifictivos del vencimiento
                            Sql = Mid(Sql, 89, 27)
                            Listado.Add Sql
                            Sql = ""
                        Else
                            If C2 = "599" Then 'antes 5990
                                Fin = True
                                impo = Val(Mid(Sql, 89, 10)) / 100
                            Else
                                Sql = ""
                            End If
                        End If
                    
                    End If
                    
                End If
                    
        Loop Until Registro = ""
            
        'Cerramos y salimos
        Close #NF
        Exit Sub
    End If
    
    Line Input #NF, Registro
    
    'Ahora empezamos
    SumaComprobacion = 0
    Fin = False
    Sql = ""
    Do
        Line Input #NF, Registro
        If Registro <> "" Then
         
            Sql = Mid(Registro, 1, 3)
            
            If EsSepa Then
                bol = Mid(Registro, 1, 4) = "2319"
            Else
                bol = Sql = "569"
            End If
            If bol Then
                'Registro normal de devolucion
                '1... 68 carcaater
                '5690B972300800003169816315  RUANO MORENO, VICENTE                   "
                '69 .. 162
                '3082140015316981631500000350890047080000004708Fact. 2059121 31/12/2005 Tarj   9434    1
                
                'Cuestion 1:
                'Importe: 0000035089 desde la poscion  hasta la posicion
                If EsSepa Then
                    impo = Val(Mid(Registro, 89, 11)) / 100
                Else
                    impo = Val(Mid(Registro, 89, 10)) / 100
                End If
                SumaComprobacion = SumaComprobacion + impo
                
                'Cuestion 2
                'Datos identifictivos del vencimiento
                If EsSepa Then
                    Sql = Mid(Registro, 21, 35)
                Else
                    Sql = Mid(Registro, 89, 27)
                End If
                Listado.Add Sql
                Sql = ""
            Else
                
                If EsSepa Then
                    bol = Mid(Registro, 1, 2) = "99"
                Else
                    bol = Sql = "599"
                End If
                    
                If bol Then
                    Fin = True
                    If EsSepa Then
                        impo = Val(Mid(Registro, 3, 17)) / 100
                    Else
                        impo = Val(Mid(Registro, 89, 10)) / 100
                    End If
                Else
                    Sql = ""
                End If
            End If
        End If
        If EOF(NF) Then Fin = True
    Loop Until Fin
    Close #NF
    
    If Sql = "" Then
        MsgBox "No se ha leido la linea final fichero", vbExclamation
        Set Listado = Nothing
    Else
        'OK salimos
        If impo <> SumaComprobacion Then
            Sql = "Error leyendo importes. �Desea continuar con los datos obtenidos?"
            If MsgBox(Sql, vbExclamation) = vbNo Then Set Listado = Nothing
        End If
    End If
    
    
    Exit Sub
EDevRemesa1:
    MuestraError Err.Number, "Lineas devolucion"
End Sub


'------ aqui aqui aqui


        


'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************
'
'       Normas 34 y 68
'
'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************

'----------------------------------------------------------------------
'  Copia fichero generado bajo
'Public Sub CopiarFicheroNorma43(Es34 As Boolean, Destino As String)
Public Sub CopiarFicheroNormaBancaria(TipoFichero As Byte, Destino As String)
    
    'If Not CopiarEnDisquette(True, 3) Then
        AuxD = Destino
        'CopiarEnDisquette False, 0, Es34 'A disco
        CopiarEnDisquette TipoFichero
        
End Sub
'Private Function CopiarEnDisquette(A_disquetera As Boolean, Intentos As Byte, Es34 As Boolean) As Boolean
'TipoFichero
'   0- norma 34
'   1- N8
'   2- Caixa confirming
Private Function CopiarEnDisquette(TipoFichero As Byte) As Boolean
Dim i As Integer
Dim cad As String

On Error Resume Next

    CopiarEnDisquette = False
    
 
        If AuxD = "" Then
            cad = Format(Now, "ddmmyyhhnn")
            cad = App.Path & "\" & cad & ".txt"
        Else
            cad = AuxD
        End If
        Select Case TipoFichero
        Case 0
            FileCopy App.Path & "\norma34.txt", cad
        Case 1
            FileCopy App.Path & "\norma34.txt", cad
        Case 2
 
            FileCopy App.Path & "\norma68.txt", cad
        Case 3
            'vbConfirmingStd
            
            FileCopy App.Path & "\confirming.txt", cad
            
        End Select
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte t�cnico." & vbCrLf & Err.Description, vbCritical
            Err.Clear
        Else
            'MsgBox "El fichero esta guardado como: " & cad, vbInformation
        End If
            
    'End If
End Function

Public Function GeneraFicheroNorma34(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean, Anyo As String, IdFichero As String, AgrupaVtos As Boolean) As Boolean
    
    
        If vParamT.NormasFormatoXML Then
            GeneraFicheroNorma34 = GeneraFicheroNorma34SEPA_XML(CIF, Fecha, CuentaPropia, CLng(vNumeroTransferencia), Pagos, ConceptoTransferencia, Anyo, IdFichero, AgrupaVtos)
        
        Else
            MsgBox "NO disponible", vbExclamation
        End If
   
End Function


Public Function comprobarCuentasBancariasPagos(Transferencia As String, Anyo As String, Pagos As Boolean) As Boolean
Dim CC As String
Dim IBAN As String
On Error GoTo EcomprobarCuentasBancariasPagos

    comprobarCuentasBancariasPagos = False
    If Pagos Then
        Sql = "select * from pagos where nrodocum = " & Transferencia & " and anyodocum = " & DBSet(Anyo, "N")
    Else
        'ABONOS
        Sql = "Select * "
        Sql = Sql & " FROM cobros where transfer=" & Transferencia
        Sql = Sql & " and anyorem = " & DBSet(Anyo, "N")
    End If
    
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    NF = 0
    While Not miRsAux.EOF

        If DBLet(miRsAux!IBAN, "T") = "" Or Len(DBLet(miRsAux!IBAN, "T")) <> 24 Then
            Sql = ""
        Else
            Sql = "D"
        End If

    
        If Sql = "" Then
             Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!numfactu & "-" & miRsAux!numorden
             If NF < 2 Then
                Registro = Registro & "         "
                NF = NF + 1
             Else
                Registro = Registro & vbCrLf
                NF = 0
            End If
    
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Registro <> "" Then
        Sql = "Los siguientes vencimientos no tienen la cuenta bancaria con todos los datos." & vbCrLf & Registro
        MsgBox Sql, vbExclamation
        Exit Function
    End If
    
    
    'Si llega aqui es que todos tienen DATOS
    If Pagos Then
        Sql = "select iban from pagos where nrodocum = " & Transferencia & " and anyodocum = " & DBSet(Anyo, "N")
        Sql = Sql & " GROUP BY mid(iban,5,4),mid(iban,9,4),mid(iban,15,10),mid(iban,13,2)"
    Else
        Sql = "SELECT iban"
        Sql = Sql & " FROM cobros where transfer=" & Transferencia & " and anyorem = " & DBSet(Anyo, "N")
        Sql = Sql & " GROUP BY mid(iban,5,4),mid(iban,9,4),mid(iban,15,10),mid(iban,13,2)"
    End If
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                Sql = Mid(miRsAux!IBAN, 5, 4) ' C�digo de entidad receptora
                Sql = Sql & Mid(miRsAux!IBAN, 9, 4) ' C�digo de oficina receptora
                
                Sql = Sql & Mid(miRsAux!IBAN, 15, 10) ' C�digo de cuenta
                
                CC = Mid(miRsAux!IBAN, 13, 2) ' D�gitos de control
                
                'Este lo mando.
                IBAN = Mid(Sql, 1, 8) & CC & Mid(Sql, 9)
                
                Sql = CodigoDeControl(Sql)
                If Sql <> CC Then
                    
                    Sql = " - " & Mid(miRsAux!IBAN, 13, 2) & "- " & Mid(miRsAux!IBAN, 15, 10) & " --> CC. correcto:" & Sql
                    Sql = Mid(miRsAux!IBAN, 5, 4) & " - " & Mid(miRsAux!IBAN, 9, 4) & Sql
                    Registro = Registro & Sql & vbCrLf
                End If
                
                
                'Noviembre 2013
                'IBAN
                If vParamT.NuevasNormasSEPA Then
                        Sql = "ES"
                        If DBLet(miRsAux!IBAN, "T") <> "" Then Sql = Mid(miRsAux!IBAN, 1, 2)
                    
                
                        If Not DevuelveIBAN2(Sql, IBAN, IBAN) Then
                            
                            Sql = "Error calculo"
                        Else
                            Sql = Sql & IBAN
                            If Mid(DBLet(miRsAux!IBAN, "T"), 1, 4) <> Sql Then
                                Sql = "Error IBAN. Calculado " & Sql & " / " & Mid(DBLet(miRsAux!IBAN, "T"), 1, 4)
                            Else
                                'OK
                                Sql = ""
                            End If
                        End If
                        
                        If Sql <> "" Then
                            Sql = Sql & " - " & Mid(miRsAux!IBAN, 13, 2) & "- " & Mid(miRsAux!IBAN, 15, 10) & " --> CC. correcto:" & Sql
                            Sql = Mid(miRsAux!IBAN, 5, 4) & " - " & Mid(miRsAux!IBAN, 9, 4) & Sql
                            Registro = Registro & "Error obteniendo IBAN: " & Sql & vbCrLf
                        End If
                End If
                
                
                miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Registro <> "" Then
        Sql = "Generando diskette." & vbCrLf & vbCrLf
        Sql = Sql & "Las siguientes cuentas no son correctas.:" & vbCrLf & Registro
        Sql = Sql & vbCrLf & "�Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    comprobarCuentasBancariasPagos = True
    Exit Function
EcomprobarCuentasBancariasPagos:
    MuestraError Err.Number, "comprobar Cuentas Bancarias pagos"
End Function



Public Function RellenaABlancos(Cadena As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Space(Longitud)
    If PorLaDerecha Then
        cad = Cadena & cad
        RellenaABlancos = Left(cad, Longitud)
    Else
        cad = cad & Cadena
        RellenaABlancos = Right(cad, Longitud)
    End If
    
End Function



Private Function RellenaAceros(Cadena As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        cad = Cadena & cad
        RellenaAceros = Left(cad, Longitud)
    Else
        cad = cad & Cadena
        RellenaAceros = Right(cad, Longitud)
    End If
    
End Function





'******************************************************************************************************************
'******************************************************************************************************************
'
'       Genera fichero CAIXACONFIRMING
'
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroCaixaConfirming(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroCaixaConfirming = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 9)
    Aux = Mid(CIF & Space(10), 1, 9)
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!IBAN) Then
            cad = ""
        Else
            
            CodigoOrdenante = Mid(DBLet(Rs!IBAN), 4, 20) 'Format(RS!Entidad, "0000") & Format(DBLet(RS!oficina, "N"), "0000") & Format(DBLet(RS!Control, "N"), "00") & Format(DBLet(RS!CtaBanco, "T"), "0000000000")
            
            If Not DevuelveIBAN2("ES", CodigoOrdenante, cad) Then cad = ""
            CuentaPropia = "ES" & cad & CodigoOrdenante
                        
            'Esta variable NO se utiliza. La cojo "prestada"
            'Guardare el numero de contrato de CAIXACONFIRMING
            ' Sera, un char de 14
            ' Si no pone nada sera oficnacuenta  Total 14 posiciones
            ConceptoTr_ = Trim(DBLet(Rs!CaixaConfirming, "T"))
            If ConceptoTr_ = "" Then ConceptoTr_ = Mid(CodigoOrdenante, 5, 4) & Mid(CodigoOrdenante, 11, 10)
            
            '                ENTIDAD
            ConceptoTr_ = Mid(CodigoOrdenante, 1, 4) & ConceptoTr_
        End If
        
        
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
  
    Set Rs = New ADODB.Recordset
    
    'CABECERA
    'UNo
    Aux = "0156" & CodigoOrdenante & Space(12) & "001" & Format(Fecha, "ddmmyy") & Space(6)
    Aux = Aux & ConceptoTr_ & "1" & "EUR" & Space(9)   'Ya esta. Ya he utlizado la variable ConceptoTr_. Nada mas
    Print #NFich, Aux
    'Nombre
    Aux = "0156" & CodigoOrdenante & Space(12) & "002" & FrmtStr(vEmpresa.nomempre, 36) & Space(7)
    Print #NFich, Aux
    
    'Registros obligatorios  3 4
    Aux = "Select pobempre, provempre from empresa2"
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    For Regs = 0 To 1
        Aux = "0156" & CodigoOrdenante & Space(12) & Format(Regs + 3, "000") & FrmtStr(DBLet(Rs.Fields(Regs), "T"), 36) & Space(7)
        Print #NFich, Aux
    Next
    Rs.Close
    
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    
    Aux = "Select pagos.* from pagos"
    Aux = Aux & " where  nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum = " & DBSet(vAnyoTransferencia, "N")
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not Rs.EOF
                '*********************************************************
                'Suposicion 1,. TODOS son nacionales
                '*********************************************************
                
                'Im = DBLet(Rs!imppagad, "N")
                Im = 0
                Im = Rs!ImpEfect - Im
                Aux = RellenaABlancos(Rs!NifProve, True, 12)
                
                    
                'Reg 010
                Aux = "0656" & CodigoOrdenante & Aux & "010"
                Aux = Aux & RellenaAceros(CStr(Im * 100), False, 12)
                Aux = Aux & FrmtStr(Mid(DBLet(Rs!IBAN, "N"), 5, 4), 4) & FrmtStr(Mid(DBLet(Rs!IBAN, "N"), 9, 4), 4)
                Aux = Aux & FrmtStr(Mid(DBLet(Rs!IBAN, "N"), 15, 10), 10) & "1" & "9" & "  " & FrmtStr(Mid(DBLet(Rs!IBAN, "N"), 13, 2), 2)
                Aux = Aux & "N" & "C" & "EUR  "
                Print #NFich, Aux
                
        
           
           
                'nomprove  domprove  pobprove  cpprove  proprove  nifprove  codpais
                'OBligaorio 011   Nombre
                Aux = RellenaABlancos(Rs!NifProve, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "011"
                Aux = Aux & FrmtStr(DBLet(Rs!nomprove, "T"), 36) & Space(7)
                Print #NFich, Aux
           
                'OBligaorio 012   direccion
                Aux = RellenaABlancos(Rs!NifProve, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "012"
                Aux = Aux & FrmtStr(DBLet(Rs!domprove, "T"), 36) & Space(7)
                Print #NFich, Aux
           
                'OBligaorio 014   cpos provi
                Aux = RellenaABlancos(Rs!NifProve, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "014"
                Aux = Aux & FrmtStr(DBLet(Rs!cpprove, "N"), 5) & FrmtStr(DBLet(Rs!pobprove, "T"), 31) & Space(7)
                Print #NFich, Aux
                
                'OBligaorio 016   ID factura
                Aux = RellenaABlancos(Rs!NifProve, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "016"
                Aux = Aux & "T" & Format(Rs!FecFactu, "ddmmyy") & FrmtStr(Rs!numfactu, 15) & Format(Rs!fecefect, "ddmmyy") & Space(15)
                Print #NFich, Aux
           
                 
        
               'Totales
               Importe = Importe + Im
               Regs = Regs + 1
               Rs.MoveNext
        Wend
        'Imprimimos totales
        Aux = "08" & "56"
        Aux = Aux & CodigoOrdenante    'llevara tb la ID del socio
        Aux = Aux & Space(15)
        Aux = Aux & RellenaAceros(CStr(Int(Round(Importe * 100, 2))), False, 12)
        Aux = Aux & RellenaAceros(CStr((Regs)), False, 8)
        Aux = Aux & RellenaAceros(CStr((Regs * 5) + 4 + 1), False, 10)    '4 de cabecera + uno de totales
        Aux = RellenaABlancos(Aux, True, 72)
        Print #NFich, Aux
        
        
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    If Regs > 0 Then
        GeneraFicheroCaixaConfirming = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function






'Grupos santander. Confirming
'******************************************************************************************************************
'******************************************************************************************************************
'
'       Genera fichero
'
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroGrSantanderConfirming(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroGrSantanderConfirming = False
    
    NumeroTransferencia = vNumeroTransferencia
    NFich = -1
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 9)
    Aux = Mid(CIF & Space(10), 1, 9)
    If Rs.EOF Then
        cad = ""
    Else

            
            'CodigoOrdenante = Mid(DBLet(Rs!IBAN), 4, 20) 'Format(RS!Entidad, "0000") & Format(DBLet(RS!oficina, "N"), "0000") & Format(DBLet(RS!Control, "N"), "00") & Format(DBLet(RS!CtaBanco, "T"), "0000000000")
            
            'If Not DevuelveIBAN2("ES", CodigoOrdenante, Cad) Then Cad = ""
            'CuentaPropia = "ES" & Cad & CodigoOrdenante
                        
            'Esta variable NO se utiliza. La cojo "prestada"
            'Guardare el numero de contrato de CAIXACONFIRMING
            ' Sera, un char de 14
            ' Si no pone nada sera oficnacuenta  Total 14 posiciones
            'ConceptoTr_ = Trim(DBLet(Rs!caixaconfirming, "T"))
            'If ConceptoTr_ = "" Then ConceptoTr_ = Mid(CodigoOrdenante, 5, 4) & Mid(CodigoOrdenante, 11, 10)
           '
            '                ENTIDAD
            'ConceptoTr_ = Mid(CodigoOrdenante, 1, 4) & ConceptoTr_
       
            CodigoOrdenante = Mid(DBLet(Rs!CaixaConfirming, "T") & Space(16), 1, 16)
        
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    '1PPK             B20899563                EULER POMPAK, S.L. EULER POMPAK, S.L.   20181001CONF.LIKSUR30OCTEUR
                                             ' 1234567890123456789012345678901234567890
    
    'Resgristro 1 de cabecera
    
    
    Aux = Mid(Aux & Space(25), 1, 25)   'CIF +25
    Aux = "1" & CodigoOrdenante & Aux
    Aux = Aux & Mid(vEmpresa.NombreEmpresaOficial & Space(40), 1, 40)
    Aux = Aux & Format(Fecha, "yyyymmdd") & Mid(ConceptoTr_ & Space(16), 1, 16) & "EUR"
    Aux = Mid(Aux & Space(578), 1, 578)
    Print #NFich, Aux
    
    
    
    
    Set Rs = New ADODB.Recordset
    
    
    
    'Imprimimos las lineas
    cad = "Abriendo RS"
    Aux = "Select pagos.*,maidatos ,telefonocta from pagos left join cuentas on cuentas.codmacta=pagos.codmacta"
    Aux = Aux & " where  nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum = " & DBSet(vAnyoTransferencia, "N")
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not Rs.EOF
                '*********************************************************
                'Suposicion 1,. TODOS son nacionales
                '*********************************************************
               
                Im = 0
                Im = Rs!ImpEfect - Im
                
                cad = "NIF"
                Aux = "2" & RellenaABlancos(Rs!NifProve, True, 15)   'pos :16
                Aux = Aux & "1" & RellenaABlancos(Rs!NifProve, True, 22)   'pos :16
                
                cad = "NOmbre prov"
                Aux = Aux & "J" & RellenaABlancos(Rs!nomprove, True, 90)
                'Pos:131
                Aux = Aux & "ESCL"
                'Pos:135
                Aux = Aux & RellenaABlancos(DBLet(Rs!domprove, "T"), True, 88)
                'Pos:223
                Aux = Aux & RellenaABlancos(DBLet(Rs!pobprove, "T"), True, 25)
                'Pos:248
                Aux = Aux & RellenaABlancos(DBLet(Rs!cpprove, "T"), True, 8)
                'Pos:256
                Aux = Aux & RellenaABlancos(DBLet(Rs!proprove, "T"), True, 25)
                '
                Aux = Aux & "ES" & RellenaABlancos(DBLet(Rs!telefonocta, "T"), True, 28)
                '311
                Aux = Aux & RellenaABlancos(DBLet(Rs!maidatos, "T"), True, 60)
                
                'NOVIEMBRE 2018
                cad = Mid(Mid(Rs!IBAN, 5) & Space(30), 1, 30)
                Aux = Aux & "T" & cad
                
                
                ' Si se indica IBAN en las posiciones 414 a 461,
                'entonces deb�is indicar tambi�n el swfit.
                
                cad = Mid(Rs!IBAN, 5, 4)
                cad = DevuelveDesdeBD("bic", "bics", "entidad", cad)
                If cad = "" Then Err.Raise 513, , "Error obteniendo BIC para IBAN " & Rs!IBAN
                Aux = Aux & RellenaABlancos(cad, True, 12)
                
                '414
                cad = "IBAN"
                Aux = Aux & RellenaABlancos(Rs!IBAN, True, 47)
                cad = "N� factura"
                Aux = Aux & "EUR" & "F" & RellenaABlancos(Rs!numfactu, True, 15)
                
                '480
                Aux = Aux & RellenaAceros(CStr(Im * 100), False, 15)
                '495
                cad = "Fecha factura-vto"
                Aux = Aux & Format(Rs!FecFactu, "yyyymmdd") & Format(Rs!fecefect, "yyyymmdd")
                Aux = Aux & Space(45) & "E" & Space(22)
                        
                Print #NFich, Aux
                
               'Totales
               Importe = Importe + Im
               Regs = Regs + 1
               Rs.MoveNext
        Wend
        'Imprimimos totales
        Aux = "3"
        Aux = Aux & RellenaAceros(CStr(Regs), False, 6)
        Aux = Aux & RellenaAceros(CStr(Importe * 100), False, 15)
        Aux = Mid(Aux & Space(560), 1, 578)
        Print #NFich, Aux
        
        
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    NFich = -1
    If Regs > 0 Then
        GeneraFicheroGrSantanderConfirming = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description, cad
     If NFich > 0 Then Close (NFich)
End Function

Public Function GeneraFicheroGrSantanderConfirmingC12(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroGrSantanderConfirmingC12 = False
    
    NumeroTransferencia = vNumeroTransferencia
    NFich = -1
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    Aux = Mid(CIF & Space(10), 1, 12)
    If Rs.EOF Then
        cad = ""
    Else
       
            CodigoOrdenante = Mid(DBLet(Rs!CaixaConfirming, "T") & Space(3), 1, 3)
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    
    'Resgristro 1 de cabecera

    
   
    Aux = "1     " & Aux & CodigoOrdenante
    Aux = Aux & Mid(vEmpresa.NombreEmpresaOficial & Space(40), 1, 40)
    Aux = Aux & Format(Fecha, "yyyymmdd")
    Aux = Mid(Aux & Space(273), 1, 273)
    Print #NFich, Aux
    
    
    
    
    Set Rs = New ADODB.Recordset
    
    
    
    'Imprimimos las lineas
    cad = "Abriendo RS"
    Aux = "Select pagos.*,maidatos ,telefonocta from pagos left join cuentas on cuentas.codmacta=pagos.codmacta"
    Aux = Aux & " where  nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum = " & DBSet(vAnyoTransferencia, "N")
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not Rs.EOF
                '*********************************************************
                'Suposicion 1,. TODOS son nacionales
                '*********************************************************
               
                Im = 0
                Im = Rs!ImpEfect - Im
                
                Aux = "2" & "000"
                
                cad = "Proveedor"
                Aux = Aux & RellenaABlancos(Rs!NifProve, True, 12)
                Aux = Aux & RellenaABlancos(Rs!codmacta, True, 15)
                
                
                Aux = Aux & RellenaABlancos(Rs!nomprove, True, 40)
                Aux = Aux & RellenaABlancos(DBLet(Rs!domprove, "T"), True, 30)
                Aux = Aux & RellenaABlancos(DBLet(Rs!pobprove, "T"), True, 30)
                Aux = Aux & Mid(Format(DBLet(Rs!cpprove, "N"), "00000"), 1, 5)
                Aux = Aux & "ES " & RellenaABlancos(DBLet(Rs!telefonocta, "T"), True, 10)
                Aux = Aux & Space(10) & Space(10) & "N"
                
                'NOVIEMBRE 2018
                cad = Rs!IBAN
                Aux = Aux & Mid(cad, 5, 20)
                Aux = Aux & "C"
                
                
                cad = "Fac "
                Aux = Aux & RellenaABlancos(Rs!numfactu, True, 15)
                Aux = Aux & RellenaAceros(Replace(CStr(Abs(Im)), ",", ""), False, 17)
                Aux = Aux & IIf(Im < 0, "H", "D")
                Aux = Aux & "0978"
                Aux = Aux & Format(Rs!FecFactu, "yyyymmdd") & Format(Rs!fecefect, "yyyymmdd")
                Aux = Aux & Space(23)
                Aux = Aux & Space(6)
                Print #NFich, Aux
                
               'Totales
               Importe = Importe + Im
               Regs = Regs + 1
               Rs.MoveNext
        Wend
        'Imprimimos totales
        Aux = "4"
        'Aux = Aux & RellenaAceros(CStr(Regs), False, 6)
        Aux = Aux & Space(18)
        Aux = Aux & RellenaAceros(CStr(Importe * 100), False, 17)
        Aux = Mid(Aux & Space(273), 1, 273)
        Print #NFich, Aux
        
        
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    NFich = -1
    If Regs > 0 Then
        GeneraFicheroGrSantanderConfirmingC12 = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description, cad
     If NFich > 0 Then Close (NFich)
End Function


'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
'SABADEL
'
Public Function GeneraFicheroSabadellConfirming(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim miBanco As String
Dim Importe As Currency
Dim CodigoOrdenante As String
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroSabadellConfirming = False
    
    NumeroTransferencia = vNumeroTransferencia
    NFich = -1
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 9)
    Aux = Mid(CIF & Space(10), 1, 9)
    If Rs.EOF Then
        cad = ""
    Else

        miBanco = Mid(Rs!IBAN, 5)
       
        CodigoOrdenante = Mid(DBLet(Rs!CaixaConfirming, "T") & Space(12), 1, 12)
        
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    
    'Resgristro 1 de cabecera
    
    
    Aux = "1  "
    Aux = Aux & Mid(vEmpresa.NombreEmpresaOficial & Space(40), 1, 40)
    Aux = Aux & Format(Fecha, "yyyymmdd") & RellenaABlancos(vEmpresa.NIF, True, 9)
    
    Aux = Aux & "65B" & miBanco & CodigoOrdenante & "KF01" & "EUR" & Space(198)
    
    
    Print #NFich, Aux
    
    
    
    
    Set Rs = New ADODB.Recordset
    
    
    
    'Imprimimos las lineas
    cad = "Abriendo RS"
    Aux = "Select pagos.*,maidatos ,telefonocta from pagos left join cuentas on cuentas.codmacta=pagos.codmacta"
    Aux = Aux & " where  nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum = " & DBSet(vAnyoTransferencia, "N")
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not Rs.EOF
                '*********************************************************
                'Suposicion 1,. TODOS son nacionales
                '*********************************************************
               
                Im = 0
                Im = Rs!ImpEfect - Im
                
                cad = "Identificador"
                Aux = "2" & RellenaABlancos(Rs!codmacta, True, 15)
                Aux = Aux & "01" & RellenaABlancos(Rs!NifProve, True, 12)
                
                Aux = Aux & "T" 'transferencia
                
                cad = "Cuenta abono prov"
                Aux = Aux & Mid(Mid(Rs!IBAN, 5) & Space(20), 1, 20) 'CCC
                
                'Numero e importe factura
                Aux = Aux & RellenaABlancos(Rs!numfactu, True, 15)
                Aux = Aux & RellenaAceros(CStr(Abs(Im) * 100), False, 14)
                Aux = Aux & IIf(Im < 0, "-", "+")
                
                '495
                cad = "Fecha factura-vto"
                Aux = Aux & Format(Rs!FecFactu, "yyyymmdd") & Format(Rs!fecefect, "yyyymmdd")
                
                
                Aux = Aux & Space(30) & "N" & Space(16)   'Posiciones  98 a 137
                Aux = Aux & "N"
                
                
                '414
                cad = "IBAN"
                Aux = Aux & RellenaABlancos(Rs!IBAN, True, 30)
                
                Aux = Mid(Aux & Space(300), 1, 300)
                        
                Print #NFich, Aux
                
                
                
                'Datos complemntarios
                cad = "Datos complementarios"
                Aux = "3"
                cad = "NOmbre prov"
                Aux = Aux & RellenaABlancos(Rs!nomprove, True, 40)
                Aux = Aux & "08"
                
                Aux = Aux & RellenaABlancos(DBLet(Rs!domprove, "T"), True, 67)
                'Pos:223
                Aux = Aux & RellenaABlancos(DBLet(Rs!pobprove, "T"), True, 40)
                'Pos:248
                Aux = Aux & RellenaABlancos(DBLet(Rs!cpprove, "T"), True, 5)
                Aux = Aux & Space(6) 'reservador la 156
                Aux = Aux & RellenaABlancos(DBLet(Rs!telefonocta, "T"), True, 15)
                Aux = Aux & RellenaABlancos("", True, 15)
                
                Aux = Aux & RellenaABlancos(DBLet(Rs!maidatos, "T"), True, 60)
                Aux = Aux & "1" 'correo
                Aux = Aux & "ES"
                Aux = Aux & Space(46) 'pais residencia y reservado
                
               
                
                Print #NFich, Aux
                
               'Totales
               Importe = Importe + Im
               Regs = Regs + 1
               Rs.MoveNext
        Wend
        'Imprimimos totales
        Aux = "5"
        Aux = Aux & RellenaABlancos(vEmpresa.NIF, True, 9)
        Aux = Aux & RellenaAceros(CStr(Regs), False, 7)
        Aux = Aux & RellenaAceros(CStr(Importe * 100), False, 14) & "+"
        Aux = Aux & Space(268)
        Print #NFich, Aux
        
        
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    NFich = -1
    If Regs > 0 Then
        GeneraFicheroSabadellConfirming = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description, cad
     If NFich > 0 Then Close (NFich)
End Function






















'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'
'
'
'
'            SSSSSS         EEEEEEEE             PPPPPPP                 A
'           SS              EE                   PP     P               A A
'            SS             EE                   PP     P              A   A
'              SSS          EEEEEEEE             PPPPPPP              AAAAAAA
'                SS         EE                   PP                  A       A
'               SS          EE                   PP                 A         A
'           SSSSS           EEEEEEEE             PP                A           A
'
'
'
'
'
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
', DatosExtra As String
' N19Punto19  -> True.  19.14
'             -> False. 19.44


'SEPA XML:   Hay un modulo donde genera el fichero. Las comprobaciones iniciales son las mismas
'             para ambos modulos
'
' En funcion del parametro llamara a uno u a otro

'Si viene FECHACOBRO es que todos los vencimientos van a esa FECHA
'       si no , cada vto lleva su fecha
' Arupa VTOS .
'  ResumidoIDFichero : en la etiqueta del XML ID grabamos texto corto
Private Function GrabarFicheroNorma19SEPA(NomFichero As String, Remesa As String, FecPre As String, TipoReferenciaCliente As Byte, Sufijo As String, FechaCobro As String, SEPA_EmpresasGraboNIF As Boolean, Norma19_15 As Boolean, FormatoXML As Boolean, esAnticipoCredito As Boolean, IdGrabadoEnFichero As String, AgruparVtos As Boolean) As Boolean
Dim B As Boolean


    '-- Genera_Remesa: Esta funci�n genera la remesa indicada, en el fichero correspondiente

    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim NifEmpresa_ As String
    
    '-- Primero comprobamos que la remesa no haya sido enviada ya
    Sql = "SELECT * FROM remesas,bancos WHERE codigo = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyo = " & RecuperaValor(Remesa, 2) & " AND remesas.codmacta = bancos.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If miRsAux!Situacion >= "C" Then
            MsgBox "La remesa ya figura como enviada", vbCritical
            
        Else
            'Cargo algunos de los datos de la remesa
            DatosBanco = miRsAux!IBAN
            
             'En datos extra dejo el CONCEPTO PPAL
             'DatosExtra = RecuperaValor(DatosExtra, 2)
        End If
    Else
        MsgBox "La remesa solicitada no existe", vbCritical
    End If
    miRsAux.Close
    
    If DatosBanco = "" Then Exit Function
    
    If Not comprobarCuentasBancariasRecibos_NIF(Remesa) Then Exit Function




    'Si es el campo referencia del fichero de cobros, entonces hay que comprobar que es obligado
    If TipoReferenciaCliente = 2 Then
        'Campo REFERENCAI como identificador
        If Not ComprobarCampoReferenciaRemesaNorma19(Remesa) Then Exit Function
    End If


    'Ahora cargare el NIF y la empresa
    Sql = "Select * from empresa2"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NifEmpresa_ = ""
    If Not miRsAux.EOF Then
        NifEmpresa_ = DBLet(miRsAux!nifempre, "T")
    End If
    miRsAux.Close
    If NifEmpresa_ = "" Then
        MsgBox "Datos empresa MAL configurados", vbExclamation
        Exit Function
    End If
    
    'Desde aqui, cada norma sigue su camino, generando un fichero al final
    
    If FormatoXML Then
        
        'El que habia
        B = GrabarDisketteNorma19_SEPA_XML(NomFichero, Remesa, FecPre, TipoReferenciaCliente, Sufijo, FechaCobro, SEPA_EmpresasGraboNIF, Norma19_15, DatosBanco, NifEmpresa_, esAnticipoCredito, IdGrabadoEnFichero, AgruparVtos)
        
    End If
    GrabarFicheroNorma19SEPA = B
End Function





'miRsAux no lo paso pq es GLOBAL
'TipoRegistro
'   0: Cabecera deudor
'   1. Total deudor/FECHA
'   2. Total deudor
'   3. Total general
Private Sub ImprimiSEPA_ProveedorFecha2(TipoRegistro As Byte, IdDeudorAcreedor As String, Fecha As Date, Registros003 As Integer, Suma As Currency, NumeroLineasTotalesSinCabceraPresentador As Integer, IdNorma As String)
Dim cad As String

    Select Case TipoRegistro
    Case 0
        'Cabecera de ACREEDOR-FECHA
        cad = "02" & IdNorma & "002"   '19143-> Podria ser 19154 ver pdf
        cad = cad & IdDeudorAcreedor
        
        'Fecha cobro
        cad = cad & Format(miRsAux!FecVenci, "yyyymmdd")
        
        'Nomprove
        cad = cad & DatosBasicosDelAcreedor
        'EN SQL llevamos el IBAN completo del acredor, es decir, de la empresa presentardora que le deben los deudores
        cad = cad & Sql & Space(10)  'El iban son 24 y dejan hasta 34 psociones
        '
        cad = cad & Space(301)
        
    Case 1
        'total x fecha -deudor
        cad = "04"
        cad = cad & IdDeudorAcreedor

        'Fecha cobro
        cad = cad & Format(Fecha, "yyyymmdd")

        cad = cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        cad = cad & Format(Registros003, "00000000")
        cad = cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") ' +cabecera y pie
        cad = cad & FrmtStr(" ", 520) ' LIBRE

        
        
    Case 2
        'total deudor
        cad = "05"
        cad = cad & IdDeudorAcreedor

        cad = cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        cad = cad & Format(Registros003, "00000000")
        cad = cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") '
        cad = cad & FrmtStr(" ", 528) ' LIBRE
      
    Case 3
        'total general
        cad = "99"
        cad = cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        cad = cad & Format(Registros003, "00000000")
        cad = cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") ' +cabecera y pie
        cad = cad & FrmtStr(" ", 563) ' LIBRE
      
    End Select
        
    Print #NF, cad
        
        
End Sub

' AT-09.   70 + 50 + 50 + 40 +2
Private Function DatosBasicosDelDeudor() As String
        DatosBasicosDelDeudor = FrmtStr(miRsAux!Nommacta, 70)
        'dirdatos,codposta,despobla,pais desprovi
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(DBLet(miRsAux!dirdatos, "T"), 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!desPobla, "T")), 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(DBLet(miRsAux!desProvi, "T"), 40)
        
        If IsNull(miRsAux!PAIS) Then
            DatosBasicosDelDeudor = DatosBasicosDelDeudor & "ES"
        Else
            DatosBasicosDelDeudor = DatosBasicosDelDeudor & Mid(miRsAux!PAIS, 1, 2)
        End If
End Function


'NUestros datos basicos
' AT-09.   70 + 50 + 50 + 40 +2
Private Function DatosBasicosDelAcreedor() As String
Dim RN As ADODB.Recordset

        'NO PUEDE SER EOF
        Set RN = New ADODB.Recordset
        RN.Open "Select * from empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText


        'siglasvia direccion  numero puerta  codpos poblacion provincia


        DatosBasicosDelAcreedor = FrmtStr(vEmpresa.nomempre, 70)
        'dirdatos,codposta,despobla,pais desprovi
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & FrmtStr(Trim(DBLet(RN!siglasvia, "T") & " " & DBLet(RN!Direccion, "T") & ", " & DBLet(RN!numero, "T") & " " & DBLet(RN!puerta, "T")), 50)
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & FrmtStr(Trim(DBLet(RN!codpos, "T") & " " & DBLet(RN!Poblacion, "T")), 50)
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & FrmtStr(DBLet(RN!provincia, "T"), 40)
         
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & "ES"
        
        
        RN.Close
        Set RN = Nothing
End Function





Private Sub ImprimeEnXML(Anidacion As Byte, Fich As Integer, Etiqueta As String)

End Sub











'---------------------------------------------------------------------
'  DEVOLUCION FICHERO  SEPA
'---------------------------
Public Sub ProcesaCabeceraFicheroDevolucionSEPA(Fichero As String, ByRef Remesa As String)
Dim aux2 As String  'Para buscar los vencimientos
Dim FinLecturaLineas As Boolean
Dim TodoOk As Boolean
Dim ErroresVto As String
Dim Cuantos As Integer
Dim Bien As Integer
Dim LinDelFichero As Collection


    On Error GoTo eProcesaCabeceraFicheroDevolucionSEPA
    Remesa = ""
    
    
    
    
    'ANTES nov 2012
    '
    'nf = FreeFile
    'Open Fichero For Input As #nf
    Registro = Fichero 'para no pasr mas variables al proceso
    ProcesoFicheroDevolucion 0, LinDelFichero 'abrir el fichero y volcarlo sobre un Collection
    
    'Proceso la primera linea. A veriguare a que norma pertenece
    ' y hallare la remesa
    'Line Input #nf, Registro
    ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
    
    'Comproamos ciertas cosas
    Sql = "Linea 1 vacia"
    If Registro <> "" Then
        
        
        
        'Tiene valor
        If Len(Registro) <> 600 Then
            Sql = "Longitud linea incorrecta(600)"
        Else
            'Febrero 2014
            'Devolucion:2119
            'Rechazo:   1119
            'Antes: Mid(Registro, 1, 4) <> "2119"
            
            If Mid(Registro, 2, 3) <> "119" Then
                Sql = "Cadena control incorrecta(?119)"
            Else
                Sql = ""
            End If
        End If
    End If
    
    If Sql = "" Then
    
        'Segunda LINEA.
        'Line Input #nf, Registro
        ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
        
        Sql = "Linea 2 vacia"
        If Registro <> "" Then
            
           
            
            
            'Tiene valor
            If Len(Registro) <> 600 Then
                Sql = "Longitud linea incorrecta(600)"
            Else
                'Devolucion:2219
                'Rechazo:   1119
                'Antes: Mid(Registro, 1, 4) <> "2119"
                
                If Mid(Registro, 2, 3) <> "219" Then
                    Sql = "Cadena control incorrecta(?219)"
                Else
                    
                    Sql = "Falta linea 2319"  'la que lleva los vtos
                    Remesa = ""
                    Do
                        ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                        
                        If Registro <> "" Then
                            Sql = "FIN LINEAS. No se ha encontrado linea: 2319"
                            Remesa = "NO"
                        Else
                            'Line Input #nf, Registro
                            ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            
                            'BUsco la linea:
                            '5690
                            If Registro <> "" Then
                                '2319  Lleva los vtos
                                '1319 en devoluciones
                                If Mid(Registro, 2, 3) = "319" Then
                                    Sql = ""
                                    Remesa = "NO"
                                End If
                            End If
                        End If
                        
                    Loop Until Remesa <> ""
                    Remesa = ""
                    
                    If Sql = "" Then
                        'VAMOS BIEN. Veremos si a partir de los datos del recibo nos dan la remesa
                        'Para ello bucaremos en registro, la cadena que contiene los datos
                        'del vencimiento
                        'Registro=
                        '2319143003430000061 M  0330047820131201001   430000061 M  0330047820131201001
                        'sigue arriba RCURTRAD0000001210020091031CCRIES2AXXXCOANNA, COOP. V.                                                      CAMINO HONDO, 1                                   46820                                                                                     ES1IF46024493                          F46024493                          AES1830820134930330000488          TRADFACTURA: M-3300478 de Fecha 01 dic 2013                                                                                                     MD0120131230
                        Set miRsAux = New ADODB.Recordset
                        ErroresVto = ""
                        FinLecturaLineas = False
                        Cuantos = 0
                        Bien = 0
                        Do
                            'Devolucion:2319
                            'Rechazo:   1319
                            'Antes: Mid(Registro, 1, 4) <> "2119"
            
                            If Mid(Registro, 2, 3) = "319" Then
                                'Los vtos vienen en estas lineas
                                Cuantos = Cuantos + 1
                                Registro = Mid(Registro, 21, 35)
                                'M  0330047820131201001
                                Sql = "Select codrem,anyorem,siturem from cobros where fecfactu='" & Mid(Registro, 12, 4) & "-" & Mid(Registro, 16, 2) & "-" & Mid(Registro, 18, 2)
                                
                                Sql = Sql & "' AND numserie = '" & Trim(Mid(Registro, 1, 3)) & "' AND numfactu = " & Val(Mid(Registro, 4, 8)) & " AND numorden=" & Mid(Registro, 20, 3)
                                
                                
                                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                TodoOk = False
                                Sql = "Vencimiento no encontrado: " & Registro
                                If Not miRsAux.EOF Then
                                    If IsNull(miRsAux!Codrem) Then
                                        Sql = "Vencimiento sin Remesa: " & Registro
                                    Else
                                        Sql = miRsAux!Codrem & "|" & miRsAux!Anyorem & "|�"
                                        
                                        If InStr(1, Remesa, Sql) = 0 Then Remesa = Remesa & Sql
                                        Sql = ""
                                        TodoOk = True
                                    End If
                                End If
                                miRsAux.Close
                                
                               
                                
                                
                                
                                If Sql <> "" Then
                                    ErroresVto = ErroresVto & vbCrLf & Sql
                                Else
                                    Bien = Bien + 1
                                End If
                            Else
                                'La linea no empieza por 569
                                'veremos los totales
                                
                                If Mid(Registro, 1, 2) = "99" Then
                                    'TOTAL TOTAL
                                    Sql = Mid(Registro, 20, 8)
                                    If Val(Sql) <> Cuantos Then ErroresVto = "Fichero: " & Sql & "   Leidos" & Cuantos & vbCrLf & ErroresVto & vbCrLf & Sql
                                End If
                            End If
                            
                            'Siguiente linea
                            ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                            
                            If Registro <> "" Then
                                FinLecturaLineas = True
                            Else
                                'Line Input #nf, Registro
                                ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            End If
                            
                        Loop Until FinLecturaLineas
                        
                        If Cuantos <> Bien Then ErroresVto = ErroresVto & vbCrLf & "Total: " & Cuantos & "   Correctos:" & Bien
                        
                        Sql = ErroresVto
                        Set miRsAux = Nothing
                    
                    End If
                End If  'Control SEGUNDA LINEA
        
        
            End If
        End If
    
    End If  'DE SEGUNDA LINEA
    
    ProcesoFicheroDevolucion 4, LinDelFichero
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
    Else
        'Remesa = Mid(Registro, 1, 4) & "|" & Mid(Registro, 5) & "|"
        
        
        'Ahora comprobaremos que para cada remesa  veremos si existe y si la situacion es la contabilizadxa
        Sql = Remesa
        Registro = "" 'Cadena de error de situacion remesas
        Set miRsAux = New ADODB.Recordset
        Do
            Cuantos = InStr(1, Sql, "�")
            If Cuantos = 0 Then
                Sql = ""
            Else
                aux2 = Mid(Sql, 1, Cuantos - 1)
                Sql = Mid(Sql, Cuantos + 1)
                
                
                'En aux2 tendre codrem|an�orem|
                aux2 = RecuperaValor(aux2, 1) & " AND anyo = " & RecuperaValor(aux2, 2)
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
                    aux2 = Replace(aux2, "anyo", "a�o")
                    Registro = Registro & vbCrLf & aux2
                End If
                miRsAux.Close
            End If
        Loop Until Sql = ""
        Set miRsAux = Nothing
        
        
        If Registro <> "" Then
            Registro = "Error remesas " & vbCrLf & String(30, "=") & Registro
            MsgBox Registro, vbExclamation
            
            'Pongo REMESA=""
            Remesa = "" 'para que no continue el preoceso de devolucion
        End If
        
    End If
    
    Exit Sub
eProcesaCabeceraFicheroDevolucionSEPA:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion SEPA"
End Sub




Public Function EsFicheroDevolucionSEPA2(elpath As String) As Byte
Dim NF As Integer

    On Error GoTo eEsFicheroDevolucionSEPA
    EsFicheroDevolucionSEPA2 = 0   'N19 Antiquisima      1.- SEPA txt    2 SEPA xml
    NF = FreeFile
    Open elpath For Input As #NF
    If Not EOF(NF) Then
        Line Input #NF, Sql
        If Sql <> "" Then
            '                 DEVOLUCION                RECHAZO
            If LCase(Mid(Sql, 1, 5)) = "<?xml" Then
                EsFicheroDevolucionSEPA2 = 2
            Else
                If Mid(Sql, 1, 2) = "21" Or Mid(Sql, 1, 2) = "11" Then
                    EsFicheroDevolucionSEPA2 = 1
                Else
                    EsFicheroDevolucionSEPA2 = 0
                End If
            End If
        End If
    End If
    Close #NF
eEsFicheroDevolucionSEPA:
    Err.Clear
End Function



'******************************************************************************************************************
'******************************************************************************************************************
'
'       Genera fichero NORMA 68
'
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma68(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String
Dim PagosJuntos As Boolean
Dim pagosAux As Currency
    On Error GoTo EGen
    GeneraFicheroNorma68 = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 9)
    Aux = Mid(CIF & Space(10), 1, 9)
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!IBAN) Then
            cad = ""
        Else
            
            'CodigoOrdenante = Mid(DBLet(Rs!IBAN, "T"), 5, 20) 'Format(Rs!Entidad, "0000") & Format(DBLet(Rs!oficina, "N"), "0000") & Format(DBLet(Rs!Control, "N"), "00") & Format(DBLet(Rs!CtaBanco, "T"), "0000000000")
            'If Not DevuelveIBAN2("ES", CodigoOrdenante, Cad) Then Cad = ""
            'CuentaPropia = "ES" & Cad & CodigoOrdenante
            
            CuentaPropia = Rs!IBAN
            
            
        End If
        
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma68.txt" For Output As #NFich
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 9)  'CIF EMPRESA
    CodigoOrdenante = CodigoOrdenante & "000" 'el sufijo
    
    'CABECERA
    Cabecera1_68 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
   
    Aux = DevuelveDesdeBDNew(cConta, "transferencias", "concepto", "codigo", CStr(vNumeroTransferencia), "N", , "anyo", vAnyoTransferencia, "N")
    PagosJuntos = Aux = "1"
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set Rs = New ADODB.Recordset
    Aux = "Select pagos.*,nommacta,dirdatos,codposta,dirdatos,despobla,nifdatos,razosoci,desprovi,cuentas.codpais from pagos,cuentas"
    Aux = Aux & " where pagos.codmacta=cuentas.codmacta and nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum =" & vAnyoTransferencia
    Rs.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hay ningun registro
        
    Else
    
        '-----------------------------
        
         
        
        Regs = 0
        While Not Rs.EOF
            
                'Junio18
                'Im = DBLet(Rs!imppagad, "N")
                Im = 0
                Im = Rs!ImpEfect - Im
                Aux = RellenaABlancos(Rs!nifdatos, True, 12)
            

            
            
            Aux = "06" & "59" & CodigoOrdenante & Aux   'Ordenante y nifprove
        
            Linea1_68 NFich, Aux, Rs, cad
            Linea2_68 NFich, Aux, Rs, cad
            Linea3_68 NFich, Aux, Rs, cad
            Linea4_68 NFich, Aux, Rs, cad
            
            
            
            '13/04/2016
              
            pagosAux = CCur("9000" & Format(Regs + 1, "0000000")) 'concatenamos 9000 con el numero de pago domiciliado
            AuxD = pagosAux - (Int(pagosAux / 7) * 7)
            AuxD = Format(Regs + 1, "0000000") & AuxD 'cargamos el numero del pago para NO pasarlo a las lineas 5  6

            
            
            'Antes
            'Linea5_68 NFich, AUX, RS, Cad, Fecha, Im
            'Ahora en funcion de si los queremos todos juntos o cada uno a su vto
            
            Linea5_68 NFich, Aux, Rs, cad, IIf(PagosJuntos, Fecha, Rs!fecefect), Im
            
            
            Linea6_68 NFich, Aux, Rs, Im, cad, ConceptoTr
            'If Pagos Then Linea7 NFich, Aux, RS, Cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            Rs.MoveNext
        Wend
        'Imprimimos totales
        Totales68 NFich, CodigoOrdenante, Importe, Regs, cad
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    If Regs > 0 Then
        GeneraFicheroNorma68 = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function


Private Sub Cabecera1_68(NF As Integer, ByRef CodOrde As String, Fecha As Date, IBAN As String, ByRef cad As String)

    cad = "03"
    cad = cad & "59"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "001"
    
    cad = cad & Format(Fecha, "ddmmyy")
    
    'Cuenta bancaria
    cad = cad & Space(9)
    cad = cad & IBAN
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub







Private Sub Linea1_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "010"
    If IsNull(RS1!razosoci) Then
        cad = cad & RellenaABlancos(RS1!Nommacta, True, 40)
    Else
        cad = cad & RellenaABlancos(RS1!razosoci, True, 40)
    End If
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub


Private Sub Linea2_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "011"
    cad = cad & RellenaABlancos(DBLet(RS1!dirdatos, "T"), True, 45)
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub





Private Sub Linea3_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "012"
    cad = cad & RellenaABlancos(DBLet(RS1!codposta, "T"), False, 5)
    cad = cad & RellenaABlancos(DBLet(RS1!desPobla, "T"), True, 40)
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub

Private Sub Linea4_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "013"
    'De mommento pongo balancos, ya que es para extranjero
    cad = cad & RellenaABlancos(DBLet(RS1!codposta, "T"), True, 9)
    cad = cad & RellenaABlancos(DBLet(RS1!desProvi, "T"), True, 30)   'desprovi,pais
    cad = cad & RellenaABlancos(DBLet(RS1!codpais, "T"), True, 20)   'desprovi,pais
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub

' Febrero 2016.
' En la cabecera llevamos si queremos todos los pagos a una fecha o cada uno en su vencimiento
' con lo cual aqui siempre enviaremos el valor fecha que ya llevara uno u otro
Private Sub Linea5_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String, ByRef Fechapag As Date, ByRef Importe1 As Currency)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "014"

    cad = cad & AuxD '13/04/16  Lo cargamos antes de recorrer el RS
    
    cad = cad & Format(Fechapag, "ddmmyyyy")
    'Cad = Cad & Format(RS1!Fecefect, "ddmmyyyy") 'fecha vencimiento de cada recibo   'YA VIENEN CARGADA en fecha doc lo que corresponda
   
    cad = cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    cad = cad & "0" 'presentacion
    'Cad = Cad & "ES1" 'presentacion
    cad = cad & "   " 'En el pdf pone que el pais es para NO residentes
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 99) & " "   'Antes ponia un 1. Mayo16'
    Print #NF, cad
End Sub


Private Sub Linea6_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Importe1 As Currency, ByRef cad As String, vConceptoTransferencia As String)


   
    '
    cad = CodOrde   'llevara tb la ID del socio
    cad = cad & "015"
    cad = cad & AuxD 'Numero de pago domiciliado
    cad = cad & RellenaABlancos(RS1!numfactu, False, 12)
    cad = cad & Format(RS1!FecFactu, "ddmmyyyy") 'fecha fac

    cad = cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    
    cad = cad & "H"
    'Cad = Cad & RellenaABlancos(vConceptoTransferencia, False, 26)
    cad = cad & "PAGO FACTURA   " & RS1!numfactu
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub



Private Sub Totales68(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef cad As String)
    cad = "08" & "59"
    cad = cad & CodOrde    'llevara tb la ID del socio
    cad = cad & Space(15)
    cad = cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    'Cad = Cad & RellenaAceros(CStr(Registros), False, 8)
    cad = cad & RellenaAceros(CStr((Registros * 6) + 1 + 1), False, 10)
    cad = RellenaABlancos(cad, True, 100)
    Print #NF, cad
End Sub


'**************************************************************************************************************
'**************************************************************************************************************
'
'   CONFIRMING estandar.
'
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************

Public Function GeneraFicheroConfirmingSt(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante2 As String
Dim Importe As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String
Dim NifProve As String
Dim IbanPRov As String   'Para garantizar que estan todos lo de un proveedor al mismo  IBAN
Dim Fin As Boolean


    On Error GoTo EGen
    GeneraFicheroConfirmingSt = False
    
    NumeroTransferencia = vNumeroTransferencia
    
        
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodigoOrdenante2 = "000"
    Aux = Right("    " & CIF, 9)
    Aux = Mid(CIF & Space(10), 1, 9)
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!IBAN) Then
            cad = ""
        Else
            CuentaPropia = Rs!IBAN
             
        End If
        If Not IsNull(Rs!sufijoconfirming) Then CodigoOrdenante2 = Mid(Rs!sufijoconfirming & CodigoOrdenante2, 1, 3)
    End If
    Rs.Close
    
    If cad = "" Then Err.Raise 513, , "Error leyendo datos para: " & CuentaPropia
        
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    
    Aux = Right(Space(10) & CIF, 10)
    CIF = Aux
    
    'CABECERA
    '-----------------------------------------------------
    
    CodigoOrdenante2 = "13" & "70" & CIF & CodigoOrdenante2 & Space(9)    '000-> sufijo      Libre(9)
        
        
    '1er registro 001
    Aux = DevuelveDesdeBDNew(cConta, "transferencias", "fecha", "codigo", CStr(vNumeroTransferencia), "N", , "anyo", vAnyoTransferencia, "N")
    Aux = CodigoOrdenante2 & "001" & Format(Now, "ddmmyyyy") & Format(Aux, "ddmmyyyy")
    Aux = Aux & Mid(CuentaPropia, 5, 4) & Mid(CuentaPropia, 9, 4) & Mid(CuentaPropia, 15, 10) & " " & "  "   'F7 dos carcateres alfa pactados con
    Aux = Aux & " " & Mid(CuentaPropia, 13, 2) & Space(3)
    Print #NFich, Aux
        
    ' Registro 2 al 4
    cad = DBSet(vEmpresa.nomempre, "T")
                '   noempres              direccion                       codposta prob prov
    cad = "select " & cad & " , concat(direccion,' ',numero) , concat(codpos,' ',poblacion,' ',provincia) From empresa2"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then Err.Raise 513, , "Error leyendo datos empresa"
    
    For i = 0 To 2
        Aux = DBLet(Rs.Fields(i), "T")
        Aux = CodigoOrdenante2 & Format(i + 2, "000") & RellenaABlancos(Aux, True, 36) & Space(7)
        Print #NFich, Aux
    Next
    Rs.Close
    Regs = 4  'Numero total de registros (incluidos cabecera y pie). Llevamos 4 ya
    
    'Imprimimos las lineas
    
    Aux = "Select pagos.*,nommacta,dirdatos,codposta,dirdatos,despobla,nifdatos,razosoci,desprovi,cuentas.codpais from pagos,cuentas"
    Aux = Aux & " where pagos.codmacta=cuentas.codmacta and nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum =" & vAnyoTransferencia
    Aux = Aux & " ORDER BY nifprove"
    Rs.Open Aux, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Importe = 0
    If Rs.EOF Then
        'No hay ningun registro
        
    Else
        
        NifProve = ""
        J = 0  'Numero total de 010
        While Not Rs.EOF
            
            'PARA CADA PROVEEDOR
            If NifProve <> Rs!NifProve Then
                
                
                'Vamos a ver el importe total del proveedor, y comprobaremos su IBAN
                Im = 0
                Fin = False
                cad = Rs!NifProve 'A tratar
                While Not Fin
                    Im = Im + Rs!ImpEfect
                    Rs.MoveNext
                    If Rs.EOF Then
                        Fin = True
                    Else
                        If Rs!NifProve <> cad Then Fin = True
                    End If
                Wend
                'Dejamos el cursor(recordset) andestaba
                If NifProve = "" Then
                    'Es el primer proveedor
                    Rs.MoveFirst
                Else
                    Do
                        Rs.MovePrevious
                    Loop Until Rs!NifProve = NifProve  'encontramos el ultimo anterior
                    Rs.MoveNext
                End If
                
                cad = RellenaABlancos(Rs!nifdatos, True, 12)
                CodigoOrdenante2 = "16" & "70" & CIF & cad   '26 comunes a todas las lineas
                
                'Registro bene 1
                Aux = CStr(Im * 100)
                cad = CodigoOrdenante2 & "010" & RellenaAceros(Aux, False, 12)
                
                If DBLet(Rs!IBAN, "T") = "" Then Err.Raise 513, , "IBAN incorrecto.  Factura " & Rs!numfactu & "   " & Rs!nomprove
                Aux = Mid(Rs!IBAN, 5, 4) & Mid(Rs!IBAN, 9, 4) & Mid(Rs!IBAN, 15, 10)
                cad = cad & Aux & "2" & "T" & " " & Mid(Rs!IBAN, 13, 2) & Space(8)  'F5 , F6 ,F7,F8
                Print #NFich, cad
                
                'Registro bene 2
                Print #NFich, CodigoOrdenante2 & "011" & RellenaABlancos(Rs!nomprove, True, 36) & Space(7)
                'Registro bene 3
                Print #NFich, CodigoOrdenante2 & "012" & RellenaABlancos(Rs!domprove, True, 36) & Space(7)
                'Registro bene  5
                Aux = DBLet(Rs!cpprove, "T") & " " & DBLet(Rs!pobprove, "T")
                Print #NFich, CodigoOrdenante2 & "014" & RellenaABlancos(Aux, True, 36) & Space(7)
                Regs = Regs + 4
                'Opcional 6 proveincia
                If DBLet(Rs!proprove, "T") <> "" Then
                    If Rs!proprove <> DBLet(Rs!pobprove, "T") Then
                        'Provincia distionto pioblacion
                        cad = CodigoOrdenante2 & "015" & RellenaABlancos(Rs!proprove, True, 26)
                        cad = cad & RellenaAceros(Rs!NifProve, False, 10) & Space(7)
                        Print #NFich, cad
                        Regs = Regs + 1
                    End If
                End If
                    
                
                cad = RellenaABlancos(Rs!nifdatos, True, 12)
                CodigoOrdenante2 = "17" & "70" & CIF & cad   '26 comunes a todas las lineas de factura
                
                NifProve = Rs!NifProve
                IbanPRov = Rs!IBAN
                J = J + 1 'Total de beneficiarios
                i = 0
            End If
            'Comprobacion
            If IbanPRov <> Rs!IBAN Then Err.Raise 513, , "Iban distinto: " & Rs!nomprove
            'Contadores
            Im = Rs!ImpEfect
            Importe = Importe + Im
            Regs = Regs + 1
            i = i + 1 'Como mucho podemos pagar 275 vtos
            Aux = CStr(Im * 100)
            
        
            
            cad = CodigoOrdenante2 & Format(i, "000")
            cad = cad & Format(Rs!FecFactu, "ddmmyyyy") & Format(Rs!fecefect, "ddmmyyyy")
            cad = cad & RellenaABlancos(Rs!numfactu, True, 14) & RellenaAceros(Aux, False, 12) & IIf(Im < 0, "-", " ")
            Print #NFich, cad
            
            Rs.MoveNext
        Wend
        
        'Imprimimos totales
        Regs = Regs + 1
        cad = "18" & "70" & CIF & Space(12) & Space(3)
        Aux = CStr(Importe * 100)
        cad = cad & RellenaAceros(Aux, False, 12) & RellenaAceros(CStr(J), False, 8) & RellenaAceros(CStr(Regs), False, 10) & Space(6) & Space(7)
        Print #NFich, cad
    End If
    Rs.Close
    Set Rs = Nothing
    Close (NFich)
    If Regs > 0 Then
        GeneraFicheroConfirmingSt = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
    IntentaCErrar NFich
End Function

Private Sub IntentaCErrar(NumeroFichero As Integer)
 On Error Resume Next
 Close (NumeroFichero)
 Err.Clear
End Sub





'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
' CAIXA RURAL
'
Public Function GeneraFicheroCaixaRural(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Impor2 As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String
Dim RefereProve As String
Dim colVtos As Collection
Dim Fin As Boolean
Dim IBAN As String
Dim NumeroDePago As Long

    On Error GoTo EGen
    GeneraFicheroCaixaRural = False
    
    NumeroTransferencia = vNumeroTransferencia
    NFich = -1
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    Aux = Mid(CIF & Space(10), 1, 9)
    If Rs.EOF Then
        cad = ""
    Else
       
        CodigoOrdenante = Mid(DBLet(Rs!sufijoconfirming, "T") & Space(3), 1, 3)
        IBAN = Rs!IBAN
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    
    'Resgristro 1 de cabecera

    CodigoOrdenante = Aux & CodigoOrdenante
   
    Aux = "0359" & CodigoOrdenante
    Aux = Aux & Space(12) & "001"
    'Aux = Aux & Mid(vEmpresa.NombreEmpresaOficial & Space(40), 1, 40)
    Aux = Aux & Format(Fecha, "ddmmyy") & Space(9)
    Aux = Aux & IBAN
    Aux = Mid(Aux & Space(100), 1, 100)
    Print #NFich, Aux
    Regs = 1
    
    
    
    Set Rs = New ADODB.Recordset
    
    
    
    'Imprimimos las lineas
    cad = "Abriendo RS"
    Aux = "Select pagos.*,maidatos ,telefonocta from pagos left join cuentas on cuentas.codmacta=pagos.codmacta"
    Aux = Aux & " where  nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum = " & DBSet(vAnyoTransferencia, "N")
    Aux = Aux & " ORDER BY nifprove"
    Rs.Open Aux, Conn, adOpenKeyset, adLockPessimistic, adCmdText   'Puede ir ir hacia adelante o atras
    Importe = 0
    RefereProve = ""
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        
        While Not Rs.EOF
                '*********************************************************
                'Suposicion 1,. TODOS son nacionales
                '*********************************************************
               
                
                RefereProve = RellenaABlancos(Rs!NifProve, True, 12)
                
                'Registro 1
                
                Aux = "0659" & CodigoOrdenante & RefereProve
                cad = Rs!nomprove
                FormatearTextoParaInformativas cad
                Aux = Aux & "010" & RellenaABlancos(cad, True, 40)
                Aux = Aux & Space(29)
                Print #NFich, Aux
                
                
                
                Aux = "0659" & CodigoOrdenante & RefereProve & "011"
                cad = DBLet(Rs!domprove, "T")
                FormatearTextoParaInformativas cad
                Aux = Aux & RellenaABlancos(cad, True, 45)
                Aux = Aux & Space(24)
                Print #NFich, Aux
                
                
                Aux = "0659" & CodigoOrdenante & RefereProve & "012"
                Aux = Aux & Mid(Format(DBLet(Rs!cpprove, "N"), "00000"), 1, 5)
                Aux = Aux & RellenaABlancos(DBLet(Rs!pobprove, "T"), True, 40)
                Aux = Aux & Space(24)
                Print #NFich, Aux
                
                Aux = "0659" & CodigoOrdenante & RefereProve & "013"
                Aux = Aux & Space(9) & RellenaABlancos(DBLet(Rs!proprove, "T"), True, 30)
                Aux = Aux & RellenaABlancos("ESPA�A", True, 20)
                Aux = Aux & Space(10)
                Print #NFich, Aux
                
                Regs = Regs + 4
                
                'tipo registro 5
                'Llevara los totoales del BENEFICIIARIO
                
                ' en la linea
                ' 7 digitos + 2 de control. para cada proveedor NO puede repetirse en toda su vida.
                'Puede llevar en este fichero varios pagos juntos, pero en proximos ficheros NO se repite
                'Para ello vamos a hacer lo siguiente
                '    XXXXXXX
                '    123    -> right (codmacta priove)
                '       X   -> ultimo digito de a�o
                '        SSD-> dia del a�o   SS semana dia de seaman
                cad = Right(Rs!codmacta, 3) & Right(CStr(Year(Now)), 1) & Format(Rs!fecefect, "ww") & WeekDay(Rs!fecefect)
                Aux = Val(CadenaTextoMod97(CStr(cad)))
                
                NumeroDePago = Val(Aux)
                If NumeroDePago > 9 Then NumeroDePago = NumeroDePago \ 10
                Aux = NumeroDePago
                
                    
                NumeroDePago = CLng(cad & Aux)
                Set colVtos = New Collection
                Impor2 = 0
                Fin = False
                cad = Rs!NifProve
                
                Do
                    'Guardo linea importe del proveedor
                    Im = 0
                    Im = Rs!ImpEfect - Im
                    Impor2 = Impor2 + Im
                    Importe = Importe + Im
                    Aux = "0659" & CodigoOrdenante & RefereProve & "015" & Format(NumeroDePago, "00000000")
                    
                    Aux = Aux & RellenaABlancos(Right(Rs!codmacta, 5) & Rs!numfactu & Rs!numorden, False, 12)
                    Aux = Aux & Format(Rs!FecFactu, "ddmmyyyy")
                    Aux = Aux & RellenaAceros(Replace(CStr(Abs(Im)), ",", ""), False, 12)
                    Aux = Aux & IIf(Im < 0, "D", "H")
                    cad = Mid(Rs!numfactu & " " & Rs!FecFactu & " " & Rs!numorden, 1, 26)
                    Aux = Aux & RellenaABlancos(cad, True, 28)
                    colVtos.Add Aux
                    
                    Rs.MoveNext
                    If Rs.EOF Then
                        Fin = True
                    
                    Else
                        If Rs!NifProve <> cad Then Fin = True
                    End If
                    Rs.MovePrevious
                    
                Loop Until Fin
                
                'Metemos el registro 5 cabecera del pago
                Aux = "0659" & CodigoOrdenante & RefereProve & "014" & Format(NumeroDePago, "00000000")
                Aux = Aux & Format(Rs!fecefect, "ddmmyyyy")
                Aux = Aux & RellenaAceros(Replace(CStr(Abs(Impor2)), ",", ""), False, 12)
                Aux = Aux & "0  " & "      "
                Aux = Aux & Space(32)
                Print #NFich, Aux
                Regs = Regs + 1
                
                'Ahora los vencimientos
                For NumeroDePago = 1 To colVtos.Count
                    Aux = colVtos.Item(NumeroDePago)
                    Print #NFich, Aux
                    Regs = Regs + 1
                Next
                
                Aux = "0659" & CodigoOrdenante & RefereProve & "044"
                Aux = Aux & RellenaABlancos(Rs!IBAN, True, 34)
                cad = Mid(Rs!IBAN, 5, 4)
                cad = DevuelveDesdeBD("bic", "bics", "entidad", cad)
                If cad = "" Then Err.Raise 513, , "IBAN entidad: " & Mid(IBAN, 5, 4) & " Banco: " & IBAN
                Aux = Aux & RellenaABlancos(cad, True, 35)
                Print #NFich, Aux
                Regs = Regs + 1
                
               
               Rs.MoveNext
        Wend
        'Imprimimos totales
      
        
        
        Aux = "0859" & CodigoOrdenante
        Aux = Aux & Space(15)
        Aux = Aux & RellenaAceros(Replace(Importe, ",", ""), False, 12)
        Regs = Regs + 1
        Aux = Aux & RellenaAceros(CStr(Regs), False, 10)
        Aux = Mid(Aux & Space(100), 1, 100)
        Print #NFich, Aux
        
        
            
        
        
        
        
    End If
    Rs.Close
    Set Rs = Nothing
    Set colVtos = Nothing
    Close (NFich)
    NFich = -1
    If Regs > 0 Then
        GeneraFicheroCaixaRural = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description, cad
     If NFich > 0 Then Close (NFich)
     Set colVtos = Nothing
End Function









'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
' CAIXA POPULAR
'
Public Function GeneraFicheroCaixaPipular(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Impor2 As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String



    On Error GoTo EGen
    GeneraFicheroCaixaPipular = False
    
    NumeroTransferencia = vNumeroTransferencia
    NFich = -1
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    Aux = Mid(CIF & Space(10), 1, 10)
    If Rs.EOF Then
        cad = ""
    Else
        'Enpipados:   contrato confirming | sufijoconfirming
        CodigoOrdenante = DBLet(Rs!CaixaConfirming, "T") & "|" & Mid(DBLet(Rs!sufijoconfirming, "T") & Space(2), 1, 2) & "|" & Rs!IBAN & "|"
        
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    
    'Resgristro de cabecera

    'Tipo    n1clie CIF    nombre
    Aux = "1" & FrmtStr(RecuperaValor(CodigoOrdenante, 1), 8) & Aux
    Aux = Aux & FrmtStr(vEmpresa.NombreEmpresaOficial, 36) & RecuperaValor(CodigoOrdenante, 2)
    
    cad = "codigo=" & vNumeroTransferencia & " AND anyo"
    cad = DevuelveDesdeBD("importe", "transferencias", cad, CStr(vAnyoTransferencia))
    If Val(cad) <= 0 Then Err.Raise 513, , "Importe confirming CERO. " & cad & "."
    Impor2 = CCur(cad)
    
    
    Aux = Aux & Format(Fecha, "yyyymmdd") & FrmtCurren(Impor2, 11)
    Aux = Aux & "RCCR-51" & "EUR" & "N" & "N" & "   " & "        "
    cad = RecuperaValor(CodigoOrdenante, 3)
    Aux = Aux & FrmtStr(cad, 24)
    cad = FrmtStr(Aux, 414)
    Print #NFich, cad
    Regs = 0
    
    
    Set Rs = New ADODB.Recordset
    
    
    'Imprimimos las lineas
    cad = "Abriendo RS"
    Aux = "Select pagos.*,maidatos ,telefonocta from pagos left join cuentas on cuentas.codmacta=pagos.codmacta"
    Aux = Aux & " where  nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum = " & DBSet(vAnyoTransferencia, "N")
    Aux = Aux & " ORDER BY nifprove"
    Rs.Open Aux, Conn, adOpenKeyset, adLockPessimistic, adCmdText   'Puede ir ir hacia adelante o atras
    Importe = 0
    
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        
        While Not Rs.EOF
                '*********************************************************
                'Suposicion 1,. TODOS son nacionales
                '*********************************************************
                Regs = Regs + 1
                
                cad = "2" & RellenaABlancos(Rs!NifProve, True, 15)
                Aux = DBLet(Rs!nomprove, "T")
                FormatearTextoParaInformativas Aux
                cad = cad & RellenaABlancos(Aux, True, 36) & " "
                                
                Aux = DBLet(Rs!domprove, "T")
                FormatearTextoParaInformativas Aux
                cad = cad & RellenaABlancos(Aux, True, 50)
                
                Aux = DBLet(Rs!cpprove, "T")
                If Aux = "" Then Aux = "46000"
                cad = cad & RellenaABlancos(Aux, True, 15)
                
                cad = cad & RellenaABlancos(DBLet(Rs!pobprove, "T"), True, 30) & RellenaABlancos(DBLet(Rs!proprove, "T"), True, 30) & "ES"
                cad = cad & RellenaABlancos(DBLet(Rs!telefonocta, "T"), True, 15) & Space(15) & RellenaABlancos(DBLet(Rs!maidatos, "T"), True, 50) & "ES"  'pais destino
                cad = cad & RellenaABlancos(DBLet(Rs!IBAN, "T"), True, 34) & Space(11) 'BIC
                
                'Datos vto
                cad = cad & RellenaABlancos(Rs!numfactu, True, 15) & Format(Rs!FecFactu, "yyyymmdd")
                Im = Rs!ImpEfect
                Importe = Importe + Im
                cad = cad & FrmtCurren(Im, 11) & Format(Rs!fecefect, "yyyymmdd")
                
                cad = cad & Space(65)  'estadistico reservado reservado 2�vto  reservado    LENGTH 414
                
                Print #NFich, cad
                
                
                
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing
    
    Close (NFich)
    NFich = -1
    If Regs > 0 Then
        If Importe <> Impor2 Then
            MsgBox "Importes vencimientos distinto del confirming", vbExclamation
        Else
            GeneraFicheroCaixaPipular = True
        End If
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description, cad
     If NFich > 0 Then Close (NFich)
     
End Function












'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'
' BANCA MARCH
'
Public Function GeneraFicheroBancaMarch(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String, vAnyoTransferencia As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Impor2 As Currency
Dim Im As Currency
Dim Rs As ADODB.Recordset
Dim Aux As String
Dim cad As String

Dim LineaProv As String
Dim K As Byte
    'El bucle de vencimientos se hara 2 veces, 1 para proveedores y otra para los vtos de esos proveedores
    '   Reg 1       Cabecera        1 linea
    '   Reg 2       Proveedores     * 1 o mas
    '   Reg 3       Pagos       * 1 o mas
    '   REg 4       Totales     1



    On Error GoTo EGen
    GeneraFicheroBancaMarch = False
    
    NumeroTransferencia = vNumeroTransferencia
    NFich = -1
    
    'Cargamos la cuenta
    cad = "Select * from bancos where codmacta='" & CuentaPropia & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
    Aux = Mid(CIF & Space(9), 1, 9)  'CIF
    If Rs.EOF Then
        cad = ""
    Else
        'Enpipados:   contrato confirming | sufijoconfirming
        CodigoOrdenante = DBLet(Rs!CaixaConfirming, "T") & "|" & Mid(DBLet(Rs!sufijoconfirming, "T") & Space(2), 1, 2) & "|" & Rs!IBAN & "|"
        
    End If
    Rs.Close
    Set Rs = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\confirming.txt" For Output As #NFich
    
    
    
    'Resgristro de cabecera

    'Tipo    n1clie CIF    nombre
    Aux = "1" & FrmtStr(RecuperaValor(CodigoOrdenante, 1), 8) & Aux
    Aux = Aux & FrmtStr(vEmpresa.NombreEmpresaOficial, 36) & RecuperaValor(CodigoOrdenante, 2)
    
    cad = "codigo=" & vNumeroTransferencia & " AND anyo"
    cad = DevuelveDesdeBD("importe", "transferencias", cad, CStr(vAnyoTransferencia))
    If Val(cad) <= 0 Then Err.Raise 513, , "Importe confirming CERO. " & cad & "."
    Impor2 = CCur(cad)
    
    
    Aux = Aux & Format(Fecha, "yyyymmdd") & FrmtCurren(Impor2, 11)
    Aux = Aux & "RCCR-51" & "EUR" & "N" & "N" & "   " & "        "
    cad = RecuperaValor(CodigoOrdenante, 3)
    Aux = Aux & FrmtStr(cad, 24)
    cad = FrmtStr(Aux, 293)
    Print #NFich, cad
    Regs = 0
    
    
    Set Rs = New ADODB.Recordset
    
    
    'Imprimimos las lineas
    cad = "Abriendo RS"
    Aux = "Select pagos.*,maidatos ,telefonocta from pagos left join cuentas on cuentas.codmacta=pagos.codmacta"
    Aux = Aux & " where  nrodocum =" & NumeroTransferencia
    Aux = Aux & " and anyodocum = " & DBSet(vAnyoTransferencia, "N")
    Aux = Aux & " ORDER BY nifprove"
    Rs.Open Aux, Conn, adOpenKeyset, adLockPessimistic, adCmdText   'Puede ir ir hacia adelante o atras
    Importe = 0
    NumRegElim = 0 'cuantos proveedores
    Aux = ""
    cad = ""
    If Rs.EOF Then
        'No hayningun registro
        
    Else
        
        For K = 1 To 2
            Rs.MoveFirst
            Regs = 0
            
            While Not Rs.EOF
            
            
            
                    If K = 1 Then
                        'Primera pasada, PROVEEDORES
                    
                        cad = RellenaABlancos(Rs!NifProve, True, 9)
                        If cad <> Aux Then
                            NumRegElim = NumRegElim + 1
                            If Aux <> "" Then Print #NFich, LineaProv
                            Aux = cad
                         
                            
                            
                        
                                               
                            LineaProv = "2" & cad
                            cad = DBLet(Rs!nomprove, "T")
                            FormatearTextoParaInformativas Aux
                            LineaProv = LineaProv & RellenaABlancos(Aux, True, 40)
                                            
                            cad = DBLet(Rs!domprove, "T")
                            FormatearTextoParaInformativas cad
                            LineaProv = LineaProv & RellenaABlancos(cad, True, 50)
                            LineaProv = LineaProv & RellenaABlancos("", True, 8)
                            LineaProv = LineaProv & RellenaABlancos(DBLet(Rs!pobprove, "T"), True, 30) & RellenaABlancos(DBLet(Rs!proprove, "T"), True, 25)
                            LineaProv = LineaProv & Format(DBLet(Rs!cpprove, "N"), "00000") & "ES"
                            LineaProv = LineaProv & RellenaABlancos(DBLet(Rs!IBAN, "T"), True, 24)
                            LineaProv = LineaProv & RellenaABlancos(DBLet(Rs!maidatos, "T"), True, 50) & RellenaABlancos(DBLet(Rs!telefonocta, "T"), True, 15)
                            LineaProv = RellenaABlancos(LineaProv, True, 293)
                        End If
                    
                    Else
                        Regs = Regs + 1
                        
                        Im = Rs!ImpEfect
                        Importe = Importe + Im 'sumatorio
                        
                        'Datos vto
                        cad = "3" & Space(15) & IIf(Im < 0, "A", "F")
                        cad = cad & RellenaABlancos(Rs!numfactu, True, 20) & Format(Rs!FecFactu, "yyyymmdd")
                        
                        
                        cad = cad & FrmtCurren(Im, 15) & Format(Rs!fecefect, "yyyymmdd")
                        cad = FrmtStr(cad, 293)
                    
                        Print #NFich, cad
                    
                    End If
                    
                Rs.MoveNext
                
                If Rs.EOF And K = 1 Then
                    'El ultimo PRoveedor hay que meterlo
                    Print #NFich, LineaProv
                End If
            Wend
        Next
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    If Regs > 0 Then
        'registro totales
        LineaProv = "4" & Format(Regs, "000000") & Format(NumRegElim, "00000") & " " & FrmtCurren(Importe, 15)
        LineaProv = FrmtStr(LineaProv, 293)
        Print #NFich, LineaProv
    End If
    
    
    Close (NFich)
    NFich = -1
    If Regs > 0 Then
        
        If Importe <> Impor2 Then
            MsgBox "Importes vencimientos distinto del confirming", vbExclamation
        Else
            GeneraFicheroBancaMarch = True
        End If
        
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description, cad
     If NFich > 0 Then Close (NFich)
     
End Function










