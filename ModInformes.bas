Attribute VB_Name = "ModInformes"
Option Explicit


Public AbiertoOtroFormEnListado As Boolean  'Para saber si ha abieto un from desde el forms de listados



'Los reports
Public cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Public cadParam As String 'Cadena con los parametros para Crystal Report
Public numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Public cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Public cadNomRPT As String 'Nombre del informe a Imprimir
Public conSubRPT As Boolean 'Si el informe tiene subreports

Public cadPDFrpt As String 'Nombre del informe a enviar por email
Public vMostrarTree As Boolean
Public ExportarPDF As Boolean
Public SoloImprimir As Boolean


Dim Rs As Recordset
Dim cad As String
Dim Sql As String
Dim i As Integer


'Esto sera para el pb general
Dim TotalReg As Long
Dim Actual As Long


'Esta funcion lo que hace es genera el valor del campo
'El campo lo coge del recordset, luego sera field(i), y el tipo es para añadirle
'las coimllas, o quitarlas comas
'  Si es numero viene un 1 si no nada
Private Function ParaBD(ByRef Campo As ADODB.Field, Optional EsNumerico As Byte) As String
    
    If IsNull(Campo) Then
        ParaBD = "NULL"
    Else
        Select Case EsNumerico
        Case 1
            ParaBD = TransformaComasPuntos(CStr(Campo))
        Case 2
            'Fechas
            ParaBD = "'" & Format(CStr(Campo), "dd/mm/yyyy") & "'"
        Case Else
            ParaBD = "'" & Campo & "'"

            
        End Select
    End If
    ParaBD = "," & ParaBD
End Function


'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------

'   En este modulo se crearan los datos para los informes
'   Con lo cual cada Function Generara los datos en la tabla



'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------











Public Function InformeConceptos(ByRef vSql As String) As Boolean

On Error GoTo EGI_Conceptos
    InformeConceptos = False
    'Borramos los anteriores
    Conn.Execute "Delete from Usuarios.zconceptos where codusu = " & vUsu.Codigo
    cad = "INSERT INTO Usuarios.zconceptos (codusu, codconce, nomconce,tipoconce) VALUES ("
    cad = cad & vUsu.Codigo & ",'"
    Set Rs = New ADODB.Recordset
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Sql = cad & Format(Rs.Fields(0), "000")
        Sql = Sql & "','" & Rs.Fields(1) & "','" & Rs.Fields(3) & "')"
        Conn.Execute Sql
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    InformeConceptos = True
Exit Function
EGI_Conceptos:
    MuestraError Err.Number
    Set Rs = Nothing
End Function





Public Function ListadoEstadisticas(ByRef vSql As String) As Boolean
On Error GoTo EListadoEstadisticas
    ListadoEstadisticas = False
    Conn.Execute "Delete from Usuarios.zestadinmo1 where codusu = " & vUsu.Codigo
    'Sentencia insert
    cad = "INSERT INTO Usuarios.zestadinmo1 (codusu, codigo, codconam, nomconam, codinmov, nominmov,"
    cad = cad & "tipoamor, porcenta, codprove, fechaadq, valoradq, amortacu, fecventa, impventa) VALUES ("
    cad = cad & vUsu.Codigo & ","
    
    'Empezamos
    Set Rs = New ADODB.Recordset
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 1
    While Not Rs.EOF
        
        Sql = i & ParaBD(Rs!codconam, 1) & ParaBD(Rs!nomconam)
        Sql = Sql & ParaBD(Rs!Codinmov) & ",'" & DevNombreSQL(Rs!nominmov) & "'"
        
        Sql = Sql & ParaBD(Rs!tipoamor) & ParaBD(Rs!coeficie) & ParaBD(Rs!codprove)
        Sql = Sql & ParaBD(Rs!fechaadq, 2) & ParaBD(Rs!valoradq, 1) & ParaBD(Rs!amortacu, 1)
        Sql = Sql & ParaBD(Rs!fecventa, 2) 'FECHA
        Sql = Sql & ParaBD(Rs!impventa, 1) & ")"
        Conn.Execute cad & Sql
        
        'Sig
        Rs.MoveNext
        i = i + 1
    Wend
    ListadoEstadisticas = True
EListadoEstadisticas:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function





Public Function ListadoFichaInmo(ByRef vSql As String) As Boolean
On Error GoTo Err1
    ListadoFichaInmo = False
    Conn.Execute "Delete from Usuarios.zfichainmo where codusu = " & vUsu.Codigo
    'Sentencia insert
    cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, Fechaamor,Importe, porcenta) VALUES ("
    cad = cad & vUsu.Codigo & ","
    
    'Empezamos
    Set Rs = New ADODB.Recordset
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 1
    While Not Rs.EOF
        Sql = Rs!nominmov
        NombreSQL Sql
        Sql = i & ParaBD(Rs!Codinmov) & ",'" & Sql & "'"
        Sql = Sql & ParaBD(Rs!fechaadq, 2) & ParaBD(Rs!valoradq, 1) & ParaBD(Rs!fechainm, 2)
        Sql = Sql & ParaBD(Rs!imporinm, 1) & ParaBD(Rs!porcinm, 1)
        Sql = Sql & ")"
        Conn.Execute cad & Sql
        
        'Sig
        Rs.MoveNext
        i = i + 1
    Wend
    Rs.Close
    If i > 1 Then
        ListadoFichaInmo = True
    Else
        MsgBox "Ningún registro con esos valores", vbExclamation
    End If
Err1:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function




Public Function GenerarDiarios() As Boolean
On Error GoTo EGen
    GenerarDiarios = False
    cad = "Delete FROM Usuarios.ztiposdiario where codusu =" & vUsu.Codigo
    Conn.Execute cad
    cad = "INSERT INTO Usuarios.ztiposdiario (codusu, numdiari, desdiari)"
    cad = cad & " SELECT " & vUsu.Codigo & ",d.numdiari,d.desdiari"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".tiposdiario as d;"
    Conn.Execute cad
    GenerarDiarios = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function



Public Function IAsientosErrores(ByRef vSql As String) As Boolean
On Error GoTo EGen
    IAsientosErrores = False
    cad = "Delete FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'Las lineas
    cad = "INSERT INTO Usuarios.zdiapendact (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum,"
    cad = cad & " ampconce, timporteD, timporteH, codccost)"
    cad = cad & " SELECT " & vUsu.Codigo
    cad = cad & ",linapue.numdiari, tiposdiario.desdiari, linapue.fechaent, linapue.numasien, linapue.linliapu, linapue.codmacta, cuentas.nommacta, linapue.numdocum, linapue.ampconce, linapue.timporteD, linapue.timporteH, linapue.codccost"
    cad = cad & " FROM (linapue LEFT JOIN tiposdiario ON linapue.numdiari = tiposdiario.numdiari) LEFT JOIN cuentas ON linapue.codmacta = cuentas.codmacta"
    If vSql <> "" Then cad = cad & " WHERE " & vSql
    
    
    
    
    Conn.Execute cad
    
    Set Rs = New ADODB.Recordset
    cad = "select count(*) FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then cad = ""
    End If
    Rs.Close
    Set Rs = Nothing
    
    If cad <> "" Then
        MsgBox "Ningun registro por mostrar.", vbExclamation
        Exit Function
    End If
    IAsientosErrores = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function



Public Function ITotalesCtaConcepto(ByRef vSql As String, tabla As String) As Boolean
On Error GoTo EGen
    ITotalesCtaConcepto = False
    cad = "Delete FROM Usuarios.ztotalctaconce  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'Las lineas
    cad = "INSERT INTO Usuarios.ztotalctaconce (codusu, codmacta, nommacta, nifdatos, fechaent, timporteD, timporteH, codconce)"
    cad = cad & " SELECT " & vUsu.Codigo
    cad = cad & " ," & tabla & ".codmacta, nommacta, nifdatos, fechaent,"
    cad = cad & " timporteD,timporteH, codconce"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".cuentas ,"
    cad = cad & vUsu.CadenaConexion & "." & tabla & " WHERE cuentas.codmacta = " & tabla & ".codmacta"
    If vSql <> "" Then cad = cad & " AND " & vSql
    
    
    
    
    Conn.Execute cad
    
    
    'Inserto en ztmpdiarios para los que sean TODOS los conceptos
    cad = "Delete FROM Usuarios.ztiposdiario  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    cad = "INSERT INTO Usuarios.ztiposdiario SELECT " & vUsu.Codigo & ",codconce,nomconce FROM conceptos"
    Conn.Execute cad
    
    
    'Contamos para ver cuantos hay
    cad = "Select count(*) from Usuarios.ztotalctaconce WHERE codusu =" & vUsu.Codigo
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then i = 1
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    If i > 0 Then
        ITotalesCtaConcepto = True
    Else
        MsgBox "Ningún registro con esos valores.", vbExclamation
    End If
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Public Function IAsientosPre(ByRef vSql As String) As Boolean
On Error GoTo EGen
    IAsientosPre = False
    cad = "Delete FROM Usuarios.zasipre  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "INSERT INTO Usuarios.zasipre (codusu, numaspre, nomaspre, linlapre, codmacta, nommacta"
    cad = cad & ", ampconce, timporteD, timporteH, codccost)"

    cad = cad & " SELECT " & vUsu.Codigo
    cad = cad & ", t1.numaspre, t1.nomaspre, t2.linlapre,t2.codmacta, t3.nommacta,t2.ampconce,"
    cad = cad & "t2.timported,t2.timporteh,t2.codccost FROM "
    cad = cad & vUsu.CadenaConexion & ".asipre as t1,"
    cad = cad & vUsu.CadenaConexion & ".asipre_lineas as t2,"
    cad = cad & vUsu.CadenaConexion & ".cuentas as t3 WHERE "
    cad = cad & " t1.numaspre=t2.numaspre AND t2.codmacta=t3.codmacta"
    If vSql <> "" Then cad = cad & " AND " & vSql
    
    
    Conn.Execute cad
    IAsientosPre = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function






Public Function GeneraDatosHcoInmov(ByRef vSql As String) As Boolean

On Error GoTo EGeneraDatosHcoInmov
    GeneraDatosHcoInmov = False
        
    'Borramos tmp
    cad = "Delete from Usuarios.zfichainmo where codusu = " & vUsu.Codigo
    Conn.Execute cad
    'Abrimos datos
    Set Rs = New ADODB.Recordset
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ningún dato a mostrar", vbExclamation
    Else
        cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, fechaamor, Importe, porcenta) VALUES (" & vUsu.Codigo & ","
        TotalReg = 0
        While Not Rs.EOF
           
            TotalReg = TotalReg + 1
            'Metemos los nuevos datos
            Sql = TotalReg & ParaBD(Rs!Codinmov, 1) & ",'" & DevNombreSQL(CStr(Rs!nominmov)) & "'" & ParaBD(Rs!fechainm, 2)
            Sql = Sql & ",NULL,NULL" & ParaBD(Rs!imporinm, 1) & ParaBD(Rs!porcinm, 1) & ")"
            Sql = cad & Sql
            Conn.Execute Sql
            Rs.MoveNext
        Wend
        GeneraDatosHcoInmov = True
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EGeneraDatosHcoInmov:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function





Public Function GeneraDatosConceptosInmov() As Boolean

On Error GoTo EGeneraDatosConceptosInmov
    GeneraDatosConceptosInmov = False
        
    'Borramos tmp
    cad = "Delete from Usuarios.ztmppresu1 where codusu = " & vUsu.Codigo
    Conn.Execute cad
    'Abrimos datos
    Sql = "Select * from inmovcon"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ningún dato a mostrar", vbExclamation
    Else
        cad = "INSERT INTO Usuarios.ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (" & vUsu.Codigo
        While Not Rs.EOF
            'Metemos los nuevos datos
            Sql = ParaBD(Rs!codconam, 1) & ",'" & Format(Rs!codconam, "0000") & "'" & ParaBD(Rs!nomconam)
            Sql = Sql & ",0" & ParaBD(Rs!perimaxi, 1) & ParaBD(Rs!coefimaxi, 1) & ")"
            Sql = cad & Sql
            Conn.Execute Sql
            Rs.MoveNext
        Wend
        GeneraDatosConceptosInmov = True
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EGeneraDatosConceptosInmov:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function


'#################################################
'###########    A Ñ A D I D O     ################  DE  NUEVA CONTA DE DAVID
'#################################################
Public Sub PonerDatosPorDefectoImpresion(ByRef formu As Form, SoloImpresora As Boolean, Optional NombreArchivoEx As String)
On Error Resume Next
'        AbiertoOtroFormEnListado = False
        
        formu.txtTipoSalida(0).Text = Printer.DeviceName
        If Err.Number <> 0 Then
            formu.txtTipoSalida(0).Text = "No hay impresora instalada"
            Err.Clear
        End If
        If SoloImpresora Then Exit Sub
        
        formu.txtTipoSalida(1).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".csv"
        formu.txtTipoSalida(2).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".pdf"
        
        If Err.Number <> 0 Then Err.Clear
    
End Sub


'PDF=true   CSV=false
Public Function EliminarDocum(PDF As Boolean) As Boolean
    On Error Resume Next
    If PDF Then
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    Else
        If Dir(App.Path & "\docum.csv", vbArchive) <> "" Then Kill App.Path & "\docum.csv"
    End If
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Err.Clear
        EliminarDocum = False
    Else
        EliminarDocum = True
    End If
End Function


Public Sub ponerLabelBotonImpresion(ByRef BotonAcept As CommandButton, ByRef BotonImpr As CommandButton, SelectorImpresion As Integer)
    If SelectorImpresion = 0 Then
        BotonAcept.Caption = "&Vista previa"
    Else
        BotonAcept.Caption = "&Aceptar"
    End If
    BotonImpr.Visible = SelectorImpresion = 0
End Sub

Public Function PonerDesdeHasta(Campo As String, Tipo As String, ByRef Desde As TextBox, ByRef DesD As TextBox, ByRef Hasta As TextBox, ByRef DesH As TextBox, param As String) As Boolean
Dim Devuelve As String
Dim cad As String
Dim SubTipo As String 'F: fecha   N: numero   T: texto  H: HORA



    PonerDesdeHasta = False
    
    Select Case Tipo
    Case "F", "FEC"
        'Campos fecha
        SubTipo = "F"
    
    Case "CONC", "TDIA", "BIC", "AGE", "COI", "INM", "FRA", "COD", "GTO"
        'concepto
        SubTipo = "N"
        
    Case "CTA", "BAN", "CCO", "SER", "CRY"
        SubTipo = "T"
        
    Case "ASIP", "ASI", "AGE", "DPTO"
        SubTipo = "N"
        
    Case "TIVA", "DIA"
        SubTipo = "N"
        
    Case "TPAG", "FPAG", "REM", "ANYO"
        SubTipo = "N"
        
    End Select
    
    Devuelve = CadenaDesdeHasta(Desde, Hasta, Campo, SubTipo)
    If Devuelve = "Error" Then
        PonFoco Desde
        Exit Function
    End If
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Function
    
    If Devuelve = "" Then
        PonerDesdeHasta = True
        Exit Function
    End If
    
    'QUITO LAS LLAVES
    Devuelve = Replace(Devuelve, "{", "")
    Devuelve = Replace(Devuelve, "}", "")
    
    If SubTipo <> "F" And SubTipo <> "FH" Then
        'Fecha para Crystal Report

        If Not AnyadirAFormula(cadselect, Devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(Desde.Text, Hasta.Text, Campo, SubTipo)
        cad = Replace(cad, "{", "")
        cad = Replace(cad, "}", "")
        If Not AnyadirAFormula(cadselect, cad) Then Exit Function
    End If
    
    If Devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, Desde, Hasta, DesD, DesH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Public Function AnyadirAFormula(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
    If arg = "Error" Then
        AnyadirAFormula = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " AND (" & arg & ")"
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormula = True
End Function



Private Function AnyadirParametroDH(cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
    
    If Not TextoDESDE Is Nothing Then
         If TextoDESDE.Text <> "" Then
            cad = cad & "desde " & TextoDESDE.Text
'            If TD.Caption <> "" Then Cad = Cad & " - " & TD.Caption
        End If
    End If
    If Not TextoHasta Is Nothing Then
        If TextoHasta.Text <> "" Then
            cad = cad & "  hasta " & TextoHasta.Text
'            If TH.Caption <> "" Then Cad = Cad & " - " & TH.Caption
        End If
    End If
    
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function GeneraFicheroCSV(cadSQL As String, Salida As String) As Boolean
Dim NF As Integer
Dim i  As Integer

    On Error GoTo eGeneraFicheroCSV
    GeneraFicheroCSV = False
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Ningun dato generado", vbExclamation
        cadSQL = ""
    Else
        NF = FreeFile
        Open App.Path & "\docum.csv" For Output As #NF
        'Cabecera
        cadSQL = ""
        For i = 0 To miRsAux.Fields.Count - 1
            cadSQL = cadSQL & ";""" & miRsAux.Fields(i).Name & """"
        Next i
        Print #NF, Mid(cadSQL, 2)
    
    
        'Lineas
        While Not miRsAux.EOF
            cadSQL = ""
            For i = 0 To miRsAux.Fields.Count - 1
                cadSQL = cadSQL & ";""" & DBLet(miRsAux.Fields(i).Value, "T") & """"
            Next i
            Print #NF, Mid(cadSQL, 2)
            
            
            
            miRsAux.MoveNext
        Wend
        cadSQL = "OK"
    End If
    miRsAux.Close
    Close #NF

    If cadSQL = "OK" Then
        If CopiarFicheroASalida(True, Salida) Then GeneraFicheroCSV = True
    End If
    
    Exit Function
eGeneraFicheroCSV:
    MuestraError Err.Number, Err.Description
End Function


Public Function CopiarFicheroASalida(csv As Boolean, Salida As String, Optional SinMensaje As Boolean) As Boolean
    CopiarFicheroASalida = False
    If csv Then
        FileCopy App.Path & "\docum.csv", Salida
    Else
        FileCopy App.Path & "\docum.pdf", Salida
    End If
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Copiando " & Salida
    Else
        If Not SinMensaje Then
            MsgBox "Fichero:  " & Salida & vbCrLf & "Generado con éxito.", vbInformation
        End If
        CopiarFicheroASalida = True
    End If
End Function

Public Function ImprimeGeneral() As Boolean
    
    Dim HaPulsadoImprimir As Boolean

    Screen.MousePointer = vbHourglass



    

    frmppal.SkinFramework1.AutoApplyNewWindows = False
    frmppal.SkinFramework1.AutoApplyNewThreads = False

  
    HaPulsadoImprimir = False
    cadPDFrpt = cadNomRPT
    With frmVisReport
        .Informe = App.Path & "\Informes\"
        If ExportarPDF Then
            'PDF
            .Informe = .Informe & cadPDFrpt
        Else
            'IMPRIMIR
            .Informe = .Informe & cadNomRPT
        End If
        .FormulaSeleccion = cadFormula
        .SoloImprimir = False
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .ConSubInforme = conSubRPT

        .NumCopias2 = 1

        .SoloImprimir = SoloImprimir
        .ExportarPDF = ExportarPDF
        .MostrarTree = vMostrarTree
        
        .Show vbModal
        HaPulsadoImprimir = .EstaImpreso
        
      End With
    
    
     'DAVID
     frmppal.SkinFramework1.AutoApplyNewWindows = True
     frmppal.SkinFramework1.AutoApplyNewThreads = True
    
End Function



Public Sub QuitarPulsacionMas(ByRef T As TextBox)
Dim i As Integer

    Do
        i = InStr(1, T.Text, "+")
        If i > 0 Then T.Text = Mid(T.Text, 1, i - 1) & Mid(T.Text, i + 1)
    Loop Until i = 0
        
End Sub



Public Sub LanzaProgramaAbrirOutlook(outTipoDocumento As Integer)
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    If Not PrepararCarpetasEnvioMail(True) Then Exit Sub
    
    If Not ExisteARIMAILGES Then Exit Sub

    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1
        'Conceptos
        Aux = "Conceptos.pdf"
    Case 2
        'Cuentas contables
        Aux = "Cuentas.pdf"
    Case 3
        'Asientos Predefinidos
        Aux = "Asientos Predefinidos.pdf"
    Case 4
        Aux = "Tipos Diario.pdf"
    Case 5
        Aux = "Asientos.pdf"
    Case 6
        Aux = "Tipos de Iva.pdf"
    Case 7
        Aux = "Tipos de Pago.pdf"
    Case 8
        Aux = "Formas de Pago.pdf"
    Case 9
        Aux = "Bancos.pdf"
    Case 10
        Aux = "Bic/Swift.pdf"
    Case 11
        Aux = "Agentes.pdf"
    Case 12
        Aux = "Informes.pdf"
    Case 13
        Aux = "Bancos.pdf"
    Case 14
        Aux = "AsientosHco.pdf"
    Case 15
        Aux = "Listado de Facturas de Cliente.pdf"
    Case 16
        Aux = "Relación de Clientes por Cta Ventas.pdf"
    Case 17
        Aux = "Listado de Facturas de Proveedores.pdf"
    Case 18
        Aux = "Relación de Proveedores por Cta Gastos.pdf"
    Case 19
        Aux = "Modelo 303.pdf"
    Case 20
        Aux = "Modelo 340.pdf"
    Case 21
        Aux = "Modelo 347.pdf"
    Case 51
        
    Case 100
        
    End Select
    NombrePDF = App.Path & "\temp\" & Aux
    If Dir(NombrePDF, vbArchive) <> "" Then Kill NombrePDF
    FileCopy App.Path & "\docum.pdf", NombrePDF
    
    Aux = FijaDireccionEmail(outTipoDocumento)
    Lanza = Aux & "|"
    Aux = ""
    Select Case outTipoDocumento
    Case 1
        Aux = "Conceptos"
    Case 2
        Aux = "Cuentas"
    Case 3
        'Asientos Predefinidos
        Aux = "Asientos Predefinidos"
    Case 4
        Aux = "Tipos Diario"
    Case 5
        Aux = "Asientos"
    Case 6
        Aux = "Tipos de Iva"
    Case 7
        Aux = "Tipos de Pago"
    Case 8
        Aux = "Formas de Pago"
    Case 9
        Aux = "Bancos"
    Case 10
        Aux = "Bic/Swift"
    Case 11
        Aux = "Agentes"
    Case 12
        Aux = "Informes"
    Case 13
        Aux = "Bancos"
    Case 14
        Aux = "AsientosHco"
    Case 15
        Aux = "Listado de Facturas de Cliente"
    Case 16
        Aux = "Relación de Clientes por Cta Ventas"
    Case 17
        Aux = "Listado de Facturas de Proveedores"
    Case 18
        Aux = "Relación de Proveedores por Cta Gastos"
    Case 19
        Aux = "Modelo 303"
    Case 20
        Aux = "Modelo 340"
    Case 21
        Aux = "Modelo 347"
        
        
    '--------------------------------------------------
    Case 51
        Aux = "Pedido proveedor nº: " ' & outClaveNombreArchiv
        
    Case 100
        Aux = "Factura nº" '& outClaveNombreArchiv
        
    End Select
    
    Lanza = Lanza & Aux & "|"
    
    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    Lanza = Lanza & NombrePDF & "|"
    
    Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub

Private Function FijaDireccionEmail(outTipoDocumento As Integer) As String
Dim campoemail As String
Dim otromail As String


    FijaDireccionEmail = ""
    
    
    If outTipoDocumento < 50 Then
        campoemail = ""
        
    Else
        If outTipoDocumento < 100 Then
            'Para provedores
            If outTipoDocumento = 51 Or outTipoDocumento = 52 Or outTipoDocumento = 53 Then
                campoemail = "maiprov1"
                otromail = "maiprov2"
            Else
                campoemail = "maiprov2"
                otromail = "maiprov1"
            End If
'            campoemail = DevuelveDesdeBDNew(cpconta, "proveedor", campoemail, "codprove", Me.outCodigoCliProv, "N", otromail)
            If campoemail = "" Then campoemail = otromail
        Else
            'Para Socios
            If outTipoDocumento >= 100 Then
                campoemail = "maisocio"
                otromail = "maisocio"
            Else
                campoemail = "maisocio"
                otromail = "maisocio"
            End If
'            campoemail = DevuelveDesdeBDNew(cAgro, "rsocios", campoemail, "codsocio", Me.outCodigoCliProv, "N") ' , otromail)
            If campoemail = "" Then campoemail = otromail
        End If
    End If
    FijaDireccionEmail = campoemail
End Function




'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        MkDir App.Path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.Path & "\temp\*.*", vbArchive) <> "" Then Kill App.Path & "\temp\*.*"
        End If
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas Envio Mail "
End Function


Public Function LanzaMailGnral(dirMail As String) As Boolean
'LLama al Programa de Correo (Outlook,...)
Dim Aux As String
Dim Lanza As String

On Error GoTo ELanzaHome

    LanzaMailGnral = False

    If Not ExisteARIMAILGES Then Exit Function


    If dirMail = "" Then
        MsgBox "No hay dirección e-mail a la que enviar.", vbExclamation
        Exit Function
    End If

    Aux = dirMail
    Lanza = Lanza & Aux & "||"

    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"

    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send

    'Campos reservados para el futuro
    Lanza = Lanza & "||||"

    'El/los adjuntos
    Lanza = Lanza & "|"

    Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus

    LanzaMailGnral = True

ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, vbCrLf & Err.Description
'    CadenaDesdeOtroForm = ""
End Function


Public Function ExisteARIMAILGES()
Dim Sql As String

    If Dir(App.Path & "\ArimailGes.exe") = "" Then
        MsgBox "No existe el programa ArimailGes.exe. Llame a Ariadna.", vbExclamation
        ExisteARIMAILGES = False
    Else
        ExisteARIMAILGES = True
    End If
End Function



Public Function HayRegParaInforme(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
    
    Sql = "Select count(*) FROM " & cTabla
    If cWhere <> "" Then
        Sql = Sql & " WHERE " & cWhere
    End If
    
    If TotalRegistros(Sql) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegParaInforme = False
    Else
        HayRegParaInforme = True
    End If
End Function

Public Function PonerParamRPT(Indice As String, nomDocu As String) As Boolean
Dim cad As String
Dim Encontrado As Boolean

        nomDocu = ""
        Encontrado = False
        PonerParamRPT = False
        
        cad = "select informe from scryst where codigo = " & DBSet(Indice, "T")
        
        Set Rs = New ADODB.Recordset
        Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            nomDocu = DBLet(Rs!Informe, "T")
            Encontrado = True
        End If
        
        If Encontrado = False Or nomDocu = "" Then
            cad = "No se han podido cargar los Parámetros de Tipos de Documentos." & vbCrLf
            MsgBox cad & "Debe configurar la aplicación.", vbExclamation
            PonerParamRPT = False
            Exit Function
        End If
        
        PonerParamRPT = True
    

End Function

Public Sub LanzaProgramaAbrirOutlookMasivo(outTipoDocumento As Integer, Cuerpo As String)
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    
    If Not ExisteARIMAILGES Then Exit Sub

    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1
        'Reclamacion
        Aux = "Reclamacion.pdf"
    End Select
    
    Sql = "select tmp347.*, cuentas.razosoci, cuentas.maidatos from tmp347, cuentas "
    Sql = Sql & " where codusu = " & vUsu.Codigo & " and importe <> 0 and tmp347.cta = cuentas.codmacta"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        NombrePDF = App.Path & "\temp\" & Rs!NIF
        
        'direccion email
        Aux = DBLet(Rs!maidatos)
        Lanza = Aux & "|"
        
        'asunto
        Aux = ""
        Select Case outTipoDocumento
        Case 1 ' reclamaciones
            Aux = RecuperaValor(Cuerpo, 1)
        End Select
        
        Lanza = Lanza & Aux & "|"
        
        
        
    If LCase(Mid(cadNomRPT, 1, 3)) = "esc" Then
    ' para el caso de escalona
    
        cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
        cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
        cad = cad & "<TR><TD VALIGN=""TOP""><P><FONT FACE=""Tahoma""><FONT SIZE=3>"
        cad = cad & RecuperaValor(Cuerpo, 2)
        'FijarTextoMensaje
        
        cad = cad & "</FONT></FONT></P></TD></TR><TR><TD VALIGN=""TOP"">"
        cad = cad & "<p class=""MsoNormal""><b><i>"
        cad = cad & "<span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">C."
        cad = cad & "R. Reial Séquia Escalona</span></i></b></p>"
        cad = cad & "<p class=""MsoNormal""><em><b>"
        cad = cad & "<span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">"
        cad = cad & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; La Junta</span></b></em><span style=""font-size: 10.0pt; font-family: Arial,sans-serif; color: black"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">&nbsp;</span></p>"
        cad = cad & "<p class=""MsoNormal"">"
        cad = cad & "<span style=""font-size: 13.5pt; font-family: Arial,sans-serif; color: #9999FF"">"
        cad = cad & "********************</span></p>"
        cad = cad & "<p class=MsoNormal><b>"
         cad = cad & "<span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>Confidencialidad"
         cad = cad & "</span></b><span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'><br>"
         cad = cad & "Este mensaje y sus archivos adjuntos van dirigidos exclusivamente a su destinatario, "
         cad = cad & "pudiendo contener información confidencial sometida a secreto profesional. No está permitida su reproducción o "
         cad = cad & "distribución sin la autorización expresa de Real Acequia Escalona. Si usted no es el destinatario final por favor "
         cad = cad & "elimínelo e infórmenos por esta vía.<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-top-alt:6.0pt;"
         cad = cad & "margin-right:0cm;margin-bottom:6.0pt;margin-left:0cm;text-align:justify'><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS"";color:black'>De acuerdo con la Ley 34/2002 (LSSI) y la Ley 15/1999 (LOPD), "
         cad = cad & "usted tiene derecho al acceso, rectificación y cancelación de sus datos personales informados en el fichero del que es "
         cad = cad & "titular Real Acequia Escalona. Si desea modificar sus datos o darse de baja en el sistema de comunicación electrónica "
         cad = cad & "envíe un correo a</span> <span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>"
         cad = cad & "<a href=""mailto:escalona@acequiaescalona.org"">escalona@acequiaescalona.org</a> </span><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS""'>, <span style='color:black'>indicando en la línea de <b>&#8220;Asunto&#8221;</b> el derecho "
         cad = cad & "que desea ejercitar. <o:p></o:p></span></span></p><p class=MsoNormal><o:p>&nbsp;</o> "
         
         'ahora en valenciano
         cad = cad & ""
         cad = cad & "<p class=MsoNormal><b>"
         cad = cad & "<span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>Confidencialitat"
         cad = cad & "</span></b><span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'><br>"
         cad = cad & "Aquest missatge i els seus arxius adjunts van dirigits exclusivamente al seu destinatari, "
         cad = cad & "podent contindre informació confidencial sotmesa a secret professional. No està permesa la seua reproducció o "
         cad = cad & "distribució sense la autorització expressa de Reial Séquia Escalona. Si vosté no és el destinatari final per favor "
         cad = cad & "elimíneu-lo e informe-nos per aquesta via.<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-top-alt:6.0pt;"
         cad = cad & "margin-right:0cm;margin-bottom:6.0pt;margin-left:0cm;text-align:justify'><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS"";color:black'>D'acord amb la Llei 34/2002 (LSSI) i la Llei 15/1999 (LOPD), "
         cad = cad & "vosté té dret a l'accés, rectificació i cancelació de les seues dades personals informats en el ficher del qué és "
         cad = cad & "titolar Reial Séquia Escalona. Si vol modificar les seues dades o donar-se de baixa en el sistema de comunicació electrònica "
         cad = cad & "envíe un correu a</span> <span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>"
         cad = cad & "<a href=""mailto:escalona@acequiaescalona.org"">escalona@acequiaescalona.org</a> </span><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS""'>, <span style='color:black'>indicant en la línea de <b>&#8220;Asumpte&#8221;</b> el dret "
         cad = cad & "que desitja exercitar. <o:p></o:p></span></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p> "
        
        
        cad = cad & "</TR></BODY></HTML>"
        
        
    Else
    
        cad = RecuperaValor(cad, 2)
        
    End If
        
    ' end
        
        
        'Aqui pondremos lo del texto del BODY
        
        Aux = ""
        Select Case outTipoDocumento
        Case 1 ' reclamaciones
            Aux = cad
        End Select
        Lanza = Lanza & Aux & "|"
        
        'Envio o mostrar
        Lanza = Lanza & "1"   '0. Display   1.  send
        
        'Campos reservados para el futuro
        Lanza = Lanza & "||||"
        
        'El/los adjuntos
        Lanza = Lanza & NombrePDF & "|"
        
        Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
        Shell Aux, vbNormalFocus
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub

