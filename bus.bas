Attribute VB_Name = "bus"
Option Explicit


Global I&, J&, k&                             ' Contadores
Global Msg$, MsgErr$, NumErr&                 ' Variables de control de error
Global CONT%, Opc%, Skn$, SknDir$             ' Otros contadores
Public Tmp%, m_hMod&

' a�adido por la insercion de documentos en las lineas de asientos
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public vUsu As Usuario  'Datos usuario
Public vEmpresa As Cempresa 'Los datos de la empresa
Public vParam As Cparametros  'Los parametros
Public vParamT As CparametrosT  'Los parametros
Public vControl As Control2 ' Clase de control.dat
Public vLog As cLOG   'Log de acciones

Private Procesador64bits As Boolean

'Formato de fecha
Public FormatoFecha As String
Public FormatoImporte As String
Public FormatoPrecio As String
Public FormatoDec10d2 As String
Public FormatoPorcen As String

Public DireccionAyuda As String

Public CadenaDesdeOtroForm As String
Public NumAsiPre As String
Public Ampliacion As String

'Public DB As Database
Public Conn As ADODB.Connection

Public Const cConta As Byte = 1 'trabajaremos con connConta (cxion a BD Contabilidad)

Public CadenaControl As String


'Global para n� de registro eliminado
Public NumRegElim  As Long

'Para algunos campos de texto sueltos controlarlos
Public miTag As CTag

'Variable para saber si se ha actualizado algun asiento
Public AlgunAsientoActualizado As Boolean
Public TieneIntegracionesPendientes As Boolean

Public miRsAux As ADODB.Recordset

Public AnchoLogin As String  'Para fijar los anchos de columna


Public AsientoConExtModificado As Byte


'Para ver si reviso la introduccion
Public RevisarIntroduccion As Byte

'Reorganizar iconos que se visualizan en el formulario principal
Public Reorganizar As Boolean

'He cambiado el FechaOK. Para almacenar lo que devuelve, en algunos sitios no tengo variable
'La pongo aqui y sera comun para todos
Public varFecOk As Byte
Public Const varTxtFec = "Fecha fuera de �mbito"


Public Saldo473en470 As Boolean
Public Saldo6y7en129 As Boolean




'ARIMONEY
Public Const vbTipoPagoRemesa = 4
Public Const vbEfectivo = 0
Public Const vbTransferencia = 1
Public Const vbTalon = 2
Public Const vbPagare = 3
Public Const vbTarjeta = 6
Public Const vbConfirming = 5


Public Const ContCreditoNav = 3

'++
Public teclaBuscar As Integer   'llamada desde prismaticos

Public Const vbLightBlue = &HFEEFDA
Public Const vbErrorColor = &HDFE1FF      '&HFFFFC0
Public Const vbMoreLightBlue = &HFEFBD8   ' azul clarito

'++
Public Const vbOpcionVer = 0
Public Const vbOpcionCrearEliminar = 1
Public Const vbOpcionModificar = 2
Public Const vbOpcionImprimir = 3
Public Const vbOpcionEspecial = 4


Public ValorAnterior As String

Public CadenaCambio As String

Public ContinuarCobro As Boolean
Public ContinuarPago As Boolean


Public NumConta As Integer

Public XAnt As Currency
Public YAnt As Currency


Public HaHabidoCambios As Boolean


    'Para los asientos k vemos desde la consulta de extractos
    '  0.- NADA
    '  1.- SIIII
Public Sub Main()


Dim cad As String
Dim NF As Integer

Dim Sql As String
Dim PrimeraBD As String


       Load frmIdentifica
       CadenaDesdeOtroForm = ""

       'Necesitaremos el archivo arifon.dat
       'MsgBox "Antes identifca.show"
       frmIdentifica.Show vbModal
        
       If CadenaDesdeOtroForm = "" Then
            'NO se ha identificado
            Set Conn = Nothing
            End
       End If
       
       '++
       CadenaDesdeOtroForm = vControl.UltEmpre 'ultima empresa
       vUsu.CadenaConexion = vControl.UltEmpre
       '++
       
       If CadenaDesdeOtroForm = "" Then
            'No ha seleccionado nonguna empresa
            Set Conn = Nothing
            End
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass

        ' antes de cerrar la conexion cojo de usuarios.empresasariconta la primera que encuentre
        Sql = "select min(conta) from usuarios.empresasariconta  "
        PrimeraBD = DevuelveValor(Sql)


        'Cerramos la conexion
        Conn.Close
        
        If AbrirConexion(CadenaDesdeOtroForm, True) = False Then
            CadenaDesdeOtroForm = PrimeraBD
            If AbrirConexion(CadenaDesdeOtroForm) = False Then
                End
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        LeerEmpresaParametros
        

        RevisarIntroduccion = 0
        
        'Otras acciones
        Screen.MousePointer = vbHourglass
        
        OtrasAcciones
        
        '--
        'Load frmPpal
        frmPpal.UnaVez = True
        frmPpal.nomempre = vEmpresa.nomempre
        frmPpal.Show
        
        Screen.MousePointer = vbHourglass
End Sub


Public Function LeerEmpresaParametros()
        'Abrimos la empresa
        Set vEmpresa = New Cempresa
        If vEmpresa.Leer = 1 Then
            MsgBox "No se han podido cargar datos empresa. Debe configurar la aplicaci�n.", vbExclamation
            Set vEmpresa = Nothing
        End If
            
           
        Set vParam = New Cparametros
        If vParam.Leer() = 1 Then
            MsgBox "No se han podido cargar los par�metros. Debe confgurar la aplicaci�n.", vbExclamation
            Set vParam = Nothing
        End If
        
        If Not vEmpresa Is Nothing And Not vParam Is Nothing Then
            If vEmpresa.TieneTesoreria Then
                Set vParamT = New CparametrosT
                If vParamT.Leer() = 1 Then
                    MsgBox "No se han podido cargar los par�metros de tesoreria. Debe confgurar la aplicaci�n.", vbExclamation
                    Set vParamT = Nothing
                End If
            End If
        End If
        
        'incializamos el objeto
        Set vLog = New cLOG
        'Si estamos en localhost, y el usuario es administrador
        'Haremos la opcion de volcar la
        
'        If Not (vEmpresa Is Nothing) And Not (vParam Is Nothing) Then
'
'            'PAra la consulta de extractos
'            CadenaDesdeOtroForm = "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
'            Conn.Execute CadenaDesdeOtroForm
'
'            CadenaDesdeOtroForm = "DELETE from tmpconext where codusu= " & vUsu.Codigo
'            Conn.Execute CadenaDesdeOtroForm
'
'            CadenaDesdeOtroForm = ""
'        End If
        
        
End Function

'/////////////////////////////////////////////////////////////////
'// Se trata de identificar el PC en la BD. Asi conseguiremos tener
'// los nombres de los PC para poder asignarles un codigo
'// UNa vez asignado el codigo  se lo sumaremos (x 1000) al codusu
'// con lo cual el usuario sera distinto( aunque sea con el mismo codigo de entrada)
'// dependiendo desde k PC trabaje

Public Function GestionaPC2() As Integer
CadenaDesdeOtroForm = ComputerName
If CadenaDesdeOtroForm <> "" Then
    FormatoFecha = DevuelveDesdeBD("codpc", "Usuarios.pcs", "nompc", CadenaDesdeOtroForm, "T")
    If FormatoFecha = "" Then
        NumRegElim = 0
        FormatoFecha = "Select max(codpc) from Usuarios.pcs"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open FormatoFecha, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            NumRegElim = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        NumRegElim = NumRegElim + 1
        If NumRegElim > 9999 Then
            MsgBox "Error en numero de PC's activos. Demasiados PC en BD. Llame a soporte t�cnico.", vbCritical
            End
        End If
        FormatoFecha = "INSERT INTO Usuarios.pcs (codpc, nompc) VALUES (" & NumRegElim & ", '" & CadenaDesdeOtroForm & "')"
        Conn.Execute FormatoFecha
    Else
        NumRegElim = Val(FormatoFecha)
    End If
    GestionaPC2 = NumRegElim
    
End If
End Function


Private Sub OtrasAcciones()
On Error Resume Next

    FormatoFecha = "yyyy-mm-dd"
    FormatoImporte = "#,###,###,##0.00"
    FormatoPrecio = "#,###0.000"
    FormatoDec10d2 = "##,###,##0.00"
    FormatoPorcen = "##0.00"
    
    '++
    teclaBuscar = 43

    DireccionAyuda = "http://help-ariconta.ariadnasw.com/"

    'Borramos uno de los archivos temporales
    If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
    
    
    'Borramos tmp bloqueos
    'Borramos temporal
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad(True)
    NumRegElim = Len(CadenaDesdeOtroForm)
    If NumRegElim = 0 Then
        CadenaDesdeOtroForm = ""
    Else
        CadenaDesdeOtroForm = " WHERE codusu = " & vUsu.Codigo
    End If
    Conn.Execute "Delete from zBloqueos " & CadenaDesdeOtroForm
    CadenaDesdeOtroForm = ""
    NumRegElim = 0
    
    
End Sub


'Usuario As String, Pass As String --> Directamente el usuario
Public Function AbrirConexion(BBDD As String, Optional OcultarMsg As Boolean) As Boolean
Dim cad As String
Dim Prueba As String

On Error GoTo EAbrirConexion



    AbrirConexion = False
    Prueba = "Set Conn = Nothing"
    Set Conn = Nothing
    Prueba = "Set Conn = New Connection"
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Prueba = "Conn.CursorLocation = adUseServer  "
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
    
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATA SOURCE=" & vControl.ODBC
    If BBDD <> "" Then cad = cad & ";DATABASE= " & BBDD
    cad = cad & ";UID=" & vControl.UsuarioBD
    cad = cad & ";PWD=" & vControl.PassworBD
    cad = cad & ";Persist Security Info=true"
    
    Prueba = "Conn.ConnectionString = cad"
    Conn.ConnectionString = cad
    Prueba = "Conn.open"
    Conn.Open
    AbrirConexion = True
    Exit Function
    
    
    
EAbrirConexion:
    If Not OcultarMsg Then
        MuestraError Err.Number, "Abrir conexi�n." & Prueba, Err.Description
    End If
End Function






'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo antes de bloquear
'   Prepara la conexion para bloquear
Public Sub PreparaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=0"
End Sub

'/////////////////////////////////////////////////
'   Esto lo ejecutaremos justo despues de un bloque
'   Prepara la conexion para bloquear
Public Sub TerminaBloquear()
    Conn.Execute "commit"
    Conn.Execute "set autocommit=1"
End Sub


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "," & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaPuntosComas = CADENA
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ".")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & ":" & Mid(CADENA, I + 1)
        End If
    Loop Until I = 0
    TransformaPuntosHoras = CADENA
End Function


Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"
                    DBLet = ""
                Case "N"
                    DBLet = 0
                Case "F"
                    DBLet = "0:00:00"
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

Public Function DBMemo(vData As Variant) As String
Dim C As String
    On Error Resume Next
    C = vData
    If Err.Number <> 0 Then
        'Borramos error
        Err.Clear
        C = ""
    End If
    DBMemo = C
End Function


''MODIFICADO. Conta nueva. Ambito fechas
'Public Function FechaCorrecta2(vFecha As Date) As Byte
''--------------------------------------------------------
''   Dada una fecha dira si pertenece o no
''   al intervalo de fechas que maneja la apliacion
''   Resultados:
''       0 .- A�o actual
''       1 .- Siguiente
''       2 .- Ambito fecha. Fecha menor a la del ambito !!!!! NUEVO !!!!
''       3 .- Anterior al inicio
''       4 .- Posterior al fin
''--------------------------------------------------------
'
'    If vFecha >= vParam.fechaini Then
'        'Mayor que fecha inicio
'        If vFecha >= vParam.FechaActiva Then
'            If vFecha <= vParam.fechafin Then
'                FechaCorrecta2 = 0
'            Else
'                'Compruebo si el a�o siguiente
'                If vFecha <= DateAdd("yyyy", 1, vParam.fechafin) Then
'                    FechaCorrecta2 = 1
'                Else
'                    FechaCorrecta2 = 4   'Fuera ejercicios
'                End If
'            End If
'        Else
'            FechaCorrecta2 = 2   'Menor que fecha actvia
'        End If
'    Else            '< fecha ini
'        FechaCorrecta2 = 3
'    End If
'End Function

Public Function FechaCorrecta2(vFecha As Date, Optional MostrarMensaje As Boolean) As Byte
Dim Mens As String
'--------------------------------------------------------
'   Dada una fecha dira si pertenece o no
'   al intervalo de fechas que maneja la apliacion
'   Resultados:
'       0 .- A�o actual
'       1 .- Siguiente
'       2 .- Ambito fecha. Fecha menor a la del ambito !!!!! NUEVO !!!!
'       3 .- Anterior al inicio
'       4 .- Posterior al fin
'--------------------------------------------------------
    
    If vFecha >= vParam.fechaini Then
        'Mayor que fecha inicio
        If vFecha >= vParam.FechaActiva Then 'vParamT.fechaAmbito Then --> por si no tenemos tesoreria
            If vFecha <= vParam.fechafin Then
                FechaCorrecta2 = 0
            Else
                'Compruebo si el a�o siguiente
                If vFecha <= DateAdd("yyyy", 1, vParam.fechafin) Then
                    FechaCorrecta2 = 1
                Else
                    FechaCorrecta2 = 4   'Fuera ejercicios
                    Mens = "mayor que fin ejercicios"
                End If
            End If
        Else
            Mens = "menor que fecha activa"
            FechaCorrecta2 = 2   'Menor que fecha actvia
        End If
    Else            '< fecha ini
        FechaCorrecta2 = 3
        Mens = "anterior al inicio de ejercicios"
    End If
    
    If FechaCorrecta2 > 1 Then
        If MostrarMensaje Then
            Mens = "Fecha " & Mens & ". Fecha: " & vFecha
            MsgBox Mens, vbExclamation
        End If
    End If
End Function

Public Sub MuestraError(numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim cad As String
    Dim Aux As String
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        cad = cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If
    'Numeros de errores que contolamos
    If Conn.Errors.Count > 0 Then
        ControlamosError Aux
        Conn.Errors.Clear
    Else
        Aux = ""
    End If
    If Aux <> "" Then Desc = Aux
    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    If Aux = "" Then
        If numero <> 513 Then cad = cad & "N�mero: " & numero & vbCrLf & "Descripci�n: " & Error(numero)
    End If
    MsgBox cad, vbExclamation
End Sub

Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function


Public Function RellenaCodigoCuenta(vCodigo As String) As String
    Dim I As Integer
    Dim J As Integer
    Dim CONT As Integer
    Dim cad As String
    
    RellenaCodigoCuenta = vCodigo
    If Len(vCodigo) > vEmpresa.DigitosUltimoNivel Then Exit Function
    I = 0: CONT = 0
    Do
        I = I + 1
        I = InStr(I, vCodigo, ".")
        If I > 0 Then
            If CONT > 0 Then CONT = 1000
            CONT = CONT + I
        End If
    Loop Until I = 0
    
    'Habia mas de un punto
    If CONT > 1000 Or CONT = 0 Then Exit Function
    
    'Cambiamos el punto por 0's  .-Utilizo la variable maximocaracteres, para no tener k definir mas
    I = Len(vCodigo) - 1 'el punto lo quito
    J = vEmpresa.DigitosUltimoNivel - I
    cad = ""
    For I = 1 To J
        cad = cad & "0"
    Next I
    
    cad = Mid(vCodigo, 1, CONT - 1) & cad
    cad = cad & Mid(vCodigo, CONT + 1)
    RellenaCodigoCuenta = cad
End Function


Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim Rs As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    
    If ValorCodigo = "" Then Exit Function
    
    cad = "Select " & kCampo
    If OtroCampo <> "" Then cad = cad & ", " & OtroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD." & vbCrLf & cad, Err.Description
End Function


Public Function DevuelveDesdeBDNew(vBD As Byte, Ktabla As String, kCampo As String, Kcodigo1 As String, valorCodigo1 As String, Optional tipo1 As String, Optional ByRef OtroCampo As String, Optional KCodigo2 As String, Optional ValorCodigo2 As String, Optional tipo2 As String, Optional KCodigo3 As String, Optional ValorCodigo3 As String, Optional tipo3 As String) As String
'IN: vBD --> Base de Datos a la que se accede
Dim Rs As Recordset
Dim cad As String
Dim Aux As String
    
On Error GoTo EDevuelveDesdeBDnew
    DevuelveDesdeBDNew = ""
'    If valorCodigo1 = "" And ValorCodigo2 = "" Then Exit Function
    cad = "Select " & kCampo
    If OtroCampo <> "" Then cad = cad & ", " & OtroCampo
    cad = cad & " FROM " & Ktabla
    If Kcodigo1 <> "" Then
        cad = cad & " WHERE " & Kcodigo1 & " = "
        If tipo1 = "" Then tipo1 = "N"
    Select Case tipo1
        Case "N"
            'No hacemos nada
            cad = cad & Val(valorCodigo1)
        Case "T"
            cad = cad & DBSet(valorCodigo1, "T")
        Case "F"
            cad = cad & DBSet(valorCodigo1, "F")
        Case Else
            MsgBox "Tipo : " & tipo1 & " no definido", vbExclamation
            Exit Function
    End Select
    End If
    
    If KCodigo2 <> "" Then
        cad = cad & " AND " & KCodigo2 & " = "
        If tipo2 = "" Then tipo2 = "N"
        Select Case tipo2
        Case "N"
            'No hacemos nada
            If ValorCodigo2 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo2)
            End If
        Case "T"
'            cad = cad & "'" & ValorCodigo2 & "'"
            cad = cad & DBSet(ValorCodigo2, "T")
        Case "F"
            cad = cad & "'" & Format(ValorCodigo2, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo2 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    If KCodigo3 <> "" Then
        cad = cad & " AND " & KCodigo3 & " = "
        If tipo3 = "" Then tipo3 = "N"
        Select Case tipo3
        Case "N"
            'No hacemos nada
            If ValorCodigo3 = "" Then
                cad = cad & "-1"
            Else
                cad = cad & Val(ValorCodigo3)
            End If
        Case "T"
            cad = cad & "'" & ValorCodigo3 & "'"
        Case "F"
            cad = cad & "'" & Format(ValorCodigo3, FormatoFecha) & "'"
        Case Else
            MsgBox "Tipo : " & tipo3 & " no definido", vbExclamation
            Exit Function
        End Select
    End If
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    
    Select Case vBD
        Case cConta ' Conta
            Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
    End Select
    
    If Not Rs.EOF Then
        DevuelveDesdeBDNew = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
EDevuelveDesdeBDnew:
        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
End Function





'Obvio
Public Function EsCuentaUltimoNivel(Cuenta As String) As Boolean
    EsCuentaUltimoNivel = (Len(Cuenta) = vEmpresa.DigitosUltimoNivel)
End Function


Public Function CuentaCorrectaUltimoNivel(ByRef Cuenta As String, ByRef Devuelve As String) As Boolean
'Comprueba si es numerica
Dim Sql As String

CuentaCorrectaUltimoNivel = False
If Cuenta = "" Then
    Devuelve = "Cuenta vacia"
    Exit Function
End If
If Not IsNumeric(Cuenta) Then
    Devuelve = "La cuenta debe de ser num�rica: " & Cuenta
    Exit Function
End If

'Rellenamos si procede
Cuenta = RellenaCodigoCuenta(Cuenta)

If Not EsCuentaUltimoNivel(Cuenta) Then
    Devuelve = "No es cuenta de �ltimo nivel: " & Cuenta
    Exit Function
End If

Sql = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
If Sql = "" Then
    Devuelve = "No existe la cuenta : " & Cuenta
    Exit Function
End If

'Llegados aqui, si que existe la cuenta
CuentaCorrectaUltimoNivel = True
Devuelve = Sql
End Function

'-------------------------------------------------------------------------
'
'   Es la misma solo k no si no existe cuenta no da error
Public Function CuentaCorrectaUltimoNivelSIN(ByRef Cuenta As String, ByRef Devuelve As String) As Byte
'Comprueba si es numerica
Dim Sql As String

CuentaCorrectaUltimoNivelSIN = 0
If Cuenta = "" Then
    Devuelve = "Cuenta vacia"
    Exit Function
End If
If Not IsNumeric(Cuenta) Then
    Devuelve = "La cuenta debe de ser num�rica: " & Cuenta
    Exit Function
End If

'Rellenamos si procede
Cuenta = RellenaCodigoCuenta(Cuenta)

CuentaCorrectaUltimoNivelSIN = 1
If Not EsCuentaUltimoNivel(Cuenta) Then
    Sql = "No es cuenta de �ltimo nivel"
Else
    Sql = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
    If Sql = "" Then
        Sql = "No existe la cuenta  "
    Else
        CuentaCorrectaUltimoNivelSIN = 2
    End If
End If

'Llegados aqui, si que existe la cuenta
Devuelve = Sql
End Function

Public Function CuentaCorrectaUltimoNivelTXT(TCta As TextBox, TDesc As TextBox) As Boolean
Dim C1 As String
Dim C2 As String

    C1 = TCta.Text
    C2 = ""
    CuentaCorrectaUltimoNivelTXT = CuentaCorrectaUltimoNivel(C1, C2)
    TCta.Text = C1
    TDesc.Text = C2
End Function




'Devuelve, para un nivel determinado, cuantos digitos tienen las cuentas
' a ese nivel
Public Function DigitosNivel(numnivel As Integer) As Integer
    Select Case numnivel
    Case 1
        DigitosNivel = vEmpresa.numdigi1

    Case 2
        DigitosNivel = vEmpresa.numdigi2
        
    Case 3
        DigitosNivel = vEmpresa.numdigi3
        
    Case 4
        DigitosNivel = vEmpresa.numdigi4
        
    Case 5
        DigitosNivel = vEmpresa.numdigi5
        
    Case 6
        DigitosNivel = vEmpresa.numdigi6
        
    Case 7
        DigitosNivel = vEmpresa.numdigi7
        
    Case 8
        DigitosNivel = vEmpresa.numdigi8
        
    Case 9
        DigitosNivel = vEmpresa.numdigi9
        
    Case 10
        DigitosNivel = vEmpresa.numdigi10
        
    Case Else
        DigitosNivel = -1
    End Select
End Function

Public Function NivelCuenta(CodigoCuenta As String) As Integer
Dim lon As Integer
Dim niv As Integer
Dim I As Integer
    NivelCuenta = -1
    lon = Len(CodigoCuenta)
    I = 0
    Do
       I = I + 1
       niv = DigitosNivel(I)
       If niv > 0 Then
            If niv = lon Then
                NivelCuenta = I
                I = 11 'para salir del bucle
            End If
        Else
            I = 11 'salimos pq ya no hay nveles para las cuentas de longitud lon
        End If
    Loop Until I > 10
End Function


Public Function ExistenSubcuentas(ByRef Cuenta As String, Nivel As Integer) As Boolean
Dim I As Integer
Dim B As Boolean
Dim cad As String
    
    I = DigitosNivel(Nivel)
    cad = Mid(Cuenta, 1, I)
    cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", cad, "T")
    If cad = "" Then
        'NO existe la subcuenta de nivel N
        'salimos
        ExistenSubcuentas = False
        Exit Function
    End If
    If Nivel > 1 Then
        ExistenSubcuentas = ExistenSubcuentas(Cuenta, Nivel - 1)
    Else
        ExistenSubcuentas = True
    End If
End Function


Public Function CreaSubcuentas(ByRef Cuenta, HastaNivel As Integer, TEXTO As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim cad As String
Dim Cta As String

On Error GoTo ECreaSubcuentas
CreaSubcuentas = False
For I = 1 To HastaNivel
    J = DigitosNivel(I)
    Cta = Mid(Cuenta, 1, J)
    cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cta, "T")
    If cad = "" Then
        'CreaCuenta
        cad = "INSERT INTO cuentas (codmacta, nommacta, apudirec, model347, razosoci, "
        cad = cad & " dirdatos, codposta, despobla, desprovi, nifdatos, maidatos, webdatos,"
        cad = cad & " obsdatos) VALUES ("
        cad = cad & "'" & Cta
        cad = cad & "', '" & TEXTO
        cad = cad & "', "
        cad = cad & "'N', 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
        Conn.Execute cad
    End If
Next I
CreaSubcuentas = True
Exit Function
ECreaSubcuentas:
    MuestraError Err.Number, "Creando subcuentas", Err.Description
End Function




Public Function CambiarBarrasPATH(ParaGuardarBD As Boolean, CADENA) As String
Dim I As Integer
Dim Ch As String
Dim Ch2 As String

If ParaGuardarBD Then
    Ch = "\"
    Ch2 = "/"
Else
    Ch = "/"
    Ch2 = "\"
End If
I = 0
Do
    I = I + 1
    I = InStr(1, CADENA, Ch)
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Ch2 & Mid(CADENA, I + 1)
Loop Until I = 0
CambiarBarrasPATH = CADENA
End Function


Public Function ImporteSinFormato(CADENA As String) As String
Dim I As Integer
'Quitamos puntos
Do
    I = InStr(1, CADENA, ".")
    If I > 0 Then CADENA = Mid(CADENA, 1, I - 1) & Mid(CADENA, I + 1)
Loop Until I = 0
ImporteSinFormato = TransformaPuntosComas(CADENA)
End Function



'Periodo vendran las fechas Ini y fin con pipe final
Public Sub SaldoHistorico(Cuenta As String, Periodo As String, DescCuenta As String, EsSobreEjerciciosCerrados As Boolean)
Dim Rs As Recordset
Dim Sql As String
Dim RC2 As String
Dim vImp As Currency

    Screen.MousePointer = vbHourglass
    Sql = "Select Sum(timporteD),sum(timporteH) from hlinapu"
    If EsSobreEjerciciosCerrados Then Sql = Sql & "1"
    Sql = Sql & " WHERE codmacta='" & Cuenta & "'"
    
    If Not EsSobreEjerciciosCerrados Then _
        Sql = Sql & " AND fechaent>='" & Format(vParam.fechaini, FormatoFecha) & "'"
    Sql = Sql & " AND punteada "
    Set Rs = New ADODB.Recordset
    RC2 = Cuenta & "|"
    'PUNTEADO
    Rs.Open Sql & "='S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
       RC2 = RC2 & Format(Rs.Fields(0), FormatoImporte) & "|"
       RC2 = RC2 & Format(Rs.Fields(1), FormatoImporte) & "|"
    Else
        RC2 = RC2 & "||"
    End If
    Rs.Close
    'SIN puntear
    Rs.Open Sql & "<>'S';", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
       RC2 = RC2 & Format(Rs.Fields(0), FormatoImporte) & "|"
       RC2 = RC2 & Format(Rs.Fields(1), FormatoImporte) & "|"
    Else
        RC2 = RC2 & "||"
    End If
    Rs.Close
    
    'En el periodo. Para cuando viene de puntear
    If Periodo <> "" Then
        Sql = "Select Sum(timporteD) , sum(timporteH) from hlinapu"
        If EsSobreEjerciciosCerrados Then Sql = Sql & "1"
        Sql = Sql & " WHERE codmacta='" & Cuenta & "' AND "
        Sql = Sql & Periodo
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            vImp = DBLet(Rs.Fields(0), "N")
            vImp = vImp - DBLet(Rs.Fields(1), "N")
            RC2 = RC2 & Format(vImp, FormatoImporte) & "|"
        Else
            RC2 = RC2 & "|"
        End If
    Else
        RC2 = RC2 & "|"
    End If
    RC2 = RC2 & DescCuenta & "|"
    Set Rs = Nothing
    'Mostramos la ventanita de mesaje
    frmMensajes.Opcion = 1
    frmMensajes.Parametros = RC2
    frmMensajes.Show vbModal

End Sub

'Lo que hace es comprobar que si la resolucion es mayor
'que 800x600 lo pone en el 400
Public Sub AjustarPantalla(ByRef formulario As Form)
    If Screen.Width > 13000 Then
        formulario.Top = 400
        formulario.Left = 400
    Else
        formulario.Top = 0
        formulario.Left = 0
    End If
    formulario.Width = 12000
    formulario.Height = 9000
End Sub


'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256.98
'   Tiene que venir num�rico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

If Importe = "" Then
    ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
End If
End Function





Public Function DiasMes(Mes As Byte, Anyo As Integer) As Integer
    Select Case Mes
    Case 2
        If (Anyo Mod 4) = 0 Then
            DiasMes = 29
        Else
            DiasMes = 28
        End If
    Case 1, 3, 5, 7, 8, 10, 12
        DiasMes = 31
    Case Else
        DiasMes = 30
    End Select
End Function





Public Function ComprobarEmpresaBloqueada(codusu As Long, ByRef Empresa As String) As Boolean
Dim cad As String

ComprobarEmpresaBloqueada = False

'Antes de nada, borramos las entradas de usuario, por si hubiera kedado algo
Conn.Execute "Delete from Usuarios.vBloqBD where codusu=" & codusu

'Ahora comprobamos k nadie bloquea la BD
cad = DevuelveDesdeBD("codusu", "Usuarios.vBloqBD", "conta", Empresa, "T")
If cad <> "" Then
    'En teoria esta bloqueada. Puedo comprobar k no se haya kedado el bloqueo a medias
    
    Set miRsAux = New ADODB.Recordset
    cad = "show processlist"
    miRsAux.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        If miRsAux.Fields(3) = Empresa Then
            cad = miRsAux.Fields(2)
            miRsAux.MoveLast
        End If
    
        'Siguiente
        miRsAux.MoveNext
    Wend
    
    If cad = "" Then
        'Nadie esta utilizando la aplicacion, luego se puede borrar la tabla
        Conn.Execute "Delete from Usuarios.vBloqBD where conta ='" & Empresa & "'"
        
    Else
        MsgBox "BD bloqueada.", vbCritical
        ComprobarEmpresaBloqueada = True
    End If
End If

Conn.Execute "commit"
End Function


Public Function Bloquear_DesbloquearBD(Bloquear As Boolean) As Boolean

On Error GoTo EBLo
    Bloquear_DesbloquearBD = False
    If Bloquear Then
        CadenaDesdeOtroForm = "INSERT INTO Usuarios.wBloqBD (codusu, conta) VALUES (" & vUsu.Codigo & ",'" & vUsu.CadenaConexion & "')"
    Else
        CadenaDesdeOtroForm = "DELETE FROM  Usuarios.wBloqBD WHERE codusu =" & vUsu.Codigo & " AND conta = '" & vUsu.CadenaConexion & "'"
    End If
    Conn.Execute CadenaDesdeOtroForm
    Bloquear_DesbloquearBD = True
    Exit Function
EBLo:
    'MuestraError Err.Number, "Bloq. BD"
    Err.Clear
End Function


Private Function Servidor() As String
Dim I As Integer
Dim cad As String

    On Error GoTo eServidor

    Servidor = ""

    I = InStr(1, Conn.ConnectionString, "SERVER=")
    
    If I = 0 Then Exit Function
    
    cad = Mid(Conn.ConnectionString, I, Len(Conn.ConnectionString) - I)
    
    I = InStr(1, cad, ";")
    
    Servidor = Mid(cad, 8, I - 8)  '8 es la longitud de "SERVER="
    Exit Function
    
eServidor:
    
End Function


Public Function OtrosPCsContraContabiliad(EsAlIniciar As Boolean) As String
Dim MiRS As Recordset
Dim cad As String
Dim Equipo As String
Dim EquipoConBD As Boolean

Dim SERVER As String

    On Error GoTo EOtrosPCsContraContabiliad
    
    Set MiRS = New ADODB.Recordset
    
    SERVER = Servidor
    
    EquipoConBD = (UCase(vUsu.PC) = UCase(SERVER)) Or (LCase(SERVER) = "localhost")
    
    cad = "show processlist"
    MiRS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not MiRS.EOF
        If UCase(MiRS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
            Equipo = MiRS.Fields(2)
            'Primero quitamos los dos puntos del puerot
            NumRegElim = InStr(1, Equipo, ":")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            'El punto del dominio
            NumRegElim = InStr(1, Equipo, ".")
            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
            
            Equipo = UCase(Equipo)
            
            If Equipo <> vUsu.PC Then
                    
                    NumRegElim = 0
                    If Equipo <> "LOCALHOST" Then
                        'Si no es localhost
                        NumRegElim = 1
                    Else
                        'HAy un proceso de loclahost. Luego, si el equipo no tiene la BD
                        If Not EquipoConBD Then NumRegElim = 1
                    End If
                    
                    'Si hay que insertar
                    If NumRegElim = 1 Then
                        If InStr(1, cad, Equipo & "|") = 0 Then cad = cad & Equipo & "|"
                    End If
            End If
        End If
        'Siguiente
        MiRS.MoveNext
    Wend
    NumRegElim = 0
    MiRS.Close
    Set MiRS = Nothing
    OtrosPCsContraContabiliad = cad
    Exit Function
EOtrosPCsContraContabiliad:
    MuestraError Err.Number, Err.Description, "Leyendo PROCESSLIST"
    Set MiRS = Nothing
    If EsAlIniciar Then
        OtrosPCsContraContabiliad = "LEYENDOPC|"
    Else
        cad = "�El sistema no puede determinar si hay PCs conectados. �Desea continuar igualmente?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
            OtrosPCsContraContabiliad = ""
        Else
            OtrosPCsContraContabiliad = "USUARIO ACTUAL|"
        End If
    End If
    
    
    
End Function


Public Function EsNumerico(TEXTO As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim cad As String
    
    EsNumerico = False
    cad = ""
    If Not IsNumeric(TEXTO) Then
        cad = "El campo debe ser num�rico"
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, TEXTO, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then cad = "Numero de puntos incorrecto"
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, TEXTO, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then cad = "Numero incorrecto"
        End If
        
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
    Else
        EsNumerico = True
    End If
End Function



Public Function EsFechaOK(T As TextBox) As Boolean
Dim cad As String
    
    cad = T.Text
    If InStr(1, cad, "/") = 0 Then
        If Len(T.Text) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T.Text) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/20" & Mid(cad, 5)
        End If
    End If
    
    If IsDate(cad) Then
        EsFechaOK = True
        T.Text = Format(cad, "dd/MM/yyyy")
    Else
        EsFechaOK = False
    End If
End Function



Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/20" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOKString = True
        T = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function

'Devuelve si hay archivos
'                                                        Llevara la forma: 01, 02  para la empresa 1 o 2..

'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
End Sub

Public Function DevNombreSQL(CADENA As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = CADENA
End Function



'Para los balnces
Public Function FechaInicioIGUALinicioEjerecicio(FecIni As Date, EjerciciosCerrados1 As Boolean) As Byte
Dim Fecha As Date
Dim Salir As Boolean
Dim I As Integer
On Error GoTo EfechaInicioIGUALinicioEjerecicio

    FechaInicioIGUALinicioEjerecicio = 1
    If EjerciciosCerrados1 Then
        I = -1 'En ejercicios cerrados emp�zamos mirando un a�o por debajo fecini
    Else
        I = 1
    End If
    Fecha = DateAdd("yyyy", I, vParam.fechaini)
    Salir = False
    While Not Salir
        If FecIni = Fecha Then
            'Fecha inicio del listado contiene es fecha incio ejercicio
            FechaInicioIGUALinicioEjerecicio = 0
            'Modificacion del 2 de Septiembre de 2004
            'Si la fehca es incio pero el de el ejercicio siguiente
            'entonces no te lo
            If Not EjerciciosCerrados1 Then
                'Ejerecicio actual / siguiente
                If FecIni > vParam.fechaini Then
                    'Ejercicio siguiente. Con lo cual SI tengo k poner los acumulados
                    FechaInicioIGUALinicioEjerecicio = 1
                End If
            End If
            
            
            
            Salir = True
        Else
            If FecIni < Fecha Then
                Fecha = DateAdd("yyyy", -1, Fecha)
            Else
                Salir = True
            End If
        End If
    Wend
    
    Exit Function
EfechaInicioIGUALinicioEjerecicio:
    Err.Clear  'No tiene importancia
End Function





Public Function DevuelveDigitosNivelAnterior() As Integer
Dim J As Integer
    DevuelveDigitosNivelAnterior = 3
    If vEmpresa Is Nothing Then Exit Function
    If vEmpresa.numnivel < 2 Then Exit Function
    J = vEmpresa.numnivel - 1
    J = DigitosNivel(J)
    If J < 3 Then J = 3
    DevuelveDigitosNivelAnterior = J
End Function



'--------------------------------------------------------------------------
' Los numeros vendran formateados o sin formatear, pero siempre viene texto
'
Public Function CadenaCurrency(TEXTO As String, ByRef Importe As Currency) As Boolean
Dim I As Integer

    On Error GoTo ECadenaCurrency
    Importe = 0
    CadenaCurrency = False
    If Not IsNumeric(TEXTO) Then Exit Function
    I = InStr(1, TEXTO, ",")
    If I = 0 Then
        'Significa k el numero no esta  formateado y como mucho tiene punto
        Importe = CCur(TransformaPuntosComas(TEXTO))
    Else
        Importe = ImporteFormateado(TEXTO)
    End If
    
    CadenaCurrency = True
    
    Exit Function
ECadenaCurrency:
    Err.Clear
End Function


Public Function UsuariosConectados(vMens As String, Optional DejarContinuar As Boolean) As Boolean
Dim I As Integer
Dim cad As String
Dim metag As String
Dim Sql As String
cad = OtrosPCsContraContabiliad(False)
UsuariosConectados = False
If cad <> "" Then
    UsuariosConectados = True
    I = 1
    metag = vMens
    If vMens <> "" Then metag = metag & vbCrLf
    metag = metag & vbCrLf & "Los siguientes PC's est�n conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
    
    Do
        Sql = RecuperaValor(cad, I)
        If Sql <> "" Then
            metag = metag & "    - " & Sql & vbCrLf
            I = I + 1
        End If
    Loop Until Sql = ""
    If DejarContinuar Then
        'Hare la pregunta
        metag = metag & vbCrLf & "�Continuar?"
        If MsgBox(metag, vbQuestion + vbYesNoCancel) = vbYes Then UsuariosConectados = False
    Else
        'Informa UNICAMENTE
        MsgBox metag, vbExclamation
    End If
End If
End Function




Public Function HayKHabilitarCentroCoste(ByRef Cuenta As String) As Boolean
Dim Ch As String

    HayKHabilitarCentroCoste = False
    If Cuenta <> "" Then
        'Hay cuenta
         Ch = Mid(Cuenta, 1, 1)
         If Ch = vParam.grupogto Or Ch = vParam.grupovta Or Ch = vParam.grupoord Then
            HayKHabilitarCentroCoste = True
         Else
            Ch = Mid(Cuenta, 1, 3)
            If Ch = vParam.Subgrupo1 Or Ch = vParam.Subgrupo2 Then
                HayKHabilitarCentroCoste = True
            End If
        End If
    End If
End Function


Public Function EjecutaSQL(ByRef Sql As String) As Boolean
    EjecutaSQL = False
    On Error Resume Next
    Conn.Execute Sql
    If Err.Number <> 0 Then
        Err.Clear
    Else
        EjecutaSQL = True
    End If
End Function



Public Function DirectorioEAT() As Boolean
    On Error GoTo EDirecEAT
    DirectorioEAT = False
    If Dir("C:\AEAT", vbDirectory) = "" Then
        MsgBox "No se encuentra la carpeta de la agencia tributaria.  ( C:\AEAT )", vbExclamation
    Else
        DirectorioEAT = True
    End If
    Exit Function
EDirecEAT:
    Err.Clear
End Function





Public Function EstaLaCuentaBloqueada(ByRef codmacta As String, Fecha As Date) As Boolean
Dim I As Integer

        EstaLaCuentaBloqueada = False
        If vParam.CuentasBloqueadas <> "" Then
            I = InStr(1, vParam.CuentasBloqueadas, codmacta & ":")
            If I > 0 Then
                'La cuenta esta con fecha de bloqueo
                If Fecha >= CDate(Mid(vParam.CuentasBloqueadas, I + Len(codmacta) + 1, 10)) Then EstaLaCuentaBloqueada = True
            End If
        End If
End Function


Public Sub CerrarRs(ByRef Rsss As ADODB.Recordset)
    On Error Resume Next
    Rsss.Close
    If Err.Number <> 0 Then Err.Clear
End Sub


'*******************************************************************
'*******************************************************************
'*******************************************************************
'   Septiembre 2011
'
'  Letra serie 3 Digitos
'  Con lo cual para algunas campos (numdocum de hlinapu) son un maximo de
'   10 posiciones. Como antes era un digito letra ser, formateabamos con 9
'       numerofactura debe ser NUMERICO


Public Function EsEntero(TEXTO As String) As Boolean
Dim I As Integer
Dim C As Integer
Dim L As Integer
Dim res As Boolean

    res = True
    EsEntero = False

    If Not IsNumeric(TEXTO) Then
        res = False
    Else
        'Vemos si ha puesto mas de un punto
        C = 0
        L = 1
        Do
            I = InStr(L, TEXTO, ".")
            If I > 0 Then
                L = I + 1
                C = C + 1
            End If
        Loop Until I = 0
        If C > 1 Then res = False
        
        'Si ha puesto mas de una coma y no tiene puntos
        If C = 0 Then
            L = 1
            Do
                I = InStr(L, TEXTO, ",")
                If I > 0 Then
                    L = I + 1
                    C = C + 1
                End If
            Loop Until I = 0
            If C > 1 Then res = False
        End If
        
    End If
        EsEntero = res
End Function




'#####################################################################################################
'#####################################################################################################
'#
'#
'#                          T   E   S   O   R   E   R   I   A
'#
'#
'#####################################################################################################
'#####################################################################################################

'********************************************************************************
'********************************************************************************
'   Carga iconos de un formulario
'   -----------------------------
'       Opciones:   Colection  El col de imagenes
'                   Tipo    1.- Lupa
'                           2.- Fecha
'                           3.- Ayuda
Public Sub CargaImagenesAyudas(ByRef Colec, Tipo As Byte, Optional ToolTipText_ As String)
Dim I As Image

    

    For Each I In Colec
            I.Picture = frmPpal.imgIcoForms.ListImages(Tipo).Picture
            If I.ToolTipText = "" Then
                If ToolTipText_ <> "" Then
                    I.ToolTipText = ToolTipText_
                Else
                    If Tipo = 3 Then
                        I.ToolTipText = "Ayuda"
                    ElseIf Tipo = 2 Then
                        I.ToolTipText = "Buscar fecha"
                    Else
                        I.ToolTipText = "Buscar"
                    End If
                End If
            End If
    Next
End Sub





Public Sub CargaIconoListview(ByRef QueListview As ListView)
On Error Resume Next
    If Dir(App.Path & "\listview.dat", vbArchive) <> "" Then
        QueListview.Picture = LoadPicture(App.Path & "\listview.dat")
        QueListview.PictureAlignment = lvwTopLeft
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub TirarAtrasTransaccion()
    On Error Resume Next
    Conn.RollbackTrans
    If Err.Number <> 0 Then
        If Conn.Errors(0).NativeError = 1196 Then
            'NO PASA NADA. YA sabemos que las tblas tmp no se van a hacer rollbacktrans
            Conn.Cancel
            Conn.RollbackTrans
        Else
            MsgBox "Deshaciendo transacciones:" & Err.Description, vbExclamation
        End If
        Err.Clear
        Conn.Errors.Clear
        
    End If
    
End Sub

Public Function DevuelveNombreInformeSCRYST(NumInforme As Integer, Titulo As String) As String
Dim cad As String

        DevuelveNombreInformeSCRYST = ""
        cad = DevuelveDesdeBD("informe", "scryst", "codigo", CStr(NumInforme))

        If cad = "" Then
            MsgBox "No existe el informe para: " & Titulo & " (" & NumInforme & ")", vbExclamation
            Exit Function
        End If
        
        
        If Dir(App.Path & "\InformesT\" & cad, vbArchive) = "" Then
            MsgBox "No se encuentra el archivo: " & cad & vbCrLf & "Opcion: " & Titulo, vbExclamation
            Exit Function
        End If
        DevuelveNombreInformeSCRYST = cad
            
End Function

Public Function Memo_Leer(ByRef C As ADODB.Field) As String
    On Error Resume Next
    Memo_Leer = C.Value
    If Err.Number <> 0 Then
        Err.Clear
        Memo_Leer = ""
    End If
End Function


'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'
'   Imprimir listado caja.
'
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Public Function ImpirmirListadoCaja(ByRef vSql As String, SaldoArrastrado As Boolean) As Boolean
Dim miSQL As String
Dim L As Long
Dim cad As String
Dim Caja As String
Dim CtaCaja As String
Dim Tipo As Integer
Dim RT As ADODB.Recordset

    ImpirmirListadoCaja = False
    Conn.Execute "DELETE from Usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    
    Set miRsAux = New ADODB.Recordset
    miSQL = "Select slicaja.*,nommacta from slicaja,cuentas,susucaja where slicaja.codmacta=cuentas.codmacta " & vSql
    miSQL = miSQL & " ORDER BY slicaja.codusu,feccaja,numlinea"
    miRsAux.Open miSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 1
    'INSERT INTO ztesoreriacomun (
    'codusu, fecha1,codigo, texto1, texto2,opcion, texto3, texto4, texto5, texto6,
    'importe1, importe2,   fecha3,
    'observa1, observa2) VALUES (
    
    vSql = "INSERT INTO Usuarios.ztesoreriacomun (codusu, fecha1,codigo, texto1, texto2,texto4,opcion, texto3, observa1, "
    vSql = vSql & "texto5,importe1 ,importe2,texto6 ) VALUES (" & vUsu.Codigo & ",'"
    CtaCaja = ""
    While Not miRsAux.EOF
        If miRsAux!codusu <> CtaCaja Then
            CtaCaja = miRsAux!codusu
            
            Caja = DevuelveDesdeBD("nomusu", "usuarios.usuarios", "codusu", miRsAux!codusu, "N")
            Caja = DevNombreSQL(Caja)
            Caja = ",'" & CtaCaja & "','" & Caja & "'"
            
            'Si lleva saldo arrastrado entonces lo obtengo del datos de usucaja
            If SaldoArrastrado Then
                cad = "Select saldo from susucaja where codusu =" & CtaCaja
                Set RT = New ADODB.Recordset
                RT.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                If Not RT.EOF Then
                    'Inserto una primera linea con fecha 1900 con el saldo de la caja
                    cad = "1900-01-01'," & L & Caja
                    For Tipo = 1 To 4
                        cad = cad & ",NULL"
                    Next Tipo
                      cad = cad & ",'Saldo caja :',"
                    If RT!Saldo >= 0 Then
                        cad = cad & TransformaComasPuntos(CStr(RT!Saldo)) & ",0"
                    Else
                        cad = cad & "0," & TransformaComasPuntos(CStr(Abs(RT!Saldo)))
                    End If
                    cad = vSql & cad & ",NULL)"
                    Conn.Execute cad
                    'Sumo L
                    L = L + 1
                End If
                RT.Close
                Set RT = Nothing
            End If
        End If
        
        cad = Format(miRsAux!feccaja, FormatoFecha) & "'," & L & Caja
        If miRsAux!tipomovi = 1 Then
            Tipo = 1
            'FACTURAS PROVEEDORES
            cad = cad & ",'FRAPRO',1,'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            'Numero de factura
            cad = cad & DevNombreSQL(DBLet(miRsAux!NumFacpr))
            If Not IsNull(miRsAux!numvenci) Then cad = cad & " - Vto: " & miRsAux!numvenci
            cad = cad & "',"
        Else
            If miRsAux!tipomovi >= 2 Then
                'TRASPASO o PAGO
                Tipo = Val(miRsAux!tipomovi)
                cad = cad & ",'"
                If Tipo = 2 Then
                    cad = cad & "PAGO"
                Else
                    cad = cad & "TRASPASO"
                End If
                cad = cad & "'," & Tipo & ",'"
                cad = cad & "','" & DevNombreSQL(miRsAux!Ampliaci) & "',NULL,"
                ''" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            Else
                'FACTURA CLIENTE
                Tipo = 0
                cad = cad & ",'FRACLI',0,'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                'Numero factura
                If Not IsNull(miRsAux!NUmSerie) Then cad = cad & miRsAux!NUmSerie
                If Not IsNull(miRsAux!numfaccl) Then cad = cad & Format(miRsAux!numfaccl, "0000000000")
                If Not IsNull(miRsAux!numvenci) Then cad = cad & " - Vto: " & miRsAux!numvenci
                cad = cad & "',"
            End If
        End If
        'El importe
        cad = cad & TransformaComasPuntos(CStr(DBLet(miRsAux!ImporteD, "N")))
        cad = cad & "," & TransformaComasPuntos(CStr(DBLet(miRsAux!ImporteH, "N")))
        
        
        'Texto 6: numero de linea
        cad = cad & "," & Format(miRsAux!NumLinea, "00000")
        
        cad = vSql & cad & ")"
        Conn.Execute cad
        
        miRsAux.MoveNext
        L = L + 1
    Wend
    miRsAux.Close
    '
    'INSERT INTO ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion) VALUES (
    ImpirmirListadoCaja = True
End Function


Public Function ListadoFormaPago(ByRef Sql As String) As Boolean

    On Error GoTo EListadoFormaPago
    ListadoFormaPago = False
    
    Conn.Execute "DELETE from Usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
        
    'MONTO EL SQL AL REVES. Empezando por el where

    Sql = " WHERE sforpa.tipforpa = stipoformapago.tipoformapago " & Sql
    Sql = " FROM sforpa ,stipoformapago" & Sql
    Sql = " sforpa.codforpa,sforpa.nomforpa,stipoformapago.descformapago " & Sql
    Sql = "INSERT INTO Usuarios.ztesoreriacomun(codusu,codigo,texto1,texto2) Select " & vUsu.Codigo & "," & Sql
    'INSERT INTO Usuarios.ztesoreriacomun (codusu, observa1, codigo,
    'texto1, texto2,  texto3, texto4 ,texto5) VALUES (
    
    
    
    Conn.Execute Sql
    
    Set miRsAux = New ADODB.Recordset
    Sql = "select count(*) from Usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then Sql = ""
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    If Sql <> "" Then
        MsgBox "Ningun dato se ha generado", vbExclamation
    Else
        ListadoFormaPago = True
    End If
    Exit Function
EListadoFormaPago:
    MuestraError Err.Number, "ListadoFormaPago "
End Function


Public Sub cargaEmpresasTesor(ByRef Lis As ListView)
Dim Prohibidas As String
Dim IT
Dim Aux As String

    Set miRsAux = New ADODB.Recordset

    Prohibidas = DevuelveProhibidas
    
    Lis.ListItems.Clear
    Aux = "Select * from Usuarios.empresas where tesor=1"
    
    miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
    
        Aux = "|" & miRsAux!codempre & "|"
        If InStr(1, Prohibidas, Aux) = 0 Then
            Set IT = Lis.ListItems.Add
            IT.Key = "C" & miRsAux!codempre
            If vEmpresa.codempre = miRsAux!codempre Then IT.Checked = True
            IT.Text = miRsAux!nomempre
            IT.Tag = miRsAux!codempre
        End If
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

'-------------------------------------------------------------------------
'CCargar LISTVIEW con las mempresas de tesoreria
Private Function DevuelveProhibidas() As String
Dim I As Integer


    On Error GoTo EDevuelveProhibidas
    DevuelveProhibidas = ""

    I = vUsu.Codigo Mod 100
    miRsAux.Open "Select * from usuarios.usuarioempresasariconta WHERE codusu =" & I, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    DevuelveProhibidas = ""
    While Not miRsAux.EOF
        DevuelveProhibidas = DevuelveProhibidas & miRsAux.Fields(1) & "|"
        miRsAux.MoveNext
    Wend
    If DevuelveProhibidas <> "" Then DevuelveProhibidas = "|" & DevuelveProhibidas
    miRsAux.Close
    Exit Function
EDevuelveProhibidas:
    MuestraError Err.Number, "Cargando empresas prohibidas"
    Err.Clear
End Function


Public Function ComprobarCampoENlazado(ByRef T As TextBox, TDesc As TextBox, Tipo As String) As Byte

    T.Text = Trim(T.Text)
    If T.Text = "" Then
        ComprobarCampoENlazado = 0 'NO HA PUESTO NADA
        TDesc.Text = ""
        Exit Function
    End If
    
    Select Case Tipo
    Case "N"
        If Not IsNumeric(T.Text) Then
            MsgBox "El campo debe ser num�rico: " & T.Text, vbExclamation
            TDesc.Text = ""
            T.Text = ""
            ComprobarCampoENlazado = 1
        Else
            ComprobarCampoENlazado = 2
        End If
    End Select
        
End Function


Public Function RemesaSeleccionTipoRemesa(chkEfec As Boolean, chkPaga As Boolean, chkTalon As Boolean) As String
Dim C As String
    C = ""
    
    If chkEfec And chkPaga And chkTalon Then
        'LOS QUIERE TODOS, NO hacemos nada
        
    Else
    
        If Not chkEfec And Not chkPaga And Not chkTalon Then
            'NO QUIERE NINGUNO. Tampoco hago nada
            
        Else
            
            If chkEfec Then
                If chkPaga Then
                    C = " <> 3 "
                Else
                    If chkTalon Then
                        C = " <> 2 "
                    Else
                        C = " = 1" 'Solo efectos
                    End If
                End If
            Else
                If chkPaga Then
                    If chkTalon Then
                        C = " <> 1"
                    Else
                        C = " = 2 "
                    End If
                Else
                    C = " =3 "
                End If
            End If
        End If
    End If
    If C <> "" Then C = " tiporem  " & C
    RemesaSeleccionTipoRemesa = C
End Function

Public Function TextoAimporte(Importe As String) As Currency
Dim I As Integer
    If Importe = "" Then
        TextoAimporte = 0
    Else
        If InStr(1, Importe, ",") > 0 Then
            'Primero quitamos los puntos
            Do
                I = InStr(1, Importe, ".")
                If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
            Loop Until I = 0
            TextoAimporte = Importe
        
        
        Else
            'No tiene comas. El punto es el decimal
            TextoAimporte = TransformaPuntosComas(Importe)
        End If
    End If

End Function

Public Function EjecutarSQL(CadenaSQL As String) As Boolean
    On Error Resume Next
    Conn.Execute CadenaSQL
    If Err.Number <> 0 Then
         
         MuestraError Err.Number, "Error ejecutando SQL: " & vbCrLf & CadenaSQL, Err.Description
         EjecutarSQL = False
    Else
         EjecutarSQL = True
    End If
    
End Function



