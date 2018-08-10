VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CparametrosT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
'Estos parametros los lee de la contabilidad
Private mvarfechaini As Date 'copia local
Private mvarfechafin As Date 'copia local
Private mvarfechaAmbito As Date
Private mvarabononeg As Boolean
Private mvarautocoste As Boolean

Private mvarPagaresCtaPuente As Boolean
Private mvarTalonesCtaPuente As Boolean
Private mvarConfirmingCtaPuente As Boolean

Private mvarComprobarAlInicio As Boolean

'¡Estos otros los lee de su propia tabla de asignacion
Private mvarcontapag2 As Boolean
Private mvarAgrupaBancario As Boolean

'Desde paraemtros de Contabilidad
Private mvarEnvioDesdeOutlook As Boolean

'Para cuando elimina en recepcion de documentos, para que
Private mvarEliminaRecibidosRiesgo As Boolean

'Enero 2013
'   Si los dias para el seguro cuentan desde la fecha de fra o desde la fecha de vto
'   True: De factura ALZIRA
'   False: vto      HERBELCA
Private mvarFechaSeguroEsFra As Boolean

'Operaciones aseguradas
' FechaIniSeg Fecha que se inicia las operaciones aseguradas. las facturas anteriores NO las teiene en cuenta
' DiasMaxAviso`  'Dias maximo aviso falta pago'  DESDE/HASTA
' DiasMaxSiniestro 'Dias maximo aviso siniestro
'EN BD
' `DiasMaxAvisoDesde` `DiasMaxAvisoHasta `DiasMaxSiniestroDesde`
'`DiasMaxSiniestrohasta` `FechaIniSeg`
Private mvarTieneOperacionesAseguradas As Boolean
Private mvarFechaIniSeg As Date
Private mvarDiasMaxAvisoD As Integer
Private mvarDiasMaxAvisoH As Integer
Private mvarDiasMaxSiniestroD As Integer
Private mvarDiasMaxSiniestroH As Integer
Private mvarDiasAvisoDesdeProrroga As Integer


Private mvarNorma19xFechaVto As Boolean


Private mvarRecaudacionEjecutiva As Boolean


Private mvarNuevasNormasSEPA As Boolean

Private mvarIntereseCobrosTarjeta As Currency


Private mvarNorma57 As Byte ' 0.- NO LLEVA     1.- ESCALONA

Private mvarPagosConfirmingNorma As Boolean




Private mvarRemesasPorEntidad As Boolean


Private mvarNormasFormatoXML As Boolean
Private mvarPar_pen_apli As String

Private mvarForpaPaogTarjetaAplicaIntereses As Integer



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As Recordset
Private RC As Byte



Public Property Get AgrupaBancario() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     AgrupaBancario = mvarAgrupaBancario
End Property



Public Property Let AgrupaBancario(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarAgrupaBancario = vData
End Property




Public Property Get PagaresCtaPuente() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     PagaresCtaPuente = mvarPagaresCtaPuente
End Property



Public Property Let PagaresCtaPuente(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarPagaresCtaPuente = vData
End Property



Public Property Get TalonesCtaPuente() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     TalonesCtaPuente = mvarTalonesCtaPuente
End Property



Public Property Let TalonesCtaPuente(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarTalonesCtaPuente = vData
End Property


Public Property Get ConfirmingCtaPuente() As Boolean
     ConfirmingCtaPuente = mvarConfirmingCtaPuente
End Property

Public Property Let ConfirmingCtaPuente(ByVal vData As Boolean)
     mvarConfirmingCtaPuente = vData
End Property









Public Property Let fechaini(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarfechaini = vData
End Property

Public Property Get fechaini() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     fechaini = mvarfechaini
End Property


Public Property Let fechafin(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarfechafin = vData
End Property

Public Property Get fechafin() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     fechafin = mvarfechafin
End Property




Public Property Let fechaAmbito(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarfechaAmbito = vData
End Property

Public Property Get fechaAmbito() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     fechaAmbito = mvarfechaAmbito
End Property



Public Property Let abononeg(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarabononeg = vData
End Property

Public Property Get abononeg() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     abononeg = mvarabononeg
End Property



Public Property Let contapag(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarcontapag2 = vData
End Property

Public Property Get contapag() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     contapag = mvarcontapag2
End Property



Public Property Let autocoste(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarautocoste = vData
End Property

Public Property Get autocoste() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     autocoste = mvarautocoste
End Property



Public Property Let ComprobarAlInicio(ByVal vData As Boolean)
     mvarComprobarAlInicio = vData
End Property

Public Property Get ComprobarAlInicio() As Boolean
     ComprobarAlInicio = mvarComprobarAlInicio
End Property



Public Property Let EnvioDesdeOutlook(ByVal vData As Boolean)
     mvarEnvioDesdeOutlook = vData
End Property

Public Property Get EnvioDesdeOutlook() As Boolean
     EnvioDesdeOutlook = mvarEnvioDesdeOutlook
End Property




Public Property Let EliminaRecibidosRiesgo(ByVal vData As Boolean)
     mvarEliminaRecibidosRiesgo = vData
End Property

Public Property Get EliminaRecibidosRiesgo() As Boolean
     EliminaRecibidosRiesgo = mvarEliminaRecibidosRiesgo
End Property






'SOLO GET
Public Property Get TieneOperacionesAseguradas() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     TieneOperacionesAseguradas = mvarTieneOperacionesAseguradas
End Property

Public Property Get FechaIniSeg() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     FechaIniSeg = mvarFechaIniSeg
End Property
Public Property Let FechaIniSeg(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarFechaIniSeg = vData
End Property




Public Property Get DiasMaxAvisoD() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     DiasMaxAvisoD = mvarDiasMaxAvisoD
End Property
Public Property Let DiasMaxAvisoD(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarDiasMaxAvisoD = vData
End Property


'---------------------
Public Property Get DiasMaxAvisoH() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     DiasMaxAvisoH = mvarDiasMaxAvisoH
End Property
Public Property Let DiasMaxAvisoH(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarDiasMaxAvisoH = vData
End Property





'----------------------
Public Property Get DiasMaxSiniestroD() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     DiasMaxSiniestroD = mvarDiasMaxSiniestroD
End Property
Public Property Let DiasMaxSiniestroD(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarDiasMaxSiniestroD = vData
End Property


Public Property Get DiasMaxSiniestroH() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     DiasMaxSiniestroH = mvarDiasMaxSiniestroH
End Property
Public Property Let DiasMaxSiniestroH(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarDiasMaxSiniestroH = vData
End Property

Public Property Get DiasAvisoDesdeProrroga() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechafin
     DiasAvisoDesdeProrroga = mvarDiasAvisoDesdeProrroga
End Property
Public Property Let DiasAvisoDesdeProrroga(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechafin = 5
     mvarDiasAvisoDesdeProrroga = vData
End Property






Public Property Let Norma19xFechaVto(ByVal vData As Boolean)
     mvarNorma19xFechaVto = vData
End Property

Public Property Get Norma19xFechaVto() As Boolean
     Norma19xFechaVto = mvarNorma19xFechaVto
End Property





Public Property Let RecaudacionEjecutiva(ByVal vData As Boolean)
     mvarRecaudacionEjecutiva = vData
End Property

Public Property Get RecaudacionEjecutiva() As Boolean
     RecaudacionEjecutiva = mvarRecaudacionEjecutiva
End Property


Public Property Let FechaSeguroEsFra(ByVal vData As Boolean)
     mvarFechaSeguroEsFra = vData
End Property

Public Property Get FechaSeguroEsFra() As Boolean
     FechaSeguroEsFra = mvarFechaSeguroEsFra
End Property





'SOLO GET
Public Property Get NuevasNormasSEPA() As Boolean
     NuevasNormasSEPA = mvarNuevasNormasSEPA
End Property


'NAVARRES
'Los cobros a credito se pagaran con intereses.
'Grabara en gastos el valor de los intereses
' y cuando de por aceptado lo guarda
Public Property Get IntereseCobrosTarjeta2() As Currency
     IntereseCobrosTarjeta2 = mvarIntereseCobrosTarjeta
End Property
Public Property Let IntereseCobrosTarjeta2(ByVal vData As Currency)
     mvarIntereseCobrosTarjeta = vData
End Property
Public Property Get FormaPagoInterTarjeta() As Integer
     FormaPagoInterTarjeta = mvarForpaPaogTarjetaAplicaIntereses
End Property



'Solo en Base de datos
' 0.- NO LLEVA     1.- ESCALONA
Public Property Get Norma57() As Byte
     Norma57 = mvarNorma57
End Property

'Solo get
Public Property Get PagosConfirmingCaixa() As Boolean
     PagosConfirmingCaixa = mvarPagosConfirmingNorma
End Property


'No lleva campo en BD
'Hace un select leyendo parametros
Public Property Get RemesasPorEntidad() As Boolean
     RemesasPorEntidad = mvarRemesasPorEntidad
End Property

Public Property Get NormasFormatoXML() As Boolean
     NormasFormatoXML = mvarNormasFormatoXML
End Property


Public Property Get Par_pen_apli() As String
    Par_pen_apli = mvarPar_pen_apli
End Property




'Public Property Let nctafact(ByVal vData As String)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.nctafact = 5
'     mvarnctafact = vData
'End Property
'
'Public Property Get nctafact() As String
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.nctafact
'     nctafact = mvarnctafact
'End Property


Public Function Leer() As Byte

    On Error GoTo ELeer
    Leer = 1
    
    
    'Se divide en dos, una para leer de una tabla
    
    If Leer1 = 1 Then Exit Function
    
    If Leer2 = 1 Then Exit Function
    
    Leer = 0
    Exit Function
ELeer:
    
    MuestraError Err.Number
    Set Rs = Nothing

End Function




'****   LEER  ***
Private Function Leer1() As Byte
    
    Leer1 = 1
    
    SQL = "SELECT * FROM paramtesor"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic
    If Rs.EOF Then
        'ELSE
        Else
             
            mvarcontapag2 = (Rs!contapag = 1) ' TINYINT DEFAULT '0'  NOT NULL,
            ParametrosTesoreriaNuevos
            Leer1 = 0
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

End Function


Private Function ParametrosTesoreriaNuevos()
    On Error GoTo EP
    
    'Fijo algunos valores por defecto
    mvarTieneOperacionesAseguradas = False
    mvarRecaudacionEjecutiva = False
    mvarNuevasNormasSEPA = False
    mvarIntereseCobrosTarjeta = 0
    mvarForpaPaogTarjetaAplicaIntereses = -1
    mvarNorma57 = 0
    mvarNormasFormatoXML = False
    
    SQL = "generactrpar"
    mvarAgrupaBancario = (DBLet(Rs.Fields(SQL), "N") = 1)



    SQL = "contapagarepte"
    mvarPagaresCtaPuente = (DBLet(Rs.Fields(SQL), "N") = 1)
    

    SQL = "contatalonpte"
    mvarTalonesCtaPuente = (DBLet(Rs.Fields(SQL), "N") = 1)

    
    SQL = "contaconfirmpte"
    mvarConfirmingCtaPuente = (DBLet(Rs.Fields(SQL), "N") = 1)

    
    
    SQL = "comprobarinicio"
    mvarComprobarAlInicio = (DBLet(Rs.Fields(SQL), "N") = 1)
    
    SQL = "EliminaRecibidosRiesgo"
    mvarEliminaRecibidosRiesgo = (DBLet(Rs.Fields(SQL), "N") = 1)
    
    
    'Operaciones aseguradas
    ' `DiasMaxAvisoDesde` `DiasMaxAvisoHasta `DiasMaxSiniestroDesde`
    '`DiasMaxSiniestrohasta` `FechaIniSeg`
    
    SQL = "FechaIniSeg"
    If Not IsNull(Rs.Fields(SQL)) Then
        mvarTieneOperacionesAseguradas = True
        mvarFechaIniSeg = Rs.Fields(SQL)
        
        SQL = "DiasMaxAvisoDesde"
        mvarDiasMaxAvisoD = DBLet(Rs.Fields(SQL), "N")
        SQL = "DiasMaxAvisoHasta"
        mvarDiasMaxAvisoH = DBLet(Rs.Fields(SQL), "N")
        SQL = "DiasMaxSiniestroDesde"
        mvarDiasMaxSiniestroD = DBLet(Rs.Fields(SQL), "N")
        SQL = "DiasMaxSiniestrohasta"
        mvarDiasMaxSiniestroH = DBLet(Rs.Fields(SQL), "N")

        SQL = "DiasAvisoDesdeProrroga"
        mvarDiasAvisoDesdeProrroga = DBLet(Rs.Fields(SQL), "N")
        

    End If
    
    
    SQL = "Nor19xVto"
    mvarNorma19xFechaVto = (DBLet(Rs.Fields(SQL), "N") = 1)
    SQL = "RecaudacionEjecutiva"
    mvarRecaudacionEjecutiva = (DBLet(Rs.Fields(SQL), "N") = 1)
    SQL = "FechaAsegEsFra"
    mvarFechaSeguroEsFra = (DBLet(Rs.Fields(SQL), "N") = 1)
    
    'Noviembre 2013
    SQL = "Norma19_34Nueva"
    mvarNuevasNormasSEPA = (DBLet(Rs.Fields(SQL), "N") = 1)
    
    'Abril 2014
    SQL = "Intereses cobro tarjeta"
    mvarIntereseCobrosTarjeta = DBLet(Rs!InteresesCobrosTarjeta, "N")
    
    
    'Octubre 2014
    SQL = "Norma57. Pagos porventanilla"
    mvarNorma57 = DBLet(Rs!Norma57, "N")
    
    
    'Marzo 2015
    SQL = "Pagos confirming"
    mvarPagosConfirmingNorma = DBLet(Rs!CaixaConfirming, "N") = 1
    
    'Mayo 2015
    mvarRemesasPorEntidad = (DBLet(Rs!RemesasEntidad, "N") = 1)
    
    'Sept. 2015
    ' Si la SEPA se presenta enm XML o fichero texto
    SQL = "SEPA"
    mvarNormasFormatoXML = DBLet(Rs!XML, "N") = 1
    
    
    mvarPar_pen_apli = DBLet(Rs!Par_pen_apli, "T")
    
    SQL = "FormaPagoTarjeta"
    If Not IsNull(Rs!FormaPagoTarjeta) Then mvarForpaPaogTarjetaAplicaIntereses = Rs!FormaPagoTarjeta
  
    
    
    Exit Function
EP:
    MsgBox "Error leyendo parametros: " & SQL, vbExclamation
End Function

'****   LEER  ***
Private Function Leer2() As Byte
    
    
    Leer2 = 1
    SQL = "SELECT * FROM parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic
    If Rs.EOF Then
    
        'ELSE
        Else
            mvarfechaini = Rs!fechaini
            mvarfechafin = Rs!fechafin
            If IsNull(Rs!FechaActiva) Then
                mvarfechaAmbito = mvarfechaini
            Else
                mvarfechaAmbito = Rs!FechaActiva
            End If
            mvarabononeg = (Rs!abononeg = 1)
            'Centros de coste
            mvarautocoste = Rs!autocoste = 1
            ParametrosNuevos2
            
            Leer2 = 0
    End If

    Rs.Close
    

End Function



Private Sub ParametrosNuevos2()


     On Error Resume Next
     
   
     mvarEnvioDesdeOutlook = False
     
        'Controlamos el error
'            mvarCCenFacturas = False
'            mvarSubgrupo1 = ""
'            mvarSubgrupo2 = ""
'
'            On Error Resume Next
'            mvarCCenFacturas = DBLet(RS!CCenFacturas, "B")
'            mvarSubgrupo1 = DBLet(RS!Subgrupo1, "T")
'            mvarSubgrupo2 = DBLet(RS!Subgrupo2, "T")

        mvarEnvioDesdeOutlook = (DBLet(Rs!EnvioDesdeOutlook, "N") = 1)

   
    
       


        If Err.Number <> 0 Then Err.Clear
End Sub

