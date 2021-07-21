VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarfechaini As Date 'copia local
Private mvarfechafin As Date 'copia local
Private mvarautocoste As Boolean 'copia local
Private mvaremitedia As Boolean 'copia local
Private mvarlistahco As Boolean 'copia local
Private mvarnumdiapr As Integer 'copia local
Private mvarconcefpr As Integer 'copia local
Private mvarconceapr As Integer 'copia local
Private mvarnumdiacl As Integer 'copia local
Private mvarconcefcl As Integer 'copia local
Private mvarconceacl As Integer 'copia local
Private mvarlimimpcl As Currency 'copia local
'Private mvarconpresu As Boolean 'copia local
Private mvarperiodos As String 'copia local
Private mvargrupogto As String 'copia local
Private mvargrupovta As String 'copia local
Private mvarctaperga As String 'copia local4
Private mvarabononeg As Boolean 'copia local
Private mvargrupoord As String 'copia local
Private mvartinumfac As String 'copia local
Private mvarmodhcofa As Boolean 'copia local
Private mvaranofactu As Integer 'copia local
Private mvarperfactu As Integer 'copia local
Private mvarnctafact As String 'copia local
Private mvarCodiNume As Byte
Private mvarAsienActAuto As Boolean 'copia local

'Private "mvar Constructoras" As Boolean
Private mvarIvaEnFechaPago As Boolean

'Modificacion del 28 de Octubre de 2003
Private mvarContabilizaFactura As Boolean
Private mvarNroAriges As Integer

'Private mvardiremail As String 'copia local
'Private mvarSmtpHost As String 'copia local
'Private mvarSmtpUser As String 'copia local
'Private mvarSmtpPass As String 'copia local


'Modificacion 19 Diciembre


'Modificacion 2 Agosto 2004
Private mvarCCenFacturas  As Boolean
Private mvarSubgrupo1  As String 'copia local
Private mvarSubgrupo2 As String 'copia local


'Texto para la ampliacion de facturas.
' Descripcion concepto normal| ABONO
Private mvarAmpliacionFacurasCli As String
Private mvarAmpliacionFacurasPro As String



Private mvarEnlazaCtasMultibase As String

'27 Marzo 2007
Private mvarAgenciaViajes As Boolean

'Nuevo plan contable
Private mvarNuevoPlanContable As Boolean
Private mvarFechaActiva As Date 'copia local




'Para autoocion. Ellos el grupo de analitica 9 VAN a seguir llevandolo
Private mvarAutomocion As String

Private mvarTicketsEn340LetraSerie As Boolean
Private mvarPresentacion349Mensual As Boolean


Private mvarEnviarPorOutlook As Boolean


Private mvarGranEmpresa As Boolean   'Por si utiliza las cuentas 8 y 9 con otro fin ... para que no
                                     'se de gran empresa

'ALZIRA  Ene 2013
' Solo lleva GET.  Se cambia en BD
' Sera el numero contabilidad destino para los trasapaso de la cta bancos 555
Private mvarTraspasCtasBanco As Integer


Private mvarPathAyuda As String
Private mvarFecEjerAct As Boolean

Private mvarMailSoporte As String
Private mvarWebSoporte As String

Private mvarctahpdeudor As String 'cuenta de hda publica deudora
Private mvarctahpacreedor As String 'cuenta de hda publica acreedora
Private mvarnumdia303 As Integer 'copia local
Private mvarconce303 As Integer 'copia local




Private Sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As Recordset
Private RC As Byte


Private mvarEsMultiseccion As Boolean
Private mvarContabApteIva0 As Boolean





Private mvarTieneSII As Boolean
Private mvarFechaInicioSII As Date
Private mvarDiasAvisoSII As Integer

Private mvarVersionSII As String
Private mvarSII_Periodo_DesdeLiq As Boolean

Private mvarModificarIvaLineasFraCli As Boolean


Private mvarPathFicherosInteg As String

Private mvarPathAlbaranesProv As String
Private mvarPathFacturasProv As String
Private mvarEliminaPdfOriginal As Boolean


Private mvarAplicarFiltrosEnCuentas As Boolean   'Si valor en BD. De momento es que monti NO quiere filtros en cuentas

'Ene 19
Private mvarOrdenIvas303Aeat As String
Private mvarTextoInmoSubencionado As String


'No lleva campo en BD. Cuando lee los parametros hace el select
Private mvarFechaMinimaEnHlinapu As Date

Private mvarTipoIntegracionSeleccionable As Byte
    '0 nada
    '1 TAXCO. Tipo SAGE
    '2 Coarval


Private mvarRectificativasSeparadas303 As Boolean
Private mvarExcluirBasesIvaCeroRecibidas303 As Boolean



Private mvarComprobarAsientosIncio As Boolean
Private mvarcaptionFraProFecRegis As String

Private mvarSII_ProvDesdeFechaRecepcion As Boolean
    '  Siempre que no tengan marcado el parametro de mvarSII_Periodo_DesdeLiq
    'TRUE:  Como lo hacia antes de poner la fecha de registro contable, trata la fecha de rececpion
    'FALSE:  Como ahora. Va desde fecha registro contable
    
    
    
Private mvarinscritoDeclarDUA As Boolean
    
Private mvarRemarcaTotalExtr As Boolean
Private mvarcolCtasConSaldos As Boolean
Private mvarAnyadeNumFaproAmpliacApunte As Boolean

Private mvarObsrvacionesLineasFrapro As Boolean



'------------------------------------------------
'Propiedades del modulo Cparametros
'------------------------------------------------

Public Property Let IvaEnFechaPago(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarIvaEnFechaPago = vData
End Property

Public Property Get IvaEnFechaPago() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     IvaEnFechaPago = mvarIvaEnFechaPago
End Property


Public Property Let AgenciaViajes(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarAgenciaViajes = vData
End Property

Public Property Get AgenciaViajes() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     AgenciaViajes = mvarAgenciaViajes
End Property


'''Public Property Let diario43(ByVal vData As Integer)
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado izquierdo de la asignacion
''''Syntax: X.numdiapr = 5
'''     mvardiario43 = vData
'''End Property
'''
'''Public Property Get diario43() As Integer
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado derecho de la asignacion
''''Syntax: Debug.Print X.numdiapr
'''     diario43 = mvardiario43
'''End Property
'''
'''
'''Public Property Let conce43(ByVal vData As Integer)
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado izquierdo de la asignacion
''''Syntax: X.concefpr = 5
'''     mvarconce43 = vData
'''End Property
'''
'''Public Property Get conce43() As Integer
''''se usa cuando se asigna un valor a la propiedad, en el
''''lado derecho de la asignacion
''''Syntax: Debug.Print X.concefpr
'''     conce43 = mvarconce43
'''End Property



Public Property Let CCenFacturas(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarCCenFacturas = vData
End Property

Public Property Get CCenFacturas() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     CCenFacturas = mvarCCenFacturas
End Property

Public Property Let Subgrupo1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarSubgrupo1 = vData
End Property

Public Property Get Subgrupo1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     Subgrupo1 = mvarSubgrupo1
End Property

Public Property Let Subgrupo2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarSubgrupo2 = vData
End Property

Public Property Get Subgrupo2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     Subgrupo2 = mvarSubgrupo2
End Property

'Private mvarAmpliacionFacurasCli As String
'Private mvarAmpliacionFacurasPro As String

Public Property Get AmpliacionFacurasCli() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     AmpliacionFacurasCli = mvarAmpliacionFacurasCli
End Property

Public Property Get AmpliacionFacurasPro() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     AmpliacionFacurasPro = mvarAmpliacionFacurasPro
End Property



Public Property Let ContabilizaFactura(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarContabilizaFactura = vData
End Property

Public Property Get ContabilizaFactura() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     ContabilizaFactura = mvarContabilizaFactura
End Property



Public Property Let NroAriges(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarNroAriges = vData
End Property

Public Property Get NroAriges() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     NroAriges = mvarNroAriges
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


Public Property Let autocoste(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.autocoste = 5
     mvarautocoste = vData
End Property

Public Property Get autocoste() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.autocoste
     autocoste = mvarautocoste
End Property


Public Property Let emitedia(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.emitedia = 5
     mvaremitedia = vData
End Property

Public Property Get emitedia() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.emitedia
     emitedia = mvaremitedia
End Property


Public Property Let listahco(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.listahco = 5
     mvarlistahco = vData
End Property

Public Property Get listahco() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.listahco
     listahco = mvarlistahco
End Property


Public Property Let numdiapr(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdiapr = 5
     mvarnumdiapr = vData
End Property

Public Property Get numdiapr() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdiapr
     numdiapr = mvarnumdiapr
End Property


Public Property Let concefpr(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.concefpr = 5
     mvarconcefpr = vData
End Property

Public Property Get concefpr() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.concefpr
     concefpr = mvarconcefpr
End Property


Public Property Let conceapr(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.conceapr = 5
     mvarconceapr = vData
End Property

Public Property Get conceapr() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.conceapr
     conceapr = mvarconceapr
End Property


Public Property Let numdiacl(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdiacl = 5
     mvarnumdiacl = vData
End Property

Public Property Get numdiacl() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdiacl
     numdiacl = mvarnumdiacl
End Property


Public Property Let concefcl(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.concefcl = 5
     mvarconcefcl = vData
End Property

Public Property Get concefcl() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.concefcl
     concefcl = mvarconcefcl
End Property


Public Property Let conceacl(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.conceacl = 5
     mvarconceacl = vData
End Property

Public Property Get conceacl() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.conceacl
     conceacl = mvarconceacl
End Property
Public Property Let EnlazaCtasMultibase(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.limimpcl = 5
     mvarEnlazaCtasMultibase = vData
End Property

Public Property Get EnlazaCtasMultibase() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.limimpcl
     EnlazaCtasMultibase = mvarEnlazaCtasMultibase
End Property



Public Property Let limimpcl(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.limimpcl = 5
     mvarlimimpcl = vData
End Property

Public Property Get limimpcl() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.limimpcl
     limimpcl = mvarlimimpcl
End Property


'Public Property Let conpresu(ByVal vData As Boolean)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.conpresu = 5
'     mvarconpresu = vData
'End Property
'
'Public Property Get conpresu() As Boolean
''se usa cuando se asigna un valor a la propiedad, en el
''lado derecho de la asignacion
''Syntax: Debug.Print X.conpresu
'     conpresu = mvarconpresu
'End Property


Public Property Let periodos(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarperiodos = vData
End Property

Public Property Get periodos() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     periodos = mvarperiodos
End Property


Public Property Let GrupoGto(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupogto = 5
     mvargrupogto = vData
End Property

Public Property Get GrupoGto() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupogto
     GrupoGto = mvargrupogto
End Property


Public Property Let GrupoVta(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupovta = 5
     mvargrupovta = vData
End Property

Public Property Get GrupoVta() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupovta
     GrupoVta = mvargrupovta
End Property


Public Property Let ctaperga(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ctaperga = 5
     mvarctaperga = vData
End Property

Public Property Get ctaperga() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ctaperga
     ctaperga = mvarctaperga
End Property


Public Property Let abononeg(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarabononeg = vData
End Property

Public Property Get abononeg() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     abononeg = mvarabononeg
End Property



Public Property Let NuevoPlanContable(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarNuevoPlanContable = vData
End Property

Public Property Get NuevoPlanContable() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     NuevoPlanContable = mvarNuevoPlanContable
End Property




Public Property Let FechaActiva(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarFechaActiva = vData
End Property

Public Property Get FechaActiva() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     FechaActiva = mvarFechaActiva
End Property



Public Property Let GrupoOrd(ByVal vData As String)
     mvargrupoord = vData
End Property

Public Property Get GrupoOrd() As String
     GrupoOrd = mvargrupoord
End Property


Public Property Let tinumfac(ByVal vData As String)
     mvartinumfac = vData
End Property

Public Property Get tinumfac() As String
     tinumfac = mvartinumfac
End Property


Public Property Let modhcofa(ByVal vData As Boolean)
     mvarmodhcofa = vData
End Property

Public Property Get modhcofa() As Boolean
     modhcofa = mvarmodhcofa
End Property



Public Property Let Anofactu(ByVal vData As Integer)
     mvaranofactu = vData
End Property

Public Property Get Anofactu() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.anofactu
     Anofactu = mvaranofactu
End Property


Public Property Let perfactu(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.perfactu = 5
     mvarperfactu = vData
End Property

Public Property Get perfactu() As Integer
     perfactu = mvarperfactu
End Property


'Public Property Let nctafac_t(ByVal vData As String)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.nctafact = 5
'     mvarnctafact = vData
'End Property

Public Property Get nctafact() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nctafact
     nctafact = mvarnctafact
End Property


Public Property Let AsienActAuto(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.AsienActAuto = 5
     mvarAsienActAuto = vData
End Property

Public Property Get AsienActAuto() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.AsienActAuto
     AsienActAuto = mvarAsienActAuto
End Property


Public Property Let CodiNume(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarCodiNume = vData
End Property

Public Property Get CodiNume() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     CodiNume = mvarCodiNume
End Property




Public Property Let Automocion(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarAutomocion = vData
End Property

Public Property Get Automocion() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     Automocion = mvarAutomocion
End Property




Public Property Let TicketsEn340LetraSerie(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarTicketsEn340LetraSerie = vData
End Property

Public Property Get TicketsEn340LetraSerie() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     TicketsEn340LetraSerie = mvarTicketsEn340LetraSerie
End Property





Public Property Let Presentacion349Mensual(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarPresentacion349Mensual = vData
End Property

Public Property Get Presentacion349Mensual() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     Presentacion349Mensual = mvarPresentacion349Mensual
End Property



Public Property Get GranEmpresa() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     GranEmpresa = mvarGranEmpresa
End Property



Public Property Let GranEmpresa(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi10 = 5
     mvarGranEmpresa = vData
End Property


Public Property Let EnviarPorOutlook(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.abononeg = 5
     mvarEnviarPorOutlook = vData
End Property

Public Property Get EnviarPorOutlook() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     EnviarPorOutlook = mvarEnviarPorOutlook
End Property



'Solo GET
Public Property Get TraspasCtasBanco() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.abononeg
     TraspasCtasBanco = mvarTraspasCtasBanco
End Property


Public Property Let PathAyuda(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fechaini = 5
     mvarPathAyuda = vData
End Property

Public Property Get PathAyuda() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fechaini
     PathAyuda = mvarPathAyuda
End Property


Public Property Let FecEjerAct(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarFecEjerAct = vData
End Property

Public Property Get FecEjerAct() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     FecEjerAct = mvarFecEjerAct
End Property


Public Property Let MailSoporte(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarMailSoporte = vData
End Property

Public Property Get MailSoporte() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     MailSoporte = mvarMailSoporte
End Property

Public Property Let WebSoporte(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarWebSoporte = vData
End Property

Public Property Get WebSoporte() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     WebSoporte = mvarWebSoporte
End Property


Public Property Get EsMultiseccion() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     EsMultiseccion = mvarEsMultiseccion
End Property



Public Property Let EsMultiseccion(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi10 = 5
     mvarEsMultiseccion = vData
End Property


' Contabiliza o no el apunte de iva 0
Public Property Get ContabApteIva0() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     ContabApteIva0 = mvarContabApteIva0
End Property


Public Property Let ContabApteIva0(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi10 = 5
     mvarContabApteIva0 = vData
End Property


Public Property Let CtaHPDeudor(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarctahpdeudor = vData
End Property

Public Property Get CtaHPDeudor() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     CtaHPDeudor = mvarctahpdeudor
End Property

Public Property Let CtaHPAcreedor(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.periodos = 5
     mvarctahpacreedor = vData
End Property

Public Property Get CtaHPAcreedor() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     CtaHPAcreedor = mvarctahpacreedor
End Property


Public Property Let numdia303(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdiapr = 5
     mvarnumdia303 = vData
End Property

Public Property Get numdia303() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdiapr
     numdia303 = mvarnumdia303
End Property


Public Property Let conce303(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.concefpr = 5
     mvarconce303 = vData
End Property

Public Property Get conce303() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.concefpr
     conce303 = mvarconce303
End Property



Public Property Get SIITiene() As Boolean
     SIITiene = mvarTieneSII
End Property
Public Property Get SIIFechaInicio() As Date
     SIIFechaInicio = mvarFechaInicioSII
End Property
Public Property Get SIIDiasAviso() As Integer
     SIIDiasAviso = mvarDiasAvisoSII
End Property

Public Property Get SII_Version() As String
     SII_Version = mvarVersionSII
End Property

Public Property Get ModificarIvaLineasFraCli() As Boolean
     ModificarIvaLineasFraCli = mvarModificarIvaLineasFraCli
End Property

Public Property Get SII_Periodo_DesdeLiq() As Boolean
     SII_Periodo_DesdeLiq = mvarSII_Periodo_DesdeLiq
End Property


Public Property Let PathFicherosInteg(ByVal vData As String)
     mvarPathFicherosInteg = vData
End Property

Public Property Get PathFicherosInteg() As String
     PathFicherosInteg = mvarPathFicherosInteg
End Property

Public Property Get AplicarFiltrosEnCuentas() As Boolean
     AplicarFiltrosEnCuentas = mvarAplicarFiltrosEnCuentas
End Property


'SOLO GET
Public Property Get FechaMinimaEnHlinapu() As Date
     FechaMinimaEnHlinapu = mvarFechaMinimaEnHlinapu
End Property

Public Property Get PathAlbaranesProv() As String
     PathAlbaranesProv = mvarPathAlbaranesProv
End Property

Public Property Get PathFacturasProv() As String
     PathFacturasProv = mvarPathFacturasProv
End Property

Public Property Get EliminaPdfOriginal() As Boolean
     EliminaPdfOriginal = mvarEliminaPdfOriginal
End Property


Public Property Get OrdenIvas303Aeat() As String
     OrdenIvas303Aeat = mvarOrdenIvas303Aeat
End Property

Public Property Get TextoInmoSubencionado() As String
     TextoInmoSubencionado = mvarTextoInmoSubencionado
End Property

'0 nada       1 TAXCO. Tipo SAGE        2 Coarval
Public Property Get TipoIntegracionSeleccionable() As Byte
     TipoIntegracionSeleccionable = mvarTipoIntegracionSeleccionable
End Property


'Descuadre Asientos al princio

Public Property Let ComprobarAsientosIncio(ByVal vData As Boolean)
     mvarComprobarAsientosIncio = vData
End Property

Public Property Get ComprobarAsientosIncio() As Boolean
     ComprobarAsientosIncio = mvarComprobarAsientosIncio
End Property


Public Property Get RectificativasSeparadas303() As Boolean
     RectificativasSeparadas303 = mvarRectificativasSeparadas303
End Property

Public Property Get ExcluirBasesIvaCeroRecibidas303() As Boolean
     ExcluirBasesIvaCeroRecibidas303 = mvarExcluirBasesIvaCeroRecibidas303
End Property

Public Property Get captionFraProFecRegis() As String
     captionFraProFecRegis = mvarcaptionFraProFecRegis
End Property


Public Property Get SII_ProvDesdeFechaRecepcion() As Boolean
     SII_ProvDesdeFechaRecepcion = mvarSII_ProvDesdeFechaRecepcion
End Property

Public Property Get InscritoDeclarDUA() As Boolean
     InscritoDeclarDUA = mvarinscritoDeclarDUA
End Property

Public Property Get RemarcaTotalExtr() As Boolean
     RemarcaTotalExtr = mvarRemarcaTotalExtr
End Property

Public Property Get colCtasConSaldos() As Boolean
     colCtasConSaldos = mvarcolCtasConSaldos
End Property


Public Property Get AnyadeNumFaproAmpliacApunte() As Boolean
    AnyadeNumFaproAmpliacApunte = mvarAnyadeNumFaproAmpliacApunte
End Property


Public Property Get ObsrvacionesLineasFrapro() As Boolean
    ObsrvacionesLineasFrapro = mvarObsrvacionesLineasFrapro
End Property



'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------



'****   LEER  ***
Public Function Leer() As Byte
    On Error GoTo Err3
    Sql = "SELECT * FROM parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic
    If Rs.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarfechaini = Rs!fechaini
            mvarfechafin = Rs!fechafin
            mvarautocoste = Rs!autocoste
            mvaremitedia = Rs!emitedia
            mvarlistahco = Rs!listahco
            mvarnumdiapr = Rs!numdiapr
            mvarconcefpr = DBLet(Rs!concefpr, "N")
            mvarconceapr = DBLet(Rs!conceapr, "N")
            mvarnumdiacl = DBLet(Rs!numdiacl, "N")
            mvarconcefcl = DBLet(Rs!concefcl, "N")
            mvarconceacl = DBLet(Rs!conceacl, "N")
            mvarlimimpcl = Rs!limimpcl
            'mvarconpresu = RS!conpresu
            mvarperiodos = Rs!periodos
            mvargrupogto = DBLet(Rs!GrupoGto)
            mvargrupovta = DBLet(Rs!GrupoVta)
            mvarctaperga = DBLet(Rs!ctaperga)
            mvarabononeg = Rs!abononeg
            mvargrupoord = DBLet(Rs!GrupoOrd)
            mvartinumfac = Rs!tinumfac
            mvarmodhcofa = Rs!modhcofa
            mvaranofactu = Rs!Anofactu
            mvarperfactu = Rs!perfactu
            mvarnctafact = Rs!nctafact
            mvarAsienActAuto = Rs!AsienActAuto
            mvarCodiNume = Rs!CodiNume
            
            'Modificacion del 28 de Octubre de 2003
            mvarContabilizaFactura = DBLet(Rs!ContabilizaFact, "N")
            
            mvarIvaEnFechaPago = (Rs!Constructoras = 1)
            
            '++
            mvarPathAyuda = DBLet(Rs!PathAyuda, "T")
            mvarFecEjerAct = DBLet(Rs!FecEjerAct, "N")
            mvarNroAriges = DBLet(Rs!NroAriges, "N")
            
            mvarWebSoporte = DBLet(Rs!WebSoporte, "T")
            mvarMailSoporte = DBLet(Rs!MailSoporte, "T")
            
            '++
            mvarctahpdeudor = DBLet(Rs!CtaHPDeudor, "T")
            mvarctahpacreedor = DBLet(Rs!CtaHPAcreedor, "T")
            mvarnumdia303 = DBLet(Rs!diario303, "N")
            mvarconce303 = DBLet(Rs!conce303, "N")
            
            ParametrosNuevos

            mvarTieneSII = False
            LeerDatosSII


            If ASignarConceptosFacturas Then Leer = 0
    End If

    Rs.Close
    
    
                    
    If Leer = 0 Then MinimaFechaEnApuntes
    
    
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    On Error Resume Next
    CerrarRs Rs
    Leer = 1
    Err.Clear
End Function



Private Sub ParametrosNuevos()
Dim Nuevos As String

                '" agosto 2004
            mvarCCenFacturas = False
            mvarSubgrupo1 = ""
            mvarSubgrupo2 = ""
            mvarEnlazaCtasMultibase = ""
            mvarAgenciaViajes = False
            mvarNuevoPlanContable = False
            mvarFechaActiva = mvarfechaini   'Por fecto se carga con la fecha inicio
            mvarAutomocion = ""
            mvarEnviarPorOutlook = False
            mvarEsMultiseccion = False
            mvarContabApteIva0 = False
            mvarModificarIvaLineasFraCli = False
            mvarEsMultiseccion = False
            mvarPathFicherosInteg = ""
            mvarPathAlbaranesProv = ""
            mvarPathFacturasProv = ""
            mvarEliminaPdfOriginal = False
            mvarComprobarAsientosIncio = False
            mvarRectificativasSeparadas303 = False
            mvarExcluirBasesIvaCeroRecibidas303 = False
            
            '0 nada     1 TAXCO. Tipo SAGE      2 Coarval
            mvarTipoIntegracionSeleccionable = 0
            
            
            mvarinscritoDeclarDUA = False
            mvarRemarcaTotalExtr = False
            mvarcolCtasConSaldos = False
            mvarAnyadeNumFaproAmpliacApunte = False
            mvarObsrvacionesLineasFrapro = False


            
            Nuevos = ""
            
            
            On Error Resume Next
            
            mvarCCenFacturas = DBLet(Rs!CCenFacturas, "B")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "CCenFacturas"
            
            mvarSubgrupo1 = DBLet(Rs!Subgrupo1, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Subgrupo1"
            
            mvarSubgrupo2 = DBLet(Rs!Subgrupo2, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Subgrupo2"
            
            If Not IsNull(Rs!FechaActiva) Then mvarFechaActiva = Rs!FechaActiva
            If Err.Number <> 0 Then AgregarMsg Nuevos, "FechaActiva"
            
            
            mvarEnlazaCtasMultibase = DBLet(Rs!enlaza_cta, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "enlaza_cta"
            
            mvarAgenciaViajes = (DBLet(Rs!AgenciaViajes, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "AgenciaViajes"
            
            mvarNuevoPlanContable = (DBLet(Rs!NuevaPlan, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "NuevaPlan"
            
            mvarAutomocion = DBLet(Rs!Automocion, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Automocion"
            
            
            
            mvarTicketsEn340LetraSerie = (DBLet(Rs!TicketsEn340LetraSerie, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "TicketsEn340LetraSerie"
            
            mvarPresentacion349Mensual = (DBLet(Rs!Presentacion349Mensual, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Presentacion349Mensual"
            
            
            mvarEnviarPorOutlook = (DBLet(Rs!EnvioDesdeOutlook, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "EnvioDesdeOutlook"
            
            mvarGranEmpresa = (DBLet(Rs!GranEmpresa, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "GranEmpresa"
            
            mvarTraspasCtasBanco = DBLet(Rs!TraspasCtasBanco, "N")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Trapaso bancos entre secciones"
            
            mvarEsMultiseccion = (DBLet(Rs!EsMultiseccion, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Multisección"
            
            mvarContabApteIva0 = (DBLet(Rs!ContabApteIva0, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Contabilizacion Apte Iva 0"
            
            mvarModificarIvaLineasFraCli = (DBLet(Rs!MoficaImporIVALineas, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Permite modificar iva lineas factura cliente"
            
            mvarPathFicherosInteg = DBLet(Rs!PathFicherosInteg, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Carpeta fichero integraciones"
            
           
            mvarPathAlbaranesProv = DBLet(Rs!PathAlbaranesProv, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Carpeta fichero albaranes"
            
            mvarPathFacturasProv = DBLet(Rs!PathFacturasProv, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Carpeta ficheros facturas PDF"
            
            mvarEliminaPdfOriginal = (DBLet(Rs!EliminaPdfOriginal, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Parametro eliminar albaranes/facturas PDF"
                
                    
            
            mvarOrdenIvas303Aeat = DBLet(Rs!OrdenIvas303Aeat, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "IVAS 303"
            If mvarOrdenIvas303Aeat = "" Then mvarOrdenIvas303Aeat = "4|10|21|"
            
            mvarTextoInmoSubencionado = DBLet(Rs!TextoInmoSubencionado, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Lbl subvencionado"
            If mvarTextoInmoSubencionado = "" Then mvarTextoInmoSubencionado = "Subvencionado"
            
            mvarAplicarFiltrosEnCuentas = True
            
            
            '                         1 TAXCO. Tipo SAGE      2 Coarval
            If InStr(1, vEmpresa.nomresum, "TAXCO") > 0 Then mvarTipoIntegracionSeleccionable = 1
                            
            mvarComprobarAsientosIncio = DBLet(Rs!ComprobarAsientosIncio, "N") = 1
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Comprobar descuadre inicio"
            
            mvarRectificativasSeparadas303 = DBLet(Rs!RectificativasSeparadas303, "N") = 1
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Rectificativas Separadas 303"
                        
            mvarExcluirBasesIvaCeroRecibidas303 = DBLet(Rs!ExcluirBasesIvaCeroRecibidas303, "N") = 1
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Excluir en 303 bases cero Fra recibidas "
            
            
            
            mvarcaptionFraProFecRegis = DBLet(Rs!captionFraProFecRegis, "T")
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Caption factura proveedor fecha registro"
            
            If mvarcaptionFraProFecRegis = "" Then mvarcaptionFraProFecRegis = "F.recepción"
                        
            mvarinscritoDeclarDUA = DBLet(Rs!InscritoDeclarDUA, "N") = 1
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Inscrito registro AEAT declar. DUA"
            
            mvarRemarcaTotalExtr = DBLet(Rs!RemarcaTotalExtr, "N") = 1
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Remarca total consulta extracto"
            
            mvarcolCtasConSaldos = (DBLet(Rs!colCtasConSaldos, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Formulario cuentas con saldos"
            
            mvarAnyadeNumFaproAmpliacApunte = (DBLet(Rs!anyadenumfacproContab, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Añadir NºFac pro en ampliacion"
            
            mvarObsrvacionesLineasFrapro = (DBLet(Rs!obsrlinfact, "N") = 1)
            If Err.Number <> 0 Then AgregarMsg Nuevos, "Observaciones lineas factura proveedor"
            
            


            
            
        If Err.Number <> 0 Then Err.Clear
        If Nuevos <> "" Then MsgBox "Faltan parámetros. El programa continuará. Avise soporte técnico" & vbCrLf & vbCrLf & Nuevos, vbExclamation
    
End Sub

Private Sub AgregarMsg(Cadena As String, Trozo As String)
    
        Cadena = Cadena & vbCrLf & Trozo
        Err.Clear
End Sub

Private Function ASignarConceptosFacturas() As Boolean
Dim Aux As String
    
    On Error GoTo EASignarConceptosFacturas
    'Private mvarconcefpr As Integer
    'Private mvarconceapr As Integer
    '
    'Private mvarconcefcl As Integer
    'Private mvarconceacl As Integer
    
    ASignarConceptosFacturas = False
    
    'Private mvarAmpliacionFacurasCli As String
    'Private mvarAmpliacionFacurasPro As String
    
    'FACTURAS PROVEEDORES
    Aux = "tipoconce"
    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconcefpr), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasPro = Sql & "|"
        
    Aux = "tipoconce"
    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconceapr), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasPro = mvarAmpliacionFacurasPro & Sql & "|"
    
    'FACTURAS CLIENTES
    Aux = "tipoconce"
    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconcefcl), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasCli = Sql & "|"
        
    Aux = "tipoconce"
    Sql = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconceacl), "N", Aux)
    If Sql = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasCli = mvarAmpliacionFacurasCli & Sql & "|"
    
    
    
    
    ASignarConceptosFacturas = True
    Exit Function
EASignarConceptosFacturas:
    MuestraError Err.Number, "ASignar Conceptos Facturas"
End Function





Private Sub MinimaFechaEnApuntes()

    
    On Error GoTo EmvarCuentasBloqueadas
    
    mvarFechaMinimaEnHlinapu = mvarfechaini
    Set Rs = Nothing
    Set Rs = New ADODB.Recordset
    Sql = "Select min(fechaent) from hlinapu WHERE fechaent<'" & Format(mvarfechaini, "yyyy-mm-dd") & "'"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then mvarFechaMinimaEnHlinapu = Rs.Fields(0)
    End If
    Rs.Close
    
    
EmvarCuentasBloqueadas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener CuentasBloqueadas"


End Sub


Private Sub LeerDatosSII()

On Error GoTo eLeerDatosSII

    mvarTieneSII = False
    mvarSII_ProvDesdeFechaRecepcion = False
    mvarSII_Periodo_DesdeLiq = False
    
    If Not IsNull(Rs!SIIFechaInicio) Then
        mvarFechaInicioSII = Rs!SIIFechaInicio
        mvarDiasAvisoSII = Rs!SIIDiasAviso
        mvarTieneSII = True
        mvarVersionSII = Rs!VersionSII
        mvarSII_Periodo_DesdeLiq = DBLet(Rs!SII_Periodo_DesdeLiq, "N") = 1
        mvarSII_ProvDesdeFechaRecepcion = DBLet(Rs!SII_FraPro_Fecregis, "N") = 1
    End If
    
    
    
    
    
    
    


    Exit Sub
eLeerDatosSII:
    mvarTieneSII = False
    Conn.Errors.Clear
    Err.Clear
End Sub


Public Sub FijarAplicarFiltrosEnCuentas(ByRef NomEmpresa As String)
    mvarAplicarFiltrosEnCuentas = True
    If InStr(1, NomEmpresa, "MONTI") > 0 Then mvarAplicarFiltrosEnCuentas = False
End Sub

Private Sub Class_Terminate()
    Set Rs = Nothing
End Sub
