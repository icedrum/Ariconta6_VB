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
Private mvarctaperga As String 'copia local
Private mvarabononeg As Boolean 'copia local
Private mvargrupoord As String 'copia local
Private mvartinumfac As String 'copia local
Private mvarmodhcofa As Boolean 'copia local
Private mvaranofactu As Integer 'copia local
Private mvarperfactu As Integer 'copia local
Private mvarnctafact As String 'copia local
Private mvarCodiNume As Byte
Private mvarAsienActAuto As Boolean 'copia local
Private mvarConstructoras As Boolean


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

'Bloqueo de cuentas a partir de una fecha. EN la introduccion (facturas,astos,inmov, presup,banco....)
'Para ello tendremos
'                      - mvarCuentasBloq    codmacta:fecbloq|codmacta:....
'                   Habra una funcion privada que rellenara esta variable (DE SOLO LECTURA)
Private mvarCuentasBloqueadas As String



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




Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As Recordset
Private RC As Byte


Private mvarEsMultiseccion As Boolean
Private mvarContabApteIva0 As Boolean





Private mvarTieneSII As Boolean
Private mvarFechaInicioSII As Date
Private mvarDiasAvisoSII As Integer






'------------------------------------------------
'Propiedades del modulo Cparametros
'------------------------------------------------

Public Property Let Constructoras(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.contabil = 5
     mvarConstructoras = vData
End Property

Public Property Get Constructoras() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.contabil
     Constructoras = mvarConstructoras
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



'Solo GET

Public Property Get CuentasBloqueadas() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.periodos
     CuentasBloqueadas = mvarCuentasBloqueadas
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


Public Property Let grupogto(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupogto = 5
     mvargrupogto = vData
End Property

Public Property Get grupogto() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupogto
     grupogto = mvargrupogto
End Property


Public Property Let grupovta(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupovta = 5
     mvargrupovta = vData
End Property

Public Property Get grupovta() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupovta
     grupovta = mvargrupovta
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



''''Public Property Let intfaccl(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intfaccl = 5
''''     mvarintfaccl = vData
''''End Property
''''
''''Public Property Get intfaccl() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intfaccl
''''     intfaccl = mvarintfaccl
''''End Property
''''
''''
''''Public Property Let intfacpr(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intfacpr = 5
''''     mvarintfacpr = vData
''''End Property
''''
''''Public Property Get intfacpr() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intfacpr
''''     intfacpr = mvarintfacpr
''''End Property
''''
''''
''''Public Property Let intliapu(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intliapu = 5
''''     mvarintliapu = vData
''''End Property
''''
''''Public Property Get intliapu() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intliapu
''''     intliapu = mvarintliapu
''''End Property
''''
''''
''''Public Property Let intlihap(ByVal vData As String)
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado izquierdo de la asignacion
'''''Syntax: X.intlihap = 5
''''     mvarintlihap = vData
''''End Property
''''
''''Public Property Get intlihap() As String
'''''se usa cuando se asigna un valor a la propiedad, en el
'''''lado derecho de la asignacion
'''''Syntax: Debug.Print X.intlihap
''''     intlihap = mvarintlihap
''''End Property


Public Property Let grupoord(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.grupoord = 5
     mvargrupoord = vData
End Property

Public Property Get grupoord() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.grupoord
     grupoord = mvargrupoord
End Property


Public Property Let tinumfac(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tinumfac = 5
     mvartinumfac = vData
End Property

Public Property Get tinumfac() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tinumfac
     tinumfac = mvartinumfac
End Property


Public Property Let modhcofa(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.modhcofa = 5
     mvarmodhcofa = vData
End Property

Public Property Get modhcofa() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.modhcofa
     modhcofa = mvarmodhcofa
End Property



Public Property Let anofactu(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.anofactu = 5
     mvaranofactu = vData
End Property

Public Property Get anofactu() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.anofactu
     anofactu = mvaranofactu
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


Public Property Let nctafact(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nctafact = 5
     mvarnctafact = vData
End Property

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


'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------



''****   AGREGAR  ***
'Public Function Agregar() As Byte
'    On Error GoTo Err1
'    SQL = "SELECT * FROM parametros"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " fechaini = " & mvarfechaini
'   ' Set Rs = DB.OpenRecordSet(Sql, dbOpenDynaset)
'    If Rs.RecordCount > 0 Then
'        Mens = "Ya existe un Registro en la Tabla parametros con esos valores."
'        MsgBox Mens, vbExclamation
'        Agregar = 1
'        'ELSE
'        Else
'            Rs.AddNew
'            Rs!fechaini = mvarfechaini
'            Rs!fechafin = mvarfechafin
'            Rs!autocoste = mvarautocoste
'            Rs!emitedia = mvaremitedia
'            Rs!contabil = mvarcontabil
'            Rs!listahco = mvarlistahco
'            Rs!numdiapr = mvarnumdiapr
'            Rs!concefpr = mvarconcefpr
'            Rs!conceapr = mvarconceapr
'            Rs!numdiacl = mvarnumdiacl
'            Rs!concefcl = mvarconcefcl
'            Rs!conceacl = mvarconceacl
'            Rs!limimpcl = mvarlimimpcl
'            'RS!conpresu = mvarconpresu
'            Rs!periodos = mvarperiodos
'            Rs!grupogto = mvargrupogto
'            Rs!grupovta = mvargrupovta
'            Rs!ctaperga = mvarctaperga
'            Rs!abononeg = mvarabononeg
'            Rs!amortiza = mvaramortiza
'            Rs!intfaccl = mvarintfaccl
'            Rs!intfacpr = mvarintfacpr
'            Rs!intliapu = mvarintliapu
'            Rs!intlihap = mvarintlihap
'            Rs!grupoord = mvargrupoord
'            Rs!tinumfac = mvartinumfac
'            Rs!modhcofa = mvarmodhcofa
'            Rs!anofactu = mvaranofactu
'            Rs!perfactu = mvarperfactu
'            Rs!nctafact = mvarnctafact
'            Rs!AsienActAuto = mvarAsienActAuto
'            Rs.Update
'            Agregar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err1:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Rs.Close
'    Agregar = 1
'End Function


''****   ELIMINAR  ***
'Public Function Eliminar() As Byte
'    On Error GoTo Err2
'    SQL = "SELECT * FROM parametros"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " fechaini = " & mvarfechaini
' '   Set Rs = DB.OpenRecordSet(Sql, dbOpenDynaset)
'    If Rs.RecordCount = 0 Then
'        Mens = "No existe un Registro en la Tabla parametros con esos valores."
'        MsgBox Mens, vbExclamation
'        Eliminar = 1
'        'ELSE
'        Else
'            Rs.Delete
'            Rs.MoveNext
'            If Rs.EOF Then Rs.MoveFirst
'            Eliminar = 0
'    End If
'
'    Rs.Close
'    Exit Function
'
'Err2:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Rs.Close
'    Eliminar = 1
'End Function


'****   LEER  ***
Public Function Leer() As Byte
    On Error GoTo Err3
    SQL = "SELECT * FROM parametros"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic
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
            mvargrupogto = DBLet(Rs!grupogto)
            mvargrupovta = DBLet(Rs!grupovta)
            mvarctaperga = DBLet(Rs!ctaperga)
            mvarabononeg = Rs!abononeg
            mvargrupoord = DBLet(Rs!grupoord)
            mvartinumfac = Rs!tinumfac
            mvarmodhcofa = Rs!modhcofa
            mvaranofactu = Rs!anofactu
            mvarperfactu = Rs!perfactu
            mvarnctafact = Rs!nctafact
            mvarAsienActAuto = Rs!AsienActAuto
            mvarCodiNume = Rs!CodiNume
            
            'Modificacion del 28 de Octubre de 2003
            mvarContabilizaFactura = DBLet(Rs!ContabilizaFact, "N")
            
            mvarConstructoras = (Rs!Constructoras = 1)
            
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
    
    
                    'Para que lea si hay bloqueo de cuentas
    If Leer = 0 Then ObtenerCuentasBloqueadas
    
    
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
            
            
            
        If Err.Number <> 0 Then Err.Clear
        If Nuevos <> "" Then MsgBox "Faltan parametros. El programa continuara. Avise soporte tecnico" & vbCrLf & vbCrLf & Nuevos, vbExclamation
    
End Sub

Private Sub AgregarMsg(CADENA As String, Trozo As String)
    
        CADENA = CADENA & vbCrLf & Trozo
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
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconcefpr), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasPro = SQL & "|"
        
    Aux = "tipoconce"
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconceapr), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasPro = mvarAmpliacionFacurasPro & SQL & "|"
    
    'FACTURAS CLIENTES
    Aux = "tipoconce"
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconcefcl), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasCli = SQL & "|"
        
    Aux = "tipoconce"
    SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(mvarconceacl), "N", Aux)
    If SQL = "" And Aux = "tipoconce" Then Exit Function
    mvarAmpliacionFacurasCli = mvarAmpliacionFacurasCli & SQL & "|"
    
    
    
    
    ASignarConceptosFacturas = True
    Exit Function
EASignarConceptosFacturas:
    MuestraError Err.Number, "ASignar Conceptos Facturas"
End Function





'Metodo publico para la obtencion de cuentas bloquedas
Public Sub ObtenerCuentasBloqueadas()

    
    On Error GoTo EmvarCuentasBloqueadas
    
    mvarCuentasBloqueadas = ""
    Set Rs = Nothing
    Set Rs = New ADODB.Recordset
    SQL = "Select codmacta,fecbloq from cuentas where apudirec = 'S' and not (fecbloq is null)"
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        mvarCuentasBloqueadas = mvarCuentasBloqueadas & Rs!codmacta & ":" & Rs!FecBloq & "|"
        Rs.MoveNext
    Wend
    Rs.Close
    
EmvarCuentasBloqueadas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener CuentasBloqueadas"


End Sub



Private Sub LeerDatosSII()

On Error GoTo eLeerDatosSII

    mvarTieneSII = False
    
    

    If Not IsNull(Rs!SIIFechaInicio) Then
        mvarFechaInicioSII = Rs!SIIFechaInicio
        mvarDiasAvisoSII = Rs!SIIDiasAviso
        mvarTieneSII = True
    
    End If
    
    
    
    
    
    


    Exit Sub
eLeerDatosSII:
    mvarTieneSII = False
    Conn.Errors.Clear
    Err.Clear
End Sub


