VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#16.2#0"; "Codejock.SkinFramework.v16.2.0.ocx"
Begin VB.MDIForm frmPpalOld 
   BackColor       =   &H00858585&
   Caption         =   "MDIForm1"
   ClientHeight    =   8955
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   12150
   Icon            =   "frmPpalOld.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmPpalOld.frx":030A
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   6360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Plan contable"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Introduccion de apuntes"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Histórico de apuntes"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consulta extractos"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cuenta de explotacion"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas clientes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas proveedores"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Liquidacion IVA"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Balance de sumas y saldos mensual"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Balance de situación"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cuenta de Pérdidas y ganancias"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambio Empresa"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar impresora seleccionada"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Integraciones pendientes"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Integraciones erroneas"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   8370
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpalOld.frx":EEC8
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13361
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "12:53"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   6330
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   6300
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   7140
      Top             =   1830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   7110
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   7110
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   4500
      Top             =   3240
      _Version        =   1048578
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnDatos 
      Caption         =   "D&atos generales"
      Begin VB.Menu mnPlanContable 
         Caption         =   "&Plan contable"
      End
      Begin VB.Menu mnConceptos 
         Caption         =   "&Conceptos"
      End
      Begin VB.Menu mnTiposDiario 
         Caption         =   "Tipos &diarios"
      End
      Begin VB.Menu mnContadores 
         Caption         =   "C&ontadores"
      End
      Begin VB.Menu mnTiposIVA 
         Caption         =   "Tipos de &I.V.A."
      End
      Begin VB.Menu mnbarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnConfiguracionAplicacion 
         Caption         =   "Confi&guracion"
         Begin VB.Menu mnEmpresa 
            Caption         =   "&Datos empresa"
         End
         Begin VB.Menu mnParametros 
            Caption         =   "&Parametros"
         End
         Begin VB.Menu mnInformesScrystal 
            Caption         =   "Informes"
         End
         Begin VB.Menu mnUsuarios 
            Caption         =   "Mantenimiento &Usuarios"
         End
         Begin VB.Menu mnNuevaEmpresa 
            Caption         =   "Creacion nueva &empresa"
         End
         Begin VB.Menu mnbarra15 
            Caption         =   "-"
         End
         Begin VB.Menu mnCambioPassword 
            Caption         =   "Cambiar password"
         End
      End
      Begin VB.Menu mnBarra5 
         Caption         =   "-"
      End
      Begin VB.Menu mnCambioUsuario 
         Caption         =   "Cambiar  empresa"
      End
      Begin VB.Menu mnbarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSeleccionarImpresora 
         Caption         =   "Seleccionar impresora"
      End
      Begin VB.Menu mnbarra10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSal 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnDiario 
      Caption         =   "&Diario"
      Begin VB.Menu mnIntroducirAsientos 
         Caption         =   "Introducir &asientos"
      End
      Begin VB.Menu mnActalizacionAsientos 
         Caption         =   "A&ctualización de asientos"
      End
      Begin VB.Menu mnConsultaExtractos 
         Caption         =   "&Consulta de extractos"
      End
      Begin VB.Menu mnPunteoExtractos 
         Caption         =   "Punteo &de extractos"
      End
      Begin VB.Menu mnListadoMayor 
         Caption         =   "Listado de extracto de cuentas"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnAsientosPredefinidos 
         Caption         =   "Asientos &predefinidos"
      End
   End
   Begin VB.Menu mnMenuIVA 
      Caption         =   "I&va"
      Begin VB.Menu mnClientes 
         Caption         =   "Clientes"
         Begin VB.Menu mnRegFacCli 
            Caption         =   "Registro facturas"
         End
         Begin VB.Menu mnContFactCli 
            Caption         =   "Contabilizar facturas"
         End
         Begin VB.Menu mnListFactCli 
            Caption         =   "Listado facturas"
         End
         Begin VB.Menu mnRelacionClientesVentas 
            Caption         =   "Relacion clientes por cuenta ventas"
         End
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnMenuProveedores 
         Caption         =   "Proveedores"
         Begin VB.Menu mnRegFac 
            Caption         =   "Registro Facturas"
         End
         Begin VB.Menu mnContFactProv 
            Caption         =   "Contabilizar"
         End
         Begin VB.Menu mnListFacProv 
            Caption         =   "Listado facturas"
         End
         Begin VB.Menu mnRelacionProveGastos 
            Caption         =   "Relación proveedores por Cta. gastos"
         End
      End
      Begin VB.Menu mnbarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnLiquidacion 
         Caption         =   "Liquidacion I.V.A."
      End
      Begin VB.Menu mnListINTRASTAD 
         Caption         =   "Certificado declaración IVA"
      End
      Begin VB.Menu mnModelo340 
         Caption         =   "Modelo 340"
      End
      Begin VB.Menu mnModelo347 
         Caption         =   "Modelo 347"
      End
      Begin VB.Menu mnModelo349 
         Caption         =   "Modelo 349"
      End
      Begin VB.Menu mnbarra101 
         Caption         =   "-"
      End
      Begin VB.Menu mnDatosExternos347 
         Caption         =   "Datos externos 347"
         Begin VB.Menu mnDatosExternos347Manten 
            Caption         =   "Mantenimiento"
         End
         Begin VB.Menu mnDatosExternos347Importar 
            Caption         =   "Importar fichero datos ext. 347"
         End
      End
   End
   Begin VB.Menu mnHcoApuntes 
      Caption         =   "&Histórico"
      Begin VB.Menu mnVerHistoricoApuntes 
         Caption         =   "Histórico de apuntes"
      End
      Begin VB.Menu mnReemision 
         Caption         =   "Reemisión de diario"
      End
      Begin VB.Menu mnTotalesCtaConcepto 
         Caption         =   "Totales por cuenta y concepto"
      End
      Begin VB.Menu mnBalanceMensual 
         Caption         =   "Balance de sumas y saldos"
      End
      Begin VB.Menu mnEvolmensualSaldos 
         Caption         =   "Evolución mensual de saldos"
      End
      Begin VB.Menu mnBalancesituacion 
         Caption         =   "Balance de situacion"
      End
      Begin VB.Menu mnPerdyGan 
         Caption         =   "Cuenta pérdidas y ganancias"
      End
      Begin VB.Menu mnCtaExplotacion 
         Caption         =   "Cuenta de explotación"
      End
      Begin VB.Menu mnCtaExplotacionComparativaActual 
         Caption         =   "Cuenta de explotación comparativa"
      End
      Begin VB.Menu mnRatios 
         Caption         =   "Ratios y gráficas"
      End
      Begin VB.Menu mnConsolidado 
         Caption         =   "Consolidado"
         Begin VB.Menu mnBalanceConsolidado 
            Caption         =   "Balance de sumas y saldos"
         End
         Begin VB.Menu mnCtaexplotaConsol 
            Caption         =   "Cuenta de explotación"
         End
         Begin VB.Menu mnBalSituConsolidado 
            Caption         =   "Balance situación"
         End
         Begin VB.Menu mnPyGConso 
            Caption         =   "Pérdidas y ganancias"
         End
         Begin VB.Menu mnConsFacturasProv 
            Caption         =   "Listado Facturas Proveedores"
         End
         Begin VB.Menu mnConsoFacturasCli 
            Caption         =   "Listado Facturas  Cliente"
         End
      End
   End
   Begin VB.Menu mnCierre 
      Caption         =   "&Cierre"
      Begin VB.Menu mnRenumeracion 
         Caption         =   "Renumeración de asientos"
      End
      Begin VB.Menu mnSimulaCierre 
         Caption         =   "Simulación del cierre"
      End
      Begin VB.Menu mnAsiePerdyGana 
         Caption         =   "Cierre de ejercicio"
      End
      Begin VB.Menu mnDescierre 
         Caption         =   "Deshacer ejercicio cerrado"
      End
      Begin VB.Menu mnDiarioOficial 
         Caption         =   "Diario oficial"
      End
      Begin VB.Menu mnDiarioResumen 
         Caption         =   "Diario resumen"
      End
      Begin VB.Menu mnPresetaCuentas 
         Caption         =   "Presentación cuentas anuales"
      End
      Begin VB.Menu mnPresentaTelematica 
         Caption         =   "Presentacion telemática de libros"
      End
      Begin VB.Menu mnbarra1_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnTraspasoACerrados 
         Caption         =   "Traspaso a ejercicios cerrados"
      End
      Begin VB.Menu mnTraerDeCerrados 
         Caption         =   "Traer de ejercicios cerrados"
      End
      Begin VB.Menu mnBorrarRegClientes 
         Caption         =   "Borre registro clientes"
      End
      Begin VB.Menu mnBorrarProveedores 
         Caption         =   "Borre registro proveedores"
      End
      Begin VB.Menu mnGenerarMemoria 
         Caption         =   "Generar memoria"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnEjereciciosCerrados 
      Caption         =   "&Ejer. cerrados"
      Begin VB.Menu mnTotalMensyAcumuladosCerrados 
         Caption         =   "Consulta totales mensuales y saldos acumulados"
         Visible         =   0   'False
      End
      Begin VB.Menu mnTotalCtaConceptoCerrado 
         Caption         =   "Total cuenta y concepto"
      End
      Begin VB.Menu mnConsultaExtractoCtaCerrado 
         Caption         =   "Consulta extractos cuentas"
      End
      Begin VB.Menu mnPunteoCerrados 
         Caption         =   "Punteo extractos cuentas"
      End
      Begin VB.Menu mnListadoExtracotCuentas 
         Caption         =   "Listado extractos de cuentas"
      End
      Begin VB.Menu mnConsultaHcoApuntes 
         Caption         =   "Consulta histórico apuntes"
      End
      Begin VB.Menu mnReemisionDiarios 
         Caption         =   "Reemisión de diarios"
      End
      Begin VB.Menu mnBalanceSumasSaldosMensual 
         Caption         =   "Balance sumas y saldos"
      End
      Begin VB.Menu mnListadoCuentaExplotacion 
         Caption         =   "Listado cuenta explotación"
      End
      Begin VB.Menu mnCtaExplotaComparativa 
         Caption         =   "Cuenta de explotación comparativa"
      End
      Begin VB.Menu mnListadoDiarioOficial 
         Caption         =   "Listado del diario oficial"
      End
      Begin VB.Menu mnListadoDiarioResumen 
         Caption         =   "Listado del diario resumen"
      End
      Begin VB.Menu mnBorreEjerciciosCerrados 
         Caption         =   "Borre de ejercicios cerrados"
      End
   End
   Begin VB.Menu mnAnalitica 
      Caption         =   "A&nálitica"
      Begin VB.Menu mnCentosDeCoste 
         Caption         =   "Centros de coste"
      End
      Begin VB.Menu mnConsultaSaldosAnal 
         Caption         =   "Consulta de saldos"
      End
      Begin VB.Menu mnCtaExplotacionAnal 
         Caption         =   "Cuenta de explotación"
      End
      Begin VB.Menu mnCentrosCostexCuenta 
         Caption         =   "Centros de coste por cuenta"
      End
      Begin VB.Menu mnDetalleExplotacion 
         Caption         =   "Detalle de explotación"
      End
   End
   Begin VB.Menu mnPresupuestaria 
      Caption         =   "&Presupuestaria"
      Begin VB.Menu mnPresupuestos 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu mnListadoPresupuestos 
         Caption         =   "Listado presupuestos"
      End
      Begin VB.Menu mnBalPresupuestos 
         Caption         =   "Balance presupuestario"
      End
   End
   Begin VB.Menu mnInmovilizdo 
      Caption         =   "&Inmovilizado"
      Begin VB.Menu mnParametrosInmo 
         Caption         =   "Parámetros"
      End
      Begin VB.Menu mnConceptosInmo 
         Caption         =   "Conceptos"
      End
      Begin VB.Menu mnElementosInmo 
         Caption         =   "Elementos"
      End
      Begin VB.Menu mnFichaElto 
         Caption         =   "Ficha de elementos"
      End
      Begin VB.Menu mnEstadisticasInmo 
         Caption         =   "Estadísticas"
      End
      Begin VB.Menu mnEstadFechsInmo 
         Caption         =   "Estadísticas entre fechas"
      End
      Begin VB.Menu mnHcoInmo 
         Caption         =   "Histórico inmovilizado"
      End
      Begin VB.Menu mnSimulAmortiza 
         Caption         =   "Simulación próxima amortización"
      End
      Begin VB.Menu mnCaluloYContabilizacion 
         Caption         =   "Cálculo y contabilización"
      End
      Begin VB.Menu mnVentaBajaInmo 
         Caption         =   "Venta/Baja Inmovilizado"
      End
      Begin VB.Menu mnDeshacerAmortizacion 
         Caption         =   "Deshacer última amortización"
      End
   End
   Begin VB.Menu mnEnlacebancos 
      Caption         =   "Enlace &bancos"
      Begin VB.Menu mnConfigurarCtas 
         Caption         =   "Configurar cuentas"
      End
      Begin VB.Menu mnImportarNorma43 
         Caption         =   "Importar fichero Norma 43"
      End
      Begin VB.Menu mnPunteoBancario 
         Caption         =   "Punteo extractos"
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnCalcularSaldos 
         Caption         =   "Comprobación Cuadre"
      End
      Begin VB.Menu mnRecalculoSaldos 
         Caption         =   "Recálculo de saldos"
      End
      Begin VB.Menu mnImportarDatosFiscales 
         Caption         =   "Importar datos fiscales"
      End
      Begin VB.Menu mnRevisarMultibase 
         Caption         =   "Revisar caracteres especiales"
      End
      Begin VB.Menu mnMemoriaPagos 
         Caption         =   "Memoria de pagos a proveedores"
      End
      Begin VB.Menu mnPGC2008 
         Caption         =   "TRASPASO P.G.C. 2008"
         Begin VB.Menu mnTraspasoPGC2008 
            Caption         =   "Digitos 3 y 4"
         End
         Begin VB.Menu mnTraspasoPGC2008UltNivel 
            Caption         =   "Ultimo nivel"
         End
      End
      Begin VB.Menu mnAgrCta 
         Caption         =   "Agrupacion cuentas"
         Begin VB.Menu mnAgruparCtasBalance 
            Caption         =   "Agrupar en balance"
         End
         Begin VB.Menu mnExclusionConsol 
            Caption         =   "Exclusion en consolidado"
         End
      End
      Begin VB.Menu mnBusquedas 
         Caption         =   "Buscar ..."
         Begin VB.Menu mnBuscarAsientosDescuadrados 
            Caption         =   "Asientos descuadrados"
         End
         Begin VB.Menu mnCabecerasErroneas 
            Caption         =   "Cabeceras asiento incorretas"
         End
         Begin VB.Menu mnCtasSinMovs 
            Caption         =   "Cuentas sin movimientos"
         End
         Begin VB.Menu mnBuscarHuecosCtas 
            Caption         =   "Buscar nº cuentas libres"
         End
         Begin VB.Menu mnErroresFactura 
            Caption         =   "Errores nº factura"
            Begin VB.Menu mnBusFacCli 
               Caption         =   "Cliente"
            End
            Begin VB.Menu mnBusFacProv 
               Caption         =   "Proveedores"
            End
         End
      End
      Begin VB.Menu mnMenConfigurar 
         Caption         =   "Configurar"
         Begin VB.Menu mnConfigBalPeryGan 
            Caption         =   "Balances"
         End
         Begin VB.Menu mnConfigMemoria 
            Caption         =   "Memoria"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnTraspasos 
         Caption         =   "Traspasos"
         Begin VB.Menu mnTrasapasoAce 
            Caption         =   "Ace"
         End
         Begin VB.Menu mnTraspasoPersa 
            Caption         =   "Persa"
         End
         Begin VB.Menu mnbarra31 
            Caption         =   "-"
         End
         Begin VB.Menu mnTraspasoFacturas 
            Caption         =   "Facturas"
            Begin VB.Menu mnTraspasoExportar 
               Caption         =   "Exportar"
            End
            Begin VB.Menu mnTraspasoImportar 
               Caption         =   "Importar"
            End
         End
         Begin VB.Menu mnTraspasoEntreSecciones 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnTraspasoEntreSecciones 
            Caption         =   "Traspaso cuentas banco"
            Index           =   1
         End
      End
      Begin VB.Menu mnbarra12 
         Caption         =   "-"
      End
      Begin VB.Menu mnHerramientasAriadnaCC 
         Caption         =   "Herramientas CC"
         Begin VB.Menu mnHerrAriadnaCC 
            Caption         =   "Desbloquear asientos"
            Index           =   0
         End
         Begin VB.Menu mnHerrAriadnaCC 
            Caption         =   "Mover cuentas"
            HelpContextID   =   1
            Index           =   1
         End
         Begin VB.Menu mnHerrAriadnaCC 
            Caption         =   "Mod.  nº registro frapro"
            HelpContextID   =   5
            Index           =   2
         End
         Begin VB.Menu mnHerrAriadnaCC 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnHerrAriadnaCC 
            Caption         =   "Aumentar digitos ultimo nivel"
            HelpContextID   =   3
            Index           =   4
         End
         Begin VB.Menu mnHerrAriadnaCC 
            Caption         =   "Cambio IVA"
            HelpContextID   =   4
            Index           =   5
         End
      End
      Begin VB.Menu mnbarra11 
         Caption         =   "-"
      End
      Begin VB.Menu mnBackUp 
         Caption         =   "Copia seguridad local"
      End
      Begin VB.Menu mnVerLog 
         Caption         =   "Ver LOG"
      End
      Begin VB.Menu mnUsuariosActivos 
         Caption         =   "Usuarios activos"
      End
   End
   Begin VB.Menu mnPuntoFinal 
      Caption         =   "&Soporte"
      Begin VB.Menu mnAyuda 
         Caption         =   "Ayuda"
      End
      Begin VB.Menu mnbarra7_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEnviarMail 
         Caption         =   "Enviar Mail"
      End
      Begin VB.Menu mnWeb 
         Caption         =   "Web Ariadna Software"
      End
      Begin VB.Menu mnCheckVersion 
         Caption         =   "Comprobar version operativa"
      End
      Begin VB.Menu mnBarra7_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnAcercaDE 
         Caption         =   "Acerca de ..."
      End
   End
End
Attribute VB_Name = "frmPpalOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'

Dim PrimeraVez As Boolean
 
Dim TieneEditorDeMenus As Boolean

Private Sub MDIForm_Activate()

    If PrimeraVez Then
           
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
        
        'CadenaDesdeOtroForm = Format(Now, "Short date")
        
        
        espera 0.1
        EliminarAlgunosDatos False
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub MDIForm_Load()
    PrimeraVez = True
    'Dim T1 As Single
    
    'Cargamos los iconos desde la DLL
    'T1 = Timer
    ImageList1.ImageHeight = 32
    ImageList1.ImageWidth = 32
    GetIconsFromLibrary App.path & "\icoconppal.dll", 1, 32

    imgListComun.ImageHeight = 24
    imgListComun.ImageWidth = 24
    GetIconsFromLibrary App.path & "\icolistcon.dll", 2, 24
    
    '++
    imgListComun_BN.ImageHeight = 24
    imgListComun_BN.ImageWidth = 24
    GetIconsFromLibrary App.path & "\icolistcon_BN.dll", 3, 24
    imgListComun_OM.ImageHeight = 24
    imgListComun_OM.ImageWidth = 24
    GetIconsFromLibrary App.path & "\icolistcon_OM.dll", 4, 24
    
    GetIconsFromLibrary App.path & "\icolistcon.dll", 5, 16
    GetIconsFromLibrary App.path & "\icolistcon_BN.dll", 6, 16
    GetIconsFromLibrary App.path & "\icolistcon_OM.dll", 7, 16
    '++
    
    'MsgBox Timer - T1
    'Botones
    With Me.Toolbar1
        .ImageList = Me.ImageList1
        .Buttons(1).Image = 1   'Plan contable
        '---
        .Buttons(3).Image = 2   'Diario
        .Buttons(4).Image = 3   'Hco
        .Buttons(5).Image = 4   'Con extractos
        .Buttons(6).Image = 19   'CTA EXPLOTACION
        '----
        .Buttons(8).Image = 16   'Fac CLI
        .Buttons(9).Image = 18   'Fac PRO
        .Buttons(10).Image = 13   'Liquidacion IVA
        '----
        .Buttons(12).Image = 7  'Balance
        .Buttons(13).Image = 14   'Balance situacion
        .Buttons(14).Image = 17   'Cuenta P y G
        '----
        .Buttons(16).Image = 8  'Usuarios
        .Buttons(17).Image = 9  'Impresora
        '----
        '----
        .Buttons(22).Image = 10 'Salir
     
    End With
    LeerEditorMenus
    PonerDatosFormulario
       
   
   EstablecerSkin CInt(1)
       
       
End Sub


Public Sub EstablecerSkin(QueSkin As Integer)


  FijaSkin QueSkin

  ' Cargando el archivo del Skin
  ' ============================
    frmPpal.SkinFramework.LoadSkin Skn$, ""
    frmPpal.SkinFramework.ApplyWindow frmPpal.hWnd
    frmPpal.SkinFramework.ApplyOptions = frmPpal.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
'
    
End Sub

Private Function FijaSkin(Cual As Integer)
  Select Case (Cual)
    Case 0:     ' Windows Luna XP Modificado
      Skn$ = CStr(App.path & "\Styles\WinXP.Luna.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
    Case 1:     ' Windows Royale Modificado
      Skn$ = CStr(App.path & "\Styles\WinXP.Royale.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalRoyale.ini"
    Case 2:     ' Microsoft Office 2007
      Skn$ = CStr(App.path & "\Styles\Office2007.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
    Case 3:     ' Windows Vista Sencillo
      Skn$ = CStr(App.path & "\Styles\Vista.cjstyles")
      frmPpal.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
  End Select

End Function



Private Sub PonerDatosFormulario()
Dim Config As Boolean

    Config = (vParam Is Nothing) Or (vEmpresa Is Nothing)
    
    If Not Config Then HabilitarSoloPrametros_o_Empresas True
    
    'FijarConerrores
    CadenaDesdeOtroForm = ""
    
    'Poner datos visible del form
    PonerDatosVisiblesForm
    'Poner opciones de nivel de usuario
    PonerOpcionesUsuario
    
    
    If Not Config Then
        Me.mnTraspasoEntreSecciones(0).Visible = vParam.TraspasCtasBanco > 0
        Me.mnTraspasoEntreSecciones(1).Visible = mnTraspasoEntreSecciones(0).Visible
    End If
    'Habilitar
    If Config Then HabilitarSoloPrametros_o_Empresas False
    'Panel con el nombre de la empresa
    If Not vEmpresa Is Nothing Then
        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
    Else
        Me.StatusBar1.Panels(2).Text = "Falta configurar"
    End If
    
    'Primero los pongo a visible
    mnDatosExternos347.Visible = True
    mnbarra101.Visible = True
    
    
    
    
    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    
     mnCheckVersion.Visible = False 'Siempre oculto
    
        
    If Not Config Then
        mnDatosExternos347.Visible = mnDatosExternos347.Visible And vParam.AgenciaViajes
        mnbarra101.Visible = mnbarra101.Visible And vParam.AgenciaViajes
    End If
    '---------------------------------------------------
    'Las asociaciones entre menu y botones  del TOOLBAR
    With Me.Toolbar1
        .Buttons(1).Visible = mnDatos.Visible And Me.mnPlanContable.Visible
        '---
        .Buttons(3).Visible = mnDiario.Visible And Me.mnIntroducirAsientos.Visible    'Diario
        .Buttons(4).Visible = mnHcoApuntes.Visible And mnVerHistoricoApuntes.Visible    'Hco
        .Buttons(5).Visible = mnDiario.Visible And mnConsultaExtractos.Visible   'Con extractos
        .Buttons(6).Visible = mnHcoApuntes.Visible And mnCtaExplotacion.Visible   'CTA EXPLOTACION
        '----
        .Buttons(8).Visible = mnMenuIVA.Visible And mnClientes.Visible And Me.mnRegFacCli.Visible     'Fac CLI
        .Buttons(9).Visible = mnMenuIVA.Visible And mnMenuProveedores.Visible And Me.mnRegFac.Visible    'Fac PRO
        .Buttons(10).Visible = mnMenuIVA.Visible And Me.mnLiquidacion.Visible   'Liquidacion IVA
        '----
        .Buttons(12).Visible = mnHcoApuntes.Visible And mnBalanceMensual.Visible  'Balance
        .Buttons(13).Visible = mnHcoApuntes.Visible And mnBalancesituacion.Visible
        .Buttons(14).Visible = mnHcoApuntes.Visible And Me.mnPerdyGan.Visible  'Cuenta P y G
        '----
        .Buttons(16).Image = 8  'Usuarios
        .Buttons(17).Image = 9  'Impresora
        '----
        .Buttons(19).Visible = TieneIntegracionesPendientes
        .Buttons(19).Image = 11
        'Antes
        .Buttons(20).Visible = False
        '.Buttons(20).Visible = BuscarIntegraciones(True)
        .Buttons(20).Image = 12
        '----
        .Buttons(22).Image = 10 'Salir
    End With
        
    'Si el usuario tiene permiso para ver los balances, le dejo las graficas
    Me.mnRatios.Visible = Toolbar1.Buttons(12).Visible
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim Cad As String


    'Alguna cosilla antes de cerrar. Eliminar bloqueos
    Cad = "Delete from zBloqueos where codusu = " & vUsu.Codigo
    Conn.Execute Cad

    'Elimnar bloquo BD
    Cad = DevuelveDesdeBD("codusu", "Usuarios.vBloqBD", "codusu", vUsu.Codigo)
    If Cad <> "" Then Conn.Execute "Delete from Usuarios.vBloqBD where codusu=" & vUsu.Codigo
        
    Conn.Close
    Set Conn = Nothing
    
End Sub

Private Sub mnAcercaDE_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.opcion = 6
    frmMensajes.Show vbModal
End Sub

Private Sub mnActalizacionAsientos_Click()
    Screen.MousePointer = vbHourglass
    frmActualizar.OpcionActualizar = 5
    AlgunAsientoActualizado = False
    frmActualizar.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnAgruparCtasBalance_Click()
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 2
    frmUtilidades.Show vbModal
End Sub


Private Sub mnASientosPredefinidos_Click()
    Screen.MousePointer = vbHourglass
    frmAsiPre.Show vbModal
End Sub


Private Sub mnAsiePerdyGana_Click()
    frmCierre.opcion = 1
    frmCierre.Show vbModal
End Sub

Private Sub mnBackUp_Click()
    frmBackUP.Show vbModal
End Sub

Private Sub mnBalanceConsolidado_Click()
    AbrirListado 24, False
End Sub

Private Sub mnBalanceMensual_Click()
    AbrirListado 5, False
End Sub

Private Sub mnBalancesituacion_Click()
    AbrirListado 26, False
End Sub

Private Sub mnBalanceSumasSaldosMensual_Click()
 AbrirListado 5, True
End Sub

Private Sub mnBalPresupuestos_Click()
    AbrirListado 10, False
End Sub


Private Sub mnBalSituConsolidado_Click()

    AbrirListado 51, False
End Sub

Private Sub mnBorrarProveedores_Click()
    AbrirListado 23, False
End Sub

Private Sub mnBorrarRegClientes_Click()
    AbrirListado 22, False
End Sub

Private Sub mnBorreEjerciciosCerrados_Click()
    frmCierre.opcion = 6
    frmCierre.Show vbModal
End Sub

Private Sub mnBuscarAsientosDescuadrados_Click()
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 1
    frmUtilidades.Show vbModal
End Sub

Private Sub mnBuscarHuecosCtas_Click()
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 4
    frmUtilidades.Show vbModal
End Sub

Private Sub mnBusFacCli_Click()
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 5
    frmUtilidades.Show vbModal
End Sub

Private Sub mnBusFacProv_Click()
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 6
    frmUtilidades.Show vbModal
End Sub

Private Sub mnCabecerasErroneas_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.opcion = 12
    frmMensajes.Show vbModal
End Sub

Private Sub mnCalcularSaldos_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.opcion = 2
    frmMensajes.Show vbModal
End Sub

Private Sub mnCaluloYContabilizacion_Click()
    frmInmov.opcion = 2
    frmInmov.Show vbModal
End Sub

Private Sub mnCambioPassword_Click()
    frmMensajes.opcion = 15
    frmMensajes.Show vbModal
End Sub

Private Sub mnCambioUsuario_Click()

    
    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If
    
    'Borramos temporal
    EjecutaSQL "Delete from zBloqueos where codusu = " & vUsu.Codigo

    
    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"

    frmLogin.Show vbModal

    
    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    Conn.Close

    
    If AbrirConexion("") = False Then
        MsgBox "La apliacación no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If
    
    'Borramos Conextr
    EliminarAlgunosDatos True

    
    Set vParam = Nothing
    Set vEmpresa = Nothing
    LeerEmpresaParametros
    PonerDatosFormulario
    
    'Ponemos primera vez a false
    PrimeraVez = True
    Me.SetFocus
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub mnCentosDeCoste_Click()
    frmCCoste.DatosADevolverBusqueda = ""
    frmCCoste.Show vbModal
End Sub

Private Sub mnCentrosCostexCuenta_Click()
    AbrirListado 17, False
End Sub

Private Sub mnCheckVersion_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "webversion"
    espera 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnConceptos_Click()
    Screen.MousePointer = vbHourglass
    frmConceptos.Show vbModal
End Sub



Private Sub mnConceptosInmo_Click()
    Screen.MousePointer = vbHourglass
    frmConceptosInmo.Show vbModal
End Sub

Private Sub mnConfigBalPeryGan_Click()
    Screen.MousePointer = vbHourglass
    frmColBalan.Show vbModal
End Sub

Private Sub mnConfigBalSituacion_Click()
    Screen.MousePointer = vbHourglass
    frmBalances.Show vbModal
End Sub

Private Sub mnConfigMemoria_Click()
    Screen.MousePointer = vbHourglass
    frmMemoria.opcion = 0
    frmMemoria.Show vbModal
End Sub

Private Sub mnConfigurarCtas_Click()
    Screen.MousePointer = vbHourglass
    frmCuentasBancarias.Show vbModal
End Sub

Private Sub mnConsFacturasProv_Click()
    AbrirListado 52, False
End Sub

Private Sub mnConsoFacturasCli_Click()
    AbrirListado 53, False
End Sub

Private Sub mnConsultaExtractoCtaCerrado_Click()
    Screen.MousePointer = vbHourglass
    frmConExtr.EjerciciosCerrados = True
    frmConExtr.Cuenta = ""
    frmConExtr.Show vbModal
End Sub

Private Sub mnConsultaExtractos_Click()
    Screen.MousePointer = vbHourglass
    frmConExtr.EjerciciosCerrados = False
    frmConExtr.Cuenta = ""
    frmConExtr.Show vbModal
End Sub

Private Sub mnConsultaHcoApuntes_Click()
    Screen.MousePointer = vbHourglass
    frmHcoApuntes.EjerciciosCerrados = True
    frmHcoApuntes.ASIENTO = ""
    frmHcoApuntes.LINASI = 0
    frmHcoApuntes.Show vbModal
End Sub

Private Sub mnConsultaSaldosAnal_Click()
    AbrirListado 15, False
End Sub

Private Sub mnContadores_Click()
    Screen.MousePointer = vbHourglass
    frmContadores.Show vbModal
End Sub

Private Sub mnContFactCli_Click()
    Screen.MousePointer = vbHourglass
    frmActualizar.OpcionActualizar = 10
    AlgunAsientoActualizado = False
    frmActualizar.Show vbModal
End Sub

Private Sub mnContFactProv_Click()
    Screen.MousePointer = vbHourglass
    frmActualizar.OpcionActualizar = 11
    AlgunAsientoActualizado = False
    frmActualizar.Show vbModal
End Sub

Private Sub mnCtaExplotacion_Click()
    AbrirListado 7, False
End Sub


Private Sub mnCtaExplotacionAnal_Click()
    AbrirListado 16, False
End Sub

Private Sub mnCtaExplotacionComparativaActual_Click()
    AbrirListado 21, False
End Sub

Private Sub mnCtaExplotaComparativa_Click()
 'MsgBox "NO TA ENTOAVIA", vbMsgBoxRight + vbCritical
    AbrirListado 21, True
End Sub

Private Sub mnCtaexplotaConsol_Click()
    AbrirListado 31, False
End Sub

Private Sub mnCtasSinMovs_Click()
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 0
    frmUtilidades.Show vbModal
End Sub

Private Sub mnDatosExternos347Importar_Click()
    frmMensajes.opcion = 20
    frmMensajes.Show vbModal
End Sub

Private Sub mnDatosExternos347Manten_Click()
    frmDatosExt347.Show vbModal
End Sub

Private Sub mnDescierre_Click()
    frmCierre.opcion = 5
    frmCierre.Show vbModal
End Sub

Private Sub mnDeshacerAmortizacion_Click()
    frmInmov.opcion = 10
    frmInmov.Show vbModal
End Sub

Private Sub mnDetalleExplotacion_Click()
    AbrirListado 19, False
End Sub

Private Sub mnDiarioOficial_Click()
  AbrirListado 14, False
End Sub

Private Sub mnDiarioResumen_Click()
    AbrirListado 18, False
End Sub

Private Sub mnElementosInmo_Click()
    frmEltoInmo.DatosADevolverBusqueda = ""
    frmEltoInmo.Show vbModal
End Sub

Private Sub mnEmpresa_Click()
    frmempresa.Show
End Sub


Private Sub mnEnviarMail_Click()
    frmEMail.opcion = 1
    frmEMail.Show vbModal
End Sub

Private Sub mnEstadFechsInmo_Click()
    frmInmov.opcion = 6
    frmInmov.Show vbModal
End Sub

Private Sub mnEstadisticasInmo_Click()
    frmInmov.opcion = 4
    frmInmov.Show vbModal
End Sub

Private Sub mnEvolmensualSaldos_Click()
    AbrirListado 54, False
End Sub

Private Sub mnExclusionConsol_Click()
    Screen.MousePointer = vbHourglass
    frmUtilidades.opcion = 3
    frmUtilidades.Show vbModal
End Sub

Private Sub mnFichaElto_Click()
    frmInmov.opcion = 5
    frmInmov.Show vbModal
End Sub

Private Sub mnGenerarMemoria_Click()
    Screen.MousePointer = vbHourglass
    frmMemoria.opcion = 1
    frmMemoria.Show vbModal
End Sub

Private Sub mnHcoInmo_Click()
    Screen.MousePointer = vbHourglass
    frmHcoInmo.Show vbModal
End Sub

Private Sub mnHerrAriadnaCC_Click(Index As Integer)
 
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene permisos", vbExclamation
            Exit Sub
        End If
        'El index 3 , que es la barra, en frmCC es la opcion de NUEVA EMPRESA
        ' y no se llma desde aqui, con lo cual no hay problemo
        'Para el restro cojo el valor del helpidi
        
        frmCentroControl.opcion = mnHerrAriadnaCC(Index).HelpContextID
        frmCentroControl.Show vbModal
    
End Sub

Private Sub mnImportarDatosFiscales_Click()
    frmMensajes.opcion = 13
    frmMensajes.Show vbModal
End Sub

Private Sub mnImportarNorma43_Click()
    frmUtiliBanco.Show vbModal
End Sub

Private Sub mnInformesScrystal_Click()
    frmCrystal.Show vbModal
End Sub

Private Sub mnIntroducirAsientos_Click()
    Screen.MousePointer = vbHourglass
    frmAsientos.ASIENTO = ""
    frmAsientos.Show vbModal
End Sub

Private Sub mnLiquidacion_Click()
    AbrirListado 12, False
End Sub

Private Sub mnListadoCuentaExplotacion_Click()
 AbrirListado 7, True
End Sub

Private Sub mnListadoDiarioOficial_Click()
    AbrirListado 14, True
End Sub

Private Sub mnListadoDiarioResumen_Click()
    AbrirListado 18, True
End Sub

Private Sub mnListadoExtracotCuentas_Click()
    AbrirListado 1, True
End Sub

Private Sub mnListadoMayor_Click()
    AbrirListado 1, False
End Sub

Private Sub mnListadoPresupuestos_Click()
    AbrirListado 9, False
End Sub

Private Sub mnListFacProv_Click()
    AbrirListado 13, False
End Sub

Private Sub mnListFactCli_Click()
    AbrirListado 8, False
End Sub

Private Sub mnListINTRASTAD_Click()
    AbrirListado 11, False
End Sub

Private Sub mnMemoriaPagos_Click()
    frmListado2.opcion = 3
    frmListado2.Show vbModal
End Sub

Private Sub mnModelo340_Click()
    frmListado2.opcion = 0
    frmListado2.Show vbModal
End Sub

Private Sub mnModelo347_Click()
    AbrirListado 20, False
End Sub

Private Sub mnModelo349_Click()
    AbrirListado 28, False
End Sub

Private Sub mnNuevaEmpresa_Click()
    'NUEVO
    If vUsu.Nivel > 1 Then Exit Sub
    
    frmCentroControl.opcion = 2
    frmCentroControl.Show vbModal
    'ANTES
    'CadenaDesdeOtroForm = App.path & "\AriadnaCC.exe"
    'If Dir(CadenaDesdeOtroForm) = "" Then
    '    MsgBox "No esta bien referenciado el CENTRO DE CONTROL", vbExclamation
    'Else
    '    Shell CadenaDesdeOtroForm & " /E:" & vEmpresa.codempre, vbNormalFocus
    'End If
End Sub

Private Sub mnparametros_Click()
    If Not (vEmpresa Is Nothing) Then
        frmparametros.Show
    End If
End Sub

Private Sub mnParametrosInmo_Click()
    frmInmov.opcion = 0
    frmInmov.Show vbModal
End Sub

'Private Sub mnPedirPwd_Click()
'Dim Anterior As Boolean
'
'    Anterior = Me.mnPedirPwd.Checked
'    vConfig.PedirPasswd = Not Anterior
'    If vConfig.Grabar = 1 Then
'        Me.mnPedirPwd.Checked = Anterior
'    Else
'        Me.mnPedirPwd.Checked = Not Anterior
'    End If
'End Sub

Private Sub mnPerdyGan_Click()
    AbrirListado 27, False
End Sub

Private Sub mnPlanContable_Click()
'    Dim i As Integer
'    Dim SQL As String
'      i = Year(DateAdd("yyyy", 1, vParam.fechafin))
'        If i > 0 Then
'            If i > 2010 Then
'                SQL = "3"
'            Else
'                SQL = "2"
'            End If
'            SQL = SQL & Mid(CStr(i), 4, 1)
'            SQL = SQL & "00000"
'
'            SQL = "UPDATE contadores SET contado2 = " & SQL & " WHERE tiporegi='1'"
'            If Not EjecutaSQL(SQL) Then MsgBox "Error updateando contador proveeedores: " & vbCrLf & SQL, vbExclamation
'        End If
'
'    Exit Sub
    Screen.MousePointer = vbHourglass
    frmColCtas.ConfigurarBalances = 0
    frmColCtas.DatosADevolverBusqueda = ""
    frmColCtas.Show vbModal
End Sub

Private Sub mnPresentaTelematica_Click()
    Telematica 1
End Sub

Private Sub mnPresetaCuentas_Click()
    Telematica 0
End Sub

Private Sub Telematica(Caso As Integer)
        Me.Enabled = False
        frmTelematica.opcion = Caso
        frmTelematica.Show
End Sub


Private Sub mnPresupuestos_Click()
    Screen.MousePointer = vbHourglass
    frmColPresu.Show vbModal
End Sub

Private Sub mnPunteoBancario_Click()
    frmPunteoBanco.Show vbModal
End Sub

Private Sub mnPunteoCerrados_Click()
    Screen.MousePointer = vbHourglass
    frmPuntear.EjerciciosCerrados = True
    frmPuntear.Show vbModal
End Sub

Private Sub mnPunteoExtractos_Click()
    Screen.MousePointer = vbHourglass
    frmPuntear.EjerciciosCerrados = False
    frmPuntear.Show vbModal
End Sub

Private Sub mnPyGConso_Click()
    AbrirListado 50, False
End Sub

Private Sub mnRatios_Click()
    frmRatios.Show vbModal
End Sub

Private Sub mnRecalculoSaldos_Click()
    frmActualizar.OpcionActualizar = 12
    frmActualizar.NumAsiento = 0
    frmActualizar.FechaAsiento = Now
    frmActualizar.NumDiari = 1
    AlgunAsientoActualizado = False
    frmActualizar.Show vbModal
End Sub

Private Sub mnReemision_Click()
    AbrirListado 6, False
End Sub

Private Sub mnReemisionDiarios_Click()
     AbrirListado 6, True
End Sub

Private Sub mnRegFac_Click()
    Screen.MousePointer = vbHourglass
    frmFacturProv.Show vbModal
End Sub

Private Sub mnRegFacCli_Click()
    Screen.MousePointer = vbHourglass
    frmFacturas.Show vbModal
    CerrarFormularios 1
End Sub

Private Sub mnRelacionClientesVentas_Click()
    AbrirListado 55, False
End Sub

Private Sub mnRelacionProveGastos_Click()
    AbrirListado 56, False
End Sub

Private Sub mnRenumeracion_Click()
    frmCierre.opcion = 0
    frmCierre.Show vbModal
End Sub

Private Sub mnRevisarMultibase_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.opcion = 14
    frmMensajes.Show vbModal
End Sub

Private Sub mnSeleccionarImpresora_Click()
    Screen.MousePointer = vbHourglass
    Me.CommonDialog1.Flags = cdlPDPrintSetup
    Me.CommonDialog1.ShowPrinter
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnSimulaCierre_Click()
    frmCierre.opcion = 4
    frmCierre.Show vbModal
End Sub

Private Sub mnSimulAmortiza_Click()
    frmInmov.opcion = 1
    frmInmov.Show vbModal
End Sub

Private Sub mnTiposDiario_Click()
    Screen.MousePointer = vbHourglass
    frmTiposDiario.Show vbModal
End Sub

Private Sub mnTiposIVA_Click()
    Screen.MousePointer = vbHourglass
    frmIVA.Show vbModal
End Sub

Private Sub mnTotalCtaConceptoCerrado_Click()
    AbrirListado 4, True
End Sub

Private Sub mnTotalesCtaConcepto_Click()
    AbrirListado 4, False
End Sub

Private Sub mnTotalMensyAcumuladosCerrados_Click()
    MsgBox "No esta disponible", vbMsgBoxRight + vbCritical
End Sub

Private Sub mnTraerDeCerrados_Click()
    frmCierre.opcion = 7
    frmCierre.Show vbModal
End Sub

Private Sub mnTrasapasoAce_Click()
    AbrirListado 30, False
End Sub

Private Sub mnTraspasoACerrados_Click()
    frmCierre.opcion = 3
    frmCierre.Show vbModal
End Sub





Private Sub mnTraspasoEntreSecciones_Click(Index As Integer)
    
    'Por si acaso
    If vUsu.Nivel > 1 Then Exit Sub
    
    frmListado2.opcion = 1
    frmListado2.Show vbModal
End Sub

Private Sub mnTraspasoExportar_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.opcion = 16
    frmMensajes.Show vbModal
End Sub

Private Sub mnTraspasoImportar_Click()
    Screen.MousePointer = vbHourglass
    frmMensajes.opcion = 17
    frmMensajes.Show vbModal
End Sub

Private Sub mnTraspasoPersa_Click()
    AbrirListado 29, False
End Sub

Private Sub mnTraspasoPGC2008_Click()
    frmTraspaso.UltimoNivel = False
    frmTraspaso.Show vbModal
End Sub

Private Sub mnTraspasoPGC2008UltNivel_Click()
    frmTraspaso.UltimoNivel = True
    frmTraspaso.Show vbModal
End Sub

Private Sub mnuSal_Click()
    Unload Me
End Sub

Private Sub mnUsuarios_Click()

    frmMantenusu.Show vbModal

End Sub

Private Sub mnUsuariosActivos_Click()
Dim SQL As String
Dim i As Integer
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad(False)
    If CadenaDesdeOtroForm <> "" Then
        i = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            SQL = RecuperaValor(CadenaDesdeOtroForm, i)
            If SQL <> "" Then Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
            i = i + 1
        Loop Until SQL = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningun usuario, además de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    CadenaDesdeOtroForm = ""
End Sub

Private Sub mnVentaBajaInmo_Click()
    frmInmov.opcion = 3
    frmInmov.Show vbModal
End Sub

Private Sub mnVerHistoricoApuntes_Click()
    Screen.MousePointer = vbHourglass
    frmHcoApuntes.EjerciciosCerrados = False
    frmHcoApuntes.ASIENTO = ""
    frmHcoApuntes.LINASI = 0
    frmHcoApuntes.Show vbModal
End Sub

Private Sub mnVerLog_Click()
    Screen.MousePointer = vbHourglass
    Load frmLog
    DoEvents
    frmLog.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnWeb_Click()
    Screen.MousePointer = vbHourglass
    LanzaHome "websoporte"
    espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 1
    
            
    
        'Cosas
'        Dim Cad As String
'        Dim Ini As String
'        Dim RS As ADODB.Recordset
'        Dim I As Integer
'        Dim L As Long
'        Cad = "    select * from hlinapu where codmacta='60300001'"
'        Cad = Cad & " and fechaent>='2006-01-01' and fechaent<='2006-12-31'"
'        Set RS = New ADODB.Recordset
'        RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        While Not RS.EOF
'            Cad = Trim(RS!ampconce)
'            I = InStr(1, Cad, "(")
'            If I > 0 Then
'                Cad = Mid(Cad, 1, I - 1)
'            Else
'
'                Debug.Print "i:0    " & Cad
'            End If
'
'
'            'Ahora veremos los casos
'            'SU FRA. Nº 5218746
'            'ABONO
'            Ini = Mid(Cad, 1, 5)
'            If Ini = "ABONO" Then
'                'ES UN ABONO
'                L = CLng(Mid(Cad, 6))
'            Else
'                If Ini = "SU FR" Then
'                    I = InStr(5, Cad, "/")
'                    If I > 0 Then
'                        L = CLng(Mid(Cad, I + 1))
'                    Else
'                        L = CLng(Mid(Cad, 11))
'                    End If
'
'                Else
'                    Stop
'                End If
'            End If
'            If L = 0 Then Stop
'
'            Cad = "INSERT INTO bdatos (h, fechaent, numdocum, timported, timporteh, importe1) VALUES ("
'            Cad = Cad & "'h','" & Format(RS!fechaent, FormatoFecha) & "','" & CStr(L) & "',"
'            If IsNull(RS!timported) Then
'                Cad = Cad & "0"
'            Else
'                Cad = Cad & TransformaComasPuntos(CStr(RS!timported))
'            End If
'            Cad = Cad & ","
'            If IsNull(RS!timporteh) Then
'                Cad = Cad & "0"
'            Else
'                Cad = Cad & TransformaComasPuntos(CStr(RS!timporteh))
'            End If
'            Cad = Cad & ",0)"
'            Conn.Execute Cad
'            RS.MoveNext
'        Wend
'        RS.Close
'

'       Exit Sub
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
    
        
        'Cuentas
        mnPlanContable_Click
    Case 3
        'Diario
        mnIntroducirAsientos_Click
    Case 4
        mnVerHistoricoApuntes_Click
    Case 5
        mnConsultaExtractos_Click
    Case 6
        AbrirListado 7, False
    Case 8
        mnRegFacCli_Click
    Case 9
        mnRegFac_Click
    Case 10
        'Liquidacion del IVA
        AbrirListado 12, False
    
    Case 12
        mnBalanceMensual_Click
    Case 13
        AbrirListado 26, False
        
    Case 14
        AbrirListado 27, False
        
    Case 16
        mnCambioUsuario_Click
    Case 17
        mnSeleccionarImpresora_Click
        
    Case 22
        Unload Me
    End Select
End Sub


Private Sub PonerDatosVisiblesForm()
Dim Cad As String
    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    Cad = Cad & ", " & Format(Now, "d")
    Cad = Cad & " de " & Format(Now, "mmmm")
    Cad = Cad & " de " & Format(Now, "yyyy")
    Cad = "    " & Cad & "    "
    Me.StatusBar1.Panels(5).Text = Cad
    If vEmpresa Is Nothing Then
        Caption = "ARICONTA" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
    Else
        'Caption = "ARICONTA" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vEmpresa.nomempre & "  -    Usuario: " & vUsu.Nombre
        Caption = "ARICONTA" & " Ver. " & App.Major & "." & App.Minor & "." & App.Revision & "    " & vEmpresa.nomresum & "     Usuario: " & vUsu.Nombre
    End If
End Sub


Private Sub AbrirListado(numero As Byte, Cerrado As Boolean)
    Screen.MousePointer = vbHourglass
    frmListado.EjerciciosCerrados = Cerrado
    frmListado.opcion = numero
    frmListado.Show vbModal
End Sub



Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim Cad As String

    On Error Resume Next
    For Each T In Me
        Cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            If LCase(Mid(T.Name, 1, 6)) <> "mnbarr" Then T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.Visible = Habilitar
    mnParametros.Enabled = True
    mnEmpresa.Enabled = True
    Me.mnParametros.Enabled = True
    Me.mnConfiguracionAplicacion.Enabled = True
    mnDatos.Enabled = True
    Me.mnuSal.Enabled = True
    Me.mnCambioUsuario.Enabled = True
End Sub



Private Sub PonerOpcionesUsuario()
    Dim B As Boolean


    'SOLO ROOT
    B = (vUsu.Codigo Mod 1000) = 0
    Me.mnTraerDeCerrados.Visible = B
    Me.mnUsuarios.Enabled = B
    
    B = vUsu.Nivel < 2  'Administradores y root
    Me.mnParametros.Enabled = B
    Me.mnEmpresa.Enabled = B
    Me.mnParametrosInmo.Enabled = B
    Me.mnHerramientasAriadnaCC.Enabled = B
    If B Then
        'Si tiene permiso solo admin podra  subir ctas
        
        
    End If
        
        
        
    mnAsiePerdyGana.Enabled = B
    mnRenumeracion.Enabled = B
    mnTraspasoACerrados.Enabled = B
    mnBorrarProveedores.Enabled = B
    mnBorrarRegClientes.Enabled = B
    mnDescierre.Enabled = B
    mnVentaBajaInmo.Enabled = B
    mnCaluloYContabilizacion.Enabled = B
    mnDeshacerAmortizacion.Enabled = B
    mnNuevaEmpresa.Enabled = B
    mnRecalculoSaldos.Enabled = B
    mnInformesScrystal.Enabled = B
    Me.mnImportarDatosFiscales.Enabled = B
    
    mnVerLog.Visible = B
    
    'mnPedirPwd.Enabled = B
    B = vUsu.Nivel = 3  'Es usuario de consultas
    If B Then
        mnBorreEjerciciosCerrados.Enabled = False
        mnDiarioOficial.Enabled = False
        mnActalizacionAsientos.Enabled = False
        mnAsientosPredefinidos.Enabled = False
        Me.mnConfigBalPeryGan.Enabled = False
        Me.mnContFactCli.Enabled = False
        Me.mnContFactProv.Enabled = False
        Me.mnPunteoExtractos.Enabled = False
        Me.mnImportarNorma43.Enabled = False
        Me.mnPunteoBancario.Enabled = False
        Me.mnImportarDatosFiscales.Enabled = False
    End If
End Sub



'Private Sub FijarConerrores()
'
'    If (vParam Is Nothing) Or (vEmpresa Is Nothing) Then Exit Sub
'
'    Set miRsAux = New ADODB.Recordset
'    'Asierr
'    CadenaDesdeOtroForm = "Select numasien from linapue"
'    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Me.mnASientosConErrores.Checked = Not miRsAux.EOF
'    miRsAux.Close
'
'    'Clientes con errores
'    CadenaDesdeOtroForm = "Select anofaccl from cabfacte"
'    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Me.mnFacErrCli.Checked = Not miRsAux.EOF
'    miRsAux.Close
'
'    'Facturas con errores
'    CadenaDesdeOtroForm = "Select anofacpr from cabfactprove"
'    miRsAux.Open CadenaDesdeOtroForm, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Me.mnFactErrProv.Checked = Not miRsAux.EOF
'    miRsAux.Close
'    Set miRsAux = Nothing
'End Sub

Private Sub LanzaHome(opcion As String)
    Dim i As Integer
    Dim Cad As String
    On Error GoTo ELanzaHome
    
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBD(opcion, "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en parametros.", vbExclamation
        Exit Sub
    End If
        
    If opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision
        
    i = FreeFile
    Cad = ""
    Open App.path & "\lanzaexp.dat" For Input As #i
    Line Input #i, Cad
    Close #i
    
    'Lanzamos
    If Cad <> "" Then Shell Cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
    
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, Cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub





'Opcions
'   1.- Asiento
'   2.- Fac clientes
'   3.- Fac proveed
Private Function TienenErrores(opcion As Byte) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

    TienenErrores = False
    Set Rs = New ADODB.Recordset
    SQL = "Select count(*) from "
    Select Case opcion
    Case 2
        SQL = SQL & " cabfacte"
    Case 3
        SQL = SQL & " cabfactprove"
    Case Else
        SQL = SQL & " cabapue"
    End Select
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then TienenErrores = True
        End If
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Private Sub EliminarAlgunosDatos(CambioEmpresa As Boolean)
Dim i As Integer
Dim C As String

    On Error GoTo EEliminar

    
    Me.StatusBar1.Panels(2).Text = "Preparando datos"
    Me.StatusBar1.Refresh

    If Not CambioEmpresa Then
        Conn.Execute "DELETE from tmpconextcab where codusu= " & vUsu.Codigo
         
        Conn.Execute "DELETE from tmpconext where codusu= " & vUsu.Codigo
        
    End If
    
    
    'Si ha cambiado de empresa entonces aprovecho y
    'elimino unos cuantos mas
    If CambioEmpresa Then
            'Elimino datos de unos cuantos temporales
            C = "tmplinfactura|"
            
            'A partir de la 2 van a BD usuarios
            C = C & "zhistoapu|zlinccexplo|ztmpconext|ztmpconextcab|ztmpfaclin|ztmpfaclinprov|"
            NumRegElim = 0
            Do
                NumRegElim = NumRegElim + 1
                i = InStr(1, C, "|")
                
                Me.StatusBar1.Panels(2).Text = "Tablas (" & NumRegElim & "/8)"
                Me.StatusBar1.Refresh
                If i > 0 Then
                    'Monto el sql
                    CadenaDesdeOtroForm = Mid(C, 1, i - 1) 'la tabla
                    C = Mid(C, i + 1)
                    If NumRegElim > 1 Then CadenaDesdeOtroForm = "Usuarios." & CadenaDesdeOtroForm
                    CadenaDesdeOtroForm = "DELETE FROM " & CadenaDesdeOtroForm & " WHERE codusu = " & vUsu.Codigo
                
                    EjecutaSQL CadenaDesdeOtroForm
                End If
    
    
    
            Loop Until i = 0
    
    
    End If
    Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
    Me.StatusBar1.Refresh
    CadenaDesdeOtroForm = ""
    
    Exit Sub
EEliminar:
    Err.Clear
End Sub



Private Sub LeerEditorMenus()
Dim SQL As String
    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    SQL = "Select count(*) from usuarios.appmenus where aplicacion='conta'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim SQL As String
Dim C As String

    On Error GoTo ELeerEditorMenus
    
    SQL = "Select * from usuarios.appmenususuario where aplicacion='conta' and codusu = " & Val(Right(CStr(vUsu.Codigo), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            SQL = SQL & miRsAux.Fields(3) & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If SQL <> "" Then
        SQL = "·" & SQL
        For Each T In Me.Controls
            If TypeOf T Is menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                If InStr(1, SQL, C) > 0 Then T.Visible = False
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index & "|"
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function






'''ICONOS
Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub


'Esto es HEAVY heavy
'Esto es HEAVY heavy
'Esto es HEAVY heavy
Private Sub CerrarFormularios(N As Byte)
    On Error GoTo ECerrarFormularios
    
    If N = 1 Then Unload frmFacturas
    
    
    Exit Sub
ECerrarFormularios:
    Err.Clear
End Sub


''_________________________________________________________________-
''
''ERRORES EN EL TRASPASO DE LA CONTABILIDAD AL PGC 2008
''
''_________________________________________________________________-
'Private Sub HacerCambioCuentasInmovilizado()
'Dim SQl As String
'Dim DatosDelTraspaso As Collection
'Dim N As Integer
'
'    Set DatosDelTraspaso = New Collection
'    DatosDelTraspaso.Add "28200|28180|"
'    DatosDelTraspaso.Add "28201|28181|"
'    DatosDelTraspaso.Add "28202|28182|"
'    DatosDelTraspaso.Add "28203|28130|"
'    DatosDelTraspaso.Add "28204|28140|"
'    DatosDelTraspaso.Add "28205|28150|"
'    DatosDelTraspaso.Add "28206|28160|"
'    DatosDelTraspaso.Add "28207|28170|"
'    DatosDelTraspaso.Add "28209|28190|"
'
'    'Para cada subgrupo iremos obteniendo la rista de hlinapu E INMOVILIZADO
'    'y actualizando
'
'    For N = 1 To DatosDelTraspaso.Count
'        Set miRsAux = New ADODB.Recordset
'        HacerCambiosParaGrupo DatosDelTraspaso(N)
'        Set miRsAux = Nothing
'    Next N
'
'
'End Sub
'
'
'Private Sub HacerCambiosParaGrupo(CADENA As String)
'Dim SQl As String
'Dim NuevaCuenta As String
'
'    SQl = "'" & RecuperaValor(CADENA, 1) & "%'"
'    SQl = "Select codmacta,nommacta from cuentas where apudirec='S' AND codmacta like " & SQl
'    miRsAux.Open SQl, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not miRsAux.EOF
'        'Obtenemos nueva cuenta
'        SQl = RecuperaValor(CADENA, 2)
'        NuevaCuenta = SQl & Mid(miRsAux!Codmacta, Len(SQl) + 1)
'
'
'        'Insertamos en cuentas. Si da error NO pasa nada
'        SQl = "INSERT INTO cuentas (codmacta,Nommacta,apudirec) VALUES ('" & NuevaCuenta & "','" & DevNombreSQL(miRsAux!nommacta) & "','S')"
'        If Not EjecutaSQL(SQl) Then MsgBox "Error : " & SQl, vbInformation
'
'        'UPDATEO hlinapu
'        HazCambioTabla "hlinapu", "codmacta|ctacontr|", 2, NuevaCuenta
'
'        'En sinmov
'        HazCambioTabla "sinmov", "codmact1|codmact2|codmact3|", 3, NuevaCuenta
'
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'
'End Sub
'
'Private Sub HazCambioTabla(Tabla As String, Campos As String, NumeroCampos As Integer, NuevaCuenta As String)
'Dim I As Integer
'Dim Campo As String
'Dim SQl As String
'
'    For I = 1 To NumeroCampos
'        Campo = RecuperaValor(Campos, I)
'        SQl = MontaSQLCambioCuenta(Tabla, Campo, NuevaCuenta, CStr(miRsAux!Codmacta))
'        Conn.Execute SQl
'    Next I
'
'End Sub
'
'Private Function MontaSQLCambioCuenta(Ta As String, Campo1, ValorNuevo, valorAntiguo) As String
'    MontaSQLCambioCuenta = "UPDATE " & Ta & " SET " & Campo1 & " = '" & ValorNuevo & "' WHERE " & Campo1 & " = '" & valorAntiguo & "'"
'End Function
