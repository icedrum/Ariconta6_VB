VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#16.2#0"; "Codejock.SkinFramework.v16.2.0.ocx"
Begin VB.Form frmPpalMenu 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmPpalMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   7858
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin VB.Frame FrameSeparador 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      DragMode        =   1  'Automatic
      Height          =   3015
      Left            =   0
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   0
      Width           =   45
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5535
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9763
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   1455
      Left            =   3000
      TabIndex        =   6
      Top             =   6840
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   2566
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1695
      Left            =   2940
      TabIndex        =   7
      Top             =   6840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mensaje"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   1695
      Left            =   7380
      TabIndex        =   8
      Top             =   6840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   1230
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   390
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   1200
      Top             =   7140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   1170
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   2010
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   1980
      Top             =   7140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   1980
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   510
      Top             =   7320
      _Version        =   1048578
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblMsgUsu 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4620
      TabIndex        =   12
      Top             =   6690
      Width           =   855
   End
   Begin VB.Label lblMsgUsu 
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6060
      TabIndex        =   11
      Top             =   6690
      Width           =   855
   End
   Begin VB.Label lblMsgApli 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7620
      TabIndex        =   10
      Top             =   6690
      Width           =   855
   End
   Begin VB.Label lblMsgApli 
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8940
      TabIndex        =   9
      Top             =   6690
      Width           =   855
   End
   Begin VB.Label lblMsgUsu 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   6600
      Width           =   975
   End
   Begin VB.Image ImageLogo 
      Height          =   570
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label33 
      BackColor       =   &H004D2C1D&
      Caption         =   "   Ariconta2014"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9195
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   7695
      Left            =   -120
      Top             =   840
      Width           =   10455
   End
   Begin VB.Label Label22 
      BackColor       =   &H004D2C1D&
      Height          =   690
      Left            =   7440
      TabIndex        =   4
      Top             =   -120
      Width           =   3135
   End
   Begin VB.Menu mnPopUp 
      Caption         =   "mnPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnPopUp1 
         Caption         =   "Editar"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnPopUp1 
         Caption         =   "Eliminar"
         Index           =   1
      End
      Begin VB.Menu mnPopUp1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnPopUp1 
         Caption         =   "Organizar"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmPpalMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nomempre As String  'Vendra con los parametrros,


Public UnaVez As Boolean
Dim Base
Dim AnchoListview As Integer

Dim PrimeraVez As Boolean


Private Sub Form_Activate()

    Screen.MousePointer = vbHourglass
    If UnaVez Then
        UnaVez = False
        CargaMenu "ariconta", Me.TreeView1
'        CargaMenu "introcon", Me.TreeView2
        
        MenuComoEstaba
'--
'        CargaShortCuts 0
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
        
'    Me.Icon = frmEntrada.Icon
    PrimeraVez = True
    
    ImageList1.ImageHeight = 48
    ImageList1.ImageWidth = 48
    GetIconsFromLibrary App.path & "\icoconppal.dll", 1, 48

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

    
    ' sirve para calcular despues el width
    Base = 1290
    Base = Base + 550 '550 es lo k mide de alto la imagen de ariadna
   

'    Me.ListView1.Icons = frmEntrada.ImaListPersonalizaIco
    
    PonerCaption
       
    PonerDatosFormulario

    EstablecerSkin CInt(2)
    

End Sub



Private Sub PonerCaption()
        Caption = "AriCONTA 14    V-" & App.Major & "." & App.Minor & "." & App.Revision & "    usuario: " & vUsu.Nombre
        Label33.Caption = "   " & vEmpresa.nomempre
End Sub



Private Sub Form_Resize()
    Dim X, Y As Integer
Dim V ''


If WindowState = 1 Then Exit Sub ' ha pulsado minimizar
X = Me.Width
Y = Me.Height
If X < 5990 Then Me.Width = 5990
If Y < 4100 Then Me.Height = 4100
ImageLogo.Left = Me.Width - ImageLogo.Width - 240
Label33.Left = 30
'Text1.Width = Me.Width - Text1.Left - 250
X = Me.Height - Base

TreeView1.Height = X
X = X \ 6
ListView1.Height = X * 4

ListView2.Top = ListView1.Top + ListView1.Height + 500
ListView2.Height = Me.Height - ListView2.Top - 850
ListView3.Top = ListView2.Top
ListView3.Height = ListView2.Height








Y = Me.Width - 200
Y = ((30 / 100) * Y)

TreeView1.Left = 30
TreeView1.Width = Y - 30

'Separador
Me.FrameSeparador.Left = Y + 15
Me.FrameSeparador.Top = TreeView1.Top
Me.FrameSeparador.Height = Me.TreeView1.Height

ListView1.Left = Y + 60
Me.ListView2.Left = Y + 60


AnchoListview = Me.Width - 200 - Y - 30
ListView1.Width = AnchoListview
V = Me.ImageLogo.Left
Label33.Width = V + 20
Label33.Left = -15
Label22.Left = Label33.Width - 120
Label22.Width = Me.Width - Label22.Left
Label22.Top = 0


X = AnchoListview \ 3
ListView2.Width = 2 * X
Me.ListView3.Left = Me.ListView2.Left + Me.ListView2.Width + 30
ListView3.Width = X

'Dos listiview
For X = 0 To 2
    lblMsgUsu(X).Top = ListView2.Top - 240
    If X < 2 Then lblMsgApli(X).Top = ListView2.Top - 240
Next
'Left
lblMsgUsu(0).Left = ListView2.Left + 60
lblMsgUsu(1).Left = lblMsgUsu(0).Left + lblMsgUsu(0).Width + 120
lblMsgUsu(2).Left = lblMsgUsu(1).Left + lblMsgUsu(1).Width + 120
lblMsgApli(0).Left = ListView3.Left + 60
lblMsgApli(1).Left = ListView3.Left + lblMsgApli(0).Width + 30





Shape1.Width = Me.Width - Shape1.Left - 50
Shape1.Height = Me.Height - Shape1.Top - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    FijarUltimoSkin False
'    FreeLibrary m_hMod: UnloadApp: End
    ActualizarExpansionMenus vUsu.Id, Me.TreeView1, "ariconta"
    
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    AbrirFormularios CLng(Mid(ListView1.SelectedItem.Key, 3))
End Sub

Private Sub AbrirFormularios(Accion As Long)

    Select Case Accion
    Case 101 ' empresa
        frmempresa.Show vbModal
    Case 102 ' parametros contabilidad
        If Not (vEmpresa Is Nothing) Then
            frmparametros.Show vbModal
        End If
    Case 103 ' parametros tesoreria
    Case 104 ' contadores
        Screen.MousePointer = vbHourglass
        frmContadores.Show vbModal
    Case 105 ' usuarios
        frmMantenusu.Show vbModal
    Case 106 ' informes
        frmCrystal.Show vbModal
    Case 107 ' crear nueva empresa
        If vUsu.Nivel > 1 Then Exit Sub
        
        frmCentroControl.Opcion = 2
        frmCentroControl.Show vbModal
    Case 108 ' acerca de...
        Screen.MousePointer = vbHourglass
        frmMensajes.Opcion = 6
        frmMensajes.Show vbModal
    Case 201 ' plan contable
        Screen.MousePointer = vbHourglass
        frmColCtas.ConfigurarBalances = 0
        frmColCtas.DatosADevolverBusqueda = ""
        frmColCtas.Show vbModal
    Case 202 ' tipos de diario
        Screen.MousePointer = vbHourglass
        frmTiposDiario.Show vbModal
    Case 203 ' conceptos
        Screen.MousePointer = vbHourglass
        frmConceptos.Show vbModal
    Case 204 ' tipos de iva
        Screen.MousePointer = vbHourglass
        frmIVA.Show vbModal
    Case 205 ' tipos de pago
    Case 206 ' formas de pago
    Case 207 ' bancos
    Case 209 ' agentes
    Case 210 ' departamentos
    Case 211 ' asientos predefinidos
        Screen.MousePointer = vbHourglass
        frmAsiPre.Show vbModal
    Case 212 ' cartas de reclamacion
    
    Case 301 ' asientos
        Screen.MousePointer = vbHourglass
        frmAsientos.ASIENTO = ""
        frmAsientos.Show vbModal
    Case 302 ' historico
        Screen.MousePointer = vbHourglass
        frmHcoApuntes.EjerciciosCerrados = True
        frmHcoApuntes.ASIENTO = ""
        frmHcoApuntes.LINASI = 0
        frmHcoApuntes.Show vbModal
    Case 303 ' extractos
        Screen.MousePointer = vbHourglass
        frmConExtr.EjerciciosCerrados = False
        frmConExtr.Cuenta = ""
        frmConExtr.Show vbModal
    Case 304 ' punteo
        Screen.MousePointer = vbHourglass
        frmPuntear.EjerciciosCerrados = False
        frmPuntear.Show vbModal
    Case 305 ' reemision de diarios
        AbrirListado 6, False
    Case 306 ' sumas y saldos
        AbrirListado 5, False
    Case 307 ' cuenta de explotacion
        AbrirListado 7, False
    Case 308 ' balance de situacion
        AbrirListado 26, False
    Case 309 ' perdidas y ganancias
        AbrirListado 27, False
    Case 310 ' totales por concepto
        AbrirListado 4, False
    Case 311 ' evolucion de saldos
        AbrirListado 54, False
    Case 312 ' ratios y graficas
        frmRatios.Show vbModal
    Case 313 ' importar N43
        frmUtiliBanco.Show vbModal
    Case 314 ' puntero extracto bancario
        frmPunteoBanco.Show vbModal
    Case 401 ' emitidas
        Screen.MousePointer = vbHourglass
        frmFacturas.Show vbModal
        CerrarFormularios 1
    Case 402 ' libro emitidas
        AbrirListado 8, False
    Case 403 ' relacion clientes por cuenta
        AbrirListado 55, False
    Case 404 ' recibidas
        Screen.MousePointer = vbHourglass
        frmFacturProv.Show vbModal
    Case 405 ' libro recibidas
        AbrirListado 13, False
    Case 406 ' relacion proveedores por cuenta
        AbrirListado 56, False
    Case 407 ' liquidacion iva
        AbrirListado 12, False
    Case 408 ' certificado iva
        AbrirListado 11, False
    Case 409 ' modelo 340
        frmListado2.Opcion = 0
        frmListado2.Show vbModal
    Case 410 ' modelo 347
        AbrirListado 20, False
    Case 411 ' modelo 349
        AbrirListado 28, False
    
    Case 501 ' parametros inmovilizado
        frmInmov.Opcion = 0
        frmInmov.Show vbModal
    Case 502 ' conceptos
        Screen.MousePointer = vbHourglass
        frmConceptosInmo.Show vbModal
    Case 503 ' elementos
        frmEltoInmo.DatosADevolverBusqueda = ""
        frmEltoInmo.Show vbModal
    Case 504 ' ficha de elementos
        frmInmov.Opcion = 5
        frmInmov.Show vbModal
    Case 505 ' estadistica
        frmInmov.Opcion = 4
        frmInmov.Show vbModal
    Case 506 ' estadistica entre fechas
        frmInmov.Opcion = 6
        frmInmov.Show vbModal
    Case 507 ' historico inmovilizado
        Screen.MousePointer = vbHourglass
        frmHcoInmo.Show vbModal
    Case 508 ' simulacion
        frmInmov.Opcion = 1
        frmInmov.Show vbModal
    Case 509 ' calculo y contabilizacion
        frmInmov.Opcion = 2
        frmInmov.Show vbModal
    Case 510 ' deshacer amortizacion
        frmInmov.Opcion = 10
        frmInmov.Show vbModal
    Case 511 ' venta-baja inmmovilizado
        frmInmov.Opcion = 3
        frmInmov.Show vbModal
    Case 601 ' cartera de cobros
    Case 602 ' informe de cobros pendientes
    Case 603 ' impresion de recibos
    Case 604 ' realizar cobro
    Case 605 ' transferencia abonos
    Case 606 ' compensaciones
    Case 607 ' compensar cliente
    Case 608 ' reclamaciones
    
    Case 701 ' remesas
    Case 702 ' cancelacion cliente
    Case 703 ' abono remesa
    Case 704 ' Devoluciones
    Case 704 ' Devoluciones
    Case 705 ' Eliminar riesgo
    Case 706 ' Informe Impagados
    Case 707 ' Recepción Talón-Pagaré
    Case 708 ' Remesas Talón-Pagaré
    Case 709 ' Abono remesa
    Case 710 ' Devoluciones
    Case 711 ' Eliminar riesgo
    
    Case 801 ' Cartera de Pagos
    Case 802 ' Informe Pagos pendientes
    Case 803 ' Informe Pagos bancos
    Case 804 ' Realizar Pago
    Case 805 ' Transferencias
    Case 806 ' Pagos domiciliados
    Case 807 ' Gastos Fijos
    Case 808 ' Memoria Pagos proveedores
    
    Case 901 ' Informe por NIF
    Case 902 ' Informe por cuenta
    Case 903 ' Situación Tesoreria
    
    ' Analitica
    Case 1001 ' Centros de Coste
        frmCCoste.DatosADevolverBusqueda = ""
        frmCCoste.Show vbModal
    Case 1002 ' Consulta de Saldos
        AbrirListado 15, False
    Case 1003 ' Cuenta de Explotación
        AbrirListado 16, False
    Case 1004 ' Centros de coste por cuenta
        AbrirListado 17, False
    Case 1005 ' Detalle de explotación
        AbrirListado 19, False
        
    ' Presupuestaria
    Case 1101 ' Presupuestos
        Screen.MousePointer = vbHourglass
        frmColPresu.Show vbModal
    Case 1102 ' Listado de Presupuestos
        AbrirListado 9, False
    Case 1103 ' Balance Presupuestario
        AbrirListado 10, False
        
    ' Consolidado
    Case 1201 ' Sumas y Saldos
        AbrirListado 24, False
    Case 1202 ' Balance de Situación
        AbrirListado 51, False
    Case 1203 ' Pérdidas y Ganancias
        AbrirListado 50, False
    Case 1204 ' Cuenta de Explotación
        AbrirListado 31, False
    Case 1205 ' Listado Facturas Clientes
        AbrirListado 53, False
    Case 1206 ' Listado Facturas Proveedores
        AbrirListado 52, False
    
    ' Cierre de Ejercicio
    Case 1301 ' Renumeración de asientos
        frmCierre.Opcion = 0
        frmCierre.Show vbModal
    Case 1302 ' Simulación de cierre
        frmCierre.Opcion = 4
        frmCierre.Show vbModal
    Case 1303 ' Cierre de Ejercicio
        frmCierre.Opcion = 1
        frmCierre.Show vbModal
    Case 1304 ' Deshacer cierre
        frmCierre.Opcion = 5
        frmCierre.Show vbModal
    Case 1305 ' Diario Oficial
        AbrirListado 14, False
    Case 1306 ' Diario Oficial Resumen
        AbrirListado 18, False
    Case 1307 ' Presentación cuentas anuales
        Telematica 0
    Case 1308 ' Presentación Telemática de Libros
        Telematica 1
    
    ' Utilidades
    Case 1401 ' Comprobar cuadre
        Screen.MousePointer = vbHourglass
        frmMensajes.Opcion = 2
        frmMensajes.Show vbModal
    Case 1402 ' Recalculo de Saldos
        frmActualizar.OpcionActualizar = 12
        frmActualizar.NumAsiento = 0
        frmActualizar.FechaAsiento = Now
        frmActualizar.NumDiari = 1
        AlgunAsientoActualizado = False
        frmActualizar.Show vbModal
    Case 1403 ' Revisar caracteres especiales
        Screen.MousePointer = vbHourglass
        frmMensajes.Opcion = 14
        frmMensajes.Show vbModal
    
    Case 1404 ' Agrupacion cuentas
    Case 1405 'Buscar ...
    
    Case 1406 'Configurar Balances
        Screen.MousePointer = vbHourglass
        frmColBalan.Show vbModal
    Case 1407 'Desbloquear asientos
        mnHerrAriadnaCC_Click (0)
    Case 1408 'Mover cuentas
        mnHerrAriadnaCC_Click (1)
    Case 1409 'Renumerar registros proveedor
        mnHerrAriadnaCC_Click (2)
    Case 1410 'Aumentar dígitos contables
        mnHerrAriadnaCC_Click (4)
    Case 1411 'cambio de iva
        mnHerrAriadnaCC_Click (5)
    Case 1412 'log de acciones
        Screen.MousePointer = vbHourglass
        Load frmLog
        DoEvents
        frmLog.Show vbModal
        Screen.MousePointer = vbDefault
    Case 1413 'usuarios activos
        mnUsuariosActivos_Click
    
    Case Else
  
    End Select

End Sub

Private Sub mnHerrAriadnaCC_Click(Index As Integer)
 
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene permisos", vbExclamation
            Exit Sub
        End If
        'El index 3 , que es la barra, en frmCC es la opcion de NUEVA EMPRESA
        ' y no se llma desde aqui, con lo cual no hay problemo
        'Para el restro cojo el valor del helpidi
        
        frmCentroControl.Opcion = Index
        frmCentroControl.Show vbModal
    
End Sub

Private Sub Telematica(Caso As Integer)
        Me.Enabled = False
        frmTelematica.Opcion = Caso
        frmTelematica.Show
End Sub
    

'El usuarios si tiene maximizada unas cosas y minimiazadas otras se las guardaremos
Private Sub MenuComoEstaba()
Dim N As Node
Dim SQL As String

    For I = 1 To Me.TreeView1.Nodes.Count
        SQL = "select expandido from menus_usuarios where codusu = " & DBSet(vUsu.Id, "N") & " and codigo in (select codigo from menus where descripcion = " & DBSet(Me.TreeView1.Nodes(I), "T") & ")"

        If DevuelveValor(SQL) = 0 Then
            Me.TreeView1.Nodes(I).Expanded = False
        Else
            Me.TreeView1.Nodes(I).Expanded = True
        End If
                
    Next I

'-- de David
'    Set N = Me.TreeView1.Nodes(1)
'    While Not N Is Nothing
'        N.Expanded = True
'        Set N = N.Next
'    Wend
End Sub

Private Sub OcultarHijos(Padre As String)
Dim SQL As String

    SQL = "update menus_usuarios set ver = 0 where codusu = " & vUsu.Id & " and padre = " & DBSet(Padre, "N")

    Conn.Execute SQL
    
End Sub



Private Sub CargaShortCuts(Seleccionado As Long)
Dim AUx As String

    'Para cada usuarios, y a partir del menu del que disponga
    Set miRsAux = New ADODB.Recordset
    AUx = "Select * from  usuarios.usuariosiconosppal WHERE codusu =" & vUsu.Codigo & " AND aplicacion='ariconta'"
    miRsAux.Open AUx, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AnchoListview = 0
    ListView1.ListItems.Clear
    While Not miRsAux.EOF
            AnchoListview = AnchoListview + 1
            Me.ListView1.ListItems.Add , CStr("LW" & Format(miRsAux!PuntoMenu, "000000")), miRsAux!TextoVisible, CInt(miRsAux!icono)
            If miRsAux!PuntoMenu = Seleccionado Then Set ListView1.SelectedItem = Me.ListView1.ListItems(AnchoListview)
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListView1_DblClick
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then PopupMenu mnPopUp

End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    
    Caption = Data.GetData(1)
    
    If ListView1.ListItems.Count > 8 Then
        MsgBox "Numero maximo de accesos directos superado", vbExclamation
        Exit Sub
    End If
    
    
    'Aqui tendrmos la configuracion perosnalizada
    If TreeView1.SelectedItem = Data.GetData(1) Then
        'OK. El nodo selecionado es el que estamos moviendo
        If TreeView1.SelectedItem.Children > 0 Then
            MsgBox "Solo ultimo nivel", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Error en drag/drop", vbExclamation
        Exit Sub
    End If
    
    '
    LanzaPersonalizarEdicion Val(Mid(TreeView1.SelectedItem.Key, 3)), TreeView1.SelectedItem.Text
    
    
End Sub



Private Sub LanzaPersonalizarEdicion(Valor As Long, TextoInicio As String)
    'AHORA, de momento NO dejamos personalizar los ICONOS NI LOS textos
    'El el form al final hace un:
    '   REPLACE INTO usuarios.usuariosiconosppal(codusu,aplicacion,PuntoMenu,icono,TextoOrigen,TextoVisible) VALUES (1,'ariconta',7,1,'Parámetros','Parámetros')
    
    'frmMenusPersonalizaIconos.TextoMenu = TextoInicio
    'frmMenusPersonalizaIconos.idPuntoMenu = valor
    'frmMenusPersonalizaIconos.Show vbModal
    Msg$ = "REPLACE INTO usuarios.usuariosiconosppal(codusu,aplicacion,PuntoMenu,icono,TextoOrigen,TextoVisible) "
    Msg = Msg$ & " VALUES (" & vUsu.Codigo & ",'ariconta'," & Valor & "," & Valor
    Msg = Msg$ & "," & DBSet(TextoInicio, "T") & "," & DBSet(TextoInicio, "T") & ")"
    
    Conn.Execute Msg$
    espera 0.2
    CargaShortCuts Valor
    
End Sub


Private Sub mnPopUp1_Click(Index As Integer)
    If Index <= 1 Then
        If ListView1.SelectedItem Is Nothing Then Exit Sub
    End If
    
    
    Select Case Index
    Case 0
        'LanzaPersonalizarEdicion Val(Mid(ListView1.SelectedItem.Key, 3)), ""
    Case 1
        If MsgBox("Desea eliminar el acceso directo: " & Me.ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbYes Then
            
            Conn.Execute "DELETE from  usuarios.usuariosiconosppal WHERE codusu =" & vUsu.Codigo & " AND aplicacion='ariconta' AND PuntoMenu =" & Mid(ListView1.SelectedItem.Key, 3)
            CargaShortCuts 0
        End If
    Case 3
    
    End Select
        
    
End Sub

Private Sub TreeView1_DblClick()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Children > 0 Then Exit Sub
    
    AbrirFormularios CLng(Mid(TreeView1.SelectedItem.Key, 3))
    
End Sub


Private Sub CambiarEmpresa()



    CadenaDesdeOtroForm = vUsu.Login & "|" & vEmpresa.codempre & "|"
'--
'    frmPrimerLogin.Show vbModal
    If vUsu Is Nothing Then
        'NO HA CAMBIADO DE Empresa
        Set vUsu = New Usuario
        vUsu.Leer RecuperaValor(CadenaDesdeOtroForm, 1)
        
        AbrirConexion  '"ariconta" & RecuperaValor(CadenaDesdeOtroForm, 2)
        Set vEmpresa = New Cempresa
        Set vParam = New Cparametros
        'NO DEBERIAN DAR ERROR
        vEmpresa.Leer
        vParam.Leer
        
        
        vUsu.CargaPermisosEspeciales  'los cargamos aqui
    Else
        'SI ha cabiado de emrpesa
        PonerCaption
    End If
End Sub

Private Sub TreeView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    
    If Not TreeView1.SelectedItem Is Nothing Then
        If TreeView1.SelectedItem.Children > 0 Then TreeView1.Drag vbCancel
    End If
            
    
End Sub


Private Sub CerrarFormularios(N As Byte)
    On Error GoTo ECerrarFormularios
    
    If N = 1 Then Unload frmFacturas
    
    
    Exit Sub
ECerrarFormularios:
    Err.Clear
End Sub

Private Sub mnUsuariosActivos_Click()
Dim SQL As String
Dim I As Integer
    CadenaDesdeOtroForm = OtrosPCsContraContabiliad(False)
    If CadenaDesdeOtroForm <> "" Then
        I = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            SQL = RecuperaValor(CadenaDesdeOtroForm, I)
            If SQL <> "" Then Me.Tag = Me.Tag & "    - " & SQL & vbCrLf
            I = I + 1
        Loop Until SQL = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningun usuario, además de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    CadenaDesdeOtroForm = ""
End Sub


Private Sub AbrirListado(numero As Byte, Cerrado As Boolean)
    Screen.MousePointer = vbHourglass
    frmListado.EjerciciosCerrados = Cerrado
    frmListado.Opcion = numero
    frmListado.Show vbModal
End Sub


Private Sub PonerDatosFormulario()
Dim Config As Boolean

'    Config = (vParam Is Nothing) Or (vEmpresa Is Nothing)
'
'    If Not Config Then HabilitarSoloPrametros_o_Empresas True
'
'    'FijarConerrores
'    CadenaDesdeOtroForm = ""
'
'    'Poner datos visible del form
'    PonerDatosVisiblesForm
'    'Poner opciones de nivel de usuario
'    PonerOpcionesUsuario
'
'
'    If Not Config Then
'        Me.mnTraspasoEntreSecciones(0).Visible = vParam.TraspasCtasBanco > 0
'        Me.mnTraspasoEntreSecciones(1).Visible = mnTraspasoEntreSecciones(0).Visible
'    End If
'    'Habilitar
'    If Config Then HabilitarSoloPrametros_o_Empresas False
'    'Panel con el nombre de la empresa
'    If Not vEmpresa Is Nothing Then
'        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
'    Else
'        Me.StatusBar1.Panels(2).Text = "Falta configurar"
'    End If
'
'    'Primero los pongo a visible
'    mnDatosExternos347.Visible = True
'    mnbarra101.Visible = True
'
'
'
'
'    'Si tiene editor de menus
'    If TieneEditorDeMenus Then PoneMenusDelEditor
'
'     mnCheckVersion.Visible = False 'Siempre oculto
'
'
'    If Not Config Then
'        mnDatosExternos347.Visible = mnDatosExternos347.Visible And vParam.AgenciaViajes
'        mnbarra101.Visible = mnbarra101.Visible And vParam.AgenciaViajes
'    End If
'    '---------------------------------------------------
'    'Las asociaciones entre menu y botones  del TOOLBAR
'    With Me.Toolbar1
'        .Buttons(1).Visible = mnDatos.Visible And Me.mnPlanContable.Visible
'        '---
'        .Buttons(3).Visible = mnDiario.Visible And Me.mnIntroducirAsientos.Visible    'Diario
'        .Buttons(4).Visible = mnHcoApuntes.Visible And mnVerHistoricoApuntes.Visible    'Hco
'        .Buttons(5).Visible = mnDiario.Visible And mnConsultaExtractos.Visible   'Con extractos
'        .Buttons(6).Visible = mnHcoApuntes.Visible And mnCtaExplotacion.Visible   'CTA EXPLOTACION
'        '----
'        .Buttons(8).Visible = mnMenuIVA.Visible And mnClientes.Visible And Me.mnRegFacCli.Visible     'Fac CLI
'        .Buttons(9).Visible = mnMenuIVA.Visible And mnMenuProveedores.Visible And Me.mnRegFac.Visible    'Fac PRO
'        .Buttons(10).Visible = mnMenuIVA.Visible And Me.mnLiquidacion.Visible   'Liquidacion IVA
'        '----
'        .Buttons(12).Visible = mnHcoApuntes.Visible And mnBalanceMensual.Visible  'Balance
'        .Buttons(13).Visible = mnHcoApuntes.Visible And mnBalancesituacion.Visible
'        .Buttons(14).Visible = mnHcoApuntes.Visible And Me.mnPerdyGan.Visible  'Cuenta P y G
'        '----
'        .Buttons(16).Image = 8  'Usuarios
'        .Buttons(17).Image = 9  'Impresora
'        '----
'        .Buttons(19).Visible = TieneIntegracionesPendientes
'        .Buttons(19).Image = 11
'        'Antes
'        .Buttons(20).Visible = False
'        '.Buttons(20).Visible = BuscarIntegraciones(True)
'        .Buttons(20).Image = 12
'        '----
'        .Buttons(22).Image = 10 'Salir
'    End With
'
'    'Si el usuario tiene permiso para ver los balances, le dejo las graficas
'    Me.mnRatios.Visible = Toolbar1.Buttons(12).Visible
'
End Sub


Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim Cad As String
'
'    On Error Resume Next
'    For Each T In Me
'        Cad = T.Name
'        If Mid(T.Name, 1, 2) = "mn" Then
'            If LCase(Mid(T.Name, 1, 6)) <> "mnbarr" Then T.Enabled = Habilitar
'        End If
'    Next
'    Me.Toolbar1.Enabled = Habilitar
'    Me.Toolbar1.Visible = Habilitar
'    mnParametros.Enabled = True
'    mnEmpresa.Enabled = True
'    Me.mnParametros.Enabled = True
'    Me.mnConfiguracionAplicacion.Enabled = True
'    mnDatos.Enabled = True
'    Me.mnuSal.Enabled = True
'    Me.mnCambioUsuario.Enabled = True
End Sub

Private Sub PonerDatosVisiblesForm()
Dim Cad As String
'    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
'    Cad = Cad & ", " & Format(Now, "d")
'    Cad = Cad & " de " & Format(Now, "mmmm")
'    Cad = Cad & " de " & Format(Now, "yyyy")
'    Cad = "    " & Cad & "    "
'    Me.StatusBar1.Panels(5).Text = Cad
'    If vEmpresa Is Nothing Then
'        Caption = "ARICONTA" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
'    Else
'        'Caption = "ARICONTA" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vEmpresa.nomempre & "  -    Usuario: " & vUsu.Nombre
'        Caption = "ARICONTA" & " Ver. " & App.Major & "." & App.Minor & "." & App.Revision & "    " & vEmpresa.nomresum & "     Usuario: " & vUsu.Nombre
'    End If
End Sub


Private Sub PonerOpcionesUsuario()
    Dim B As Boolean

'
'    'SOLO ROOT
'    B = (vUsu.Codigo Mod 1000) = 0
'    Me.mnTraerDeCerrados.Visible = B
'    Me.mnUsuarios.Enabled = B
'
'    B = vUsu.Nivel < 2  'Administradores y root
'    Me.mnParametros.Enabled = B
'    Me.mnEmpresa.Enabled = B
'    Me.mnParametrosInmo.Enabled = B
'    Me.mnHerramientasAriadnaCC.Enabled = B
'    If B Then
'        'Si tiene permiso solo admin podra  subir ctas
'
'
'    End If
'
'
'
'    mnAsiePerdyGana.Enabled = B
'    mnRenumeracion.Enabled = B
'    mnTraspasoACerrados.Enabled = B
'    mnBorrarProveedores.Enabled = B
'    mnBorrarRegClientes.Enabled = B
'    mnDescierre.Enabled = B
'    mnVentaBajaInmo.Enabled = B
'    mnCaluloYContabilizacion.Enabled = B
'    mnDeshacerAmortizacion.Enabled = B
'    mnNuevaEmpresa.Enabled = B
'    mnRecalculoSaldos.Enabled = B
'    mnInformesScrystal.Enabled = B
'    Me.mnImportarDatosFiscales.Enabled = B
'
'    mnVerLog.Visible = B
'
'    'mnPedirPwd.Enabled = B
'    B = vUsu.Nivel = 3  'Es usuario de consultas
'    If B Then
'        mnBorreEjerciciosCerrados.Enabled = False
'        mnDiarioOficial.Enabled = False
'        mnActalizacionAsientos.Enabled = False
'        mnAsientosPredefinidos.Enabled = False
'        Me.mnConfigBalPeryGan.Enabled = False
'        Me.mnContFactCli.Enabled = False
'        Me.mnContFactProv.Enabled = False
'        Me.mnPunteoExtractos.Enabled = False
'        Me.mnImportarNorma43.Enabled = False
'        Me.mnPunteoBancario.Enabled = False
'        Me.mnImportarDatosFiscales.Enabled = False
'    End If
End Sub

'''ICONOS
Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim I As Integer
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

Public Sub EstablecerSkin(QueSkin As Integer)


  FijaSkin QueSkin

  ' Cargando el archivo del Skin
  ' ============================
    frmPpalMenu.SkinFramework.LoadSkin Skn$, ""
    frmPpalMenu.SkinFramework.ApplyWindow frmPpalMenu.hWnd
    frmPpalMenu.SkinFramework.ApplyOptions = frmPpalMenu.SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
'
    
End Sub

Private Function FijaSkin(Cual As Integer)
  Select Case (Cual)
    Case 0:     ' Windows Luna XP Modificado
      Skn$ = CStr(App.path & "\Styles\WinXP.Luna.cjstyles")
      frmPpalMenu.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
    Case 1:     ' Windows Royale Modificado
      Skn$ = CStr(App.path & "\Styles\WinXP.Royale.cjstyles")
      frmPpalMenu.SkinFramework.LoadSkin Skn$, "NormalRoyale.ini"
    Case 2:     ' Microsoft Office 2007
      Skn$ = CStr(App.path & "\Styles\Office2007.cjstyles")
      frmPpalMenu.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
    Case 3:     ' Windows Vista Sencillo
      Skn$ = CStr(App.path & "\Styles\Vista.cjstyles")
      frmPpalMenu.SkinFramework.LoadSkin Skn$, "NormalBlue.ini"
  End Select

End Function


