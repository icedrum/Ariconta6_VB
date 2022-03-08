VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESGastosFijos 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos Fijos"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmTESGastosFijos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6930
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   8475
      Begin VB.Frame FrameFiltro 
         Height          =   705
         Left            =   4950
         TabIndex        =   10
         Top             =   210
         Width           =   2415
         Begin VB.ComboBox cboFiltro 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmTESGastosFijos.frx":000C
            Left            =   90
            List            =   "frmTESGastosFijos.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   210
            Width           =   2235
         End
      End
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   240
         TabIndex        =   8
         Top             =   3810
         Width           =   2415
         Begin MSComctlLib.Toolbar ToolbarAux 
            Height          =   330
            Left            =   180
            TabIndex        =   9
            Top             =   150
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Generar gastos fijos"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Contabilizar"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   300
         TabIndex        =   5
         Top             =   8310
         Width           =   1755
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   210
            Width           =   1200
         End
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   2
         Top             =   180
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   210
            TabIndex        =   0
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   10
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
                  Object.Width           =   1e-4
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Buscar"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Ver Todos"
                  Object.Tag             =   "0"
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Imprimir"
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Salir"
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Style           =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   7740
         TabIndex        =   4
         Top             =   330
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lw2 
         Height          =   3645
         Left            =   210
         TabIndex        =   7
         Top             =   4530
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   6429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   2685
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   4736
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmTESGastosFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)



Private Const SaltoLinea = """ + chr(13) + """

Private Const IdPrograma = 807


Public Tipo As Integer
Public vSql As String
Public Opcion As Byte      ' 0.- Nueva remesa    1.- Modifcar remesa
                           ' 2.- Devolucion remesa
Public vRemesa As String   ' nºremesa|fecha remesa
Public ImporteRemesa As Currency

Public ValoresDevolucionRemesa As String
        'NOV 2009
        'antes: 4 campos     AHORA 5 campos
        'Concepto|ampliacion|
        'Concepto banco|ampliacion banco|
        'ahora+ Agrupa vtos

Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1

Private frmAux As frmTESGastosFijos2
Attribute frmAux.VB_VarHelpID = -1


Private frmMens3 As frmMensajes
Private frmMens2 As frmMensajes
Attribute frmMens2.VB_VarHelpID = -1
Private frmMens As frmMensajes

Dim SQL As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim i As Integer
Dim TotalReg As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Dim CampoOrden As String
Dim Orden As Boolean
Dim CampoOrden2 As String
Dim Orden2 As Boolean
Dim Modo As Byte

Dim Txt33Csb As String
Dim Txt41Csb As String

Dim Indice As Integer
Dim Codigo As Long

Dim SubTipo As Integer

Dim ModoInsertar As Boolean

Dim IndCodigo As Integer
Dim cadFiltro As String


Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub


Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function


Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub
    CargaList2
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub








Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        cboFiltro.ListIndex = vUsu.FiltroGastosFijos
        
        CargaList
        If lw1.ListItems.Count > 0 Then lw1.ListItems(1).EnsureVisible
        CargaList2
        PonleFoco lw1
    End If
    Screen.MousePointer = vbDefault
End Sub
    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmppal.Icon
    
    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With
    
    With Me.ToolbarAux
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 31
        .Buttons(7).Image = 37
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    CargaFiltros
    
    'Limpiamos el tag
    PrimeraVez = True
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    

    Me.lblIndicador.Caption = ""
    
    Orden = True
    CampoOrden = "gastosfijos.codigo"
    
    
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    vUsu.ActualizarFiltro "ariconta", IdPrograma, Me.cboFiltro.ListIndex
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub


Private Sub Image3_Click(Index As Integer)

    Select Case Index
        Case 1 ' cuenta contable
            Screen.MousePointer = vbHourglass
            
            Set frmCtas = New frmColCtas
            RC = Index
            frmCtas.DatosADevolverBusqueda = "0|1"
            frmCtas.ConfigurarBalances = 3
            frmCtas.Show vbModal
            Set frmCtas = Nothing
    
    End Select
End Sub



Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
    
    Select Case ColumnHeader
        Case "Código"
            CampoOrden = "gastosfijos.codigo"
        Case "Descripción"
            CampoOrden = "gastosfijos.descripcion"
        Case "Cta.Prevista"
            CampoOrden = "gastosfijos.ctaprevista"
        Case "Contrapartida"
            CampoOrden = "gastosfijos.contrapar"
    End Select
    CargaList


End Sub


Private Sub HacerToolBar(Boton As Integer)
Dim frmTESGtosList As frmTESGastosFijosList
    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
            BotonModificar
        Case 3
            BotonEliminar
        Case 5
'            BotonBuscar
        Case 6 ' ver todos
            CargaList
        Case 8
            'Imprimir factura
            Set frmTESGtosList = New frmTESGastosFijosList
            
            If Not lw1.SelectedItem Is Nothing Then frmTESGtosList.numero = lw1.SelectedItem.Text
            frmTESGtosList.Tipo = Me.cboFiltro.ListIndex
            
            frmTESGtosList.Show vbModal
            Set frmTESGtosList = Nothing
            
    End Select
End Sub

Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    
    SepuedeBorrar = False

    If lw1.SelectedItem.SubItems(8) = "Q" Then
        MsgBox "No se pueden modificar ni eliminar remesas en situación abonada.", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True

End Function


Private Sub lw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not lw1.SelectedItem Is Nothing Then
        Me.lblIndicador.Caption = lw1.SelectedItem.Index & " de " & Me.lw1.ListItems.Count
        CargaList2
    End If
End Sub

Private Sub lw1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub


Private Sub BotonAnyadir()
    
    Set frmAux = New frmTESGastosFijos2
    
    frmAux.Opcion = 1
    frmAux.Show vbModal
    
    Set frmAux = Nothing

    If CadenaDesdeOtroForm = "OK" Then CargaList

End Sub


Private Sub BotonEliminar()
    SQL = "¿Seguro que desea eliminar el gasto?"
    SQL = SQL & vbCrLf & "Código: " & lw1.SelectedItem.Text & "  " & lw1.SelectedItem.SubItems(1)
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        
        SQL = "DELETE FROM gastosfijos_recibos "
        SQL = SQL & " where codigo = " & lw1.SelectedItem.Text
        
        Conn.Execute SQL
        
        SQL = "DELETE FROM gastosfijos "
        SQL = SQL & " where codigo = " & lw1.SelectedItem.Text
        
        Conn.Execute SQL
        
        
        CargaList
    End If

End Sub


Private Sub BotonModificar()
Dim SQL As String
    
 
    If lw1.SelectedItem Is Nothing Then Exit Sub
    If lw1.SelectedItem = 0 Then Exit Sub
    
    CadenaDesdeOtroForm = ""
    

    Set frmAux = New frmTESGastosFijos2
    
    frmAux.Opcion = 2
    frmAux.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(2) & "|" & lw1.SelectedItem.SubItems(3) & "|" & lw1.SelectedItem.SubItems(4) & "|" & lw1.SelectedItem.SubItems(5) & "|"
    frmAux.Show vbModal
    
    Set frmAux = Nothing
    
    If CadenaDesdeOtroForm = "OK" Then CargaList
    
End Sub


Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    PonerIndicador lblIndicador, Modo
    
End Sub


Private Sub ToolbarAux_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea
        Case 2
            BotonModificarLinea
        Case 3
            BotonEliminarLinea
        Case 5
            BotonGeneraGastos
        Case 7
            BotonContabilizarLinea
    End Select
    If CadenaDesdeOtroForm = "OK" Then CargaList2
End Sub

Private Sub BotonAnyadirLinea()

    Set frmAux = New frmTESGastosFijos2
    
    frmAux.Opcion = 3
    frmAux.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|"
    frmAux.Show vbModal
    
    Set frmAux = Nothing
    

End Sub

Private Sub BotonModificarLinea()

    If lw2.SelectedItem Is Nothing Then Exit Sub
    
    Set frmAux = New frmTESGastosFijos2
    
    frmAux.Opcion = 4
    frmAux.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw2.SelectedItem.Text & "|" & lw2.SelectedItem.SubItems(1) & "|"
    frmAux.Show vbModal
    
    Set frmAux = Nothing

End Sub

Private Sub BotonEliminarLinea()

    If lw2.SelectedItem Is Nothing Then Exit Sub

    SQL = "¿Seguro que desea eliminar la línea de gasto?"
    SQL = SQL & vbCrLf & "Código: " & lw1.SelectedItem.Text & " de fecha " & lw2.SelectedItem.Text
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        SQL = "DELETE FROM gastosfijos_recibos "
        SQL = SQL & " where codigo = " & lw1.SelectedItem.Text & " and fecha = " & DBSet(lw2.SelectedItem.Text, "F")
        
        Conn.Execute SQL
        
        lw2.ListItems.Remove lw2.SelectedItem.Index
        
    End If

End Sub


Private Sub BotonContabilizarLinea()

    If lw2.SelectedItem Is Nothing Then Exit Sub
    
    If lw2.SelectedItem.SubItems(3) = 1 Then Exit Sub
    
    Set frmAux = New frmTESGastosFijos2
    
    frmAux.Opcion = 5
    frmAux.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|" & lw1.SelectedItem.SubItems(2) & "|" & lw1.SelectedItem.SubItems(4) & "|" & lw1.SelectedItem.SubItems(3) & "|" & lw1.SelectedItem.SubItems(5) & "|" & lw2.SelectedItem.Text & "|" & lw2.SelectedItem.SubItems(1) & "|"
    frmAux.Show vbModal
    
    Set frmAux = Nothing

End Sub

Private Sub BotonGeneraGastos()

    If lw1.SelectedItem Is Nothing Then Exit Sub
    
    Set frmAux = New frmTESGastosFijos2
    
    frmAux.Opcion = 6
    frmAux.Parametros = lw1.SelectedItem.Text & "|" & lw1.SelectedItem.SubItems(1) & "|"
    frmAux.Show vbModal
    
    Set frmAux = Nothing

End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (lw1.ListItems.Count > 0)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And (lw1.ListItems.Count > 0)
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N")
        
        ToolbarAux.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Not lw1.SelectedItem Is Nothing)
        ToolbarAux.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And lw2.ListItems.Count > 0
        ToolbarAux.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And lw2.ListItems.Count > 0
   
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


Private Sub CargaList()
Dim IT

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select gastosfijos.codigo, gastosfijos.descripcion, gastosfijos.ctaprevista,  gastosfijos.contrapar, cuentas1.nommacta, cuentas2.nommacta "
    Cad = Cad & " from gastosfijos left join cuentas cuentas2  on cuentas2.codmacta = gastosfijos.contrapar  ,cuentas cuentas1"
    Cad = Cad & " where cuentas1.codmacta = gastosfijos.ctaprevista  "
    
    If CampoOrden = "" Then CampoOrden = "gastosfijos.codigo " 'gastos fijos
    Cad = Cad & " ORDER BY " & CampoOrden ' remesas.anyo desc,
    If Orden Then Cad = Cad & " DESC"
    
    lw1.ColumnHeaders.Clear
    
    lw1.ColumnHeaders.Add , , "Código", 900
    lw1.ColumnHeaders.Add , , "Descripción", 3600
    lw1.ColumnHeaders.Add , , "Cta.Prevista", 1500
    lw1.ColumnHeaders.Add , , "Contrapartida", 1700
    lw1.ColumnHeaders.Add , , "Nombre", 0
    lw1.ColumnHeaders.Add , , "Nombre Contrapartida", 0
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = DBLet(miRsAux!Codigo, "N")
        IT.SubItems(1) = miRsAux!Descripcion
        IT.SubItems(2) = DBLet(miRsAux!Ctaprevista, "T")
        IT.ListSubItems(2).ToolTipText = DBLet(miRsAux.Fields(4), "T")
        IT.SubItems(3) = DBLet(miRsAux!contrapar, "T")
        IT.ListSubItems(3).ToolTipText = DBLet(miRsAux.Fields(5), "T")
        IT.SubItems(4) = DBLet(miRsAux.Fields(4), "T")
        IT.SubItems(5) = DBLet(miRsAux.Fields(5), "T")
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    If lw1.ListItems.Count > 0 Then
'        lw1.SelectedItem = lw1.ListItems(1)
'        lw1.SelectedItem.EnsureVisible

        PonerFocoLw lw1
        lw1_ItemClick lw1.SelectedItem
        
    End If
    CargaList2
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub


Private Sub CargaList2()
Dim IT

    lw2.ColumnHeaders.Clear
    
    lw2.ColumnHeaders.Add , , "Fecha", 1550
    lw2.ColumnHeaders.Add , , "Importe", 3200, 1
    lw2.ColumnHeaders.Add , , "Codigo", 0
    lw2.ColumnHeaders.Add , , "Contabilizado", 0
    
    lw2.ListItems.Clear

    If lw1.SelectedItem Is Nothing Then Exit Sub

    Set Me.lw2.SmallIcons = frmppal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select gastosfijos_recibos.codigo, gastosfijos_recibos.fecha, gastosfijos_recibos.importe, gastosfijos_recibos.contabilizado "
    Cad = Cad & " from gastosfijos_recibos "
    Cad = Cad & " where gastosfijos_recibos.codigo = " & lw1.SelectedItem.Text
    ' añadido el filtro
    
    CargarSqlFiltro
    
    
    If cadFiltro <> "" Then Cad = Cad & " and " & cadFiltro
    
    If CampoOrden2 = "" Then CampoOrden2 = "gastosfijos_recibos.fecha " 'gastos fijos
    Cad = Cad & " ORDER BY " & CampoOrden2 ' remesas.anyo desc,
    If Orden2 Then Cad = Cad & " DESC"
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw2.ListItems.Add()
        IT.Text = DBLet(miRsAux!Fecha, "F")
        IT.SubItems(1) = Format(miRsAux!Importe, "###,###,##0.00")
        IT.SubItems(2) = DBLet(miRsAux!Codigo, "N")
        IT.SubItems(3) = DBLet(miRsAux!Contabilizado, "N")
        
        If DBLet(miRsAux!Contabilizado, "N") = 1 Then IT.SmallIcon = 4
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    PonerFocoLw lw1
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub


Private Sub CargaFiltros()
Dim Aux As String
    

    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Ejercicios Abiertos "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Ejercicio Actual "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2
    cboFiltro.AddItem "Ejercicio Siguiente "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 3

End Sub
    
    
Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    cadFiltro = ""
    
    Select Case Me.cboFiltro.ListIndex
        Case 0 ' sin filtro
            cadFiltro = "(1=1)"
        
        Case 1 ' ejercicios abiertos
            cadFiltro = "gastosfijos_recibos.fecha >= " & DBSet(vParam.fechaini, "F")
        
        Case 2 ' ejercicio actual
            cadFiltro = "gastosfijos_recibos.fecha between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
        
        Case 3 ' ejercicio siguiente
            cadFiltro = "gastosfijos_recibos.fecha > " & DBSet(vParam.fechafin, "F")
    
    End Select
    
    Screen.MousePointer = vbDefault


End Sub
