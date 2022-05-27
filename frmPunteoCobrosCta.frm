VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPunteoCobrosCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobros // pagos"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPunteoCobrosCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   2510
      TabIndex        =   12
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   150
         TabIndex        =   13
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Serie"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Factura"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha.F"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha Vto."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Nº"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Importe"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Gastos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Pendiente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "cta"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
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
      Index           =   1
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1260
      Width           =   5625
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Tag             =   "Cta. cliente|T|N|||cobros|codmacta|||"
      Text            =   "Text1"
      Top             =   1260
      Width           =   1350
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   2385
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3750
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Dividir"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Añadir de otra cuenta"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar apunte  Ctrl + G"
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   6390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   6390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   6390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   2865
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
         Top             =   120
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image imgCuentas 
      Height          =   240
      Left            =   1080
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cuenta "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   11040
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   855
      Index           =   1
      Left            =   8160
      TabIndex        =   15
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Cobros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   855
      Index           =   0
      Left            =   3720
      TabIndex        =   14
      Top             =   6240
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   780
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmPunteoCobrosCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public ImporteBanco As Currency

Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Dim PrimeraVez As Boolean
Private CadB As String

'Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim Importe As Currency

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas BUSCAR
'   Modo 2 -> Recorrer registros
'   Modo 3 -> Lineas  INSERTAR
'   Modo 4 -> Lineas MODIFICAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Integer

    Modo = vModo

    B = (Modo = 2) Or Modo = 0
    
    Me.lblIndicador.Caption = ""
    
    'Prueba
    BloqueaTXT Text1(0), B
    BloqueaTXT Text1(1), B
    
    
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub




Private Sub BotonBuscar()
    'Buscar
    ListView1.ListItems.Clear
    PonerModo 1
    DespalzamientoVisible False
    lblIndicador.Caption = "Buscar"
    Text1(0).Text = ""
    Text1(1).Text = ""
    Label1(0).visible = False
    PonFoco Text1(1)
End Sub



Private Sub cmdAceptar_Click()
    
    Text1(0).Text = Trim(Text1(0).Text)
    Text1(1).Text = Trim(Text1(1).Text)
    If Text1(0).Text = "" And Text1(1).Text = "" Then Exit Sub
    
    
    'Hago busqueda
    CadB = ""
    If Text1(0).Text <> "" Then CadB = CadB & " AND cobros.codmacta like " & DBSet(Text1(0).Text, "T")
    If Text1(1).Text <> "" Then
        CadB = CadB & " AND (cobros.nomclien like " & DBSet("%" & Text1(1).Text & "%", "T")
        CadB = CadB & " OR (cobros.nomclien is null and cuentas.codmacta like " & DBSet("%" & Text1(1).Text & "%", "T") & "))"
    End If
    HacerBusqueda2 False
    
    
End Sub

Private Sub cmdCancelar_Click()
    PonerModo 0
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        BotonBuscar
    End If

    Screen.MousePointer = vbDefault
End Sub
'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++
Private Sub Form_Load()

    Me.Icon = frmppal.Icon

    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 1
        .Buttons(3).Image = 11 '25
        .Buttons(4).Image = 13
        .Buttons(6).Image = 25
        '.Buttons(8).Image = 16
    End With

' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    '## A mano
    
    imgCuentas.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    
    cmdRegresar.visible = (DatosADevolverBusqueda <> "")
    Label2(1).Caption = Format(ImporteBanco, FormatoImporte)
    PonerModo 0
    DespalzamientoVisible False
    PrimeraVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub




Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    CadB = CadenaSeleccion
End Sub

Private Sub imgCuentas_Click()
    If Modo <> 1 Then Exit Sub
     If itemsSeleccionados(True) Then Exit Sub
     
     CadB = ""
     AbrirFormCuenta
     If CadB <> "" Then
        Text1(0).Text = RecuperaValor(CadB, 1)
        cmdAceptar_Click
    End If
End Sub


Private Sub AbrirFormCuenta()

        Set frmCCtas = New frmColCtas
        frmCCtas.FILTRO = "1"
        frmCCtas.DatosADevolverBusqueda = "0"
        frmCCtas.Show vbModal
        Set frmCCtas = Nothing
        
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Importe = 0
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then Importe = Importe + ImporteFormateado(ListView1.ListItems(I).SubItems(7))
    Next
    Label1(0).visible = Importe <> 0
    Label1(0).Caption = Format(Importe, FormatoImporte)
    
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyG Then Hacertool 6
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
Unload Me
End Sub

'----------------------------------------------------------------

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Hacertool Button.Index
    
End Sub

Private Sub Hacertool(Indice As Integer)

Dim Aux As String
    Select Case Indice
    Case 1
        BotonBuscar
    Case 3
        DividirVencimiento
    Case 4
        If Modo = 2 Then
            
            
            
            CadB = ""
            AbrirFormCuenta
            If CadB <> "" Then
                            
                'Primero compruebo que no sea la misma
                Aux = RecuperaValor(CadB, 1)
                For J = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(J).SubItems(8) = Aux Then Exit For
                Next
                If J <= ListView1.ListItems.Count Then
                    MsgBoxA "Esta cuenta ya ha sido agregada anteriormente", vbExclamation
                    Exit Sub
                End If
            
                'CadB = "Select NUmSerie ,numfactu, FecFactu,FecVenci, numorden, ImpVenci ,gastos ,impcobro"
                Aux = RecuperaValor(CadB, 1)
                Aux = " FROM cobros WHERE codmacta =   " & DBSet(Aux, "T") & " AND situacion=0 AND codrem is null  group BY codmacta"
                Aux = "Select count(*) cuantos, sum(ImpVenci + coalesce(gastos,0) - coalesce(impcobro,0)) deuda  " & Aux
                
                Set miRsAux = New ADODB.Recordset
                miRsAux.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Aux = ""
                If Not miRsAux.EOF Then
                    If DBLet(miRsAux!Cuantos, "N") > 0 Then
                        Aux = RecuperaValor(CadB, 2) & vbCrLf
                        Aux = Aux & "Vtos:   " & Format(miRsAux!Cuantos, "0") & vbCrLf
                        Aux = Aux & "Deuda:   " & Format(miRsAux!deuda, FormatoImporte) & vbCrLf
                        
                        Aux = Aux & vbCrLf & "¿Continuar?"
                        If MsgBoxA(Aux, vbQuestion + vbYesNoCancel) <> vbYes Then Aux = ""
                        
                    End If
                End If
                miRsAux.Close
                Set miRsAux = Nothing
                If Aux = "" Then Exit Sub
                Text1(0).Text = RecuperaValor(CadB, 1)
                CadB = " AND cobros.codmacta = " & DBSet(Text1(0).Text, "T")
                HacerBusqueda2 True
            End If
        
        End If
    Case 6
        If PreparaDevolverVtos Then
            CadB = Mid(CadB, 2)
            CadenaDesdeOtroForm = " (numserie , numfactu, fecfactu, numorden) IN (" & CadB & ")"
            Unload Me
        Else
            CadenaDesdeOtroForm = ""
        End If
    End Select
        
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next
    'IdPrograma = 601  Ver cobros
    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(601, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!CrearEliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!CrearEliminar, "N") And Me.ListView1.ListItems.Count > 0
        Toolbar1.Buttons(4).Enabled = Toolbar1.Buttons(3).Enabled
        
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(3).Enabled
        
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub





Private Sub HacerBusqueda2(Anyade As Boolean)

    If CadB <> "" Then
        CadB = " situacion = 0  AND codrem is null   " & CadB
        CadB = " INNER JOIN cuentas on cuentas.codmacta=cobros.codmacta WHERE " & CadB
        CadB = "select cobros.codmacta, cobros.nomclien,cuentas.nommacta from cobros " & CadB & " group  BY cobros.codmacta"
        
        
        PonerCadenaBusqueda Anyade
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonFoco Text1(0)
        ' **********************************************************************
    End If
    

    

End Sub

Private Sub PonerCadenaBusqueda(Anyade As Boolean)
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq
    
    adodc1.ConnectionString = Conn.ConnectionString
    adodc1.RecordSource = CadB
    adodc1.Refresh
    If adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro   ", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    
        Else
            PonerModo 2
            adodc1.Recordset.MoveFirst
            PonerCampos Anyade
            PonerModoUsuarioGnral Modo, "ariconta"
    End If
    

Screen.MousePointer = vbDefault
Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos(AnyadirCta As Boolean)
Dim ItmX

    
    Text1(0).Text = adodc1.Recordset!codmacta
    Text1(1).Text = DBLet(adodc1.Recordset!nomclien, "T")
    If Text1(1).Text = "" Then Text1(1).Text = DBLet(adodc1.Recordset!Nommacta, "T")


    CadB = "Select NUmSerie ,numfactu, FecFactu,FecVenci, numorden, ImpVenci ,gastos ,impcobro"
    CadB = CadB & " FROM cobros WHERE codmacta =   " & DBSet(adodc1.Recordset!codmacta, "T") & " AND situacion=0 AND codrem is null  ORDER BY fecfactu,numserie,numfactu"
    Set miRsAux = New ADODB.Recordset
    If Not AnyadirCta Then
        ListView1.ListItems.Clear
        Me.ListView1.ColumnHeaders(8).Width = 0
    Else
        Me.ListView1.ColumnHeaders(8).Width = 1400
    End If
    miRsAux.Open CadB, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set ItmX = ListView1.ListItems.Add()
        
        ItmX.Text = miRsAux!NUmSerie
        ItmX.SubItems(1) = miRsAux!numfactu
        ItmX.SubItems(2) = Format(miRsAux!Fecfactu, "dd/mm/yyyy")
        ItmX.SubItems(3) = Format(miRsAux!FecVenci, "dd/mm/yyyy")
        ItmX.SubItems(4) = miRsAux!numorden
        
        ItmX.SubItems(5) = Format(miRsAux!ImpVenci, FormatoImporte)
        'Gastos
        Importe = DBLet(miRsAux!Gastos, "N")
        ItmX.SubItems(6) = Format(Importe, FormatoImporte)
        Importe = Importe + miRsAux!ImpVenci
    
        If DBLet(miRsAux!impcobro, "N") <> 0 Then Importe = Importe - miRsAux!impcobro
        ItmX.SubItems(7) = Format(Importe, FormatoImporte)
        ItmX.SubItems(8) = adodc1.Recordset!codmacta
        ItmX.ToolTipText = Text1(1).Text
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing



    DespalzamientoVisible adodc1.Recordset.RecordCount > 1
    lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount



End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
Private Sub Desplazamiento(Index As Integer)
    Select Case Index
        Case 1
            
            adodc1.Recordset.MoveFirst
        Case 2
            adodc1.Recordset.MovePrevious
            If adodc1.Recordset.BOF Then adodc1.Recordset.MoveFirst
        Case 3
            adodc1.Recordset.MoveNext
            If adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
        Case 4
            adodc1.Recordset.MoveLast
    End Select
    PonerCampos False
    lblIndicador.Caption = adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
End Sub


Private Function itemsSeleccionados(HazPregunta As Boolean) As Boolean
    
    itemsSeleccionados = False
    For J = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(J).Checked Then Exit For
            
    Next
    
    If J <= ListView1.ListItems.Count Then
        If MsgBox("Perderá la seleccion.  ¿Continuar?", vbQuestion + vbYesNo) = vbYes Then Exit Function
        itemsSeleccionados = True
    End If
    
End Function




Private Sub DividirVencimiento()
    
    If Modo <> 2 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    If Me.ListView1.ColumnHeaders(8).Width <> 0 Then
        MsgBoxA "Está realizando cobro sobre varias cuentas", vbExclamation
        Exit Sub
    End If
    
           'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
    Importe = ImporteFormateado(ListView1.SelectedItem.SubItems(7))
    CadenaDesdeOtroForm = "numserie = '" & ListView1.SelectedItem.Text & "' AND numfactu = " & ListView1.SelectedItem.SubItems(1)
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND fecfactu = '" & Format(ListView1.SelectedItem.SubItems(2), FormatoFecha) & "'|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView1.SelectedItem.SubItems(4) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CStr(Importe) & "|"
    
    
    
    frmTESCobrosDivVto.Opcion = 27
    frmTESCobrosDivVto.Label4(56).Caption = Text1(1).Text
    frmTESCobrosDivVto.txtCodigo(2).Text = ListView1.SelectedItem.SubItems(3)
    frmTESCobrosDivVto.Label4(57).Caption = ListView1.SelectedItem.Text & ListView1.SelectedItem.SubItems(1) & " / " & ListView1.SelectedItem.SubItems(4) & "      de " & ListView1.SelectedItem.SubItems(2)
    frmTESCobrosDivVto.txtCodigo(1).Text = ListView1.SelectedItem.SubItems(7)

    frmTESCobrosDivVto.Show vbModal
    
    If CadenaDesdeOtroForm <> "" Then
        CadB = " AND cobros.codmacta = " & DBSet(Text1(0).Text, "T")
        HacerBusqueda2 False
    End If
    CadenaDesdeOtroForm = ""
End Sub


Private Function PreparaDevolverVtos() As Boolean
    PreparaDevolverVtos = False
    
    Importe = 0
    CadB = ""
    For I = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then Importe = Importe + ImporteFormateado(ListView1.ListItems(I).SubItems(7)): CadB = CadB & "X"
    Next
    
    If CadB = "" Then
        CadB = "Seleccione algun vencimiento"
    Else
        If ImporteBanco <> Importe Then
            CadB = "Apunte banco: " & ImporteBanco & vbCrLf & "Seleccionado: " & Importe
        Else
            CadB = ""
        End If
    End If
    
    If CadB <> "" Then
        MsgBoxA CadB, vbExclamation
        Exit Function
    End If
    
    
    
    If MsgBox("¿Desea realizar el cobro?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    CadB = ""
    For I = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            With ListView1.ListItems(I)
                'numserie numfactu fecfactu numorden
                CadB = CadB & ", ('" & .Text & "' , " & .SubItems(1) & ", '" & Format(.SubItems(2), FormatoFecha) & "'," & .SubItems(4) & ")"
            End With
        End If
    Next
    
    PreparaDevolverVtos = True
    
End Function
