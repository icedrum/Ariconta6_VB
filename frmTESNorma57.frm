VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESNorma57 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Norma 57 - Pagos por Ventanilla"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   Icon            =   "frmTESNorma57.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameNorma57Importar 
      Height          =   6615
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12015
      Begin VB.TextBox txtDescBanc 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   3900
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   6150
         Width           =   3765
      End
      Begin VB.TextBox txtCtaBanc 
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
         Index           =   5
         Left            =   2430
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   6150
         Width           =   1365
      End
      Begin VB.CommandButton cmdNoram57Fich 
         Height          =   375
         Left            =   8400
         Picture         =   "frmTESNorma57.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Leer"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdContabilizarNorma57 
         Caption         =   "Aceptar"
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
         Left            =   9210
         TabIndex        =   7
         Top             =   5970
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
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
         Height          =   375
         Index           =   42
         Left            =   10530
         TabIndex        =   5
         Top             =   5970
         Width           =   1095
      End
      Begin MSComctlLib.ListView lwNorma57Importar 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   3836
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serie"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº Fact"
            Object.Width           =   2718
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fec. fact."
            Object.Width           =   2735
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Orden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cliente"
            Object.Width           =   6086
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fec Cobro"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView lwNorma57Importar 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   3836
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº Fact"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Motivo"
            Object.Width           =   7832
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   11220
         TabIndex        =   17
         Top             =   210
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
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   5
         Left            =   2070
         Top             =   6150
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   78
         Left            =   240
         TabIndex        =   12
         Top             =   6120
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Leer fichero bancario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5220
         TabIndex        =   9
         Top             =   330
         Width           =   3000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimientos erroneos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   77
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimientos encontrados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Index           =   76
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Importar fichero norma 57"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   23
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4890
      End
   End
   Begin VB.Frame FrameProgreso 
      Height          =   1935
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   4095
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label lblPPAL 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lbl2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmTESNorma57"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Private Const IdPrograma = 613


Public Opcion As Byte
    '42.-   IMportar fichero norma 57 (recibos al cobro en ventanilla)
    
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1

Dim Sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim cad As String
Dim CONT As Long
Dim i As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

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


Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdContabilizarNorma57_Click()
Dim frmTESRealCob As frmTESRealizarCobros


    Sql = ""
    If Me.lwNorma57Importar(0).ListItems.Count = 0 Then Sql = Sql & "-Ningun vencimiento desde el fichero" & vbCrLf
    If Me.txtCtaBanc(5).Text = "" Then Sql = Sql & "-Cuenta bancaria" & vbCrLf
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    
    'La madre de las batallas
    'El sql que mando
    Sql = "(numserie ,numfactu,fecfactu,numorden ) IN (select ccost,pos,nomdocum,numdiari from tmpconext "
    Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " and numasien=0 ) "
    'CUIDADO. El trozo 'from tmpconext  WHERE codusu' tiene que estar extamente ASI
    '  ya que en ver cobros, si encuentro esto, pong la fecha de vencimiento la del PAGO por
    ' ventanilla que devuelve el banco y contabilizamos en funcion de esa fecha
            
            
    cad = Format(Now, "dd/mm/yyyy") & "|" & Me.txtCtaBanc(5).Text & "|" & Me.txtDescBanc(5).Text & "|0|"  'efectivo

    Set frmTESRealCob = New frmTESRealizarCobros

    With frmTESRealCob
        .ImporteGastosTarjeta_ = 0
        '--.OrdenacionEfectos = 3
        .vSql2 = Sql
        .OrdenarEfecto = True
        .Regresar = False
        .ContabTransfer = False
        .Cobros = True
        '.Tipo = 0
        .SegundoParametro = ""
        'Los textos
        .vTextos = cad
        .CodmactaUnica = ""
        .VieneDesdeNorma57 = True
        .Show vbModal
    End With
        
    Set frmTESRealCob = Nothing
        
    'Borro haya cancelado o no
    LimpiarDelProceso
End Sub

Private Sub cmdNoram57Fich_Click()

    If Me.lwNorma57Importar(0).ListItems.Count > 0 Or lwNorma57Importar(1).ListItems.Count > 0 Then
        Sql = "Ya hay un proceso . ¿ Desea importar otro archivo?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    Me.cmdContabilizarNorma57.Visible = False
    
    cd1.FileName = ""
    cd1.ShowOpen
    If cd1.FileName = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    LimpiarDelProceso
    Me.Refresh
   
    If procesarficheronorma57 Then
        
        'El fichero que ha entrado es correcto.
        'Ahora vamos a buscar los vencimientos
        If BuscarVtosNorma57 Then
            
            'AHORA cargamos los listviews
            CargaLWNorma57 True   'los correctos 'Si es que hay
            
            'Los errores
            CargaLWNorma57 False
    
    
    
            Me.cmdContabilizarNorma57.Visible = Me.lwNorma57Importar(0).ListItems.Count > 0
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmppal.Icon
    CargaImagenesAyudas Me.imgCtaBanc, 1, "Cuenta contable bancaria"
    
    
    
    
    'Limpiamos el tag
    PrimeraVez = True
    FrameNorma57Importar.Visible = False
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    
    
    
    Select Case Opcion
    Case 42
        H = FrameNorma57Importar.Height + 120
        W = FrameNorma57Importar.Width
        FrameNorma57Importar.Visible = True
    
    
    End Select
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    i = Opcion
    If Opcion = 13 Or i = 43 Or i = 44 Then i = 11
    
    
    PonerFrameProgreso

End Sub


Private Sub PonerFrameProgreso()
Dim i As Integer

    'Ponemos el frame al pricnipio de todo
    FrameProgreso.Visible = False
    FrameProgreso.ZOrder 0
    
    'lo ubicamos
    'Posicion horizintal WIDTH
    i = Me.Width - FrameProgreso.Width
    If i > 100 Then
        i = i \ 2
    Else
        i = 0
    End If
    FrameProgreso.Left = i
    'Posicion  VERTICAL HEIGHT
    i = Me.Height - FrameProgreso.Height
    If i > 100 Then
        i = i \ 2
    Else
        i = 0
    End If
    FrameProgreso.top = i
End Sub








Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    Sql = CadenaSeleccion
End Sub



Private Sub imgCtaBanc_Click(Index As Integer)
    Sql = ""
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
    If Sql <> "" Then
        txtCtaBanc(Index).Text = RecuperaValor(Sql, 1)
        Me.txtDescBanc(Index).Text = RecuperaValor(Sql, 2)
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtCtaBanc_GotFocus(Index As Integer)
    PonFoco txtCtaBanc(Index)
End Sub

Private Sub txtCtaBanc_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaBanc_LostFocus(Index As Integer)
    txtCtaBanc(Index).Text = Trim(txtCtaBanc(Index).Text)
    If txtCtaBanc(Index).Text = "" Then
        txtDescBanc(Index).Text = ""
        Exit Sub
    End If
    
    cad = txtCtaBanc(Index).Text
    i = CuentaCorrectaUltimoNivelSIN(cad, Sql)
    If i = 0 Then
        MsgBox "NO existe la cuenta: " & txtCtaBanc(Index).Text, vbExclamation
        Sql = ""
        cad = ""
    Else
        cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", cad, "T")
        If cad = "" Then
            MsgBox "Cuenta no asoaciada a ningun banco", vbExclamation
            Sql = ""
            i = 0
        End If
    End If
    
    txtCtaBanc(Index).Text = cad
    Me.txtDescBanc(Index).Text = Sql
    If i = 0 Then PonFoco txtCtaBanc(Index)
    
End Sub




Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


'Si tiene valor el campo fecha, entonces lo ponemos con el BD
Private Function CampoABD(ByRef T As TextBox, Tipo As String, CampoEnLaBD, Mayor_o_Igual As Boolean) As String

    CampoABD = ""
    If T.Text <> "" Then
        If Mayor_o_Igual Then
            CampoABD = " >= "
        Else
            CampoABD = " <= "
        End If
        Select Case Tipo
        Case "F"
            CampoABD = CampoEnLaBD & CampoABD & "'" & Format(T.Text, FormatoFecha) & "'"
        Case "T"
            CampoABD = CampoEnLaBD & CampoABD & "'" & T.Text & "'"
        Case "N"
            CampoABD = CampoEnLaBD & CampoABD & T.Text
        End Select
    End If
End Function



Private Function CampoBD_A_SQL(ByRef C As ADODB.Field, Tipo As String, Nulo As Boolean) As String

    If IsNull(C) Then
        If Nulo Then
            CampoBD_A_SQL = "NULL"
        Else
            If Tipo = "T" Then
                CampoBD_A_SQL = "''"
            Else
                CampoBD_A_SQL = "0"
            End If
        End If

    Else
    
        Select Case Tipo
        Case "F"
            CampoBD_A_SQL = "'" & Format(C.Value, FormatoFecha) & "'"
        Case "T"
            CampoBD_A_SQL = "'" & DevNombreSQL(C.Value) & "'"
        Case "N"
            CampoBD_A_SQL = TransformaComasPuntos(CStr(C.Value))
        End Select
    End If
End Function



Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.lblPPAL.Caption = TEXTO
        Me.lbl2.Caption = ""
        Me.ProgressBar1.Value = 0
        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub






'****************************************************************************************
'****************************************************************************************
'
'       NORMA 57
'
'****************************************************************************************
'****************************************************************************************
Private Function procesarficheronorma57() As Boolean
Dim Estado As Byte  '0  esperando cabcerea
                    '1  esperando pie (leyendo lineas)
    
    On Error GoTo eprocesarficheronorma57
    
    
    'insert into tmpconext(codusu,cta,fechaent,Pos)
    Conn.Execute "DELETE FROM tmpconext WHERE codusu = " & vUsu.Codigo
    procesarficheronorma57 = False
    i = FreeFile
    Open cd1.FileName For Input As #i
    Sql = ""
    Estado = 0
    Importe = 0
    TotalRegistros = 0
    While Not EOF(i)
            Line Input #i, Sql
            RC = Mid(Sql, 1, 4)
            Select Case Estado
            Case 0
                'Para saber que el fichero tiene el formato correcto
                If RC = "0270" Then
                        Estado = 1
                        'Voy a buscar si hay un banco
                        
                        RC = "select cuentas.codmacta,nommacta from bancos,cuentas where bancos.codmacta="
                        RC = RC & "cuentas.codmacta AND mid(bancos.iban,5,4) = " & Trim(Mid(Sql, 23, 4))
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        TotalRegistros = 0
                        While Not miRsAux.EOF
                            RC = miRsAux!codmacta & "|" & miRsAux!Nommacta & "|"
                            TotalRegistros = TotalRegistros + 1
                            miRsAux.MoveNext
                        Wend
                        miRsAux.Close
                        If TotalRegistros = 1 Then
                            Me.txtCtaBanc(5).Text = RecuperaValor(RC, 1)
                            Me.txtDescBanc(5).Text = RecuperaValor(RC, 2)
                        End If
                        TotalRegistros = 0
                End If
            Case 1
                If RC = "6070" Then
                    'Linea con recibo
                    'Ejemplo:
                    '   6070      46076147000130582263151014000000014067003059                      0000000516142
                    '                                  fecha       impot   socio                      fra      CC codigo de control del codigo de barra
                    'Fecha pago
                    RC = Mid(Sql, 31, 2) & "/" & Mid(Sql, 33, 2) & "/20" & Mid(Sql, 35, 2)
                    Fecha = CDate(RC)
                    'IMporte
                    RC = Mid(Sql, 37, 12)
                    cad = CStr(CCur(Val(RC) / 100))
                    'FRA
                    RC = Mid(Sql, 77, 11)
                    CONT = Val(RC)
                    'Socio
                    RC = Val(Mid(Sql, 50, 6))
                        
                    'Insertamos en tmp
                    TotalRegistros = TotalRegistros + 1
                    Sql = "INSERT INTO tmpconext(codusu,cta,fechaent,Pos,TimporteD,linliapu) VALUES (" & vUsu.Codigo & ",'"
                    Sql = Sql & RC & "','" & Format(Fecha, FormatoFecha) & "'," & CONT & "," & TransformaComasPuntos(cad) & "," & TotalRegistros & ")"
                    Conn.Execute Sql
                    
                    Importe = Importe + CCur(TransformaPuntosComas(cad))
                ElseIf RC = "8070" Then
                    'OK. Final de linea.
                    '
                    'Comprobacion BASICA
                    '8070      46076147000 000010        000000028440
                    '                       vtos-2           importe
                    
                    RC = ""
                    
                    'numero registros
                    cad = Val(Mid(Sql, 24, 5))
                    If Val(cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. Nº registros cero. " & Sql
                    Else
                        If Val(cad) - 2 <> TotalRegistros Then RC = "Contador de registros incorrecto"
                    End If
                    'Suma importes
                    cad = CStr(CCur(Mid(Sql, 37, 12) / 100))
                    
                    If CCur(cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. Suma importes cero. " & Sql
                    Else
                        If CCur(cad) <> Importe Then RC = RC & vbCrLf & "Suma importes incorrecta"
                    End If
                    
                    If RC <> "" Then
                        Err.Raise 513, , RC
                    Else
                        Estado = 2
                    End If
                End If
            End Select
    Wend
    Close #i
    i = 0 'para que no vuelva a cerrar el fichero
    
    If Estado < 2 Then
        'Errores procesando fichero
        If Estado = 0 Then
            Sql = "No se encuetra la linea de inicio de declarante(6070)"
        Else
            Sql = "No se encuetra la linea de totales(8070)"
        End If

        MsgBox "Error procesando el fichero." & vbCrLf & Sql, vbExclamation
    Else
        espera 0.5
        procesarficheronorma57 = True
    End If
eprocesarficheronorma57:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    If i > 0 Then Close #i
End Function


Private Function BuscarVtosNorma57() As Boolean

    BuscarVtosNorma57 = False
    
    Set miRsAux = New ADODB.Recordset
    
    'Dependiendo del parametro....
    If vParamT.Norma57 = 1 Then
        'ESCALONA.
        'Viene el socio y el numero de factura e importe.
        'Habra que buscar
        BuscarVtosNorma57 = VtosNorma57Escalona

    Else
        MsgBox "En desarrollo", vbExclamation
    End If
    
    Set miRsAux = Nothing
End Function

Private Function VtosNorma57Escalona() As Boolean
Dim RN As ADODB.Recordset
Dim Fin As Boolean
Dim NoEncontrado As Byte
Dim AlgunVtoNoEncontrado As Boolean
On Error GoTo eVtosNorma57Escalona
    
    VtosNorma57Escalona = False
    Set RN = New ADODB.Recordset
    Sql = "select * from tmpconext WHERE codusu =" & vUsu.Codigo & " order by cta,pos "
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AlgunVtoNoEncontrado = False
    While Not miRsAux.EOF
        'Vto a vto
        '
        RC = RellenaCodigoCuenta("430." & miRsAux!Cta)
        Sql = "Select * from cobros where codmacta = '" & RC & "' AND numfactu =" & miRsAux!Pos & " and impvenci>0"
        RN.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CONT = 0
        If RN.EOF Then
            cad = "NO encontrado"
            NoEncontrado = 2
        Else
            'OK encontrado.
            Fin = False
            i = 0
            NoEncontrado = 1
            cad = ""
            
            While Not Fin
            
                i = i + 1
                
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                If Not Fin Then
                    RN.MoveNext
                    If RN.EOF Then Fin = True
                End If
            Wend
        End If
        RN.Close
        Sql = "UPDATE tmpconext SET "
        If CONT = 1 Then
            'OK este es el vto
            'NO hacemos nada. Updateamos los campos de la tmp
            'para buscar despues
            'numdiari numorden       numdocum=fecfaccl     ccost numserie
            Sql = Sql & " nomdocum ='" & Format(Fecha, FormatoFecha)
            Sql = Sql & "', ccost ='" & DevfrmCCtas
            Sql = Sql & "', numdiari = " & i
            Sql = Sql & ", contra = '" & RC & "'"
        Else
            If i > 1 Then cad = "(+1) " & cad
            Sql = Sql & " numasien=  " & NoEncontrado  'para vtos no encontrados o erroneos
            Sql = Sql & ", ampconce ='" & DevNombreSQL(cad) & "'"
            If NoEncontrado = 2 Then AlgunVtoNoEncontrado = True
        End If
        Sql = Sql & " WHERE codusu =" & vUsu.Codigo & " AND linliapu = " & miRsAux!Linliapu
        Conn.Execute Sql
            
 
        
        'Sig
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    
    
    
    If AlgunVtoNoEncontrado Then
        'Lo buscamos al reves
        espera 0.5
        Sql = "select * from  tmpconext  WHERE codusu =" & vUsu.Codigo & " AND numasien=2"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Miguel angel
            'Puede que en algunos recibos las posciones del fichero vengan cambiadas
            'Donde era la factura es la cta y al reves
            RC = RellenaCodigoCuenta("430." & miRsAux!Pos)
            Sql = "Select * from cobros where codmacta = '" & RC & "' AND numfactu =" & Val(miRsAux!Cta) & " and impvenci>0"
            RN.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RN.EOF Then
        
                'OK encontrado.
                Fin = False
                CONT = 0
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                
            
            
                'OK este es el vto
                'NO hacemos nada. Updateamos los campos de la tmp
                'para buscar despues
                'numdiari numorden       numdocum=fecfaccl     ccost numserie
                If CONT = 1 Then
                    Sql = Sql & " nomdocum ='" & Format(Fecha, FormatoFecha)
                    Sql = Sql & "', ccost ='" & DevfrmCCtas
                    Sql = Sql & "', numdiari = " & i
                    Sql = Sql & ", contra = '" & RC & "'"
                    Sql = "UPDATE tmpconext SET "
                End If
            End If
            RN.Close
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    VtosNorma57Escalona = True
eVtosNorma57Escalona:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, "Buscando vtos Escalona"
    Set RN = Nothing
End Function


Private Sub Norma57VencimientoEncontradoEsCorrecto(ByRef Rss As ADODB.Recordset, ByRef Final As Boolean)
        
        'Ha encontrado el vencimiento. Falta ver si no esta en remesa....
        If Not IsNull(Rss!CodRem) Then
            cad = "En la remesa " & Rss!CodRem
        
        Else
            If Not IsNull(Rss!transfer) Then
                cad = "Transferencia " & Rss!transfer
            Else
                Importe = Rss!ImpVenci + DBLet(Rss!Gastos, "N") - DBLet(Rss!impcobro, "N")
                If Importe <> miRsAux!timported Then
                    'Importe distinto
                    'Veamos si es que esta
                    cad = "Importe distinto"
                Else
                    'OK. Misma factura, socio, importe. SAlimos ya poniendo ""
                    Fecha = Rss!FecFactu
                    DevfrmCCtas = Rss!NUmSerie
                    i = Rss!numorden
                    cad = ""
                    Final = True
                    CONT = 1
                End If
            End If
        End If
End Sub

Private Sub CargaLWNorma57(Correctos As Boolean)
Dim IT As ListItem

    Set miRsAux = New ADODB.Recordset
    If Correctos Then
        Sql = "select tmpconext.*,nommacta from tmpconext left join cuentas on tmpconext.contra=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        Sql = Sql & " and numasien=0 order by  ccost,pos  "
    Else
        Sql = "select * from tmpconext WHERE codusu = " & vUsu.Codigo & " and numasien > 0 order by cta,pos "
    End If
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Correctos Then
            Set IT = Me.lwNorma57Importar(0).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!CCost
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!nomdocum, "dd/mm/yyyy")
            IT.SubItems(3) = miRsAux!Linliapu
            If IsNull(miRsAux!Nommacta) Then
                Sql = "ERRROR GRAVE"
            Else
                Sql = miRsAux!Nommacta
            End If
            IT.SubItems(4) = Sql
            IT.SubItems(5) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(6) = Format(miRsAux!FechaEnt, "dd/mm/yyyy")
            IT.Checked = True
        Else
            'ERRORES
            Set IT = Me.lwNorma57Importar(1).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!Cta
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(3) = miRsAux!Ampconce
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub


Private Sub LimpiarDelProceso()
    lwNorma57Importar(0).ListItems.Clear
    lwNorma57Importar(1).ListItems.Clear
    Me.txtCtaBanc(5).Text = ""
    Me.txtDescBanc(5).Text = ""
End Sub

