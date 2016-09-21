VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmColBalan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurador de Balances"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmColBalan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   4590
      TabIndex        =   8
      Top             =   90
      Width           =   2865
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
         ItemData        =   "frmColBalan.frx":000C
         Left            =   120
         List            =   "frmColBalan.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   2625
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   2490
      TabIndex        =   5
      Top             =   90
      Width           =   1275
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   1410
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Comprobar"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copiar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   2205
      Begin VB.CheckBox Check2 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   3
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Width           =   1995
         _ExtentX        =   3519
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Ver todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6060
      Top             =   4080
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColBalan.frx":0050
      Height          =   3885
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   6853
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   7560
      TabIndex        =   10
      Top             =   120
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
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   4950
      Width           =   7755
   End
End
Attribute VB_Name = "frmColBalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean

Private Const IdPrograma = 108



Private Sub cboFiltro_Click()
    CargaGrid
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.cboFiltro.ListIndex = 0
        CargaGrid
    
        PonerModoUsuarioGnral 0, "ariconta"
    
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True
      
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With
    
    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 37 'comprobar
        .Buttons(2).Image = 35 'copiar
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    CargaCombo
    
   
End Sub


Private Sub CargaCombo()
Dim Aux As String
    
    cboFiltro.Clear

    'tipos de balance
    cboFiltro.AddItem "Todos"
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Perdidas y Ganancias"
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Situación"
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2
    
End Sub



Private Sub CargaGrid()
Dim sql As String
    
    
    sql = ""
    If cboFiltro.ListIndex = 1 Then
        sql = " WHERE Perdidas = 1"
    Else
        If cboFiltro.ListIndex = 2 Then
            sql = " WHERE Perdidas = 0"
        End If
    End If
    sql = sql & " ORDER BY Numbalan"
    sql = "select numbalan,nombalan, if(perdidas=1,'SI','NO') as Perd ,if(predeterminado=1,'*','') as Pre  from balances " & sql
    Adodc1.RecordSource = sql
    Adodc1.ConnectionString = Conn
    Adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    DataGrid1.Columns(0).Caption = "Número"
    DataGrid1.Columns(0).Width = 1000
    
    DataGrid1.Columns(1).Caption = "Nombre"
    DataGrid1.Columns(1).Width = 4750

    DataGrid1.Columns(2).Caption = "P y G"
    DataGrid1.Columns(2).Width = 700
    
    DataGrid1.Columns(3).Caption = "Pred."
    DataGrid1.Columns(3).Width = 700
    
End Sub
    

Private Function ObtenerSiguiente() As Long

    ObtenerSiguiente = 0
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select max(numbalan) from balances", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        ObtenerSiguiente = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    ObtenerSiguiente = ObtenerSiguiente + 1
End Function



Private Sub EliminarBalance()
Dim sql As String
    sql = "Seguro que desea eliminar el balance: " & Adodc1.Recordset!NomBalan & "?"
    If MsgBox(sql, vbExclamation + vbYesNo) <> vbYes Then Exit Sub
    
    'Eliminamos las cuentas
    sql = "DELETE FROM balances_ctas WHere numbalan=" & Adodc1.Recordset!NumBalan
    Conn.Execute sql
    
    'Eliminamos las lineas del balance
    sql = "DELETE FROM balances_texto WHere numbalan=" & Adodc1.Recordset!NumBalan
    Conn.Execute sql
    
    'Eliminamos el balance
    sql = "DELETE FROM balances WHere numbalan=" & Adodc1.Recordset!NumBalan
    Conn.Execute sql
    
End Sub



Private Sub ComprobarBalance(NumBal, EsPerdidas As Boolean)
Dim cad As String
    
    
    'UPDATEAMOS TIENE CUENTAS A 0
    Conn.Execute "UPDATE balances_texto SET tienenctas=0 where numbalan=" & NumBal
    
    Set miRsAux = New ADODB.Recordset
    cad = "select numbalan,pasivo,codigo from balances_ctas group by"
    cad = cad & " numbalan,pasivo,codigo having numbalan=" & NumBal
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = "UPDATE balances_texto set tienenctas=1 where numbalan=" & miRsAux!NumBalan
        cad = cad & " and pasivo='" & miRsAux!Pasivo & "' AND codigo = " & miRsAux!Codigo
        Conn.Execute cad
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    'Haremos una segunda comprobacion
    Label1.Caption = "Obteniendo cuentas ejercicio actual"
    Me.Refresh
    DoEvents
            
            Conn.Execute "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
            cad = "INSERT INTO tmpcierre1(codusu,cta) "
            cad = cad & "Select " & vUsu.Codigo & ",codmacta from hlinapu where "
            cad = cad & " fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            cad = cad & "' AND fechaent<='" & Format(vParam.fechafin, FormatoFecha) & "' AND "

            If Not EsPerdidas Then cad = cad & " NOT "
            cad = cad & "(codmacta like '" & vParam.grupogto & "%' OR "
            cad = cad & "codmacta like '" & vParam.grupovta & "%'"
            If vParam.Subgrupo1 <> "" Then cad = cad & " OR " & "codmacta like '" & vParam.Subgrupo1 & "%'"
            cad = cad & ") GROUP BY codmacta"
            Conn.Execute cad
            
            
            'Ya tengo todas las cuentas que entran en hlinapu
            'Cogere la configuracion y para cada cuenta ire quitando
            'Las que esten configuradas
            Label1.Caption = "Comprobando configuracion"
            Me.Refresh
            DoEvents
    
            
            cad = "Select * from balances_ctas where numbalan=" & NumBal
            miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo & " AND cta like '"
            While Not miRsAux.EOF
                Conn.Execute cad & miRsAux!codmacta & "%'"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
                
            'Veremos las que queden
            cad = "SELECT * FROM tmpcierre1 where codusu =" & vUsu.Codigo
            miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NumRegElim = 0
            cad = ""
            While Not miRsAux.EOF
                cad = cad & "      " & miRsAux!Cta
                NumRegElim = NumRegElim + 1
                If (NumRegElim Mod 11) = 0 Then cad = cad & vbCrLf
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If NumRegElim > 0 Then
                cad = "Hay cuentas(" & NumRegElim & ") que parecen no haber sido configuradas en el balance" & vbCrLf & vbCrLf & cad
                MsgBox cad, vbInformation
                NumRegElim = 0
            Else
                
                'Haremos una tercera comprobacion
                Label1.Caption = "Nodos superiores con cuentas"
                Me.Refresh
                DoEvents
            
                cad = "select * from balances_texto where numbalan=1 and tienenctas=1 and (pasivo,codigo) in ("
                cad = cad & " select pasivo,padre from balances_texto where numbalan=1 and padre>=0) order by pasivo,codigo"
                miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                NumRegElim = 0
                cad = ""
                While Not miRsAux.EOF
                    cad = cad & "      -" & miRsAux!Pasivo & miRsAux!Codigo & "     " & miRsAux!deslinea & vbCrLf
                    NumRegElim = NumRegElim + 1
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
                
                If NumRegElim > 0 Then
                    cad = "No es nodo de ultimo nivel y tienen cuentas configuradas" & vbCrLf & cad
                    MsgBox cad, vbExclamation
                Else
                    MsgBox "Comprobacion  finalizada", vbInformation
                End If
            End If
    Set miRsAux = Nothing
    

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim C As String
    WheelUnHook
    Select Case Button.Index
        Case 1
                BotonAnyadir
        Case 2
                BotonModificar
        Case 3
                BotonEliminar
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
        Case 8
                'Imprimimos el listado
                BotonImprimir
        Case Else
        
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim C As String
    WheelUnHook
    
    Select Case Button.Index
        Case 1 'comprobar
            BotonComprobar

        Case 2 'copiar
            BotonCopiar
        Case Else
    End Select
End Sub

Private Sub BotonImprimir()
Dim CodigoBalanceBuscar As Integer
    
    
    frmColBalanList.NumBalan = Adodc1.Recordset!NumBalan
    frmColBalanList.NomBalan = Adodc1.Recordset!NomBalan
    frmColBalanList.Show vbModal
    
    Screen.MousePointer = vbHourglass
    CargaGrid
    Adodc1.Recordset.Find "Numbalan = " & CodigoBalanceBuscar
    Screen.MousePointer = vbDefault

End Sub

Private Sub BotonCopiar()
Dim CodigoBalanceBuscar As Integer
    
    If Adodc1.Recordset.EOF Then Exit Sub
    CodigoBalanceBuscar = Adodc1.Recordset!NumBalan
    CadenaDesdeOtroForm = Adodc1.Recordset!NumBalan & "|" & Adodc1.Recordset!NomBalan & "|"
    frmBalancesCopy.Opcion = 57
    frmBalancesCopy.Show vbModal
    CargaGrid
    
    Screen.MousePointer = vbHourglass
    CargaGrid
    Adodc1.Recordset.Find "Numbalan = " & CodigoBalanceBuscar
    Screen.MousePointer = vbDefault

End Sub

Private Sub BotonComprobar()
Dim CodigoBalanceBuscar As Integer
    
    If Adodc1.Recordset.EOF Then Exit Sub
    CodigoBalanceBuscar = Adodc1.Recordset!NumBalan


    Screen.MousePointer = vbHourglass
    Label1.Tag = Label1.Caption
    Label1.Caption = "Comprobaciones ....."
    Label1.Refresh
    ComprobarBalance Adodc1.Recordset!NumBalan, Adodc1.Recordset!perd = "SI"
    Label1.Caption = Label1.Tag
    Label1.Tag = ""
    Screen.MousePointer = vbDefault
    

    Screen.MousePointer = vbHourglass
    CargaGrid
    Adodc1.Recordset.Find "Numbalan = " & CodigoBalanceBuscar
    Screen.MousePointer = vbDefault

End Sub


Private Sub BotonAnyadir()
Dim CodigoBalanceBuscar As Integer
    
    Screen.MousePointer = vbHourglass
    
    NumRegElim = ObtenerSiguiente
    frmBalances.numBalance = 0
    frmBalances.Show vbModal
    CodigoBalanceBuscar = NumRegElim

    Screen.MousePointer = vbHourglass
    CargaGrid
    Adodc1.Recordset.Find "Numbalan = " & CodigoBalanceBuscar
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub BotonBuscar()

End Sub

Private Sub BotonVerTodos()

End Sub


Private Sub BotonModificar()
Dim CodigoBalanceBuscar As Integer
    
    If Adodc1.Recordset.EOF Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    'Nuevo balance   y modificar
    NumRegElim = 0
    frmBalances.numBalance = Adodc1.Recordset!NumBalan
    frmBalances.Show vbModal
    'Para luego hace la busqueda
    NumRegElim = Adodc1.Recordset!NumBalan
    CodigoBalanceBuscar = NumRegElim
    Screen.MousePointer = vbHourglass
    CargaGrid
    Adodc1.Recordset.Find "Numbalan = " & CodigoBalanceBuscar
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub BotonEliminar()
Dim CodigoBalanceBuscar As Integer
Dim sql As String
    
    On Error GoTo Error2
    
    If Adodc1.Recordset.EOF Then Exit Sub
    
    EliminarBalance
    
    Screen.MousePointer = vbHourglass
    CargaGrid
    Screen.MousePointer = vbDefault
    
Error2:

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And Me.Adodc1.Recordset.Fields(0) <> 0
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And Me.Adodc1.Recordset.Fields(0) <> 0
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Modificar, "N") And Me.Adodc1.Recordset.Fields(0) <> 0
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Modificar, "N") And Me.Adodc1.Recordset.Fields(0) <> 0
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And Me.Adodc1.Recordset.Fields(0) <> 0
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


