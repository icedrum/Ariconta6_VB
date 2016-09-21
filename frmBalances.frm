VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalances22 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances.frx":629A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances.frx":6CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances.frx":76BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances.frx":D2E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalances.frx":12952
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5400
      Width           =   6135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pasivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Activo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   7200
      TabIndex        =   1
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   8705
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7858
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   6375
   End
   Begin VB.Menu menuListview 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnInsertarNuevaCuenta 
         Caption         =   "Insertar cuenta"
         Begin VB.Menu mnInsertarSaldo 
            Caption         =   "Saldo"
         End
         Begin VB.Menu mnInsertarHaber 
            Caption         =   "Haber"
         End
         Begin VB.Menu mnInsertDebe 
            Caption         =   "Debe"
         End
      End
      Begin VB.Menu mnEliminarCuena 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu menuTree 
      Caption         =   "Tree"
      Visible         =   0   'False
      Begin VB.Menu mnNuevo 
         Caption         =   "Nuevo"
         Begin VB.Menu mnNuevoGrupo 
            Caption         =   "Elto. del grupo"
         End
         Begin VB.Menu mnInsertarSubGrupo 
            Caption         =   "Sub grupo"
         End
      End
      Begin VB.Menu mnEliminarGrupo 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mnDescLin 
         Caption         =   "Descripcion linea"
      End
      Begin VB.Menu mnbarra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnModificarTexto 
         Caption         =   "Modificar texto asociado"
      End
      Begin VB.Menu mnbarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnArriba 
         Caption         =   "Hacia arriba"
      End
      Begin VB.Menu mnAbajo 
         Caption         =   "Hacia abajo"
      End
   End
End
Attribute VB_Name = "frmBalances22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte  '0.- De situacion  1.- Perdidas y ganancias

Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim SQL As String
Dim Rs As Recordset
Dim NodoArbol As Node
Dim I As Integer
Dim Devolucion As String
Dim Clave As String

Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False
    If Opcion = 0 Then
    'Cargamos el tree
        CargaSituacion
    Else
        'cargaperd
        CargaPerdidas
    End If
    
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Top = 1000
    Left = 1000
    PrimeraVez = True
    'Opcion = 1
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 15
    End With
    TreeView1.ImageList = Me.ImageList1
    ListView1.Icons = Me.ImageList1
    If Opcion = 0 Then
        Caption = "Configurador balance situación"
        Label1.Caption = "Situación"
    Else
        Caption = "Configurador balance pérdidas y ganancias"
        Label1.Caption = "Pérdidas y ganancias"
    End If
    Text1.Text = ""
End Sub



Private Sub CargaSituacion()
On Error GoTo Ecargasituacion
Dim Padre As String

    TreeView1.Nodes.Clear
    SQL = "SELECT * FROM sbalan WHERE numlinea ='"
    If Option1(0).Value Then
        SQL = SQL & "A"
    Else
        SQL = SQL & "P"
    End If
    SQL = SQL & "' ORDER BY numlinea ASC, numline1 ASC, numline2 ASC, numline3 ASC"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        Rs.Close
        Exit Sub
    End If
    
    'Haremos tres pasadas, primero buscamos el nivel numline2=0, numline3=0
    Rs.MoveFirst
    While Not Rs.EOF
        If Rs!numline2 = "0" And Rs!Numline3 = "0" Then
            'Nodo principal
            Clave = Rs!NumLinea & "|" & Rs!numline1 & "|" & Rs!numline2 & "|" & Rs!Numline3 & "|"
            Set NodoArbol = TreeView1.Nodes.Add(, , Clave, Rs!deslinea)
            NodoArbol.Tag = DBLet(Rs!texlinea)
            NodoArbol.Image = 1
        End If
        Rs.MoveNext
    Wend
    
    'Segunda pasada, los niveles 2
    Rs.MoveFirst
    While Not Rs.EOF
        If Rs!numline2 <> "0" And Rs!Numline3 = "0" Then
            'Clave padre
            Padre = Rs!NumLinea & "|" & Rs!numline1 & "|0|" & Rs!Numline3 & "|"
            'Nodo principal
            Clave = Rs!NumLinea & "|" & Rs!numline1 & "|" & Rs!numline2 & "|" & Rs!Numline3 & "|"
            Set NodoArbol = TreeView1.Nodes.Add(Padre, tvwChild, Clave, Rs!deslinea)
            NodoArbol.Tag = DBLet(Rs!texlinea)
            I = NodoArbol.Index
            NodoArbol.Image = 1
            TreeView1.Nodes(I).EnsureVisible
        End If
        Rs.MoveNext
    Wend
        
        
    'Tercera pasada, los niveles 2
    Rs.MoveFirst
    While Not Rs.EOF
        If Rs!numline2 <> "0" And Rs!Numline3 <> "0" Then
            'Clave padre
            Padre = Rs!NumLinea & "|" & Rs!numline1 & "|" & Rs!numline2 & "|0|"
            'Nodo principal
            Clave = Rs!NumLinea & "|" & Rs!numline1 & "|" & Rs!numline2 & "|" & Rs!Numline3 & "|"
            Set NodoArbol = TreeView1.Nodes.Add(Padre, tvwChild, Clave, Rs!deslinea)
            NodoArbol.Tag = Rs!texlinea
            NodoArbol.Image = 2
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    TreeView1.Nodes(1).EnsureVisible
    Set TreeView1.SelectedItem = TreeView1.Nodes(1)
    CargaListview
    Set NodoArbol = Nothing
Exit Sub
Ecargasituacion:
    MuestraError Err.Number, "Carga arbol"
End Sub




Private Sub CargaPerdidas()
On Error GoTo ECargaPerdidas
Dim Padre As String

    TreeView1.Nodes.Clear
    SQL = "SELECT * FROM sperdi WHERE numlinea ='"
    If Option1(0).Value Then
        SQL = SQL & "A"
    Else
        SQL = SQL & "B"
    End If
    SQL = SQL & "' ORDER BY numlinea ASC, numline1 ASC, orden ASC"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        Rs.Close
        Exit Sub
    End If
    
    'Haremos 2 pasadas, primero buscamos el nivel numline2=0
    Rs.MoveFirst
    While Not Rs.EOF
        If Rs!numline1 = "0" And Rs!numline2 = "0" Then
            'Nodo principal
            Clave = Rs!NumLinea & "|" & Rs!numline1 & "|" & Rs!numline2 & "|"
            Set NodoArbol = TreeView1.Nodes.Add(, , Clave, Rs!deslinea)
            NodoArbol.Tag = DBLet(Rs!texlinea)
            NodoArbol.Image = 1
        End If
        Rs.MoveNext
    Wend
    
    'Segunda pasada, los niveles 2
    Rs.MoveFirst
    While Not Rs.EOF
        If Rs!numline1 <> "0" And Rs!numline2 = "0" Then
            'Clave padre
            Padre = Rs!NumLinea & "|0|0|"
            'Nodo principal
            Clave = Rs!NumLinea & "|" & Rs!numline1 & "|" & Rs!numline2 & "|"
            Set NodoArbol = TreeView1.Nodes.Add(Padre, tvwChild, Clave, Rs!deslinea)
            NodoArbol.Tag = DBLet(Rs!texlinea)
            I = NodoArbol.Index
            NodoArbol.Image = 2
            TreeView1.Nodes(I).EnsureVisible
        End If
        Rs.MoveNext
    Wend
        
    Rs.Close
    Set Rs = Nothing
    
    TreeView1.Nodes(1).EnsureVisible
    Set TreeView1.SelectedItem = TreeView1.Nodes(1)
    CargaListview
    Set NodoArbol = Nothing
Exit Sub
ECargaPerdidas:
    MuestraError Err.Number, "Carga arbol"
End Sub





Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
'Comprobamos k la cuenta no esta
Devolucion = RecuperaValor(CadenaSeleccion, 1)
End Sub


Private Function CompruebaCuenta(ByRef Cta As String) As Boolean

    If Opcion = 0 Then
        SQL = "Select * from sbala2"
    Else
        SQL = "Select * from sperd2"
    End If
    SQL = SQL & " WHERE codmacta='" & Cta & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not Rs.EOF
        If Rs!tipsaldo <> "S" Then _
            SQL = SQL & Rs!NumLinea & " - " & Rs!numline1 & " - " & Rs!numline2 & " - " & Rs!tipsaldo & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If SQL <> "" Then
        MsgBox "La cuenta ya en los registros: " & vbCrLf & SQL, vbExclamation
        CompruebaCuenta = False
    Else
        CompruebaCuenta = True
    End If
        
End Function



Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.menuListview
End Sub

Private Sub mnDescLin_Click()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    TreeView1.StartLabelEdit
End Sub

Private Sub InsertarNueva(Tipo As String)
    If TreeView1.Nodes.Count = 0 Then Exit Sub
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Children > 0 Then
        MsgBox "Tiene subniveles", vbExclamation
        Exit Sub
    End If
    Set frmCta = New frmColCtas
    frmCta.ConfigurarBalances = 1
    frmCta.DatosADevolverBusqueda = "0|1"
    Devolucion = ""
    frmCta.Show vbModal
    Set frmCta = Nothing
    If Devolucion <> "" Then
        If CompruebaCuenta(Devolucion) Then
            'OK, hay que insertar
            InsertaLaCuenta Tipo
            CargaListview
        End If
    End If
End Sub



Private Sub InsertaLaCuenta(ByRef Tipo As String)
    On Error Resume Next
    If Opcion = 0 Then
        SQL = "INSERT INTO sbala2 (numlinea, numline1, numline2, numline3, codmacta, tipsaldo) VALUES ("
    Else
        SQL = "INSERT INTO sperd2 (numlinea, numline1, numline2, codmacta, tipsaldo) VALUES ("
    End If
    SQL = SQL & "'" & RecuperaValor(TreeView1.SelectedItem.Key, 1) & "',"
    SQL = SQL & "'" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "',"
    SQL = SQL & "'" & RecuperaValor(TreeView1.SelectedItem.Key, 3) & "',"
    If Opcion = 0 Then SQL = SQL & "'" & RecuperaValor(TreeView1.SelectedItem.Key, 4) & "',"
    SQL = SQL & "'" & Devolucion & "',"
    SQL = SQL & "'" & Tipo & "')"
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Insertando cuenta"
    End If
End Sub

Private Sub mnEliminarCuena_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    SQL = "Seguro que desea eliminar la cuenta " & ListView1.SelectedItem.Text
    SQL = SQL & "       de '" & TreeView1.SelectedItem.Text & "'?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    If Opcion = 0 Then
        SQL = "DELETE FROM sbala2 "
    Else
        SQL = "DELETE FROM sperd2 "
    End If
    SQL = SQL & PonSQLdelWHERE
    SQL = SQL & " AND codmacta ='" & ListView1.SelectedItem.Text & "'"
    Conn.Execute SQL
    CargaListview
    ListView1.SetFocus
End Sub

Private Sub mnEliminarGrupo_Click()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Children > 0 Then
        MsgBox "Cuelgan entradas de este nodo", vbExclamation
        Exit Sub
    End If
    
    
    'Eliminamos
    
    
End Sub

Private Sub mnInsertarHaber_Click()
    InsertarNueva "H"
End Sub


Private Sub mnInsertarSaldo_Click()
    InsertarNueva "S"
End Sub

Private Sub mnInsertarSubGrupo_Click()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If Opcion = 1 Then
        'para perdidas y ganacias es el ultimo novil k admite subgrpos
        If RecuperaValor(TreeView1.SelectedItem.Key, 3) <> "0" Then Exit Sub
    Else
        'Para balance de situacion
        If RecuperaValor(TreeView1.SelectedItem.Key, 4) <> "0" Then Exit Sub
    End If
    If Opcion = 0 Then
        NuevoEltoGrupoSituacion
    Else
        
    End If

End Sub

Private Sub mnInsertDebe_Click()
    InsertarNueva "D"
End Sub

Private Sub mnModificarTexto_Click()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    'Habilitampos el text
    Text1.Enabled = True
    Text1.SetFocus
End Sub

Private Sub mnNuevo_Click()
    'Este cotiene dos opciones
End Sub

Private Sub mnNuevoGrupo_Click()
If Opcion = 0 Then
    NuevoSituacion
Else
    'nuevoPyG
End If

End Sub

Private Sub Option1_Click(Index As Integer)
    If PrimeraVez Then Exit Sub
    If Opcion = 0 Then
        CargaSituacion
    Else
        CargaPerdidas
    End If
End Sub



Private Sub CargaListview()
Dim Itmx As ListItem
    ListView1.ListItems.Clear
    Text1.Text = ""
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    Text1.Text = TreeView1.SelectedItem.Tag
'Comun

    If Opcion = 0 Then
        SQL = "Select * from sbala2 "
    Else
        SQL = "Select * from sperd2 "
    End If
    SQL = SQL & PonSQLdelWHERE
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set Itmx = ListView1.ListItems.Add(, , Rs!codmacta)   ' Autor.
        Select Case Rs!tipsaldo
        Case "D"
             Itmx.Icon = 4
        Case "H"
              Itmx.Icon = 5
        Case Else
              Itmx.Icon = 6
        End Select
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End Sub

Private Function PonSQLdelWHERE() As String
Dim Cad As String

    If Me.Option1(0).Value Then
        Cad = "'A'"
    Else
        Cad = "'P'"
        If Opcion = 1 Then Cad = "'B'"
    End If
    If Opcion = 0 Then
        Cad = "Where Numlinea = " & Cad
        Cad = Cad & " AND numline1 ='" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "'"
        Cad = Cad & " AND numline2 ='" & RecuperaValor(TreeView1.SelectedItem.Key, 3) & "'"
        Cad = Cad & " AND numline3 ='" & RecuperaValor(TreeView1.SelectedItem.Key, 4) & "'"
    Else
        Cad = "Where Numlinea = " & Cad
        Cad = Cad & " AND numline1 ='" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "'"
        Cad = Cad & " AND numline2 ='" & RecuperaValor(TreeView1.SelectedItem.Key, 3) & "'"
    End If
    PonSQLdelWHERE = Cad
End Function


Private Sub Text1_GotFocus()
With Text1
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus()
    'Modificamos el texto
    Text1.Enabled = False
    'Actualizamos el tag del tree y la BD
    TreeView1.SelectedItem.Tag = Trim(Text1.Text)
    If Opcion = 0 Then
        SQL = "UPDATE sbalan"
    Else
        SQL = "UPDATE sperdi"
    End If
    SQL = SQL & " SET texlinea = '" & Text1.Text & "' "
    SQL = SQL & PonSQLdelWHERE
    Conn.Execute SQL
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Unload Me
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    I = Len(Trim(NewString))
    If I < 1 Or I > 60 Then
        MsgBox "La cadena ocupar entre 1 y 60 caracteres", vbExclamation
        Cancel = 1
        Exit Sub
    End If
    'Updateamos la BD
    If Opcion = 0 Then
        SQL = "UPDATE sbalan"
    Else
        SQL = "UPDATE sperdi"
    End If
    SQL = SQL & " SET deslinea = '" & NewString & "' "
    SQL = SQL & PonSQLdelWHERE
    Conn.Execute SQL
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Me.menuTree
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    CargaListview
End Sub




Private Sub NuevoSituacion()
Dim vSql As String
Dim Letra As String
'Dim GrupoRoot As Boolean

    If Option1(0).Value Then
        Letra = "A"
    Else
        Letra = "P"
    End If
    
    If TreeView1.Nodes.Count = 0 Then
        'GrupoRoot = True
        vSql = "'" & Letra & "','A','0','0'"
        I = 1
    Else
        If TreeView1.SelectedItem.Parent Is Nothing Then
           ' GrupoRoot = True
            'Primer nivel
            Clave = ObtenerSiguiente(1)
            vSql = "'" & Letra & "','" & Clave & "','0','0'"
            Clave = Letra & "|" & Clave & "|0|0|"
            I = 1
            Else
                If TreeView1.SelectedItem.Parent.Parent Is Nothing Then
                    'Segundo nivel
                    Clave = ObtenerSiguiente(2)
                    vSql = "'" & Letra & "','" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "','" & Clave & "','0'"
                    Clave = Letra & "|" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "|" & Clave & "|0|"
                    I = 1
                Else
                    'Tercer nivel.
                    Clave = ObtenerSiguiente(2)
                    vSql = "'" & Letra & "','" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "','" & RecuperaValor(TreeView1.SelectedItem.Key, 3) & "','" & Clave & "'"
                    Clave = Letra & "|" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "|" & RecuperaValor(TreeView1.SelectedItem.Key, 3) & "|" & ObtenerSiguiente(2) & "|"
                    I = 2
                End If
        End If
    End If
    
    SQL = "INSERT INTO sbalan (numlinea, numline1, numline2, numline3, deslinea, texlinea) VALUES ("
    
    SQL = SQL & vSql & ",'Nuevo','')"
    Conn.Execute SQL
    ListView1.ListItems.Clear
    Text1.Text = ""
    
    If TreeView1.Nodes.Count = 0 Then
        SQL = ""
        Set NodoArbol = TreeView1.Nodes.Add(, , Clave, "Nuevo elemento", 2)
    Else
        
        If TreeView1.SelectedItem.Parent Is Nothing Then
            Set NodoArbol = TreeView1.Nodes.Add(, , Clave, "Nuevo elemento")
        Else
            SQL = TreeView1.SelectedItem.Parent.Key
            Set NodoArbol = TreeView1.Nodes.Add(SQL, tvwChild, Clave, "Nuevo elemento", 2)
        End If
    End If
    NodoArbol.Image = I
    Set TreeView1.SelectedItem = NodoArbol
    Set NodoArbol = Nothing
    TreeView1.StartLabelEdit
End Sub




Private Sub NuevoEltoGrupoSituacion()
Dim vSql As String
Dim Letra As String
'Dim GrupoRoot As Boolean
Dim ICono As Integer
    If Option1(0).Value Then
        Letra = "A"
    Else
        Letra = "P"
    End If
    
    If TreeView1.Nodes.Count = 0 Then
        MsgBox "No hay nodos de primer nivel", vbExclamation
        Exit Sub
    Else
        If TreeView1.SelectedItem.Parent Is Nothing Then
            'INSERTAR NODO Segundo nivel
            Clave = ObtenerSiguiente(2)
            vSql = "'" & Letra & "','" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "','" & Clave & "','0'"
            Clave = Letra & "|" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "|" & Clave & "|0|"
            I = 1
        Else
                If TreeView1.SelectedItem.Parent.Parent Is Nothing Then
                    'Tercer nivel.
                    Clave = ObtenerSiguiente(3)
                    vSql = "'" & Letra & "','" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "','" & RecuperaValor(TreeView1.SelectedItem.Key, 3) & "','" & Clave & "'"
                    Clave = Letra & "|" & RecuperaValor(TreeView1.SelectedItem.Key, 2) & "|" & RecuperaValor(TreeView1.SelectedItem.Key, 3) & "|" & ObtenerSiguiente(2) & "|"
                    I = 1
                Else
                    MsgBox "no se pueden insertar subgrupos al ultimo nivel", vbExclamation
                    Exit Sub
                End If
        End If
    End If
    
    SQL = "INSERT INTO sbalan (numlinea, numline1, numline2, numline3, deslinea, texlinea) VALUES ("
    
    SQL = SQL & vSql & ",'Nuevo','')"
    Conn.Execute SQL
    ListView1.ListItems.Clear
    Text1.Text = ""
    
    If TreeView1.Nodes.Count = 0 Then
        SQL = ""
        Set NodoArbol = TreeView1.Nodes.Add(, , Clave, "Nuevo elemento", 1)
    Else
        
            SQL = TreeView1.SelectedItem.Key
            Set NodoArbol = TreeView1.Nodes.Add(SQL, tvwChild, Clave, "Nuevo elemento", 1)
    End If
    NodoArbol.Image = I
    Set TreeView1.SelectedItem = NodoArbol
    Set NodoArbol = Nothing
    TreeView1.StartLabelEdit
End Sub


















Private Function ObtenerSiguiente(Nivel As Byte) As String
Select Case Nivel
Case 2
    SQL = "numline2"
Case 3
    SQL = "numline3"
Case Else
    SQL = "numline1"
End Select
SQL = "Select max(" & SQL & ") from sbalan"
If Opcion = 0 Then
    If Option1(0).Value Then
        SQL = SQL & " WHERE numlinea='A'"
    Else
        SQL = SQL & " WHERE numlinea='P'"
    End If
End If
Set Rs = New ADODB.Recordset
Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Rs.EOF Then
    SQL = "A"
Else
    SQL = Rs.Fields(0)
    SQL = Chr(Asc(SQL) + 1)
End If
Rs.Close
Set Rs = Nothing
ObtenerSiguiente = SQL
End Function
