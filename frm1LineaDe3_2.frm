VERSION 5.00
Begin VB.Form frm1LineaDe3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd1 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   290
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2865
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   290
      Left            =   4755
      Picture         =   "frm1LineaDe3_2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   15
      Width           =   315
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Height          =   290
      Left            =   4395
      Picture         =   "frm1LineaDe3_2.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   15
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   290
      Index           =   1
      Left            =   1425
      TabIndex        =   1
      Top             =   0
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   290
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   800
   End
   Begin VB.Image imgMod 
      Height          =   240
      Left            =   15
      Picture         =   "frm1LineaDe3_2.frx":0204
      Top             =   30
      Width           =   240
   End
   Begin VB.Image imgNuevo 
      Height          =   240
      Left            =   60
      Picture         =   "frm1LineaDe3_2.frx":0306
      Top             =   30
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   0
      Top             =   6
      Width           =   510
   End
End
Attribute VB_Name = "frm1LineaDe3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vTIPO As Byte   'Para saber que es lo quien lo ha llamado
    '0 .- Conceptos
    '1 .- Diarios
    '2 .- Cuentas
    '3 .- Porcentaje



Public vCamposHabilitados As String
Public vBotonesVisibles As String
Public vAnchoCampos As String
Public vModo As Byte '0.- Nuevo  1.- Modificar  2.- Buscar
Public vTop As Long
Public vLeft As Long
Public vCadena As String
'Public vFac As Long
Public vFac As String
Public vLinea As String

Dim Sql As String
Dim PrimeraVez As Boolean
Dim Rs As Recordset
Dim sng As Single

Dim MaxValor As Single  'Valor maximo para
                        'Cuando % en centro de coste



Private Sub cmdAceptar_Click()
Dim Cad As String
Dim Rc As Byte
On Error GoTo ECmd1

    'Compobaremos que los datos son correctos
    Cad = DatosOk(vModo = 2)
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
    Cad = ""
    'Si son correctos haremos las operaciones que correspondan segun el boton pulsado
    Select Case vTIPO
    Case 0
    '-----------------------------------------------------------
    '   CONCEPTOS
    '-----------------------------------------------------------
        If vModo < 2 Then
            If Not InsertarModificarConcepto Then Exit Sub
            'Correcto
            FormularioHijoModificado = True
        Else
            'Buscar, generar la cadena de busqueda
            Rc = SeparaCampoBusqueda("N", "conceptos.codconce", Text1(0).Text, Cad)
            If Rc = 0 Then
               CadenaDevueltaFormHijo = Cad
            End If
            'El campo NOMBRE
            If Text1(1).Text <> "" Then
                Rc = SeparaCampoBusqueda("T", "conceptos.nomconce", Text1(1).Text, Cad)
                If Rc = 0 Then
                    If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                    CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & Cad
                End If
            End If
                  
            'El tipo de concepto
            If Combo1.ListIndex >= 0 Then
                Rc = SeparaCampoBusqueda("N", "tipoconceptos.tipoconce", CStr(Combo1.ItemData(Combo1.ListIndex)), Cad)
                If Rc = 0 Then
                    If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                    CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & Cad
                End If
            End If
            
            If CadenaDevueltaFormHijo = "" Then
                MsgBox "Error generando consulta. Compruebe que los campos de busqueda son correctos.", vbExclamation
                Exit Sub
            End If
              
            'Correcto
            FormularioHijoModificado = True
        End If 'modo <2
    
    '-----------------------------------------------------------
    '   DIARIO
    '-----------------------------------------------------------
    Case 1
        If vModo < 2 Then
            If Not InsertarModificarDiario Then Exit Sub
            'Correcto
            FormularioHijoModificado = True
        Else
            'Buscar, generar la cadena de busqueda
            Rc = SeparaCampoBusqueda("N", "tiposdiario.numdiari", Text1(0).Text, Cad)
            If Rc = 0 Then
               CadenaDevueltaFormHijo = Cad
            End If
            'El campo NOMBRE
            If Text1(1).Text <> "" Then
                Rc = SeparaCampoBusqueda("T", "tiposdiario.desdiari", Text1(1).Text, Cad)
                If Rc = 0 Then
                    If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                    CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & Cad
                End If
            End If
            If CadenaDevueltaFormHijo = "" Then
                MsgBox "Error generando consulta. Compruebe que los campos de busqueda son correctos.", vbExclamation
                Exit Sub
            End If
              
            'Correcto
            FormularioHijoModificado = True
        End If 'modo <2
        
    
    
    
    
    '-----------------------------------------------------------
    '   COLECCION DE CUENTAS
    '-----------------------------------------------------------
    Case 2
        If vModo = 2 Then
            'Buscar, generar la cadena de busqueda
            'El campo codcta
            If Text1(0).Text <> "" Then
                Rc = SeparaCampoBusqueda("T", "cuentas.codmacta", Text1(0).Text, Cad)
                If Rc = 0 Then CadenaDevueltaFormHijo = Cad
            End If
            'El campo NOMBRE
            If Text1(1).Text <> "" Then
                Rc = SeparaCampoBusqueda("T", "cuentas.nommacta", Text1(1).Text, Cad)
                If Rc = 0 Then
                    If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                    CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & Cad
                End If
            End If
            If Combo1.ListIndex > -1 Then
                    'Solo los apuntes directos
                    If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                    If Combo1.ListIndex = 0 Then
                        Cad = "'S'"
                    Else
                        Cad = "'N'"
                    End If
                    CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & "apudirec = " & Cad
            End If
            If CadenaDevueltaFormHijo = "" Then
                MsgBox "Error generando consulta. Compruebe que los campos de busqueda son correctos.", vbExclamation
                Exit Sub
            End If
              
            'Correcto
            FormularioHijoModificado = True
        End If 'modo <2
    
        '-----------------------------------------------------------------
        Case 3
            If vModo < 2 Then
                If Not InsertarModificarCCoste Then Exit Sub
                'Correcto
                FormularioHijoModificado = True
            Else
                'Buscar, generar la cadena de busqueda
                Rc = SeparaCampoBusqueda("N", "conceptos.codconce", Text1(0).Text, Cad)
                If Rc = 0 Then
                   CadenaDevueltaFormHijo = Cad
                End If
                'El campo NOMBRE
                If Text1(1).Text <> "" Then
                    Rc = SeparaCampoBusqueda("T", "conceptos.nomconce", Text1(1).Text, Cad)
                    If Rc = 0 Then
                        If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                        CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & Cad
                    End If
                End If
                      
                'El tipo de concepto
                If Combo1.ListIndex >= 0 Then
                    Rc = SeparaCampoBusqueda("N", "tipoconceptos.tipoconce", CStr(Combo1.ItemData(Combo1.ListIndex)), Cad)
                    If Rc = 0 Then
                        If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                        CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & Cad
                    End If
                End If
                
                If CadenaDevueltaFormHijo = "" Then
                    MsgBox "Error generando consulta. Compruebe que los campos de busqueda son correctos.", vbExclamation
                    Exit Sub
                End If
                  
                'Correcto
                FormularioHijoModificado = True
            End If 'modo <2
    
    End Select
    Unload Me
Exit Sub
ECmd1:
    MuestraError Err.Number, "Error general en modulo frm1LinD6. Boton aceptar. " & vbCrLf & vbCrLf & Err.Description
End Sub


Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim i As Integer
If PrimeraVez Then
    PrimeraVez = False
    Select Case vModo
    Case 0
        Text1(0).Text = vCadena
        Text1(0).SetFocus
    Case 1
            Text1(0).Locked = True
            For i = 0 To Text1.Count - 1
                Text1(i).Text = RecuperaValor(vCadena, i + 1)
            Next i
            'Ponemos valor al combo1
            i = Val(RecuperaValor(vCadena, 3))
            PonvalorCombo (i)
            cmdCancelar.SetFocus
    Case 2
        'Buscar
    End Select
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim aux As Single
Dim bol As Boolean

'Situamos el form
Top = vTop
Left = vLeft
PrimeraVez = True

imgMod.Visible = vModo = 1
imgNuevo.Visible = vModo = 0
sng = 290
'Situamos todos los campos, en funcion de si hay botones visbles
'El primer o es fijo, de ancho determinado
    aux = CSng(RecuperaValor(vAnchoCampos, 1))
    Text1(0).Left = sng
    Text1(0).Width = aux - 6
    sng = sng + aux
    
'A partir del segundo ya vamos mirando
'Vemos que camos van en blancos y esta habilidados y cuales van en pastel y no
    aux = CSng(RecuperaValor(vAnchoCampos, 2))
    Text1(1).Left = sng
    Text1(1).Width = aux - 6
    sng = sng + aux '+ 15
    
    'Ahora vemos cuales estan habilitados y cuales no
    For i = 0 To 1
        bol = RecuperaValor(vCamposHabilitados, i + 1) = "S" 'NO HABILITADOS
        Text1(i).Enabled = bol
        If bol Then
            Text1(i).BackColor = vbWhite
        Else
            Text1(i).BackColor = -2147483624
        End If
    Next i
    
    '-------------------
    'El combo solo si el tipo es 0 o 2
    aux = CSng(RecuperaValor(vAnchoCampos, 3))
    cmd1.Visible = False
    Select Case vTIPO
    Case 0, 2
        ' El; combo
        CargaCombo
        Combo1.Visible = True
        Text1(2).Visible = False
        Combo1.Left = sng
'''        aux = aux + 200 'El botoncito del combo
        Combo1.Width = aux
        
        
    
        'El combo
        bol = RecuperaValor(vCamposHabilitados, 3) = "S" 'NO HABILITADOS
        Combo1.Enabled = bol
        If bol Then
            Combo1.BackColor = vbWhite
        Else
            Combo1.BackColor = -2147483624
        End If
    
    Case 3
        'Visible el text
        Combo1.Visible = False
        Text1(2).Visible = True
        Text1(2).Left = sng
        Text1(2).Width = aux
        'Para el Centro de coste
        cmd1.Left = Text1(1).Left
        Text1(1).Left = cmd1.Left + cmd1.Width '+ 15
        Text1(1).Width = Text1(1).Width - cmd1.Width - 30
        cmd1.Visible = True
    Case 1
        aux = 0
        Combo1.Visible = False
        Text1(2).Visible = False
    End Select
    'Fijamos el ancho total
    sng = sng + aux '+ 60
    cmdAceptar.Left = sng
    cmdCancelar.Left = sng + 330
    'y la posicion de los aceptar y cancelar
    sng = sng + 660 'Ancho 2 botones +espacios intermedios
    Me.Width = sng
    If vTIPO = 3 Then
        ObtenerMaxValor
        Text1(2).Text = MaxValor
    End If
End Sub





Private Sub Text1_GotFocus(Index As Integer)
With Text1(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)

    'Comprobaremos ciertos valores
    Text1(Index).Text = Trim(Text1(Index).Text)

    'Comun a todos
    If Text1(Index).Text = "" Then Exit Sub
    
    Select Case vTIPO
    Case 0
        LostFocusTipo0 Index
    Case 2
          If Index = 0 Then
              'Ha perdido el foco el primer campo
              'cuando estamos con cuentas. Luego tengo que
              'ver si ha puesto el punto
              Text1(0).Text = RellenaCodigoCuenta(Text1(0).Text)
          End If
    Case 3
        LostFocusTipo3 Index
    End Select
End Sub

Private Sub LostFocusTipo0(Campo As Integer)

End Sub

Private Sub LostFocusTipo3(Campo As Integer)
If Campo = 0 Then
    Text1(0).Text = UCase(Text1(0).Text)
        'Reutilizacion de variables
    vAnchoCampos = "idsubcos"
    Sql = DevuelveDesdeBD("nomccost", "cabccost", "codccost", Text1(0).Text, "T", vAnchoCampos)
    sng = 0
    If Sql <> "" Then
        If vAnchoCampos = "0" Then  'Tiene sub centro de coste
            Text1(1).Text = Sql
        Else
            Sql = "El subcentro de coste tiene reparto"
            sng = 2
        End If
    Else
        Sql = "No existe el subcentro de coste para: " & Text1(0).Text
        sng = 2
    End If
    If sng > 1 Then
        MsgBox Sql, vbExclamation
        Text1(0).Text = ""
        Text1(1).Text = ""
    End If
Else
    If Campo = 2 Then
        If Not IsNumeric(Text1(2).Text) Then
            MsgBox "El % de reparto debe de ser numérico", vbExclamation
            Exit Sub
        End If
        Text1(2).Text = TransformaPuntosComas(Text1(2).Text)
        Text1(2).Text = Format(Text1(2).Text, "0.00")
    End If
End If
End Sub

Private Function DatosOk(ParaBusqueda As Boolean) As String
Dim Cad As String
Cad = ""
Select Case vTIPO
Case 0
    'Lineas factura
    Cad = DatosOkConcepto(ParaBusqueda)
Case 1
    Cad = DatosOkDiario(ParaBusqueda)
Case 2
    'Es para COLCUENTAS
    'Luego ponga lo que ponga es correcto
Case 3
    Cad = DatosOkCCoste(ParaBusqueda)
Case Else
    Cad = "Error en el tipo (vTIPO incorrecto)"
End Select
DatosOk = Cad
End Function





'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'
'
'                   Todas estas lineas son a mano#
'                   Dependera de donde se llaman
'                   para realizar unas cosas u otras
'
'
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Public Function DatosOkConcepto(ParaBusqueda As Boolean) As String
Dim i As Integer

DatosOkConcepto = ""

For i = 0 To Text1.Count - 1
    Text1(i).Text = Trim(Text1(i).Text)
Next i


If Not ParaBusqueda Then
    For i = 0 To Text1.Count - 2   'Menos dos pq el ultimo no lo utiliza
       If Text1(i).Text = "" Then
            DatosOkConcepto = "Ningún dato puede estar vacio"
            Exit Function
        End If
    Next i
    If Combo1.ListIndex < 0 Then
        DatosOkConcepto = "Seleccione un tipo de concepto"
        Exit Function
    End If
    If Not IsNumeric(Text1(0).Text) Then
        DatosOkConcepto = "El campo cod debe de ser numérico"
        Exit Function
    End If
    i = CInt(Text1(0).Text)
    If i < 0 Or i > 1000 Then
        DatosOkConcepto = "El campo cod debe de estar entre 0 y 999"
        Exit Function
    End If
End If


'Comprobamos que el clave primaria
If vModo = 0 Then
    Set Rs = New ADODB.Recordset
    Sql = "Select * from conceptos where codconce=" & Text1(0).Text
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DatosOkConcepto = "Ya existe un registro para codigo: " & Text1(0).Text
    End If
    Set Rs = Nothing
End If
End Function


Private Sub PonvalorCombo(indice As Integer)
Dim i As Integer
For i = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(i) = indice Then
        Combo1.ListIndex = i
        Exit For
    End If
Next i
End Sub


'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------

Private Sub CargaCombo()
Combo1.Clear
Select Case vTIPO
Case 0
    'Conceptos
    Combo1.AddItem "Debe"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "Haber"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
    Combo1.AddItem "Decide asisento"
    Combo1.ItemData(Combo1.NewIndex) = 3
    
    
Case 2
    'Buscar cuentas
    Combo1.AddItem "S"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "N"
    Combo1.ItemData(Combo1.NewIndex) = 1
End Select

End Sub
Private Function InsertarModificarConcepto() As Boolean
On Error GoTo EInsertarModificarCliente
    InsertarModificarConcepto = False
    If vModo = 0 Then
        Sql = "INSERT INTO conceptos VALUES ("
        Sql = Sql & Text1(0).Text & ",'"
        Sql = Sql & Text1(1).Text & "',"
        Sql = Sql & Combo1.ItemData(Combo1.ListIndex) & ")"
        Else
        Sql = "UPDATE conceptos Set nomconce = '" & Text1(1).Text
        Sql = Sql & "', tipoconce = " & Combo1.ItemData(Combo1.ListIndex)
        Sql = Sql & " WHERE codconce =" & Text1(0).Text
    End If
    Conn.Execute Sql
    InsertarModificarConcepto = True
    Exit Function
EInsertarModificarCliente:
    MuestraError Err.Number, "Insertar/Modificar Cliente" & vbCrLf & Err.Description
End Function



Private Function InsertarModificarDiario() As Boolean
On Error GoTo EInsertarModificarDiario
    InsertarModificarDiario = False
    If vModo = 0 Then
        Sql = "INSERT INTO tiposdiario VALUES ("
        Sql = Sql & Text1(0).Text & ",'"
        Sql = Sql & Text1(1).Text & "')"
        Else
        Sql = "UPDATE tiposdiario Set desdiari = '" & Text1(1).Text
        Sql = Sql & "' WHERE numdiari =" & Text1(0).Text
    End If
    Conn.Execute Sql
    InsertarModificarDiario = True
    Exit Function
EInsertarModificarDiario:
    MuestraError Err.Number, "Insertar/Modificar Diario" & vbCrLf & Err.Description
End Function




Private Function InsertarModificarCCoste() As Boolean
On Error GoTo EInsertarModificarCCoste
    InsertarModificarCCoste = False
    If vModo = 0 Then
        Sql = "INSERT INTO linccost VALUES ("
        Sql = Sql & "'" & vFac & "',"
        Sql = Sql & vLinea & ",'"
        Sql = Sql & Text1(0).Text & "',"
        Sql = Sql & TransformaComasPuntos(Text1(2).Text) & ")"
        Else
        Sql = "UPDATE linccost Set subccost = '" & Text1(0).Text
        Sql = Sql & "', porccost = " & TransformaComasPuntos(Text1(2).Text)
        Sql = Sql & " WHERE codccost ='" & vFac & "'"
        Sql = Sql & " AND linscost=" & vLinea
    End If
    Conn.Execute Sql
    InsertarModificarCCoste = True
    Exit Function
EInsertarModificarCCoste:
    MuestraError Err.Number, "Insertar/Modificar Cliente" & vbCrLf & Err.Description
End Function



Public Function DatosOkDiario(ParaBusqueda As Boolean) As String
Dim i As Integer

DatosOkDiario = ""

For i = 0 To Text1.Count - 1
    Text1(i).Text = Trim(Text1(i).Text)
Next i


If Not ParaBusqueda Then
    For i = 0 To Text1.Count - 2 'El ultimo no lo utilza
       If Text1(i).Text = "" Then
            DatosOkDiario = "Ningún dato puede estar vacio"
            Exit Function
        End If
    Next i
End If
If Text1(0).Text <> "" And vModo < 2 Then
    If Not IsNumeric(Text1(0).Text) Then
        DatosOkDiario = "El campo Nº diario debe de er numérico"
        Exit Function
    End If
End If

'Comprobamos que el clave primaria
If vModo = 0 Then
    Set Rs = New ADODB.Recordset
    Sql = "Select * from tiposdiario where numdiari=" & Text1(0).Text
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DatosOkDiario = "Ya existe un registro para Nº diario: " & Text1(0).Text
    End If
    Set Rs = Nothing
End If
End Function


Public Function DatosOkCCoste(ParaBusqueda As Boolean) As String
Dim i As Integer
Dim Por As Single

DatosOkCCoste = ""

For i = 0 To Text1.Count - 1
    Text1(i).Text = Trim(Text1(i).Text)
Next i


If Not ParaBusqueda Then
    For i = 0 To Text1.Count - 2   'Menos dos pq el ultimo no lo utiliza
       If Text1(i).Text = "" Then
            DatosOkCCoste = "Ningún dato puede estar vacio"
            Exit Function
        End If
    Next i
    
    If Not IsNumeric(Text1(2).Text) Then
        DatosOkCCoste = "El campo % reparto debe ser numérico"
        Exit Function
    End If
    sng = CSng(Text1(2).Text)
    If sng < 0 Or sng > 100 Then
        DatosOkCCoste = "El campo % reparto debe de estar entre 0 y 100"
        Exit Function
    End If
    
    If sng > MaxValor Then DatosOkCCoste = "El valor máximo permitido es : " & MaxValor
End If
    
End Function


Private Sub ObtenerMaxValor()
    Set Rs = New ADODB.Recordset
    Sql = "SELECT * FROM linccost where codccost='" & vFac & "'"
    If vModo = 1 Then
        'Como estamos seleccionando todos menos la linea
        Sql = Sql & " AND linscost <> " & vLinea
    End If

    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    sng = 0
    While Not Rs.EOF
        sng = sng + Rs!porccost
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If sng >= 100 Then MsgBox "El valor total del % de reparto de los subcentros de coste excede o es igual a 100", vbCritical
    
    MaxValor = Round(100 - sng, 2)

End Sub
