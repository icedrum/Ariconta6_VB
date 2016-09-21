VERSION 5.00
Begin VB.Form frm1LineaDe2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   2715
      Picture         =   "frm1LineaDe2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   -15
      Width           =   315
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   315
      Left            =   2325
      Picture         =   "frm1LineaDe2.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   -15
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   1
      Left            =   825
      TabIndex        =   1
      Top             =   0
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   800
   End
End
Attribute VB_Name = "frm1LineaDe2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vTIPO As Byte   'Para saber que es lo quien lo ha llamado

Public vCamposHabilitados As String
Public vBotonesVisibles As String
Public vAnchoCampos As String
Public vModo As Byte '0.- Nuevo  1.- Modificar  2.- Buscar
Public vTop As Long
Public vLeft As Long
Public vCadena As String
Public vFac As Long

Dim SQL As String
Dim PrimeraVez As Boolean
Dim Rs As Recordset
Dim sng As Single





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
        'Llamado desde clientes
        If vModo < 2 Then
            If Not InsertarModificarCliente Then Exit Sub
            'Correcto
            FormularioHijoModificado = True
        Else
            'Buscar, generar la cadena de busqueda
            Rc = SeparaCampoBusqueda("N", "Clientes.Codcli", Text1(0).Text, Cad)
            If Rc = 0 Then
               CadenaDevueltaFormHijo = Cad
            End If
            'El campo NOMBRE
            Rc = SeparaCampoBusqueda("T", "Clientes.nomcli", Text1(1).Text, Cad)
            If Rc = 0 Then
                If CadenaDevueltaFormHijo <> "" Then CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & " AND "
                CadenaDevueltaFormHijo = CadenaDevueltaFormHijo & Cad
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

Top = vTop
Left = vLeft
PrimeraVez = True

sng = 15
'Situamos todos los campos, en funcion de si hay botones visbles
'El primer o es fijo, de ancho determinado
    aux = CSng(RecuperaValor(vAnchoCampos, 1))
    Text1(0).Width = aux
    sng = sng + aux + 15
    
'A partir del segundo ya vamos mirando
'Vemos que camos van en blancos y esta habilidados y cuales van en pastel y no
'For i = 1 To 5
i = 1
    aux = CSng(RecuperaValor(vAnchoCampos, i + 1))
    'bol = RecuperaValor(vBotonesVisibles, i + 1) = "S"
    'El de dos campos no tiene botones
    'cmdBoton(i - 1).Visible = bol
    'If bol Then
    '    cmdBoton(i - 1).Left = sng
    '    sng = sng + 330 ''315 del ancho del boton y 15 de margen
    '    aux = aux - 330 - 15 '315 del ancho del boton y 15 de margen
    'End If
    Text1(i).Left = sng
    Text1(i).Width = aux
    sng = sng + aux + 15
'Next i
    


'Fijamos el ancho total
cmdAceptar.Left = sng
cmdCancelar.Left = sng + 330
'y la posicion de los aceptar y cancelar
sng = sng + 675 'Ancho 2 botones +espacios intermedios
Me.Width = sng
End Sub




Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)

'Por si acaso se puede llamar desde varios vTIPO
'Ponemos los iF , o Select case correspondiente
    If vTIPO = 0 Then
       Text1(1).Text = RecuperaValor(CadenaSeleccion, 1)
       Text1(2).Text = RecuperaValor(CadenaSeleccion, 2)
       Text1(4).Text = RecuperaValor(CadenaSeleccion, 3)
       Text1(5).Text = ""
       If Text1(3).Text <> "" Then
                sng = CSng(Text1(4).Text) * Val(Text1(3).Text)
                Text1(5).Text = Format(sng, "0.00")
       End If
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
    
    If vTIPO = 0 Then LostFocusTipo0 Index
End Sub


Private Function DatosOk(ParaBusqueda As Boolean) As String
Dim Cad As String
Cad = ""
Select Case vTIPO
Case 0
    'Lineas factura
    Cad = DatosOkCliente(ParaBusqueda)
    
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

Public Function DatosOkCliente(ParaBusqueda As Boolean) As String
Dim i As Integer

DatosOkCliente = ""

For i = 0 To Text1.Count - 1
    Text1(i).Text = Trim(Text1(i).Text)
Next i


If Not ParaBusqueda Then
    For i = 0 To Text1.Count - 1
       If Text1(i).Text = "" Then
            DatosOkCliente = "ningún dato puede estar vacio"
            Exit Function
        End If
    Next i
End If



'Comprobamos que el clave primaria
If vModo = 0 Then
    Set Rs = New ADODB.Recordset
    
    Set Rs = Nothing
End If
End Function




Private Sub LostFocusTipo0(Index As Integer)

End Sub



'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------


Private Function InsertarModificarCliente() As Boolean
On Error GoTo EInsertarModificarCliente
    InsertarModificarCliente = False
    If vModo = 0 Then
        SQL = "INSERT INTO Clientes VALUES ("
        SQL = SQL & Text1(0).Text & ",'"
        SQL = SQL & Text1(1).Text & "')"
        Else
        SQL = "UPDATE Clientes Set nomcli = '" & Text1(1).Text
        SQL = SQL & "' WHERE codcli =" & Text1(0).Text
    End If
    Conn.Execute SQL
    InsertarModificarCliente = True
    Exit Function
EInsertarModificarCliente:
    MuestraError Err.Number, "Insertar/Modificar Cliente" & vbCrLf & Err.Description
    
End Function
