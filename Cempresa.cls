VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodempre As String 'copia local
Private mvarnomempre As String 'copia local
Private mvarnomresum As String 'copia local
Private mvarnumnivel As Integer 'copia local
Private mvarnumdigi1 As Integer 'copia local
Private mvarnumdigi2 As Integer 'copia local
Private mvarnumdigi3 As Integer 'copia local
Private mvarnumdigi4 As Integer 'copia local
Private mvarnumdigi5 As Integer 'copia local
Private mvarnumdigi6 As Integer 'copia local
Private mvarnumdigi7 As Integer 'copia local
Private mvarnumdigi8 As Integer 'copia local
Private mvarnumdigi9 As Integer 'copia local
Private mvarnumdigi10 As Integer 'copia local

'Campo calculado de numero de digitos en la ULTIMA NIVEL CONTABLE
Private mvarDigitosUltimoNivel As Byte


'Campo que nos indicara si esta empresa tiene teseoreria
Private mvarTieneTesoreria As Boolean
'Campo que nos indicara si esta empresa tiene contabilidad
Private mvarTieneContabilidad As Boolean

'Es gran Empresa. Sera a efectos
Private mvarGranEmpresa As Boolean



'NIF empresa. Esta en la tabla empresa2
Private mvarNIF As String

Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset
Private RC As Byte


'------------------------------------------------
'Propiedades del modulo Cempresa
'------------------------------------------------


Public Property Let codempre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codempre = 5
     mvarcodempre = vData
End Property

Public Property Get codempre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codempre
     codempre = mvarcodempre
End Property


Public Property Let nomempre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nomempre = 5
     mvarnomempre = vData
End Property

Public Property Get nomempre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     nomempre = mvarnomempre
End Property


Public Property Let nomresum(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nomresum = 5
     mvarnomresum = vData
End Property

Public Property Get nomresum() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomresum
     nomresum = mvarnomresum
End Property




Public Property Let numnivel(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numnivel = 5
     mvarnumnivel = vData
End Property

Public Property Get numnivel() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numnivel
     numnivel = mvarnumnivel
End Property


Public Property Let numdigi1(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi1 = 5
     mvarnumdigi1 = vData
End Property

Public Property Get numdigi1() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi1
     numdigi1 = mvarnumdigi1
End Property


Public Property Let numdigi2(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi2 = 5
     mvarnumdigi2 = vData
End Property

Public Property Get numdigi2() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi2
     numdigi2 = mvarnumdigi2
End Property


Public Property Let numdigi3(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi3 = 5
     mvarnumdigi3 = vData
End Property

Public Property Get numdigi3() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi3
     numdigi3 = mvarnumdigi3
End Property


Public Property Let numdigi4(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi4 = 5
     mvarnumdigi4 = vData
End Property

Public Property Get numdigi4() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi4
     numdigi4 = mvarnumdigi4
End Property


Public Property Let numdigi5(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi5 = 5
     mvarnumdigi5 = vData
End Property

Public Property Get numdigi5() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi5
     numdigi5 = mvarnumdigi5
End Property


Public Property Let numdigi6(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi6 = 5
     mvarnumdigi6 = vData
End Property

Public Property Get numdigi6() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi6
     numdigi6 = mvarnumdigi6
End Property


Public Property Let numdigi7(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi7 = 5
     mvarnumdigi7 = vData
End Property

Public Property Get numdigi7() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi7
     numdigi7 = mvarnumdigi7
End Property


Public Property Let numdigi8(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi8 = 5
     mvarnumdigi8 = vData
End Property

Public Property Get numdigi8() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi8
     numdigi8 = mvarnumdigi8
End Property


Public Property Let numdigi9(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi9 = 5
     mvarnumdigi9 = vData
End Property

Public Property Get numdigi9() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi9
     numdigi9 = mvarnumdigi9
End Property


Public Property Let numdigi10(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.numdigi10 = 5
     mvarnumdigi10 = vData
End Property

Public Property Get numdigi10() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.numdigi10
     numdigi10 = mvarnumdigi10
End Property

'Solo son de lectura
Public Property Get DigitosUltimoNivel() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     DigitosUltimoNivel = mvarDigitosUltimoNivel
End Property



Public Property Get TieneTesoreria() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     TieneTesoreria = mvarTieneTesoreria
End Property

Public Property Get TieneContabilidad() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nomempre
     TieneContabilidad = mvarTieneContabilidad
End Property


Public Property Get NIF() As String
    NIF = mvarNIF
End Property

'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------



'****   AGREGAR  ***
Public Function Agregar() As Byte
    On Error GoTo Err1

    SQL = "INSERT INTO empresa (codempre, nomempre, nomresum, numnivel, numdigi1,"
    SQL = SQL & "numdigi2, numdigi3, numdigi4, numdigi5, numdigi6, numdigi7,"
    SQL = SQL & " numdigi8, numdigi9, numdigi10) VALUES ("

    'Datos
    SQL = SQL & mvarcodempre
    SQL = SQL & ",'" & mvarnomempre
    SQL = SQL & "','" & mvarnomresum
    SQL = SQL & "'," & mvarnumnivel
    SQL = SQL & "," & mvarnumdigi1
    SQL = SQL & "," & mvarnumdigi2
    SQL = SQL & "," & mvarnumdigi3
    SQL = SQL & "," & mvarnumdigi4
    SQL = SQL & "," & mvarnumdigi5
    SQL = SQL & "," & mvarnumdigi6
    SQL = SQL & "," & mvarnumdigi7
    SQL = SQL & "," & mvarnumdigi8
    SQL = SQL & "," & mvarnumdigi9
    SQL = SQL & "," & mvarnumdigi10 & ")"
    
    Conn.Execute SQL
    
    Exit Function
Err1:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Agregar = 1
End Function




'****   LEER  ***
Public Function Leer(Optional CONTA As String) As Byte
    On Error GoTo Err3
    SQL = "SELECT * FROM "
    If CONTA <> "" Then
        SQL = SQL & "ariconta" & CONTA & "."
    End If
    SQL = SQL & "Empresa "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If Rs.EOF Then
        Leer = 1
        Rs.Close
        'ELSE
        Else
            Mens = "Datos basicos"
            mvarcodempre = Rs!codempre
            mvarnomempre = Rs!nomempre
            mvarnomresum = DBLet(Rs!nomresum)
            Mens = "Datos contables"
            mvarnumnivel = Rs!numnivel
            mvarnumdigi1 = Rs!numdigi1
            mvarnumdigi2 = DBLet(Rs!numdigi2, "N")
            mvarnumdigi3 = DBLet(Rs!numdigi3, "N")
            mvarnumdigi4 = DBLet(Rs!numdigi4, "N")
            mvarnumdigi5 = DBLet(Rs!numdigi5, "N")
            mvarnumdigi6 = DBLet(Rs!numdigi6, "N")
            mvarnumdigi7 = DBLet(Rs!numdigi7, "N")
            mvarnumdigi8 = DBLet(Rs!numdigi8, "N")
            mvarnumdigi9 = DBLet(Rs!numdigi9, "N")
            mvarnumdigi10 = DBLet(Rs!numdigi10, "N")
            'Fijamos la longitud maxima de digitos en ctas de ultimo nivel
            FijaLongitudMaxima
            
            Rs.Close
            
            'Leemos si tiene tesoreria
            LeerTesoreria
            
            'Vemos si es gran empresa o no
            LeerGranEmpresa
            
            
            mvarNIF = DevuelveDesdeBD("nifempre", "empresa2", "1", "1")
            
            
            Leer = 0
    End If

    
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    Leer = 1
End Function


'****   MODIFICAR  ***
Public Function Modificar() As Byte
    On Error GoTo Err4


    SQL = "UPDATE empresa SET "
    SQL = SQL & " nomempre = '" & mvarnomempre
    SQL = SQL & "', nomresum = '" & mvarnomresum
    SQL = SQL & "', numnivel =" & mvarnumnivel
    SQL = SQL & ", numdigi1 =" & mvarnumdigi1
    SQL = SQL & ", numdigi2 =" & mvarnumdigi2
    SQL = SQL & ", numdigi3 =" & mvarnumdigi3
    SQL = SQL & ", numdigi4 =" & mvarnumdigi4
    SQL = SQL & ", numdigi5 =" & mvarnumdigi5
    SQL = SQL & ", numdigi6 =" & mvarnumdigi6
    SQL = SQL & ", numdigi7 =" & mvarnumdigi7
    SQL = SQL & ", numdigi8 =" & mvarnumdigi8
    SQL = SQL & ", numdigi9 =" & mvarnumdigi9
    SQL = SQL & ", numdigi10 =" & mvarnumdigi10
    SQL = SQL & " WHERE codempre =" & mvarcodempre & ";"
    
    Conn.Execute SQL
    
    
    'Modificamos en la tabla Usuarios.empresa
    SQL = "UPDATE Usuarios.empresasariconta SET nomempre= '" & mvarnomempre
    SQL = SQL & "', nomresum= '" & mvarnomresum
    SQL = SQL & "' WHERE codempre=" & mvarcodempre
    Conn.Execute SQL
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    Modificar = 1
End Function


'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------



Private Sub FijaLongitudMaxima()
Dim i As Byte
        Select Case mvarnumnivel
        Case 1
                i = mvarnumdigi1
        Case 2
                i = mvarnumdigi2
        Case 3
                i = mvarnumdigi3
        Case 4
                i = mvarnumdigi4
        Case 5
                i = mvarnumdigi5
        Case 6
                i = mvarnumdigi6
        Case 7
                i = mvarnumdigi7
        Case 8
                i = mvarnumdigi8
        Case 9
                i = mvarnumdigi9
        Case 10
                i = mvarnumdigi10
        End Select
        mvarDigitosUltimoNivel = i
End Sub




Private Sub LeerTesoreria()

On Error GoTo EmvarTieneTesoreria


    mvarTieneTesoreria = False
    mvarTieneContabilidad = False
    SQL = "Select tesor from usuarios.empresasariconta where codempre = " & DBSet(mvarcodempre, "N")
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            mvarTieneTesoreria = ((Rs.Fields(0) = 1) Or (Rs.Fields(0) = 2))
            mvarTieneContabilidad = (Rs.Fields(0) <> 2)
        End If
    End If
    Rs.Close
    
EmvarTieneTesoreria:
    If Err.Number <> 0 Then Err.Clear
    Set Rs = Nothing
End Sub



Private Sub LeerGranEmpresa()
    
    On Error GoTo ELeerGranEmpresa
    
    mvarGranEmpresa = False
    Set Rs = New ADODB.Recordset
    SQL = "Select codmacta from cuentas where apudirec='S' and (codmacta like '8%' or codmacta like '9%') "
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then mvarGranEmpresa = True
    End If
    Rs.Close
    
ELeerGranEmpresa:
    If Err.Number <> 0 Then Err.Clear
    Set Rs = Nothing
    
End Sub

Public Sub SoloTesoreria()

On Error GoTo eSoloTesoreria

    mvarTieneTesoreria = True
    mvarTieneContabilidad = False

eSoloTesoreria:
    If Err.Number <> 0 Then Err.Clear
    Set Rs = Nothing

End Sub
