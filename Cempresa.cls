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
Private mvarNombreEmpresaOficial As String
Private mvarImagenFondo As Byte





Private Sql As String  'Para la sentencia
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


Public Property Get NombreEmpresaOficial() As String
    NombreEmpresaOficial = mvarNombreEmpresaOficial
End Property



Public Property Let ImagenFondo(ByVal vData As Byte)
     mvarImagenFondo = vData
End Property

Public Property Get ImagenFondo() As Byte
     ImagenFondo = mvarImagenFondo
End Property




'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------



'****   AGREGAR  ***
Public Function Agregar() As Byte
    On Error GoTo Err1

    Sql = "INSERT INTO empresa (codempre, nomempre, nomresum, numnivel, numdigi1,"
    Sql = Sql & "numdigi2, numdigi3, numdigi4, numdigi5, numdigi6, numdigi7,"
    Sql = Sql & " numdigi8, numdigi9, numdigi10) VALUES ("

    'Datos
    Sql = Sql & mvarcodempre
    Sql = Sql & ",'" & mvarnomempre
    Sql = Sql & "','" & mvarnomresum
    Sql = Sql & "'," & mvarnumnivel
    Sql = Sql & "," & mvarnumdigi1
    Sql = Sql & "," & mvarnumdigi2
    Sql = Sql & "," & mvarnumdigi3
    Sql = Sql & "," & mvarnumdigi4
    Sql = Sql & "," & mvarnumdigi5
    Sql = Sql & "," & mvarnumdigi6
    Sql = Sql & "," & mvarnumdigi7
    Sql = Sql & "," & mvarnumdigi8
    Sql = Sql & "," & mvarnumdigi9
    Sql = Sql & "," & mvarnumdigi10 & ")"
    
    Conn.Execute Sql
    
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
    Sql = "SELECT * FROM "
    If CONTA <> "" Then
        Sql = Sql & "ariconta" & CONTA & "."
    End If
    Sql = Sql & "empresa "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
                
            'Valores en empres2   NIF, nombre oficial  fondo
            ValoresEmpresa2
            
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


Private Sub ValoresEmpresa2()
    On Error GoTo eValoresEmpresa2
    mvarNombreEmpresaOficial = ""
    Set Rs = Nothing
    Set Rs = New ADODB.Recordset

            
    Sql = "Select * from empresa2 "
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    mvarNIF = DBLet(Rs!nifempre, "T")
    mvarNombreEmpresaOficial = DBLet(Rs!NomEmpresaOficial, "T")
    mvarImagenFondo = 0
    mvarImagenFondo = DBLet(Rs!ImagenFondo, "N")
    Rs.Close
    
eValoresEmpresa2:
    If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
    Set Rs = Nothing
End Sub


'****   MODIFICAR  ***
Public Function Modificar() As Byte
    On Error GoTo Err4


    Sql = "UPDATE empresa SET "
    Sql = Sql & " nomempre = '" & mvarnomempre
    Sql = Sql & "', nomresum = '" & mvarnomresum
    Sql = Sql & "', numnivel =" & mvarnumnivel
    Sql = Sql & ", numdigi1 =" & mvarnumdigi1
    Sql = Sql & ", numdigi2 =" & mvarnumdigi2
    Sql = Sql & ", numdigi3 =" & mvarnumdigi3
    Sql = Sql & ", numdigi4 =" & mvarnumdigi4
    Sql = Sql & ", numdigi5 =" & mvarnumdigi5
    Sql = Sql & ", numdigi6 =" & mvarnumdigi6
    Sql = Sql & ", numdigi7 =" & mvarnumdigi7
    Sql = Sql & ", numdigi8 =" & mvarnumdigi8
    Sql = Sql & ", numdigi9 =" & mvarnumdigi9
    Sql = Sql & ", numdigi10 =" & mvarnumdigi10
    Sql = Sql & " WHERE codempre =" & mvarcodempre & ";"
    
    Conn.Execute Sql
    
    
    'Modificamos en la tabla usuarios.empresa
    Sql = "UPDATE usuarios.empresasariconta SET nomempre= '" & mvarnomempre
    Sql = Sql & "', nomresum= '" & mvarnomresum
    Sql = Sql & "' WHERE codempre=" & mvarcodempre
    Conn.Execute Sql
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
Dim I As Byte
        Select Case mvarnumnivel
        Case 1
                I = mvarnumdigi1
        Case 2
                I = mvarnumdigi2
        Case 3
                I = mvarnumdigi3
        Case 4
                I = mvarnumdigi4
        Case 5
                I = mvarnumdigi5
        Case 6
                I = mvarnumdigi6
        Case 7
                I = mvarnumdigi7
        Case 8
                I = mvarnumdigi8
        Case 9
                I = mvarnumdigi9
        Case 10
                I = mvarnumdigi10
        End Select
        mvarDigitosUltimoNivel = I
End Sub




Private Sub LeerTesoreria()

On Error GoTo EmvarTieneTesoreria


    mvarTieneTesoreria = False
    mvarTieneContabilidad = False
    Sql = "Select tesor from usuarios.empresasariconta where codempre = " & DBSet(mvarcodempre, "N")
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    Sql = "Select codmacta from cuentas where apudirec='S' and (codmacta like '8%' or codmacta like '9%') "
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
