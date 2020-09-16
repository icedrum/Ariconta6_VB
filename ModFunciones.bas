Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const FormatoFechaHora = "yyyy-mm-dd hh:nn:ss"
Public Const ValorNulo = "Null"

Public Function CompForm(ByRef formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Carga As Boolean
    Dim Correcto As Boolean
       
    Dim HayCamposIncorrectos As Boolean
    Dim CampoIncorrecto As String
    
    HayCamposIncorrectos = False
    CampoIncorrecto = ""
       
    CompForm = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control, True)
'                If Not Correcto Then Exit Function
                If Not Correcto Then
                    Control.BackColor = vbErrorColor
                    HayCamposIncorrectos = True
                    CampoIncorrecto = Control.Name
                    If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                Else
                    Control.BackColor = vbWhite
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
'                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
'                            Exit Function
                        Control.BackColor = vbErrorColor
                        HayCamposIncorrectos = True
                        CampoIncorrecto = Control.Name
                        If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                    Else
                        Control.BackColor = vbWhite
                    End If
                End If
            End If
        End If
    Next Control
    
    If HayCamposIncorrectos Then
        MsgBox "Revise datos obligatorios o incorrectos", vbExclamation
    End If
    CompForm = Not HayCamposIncorrectos
    
End Function

Public Function CompForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Carga As Boolean
Dim Correcto As Boolean

    Dim HayCamposIncorrectos As Boolean
    Dim CampoIncorrecto As String
    
    HayCamposIncorrectos = False
    CampoIncorrecto = ""



    CompForm2 = False
    Set mTag = New CTag
    For Each Control In formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                Carga = mTag.Cargar(Control)
                
                
                If Carga = True Then
                   
                    Correcto = mTag.Comprobar(Control, True)
'                    If Not Correcto Then Exit Function
                    If Not Correcto Then
                        Control.BackColor = vbErrorColor
                        HayCamposIncorrectos = True
                        CampoIncorrecto = Control.Name
                        If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                    Else
                        If Control.Tag <> "" Then Control.BackColor = vbWhite
                    End If
    
                Else
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                End If
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.visible = True Then
                'Comprueba que los campos estan bien puestos
                If Control.Tag <> "" Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        Carga = mTag.Cargar(Control)
                        If Carga = False Then
                            MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                            Exit Function
        
                        Else
                            If mTag.Vacio = "N" And Control.ListIndex < 0 Then
'                                    MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
'                                    Exit Function
                                Control.BackColor = vbErrorColor
                                HayCamposIncorrectos = True
                                CampoIncorrecto = Control.Name
                                If IsArray(Control) Then CampoIncorrecto = CampoIncorrecto & "(" & Control.Index & ")"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next Control
    
    If HayCamposIncorrectos Then
        MsgBox "Revise datos obligatorios o incorrectos", vbExclamation
    End If
    CompForm2 = Not HayCamposIncorrectos
'    CompForm2 = True
End Function

Public Sub Limpiar(ByRef formulario As Form)
    Dim Control As Object
    
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub


Public Function CampoSiguiente(ByRef formulario As Form, Valor As Integer) As Control
Dim Fin As Boolean
Dim Control As Object

On Error GoTo ECampoSiguiente

    'Debug.Print "Llamada:  " & Valor
    'Vemos cual es el siguiente
    Do
        Valor = Valor + 1
        For Each Control In formulario.Controls
            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
            'Si es texto monta esta parte de sql
            If Control.TabIndex = Valor Then
                    Set CampoSiguiente = Control
                    Fin = True
                    Exit For
            End If
        Next Control
        If Not Fin Then
            Valor = -1
        End If
    Loop Until Fin
    Exit Function
ECampoSiguiente:
    Set CampoSiguiente = Nothing
    Err.Clear
End Function


Private Function ValorParaSQL(Valor, ByRef vTag As CTag) As String
Dim Dev As String
Dim d As Single
Dim i As Integer
Dim v
    Dev = ""
    If Valor <> "" Then
        Select Case vTag.TipoDato
        Case "N"
            v = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'Ademas de la coma lleva puntos
                    v = ImporteFormateado(CStr(Valor))
                    Valor = v
                Else
                    v = CCur(Valor)
                    Valor = v
                End If
            Else
         
            End If
            Dev = TransformaComasPuntos(CStr(Valor))
            
        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
            
        Case "FH"
            Dev = "'" & Format(Valor, "yyyy-mm-dd hh:mm:ss") & "'"
        
        Case Else
            Dev = "'" & Valor & "'"
        End Select
        
    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vTag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function

Public Function InsertarDesdeForm(ByRef formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Izda As String
    Dim Der As String
    Dim cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                    
                        'Parte VALUES
                        cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.Columna & ""
                If Control.Value = 1 Then
                    cad = "1"
                    Else
                    cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                Der = Der & cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.ListIndex = -1 Then
                        cad = ValorNulo
                        Else
                        cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    
   
    Conn.Execute cad, , adCmdText
    
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
    
End Function

Public Function InsertarDesdeForm2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim Der As String
Dim cad As String
    
    On Error GoTo EInsertarF
    
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm2 = False
    Der = ""
    Izda = ""
    
    For Each Control In formulario.Controls
    
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" Then
                            If Izda <> "" Then Izda = Izda & ","
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.Columna & ""
                        
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Izda <> "" Then Izda = Izda & ","
                    'Access
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.Value = 1 Then
                        cad = "1"
                        Else
                        cad = "0"
                    End If
                    If Der <> "" Then Der = Der & ","
                    If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                    Der = Der & cad
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        If Der <> "" Then Der = Der & ","
                        Der = Der & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & ","
                            Izda = Izda & "" & mTag.Columna & ""
                            cad = Control.Index
                            If Der <> "" Then Der = Der & ","
                            Der = Der & cad
                        End If
                    End If
                End If
            End If
            
        End If
        
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    cad = "INSERT INTO " & mTag.tabla
    cad = cad & " (" & Izda & ") VALUES (" & Der & ");"
    Conn.Execute cad, , adCmdText
    
     ' ### [Monica] 18/12/2006
    CadenaCambio = cad
   
    InsertarDesdeForm2 = True
Exit Function

EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function




Public Function PonerCamposForma(ByRef formulario As Form, ByRef vData As Adodc) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim cad As String
    Dim Valor As Variant
    Dim Campo As String  'Campo en la base de datos
    Dim i As Integer

    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In formulario.Controls
        'TEXTO
        'Debug.Print Control.Tag
        If TypeOf Control Is TextBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.Columna <> "" Then
                        Campo = mTag.Columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(Campo))
                        Else
                            Valor = vData.Recordset.Fields(Campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    Campo = mTag.Columna
                    If mTag.Vacio = "S" Then
                        Valor = DBLet(vData.Recordset.Fields(Campo), mTag.TipoDato)
                    Else
                        Valor = vData.Recordset.Fields(Campo)
                    End If
                    Else
                        Valor = 0
                End If
                Control.Value = Valor
            End If
            
         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Campo = mTag.Columna
                    
                    If mTag.Vacio = "S" Then
                        If IsNull(vData.Recordset.Fields(Campo)) Then
                            Valor = -1
                        Else
                            Valor = vData.Recordset.Fields(Campo)
                        End If
                    Else
                        Valor = vData.Recordset.Fields(Campo)
                    End If
                    i = 0
                    For i = 0 To Control.ListCount - 1
                        If Control.ItemData(i) = Val(Valor) Then
                            Control.ListIndex = i
                            Exit For
                        End If
                    Next i
                    If i = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control
    
    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

Public Function PonerCamposForma2(ByRef formulario As Form, ByRef vData As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim cad As String
Dim Valor As Variant
Dim Campo As String  'Campo en la base de datos
Dim i As Integer
    On Error GoTo EPonerCamposForma2
    
    Set mTag = New CTag
    PonerCamposForma2 = False
    For Each Control In formulario.Controls
        'TEXTO
        If (TypeOf Control Is TextBox) Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    If mTag.Cargado Then
                        'If Control.Index = 29 Then St op
                    
                        'Columna en la BD
                        If mTag.Columna <> "" Then
                            Campo = mTag.Columna
                            If mTag.Vacio = "S" Then
                                Valor = DBLet(vData.Recordset.Fields(Campo))
                            Else
                                Valor = vData.Recordset.Fields(Campo)
                            End If
                            If mTag.Formato <> "" And CStr(Valor) <> "" Then
                                If mTag.TipoDato = "N" Then
                                    'Es numerico, entonces formatearemos y sustituiremos
                                    ' La coma por el punto
                                    cad = Format(Valor, mTag.Formato)
                                    'Antiguo
                                    'Control.Text = TransformaComasPuntos(cad)
                                    'nuevo
                                    Control.Text = cad
                                Else
                                    Control.Text = Format(Valor, mTag.Formato)
                                End If
                            Else
                                Control.Text = Valor
                            End If
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf (TypeOf Control Is CheckBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        Campo = mTag.Columna
                        Valor = vData.Recordset.Fields(Campo)
                    Else
                        Valor = 0
                    End If
                    If IsNull(Valor) Then Valor = 0
                    Control.Value = Valor
                End If
            End If

         'COMBOBOX
         ElseIf (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        Campo = mTag.Columna
                        Valor = DBLet(vData.Recordset.Fields(Campo))
                        i = 0
                        For i = 0 To Control.ListCount - 1
                            If Control.ItemData(i) = Val(Valor) Then
                                Control.ListIndex = i
                                Exit For
                            End If
                        Next i
                        If i = Control.ListCount Then Control.ListIndex = -1
                    End If 'de cargado
                End If
            End If 'de <>""
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Columna en la BD
                        Campo = mTag.Columna
                        Valor = vData.Recordset.Fields(Campo)
                        If IsNull(Valor) Then Valor = 0
                        If Control.Index = Valor Then
                            Control.Value = True
                        Else
                            Control.Value = False
                        End If
                    End If
                End If
            End If
            
        End If
    Next Control

    'Veremos que tal
    PonerCamposForma2 = True
Exit Function
EPonerCamposForma2:
    MuestraError Err.Number, "Poner campos formulario 2. "
End Function





Private Function ObtenerMaximoMinimo(ByRef vSql As String) As String
Dim Rs As Recordset
ObtenerMaximoMinimo = ""
Set Rs = New ADODB.Recordset
Rs.Open vSql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.EOF) Then
        ObtenerMaximoMinimo = CStr(Rs.Fields(0))
    End If
End If
Rs.Close
Set Rs = Nothing
End Function


Public Function ObtenerBusqueda(ByRef formulario As Form) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        cad = " MAX(" & mTag.Columna & ")"
                    Else
                        cad = " MIN(" & mTag.Columna & ")"
                    End If
                    Sql = "Select " & cad & " from " & mTag.tabla
                    Sql = ObtenerMaximoMinimo(Sql)
                    Select Case mTag.TipoDato
                    Case "N"
                        Sql = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(Sql)
                    Case "F"
                        Sql = mTag.tabla & "." & mTag.Columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                    Case Else
                        Sql = mTag.tabla & "." & mTag.Columna & " = '" & Sql & "'"
                    End Select
                    Sql = "(" & Sql & ")"
                End If
            End If
        End If
    Next

    
    
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            'If Control.Text <> "" Then St op
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    Sql = mTag.tabla & "." & mTag.Columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                Aux = Trim(Control.Text)
                'If Control.Text <> "" Then St op
                If Aux <> "" Then
                    If mTag.tabla <> "" Then
                        tabla = mTag.tabla & "."
                        Else
                        tabla = ""
                    End If
                    RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad, mTag.EsClave)
                    If RC = 0 Then
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    cad = Control.ItemData(Control.ListIndex)
                    cad = mTag.tabla & "." & mTag.Columna & " = " & cad
                    If Sql <> "" Then Sql = Sql & " AND "
                    Sql = Sql & "(" & cad & ")"
                End If
            End If
        
        
        'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.Value = 1 Then
                        cad = mTag.tabla & "." & mTag.Columna & " = 1"
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
        End If

        
    Next Control
    ObtenerBusqueda = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function


Public Function ObtenerBusqueda2(ByRef formulario As Form, Optional check As String, Optional opcio As Integer, Optional nom_frame As String) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim cad As String
    Dim Sql As String
    Dim tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda2 = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Aux = ">>" Then
                            cad = " MAX(" & mTag.Columna & ")"
                        Else
                            cad = " MIN(" & mTag.Columna & ")"
                        End If
                        Sql = "Select " & cad & " from " & mTag.tabla
                        Sql = ObtenerMaximoMinimo(Sql)
                        Select Case mTag.TipoDato
                        Case "N"
                            Sql = mTag.tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(Sql)
                        Case "F"
                            Sql = mTag.tabla & "." & mTag.Columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                        Case Else
                            '[Monica]04/03/2013: quito las comillas y pongo el dbset
                            Sql = mTag.tabla & "." & mTag.Columna & " = " & DBSet(Sql, "T") ' & "'"
                        End Select
                        Sql = "(" & Sql & ")"
                    End If
                End If
            End If
        End If
    Next

'++monica: lo he añadido del anterior obtenerbusqueda
    'Recorremos los text en busca del NULL
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga And mTag.Columna <> "" Then

                    Sql = mTag.tabla & "." & mTag.Columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    'Recorremos los textbox
    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
          If Control.Tag <> "" Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    Aux = Trim(Control.Text)
                    If Aux <> "" Then
                        If mTag.tabla <> "" Then
                            tabla = mTag.tabla & "."
                            Else
                            tabla = ""
                        End If
                        RC = SeparaCampoBusqueda(mTag.TipoDato, tabla & mTag.Columna, Aux, cad, mTag.EsClave)
                        If RC = 0 Then
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        End If
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then ' +-+- 12/05/05: canvi de Cèsar, no te sentit passar-li un control que no té TAG +-+-
                mTag.Cargar Control
                If mTag.Cargado Then
                    If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                        If Control.ListIndex > -1 Then
                            cad = Control.ItemData(Control.ListIndex)
                            cad = mTag.tabla & "." & mTag.Columna & " = " & cad
                            If Sql <> "" Then Sql = Sql & " AND "
                            Sql = Sql & "(" & cad & ")"
                        End If
                    End If
                End If
            End If
            
         ElseIf TypeOf Control Is CheckBox Then
            '=============== Añade: Laura, 27/04/05
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    ' añadido 12022007
                    Aux = ""
                    If check <> "" Then
                        tabla = DBLet(Control.Index, "T")
                        If tabla <> "" Then tabla = "(" & tabla & ")"
                        tabla = Control.Name & tabla & "|"
                        If InStr(1, check, tabla, vbTextCompare) > 0 Then Aux = Control.Value
                    Else
                        If Control.Value = 1 Then Aux = "1"
                    End If
                    If Aux <> "" Then
'                    If Control.Value = 1 Then
                        cad = Control.Value
                        cad = mTag.tabla & "." & mTag.Columna & " = " & cad
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & cad & ")"
                    End If
                End If
            End If
            '===================
        End If
    Next Control
    ObtenerBusqueda2 = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda2 = ""
    MuestraError Err.Number, "Obtener búsqueda. " & vbCrLf & Err.Description
End Function



Public Function ModificaDesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.Columna & " = " & Aux & ")"
                             
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    Conn.Execute Aux, , adCmdText



ModificaDesdeFormulario = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function


Public Function ModificaDesdeFormulario2(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                            If mTag.EsClave Then
                                'Lo pondremos para el WHERE
                                 If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                 cadwhere = cadwhere & "(" & mTag.Columna & " = " & Aux & ")"
    
                            Else
                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                            End If
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Control.Value = 1 Then
                        Aux = "TRUE"
                    Else
                        Aux = "FALSE"
                    End If
                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'Esta es para access
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.Columna & " = " & Aux & ")"
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
'
'
'                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            'dentro del WHERE
                              If mTag.EsClave Then
                                  'Lo pondremos para el WHERE
                                   If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                                   cadwhere = cadwhere & "(" & mTag.Columna & " = " & Aux & ")"
                              Else
                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                                  cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                              End If
                        End If
                    End If
                End If
            End If
            
        End If
    Next Control

    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    Conn.Execute Aux, , adCmdText

    ' ### [Monica] 18/12/2006
    CadenaCambio = cadUPDATE

    ModificaDesdeFormulario2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2. " & Err.Description
End Function





Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim cad As String

'Montamos al final: "Cod Diag.|idDiag|N|10·"

ParaGrid = ""
cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Desc <> "" Then
                cad = Desc
            Else
                cad = mTag.Nombre
            End If
            cad = cad & "|"
            cad = cad & mTag.Columna & "|"
            cad = cad & mTag.TipoDato & "|"
            cad = cad & AnchoPorcentaje & "·"
            
                
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            
        ElseIf TypeOf Control Is ComboBox Then
        
        
        End If 'De los elseif
    End If
Set mTag = Nothing
ParaGrid = cad
End If



End Function

'////////////////////////////////////////////////////
' Monta a partir de una cadena devuelta por el formulario
'de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As CTag
Dim cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

ValorDevueltoFormGrid = ""
cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            Aux = RecuperaValor(CadenaDevuelta, Orden)
            If Aux <> "" Then cad = mTag.Columna & " = " & ValorParaSQL(Aux, mTag)
                
            
            
                
        'CheckBOX
       ' ElseIf TypeOf Control Is CheckBox Then
       '
       ' ElseIf TypeOf Control Is ComboBox Then
       '
       '
        End If 'De los elseif
    End If
End If
Set mTag = Nothing
ValorDevueltoFormGrid = cad
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                cad = TransformaPuntosComas(vTex.Text)
                cad = Format(cad, mTag.Formato)
                vTex.Text = cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim i As Integer
Dim J As Integer
Dim CONT As Integer
Dim cad As String

i = 0
CONT = 1
cad = ""
Do
    J = i + 1
    i = InStr(J, Cadena, "|")
    If i > 0 Then
        If CONT = Orden Then
            cad = Mid(Cadena, J, i - J)
            i = Len(Cadena) 'Para salir del bucle
            Else
                CONT = CONT + 1
        End If
    End If
Loop Until i = 0
RecuperaValor = cad
End Function

'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValorNew(ByRef Cadena As String, Separador As String, Orden As Integer) As String
Dim i As Integer
Dim J As Integer
Dim CONT As Integer
Dim cad As String

    i = 0
    CONT = 1
    cad = ""
    Do
        J = i + 1
        i = InStr(J, Cadena, Separador)
        If i > 0 Then
            If CONT = Orden Then
                cad = Mid(Cadena, J, i - J)
                i = Len(Cadena) 'Para salir del bucle
                Else
                    CONT = CONT + 1
            End If
        End If
    Loop Until i = 0
    RecuperaValorNew = cad
End Function





'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function InsertaValor(ByRef Cadena As String, Orden As Integer, SubCadena As String) As String
Dim i As Integer
Dim J As Integer
Dim CONT As Integer
Dim cad As String
Dim Cad2 As String
i = 0
CONT = 1
cad = ""
Do
    J = i + 1
    i = InStr(J, Cadena, "|")
    If i > 0 Then
        If CONT = Orden Then
            cad = Mid(Cadena, J, i - J)
            
            Cad2 = Mid(Cadena, 1, J - 1) & SubCadena & Mid(Cadena, i, Len(Cadena))
            
            i = Len(Cadena) 'Para salir del bucle
            Else
                CONT = CONT + 1
        End If
    End If
Loop Until i = 0
InsertaValor = Cad2
End Function






'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef formulario As Form)
Dim i As Integer
Dim J As Integer


On Error GoTo EPonerOpcionesMenuGeneral


'Añadir, modificar y borrar deshabilitados si no nivel
With formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For i = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(i).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(i).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(i).Enabled = False
            End If
        End If
    Next i
    
    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next
    
    On Error Resume Next
    
    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnNuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnNuevo.Enabled = False
    
    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False
    
    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    On Error GoTo 0
End With




Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim cadUPDATE As String
Dim i As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadwhere = Claves
    'Construimos el SQL
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadwhere
    Conn.Execute Aux, , adCmdText






ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

'Este será el bueno
Public Function ModificaDesdeFormularioClaves2(ByRef formulario As Form, opcio As Integer, nom_frame As String, Clave As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves2 = False
    Set mTag = New CTag
    Aux = ""
    
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" Then
                            'Sea para el where o para el update esto lo necesito
                            Aux = ValorParaSQL(Control.Text, mTag)
                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
                            
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Control.Value = 1 Then
                        Aux = "TRUE"
                    Else
                        Aux = "FALSE"
                    End If
                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'Esta es para access
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If

        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.ListIndex = -1 Then
                            Aux = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            Aux = Control.ItemData(Control.ListIndex)
                        Else
                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        
                        
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        
                    End If
                End If
            End If
            
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If Control.Value Then
                            Aux = Control.Index
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
                    End If
                End If
            End If
            
        End If
    Next Control
  
    Aux = "UPDATE " & mTag.tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & Clave
    Conn.Execute Aux, , adCmdText


    CadenaCambio = cadUPDATE

    ModificaDesdeFormularioClaves2 = True
    Exit Function
    
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar 2 claves. " & Err.Description
End Function










Public Function BLOQUEADesdeFormulario(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                         cadwhere = cadwhere & "(" & mTag.Columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control
    
    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"
        
        'Intenteamos bloquear
        PreparaBloquear
        Conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BLOQUEADesdeFormulario2(ByRef formulario As Form, ByRef ado As Adodc, Optional opcio As Integer, Optional nom_frame As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadwhere As String
Dim AntiguoCursor As Byte
Dim nomcamp As String

    On Error GoTo EBLOQUEADesdeFormulario2
    
    BLOQUEADesdeFormulario2 = False
    Set mTag = New CTag
    Aux = ""
    cadwhere = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If (TypeOf Control Is TextBox) Or (TypeOf Control Is ComboBox) Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        'Sea para el where o para el update esto lo necesito
                        'Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            Aux = ValorParaSQL(CStr(ado.Recordset.Fields(mTag.Columna)), mTag)
                            'Lo pondremos para el WHERE
                             If cadwhere <> "" Then cadwhere = cadwhere & " AND "
                             cadwhere = cadwhere & "(" & mTag.Columna & " = " & Aux & ")"
                        End If
                    End If
                End If
            End If
        End If
    Next Control

    If cadwhere = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
    Else
        Aux = "select * FROM " & mTag.tabla
        Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"

        'Intenteamos bloquear
        PreparaBloquear
        Conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario2 = True
    End If
    
EBLOQUEADesdeFormulario2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla 2"
'        BLOQUEADesdeFormulario2 = False
        TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function BloqueaRegistroForm(ByRef formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim AuxDef As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQ
    BloqueaRegistroForm = False
    Set mTag = New CTag
    Aux = ""
    AuxDef = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        Aux = ValorParaSQL(Control.Text, mTag)
                        AuxDef = AuxDef & Aux & "|"
                    End If
                End If
            End If
        End If
    Next Control
    
    If AuxDef = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        Aux = "Insert into zbloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.tabla
        Aux = Aux & "',""" & AuxDef & """)"
        Conn.Execute Aux
        BloqueaRegistroForm = True
    End If
EBLOQ:
    If Err.Number <> 0 Then
        Aux = ""
        If Conn.Errors.Count > 0 Then
            If Conn.Errors(0).NativeError = 1062 Then
                '¡Ya existe el registro, luego esta bloqueada
                Aux = "BLOQUEO"
            End If
        End If
        If Aux = "" Then
            MuestraError Err.Number, "Bloqueo tabla"
        Else
            MsgBox "Registro bloqueado por otro usuario", vbExclamation
            Err.Clear
        End If
    End If
    Screen.MousePointer = AntiguoCursor
End Function


Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
Dim mTag As CTag
Dim Sql As String

'Solo me interesa la tabla
On Error Resume Next
    Set mTag = New CTag
    mTag.Cargar TextBoxConTag
    If mTag.Cargado Then
        Sql = "DELETE from zbloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.tabla & "'"
        Conn.Execute Sql
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    Set mTag = Nothing
End Function



'----------------------------------------------------------
'----------------------------------------------------------
'
'       Funciones comunes en todos los formularios
'           del tipo KEY PRESS, gotfocus
'

Public Sub KEYpressGnral(KeyAscii As Integer, Modo As Byte, cerrar As Boolean)
'IN: codigo keyascii tecleado, y modo en que esta el formulario
'OUT: si se tiene que cerrar el formulario o no
    cerrar = False
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        CreateObject("WScript.Shell").SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then cerrar = True
        End If
End Sub



Public Sub KEYdown(KeyCode As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
On Error Resume Next
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
            CreateObject("WScript.Shell").SendKeys "+{tab}"
        Case 40 'Desplazamiento Flecha Hacia Abajo
            CreateObject("WScript.Shell").SendKeys "{tab}"
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonFoco(ByRef T As TextBox)
On Error Resume Next
'        T.SelStart = 0
'        T.SelLength = Len(T.Text)
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub BloqueaTXT(ByRef T As TextBox, bloquear As Boolean)
On Error Resume Next
    
    T.Locked = bloquear
    T.BackColor = IIf(bloquear, &H80000018, &H80000005)
    If Err.Number <> 0 Then Err.Clear
End Sub




Public Sub PonerFocoGrid(ByRef DGrid As DataGrid)
    On Error Resume Next
    DGrid.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoCmb(ByRef combo As ComboBox)
On Error Resume Next
    combo.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonerFocoChk(ByRef chk As CheckBox)
On Error Resume Next
    chk.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoLw(ByRef LW As ListView)
On Error Resume Next
    LW.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PonerFocoBtn(ByRef Btn As CommandButton)
    On Error Resume Next
    Btn.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Sub DeseleccionaGrid(ByRef DataGrid1 As DataGrid)
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear

End Sub


Public Sub PonLblIndicador(ByRef L As Label, ByRef AdodcX As Adodc)
    On Error Resume Next
    L.Caption = AdodcX.Recordset.AbsolutePosition & " de " & AdodcX.Recordset.RecordCount
    If Err.Number <> 0 Then
        Err.Clear
        L.Caption = ""
    End If
    If InStr(L.Caption, "-1") > 0 Then L.Caption = ""
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub PonleFoco(Ob As Object)
    On Error Resume Next
    Ob.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



'Devuelve la variable parafijar la altura donde empiezan los txtaux
' y cuadren con el datagrid
Public Function FijarVariableAnc(ByRef DTGRD1 As DataGrid) As Single
Dim i As Integer

    If DTGRD1.Row < 0 Then
        i = 0
        Else
        i = DTGRD1.Row
    End If
    FijarVariableAnc = DTGRD1.RowTop(i) + DTGRD1.top + 15
    
End Function


Public Function BloqueaRegistro(cadTABLA As String, cadwhere As String) As Boolean
Dim Aux As String

    On Error GoTo EBloqueaRegistro
        
    BloqueaRegistro = False
    Aux = "select * FROM " & cadTABLA
    Aux = Aux & " WHERE " & cadwhere & " FOR UPDATE"

    'Intenteamos bloquear
    PreparaBloquear
    Conn.Execute Aux, , adCmdText
    BloqueaRegistro = True

EBloqueaRegistro:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        TerminaBloquear
    End If
End Function

'++++++++++++++++++++++++++++++++++++
'++       FUNCIONES AÑADIDAS
'++++++++++++++++++++++++++++++++++++
Public Function SituarData(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
    On Error GoTo ESituarData

        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        
        'El sql para que se situe en el registro en especial es el siguiente
        vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then
            If vData.Recordset.RecordCount > 0 Then vData.Recordset.MoveFirst
            GoTo ESituarData
        End If
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarData = True
        Exit Function

ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarData = False
End Function

Public Function SituarDataMULTI(ByRef vData As Adodc, vWhere As String, ByRef Indicador As String, Optional NoRefresca As Boolean) As Boolean
'Situa un DataControl en el registo que cumple vwhere
On Error GoTo ESituarData
        'Actualizamos el recordset
        If Not NoRefresca Then vData.Refresh
        'El sql para que se situe en el registro en especial es el siguiente
        Multi_Find vData.Recordset, vWhere
        'vData.Recordset.Find vWhere
        If vData.Recordset.EOF Then GoTo ESituarData
        Indicador = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
        SituarDataMULTI = True
        Exit Function
ESituarData:
        If Err.Number <> 0 Then Err.Clear
        SituarDataMULTI = False
End Function


Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)

    Dim clone_rs As ADODB.Recordset
    Set clone_rs = oRs.Clone
    
    clone_rs.Filter = sCriteria
    
    If clone_rs.EOF Or clone_rs.BOF Then
     oRs.MoveLast
     oRs.MoveNext
    Else
     oRs.Bookmark = clone_rs.Bookmark
    End If
    
    clone_rs.Close
    Set clone_rs = Nothing

End Sub


Public Sub PonerIndicador(ByRef lblIndicador As Label, Modo As Byte, Optional ModoLineas As Byte)
'Pone el titulo del label lblIndicador
    Select Case Modo
        Case 0    'Modo Inicial
            lblIndicador.Caption = ""
        Case 1 'Modo Buscar
            lblIndicador.Caption = "BUSQUEDA"
        Case 2    'Preparamos para que pueda Modificar
            lblIndicador.Caption = ""

        Case 3 'Modo Insertar
            lblIndicador.Caption = "INSERTAR"
        Case 4 'MODIFICAR
            lblIndicador.Caption = "MODIFICAR"
            
        Case 5 'Modo Lineas
            If ModoLineas = 1 Then
                lblIndicador.Caption = "INSERTAR LINEA"
            ElseIf ModoLineas = 2 Then
                lblIndicador.Caption = "MODIFICAR LINEA"
            End If
        Case Else
            lblIndicador.Caption = ""
    End Select
End Sub


Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
'Tipos
'       T
'       N
'       F
'       H
'       FH
'       B
'       S   single O DOUBLE. sINGLE DE MOMENTO.    MAYO 2009
Dim cad As String
Dim ValorNumericoCero As Boolean

    On Error GoTo Error1

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If
    
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N", "S"   'Numero  y  SINGLE
                    
                    If CStr(vData) = "" Then
                        ValorNumericoCero = True
                    
                    Else
                        If Tipo = "S" Then
                            ValorNumericoCero = CSng(vData) = 0
                        Else
                            ValorNumericoCero = CCur(vData) = 0
                        End If
                    End If
                    
                    If ValorNumericoCero Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        If Tipo = "N" Then
                            cad = CStr(ImporteFormateado(CStr(vData)))
                        Else
                            'Sngle
                            cad = CStr(ImporteFormateadoSingle(CStr(vData)))
                        End If
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If

                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If

                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
Error1:
    If Err.Number <> 0 Then MuestraError Err.Number, "Formato para la BD.", Err.Description
End Function

Public Function ImporteFormateadoSingle(Importe As String) As Single
Dim i As Integer

    If Importe = "" Then
        ImporteFormateadoSingle = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateadoSingle = Importe
    End If
End Function

Public Function DevuelveValor(vSql As String) As Variant
'Devuelve el valor de la SQL
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    DevuelveValor = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0).Value) Then DevuelveValor = Rs.Fields(0).Value   'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        DevuelveValor = 0
        Err.Clear
        Conn.Errors.Clear
    End If
End Function

Public Sub PosicionarCombo(ByRef Combo1 As ComboBox, Valor As Integer)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(J) = Valor Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PosicionarCombo2(ByRef Combo1 As ComboBox, Valor As String)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Trim(Combo1.List(J)) = Trim(Valor) Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PosicionarCombo3(ByRef Combo1 As ComboBox, Valor As String)
'Situa el combo en la posicion de un valor concreto
Dim J As Integer

    On Error GoTo EPosCombo
    
    For J = 0 To Combo1.ListCount - 1
        If Mid(Trim(Combo1.List(J)), 1, 3) = Trim(Valor) Then
            Combo1.ListIndex = J
            Exit For
        End If
    Next J

EPosCombo:
    If Err.Number <> 0 Then Err.Clear
End Sub



Public Function ValorCombo(ByRef Cbo As ComboBox) As Integer
'obtiene el valor del combo de la posicion en la q se encuentra

    On Error GoTo EValCombo
    
    If Cbo.ListIndex < 0 Then
        ValorCombo = -1
    Else
        ValorCombo = Cbo.ItemData(Cbo.ListIndex)
    End If
    Exit Function

EValCombo:
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function TextoCombo(ByRef Cbo As ComboBox) As String
'obtiene la descripcion del combo de la posicion en la q se encuentra

    On Error GoTo ErrTexCombo
    
    If Cbo.ListIndex < 0 Then
        TextoCombo = ""
    Else
        TextoCombo = Cbo.List(Cbo.ListIndex)
    End If
    Exit Function

ErrTexCombo:
    If Err.Number <> 0 Then Err.Clear
End Function

Public Sub SituarItemList(ByRef LView As ListView)
'Subir el item seleccionado del listview una posicion
Dim i As Byte, Item As Byte
Dim Aux As String
On Error Resume Next
   
    For i = 1 To LView.ListItems.Count
        If LView.ListItems(i).ToolTipText = vUsu.CadenaConexion Then
            LView.ListItems(i).Selected = True
            LView.ListItems(i).Bold = True
            LView.ListItems(i).ListSubItems(1).Bold = True
            LView.ListItems(i).ListSubItems(2).Bold = True
            LView.ListItems(i).EnsureVisible
            Exit For
        End If
    Next i
    LView.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function QuitarCaracterNULL(vCad As String) As String
Dim i As Integer

    Do
        i = InStr(1, vCad, vbNullChar)
        If i > 0 Then 'Hay null
            vCad = Mid(vCad, 1, i - 1) & Mid(vCad, i + 2)
        End If
    Loop Until i = 0
    QuitarCaracterNULL = vCad
End Function

   

Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then 'Modo 1: Busqueda
            Text.BackColor = vbLightBlue 'vbYellow
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ConseguirFocoLin(ByRef Text As TextBox)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

    With Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ConseguirFocoChk(Modo As Byte)
'Acciones que se realizan en el evento:GotFocus de los TextBox:TxtAux para LINEAS
'en los formularios de Mantenimiento
On Error Resume Next

    If Modo = 0 Or Modo = 2 Then
'        KEYpress 13
        CreateObject("WScript.Shell").SendKeys "{tab}"

        
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function CadenaDesdeHasta(ByRef TD As TextBox, TH As TextBox, Campo As String, TipoCampo As String, Optional NomCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= cadDesde and campo<=cadHasta) "
'para Crystal Report
Dim CadAux As String
Dim cadDesde As String, cadHasta As String
On Error GoTo ErrDH

    cadDesde = "": cadHasta = ""
    If Not TD Is Nothing Then cadDesde = TD.Text
    If Not TH Is Nothing Then cadHasta = TH.Text
    
    Campo = "{" & Campo & "}"
    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            CadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    CadAux = Campo & " >= " & Val(cadDesde)
                Case "T"
                    CadAux = Campo & " >= """ & cadDesde & """"
                Case "F"
                    CadAux = Campo & " >= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
                Case "FH"
                    CadAux = Campo & " >= DateTime(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & "," & Hour(cadDesde) & "," & Minute(cadDesde) & "," & Second(cadDesde) & ")"
                    
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If CadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If cadDesde > cadHasta Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                        End If
                        
                    Case "FH"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                                   
                            CadAux = CadAux & " AND " & Campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                CadAux = CadAux & "23,59,59"
                            Else
                                CadAux = CadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            CadAux = CadAux & ")"
                        End If
                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        CadAux = Campo & " <= " & Val(cadHasta)
                    Case "T"
                        CadAux = Campo & " <= """ & cadHasta & """"
                    Case "F"
                        CadAux = Campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                    Case "FH"
                            CadAux = Campo & " <= DateTime(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ","
                            If Len(cadHasta) = 10 Then
                                CadAux = CadAux & "23,59,59"
                            Else
                                CadAux = CadAux & Hour(cadHasta) & "," & Minute(cadHasta) & "," & Second(cadHasta)
                            End If
                            CadAux = CadAux & ")"
                        
                End Select
            End If
        End If
    End If
    If CadAux <> "" And CadAux <> "Error" Then CadAux = "(" & CadAux & ")"
    CadenaDesdeHasta = CadAux
ErrDH:
    If Err.Number <> 0 Then CadenaDesdeHasta = "Error"
End Function


Public Function CadenaDesdeHastaBD(cadDesde As String, cadHasta As String, Campo As String, TipoCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= valor1 and campo<=valor2) "
'Para MySQL
Dim CadAux As String

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            CadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    CadAux = Campo & " >= " & Val(cadDesde)
                Case "T"
                    CadAux = Campo & " >= """ & cadDesde & """"
                Case "F"
                    CadAux = "(" & Campo & " >= '" & Format(cadDesde, FormatoFecha) & "')"
                Case "FH"
                    If Len(cadDesde) = 10 Then cadDesde = cadDesde & " 00:00:00"
                    CadAux = "(" & Campo & " >= '" & Format(cadDesde, FormatoFechaHora) & "')"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If CadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and " & Campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " and (" & Campo & " <= '" & Format(cadHasta, FormatoFecha) & "')"
                        End If
                    Case "FH"
                        If Len(cadHasta) = 10 Then cadHasta = cadHasta & " 23:59:59"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            CadAux = "Error"
                        Else
                            CadAux = CadAux & " AND (" & Campo & " <= '" & Format(cadHasta, FormatoFechaHora) & "')"
                        End If

                    

                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        CadAux = Campo & " <= " & Val(cadHasta)
                    Case "T"
                        CadAux = Campo & " <= """ & cadHasta & """"
                    Case "F"
                        CadAux = Campo & " <= '" & Format(cadHasta, FormatoFecha) & "'"
                End Select
            End If
        End If
    End If
    If CadAux <> "" And CadAux <> "Error" Then CadAux = "(" & CadAux & ")"
    CadenaDesdeHastaBD = CadAux
End Function

Public Function TotalRegistros(vSql As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Set Rs = New ADODB.Recordset
    Rs.Open vSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalRegistros = Rs.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function

Public Function TotalRegistrosConsulta(cadSQL) As Long
Dim cad As String
Dim Rs As ADODB.Recordset

    On Error GoTo ErrTotReg
    cad = "SELECT count(*) FROM (" & cadSQL & ") x"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not Rs.EOF Then
        TotalRegistrosConsulta = DBLet(Rs.Fields(0).Value, "N")
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ErrTotReg:
    MuestraError Err.Number, "", Err.Description
End Function


Public Sub PonerLongCamposGnral(ByRef formulario As Form, Modo As Byte, Opcion As Byte)
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'ya que en busqueda se permite introducir criterios más largos del tamaño del campo
'en busqueda permitimos escribir: "0001:0004"
'en cambio al insertar o modificar la longitud solo debe permitir ser: "0001"
'(IN) formulario y Modo en que se encuentra el formulario
'(IN) Opcion : 1 para los TEXT1, 3 para los txtAux

    Dim i As Integer
    
    On Error Resume Next

    With formulario
        If Modo = 1 Then 'BUSQUEDA
            Select Case Opcion
                Case 1 'Para los TEXT1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tamaño infinito
                            End If
                        End With
                    Next i
                
                Case 3 'para los TXTAUX
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .MaxLength <> 0 Then
                               .HelpContextID = .MaxLength 'guardamos es maxlenth para reestablecerlo despues
                                .MaxLength = 0 'tamaño infinito
                            End If
                        End With
                    Next i
            End Select
            
        Else 'resto de modos
            Select Case Opcion
                Case 1 'par los Text1
                    For i = 0 To .Text1.Count - 1
                        With .Text1(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
                Case 3 'para los txtAux
                    For i = 0 To .txtAux.Count - 1
                        With .txtAux(i)
                            If .HelpContextID <> 0 Then
                                .MaxLength = .HelpContextID 'volvemos a poner el valor real del maxlenth
                                .HelpContextID = 0
                            End If
                        End With
                    Next i
            End Select
        End If
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub CargaGridGnral(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Sql As String, PrimeraVez As Boolean)
    On Error GoTo ECargaGRid

    vDataGrid.Enabled = True
    '    vdata.Recordset.Cancel
    vData.ConnectionString = Conn
    vData.RecordSource = Sql
    vData.CursorType = adOpenDynamic
    vData.LockType = adLockPessimistic
    vDataGrid.ScrollBars = dbgNone
    vData.Refresh
    
    Set vDataGrid.DataSource = vData
    vDataGrid.AllowRowSizing = False
    vDataGrid.RowHeight = 350 '350
    
    If PrimeraVez Then
        vDataGrid.ClearFields
        vDataGrid.ReBind
        vDataGrid.Refresh
    End If
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "CargaGrid", Err.Description
End Sub

'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Public Function SugerirCodigoSiguienteStr(NomTabla As String, NomCodigo As String, Optional CondLineas As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo ESugerirCodigo

    'SQL = "Select Max(codtipar) from stipar"
    Sql = "Select Max(" & NomCodigo & ") from " & NomTabla
    If CondLineas <> "" Then
        Sql = Sql & " WHERE " & CondLineas
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, , , adCmdText
    Sql = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If IsNumeric(Rs.Fields(0)) Then
                Sql = CStr(Rs.Fields(0) + 1)
            Else
                If Asc(Left(Rs.Fields(0), 1)) <> 122 Then 'Z
                Sql = Left(Rs.Fields(0), 1) & CStr(Asc(Right(Rs.Fields(0), 1)) + 1)
                End If
            End If
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    SugerirCodigoSiguienteStr = Sql
ESugerirCodigo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Public Sub CargarValoresAnteriores(formulario As Form, Optional opcio As Integer, Optional nom_frame As String)
Dim Control As Object
Dim mTag As CTag
Dim Izda As String
Dim cad As String
    Set mTag = New CTag

    For Each Control In formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado Then
                        If mTag.Columna <> "" And Not mTag.EsClave Then
                            If Izda <> "" Then Izda = Izda & " , "
                            'Access
                            'Izda = Izda & "[" & mTag.Columna & "]"
                            Izda = Izda & "" & mTag.Columna & " = "
                            'Parte VALUES
                            cad = ValorParaSQL(Control.Text, mTag)
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If Not mTag.EsClave Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & " = "
                        If Control.Value = 1 Then
                            cad = "1"
                            Else
                            cad = "0"
                        End If
                        If mTag.TipoDato = "N" Then cad = Abs(CBool(cad))
                        Izda = Izda & cad
                    End If
                End If
            End If
            
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
                    mTag.Cargar Control
                    If mTag.Cargado And Not mTag.EsClave Then
                        If Izda <> "" Then Izda = Izda & " , "
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & " = "
                        If Control.ListIndex = -1 Then
                            cad = ValorNulo
                        ElseIf mTag.TipoDato = "N" Then
                            cad = Control.ItemData(Control.ListIndex)
                        Else
                            cad = ValorParaSQL(Control.List(Control.ListIndex), mTag)
                        End If
                        Izda = Izda & cad
                    End If
                End If
            End If
            
        'OPTION BUTTON
        ElseIf TypeOf Control Is OptionButton Then
            If Control.Tag <> "" Then
                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
                    mTag.Cargar Control
                    If mTag.Cargado And Not mTag.EsClave Then
                        If Control.Value Then
                            If Izda <> "" Then Izda = Izda & " , "
                            Izda = Izda & "" & mTag.Columna & " = "
                            cad = Control.Index
                            Izda = Izda & cad
                        End If
                    End If
                End If
            End If
            
'        ElseIf TypeOf Control Is DTPicker Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado And Not mTag.EsClave Then
'                        If Izda <> "" Then Izda = Izda & " , "
'                        Izda = Izda & "" & mTag.Columna & " = "
'
'                        'Parte VALUES
'                        If Control.Visible Then
'                            cad = ValorParaSQL(Control.Value, mTag)
'                        Else
'                            cad = ValorNulo
'                        End If
'                        Izda = Izda & cad
'                    End If
'                End If
'            End If
        End If
        
    Next Control

    ValorAnterior = Izda

End Sub

Public Function SituarDataTrasEliminar(ByRef vData As Adodc, NumReg, Optional no_refre As Boolean) As Boolean
    On Error GoTo ESituarDataElim

    If Not no_refre Then vData.Refresh 'quan siga False o no es passe a la funció, es refrescarà. Hi ha que passar-lo com a True quan el manteniment siga Grid per a que no refresque
    
    If Not vData.Recordset.EOF Then    'Solo habia un registro
        If NumReg > vData.Recordset.RecordCount Then
            vData.Recordset.MoveLast
        Else
            vData.Recordset.MoveFirst
            vData.Recordset.Move NumReg - 1
        End If
        SituarDataTrasEliminar = True
    Else
        SituarDataTrasEliminar = False
    End If
        
ESituarDataElim:
    If Err.Number <> 0 Then
        Err.Clear
        SituarDataTrasEliminar = False
    End If
End Function


Public Sub AnyadirLinea(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
On Error Resume Next

    vDataGrid.AllowAddNew = True
    If vData.Recordset.RecordCount > 0 Then
        vDataGrid.HoldFields
        vData.Recordset.MoveLast
        vDataGrid.Row = vDataGrid.Row + 1
    End If
    vDataGrid.Enabled = False
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub LimpiarLin(ByRef formulario As Form, nomframe As String)
'Limpiar los controles Text que esten dentro del frame nomFrame
    Dim Control As Object

    For Each Control In formulario.Controls
        If TypeOf Control Is TextBox Then
            If Control.Container.Name = nomframe Then
                Control.Text = ""
            End If
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Container.Name = nomframe Then
                Control.ListIndex = -1
            End If
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Container.Name = nomframe Then
                Control.Value = 0
            End If
        End If
    Next Control
End Sub


Public Function PonerFormatoFecha(ByRef T As TextBox) As Boolean
Dim cad As String

    cad = T.Text
    If cad <> "" Then
        If Not EsFechaOK(T) Then
            MsgBox "Fecha incorrecta. (dd/MM/yyyy)", vbExclamation
            cad = "mal"
        End If
        If cad <> "" And cad <> "mal" Then
'            T.Text = Cad
            PonerFormatoFecha = True
        Else
            PonFoco T
        End If
    End If
End Function



Public Function ComprobarCero(Valor As String) As String
    If Valor = "" Then
        ComprobarCero = "0"
    Else
        ComprobarCero = Valor
    End If
End Function

Public Function PonerFormatoDecimal(ByRef T As TextBox, tipoF As Single) As Boolean
'tipoF: tipo de Formato a aplicar
'  1 -> Decimal(12,2)
'  2 -> Decimal(8,3)
'  3 -> Decimal(10,2)
'  4 -> Decimal(5,2)

' 6--> Decimal 10,4

Dim Valor As Double
Dim PEntera As Currency
Dim NoOK As Boolean
Dim i As Byte
Dim cadEnt As String
'Dim mTas As CTag

    If T.Text = "" Then Exit Function
    PonerFormatoDecimal = False
    NoOK = False
    With T
        If Not EsNumerico(CStr(.Text)) Then
            PonFoco T
            Exit Function
        End If


        If InStr(1, .Text, ",") > 0 Then
            Valor = ImporteFormateado(.Text)
        Else
            cadEnt = .Text
            i = InStr(1, cadEnt, ".")
            If i > 0 Then cadEnt = Mid(cadEnt, 1, i - 1)
            If tipoF = 1 And Len(cadEnt) > 10 Then
                MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                NoOK = True
            End If
            If NoOK Then
'                    .Text = ""
                T.SetFocus
                Exit Function
            End If
            Valor = CDbl(TransformaPuntosComas(.Text))
        End If
            
        'Comprobar la longitud de la Parte Entera
        PEntera = Int(Valor)
        Select Case tipoF 'Comprobar longitud
            Case 1 'Decimal(12,2)
                If Len(CStr(PEntera)) > 10 Then
                    MsgBox "El valor no puede ser mayor de 9999999999,99", vbExclamation
                    NoOK = True
                End If
            Case 2 'Decimal(8,3)
                If Len(CStr(PEntera)) > 5 Then
                    MsgBox "El valor no puede ser mayor de 99999,999", vbExclamation
                    NoOK = True
                End If
            Case 3 'Decimal(10,2)
                If Len(CStr(PEntera)) > 8 Then
                    MsgBox "El valor no puede ser mayor de 99999999,99", vbExclamation
                    NoOK = True
                End If
            Case 4 'Decimal(5,2)
                If Len(CStr(PEntera)) > 3 Then
                    MsgBox "El valor no puede ser mayor de 999,99", vbExclamation
                    NoOK = True
                End If
            
            
        End Select

            
            If NoOK Then
                PonerFormatoDecimal = False
                T.SetFocus
                Exit Function
            End If
            
            'Poner el Formato
            Select Case tipoF
                Case 1 'Formato Decimal(12,2)
                    .Text = Format(Valor, FormatoImporte)
                Case 2 'Formato Decimal(8,3)
                    .Text = Format(Valor, FormatoPrecio)
                Case 3 'Formato Decimal(10,2)
                    .Text = Format(Valor, FormatoDec10d2)
                Case 4 'Formato Decimal(5,2)
                    .Text = Format(Valor, FormatoPorcen)
                
                Case 6
                    .Text = Format(Valor, "##,###,##0.0000")
                
            End Select
            PonerFormatoDecimal = True
    End With
End Function

Public Function PonerNombreDeCod(ByRef Txt As TextBox, tabla As String, Campo As String, Optional Codigo As String, Optional Tipo As String, Optional cBD As Byte, Optional codigo2 As String, Optional Valor2 As String, Optional tipo2 As String) As String
'Devuelve el nombre/Descripción asociado al Código correspondiente
'Además pone formato al campo txt del código a partir del Tag
Dim Sql As String
Dim Devuelve As String
Dim vTag As CTag
Dim ValorCodigo As String

    On Error GoTo EPonerNombresDeCod

    ValorCodigo = Txt.Text
    If ValorCodigo <> "" Then
        Set vTag = New CTag
        If vTag.Cargar(Txt) Then
            If Codigo = "" Then Codigo = vTag.Columna
            If Tipo = "" Then Tipo = vTag.TipoDato
            
            Sql = DevuelveDesdeBDNew(cConta, tabla, Campo, Codigo, ValorCodigo, Tipo, , codigo2, Valor2, tipo2)
            If vTag.TipoDato = "N" Then ValorCodigo = Format(ValorCodigo, vTag.Formato)
            Txt.Text = ValorCodigo 'Valor codigo formateado
            If Sql = "" Then
            
            Else
                PonerNombreDeCod = Sql 'Descripcion del codigo
            End If
        End If
        Set vTag = Nothing
    Else
        PonerNombreDeCod = ""
    End If
'    Exit Function
EPonerNombresDeCod:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Nombre asociado a código: " & Codigo, Err.Description
End Function

Public Sub MostrarObservaciones(Cuenta As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Var As String
    On Error Resume Next

    If Cuenta = "" Then Exit Sub

    Sql = "select obsdatos from cuentas where codmacta = " & DBSet(Cuenta, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Var = DBMemo(Rs.Fields(0).Value)
    If Trim(Var) <> "" Then
        If Asc(Var) <> 13 Then
            MsgBox Var, vbInformation
        End If
    End If
    
    Set Rs = Nothing
    
End Sub

Public Function EsMultiseccion(vNombre As String) As Boolean
Dim Sql As String

    Sql = "select esmultiseccion from " & vNombre & ".parametros "
    EsMultiseccion = (DevuelveValor(Sql) = 1)

End Function

Public Function TieneInmovilizado() As Boolean
Dim Sql As String
    
    Sql = "select * from paramamort "
    TieneInmovilizado = (TotalRegistrosConsulta(Sql) <> 0)

End Function


Public Sub CargarProgres(ByRef PBar As ProgressBar, Valor As Long)
On Error Resume Next
    
    PBar.Value = 0
    PBar.Max = 100
    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub IncrementarProgres(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function Ejecuta(ByRef Sql As String, Optional OcultarMsgbox As Boolean) As Boolean

    On Error Resume Next
    Conn.Execute Sql
    If Err.Number <> 0 Then
        If Not OcultarMsgbox Then MuestraError Err.Number, "Cadena: " & Sql & vbCrLf & Err.Description
        Ejecuta = False
    Else
        Ejecuta = True
    End If
End Function


Public Function ComprobarContabilizacionFrasCliProv(Escliente As Boolean, NumConta As Integer, Optional DetenerProceso As Boolean) As Boolean
Dim Sql As String
Dim Nregs As Long
Dim Rs As ADODB.Recordset
Dim vCadena As String
    
    On Error GoTo eComprobarContabilizacionFrasCliProv


    ComprobarContabilizacionFrasCliProv = False

    If Escliente Then
        Sql = "select * from ariconta" & NumConta & ".factcli where (numasien is null or numasien = 0 or fechaent is null or numdiari is null) "
        Sql = Sql & " AND fecfactu>=" & DBSet(vParam.fechaini, "F")
    Else
        Sql = "select * from ariconta" & NumConta & ".factpro where (numasien is null or numasien = 0 or fechaent is null or numdiari is null) "
        Sql = Sql & " AND fecharec>=" & DBSet(vParam.fechaini, "F")
    End If
    If TotalRegistrosConsulta(Sql) > 0 Then
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Escliente Then
            vCadena = "Facturas de Cliente sin Nro.Asiento:" & vbCrLf & vbCrLf
        Else
            vCadena = "Facturas de Proveedor sin Nro.Asiento:" & vbCrLf & vbCrLf
        End If
        
        Nregs = 1
        While Not Rs.EOF
            If Escliente Then
                vCadena = vCadena & "Fra. " & DBLet(Rs!NUmSerie) & " " & Format(DBLet(Rs!numfactu), "0000000") & " " & DBLet(Rs!FecFactu, "F")
            Else
                vCadena = vCadena & "Fra.Reg. " & DBLet(Rs!NUmSerie) & " " & Format(DBLet(Rs!Numregis), "0000000") & " " & DBLet(Rs!fecharec, "F")
            
            End If
            
            If (Nregs Mod 2) = 0 Then
                vCadena = vCadena & vbCrLf
            Else
                vCadena = vCadena & "  "
            End If
            
            Nregs = Nregs + 1
            
            Rs.MoveNext
        Wend
    
        If DetenerProceso Then
            MsgBox vCadena & vbCrLf & vbCrLf & "Revise.", vbExclamation
            Exit Function
        Else
            MsgBox vCadena, vbExclamation
        End If
    End If

    ComprobarContabilizacionFrasCliProv = True
    Exit Function

eComprobarContabilizacionFrasCliProv:
    MuestraError Err.Number, "Comprobar Contabilizacion", Err.Description
End Function





'Public Function CambiosEnFormulario(ByRef formulario As Form, Optional opcio As Integer, Optional nom_frame As String, Optional QueCampos As String, Optional QueIndices As String) As Sql
'Dim Control As Object
'Dim mTag As CTag
'Dim mTag1 As CTag
'Dim Aux As String
'Dim CadWhere As String
'Dim cadUPDATE As String
'
'Dim Result As String
'
'On Error GoTo ECambiosEnFormulario
'
'    CambiosEnFormulario = ""
'
'    Set mTag = New CTag
'
'    Aux = ""
'    CadWhere = ""
'    For Each Control In formulario.Controls
'        'Si es texto monta esta parte de sql
'        If TypeOf Control Is TextBox Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        If mTag.Columna <> "" Then
'
'
'                            'Sea para el where o para el update esto lo necesito
'                            Aux = ValorParaSQL(Control.Text, mTag)
'                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
'                            'dentro del WHERE
'                            If mTag.EsClave Then
'                                'Lo pondremos para el WHERE
'                                 If CadWhere <> "" Then CadWhere = CadWhere & " AND "
'                                 CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
'
'                            Else
'                                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        'CheckBOX
'        ElseIf TypeOf Control Is CheckBox Then
'            'Partimos de la base que un booleano no es nunca clave primaria
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If Control.Value = 1 Then
'                        Aux = "TRUE"
'                    Else
'                        Aux = "FALSE"
'                    End If
'                    If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
'                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                    'Esta es para access
'                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
'                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
'                End If
'            End If
'
'        ElseIf TypeOf Control Is ComboBox Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Or ((opcio = 2) And (Control.Container.Name = nom_frame)) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        If Control.ListIndex = -1 Then
'                            Aux = ValorNulo
'                        ElseIf mTag.TipoDato = "N" Then
'                            Aux = Control.ItemData(Control.ListIndex)
'                        Else
'                            Aux = ValorParaSQL(Control.List(Control.ListIndex), mTag)
'                        End If
'
'                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
'                        'dentro del WHERE
'                        If mTag.EsClave Then
'                            'Lo pondremos para el WHERE
'                             If CadWhere <> "" Then CadWhere = CadWhere & " AND "
'                             CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
'                        Else
'                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
'                        End If
''
''
''                        If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
''                        'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
''                        cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
'                    End If
'                End If
'            End If
'
'        ElseIf TypeOf Control Is OptionButton Then
'            If Control.Tag <> "" Then
'                If (opcio = 0) Or ((opcio = 1) And (InStr(1, Control.Container.Name, "FrameAux")) = 0) Then
'                    mTag.Cargar Control
'                    If mTag.Cargado Then
'                        If Control.Value Then
'                            Aux = Control.Index
'                            'Si es campo clave NO se puede modificar y se utiliza como busqueda
'                            'dentro del WHERE
'                              If mTag.EsClave Then
'                                  'Lo pondremos para el WHERE
'                                   If CadWhere <> "" Then CadWhere = CadWhere & " AND "
'                                   CadWhere = CadWhere & "(" & mTag.Columna & " = " & Aux & ")"
'                              Else
'                                  If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
'                                  cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
'                              End If
'                        End If
'                    End If
'                End If
'            End If
'
'        End If
'    Next Control
'    'Construimos el SQL
'    'Ejemplo:
'    'Update Pedidos
'    'SET ImportePedido = ImportePedido * 1.1,
'    'Cargo = Cargo * 1.03
'    'WHERE PaísDestinatario = 'México';
'    If CadWhere = "" Then
'        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
'        Exit Function
'    End If
'    Aux = "UPDATE " & mTag.tabla
'    Aux = Aux & " SET " & cadUPDATE & " WHERE " & CadWhere
'    Conn.Execute Aux, , adCmdText
'
'    ' ### [Monica] 18/12/2006
'    CadenaCambio = cadUPDATE
'
'    CambiosEnFormulario = True
'    Exit Function
'
'ECambiosEnFormulario:
'    MuestraError Err.Number, "Cambios en formulario" & Err.Description
'End Function
'
'

Public Sub CargarCombo_Tabla(ByRef Cbo As ComboBox, NomTabla As String, NomCodigo As String, nomDescrip As String, Optional strWhere As String, Optional ItemNulo As Boolean, Optional Ordenacion As String)
'Carga un objeto ComboBox con los registros de una Tabla
'(IN) cbo: ComboBox en el q se van a cargar los datos
'(IN) nomTabla: nombre de la tabla de la q leeremos los datos a cargar
'(IN) nomCodigo: nombre del campo codigo de la tabla q queremos cargar
'(IN) nomDescrip: nombre del campo descripcion de la tabla a cargar
'(IN) strWhere: para filtrar los registros de la tabla q queremos cargar
'(IN) ItemNulo: si es true se añade el primer item con linea en blanco
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCombo
    
    Cbo.Clear
    
    Sql = "SELECT " & NomCodigo & "," & nomDescrip & " FROM " & NomTabla
    If strWhere <> "" Then Sql = Sql & " WHERE " & strWhere
    Sql = Sql & " ORDER BY "
    If Ordenacion <> "" Then
        Sql = Sql & Ordenacion
    Else
        Sql = Sql & nomDescrip
    End If
    
'    If AbrirRecordset(SQL, RS) Then
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    '- si valor del parametro ItemNulo=true hay que añadir linea en blanco
    If Not Rs.EOF And ItemNulo Then
        Cbo.AddItem "  "
        Cbo.ItemData(Cbo.NewIndex) = 0
    End If
    
    If Not Rs.EOF Then
        If IsNumeric(Rs.Fields(0).Value) Then
            '- si el codigo NomCodigo es numerico en el ItemData se carga el campo clave primaria
            '- y en List la descripcion NomDescrip
            While Not Rs.EOF
              Cbo.AddItem Rs.Fields(1).Value 'descrip
              Cbo.ItemData(Cbo.NewIndex) = Rs.Fields(0).Value 'codigo
              Rs.MoveNext
            Wend
        Else
            '- si el codigo NomCodigo en alfanumerico no se puede cargar
            '- el codigo en ItemData y cargamos un indice ficticio
            '- y en el List el campo codigo NomCodigo
            i = 1
            While Not Rs.EOF
              Cbo.AddItem Rs.Fields(0).Value 'campo del codigo
              Cbo.ItemData(Cbo.NewIndex) = i
              i = i + 1
              Rs.MoveNext
            Wend
        End If
    End If
'    End If
    
'    CerrarRecordset RS
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
ErrCombo:
    MuestraError Err.Number, "Cargar combo." & NomTabla, Err.Description
End Sub

'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'
' Cosas comunes sencilass con los listview
'
'**************************************************************************************************************
'**************************************************************************************************************

'El sql llevará los dos campos. Codigo, descripcion.
' En el orden que TOCA
Public Sub CargaListviewCodigoDescripcion(ByRef LW As ListView, Sql As String, Checked As Boolean, PorcentajeAnchoCampo1 As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim Ancho As Integer
    On Error GoTo ECargarList

    'Los encabezados
    LW.ColumnHeaders.Clear
    
    
    Ancho = (LW.Width - 320) * (PorcentajeAnchoCampo1 / 100)
    LW.ColumnHeaders.Add , , "Código", Ancho
    Ancho = (LW.Width - 320) - Ancho
    LW.ColumnHeaders.Add , , "Descripción", Ancho
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = LW.ListItems.Add
        
        ItmX.Checked = Checked
        ItmX.Text = Rs.Fields(0).Value
        ItmX.SubItems(1) = Rs.Fields(1).Value
        Rs.MoveNext
    Wend
    Rs.Close
    

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar listview", Err.Description
    End If
    Set Rs = Nothing


End Sub

Public Sub ListviewSelecDeselec(ByRef LW As ListView, check As Boolean)
Dim N As Integer
    For N = 1 To LW.ListItems.Count
        LW.ListItems(N).Checked = check
    Next
End Sub
