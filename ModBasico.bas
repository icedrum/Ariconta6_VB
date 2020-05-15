Attribute VB_Name = "ModBasico"
Option Explicit

'
'
'   ******  Cuidado  con los nombres de los objetos. Con las mayusculas minusculas
'
Public Sub arregla(ByRef tots As String, ByRef grid As DataGrid, ByRef formu As Form)
    'Dim tots As String
    Dim camp As String
    Dim Mens As String
    Dim difer As Integer
    Dim I As Integer
    Dim K As Integer
    Dim posi As Integer
    Dim posi2 As Integer
    Dim fil As Integer
    Dim C As Integer
    Dim o As Integer
    Dim A() As Variant 'per als 5 parametres
    'Dim grid As DataGrid
    Dim Obj As Object
    Dim obj_ant As Object
    Dim primer As Boolean
    Dim TotalAncho As Integer
    
    grid.AllowRowSizing = False
    grid.RowHeight = 350
    
    '***********
    difer = 563 'dirència recomanda entre l'ample del Datagrid i la suma dels amples de les columnes
    '***********
    
    TotalAncho = 0
    primer = False
'    Set grid = DataGrid1 'nom del DataGrid
    fil = -1 'fila a -1
    C = -1 'columna del datagrid a 0
    'tots = "S|txtAux(0)|T|Código|700|;S|txtAux(1)|T|Descripción|3000|;"
    
    While (tots <> "") 'bucle per a recorrer els distins camps
        Set Obj = Nothing
        Set obj_ant = Nothing
    
        fil = fil + 1
        'ReDim Preserve A(6, fil)
        ReDim Preserve A(5, fil)
        'fila i columna a 0 (NOTA: les files es numeren a partir d'1 i les columnes a partir de 0)
        posi = InStr(tots, ";") '1ª posicio del ;
        camp = Left(tots, posi - 1)
        tots = Right(tots, Len(tots) - posi) 'lleve el camp actual
        'For k = 0 To 5
        For K = 0 To 4
          posi2 = InStr(camp, "|") '1ª posició del |
          A(K, fil) = Left(camp, posi2 - 1)
          camp = Right(camp, Len(camp) - posi2) 'lleve l'argument actual
        Next K 'quan acabe el for tinc en A el camp actual
        
        'només incremente el nº de la columna si no es un boto
        If A(2, fil) <> "B" Then C = C + 1
        
        If A(0, fil) = "N" Then 'no visible
            grid.Columns(C).visible = False
            grid.Columns(C).Width = 0 'si no es visible, pose a 0 l'ample
        ElseIf A(0, fil) = "S" Then 'visible
            ' ********* CAPTION I WIDTH DE L'OBJECTE ************
            
            Select Case A(2, fil) 'tipo (T, C o B) (o CB=CheckBox ) (DT=DTPicker)
                Case "T"
                    grid.Columns(C).visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil))
'                    If A(5, fil) <> "" Then
'                        grid.Columns(c).NumberFormat = A(5, fil)
'                    Else
'                        grid.Columns(c).NumberFormat = ""
'                    End If
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                Case "C"
                    grid.Columns(C).visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil)) - 10
'                    If A(5, fil) <> "" Then
'                        grid.Columns(c).NumberFormat = A(5, fil)
'                    Else
'                        grid.Columns(c).NumberFormat = ""
'                    End If
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                Case "B"
                
               '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                Case "CB"
                    grid.Columns(C).visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil))
                    TotalAncho = TotalAncho + CInt(A(4, fil))
               '===============================================
               '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                Case "DT"
                    grid.Columns(C).visible = True
                    If A(3, fil) <> "" Then grid.Columns(C).Caption = A(3, fil)
                    If A(4, fil) <> "" Then grid.Columns(C).Width = CInt(A(4, fil))
                    TotalAncho = TotalAncho + CInt(A(4, fil))
                '==============================================
            End Select
                       
            ' ********* CARREGUE L'OBJECTE ************
            Set Obj = eval(formu, CStr(A(1, fil)))
            
            ' ********* NUMBERFORMAT i ALIGNMENT DE L'OBJECTE ************
            If (A(2, fil) = "T") Or (A(2, fil) = "C") Or (A(2, fil) = "DT") Then 'el numberformat només es per a text o combo
                If Obj.Tag <> "" Then
                    grid.Columns(C).NumberFormat = FormatoCampo2(Obj)
                    If TipoCamp(Obj) = "N" Then
                        If (A(2, fil) = "T") Then _
                            grid.Columns(C).Alignment = dbgRight ' el Alignment només per a Text
                        grid.Columns(C).NumberFormat = grid.Columns(C).NumberFormat & " "
                    End If
                Else
                    grid.Columns(C).NumberFormat = ""
                End If
            End If
            
            ' ********* WIDTH I LEFT DE L'OBJECTE ************
            Select Case A(2, fil) 'tipo (T, C o B)
                Case "T"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 60
                        'obj.Width = grid.Columns(c).Width - 8
                        Obj.Left = grid.Left + 340
                        'obj.Left = grid.Left + 308
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a text es text
                                Obj.Width = grid.Columns(C).Width - 60
                                'obj.Width = grid.Columns(c).Width - 38
                                Obj.Left = obj_ant.Left + obj_ant.Width + 60
                                'obj.Left = obj_ant.Left + obj_ant.Width + 38
                            Case "C" 'objecte anterior a text es combo
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 30
                            Case "B" 'objecte anterior a text es un boto
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 30
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es CheckBox
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 60
                            '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                            Case "DT" 'anterior es un DTPicker
                                Obj.Width = grid.Columns(C).Width - 60
                                Obj.Left = obj_ant.Left + obj_ant.Width + 60
                        End Select
                    End If
                    
                Case "C"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 10
                        Obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                Obj.Width = grid.Columns(C).Width - 20
                                Obj.Left = obj_ant.Left + obj_ant.Width + 40
                            Case "C" 'objecte anterior a combo es combo
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
'                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
'                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                                '=== LAURA (14/09/06): añadir este caso
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width + 10
                            
                            '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es CheckBox (falta comprobar)
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                        End Select
                    End If
                    
                Case "B"
                    If Not primer Then 'es el primer objecte visible
                        ' *** FALTA PER A QUAN UN BOTO ES EL PRIMER OBJECTE VISIBLE
                        Mens = "Falta programar en arreglaGrid per al cas que un Button es el primer objete visible d'un Datagrid"
                        MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a boto es text
                                obj_ant.Width = obj_ant.Width - Obj.Width + 30 '1r faig més curt l'objecte de text
                                Obj.Left = obj_ant.Left + obj_ant.Width
                                'obj.Left = obj_ant.Left + obj_ant.Width - obj.Width
                            Case "C" 'objecte anterior a boto es combo
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN BOTO ES UN COMBO
                                Mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un Button es un ComboBox"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN BOTO ES UN BOTO
                                Mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un Button es un Button"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                        End Select
                    End If
                    
                 '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                Case "CB"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 10
                        Obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                Obj.Width = grid.Columns(C).Width - (grid.Columns(C).Width / 3)
                                Obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(C).Width / 3) - 10
                            Case "C" 'objecte anterior a combo es combo
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
'Laura: 140508
'                                mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
'                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
                                
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width + 10
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es un ChekBox
                                Obj.Width = grid.Columns(C).Width - (grid.Columns(C).Width / 3)
                                Obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(C).Width / 3)
                        End Select
                    End If
                
                
                 '=== LAURA (14/07/06): añadir tipo DT=DTPicker
                Case "DT"
                    If Not primer Then 'es el primer objecte visible
                        Obj.Width = grid.Columns(C).Width - 10
                        Obj.Left = grid.Left + 320
                    Else
                        o = 0
                        While obj_ant Is Nothing
                            o = o + 1
                            If A(0, fil - o) = "S" Then
                                Set obj_ant = eval(formu, CStr(A(1, fil - o)))
                            End If
                        Wend
                        'en obj_ant tindré el 1r objecte per darrere que siga visible
                        Select Case A(2, fil - o)
                            Case "T" 'objecte anterior a combo es text
                                Obj.Width = grid.Columns(C).Width - 40
                                Obj.Left = obj_ant.Left + obj_ant.Width + 40
                            Case "C" 'objecte anterior a combo es combo
                                Obj.Width = grid.Columns(C).Width
                                Obj.Left = obj_ant.Left + obj_ant.Width
                            Case "B" 'objecte anterior a combo es un boto
                                ' *** FALTA PER A QUAN L'OBJECTE ANTERIOR A UN COMBO ES UN BOTO
                                Mens = "Falta programar en arreglaGrid per al cas que un l'objecte anterior a un ComboBox es un Button"
                                MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & Mens
                                
                             '=== LAURA (07/04/06): añadir tipo CB=CheckBox
                            Case "CB" 'anterior es un ChekBox
                                Obj.Width = grid.Columns(C).Width - (grid.Columns(C).Width / 3)
                                Obj.Left = obj_ant.Left + obj_ant.Width + (grid.Columns(C).Width / 3)
                            Case "DT" 'anterior es un DTPicker
                                Obj.Width = grid.Columns(C).Width - 40
                                Obj.Left = obj_ant.Left + obj_ant.Width + 40
                        End Select
                    End If
                Case Else
                    MsgBox "No existix el tipo de control " & A(2, fil)
            End Select
            
        primer = True
        End If
                
    Wend

    'No permitir canviar tamany de columnes
    For I = 0 To grid.Columns.Count - 1
         grid.Columns(I).AllowSizing = False
    Next I

'    If grid.Width - TotalAncho <> difer Then
'        mens = "Es recomana que el total d'amples de les columnes per a este DataGrid siga de "
'        mens = mens & CStr(grid.Width - difer)
'        MsgBox "MÒDUL arreglaGrid:" & vbCrLf & "-----------------------" & vbCrLf & vbCrLf & mens
'    End If
End Sub

Public Function eval(ByRef formu As Form, nom_camp As String) As Control
Dim Ctrl As Control
Dim nom_camp2 As String
Dim nou_i As Integer
Dim J As Integer

    Set eval = Nothing
    J = InStr(1, nom_camp, "(")
    If J = 0 Then
        nou_i = -1
    Else
        nom_camp = Left(nom_camp, Len(nom_camp) - 1)
        nou_i = Val(Mid(nom_camp, J + 1))
        nom_camp = Left(nom_camp, J - 1)
    End If
    
    For Each Ctrl In formu.Controls
        If Ctrl.Name = nom_camp Then
            If nou_i >= 0 Then
                If nou_i = Ctrl.Index Then
                    J = 1 'coincidix el nom i l'index
                Else
                    J = 0 'coincidix el nom però no l'index
                End If
            Else
                J = 1 'coincidix el nom i no te index
            End If
        Else
            J = -1 'no coincidix el nom
        End If
        
        If J > 0 Then
            Set eval = Ctrl
            Exit For
        End If
    Next Ctrl
End Function


Public Function PerderFocoGnral(ByRef Text As TextBox, Modo As Byte) As Boolean
Dim Comprobar As Boolean
'Dim mTag As CTag

    On Error Resume Next

    If Screen.ActiveForm.ActiveControl.Name = "cmdCancelar" Then
        PerderFocoGnral = False
        Exit Function
    End If

    With Text
        'Quitamos blancos por los lados
        .Text = Trim(.Text)
        
        
         If .BackColor = vbLightBlue Then
            If .Locked Then
                .BackColor = vbLightBlue '&H80000018
            Else
                .BackColor = vbWhite
            End If
        End If
        
        
        'Si no estamos en modo: 3=Insertar o 4=Modificar o 1=Busqueda, no hacer ninguna comprobacion
        If (Modo <> 3 And Modo <> 4 And Modo <> 1 And Modo <> 5) Then
            PerderFocoGnral = False
            Exit Function
        End If
        
        If Modo = 1 Then
            'Si estamos en modo busqueda y contiene un caracter especial no realizar
            'las comprobaciones
            Comprobar = ContieneCaracterBusqueda(.Text)
            If Comprobar Then
                PerderFocoGnral = False
                Exit Function
            End If
        End If
        PerderFocoGnral = True
    End With
    
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function PonerFormatoEntero(ByRef T As TextBox) As Boolean
'Comprueba que el valor del textbox es un entero y le pone el formato
Dim mTag As CTag
Dim cad As String
Dim Formato As String
On Error GoTo EPonerFormato

    
    If T.Text = "" Then Exit Function
    PonerFormatoEntero = True
    
    Set mTag = New CTag
    mTag.Cargar T
    If mTag.Cargado Then
       cad = mTag.Nombre 'descripcion del campo
       Formato = mTag.Formato
    End If
    Set mTag = Nothing

    If Not EsEntero(T.Text) Then
        PonerFormatoEntero = False
        MsgBox "El campo " & cad & " tiene que ser numérico.", vbExclamation
        PonFoco T
    Else
         'T.Text = Format(T.Text, Formato)
         ' **** 21-11-2005 Canvi de Cèsar. Per a que formatetge be si es posa un
         ' número negatiu, li lleve un 0 a la màscara per a que el número
         ' càpiga dins del textbox en el maxlength asignat.
         ' Si es crida a esta funció la màscara es del tipo 0000
         If T.Text < 0 Then _
            Formato = Replace(Formato, "0", "", 1, 1)
        ' *************************************************************************
         
         T.Text = Format(T.Text, Formato)
    End If
    
EPonerFormato:
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function ExisteCP(T As TextBox) As Boolean
'comprueba para un campo de texto que sea clave primaria, si ya existe un
'registro con ese valor
Dim vTag As CTag
Dim Devuelve As String

    On Error GoTo ErrExiste

    ExisteCP = False
    If T.Text <> "" Then
        If T.Tag <> "" Then
            Set vTag = New CTag
            If vTag.Cargar(T) Then
'                If vtag.EsClave Then
                    Devuelve = DevuelveDesdeBD(vTag.Columna, vTag.tabla, vTag.Columna, T.Text, vTag.TipoDato)
                    If Devuelve <> "" Then
    '                    MsgBox "Ya existe un registro para " & vtag.Nombre & ": " & T.Text, vbExclamation
                        MsgBox "Ya existe el " & vTag.Nombre & ": " & T.Text, vbExclamation
                        ExisteCP = True
                        PonFoco T
                    End If
'                End If
            End If
            Set vTag = Nothing
        End If
    End If
    Exit Function
    
ErrExiste:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar código.", Err.Description
End Function

Public Function PonerContRegistros(ByRef vData As Adodc) As String
'indicador del registro donde nos encontramos: "1 de 20"
    On Error GoTo EPonerReg
    
    If Not vData.Recordset.EOF Then
        PonerContRegistros = vData.Recordset.AbsolutePosition & " de " & vData.Recordset.RecordCount
    Else
        PonerContRegistros = ""
    End If
    
EPonerReg:
    If Err.Number <> 0 Then
        Err.Clear
        PonerContRegistros = ""
    End If
End Function


Public Function FormatoCampo2(ByRef objec As Object) As String
'Devuelve el formato del campo en el TAg: "0000"
Dim mTag As CTag
Dim cad As String

    On Error GoTo EFormatoCampo2

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        FormatoCampo2 = mTag.Formato
    End If
    
EFormatoCampo2:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function

Public Function TipoCamp(ByRef objec As Object) As String
Dim mTag As CTag
Dim cad As String

    On Error GoTo ETipoCamp

    Set mTag = New CTag
    mTag.Cargar objec
    If mTag.Cargado Then
        TipoCamp = mTag.TipoDato
    End If

ETipoCamp:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Function


Public Function ContieneCaracterBusqueda(CADENA As String) As Boolean
'Comprueba si la cadena contiene algun caracter especial de busqueda
' >,>,>=,: , ....
'si encuentra algun caracter de busqueda devuelve TRUE y sale
Dim B As Boolean
Dim I As Integer
Dim Ch As String

    'For i = 1 To Len(cadena)
    I = 1
    B = False
    Do
        Ch = Mid(CADENA, I, 1)
        Select Case Ch
            Case "<", ">", ":", "="
                B = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                B = True
            Case Else
                B = False
        End Select
    'Next i
        I = I + 1
    Loop Until (B = True) Or (I > Len(CADENA))
    ContieneCaracterBusqueda = B
End Function


Public Sub AyudaAgentes(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Nombre|5230|;"
    frmBas.CadenaConsulta = "SELECT agentes.codigo, agentes.nombre "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM agentes "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||agentes|codigo||S|"
    frmBas.Tag2 = "Nombre|T|N|||agentes|nombre|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "agentes"
    frmBas.CampoCP = "codigo"
    frmBas.Caption = "Agentes"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub




Public Sub AyudaTiposIva(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;S|txtAux(2)|T|IVA|900|;"
    frmBas.CadenaConsulta = "SELECT tiposiva.codigiva, tiposiva.nombriva, tiposiva.porceiva "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM tiposiva "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|N|N|||tiposiva|codigiva|##0|S|"
    frmBas.Tag2 = "Nombre|T|N|||tiposiva|nombriva|||"
    frmBas.Tag3 = "%IVA|T|N|||tiposiva|porceiva|##0.00||"
    
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 15
    frmBas.Maxlen3 = 6
    
    frmBas.tabla = "tiposiva"
    frmBas.CampoCP = "codigiva"
    frmBas.Caption = "Tipos de Iva"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub



Public Sub AyudaTPago(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT tipofpago.tipoformapago, tipofpago.descformapago "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM tipofpago "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|N|N|||tipofpago|tipoformapago|##0|S|"
    frmBas.Tag2 = "Nombre|T|N|||tipofpago|descformapago|||"
    
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 15
    
    frmBas.tabla = "tipofpago"
    frmBas.CampoCP = "tipoformapago"
    frmBas.Caption = "Tipos Forma de Pago"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub







Public Sub AyudaCartas(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT cartas.codcarta, cartas.descarta "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM cartas "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|N|N|||cartas|codcarta|##0|S|"
    frmBas.Tag2 = "Nombre|T|N|||cartas|descarta|||"
    
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 15
    
    frmBas.tabla = "cartas"
    frmBas.CampoCP = "codcarta"
    frmBas.Caption = "Tipos de Cartas"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub



Public Sub AyudaBanco(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Descripción|2595|;S|txtAux(2)|T|IBAN|3000|;"
    frmBas.CadenaConsulta = "SELECT bancos.codmacta, bancos.descripcion, bancos.iban "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM bancos "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Cta. contable|T|N|||bancos|codmacta||S|"
    frmBas.Tag2 = "Descripcion|T|S|||bancos|descripcion|||"
    frmBas.Tag3 = "IBAN|T|N|||bancos|iban|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 40
    
    frmBas.tabla = "bancos"
    frmBas.CampoCP = "codmacta"
    frmBas.Caption = "Bancos Propios"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaRemesa(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

 
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Año|1000|;S|txtAux(2)|T|Fecha     Banco|4100|;"
    frmBas.CadenaConsulta = "select codigo,anyo,concat( DATE_FORMAT(fecremesa,'%Y-%m-%d'),'     '  ,nommacta)"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from remesas left join cuentas on remesas.codmacta=cuentas.codmacta "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE situacion='Q'  and tipo=0 and fecremesa>=DATE_ADD(now(), INTERVAL -1 YEAR)"
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Codigo|N|N|||remesas|codigo|000|S|"
    frmBas.Tag2 = "Año|N|S|||remesas|anyo|0000|S|"
    frmBas.Tag3 = "Fecha      Banco|T|N|||cuentas|nommacta|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 80
    
    frmBas.tabla = "remesas"
    frmBas.CampoCP = "codigo"
    frmBas.Caption = "Remesas"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub





Public Sub AyudaPais(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Descripción|4695|;S|txtAux(2)|T|Intracom|900|;"
    frmBas.CadenaConsulta = "SELECT paises.codpais, paises.nompais, if(paises.intracom=0,'No','Si') intracom "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM paises "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Codigo|T|N|||paises|codpais||S|"
    frmBas.Tag2 = "Descripcion|T|S|||paises|nompais|||"
    frmBas.Tag3 = "Intracom.|T|N|||paises|intracom|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 40
    frmBas.Maxlen3 = 4
    
    frmBas.tabla = "paises"
    frmBas.CampoCP = "codpais"
    frmBas.Caption = "Países"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub



Public Sub AyudaAsientosP(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT asipre.numaspre, asipre.nomaspre "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM asipre "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Nº Asiento|N|N|||asipre|numaspre|0000|S|"
    frmBas.Tag2 = "Nombre Asiento|T|N|||asipre|nomaspre|||"
    
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 40
    
    frmBas.tabla = "asipre"
    frmBas.CampoCP = "numaspre"
    frmBas.Caption = "Asientos Predefinidos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaCC(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT ccoste.codccost, ccoste.nomccost "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ccoste "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||ccoste|codccost||S|"
    frmBas.Tag2 = "Descripción|T|N|||ccoste|nomccost|||"
    
    frmBas.Maxlen1 = 4
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "ccoste"
    frmBas.CampoCP = "codccost"
    frmBas.Caption = "Centros de Coste"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaAsientos(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|2405|;S|txtAux(1)|T|Fecha|2695|;S|txtAux(2)|T|Diario|1900|;"
    frmBas.CadenaConsulta = "SELECT hcabapu.numasien, hcabapu.fechaent, hcabapu.numdiari "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM hcabapu "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Nº asiento|N|S|0||hcabapu|numasien|######0|S|"
    frmBas.Tag2 = "Fecha entrada|F|N|||hcabapu|fechaent|dd/mm/yyyy|S|"
    frmBas.Tag3 = "Diario|N|N|0||hcabapu|numdiari|#0|S|"

    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen2 = 7
    
    frmBas.tabla = "hcabapu"
    frmBas.CampoCP = "numasien"
    frmBas.Caption = "Asientos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaFPago(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|805|;S|txtAux(1)|T|Descripción|4295|;S|txtAux(2)|T|Tipo Pago|1900|;"
    frmBas.CadenaConsulta = "SELECT formapago.codforpa, formapago.nomforpa, tipofpago.descformapago "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM formapago ,tipofpago "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE formapago.tipforpa = tipofpago.tipoformapago "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|N|N|0||formapago|codforpa|000|S|"
    frmBas.Tag2 = "Denominación|T|N|||formapago|nomforpa|||"
    frmBas.Tag3 = "Tipo de pago|T|N|||tipofpago|descformapago|||"
    
    frmBas.Maxlen1 = 3
    frmBas.Maxlen2 = 30
    frmBas.Maxlen3 = 20
    
    frmBas.tabla = "formapago"
    frmBas.CampoCP = "codforpa"
    frmBas.Caption = "Formas de Pago"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.DataGrid1.Columns(2).Alignment = dbgLeft
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaEleInmo(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "N||||0|;S|txtAux(0)|T|Elemento|3805|;S|txtAux(1)|T|Fecha|1295|;S|txtAux(2)|T|Valor Adq.|1900|;"
    frmBas.CadenaConsulta = "SELECT inmovele.codinmov, inmovele.nominmov, inmovele.fechaadq, inmovele.valoradq "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM inmovele "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Descripcion|T|N|||inmovele|nominmov|||"
    frmBas.Tag2 = "Fecha adquisición|F|N|||inmovele|fechaadq|dd/mm/yyyy||"
    frmBas.Tag3 = "Valor adquisición|N|N|0||inmovele|valoradq|#,###,##0.00||"

    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 10
    frmBas.Maxlen2 = 10
    
    frmBas.tabla = "inmovele"
    frmBas.CampoCP = "nominmov"
    frmBas.Caption = "Elementos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaCuentasBancarias(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Cuenta|1350|;S|txtAux(1)|T|Descripción|2550|;S|txtAux(2)|T|IBAN|3100|;"
    frmBas.CadenaConsulta = "SELECT bancos.codmacta, bancos.descripcion, bancos.iban "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM bancos "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Cuenta|T|N|||bancos|codmacta||S|"
    frmBas.Tag2 = "Descripción|T|N|||bancos|descripcion|||"
    frmBas.Tag3 = "IBAN|T|S|||bancos|iban|||"

    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 40
    frmBas.Maxlen2 = 40
    
    frmBas.tabla = "bancos"
    frmBas.CampoCP = "codmacta"
    frmBas.Caption = "Cuentas Bancarias"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaFacturasCli(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Serie|2405|;S|txtAux(1)|T|Factura|2195|;S|txtAux(2)|T|Fecha|2400|;"
    frmBas.CadenaConsulta = "SELECT factcli.numserie, factcli.numfactu, factcli.fecfactu "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM factcli "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Nº Serie|T|N|||factcli|numserie||S|"
    frmBas.Tag2 = "Factura|N|N|||factcli|numfactu|0000000|S|"
    frmBas.Tag3 = "Fecha Factura|F|N|||factcli|fecfactu|dd/mm/yyyy|S|"

    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 10
    frmBas.Maxlen2 = 7
    
    frmBas.tabla = "factcli"
    frmBas.CampoCP = "numserie"
    frmBas.Caption = "Facturas de Cliente"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaContadores(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT contadores.tiporegi, contadores.nomregis "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM contadores "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Contador|T|N|||contadores|tiporegi||S|"
    frmBas.Tag2 = "Nombre Contador|T|N|||contadores|nomregis|||"
    
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 40
    
    frmBas.tabla = "contadores"
    frmBas.CampoCP = "tiporegi"
    frmBas.Caption = "Contadores"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub



Public Sub AyudaDepartamentos(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT departamentos.dpto, departamentos.descripcion  "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM departamentos "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||departamentos|dpto|0000|S|"
    frmBas.Tag2 = "Descripcion|T|N|||departamentos|descripcion|||"
    
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 40
    
    frmBas.tabla = "departamento"
    frmBas.CampoCP = "dpto"
    frmBas.Caption = "Departamentos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaCuentas(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String, Optional Empresa As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Cuenta|1350|;S|txtAux(1)|T|Descripción|4150|;S|txtAux(2)|T|NIF|1500|;"
    frmBas.CadenaConsulta = "SELECT cuentas.codmacta, cuentas.nommacta, cuentas.nifdatos "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM "
    If Empresa <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " ariconta" & Empresa & "."
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & "cuentas WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    
    frmBas.Tag1 = "Cuenta|T|N|||cuentas|codmacta||S|"
    frmBas.Tag2 = "Descripción|T|N|||cuentas|nommacta|||"
    frmBas.Tag3 = "NIF|T|S|||cuentas|nifdatos|||"

    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 50
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "cuentas"
    frmBas.CampoCP = "codmacta"
    frmBas.Caption = "Cuentas Contables"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|2|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

'PyG_Situacion:
'           0:  seran los menores de 50
'           1: Ratios
'           2: Personalizables

Public Sub AyudaBalances(frmBas As frmBasico, PyG_Situacion As Byte, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT balances.numbalan, balances.nombalan  "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM balances "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If PyG_Situacion = 0 Then
        If cWhere <> "" Then cWhere = cWhere & " AND "
        cWhere = cWhere & " numbalan < 50"
    ElseIf PyG_Situacion = 1 Then
        If cWhere <> "" Then cWhere = cWhere & " AND "
        cWhere = cWhere & " numbalan between 50 AND 59 "
    
    ElseIf PyG_Situacion = 2 Then
        If cWhere <> "" Then cWhere = cWhere & " AND "
        cWhere = cWhere & " numbalan between 60 AND 99 "
    End If
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|N|N|||balances|numbalan|000|S|"
    frmBas.Tag2 = "Descripcion|T|N|||balances|nombalan|||"
    
    frmBas.Maxlen1 = 7
    frmBas.Maxlen2 = 40
    
    frmBas.tabla = "balances"
    frmBas.CampoCP = "numbalan"
    frmBas.Caption = "Balances"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaDevolucion(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT codigo, descripcion "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM usuarios.wdevolucion "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||usuarios.wdevolucion|codigo||S|"
    frmBas.Tag2 = "Nombre|T|N|||usuarios.wdevolucion|descripcion|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 100
    
    frmBas.tabla = "usuarios.wdevolucion"
    frmBas.CampoCP = "codigo"
    frmBas.Caption = "Conceptos Devolución"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


Public Sub AyudaGastosFijos(frmBas As frmBasico, Optional CodActual As String, Optional cWhere As String)

    frmBas.CadenaTots = "S|txtAux(0)|T|Código|870|;S|txtAux(1)|T|Descripción|5230|;"
    frmBas.CadenaConsulta = "SELECT codigo, descripcion "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM gastosfijos "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Código|T|N|||gastosfijos|codigo||S|"
    frmBas.Tag2 = "Nombre|T|N|||gastosfijos|descripcion|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 30
    
    frmBas.tabla = "gastosfijos"
    frmBas.CampoCP = "codigo"
    frmBas.Caption = "Gastos Fijos"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub




Public Sub AyudaTrasnferencia(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)

 
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1405|;S|txtAux(1)|T|Año|1000|;S|txtAux(2)|T|Fecha            Banco|4100|;"
    frmBas.CadenaConsulta = "select codigo,anyo,concat( DATE_FORMAT(fecha,'%Y-%m-%d'),'     '  ,nommacta)"
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " from transferencias left join cuentas on transferencias.codmacta=cuentas.codmacta "
    frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE situacion='Q'  and  transferencias.fecha>=DATE_ADD(now(), INTERVAL -2 YEAR)"
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Codigo|N|N|||transferencias|codigo|000|S|"
    frmBas.Tag2 = "Año|N|S|||transferencias|anyo|0000|S|"
    frmBas.Tag3 = "Fecha      Banco|T|N|||cuentas|nommacta|||"
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 10
    frmBas.Maxlen3 = 80
    
    frmBas.tabla = "transferencias"
    frmBas.CampoCP = "codigo"
    frmBas.Caption = "Transferencias"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub



Public Sub AyudaImporNavarresSeccion(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
 
    
 
    frmBas.CadenaTots = "S|txtAux(0)|T|concepto|1405|;S|txtAux(1)|T|Descripcion|3000|;"
    frmBas.CadenaConsulta = "select concepto,Descripcion from importnavconceptos"
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Codigo|N|N|||importnavconceptos|concepto|000|S|"
    frmBas.Tag2 = "A|T|S|||transferencias|Descripcion||S|"

    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 100
   
    
    frmBas.tabla = "importnavconceptos"
    frmBas.CampoCP = "concepto"
    frmBas.Caption = "Conceptos facturas"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub

Public Sub AyudaImporNavarresCentro(frmBas As frmBasico2, Optional CodActual As String, Optional cWhere As String)
 
    frmBas.CadenaTots = "S|txtAux(0)|T|Código|1005|;S|txtAux(1)|T|Descripcion|3000|;"
    frmBas.CadenaConsulta = "select CodCentro,descripcion from importnavcentros"
    
    
    If cWhere <> "" Then frmBas.CadenaConsulta = frmBas.CadenaConsulta & " and " & cWhere
    frmBas.Tag1 = "Codigo|N|N|0||importnavcentros|codigo|000|S|"
    frmBas.Tag2 = "A|T|S|||importnavcentros|anyo||S|"
    
    
    frmBas.Maxlen1 = 10
    frmBas.Maxlen2 = 100
    
    
    frmBas.tabla = "importnavcentros"
    frmBas.CampoCP = "CodCentro"
    frmBas.Caption = "Centros Consum"
    frmBas.DeConsulta = True
    frmBas.DatosADevolverBusqueda = "0|1|"
    frmBas.CodigoActual = 0
    If CodActual <> "" Then frmBas.CodigoActual = CodActual
    frmBas.Show vbModal
    
End Sub


