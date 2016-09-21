Attribute VB_Name = "libFiltro"
Option Explicit





Public Sub CargaVectoresFiltro(NumFiltros As Integer, Textox As String, ByRef cboFiltro_ As ComboBox)  'Vendran empipados
Dim i As Integer
    
    
    cboFiltro_.Clear
    cboFiltro_.AddItem "Sin filtro"
    For i = 1 To NumFiltros
        cboFiltro_.AddItem RecuperaValor(Textox, i)
    Next i
    
    
    
End Sub


'
Public Sub ValorFiltroPorDefecto(Leer As Boolean, NombreForm As String, opcion As Integer, ByRef ColumnaOrden As Integer, ByRef OtroDatos As String, ByRef AscenDes As Boolean, ByRef CadenaParaGuardarFiltro As String)
Dim RN As ADODB.Recordset
Dim Cad As String

    If Leer Then
        Cad = "Select * from usuarios.usuariosvaloresdefecto WHERE"
        Cad = Cad & " aplicacion='ariconta' AND codusu=1" 'IRA codusu
        Cad = Cad & " AND formulario='" & NombreForm & "' AND opcion= "
        Cad = Cad & " " & opcion
            
        Set RN = New ADODB.Recordset
        RN.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RN.EOF Then
            AscenDes = DBLet(RN!ascendescen, "N") = 1
            ColumnaOrden = DBLet(RN!Columna, "N")
            OtroDatos = DBLet(RN!otros, "T")
            CadenaParaGuardarFiltro = ColumnaOrden & "|" & AscenDes & "|" & OtroDatos  'La cadena para comprara al guardar
        Else
            AscenDes = True
            ColumnaOrden = 1
            CadenaParaGuardarFiltro = "-1"
        End If
        RN.Close
        
    Else
        'INSERT
        Cad = ColumnaOrden & "|" & AscenDes & "|" & OtroDatos
        If Cad <> CadenaParaGuardarFiltro Then
            'UPDATE /INSERT
            Cad = "REPLACE usuarios.usuariosvaloresdefecto(aplicacion,codusu,formulario,opcion,columna,otros,ascendescen) VALUES ("
            Cad = Cad & "'ariconta',1,'" & NombreForm & "',"
            Cad = Cad & opcion
            Cad = Cad & "," & ColumnaOrden & ",'" & OtroDatos & "'," & CStr(Abs(AscenDes)) & ")"
            Conn.Execute Cad
            
        End If
    End If

End Sub


Public Function DevuelveFiltro(Prg As Integer, Aplicacion As String) As Integer
Dim Sql As String
    
    Sql = "select filtro from menus_usuarios where codigo = " & Prg
    Sql = Sql & " and aplicacion = " & DBSet(Aplicacion, "T") & " and codusu =" & vUsu.Id
    
    DevuelveFiltro = DevuelveValor(Sql)

End Function

