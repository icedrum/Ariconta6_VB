Attribute VB_Name = "modBackup"
Option Explicit


Public Sub BACKUP_TablaIzquierda(ByRef RS As ADODB.Recordset, ByRef CADENA As String)
Dim I As Integer
Dim nexo As String

    CADENA = ""
    nexo = ""
    For I = 0 To RS.Fields.Count - 1
        CADENA = CADENA & nexo & RS.Fields(I).Name
        nexo = ","
    Next I
    CADENA = "(" & CADENA & ")"
End Sub





'---------------------------------------------------
'El fichero siempre sera NF
Public Sub BACKUP_Tabla(ByRef RS As ADODB.Recordset, ByRef Derecha As String)
Dim I As Integer
Dim nexo As String
Dim Valor As String
Dim Tipo As Integer


    On Error GoTo EBACKUP

    Derecha = ""
    nexo = ""
    For I = 0 To RS.Fields.Count - 1
        Tipo = RS.Fields(I).Type
        
        If IsNull(RS.Fields(I)) Then
            Valor = "NULL"
        Else
            
            'pruebas
            Select Case Tipo
            'TEXTO
            Case 129, 200, 201
                Valor = RS.Fields(I)
                NombreSQL Valor    '.-----------> 23 Octubre 2003.
                Valor = "'" & Valor & "'"
            'Fecha
            Case 133
                Valor = CStr(RS.Fields(I))
                Valor = "'" & Format(Valor, FormatoFecha) & "'"
            
            Case 135
                Valor = CStr(RS.Fields(I))
                Valor = "'" & Format(Valor, "yyyy-mm-dd hh:mm:ss") & "'"
            'Numero normal, sin decimales
            Case 2, 3, 16 To 19
                Valor = RS.Fields(I)
            
            'Numero con decimales
            Case 131
                Valor = CStr(RS.Fields(I))
                Valor = TransformaComasPuntos(Valor)
            Case Else
                Valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                Valor = Valor & vbCrLf & "SQL: " & RS.Source
                Valor = Valor & vbCrLf & "Pos: " & I
                Valor = Valor & vbCrLf & "Campo: " & RS.Fields(I).Name
                Valor = Valor & vbCrLf & "Valor: " & RS.Fields(I)
                MsgBox Valor, vbExclamation
                MsgBox "El programa finalizará. Avise al soporte técnico.", vbCritical
                End
            End Select
        End If
        Derecha = Derecha & nexo & Valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
    
    
    Exit Sub
    
EBACKUP:
    MuestraError Err.Number, "Tipo dato: " & Tipo & "     Valor: " & Valor & vbCrLf & vbCrLf & RS.Source
End Sub
