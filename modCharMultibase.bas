Attribute VB_Name = "modCharMultibase"



Public Function RevisaCaracterMultibase(CADENA As String) As String
Dim I As Integer
Dim J As Integer
Dim L As String
Dim C As String

    L = ""
    For I = 1 To Len(CADENA)
        C = Mid(CADENA, I, 1)
        J = Asc(C)
        If J > 125 Then
            Select Case J
            Case 128
                C = "Ç"
            Case 164  'ñ minuscula
                C = "ñ"
            Case 165
                'Es la Ñ
                C = "Ñ"
            Case 166
                C = "ª"
            Case 167, 186
                C = "º"
            Case 209
            
            Case Else
                
            End Select
        End If
        L = L & C
    Next I
    RevisaCaracterMultibase = L

End Function
