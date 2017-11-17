Attribute VB_Name = "libIBAN"
Option Explicit




'A partir de una cuenta banco formateada y todos los numeros juntos (chr(20))
'  devuelve DOS(2) caracteres del IBAN
'  calculados como dice la formula
'  i=ctabanco_con ES... mod 97
'  i = 98-i
' format(i,"00"             'para que copie                     'es lo que devuelve
'
'Puede NO poner pais. Sera ES
Public Function DevuelveIBAN2(PAIS As String, ByVal CtaBancoFormateada As String, DosCaracteresIBAN As String) As Boolean
Dim Aux As String
Dim N As Long
Dim CadenaPais As String
On Error GoTo EDevuelveIBAN
    DevuelveIBAN2 = False
    DosCaracteresIBAN = ""
    
    
    
    If PAIS = "" Then
        PAIS = "ES"
    Else
        If Len(PAIS) <> 2 Then
            PAIS = "ES"
        Else
            PAIS = UCase(PAIS)
        End If
    End If
    
    
    'Ejemplo mio: 20770294901101867914  IBAN: 41
    'Construir el IBAn:
    'A la derecha de la cuenta se pone
    '   el ES00-->   20770294961101915202ES00 ->92
    'Se transforma las letras ES a numeros.
    ' E=14 S=28
    '           ->>> 20770294961101915202 142800
    If PAIS = "ES" Then
        CadenaPais = "1428"
    Else
        N = Asc(Mid(PAIS, 1, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CStr(N)
        N = Asc(Mid(PAIS, 2, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CadenaPais & CStr(N)
    End If
    'Se le añaden 2 ceros al final
    CadenaPais = CadenaPais & "00"
    'Esta es la cadena para ES. SiCadenaPais  fuera otro pais es aqui donde hay que cambiar
    CtaBancoFormateada = CtaBancoFormateada & CadenaPais
    Aux = ""
    While CtaBancoFormateada <> ""
        If Len(CtaBancoFormateada) >= 6 Then
            Aux = Aux & Mid(CtaBancoFormateada, 1, 6)
            CtaBancoFormateada = Mid(CtaBancoFormateada, 7)
        Else
            Aux = Aux & CtaBancoFormateada
            CtaBancoFormateada = ""
        End If
        
        N = CLng(Aux)
        N = N Mod 97
        
        Aux = CStr(N)
    Wend
        
    N = 98 - N
    
    DosCaracteresIBAN = Format(N, "00")
    DevuelveIBAN2 = True
    Exit Function
EDevuelveIBAN:
    CadenaPais = Err.Description
    CadenaPais = Err.Number & "   " & CadenaPais
    MsgBox "Devuelve IBAN. " & vbCrLf & CadenaPais, vbExclamation
    Err.Clear
End Function




Public Function IBAN_Correcto(IBAN As String) As Boolean
Dim Aux As String
    IBAN_Correcto = False
    Aux = ""
    If Len(IBAN) <> 4 Then
        Aux = "Longitud incorrecta"
    Else
        If IsNumeric(Mid(Aux, 3, 2)) Then
            Aux = "Digitos 3 y 4 deben ser numericos"
        Else
            'Podriamos comprobar lista de paises
    
        End If
    End If
    If Aux <> "" Then
        MsgBox Aux, vbExclamation
    Else
        IBAN_Correcto = True
    End If
End Function




'A partir de una cadena, con letras y numeros convertira
'en mod 97,10 Norma ISO 7064
'Para ello los caracteres se pasan a dos digitos
Public Function CadenaTextoMod97(CADENA As String) As String
Dim i As Integer
Dim C As String
Dim N As Long

    CADENA = Trim(CADENA)
    C = ""
    'Substitucion de texto por caracteres
    For i = 1 To Len(CADENA)
        N = Asc(Mid(CADENA, i, 1))
        If N >= 48 Then
            If N <= 57 Then
                'Es numerico 0..9
                'C = C & CStr(N)
            Else
                If N < 65 Or N > 90 Then
                    'MAL. No es un caracter ASCII entre A..Z  (10..35)
                    N = 0
                Else
                    N = N - 55  'el ascci menos 55 (0...35)
                End If
            End If
        End If
        If N = 0 Then
            CadenaTextoMod97 = "Caracter NO valido: " & Mid(CADENA, i, 1) & " --- " & CADENA
            Exit Function
        Else
            If N >= 48 Then
                'Es un numero
                C = C & Chr(N)
            Else
                C = C & CStr(N)
            End If
        End If
        
    Next
    
    
    
    'Ya tengo C que es numerica
    CADENA = C
    C = ""
    While CADENA <> ""
        If Len(CADENA) >= 6 Then
            C = C & Mid(CADENA, 1, 6)
            CADENA = Mid(CADENA, 7)
        Else
            C = C & CADENA
            CADENA = ""
        End If
        
        N = CLng(C)
        N = N Mod 97
        
        C = CStr(N)
    Wend
        
    N = 98 - N
    CadenaTextoMod97 = Format(N, "00")
End Function




'Validador para todos los paises
Public Function EsIBAN_Correcto(ByVal CtaBancoFormateada As String) As Boolean
Dim Aux As String
Dim N As Long
Dim CadenaPais As String
Dim PAIS As String
Dim Control As String

On Error GoTo EDevuelveIBAN

    EsIBAN_Correcto = False
    
    CtaBancoFormateada = Replace(CtaBancoFormateada, " ", "")
    PAIS = Mid(CtaBancoFormateada, 1, 2)
    Control = Mid(CtaBancoFormateada, 3, 2)
    If PAIS = "" Then
        PAIS = "ES"
    Else
        If Len(PAIS) <> 2 Then
            PAIS = "ES"
        Else
            PAIS = UCase(PAIS)
        End If
    End If
    
    
    'Ejemplo mio: 20770294901101867914  IBAN: 41
    'Construir el IBAn:
    'A la derecha de la cuenta se pone
    '   el ES00-->   20770294961101915202ES00 ->92
    'Se transforma las letras ES a numeros.
    ' E=14 S=28
    '           ->>> 20770294961101915202 142800
    If PAIS = "ES" Then
        CadenaPais = "1428"
    Else
        N = Asc(Mid(PAIS, 1, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CStr(N)
        N = Asc(Mid(PAIS, 2, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CadenaPais & CStr(N)
    End If
    'Se le añaden 2 ceros al final
    CadenaPais = CadenaPais & "00"
    'Esta es la cadena para ES. SiCadenaPais  fuera otro pais es aqui donde hay que cambiar
    CtaBancoFormateada = Mid(CtaBancoFormateada, 5) & CadenaPais
    
    CadenaPais = ""
    For i = 1 To Len(CtaBancoFormateada)
        Aux = Mid(CtaBancoFormateada, i, 1)
        If Not IsNumeric(Aux) Then Aux = DevuelveLetraNumero(Aux)
        CadenaPais = CadenaPais & Aux
    Next
    CtaBancoFormateada = CadenaPais
    Aux = ""
    While CtaBancoFormateada <> ""
        If Len(CtaBancoFormateada) >= 6 Then
            Aux = Aux & Mid(CtaBancoFormateada, 1, 6)
            CtaBancoFormateada = Mid(CtaBancoFormateada, 7)
        Else
            Aux = Aux & CtaBancoFormateada
            CtaBancoFormateada = ""
        End If
        
        N = CLng(Aux)
        N = N Mod 97
        
        Aux = CStr(N)
    Wend
        
    N = 98 - N
    
    Aux = Format(N, "00")
    EsIBAN_Correcto = Aux = Control
    Exit Function
EDevuelveIBAN:
    CadenaPais = Err.Description
    CadenaPais = Err.Number & "   " & CadenaPais
    MsgBox "Devuelve IBAN. " & vbCrLf & CadenaPais, vbExclamation
    Err.Clear
End Function


'A=10 B=11 C=12 D=13 E=14 F=15 G=16 H=17 I=18 J=19 K=20 L=21 M=22 N=23 O=24 P=25 Q=26 R=27 S=28 T=29U=30 V=31
'W=32 X=33 Y=34 Z=35
Private Function DevuelveLetraNumero(LEtra As String) As String
Dim N As Integer
    N = Asc(Mid(LEtra, 1, 1))
    If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra en cadena IBAN " & LEtra
    DevuelveLetraNumero = CStr(N - 55)
End Function






'Devolvera el texto el IBAN separado en bloques de 4 , siempre que sea ES?????
Public Function DevuelveIBANSeparado(ByVal IBAN_junto As String) As String
Dim J As Integer

    IBAN_junto = Trim(IBAN_junto)
    If IBAN_junto = "" Then
        DevuelveIBANSeparado = ""
    Else
        DevuelveIBANSeparado = IBAN_junto
        If UCase(Mid(IBAN_junto, 1, 2)) = "ES" And Len(IBAN_junto) = 24 Then
            DevuelveIBANSeparado = Mid(IBAN_junto, 1, 4) & " " & Mid(IBAN_junto, 5, 4) & " " & Mid(IBAN_junto, 9, 4) & " " & Mid(IBAN_junto, 13, 4) & " " & Mid(IBAN_junto, 17, 4) & " " & Mid(IBAN_junto, 21, 4)
        End If
    End If
End Function
