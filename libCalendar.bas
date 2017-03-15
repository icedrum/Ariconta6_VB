Attribute VB_Name = "libCalendar"
'/// Si falta algo

Public Sub FillStandardDurations_0m_2w(cmbDuratio2 As ComboBox, bSnoozeBox As Boolean)
    
    If Not bSnoozeBox Then
        cmbDuratio2.AddItem "0 minutos"
        cmbDuratio2.AddItem "1 minuto"
    End If
    
    cmbDuratio2.AddItem "5 minutos"
    cmbDuratio2.AddItem "10 minutos"
    cmbDuratio2.AddItem "15 minutos"
    cmbDuratio2.AddItem "30 minutos"
    
    cmbDuratio2.AddItem "1 hora"
    cmbDuratio2.AddItem "2 horas"
    cmbDuratio2.AddItem "4 horas"
    cmbDuratio2.AddItem "8 horas"
    
    cmbDuratio2.AddItem "0.5 dia"
    cmbDuratio2.AddItem "1 dia"
    cmbDuratio2.AddItem "2 dias"
    cmbDuratio2.AddItem "3 dias"
    cmbDuratio2.AddItem "4 dias"
    
    cmbDuratio2.AddItem "1 semana"
    cmbDuratio2.AddItem "2 semanas"
End Sub

Public Function CalcStandardDurations_0m_2wLong(sDuration As String) As Long
    Select Case sDuration
        Case "0 minutos":
            CalcStandardDurations_0m_2wLong = 0
        Case "1 minuto":
            CalcStandardDurations_0m_2wLong = 1
        Case "5 minutos":
            CalcStandardDurations_0m_2wLong = 5
        Case "10 minutos":
            CalcStandardDurations_0m_2wLong = 10
        Case "15 minutos":
            CalcStandardDurations_0m_2wLong = 15
        Case "30 minutos":
            CalcStandardDurations_0m_2wLong = 30
        
        Case "1 hora":
            CalcStandardDurations_0m_2wLong = 60
        Case "2 horas":
            CalcStandardDurations_0m_2wLong = 60 * 2
        Case "4 horas":
            CalcStandardDurations_0m_2wLong = 60 * 4
        Case "8 horas":
            CalcStandardDurations_0m_2wLong = 60 * 8
        
        Case "0.5 dia":
            CalcStandardDurations_0m_2wLong = 60 * 12
        Case "1 dia":
            CalcStandardDurations_0m_2wLong = 60 * 24
        Case "2 dias":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 2
        Case "3 dias":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 3
        Case "4 dias":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 4
        
        Case "1 semana":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 7
        Case "2 semanas":
            CalcStandardDurations_0m_2wLong = 60 * 24 * 7 * 2
    End Select
End Function

Public Function CalcStandardDurations_0m_2wString(lDuration As Long) As String
    Select Case lDuration
        Case 0:
            CalcStandardDurations_0m_2wString = "0 minutos"
        Case 1:
            CalcStandardDurations_0m_2wString = "1 minuto"
        Case 5:
            CalcStandardDurations_0m_2wString = "5 minutos"
        Case 10:
            CalcStandardDurations_0m_2wString = "10 minutos"
        Case 15:
            CalcStandardDurations_0m_2wString = "15 minutos"
        Case 30:
            CalcStandardDurations_0m_2wString = "30 minutos"
        
        Case 60:
            CalcStandardDurations_0m_2wString = "1 hora"
        Case (60 * 2):
            CalcStandardDurations_0m_2wString = "2 horas"
        Case (60 * 4):
            CalcStandardDurations_0m_2wString = "4 horas"
        Case (60 * 8):
            CalcStandardDurations_0m_2wString = "8 horas"
        
        Case (60 * 12):
            CalcStandardDurations_0m_2wString = "0.5 dia"
        Case (60 * 24):
            CalcStandardDurations_0m_2wString = "1 dia"
        Case (60 * 24 * 2):
            CalcStandardDurations_0m_2wString = "2 dias"
        Case (60 * 24 * 3):
            CalcStandardDurations_0m_2wString = "3 dias"
        Case (60 * 24 * 4):
            CalcStandardDurations_0m_2wString = "4 dias"
        
        Case (60 * 24 * 7):
            CalcStandardDurations_0m_2wString = "1 semana"
        Case (60 * 24 * 7 * 2):
            CalcStandardDurations_0m_2wString = "2 semanas"
    End Select
End Function

Public Function FormatTimeDuration(ByVal nMinutes As Long, ByVal bAprox As Boolean) As String
    Dim nWeeks As Long, nDays As Long, nHours As Long
    
    nWeeks = nMinutes / (7 * 24 * 60)
    nDays = nMinutes / (24 * 60)
    nHours = nMinutes / 60

    Dim strDuration As String
    
    If (bAprox Or (nMinutes Mod (7 * 24 * 60)) = 0) And nWeeks > 0 Then
        strDuration = nWeeks & " semana" & IIf(nWeeks > 1, "s", "")
    
    ElseIf (bAprox Or (nMinutes Mod (24 * 60)) = 0) And nDays > 0 Then
        strDuration = nDays & " dia" & IIf(nDays > 1, "s", "")
        
    ElseIf (bAprox Or (nMinutes Mod 60) = 0) And nHours > 0 Then
        strDuration = nHours & " hora" & IIf(nHours > 1, "s", "")
        
    Else
        strDuration = nMinutes & " minuto" & IIf(nMinutes > 1, "s", "")
    End If

    FormatTimeDuration = strDuration
End Function




Public Function ParseTimeDuration(ByVal strTime As String, ByRef pnMinutes As Long) As Boolean
    pnMinutes = 0
    ParseTimeDuration = False
        
    Dim nI As Long, nLen As Long
    Dim nMeasureStart As Long, nFIdx As Long
    Dim strChI As String
        
    strTime = Trim(strTime)
    nLen = Len(strTime)
    
    If nLen = 0 Then
        Exit Function
    End If
    
    '------------------------------------------
    nMeasureStart = -1
    For nI = 1 To nLen
        strChI = Mid(strTime, nI, 1)
        nFIdx = InStr(1, "-+.,0123456789", strChI)
        If nFIdx <= 0 Then
            nMeasureStart = nI
            Exit For
        End If
    Next
    
    Dim strNumber As String, strMeasure As String
    Dim nMultiplier As Long
            
    If nMeasureStart > 0 Then
        strNumber = Left(strTime, nMeasureStart - 1)
        strMeasure = Mid(strTime, nMeasureStart)
        strMeasure = Trim(strMeasure)
    Else
        strNumber = strTime
    End If
    
    If Len(strNumber) = 0 Then
        Exit Function
    End If

    Dim strM0 As String
    strM0 = Left(strMeasure, 1)
    
    nMultiplier = 1
    If strM0 = "m" Or strM0 = "M" Then
        nMultiplier = 1
    ElseIf strM0 = "h" Or strM0 = "H" Then
        nMultiplier = 60
    ElseIf strM0 = "d" Or strM0 = "D" Then
        nMultiplier = 60 * 24
    ElseIf strM0 = "w" Or strM0 = "W" Then
        nMultiplier = 60 * 24 * 7
    End If

    Dim dblTime As Double
    dblTime = Val(strNumber)
    
    pnMinutes = dblTime * nMultiplier
    ParseTimeDuration = True
End Function

