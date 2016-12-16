Attribute VB_Name = "libCalendar"
Public Sub OpenProvider(ByVal eDataProviderType As CodeJockCalendarDataType, ByVal strConnectionString As String, vHacerGetDSN As Boolean)
    
    Set m_pCustomDataHandler = Nothing
    
    ' SQL Server provider.   Abria que traer el modulo de clase que lo gestiona
    'If eDataProviderType = cjCalendarData_SQLServer Then
    '    Set m_pCustomDataHandler = New providerSQLServer
        '' Create DSN "Calendar_SQLServer" to connect to SQL Server Calendar DB
    '    m_pCustomDataHandler.OpenDB strConnectionString
        
    '    m_pCustomDataHandler.SetCalendar CalendarControl
    'End If
    
    ' MySQL provider
    If eDataProviderType = cjCalendarData_MySQL Then
        Set m_pCustomDataHandler = New providerMySQL
        m_pCustomDataHandler.OpenDB strConnectionString, vHacerGetDSN
        
        m_pCustomDataHandler.SetCalendar CalendarControl
    End If
                
    
    'Si pongo PROVIDER=Custom funciona bien, aunque en el connection string le haya dicho la empresa que es
    CalendarControl.SetDataProvider strConnectionString
    CalendarControl.SetDataProvider "Provider=Custom;DSN=vAriges"
    If eDataProviderType = cjCalendarData_SQLServer Or eDataProviderType = cjCalendarData_MySQL Then
        CalendarControl.DataProvider.CacheMode = xtpCalendarDPCacheModeOnRepeat
    End If
    
    If Not CalendarControl.DataProvider.Open Then
        CalendarControl.DataProvider.Create
    End If
    
    m_eActiveDataProvider = eDataProviderType
        
    CalendarControl.Populate
    wndDatePicker.RedrawControl

End Sub

