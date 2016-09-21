VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport 
   Caption         =   "Visor de informes"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "frmVisReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   14775
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      lastProp        =   600
      _cx             =   14208
      _cy             =   9551
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Informe As String
Public ConSubInforme As Boolean 'Si tiene subinforme ejecta la funcion AbrirSubInforme para enlazar esta a la BD correspondiente

'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean
Public opcion As Integer
Public ExportarPDF As Boolean

Public EstaImpreso As Boolean
Public NumCopias2 As Integer


Dim mapp As CRAXDRT.Application
Public mrpt As CRAXDRT.Report
Dim smrpt As CRAXDRT.Report


Dim Argumentos() As String
Dim PrimeraVez As Boolean

Private Sub CRViewer1_CloseButtonClicked(UseDefault As Boolean)
    UseDefault = False
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    If True Then Exit Sub
    UseDefault = False
    'mrpt.PrinterSetup Me.hWnd
    mrpt.PrinterSetupEx Me.hWnd
    'Para que no pregunte
    mrpt.PrintOut False
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
            Screen.MousePointer = vbHourglass
            Unload Me
            
        Else
'            PonerFocoBtn Me.Command1
'            Dim LoSabe As Boolean
'            Dim Pag As Long
'
'            CRViewer1.ShowLastPage
'            CRViewer1.GetLastPageNumber Pag, LoSabe
'            'Exportar
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

On Error GoTo Err_Carga
        Me.Icon = frmPpal.Icon
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    Set mapp = CreateObject("CrystalRuntime.Application")
    'Informe = "C:\Programas\Conta\Contabilidad\Informes\sumas12.rpt"
    Set mrpt = mapp.OpenReport(Informe)
    
'--
'    For I = 1 To mrpt.Database.Tables.Count
'       mrpt.Database.Tables(I).SetLogOnInfo "vUsuarios", "Usuarios", vConfig.User, vConfig.password
'    Next I

'++
    'Conectar a la BD de la Empresa
    For i = 1 To mrpt.Database.Tables.Count

        'NUEVO 21 Mayo 2008
        'Puede que alguna tabla este vinculada a ARICONTA
        If CStr(mrpt.Database.Tables(i).ConnectionProperties.Item("DSN")) = "Ariconta6" Then
            'A conta
            mrpt.Database.Tables(i).SetLogOnInfo "Ariconta6", "ariconta" & vEmpresa.codempre ', vConfig.User, vConfig.password
            'If (InStr(1, mrpt.Database.Tables(i).Name, "_") = 0) Then
            If RedireccionamosTabla(CStr(mrpt.Database.Tables(i).Name)) Then
               mrpt.Database.Tables(i).Location = "ariconta" & vEmpresa.codempre & "." & mrpt.Database.Tables(i).Name
            End If
        Else
            mrpt.Database.Tables(i).SetLogOnInfo "vUsuarios", "Usuarios" ', vConfig.User, vConfig.password
        End If
    Next i
'++
    
    'Con LOGO pequeño de 1.50x1.50 cms
    
    AbrirSubreportNuevo
    

    PrimeraVez = True
    CargaArgumentos
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    
    mrpt.RecordSelectionFormula = FormulaSeleccion
    
    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If
    
    'lOS MARGENES
    PonerMargen
    
    CRViewer1.ReportSource = mrpt
    If SoloImprimir Then
        If NumCopias2 = 0 Then
            mrpt.PrintOut False
        Else
            mrpt.PrintOut False, NumCopias2
        End If
    Else
        CRViewer1.ViewReport
    End If
    Exit Sub

Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical
    Set mapp = Nothing
    Set mrpt = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

'Private Sub CargaArgumentos()
'Dim Parametro As String
'Dim I As Integer
'    'El primer parametro es el nombre de la empresa para todas las empresas
'    ' Por lo tanto concaatenaremos con otros parametros
'    ' Y sumaremos uno
'    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
'    ' La cadena de parametros
'    'Si esta asignaremos su valor
'
'    OtrosParametros = "|Emp= """ & vEmpresa.nomempre & """|" & OtrosParametros
'    NumeroParametros = NumeroParametros + 1
'
'    For I = 1 To mrpt.FormulaFields.Count
'        Parametro = mrpt.FormulaFields(I).Name
'        Parametro = Mid(Parametro, 3)  'Quitamos el {@
'        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
'        'Debug.Print Parametro
'        If DevuelveValor(Parametro) Then mrpt.FormulaFields(I).Text = Parametro
'        'Debug.Print " -- " & Parametro
'    Next I
'
'End Sub

Private Sub CargaArgumentos()
Dim Parametro As String
Dim i As Integer
    'El primer parametro es el nombre de la empresa para todas las empresas
    ' Por lo tanto concaatenaremos con otros parametros
    ' Y sumaremos uno
    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
    ' La cadena de parametros
    'Si esta asignaremos su valor
    
'    OtrosParametros = "|Emp= """ & vEmpresa.nomempre & """|" & OtrosParametros
    NumeroParametros = NumeroParametros + 1
    
    For i = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(i).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(i).Text = Parametro
        End If
    Next i
'    mrpt.RecordSelectionFormula

End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
End Sub


Private Function DevuelveValor(ByRef Valor As String) As Boolean
Dim i As Integer
Dim J As Integer
    Valor = "|" & Valor & "="
    DevuelveValor = False
    i = InStr(1, OtrosParametros, Valor, vbTextCompare)
    If i > 0 Then
        i = i + Len(Valor) + 1
        J = InStr(i, OtrosParametros, "|")
        If J > 0 Then
            Valor = Mid(OtrosParametros, i, J - i)
            If Valor = "" Then
                Valor = " "
            Else
                'Si no tiene el salto
                If InStr(1, Valor, "chr(13)") = 0 Then CompruebaComillas Valor
            End If
            DevuelveValor = True
        End If
    End If
End Function


Private Sub CompruebaComillas(ByRef Valor1 As String)
Dim AUx As String
Dim J As Integer
Dim i As Integer

    If Mid(Valor1, 1, 1) = Chr(34) Then
        'Tiene comillas. Con lo cual tengo k poner las dobles
        AUx = Mid(Valor1, 2, Len(Valor1) - 2)
        i = -1
        Do
            J = i + 2
            i = InStr(J, AUx, """")
            If i > 0 Then
              AUx = Mid(AUx, 1, i - 1) & """" & Mid(AUx, i)
            End If
        Loop Until i = 0
        AUx = """" & AUx & """"
        Valor1 = AUx
    End If
End Sub

Private Sub Exportar()
    
    mrpt.ExportOptions.DiskFileName = App.Path & "\docum.pdf"
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    mrpt.ExportOptions.PDFExportAllPages = True
    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
    mrpt.DisplayProgressDialog = False
    mrpt.Export False
    'Si ha generado bien entonces
    CadenaDesdeOtroForm = "OK"
End Sub

Private Sub PonerMargen()
Dim Cad As String
Dim i As Integer
    On Error GoTo EPon
    Cad = Dir(App.Path & "\*.mrg")
    If Cad <> "" Then
        i = InStr(1, Cad, ".")
        If i > 0 Then
            Cad = Mid(Cad, 1, i - 1)
            If IsNumeric(Cad) Then
                If Val(Cad) > 4000 Then Cad = "4000"
                If Val(Cad) > 0 Then
                    mrpt.BottomMargin = mrpt.BottomMargin + Val(Cad)
                End If
            End If
        End If
    End If
    
    Exit Sub
EPon:
    Err.Clear
End Sub


Private Sub AbrirSubreportNuevo()
Dim crxSection As CRAXDRT.Section
Dim crxObject As Object
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim i As Byte

    For Each crxSection In mrpt.Sections
        For Each crxObject In crxSection.ReportObjects
             If TypeOf crxObject Is SubreportObject Then
                Set crxSubreportObject = crxObject
                Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
                For i = 1 To smrpt.Database.Tables.Count 'para cada tabla
                    
                        smrpt.Database.Tables(i).SetLogOnInfo "Ariconta6", "ariconta" & vEmpresa.codempre
                        'If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                        If RedireccionamosTabla(CStr(smrpt.Database.Tables(i).Name)) Then smrpt.Database.Tables(i).Location = "ariconta" & vEmpresa.codempre & "." & smrpt.Database.Tables(i).Name
   
                Next i
             End If
        Next crxObject
    Next crxSection
    
    Set crxSubreportObject = Nothing
End Sub


Private Function RedireccionamosTabla(tabla As String) As Boolean
    'If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
    If InStr(1, tabla, "_") = 0 Then
        RedireccionamosTabla = True
    Else
        If Mid(tabla, 1, 3) = "tel" Then
            'tablas telefonia
            RedireccionamosTabla = True
        Else
            'resto
            RedireccionamosTabla = False
        End If
    End If
    
    
End Function

