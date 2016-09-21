VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReportN 
   Caption         =   "Visor de informes"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   5925
   Icon            =   "frmVisReportN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5925
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   615
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   540
      Width           =   4695
      lastProp        =   600
      _cx             =   8281
      _cy             =   5318
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
      EnableRefreshButton=   -1  'True
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
Attribute VB_Name = "frmVisReportN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const NumeroConta = 1  'Esto estara en el parametro


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
'para saber si ha pulsado imprimir
Public Desde As Integer   'Si desde =0 esta cancelado
Public Hasta As Integer   'Si es -1 es que no se especfica, es decir, TODAS

'Dim Argumentos() As String
Dim PrimeraVez As Boolean




Private Sub Command1_Click()
    If PrimeraVez Then Exit Sub
     Unload Me
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
Dim i As Long
'    If True Then Exit Sub
    UseDefault = False
    If NumCopias2 <> 1 Then
        mrpt.PrintOut False, CInt(NumCopias2)
    Else
        i = mrpt.PrinterSetupEx(Me.hWnd)
        If i = 0 Then mrpt.PrintOut False
    End If

'    UseDefault = False
'
'
'    Me.Desde = 1
'    frmImprimirSel.Text1 = Printer.DeviceName
'    frmImprimirSel.Text2(0).Text = NumCopias2 'numcopias
'    frmImprimirSel.Text2(1).Text = 1 'desde
'    frmImprimirSel.Text2(2).Text = "" 'hasta
'    frmImprimirSel.Show vbModal
'
'
'
'    If Desde > 0 Then
'        'NO HAN CANCELADO
'
'        If Hasta > 0 Then
'            mrpt.PrintOut False, CInt(NumCopias2), False, Desde, Hasta
'        Else
'            mrpt.PrintOut False, CInt(NumCopias2), False, Desde
'        End If
'    End If
    
End Sub



Private Function PuedoCerrar(SegundoIncial As Single) As Boolean
Dim C As Integer
    PuedoCerrar = False
    If Not mrpt Is Nothing Then
        C = mrpt.PrintingStatus.Progress
        Debug.Print Now & " e:" & C
    Else
        C = 1
    End If
    
    If C = 2 Then
        DoEvents
        If Timer - SegundoIncial < 20 Then
            Screen.MousePointer = vbHourglass
            espera 1
            'If Timer - SegundoIncial > 5 Then
        Else
            PuedoCerrar = True
        End If
    Else
        PuedoCerrar = True
    End If
End Function


Private Sub Form_Activate()
Dim Incio As Single
Dim Fin As Boolean
    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
           
        
            Screen.MousePointer = vbHourglass
            If SoloImprimir Then
                Incio = Timer
                Do
                    Fin = PuedoCerrar(Incio)
                Loop Until Fin
                Set mrpt = Nothing
                Set mapp = Nothing
            End If
            Unload Me
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim i As Integer
Dim J As Integer
Dim NomImpre As String

    On Error GoTo Err_Carga
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    Screen.MousePointer = vbHourglass
    Set mapp = CreateObject("CrystalRuntime.Application")
    Set mrpt = mapp.OpenReport(Informe)
       
    Me.Command1.Top = 20000
       
    'Conectar a la BD de la Empresa
    For i = 1 To mrpt.Database.Tables.Count
    
        'NUEVO 21 Mayo 2008
        'Puede que alguna tabla este vinculada a ARICONTA
        If LCase(CStr(mrpt.Database.Tables(i).ConnectionProperties.Item("DSN"))) = "ariconta6" Then
            'A conta
            mrpt.Database.Tables(i).SetLogOnInfo "Ariconta6", "ariconta" & vEmpresa.codempre
            'If (InStr(1, mrpt.Database.Tables(i).Name, "_") = 0) Then
            If RedireccionamosTabla(CStr(mrpt.Database.Tables(i).Name)) Then
               mrpt.Database.Tables(i).Location = "ariconta" & vEmpresa.codempre & "." & mrpt.Database.Tables(i).Name
            End If
    
    
        End If
    Next i
    
    'Con LOGO pequeño de 1.50x1.50 cms
    
    AbrirSubreportNuevo
    
    PrimeraVez = True
    
    CargaArgumentos
    
    mrpt.RecordSelectionFormula = FormulaSeleccion

    
    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If
    
     'lOS MARGENES
'    PonerMargen
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    
    
    EstaImpreso = False
    
    CRViewer1.ReportSource = mrpt
   
    If SoloImprimir Then
        If NumCopias2 = 0 Then
            mrpt.PrintOut False
        Else
            mrpt.PrintOut False, NumCopias2
        End If
        EstaImpreso = True
        
    Else
        CRViewer1.ViewReport
    End If
    

    
    Exit Sub
    
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical
    Set mapp = Nothing
    Set mrpt = Nothing
    Set smrpt = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

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
Select Case NumeroParametros
Case 0
    '====Comenta: LAura
    'Solo se vacian los campos de formula que empiezan con "p" ya que estas
    'formulas se corresponden con paso de parametros al Report
    For i = 1 To mrpt.FormulaFields.Count
        If Left(Mid(mrpt.FormulaFields(i).Name, 3), 1) = "p" Then
            mrpt.FormulaFields(i).Text = """"""
        End If
    Next i
    '====
Case 1
    
    For i = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(i).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        'Debug.Print Parametro
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(i).Text = Parametro
        Else
'            mrpt.FormulaFields(I).Text = """"""
        End If
    Next i
    
Case Else
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
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
    Set smrpt = Nothing

End Sub


Private Function DevuelveValor(ByRef Valor As String) As Boolean
Dim i As Long
Dim J As Long

    Valor = "|" & Valor & "="
    DevuelveValor = False
    i = InStr(1, OtrosParametros, Valor, vbTextCompare)
    If i > 0 Then
        i = i + Len(Valor)
        J = InStr(i, OtrosParametros, "|")
        If J > 0 Then
            Valor = Mid(OtrosParametros, i, J - i)
            If Valor = "" Then
                Valor = " "
            Else
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
    mrpt.Export False

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
                     If smrpt.Database.Tables(i).ConnectionProperties.Item("DSN") = "Ariconta6" Then
                        smrpt.Database.Tables(i).SetLogOnInfo "Ariconta6", "ariconta" & vEmpresa.codempre
                        'If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                        If RedireccionamosTabla(CStr(smrpt.Database.Tables(i).Name)) Then smrpt.Database.Tables(i).Location = "ariconta" & vEmpresa.codempre & "." & smrpt.Database.Tables(i).Name
                     End If
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
