VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmvisreport 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
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
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmvisreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim Argumentos() As String

Private Sub Form_Load()
On Error GoTo Err_Carga
    
    Dim I As Integer
    Set mapp = CreateObject("CrystalRuntime.Application")
    Set mrpt = mapp.OpenReport("C:\Programas\Conta\Contabilidad\InformesD\Sumas1.rpt")
    
    For I = 1 To mrpt.Database.Tables.Count
        mrpt.Database.Tables(I).SetLogOnInfo "", "", vConfig.User, vConfig.password
    Next I
    
'    If Me.NumeroParametros > 0 Then
'    For I = 1 To Me.NumeroParametros
'            'CR1.Formulas(i) = RecuperaValor(Me.OtrosParametros, i)
'            mrpt.FormulaFields(I) = RecuperaValor(Me.OtrosParametros, I)
'        Next I
'    End If
'    CRViewer1.EnableGroupTree = MostrarTree
            
        
   ' mrpt.RecordSelectionFormula = FormulaSeleccion
    CRViewer1.ReportSource = mrpt
     
    CRViewer1.ViewReport
    Exit Sub
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
