VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmSII_Avisos 
   Caption         =   "Comunicaci�n datos SII"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   16080
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   6735
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   15615
      _Version        =   1114114
      _ExtentX        =   27543
      _ExtentY        =   11880
      _StockProps     =   64
      MultipleSelection=   0   'False
      FreezeColumnsAbs=   0   'False
      MultiSelectionMode=   -1  'True
   End
   Begin VB.Frame FrameAux 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   15735
      Begin VB.CommandButton cmdExportar 
         Height          =   495
         Left            =   11040
         Picture         =   "frmSII_Avisos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exportar"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdRefrescar 
         Height          =   495
         Left            =   9840
         Picture         =   "frmSII_Avisos.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Refrescar datos"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   6840
         TabIndex        =   3
         Top             =   480
         Width           =   1305
      End
      Begin VB.ComboBox cboFiltro 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSII_Avisos.frx":0F8C
         Left            =   4080
         List            =   "frmSII_Avisos.frx":0F8E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2610
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSII_Avisos.frx":0F90
         Left            =   2400
         List            =   "frmSII_Avisos.frx":0F92
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1650
      End
      Begin VB.ComboBox cboFacturas 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSII_Avisos.frx":0F94
         Left            =   120
         List            =   "frmSII_Avisos.frx":0F9B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   2130
      End
      Begin VB.CommandButton cmdComunicar 
         Height          =   495
         Left            =   9240
         Picture         =   "frmSII_Avisos.frx":0FAB
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Comunicar facturas"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdModificar 
         Height          =   495
         Left            =   10440
         Picture         =   "frmSII_Avisos.frx":19AD
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ver/Modificar factura"
         Top             =   240
         Width           =   495
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   15120
         TabIndex        =   14
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filtro"
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   2
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   1
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
         ForeColor       =   &H00404040&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   630
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   7680
         Picture         =   "frmSII_Avisos.frx":23AF
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6840
         TabIndex        =   10
         Top             =   240
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   540
         Left            =   14280
         Picture         =   "frmSII_Avisos.frx":243A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   540
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   13680
         Picture         =   "frmSII_Avisos.frx":2F93
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   8400
         Picture         =   "frmSII_Avisos.frx":3B7E
         ToolTipText     =   "Seleccionar todo"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   8760
         Picture         =   "frmSII_Avisos.frx":3CC8
         ToolTipText     =   "Quitar seleccion"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "aa"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   11640
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   8130
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSII_Avisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public QueMostrarDeSalida As Byte
    '0  Nada
    '1  Clientes  pdtes
    '2  proveed   pdtes
    '3  ambos
    '
    '11 Enviadas=0  cliente
    '12 Enviadas=0  provee
    '13 ambas
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim PrimeraVez As Boolean



Dim fntBold As StdFont
Dim fntStrike As StdFont

Dim MarcarComoCheck As Boolean
Dim FechaRoja As Date

Dim R340 As ADODB.Recordset
Dim Rwtipopera As ADODB.Recordset

Dim FacturasEnviadasYConCSV As String


'Montifruit NO las quiere marcadas
Dim Marcadas As Boolean



Dim ActualizarAIncorrectos As String
Dim ActualizarAAceptada As String



Private Sub VerTodos()

    CargaDatos

End Sub



Private Sub cboFacturas_Click()
    If PrimeraVez Then Exit Sub
    If cboFacturas.Tag = cboFacturas.ListIndex Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    CargaCombos
    
    MarcarComoCheck = True

    
    PrimeraVez = False
    
    
    CreateReportControl
    CargaDatos
    wndReportControl.Populate
    
    imgCheck(0).visible = cboFacturas.ListIndex = 0
    imgCheck(1).visible = cboFacturas.ListIndex = 0
    Me.cmdModificar.visible = cboFacturas.ListIndex = 0
    Me.cmdComunicar.visible = cboFacturas.ListIndex = 0
    
    cboFacturas.Tag = cboFacturas.ListIndex
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    If cboFiltro.Tag = cboFiltro.ListIndex Then Exit Sub
    Screen.MousePointer = vbHourglass
    CargaDatos
    wndReportControl.Populate
    cboFiltro.Tag = cboFiltro.ListIndex
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboTipo_Click()
     If PrimeraVez Then Exit Sub
    If cboTipo.Tag = cboTipo.ListIndex Then Exit Sub
    Screen.MousePointer = vbHourglass
    CreateReportControl
    Me.Refresh
    Screen.MousePointer = vbHourglass
    CargaDatos
    wndReportControl.Populate
    cboTipo.Tag = cboTipo.ListIndex
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdComunicar_Click()

    J = 0
    
    For i = 0 To Me.wndReportControl.Records.Count - 1
        If wndReportControl.Records(i).Item(0).HasCheckbox Then
            If wndReportControl.Records(i).Item(0).Checked Then J = J + 1
        End If
    Next

    If J = 0 Then
        MsgBox "Seleccione alguna factura", vbExclamation
        Exit Sub
    End If
    
    Msg = ""
    If J > 1 Then Msg = "s"
    Msg = "Va a traspasar " & J & " factura" & Msg & " al programa de comunicaci�n electronica ASWSII(Ariadna Software). " & vbCrLf & "�Continuar?"
    
    If MsgBox(Msg, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub


    If BloqueoManual(True, "SII", "1") Then
        Screen.MousePointer = vbHourglass
        Label True
        Set miRsAux = New ADODB.Recordset
        HacerGrabacionTablas_ASWISII
        Set miRsAux = Nothing
        
        BloqueoManual False, "SII", ""
        
        
        'volvemos a cargar
        DoEvent2
        Screen.MousePointer = vbHourglass

        MarcarComoCheck = False
        CargaDatos
        
        
        Label False
        
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Proceso bloqueado por otro usuario", vbExclamation
    End If


End Sub

Private Sub cmdExportar_Click()
    
    If Me.wndReportControl.Records.Count = 0 Then Exit Sub
    
    ExportarCSV
    
End Sub


Private Sub cmdModificar_Click()
Dim SQL As String
Dim Todo As Boolean

    If Me.wndReportControl.Records.Count = 0 Then Exit Sub
    
    If Me.wndReportControl.SelectedRows(0) Is Nothing Then Exit Sub
    
    If Me.cboFacturas.ListIndex > 0 Then Exit Sub
    
    If Me.cboTipo.ListIndex = 0 Then
    
        frmFacturasCli.FACTURA = Me.wndReportControl.SelectedRows(0).Record.Tag
        frmFacturasCli.Show vbModal
        
        
        'Veremos si ha modificado algo de
        
        
    Else
        frmFacturasPro.FACTURA = Me.wndReportControl.SelectedRows(0).Record.Tag
        frmFacturasPro.Show vbModal
    End If
    
    
    SQL = SQL_
    'Solo esta factura
    i = InStr(1, SQL, " WHERE ")
    Msg = Mid(SQL, i)
    SQL = Mid(SQL, 1, i - 1)
    SQL = SQL & " WHERE numserie = " & DBSet(RecuperaValor(Me.wndReportControl.SelectedRows(0).Record.Tag, 1), "T")
    SQL = SQL & " AND " & IIf(cboTipo.ListIndex = 0, "numfactu", "numregis") & " = "
    SQL = SQL & RecuperaValor(Me.wndReportControl.SelectedRows(0).Record.Tag, 2)
    SQL = SQL & " AND  anofactu = " & RecuperaValor(Me.wndReportControl.SelectedRows(0).Record.Tag, 3)
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        'Ha cambiado algo "importante. Volvemos a cargar todo
        Todo = True
    Else
        Todo = False
        If cboTipo.ListIndex = 0 Then
            'CLiente
            If vParam.SII_Periodo_DesdeLiq Then
                wndReportControl.SelectedRows(0).Record(3).Caption = Format(miRsAux!fecliqcl, "dd/mm/yyyy")
                wndReportControl.SelectedRows(0).Record(3).Value = Format(miRsAux!fecliqcl, "yyyymmdd")
            Else
                wndReportControl.SelectedRows(0).Record(3).Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
                wndReportControl.SelectedRows(0).Record(3).Value = Format(miRsAux!FecFactu, "yyyymmdd")
            End If
            wndReportControl.SelectedRows(0).Record(4).Caption = miRsAux!codmacta
            wndReportControl.SelectedRows(0).Record(5).Caption = miRsAux!Nommacta
            
            Rwtipopera.Find "codigo = " & DBSet(DBLet(miRsAux!CodOpera, "T"), "T"), , adSearchForward, 1
            wndReportControl.SelectedRows(0).Record(6).Caption = DBLet(Rwtipopera!denominacion, "")
            R340.Find "codigo = " & DBSet(DBLet(miRsAux!codconce340, "T"), "T"), , adSearchForward, 1
            wndReportControl.SelectedRows(0).Record(7).Caption = DBLet(R340!Descripcion, "")
            
            wndReportControl.SelectedRows(0).Record(8).Caption = Format(miRsAux!totfaccl, FormatoImporte)
            wndReportControl.SelectedRows(0).Record(8).Value = miRsAux!totfaccl * 100
            
            
        Else
            'Proveedor
            If vParam.SII_Periodo_DesdeLiq Then
                wndReportControl.SelectedRows(0).Record(2).Caption = Format(miRsAux!fecliqpr, "dd/mm/yyyy")
                wndReportControl.SelectedRows(0).Record(2).Value = Format(miRsAux!fecliqpr, "yyyymmdd")
            Else
                If vParam.SII_ProvDesdeFechaRecepcion Then
                    wndReportControl.SelectedRows(0).Record(2).Caption = Format(miRsAux!fecharec, "dd/mm/yyyy")
                    wndReportControl.SelectedRows(0).Record(2).Value = Format(miRsAux!fecharec, "yyyymmdd")
                Else
                    wndReportControl.SelectedRows(0).Record(2).Caption = Format(miRsAux!fecregcontable, "dd/mm/yyyy")
                    wndReportControl.SelectedRows(0).Record(2).Value = Format(miRsAux!fecregcontable, "yyyymmdd")
                End If
            End If
            wndReportControl.SelectedRows(0).Record(3).Value = miRsAux!numfactu
            wndReportControl.SelectedRows(0).Record(4).Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            wndReportControl.SelectedRows(0).Record(4).Value = Format(miRsAux!FecFactu, "yyyymmdd")
            wndReportControl.SelectedRows(0).Record(5).Caption = miRsAux!codmacta
            wndReportControl.SelectedRows(0).Record(6).Caption = miRsAux!Nommacta
            Rwtipopera.Find "codigo = " & DBSet(DBLet(miRsAux!CodOpera, "T"), "T"), , adSearchForward, 1
            wndReportControl.SelectedRows(0).Record(7).Caption = DBLet(Rwtipopera!denominacion, "")
            R340.Find "codigo = " & DBSet(DBLet(miRsAux!codconce340, "T"), "T"), , adSearchForward, 1
            If R340.EOF Then
                wndReportControl.SelectedRows(0).Record(8).Caption = DBLet("N/D", "T")
            Else
                wndReportControl.SelectedRows(0).Record(8).Caption = DBLet(R340!Descripcion, "T")
            End If
            wndReportControl.SelectedRows(0).Record(9).Caption = Format(miRsAux!totfacpr, FormatoImporte)
            wndReportControl.SelectedRows(0).Record(9).Value = miRsAux!totfacpr * 100
    
        End If
        wndReportControl.Populate
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If Todo Then CargaDatos
    
End Sub

Private Sub cmdRefrescar_Click()
    CargaDatos
End Sub

Private Sub Form_Activate()
    Dim Record As ReportRecord

    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        DoEvent2
        CargaDatos
        'CargaDatosPrueba "", False
     
    End If
End Sub


Private Sub AbrirRss()
    Set R340 = New ADODB.Recordset
    R340.Open "Select * from usuarios.wconce340 ", Conn, adOpenKeyset, adCmdText
    
    Set Rwtipopera = New ADODB.Recordset
    Rwtipopera.Open "Select * from usuarios.wtipopera ", Conn, adOpenKeyset, adCmdText

End Sub


Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon


    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With

    AbrirRss

    MarcarComoCheck = True
    wndReportControl.Icons = ReportControlGlobalSettings.Icons
    wndReportControl.PaintManager.NoItemsText = "Ning�n registro "
    EstablecerFuente
    
    
    
    If QueMostrarDeSalida < 10 Then
        'Pendientes de enviar
        Me.cboFacturas.ListIndex = 0
        cboFacturas.Tag = 0
    Else
        'Subidas a ASWII, pero sin evniar a la AEAT
        Me.cboFacturas.ListIndex = 1
        cboFacturas.Tag = 1
    End If
    CargaCombos
    
    If QueMostrarDeSalida < 10 Then
        If QueMostrarDeSalida = 2 Then
            Me.cboTipo.ListIndex = 1
        Else
            Me.cboTipo.ListIndex = 0
        End If
    Else
        If QueMostrarDeSalida = 12 Then
            Me.cboTipo.ListIndex = 1
        Else
            Me.cboTipo.ListIndex = 0
        End If
    End If
    cboTipo.Tag = cboTipo.ListIndex
    
    
    'FechaRoja = DateAdd("d", -vParam.SIIDiasAviso, Now)
    FechaRoja = UltimaFechaCorrectaSII(vParam.SIIDiasAviso, Now)
    
    CreateReportControl
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    R340.Close
    Rwtipopera.Close
    Set R340 = Nothing
    Set Rwtipopera = Nothing
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "416.html"
    End Select
End Sub


Private Sub EstablecerFuente()

    On Error GoTo eEstablecerFuente
    'The following illustrate how to change the different fonts used in the ReportControl
    Dim TextFont As StdFont
    Set TextFont = Me.Font
    TextFont.SIZE = 9
    Set wndReportControl.PaintManager.TextFont = TextFont
    Set wndReportControl.PaintManager.CaptionFont = TextFont
    Set wndReportControl.PaintManager.PreviewTextFont = TextFont
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    'Set fntStrike = wndReportControl.PaintManager.TextFont
    'fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    'Set fntBold = wndReportControl.PaintManager.TextFont
    'fntBold.Bold = True


    Exit Sub
eEstablecerFuente:
    MuestraError Err.Number, Err.Description

End Sub

Private Sub Form_Resize()
Dim J As Integer
    On Error Resume Next
    FrameAux.Move 0, 0, Me.Width, FrameAux.Height
    J = FrameAux.Width - Me.ToolbarAyuda.Width - 320
    If J > 0 Then
        
         
        ToolbarAyuda.Left = J
        J = ToolbarAyuda.Left - Me.Image1.Width - 120
        
        Image1.Left = J
        Image2.Left = Image1.Left - Image2.Width - 30
    End If
    
    wndReportControl.Move 0, FrameAux.Height + FrameAux.top + 60, Me.Width - 240, Me.Height - statusBar.Height - FrameAux.Height - FrameAux.top - 290
    
    
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub






























Public Sub CreateReportControl()

    wndReportControl.Records.DeleteAll
    If Me.cboFacturas.ListIndex = 0 Then
        CreateReportControlPendientes
    Else
        CreateReportControlASWSII
    End If
End Sub


Public Sub CreateReportControlPendientes()
    'Start adding columns
    Dim Column As ReportColumn
    wndReportControl.Columns.DeleteAll
    
    'Chekc
    Set Column = wndReportControl.Columns.Add(0, "", 18, False)
    Column.Icon = COLUMN_CHECK_ICON
     
    'Status factura SII
    Set Column = wndReportControl.Columns.Add(1, "Situacion", 18, False)
    Column.Icon = COLUMN_ATTACHMENT_NORMAL_ICON
    
    
    If Me.cboTipo.ListIndex = 1 Then
        'Set Column = wndReportControl.Columns.Add(2, IIf(vParam.SII_Periodo_DesdeLiq, "F.Liq.", "F.Recep"), 13, True)
        
        
        If vParam.SII_Periodo_DesdeLiq Then
            Msg = "F.Liq."
        Else
            If vParam.SII_ProvDesdeFechaRecepcion Then
                Msg = "F.recep"
            Else
                Msg = "F.Registro"
            End If
        End If
        Set Column = wndReportControl.Columns.Add(2, CStr(Msg), 13, True)
        Set Column = wndReportControl.Columns.Add(3, "N�Factura", 15, True)
        Set Column = wndReportControl.Columns.Add(4, "F. Factura", 13, True)
        Set Column = wndReportControl.Columns.Add(5, "Proveedor", 10, True)
        Set Column = wndReportControl.Columns.Add(6, "Nombre", 30, True)
        Set Column = wndReportControl.Columns.Add(7, "Operacion", 17, True)
        Set Column = wndReportControl.Columns.Add(8, "Tipo", 30, True)
        Set Column = wndReportControl.Columns.Add(9, "Total", 15, True)
        Column.Icon = 12
        Column.Alignment = xtpAlignmentRight
        
    Else
        Set Column = wndReportControl.Columns.Add(2, "N�Factura", 11, True)
        Set Column = wndReportControl.Columns.Add(3, "F.Factura", 10, True)
        Set Column = wndReportControl.Columns.Add(4, "Cliente", 11, True)
        Set Column = wndReportControl.Columns.Add(5, "Nombre", 30, True)
        Set Column = wndReportControl.Columns.Add(6, "Operacion", 20, True)
        Set Column = wndReportControl.Columns.Add(7, "Tipo", 20, True)
        Set Column = wndReportControl.Columns.Add(8, "Total", 15, True)
        Column.Alignment = xtpAlignmentRight
        
        
    End If
     
    

    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    Set fntStrike = wndReportControl.PaintManager.TextFont
    fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    Set fntBold = wndReportControl.PaintManager.TextFont
    fntBold.Bold = True
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
End Sub


Public Sub CreateReportControlASWSII()
    'Start adding columns
    Dim Column As ReportColumn
    wndReportControl.Columns.DeleteAll
    
    
    
    'Status factura SII
    Set Column = wndReportControl.Columns.Add(0, "Situ", 18, False)
    Column.Icon = COLUMN_IMPORTANCE_ICON
    
    Set Column = wndReportControl.Columns.Add(1, "Registrada", 13, True)
    Set Column = wndReportControl.Columns.Add(2, "Enviada", 9, True)
    Set Column = wndReportControl.Columns.Add(3, "Resultado", 13, True)
    
    Set Column = wndReportControl.Columns.Add(4, "N�Factura", 15, True)
    Set Column = wndReportControl.Columns.Add(5, "Fec. factura", 15, True)
    Set Column = wndReportControl.Columns.Add(6, "NIF", 15, True)
    Set Column = wndReportControl.Columns.Add(7, "Nombre", 35, True)
    Set Column = wndReportControl.Columns.Add(8, "Total", 15, True)
    Column.Alignment = xtpAlignmentRight
    

    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    Set fntStrike = wndReportControl.PaintManager.TextFont
    fntStrike.Strikethrough = True
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Bold to the currently set text font
    Set fntBold = wndReportControl.PaintManager.TextFont
    fntBold.Bold = True
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
End Sub









Private Sub Label(visible As Boolean)
    If visible Then
        Label1.Caption = "Leyendo BD"
    Else
        Label1.Caption = ""
    End If
    Label1.Refresh
End Sub


Private Function SQL_ASWSII()


    If Me.cboTipo.ListIndex = 0 Then
        'Emitidas
        SQL_ASWSII = "select IDEnvioFacturasEmitidas id,FechaHoraCreacion fcre,Enviada,Resultado,REG_IDF_NumSerieFacturaEmisor,REG_FE_ImporteTotal impor,REG_IDF_FechaExpedicionFacturaEmisor fecha1,"
        SQL_ASWSII = SQL_ASWSII & " REG_FE_CNT_NombreRazon nombre , REG_FE_CNT_NIF nif,Mensaje"
        SQL_ASWSII = SQL_ASWSII & " From aswsii.envio_facturas_emitidas "
    Else
        'Recibidas
        SQL_ASWSII = "select IDEnvioFacturasRecibidas id,FechaHoraCreacion fcre,Enviada,Resultado,REG_IDF_NumSerieFacturaEmisor,REG_FR_ImporteTotal impor,"
        SQL_ASWSII = SQL_ASWSII & " REG_IDF_FechaExpedicionFacturaEmisor fecha1, REG_FR_CNT_NombreRazon nombre,REG_FR_CNT_NIF nif ,Mensaje "
        SQL_ASWSII = SQL_ASWSII & " FROM aswsii.envio_facturas_recibidas "
        
    End If
    
    SQL_ASWSII = SQL_ASWSII & " WHERE "
    'Multi empresa, vamos a filtrar por NIF
    SQL_ASWSII = SQL_ASWSII & " CAB_Titular_NIF =" & DBSet(vEmpresa.NIF, "T") & " AND "

    SQL_ASWSII = SQL_ASWSII & " enviada = 1 and not csv is null and FechaHoraCreacion>= " & DBSet(DateAdd("m", -1, Now), "F")

    If Text3(0).Text <> "" Then SQL_ASWSII = SQL_ASWSII & " AND date(FechaHoraCreacion) = " & DBSet(Text3(0).Text, "F")
    SQL_ASWSII = SQL_ASWSII & " ORDER BY FechaHoraCreacion desc        "
End Function


'Private Sub CargaRs()
'
'End Sub

Private Function SQL_() As String
Dim SQL As String
Dim Aux As String
    If Me.cboTipo.ListIndex = 0 Then
        'EMITIDAS
        SQL = "select factcli.numserie,factcli.numfactu, factcli.fecfactu,factcli.codmacta,factcli.nommacta"
        'SQL = SQL & " ,wtipopera.denominacion,wconce340.descripcion,factcli.totfaccl"
        SQL = SQL & " , factcli.codopera,factcli.codconce340,factcli.totfaccl"
        SQL = SQL & " ,anofactu,nifdatos,SII_ID ,resultado,csv,enviada,fecliqcl, SII_estado "
        SQL = SQL & " from factcli left join aswsii.envio_facturas_emitidas "
        SQL = SQL & " on factcli.SII_ID = envio_facturas_emitidas.IDEnvioFacturasEmitidas"
       
        Aux = "fecfactu"
        If vParam.SII_Periodo_DesdeLiq Then Aux = "fecliqcl"
        
    Else
        'RECIBIDAS
        SQL = " select factpro.numserie,factpro.numfactu, factpro.fecfactu,factpro.codmacta,factpro.nommacta"
        SQL = SQL & " ,codopera, factpro.codconce340,factpro.totfacpr"
        SQL = SQL & " ,numregis,anofactu,fecharec,SII_ID,nifdatos ,resultado,csv,enviada,fecliqpr,fecregcontable, SII_estado "
        SQL = SQL & " from factpro left join aswsii.envio_facturas_recibidas "
        SQL = SQL & " on factpro.SII_ID = envio_facturas_recibidas.IDEnvioFacturasRecibidas"
       
        
        'Enero 2020
        'a�adimos fechar geistro contable, que ser� la que sube al SII
        'Aux = "fecharec"
        If vParam.SII_ProvDesdeFechaRecepcion Then
            Aux = "fecharec"
        Else
            Aux = "date(fecregcontable)"
        End If
        If vParam.SII_Periodo_DesdeLiq Then Aux = "fecliqpr"
        
    End If
    SQL = SQL & " WHERE " & Aux & " >=" & DBSet(vParam.SIIFechaInicio, "F")
    SQL = SQL & " and " & Aux & " >=" & DBSet(vParam.fechaini, "F")
    SQL = SQL & " AND " & Aux & " <= " & DBSet(Now, "F")
    
    'Enero 2021
    SQL = SQL & " AND SII_estado<9 "   '8-> MODIFICANDO
    SQL = SQL & " and (csv is null or resultado='AceptadoConErrores')"
    
    If Text3(0).Text <> "" Then SQL = SQL & " AND " & Aux & " = " & DBSet(Text3(0).Text, "F")
   
    
    'Enero 2021. Las que estan em proceso de modificacion
    SQL = SQL & " UNION "
    
    If Me.cboTipo.ListIndex = 0 Then
        SQL = SQL & " select numserie,numfactu, fecfactu,codmacta,nommacta , codopera,codconce340,totfaccl"
        SQL = SQL & " ,anofactu,nifdatos,SII_ID ,'' resultado,'' csv,'' enviada,  '' fecliqcl , SII_estado  from factcli"
        
    Else
        SQL = SQL & " select numserie,numfactu, fecfactu,codmacta,nommacta , codopera,codconce340,totfacpr,"
        SQL = SQL & "numregis,anofactu, fecharec ,SII_ID ,nifdatos,'' resultado,'' csv,'' enviada,  '' fecliqpr "
        SQL = SQL & " ,fecregcontable , SII_estado  from factpro "
    End If
    
    SQL = SQL & "  WHERE SII_estado =8 AND "
    SQL = SQL & Aux & " >=" & DBSet(vParam.SIIFechaInicio, "F")
    SQL = SQL & " and " & Aux & " >=" & DBSet(vParam.fechaini, "F")
    If Text3(0).Text <> "" Then SQL = SQL & " AND " & Aux & " = " & DBSet(Text3(0).Text, "F")
   
   
    
    'ORdenacion
    SQL = SQL & " ORDER BY " & Aux & ", numserie, numfactu"
    
    SQL_ = SQL

End Function


'Cuando modifiquemos o insertemos, pondremos el SQL entero
Private Sub CargaDatos()
    Dim Item As ReportRecordItem
    Dim Record
    
    
    
    If Me.cboFacturas.ListIndex = 0 Then
        CargaDatosPendientes
    Else
        CargaDatospendientesASWSII
    End If
    
        
    
End Sub

Private Sub CargaDatosPendientes()
Dim SQL As String

    On Error GoTo ECargaDatos

    Screen.MousePointer = vbHourglass
    statusBar.Panels(1).Text = "Leyendo BD"
    
    Label True
    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
    
    
    
    'Monti NO las quiere marcadas por defecto
    Marcadas = True
    If InStr(1, UCase(vEmpresa.nomempre), "MONTI") > 0 Then Marcadas = False
    
                
    SQL = SQL_
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        If Me.cboTipo.ListIndex = 0 Then
            AddRecordCli
        Else
            AddRecordPro
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    wndReportControl.Populate
    
ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, SQL
    
    
    
    statusBar.Panels(1).Text = ""
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub




'socio, pendiente , nombre, matricula,licencia
'Leera los datos de mirsaux
Private Sub AddRecordCli()
Dim SQL As String
Dim Record As ReportRecord
Dim PuedeEnviar As Boolean
Dim ItemToolTip As String
Dim ItemIcon As Integer
Dim Color As Long
    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    PuedeEnviar = True
    If IsNull(miRsAux!enviada) Then
        'PENDIENTE DE grabar en ASWII
        i = 0
        ItemIcon = -1
        ItemToolTip = "Pendiente"
        
    Else
        If miRsAux!enviada = 0 Then
            'OK Falta enviar desdeel ASWSI a la AEAT
            i = 1
            PuedeEnviar = False
            ItemIcon = RECORD_REPLIED_ICON
            ItemToolTip = "En progreso"
      
        Else
        
            'ENERO
            If miRsAux!SII_estado = 8 Then
                i = 3
                ItemIcon = RECORD_READ_MAIL_ICON
                ItemToolTip = "Mofificando factura"
                PuedeEnviar = False

            
            
            Else
        
                'ENVIADA. Estara incorrecta
                If DBLet(miRsAux!Resultado, "T") = "AceptadoConErrores" Then
                    i = 3
                    ItemIcon = RECORD_IMPORTANCE_HIGH_ICON
                    ItemToolTip = "Aceptada con errores"
                    PuedeEnviar = False
    
                Else
                    i = 2
                    ItemIcon = RECORD_UNREAD_MAIL_ICON
                    ItemToolTip = "Errores"
                    
                End If
            End If  'EDITANDO
        End If
    End If
    
        'Check
    Set Item = Record.AddItem("")
    If PuedeEnviar Then
        Item.Checked = MarcarComoCheck And Marcadas
        Item.HasCheckbox = i = 0 Or i = 2
    Else
        Item.Checked = False
        Item.HasCheckbox = False
    End If
    Item.TristateCheckbox = False
    
    
    'Estado
    Set Item = Record.AddItem("")
    Item.Icon = ItemIcon
    Item.ToolTip = ItemToolTip
    Item.SortPriority = i
    Item.Caption = ""
    
    
    
    
    SQL = miRsAux!NUmSerie & Format(miRsAux!numfactu, "0000000")
    Set Item = Record.AddItem(SQL)
    
    Set Item = Record.AddItem(Format(miRsAux!FecFactu, "yyyymmdd"))
    Item.Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
    
    
    Set Item = Record.AddItem(CStr(miRsAux!codmacta))
    Item.Value = CStr(miRsAux!codmacta)
    Set Item = Record.AddItem(DBLet(miRsAux!Nommacta, "T"))
    
    If DBLet(miRsAux!nifdatos, "T") = "" Then
        Item.BackColor = vbRed
    End If
    If Val(DBLet(miRsAux!CodOpera, "N")) <> Val(Rwtipopera!Codigo) Then
        Rwtipopera.Find "codigo = " & DBLet(miRsAux!CodOpera, "N"), , adSearchForward, 1
        If Rwtipopera.EOF Then
            SQL = "N/D"
            Rwtipopera.MoveFirst
        Else
            SQL = DBLet(Rwtipopera!denominacion, "T")
        End If
    Else
        SQL = DBLet(Rwtipopera!denominacion, "T")
    End If
    Record.AddItem CStr(SQL)
    
    If DBLet(miRsAux!codconce340, "T") <> R340!Codigo Then
        R340.Find "codigo = " & DBSet(DBLet(miRsAux!codconce340, "T"), "T"), , adSearchForward, 1
        If R340.EOF Then
            SQL = "N/D"
            R340.MoveFirst
        Else
            SQL = DBLet(R340!Descripcion, "T")
        End If
    Else
        SQL = DBLet(R340!Descripcion, "T")
    End If
    Record.AddItem CStr(SQL)
    
    Set Item = Record.AddItem(miRsAux!totfaccl * 100)
    Item.Caption = Format(miRsAux!totfaccl, FormatoImporte)
    
    Record.Tag = miRsAux!NUmSerie & "|" & miRsAux!numfactu & "|" & miRsAux!Anofactu & "|" & miRsAux!SII_ID & "|"
                
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = "ID: " & miRsAux!Nommacta
    
    If PuedeEnviar Then
        Color = -1
        If miRsAux!FecFactu < FechaRoja Then
            Color = vbRed
        Else
            If miRsAux!FecFactu = FechaRoja Then Color = vbBlue
        End If
        If Color <> -1 Then
            For J = 2 To Record.ItemCount - 1
                Record.Item(J).ForeColor = Color
            Next
        End If
    End If
    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
End Sub


Private Sub AddRecordPro()
Dim SQL As String
Dim Record As ReportRecord
Dim PuedeEnviar As Boolean
Dim ItemToolTip As String
Dim ItemIcon As Integer

Dim Color As Long
    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    
    
    PuedeEnviar = True
    If IsNull(miRsAux!enviada) Then
        'PENDIENTE DE grabar en ASWII
        i = 0
        ItemIcon = -1
        ItemToolTip = "Pendiente"
        
    Else
        If miRsAux!enviada = 0 Then
            'OK Falta enviar desdeel ASWSI a la AEAT
            i = 1
            PuedeEnviar = False
            ItemIcon = RECORD_REPLIED_ICON
            ItemToolTip = "En progreso"
      
        Else
        
                    'ENERO
            If miRsAux!SII_estado = 8 Then
                i = 3
                ItemIcon = RECORD_READ_MAIL_ICON
                ItemToolTip = "Mofificando factura"
                PuedeEnviar = False

            
            
            Else
            
                'ENVIADA. Estara incorrecta
                If DBLet(miRsAux!Resultado, "T") = "AceptadoConErrores" Then
                    i = 3
                    ItemIcon = RECORD_IMPORTANCE_HIGH_ICON
                    ItemToolTip = "Aceptada con errores"
                    PuedeEnviar = False
                Else
                    i = 2
                    ItemIcon = RECORD_UNREAD_MAIL_ICON
                    ItemToolTip = "Errores"
                End If
            End If
        End If
    End If
    
        'Check
    Set Item = Record.AddItem("")
    If PuedeEnviar Then
        Item.Checked = MarcarComoCheck And Marcadas
        Item.HasCheckbox = i = 0 Or i = 2
    Else
        Item.Checked = False
        Item.HasCheckbox = False
    End If
    Item.TristateCheckbox = False
    
    
    'Estado
    Set Item = Record.AddItem("")
    Item.Icon = ItemIcon
    Item.ToolTip = ItemToolTip
    Item.SortPriority = i
    Item.Caption = ""
    
        
    
    If vParam.SII_Periodo_DesdeLiq Then
        Set Item = Record.AddItem(Format(miRsAux!fecliqpr, "yyyymmdd"))
        Item.Caption = Format(miRsAux!fecliqpr, "dd/mm/yyyy")
    Else
        If vParam.SII_ProvDesdeFechaRecepcion Then
            Set Item = Record.AddItem(Format(miRsAux!fecharec, "yyyymmdd"))
            Item.Caption = Format(miRsAux!fecharec, "dd/mm/yyyy")
        Else
            Set Item = Record.AddItem(Format(miRsAux!fecregcontable, "yyyymmdd"))
            Item.Caption = Format(miRsAux!fecregcontable, "dd/mm/yyyy")
        End If
    End If
    
    If miRsAux!NUmSerie = "1" Then
        SQL = ""
    Else
        SQL = miRsAux!NUmSerie & " "
    End If
    SQL = SQL & miRsAux!numfactu
    Set Item = Record.AddItem(SQL)
    
    Set Item = Record.AddItem(Format(miRsAux!FecFactu, "yyyymmdd"))
    Item.Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
    
    
    Set Item = Record.AddItem(CStr(miRsAux!codmacta))
    Item.Value = CStr(miRsAux!codmacta)
    Set Item = Record.AddItem(DBLet(miRsAux!Nommacta, "T"))
    
    If Val(DBLet(miRsAux!CodOpera, "N")) <> Val(Rwtipopera!Codigo) Then
        Rwtipopera.Find "codigo = " & DBLet(miRsAux!CodOpera, "N"), , adSearchForward, 1
        If Rwtipopera.EOF Then
            SQL = "N/D"
            Rwtipopera.MoveFirst
        Else
            SQL = DBLet(Rwtipopera!denominacion, "T")
        End If
    Else
        SQL = DBLet(Rwtipopera!denominacion, "T")
    End If
    Record.AddItem CStr(SQL)
    
    If DBLet(miRsAux!codconce340, "T") <> R340!Codigo Then
        R340.Find "codigo = " & DBSet(DBLet(miRsAux!codconce340, "T"), "T"), , adSearchForward, 1
        If R340.EOF Then
            SQL = "N/D"
            R340.MoveFirst
        Else
            SQL = DBLet(R340!Descripcion, "T")
        End If
    Else
        SQL = DBLet(R340!Descripcion, "T")
    End If
    Record.AddItem CStr(SQL)
    
    
    
    
    
    Set Item = Record.AddItem(miRsAux!totfacpr * 100)
    Item.Caption = Format(miRsAux!totfacpr, FormatoImporte)
    
    Record.Tag = miRsAux!NUmSerie & "|" & miRsAux!Numregis & "|" & miRsAux!Anofactu & "|" & miRsAux!SII_ID & "|"
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = "ID: " & miRsAux!Nommacta
    
    Color = -1
    If PuedeEnviar Then
        If vParam.SII_Periodo_DesdeLiq Then
            If miRsAux!fecliqpr < FechaRoja Then
                Color = vbRed
            Else
                If miRsAux!fecliqpr = FechaRoja Then Color = vbBlue
            End If
        Else
            'Enero 2020
            If vParam.SII_ProvDesdeFechaRecepcion Then
                If miRsAux!fecharec < FechaRoja Then
                    Color = vbRed
                Else
                    If miRsAux!fecharec = FechaRoja Then Color = vbBlue
                End If
            Else
                If CDate(miRsAux!fecregcontable) < FechaRoja Then
                    Color = vbRed
                Else
                    If CDate(miRsAux!fecregcontable) = FechaRoja Then Color = vbBlue
                End If
            End If
        End If
    End If
    If Color <> -1 Then
        For J = 2 To Record.ItemCount - 1
            Record.Item(J).ForeColor = Color
        Next
    End If
    
    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
End Sub


Private Sub AddRecordASWSII()
Dim SQL As String
Dim Record As ReportRecord
    
    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    Set Item = Record.AddItem("")
    SQL = DBLet(miRsAux!Resultado, "T")
    Item.SortPriority = IIf(SQL <> "", 1, 0)
    If SQL = "Correcto" Then SQL = ""
    
    If SQL = "" Then
        'Pendiente
        Item.Icon = -1
        Item.ToolTip = "Pendiente"

    Else
        'incorrect
        Item.Icon = RECORD_IMPORTANCE_HIGH_ICON
        Item.ToolTip = "Incorrecta"
    End If
    
    
    Set Item = Record.AddItem(Format(miRsAux!fcre, "yyyymmddmmnn"))
    Item.Caption = Format(miRsAux!fcre, "dd/mm/yyyy")
    
    
    Set Item = Record.AddItem(IIf(miRsAux!enviada = 0, "NO", ""))
    
    
    Set Item = Record.AddItem(Mid(SQL, 1, 12)) 'resultado
    Item.ToolTip = DBLet(miRsAux!Mensaje, "T")
    
    Set Item = Record.AddItem(DBLet(miRsAux!REG_IDF_NumSerieFacturaEmisor, "T"))
    Set Item = Record.AddItem(Format(miRsAux!fecha1, "yyyymmdd"))
    Item.Caption = Format(miRsAux!fecha1, "dd/mm/yyyy")
    
    
    
    Record.AddItem DBLet(miRsAux!NIF, "T")
    Record.AddItem CStr(DBLet(miRsAux!Nombre, "T"))
    
    Set Item = Record.AddItem(miRsAux!Impor * 100)
    Item.Caption = Format(miRsAux!Impor, FormatoImporte)
    
    Record.Tag = miRsAux!Id 'CLAVE
                
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = "ID: " & miRsAux!Nombre
    
    
    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
End Sub




Private Sub frmC_Selec(vFecha As Date)
    Msg = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    
    For i = 0 To Me.wndReportControl.Records.Count - 1
           
        wndReportControl.Records(i).Item(0).Checked = Index = 1
    Next
     Screen.MousePointer = vbHourglass
    wndReportControl.Populate
    Screen.MousePointer = vbDefault
End Sub











Private Sub imgFecha_Click(Index As Integer)
    Msg = Format(Now, "dd/mm/yyyy")
    If IsDate(Text3(Index).Text) Then Msg = Format(Text3(Index).Text, "dd/mm/yyyy")
    Set frmC = New frmCal
    frmC.Fecha = CDate(Msg)
    Msg = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If Msg <> "" Then
        Text3(Index).Text = Msg
        Text3_LostFocus Index
    End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    If Text3(Index).Text = Text3(Index).Tag Then Exit Sub
    
    If Not EsFechaOK(Text3(Index)) Then Text3(Index).Text = ""
    
    Text3(Index).Tag = Text3(Index).Text
    CargaDatos
    
End Sub

Private Sub wndReportControl_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = 32 Then
        'SAPACE BAR
        
        If wndReportControl.SelectedRows.Count > 0 Then
            If wndReportControl.SelectedRows(0).Record(0).HasCheckbox Then
                wndReportControl.SelectedRows(0).Record(0).Checked = Not wndReportControl.SelectedRows(0).Record(0).Checked
                wndReportControl.SelectedRows(0).Selected = True
                wndReportControl.SetFocus
            End If
        Else
            
            If wndReportControl.FocusedRow Is Nothing Then
            
            Else
                If wndReportControl.FocusedRow.Record(0).HasCheckbox Then wndReportControl.FocusedRow.Record(0).Checked = Not wndReportControl.FocusedRow.Record(0).Checked
            End If
            
           
        End If
    End If
End Sub

Private Sub wndReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
 Dim SQL As String
    If Row Is Nothing Then Exit Sub
    
    If cboFacturas.ListIndex = 0 Then
        If Row.Record.Item(1).Icon = RECORD_UNREAD_MAIL_ICON Or Row.Record.Item(1).Icon = RECORD_IMPORTANCE_HIGH_ICON Then
        

            If Me.cboTipo.ListIndex = 0 Then
                SQL = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_emitidas", "IDEnvioFacturasEmitidas", RecuperaValor(Row.Record.Tag, 4))
            Else
                SQL = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_recibidas", "IDEnvioFacturasRecibidas", RecuperaValor(Row.Record.Tag, 4))
            End If
            If Row.Record.Item(1).Icon = RECORD_UNREAD_MAIL_ICON Then
                MsgBox "Factura incorrecta: " & vbCrLf & SQL, vbExclamation
            Else
                MsgBox "ACEPTADA con errores: " & vbCrLf & SQL, vbExclamation
            End If
       End If
   Else
        'Facturas ya subidas al ASWII
        SQL = ""
        If Me.cboFiltro.ListIndex = 1 Then
        '    SQL = "S"
        Else
            If Row.Record.Item(0).Icon = RECORD_IMPORTANCE_HIGH_ICON Then SQL = "s"
        End If
        If SQL <> "" Then
            If Me.cboTipo.ListIndex = 0 Then
                SQL = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_emitidas", "IDEnvioFacturasEmitidas", Row.Record.Tag)
            Else
                SQL = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_recibidas", "IDEnvioFacturasRecibidas", Row.Record.Tag)
            End If
            MsgBox "ASWSII. Factura incorrecta: " & vbCrLf & SQL, vbExclamation
        End If
   End If

    
End Sub






Private Sub HacerGrabacionTablas_ASWISII()
Dim SQL As String
Dim ID_ASWSII As Long
Dim CuantasLlevo As Integer
Dim B As Boolean


    'Insertamos en tmpfaclin
    Label1.Caption = "Prepara datos"
    Label1.Refresh
    Conn.Execute "DELETE FROM tmpfaclin WHERE codusu = " & vUsu.Codigo
    
    SQL = ""
    For i = 0 To Me.wndReportControl.Records.Count - 1
        'codusu , Codigo, NUmSerie, NumFac, Fecha, Codigo
        'tmpfaclin
        If wndReportControl.Records(i).Item(0).HasCheckbox Then
        
            If wndReportControl.Records(i).Item(0).Checked Then
                Msg = wndReportControl.Records(i).Tag
                SQL = SQL & ", (" & vUsu.Codigo & "," & DBSet(RecuperaValor(Msg, 1), "T")
                SQL = SQL & "," & RecuperaValor(Msg, 2) & ",'" & RecuperaValor(Msg, 3) & "-01-01',0)"
                
                
                If Len(SQL) > 600 Then
                    SQL = Mid(SQL, 2)
                    SQL = "INSERT INTO tmpfaclin(codusu ,  NUmSerie, NumFac, Fecha, Codigo) VALUES " & SQL
                    Conn.Execute SQL
                    SQL = ""
                End If
            End If
        End If
    Next
    
    If SQL <> "" Then
        SQL = Mid(SQL, 2)
        SQL = "INSERT INTO tmpfaclin(codusu ,  NUmSerie, NumFac, Fecha, Codigo) VALUES " & SQL
        Conn.Execute SQL
    End If
    
    
    
    'Ya tenemos las facturas que vamos a pasar
    Label1.Caption = "Inserta registros"
    Label1.Refresh
    
        
     
   

    
    SQL = "Select * from tmpfaclin where codusu = " & vUsu.Codigo & " ORDER BY numserie,numfac"
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic
    While Not miRsAux.EOF
    
        'Obtengo contador
         If Me.cboTipo.ListIndex = 0 Then
            SQL = DevuelveDesdeBD("max(IDEnvioFacturasEmitidas)", "aswsii.envio_facturas_emitidas", "1", "1")
        Else
            SQL = DevuelveDesdeBD("max(IDEnvioFacturasRecibidas)", "aswsii.envio_facturas_recibidas", "1", "1")
        End If
        If SQL = "" Then SQL = "0"
        
        ID_ASWSII = Val(SQL) + 1
    
    
    
        If Me.cboTipo.ListIndex = 0 Then
    
            Label1.Caption = "Fac:" & miRsAux!NUmSerie & miRsAux!NumFac
            Label1.Refresh
            B = Sii_FraCLI(miRsAux!NUmSerie, miRsAux!NumFac, Year(miRsAux!Fecha), ID_ASWSII, SQL, False)
            
        Else
            Label1.Caption = "Reg: " & miRsAux!NUmSerie & " :  " & miRsAux!NumFac
            Label1.Refresh
            B = Sii_FraPRO(miRsAux!NUmSerie, miRsAux!NumFac, Year(miRsAux!Fecha), ID_ASWSII, SQL, False)
            
        End If
        
        If B Then
            
            If Ejecuta(SQL) Then
                
                If Me.cboTipo.ListIndex = 0 Then
                    'UPDATEAMOS EN cabfact
                    SQL = "UPDATE factcli SET SII_ID =" & ID_ASWSII
                    SQL = SQL & " , SII_estado = 1"
                    SQL = SQL & " where numserie =" & DBSet(miRsAux!NUmSerie, "T") & " AND numfactu =" & miRsAux!NumFac & " AND anofactu =" & Year(miRsAux!Fecha)
                    
                Else
                    SQL = "UPDATE factpro SET SII_ID =" & ID_ASWSII & " ,fecregcontable=fecregcontable "
                    SQL = SQL & " , SII_estado = 1"
                    SQL = SQL & " where numserie =" & DBSet(miRsAux!NUmSerie, "T") & " AND numregis =" & miRsAux!NumFac & " AND anofactu =" & Year(miRsAux!Fecha)
                End If
                
                
                Conn.Execute SQL

               
                
                
                
                CuantasLlevo = CuantasLlevo + 1
                If CuantasLlevo > 40 Then
                    Screen.MousePointer = vbHourglass
                    DoEvent2
                    CuantasLlevo = 0
                    Label1.Caption = "Leyendo BD"
                    Label1.Refresh
                    espera 0.25
                    
      
                    
                    
                    
                    
                End If
                
            End If
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    DoEvent2
    Screen.MousePointer = vbHourglass
    Label1.Caption = "Leyendo BD"
    Label1.Refresh
    espera 0.5
    

End Sub


Private Sub CargaCombos()
Dim FiltroTipo As Integer
    
    Me.cboFiltro.Clear
    If cboTipo.ListCount = 0 Then
        cboTipo.AddItem "Emitidas"
        cboTipo.AddItem "Recibidas"
        cboTipo.ListIndex = 0
        cboTipo.Tag = 0
    End If
    
    
    'Filtro
    If Me.cboFacturas.ListIndex = 0 Then
        'Para pendientes
        Me.cboFiltro.AddItem "Todas"
        cboFiltro.AddItem "Erroneas"
        cboFiltro.AddItem "Pendientes"
        
    Else
        'Para comunicadas
'        cboFiltro.AddItem "Pendiente"
'        cboFiltro.AddItem "Erroneas"
        cboFiltro.AddItem "Correctas ultimo mes"
    End If
    cboFiltro.ListIndex = 0
    cboFiltro.Tag = 0
End Sub



Private Sub CargaDatospendientesASWSII()
Dim SQL As String
    On Error GoTo eCargaDatospendientesASWSII
    Screen.MousePointer = vbHourglass
    statusBar.Panels(1).Text = "Leyendo BD"
    
    Label True
    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
               
    SQL = SQL_ASWSII
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        AddRecordASWSII
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    wndReportControl.Populate
    
    

    
eCargaDatospendientesASWSII:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, SQL
    
    statusBar.Panels(1).Text = ""
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub


Private Sub ExportarCSV()
Dim SQL As String
    
    On Error GoTo eExportarCsv
    
    SQL = SQL_
        
    'LLamos a la funcion
    Screen.MousePointer = vbHourglass
    If GeneraFicheroCSV(SQL, App.Path & "\DatosPdeSII.csv", True) Then
        LanzaVisorMimeDocumento Me.hwnd, App.Path & "\DatosPdeSII.csv"
    End If

    

eExportarCsv:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Screen.MousePointer = vbDefault
End Sub



