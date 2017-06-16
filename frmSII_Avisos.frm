VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmSII_Avisos 
   Caption         =   "Comunicación datos SII"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15630
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
   ScaleHeight     =   9555
   ScaleWidth      =   15630
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   6735
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   14295
      _Version        =   1114114
      _ExtentX        =   25215
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
      Width           =   14295
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
         ItemData        =   "frmSII_Avisos.frx":0000
         Left            =   4440
         List            =   "frmSII_Avisos.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2250
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
         ItemData        =   "frmSII_Avisos.frx":0004
         Left            =   2400
         List            =   "frmSII_Avisos.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1890
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
         ItemData        =   "frmSII_Avisos.frx":0008
         Left            =   120
         List            =   "frmSII_Avisos.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   2130
      End
      Begin VB.CommandButton cmdComunicar 
         Height          =   495
         Left            =   9240
         Picture         =   "frmSII_Avisos.frx":002F
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Comunicar facturas"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdModificar 
         Height          =   495
         Left            =   9840
         Picture         =   "frmSII_Avisos.frx":6881
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   495
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
         Picture         =   "frmSII_Avisos.frx":7283
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
         Left            =   13440
         Picture         =   "frmSII_Avisos.frx":730E
         Stretch         =   -1  'True
         Top             =   210
         Width           =   540
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   12720
         Picture         =   "frmSII_Avisos.frx":7E67
         Stretch         =   -1  'True
         Top             =   180
         Width           =   615
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   8400
         Picture         =   "frmSII_Avisos.frx":8A52
         ToolTipText     =   "Seleccionar todo"
         Top             =   480
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   8760
         Picture         =   "frmSII_Avisos.frx":8B9C
         ToolTipText     =   "Quitar seleccion"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "aa"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   10440
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   9180
      Width           =   15630
      _ExtentX        =   27570
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



'Dim iconArray(0 To 9) As Long
'Dim RowExpanded(0 To 49) As Boolean
'Dim RowVisible(0 To 49) As Boolean
'Dim MaxRowIndex As Long
Dim fntBold As StdFont
Dim fntStrike As StdFont










Private Sub VerTodos()

    CargaDatos

End Sub
















Private Sub cboFacturas_Click()
    If PrimeraVez Then Exit Sub
    If cboFacturas.Tag = cboFacturas.ListIndex Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    CargaCombos
    
    
    ComprobarResultadoEnviadasASWII IIf(Me.cboTipo.ListIndex = 1, 2, 1), Me.Label1
    
    PrimeraVez = False
    
    
    CreateReportControl
    CargaDatos
    wndReportControl.Populate
    
    imgCheck(0).Visible = cboFacturas.ListIndex = 0
    imgCheck(1).Visible = cboFacturas.ListIndex = 0
    Me.cmdModificar.Visible = cboFacturas.ListIndex = 0
    Me.cmdComunicar.Visible = cboFacturas.ListIndex = 0
    
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
        If wndReportControl.Records(i).Item(0).Checked Then J = J + 1
    Next

    If J = 0 Then
        MsgBox "Seleccione alguna factura", vbExclamation
        Exit Sub
    End If
    
    Msg = ""
    If J > 1 Then Msg = "s"
    Msg = "Va a traspasar " & J & " factura" & Msg & " al programa de comunicación electronica ASWSII(Ariadna Software). " & vbCrLf & "¿Continuar?"
    
    If MsgBox(Msg, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub


    If BloqueoManual(True, "SII", "1") Then
        Screen.MousePointer = vbHourglass
        Label True
        Set miRsAux = New ADODB.Recordset
        HacerGrabacionTablas_ASWISII
        Set miRsAux = Nothing
        Label False
        BloqueoManual False, "SII", ""
        
        
        'volvemos a cargar
        CargaDatos
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Proceso bloqueado por otro usuario", vbExclamation
    End If


End Sub

Private Sub cmdModificar_Click()
Dim Sql As String
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
    
    
    Sql = SQL_
    'Solo esta factura
    i = InStr(1, Sql, " ORDER BY ")
    Msg = Mid(Sql, i)
    Sql = Mid(Sql, 1, i - 1)
    Sql = Sql & " AND numserie = " & DBSet(RecuperaValor(Me.wndReportControl.SelectedRows(0).Record.Tag, 1), "T")
    Sql = Sql & " AND " & IIf(cboTipo.ListIndex = 0, "numfactu", "numregis") & " = "
    Sql = Sql & RecuperaValor(Me.wndReportControl.SelectedRows(0).Record.Tag, 2)
    Sql = Sql & " AND  anofactu = " & RecuperaValor(Me.wndReportControl.SelectedRows(0).Record.Tag, 3)
    Sql = Sql & Msg
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        'Ha cambiado algo "importante. Volvemos a cargar todo
        Todo = True
    Else
        Todo = False
        If cboTipo.ListIndex = 0 Then
            'CLiente
            wndReportControl.SelectedRows(0).Record(3).Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            wndReportControl.SelectedRows(0).Record(3).Value = Format(miRsAux!FecFactu, "yyyymmdd")
            wndReportControl.SelectedRows(0).Record(4).Caption = miRsAux!codmacta
            wndReportControl.SelectedRows(0).Record(5).Caption = miRsAux!Nommacta
            wndReportControl.SelectedRows(0).Record(6).Caption = DBLet(miRsAux!denominacion, "")
            wndReportControl.SelectedRows(0).Record(7).Caption = DBLet(miRsAux!Descripcion, "")
            
            wndReportControl.SelectedRows(0).Record(8).Caption = Format(miRsAux!totfaccl, "")
            wndReportControl.SelectedRows(0).Record(8).Value = miRsAux!totfaccl * 100
            
            
        Else
            'Proveedor
            wndReportControl.SelectedRows(0).Record(2).Caption = Format(miRsAux!fecharec, "dd/mm/yyyy")
            wndReportControl.SelectedRows(0).Record(2).Value = Format(miRsAux!fecharec, "yyyymmdd")
            wndReportControl.SelectedRows(0).Record(3).Value = miRsAux!NumFactu
            wndReportControl.SelectedRows(0).Record(4).Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
            wndReportControl.SelectedRows(0).Record(4).Value = Format(miRsAux!FecFactu, "yyyymmdd")
            wndReportControl.SelectedRows(0).Record(5).Caption = miRsAux!codmacta
            wndReportControl.SelectedRows(0).Record(6).Caption = miRsAux!Nommacta
            wndReportControl.SelectedRows(0).Record(7).Caption = DBLet(miRsAux!denominacion, "")
            wndReportControl.SelectedRows(0).Record(8).Caption = DBLet(miRsAux!Descripcion, "")
            
            wndReportControl.SelectedRows(0).Record(9).Caption = Format(miRsAux!totfacpr, "")
            wndReportControl.SelectedRows(0).Record(9).Value = miRsAux!totfacpr * 100
    
        End If
        wndReportControl.Populate
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If Todo Then CargaDatos
    
End Sub

Private Sub Form_Activate()
    Dim Record As ReportRecord

    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        DoEvents
        CargaDatos
        'CargaDatosPrueba "", False
     
    End If
End Sub





Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon


    
    wndReportControl.Icons = ReportControlGlobalSettings.Icons
    wndReportControl.PaintManager.NoItemsText = "Ningún registro "
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
    
    
    
    CreateReportControl
    
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
    J = FrameAux.Width - Me.Image1.Width - 320
    If J > 0 Then
        Image1.Left = J
        Image2.Left = Image1.Left - Image2.Width - 30
    End If
    
    wndReportControl.Move 0, FrameAux.Height + FrameAux.top + 60, Me.Width - 120, Me.Height - statusBar.Height - FrameAux.Height - FrameAux.top - 290
    
    
    
    
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
        
        Set Column = wndReportControl.Columns.Add(2, "F.Recep", 13, True)
        Set Column = wndReportControl.Columns.Add(3, "NºFactura", 15, True)
        Set Column = wndReportControl.Columns.Add(4, "F. Factura", 13, True)
        Set Column = wndReportControl.Columns.Add(5, "Proveedor", 10, True)
        Set Column = wndReportControl.Columns.Add(6, "Nombre", 30, True)
        Set Column = wndReportControl.Columns.Add(7, "Operacion", 17, True)
        Set Column = wndReportControl.Columns.Add(8, "Tipo", 30, True)
        Set Column = wndReportControl.Columns.Add(9, "Total", 15, True)
        Column.Icon = 12
        Column.Alignment = xtpAlignmentRight
        
    Else
        Set Column = wndReportControl.Columns.Add(2, "NºFactura", 11, True)
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
    
    Set Column = wndReportControl.Columns.Add(4, "NºFactura", 15, True)
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









Private Sub Label(Visible As Boolean)
    If Visible Then
        Label1.Caption = "Leyendo registros BD"
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
    If cboFiltro.ListIndex = 0 Then
        SQL_ASWSII = SQL_ASWSII & " enviada = 0 "
        'Y las que esten pendientes de
        SQL_ASWSII = SQL_ASWSII & " OR substring(resultado,1,1)='A'"
    Else
        If cboFiltro.ListIndex = 1 Then
            SQL_ASWSII = SQL_ASWSII & " enviada = 1 and resultado = 'Incorrecto'"
            SQL_ASWSII = SQL_ASWSII & " AND  FechaHoraCreacion>= '" & Format(DateAdd("m", -1, Now), "yyyy-mm") & "-01'"
            
        Else
            SQL_ASWSII = SQL_ASWSII & " enviada = 1 and resultado='Correcto' and FechaHoraCreacion>= " & DBSet(DateAdd("m", -1, Now), "F")
        End If
    End If
    If Text3(0).Text <> "" Then SQL_ASWSII = SQL_ASWSII & " AND date(FechaHoraCreacion) = " & DBSet(Text3(0).Text, "F")
        
End Function




Private Function SQL_() As String
Dim Sql As String
Dim Aux As String
    If Me.cboTipo.ListIndex = 0 Then
        'EMITIDAS
        Sql = "select factcli.numserie,factcli.numfactu, factcli.fecfactu,factcli.codmacta,factcli.nommacta"
        Sql = Sql & " ,wtipopera.denominacion,wconce340.descripcion,factcli.totfaccl"
        Sql = Sql & " ,anofactu,SII_ID ,nifdatos,SII_status "
        Sql = Sql & " from factcli left join usuarios.wtipopera  on factcli.codopera = wtipopera.codigo"
        Sql = Sql & " left join usuarios.wconce340  on factcli.codconce340 = wconce340.codigo"

        Aux = "fecfactu"

        
        
    Else
        'RECIBIDAS
        Sql = " select factpro.numserie,factpro.numfactu, factpro.fecfactu,factpro.codmacta,factpro.nommacta"
        Sql = Sql & " ,wtipopera.denominacion,wconce340.descripcion,factpro.totfacpr"
        Sql = Sql & " ,numregis,anofactu,fecharec,SII_ID,nifdatos, SII_status "
        Sql = Sql & " from factpro left join usuarios.wtipopera  on factpro.codopera = wtipopera.codigo"
        Sql = Sql & " left join usuarios.wconce340  on factpro.codconce340 = wconce340.codigo"
        Aux = "fecharec"
        
    End If
    Sql = Sql & " WHERE " & Aux & " >=" & DBSet(vParam.SIIFechaInicio, "F")
    Sql = Sql & " AND " & Aux & " <= " & DBSet(DateAdd("d", vParam.SIIDiasAviso, Now), "F")
    'que no esten comunicadas SII_ID  SII_status=1"
    If cboFiltro.ListIndex = 0 Then
        'TODAS
        Sql = Sql & " AND (coalesce(SII_ID,0)=0 OR (SII_ID >0 and SII_status IN (1,2) ))"
    ElseIf cboFiltro.ListIndex = 1 Then
        Sql = Sql & " AND (SII_ID >0 and SII_status IN (1,2) )"
    Else
        Sql = Sql & " AND coalesce(SII_ID,0)=0 "
    End If
    
    
    
    If Text3(0).Text <> "" Then Sql = Sql & " AND " & Aux & " = " & DBSet(Text3(0).Text, "F")
   
    
    
    'ORdenacion
    Sql = Sql & " ORDER BY " & Aux & ", numserie, numfactu"
    
    
    SQL_ = Sql

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
Dim Sql As String

    On Error GoTo ECargaDatos

    Screen.MousePointer = vbHourglass
    statusBar.Panels(1).Text = "Leyendo BD"
    
    Label True
    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
                
    
                
    Sql = SQL_
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, Sql
    
    
    
    statusBar.Panels(1).Text = ""
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub




'socio, pendiente , nombre, matricula,licencia
'Leera los datos de mirsaux
Private Sub AddRecordCli()
Dim Sql As String
Dim Record As ReportRecord
    
    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    
    'Check
    Set Item = Record.AddItem("")
    Item.Checked = True
    Item.HasCheckbox = True
    Item.TristateCheckbox = False
    
    'Estado
    Set Item = Record.AddItem("")
    Item.SortPriority = IIf(DBLet(miRsAux!sii_status, "N") = 0, 1, 0)
    i = DBLet(miRsAux!sii_status, "N")
    If i = 0 Then
        'Pendiente
        Item.Icon = -1
        Item.ToolTip = "Pendiente"
        
    ElseIf i < 3 Then
        'incorrect
        Item.Icon = RECORD_UNREAD_MAIL_ICON
        Item.ToolTip = "Incorrecta"
        
    Else
        'Correcta
        Item.Icon = COLUMN_ATTACHMENT_ICON
        Item.ToolTip = "OK"
        
    End If
    Item.Caption = ""
    
    
    
    
    Sql = miRsAux!NUmSerie & Format(miRsAux!NumFactu, "0000000")
    Set Item = Record.AddItem(Sql)
    
    Set Item = Record.AddItem(Format(miRsAux!FecFactu, "yyyymmdd"))
    Item.Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
    
    
    Set Item = Record.AddItem(CStr(miRsAux!codmacta))
    Item.Value = CStr(miRsAux!codmacta)
    Set Item = Record.AddItem(DBLet(miRsAux!Nommacta, "T"))
    
    If DBLet(miRsAux!nifdatos, "T") = "" Then Item.BackColor = vbRed
    
    Record.AddItem DBLet(miRsAux!denominacion, "T")
    Record.AddItem CStr(DBLet(miRsAux!Descripcion, "T"))
    
    Set Item = Record.AddItem(miRsAux!totfaccl * 100)
    Item.Caption = Format(miRsAux!totfaccl, FormatoImporte)
    
    Record.Tag = miRsAux!NUmSerie & "|" & miRsAux!NumFactu & "|" & miRsAux!anofactu & "|" & miRsAux!SII_ID & "|"
                
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = "ID: " & miRsAux!Nommacta
    
    
    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
End Sub


Private Sub AddRecordPro()
Dim Sql As String
Dim Record As ReportRecord
    
    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
        
    'Check
    Set Item = Record.AddItem("")
    Item.Checked = True
    Item.HasCheckbox = True
    Item.TristateCheckbox = False
    
    'Estado
    'Estado
    Set Item = Record.AddItem("")
    Item.SortPriority = IIf(DBLet(miRsAux!sii_status, "N") = 0, 1, 0)
    i = DBLet(miRsAux!sii_status, "N")
    If i = 0 Then
        'Pendiente
        Item.Icon = -1
        Item.ToolTip = "Pendiente"
        
    ElseIf i < 3 Then
        'incorrect
        Item.Icon = RECORD_UNREAD_MAIL_ICON
        Item.ToolTip = "Incorrecta"
        
    Else
        'Correcta
        Item.Icon = COLUMN_ATTACHMENT_ICON
        Item.ToolTip = "OK"
        
    End If
    Item.Caption = ""
    
    
    
    Set Item = Record.AddItem(Format(miRsAux!fecharec, "yyyymmdd"))
    Item.Caption = Format(miRsAux!fecharec, "dd/mm/yyyy")
    
    If miRsAux!NUmSerie = "1" Then
        Sql = ""
    Else
        Sql = miRsAux!NUmSerie & " "
    End If
    Sql = Sql & miRsAux!NumFactu
    Set Item = Record.AddItem(Sql)
    
    Set Item = Record.AddItem(Format(miRsAux!FecFactu, "yyyymmdd"))
    Item.Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
    
    
    Set Item = Record.AddItem(CStr(miRsAux!codmacta))
    Item.Value = CStr(miRsAux!codmacta)
    Set Item = Record.AddItem(DBLet(miRsAux!Nommacta, "T"))
    
    Record.AddItem DBLet(miRsAux!denominacion, "T")
    Record.AddItem CStr(DBLet(miRsAux!Descripcion, "T"))
    
    Set Item = Record.AddItem(miRsAux!totfacpr * 100)
    Item.Caption = Format(miRsAux!totfacpr, FormatoImporte)
    
    Record.Tag = miRsAux!NUmSerie & "|" & miRsAux!numregis & "|" & miRsAux!anofactu & "|" & miRsAux!SII_ID & "|"
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = "ID: " & miRsAux!Nommacta
    
    
    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
End Sub


Private Sub AddRecordASWSII()
Dim Sql As String
Dim Record As ReportRecord
    
    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    Set Item = Record.AddItem("")
    Sql = DBLet(miRsAux!Resultado, "T")
    Item.SortPriority = IIf(Sql <> "", 1, 0)
    If Sql = "Correcto" Then Sql = ""
    
    If Sql = "" Then
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
    
    
    Set Item = Record.AddItem(IIf(miRsAux!Enviada = 0, "NO", ""))
    
    
    Set Item = Record.AddItem(Mid(Sql, 1, 12)) 'resultado
    Item.ToolTip = DBLet(miRsAux!Mensaje, "T")
    
    Set Item = Record.AddItem(DBLet(miRsAux!REG_IDF_NumSerieFacturaEmisor, "T"))
    Set Item = Record.AddItem(Format(miRsAux!fecha1, "yyyymmdd"))
    Item.Caption = Format(miRsAux!fecha1, "dd/mm/yyyy")
    
    
    
    Record.AddItem DBLet(miRsAux!NIF, "T")
    Record.AddItem CStr(DBLet(miRsAux!Nombre, "T"))
    
    Set Item = Record.AddItem(miRsAux!Impor * 100)
    Item.Caption = Format(miRsAux!Impor, FormatoImporte)
    
    Record.Tag = miRsAux!id 'CLAVE
                
    
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
            
            wndReportControl.SelectedRows(0).Record(0).Checked = Not wndReportControl.SelectedRows(0).Record(0).Checked
            wndReportControl.SelectedRows(0).Selected = True
            wndReportControl.SetFocus
        Else
            
            If wndReportControl.FocusedRow Is Nothing Then
            
            Else
                 wndReportControl.FocusedRow.Record(0).Checked = Not wndReportControl.FocusedRow.Record(0).Checked
            End If
            
           
        End If
    End If
End Sub

Private Sub wndReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
 Dim Sql As String
    If Row Is Nothing Then Exit Sub
    
    If cboFacturas.ListIndex = 0 Then
        If Row.Record.Item(1).Icon = RECORD_UNREAD_MAIL_ICON Then
        

            If Me.cboTipo.ListIndex = 0 Then
                Sql = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_emitidas", "IDEnvioFacturasEmitidas", RecuperaValor(Row.Record.Tag, 4))
            Else
                Sql = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_recibidas", "IDEnvioFacturasRecibidas", RecuperaValor(Row.Record.Tag, 4))
            End If
            MsgBox "Factura incorrecta: " & vbCrLf & Sql, vbExclamation
        End If
   Else
        'Facturas ya subidas al ASWII
        Sql = ""
        If Me.cboFiltro.ListIndex = 1 Then
            Sql = "S"
        Else
            If Row.Record.Item(0).Icon = RECORD_IMPORTANCE_HIGH_ICON Then Sql = "s"
        End If
        If Sql <> "" Then
            If Me.cboTipo.ListIndex = 0 Then
                Sql = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_emitidas", "IDEnvioFacturasEmitidas", Row.Record.Tag)
            Else
                Sql = DevuelveDesdeBD("Mensaje", "aswsii.envio_facturas_recibidas", "IDEnvioFacturasRecibidas", Row.Record.Tag)
            End If
            MsgBox "ASWSII. Factura incorrecta: " & vbCrLf & Sql, vbExclamation
        End If
   End If

    
End Sub






Private Sub HacerGrabacionTablas_ASWISII()
Dim Sql As String
Dim ID_ASWSII As Long
Dim CuantasLlevo As Integer
Dim B As Boolean
    
    'Insertamos en tmpfaclin
    Label1.Caption = "Preparando campos"
    Label1.Refresh
    Conn.Execute "DELETE FROM tmpfaclin WHERE codusu = " & vUsu.Codigo
    
    Sql = ""
    For i = 0 To Me.wndReportControl.Records.Count - 1
        'codusu , Codigo, NUmSerie, NumFac, Fecha, Codigo
        'tmpfaclin
        If wndReportControl.Records(i).Item(0).Checked Then
            Msg = wndReportControl.Records(i).Tag
            Sql = Sql & ", (" & vUsu.Codigo & "," & DBSet(RecuperaValor(Msg, 1), "T")
            Sql = Sql & "," & RecuperaValor(Msg, 2) & ",'" & RecuperaValor(Msg, 3) & "-01-01',0)"
            
            
            If Len(Sql) > 600 Then
                Sql = Mid(Sql, 2)
                Sql = "INSERT INTO tmpfaclin(codusu ,  NUmSerie, NumFac, Fecha, Codigo) VALUES " & Sql
                Conn.Execute Sql
                Sql = ""
            End If
        End If
    Next
    
    If Sql <> "" Then
        Sql = Mid(Sql, 2)
        Sql = "INSERT INTO tmpfaclin(codusu ,  NUmSerie, NumFac, Fecha, Codigo) VALUES " & Sql
        Conn.Execute Sql
    End If
    
    
    
    'Ya tenemos las facturas que vamos a pasar
    Label1.Caption = "Insertando registros"
    Label1.Refresh
    
    
     
    If Me.cboTipo.ListIndex = 0 Then
        Sql = DevuelveDesdeBD("max(IDEnvioFacturasEmitidas)", "aswsii.envio_facturas_emitidas", "1", "1")
    Else
        Sql = DevuelveDesdeBD("max(IDEnvioFacturasRecibidas)", "aswsii.envio_facturas_recibidas", "1", "1")
    End If
    If Sql = "" Then Sql = "0"
    
    ID_ASWSII = Val(Sql) + 1
    
    
    Sql = "Select * from tmpfaclin where codusu = " & vUsu.Codigo & " ORDER BY numserie,numfac"
    
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic
    While Not miRsAux.EOF
    
    
        If Me.cboTipo.ListIndex = 0 Then
    
            Label1.Caption = "Fac:" & miRsAux!NUmSerie & miRsAux!NumFac
            Label1.Refresh
            B = Sii_FraCLI(miRsAux!NUmSerie, miRsAux!NumFac, Year(miRsAux!Fecha), ID_ASWSII, Sql)
            
        Else
            Label1.Caption = "Reg: " & miRsAux!NUmSerie & " :  " & miRsAux!NumFac
            Label1.Refresh
            B = Sii_FraPRO(miRsAux!NUmSerie, miRsAux!NumFac, Year(miRsAux!Fecha), ID_ASWSII, Sql)
            
        End If
        
        If B Then
            
            If Ejecuta(Sql) Then
                
                If Me.cboTipo.ListIndex = 0 Then
                    'UPDATEAMOS EN cabfact
                    Sql = "UPDATE factcli SET SII_ID =" & ID_ASWSII & ", SII_status=0"
                    Sql = Sql & " where numserie =" & DBSet(miRsAux!NUmSerie, "T") & " AND numfactu =" & miRsAux!NumFac & " AND anofactu =" & Year(miRsAux!Fecha)
                    
                Else
                    Sql = "UPDATE factpro SET SII_ID =" & ID_ASWSII & ", SII_status=0"
                    Sql = Sql & " where numserie =" & DBSet(miRsAux!NUmSerie, "T") & " AND numregis =" & miRsAux!NumFac & " AND anofactu =" & Year(miRsAux!Fecha)
                End If
                
                
                Conn.Execute Sql
                'Incrementamos contador
                ID_ASWSII = ID_ASWSII + 1
                
                
                
                
                CuantasLlevo = CuantasLlevo + 1
                If CuantasLlevo > 40 Then
                    Screen.MousePointer = vbHourglass
                    DoEvents
                    CuantasLlevo = 0
                    Label1.Caption = "Comprobando tabla"
                    Label1.Refresh
                    espera 0.5
                    
                    ComprobarResultadoEnviadasASWII IIf(Me.cboTipo.ListIndex = 1, 2, 1), Me.Label1
                    
                    
                    
                    
                End If
                
            End If
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    espera 1
    'Si envia directamente compruebo como estan las que hayan subido
    ComprobarResultadoEnviadasASWII IIf(Me.cboTipo.ListIndex = 1, 2, 1), Me.Label1
    
End Sub


Private Sub CargaCombos()

    Me.cboFiltro.Clear
    Me.cboTipo.Clear
    If Me.cboFacturas.ListIndex = 1 Then
        'YA COMUNICADAS
        cboTipo.AddItem "Emitidas"
        cboTipo.AddItem "Recibidas"
    Else
        'PEndientes
        
        cboTipo.AddItem "Emitidas"
        cboTipo.AddItem "Recibidas"
    End If
    
    
    'Filtro
    If Me.cboFacturas.ListIndex = 0 Then
        'Para pendientes
        Me.cboFiltro.AddItem "Todas"
        cboFiltro.AddItem "Erroneas"
        cboFiltro.AddItem "Pendientes"
        
    Else
        'Para comunicadas
        cboFiltro.AddItem "Pendiente"
        cboFiltro.AddItem "Erroneas"
        cboFiltro.AddItem "Correctas ultimo mes"
    End If


    cboTipo.ListIndex = 0
    cboTipo.Tag = 0
    cboFiltro.ListIndex = 0
    cboFiltro.Tag = 0
End Sub



Private Sub CargaDatospendientesASWSII()
Dim Sql As String
    On Error GoTo eCargaDatospendientesASWSII
    Screen.MousePointer = vbHourglass
    statusBar.Panels(1).Text = "Leyendo BD"
    
    Label True
    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
               
    Sql = SQL_ASWSII
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        AddRecordASWSII
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    wndReportControl.Populate
    
    

    
eCargaDatospendientesASWSII:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, Sql
    
    statusBar.Panels(1).Text = ""
    Label1.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

































'''***********************************************************************************
''
''
''
'''Leera los datos de mirsaux
''Private Sub AddRecord2()
''
''Dim Record As ReportRecord
''Dim Socio As Boolean
''Dim OtroIcono As Boolean
''Dim NoActivo As Boolean
''
''    On Error GoTo eAddRecord2
''
''    'Adds a new Record to the ReportControl's collection of records, this record will
''    'automatically be attached to a row and displayed with the Populate method
''    Set Record = wndReportControl.Records.Add()
''
''    Dim Item As ReportRecordItem
''    'Socio
''    Set Item = Record.AddItem("")
''    Socio = miRsAux!esSocio
''    NoActivo = DBLet(miRsAux!situclien, "N")
''
''    Item.SortPriority = IIf(Socio, 1, 0)
''
''
''    If NoActivo Then
''        Item.Icon = 7
''    Else
''        Item.Icon = IIf(Socio, RECORD_UNREAD_MAIL_ICON, -1)
''
''    End If
''    'Cuota
''    Set Item = Record.AddItem("")
''    OtroIcono = False
''    Item.SortPriority = 0
''    Item.Icon = 1
''
''    'Laboral
''    Set Item = Record.AddItem("")
''    OtroIcono = False
''    Item.SortPriority = 0
''    Item.Icon = 2
''    'Fiscal
''    Set Item = Record.AddItem("")
''    OtroIcono = False
''    Item.SortPriority = 0
''    Item.Icon = 3
''
''
''
''    ' '  codclien,nomclien,nifclien,matricula,licencia,essocio "
''    Set Item = Record.AddItem(CStr(miRsAux!CodClien))
''    Item.Value = CLng(miRsAux!CodClien)
''    Set Item = Record.AddItem(DBLet(miRsAux!NomClien, "T"))
''    If NoActivo Then Item.ForeColor = vbRed
''
''    Record.AddItem CStr(miRsAux!NIFClien)
''    Record.AddItem CStr(DBLet(miRsAux!Matricula, "T"))
''
''    Set Item = Record.AddItem(DBLet(miRsAux!licencia, "T"))
''    Item.Value = CLng(DBLet(miRsAux!licencia, "N"))
''
''
''    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
''    Record.PreviewText = "ID: " & miRsAux!CodClien
''
''
''
''
''eAddRecord2:
''
''End Sub
''
''
''
''
''
''Public Sub CreateReportControlPrueba()
''    'Start adding columns
''    Dim Column As ReportColumn
''    wndReportControl.Columns.DeleteAll
''    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
''    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "Socio", 18, False)
''    Column.Icon = COLUMN_IMPORTANCE_ICON
''    Set Column = wndReportControl.Columns.Add(COLUMN_ICON, "Cuotas", 18, False)
''    Column.Icon = COLUMN_MAIL_ICON
''    Set Column = wndReportControl.Columns.Add(COLUMN_ATTACHMENT, "Laboral", 18, False)
''    Column.Icon = COLUMN_ATTACHMENT_ICON
''    Set Column = wndReportControl.Columns.Add(3, "Fiscal", 18, False)
''    Column.Icon = COLUMN_ATTACHMENT_ICON
''
''
''    Set Column = wndReportControl.Columns.Add(4, "ID", 30, True)
''    Column.Alignment = xtpAlignmentRight
''
''    Set Column = wndReportControl.Columns.Add(5, "Nombre", 200, True)
''    Set Column = wndReportControl.Columns.Add(6, "DNI", 60, True)
''    Set Column = wndReportControl.Columns.Add(7, "Matricula", 55, True)
''    Set Column = wndReportControl.Columns.Add(8, "Licencia", 55, True)
''
''
''
''    wndReportControl.PaintManager.MaxPreviewLines = 1
''    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
''
''
''    'This font will be used in the BeforeDrawRow when automatic formatting is selected
''    'This simply applies Strikethrough to the currently set text font
''    Set fntStrike = wndReportControl.PaintManager.TextFont
''    fntStrike.Strikethrough = True
''
''    'This font will be used in the BeforeDrawRow when automatic formatting is selected
''    'This simply applies Bold to the currently set text font
''    Set fntBold = wndReportControl.PaintManager.TextFont
''    fntBold.Bold = True
''
''    'Any time you add or delete rows(by removing the attached record), you must call the
''    'Populate method so the ReportControl will display the changes.
''    'If rows are added, the rows will remain hidden until Populate is called.
''    'If rows are deleted, the rows will remain visible until Populate is called.
''    wndReportControl.Populate
''
''    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
''End Sub
''
''
''
''
''
''
''
''
''
'''Cuando modifiquemos o insertemos, pondremos el SQL entero
''Public Sub CargaDatosPrueba(ByVal Sql As String, EsTodoSQL As Boolean)
''Dim Aux  As String
''Dim Inicial As Integer
''Dim N As Integer
''Dim V As Boolean
''Dim T1 As Single
''
''
''
''    On Error GoTo eCargaDatos
''
''    Screen.MousePointer = vbHourglass
''
''    V = True
''
''    wndReportControl.ShowItemsInGroups = False
''    wndReportControl.Records.DeleteAll
''    wndReportControl.Populate
''
''    Set miRsAux = New ADODB.Recordset
''
''    If EsTodoSQL Then
''        Stop
''    Else
''        If Sql <> "" Then Sql = " WHERE " & Sql
''
''        Sql = " FROM arigestion1.clientes" & Sql
''        Sql = "SELECT codclien,nomclien,nifclien,matricula,licencia,essocio,situclien " & Sql
''
''        Sql = Sql & " ORDER BY codclien"
''    End If
''
''    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''    Inicial = 0
''
''    T1 = Timer
''    While Not miRsAux.EOF
''        AddRecord2
''
''        miRsAux.MoveNext
''    Wend
''    miRsAux.Close
''        Me.wndReportControl.Populate
''
''
''eCargaDatos:
''    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, Sql
''
''
''
''
''    Screen.MousePointer = vbDefault
''End Sub
''
''
