VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmTesCobrosAgente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros agente"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   16995
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   6735
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   16695
      _Version        =   1114114
      _ExtentX        =   29448
      _ExtentY        =   11880
      _StockProps     =   64
      MultipleSelection=   0   'False
      FreezeColumnsAbs=   0   'False
      MultiSelectionMode=   -1  'True
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
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
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   13
      Top             =   360
      Width           =   5085
   End
   Begin VB.TextBox Text2 
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
      Index           =   1
      Left            =   8040
      MaxLength       =   30
      TabIndex        =   2
      Top             =   360
      Width           =   1485
   End
   Begin VB.TextBox Text2 
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
      Left            =   6360
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   14880
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   10
      Top             =   360
      Width           =   1485
   End
   Begin VB.Frame FrameUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox Combo1 
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
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1365
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Impresion "
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Contabilizar"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   16440
      TabIndex        =   8
      Top             =   120
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
   Begin VB.Image imgCuentas 
      Height          =   240
      Index           =   0
      Left            =   9360
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Left            =   7560
      Picture         =   "frmTesCobrosParciales.frx":0000
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta cobro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8040
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "€ Cobrado "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   14880
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha cobro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmTesCobrosAgente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim Cad As String
Dim PrimeraVez As Boolean



Private Sub Combo1_Click()
 If PrimeraVez Then Exit Sub
    CargaCobrosParciales
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        DoEvents
        CargaCobrosParciales
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmppal.Icon
    PrimeraVez = True
    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 16
        .Buttons(2).Image = 37
    End With
    
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    Me.imgCuentas(0).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    
    Set miRsAux = New ADODB.Recordset
    Combo1.AddItem "Todos"
    
    Cad = "Select codusu from cobros_parciales GROUP BY 1"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux.Fields(0)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Combo1.ListIndex = 0
    CreateReportControlPendientes
    Text2(1).Text = ""
    Text4.Text = ""
    Text2(0).Text = Format(Now, "dd/mm/yyyy")
End Sub


Public Sub CreateReportControlPendientes()
    'Start adding columns
    Dim Column As ReportColumn
    wndReportControl.Columns.DeleteAll
    EstablecerFuente
    'Chekc
    Set Column = wndReportControl.Columns.Add(0, "", 15, False)
    Column.Icon = COLUMN_CHECK_ICON
     
   
    Set Column = wndReportControl.Columns.Add(1, "Fecha", 140, False)
    Column.Icon = COLUMN_ATTACHMENT_NORMAL_ICON
    
   
        Set Column = wndReportControl.Columns.Add(2, "NºFactura", 11, True)
        Set Column = wndReportControl.Columns.Add(3, "F.Factura", 10, True)
        Set Column = wndReportControl.Columns.Add(4, "Cliente", 11, True)
        Set Column = wndReportControl.Columns.Add(5, "Nombre", 30, True)
        Set Column = wndReportControl.Columns.Add(6, "Total factura", 10, True)
        Column.Alignment = xtpAlignmentRight
        Set Column = wndReportControl.Columns.Add(7, "Agente", 10, True)
        Set Column = wndReportControl.Columns.Add(8, "Tipo", 10, True)
        Set Column = wndReportControl.Columns.Add(9, "Pago €", 15, True)
        Column.Alignment = xtpAlignmentRight
        Set Column = wndReportControl.Columns.Add(10, "Gastos€", 0, False)
        Column.visible = False
    

    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    'Any time you add or delete rows(by removing the attached record), you must call the
    'Populate method so the ReportControl will display the changes.
    'If rows are added, the rows will remain hidden until Populate is called.
    'If rows are deleted, the rows will remain visible until Populate is called.
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
End Sub

Private Sub CargaCobrosParciales()
Dim SQL As String
Dim impo As Currency

    On Error GoTo ECargaDatos

    Screen.MousePointer = vbHourglass
    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
    
    'Monti NO las quiere marcadas por defecto
    
    Text1.Text = ""
    Text1.Tag = 0
    
    SQL = "select id,cobros_parciales.numserie,cobros_parciales.numfactu,cobros_parciales.fecfactu,cobros_parciales.numorden,"
    SQL = SQL & " cobros_parciales.tipoformapago,descformapago,fecha,impcobrado,cobros_parciales.codusu,cobros_parciales.observa,codmacta,nomclien,"
    SQL = SQL & " impvenci+coalesce(gastos,0)-coalesce(impcobro,0) Pdte,gastos "
    
    
    SQL = SQL & " from cobros_parciales left join cobros on  cobros_parciales.numserie =cobros.numserie"
    SQL = SQL & " and cobros_parciales.numfactu =cobros.numfactu and cobros_parciales.fecfactu =cobros.fecfactu and cobros_parciales.numorden =cobros.numorden"
    SQL = SQL & " left join tipofpago on  cobros_parciales.tipoformapago =tipofpago.tipoformapago"

    If Combo1.ListIndex > 0 Then SQL = SQL & " WHERE cobros_parciales.codusu = " & DBSet(Combo1.Text, "T")
    
    SQL = SQL & " order by fecha"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    impo = 0
    While Not miRsAux.EOF
        impo = impo + miRsAux!impcobrado
        AddRecordCli
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Text1.Text = Format(impo, FormatoImporte)
    Text1.Tag = impo
    wndReportControl.Populate
    
    

    
ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, SQL
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub




'socio, pendiente , nombre, matricula,licencia
'Leera los datos de mirsaux
Private Sub AddRecordCli()
Dim SQL As String
Dim Record As ReportRecord
Dim ItemToolTip As String
Dim ItemIcon As Integer
Dim Color As Long
Dim Impor As Currency

    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    
    
  
  
  
    Set Item = Record.AddItem("")
    If Combo1.ListIndex > 0 Then
        Item.TristateCheckbox = False
        Item.HasCheckbox = True
        Item.Checked = True
    Else
        Item.HasCheckbox = False
    End If
    
    
    Set Item = Record.AddItem(Format(miRsAux!Fecha, "yyyymmddhhnnss"))
    Item.Caption = Format(miRsAux!Fecha, "dd/mm/yyyy hh:nn")
        
    
    
    
    SQL = miRsAux!NUmSerie & "|" & Format(miRsAux!NumFactu, "0000000") & "|" & Format(miRsAux!numorden, "00") & "|"
    Set Item = Record.AddItem(SQL)
    Item.Caption = miRsAux!NUmSerie & Format(miRsAux!NumFactu, "0000000")
    
    Set Item = Record.AddItem(Format(miRsAux!FecFactu, "yyyymmdd"))
    Item.Caption = Format(miRsAux!FecFactu, "dd/mm/yyyy")
    
    
    Set Item = Record.AddItem(CStr(miRsAux!codmacta))
    Item.Value = CStr(miRsAux!codmacta)
    Set Item = Record.AddItem(DBLet(miRsAux!nomclien, "T"))
    
    If DBLet(miRsAux!nomclien, "T") = "" Then
        Item.Caption = "ERROR"
        Item.BackColor = vbRed
        
    End If
    
  
    Set Item = Record.AddItem(miRsAux!pdte * 100)
    Item.Caption = Format(miRsAux!pdte, FormatoImporte)
    
  
    SQL = miRsAux!codusu
    Record.AddItem CStr(SQL)
    
   
    SQL = DBLet(miRsAux!descformapago, "T")
    Set Item = Record.AddItem(DBLet(miRsAux!tipoformapago, "N"))
    Item.Caption = SQL
    
    
    
    Set Item = Record.AddItem(miRsAux!impcobrado * 100)
    Item.Caption = Format(miRsAux!impcobrado, FormatoImporte)
    Impor = miRsAux!pdte - miRsAux!impcobrado
    If DBLet(miRsAux!Gastos, "N") <> 0 Then
        
        If Impor > 0 Then
            Item.ToolTip = "Gastos: " & miRsAux!Gastos & "€      No paga total pendiente"
            'Tiene gastos y NO paga el total
            Item.ForeColor = vbRed
        Else
            Item.ToolTip = "Gastos: " & miRsAux!Gastos & "€"
            If Impor < 0 Then Item.ToolTip = Item.ToolTip & "    Importe > Pdte"
            Item.ForeColor = vbBlue
        End If
    Else
        If Impor < 0 Then
            Item.ToolTip = "Pago mayor importe pdte."
            Item.ForeColor = vbBlue
        ElseIf Impor > 0 Then Item.ToolTip = "Importe a cuenta"
            
        End If
    End If
    Set Item = Record.AddItem(DBLet(miRsAux!Gastos, "N"))
    
    
    
    
    Record.Tag = miRsAux!Id
                
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = "ID: "

    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
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



Private Function CargaTEmporalCobros() As Boolean
Dim B As Boolean


    On Error GoTo eCargaTEmporalCobros
    
    Conn.Execute "DELETE FROM tmpcobros2 WHERE codusu = " & vUsu.Codigo
    
    
    '
    
    Cad = ""
    NumRegElim = 0
    For i = 0 To Me.wndReportControl.Records.Count - 1
        B = True
        If Combo1.ListIndex > 0 Then
            If Not wndReportControl.Records(i).Item(0).Checked Then B = False
        End If
        If B Then
            NumRegElim = NumRegElim + 1
            
            'codusu,numserie,numfactu,fecfactu,numorden,
            Msg = wndReportControl.Records(i).Item(2).Value
            
            Cad = Cad & ", (" & vUsu.Codigo & "," & DBSet(RecuperaValor(Msg, 1), "T")
            Cad = Cad & "," & RecuperaValor(Msg, 2) & "," & DBSet(wndReportControl.Records(i).Item(3).Caption, "F")
            Cad = Cad & "," & RecuperaValor(Msg, 3)
            'codmacta,codforpa,referencia,cliente,
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(4).Caption, "T")
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(8).Value, "N")
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(8).Caption, "T")
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(5).Caption, "T")
            
            
            'fecvenci,impvenci,reftalonpag,gastos,text33csb
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(1).Caption, "F")
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(9).Caption, "N")
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(3).Caption, "T")  ' fecha factura
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(6).Caption, "N")  ' importe pdte factura
            Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(7).Caption, "T")   ' usuario
            Cad = Cad & ")"
        End If
        'Ejecut sql
        B = False
        If Len(Cad) > 6000 Then
            B = True
        Else
            If i = wndReportControl.Records.Count - 1 Then B = True
        End If
        If B Then
            Msg = " INSERT INTO tmpcobros2(codusu,numserie,numfactu,fecfactu,numorden,codmacta,codforpa,referencia,cliente,fecvenci,impvenci,reftalonpag,gastos,text33csb)  "
            Cad = Mid(Cad, 2)
            Cad = Msg & " VALUES " & Cad
            Conn.Execute Cad
            Cad = ""

        End If
    Next i
    
    If NumRegElim > 0 Then CargaTEmporalCobros = True
    
eCargaTEmporalCobros:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
End Function

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    Cad = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCuentas_Click(Index As Integer)
    Set frmBan = New frmBasico2
    Cad = ""
    AyudaBanco frmBan
    Set frmBan = Nothing
    If Cad <> "" Then
        Text2(1).Text = RecuperaValor(Cad, 1)
        Text4.Text = RecuperaValor(Cad, 2)
    End If
End Sub

Private Sub imgFecha_Click()
    Cad = Now
    If Text2(0).Text <> "" Then
        If IsDate(Text2(0).Text) Then Cad = CDate(Text2(0).Text)
    End If
    Cad = ""
    Set frmC = New frmCal
    frmC.Fecha = CDate(Cad)
    Cad = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If Cad <> "" Then
        
        If FechaCorrecta2(CDate(Cad), True) < 2 Then Text2(0).Text = Cad
        
        
    End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    ConseguirFoco Text2(Index), 3
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    If Index = 0 Then
        If Text2(0).Text <> "" Then
            If Not EsFechaOK(Text2(0)) Then
                MsgBox "Fecha incorrecta ", vbExclamation
                Text2(0).Text = ""
                PonFoco Text2(0)
            End If
        End If
    Else
        Cad = Trim(Text2(1).Text)
        Msg = ""
        If Cad <> "" Then
            If CuentaCorrectaUltimoNivel(Cad, Msg) Then
                Cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", Cad, "T")
                If Cad = "" Then
                    Msg = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            Else
                MsgBox Msg, vbExclamation
                Cad = ""
                Msg = ""
            End If
        End If
        If Text2(1).Text <> "" And Cad = "" Then PonFoco Text2(1)
        Text2(1).Text = Cad
        Text4.Text = Msg
       
    End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    If wndReportControl.Records.Count = 0 Then Exit Sub
    
    

    If Button.Index = 1 Then
        CargaTEmporalCobros
        InicializarVblesInformesGeneral True
         
        Cad = ""
        If Combo1.ListIndex > 0 Then
            Cad = "AGENTE: " & Combo1.Text
            If NumRegElim < wndReportControl.Records.Count Then Cad = Cad & "      Seleccionados(" & NumRegElim & "/" & wndReportControl.Records.Count & ")"
        End If
        cadParam = cadParam & "pDHCuentas=""" & Cad & """|"
        numParam = 2
        cadNomRPT = "rCobrosAgente.rpt"
    
        cadFormula = "{tmpcobros2.codusu}=" & vUsu.Codigo
    
        ImprimeGeneral
    
    Else
        If Combo1.ListIndex = 0 Then
            MsgBoxA "Seleccione un único agente", vbExclamation
            Exit Sub
        End If
        
        'Comporbaciones
        Cad = ""
        If Text2(0).Text = "" Then
            Cad = "-Falta fecha contabilizacion"
        Else
            i = CInt(FechaCorrecta2(CDate(Text2(0).Text)))
            If i > 2 Then Cad = "-Fecha fuera de ejercicios"
        End If
        If Text2(1).Text = "" Then
            Cad = Cad & vbCrLf & "-Cuenta cobro"
        Else
            If Text4.Text = "" Then Cad = Cad & vbCrLf & "-Cuenta cobro(2)"
        End If
        
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
        
        Cad = ""
        NumRegElim = 0
        Msg = ""
        XAnt = 0
        YAnt = 0
        For i = 0 To Me.wndReportControl.Records.Count - 1
            If wndReportControl.Records(i).Item(0).Checked Then
                
                NumRegElim = NumRegElim + 1
                If wndReportControl.Records(i).Item(9).ForeColor = vbRed Then
                    'SI el importe NO es por el total
                    If wndReportControl.Records(i).Item(9).Caption <> wndReportControl.Records(i).Item(6).Caption Then
                        Cad = "Tiene gastos y no es el total pendiente: " & wndReportControl.Records(i).Item(2).Caption & vbCrLf
                    End If
                    
                End If
                YAnt = YAnt + wndReportControl.Records(i).Item(9).Caption
                'Si paga con documento
                If Not (wndReportControl.Records(i).Item(8).Value = 0 Or wndReportControl.Records(i).Item(8).Value = 6) Then
                    'El pago fue mediante talon, pagare... algo que no fuera efectivo(tarjeta)
                    XAnt = XAnt + wndReportControl.Records(i).Item(9).Caption
                    Msg = Msg & "X"
                End If
                
                
            End If
        Next
        If XAnt > 0 Then
            Msg = Len(Msg)
            Msg = "Documentos entregados: " & Msg & "   Tot:" & Format(XAnt, FormatoImporte)
              
        End If
        
        YAnt = 0: XAnt = 0
        If NumRegElim = 0 Then Cad = "Seleccione algun recibo"
        If Cad <> "" Then
            MsgBox Cad, vbExclamation
            Exit Sub
        End If
        If NumRegElim = wndReportControl.Records.Count Then
            Cad = NumRegElim
        Else
            Cad = NumRegElim & " de " & wndReportControl.Records.Count
        End If
        Cad = "Importe : " & Text1.Text & "     Cobros: " & Cad
        If Msg <> "" Then Cad = Cad & vbCrLf & Msg
        Cad = Cad & vbCrLf & vbCrLf
        Cad = "Va a contabilizar cobros del agente " & Combo1.Text & vbCrLf & vbCrLf & Cad & "¿Continuar?"
        If MsgBoxA(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
     
        
        Screen.MousePointer = vbHourglass
        Conn.BeginTrans
        
        If ContabilizarCobros(CDate(Text2(0).Text), Text2(1).Text) Then
            Conn.CommitTrans
            CargaCobrosParciales
        Else
            Conn.RollbackTrans
            
        End If
        
        
        Screen.MousePointer = vbDefault
    End If
End Sub



Private Function ContabilizarCobros(FechaApunte As Date, CtaBanco As String) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Linea  As Integer
Dim Mc As Contadores
Dim impo As Currency
Dim FP As Ctipoformapago
Dim Ampliacion As String
Dim Numdocum As String
Dim Conce As Integer
Dim LlevaContr As Boolean
Dim ElConcepto
Dim Gastos As Currency
Dim IngresosExtra As Currency
Dim CtaIngresosBanco As String
Dim ImporteApunte As Currency
    
    On Error GoTo eContabilizarCobros
    ContabilizarCobros = False
    
    
    
        CtaIngresosBanco = DevuelveDesdeBD("ctaingreso", "bancos", "codmacta", CtaBanco, "T")
        If CtaIngresosBanco = "" Then Err.Raise 513, , "Cuenta intresos banco sin configurar"
    
        Set Mc = New Contadores
        Mc.ConseguirContador "0", FechaApunte <= vParam.fechaini, True
        
        Set FP = New Ctipoformapago
        FP.tipoformapago = -1
        
        
        NumRegElim = 0
        For i = 0 To Me.wndReportControl.Records.Count - 1
            If wndReportControl.Records(i).Item(0).Checked Then
                impo = impo + ImporteFormateado(wndReportControl.Records(i).Item(9).Caption)
                NumRegElim = NumRegElim + 1
            End If
        Next i
        Ampliacion = "Agente: " & Combo1.Text & "       Cobros: " & NumRegElim & " de " & wndReportControl.Records.Count & "     Importe: " & Format(impo, FormatoImporte) & "€"
       
        Ampliacion = "Generado desde Tesorería el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & vbCrLf & vbCrLf & Ampliacion
        
        
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
        SQL = SQL & "1,'" & Format(FechaApunte, FormatoFecha) & "'," & Mc.Contador
        SQL = SQL & ",  " & DBSet(Ampliacion, "T") & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar cobros agentes'"

        
        SQL = SQL & ")"
        
        'Nada mas vaya a meter la primeralinea lo creo
        'Podria ser que le hubieran entragado solo talones(o pagares) y no tuviera que hacer el apunte desde aqui
        CadenaDesdeOtroForm = SQL    'Conn.Execute SQL
        
        
        Msg = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        Msg = Msg & "codmacta, numdocum, codconce, ampconce,timporteD,"
        Msg = Msg & " timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfaccl, fecfactu, numorden, tipforpa) "
            
        
        
        Linea = 0
        Gastos = 0
        ImporteApunte = 0
        For i = 0 To Me.wndReportControl.Records.Count - 1
            
            
            If wndReportControl.Records(i).Item(0).Checked Then
                'Se contabilizaran formas de pago que no sean: 0,6  EFECTIVO y Tarjeta
                If wndReportControl.Records(i).Item(8).Value = 0 Or wndReportControl.Records(i).Item(8).Value = 6 Then
                    B = True
                Else
                    B = False
                End If
            Else
                B = False
            End If
            
            If wndReportControl.Records(i).Item(0).Checked Then
            
                
                If B Then    'OK forma de pago efectivo / tarjeta
                   
                   
                   If Linea = 0 Then
                        'Inserto cabapu
                        Conn.Execute CadenaDesdeOtroForm
                   
                   
                   End If
                   'importe cobrado
                   impo = ImporteFormateado(wndReportControl.Records(i).Item(9).Caption)  'Total cobrado
                   Linea = Linea + 1
                   ImporteApunte = ImporteApunte + impo
                   
                   'Gsastos: Cojo un segundo. Total pdte
                   Gastos = ImporteFormateado(wndReportControl.Records(i).Item(6).Caption) 'Total cobrado
                   If impo > Gastos Then
                       'Ha pagodo mas que el pendiente
                       IngresosExtra = impo - Gastos
                   Else
                       IngresosExtra = 0
                   End If
                   
                   'Gastos del RECIBO
                   Gastos = ImporteFormateado(wndReportControl.Records(i).Item(10).Caption)
                   
                   impo = impo - Gastos - IngresosExtra
                   If impo < 0 Then Err.Raise 513, , "Importe negativo despues gastos"
                   
                   If FP.tipoformapago <> wndReportControl.Records(i).Item(8).Value Then
                       If FP.Leer(wndReportControl.Records(i).Item(8).Value) <> 0 Then Err.Raise 513, , "Forma de pago erronea: " & wndReportControl.Records(i).Item(8).Caption
                       FP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.conhacli)
                       FP.CadenaAuxiliar = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.condecli)
                       
                       ElConcepto = FP.conhacli
                       Conce = FP.amphacli
                       LlevaContr = FP.ctrhacli = 1
                       
                       
                   End If
                   'Inserto en las lineas de apuntes
                   SQL = Msg & " VALUES ("
                   SQL = SQL & "1"
                   SQL = SQL & ",'" & Format(Now, FormatoFecha) & "'," & Mc.Contador & ","
                   
                   
                   'numdocum
                   Numdocum = wndReportControl.Records(i).Item(2).Caption  ' letra de serie y factura
                   
                  
            
                          
               
                   Ampliacion = FP.descformapago & " "
                   If Conce = 2 Then
                       Ampliacion = Ampliacion & wndReportControl.Records(i).Item(2) 'DBLet(Rs!FecVenci)  'Fecha vto
                   ElseIf Conce = 4 Then
                       'Contra partida
                       Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaBanco, "T")
                   Else
                       
                      If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
                      Ampliacion = Ampliacion & wndReportControl.Records(i).Item(2).Caption      'serie factura Text1(2).Text & "/" & Text1(0).Text 'RecuperaValor(Vto, 1) & "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
                   End If
                   
                   'Fijo en concepto el codconce
                   Ampliacion = Mid(Ampliacion, 1, 50)
                   
                   'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
                   'Cuenta Cliente
                   Cad = Linea & ",'" & wndReportControl.Records(i).Item(4).Caption & "','" & Numdocum & "'," & FP.conhacli & ",'" & DevNombreSQL(Ampliacion) & "',"
                   
                   Cad = Cad & "NULL," & TransformaComasPuntos(CStr(impo))
                   
                   'Codccost
                   Cad = Cad & ",NULL,"
                   If LlevaContr Then
                       Cad = Cad & "'" & CtaBanco & "'"
                   Else
                       Cad = Cad & "NULL"
                   End If
                   Ampliacion = wndReportControl.Records(i).Item(2).Value  'NUMSERIE|Numfactu|numorden
                   Cad = Cad & ",'COBROS',0," & DBSet(RecuperaValor(Ampliacion, 1), "T") & "," & DBSet(RecuperaValor(Ampliacion, 2), "N")
                   Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(3).Caption, "F") & "," & DBSet(RecuperaValor(Ampliacion, 3), "N") & "," & DBSet(FP.tipoformapago, "N") & ")"
                   Cad = SQL & Cad
                   Conn.Execute Cad
                   
                   
                   'Aqui, si tiene gastos VAN separados
                   
                   If Gastos > 0 Or IngresosExtra > 0 Then
                       
                       'Hace el mismo apunte bien sea porque el vto tenia gastos o pq me ha pagado mas
                       'La diferencia la pondre en la ampliacion añadiendo un [Gast] o [Exc]
                       
                       For J = 1 To 2
                           
                           If J = 1 Then
                               impo = Gastos
                           Else
                               impo = IngresosExtra
                           End If
                           
                           If impo > 0 Then
                               'Inserto en las lineas de apuntes
                               Linea = Linea + 1
                               SQL = Msg & " VALUES ("
                               SQL = SQL & "1"
                               SQL = SQL & ",'" & Format(Now, FormatoFecha) & "'," & Mc.Contador & ","
                               
                               Ampliacion = FP.descformapago & " "
                               If Conce = 2 Then
                                   Ampliacion = Ampliacion & wndReportControl.Records(i).Item(2) 'DBLet(Rs!FecVenci)  'Fecha vto
                               ElseIf Conce = 4 Then
                                   'Contra partida
                                   Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaBanco, "T")
                               Else
                                  If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
                                  Ampliacion = Ampliacion & wndReportControl.Records(i).Item(2).Caption      'serie factura Text1(2).Text & "/" & Text1(0).Text 'RecuperaValor(Vto, 1) & "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
                               End If
                               
                               'Fijo en concepto el codconce
                               Ampliacion = Trim(Mid(Ampliacion, 1, 45))
                               Ampliacion = Ampliacion & " " & IIf(J = 1, "[Gast]", "[Exce]")
                               'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
                               'Cuenta Cliente
                               Cad = Linea & ",'" & CtaIngresosBanco & "','" & Numdocum & "'," & FP.conhacli & ",'" & DevNombreSQL(Ampliacion) & "',"
                               
                               'Codccost
                               Cad = Cad & "NULL," & TransformaComasPuntos(CStr(impo))
                               
                               'Codccost
                               Cad = Cad & ",NULL,"
                               'ctrapar
                               If LlevaContr Then
                                   Cad = Cad & "'" & CtaBanco & "'"
                               Else
                                   Cad = Cad & "NULL"
                               End If
                              
                               Ampliacion = wndReportControl.Records(i).Item(2).Value  'NUMSERIE|Numfactu|numorden
                               Cad = Cad & ",'COBROS',0," & DBSet(RecuperaValor(Ampliacion, 1), "T") & "," & DBSet(RecuperaValor(Ampliacion, 2), "N")
                               Cad = Cad & "," & DBSet(wndReportControl.Records(i).Item(3).Caption, "F") & "," & DBSet(RecuperaValor(Ampliacion, 3), "N") & "," & DBSet(FP.tipoformapago, "N") & ")"
                               Cad = SQL & Cad
                               Conn.Execute Cad
                              
                           End If
                           
                       Next J
                       
                       
                   
                   End If
                   
                   
                   
                   'UPDATEAR EL PAGO
                   
                   
                   'Veremos, si lo ha pagado todo o no
                   Gastos = ImporteFormateado(wndReportControl.Records(i).Item(9).Caption)  'Total cobrado
                   impo = ImporteFormateado(wndReportControl.Records(i).Item(6).Caption) 'Total pdte
                   
                   If Gastos >= impo Then
                       'LO HA PAGADO TODO
                        Cad = "impcobro = coalesce(gastos,0) + impvenci,situacion=1"
                   Else
                       'Hay pdte
                       Cad = "impcobro = coalesce(impcobro,0) + " & DBSet(Gastos, "N")
                   End If
                   
                   SQL = "UPDATE cobros SET fecultco=" & DBSet(Text2(0).Text, "F") & "," & Cad
                   
                   
                   Ampliacion = wndReportControl.Records(i).Item(2).Value  'NUMSERIE|Numfactu|numorden
                   SQL = SQL & " WHERE numserie = " & DBSet(RecuperaValor(Ampliacion, 1), "T") & " AND numfactu = " & DBSet(RecuperaValor(Ampliacion, 2), "N")
                   SQL = SQL & " AND fecfactu = " & DBSet(wndReportControl.Records(i).Item(3).Caption, "F") & " and numorden = " & DBSet(RecuperaValor(Ampliacion, 3), "N")
                   Conn.Execute SQL
            End If 'Forma de pago efectivo tarjeta . Para el resto NO hacemos apunte
            
            'Borrar el cobro parcial
            SQL = "DELETE from cobros_parciales WHERE id = " & wndReportControl.Records(i).Tag
            Conn.Execute SQL
        End If
    Next i
    

    'CUADRE APUNTE BANCO
    If Linea > 0 Then
        'Ha creado lineas
        Linea = Linea + 1
        impo = ImporteApunte
        SQL = Msg & " VALUES ("
        SQL = SQL & "1"
        SQL = SQL & ",'" & Format(Now, FormatoFecha) & "'," & Mc.Contador & ","
        
        If NumRegElim = wndReportControl.Records.Count Then
            Ampliacion = NumRegElim
        Else
            Ampliacion = NumRegElim & "/" & wndReportControl.Records.Count
        End If
        Ampliacion = "Cobro agente " & Combo1.Text & ".  Vtos: " & Ampliacion
        
    
        
        Ampliacion = Mid(Ampliacion, 1, 50)
        
        'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
        'Cuenta Cliente
        Cad = Linea & ",'" & CtaBanco & "','" & Numdocum & "'," & FP.condecli & ",'" & DevNombreSQL(Ampliacion) & "',"
        
        '
        Cad = Cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
        
        'Codccost  contrpar
        Cad = Cad & ",NULL,NULL"
        
        
        Cad = Cad & ",'COBROS',0,NULL,NULL,NULL,NULL,NULL)"
        Cad = SQL & Cad
        Conn.Execute Cad
        
    End If
    ContabilizarCobros = True 'TODO BIEN

    
    
    
eContabilizarCobros:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set Mc = Nothing
    Set FP = Nothing
    Set miRsAux = Nothing
    CadenaDesdeOtroForm = ""
End Function






Private Sub wndReportControl_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim impo As Currency

    
    Debug.Print Row.Record.Item(9).Caption
    impo = ImporteFormateado(Row.Record.Item(9).Caption)
    If Not Row.Record.Item(0).Checked Then impo = -impo
    impo = Text1.Tag + impo
    Text1.Tag = impo
    Text1.Text = Format(impo, FormatoImporte)
End Sub
