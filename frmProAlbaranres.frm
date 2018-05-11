VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmProAlbaranres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes proveedor"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   6735
      Left            =   120
      TabIndex        =   3
      Top             =   750
      Width           =   13095
      _Version        =   1114114
      _ExtentX        =   23098
      _ExtentY        =   11880
      _StockProps     =   64
      MultipleSelection=   0   'False
      FreezeColumnsAbs=   0   'False
      MultiSelectionMode=   -1  'True
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   12240
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1545
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   12960
      TabIndex        =   2
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
End
Attribute VB_Name = "frmProAlbaranres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Carpeta As String



Private Const IdPrograma = "404-01"


Dim fntBold As StdFont
Dim fntStrike As StdFont



Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If fntStrike Is Nothing Then
        CreateReportControlPendientes
        CargaDatosPendientes
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    
    Set fntStrike = Nothing
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 3
               
        '.Buttons(2).Image = 4
        .Buttons(2).visible = False
        .Buttons(3).Image = 5
       
       
    End With


End Sub


Public Sub CreateReportControlPendientes()
    wndReportControl.PaintManager.NoItemsText = "Ningún registro "
    wndReportControl.Records.DeleteAll
    
    
    'Start adding columns
    Dim Column As ReportColumn
    wndReportControl.Columns.DeleteAll
    
    'Chekc
    Set Column = wndReportControl.Columns.Add(0, "", 18, False)
    Column.Icon = COLUMN_CHECK_ICON
     
    
        
    Set Column = wndReportControl.Columns.Add(2, "Cuenta", 75, False)
    Set Column = wndReportControl.Columns.Add(3, "Proveedor", 10, True)
    Set Column = wndReportControl.Columns.Add(4, "Albaran", 5, True)
    Set Column = wndReportControl.Columns.Add(5, "Fecha", 95, False)
    Set Column = wndReportControl.Columns.Add(6, "Base imponible", 5, True)
    Column.Icon = 12
    Column.Alignment = xtpAlignmentRight
    
    Set Column = wndReportControl.Columns.Add(7, "Fichero", 8, True)
    Column.visible = False
    
    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
                  
    
    'This font will be used in the BeforeDrawRow when automatic formatting is selected
    'This simply applies Strikethrough to the currently set text font
    wndReportControl.PaintManager.TextFont.Name = "Verdana"
    wndReportControl.PaintManager.TextFont.SIZE = 10
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    CadenaDesdeOtroForm = ""
    If Button.Index = 1 Then
        frmAlfresQ.Carpeta = Carpeta
        frmAlfresQ.Show vbModal
        
        
    ElseIf Button.Index = 3 Then
        If Me.wndReportControl.Records.Count = 0 Then Exit Sub
        If Me.wndReportControl.SelectedRows.Count <> 1 Then
            MsgBox "Seleccione UN albaran para eliminar", vbExclamation
            Exit Sub
        End If
        Msg = ""
        For i = 1 To 5
            Msg = Msg & Me.wndReportControl.Columns(i).Caption & ": " & wndReportControl.SelectedRows(0).Record(i + 1).Caption & vbCrLf
        Next i
        
        
        Msg = "Va a eliminar el albaran: " & vbCrLf & Msg & vbCrLf & "¿Continuar?"
        If MsgBox(Msg, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        If EliminaAlbaran Then CadenaDesdeOtroForm = "OK"
        
        
    End If
    
    If CadenaDesdeOtroForm = "" Then Exit Sub
    CargaDatosPendientes
    CadenaDesdeOtroForm = ""
End Sub





Private Sub CargaDatosPendientes()
Dim SQL As String

    On Error GoTo ECargaDatos

    Screen.MousePointer = vbHourglass

    

    wndReportControl.ShowItemsInGroups = False
    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    
    Set miRsAux = New ADODB.Recordset
    
    
              
    SQL = "select factproalbaranes.*,nommacta from factproalbaranes left join cuentas on factproalbaranes.codmacta=cuentas.codmacta "
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        
        AddRecordCli
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
    wndReportControl.Populate
    
    

    
ECargaDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, SQL
    Screen.MousePointer = vbDefault
End Sub




'socio, pendiente , nombre, matricula,licencia
'Leera los datos de mirsaux
Private Sub AddRecordCli()
Dim SQL As String
Dim Record As ReportRecord
Dim ItemToolTip As String
Dim ItemIcon As Integer

    On Error GoTo eAddRecord2

    'Adds a new Record to the ReportControl's collection of records, this record will
    'automatically be attached to a row and displayed with the Populate method
    Set Record = wndReportControl.Records.Add()
    
    Dim Item As ReportRecordItem
    
    
    
    i = 1
    ItemIcon = RECORD_REPLIED_ICON
    ItemToolTip = "En progreso"



    
        'Check
    Set Item = Record.AddItem("")
  
        Item.Checked = False
        Item.HasCheckbox = False

    Item.TristateCheckbox = False
    
    
    'Estado
    Set Item = Record.AddItem("")
    Item.Icon = ItemIcon
    Item.ToolTip = ItemToolTip
    Item.SortPriority = i
    Item.Caption = ""
    
    
    
    
    SQL = miRsAux!codmacta
    Set Item = Record.AddItem(SQL)
    Item.Value = CStr(miRsAux!codmacta)
    
    If IsNull(miRsAux!Nommacta) Then
        SQL = "ERRORRRRRR"
        '.BackColor = vbRed
    Else
        SQL = miRsAux!Nommacta
    End If
    Set Item = Record.AddItem(SQL)
    
    SQL = miRsAux!numalbar
    Set Item = Record.AddItem(SQL)
    Item.Value = SQL
    
    
    Set Item = Record.AddItem(Format(miRsAux!fechaalb, "yyyymmdd"))
    Item.Caption = Format(miRsAux!fechaalb, "dd/mm/yyyy")
    
    
    Set Item = Record.AddItem(miRsAux!BIMponible * 100)
    Item.Caption = Format(miRsAux!BIMponible, FormatoImporte)
    
    Set Item = Record.AddItem(CStr(miRsAux!Nombre))
        
    'Set Item = Record.AddItem(miRsAux!totfacpr * 100)
    'Item.Caption = Format(miRsAux!totfacpr, FormatoImporte)
    Record.Tag = miRsAux!Id
                
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
    Record.PreviewText = DBLet(miRsAux!Nommacta, "T")
    
   
    
    Exit Sub
eAddRecord2:
    MuestraError Err.Number
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
End Sub

Private Sub wndReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error GoTo eDble
    
    If Row Is Nothing Then Exit Sub
    
    
    If Row.Record(7).Caption = "" Then Exit Sub
    
    Msg = Carpeta & "\" & Row.Record(7).Caption
    
    If Dir(Msg, vbArchive) = "" Then
        MsgBox "No existe el fichero: " & Msg, vbExclamation
        Exit Sub
    End If
    
    Call ShellExecute(Me.hwnd, "Open", Msg, "", "", 1)
    
    
    Exit Sub
eDble:
    MuestraError Err.Number, "Fichero: " & Msg & vbCrLf & vbCrLf & Err.Description

End Sub


Private Function EliminaAlbaran() As Boolean
    On Error GoTo eDble2
    
    EliminaAlbaran = False
    
    Msg = Carpeta & "\" & wndReportControl.SelectedRows(0).Record(i + 1).Caption
    
    If Dir(Msg, vbArchive) = "" Then
        MsgBox "No existe el fichero: " & Msg, vbExclamation
        Exit Function
    End If
    
    Kill Msg
    
    
    
    Msg = "DELETE FROM factproalbaranes where id = " & wndReportControl.SelectedRows(0).Record.Tag
    Conn.Execute Msg
    
    EliminaAlbaran = True
    
    Exit Function
eDble2:
    MuestraError Err.Number, "Fichero: " & Msg & vbCrLf & vbCrLf & Err.Description

End Function

