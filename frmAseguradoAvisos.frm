VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#17.2#0"; "Codejock.ReportControl.v17.2.0.ocx"
Begin VB.Form frmASeguradoAvisos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10380
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   19788
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   19788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   8655
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   19335
      _Version        =   1114114
      _ExtentX        =   34105
      _ExtentY        =   15266
      _StockProps     =   64
      MultipleSelection=   0   'False
      FreezeColumnsAbs=   0   'False
      MultiSelectionMode=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18240
      TabIndex        =   2
      Top             =   9840
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Aviso siniestro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Aviso falta de pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   9840
      Width           =   10965
   End
End
Attribute VB_Name = "frmASeguradoAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.- Los dos
    ' 1.- Solo Falta pago
    ' 2.- solo Siniestro

Dim PrimVez As Boolean
Dim SQL As String


Private Sub Command1_Click()

    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        
        CargaListviewAsegurados2
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimVez = True
    Me.Icon = frmppal.Icon
    
    Me.Caption = "Listado falta aviso asegurado   -   " & Format(Now, "dd/mm/yyyy")
    
    Me.Option1(0).visible = Opcion <> 2
    Me.Option1(1).visible = Opcion <> 1
    If Opcion <> 2 Then Me.Option1(0).Value = True
    
    'Set Me.ListView1.SmallIcons = frmppal.ImgListviews
    
    EstablecerFuente
    
    
     
    
    
End Sub

Private Sub CargaListviewAsegurados2()
    Set miRsAux = New ADODB.Recordset
    Label1.Caption = "Leyendo BD"
    CargaListviewAsegurados
    
    Set miRsAux = Nothing
    Label1.Caption = ""
End Sub
    
Private Sub CargaListviewAsegurados()
    On Error GoTo eCargaListviewAsegurados

    wndReportControl.Records.DeleteAll
    wndReportControl.Populate
    CargaColumnas
    wndReportControl.Populate

   

    'Monta EL SQL  '
    MontaSQLAvisosSeguros Option1(0).Value, SQL
    
    SQL = "Select numserie,numfactu,fecfactu,fecvenci, cobros.codmacta ,nommacta,impvenci,numorden,devuelto,fecprorroga " & SQL
    SQL = SQL & " ORDER  BY fecvenci ,impvenci"
    
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        AddRecord2
        
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
    Exit Sub
eCargaListviewAsegurados:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub Option1_Click(Index As Integer)
    
    CargaListviewAsegurados2
End Sub



Private Sub CargaColumnas()
Dim Columnas As String
Dim Ancho As String
Dim ALIGN As String
Dim Datos As String
Dim NCols As Integer
Dim i As Integer
Dim Column As ReportColumn



    wndReportControl.Columns.DeleteAll
    'Adds a new ReportColumn to the ReportControl's collection of columns, growing the collection by 1.
    wndReportControl.PaintManager.MaxPreviewLines = 1
    wndReportControl.PaintManager.HorizontalGridStyle = xtpGridNoLines
    'fntStrike.Strikethrough = True
    'fntBold.Bold = True
    wndReportControl.Populate
    
    wndReportControl.SetCustomDraw xtpCustomBeforeDrawRow
    
    

    Set Column = wndReportControl.Columns.Add(COLUMN_IMPORTANCE, "", 18, False)
    Column.ToolTip = "Devuelto"

     If Me.Option1(0).Value Then
         NCols = 9
         Columnas = "Serie|Factura|F. Factura|F. VTO|Dias|Cuenta|Nombre|Importe|NumVenci|"
         Ancho = "14|30|28|28|18|30|90|30|0|"
         ALIGN = "LLLLDLLDD"
         Datos = "TNFFNTTDN"
     
     Else
         NCols = 10
         Columnas = "Serie|Factura|F. Factura|F. VTO|Dias|F.Prorroga|Cuenta|Nombre|Importe|NumVenci|"
         Ancho = "14|30|28|28|18|28|30|90|30|0|"
         ALIGN = "LLLLDLDLDD"
         Datos = "TNFFNFTTDN"
     End If
     
        

        
    For i = 1 To NCols
         SQL = RecuperaValor(Columnas, i)
         If SQL <> "" Then
            
            
            K = 1 'Si sera visible
            
            Msg = RecuperaValor(Ancho, i)
            If InStr(1, Msg, "%") > 0 Then

                Msg = Mid(Msg, 1, Len(Msg) - 1)
                
            Else
                If Msg = "0" Then K = 0

            End If
            
            
            Set Column = wndReportControl.Columns.Add(Val(i), SQL, Val(Msg), True)
            Column.visible = K = 1
            'align
            SQL = Mid(ALIGN, i, 1)
            If SQL = "L" Then
                'NADA. Es valor x defecto
                Column.Alignment = xtpAlignmentLeft
            Else
                If SQL = "D" Then
                    'ColX.Alignment = lvwColumnRight
                    Column.Alignment = xtpAlignmentRight
                Else
                    'CENTER
                    Column.Alignment = xtpAlignmentLeft
                End If
            End If
            Msg = Mid(Datos, i, 1)
            Column.Tag = Msg
             
         End If
     Next i

    
    



End Sub




'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'
'
'       Window report
'
Private Sub EstablecerFuente()

    On Error GoTo eEstablecerFuente
    'The following illustrate how to change the different fonts used in the ReportControl
    Dim TextFont As StdFont
    Label1.Font.Bold = False
    Set TextFont = Label1.Font
    TextFont.SIZE = 11
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

    Label1.Font.Bold = True
    Exit Sub
eEstablecerFuente:
    MuestraError Err.Number, Err.Description

End Sub




Private Sub Label(visible As Boolean)
    If visible Then
        Label1.Caption = "Leyendo registros BD"
    Else
        Label1.Caption = ""
    End If
    Label1.Refresh
End Sub




Private Sub AddRecord2()
Dim Record As ReportRecord
Dim Item As ReportRecordItem
Dim F As Date
Dim Dias As Integer
Dim Rojo As Boolean

    On Error GoTo eAddRecord2

    Set Record = wndReportControl.Records.Add()
    Set Item = Record.AddItem("")
    
    Item.SortPriority = IIf(DBLet(miRsAux!Devuelto, "N") = 0, 1, 0)

    
    Set Item = Record.AddItem(CStr(miRsAux!NUmSerie))
   
    SQL = Format(miRsAux!NumFactu, "000000")
    Set Item = Record.AddItem(SQL)
    If False Then Item.ForeColor = vbRed
    
    Set Item = Record.AddItem(Format(miRsAux!FecFactu, "yyyymmdd"))
    Item.Value = Format(miRsAux!FecFactu, "dd/mm/yyyy")
    
    Set Item = Record.AddItem(Format(miRsAux!FecVenci, "yyyymmdd"))
    Item.Value = Format(miRsAux!FecVenci, "dd/mm/yyyy")
    
    
    i = 5
    Rojo = False
    If Option1(0).Value Then
        If vParamT.FechaSeguroEsFra Then
            F = miRsAux!FecFactu
        Else
            F = miRsAux!FecVenci
        End If
        Dias = DateDiff("d", F, Now)  'hace dias
        If Dias > vParamT.DiasMaxAvisoH Then Rojo = True
    Else
        'SINIESTRO
        If IsNull(miRsAux!fecprorroga) Then
            'NO ha prorrogado
            If vParamT.FechaSeguroEsFra Then
                F = miRsAux!FecFactu
            Else
                F = miRsAux!FecVenci
            End If
            Dias = DateDiff("d", F, Now)  'hace dias
            If Dias > vParamT.DiasMaxSiniestroH Then Rojo = True
            
        Else
           
            Dias = DateDiff("d", miRsAux!fecprorroga, Now)  'hace dias
            If Dias > vParamT.DiasAvisoDesdeProrroga Then Rojo = True
             
        End If
        i = 6
        
        
        
    End If
    
    
    
        
    
    
    Set Item = Record.AddItem(Format(Dias, "000000"))
    Item.Value = Dias
    
    If Option1(1).Value Then
        F = "01/01/1900"
        SQL = " "
        If Not IsNull(miRsAux!fecprorroga) Then
            F = miRsAux!fecprorroga
            SQL = Format(miRsAux!fecprorroga, "dd/mm/yyyy")
        End If
        
        Set Item = Record.AddItem(Format(F, "yyyymmdd"))
        Item.Value = SQL
        

    End If
    
        
    Record.AddItem (CStr(miRsAux!codmacta))
    Record.AddItem DBLet(miRsAux!Nommacta, "T")
    
    
    
    
    
    Set Item = Record.AddItem("")
    Item.Format = "%.2f"
    Item.Value = CCur(miRsAux!ImpVenci)
    Item.Caption = Format(Item.Value, FormatoImporte)
    
    
    
    
    
    Record.AddItem (CStr(miRsAux!numorden))
    
    
    If Rojo Then
        For i = 1 To Record.ItemCount
            Record.Item(i).ForeColor = vbRed
        Next
    End If
    
    
    'Adds the PreviewText to the Record.  PreviewText is the text displayed for the ReportRecord while in PreviewMode
  ' Record.PreviewText = "ID: " & miRsAux!CodClien
    
    
    
    
    
    
eAddRecord2:
    
End Sub






Private Sub wndReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    
    
        
        SQL = "numserie='" & Row.Record.Item(1).Value & "' AND numfactu = " & Row.Record.Item(2).Value & " AND fecfactu='" & Format(Row.Record.Item(3).Value, FormatoFecha) & "' AND numorden = "
        
        If Me.Option1(1).Value Then
            frmAseguradosAccion.Opcion = 1
            frmAseguradosAccion.Label1.Caption = Row.Record.Item(8).Value
            frmAseguradosAccion.lblTitulo = Me.Option1(1).Caption
            SQL = SQL & Row.Record.Item(10).Value
        Else
            frmAseguradosAccion.Opcion = 0
            frmAseguradosAccion.Label1.Caption = Row.Record.Item(7).Value
            frmAseguradosAccion.lblTitulo = Me.Option1(0).Caption
            SQL = SQL & Row.Record.Item(9).Value
        End If
        frmAseguradosAccion.Label2.Caption = Trim(Row.Record.Item(1).Value & Row.Record.Item(2).Value) & "  de " & Row.Record.Item(3).Value & "    Vto: " & Row.Record.Item(IIf(Option1(1).Value, 9, 8)).Value
        CadenaDesdeOtroForm = ""
        frmAseguradosAccion.SQLVto = SQL
        frmAseguradosAccion.Show vbModal
        
        If CadenaDesdeOtroForm <> "" Then
            'Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
            wndReportControl.RemoveRowEx Row
            wndReportControl.Populate
        End If
        
    

End Sub
