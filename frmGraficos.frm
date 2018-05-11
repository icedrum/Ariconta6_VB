VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Object = "{C215CB9A-0AE1-499F-A101-48B3C370D3DF}#17.2#0"; "Codejock.ChartPro.v17.2.0.ocx"
Begin VB.Form frmGraficos 
   Caption         =   "Graficos"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15450
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   15450
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameOpciones 
      Height          =   1815
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   10815
      Begin VB.CommandButton cmdIMprimir 
         Height          =   375
         Index           =   1
         Left            =   2640
         Picture         =   "frmGraficos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir"
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox cboPeriodo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Muestra puntos"
         Height          =   255
         Left            =   6240
         TabIndex        =   12
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkShowMarkers 
         Caption         =   "Muestra puntos"
         Height          =   255
         Left            =   8280
         TabIndex        =   11
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkShowLabels 
         Caption         =   "Muestra etiquetas"
         Height          =   255
         Left            =   6240
         TabIndex        =   10
         Top             =   840
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdIMprimir 
         Height          =   375
         Index           =   0
         Left            =   3120
         Picture         =   "frmGraficos.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Guardar como png"
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox cmbPalette 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox cmbAppearance 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin XtremeSuiteControls.ListBox ListBox1 
         Height          =   1455
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   2055
         _Version        =   1114114
         _ExtentX        =   3625
         _ExtentY        =   2566
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Style           =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   14
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3735
      End
      Begin VB.Label lblPalette 
         BackStyle       =   0  'Transparent
         Caption         =   "Tono:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblAppearance 
         BackStyle       =   0  'Transparent
         Caption         =   "Apariencia:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
   End
   Begin XtremeChartControl.ChartControl ChartControl1 
      Height          =   4335
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   5655
      _Version        =   1114114
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   0
   End
   Begin MSComctlLib.TreeView wndTreeView 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   4
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmGraficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim ActiveView As VBControlExtender


Dim PaletteList
Dim PaletteIndex As Long
Dim AppearanceList
Dim AppearanceIndex As Long


Dim Diagram As ChartDiagram2D

'me.tag :  levvara el filtro leido
Dim F As Date
Dim SQL As String
Dim CargandoGrafico As Boolean
Dim Importe As Currency





Dim PrimVez As Boolean

Private Sub cboPeriodo_Click()
    If CargandoGrafico Then Exit Sub
    CargarGraficoSeleccionado
End Sub

Private Sub chkShowLabels_Click()

On Error Resume Next:
    
    Dim X As ChartSeries
    For Each X In ChartControl1.Content.Series
        X.Style.Label.visible = chkShowLabels.Value
    Next
    

End Sub

Private Sub chkShowMarkers_Click()

On Error Resume Next:
    
    Dim X As ChartSeries
    For Each X In ChartControl1.Content.Series
        X.Style.Marker.visible = chkShowMarkers.Value
    Next
    

End Sub

Private Sub chkVarios_Click()
    If CargandoGrafico Then Exit Sub
    
    CargarGraficoSeleccionado
    
End Sub

Private Sub cmbAppearance_Click()
        AppearanceIndex = cmbAppearance.ListIndex
        ChartControl1.Content.Appearance.SetAppearance AppearanceList(AppearanceIndex)
End Sub

Private Sub cmbPalette_Click()
        PaletteIndex = cmbPalette.ListIndex
        ChartControl1.Content.Appearance.SetPalette PaletteList(PaletteIndex)
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    If CargandoGrafico Then Exit Sub

    If Index = 1 Then
        If ChartControl1.PrintPreview Then
            Me.ChartControl1.PrintChart 0

        End If
        
        
    Else
        GuardarComo
    End If
End Sub
Private Sub GuardarComo()
Dim C As String
Dim N As Integer
    
    On Error GoTo E1
    
    
    
    
    With frmppal.cd1
        .DialogTitle = "Save As"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "*.png"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
    End With
        
    C = frmppal.cd1.FileName
    If Mid(C, Len(C) - 3, 1) <> "." Then C = C + ".png"
    
    
    If Dir(C, vbArchive) <> "" Then
        If MsgBox("Ya existe el archivo  " & C & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
   End If
   
   Importe = Me.ChartControl1.Height / ChartControl1.Width
   N = Round(1000 * Importe, 0)
   
   Me.ChartControl1.SaveAsImage C, 1000, N   'CLng(ChartControl1.Width / 3), CLng(ChartControl1.Height / 3)
    
   Exit Sub
E1:
    MuestraError Err.Number, Err.Description
End Sub
Private Sub Form_Activate()
    If PrimVez Then
        PrimVez = False
        'Haremos click  en el primer nodo
        
        
    
        AjustaParaGraficos
        Me.Refresh
        Set wndTreeView.SelectedItem = wndTreeView.Nodes(2)
        wndTreeView_NodeClick wndTreeView.SelectedItem
        
    End If
End Sub


Private Sub AjustaParaGraficos()
    ChartControl1.Content.Series.DeleteAll
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl1.Content.EnableMarkup = True
    ChartControl1.Content.Titles.DeleteAll
    SQL = "Leyendo"
    ChartControl1.Content.Titles.Add SQL
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
    
    
    
   ' Dim Series As ChartSeries
   ' Set Series = ChartControl1.Content.Series.Add("Ingresos")
        
   ' Set ChartControl1.Content.Series(0).Style = New ChartBarSeriesStyle
   '
   ' Dim Diagram As ChartDiagram2D
   ' Set Diagram = ChartControl1.Content.Series(0).Diagram
   ' Diagram.AxisY.Title.visible = True
   ' Diagram.AxisY.Title.Text = "€uros"
   ' Diagram.AxisX.Title.visible = True
   ' Diagram.AxisX.Title.Text = "Ejercicio"
   '
   ' Diagram.AxisX.Label.Angle = 360

End Sub

Private Sub Form_Load()
    
    Me.Icon = frmppal.Icon
    PrimVez = True
    
    wndTreeView.Tag = ""
    
    AppearanceList = Array("Nature", "Black", "Gray")
    
    
    
    PaletteList = Array("Victorian", "Vibrant Pastel", "Vibrant", "Tropical", "Summer", "Spring Time", "Rainbow", "Purple", "Primary Colors", _
        "Postmodern", "Photodesign", "Pastel", "Office", "Orange Green", "Nature", "Natural", "Impresionism", "Illustration", "Harvest", _
        "Green Brown", "Green Blue", "Green", "Gray", "Four Color", "Fire", "Earth Tone", "Danville", "Caribbean", "Cappuccino", "Blue Gray", "Blue")

    
    
    
    
    For i = 0 To 2
        cmbAppearance.AddItem RecuperaValor("Natural|Negro|Gris|", i + 1), i 'AppearanceList(i)
    Next i
    For i = 0 To 30
        cmbPalette.AddItem PaletteList(i), i
    Next i
    
    CargandoGrafico = True 'Para que no lo carge
    ValorPorDefecto True
    cmbPalette.ListIndex = PaletteIndex
    cmbAppearance.ListIndex = AppearanceIndex
    
    CargaTree
    J = 0 'En el periodo diremeos si carga el actual o el siguiente
    CargaComboPeriodo
       
    ListBox1.Clear
    F = vParam.fechaini
    F = DateAdd("yyyy", 1, F)
    NumRegElim = 0
    
    Msg = ""
    Do
        If Year(vParam.fechaini) = Year(vParam.fechafin) Then
            SQL = Year(F)
        Else
            SQL = Year(F) & " - " & (Year(F) Mod 2000) + 1
        End If
        ListBox1.AddItem SQL
        ListBox1.ItemData(NumRegElim) = Year(F)
        If ListBox1.ListCount = 3 Then
            
            'Actual, o el sugiente
            ListBox1.Checked(J) = True
        End If
        NumRegElim = NumRegElim + 1
        F = DateAdd("yyyy", -1, F)
     Loop Until F < vParam.FechaMinimaEnHlinapu
    
    
    CargandoGrafico = False
    
    
    
End Sub

Private Sub ValorPorDefecto(Leer As Boolean)
Dim B As Byte
    If Leer Then
        AppearanceIndex = 0
        PaletteIndex = 9
        SQL = TextoValueLeer("frmgrafico")
        Me.Tag = SQL
        If SQL <> "" Then
            Msg = RecuperaValor(SQL, 1)
            If Msg <> "" Then
                If Val(Msg) > 2 Then Msg = "2"
            End If
            AppearanceIndex = Val(Msg)
            
            Msg = RecuperaValor(SQL, 2)
            If Msg <> "" Then
                If Val(Msg) > Me.cmbPalette.ListCount Then Msg = cmbPalette.ListCount
            End If
            PaletteIndex = Val(Msg)
            
        End If
    Else
        SQL = AppearanceIndex & "|" & PaletteIndex & "|"
        If SQL <> Me.Tag Then TextoValueGuardar "frmgrafico", SQL
    End If
End Sub



Private Sub CargaTree()

 Dim Group As Node, Item As Node
    
    Set Group = wndTreeView.Nodes.Add(, , , "Explotacion")
    Group.Bold = True
    Group.Expanded = True
    
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "G1", "Cuenta explotacion")
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "G2", "Cta explotacion mes")
    'Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "G3", "Ingresos por subcuenta")
    
    
    Set Group = wndTreeView.Nodes.Add(, , , "Tesoreria")
    Group.Bold = True
    Group.Expanded = True
    
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "T1", "Cobros")
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "T2", "Pagos")

    Set Group = wndTreeView.Nodes.Add(, , , "Clientes")
    Group.Bold = True
    Group.Expanded = True
    
    
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "C1", "Facturacion")
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "C2", "Top 10 ")
    
    Set Group = wndTreeView.Nodes.Add(, , , "Proveedores")
    Group.Bold = True
    Group.Expanded = True
    
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "P1", "Facturacion")
    Set Item = wndTreeView.Nodes.Add(Group, tvwChild, "P2", "Top 10")
    
    
            

End Sub

Private Sub CargaComboPeriodo()
Dim F As Date
Dim F2 As Date

    Me.cboPeriodo.Clear
    cboPeriodo.AddItem "Ejercicio"          '0
    cboPeriodo.AddItem "1er Trimestre "     '1
    cboPeriodo.AddItem "2º Trimestre "      '2
    cboPeriodo.AddItem "3er Trimestre "     '3
    cboPeriodo.AddItem "4º Trimestre "      '4
    For i = 1 To 12
        SQL = "01/" & Format(i, "00") & "/" & Year(Now)    'Enero 5 febrero 6.....
        cboPeriodo.AddItem Format(SQL, "mmmm")
    Next
        
    F = Now
   
    SQL = ""
  

    i = Month(F)
    If i >= 4 And i <= 6 Then
        SQL = "04"
    Else
        If i >= 7 And i <= 9 Then
            SQL = "07"
        Else
            If i >= 10 Then
                SQL = "10"
            Else
                SQL = "01"
            End If
        End If
    End If
    SQL = "01/" & SQL & "/" & Year(F)
    F = CDate(SQL)
    F2 = F
    F = DateAdd("d", -1, F)
    
    'Fecha a ubicar
    
    If F > vParam.fechafin Then
        J = 0
        F2 = DateAdd("yyyy", 1, vParam.fechaini)
    Else
        J = 1
        If F < vParam.fechaini Then
            J = 2
            F2 = DateAdd("yyyy", -1, vParam.fechaini)
         End If
    End If
    
    'Veremos el trimestre
    i = 0
    If F2 > F Then
        'No deberia pasar nunca
        i = 1
    Else
        Do
            i = i + 1
            F2 = DateAdd("m", 3, F2)
            
        Loop Until F2 > F
    End If
    cboPeriodo.ListIndex = i
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    FrameOpciones.Move 0, 0, Me.ScaleWidth, FrameOpciones.Height
    wndTreeView.Move 5, 5 + FrameOpciones.Height, 3000, Me.ScaleHeight - 10 - FrameOpciones.Height
    
    ChartControl1.Move wndTreeView.Width + 10, 5 + FrameOpciones.Height, Me.ScaleWidth - wndTreeView.Width - 15, Me.ScaleHeight - 10 - FrameOpciones.Height

End Sub



'  *** Cta explotacion
Private Sub GraficoCtaExplotacion()

    lblTitle.Tag = "Cuenta explotacion"
    
    
    ChartControl1.Content.Series.DeleteAll
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl1.Content.EnableMarkup = True
    
    
    AddBarSeriesCtaExplotacion
    
    Set ChartControl1.Content.Series(0).Style = New ChartBarSeriesStyle
    Set ChartControl1.Content.Series(1).Style = New ChartBarSeriesStyle
    Set ChartControl1.Content.Series(2).Style = New ChartBarSeriesStyle
    
    Dim Diagram As ChartDiagram2D
    Set Diagram = ChartControl1.Content.Series(0).Diagram
    Diagram.AxisY.Title.visible = True
    Diagram.AxisY.Title.Text = "miles €uros"
    Diagram.AxisX.Title.visible = True
    Diagram.AxisX.Title.Text = "Ejercicio"
    
    Diagram.AxisX.Label.Angle = 360

End Sub
Sub AddBarSeriesCtaExplotacion()
   
    If ChartControl1.Content.Series.Count > 0 Then
        ChartControl1.Content.Series.DeleteAll
    End If
        
    ChartControl1.Content.Titles.DeleteAll
    SQL = "Resultado "
    If Me.cboPeriodo.ListIndex > 0 Then SQL = SQL & cboPeriodo.Text
    ChartControl1.Content.Titles.Add SQL
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
    
    SQL = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Checked(i) Then SQL = SQL & "X"
    Next
    'No hay ninguno seleccionado
    If SQL = "" Then ListBox1.Checked(0) = True
    
    
    
    
    Dim Series As ChartSeries
    
    
    
    Set miRsAux = New ADODB.Recordset
    For J = 1 To 2
        If J = 1 Then
             Set Series = ChartControl1.Content.Series.Add("Ingresos")
        Else
             Set Series = ChartControl1.Content.Series.Add("Gastos")
        End If
        For i = ListBox1.ListCount - 1 To 0 Step -1
         
            If ListBox1.Checked(i) Then
                                    
                Importe = 0
                SQL = "Select sum(coalesce(timported,0)-coalesce(timporteh,0)) from hlinapu where numdiari>=0 "
                SQL = SQL & " AND fechaent " & FijarFechasEjercicios(i)
                SQL = SQL & " AND codconce<>960  AND codmacta like '" & IIf(J = 2, "6", "7") & "%'"
                
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not miRsAux.EOF Then Importe = DBLet(miRsAux.Fields(0), "N")
                
                miRsAux.Close
                If J = 1 Then Importe = -1 * Importe
                Series.Points.Add ListBox1.List(i), Round(Importe / 1000, 2)
                
               
            End If
        Next i
    Next J
    
    
    Set Series = ChartControl1.Content.Series.Add("Resultado")
    For i = 0 To ChartControl1.Content.Series(0).Points.Count - 1 'lA que acabo de crear no la cuento
        Importe = CCur(ChartControl1.Content.Series(0).Points(i).Value(0)) - CCur(ChartControl1.Content.Series(1).Points(i).Value(0))
        SQL = ChartControl1.Content.Series(0).Points(i).ArgumentValueString
        
        Series.Points.Add SQL, Importe
    Next i
    
End Sub

Private Function MesesPeriodo() As Byte
    If Me.cboPeriodo.ListIndex <= 0 Then
        MesesPeriodo = 12
    Else
        If Me.cboPeriodo.ListIndex <= 4 Then
            MesesPeriodo = 3
        Else
            MesesPeriodo = 1
        End If
    End If
    
    
    
    
    
End Function

Private Function NumeroInicioPeridodPorMeses(ItemSeleccionado) As Long
Dim F  As Date
Dim Cad As String
Dim H As Integer
     If Me.cboPeriodo.ListIndex <= 0 Then
        F = Format(vParam.fechaini, "dd/mm/") & ListBox1.ItemData(ItemSeleccionado)
        
    Else
        If Me.cboPeriodo.ListIndex <= 4 Then
            'Si año partidos
             Cad = ListBox1.ItemData(ItemSeleccionado)
            F = Format(vParam.fechaini, "dd/mm/") & Cad
            H = cboPeriodo.ListIndex - 1
            H = H * 3
            F = DateAdd("m", H, F)
           
        Else
            'mes
             Cad = ListBox1.ItemData(ItemSeleccionado)
             H = cboPeriodo.ListIndex - 4
             
             'I=mes a tratar
            If Year(vParam.fechaini) <> Year(vParam.fechafin) Then
                If H < Month(vParam.fechaini) Then Cad = Val(Cad) + 1
            End If
            F = CDate("01/" & Format(H, "00") & "/" & Cad)
            
            
           
        End If
    End If
    NumeroInicioPeridodPorMeses = CLng(Year(F) & Format(Month(F), "00"))
End Function



'----------------------------------------------
'
'    Cta explotacion mensual
Private Sub GraficoCtaExplotacionMensual()
Dim ItemSeleccionado As Integer
Dim N As Byte
Dim Meses As Byte

    
    lblTitle.Tag = "Cta explotacion mensual"
    
    ChartControl1.Content.Series.DeleteAll
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl1.Content.EnableMarkup = True
    '-------------------------------------------
    
    
    
       
    If ChartControl1.Content.Series.Count > 0 Then
        ChartControl1.Content.Series.DeleteAll
    End If
        
    ChartControl1.Content.Titles.DeleteAll
    SQL = "Explotacion mensual "
    If Me.cboPeriodo.ListIndex > 0 Then SQL = SQL & cboPeriodo.Text & " "
    ChartControl1.Content.Titles.Add SQL
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
    
    SQL = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Checked(i) Then
            SQL = SQL & "X"
            If Len(SQL) = 1 Then ChartControl1.Content.Titles(0) = ChartControl1.Content.Titles(0) & ListBox1.List(i)
            ItemSeleccionado = i
        End If
    Next
    'No hay ninguno seleccionado
    If SQL = "" Then ListBox1.Checked(0) = True
    
    
    
    Dim Series As ChartSeries
    
    
    
    Set miRsAux = New ADODB.Recordset
    For J = 1 To 2
        If J = 1 Then
             Set Series = ChartControl1.Content.Series.Add("Ingresos")
        Else
             Set Series = ChartControl1.Content.Series.Add("Gastos")
        End If
        
            
        
                                    
        Importe = 0
        SQL = "Select year(fechaent) ano,month(fechaent) mes, sum(coalesce(timported,0)-coalesce(timporteh,0)) importe from hlinapu where numdiari>=0 "
        SQL = SQL & " AND fechaent " & FijarFechasEjercicios(ItemSeleccionado)
        SQL = SQL & " AND codconce<>960  AND codmacta like '" & IIf(J = 2, "6", "7") & "%'"
        SQL = SQL & " group by 1,2 ORDER BY  1,2"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Meses = MesesPeriodo
        NumRegElim = NumeroInicioPeridodPorMeses(ItemSeleccionado)
        For N = 1 To Meses
            
            If miRsAux.EOF Then
                Importe = 0
            
            Else
                K = CLng(CStr(miRsAux!Ano) & Format(miRsAux!Mes, "00"))
                If K = NumRegElim Then
                    'Perfecto, tiene valor ese mes
                    Importe = miRsAux!Importe
                    miRsAux.MoveNext
                Else
                    Importe = 0
                End If
            End If
            If J = 1 Then Importe = -1 * Importe
            SQL = Mid(MonthName(Right(NumRegElim, 2)), 1, 3) & " " & Val(Left(NumRegElim, 4)) - 2000
            Series.Points.Add SQL, Round(Importe / 1000, 2)
            
            
            'Movemos siguiente
            If Val(Right(CStr(NumRegElim), 2)) = 12 Then
                NumRegElim = Val(Left(CStr(NumRegElim), 4)) + 1 & "01"
            Else
                NumRegElim = NumRegElim + 1
            End If
        
    
           
        Next N
        miRsAux.Close
        


    Next J
    
    
    Set Series = ChartControl1.Content.Series.Add("Resultado")
    For i = 0 To ChartControl1.Content.Series(0).Points.Count - 1 'lA que acabo de crear no la cuento
        Importe = CCur(ChartControl1.Content.Series(0).Points(i).Value(0)) - CCur(ChartControl1.Content.Series(1).Points(i).Value(0))
        SQL = ChartControl1.Content.Series(0).Points(i).ArgumentValueString
        
        Series.Points.Add SQL, Importe
    Next i
    

    
    
    
    
    
    '---------------------------------------------
    ' trozo  final grafico
    Set ChartControl1.Content.Series(0).Style = New ChartBarSeriesStyle
    Set ChartControl1.Content.Series(1).Style = New ChartBarSeriesStyle
    Set ChartControl1.Content.Series(2).Style = New ChartBarSeriesStyle
    
    Dim Diagram As ChartDiagram2D
    Set Diagram = ChartControl1.Content.Series(0).Diagram
    Diagram.AxisY.Title.visible = True
    Diagram.AxisY.Title.Text = "miles €uros"
    Diagram.AxisX.Title.visible = True
    Diagram.AxisX.Title.Text = "Ejercicio"
    
    Diagram.AxisX.Label.Angle = 360

End Sub




'----------------------------------------------
'
'    Cta explotacion mensual
Private Sub GraficoFacturacion(Cliente As Boolean)
Dim N As Byte
Dim Imp2 As Currency
Dim Meses As Byte
    lblTitle.Tag = "Facturacion " & IIf(Cliente, "cliente", "proveedor")
    
    ChartControl1.Content.Series.DeleteAll
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl1.Content.EnableMarkup = True
    '-------------------------------------------
    
    
    
       
    If ChartControl1.Content.Series.Count > 0 Then
        ChartControl1.Content.Series.DeleteAll
    End If
        
    ChartControl1.Content.Titles.DeleteAll
    SQL = "Facturacion mensual " & IIf(Me.chkVarios.Value = 1, "acumulada", "")
    If cboPeriodo.ListIndex > 0 Then SQL = SQL & " " & Me.cboPeriodo.Text
    ChartControl1.Content.Titles.Add SQL
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
    
    SQL = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Checked(i) Then SQL = SQL & "X"
    Next
    'No hay ninguno seleccionado
    If SQL = "" Then ListBox1.Checked(0) = True
    
    
            
    Dim Series As ChartSeries
    
    
    
    Set miRsAux = New ADODB.Recordset
        
        
    For i = 0 To ListBox1.ListCount - 1
         
        If ListBox1.Checked(i) Then
            
            Importe = 0
            If Cliente Then
                SQL = "Select year(fecfactu) ano,month(fecfactu) mes, sum(totbases-coalesce(trefaccl,0)) "
            Else
                SQL = "Select year(fecharec) ano,month(fecharec) mes, sum(totbases-coalesce(trefacpr,0)) "
            End If
            SQL = SQL & " importe FROM " & IIf(Cliente, "factcli", "factpro") & " where numserie<>'' "
            SQL = SQL & " AND " & IIf(Cliente, "fecfactu", "fecharec") & FijarFechasEjercicios(i)
            SQL = SQL & " group by 1,2 ORDER BY  1,2"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Set Series = ChartControl1.Content.Series.Add(ListBox1.List(i))
            NumRegElim = CLng(ListBox1.ItemData(i) & Format(Month(vParam.fechaini), "00"))
            Imp2 = 0
            Meses = MesesPeriodo
            NumRegElim = NumeroInicioPeridodPorMeses(i)
            For N = 1 To Meses
                If miRsAux.EOF Then
                    Importe = 0
                
                Else
                    K = CLng(CStr(miRsAux!Ano) & Format(miRsAux!Mes, "00"))
                    If K = NumRegElim Then
                        'Perfecto, tiene valor ese mes
                        Importe = miRsAux!Importe
                        miRsAux.MoveNext
                    Else
                        Importe = 0
                    End If
                End If
                Importe = Round(Importe / 1000, 2)
                If Me.chkVarios.Value = 1 Then
                    'Acumulado
                    Imp2 = Imp2 + Importe
                Else
                    'Mes a mes
                    Imp2 = Importe
                End If
                SQL = Mid(MonthName(Right(NumRegElim, 2)), 1, 3) '& " " & Val(Left(NumRegElim, 4)) - 2000
                Series.Points.Add SQL, Imp2
                
                
                'Movemos siguiente
                If Val(Right(CStr(NumRegElim), 2)) = 12 Then
                    NumRegElim = Val(Left(CStr(NumRegElim), 4)) + 1 & "01"
                Else
                    NumRegElim = NumRegElim + 1
                End If
            
        
               
            Next N
            miRsAux.Close
            
        End If
    Next

    For i = 0 To ChartControl1.Content.Series.Count - 1
        Set ChartControl1.Content.Series(i).Style = New ChartLineSeriesStyle
    Next
   
        'x.Style.Marker.visible = chkShowMarkers.Value
    
    Dim Diagram As ChartDiagram2D
    Set Diagram = ChartControl1.Content.Series(0).Diagram
    Diagram.AxisY.Title.visible = True
    Diagram.AxisY.Title.Text = "miles €uros"
    Diagram.AxisX.Title.visible = True
    Diagram.AxisX.Title.Text = "Ejercicio"
    
    Diagram.AxisX.Label.Angle = 360
    If Me.chkShowLabels.Value = 0 Then chkShowLabels_Click
End Sub



Private Sub VolumenFacturacion(Cliente As Boolean)
Dim Imp2 As Currency
Dim itemSel As Integer
    Me.lblTitle.Tag = "Top 10 " & IIf(Cliente, "Venta", "Compra")
    
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl1.Content.EnableMarkup = True
       


    If ChartControl1.Content.Series.Count > 0 Then
        ChartControl1.Content.Series.DeleteAll
    End If
    
    ChartControl1.Content.Titles.DeleteAll
    SQL = "Top 10 " & IIf(Cliente, "Venta", "Compra") & " " & Me.cboPeriodo.Text
    ChartControl1.Content.Titles.Add SQL
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
   
    Dim Style As ChartPieSeriesStyle
    
    Dim Series As ChartSeries
    
    Set Style = New ChartPieSeriesStyle
       
    Set Series = ChartControl1.Content.Series.Add(SQL)
    
    itemSel = -1
    SQL = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Checked(i) Then
            itemSel = i
            SQL = ListBox1.List(i)
            Exit For
        End If
    Next i
    
    If itemSel < 0 Then Exit Sub
    
    ChartControl1.Content.Titles(0).Text = ChartControl1.Content.Titles(0).Text & " " & SQL & " miles€"
   
    
    
    Set miRsAux = New ADODB.Recordset
    If Cliente Then
        SQL = "Select codmacta, sum(totfaccl) importe from factcli "
    Else
        SQL = "Select codmacta, sum(totfacpr) importe from factpro "
    End If
    SQL = SQL & " where numserie<>''  AND " & IIf(Cliente, "fecfactu", "fecharec")
    SQL = SQL & FijarFechasEjercicios(itemSel)
    SQL = SQL & " group by 1 ORDER BY  2 desc"
           
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockReadOnly, adCmdText
    Importe = 0
    J = 0
    If Not miRsAux.EOF Then
    
        
    
    
        While Not miRsAux.EOF
            J = J + 1
            Importe = Importe + miRsAux!Importe
            
            If J = 10 Then
                 miRsAux.MoveLast
            End If
            miRsAux.MoveNext
            
        Wend
        miRsAux.MoveFirst
    End If
    J = -1
    NumRegElim = 0
    
    
    For i = 1 To 10
        If Not miRsAux.EOF Then
            SQL = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux!codmacta, "T")
            SQL = Mid(SQL, 1, 10)
            Imp2 = miRsAux!Importe / Importe
            Imp2 = Round2(Imp2 * 360, 2)
            
    
            If J < 0 Then
                'Es el primero
                NumRegElim = miRsAux!Importe
            End If
            J = miRsAux!Importe \ 1000
            CreatePiePoint Series.Points, SQL, Round(miRsAux!Importe / 1000, 2)
            miRsAux.MoveNext
        End If
        
    Next
'        CreateBubblePoint Series.Points, "California", 8, 37, 11.95

    miRsAux.Close
    Set miRsAux = Nothing
                            
    Series.ArgumentScaleType = xtpChartScaleQualitative

    NumRegElim = NumRegElim \ 1000
    

    '--------------------------------------------------------------------------------------
    Set ChartControl1.Content.Series(0).Style = Style
    Style.Label.Format = "{V} m€"
               
    'Style.HolePercent = 40
    Style.Rotation = 30
    Style.Label.Antialiasing = False
    
    
    
    

End Sub

    


'  Cobros pagos
Private Sub CobrosPagos(Cobros As Boolean)
Dim Imp2 As Currency
Dim itemSel As Integer
Dim MsgFechas As String

    Me.lblTitle.Tag = IIf(Cobros, "Cobros", "Pagos")
    
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl1.Content.EnableMarkup = True
       


    If ChartControl1.Content.Series.Count > 0 Then
        ChartControl1.Content.Series.DeleteAll
    End If
    
    ChartControl1.Content.Titles.DeleteAll
    SQL = IIf(Cobros, "Cobros", "Pagos") & " " & Me.cboPeriodo.Text

    ChartControl1.Content.Titles.Add SQL & " miles€"
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
   
    Dim Style As ChartPieSeriesStyle
    
    Dim Series As ChartSeries
    
    Set Style = New ChartPieSeriesStyle
       
    Set Series = ChartControl1.Content.Series.Add(SQL)
    
    itemSel = -1
    SQL = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Checked(i) Then
            itemSel = i
            SQL = ListBox1.List(i)
            Exit For
        End If
    Next i
    
    If itemSel < 0 Then Exit Sub
    
    If cboPeriodo.ListIndex = 0 Then ChartControl1.Content.Titles(0).Text = ChartControl1.Content.Titles(0).Text & " " & SQL
    
    
    
    Set miRsAux = New ADODB.Recordset
    
    MsgFechas = FijarFechasEjercicios(itemSel)
    i = InStrRev(MsgFechas, " AND ")
    If i = 0 Then Err.Raise 513, , "Imposible obtener fecha fin periodo"
    MsgFechas = Mid(MsgFechas, i + 5)
    

    
    If Cobros Then
        SQL = "select if(fecultco is  null,0,if( fecultco<=" & MsgFechas & ",1,2 )) , sum(impvenci+coalesce(gastos,0)) importe FROM cobros "
    Else
        SQL = "select if(fecultpa is  null,0,if( fecultpa<=" & MsgFechas & ",1,2 )) , sum(impefect) importe FROM pagos "
    End If
    SQL = SQL & " where numserie<>''  AND " & IIf(Cobros, "fecvenci", "fecefect")
    SQL = SQL & FijarFechasEjercicios(itemSel)
    SQL = SQL & " group by 1 ORDER BY  1"
               
    miRsAux.Open SQL, Conn, adOpenKeyset, adLockReadOnly, adCmdText
    Importe = 0
    J = 0

    J = -1
    NumRegElim = 0
    
    
   
        While Not miRsAux.EOF
            If Cobros Then
                SQL = RecuperaValor("Pendiente|Cobrado|Fuera plazo|", miRsAux.Fields(0) + 1)
            Else
                SQL = RecuperaValor("Pendiente|Pagado|Fuera plazo|", miRsAux.Fields(0) + 1)
            End If
            SQL = Mid(SQL, 1, 10)
            
            CreatePiePoint Series.Points, SQL, Round(miRsAux!Importe / 1000, 2)
            miRsAux.MoveNext
        Wend
        
  
'        CreateBubblePoint Series.Points, "California", 8, 37, 11.95

    miRsAux.Close
    Set miRsAux = Nothing
                            
    Series.ArgumentScaleType = xtpChartScaleQualitative

    NumRegElim = NumRegElim \ 1000
    

    '--------------------------------------------------------------------------------------
    Set ChartControl1.Content.Series(0).Style = Style
    Style.Label.Format = "{V} miles€"
               
    'Style.HolePercent = 40
    Style.Rotation = 30
    Style.Label.Antialiasing = False
    
    
End Sub










'** Segundo

Private Sub Grafic2()
    
    Me.lblTitle.Tag = "Prueba"
    
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    ChartControl1.Content.EnableMarkup = True
       
    AddBubbleSeries
    
    Dim BubbleSeriesStyle As New ChartBubbleSeriesStyle
    BubbleSeriesStyle.MinSize = 1.5
    BubbleSeriesStyle.MaxSize = 3.5
    
    Set ChartControl1.Content.Series(0).Style = BubbleSeriesStyle
    BubbleSeriesStyle.Label.Format = "{V} %%"
    
    Dim Diagram As ChartDiagram2D
    Set Diagram = ChartControl1.Content.Series(0).Diagram
    Diagram.AxisY.Title.Text = "Population in Millions"
    Diagram.AxisY.Title.visible = True
    Diagram.AxisY.Range.AutoRange = False
    Diagram.AxisY.Range.MaxValue = 50
    Diagram.AxisY.Range.MinValue = 0
    Diagram.AxisX.Title.visible = False
    Diagram.AxisX.visible = True
    Diagram.AxisX.Reversed = True
    Diagram.AxisX.TickMarks.visible = False
    Diagram.AxisX.Label.visible = False
    'Diagram.AxisY.Range.ShowZeroLevel = False
    
    'txtMinBubbleSize.Text = BubbleSeriesStyle.MinSize
    'txtMaxBubbleSize.Text = BubbleSeriesStyle.MaxSize
End Sub

Sub AddBubbleSeries()
    
    ChartControl1.Content.Titles.Add "Bubble Syle"

    If ChartControl1.Content.Series.Count > 0 Then
        ChartControl1.Content.Series.DeleteAll
    End If
    
    ChartControl1.Content.Titles.DeleteAll
    ChartControl1.Content.Titles.Add "Top 10 States by Population"
    ChartControl1.Content.Legend.visible = True
    ChartControl1.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
   
    Dim Series As ChartSeries
    Set Series = ChartControl1.Content.Series.Add("Top 10 States by Population")
    
    CreateBubblePoint Series.Points, "California", 8, 37, 11.95
    CreateBubblePoint Series.Points, "Texas", 7, 25, 7.81
    CreateBubblePoint Series.Points, "New York", 6, 20, 6.31
    CreateBubblePoint Series.Points, "Florida", 5, 18, 5.97
    CreateBubblePoint Series.Points, "Illinois", 9, 13, 4.2
    CreateBubblePoint Series.Points, "Pennsylvania", 10, 12, 4.06
    CreateBubblePoint Series.Points, "Ohio", 4, 11, 3.75
    CreateBubblePoint Series.Points, "Michigan", 3, 10, 3.29
    CreateBubblePoint Series.Points, "Georgia", 2, 9, 3.12
    CreateBubblePoint Series.Points, "North Carolina", 1, 9, 3
                            
    Series.ArgumentScaleType = xtpChartScaleQualitative
           
End Sub

Sub CreateBubblePoint(ByVal pPointCollection As ChartSeriesPointCollection, lpszLegendText As String, nYear As Integer, nValue As Integer, dWidth As Double)
    Dim pPoint As ChartSeriesPoint
    Set pPoint = pPointCollection.Add2(nYear, nValue, dWidth)
    pPoint.LegendText = lpszLegendText
End Sub

Sub CreatePiePoint(ByVal pPointCollection As ChartSeriesPointCollection, Argument As String, nValue As Variant)
    Dim pPoint As ChartSeriesPoint
    Set pPoint = pPointCollection.Add(Argument, nValue)
    
End Sub



'El texto lo pondremos en tag
'Cuando sea leyendo ponemos leyendo BD
'Si no, ponemos el TAG
Private Sub PonLbl(LeyendoBD As Boolean)
    If LeyendoBD Then
        Screen.MousePointer = vbHourglass
        lblTitle.ForeColor = vbRed
        lblTitle.FontSize = 22
        lblTitle.Caption = "Leyendo BD"
        lblTitle.Refresh
        CargandoGrafico = True
    Else
        lblTitle.ForeColor = vbBlack
        lblTitle.Caption = lblTitle.Tag
        If Len(lblTitle.Caption) > 15 Then
            lblTitle.FontSize = 15
        Else
            lblTitle.FontSize = 22
        End If
        Screen.MousePointer = vbDefault
        lblTitle.Refresh
        CargandoGrafico = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ValorPorDefecto False
End Sub

Private Function SoloPuedeUnCheckFn() As Boolean
    SoloPuedeUnCheckFn = False
    If wndTreeView.Tag = "G2" Then
        SoloPuedeUnCheckFn = True
    ElseIf wndTreeView.Tag = "C2" Or wndTreeView.Tag = "P2" Then
        SoloPuedeUnCheckFn = True
    ElseIf wndTreeView.Tag = "T1" Or wndTreeView.Tag = "T2" Then
        SoloPuedeUnCheckFn = True
    End If
End Function



Private Sub ListBox1_ItemCheck(ByVal Item As Long)
Dim SoloPuedeUnCheck As Boolean
    If CargandoGrafico Then Exit Sub
    
    
    'UNos graficos NO podra seleccionar mas de una año
    SoloPuedeUnCheck = SoloPuedeUnCheckFn
    
    
    
    If SoloPuedeUnCheck Then
        'Solo puede UNO
        If ListBox1.Checked(Item) Then
            'Todos menos el, sin seleccionar
            For i = 0 To ListBox1.ListCount - 1
                If Item <> i Then
                    ListBox1.Checked(i) = False
                End If
            Next
        Else
            'Dejo el primero que llege
            J = 0
            For i = 0 To ListBox1.ListCount - 1
                If J = 0 Then
                    If ListBox1.Checked(i) Then J = 1
                Else
                    ListBox1.Checked(i) = False
                End If
            Next
            
        End If
    End If
    
    CargarGraficoSeleccionado
End Sub

Private Sub wndTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    'El nodo "padre" no tiene grafica
    If Node.Parent Is Nothing Then Exit Sub
    If Me.wndTreeView.Tag = Node.Key Then Exit Sub
    wndTreeView.Tag = Node.Key
    
    
    chkVarios.visible = False
    If wndTreeView.Tag = "C1" Or wndTreeView.Tag = "P1" Then
        chkVarios.Value = 0
        chkVarios.Caption = "Acumulado"
        chkVarios.visible = True
    End If
    
    
    If SoloPuedeUnCheckFn Then
        CargandoGrafico = True
        SQL = ""
        For i = 0 To Me.ListBox1.ListCount - 1
            If SQL = "" Then
                If ListBox1.Checked(i) Then
                    SQL = "N"
                    ListBox1.Selected(i) = True
                End If
            Else
                ListBox1.Checked(i) = False
            End If
        Next
        CargandoGrafico = False
    End If
    CargarGraficoSeleccionado
End Sub
Private Sub CargarGraficoSeleccionado()
    PonLbl True
    
    
    
    Select Case wndTreeView.Tag
    Case "G1"
        GraficoCtaExplotacion
    
    Case "G2"
        GraficoCtaExplotacionMensual
    Case "C1", "P1"
         GraficoFacturacion wndTreeView.Tag = "C1"
         
    Case "C2", "P2"
        VolumenFacturacion wndTreeView.Tag = "C2"
        
    Case "T1", "T2"
        'Tesoreria
        CobrosPagos wndTreeView.Tag = "T1"
    Case Else
            Grafic2
    
    
    End Select
    PonLbl False

End Sub

Private Function FijarFechasEjercicios(Item) As String
Dim Aux As String
Dim Cad As String
Dim H As Integer
Dim F As Date
Dim AnyosPartidos As Boolean


    AnyosPartidos = False
    If Year(vParam.fechaini) <> Year(vParam.fechafin) Then AnyosPartidos = True

    If Me.cboPeriodo.ListIndex <= 0 Then
        Cad = ListBox1.ItemData(Item)
        Aux = " between '" & Cad & Format(vParam.fechaini, "-mm-dd")
        If AnyosPartidos Then Cad = Val(Cad) + 1
        Aux = Aux & "' AND '" & Cad & Format(vParam.fechafin, "-mm-dd") & "'"
    
    Else
        If Me.cboPeriodo.ListIndex < 5 Then
            'TRIMESTRE
            Cad = ListBox1.ItemData(Item)
            F = Format(vParam.fechaini, "dd/mm/") & Cad
            H = cboPeriodo.ListIndex - 1
            H = H * 3
            F = DateAdd("m", H, F)
            
            Aux = " between " & DBSet(F, "F")
            F = DateAdd("m", 3, F)
            F = DateAdd("d", -1, F)
            Aux = Aux & " AND " & DBSet(F, "F")
        Else
            'mes
             Cad = ListBox1.ItemData(Item)
             H = cboPeriodo.ListIndex - 4
             
             'I=mes a tratar
            If AnyosPartidos Then
                If H < Month(vParam.fechaini) Then Cad = Val(Cad) + 1
            End If
            F = CDate("01/" & Format(H, "00") & "/" & Cad)
            Aux = " between " & DBSet(F, "F")
            F = DateAdd("m", 1, F)
            F = DateAdd("d", -1, F)
            Aux = Aux & " AND " & DBSet(F, "F")
            
            
        End If
    End If
   FijarFechasEjercicios = Aux
End Function



