VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESInfSituacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   Icon            =   "frmTESInfSituacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   120
      TabIndex        =   17
      Top             =   30
      Width           =   6945
      Begin VB.TextBox txtFecha 
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
         Index           =   1
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   1140
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
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
         Index           =   0
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1140
         Width           =   1305
      End
      Begin VB.TextBox txtDias 
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
         Index           =   0
         Left            =   1350
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   2010
         Width           =   1275
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   6270
         TabIndex        =   21
         Top             =   270
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
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   4860
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   4140
         TabIndex        =   25
         Top             =   780
         Width           =   2280
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   300
         TabIndex        =   23
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   270
         TabIndex        =   22
         Top             =   750
         Width           =   2280
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   1050
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Días Anteriores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   1650
         Width           =   1890
      End
      Begin VB.Label lblFecha 
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
         Left            =   2580
         TabIndex        =   19
         Top             =   3630
         Width           =   4095
      End
      Begin VB.Label lblFecha1 
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
         Index           =   4
         Left            =   2580
         TabIndex        =   18
         Top             =   3990
         Width           =   4095
      End
   End
   Begin VB.Frame FrameTipoSalida 
      Caption         =   "Tipo de salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   150
      TabIndex        =   6
      Top             =   2790
      Width           =   6915
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Impresora"
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
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Archivo csv"
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1515
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "PDF"
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
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "eMail"
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
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtTipoSalida 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   720
         Width           =   3345
      End
      Begin VB.TextBox txtTipoSalida 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
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
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   9
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   8
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButtonImpr 
         Caption         =   "Propiedades"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5190
         TabIndex        =   7
         Top             =   720
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5850
      TabIndex        =   4
      Top             =   5670
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4260
      TabIndex        =   3
      Top             =   5670
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Left            =   1560
      TabIndex        =   24
      Top             =   5700
      Width           =   2520
   End
End
Attribute VB_Name = "frmTESInfSituacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 903

' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
Public numero As String

Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private Sql As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim F1 As Date
Dim F2 As Date

Dim PrimeraVez As Boolean
Dim Cancelado As Boolean

Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub

Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
            
    MontaSQL = True
End Function


Private Sub cmdAccion_Click(Index As Integer)

    If Not DatosOK Then Exit Sub
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    tabla = " "
    
    If Not MontaSQL Then Exit Sub
    
    If Not CargarTemporales Then Exit Sub
    
    If Not HayRegParaInforme("tmpbalancesumas", "codusu = " & vUsu.Codigo) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
    
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
        txtDias(0).Text = 90
        
        PonFoco txtFecha(0)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmppal.Icon
        
    'Otras opciones
    Me.Caption = "Informe de Situación"

    For i = 0 To 1
        Me.ImgFec(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next i
     
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    Label3(0).Visible = (vParam.NroAriges > 0)
    ImgFec(1).Visible = (vParam.NroAriges > 0)
    ImgFec(1).Enabled = (vParam.NroAriges > 0)
    txtFecha(1).Visible = (vParam.NroAriges > 0)
    txtFecha(1).Enabled = (vParam.NroAriges > 0)
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1
        IndCodigo = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
    End Select
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
        frmppal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmppal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmppal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmppal.cd1.FilterIndex = 1
    frmppal.cd1.ShowSave
    If frmppal.cd1.FileTitle <> "" Then
        If Dir(frmppal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmppal.cd1.FileName
    End If
End Sub

Private Sub PushButtonImpr_Click()
    frmppal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub




Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
End Sub


Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    
    Sql = "SELECT `tmptesoreriacomun`.`texto1` cuenta , `tmptesoreriacomun`.`texto2` Conta, `tmptesoreriacomun`.`opcion` BD, `tmptesoreriacomun`.`texto5` Nombre, `tmptesoreriacomun`.`texto3` NroFra, `tmptesoreriacomun`.`fecha1` FecFra, `tmptesoreriacomun`.`fecha2` FecVto, `tmptesoreriacomun`.`importe1` Gasto, `tmptesoreriacomun`.`importe2` Recibo"
    Sql = Sql & " FROM   `tmptesoreriacomun` `tmptesoreriacomun`"
    Sql = Sql & " WHERE `tmptesoreriacomun`.codusu = " & vUsu.Codigo
    Sql = Sql & " ORDER BY `tmptesoreriacomun`.`texto1`, `tmptesoreriacomun`.`texto2`, `tmptesoreriacomun`.`opcion`, `tmptesoreriacomun`.`fecha1`"
    
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0903-00"
    
        
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "SituacionTes.rpt"
    
    cadParam = cadParam & "pFecha=""" & txtFecha(0).Text & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & "pDias=" & txtDias(0).Text & "|"
    numParam = numParam + 1
    
    
    
    
    cadFormula = "{tmpbalancesumas.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 52
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub


Private Function CargarTemporales() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim i As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim SqlInsert As String
Dim SqlValues As String
Dim TotalImp As Currency

    CargarTemporales = False
    
    
    On Error GoTo eCargarTemporales
    
    Label9.Caption = "Preparando tablas"
    Label9.Refresh
    Sql = "Delete from tmpbalancesumas where codusu =" & vUsu.Codigo
    Conn.Execute Sql
                
                
    F2 = CDate(txtFecha(0).Text)
    F1 = DateAdd("d", CLng(txtDias(0).Text) * (-1), F2)
                
                
    Sql = ""
    Screen.MousePointer = vbHourglass
    
            
    Label9.Caption = "Cálculo saldos de cuentas de banco"
    Label9.Refresh
            
    ' SITUACION DE LOS BANCOS
    Sql = "select * from bancos order by codmacta "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlInsert = "insert into tmpbalancesumas (codusu,cta,nomcta,TotalD,TotalH) values "
    SqlValues = ""
    
    While Not Rs.EOF
        Label9.Caption = "Banco: " & Rs!codmacta
        Label9.Refresh
        
        SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(Rs!Descripcion, "T") & ","
        
        Sql2 = "select sum(coalesce(timporteD,0)), sum(coalesce(timporteh,0)) from hlinapu where codmacta = " & DBSet(Rs!codmacta, "T")
        Sql2 = Sql2 & " and fechaent <= " & DBSet(F2, "F")
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            SqlValues = SqlValues & DBSet(Rs2.Fields(0).Value, "N") & "," & DBSet(Rs2.Fields(1).Value, "N") & "),"
        Else
            SqlValues = SqlValues & "0,0),"
        End If
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
            
    Set Rs = Nothing
            
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert & SqlValues
    End If
    
    'COBROS PENDIENTES
            
    Label9.Caption = "Cobros "
    Label9.Refresh
            
    SqlValues = ""
            
    Sql2 = "select sum(impvenci + coalesce(gastos,0) - coalesce(impcobro,0)) from cobros where impvenci + coalesce(gastos,0) - coalesce(impcobro,0) <> 0 "
    Sql2 = Sql2 & " and fecvenci between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
    Sql2 = Sql2 & " and situacionjuri = 0"
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs2.EOF Then
        SqlValues = "(" & vUsu.Codigo & ",'COBROS',''," & DBSet(Rs2.Fields(0).Value, "N") & ",0)"
    Else
        SqlValues = "(" & vUsu.Codigo & ",'COBROS',''," & DBSet(0, "N") & ",0)"
    End If
    Set Rs2 = Nothing
    
    Conn.Execute SqlInsert & SqlValues
    
    ' PAGOS PENDIENTES
    
    Label9.Caption = "Pagos "
    Label9.Refresh
            
    SqlValues = ""
            
    Sql2 = "select sum(impefect  - coalesce(imppagad,0)) from pagos where impefect - coalesce(imppagad,0) <> 0 "
    Sql2 = Sql2 & " and fecefect between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs2.EOF Then
        SqlValues = "(" & vUsu.Codigo & ",'PAGOS','',0," & DBSet(Rs2.Fields(0).Value, "N") & ")"
    Else
        SqlValues = "(" & vUsu.Codigo & ",'PAGOS','',0," & DBSet(0, "N") & ")"
    End If
    Set Rs2 = Nothing
    
    Conn.Execute SqlInsert & SqlValues
    
    If vParam.NroAriges > 0 Then
        ' CALCULO DE LOS ALBARANES CON IVA Y VTOS SEGUN LA FORMA DE PAGO
        If CargarTemporalFacturas Then
            If CargarTemporalCobros Then
                ' insertamos la facturacion pendiente
            
                Sql = "select sum(impvenci) from tmpcobros where codusu = " & DBSet(vUsu.Codigo, "N") & " and fecvenci between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
                TotalImp = DevuelveValor(Sql)
                
                Sql = "insert into tmpbalancesumas (codusu,cta,nomcta,TotalD,TotalH) values "
                Sql = Sql & "(" & vUsu.Codigo & ",'ALV',''," & DBSet(TotalImp, "N") & ",0)"
                Conn.Execute Sql
            
            End If
        End If
        If CargarTemporalFacturasPROV Then
            If CargarTemporalPagos Then
                ' insertamos la facturacion pendiente
            
                Sql = "select sum(impvenci) from tmpcobros where codusu = " & DBSet(vUsu.Codigo, "N") & " and fecvenci between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
                TotalImp = DevuelveValor(Sql)
                
                Sql = "insert into tmpbalancesumas (codusu,cta,nomcta,TotalD,TotalH) values "
                Sql = Sql & "(" & vUsu.Codigo & ",'ALP','',0," & DBSet(TotalImp, "N") & ")"
                Conn.Execute Sql
            
            End If
        End If
        
    End If
            
    Label9.Caption = ""
    Label9.Refresh
            
    CargarTemporales = True
    Exit Function
    
eCargarTemporales:
    MuestraError Err.Number, "Cargar Temporales", Err.Description
End Function

Private Function CargarTemporalFacturas() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Base As Currency

Dim TotalImp As Currency
Dim SqlInsert As String
Dim SqlValues As String


    On Error GoTo eCargarTemporalFacturas

    CargarTemporalFacturas = False

    Sql = "delete from tmpconext where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
                                                  '   cliente,  albaran, codtipom, codforpa, importe
    SqlInsert = "insert into tmpconext ( codusu , pos,         timported, cta,     timporteh, saldo ) values "
    
    Sql = "select codclien, codtipom, numalbar, codforpa, dtoppago, dtognral from ariges" & vParam.NroAriges & ".scaalb where factursn = 1"
    Sql = Sql & " order by 1,2,3,4"
    
    SqlValues = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        Label9.Caption = "Albaran: " & DBLet(Rs!numalbar, "N")
        Label9.Refresh
    
        TotalImp = 0
    
        Sql2 = "select tiposiva.porceiva, sum(slialb.importel) base from ariges" & vParam.NroAriges & ".slialb  slialb, ariges" & vParam.NroAriges & ".sartic sartic, tiposiva "
        Sql2 = Sql2 & " where slialb.numalbar = " & DBSet(Rs!numalbar, "N")
        Sql2 = Sql2 & " and slialb.codtipom = " & DBSet(Rs!codtipom, "T")
        Sql2 = Sql2 & " and slialb.codartic = sartic.codartic and sartic.codigiva = tiposiva.codigiva "
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            Base = DBLet(Rs2!Base, "N") * Round((1 - ((DBLet(Rs!dtoppago, "N") + DBLet(Rs!dtognral, "N")) / 100)), 2)
        
            TotalImp = TotalImp + Round(Base * (1 + (DBLet(Rs2!porceiva, "N") / 100)), 2)
        
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBSet(Rs!codclien, "N") & "," & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!codtipom, "T") & "," & DBSet(Rs!Codforpa, "N") & "," & DBSet(TotalImp, "N") & "),"
        
    
        Rs.MoveNext
    Wend
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert & SqlValues
    End If
    Set Rs = Nothing

    CargarTemporalFacturas = True
    Exit Function

eCargarTemporalFacturas:
    MuestraError Err.Number, "Cargar Temporal Facturas", Err.Description
End Function

Private Function CargarTemporalCobros() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rsvenci As ADODB.Recordset
Dim FecFactu As Date
Dim FecVenci As Date
Dim TotalFac As Currency
Dim TotalImp As Currency
Dim ImpVenci As Currency
Dim i As Integer
Dim cadvalues2 As String
Dim SqlInsert As String
Dim SqlValues As String


    On Error GoTo eCargarTemporalCobros

    CargarTemporalCobros = False

    Sql = "delete from tmpcobros where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    SqlInsert = "insert into tmpcobros ( codusu , fecvenci, impvenci ) values "
    
              '   codforpa,          importe
    Sql = "select timporteh codforpa, sum(saldo) totalfac from tmpconext where codusu = " & vUsu.Codigo
    Sql = Sql & " group by 1"
    Sql = Sql & " order by 1"
    
    cadvalues2 = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        Label9.Caption = "FP.:" & DBSet(Rs!Codforpa, "N")
        Label9.Refresh
    
        TotalImp = 0
        
        Sql = "SELECT numerove, primerve, restoven FROM ariges" & vParam.NroAriges & ".sforpa WHERE codforpa=" & DBSet(Rs!Codforpa, "N")
        Set Rsvenci = New ADODB.Recordset
        Rsvenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        FecFactu = CDate(Format(txtFecha(1).Text, "dd/mm/yyyy"))
        TotalFac = DBLet(Rs!TotalFac, "N")
        
        
        If Not Rsvenci.EOF Then
            If DBLet(Rsvenci!numerove, "N") > 0 Then
                '-------- Primer Vencimiento
                i = 1
                'FECHA VTO
                FecVenci = FecFactu
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(Rsvenci!primerve, "N"), FecVenci)
                '===
                
                
                '[Monica]03/07/2013: añado trim(codmacta)
                cadvalues2 = cadvalues2 & "(" & vUsu.Codigo & "," & DBSet(FecVenci, "F") & ", "
                
                'IMPORTE del Vencimiento
                If Rsvenci!numerove = 1 Then
                    ImpVenci = DBLet(TotalFac, "N")
                Else
                    ImpVenci = Round(TotalFac / DBLet(Rsvenci!numerove, "N"), 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * Rsvenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * DBLet(Rsvenci!numerove, "N")), 2)
                    End If
                End If
                
                cadvalues2 = cadvalues2 & DBSet(ImpVenci, "N") & "),"
                
            
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For i = 2 To Rsvenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    '===
                        
                    cadvalues2 = cadvalues2 & "(" & vUsu.Codigo & "," & DBSet(FecVenci, "F") & ", "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round(TotalFac / Rsvenci!numerove, 2)
                    cadvalues2 = cadvalues2 & DBSet(ImpVenci, "N") & "),"
                Next i
            End If
        End If
        
        Rs.MoveNext
    Wend
    
    If cadvalues2 <> "" Then
        cadvalues2 = Mid(cadvalues2, 1, Len(cadvalues2) - 1)
        Conn.Execute SqlInsert & cadvalues2
    End If
    Set Rs = Nothing



    CargarTemporalCobros = True
    Exit Function

eCargarTemporalCobros:
    MuestraError Err.Number, "Cargar Temporal Cobros", Err.Description
End Function



Private Function CargarTemporalFacturasPROV() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim SqlInsert As String
Dim SqlValues As String
Dim TotalImp As Currency
Dim Base As Currency

    On Error GoTo eCargarTemporalFacturas

    CargarTemporalFacturasPROV = False

    Sql = "delete from tmpconext where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
                                                '  proveed,albaran, fechaalb codforpa, importe
    SqlInsert = "insert into tmpconext ( codusu , pos ,    cta,     fechaent, timporteh, saldo ) values "
    
    Sql = "select codprove, numalbar, fechaalb, codforpa, dtoppago, dtognral from ariges" & vParam.NroAriges & ".scaalp "
    Sql = Sql & " order by 1,2,3,4"
    
    SqlValues = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        Label9.Caption = "Alb.Prov: " & DBLet(Rs!numalbar, "T")
        Label9.Refresh
    
        TotalImp = 0
    
        Sql2 = "select tiposiva.porceiva, sum(slialp.importel) base from ariges" & vParam.NroAriges & ".slialp  slialp, ariges" & vParam.NroAriges & ".sartic sartic, tiposiva "
        Sql2 = Sql2 & " where slialp.numalbar = " & DBSet(Rs!numalbar, "T")
        Sql2 = Sql2 & " and slialp.codprove = " & DBSet(Rs!codprove, "N")
        Sql2 = Sql2 & " and slialp.fechaalb = " & DBSet(Rs!fechaalb, "F")
        Sql2 = Sql2 & " and slialp.codartic = sartic.codartic and sartic.codigiva = tiposiva.codigiva "
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        Set Rs2 = New ADODB.Recordset
        'Rs2.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText  estaba asi. Monta el sql2 y luego abre el sql? wtf?
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF
            Base = DBLet(Rs2!Base, "N") * Round((1 - ((DBLet(Rs!dtoppago, "N") + DBLet(Rs!dtognral, "N")) / 100)), 2)
        
            TotalImp = TotalImp + Round(Base * (1 + (DBLet(Rs2!porceiva, "N") / 100)), 2)
        
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBSet(Rs!codprove, "N") & "," & DBSet(Rs!numalbar, "T") & "," & DBSet(Rs!fechaalb, "T") & "," & DBSet(Rs!Codforpa, "N") & "," & DBSet(TotalImp, "N") & "),"
        
    
        Rs.MoveNext
    Wend
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert & SqlValues
    End If
    Set Rs = Nothing

    CargarTemporalFacturasPROV = True
    Exit Function

eCargarTemporalFacturas:
    MuestraError Err.Number, "Cargar Temporal Facturas Proveedor", Err.Description
End Function

Private Function CargarTemporalPagos() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rsvenci As ADODB.Recordset
Dim FecFactu As Date
Dim TotalFac As Currency
Dim TotalImp As Currency
Dim i As Integer
Dim cadvalues2 As String
Dim SqlInsert As String
Dim SqlValues As String
Dim FecVenci As Date
Dim ImpVenci As Currency

    On Error GoTo eCargarTemporalPagos

    CargarTemporalPagos = False

    Sql = "delete from tmpcobros where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    SqlInsert = "insert into tmpcobros ( codusu , fecvenci, impvenci ) values "
    
              '   codforpa,          importe
    Sql = "select timporteh codforpa, sum(saldo) totalfac from tmpconext where codusu = " & vUsu.Codigo
    Sql = Sql & " group by 1 "
    Sql = Sql & " order by 1"
    
    cadvalues2 = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
    
        Label9.Caption = "FP.:" & DBSet(Rs!Codforpa, "N")
        Label9.Refresh
    
        TotalImp = 0
        
        Sql = "SELECT numerove, primerve, restoven FROM ariges" & vParam.NroAriges & ".sforpa WHERE codforpa=" & DBSet(Rs!Codforpa, "N")
        Set Rsvenci = New ADODB.Recordset
        Rsvenci.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        FecFactu = CDate(Format(txtFecha(1).Text, "dd/mm/yyyy")) '????
        TotalFac = DBLet(Rs!TotalFac, "N")
        
        
        If Not Rsvenci.EOF Then
            If DBLet(Rsvenci!numerove, "N") > 0 Then
                '-------- Primer Vencimiento
                i = 1
                'FECHA VTO
                FecVenci = FecFactu
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(Rsvenci!primerve, "N"), FecVenci)
                '===
                
                
                '[Monica]03/07/2013: añado trim(codmacta)
                cadvalues2 = cadvalues2 & "(" & vUsu.Codigo & "," & DBSet(FecVenci, "F") & ", "
                
                'IMPORTE del Vencimiento
                If Rsvenci!numerove = 1 Then
                    ImpVenci = DBLet(TotalFac, "N")
                Else
                    ImpVenci = Round(TotalFac / Rsvenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * Rsvenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * Rsvenci!numerove), 2)
                    End If
                End If
                
                cadvalues2 = cadvalues2 & DBSet(ImpVenci, "N") & "),"
                
            
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For i = 2 To Rsvenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(Rsvenci!restoven, "N"), FecVenci)
                    '===
                        
                    cadvalues2 = cadvalues2 & "(" & vUsu.Codigo & "," & DBSet(FecVenci, "F") & ", "
                    
                    'IMPORTE Resto de Vencimientos
                    ImpVenci = Round(TotalFac / Rsvenci!numerove, 2)
                    cadvalues2 = cadvalues2 & DBSet(ImpVenci, "N") & "),"
                Next i
            End If
        End If
        
        Rs.MoveNext
    Wend
    
    If cadvalues2 <> "" Then
        cadvalues2 = Mid(cadvalues2, 1, Len(cadvalues2) - 1)
        Conn.Execute SqlInsert & cadvalues2
    End If
    Set Rs = Nothing

    CargarTemporalPagos = True
    Exit Function

eCargarTemporalPagos:
    MuestraError Err.Number, "Cargar Temporal Pagos", Err.Description
End Function






Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtFecha(0).Text = "" Then
        MsgBox "Debe dar un valor a la fecha.", vbExclamation
        PonFoco txtFecha(0)
        Exit Function
    End If
    
    If txtDias(0).Text = "" Then
        MsgBox "Debe dar un valor a los días.", vbExclamation
        PonFoco txtDias(0)
        Exit Function
    End If
    
    If vParam.NroAriges > 0 Then
        If txtFecha(1).Text = "" Then
            MsgBox "Debe introducir un valor en la fecha de factura.", vbExclamation
            PonFoco txtFecha(1)
            Exit Function
        End If
    End If
    
    DatosOK = True


End Function


Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtFecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtDias_GotFocus(Index As Integer)
    ConseguirFoco txtDias(Index), 3
End Sub

Private Sub txtDias_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDias_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtDias(Index).Text = Trim(txtDias(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'dias
            txtDias(Index).Text = Format(txtDias(Index).Text, "####0")
            
    End Select

End Sub

