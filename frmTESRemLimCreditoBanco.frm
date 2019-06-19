VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESbancosLimiteCredi 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Ordenación"
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
      Left            =   9480
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton optVarios 
         Caption         =   "Descripción Cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   390
         TabIndex        =   22
         Top             =   1920
         Width           =   3555
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Fecha Vencimiento "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   390
         TabIndex        =   21
         Top             =   990
         Width           =   3045
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   390
         TabIndex        =   20
         Top             =   570
         Width           =   1395
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   390
         TabIndex        =   19
         Top             =   1440
         Width           =   1185
      End
   End
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
      Height          =   2835
      Left            =   120
      TabIndex        =   13
      Top             =   30
      Width           =   11355
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
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   780
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
         Index           =   1
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1200
         Width           =   1305
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2340
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   4128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha vencimiento"
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
         Height          =   240
         Index           =   0
         Left            =   7680
         TabIndex        =   26
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   7920
         TabIndex        =   25
         Top             =   840
         Width           =   690
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
         Height          =   195
         Index           =   4
         Left            =   7920
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   8610
         Picture         =   "frmTESRemLimCreditoBanco.frx":0000
         Top             =   810
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   8610
         Picture         =   "frmTESRemLimCreditoBanco.frx":008B
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmTESRemLimCreditoBanco.frx":0116
         ToolTipText     =   "Puntear al Debe"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   360
         Picture         =   "frmTESRemLimCreditoBanco.frx":0260
         ToolTipText     =   "Quitar al Debe"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Banco"
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
         Height          =   225
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1200
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
      Left            =   10200
      TabIndex        =   4
      Top             =   5850
      Width           =   1335
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
      Left            =   8520
      TabIndex        =   2
      Top             =   5850
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
      TabIndex        =   3
      Top             =   5790
      Width           =   1335
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
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   7395
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
         Left            =   5790
         TabIndex        =   16
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   7050
         TabIndex        =   15
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   7050
         TabIndex        =   14
         Top             =   1200
         Width           =   255
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
         TabIndex        =   12
         Top             =   1680
         Width           =   5145
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
         Width           =   5145
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
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   3825
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
         TabIndex        =   9
         Top             =   2160
         Width           =   975
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
         TabIndex        =   8
         Top             =   1680
         Width           =   975
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
         TabIndex        =   7
         Top             =   1200
         Width           =   1515
      End
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
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmTESbancosLimiteCredi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
'Private WithEvents frmDia As frmTiposDiario
'Private WithEvents frmC As frmColCtas

Private SQL As String
Dim cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String


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


Private Sub cmdAccion_Click(Index As Integer)
        
    cad = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            cad = cad & ", " & DBSet(ListView1.ListItems(i).SubItems(1), "T")
        End If
    Next
    If cad = "" Then
        MsgBox "Seleccione algun banco", vbExclamation
        Exit Sub
    End If
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not CargarTemporal Then Exit Sub
    
    
    
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
    Me.Caption = "Disponible para remesar por entidad"
    
    PrimeraVez = True
     txtfecha(0).Text = Format(Now, "dd/mm/yyyy")
    Set miRsAux = New ADODB.Recordset
    cad = "select bancos.codmacta,descripcion,nommacta,remesamaximo from bancos inner join cuentas on bancos.codmacta=cuentas.codmacta"
    cad = cad & " WHERE remesamaximo>0 ORDER BY descripcion,nommacta"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ListView1.ListItems.Clear
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        cad = DBLet(miRsAux!Descripcion, "T")
        If cad = "" Then cad = miRsAux!Nommacta
        cad = cad & " (" & miRsAux!codmacta & ")"
        ListView1.ListItems.Add , , cad
        ListView1.ListItems(i).SubItems(1) = miRsAux!codmacta
        ListView1.ListItems(i).ToolTipText = Format(miRsAux!remesamaximo, FormatoImporte)
        If CadenaDesdeOtroForm = miRsAux!codmacta Then ListView1.ListItems(i).Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
End Sub



Private Sub frmF_Selec(vFecha As Date)
    SQL = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = Index = 1
    Next
    
End Sub

Private Sub imgFec_Click(Index As Integer)
    'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtfecha(Index).Text <> "" Then frmF.Fecha = CDate(txtfecha(Index).Text)
        SQL = ""
        frmF.Show vbModal
        Set frmF = Nothing
        If SQL <> "" Then
            txtfecha(Index).Text = SQL
            PonFoco txtfecha(Index)
        End If
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
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



Private Sub AccionesCSV()
Dim Sql2 As String

    'Monto el SQL
    SQL = "select reftalonpag ctaBanco,text33csb Banco,text41csb Credito,fecvenci 'F.Vto', impvenci 'Imp.Vto',gastos,cliente Pdte from"
    SQL = SQL & " tmpcobros2 where codusu=" & vUsu.Codigo & " order by     reftalonpag,fecvenci asc,numfactu ,numorden"

    
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = True
        
    indRPT = "0609-02"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu

    numParam = numParam + 1
    cadFormula = "{tmpcobros2.codusu}=" & vUsu.Codigo
'
'
'    'ordenacion
'    If optVarios(0).Value Then cadParam = cadParam & "pOrden=0|"
'    If optVarios(1).Value Then cadParam = cadParam & "pOrden=1|"
'    If optVarios(2).Value Then cadParam = cadParam & "pOrden=2|"
'    If optVarios(3).Value Then cadParam = cadParam & "pOrden=3|"
'    numParam = numParam + 1
    vMostrarTree = True
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then
        If Not CopiarFicheroASalida(False, txtTipoSalida(2).Text, False) Then ExportarPDF = False
    End If
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 40
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function CargarTemporal() As Boolean
Dim cad As String
Dim ImporteLinea As Currency

    On Error GoTo eCargarTemporal

    CargarTemporal = False


    cad = "delete from tmpcobros2 where codusu= " & DBSet(vUsu.Codigo, "N")
    Conn.Execute cad
    
    RC = ""
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then RC = RC & ", '" & ListView1.ListItems(i).SubItems(1) & "'"
    Next
    RC = Mid(RC, 2)
    
    
    
        ',ctabanc1
    'cad = " select " & vUsu.Codigo & " ,ctabanc1,cobros.numserie, cobros.numfactu, cobros.fecfactu, cobros.numorden,fecvenci,impvenci,gastos,cobros.codforpa,"
    'cad = cad & " codmacta,nomclien , nomforpa"
    
    cad = " select " & vUsu.Codigo & " ,ctabanc1, cobros.fecvenci,sum(impvenci) impvenci,sum(coalesce(gastos,0)) gastos ,'' nomforpa"
    
    
    cad = cad & " from cobros left join  formapago on cobros.codforpa=formapago.codforpa"
    cad = cad & " WHERE tipforpa=4 and tiporem=1 and siturem>'B' and impcobro>0"

    cadselect = ""
    If Me.txtfecha(0).Text <> "" Then
        cad = cad & " AND fecvenci >=" & DBSet(txtfecha(0).Text, "F")
        cadselect = cadselect & " desde " & txtfecha(0).Text
    End If
    If Me.txtfecha(1).Text <> "" Then
        cad = cad & " AND fecvenci <=" & DBSet(txtfecha(1).Text, "F")
        cadselect = cadselect & " hasta " & txtfecha(1).Text
    End If
    If cadselect <> "" Then
        cadParam = cadParam & "pDH=""Fecha vencimiento: " & cadselect & """|"
        numParam = numParam + 1
    End If
    cad = cad & " and  ctabanc1 in (" & RC & ")"
    cad = cad & " group  by ctabanc1,fecvenci"
    cad = cad & " order by ctabanc1,fecvenci ,numfactu ,numorden"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    RC = ""
    cad = ""
    
    While Not miRsAux.EOF
        i = i + 1
        
        If miRsAux!ctabanc1 <> RC Then
            SQL = "remesamaximo"
            RC = "concat(descripcion,'|',nommacta,'|')"
            RC = DevuelveDesdeBD(RC, "bancos inner join cuentas on bancos.codmacta=cuentas.codmacta", "bancos.codmacta", miRsAux!ctabanc1, "T", SQL)
              
            ImporteLinea = CCur(SQL)
            SQL = RecuperaValor(RC, 1) 'descriocion banco
            If SQL = "" Then SQL = RecuperaValor(RC, 2) 'nommacta
            
            SQL = ",  (" & vUsu.Codigo & ",'" & miRsAux!ctabanc1 & "'," & DBSet(SQL, "T") & ",'" & Format(ImporteLinea, FormatoImporte) & "',"
            RC = miRsAux!ctabanc1
    
        End If
        
        'codusu,reftalonpag,text33csb,numserie,numfactu,fecfactu,numorden,
        'cad = cad & SQL & DBSet(miRsAux!NUmSerie, "T") & "," & miRsAux!NumFactu & "," & DBSet(miRsAux!FecFactu, "F") & "," & miRsAux!numorden & ","
        cad = cad & SQL & DBSet("A", "T") & "," & 0 & "," & DBSet(Now, "F") & "," & 1 & ","
        
        'fecvenci,impvenci,gastos,codforpa,codmacta,nomclien,cliente
        cad = cad & DBSet(miRsAux!FecVenci, "F") & "," & DBSet(miRsAux!ImpVenci, "N") & "," & DBSet(miRsAux!Gastos, "N", "N") & "," & 0 'miRsAux!Codforpa
        cad = cad & "," & DBSet(miRsAux!nomforpa, "T") & ",'"    'En codpais llevamos nomforpa
        ImporteLinea = ImporteLinea - (miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N"))
        'cad = cad & miRsAux!codmacta & "'," & DBSet(miRsAux!nomclien, "T") & ",'" & Format(ImporteLinea, FormatoImporte) & "')"
        cad = cad & "" & "'," & DBSet("", "T") & ",'" & Format(ImporteLinea, FormatoImporte) & "')"
        
        miRsAux.MoveNext
        If miRsAux.EOF Then i = 100
        
        If i > 99 Then
            cad = Mid(cad, 2)
            cad = "INSERT INTO tmpcobros2(codusu,reftalonpag,text33csb,text41csb,numserie,numfactu,fecfactu,numorden,fecvenci,impvenci,gastos,codforpa,codpais,codmacta,nomclien,cliente) VALUES " & cad
            Conn.Execute cad
            cad = ""
            i = 1 'para que si sale ya , no de msg de sin registros
        End If
        
    Wend
    miRsAux.Close
    
    
    If i > 0 Then
        CargarTemporal = True
    Else
        MsgBox "No existen registros con esos valores", vbExclamation
    End If
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Facturas", Err.Description
End Function



Private Sub txtfecha_LostFocus(Index As Integer)
    txtfecha(Index).Text = Trim(txtfecha(Index).Text)
    If txtfecha(Index).Text <> "" Then
        If Not EsFechaOK(txtfecha(Index)) Then txtfecha(Index).Text = ""
    End If
End Sub

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
