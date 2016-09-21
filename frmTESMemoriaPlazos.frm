VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESMemoriaPlazos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   Icon            =   "frmTESMemoriaPlazos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   13
      Top             =   2310
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   16
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   720
         Width           =   1515
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
      Height          =   2235
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6915
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
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   870
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
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1290
         Width           =   1305
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   6270
         TabIndex        =   24
         Top             =   330
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
         Height          =   195
         Index           =   18
         Left            =   300
         TabIndex        =   12
         Top             =   570
         Width           =   2280
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
         Index           =   17
         Left            =   300
         TabIndex        =   11
         Top             =   930
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
         Index           =   16
         Left            =   300
         TabIndex        =   10
         Top             =   1290
         Width           =   615
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   990
         Top             =   900
         Width           =   240
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   1
         Left            =   990
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label lblNumFactu 
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
         Left            =   2580
         TabIndex        =   9
         Top             =   2700
         Width           =   4035
      End
      Begin VB.Label lblNumFactu 
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
         Left            =   2610
         TabIndex        =   8
         Top             =   2340
         Width           =   4035
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
         TabIndex        =   7
         Top             =   3990
         Width           =   4095
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
         TabIndex        =   6
         Top             =   3630
         Width           =   4095
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
      Top             =   5130
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
      Left            =   4230
      TabIndex        =   2
      Top             =   5130
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
      Top             =   5100
      Width           =   1335
   End
End
Attribute VB_Name = "frmTESMemoriaPlazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1309


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
Public Legalizacion As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmConta As frmBasico
Attribute frmConta.VB_VarHelpID = -1
Private WithEvents frmAgen As frmBasico
Attribute frmAgen.VB_VarHelpID = -1
Private WithEvents frmDpto As frmBasico
Attribute frmDpto.VB_VarHelpID = -1
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1

Private SQL As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim PrimeraVez As Boolean

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

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    tabla = "pagos"
    
    
    If Not MontaSQL Then Exit Sub
    
    cadFormula = "{tmpimpbalance.codusu} = " & vUsu.Codigo
    
    
    If Not HayRegParaInforme("tmpimpbalance", "codusu = " & DBSet(vUsu.Codigo, "N")) Then Exit Sub
    
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
    Me.Icon = frmPpal.Icon
        
    'Otras opciones
    Me.Caption = "Memoria Plazos de Pago"

    
    For I = 0 To 1
        Me.ImgFec(I).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next I
     
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
     
    txtFecha(0).Text = Format(vParam.fechaini, "dd/mm/yyyy")
    txtFecha(1).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    
    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
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


Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
        frmPpal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmPpal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmPpal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmPpal.cd1.FilterIndex = 1
    frmPpal.cd1.ShowSave
    If frmPpal.cd1.FileTitle <> "" Then
        If Dir(frmPpal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmPpal.cd1.FileName
    End If
End Sub


Private Sub PushButtonImpr_Click()
    frmPpal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
End Sub


Private Sub AccionesCSV()
Dim SQL2 As String

    'Monto el SQL
    SQL = "SELECT cobros.codmacta Cuenta, cobros.nomclien Descripcion, hlinapu.fecdevol FecDevol, "
    SQL = SQL & "cobros.numserie Serie, cobros.numfactu Factura, cobros.fecfactu FecFra, cobros.numorden Vto, hlinapu.timporteh - hlinapu.timported Importe, "
    SQL = SQL & "hlinapu.gastodev Gastos, hlinapu.coddevol Devol, wdevolucion.descripcion Descripcion "
    SQL = SQL & " FROM  (cobros INNER JOIN hlinapu ON cobros.numserie = hlinapu.numserie AND "
    SQL = SQL & " cobros.numfactu = hlinapu.numfaccl AND cobros.fecfactu = hlinapu.fecfactu AND "
    SQL = SQL & " cobros.numorden = hlinapu.numorden) "
    SQL = SQL & "  LEFT JOIN usuarios.wdevolucion ON hlinapu.coddevol = wdevolucion.codigo "
    
    If cadselect <> "" Then SQL = SQL & " where " & cadselect
    
    
    SQL = SQL & " ORDER BY " & SQL2

            
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "1309-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "CobrosPdtes.rpt"
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 15
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub

Private Function CargarTemporal() As Boolean
Dim SQL As String
Dim vSql As String
Dim vSql1 As String
Dim vSql2 As String
Dim SqlValues As String
Dim DifDias As Long
Dim NroFras As Long
Dim Total As Currency

Dim Valor1 As Currency
Dim Importe1 As Currency
Dim Ratio1 As Currency

Dim Valor2 As Currency
Dim importe2 As Currency
Dim Ratio2 As Currency

Dim Valor1ant As Currency
Dim Importe1ant As Currency
Dim Ratio1ant As Currency

Dim Valor2ant As Currency
Dim Importe2ant As Currency
Dim Ratio2ant As Currency

Dim PM As Currency
Dim PMAnt As Currency


Dim F1Ant As String
Dim F2Ant As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False
    
    F1Ant = DateAdd("yyyy", -1, CDate(txtFecha(0).Text))
    F2Ant = DateAdd("yyyy", -1, CDate(txtFecha(1).Text))
    
    SQL = "delete from tmpimpbalance where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "insert into tmpimpbalance (codusu, codigo, descripcion, importe1, importe2) values (" & vUsu.Codigo & ","
    
    SqlValues = ""
    
        
    ' ratio de operaciones pagadas
    
    vSql = "select round(datediff(fecultpa,factpro.fecharec) * impefect,2) a, impefect b from factpro inner join pagos on factpro.numserie = pagos.numserie and factpro.numfactu = pagos.numfactu and factpro.fecfactu = pagos.fecfactu "
    vSql = vSql & " where not pagos.fecultpa is null "
    vSql = vSql & " and factpro.fecharec >=" & DBSet(txtFecha(0).Text, "F") & " and factpro.fecharec <= " & DBSet(txtFecha(1).Text, "F")
    vSql = vSql & " and pagos.fecultpa <=" & DBSet(txtFecha(1).Text, "F")
    
    vSql1 = "select sum(a) from (" & vSql & ") aaa"
    Valor1 = DevuelveValor(vSql1)
    
    vSql1 = "select sum(b) from (" & vSql & ") aaa"
    Importe1 = DevuelveValor(vSql1)
    
    Ratio1 = 0
    If Importe1 <> 0 Then Ratio1 = Round(Valor1 / Importe1, 2)
    
    ' ratio de operaciones pagadas (anterior)
    
    vSql = "select round(datediff(fecultpa,factpro.fecharec) * impefect,2) a, impefect b from factpro inner join pagos on factpro.numserie = pagos.numserie and factpro.numfactu = pagos.numfactu and factpro.fecfactu = pagos.fecfactu "
    vSql = vSql & " where not pagos.fecultpa is null "
    vSql = vSql & " and factpro.fecharec >=" & DBSet(F1Ant, "F") & " and factpro.fecharec <= " & DBSet(F2Ant, "F")
    vSql = vSql & " and pagos.fecultpa <=" & DBSet(F2Ant, "F")
    
    vSql1 = "select sum(a) from (" & vSql & ") aaa"
    Valor1ant = DevuelveValor(vSql1)
    
    vSql1 = "select sum(b) from (" & vSql & ") aaa"
    Importe1ant = DevuelveValor(vSql1)
    
    Ratio1ant = 0
    If Importe1ant <> 0 Then Ratio1ant = Round(Valor1ant / Importe1ant, 2)
    
    
    SqlValues = "2,'Ratio de operaciones pagadas'," & DBSet(Ratio1, "N") & "," & DBSet(Ratio1ant, "N") & ")"
    Conn.Execute SQL & SqlValues
    
    ' importe pagos realizados
    SqlValues = "4,'Total pagos realizados'," & DBSet(Importe1, "N") & "," & DBSet(Importe1ant, "N") & ")"
    Conn.Execute SQL & SqlValues
    
    ' ratio de operaciones pendientes de pago
    vSql = "select round(datediff(" & DBSet(txtFecha(1).Text, "F") & ",factpro.fecharec) * impefect,2) a, impefect b from factpro inner join pagos on factpro.numserie = pagos.numserie and factpro.numfactu = pagos.numfactu and factpro.fecfactu = pagos.fecfactu "
    vSql = vSql & " where factpro.fecharec >=" & DBSet(txtFecha(0).Text, "F") & " and factpro.fecharec <= " & DBSet(txtFecha(1).Text, "F")
    vSql = vSql & " and (pagos.fecultpa is null or pagos.fecultpa = '' or  pagos.fecultpa > " & DBSet(txtFecha(1).Text, "F") & ") "
    
    vSql1 = "select sum(a) from (" & vSql & ") aaa"
    Valor2 = DevuelveValor(vSql1)
    
    vSql1 = "select sum(b) from (" & vSql & ") aaa"
    importe2 = DevuelveValor(vSql1)
    
    Ratio2 = 0
    If importe2 <> 0 Then Ratio2 = Round(Valor2 / importe2, 2)
    
    ' ratio de operaciones pendientes de pago (anterior)
    vSql = "select round(datediff(" & DBSet(F2Ant, "F") & ",factpro.fecharec) * impefect,2) a, impefect b from factpro inner join pagos on factpro.numserie = pagos.numserie and factpro.numfactu = pagos.numfactu and factpro.fecfactu = pagos.fecfactu "
    vSql = vSql & " where factpro.fecharec >=" & DBSet(F1Ant, "F") & " and factpro.fecharec <= " & DBSet(F2Ant, "F")
    vSql = vSql & " and (pagos.fecultpa is null or pagos.fecultpa = '' or  pagos.fecultpa > " & DBSet(F2Ant, "F") & ") "
    
    vSql1 = "select sum(a) from (" & vSql & ") aaa"
    Valor2ant = DevuelveValor(vSql1)
    
    vSql1 = "select sum(b) from (" & vSql & ") aaa"
    Importe2ant = DevuelveValor(vSql1)
    
    Ratio2ant = 0
    If Importe2ant <> 0 Then Ratio2ant = Round(Valor2ant / Importe2ant, 2)
    
    SqlValues = "3,'Ratio de operaciones pendientes de pago'," & DBSet(Ratio2, "N") & "," & DBSet(Ratio2ant, "N") & ")"
    Conn.Execute SQL & SqlValues
    
    ' importe operaciones pendientes de pago
    SqlValues = "5,'Total pagos pendientes'," & DBSet(importe2, "N") & "," & DBSet(Importe2ant, "N") & ")"
    Conn.Execute SQL & SqlValues
    
    
    ' periodo medio de pago a proveedores
    PM = 0
    PMAnt = 0
    If (Importe1 + importe2) <> 0 Then PM = Round(((Ratio1 * Importe1) + (Ratio2 * importe2)) / (Importe1 + importe2), 2)
    If (Importe1ant + Importe2ant) <> 0 Then PMAnt = Round(((Ratio1ant * Importe1ant) + (Ratio2ant * Importe2ant)) / (Importe1ant + Importe2ant), 2)
    
    SqlValues = "1,'Periodo medio de pago a proveedores'," & DBSet(PM, "N") & "," & DBSet(PMAnt, "N") & ")"
    Conn.Execute SQL & SqlValues
    
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function

Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String
Dim I As Integer


    MontaSQL = False
    
    If Not PonerDesdeHasta(tabla & ".fecfactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
            
    If cadFormula <> "" Then cadFormula = "(" & cadFormula & ")"
    If cadselect <> "" Then cadselect = "(" & cadselect & ")"
    
    MontaSQL = CargarTemporal
    
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

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    ' debe introducir las fechas y el plazo en dias
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        MsgBox "Debe introducir obligatoriamente el rango de fechas. Reintroduzca.", vbExclamation
        PonFoco txtFecha(0)
        Exit Function
    Else
        If DateDiff("yyyy", CDate(txtFecha(0).Text), CDate(txtFecha(1).Text)) > 1 Then
            MsgBox "La diferencia entre fechas ha de ser máximo de un año. Reintroduzca.", vbExclamation
            PonFoco txtFecha(0)
            Exit Function
        End If
    End If
    
    
    DatosOK = True

End Function


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

