VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportarUtil 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar fichero facturas"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   16245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Importar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   10
      Top             =   8880
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   495
      Left            =   14640
      TabIndex        =   5
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CheckBox check1 
      Caption         =   "Primera linea encabezados"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   3885
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Tag             =   "0"
      Top             =   2040
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   11668
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15975
      Begin VB.CheckBox Check2 
         Caption         =   "Eliminar fichero al finalizar"
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
         Left            =   4200
         TabIndex        =   14
         Top             =   1320
         Value           =   1  'Checked
         Width           =   4125
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
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Asientos"
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
         Left            =   7920
         TabIndex        =   9
         Top             =   240
         Width           =   1305
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Facturas proveedor"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   2625
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Facturas de clientes"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   2505
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar fichero"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12960
         TabIndex        =   6
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   8985
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   15480
         TabIndex        =   15
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
      Begin VB.Label Label2 
         Caption         =   "Fichero"
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
         Left            =   9600
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Fichero"
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
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   690
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   1080
         Top             =   600
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero"
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
      Left            =   240
      TabIndex        =   11
      Top             =   9000
      Width           =   8370
   End
End
Attribute VB_Name = "frmImportarUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1413


Private Const NumeroCamposTratarFraCli = 20



Dim strArray() As String
Dim LineaConErrores As Boolean
Dim Cad As String
Dim Importe As Currency
Dim FacturaAnterior As String
Dim NF As Integer


Dim NumeroCamposTratarFraPro As Integer

Dim numProgres As Long
Dim TotProgres As Long

Private Sub cmdAceptar_Click()
    
    ImportarFraCLI
    
End Sub

Private Sub ImportarFraCLI()
    
    If Me.optVarios(0).Value Then
        
        If ListView1.ListItems.Count > 0 Then
            If MsgBox("¿Continuar con el proceso de importacion?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        
    ElseIf Me.optVarios(1).Value Then
        'PRoveedor
        CadenaDesdeOtroForm = ""
        If Me.cboTipo.ListIndex = 0 Then
            frmMensajes.Opcion = 59
            frmMensajes.Show vbModal
        
        Else
            If ListView1.ListItems.Count > 0 Then
                If MsgBox("¿Continuar con el proceso de importacion?", vbQuestion + vbYesNoCancel) = vbYes Then CadenaDesdeOtroForm = "OK"
            End If
        End If
        If CadenaDesdeOtroForm = "" Then Exit Sub
    
    Else
        
        'If Me.cboTipo.ListIndex Then
        CadenaDesdeOtroForm = ""
        frmMensajes.Opcion = 60
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm = "" Then Exit Sub
        
    End If
    
    
    'Proceso realmente dicho
    Screen.MousePointer = vbHourglass
    If Me.optVarios(0).Value Then
        If Me.cboTipo.ListIndex = 0 Then
            InsertarEnContabilidadFraCliStd
        Else
            InsertarEnContabilidadFraCliSAGE
        End If
        
    ElseIf Me.optVarios(1).Value Then
        If Me.cboTipo.ListIndex = 0 Then
            InsertarEnContabilidadFraProveedorNAV
        Else
            InsertarEnContabilidadFraprovSTD
        End If
    Else
        'ASientos
        InsertarAsientos
    End If
        
        If J > 0 Then
            ListView1.Tag = 0
            CargaEncabezado 2
            Set miRsAux = New ADODB.Recordset
            Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            i = 0
            While Not miRsAux.EOF
                i = i + 1
                ListView1.ListItems.Add , "K" & miRsAux!Codigo
                ListView1.ListItems(i).Text = miRsAux!texto1
                ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
                ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            Set miRsAux = Nothing
            
            
        Else
            
            MsgBox "Proceso finalizado con éxito", vbInformation
            Me.ListView1.ListItems.Clear

            
            If Me.Check2.Value = 1 Then EliminarFichero
            Text1.Text = ""
        End If
    
    Me.cmdAceptar.visible = False
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub EliminarFichero()
    On Error Resume Next
    Kill Text1.Text
    If Err.Number <> 0 Then MuestraError Err.Number
End Sub

Private Sub cmdCancelar_Click()
    If cmdAceptar.visible Then
        If ListView1.ListItems.Count > 0 Then
            Cad = "¿Desea cancelar el proceso de importación?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Command1_Click()


    If Text1.Text = "" Then
        imgppal_Click 0
        If Text1.Text = "" Then Exit Sub
    End If
    
    If ListView1.ListItems.Count > 0 Then
        'Volver a cargar
        If cmdAceptar.visible Then
            'Importacion anterior con datos correctos.
            'Preguntamos
            Cad = "Hay datos correctos pendientes de integrar. " & vbCrLf & "Cancelar proceso  anterior?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    
    
    'Precomprobacion SAGE
  
    
    
    
    
    
    
    
    cmdAceptar.visible = False
    ListView1.ListItems.Clear
    ListView1.Tag = 0 'No puede hacer dblclick
    


    Screen.MousePointer = vbHourglass
    If Me.optVarios(0).Value Then
    
        ImportarFacturasCliente
        
    Else
        If Me.optVarios(1).Value Then
            ImportacionFraPro
        Else
            ImportarAsientos
        End If
        
    End If
    Screen.MousePointer = vbDefault
    Label1.Caption = ""
    
End Sub


Private Sub PonerLabel(TEXTO As String)
    Label1.Caption = TEXTO
    Label1.Refresh
End Sub

Private Sub Form_Load()
    Me.Icon = frmppal.Icon
    For NF = 0 To imgppal.Count - 1
        imgppal(NF).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    
        ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmppal.ImgListComun
        .Buttons(1).Image = 26
    End With
    
    cmdAceptar.visible = False
    Label1.Caption = ""
    CargaCombo 0
    
    
    
    'Check2.Value = IIf(vParam.PathFicherosInteg <> "", 1, 0)
    'Check1.Value = IIf(vParam.PathFicherosInteg <> "", 0, 1)
    
    If vParam.PathFicherosInteg = "" Then
        
        
        Cad = CStr(CheckValueLeer("intetipodoc"))
        If Cad = "1" Then Me.cboTipo.ListIndex = 1
        
        Cad = CStr(CheckValueLeer("intetipodoc1"))
        If Cad <> "1" Then Cad = 0
        Me.Check1.Value = Val(Cad)
        
        Cad = CStr(CheckValueLeer("intetipodoc2"))
        If Cad <> "1" Then Cad = 0
        Me.Check2.Value = Val(Cad)
        
    Else
        Check2.Value = 1
        Check1.Value = 0
    End If
End Sub



'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'
'       Imporatr facturas cliente
'
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
Private Sub ImportarFacturasCliente()
Dim Rc As Byte
    
    'Primer paso. Lectura fichero. Comprobacion basica datos
    PonerLabel "Leyendo fichero"
    If Me.cboTipo.ListIndex = 1 Then
        'SAGE
        Rc = ProcesaFicheroClientesSAGE(Text1.Text, Label1, Check1.Value = 1)
        
        
        
        
        
    Else
        Rc = ImportarFichFracli
    End If
    
    If Rc = 2 Then Exit Sub
        
    If Rc = 1 Then
    
        'Errores en fichero
        'Ha habido errores
        CargaEncabezado 0
    
        'Cargamos datos
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto1
            ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
            ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Exit Sub
    End If
    
    'Si llega aqui siendo proceso SAGE, es uqe ha ido bien
    'SAGE. Ha ido bien
    If Me.cboTipo.ListIndex = 1 Then
        'Cargamos datos de integracion contaplus
        CargaEncabezadoSAGE
        'Datos
        Set miRsAux = New ADODB.Recordset
        Cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on tmpintefrafracli.cta_cli=cuentas.codmacta WHERE codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "C" & Format(i, "000"), miRsAux!Serie
            ListView1.ListItems(i).SubItems(1) = miRsAux!FACTURA
            ListView1.ListItems(i).SubItems(2) = miRsAux!Fecha
            ListView1.ListItems(i).SubItems(3) = DBLet(miRsAux!cta_cli, "T")
            ListView1.ListItems(i).SubItems(4) = DBLet(miRsAux!Nommacta, "T")
            ListView1.ListItems(i).SubItems(5) = " " & Format(miRsAux!calculoimponible, FormatoImporte) 'DBLet(miRsAux!calculoimponible, "N")
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!impventa, FormatoImporte) 'DBLet(miRsAux!impventa, "N")
            ListView1.ListItems(i).SubItems(7) = Format(miRsAux!ImpIva, FormatoImporte) 'DBLet(miRsAux!ImpIva, "N")
            ListView1.ListItems(i).SubItems(8) = Format(miRsAux!TotalFactura, FormatoImporte) 'DBLet(miRsAux!TotalFactura, "N")
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        
        cmdAceptar.visible = True
        Exit Sub
    
    End If
    
    
    'Segundo paso. Comprobacion datos BD (existe serie, ctas, tiposiva....
    'Fichero importado.
    'Comporbaciones BD
    PonerLabel "Comprobando en BD"
    If Not ComprobacionDatosBDFacturas Then
        'Cargaremos errores
        CargaEncabezado 1
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto2
            ListView1.ListItems(i).SubItems(1) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    
    
        Exit Sub
    End If
    
    
    'Ultima comprobacion. Que las facturas suman lo que tienen que sumar
    PonerLabel "Totales"
    If Not ComprobarTotales Then
        CargaEncabezado 2
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto1
            ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
            ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Sub
    End If
    
    
    
    'Si llega aqui, es que ha ido bien. Estamos pendientes de aceptar
    CargaEncabezado 3
    
    Set miRsAux = New ADODB.Recordset
   
    Cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on cta_cli=cuentas.codmacta where codusu= " & vUsu.Codigo
    Cad = Cad & " ORDER BY codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        ListView1.ListItems.Add , "K" & i
        If DBLet(miRsAux!Serie, "T") <> "" Then
            
            ListView1.ListItems(i).Text = miRsAux!Serie
            ListView1.ListItems(i).SubItems(1) = Format(miRsAux!FACTURA, "000000")
            ListView1.ListItems(i).SubItems(2) = miRsAux!Fecha
            
            ListView1.ListItems(i).SubItems(3) = DBLet(miRsAux!Nommacta, "T")
            
            
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!impventa, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            ListView1.ListItems(i).SubItems(7) = " " 'separado
            
            ListView1.ListItems(i).SubItems(8) = Format(miRsAux!calculoimponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(9) = Format(miRsAux!TotalFactura - miRsAux!calculoimponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(10) = Format(miRsAux!TotalFactura, FormatoImporte)
            
            If vEmpresa.TieneTesoreria Then
                Cad = " "
                If DBLet(miRsAux!IBAN, "T") <> "" Or DBLet(miRsAux!txtcsb, "T") <> "" Then Cad = "*"
                ListView1.ListItems(i).SubItems(11) = Cad
            End If
            
        Else
            
            ListView1.ListItems(i).Text = " "
            ListView1.ListItems(i).SubItems(3) = "              IVA " & miRsAux!IVA & "%"
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!impventa, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            For NF = 1 To 10
                If NF < 3 Or NF > 6 Then ListView1.ListItems(i).SubItems(NF) = " "
            Next
            
            If vEmpresa.TieneTesoreria Then ListView1.ListItems(i).SubItems(11) = " "
            
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cmdAceptar.visible = True
    
    
End Sub


Private Sub CargaEncabezado(LaOpcion As Byte)
Dim clmX As ColumnHeader
Dim B As Boolean

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.Checkboxes = False
    Select Case LaOpcion
    Case 0
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Linea"
        clmX.Width = 1000 '1500
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Error"
        clmX.Width = 3500
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Linea fichero"
        clmX.Width = 6500
        
    Case 1
        'Error en campos BD
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Error"
        clmX.Width = 4000
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Valor en fichero"
        clmX.Width = 9000
    Case 2
        'Erro en totales
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Factura"
        clmX.Width = 1200 '1500
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Error"
        clmX.Width = 3000
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Datos"
        clmX.Width = 9500
        
    Case 3
        'OK FACTURAS CLIENTE
        ListView1.Tag = 1
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Serie"
        clmX.Width = 750 '1500
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Factura"
        clmX.Width = 1400
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Fecha"
        clmX.Width = 1300
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Nombre"
        clmX.Width = 4100
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Base"
        clmX.Width = 1250
        clmX.Alignment = lvwColumnRight
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "%Iva"
        clmX.Width = 750
        clmX.Alignment = lvwColumnRight
        
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Iva"
        clmX.Width = 1100
        clmX.Alignment = lvwColumnRight
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = " "  'Separador
        clmX.Width = 400
        
        Set clmX = ListView1.ColumnHeaders.Add()
        'Para facturas clientes es BASES
        'Para frapro NAVARRES pintaremos la retencion
        B = True
        If Me.optVarios(1).Value Then
            'Proveedor. De momento para cualquiera de proveedor
            B = False
        End If
        If B Then
            clmX.Text = "Bases"
        Else
            clmX.Text = "Reten."
        End If
        clmX.Width = 1250
        clmX.Alignment = lvwColumnRight
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Total iva"
        clmX.Width = 1400
        clmX.Alignment = lvwColumnRight
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Total"
        clmX.Width = 1300
        clmX.Alignment = lvwColumnRight
                
        If vEmpresa.TieneTesoreria Then
            Set clmX = ListView1.ColumnHeaders.Add()
            clmX.Text = "€"
            clmX.Width = 370
                    
          
        End If
            
    Case 4
        'Apuntes
        ListView1.Tag = 1
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Asiento"
        clmX.Width = 1000
        
                
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Fecha"
        clmX.Width = 1300
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Cuenta"
        clmX.Width = 1400
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Nombre"
        clmX.Width = 4100
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Documento"
        clmX.Width = 1650
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Ampliacion"
        clmX.Width = 2900
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Debe"
        clmX.Width = 1600
        clmX.Alignment = lvwColumnRight
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Haber"
        clmX.Width = 1600
        clmX.Alignment = lvwColumnRight
        
        'Set clmX = ListView1.ColumnHeaders.Add()
        'clmX.Text = " "  'Separador
        'clmX.Width = 400
        
            
            
    End Select
End Sub

Private Function ImportarFichFracli() As Byte
Dim Aux As String



    On Error GoTo eImportarFicheroFracli
    ImportarFichFracli = 2  'Vacio
    
    NF = -1
    'tmpintefrafracli(codusu,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,
    'tipo_ret,ctaventas,Ccoste,impventa,iva,impiva,recargo,imprecargo,totalfactura,integracion)
    Msg$ = "Serie|Factura|Fecha|Cliente|Forma pago|Tipo opereacion|Cta retencion|Importe retencion|"
    Msg$ = Msg$ & "Tipo retencion|Cta ventas|ccoste|Importe venta|Porcentaje IVA|Importe IVA|Porcentaje recargo|Impor Recargo|Total factura|INTREGRACION|"
    '28 marzo
    Msg$ = Msg$ & "iban|txtcsb|"
    
    
    'Preparamos tabla de insercion
    Conn.Execute "DELETE FROM tmpintefrafracli WHERE codusu = " & vUsu.Codigo
    'errores
    Conn.Execute "DELETE FROM tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
    
    NF = FreeFile
    Open Text1.Text For Input As #NF
    
    
    ''tmpintefrafracli(codusu,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,
    'tipo_ret,ctaventas,Ccoste,impventa,iva,impiva,recargo,imprecargo,totalfactura,integracion)
    NumRegElim = 0
    J = 0 'Numero de error en el total de fichero. Secuencial

    FacturaAnterior = ""
    While Not EOF(NF)
        NumRegElim = NumRegElim + 1
        Line Input #NF, Cad
        
        If NumRegElim = 1 Then
            'Primera linea encabezado?
            If Me.Check1.Value = 1 Then Cad = ""
        Else
            If InStr(1, String(NumeroCamposTratarFraCli, ";"), Cad) > 0 Then Cad = "" 'todo puntos y comas
        End If
        
        
        If Cad <> "" Then
            'Procesamos linea
            
            strArray = Split(Cad, ";")
            
            If UBound(strArray) = NumeroCamposTratarFraCli - 1 Then
                'Falta el ultimo punto y coma
                Cad = Cad & ";"
                strArray = Split(Cad, ";")
            End If
            
            
            If UBound(strArray) <> NumeroCamposTratarFraCli Then
                J = J + 1
                Aux = vUsu.Codigo & "," & J & "," & NumRegElim & ",'Nº campos incorrecto'," & DBSet(Cad, "T")
                Conn.Execute "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,observa1) VALUES (" & Aux & ")"
                
            Else
                'En la sub o insertara en la tabla de correctos o en la de errores ,
                'en funcion de los tipos de datos y que sean requeridos o no
                SeparaLineaCliente
            
            
            End If
            
            
            
            
        End If
        
    Wend
    Close #NF
    If NumRegElim = 0 Then
        MsgBox "Fichero vacio", vbExclamation
    Else
        Cad = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
        If Val(Cad) > 0 Then
            ImportarFichFracli = 1 'Con errores
        Else
            Cad = DevuelveDesdeBD("count(*)", "tmpintefrafracli", "codusu", vUsu.Codigo)
            If Val(Cad) > 0 Then ImportarFichFracli = 0
        End If
    End If
    
eImportarFicheroFracli:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        CerrarFichero
    End If
End Function


Private Sub SeparaLineaCliente()
Dim NuevaLinea As Boolean

    CadenaDesdeOtroForm = "INSERT INTO tmpintefrafracli(codusu,Codigo,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ",tipo_ret,ctaventas,Ccoste,impventa,iva,impiva,recargo,imprecargo,totalfactura,integracion,iban,txtcsb"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ") VALUES (" & vUsu.Codigo & "," & NumRegElim
    

    'Vemos los campos.
    ' 0       1         2       3           4           5               6           7       8       9           10
    'SERIE  FACTURA  FECHA  CTA. CLI.    F.PAGO    TIPO OPERACION   CTA.RET.   IMP.RET   TIPO RET. CTA.VENTAS   CCOST
    '  11       12          13         14       15          16              17          18       1,9
    'IMP.VENTA  % I.V.A.  MP.IVA    % R.E.    IMP. REC.   TOTAL FACTURA   INTEGRACION   IBAN    txtcsb
    LineaConErrores = False
    
   
    
    
    
    'Precomprobacion (RAFA)
    'Puede ser que repita el numero,serie,fecha para la misma factura(con mas de una linea), ya que si se lanza el proceso
    'que genera el fichero csv es mas comodo poner todos los datos aunque se repitan
    'Por lo tanto , aqui, si los datos de factuas son los mismo, los reseteo
    NuevaLinea = True
    If FacturaAnterior <> "" Then
        If strArray(0) <> "" Or strArray(1) <> "" Or strArray(2) <> "" Then
            'Ha puesto datos de facturas
            If strArray(0) = RecuperaValor(FacturaAnterior, 1) Then
                If strArray(1) = RecuperaValor(FacturaAnterior, 2) Then
                    If strArray(2) = RecuperaValor(FacturaAnterior, 3) Then
                        'Es la misma que la de arriba. Pongo los campos a vacio
                        strArray(0) = ""
                        strArray(1) = "0"
                        strArray(2) = "1900-01-01"
                        strArray(3) = ""
                        
                        
                        
                        NuevaLinea = False
                    End If
                End If
            End If
        End If
    End If
    If NuevaLinea Then
         If strArray(0) <> "" And strArray(1) <> "" And strArray(2) <> "" Then FacturaAnterior = strArray(0) & "|" & strArray(1) & "|" & strArray(2) & "|"
    End If
    
    'Validamos campos
    For K = 0 To NumeroCamposTratarFraCli - 1
        ValidarLinea CByte(K)
    Next K
    
    
    If Not LineaConErrores Then
        
        'Si pone algun dato de la factura, debe ponerlos todos
        If strArray(0) <> "" Or strArray(1) <> "0" Or strArray(2) <> "1900-01-01" Or strArray(3) <> "" Then
            If strArray(0) = "" Or strArray(1) = "0" Or strArray(2) = "1900-01-01" Or strArray(3) = "" Then
                AnyadeEnErrores "Campos facturas. Todos o ninguno"
            End If
        End If
        
       ' 0 = TODO    1 = CONTA  2= TESORERIA
        If Val(strArray(17)) <> 1 Then
            'Queremos que meta el cobro
            If strArray(4) = "" Then
                AnyadeEnErrores "Falta forma de pago"
            End If
            
            
            
            
        End If
        
        
        'Si indica retencion tiene que indicar el tipo y bicepsversa
        
        If Val(strArray(8)) > 0 Then
            If CCur(strArray(7)) = 0 Then AnyadeEnErrores "Error importe retencion FALTA"
            If strArray(6) = "" Then AnyadeEnErrores "Error cuenta retencion FALTA"
        Else
            
                If Val(strArray(7)) > 0 Then AnyadeEnErrores "Error tipo retencion indicado "
                If strArray(6) <> "" Then AnyadeEnErrores "Error cuenta retencion indicada"

        End If
        
        
        
        
        
        
        
    End If
    
    
    If Not LineaConErrores Then
    
        
    
    
        'INSERTAMOS EN tmpintefrafracli
        Conn.Execute CadenaDesdeOtroForm & ")"
        
    
    End If
    
End Sub


Private Sub ValidarLinea(QueCampo As Byte)
Dim ValorSQL As String

    'Vemos los campos.
    ' 0       1         2       3           4           5               6           7       8       9
    'SERIE  FACTURA  FECHA  CTA. CLI.    F.PAGO    TIPO OPERACION   CTA.RET.   IMP.RET   TIPO RET. CTA.VENTAS
    '10      11       12          13         14       15          16              17          18          19
    'CC    IMP.VENTA  % I.V.A.  MP.IVA    % R.E.    IMP. REC.   TOTAL FACTURA   INTEGRACION  iban     txtcsb
    ValorSQL = "NULL"
    Select Case QueCampo
    'Numerico REQUERIDO
    Case 9, 11, 12, 13
      
        'RAFA. Pone un . decimal
        strArray(QueCampo) = Replace(strArray(QueCampo), ".", ",")
        
      
        If strArray(QueCampo) = "" Then
            AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " no puede estar vacio"
        Else
            If Not IsNumeric(strArray(QueCampo)) Then
                AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
            Else
                ValorSQL = DBSet(strArray(QueCampo), "N")
            End If
                
        End If
    
    
    
        
    
    
    
    'Numerico NO requerido
    Case 1, 4, 7, 14, 15, 16
        
        'RAFA. Pone un . decimal
        strArray(QueCampo) = Replace(strArray(QueCampo), ".", ",")
        
        If QueCampo = 1 Then
            If strArray(QueCampo) = "" Then strArray(QueCampo) = 0
            
        End If
        
        If QueCampo = 1 Or QueCampo = 14 Then
            'No es requerdio pero si es cero grabo cero
            ValorSQL = DBSet(strArray(QueCampo), "N", "N")
        Else
            ValorSQL = DBSet(strArray(QueCampo), "N", "S")
        End If
        If strArray(QueCampo) <> "" Then
            If Not IsNumeric(strArray(QueCampo)) Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
        End If
    
    'Fecha
    Case 2
        
        If strArray(QueCampo) = "" Then strArray(QueCampo) = "1900-01-01"
        
        If Not IsDate(strArray(QueCampo)) Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo fecha"
        ValorSQL = DBSet(strArray(QueCampo), "F", "S")
        
    Case 5, 8, 17
        'Tipo operacion   QUECAMPO 5
        '        0 = GENERAL
        '        1 = INTRACOMUNITARIA
        '        2 = EXPORTACION
        '        3 = INTERIOR EXENTA
        ' Tipo retencion  QUECAMPO 8
        '        0 o nada, NO TIENE
        '        1 = PROFESIONAL
        '        2 = AGRICOLA
        '        3 = ARRENDAMIENTO
        'Tipo integracion
        '        0 = TODO
        '        1 = CONTA
        '        2 = TESORERIA
    
        If strArray(QueCampo) = "" Then
            strArray(QueCampo) = "0"
            If QueCampo = 17 Then strArray(QueCampo) = "1"
        End If
        If Not IsNumeric(strArray(QueCampo)) Then
            AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
        Else
            strArray(QueCampo) = Abs(Val(strArray(QueCampo)))
            If QueCampo = 17 Then
                If Val(strArray(QueCampo)) > 2 Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " Valores: 0..2 o vacio"
            Else
                If Val(strArray(QueCampo)) > 3 Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " Valores: 0..3 o vacio"
            End If
       End If
       ValorSQL = DBSet(strArray(QueCampo), "N", "N")
    Case 10
        'Centro de coste
        'Si lleva analitica es obligatorio
        If strArray(QueCampo) = "" Then
            If vParam.autocoste Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " no puede estar vacio"
        Else
            ValorSQL = DBSet(strArray(QueCampo), "T")
        End If
            
        
        
    Case 0
        'Porque no puede ser NULO, pondremos un ''
        ValorSQL = DBSet(strArray(QueCampo), "T", "N")
    Case Else
        ValorSQL = DBSet(strArray(QueCampo), "T", "S")
    End Select
    
    
    
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "," & ValorSQL
    
End Sub


Private Sub AnyadeEnErrores(QueError As String)

    J = J + 1
    QueError = vUsu.Codigo & "," & J & "," & NumRegElim & "," & DBSet(QueError, "T") & "," & DBSet(Cad, "T")
    Conn.Execute "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,observa1) VALUES (" & QueError & ")"
    LineaConErrores = True
End Sub



Private Function ComprobacionDatosBDFacturas() As Boolean
Dim FechaMinima As Date

On Error GoTo eComprobacionDatosBD
    ComprobacionDatosBDFacturas = False
    Set miRsAux = New ADODB.Recordset
    
    'De momento solo clientes
    If Me.optVarios(0).Value Then
        Cad = "select min(fecha) minima,max(fecha) from tmpintefrafracli where factura >0 and codusu=" & vUsu.Codigo
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'no puede ser eof
        Cad = miRsAux!minima
        FechaMinima = miRsAux!minima
        If miRsAux!minima < vParam.fechaini Then
            AnyadeEnErrores "Menor que inicio ejercicio"
        Else
            If miRsAux!minima < vParam.FechaActiva Then
                AnyadeEnErrores "Menor que fecha activa"
            Else
                If miRsAux!minima <= UltimoDiaPeriodoLiquidado Then
    
                    If Me.optVarios(0).Value Then
                        'FACTURAS CLIENTE. Obliado comprobar
                        AnyadeEnErrores "Menor que ultimo periodo liquidado"
                    Else
    
                        'PROVEEDOR. Va a pedir fecha liquidacion luego. No hacemos nada"
    
    
                    End If
                End If
            End If
        End If
        miRsAux.Close
    End If
    
    'Comprobaremos que todas las SERIES estan en contadores
    Cad = "select distinct(serie) from tmpintefrafracli where serie<>''  and codusu=" & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = DevuelveDesdeBD("tiporegi", "contadores", "tiporegi", miRsAux.Fields(0), "T")
        If Cad = "" Then
            Cad = miRsAux.Fields(0)
            AnyadeEnErrores "No existe contadores"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Cuentas
    Cad = "select distinct cta_cli from tmpintefrafracli where cta_cli<>''  and codusu=" & vUsu.Codigo & " and not cta_cli in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cta cliente"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
        
    Cad = "select distinct ctaventas from tmpintefrafracli where ctaventas<>''  and codusu=" & vUsu.Codigo & " and not ctaventas in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cuenta ventas"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    'Retencion (si lleva
    If Me.optVarios(0).Value Then
        'Cuenta clientes SI pide trae la cta reten
        Cad = "select distinct ctaret from tmpintefrafracli where ctaret<>''  and codusu=" & vUsu.Codigo & " and not ctaret in (select codmacta from cuentas where apudirec='S')"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = miRsAux.Fields(0)
            AnyadeEnErrores "No existe cuenta ventas"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
         
    End If
         
         
     
    
    'Facturas que ya existen
    If Me.optVarios(1).Value Then
        Cad = "select factura,cta_cli from tmpintefrafracli where codigo>0  and codusu=" & vUsu.Codigo
        Cad = Cad & " and (serie,factura,year(fecha),cta_cli ) iN (select numserie,numfactu,anofactu,codmacta from factpro where anofactu>=" & Year(FechaMinima) & " )"
    Else
        Cad = "select factura,cta_cli as cta_cli  from tmpintefrafracli where codigo>0  and codusu=" & vUsu.Codigo
        Cad = Cad & " and (serie,factura,year(fecha)) iN (select numserie,numfactu,anofactu from factcli where anofactu>=" & Year(FechaMinima) & " )"
    End If
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!FACTURA & "  " & miRsAux!cta_cli
        AnyadeEnErrores "Ya existe factura"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    Cad = " select FACTURA,fpago from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
    Cad = Cad & " and not fpago iN (select codforpa from formapago )"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!FACTURA
        AnyadeEnErrores "No existe forma de pago " & miRsAux!fpago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si va a meter en cobros, que no existan en la tesoreria
    If Me.optVarios(1).Value Then
        Cad = "select factura from tmpintefrafracli where codigo>0  and codusu=" & vUsu.Codigo
        Cad = " select serie,factura,fecha from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
        Cad = Cad & " and  (serie,factura,fecha) iN (select numserie, numfactu, fecfactu from pagos where fecfactu>=" & DBSet(FechaMinima, "F") & " )"
    Else
        Cad = " select serie,factura,fecha from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
        Cad = Cad & " and  (serie,factura,fecha) iN (select numserie, numfactu, fecfactu from cobros )"
    End If
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!FACTURA
        AnyadeEnErrores "Ya existe vencimiento"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    If vParam.autocoste Then
        Cad = "select ccoste from tmpintefrafracli where ccoste<>''  and codusu=" & vUsu.Codigo
        Cad = Cad & " AND  not ccoste iN (select codccost from ccoste )"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = miRsAux!FACTURA
            AnyadeEnErrores "No existe centro de coste: " & miRsAux!CCoste
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    'Porcentaje IVA
    'Facturas que ya existen
    Cad = "select iva,recargo from tmpintefrafracli where codusu=" & vUsu.Codigo & " GROUP BY 1,2     "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = DBLet(miRsAux!recargo, "N") 'En tiposiva no puede ser null
        Cad = " tipodiva=0 and porcerec=" & Cad & " AND porceiva"
        Cad = DevuelveDesdeBD("codigiva", "tiposiva", Cad, DBSet(miRsAux!IVA, "N"))
        If Cad = "" Then
            Cad = "IVa: " & miRsAux!IVA & "% Rec: " & DBLet(miRsAux!recargo, "T") & "%"
            AnyadeEnErrores "No existe IVA"
            
        Else
        
                Cad = "UPDATE tmpintefrafracli SET tipoiva=" & Cad & " WHERE codusu =" & vUsu.Codigo
                Cad = Cad & " AND iva =" & DBSet(miRsAux!IVA, "N") & " AND recargo"
                If IsNull(miRsAux!recargo) Then
                    Cad = Cad & " is null"
                Else
                    Cad = Cad & "=" & DBSet(miRsAux!recargo, "N")
                End If
                Conn.Execute Cad
        
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si esta duplicado la factura
    Cad = "select serie,factura,count(*) from tmpintefrafracli where factura>0 and codusu =" & vUsu.Codigo & " group by 1,2 having count(*)>1"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = "Factura: " & miRsAux!FACTURA & "   Veces: " & miRsAux.Fields(1)
        AnyadeEnErrores "Datos duplicados"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
                



    
    'Utlima comprobacion
    If Me.optVarios(1).Value Then
        'Proveedores navarres
        If Me.cboTipo.ListIndex = 0 Then
            Cad = "select distinct iva from tmpintefrafracli where codusu=" & vUsu.Codigo
            miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad = ""
            While Not miRsAux.EOF
                Cad = Cad & "X"
                miRsAux.MoveNext
                
            Wend
            miRsAux.Close
            If Len(Cad) > 1 Then
                Cad = "REA"
                AnyadeEnErrores "Mas de un IVA en el fichero"
            Else
                Cad = "UPDATE tmpintefrafracli SET tipo_operacion =5"
                Cad = Cad & " where codusu=" & vUsu.Codigo
                Conn.Execute Cad
            End If
            'Navarres deben ser todo LIQUIDACIONES REA
            'tipo_operacion=5
            
        End If
    End If





    'Si no han habiado errores..  j=0
    ComprobacionDatosBDFacturas = J = 0


    

eComprobacionDatosBD:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function

Private Function ComprobarTotales() As Boolean
Dim FACTURA As String
Dim BaseImponible As Currency
Dim IVA As Currency
Dim TotalFac As Currency
Dim ImporteRetencion As Currency
Dim Importe As Currency
Dim TipoRetencion As Byte
Dim Ok As Boolean
Dim DiferenciaMinimaPermitida As Currency
Dim FacturasProveedorSTD As Boolean
Dim SegundoImporteAuxiliar As Currency
Dim B As Boolean
Dim EsMismaFactura As Boolean

    On Error GoTo eComprobarTotales
    ComprobarTotales = False


    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    DiferenciaMinimaPermitida = 0
    If Me.optVarios(1).Value Then
        'Proveedores navarres
        DiferenciaMinimaPermitida = 0.1
        If Me.cboTipo.ListIndex = 0 Then DiferenciaMinimaPermitida = 0.2
    Else
        DiferenciaMinimaPermitida = 0.2  'Pero sera para el calculo de IVA. Para el total factura NO hay margen
    End If
    
    Cad = "select * from tmpintefrafracli WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Dim Fin As Boolean
    Fin = False
    
    FacturasProveedorSTD = False
    If Me.optVarios(1).Value Then
        'Clientes
        If Me.cboTipo.ListIndex = 1 Then FacturasProveedorSTD = True
    
    End If
    
    While Not miRsAux.EOF
    
       
    
    
        EsMismaFactura = False
        If FacturasProveedorSTD Then
            If DBLet(miRsAux!FACTURA, "T") = "" Then EsMismaFactura = True
        Else
            If miRsAux!Serie = "" Then EsMismaFactura = True
        End If
        If EsMismaFactura Then
            'Es otra base de la misma factura
            
        Else
        
        
            If Not FacturasProveedorSTD Then
                Cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
            Else
                Cad = miRsAux!FACTURA 'Alfanumerico
            End If
            If Cad <> FACTURA Then
                If FACTURA <> "" Then
                    'Suma correcta
                    SegundoImporteAuxiliar = TotalFac - (BaseImponible + IVA - ImporteRetencion)
                    Ok = True
                    If optVarios(0).Value Then
                        'no  hay margen.
                        If Abs(SegundoImporteAuxiliar) > 0 Then Ok = False
                    Else
                        'Proveedores
                        If Abs(SegundoImporteAuxiliar) > DiferenciaMinimaPermitida Then Ok = False
                    End If
                   
                    Cad = "Calculado/fichero:  " & BaseImponible + IVA - ImporteRetencion & " / " & TotalFac
                    If ImporteRetencion <> 0 Then Cad = Cad & " Retencion Tipo" & TipoRetencion & "  Importe: " & ImporteRetencion
                    
                    If Not Ok Then
                        NumRegElim = Mid(FACTURA, 4)
                        AnyadeEnErrores "Total factura"
                        
                    Else
                        If FacturasProveedorSTD Then
                            Cad = " AND FACTURA =" & DBSet(FACTURA, "T")
                        Else
                            Cad = " AND Serie= " & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & " AND FACTURA =" & Mid(FACTURA, 4)
                        End If
                        Cad = "UPDATE tmpintefrafracli set CalculoImponible=" & DBSet(BaseImponible, "N") & "  WHERE codusu = " & vUsu.Codigo & Cad
                        Conn.Execute Cad
                    End If
                    If Not FacturasProveedorSTD Then
                        Cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
                    Else
                        Cad = miRsAux!FACTURA 'Alfanumerico
                    End If
                End If
                'Factura NUEVA. Reseteamos importes...
                TipoRetencion = DBLet(miRsAux!tipo_ret, "N")
                ImporteRetencion = 0
                If TipoRetencion <> 0 Then ImporteRetencion = DBLet(miRsAux!impret, "N")
                'NumRegElim = miRsAux!FACTURA
                BaseImponible = 0
                IVA = 0
                TotalFac = miRsAux!TotalFactura
                
                FACTURA = Cad
                
            End If
        End If
        
        Importe = (miRsAux!impventa * miRsAux!IVA) / 100
        Importe = Round2(Importe, 2)
        SegundoImporteAuxiliar = Importe - miRsAux!ImpIva
        B = False
        If Abs(SegundoImporteAuxiliar) > DiferenciaMinimaPermitida Then
            'St op
            B = True
        End If
        
        If B Then
            NumRegElim = miRsAux!Codigo
            Cad = "IVA calculado/fichero " & FACTURA & " :" & miRsAux!IVA & "%. " & Importe & " / " & miRsAux!ImpIva
            AnyadeEnErrores "Calculo  IVA"
        
            
        End If
        Importe = miRsAux!ImpIva
        BaseImponible = BaseImponible + miRsAux!impventa
        IVA = IVA + Importe
        
            
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If FACTURA <> "" Then
        'Suma correcta
        SegundoImporteAuxiliar = TotalFac - (BaseImponible + IVA - ImporteRetencion)
        B = Abs(SegundoImporteAuxiliar) > DiferenciaMinimaPermitida
        If B Then
            NumRegElim = Val(Mid(FACTURA, 4))
            Cad = "Calculado/fichero:  " & BaseImponible + IVA - ImporteRetencion & " / " & TotalFac
            AnyadeEnErrores "Total factura"
        Else
            If FacturasProveedorSTD Then
                Cad = " AND FACTURA =" & DBSet(FACTURA, "T")
            Else
                Cad = " AND Serie= " & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & " AND FACTURA =" & Mid(FACTURA, 4)
            End If
            Cad = "UPDATE tmpintefrafracli set CalculoImponible=" & DBSet(BaseImponible, "N") & "  WHERE codusu = " & vUsu.Codigo & Cad
            Conn.Execute Cad
        End If
    End If
    
    
    
    
    
    
    
    
    'Si no hay error
    ComprobarTotales = J = 0
    
    
eComprobarTotales:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function



Private Sub CerrarFichero()
    On Error Resume Next
    Close #NF
    Err.Clear
    
End Sub


Private Sub InsertarEnContabilidadFraCliSAGE()
Dim Mc As Contadores
Dim Actual As Boolean
Dim Ok As Boolean

    Set miRsAux = New ADODB.Recordset
    Cad = "Select   numasien ,fechaent from tmpintegrapu  where codusu =" & vUsu.Codigo & " GROUP BY 1,2  ORDER BY numasien"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set Mc = New Contadores
    NumRegElim = -1
    While Not miRsAux.EOF
        
        Mc.ConseguirContador "0", miRsAux!FechaEnt <= vParam.fechafin, False
        If NumRegElim < 0 Then
            NumRegElim = Mc.Contador
            Actual = miRsAux!FechaEnt <= vParam.fechafin
        End If
        Cad = "UPDATE tmpintegrapu SET numasien=" & Mc.Contador & " WHERE codusu =" & vUsu.Codigo & " AND numasien=" & miRsAux!NumAsien
        Conn.Execute Cad
        
        'Actualizamos el nuasien de las facturas
        Cad = "UPDATE tmpintefrafracli  SET txtcsb=" & Mc.Contador & " WHERE codusu =" & vUsu.Codigo & " AND txtcsb=" & miRsAux!NumAsien
        Conn.Execute Cad
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    espera 0.5
        
    Conn.BeginTrans
    Ok = TraspasaApuSAGE
    PonerLabel ""

    
    If Ok Then
        Conn.CommitTrans
        J = 0 'Sin errores
    Else
        Conn.RollbackTrans
        Cad = "Contado1"
        If Not Actual Then Cad = "Contado2"
        Cad = "UPDATE contadores set " & Cad & " = " & NumRegElim & " WHERE  tiporegi='0'"
        Conn.Execute Cad
    End If
End Sub


Private Sub InsertarEnContabilidadFraCliStd()
Dim FACTURA As String
Dim BaseImponible As Currency
Dim IVA As Currency
Dim TotalFac As Currency
Dim ColBases As Collection
Dim Fecha As Date
Dim Aux As String
Dim ColCob As Collection
Dim tCobro As String
'Tipo integracion   Viene en el fichero en la posicion 17
'        0 = TODO
'        1 = CONTA
'        2 = TESORERIA
Dim Tipointegracion As Byte

    Set miRsAux = New ADODB.Recordset
    
    Label1.Caption = "Insertando contabilidad"
    Label1.Refresh
    
    
    'Valores por defecto
    Cad = DevuelveDesdeBD("codigo", "agentes", "1", "1 ORDER BY codigo")
    CadenaDesdeOtroForm = Cad & "|"
    Cad = DevuelveDesdeBD("codforpa", "formapago", "1", "1 ORDER BY codforpa")
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad & "|"
                        
    tCobro = "select ctabanc1,count(*) from cobros where fecfactu>=" & DBSet(DateAdd("y", -1, Now), "F") & " group by 1 order by 2 desc"
    miRsAux.Open tCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    tCobro = ""
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then tCobro = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
    If tCobro = "" Then
        tCobro = DevuelveDesdeBD("codmacta", "bancos", "1", "1 ORDER BY 1 desc")
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & tCobro & "|"
    tCobro = ""
                        
    'CadenaDesdeOtroForm = agente|codforpa|"
    
    
    'codusu,Codigo,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,tipo_ret,ctaventas,Ccoste,impventa,
    'IVA , ImpIva, recargo, imprecargo, TotalFactura, integracion
    
    Cad = "select tmpintefrafracli.*,codpais,nifdatos,desprovi,despobla,codposta,dirdatos,nommacta,cuentas.iban ibancta from tmpintefrafracli left join cuentas ON cta_cli=codmacta"
    Cad = Cad & " WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotProgres = 0
    While Not miRsAux.EOF
        TotProgres = TotProgres + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    FACTURA = ""
    While Not miRsAux.EOF
    
        numProgres = numProgres + 1
        Label1.Caption = "Linea " & numProgres & "/" & TotProgres
        Label1.Refresh
    
    
    
        If miRsAux!Serie = "" Then
            'Es otra base de la misma factura
            
        Else
            Cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
            
            If Cad <> FACTURA Then
                If FACTURA <> "" Then
                   
                    InsertarFACTURA ColBases, FACTURA, Fecha, Tipointegracion, ColCob
                    
                    'Suma correcta
                    Cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
                End If
                'Factura NUEVA.
                
                'Haremos los inserts
                'INSERT INTO factcli (codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,totfaccl,trefaccl,totbasesret,totivas,
                'totbases,codagente,dpto,numasien,fecliqcl,retfaccl,cuereten,tiporeten,codopera,observa,numserie,numfactu,fecfactu,codmacta,
                'codforpa,anofactu)
                Tipointegracion = miRsAux!integracion
                
                
                NumRegElim = miRsAux!FACTURA
                BaseImponible = miRsAux!calculoimponible
                TotalFac = miRsAux!TotalFactura
                IVA = TotalFac - BaseImponible + DBLet(miRsAux!impret, "N")
                Msg$ = "(" & DBSet(miRsAux!codpais, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T", "S") & "," & DBSet(miRsAux!desProvi, "T", "S") & ","
                Msg$ = Msg$ & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S") & "," & DBSet(miRsAux!dirdatos, "T", "S") & ","
                'nommacta totfaccl,trefaccl,totbasesret,totivas,totbases,codagente,dpto,numasien,
                Msg$ = Msg$ & DBSet(miRsAux!Nommacta, "T") & "," & DBSet(miRsAux!TotalFactura, "N") & "," & DBSet(miRsAux!impret, "N") & "," & DBSet(BaseImponible, "N")
                Msg$ = Msg$ & "," & DBSet(IVA, "N") & "," & DBSet(BaseImponible, "N") & ","
                Msg$ = Msg$ & RecuperaValor(CadenaDesdeOtroForm, 1) & ",NULL,NULL," & DBSet(miRsAux!Fecha, "F") & ","
                If DBLet(DBLet(miRsAux!tipo_ret, "N"), "N") = 0 Then
                    'No lleva retencion
                    Msg$ = Msg$ & "NULL,NULL,0"
                Else
                    Msg$ = Msg$ & DBSet(miRsAux!impret, "N", "S") & ",'" & miRsAux!ctaret & "'," & DBLet(miRsAux!tipo_ret, "N")   'tiporeten=numfactura
                End If
                Msg$ = Msg$ & "," & DBLet(miRsAux!tipo_operacion, "N") & ","
                Msg$ = Msg$ & "'Importacion datos externos. " & Chr(13) & "Usuario: " & DevNombreSQL(vUsu.Nombre) & Chr(13) & "Fecha: " & Now
                Msg$ = Msg$ & "'," & DBSet(miRsAux!Serie, "T") & "," & miRsAux!FACTURA & "," & DBSet(miRsAux!Fecha, "F") & "," & DBSet(miRsAux!cta_cli, "T") & ","
                If DBLet(miRsAux!fpago, "T") = "" Then
                    'No viene forma de pago
                     K = RecuperaValor(CadenaDesdeOtroForm, 2)
                Else
                     K = miRsAux!fpago
                End If
                Msg$ = Msg$ & K
                Msg$ = Msg$ & "," & Year(miRsAux!Fecha) & ")"
                
                

                'select nommacta,dirdatos,despobla,codposta,desprovi,codpais,nifdatos from cuentas
                'Para los cobros     (cojo esta variable)
                '       insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,
                'numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,text41csb,numorden,fecvenci,impvenci)
                
                If Tipointegracion <> 1 Then
                    tCobro = " (" & RecuperaValor(CadenaDesdeOtroForm, 1) & "," & DBSet(miRsAux!Nommacta, "T") & "," & DBSet(miRsAux!dirdatos, "T", "S")
                    tCobro = tCobro & "," & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S")
                    tCobro = tCobro & "," & DBSet(miRsAux!desProvi, "T", "S") & "," & DBSet(miRsAux!codpais, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T")
                    tCobro = tCobro & "," & DBSet(miRsAux!Serie, "T") & "," & miRsAux!FACTURA & "," & DBSet(miRsAux!Fecha, "F") & "," & DBSet(miRsAux!cta_cli, "T")
                    tCobro = tCobro & "," & K & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & ","
                    
                    'Si me ha indicado IBAN, o no, en la importacion
                    If DBLet(miRsAux!IBAN, "T") = "" Then
                        Aux = DBSet(miRsAux!ibancta, "T", "S")
                    Else
                        Aux = DBSet(miRsAux!IBAN, "T", "S")
                    End If
                    tCobro = tCobro & Aux & ","
                    
                    'Si me ha indicado(o no) txtobserva text33csb text41csb
                    If DBLet(miRsAux!txtcsb, "T") = "" Then
                        Aux = DBSet("Factura : " & miRsAux!FACTURA & " de " & miRsAux!Fecha, "T") & ",null,"
                    Else
                        Aux = miRsAux!txtcsb
                        If Len(Aux) > 80 Then
                            Aux = DBSet(Mid(Aux, 1, 80), "T") & "," & DBSet(Mid(Aux, 81), "T") & ","
                        Else
                            Aux = DBSet(Aux, "T") & ",NULL,"
                        End If
                    End If
                    tCobro = tCobro & Aux
                    
                    CargarCobrosSobreCollectionConSQLInsert ColCob, CStr(K), miRsAux!Fecha, TotalFac, tCobro
                End If
                
                
                
                
                
                
                
                
                Fecha = miRsAux!Fecha
                i = 0
                FACTURA = Cad
                Set ColBases = Nothing
                Set ColBases = New Collection
                
            End If
        End If
        
        'Lineas
        'INSERT INTO factcli_lineas (aplicret,imporec,anofactu,codccost,
        'impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea,fecfactu,numserie,numfactu
        i = i + 1
        Cad = "(0," & DBSet(miRsAux!recargo, "T") & "," & Year(Fecha) & "," & DBSet(miRsAux!CCoste, "T", "S") & ","
        Cad = Cad & DBSet(miRsAux!ImpIva, "N") & "," & DBSet(miRsAux!recargo, "N", "S") & "," & DBSet(miRsAux!IVA, "N")
        Cad = Cad & "," & DBSet(miRsAux!impventa, "N", "N") & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux!ctaventas, "T") & "," & i
        Cad = Cad & "," & DBSet(Fecha, "F", "N") & "," & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & "," & DBSet(Mid(FACTURA, 4), "N") & ")"
        
        ColBases.Add Cad
            
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If FACTURA <> "" Then
        InsertarFACTURA ColBases, FACTURA, Fecha, Tipointegracion, ColCob
    End If
    
     Msg$ = ""
     Ampliacion = ""
    
    
    Label1.Caption = ""
    Label1.Refresh
    
    
    
End Sub




'Tipo integracion
'        0 = TODO
'        1 = CONTA
'        2 = TESORERIA
Private Function InsertarFACTURA(ByRef C As Collection, FacturaC As String, Fecha As Date, Tipointegracion As Byte, ByRef Cobros As Collection) As Boolean
Dim B As Boolean
 Dim RT As ADODB.Recordset
    On Error GoTo eInsertarFACTURA
    
    
    InsertarFACTURA = False
    Conn.BeginTrans

    If Tipointegracion <= 1 Then

        'Ejecutamos los sql
        If Me.optVarios(1).Value Then
            Cad = "INSERT INTO factpro(codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,"
            Cad = Cad & "codconce340,codopera,codmacta,anofactu,codforpa,totbases,totbasesret,"
            Cad = Cad & "totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten"
            Cad = Cad & ",observa ,NUmSerie , Numregis, fecharec, NumFactu, FecFactu,fecliqpr) VALUES " & Msg$
            
        Else
            'Cliente
            Cad = "INSERT INTO factcli (codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,totfaccl,trefaccl,totbasesret,totivas,"
            Cad = Cad & "totbases,codagente,dpto,numasien,fecliqcl,retfaccl,cuereten,tiporeten,codopera,observa,numserie,numfactu,fecfactu,codmacta,"
            Cad = Cad & "codforpa,anofactu) VALUES " & Msg$
        End If
        Conn.Execute Cad
        
        'Lineas
        Cad = ""
        For i = 1 To C.Count
            Cad = Cad & ", " & C.Item(i)
        Next i
        
       
        If Me.optVarios(1).Value Then
            Cad = "fecharec,numserie,numregis) VALUES " & Mid(Cad, 2)
            Cad = "INSERT INTO factpro_lineas (aplicret,imporec,anofactu,codccost,impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea," & Cad
        Else
             Cad = "fecfactu,numserie,numfactu) VALUES " & Mid(Cad, 2)
            Cad = "INSERT INTO factcli_lineas (aplicret,imporec,anofactu,codccost,impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea," & Cad
        End If
        Conn.Execute Cad
        
        'Totales
        espera 0.3
    
        
        Cad = " select codigiva, porciva, porcrec, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec "
        If Me.optVarios(1).Value Then
            Cad = Cad & " from factpro_lineas"
            Cad = Cad & " where numserie = " & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & " and numregis = " & NumRegElim & " and anofactu = " & Year(Fecha)
        Else
            Cad = Cad & " from factcli_lineas"
            Cad = Cad & " where numserie = " & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & " and numfactu = " & DBSet(Mid(FacturaC, 4), "N") & " and anofactu = " & Year(Fecha)
        End If
        Cad = Cad & " group by 1,2,3"
        Cad = Cad & " order by 1,2,3"
        Set RT = New ADODB.Recordset
        RT.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        i = 0
        Cad = ""
        While Not RT.EOF
            i = i + 1
            If Me.optVarios(1).Value Then
                Cad = Cad & ", (" & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & "," & NumRegElim & "," & DBSet(Fecha, "F") & "," & Year(Fecha)
            Else
                Cad = Cad & ", (" & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & "," & DBSet(Mid(FacturaC, 4), "N") & "," & DBSet(Fecha, "F") & "," & Year(Fecha)
            End If
            Cad = Cad & "," & i & "," & DBSet(RT!Baseimpo, "N") & "," & RT!codigiva & "," & DBSet(RT!porciva, "N", "S")
            Cad = Cad & "," & DBSet(RT!porcrec, "N") & "," & DBSet(RT!Imporiva, "N") & "," & DBSet(RT!imporrec, "N") & ")"
            RT.MoveNext  'rT!
        Wend
        RT.Close
        Set RT = Nothing
        Cad = Mid(Cad, 2)
        
         If Me.optVarios(1).Value Then
            Cad = "insert into factpro_totales (numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)  VALUES " & Cad
        Else
            Cad = "insert into factcli_totales (numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)  VALUES " & Cad
        End If
        Conn.Execute Cad
        
    
        'Creamos el apunte
         If Me.optVarios(1).Value Then
            B = IntegrarFactura(Trim(Mid(FacturaC, 1, 3)), NumRegElim, Fecha)
        Else
            B = IntegrarFactura(Trim(Mid(FacturaC, 1, 3)), CLng(Mid(FacturaC, 4)), Fecha)
        End If
        If Not B Then
            Cad = "Creando asiento. "
            AnyadeEnErrores FacturaC
        End If
    
    
    
    End If
    
    
    'Tesoreria
    If Tipointegracion <> 1 Then
        'insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,
        'numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci,ctabanc1,iban,text33csb)
        Cad = ""
        For i = 1 To Cobros.Count
            Cad = Cad & ", " & Cobros.Item(i)
        Next i
        Cad = Mid(Cad, 2)
        
        If Me.optVarios(1).Value Then
            Cad = "INSERT INTO pagos(nomprove,domprove,pobprove,cpprove,proprove,codpais,nifprove,numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text1csb,text2csb,numorden,fecefect,impefect) VALUES " & Cad
        Else
            Cad = "insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,text41csb,numorden,fecvenci,impvenci) VALUES " & Cad
        End If
        Conn.Execute Cad
    End If
    
    'Adelante transaccion
    Conn.CommitTrans
    InsertarFACTURA = True
    'Borramos de la tabla temporal
    
    
    Exit Function
eInsertarFACTURA:
    Cad = Err.Description
    Err.Clear
    Conn.RollbackTrans
    
   
    AnyadeEnErrores FacturaC
    
    
    
End Function


Private Function IntegrarFactura(Serie As String, FACTURA As Long, Fecha As Date) As Boolean
Dim SqlLog As String

    IntegrarFactura = False
    
    
    SqlLog = "Factura " & IIf(Me.optVarios(1).Value, "prov", "cli") & "  : " & Serie & " " & FACTURA & " de fecha " & Fecha

    
    
    
    With frmActualizar
        If Me.optVarios(1).Value Then
            .OpcionActualizar = 8
        Else
            .OpcionActualizar = 6
        End If
        'NumAsiento     --> CODIGO FACTURA
        'NumDiari       --> AÑO FACTURA
        'NUmSerie       --> SERIE DE LA FACTURA
        'FechaAsiento   --> Fecha factura
        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
        .NumFac = FACTURA
        'Variable reutilizada
        'Lleva el año factura
        .NumDiari = Year(Fecha)
        .NUmSerie = Serie
        .FechaAsiento = Fecha
        .FechaAnterior = Fecha
        .SqlLog = SqlLog
        .DentroBeginTrans = True
        .DiarioFacturas = IIf(Me.optVarios(1).Value, vParam.numdiapr, vParam.numdiacl)
        .NumAsiento = 0
        .Show vbModal
        
        If AlgunAsientoActualizado Then IntegrarFactura = True
        
        Screen.MousePointer = vbHourglass
        Me.Refresh
    End With
    
End Function



Private Sub Form_Unload(Cancel As Integer)
    If vParam.PathFicherosInteg = "" Then
        CheckValueGuardar "intetipodoc", CByte(Me.cboTipo.ListIndex)
        
        CheckValueGuardar "intetipodoc1", CByte(Me.Check1.Value)
        
        CheckValueGuardar "intetipodoc2", CByte(Me.Check2.Value)
    End If
End Sub

Private Sub imgppal_Click(Index As Integer)
    With frmppal.cd1
    
        If vParam.PathFicherosInteg <> "" Then .InitDir = vParam.PathFicherosInteg
        
        .FileName = ""
        .CancelError = False
        .Filter = "*.csv|*.csv"
        .FilterIndex = 1
        .DialogTitle = "Importar facturas proveedor"
        .ShowOpen
        If .FileName <> "" Then
            Text1.Text = .FileName
            Command1_Click
        End If
    End With
End Sub


Private Sub ListView1_DblClick()
    If Not vEmpresa.TieneTesoreria Then Exit Sub
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If ListView1.Tag = 0 Then Exit Sub 'El dblclick no hace nada
    
    
    If optVarios(0).Value Then
        If ListView1.SelectedItem.SubItems(11) = "*" Then
            'Han puesto observaciones y/o IBAN
            Cad = "serie =" & DBSet(ListView1.SelectedItem.Text, "T") & " AND  factura =" & ListView1.SelectedItem.SubItems(1) & " AND codusu"
            Cad = DevuelveDesdeBD("concat(coalesce(iban,''),'|',coalesce(txtcsb,''),'|')", "tmpintefrafracli", Cad, CStr(vUsu.Codigo))
            If Cad <> "" Then
                With ListView1.SelectedItem
                    CadenaDesdeOtroForm = "Cliente: " & .SubItems(3) & vbCrLf
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Factura: " & .Text & .SubItems(1) & vbCrLf & vbCrLf
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "IBAN: " & RecuperaValor(Cad, 1) & vbCrLf
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "TxtCSB: " & RecuperaValor(Cad, 2) & vbCrLf
                End With
                MsgBox CadenaDesdeOtroForm, vbInformation
            End If
        End If
    End If
End Sub

Private Sub optVarios_Click(Index As Integer)
   
    CargaCombo Index
    
End Sub

Private Sub CargaCombo(QueOpcion As Integer)

If QueOpcion < 2 Then
        cboTipo.Clear
        
        If QueOpcion = 1 Then
            'Proveedores
            cboTipo.AddItem "Ex. Access(csv)"  'NAVARRES
            
        
            cboTipo.AddItem "Proveedores"
            
            cboTipo.AddItem "CSV"
            
            
            
            
            If InStr(1, UCase(vEmpresa.nomempre), "SALUD") > 0 Then
                cboTipo.ListIndex = 0
            Else
                cboTipo.ListIndex = 1  'De momento esta por defecto
            End If
            
            
        Else
            'CLIENTES
            cboTipo.AddItem "Clientes"
            cboTipo.AddItem "CSV"
            cboTipo.ListIndex = 0
        End If
    End If
     
    Me.Label2.visible = QueOpcion < 2
    Me.cboTipo.visible = QueOpcion < 2

End Sub


'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'
'       Imporatr facturas cliente
'
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************

'--------------------------------------------------------------------
Private Sub ImportacionFraPro()

    If cboTipo.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de proceso", vbExclamation
        Exit Sub
    End If
    
    
    Select Case cboTipo.ListIndex
    Case 1
        ImportacionFacturasProveedor
    Case 2
        ProcesaFicheroClientesSAGE Text1.Text, Label1, Check1.Value = 1
        
    
    Case Else
        ImportacionNavarresFraPro
    End Select
    
End Sub


'--------------------------------------------------------------------
'TIPO NAvarres
'--------------------------------------------------------------------
Private Sub ImportacionNavarresFraPro()
Dim Rc As Byte
Dim B As Boolean

 'Primer paso. Lectura fichero. Comprobacion basica datos
    PonerLabel "Leyendo fichero"
    NumeroCamposTratarFraPro = 12
    Rc = ImportarFichFraPronav
    If Rc = 2 Then Exit Sub
        
    If Rc = 1 Then
        'Errores en fichero
        'Ha habido errores
        CargaEncabezado 0
    
        'Cargamos datos
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto1
            ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
            ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Exit Sub
    End If
    
    
    
    
    'Segundo paso. Comprobacion datos BD (existe serie, ctas, tiposiva....
    'Fichero importado.
    'Comporbaciones BD
    PonerLabel "Comprobando en BD"
    If Not ComprobacionDatosBDFacturas Then
        'Cargaremos errores
        CargaEncabezado 1
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto2
            ListView1.ListItems(i).SubItems(1) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    
    
        Exit Sub
    End If
    
    
    'Ultima comprobacion. Que las facturas suman lo que tienen que sumar
    PonerLabel "Totales"
    If Not ComprobarTotales Then
        CargaEncabezado 2
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto1
            ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
            ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Sub
    End If
    
    

    
    
    'Si llega aqui, es que ha ido bien. Estamos pendientes de aceptar
    CargaEncabezado 3
    
    
    'Que culumna pintara en Base/retencion dependiendo frapro
    B = True
    If Me.optVarios(1).Value Then B = False
    
    Set miRsAux = New ADODB.Recordset
   
    Cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on cta_cli=cuentas.codmacta where codusu= " & vUsu.Codigo
    Cad = Cad & " ORDER BY codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        ListView1.ListItems.Add , "K" & i
        If DBLet(miRsAux!Serie, "T") <> "" Then
            
            ListView1.ListItems(i).Text = miRsAux!Serie
            ListView1.ListItems(i).SubItems(1) = Format(miRsAux!FACTURA, "000000")
            ListView1.ListItems(i).SubItems(2) = miRsAux!Fecha
            
            ListView1.ListItems(i).SubItems(3) = DBLet(miRsAux!Nommacta, "T")
            
            
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!impventa, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            ListView1.ListItems(i).SubItems(7) = " " 'separado
            If B Then
                ListView1.ListItems(i).SubItems(8) = Format(miRsAux!calculoimponible, FormatoImporte)
            Else
                ListView1.ListItems(i).SubItems(8) = Format(DBLet(miRsAux!impret, "N"), FormatoImporte)
            End If
            ListView1.ListItems(i).SubItems(9) = Format(miRsAux!TotalFactura - miRsAux!calculoimponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(10) = Format(miRsAux!TotalFactura, FormatoImporte)
            
            If vEmpresa.TieneTesoreria Then
                Cad = " "
                If DBLet(miRsAux!IBAN, "T") <> "" Or DBLet(miRsAux!txtcsb, "T") <> "" Then Cad = "*"
                ListView1.ListItems(i).SubItems(11) = Cad
            End If
            
        Else
            
            ListView1.ListItems(i).Text = " "
            ListView1.ListItems(i).SubItems(3) = "              IVA " & miRsAux!IVA & "%"
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!impventa, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            For NF = 1 To 10
                If NF < 3 Or NF > 6 Then ListView1.ListItems(i).SubItems(NF) = " "
            Next
            
            If vEmpresa.TieneTesoreria Then ListView1.ListItems(i).SubItems(11) = " "
            
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cmdAceptar.visible = True
    
    
End Sub





Private Function ImportarFichFraPronav() As Byte
Dim Aux As String



    On Error GoTo eImportarFicheroFracli
    ImportarFichFraPronav = 2  'Vacio
    
    NF = -1
    'FacturaSocio   Fecha   Bruto   % IVA   IVA Base IRPF   IRPF    Fondo   Cobrado Cosecha
    Msg$ = "Factura|Socio|Fecha|Bruto|Porcentaje IVA|Importe IVA|BrutoMasIva|Retencion|ImpRetenc|Fondo|Total|Cosecha|"
  
    
    
    'Preparamos tabla de insercion
    Conn.Execute "DELETE FROM tmpintefrafracli WHERE codusu = " & vUsu.Codigo
    'errores
    Conn.Execute "DELETE FROM tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
    
    NF = FreeFile
    Open Text1.Text For Input As #NF
    
    
    ''tmpintefrafracli(codusu,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,
    'tipo_ret,ctaventas,Ccoste,impventa,iva,impiva,recargo,imprecargo,totalfactura,integracion)
    NumRegElim = 0
    J = 0 'Numero de error en el total de fichero. Secuencial

    FacturaAnterior = ""
    While Not EOF(NF)
        NumRegElim = NumRegElim + 1
        Line Input #NF, Cad
        
        If NumRegElim = 1 Then
            'Primera linea encabezado?
            If Me.Check1.Value = 1 Then Cad = ""
        Else
            If InStr(1, String(NumeroCamposTratarFraPro, ";"), Cad) > 0 Then Cad = "" 'todo puntos y comas
        End If
        
        
        If Cad <> "" Then
            'Procesamos linea
            
            strArray = Split(Cad, ";")
            
            If UBound(strArray) = NumeroCamposTratarFraPro - 1 Then
                'Falta el ultimo punto y coma
                Cad = Cad & ";"
                strArray = Split(Cad, ";")
            End If
            
            
            If UBound(strArray) <> NumeroCamposTratarFraPro Then
                J = J + 1
                Aux = vUsu.Codigo & "," & J & "," & NumRegElim & ",'Nº campos incorrecto'," & DBSet(Cad, "T")
                Conn.Execute "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,observa1) VALUES (" & Aux & ")"
                
            Else
                'En la sub o insertara en la tabla de correctos o en la de errores ,
                'en funcion de los tipos de datos y que sean requeridos o no
                SeparaLineaProveedorNavarres
            
            
            End If
            
            
            
            
        End If
        
    Wend
    Close #NF
    If NumRegElim = 0 Then
        MsgBox "Fichero vacio", vbExclamation
    Else
        Cad = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
        If Val(Cad) > 0 Then
            ImportarFichFraPronav = 1 'Con errores
        Else
            Cad = DevuelveDesdeBD("count(*)", "tmpintefrafracli", "codusu", vUsu.Codigo)
            If Val(Cad) > 0 Then ImportarFichFraPronav = 0
        End If
    End If
    
eImportarFicheroFracli:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        CerrarFichero
    End If
End Function



Private Sub SeparaLineaProveedorNavarres()
Dim NuevaLinea As Boolean
Dim Cad As String

    CadenaDesdeOtroForm = "INSERT INTO tmpintefrafracli(codusu,Codigo,serie,factura,cta_cli,fecha,impventa,iva,impiva,CalculoImponible,"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " impret,ctaret,tipo_ret,integracion , IBAN, TotalFactura, txtcsb,ctaventas"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ") VALUES (" & vUsu.Codigo & "," & NumRegElim & ",1" 'serie
    
    'Factura Socio   Fecha   Bruto   % IVA   IVA    Base IRPF   IRPF    Fondo   Cobrado    Cosecha
    '   1   8329    170407  1.044,28    12  125,32  1.169,60    23,41   0      1.146,19    OROGRANDE


    LineaConErrores = False
    NuevaLinea = True
        
    'Si pone algun dato de la factura, debe ponerlos todos
    If strArray(0) = "" Or strArray(1) = "" Or strArray(2) = "" Then AnyadeEnErrores "Campos facturas obligados"
    Cad = strArray(2)
    
    strArray(2) = Mid(Cad, 5, 2) & "/" & Mid(Cad, 3, 2) & "/20" & Mid(Cad, 1, 2)
     
    'If strArray(4) = "" Then AnyadeEnErrores "Falta forma de pago"
              
     
    For K = 0 To NumeroCamposTratarFraPro - 1
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & ","
        
        Select Case K
        Case 2
            'Fecha ...rara
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & DBSet(strArray(K), "F")
        Case 3, 5, 6, 7, 9
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & DBSet(ImporteFormateado(strArray(K)), "N")
            If K = 7 Then
                'Retencion
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & ",'RTEN',2,1"  'Tipo reten y modo integracion 'Tipo integracion
                                                                        '        0 = TODO
                                                                        '        1 = CONTA
                                                                        '        2 = TESORERIA
            End If
        Case 10, 11
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & DBSet(strArray(K), "T")
        Case 1
            'Cta
            '4030XXXXX
            'CadenaDesdeOtroForm = CadenaDesdeOtroForm & DBSet("4001041", "T")
'            If strArray(K) = "8275" Then St op
'            If strArray(K) = "8293" Then St op
'            If strArray(K) = "7529" Then St op
            
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & DBSet("4030" & Format(Trim(strArray(K)), Right("000000", vEmpresa.DigitosUltimoNivel - 4)), "T")
        Case Else
            'Numeros enteros
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & Val(strArray(K))
        End Select
    Next K
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ")"
   
            
        
        
        
        
        
        
    
    
    
    If Not LineaConErrores Then
    
        'INSERTAMOS EN tmpintefrafracli
        If Not Ejecuta(CadenaDesdeOtroForm, True) Then AnyadeEnErrores "Error insertando tmp" & CadenaDesdeOtroForm
    
    End If
    
End Sub






Private Sub InsertarEnContabilidadFraProveedorNAV()
Dim FACTURA As String
Dim BaseImponible As Currency
Dim IVA As Currency
Dim TotalFac As Currency
Dim ColBases As Collection
Dim Fecha As Date
Dim Aux As String
Dim ColCob As Collection
Dim tCobro As String
Dim Mc As Contadores
'Tipo integracion   Viene en el fichero en la posicion 17
'        0 = TODO
'        1 = CONTA
'        2 = TESORERIA
Dim Tipointegracion As Byte
Dim fRecep As Date
Dim Actual As Boolean
    Set miRsAux = New ADODB.Recordset
    
    'CadenaDesdeOtroForm  retencion|fecha|agente|forpa|
    'Valores por defecto
    Cad = DevuelveDesdeBD("codigo", "agentes", "1", "1 ORDER BY codigo")
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad & "|"
    
    
    Cad = "tipforpa=1 AND nomforpa like '%trans%' AND 1"
    Cad = DevuelveDesdeBD("codforpa", "formapago", Cad, "1 ORDER BY codforpa")
    
    'FORMA PAGO transferencia
    If Cad = "" Then
        Cad = "tipforpa=1 AND 1"
        Cad = DevuelveDesdeBD("codforpa", "formapago", Cad, "1 ORDER BY codforpa")
    End If
    If Cad = "" Then
        'Cualqueira
        Cad = "tipforpa=3 AND nomforpa like '%trans%' AND 1"
        Cad = DevuelveDesdeBD("codforpa", "formapago", Cad, "1 ORDER BY codforpa")
    End If
    
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad & "|"
    fRecep = RecuperaValor(CadenaDesdeOtroForm, 2)
     
    If Me.optVarios(1).Value Then
        tCobro = "select ctabanc1,count(*) from pagos where fecfactu>=" & DBSet(DateAdd("m", -2, Now), "F") & " group by 1 order by 2 desc"
    Else
        tCobro = "select ctabanc1,count(*) from cobros where fecfactu>=" & DBSet(DateAdd("m", -2, Now), "F") & " group by 1 order by 2 desc"
    End If
    miRsAux.Open tCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    tCobro = ""
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then tCobro = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
    If tCobro = "" Then
        tCobro = DevuelveDesdeBD("codmacta", "bancos", "1", "1 ORDER BY 1 desc")
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & tCobro & "|"
    tCobro = ""
                        
    Actual = fRecep <= vParam.fechafin
    
    
    Cad = RecuperaValor(CadenaDesdeOtroForm, 3)
    Tipointegracion = 1
    If Cad = "1" Then Tipointegracion = 0 'todo
    
    
    
    'codusu,Codigo,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,tipo_ret,ctaventas,Ccoste,impventa,
    'IVA , ImpIva, recargo, imprecargo, TotalFactura, integracion
    
    Cad = "select tmpintefrafracli.*,codpais,nifdatos,desprovi,despobla,codposta,dirdatos,nommacta,cuentas.iban ibancta from tmpintefrafracli left join cuentas ON cta_cli=codmacta"
    Cad = Cad & " WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    FACTURA = ""
    Set Mc = New Contadores
    
    
    
    
    
    While Not miRsAux.EOF
            Cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
            
            If Mc.ConseguirContador(1, Actual, False) = 1 Then Exit Sub
                  
            
            'Haremos los inserts
            'INSERT INTO factpro(codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta
            'codconce340,codopera,codmacta,anofactu,
            'codforpa,observa,totbases,totbasesret,totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten
            ',observa NUmSerie , Numregis, fecharec, NumFactu, FecFactu,fecliqpr
            
            
            NumRegElim = Mc.Contador
            BaseImponible = miRsAux!calculoimponible
            TotalFac = miRsAux!TotalFactura
            IVA = DBLet(miRsAux!ImpIva, "N")
            Msg$ = "(" & DBSet(miRsAux!codpais, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T", "S") & "," & DBSet(miRsAux!desProvi, "T", "S") & ","
            Msg$ = Msg$ & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S") & "," & DBSet(miRsAux!dirdatos, "T", "S") & ","
            
            Msg$ = Msg$ & DBSet(miRsAux!Nommacta, "T") & ",'X',5," & DBSet(miRsAux!cta_cli, "T") & "," & Year(fRecep)
            If DBLet(miRsAux!fpago, "T") = "" Then
                'No viene forma de pago
                 K = RecuperaValor(CadenaDesdeOtroForm, 5)
            Else
                 K = miRsAux!fpago
            End If
            Msg$ = Msg$ & "," & K & ","
            'totbases,totbasesret,totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten,
            Msg$ = Msg$ & DBSet(miRsAux!impventa, "N") & "," & DBSet(miRsAux!impventa + IVA, "N") & "," & DBSet(IVA, "N")
            Msg$ = Msg$ & ",NULL," & DBSet(TotalFac, "N") & ","
            If DBLet(DBLet(miRsAux!tipo_ret, "N"), "N") = 0 Then
                'No lleva retencion
                Msg$ = Msg$ & "NULL,NULL,NULL,0"
            Else
                Msg$ = Msg$ & DBSet(Round2((miRsAux!impret / (miRsAux!impventa + IVA)), 2) * 100, "N", "S") & "," & DBSet(miRsAux!impret, "N", "N")
                Msg$ = Msg$ & ",'" & RecuperaValor(CadenaDesdeOtroForm, 1) & "'," & DBLet(miRsAux!tipo_ret, "N")    'tiporeten
            End If
            ',observa NUmSerie , Numregis, fecharec, NumFactu, FecFactu,fecliqpr
            
            Msg$ = Msg$ & ",'" & DevNombreSQL(miRsAux!txtcsb) & "           "
            Msg$ = Msg$ & "Importacion datos externos. " & " " & "Usuario: " & DevNombreSQL(vUsu.Nombre) & Chr(13) & "Fecha: " & Now
            Msg$ = Msg$ & "'," & DBSet(miRsAux!Serie, "T") & "," & NumRegElim & "," & DBSet(fRecep, "F") & ","
            Msg$ = Msg$ & DBSet(Format(miRsAux!FACTURA, "000000"), "T") & "," & DBSet(miRsAux!Fecha, "F") & "," & DBSet(fRecep, "F") & ")"
            
                    
                    
                

            
            'Para los cobros     (cojo esta variable)
            '       insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,
            'numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,text41csb,numorden,fecvenci,impvenci)
            'Pagos
            'nomprove,domprove,pobprove,cpprove,proprove,codpais,nifprove,numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text1csb,text2csb,numorden,fecefect,impefect
            If Tipointegracion <> 1 Then
                    'De momento NO contemplada
                    If Me.optVarios(1).Value Then
                        tCobro = "("
                    Else
                        tCobro = " (" & RecuperaValor(CadenaDesdeOtroForm, 4) & ","  'agente
                    End If
                    tCobro = tCobro & DBSet(miRsAux!Nommacta, "T") & "," & DBSet(miRsAux!dirdatos, "T", "S")
                    tCobro = tCobro & "," & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S")
                    tCobro = tCobro & "," & DBSet(miRsAux!desProvi, "T", "S") & "," & DBSet(miRsAux!codpais, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T")
                    tCobro = tCobro & "," & DBSet(miRsAux!Serie, "T") & "," & DBSet(Format(miRsAux!FACTURA, "000000"), "T") & "," & DBSet(miRsAux!Fecha, "F") & "," & DBSet(miRsAux!cta_cli, "T")
                    tCobro = tCobro & "," & K & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 6), "T") & ","

                    'Si me ha indicado IBAN, o no, en la importacion
                    If Me.optVarios(1).Value Then
                        Aux = DBSet(miRsAux!ibancta, "T", "S")
                    Else
                    If DBLet(miRsAux!IBAN, "T") = "" Then
                        Aux = DBSet(miRsAux!ibancta, "T", "S")
                    Else
                        Aux = DBSet(miRsAux!IBAN, "T", "S")
                    End If
                    
                    End If
                    tCobro = tCobro & Aux & ","

                    'Si me ha indicado(o no) txtobserva text33csb text41csb
                    If DBLet(miRsAux!txtcsb, "T") = "" Then
                        Aux = DBSet("Factura : " & Format(miRsAux!FACTURA, "000000") & " de " & miRsAux!Fecha, "T") & ",null,"
                    Else
                        Aux = miRsAux!txtcsb
                        If Len(Aux) > 80 Then
                            Aux = DBSet(Mid(Aux, 1, 80), "T") & "," & DBSet(Mid(Aux, 81), "T") & ","
                        Else
                            Aux = DBSet(Aux, "T") & ",NULL,"
                        End If
                    End If
                    tCobro = tCobro & Aux

                    CargarCobrosSobreCollectionConSQLInsert ColCob, CStr(K), miRsAux!Fecha, TotalFac, tCobro
            End If
            

        i = 0
        FACTURA = Cad
        Set ColBases = Nothing
        Set ColBases = New Collection
                        
        'Lineas
        'INSERT INTO factpro_lineas (aplicret,imporec,anofactu,codccost,
        'impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea,fecfactu,numserie,numfactu
        i = i + 1
        Cad = "(1," & DBSet(miRsAux!recargo, "T") & "," & Year(fRecep) & "," & DBSet(miRsAux!CCoste, "T", "S") & ","
        Cad = Cad & DBSet(miRsAux!ImpIva, "N") & "," & DBSet(miRsAux!recargo, "N", "S") & "," & DBSet(miRsAux!IVA, "N")
        Cad = Cad & "," & DBSet(miRsAux!impventa, "N", "N") & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux!ctaventas, "T") & "," & i
        Cad = Cad & "," & DBSet(fRecep, "F", "N") & "," & DBSet(miRsAux!Serie, "T") & "," & NumRegElim & ")"
        
        ColBases.Add Cad
            
        If Not InsertarFACTURA(ColBases, FACTURA, fRecep, Tipointegracion, ColCob) Then
            Mc.DevolverContador Mc.TipoContador, Actual, Mc.Contador, False
        End If
        miRsAux.MoveNext
     Wend
     Msg$ = ""
     Ampliacion = ""
    
    
    
    
    
End Sub




























'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'
'       ASIENTOS
'
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
'*********************************************************************************
Private Sub ImportarAsientos()
Dim Rc As Byte
    
    'Primer paso. Lectura fichero. Comprobacion basica datos
    PonerLabel "Leyendo fichero"
    Rc = ImportarFichAsientos
    If Rc = 2 Then Exit Sub
        
    If Rc = 1 Then
        'Errores en fichero
        'Ha habido errores
        CargaEncabezado 0
    
        'Cargamos datos
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto1
            ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
            ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Exit Sub
    End If
    
    'Segundo paso. Comprobacion datos BD (existe serie, ctas, tiposiva....
    'Fichero importado.
    'Comporbaciones BD
    PonerLabel "Comprobando en BD"
    If Not ComprobacionDatosBDAsientos Then
        'Cargaremos errores
        CargaEncabezado 1
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto2
            ListView1.ListItems(i).SubItems(1) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    
    
        Exit Sub
    End If
  
    
    'Si llega aqui, es que ha ido bien. Estamos pendientes de aceptar
    CargaEncabezado 4
    
    Set miRsAux = New ADODB.Recordset
   
    Cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on cta_cli=cuentas.codmacta where codusu= " & vUsu.Codigo
    Cad = Cad & " ORDER BY codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    K = -1
    While Not miRsAux.EOF
        i = i + 1
        
       
            ListView1.ListItems.Add
            If K = miRsAux!FACTURA Then
                Cad = " "
            Else
                Cad = Format(miRsAux!FACTURA, "000")
            End If
            K = miRsAux!FACTURA
            ListView1.ListItems(i).Text = Cad
            ListView1.ListItems(i).SubItems(1) = miRsAux!Fecha
            ListView1.ListItems(i).SubItems(2) = miRsAux!cta_cli
            ListView1.ListItems(i).SubItems(3) = miRsAux!Nommacta
            ListView1.ListItems(i).SubItems(4) = DBLet(miRsAux!IBAN, "T")
            ListView1.ListItems(i).SubItems(5) = DBLet(miRsAux!txtcsb, "T")
            
            If CStr(miRsAux!Serie) = "H" Then
                ListView1.ListItems(i).SubItems(6) = " "
                ListView1.ListItems(i).SubItems(7) = Format(miRsAux!impventa, FormatoImporte)
            Else
                ListView1.ListItems(i).SubItems(6) = Format(miRsAux!impventa, FormatoImporte)
                ListView1.ListItems(i).SubItems(7) = " "
            End If
            
      
        
            miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cmdAceptar.visible = True
    
    
End Sub


Private Function ImportarFichAsientos() As Byte
Dim Aux As String



    On Error GoTo eImportarFichAsientos
    ImportarFichAsientos = 2  'Vacio
    
    NF = -1
    'tmpintefrafracli(codusu,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,
    'tipo_ret,ctaventas,Ccoste,impventa,iva,impiva,recargo,imprecargo,totalfactura,integracion)
    Msg$ = "FECHA;ASIENTO;CUENTA;cc;DOCUMENTO;AMP.CONCEPTO,;IMPORTE;D/H;CONTRAPARTIDA;CONTRAPARTIDA.AUT. (0=NO -1=SI);"
    
    
    'Preparamos tabla de insercion
    Conn.Execute "DELETE FROM tmpintefrafracli WHERE codusu = " & vUsu.Codigo
    'errores
    Conn.Execute "DELETE FROM tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
    
    NF = FreeFile
    Open Text1.Text For Input As #NF
    
    
    ''tmpintefrafracli(codusu,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,
    'tipo_ret,ctaventas,Ccoste,impventa,iva,impiva,recargo,imprecargo,totalfactura,integracion)
    NumRegElim = 0
    J = 0 'Numero de error en el total de fichero. Secuencial

    FacturaAnterior = ""
    While Not EOF(NF)
        NumRegElim = NumRegElim + 1
        Line Input #NF, Cad
        
        If NumRegElim = 1 Then
            'Primera linea encabezado?
            If Me.Check1.Value = 1 Then Cad = ""
        Else
            '10 campos a tratar
            If InStr(1, String(10, ";"), Cad) > 0 Then Cad = "" 'todo puntos y comas
        End If
        
        
        If Cad <> "" Then
            'Procesamos linea
            
            strArray = Split(Cad, ";")
            
            If UBound(strArray) = 10 - 1 Then
                'Falta el ultimo punto y coma
                Cad = Cad & ";"
                strArray = Split(Cad, ";")
            End If
            
            
            If UBound(strArray) <> 10 Then
                J = J + 1
                Aux = vUsu.Codigo & "," & J & "," & NumRegElim & ",'Nº campos incorrecto'," & DBSet(Cad, "T")
                Conn.Execute "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,observa1) VALUES (" & Aux & ")"
                
            Else
                'En la sub o insertara en la tabla de correctos o en la de errores ,
                'en funcion de los tipos de datos y que sean requeridos o no
                SeparaLineaAsiento
            
            
            End If
            
            
            
            
        End If
        
    Wend
    Close #NF
    If NumRegElim = 0 Then
        MsgBox "Fichero vacio", vbExclamation
    Else
        Cad = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
        If Val(Cad) > 0 Then
            ImportarFichAsientos = 1 'Con errores
        Else
            Cad = DevuelveDesdeBD("count(*)", "tmpintefrafracli", "codusu", vUsu.Codigo)
            If Val(Cad) > 0 Then ImportarFichAsientos = 0
        End If
    End If
    
eImportarFichAsientos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        CerrarFichero
    End If
End Function

Private Sub SeparaLineaAsiento()
Dim NuevaLinea As Boolean
Dim Linea As Integer
Dim Fin As Boolean

    
    Linea = 1
    Fin = False
    Do
    
        'Vemos los campos.
                                                            '  0       1         2       3     4
                                                        '               FECHA;ASIENTO;CUENTA; cc
        CadenaDesdeOtroForm = "INSERT INTO tmpintefrafracli(codusu , Codigo,fecha,factura,cta_cli,Ccoste,"
                                                    '      5         6      7       8            9
                                                    'AMP.CONCEPTO,;IMPORTE;D/H;CONTRAPARTIDA;CONTRAPARTIDA.AUT. (0=NO -1=SI);
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "iban,txtcsb,impventa,serie,ctaventas,ctaret)"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & " VALUES (" & vUsu.Codigo & "," & NumRegElim
    
        LineaConErrores = False
        

        
        'Validamos campos
        For K = 0 To 9
            'Campos
            ValidarLineaAsiento CByte(K)
        Next K
        
        If Not LineaConErrores And Linea = 1 Then
            If strArray(9) = "1" Then
                'QUiere generar contrapunte de la copntrapartida
                'Por lo tanto  ESTA ULTIMA DEBE EXISTIR
                If strArray(8) = "" Then AnyadeEnErrores "Contrapartida NO esta indicada"
            End If
        End If
        If Not LineaConErrores Then
            'INSERTAMOS EN tmpintefrafracli
            Conn.Execute CadenaDesdeOtroForm & ")"
            
            'Si quiere generar el contrapunte
            If Linea = 1 Then
                If strArray(9) = "1" Then
                    NumRegElim = NumRegElim + 1
                    CadenaDesdeOtroForm = strArray(2)
                    strArray(2) = strArray(8)
                    strArray(8) = CadenaDesdeOtroForm
                    'Debe-haber
                    If strArray(7) = "H" Then
                        strArray(7) = "D"
                    Else
                        strArray(7) = "H"
                    End If
                    
                    Linea = 2
                Else
                    Fin = True
                End If
            Else
                Fin = True
            End If
        Else
            Fin = True
            
        End If
    
    
    
    
    Loop Until Fin
    
    
    
End Sub



Private Sub ValidarLineaAsiento(QueCampo As Byte)
Dim ValorSQL As String

    'Vemos los campos.
    ' 0       1         2   3     4           5         6      7       8            9
    'FECHA;ASIENTO;CUENTA; cc;DOCUMENTO;AMP.CONCEPTO,;IMPORTE;D/H;CONTRAPARTIDA;CONTRAPARTIDA.AUT. (0=NO -1=SI);
    
    ValorSQL = "NULL"
    Select Case QueCampo
    'Numerico REQUERIDO
    Case 1, 6, 9
        If QueCampo = 9 And strArray(QueCampo) = "" Then strArray(QueCampo) = "0"
        strArray(QueCampo) = Replace(strArray(QueCampo), ".", ",")
        
      
        If strArray(QueCampo) = "" Then
            AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " no puede estar vacio"
        Else
            If Not IsNumeric(strArray(QueCampo)) Then
                AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
            Else
                ValorSQL = DBSet(strArray(QueCampo), "N")
            End If
                
        End If
    
    Case 7
        If UCase(strArray(QueCampo)) <> "H" And UCase(strArray(QueCampo)) <> "D" Then
            AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
        Else
            ValorSQL = DBSet(strArray(QueCampo), "T")
        End If
    'Fecha
    Case 0
        
        
        
        If Not IsDate(strArray(QueCampo)) Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo fecha"
        ValorSQL = DBSet(strArray(QueCampo), "F", "S")
        
    
    Case 3
        'Centro de coste
        'Si lleva analitica es obligatorio
        If strArray(QueCampo) = "" Then
            If vParam.autocoste Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " no puede estar vacio"
        Else
            ValorSQL = DBSet(strArray(QueCampo), "T")
        End If
 
  
    Case Else
        ValorSQL = DBSet(strArray(QueCampo), "T", "S")
    End Select
    
    
    
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "," & ValorSQL
    
End Sub





Private Function ComprobacionDatosBDAsientos() As Boolean
Dim FechaMinima As Date

On Error GoTo eComprobacionDatosBDAsientos
    ComprobacionDatosBDAsientos = False
    Set miRsAux = New ADODB.Recordset
    
    
    Cad = "select min(fecha) minima,max(fecha) maxima from tmpintefrafracli where factura >0 and codusu=" & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'no puede ser eof
    Cad = miRsAux!minima
    FechaMinima = miRsAux!minima
    If miRsAux!minima < vParam.fechaini Then
        AnyadeEnErrores "Menor que inicio ejercicio"
    Else
        If miRsAux!minima < vParam.FechaActiva Then
            AnyadeEnErrores "Menor que fecha activa"
        Else
            If miRsAux!maxima > DateAdd("yyyy", 1, vParam.fechafin) Then AnyadeEnErrores "Mayor ejercicios abiertos"

        End If
    End If
    miRsAux.Close
    
    

    'Cuentas
    Cad = "select distinct cta_cli from tmpintefrafracli where cta_cli<>''  and codusu=" & vUsu.Codigo & " and not cta_cli in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cuenta"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
        
    Cad = "select distinct ctaventas from tmpintefrafracli where ctaventas<>''  and codusu=" & vUsu.Codigo & " and not ctaventas in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cuenta"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    
         
         
    If vParam.autocoste Then
        Cad = "select ccoste from tmpintefrafracli where ccoste<>''  and codusu=" & vUsu.Codigo
        Cad = Cad & " AND  not ccoste iN (select codccost from ccoste )"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = miRsAux!FACTURA
            AnyadeEnErrores "No existe centro de coste asiento: " & miRsAux!CCoste
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    
    
    'Misma fecha para cada asiento
    'Por lo menos debe tener DOS lineas
    Cad = "select factura,max(fecha) maximo,min(fecha) Minimo,count(*) cuantos from tmpintefrafracli WHERE  codusu =" & vUsu.Codigo & " group by 1 "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = ""
        If DBLet(miRsAux!maximo, "F") <> DBLet(miRsAux!Minimo, "F") Then
            Cad = "Asiento: " & miRsAux!FACTURA & "   distintas fechas: " & miRsAux.Fields(1) & " " & miRsAux.Fields(1)
        Else
            If DBLet(miRsAux!Cuantos, "N") < 2 Then Cad = "Asiento: " & miRsAux!FACTURA & "   una linea solo"
        End If
        If Cad <> "" Then AnyadeEnErrores "Fechas/Lineas"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
                
                
    'No dejamos , de momento, que tengan fechas de un ejercicio y de otro. Ya que
    'como vamos a pasar TODO o NADA, por tema de contadores
    Cad = "select max(fecha) maximo,min(fecha) Minimo from tmpintefrafracli WHERE  codusu =" & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If miRsAux!Minimo <= vParam.fechafin Then
            If miRsAux!maximo > vParam.fechafin Then
                Cad = "Asientos en distintos ejercicios. Opcion no disponible"
                AnyadeEnErrores "Fechas ejercicios"
            End If
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
                
    'Totales
    Cad = "select factura,sum(if(serie='H',impventa,-impventa)) diferen from tmpintefrafracli WHERE  codusu =" & vUsu.Codigo & " group by 1 "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = ""
        If DBLet(miRsAux!diferen, "N") <> 0 Then
            Cad = "Asiento: " & miRsAux!FACTURA & "   descuadre: " & miRsAux.Fields(1)
            AnyadeEnErrores "Fechas/Lineas"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
                
                
                
                

    'Si no han habiado errores..  j=0
    ComprobacionDatosBDAsientos = J = 0


    

eComprobacionDatosBDAsientos:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function



Private Sub InsertarAsientos()

    'Cargaremos los nomcponce
    strArray = Split(";;", ";")
    Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", RecuperaValor(CadenaDesdeOtroForm, 1))
    strArray(0) = Cad
    Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", RecuperaValor(CadenaDesdeOtroForm, 2))
    strArray(1) = Cad
    strArray(2) = RecuperaValor(CadenaDesdeOtroForm, 3)


    Cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on cta_cli=cuentas.codmacta where codusu= " & vUsu.Codigo
    Cad = Cad & " ORDER BY codigo"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    Conn.BeginTrans
    If InsertarAsientosBD Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
End Sub


Private Function InsertarAsientosBD() As Boolean
Dim rsContador As ADODB.Recordset
Dim Contador As Long
Dim ValorEnBD As Long
Dim ACtuala As Boolean
        On Error GoTo eInsertarAsientosBD
        
        
        
        
        Set rsContador = New ADODB.Recordset
        Cad = "Select * from contadores where tiporegi='0' FOR UPDATE"
        rsContador.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        ACtuala = True
        If miRsAux!Fecha <= vParam.fechafin Then
            Contador = rsContador!Contado1
        Else
            ACtuala = False
            Contador = rsContador!Contado2
        End If
        'hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,codccost,timporteH,ctacontr,idcontab)
        i = 0
        ValorEnBD = -1
        
        While Not miRsAux.EOF
            If miRsAux!FACTURA <> ValorEnBD Then
                i = i + 1
                
                'Nuevo asiento
                If ValorEnBD > 0 Then
                    Cad = Mid(Cad, 2) 'primera com
                    'Insertamos
                    Cad = "INSERT INTO hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,codccost,ctacontr,idcontab) VALUES " & Cad
                    Conn.Execute Cad
                End If
                'hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion)
                
                Cad = "INSERT INTO hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES "
                Cad = Cad & "(" & strArray(2) & "," & DBSet(miRsAux!Fecha, "F") & "," & Contador + i & ",'Desde importacion apuntes. Fecha " & Now() & "',now()," & DBSet(vUsu.Login, "T")
                Cad = Cad & ",'Ariconta6');"
                'Inserto YA la cabecera
                Conn.Execute Cad
                K = 0
                Cad = ""
                ValorEnBD = miRsAux!FACTURA
            End If
            K = K + 1
            'cad = "hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,
            FacturaAnterior = "(" & strArray(2) & "," & DBSet(miRsAux!Fecha, "F") & "," & Contador + i & "," & K & "," & DBSet(miRsAux!cta_cli, "T") & "," & DBSet(miRsAux!IBAN, "T", "N") & ","
            'codconce,ampconce,timporteD,timporteH,,codccost,ctacontr,idcontab
            If miRsAux!Serie = "D" Then
                'Al debe
                FacturaAnterior = FacturaAnterior & RecuperaValor(CadenaDesdeOtroForm, 1) & "," & DBSet(Mid(Trim(strArray(0) & " " & miRsAux!txtcsb), 1, 50), "T", "N") & ","
                FacturaAnterior = FacturaAnterior & DBSet(miRsAux!impventa, "N") & ",NULL"
            Else
                FacturaAnterior = FacturaAnterior & RecuperaValor(CadenaDesdeOtroForm, 2) & "," & DBSet(Mid(Trim(strArray(1) & " " & miRsAux!txtcsb), 1, 50), "T", "N")
                FacturaAnterior = FacturaAnterior & ",NULL," & DBSet(miRsAux!impventa, "N")
            End If
            FacturaAnterior = FacturaAnterior & "," & DBSet(miRsAux!CCoste, "T", "S") & "," & DBSet(miRsAux!ctaventas, "T", "S") & ",'contab')"
            Cad = Cad & ", " & FacturaAnterior
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        'La ultima
        Cad = Mid(Cad, 2) 'primera com
        'Insertamos
        Cad = "INSERT INTO hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,timporteH,codccost,ctacontr,idcontab) VALUES " & Cad
        Conn.Execute Cad
        InsertarAsientosBD = True
       
        rsContador.Close
        
        
        Cad = "contado2"
        If ACtuala Then Cad = "contado1"
        
        Cad = "UPDATE contadores set " & Cad & " = " & Contador + i & " WHERE tiporegi='0'"
        Conn.Execute Cad
        
eInsertarAsientosBD:
    If Err.Number <> 0 Then
        Cad = Err.Description
        AnyadeEnErrores "ID asiento fich: " & ValorEnBD
    End If
    Set rsContador = Nothing
    Set miRsAux = Nothing
End Function

'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'  Facturas proveedore STANDARD
'--------------------------------------------------------------------
Private Sub ImportacionFacturasProveedor()
Dim Rc As Byte
Dim B As Boolean
Dim FraProStd As Boolean
Dim MismaFactura As Boolean
 'Primer paso. Lectura fichero. Comprobacion basica datos
    PonerLabel "Leyendo fichero"
    NumeroCamposTratarFraPro = 19
    Rc = ImportarFicherosFraStandard
    If Rc = 2 Then Exit Sub
        
    If Rc = 1 Then
        'Errores en fichero
        'Ha habido errores
        CargaEncabezado 0
    
        'Cargamos datos
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto1
            ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
            ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Exit Sub
    End If
    
    
    
    
    'Segundo paso. Comprobacion datos BD (existe serie, ctas, tiposiva....
    'Fichero importado.
    'Comporbaciones BD
    PonerLabel "Comprobando en BD"
    If Not ComprobacionDatosBDFacturasProveedoresSTD Then
        'Cargaremos errores
        CargaEncabezado 1
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto2
            ListView1.ListItems(i).SubItems(1) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    
    
    
        Exit Sub
    End If
    
    
    'Ultima comprobacion. Que las facturas suman lo que tienen que sumar
    PonerLabel "Totales"
    If Not ComprobarTotales Then
        CargaEncabezado 2
        Set miRsAux = New ADODB.Recordset
        Cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        i = 0
        While Not miRsAux.EOF
            i = i + 1
            ListView1.ListItems.Add , "K" & miRsAux!Codigo
            ListView1.ListItems(i).Text = miRsAux!texto1
            ListView1.ListItems(i).SubItems(1) = miRsAux!texto2
            ListView1.ListItems(i).SubItems(2) = miRsAux!observa1
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        Exit Sub
    End If
    
    
    
    'Si llega aqui, es que ha ido bien. Estamos pendientes de aceptar
    CargaEncabezado 3
    
    
    'Que culumna pintara en Base/retencion dependiendo frapro
    B = True
    FraProStd = False
    If Me.optVarios(1).Value Then
        B = False
        If Me.cboTipo.ListIndex = 1 Then FraProStd = True
    End If
    
    
    
    
    Set miRsAux = New ADODB.Recordset
   
    Cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on cta_cli=cuentas.codmacta where codusu= " & vUsu.Codigo
    Cad = Cad & " ORDER BY codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        ListView1.ListItems.Add , "K" & i
        MismaFactura = False
        If FraProStd Then
            If DBLet(miRsAux!FACTURA, "T") = "" Then MismaFactura = True
        Else
            If DBLet(miRsAux!Serie, "T") = "" Then MismaFactura = True
        End If
        If Not MismaFactura Then
            
            ListView1.ListItems(i).Text = miRsAux!Serie
            If FraProStd Then
                ListView1.ListItems(i).SubItems(1) = miRsAux!FACTURA
            Else
                ListView1.ListItems(i).SubItems(1) = Format(miRsAux!FACTURA, "000000")
            End If
            ListView1.ListItems(i).SubItems(2) = miRsAux!Fecha
            
            ListView1.ListItems(i).SubItems(3) = DBLet(miRsAux!Nommacta, "T")
            
            
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!impventa, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            ListView1.ListItems(i).SubItems(7) = " " 'separador
            If B Then
                ListView1.ListItems(i).SubItems(8) = Format(miRsAux!calculoimponible, FormatoImporte)
            Else
                ListView1.ListItems(i).SubItems(8) = Format(DBLet(miRsAux!impret, "N"), FormatoImporte)
            End If
            ListView1.ListItems(i).SubItems(9) = Format(miRsAux!TotalFactura - miRsAux!calculoimponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(10) = Format(miRsAux!TotalFactura, FormatoImporte)
            
            If vEmpresa.TieneTesoreria Then
                Cad = " "
                If DBLet(miRsAux!IBAN, "T") <> "" Or DBLet(miRsAux!txtcsb, "T") <> "" Then Cad = "*"
                ListView1.ListItems(i).SubItems(11) = Cad
            End If
            
        Else
            
            ListView1.ListItems(i).Text = " "
            ListView1.ListItems(i).SubItems(3) = "              IVA " & miRsAux!IVA & "%"
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!impventa, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            For NF = 1 To 10
                If NF < 3 Or NF > 6 Then ListView1.ListItems(i).SubItems(NF) = " "
            Next
            
            If vEmpresa.TieneTesoreria Then ListView1.ListItems(i).SubItems(11) = " "
            
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cmdAceptar.visible = True
    
    
End Sub


Private Function ImportarFicherosFraStandard() As Byte
Dim Aux As String



    On Error GoTo eImportarFicheroFracli
    ImportarFicherosFraStandard = 2 'Vacio
    
    NF = -1
    '1           2              3                   4       5           6
    'FACTURA    FECH FACT   FECHA CONTA       CTA. PRO.   F.PAGO      TIPO OPERACION
    '7             8            9
    'CTA.RET.    IMP.RET     TIPO RET.
    ' 10                11          12          13        14        15
    'CTA.COMPRAS   IMP.COMPRA     % I.V.A.    IMP.IVA   % R.E.      IMP. REC.
    '   16             17                   18      19
    'TOTAL FACTURA   INTEGRACION iban    txtcsb    ccost
    
    Msg$ = "factura|fech fact|fecha conta|cta. pro.|f.pago|tipo operacion|cta.ret|.imp.ret|tipo ret.|cta.compras|imp.compra|% i.v.a.|imp.iva|% r.e.|imp. rec.|total factura|integracion|iban|txtcsb|cc|"
    
    
    'Preparamos tabla de insercion
    Conn.Execute "DELETE FROM tmpintefrafracli WHERE codusu = " & vUsu.Codigo
    'errores
    Conn.Execute "DELETE FROM tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
    
    NF = FreeFile
    Open Text1.Text For Input As #NF
    
    
    ''tmpintefrafracli(codusu,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,
    'tipo_ret,ctaventas,Ccoste,impventa,iva,impiva,recargo,imprecargo,totalfactura,integracion)
    NumRegElim = 0
    J = 0 'Numero de error en el total de fichero. Secuencial

    FacturaAnterior = ""
    While Not EOF(NF)
        NumRegElim = NumRegElim + 1
        Line Input #NF, Cad
        
        If NumRegElim = 1 Then
            'Primera linea encabezado?
            If Me.Check1.Value = 1 Then Cad = ""
        Else
            If InStr(1, String(NumeroCamposTratarFraPro, ";"), Cad) > 0 Then Cad = "" 'todo puntos y comas
        End If
        
        
        If Cad <> "" Then
            'Procesamos linea
            
            strArray = Split(Cad, ";")
            
            If UBound(strArray) = NumeroCamposTratarFraPro - 1 Then
                'Falta el ultimo punto y coma
                Cad = Cad & ";"
                strArray = Split(Cad, ";")
            End If
            
            
            If UBound(strArray) < NumeroCamposTratarFraPro Then
                J = J + 1
                Aux = vUsu.Codigo & "," & J & "," & NumRegElim & ",'Nº campos incorrecto'," & DBSet(Cad, "T")
                Conn.Execute "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,observa1) VALUES (" & Aux & ")"
                
            Else
                If UBound(strArray) > NumeroCamposTratarFraPro + 1 Then
                    J = J + 1
                    Aux = vUsu.Codigo & "," & J & "," & NumRegElim & ",'Nº campos incorrecto +'," & DBSet(Cad, "T")
                    Conn.Execute "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,observa1) VALUES (" & Aux & ")"
                    
                Else
                    'En la sub o insertara en la tabla de correctos o en la de errores ,
                    'en funcion de los tipos de datos y que sean requeridos o no
                    SeparaLineaProveedorStandard
                
                End If
            End If
            
            
            
            
        End If
        
    Wend
    Close #NF
    If NumRegElim = 0 Then
        MsgBox "Fichero vacio", vbExclamation
    Else
        Cad = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
        If Val(Cad) > 0 Then
            ImportarFicherosFraStandard = 1 'Con errores
        Else
            Cad = DevuelveDesdeBD("count(*)", "tmpintefrafracli", "codusu", vUsu.Codigo)
            If Val(Cad) > 0 Then ImportarFicherosFraStandard = 0
        End If
    End If
    
eImportarFicheroFracli:
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        CerrarFichero
    End If
End Function



Private Sub SeparaLineaProveedorStandard()
Dim NuevaLinea As Boolean

    CadenaDesdeOtroForm = "INSERT INTO tmpintefrafracli(codusu,Codigo,serie,factura,fecha2,fecha,cta_cli,fpago,tipo_operacion,"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "ctaret,impret,tipo_ret,"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "ctaventas,impventa,iva,impiva,recargo,imprecargo,"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "totalfactura,integracion,iban,txtcsb,Ccoste"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ") VALUES (" & vUsu.Codigo & "," & NumRegElim & ",1"  'la serie tambien es fija.... de momento
    

    '1           2              3                   4       5           6
    'FACTURA    FECH FACT   FECHA CONTA       CTA. PRO.   F.PAGO      TIPO OPERACION
    '7             8            9
    'CTA.RET.    IMP.RET     TIPO RET.
    ' 10                11          12          13        14        15
    'CTA.COMPRAS   IMP.COMPRA     % I.V.A.    IMP.IVA   % R.E.      IMP. REC.
    '   16             17                   18
    'TOTAL FACTURA   INTEGRACION iban    txtcsb

    LineaConErrores = False
    
   
    
    
    
    'Precomprobacion (RAFA)
    'Puede ser que repita el numero,serie,fecha para la misma factura(con mas de una linea), ya que si se lanza el proceso
    'que genera el fichero csv es mas comodo poner todos los datos aunque se repitan
    'Por lo tanto , aqui, si los datos de factuas son los mismo, los reseteo
    NuevaLinea = True
    If FacturaAnterior <> "" Then
        If strArray(0) <> "" Or strArray(1) <> "" Or strArray(2) <> "" Then
            'Ha puesto datos de facturas
            If strArray(0) = RecuperaValor(FacturaAnterior, 1) Then
                If strArray(1) = RecuperaValor(FacturaAnterior, 2) Then
                    If strArray(3) = RecuperaValor(FacturaAnterior, 3) Then
                        'Es la misma que la de arriba. Pongo los campos a vacio
                        strArray(0) = ""
                        strArray(1) = "0"
                        strArray(2) = "1900-01-01"
                        strArray(3) = ""
                        
                        
                        
                        NuevaLinea = False
                    End If
                End If
            End If
        End If
    End If
    If NuevaLinea Then
         If strArray(0) <> "" And strArray(1) <> "" And strArray(3) <> "" Then FacturaAnterior = strArray(0) & "|" & strArray(1) & "|" & strArray(3) & "|"
    End If
    
    'Validamos campos
    For K = 0 To NumeroCamposTratarFraPro
        ValidarLineaProveStd CByte(K)
    Next K
    
    
    If Not LineaConErrores Then
        
        'Si pone algun dato de la factura, debe ponerlos todos
        If strArray(0) <> "" Or strArray(1) <> "1900-01-01" Or strArray(2) <> "1900-01-01" Then
            If strArray(0) = "" Or strArray(1) = "1900-01-01" Or strArray(2) = "1900-01-01" Then
                AnyadeEnErrores "Campos facturas. Todos o ninguno"
            End If
        End If
        
       ' 0 = TODO    1 = CONTA  2= TESORERIA
        If Val(strArray(16)) <> 1 Then
            'Queremos que meta el cobro
            If strArray(5) = "" Then
                AnyadeEnErrores "Falta forma de pago"
            End If
            
            
            
            
        End If
        
        
        'Si indica retencion tiene que indicar el tipo y bicepsversa
        
        If Val(strArray(8)) > 0 Then
            If strArray(7) = "" Then strArray(7) = "0"
            If CCur(strArray(7)) = 0 Then AnyadeEnErrores "Error importe retencion FALTA"
            If strArray(6) = "" Then AnyadeEnErrores "Error cuenta retencion FALTA"
        Else
            
                If Val(strArray(7)) > 0 Then AnyadeEnErrores "Error tipo retencion indicado "
                If strArray(6) <> "" Then AnyadeEnErrores "Error cuenta retencion indicada"

        End If
        
        
        
        
        
        
        
    End If
    
    
    If Not LineaConErrores Then
    
        
        
        
        'INSERTAMOS EN tmpintefrafracli
        Conn.Execute CadenaDesdeOtroForm & ")"
        
    
    End If
    
End Sub


Private Sub ValidarLineaProveStd(QueCampo As Byte)
Dim ValorSQL As String

    'Vemos los campos.
    '1           2              3                   4       5           6
    'FACTURA    FECH FACT   FECHA CONTA       CTA. PRO.   F.PAGO      TIPO OPERACION
    '7             8            9
    'CTA.RET.    IMP.RET     TIPO RET.
    ' 10                11          12          13        14        15
    'CTA.COMPRAS   IMP.COMPRA     % I.V.A.    IMP.IVA   % R.E.      IMP. REC.
    '   16             17                   18   19
    'TOTAL FACTURA   INTEGRACION iban    txtcsb  CC

    
    
    ValorSQL = "NULL"
    Select Case QueCampo
    'Numerico REQUERIDO
    Case 9, 10, 11, 12
      
        'RAFA. Pone un . decimal
        strArray(QueCampo) = Replace(strArray(QueCampo), ".", ",")
        
      
        If strArray(QueCampo) = "" Then
            AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " no puede estar vacio"
        Else
            If Not IsNumeric(strArray(QueCampo)) Then
                AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
            Else
                ValorSQL = DBSet(strArray(QueCampo), "N")
            End If
                
        End If
    
    
    
        
    
    
    
    'Numerico NO requerido
    Case 4, 7, 14, 15
        
       
        strArray(QueCampo) = Replace(strArray(QueCampo), ".", ",")
        
     
        
        If QueCampo = 4 Or QueCampo = 14 Then
            'No es requerdio pero si es cero grabo cero
            ValorSQL = DBSet(strArray(QueCampo), "N", "N")
        Else
            ValorSQL = DBSet(strArray(QueCampo), "N", "S")
        End If
        If strArray(QueCampo) <> "" Then
            If Not IsNumeric(strArray(QueCampo)) Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
        End If
    
    'Fecha
    Case 1, 2
        
        If strArray(QueCampo) = "" Then strArray(QueCampo) = "1900-01-01"
        
        If Not IsDate(strArray(QueCampo)) Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo fecha"
        ValorSQL = DBSet(strArray(QueCampo), "F", "S")
        
    Case 5, 8, 16
        'Tipo operacion   QUECAMPO 5
        '        0 = GENERAL
        '        1 = INTRACOMUNITARIA
        '        2 = EXPORTACION
        '        3 = INTERIOR EXENTA
        ' Tipo retencion  QUECAMPO 8
        '        0 o nada, NO TIENE
        '        1 = PROFESIONAL
        '        2 = AGRICOLA
        '        3 = ARRENDAMIENTO
        'Tipo integracion
        '        0 = TODO
        '        1 = CONTA
        '        2 = TESORERIA
    
        If strArray(QueCampo) = "" Then
            strArray(QueCampo) = "0"
            If QueCampo = 16 Then strArray(QueCampo) = "1"
        End If
        If Not IsNumeric(strArray(QueCampo)) Then
            AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo numerico"
        Else
            strArray(QueCampo) = Abs(Val(strArray(QueCampo)))
            If QueCampo = 16 Then
                If Val(strArray(QueCampo)) > 2 Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " Valores: 0..2 o vacio"
            Else
                If Val(strArray(QueCampo)) > 3 Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " Valores: 0..3 o vacio"
            End If
       End If
       ValorSQL = DBSet(strArray(QueCampo), "N", "N")
    Case 10
        'Centro de coste
        'Si lleva analitica es obligatorio
        If strArray(QueCampo) = "" Then
            If vParam.autocoste Then AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " no puede estar vacio"
        Else
            ValorSQL = DBSet(strArray(QueCampo), "T")
        End If
            
        
        
'    Case 0
'        'Texto requerido
'        'Porque no puede ser NULO
'        If strArray(QueCampo) = "" Then
'            AnyadeEnErrores RecuperaValor(Msg$, QueCampo + 1) & " campo requerido"
'        Else
'            ValorSQL = DBSet(strArray(QueCampo), "T", "N")
'        End If
    
    
    Case Else
        ValorSQL = DBSet(strArray(QueCampo), "T", "S")
    End Select
    
    
    
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "," & ValorSQL
    
End Sub




Private Function ComprobacionDatosBDFacturasProveedoresSTD() As Boolean
Dim FechaMinima As Date

On Error GoTo eComprobacionDatosBD
    ComprobacionDatosBDFacturasProveedoresSTD = False
    Set miRsAux = New ADODB.Recordset
    
    '
    Cad = "select min(fecha) minima,max(fecha) from tmpintefrafracli where factura <>'' and codusu=" & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'no puede ser eof
    Cad = miRsAux!minima
    FechaMinima = miRsAux!minima
    If miRsAux!minima < vParam.fechaini Then
        AnyadeEnErrores "Menor que inicio ejercicio"
    Else
        If miRsAux!minima < vParam.FechaActiva Then
            AnyadeEnErrores "Menor que fecha activa"
        Else
            If miRsAux!minima <= UltimoDiaPeriodoLiquidado Then
            
                If Me.optVarios(0).Value Then
                    'FACTURAS CLIENTE. Obliado comprobar
                    AnyadeEnErrores "Menor que ultimo periodo liquidado"
                Else
                   
                    'PROVEEDOR. Va a pedir fecha liquidacion luego. No hacemos nada"
                                           
                   
                End If
            End If
        End If
    End If
    miRsAux.Close
    
    
    'Comprobaremos que todas las SERIES estan en contadores
    Cad = "select distinct(serie) from tmpintefrafracli where serie<>''  and codusu=" & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = DevuelveDesdeBD("tiporegi", "contadores", "tiporegi", miRsAux.Fields(0), "T")
        If Cad = "" Then
            Cad = miRsAux.Fields(0)
            AnyadeEnErrores "No existe contadores"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Cuentas
    Cad = "select distinct cta_cli from tmpintefrafracli where cta_cli<>''  and codusu=" & vUsu.Codigo & " and not cta_cli in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cta proveedor"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
        
    Cad = "select distinct ctaventas from tmpintefrafracli where ctaventas<>''  and codusu=" & vUsu.Codigo & " and not ctaventas in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cuenta ventas"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    'Retencion (si lleva
   
        'Cuenta clientes SI pide trae la cta reten
        Cad = "select distinct ctaret from tmpintefrafracli where ctaret<>''  and codusu=" & vUsu.Codigo & " and not ctaret in (select codmacta from cuentas where apudirec='S')"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = miRsAux.Fields(0)
            AnyadeEnErrores "No existe cuenta retencion"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
         
         
         
     
    
    'Facturas que ya existen
    
        Cad = "select factura,cta_cli from tmpintefrafracli where codigo<>''  and codusu=" & vUsu.Codigo
        Cad = Cad & " and (serie,factura,year(fecha),cta_cli ) iN (select numserie,numfactu,anofactu,codmacta from factpro where anofactu>=" & Year(FechaMinima) & " )"
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!FACTURA & "  " & miRsAux!cta_cli
        AnyadeEnErrores "Ya existe factura"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    Cad = " select FACTURA,fpago from tmpintefrafracli where integracion<>1 and factura<>''  and codusu=" & vUsu.Codigo
    Cad = Cad & " and not fpago iN (select codforpa from formapago )"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!FACTURA
        AnyadeEnErrores "No existe forma de pago " & miRsAux!fpago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si va a meter en cobros, que no existan en la tesoreria
    If Me.optVarios(1).Value Then
        Cad = "select factura from tmpintefrafracli where codigo>0  and codusu=" & vUsu.Codigo
        Cad = " select serie,factura,fecha from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
        Cad = Cad & " and  (serie,factura,fecha) iN (select numserie, numfactu, fecfactu from pagos where fecfactu>=" & DBSet(FechaMinima, "F") & " )"
    Else
        Cad = " select serie,factura,fecha from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
        Cad = Cad & " and  (serie,factura,fecha) iN (select numserie, numfactu, fecfactu from cobros )"
    End If
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!FACTURA
        AnyadeEnErrores "Ya existe vencimiento"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    If vParam.autocoste Then
        Cad = "select ccoste from tmpintefrafracli where ccoste<>''  and codusu=" & vUsu.Codigo
        Cad = Cad & " AND  not ccoste iN (select codccost from ccoste )"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            Cad = miRsAux!FACTURA
            AnyadeEnErrores "No existe centro de coste: " & miRsAux!CCoste
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    'Porcentaje IVA
    'Facturas que ya existen
    Cad = "select iva,recargo from tmpintefrafracli where codusu=" & vUsu.Codigo & " GROUP BY 1,2     "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = DBLet(miRsAux!recargo, "N") 'En tiposiva no puede ser null
        Cad = " tipodiva=0 and porcerec=" & Cad & " AND porceiva"
        Cad = DevuelveDesdeBD("codigiva", "tiposiva", Cad, DBSet(miRsAux!IVA, "N"))
        If Cad = "" Then
            Cad = "IVa: " & miRsAux!IVA & "% Rec: " & DBLet(miRsAux!recargo, "T") & "%"
            AnyadeEnErrores "No existe IVA"
            
        Else
        
                Cad = "UPDATE tmpintefrafracli SET tipoiva=" & Cad & " WHERE codusu =" & vUsu.Codigo
                Cad = Cad & " AND iva =" & DBSet(miRsAux!IVA, "N") & " AND recargo"
                If IsNull(miRsAux!recargo) Then
                    Cad = Cad & " is null"
                Else
                    Cad = Cad & "=" & DBSet(miRsAux!recargo, "N")
                End If
                Conn.Execute Cad
        
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si esta duplicado la factura
    Cad = "select factura,count(*) from tmpintefrafracli where factura>0 and codusu =" & vUsu.Codigo & " group by 1 having count(*)>1"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = "Factura: " & miRsAux!FACTURA & "   Veces: " & miRsAux.Fields(1)
        AnyadeEnErrores "Datos duplicados"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
                

    'Si no han habiado errores..  j=0
    ComprobacionDatosBDFacturasProveedoresSTD = J = 0


    

eComprobacionDatosBD:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
End Function





Private Sub InsertarEnContabilidadFraprovSTD()
Dim FACTURA As String
Dim BaseImponible As Currency
Dim IVA As Currency
Dim TotalFac As Currency
Dim ColBases As Collection
Dim Fecha As Date
Dim Aux As String
Dim ColCob As Collection
Dim tCobro As String
Dim fRecep  As Date
Dim Actual As Boolean
Dim Mc As Contadores
'Tipo integracion   Viene en el fichero en la posicion 17
'        0 = TODO
'        1 = CONTA
'        2 = TESORERIA
Dim Tipointegracion As Byte

    Set miRsAux = New ADODB.Recordset
    
    Label1.Caption = "Insertando contabilidad"
    Label1.Refresh
    
    'Valores por defecto
    Cad = DevuelveDesdeBD("codigo", "agentes", "1", "1 ORDER BY codigo")
    CadenaDesdeOtroForm = Cad & "|"
    Cad = DevuelveDesdeBD("codforpa", "formapago", "1", "1 ORDER BY codforpa")
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad & "|"
                         
                        
                        
    tCobro = "select ctabanc1,count(*) from cobros where fecfactu>=" & DBSet(DateAdd("y", -1, Now), "F") & " group by 1 order by 2 desc"
    miRsAux.Open tCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    tCobro = ""
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then tCobro = miRsAux.Fields(0)
    End If
    miRsAux.Close
    
    If tCobro = "" Then
        tCobro = DevuelveDesdeBD("codmacta", "bancos", "1", "1 ORDER BY 1 desc")
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & tCobro & "|"
    tCobro = ""
                        

        
    'codusu,Codigo,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,tipo_ret,ctaventas,Ccoste,impventa,
    'IVA , ImpIva, recargo, imprecargo, TotalFactura, integracion
    
    Cad = "select tmpintefrafracli.*,codpais,nifdatos,desprovi,despobla,codposta,dirdatos,nommacta,cuentas.iban ibancta from tmpintefrafracli left join cuentas ON cta_cli=codmacta"
    Cad = Cad & " WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        TotProgres = TotProgres + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    FACTURA = ""
    numProgres = 0
    Set Mc = New Contadores
    While Not miRsAux.EOF
        numProgres = numProgres + 1
        Label1.Caption = "Linea " & numProgres & "/" & TotProgres
        Label1.Refresh
    
        If DBLet(miRsAux!FACTURA, "T") = "" Then
            'Es otra base de la misma factura
            
        Else
            Cad = miRsAux!FACTURA
            
            If Cad <> FACTURA Then
                If FACTURA <> "" Then
                   
                    InsertarFacturaProvStandard ColBases, FACTURA, Fecha, Tipointegracion, "1", ColCob
                    
                    'Suma correcta
                    Cad = miRsAux!FACTURA
                End If
                'Factura NUEVA.
                
                'Haremos los inserts
                'INSERT INTO factpro (codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,totfaccl,trefaccl,totbasesret,totivas,
                'totbases,codagente,dpto,numasien,fecliqcl,retfaccl,cuereten,tiporeten,codopera,observa,numserie,numfactu,fecfactu,codmacta,
                'codforpa,anofactu)
                fRecep = miRsAux!Fecha
                Actual = fRecep <= vParam.fechafin
                
                If Mc.ConseguirContador(1, Actual, False) = 1 Then Exit Sub
                
                Tipointegracion = miRsAux!integracion

                NumRegElim = Mc.Contador
                BaseImponible = miRsAux!calculoimponible
                TotalFac = miRsAux!TotalFactura
                IVA = DBLet(miRsAux!ImpIva, "N")
                Msg$ = "(" & DBSet(miRsAux!codpais, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T", "S") & "," & DBSet(miRsAux!desProvi, "T", "S") & ","
                Msg$ = Msg$ & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S") & "," & DBSet(miRsAux!dirdatos, "T", "S") & ","
                Msg$ = Msg$ & DBSet(miRsAux!Nommacta, "T") & ",'0',0," & DBSet(miRsAux!cta_cli, "T") & "," & Year(fRecep)
                If DBLet(miRsAux!fpago, "T") = "" Then
                    'No viene forma de pago
                     K = RecuperaValor(CadenaDesdeOtroForm, 5)
                Else
                     K = miRsAux!fpago
                End If
                Msg$ = Msg$ & "," & K & ","
                'totbases,totbasesret,totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten,
                Msg$ = Msg$ & DBSet(miRsAux!impventa, "N") & "," & DBSet(miRsAux!impventa + IVA, "N") & "," & DBSet(IVA, "N")
                Msg$ = Msg$ & ",NULL," & DBSet(TotalFac, "N") & ","
                If DBLet(DBLet(miRsAux!tipo_ret, "N"), "N") = 0 Then
                    'No lleva retencion
                    Msg$ = Msg$ & "NULL,NULL,NULL,0"
                Else
                    Msg$ = Msg$ & DBSet(Round2((miRsAux!impret / (miRsAux!impventa + IVA)), 2) * 100, "N", "S") & "," & DBSet(miRsAux!impret, "N", "N")
                    Msg$ = Msg$ & ",'" & RecuperaValor(CadenaDesdeOtroForm, 1) & "'," & DBLet(miRsAux!tipo_ret, "N")    'tiporeten
                End If
                ',observa NUmSerie , Numregis, fecharec, NumFactu, FecFactu,fecliqpr
                
                Msg$ = Msg$ & ",'Importacion datos externos. " & " " & "Usuario: " & DevNombreSQL(vUsu.Nombre) & Chr(13) & "Fecha: " & Now
                Msg$ = Msg$ & "'," & DBSet(miRsAux!Serie, "T") & "," & NumRegElim & "," & DBSet(fRecep, "F") & ","
                Msg$ = Msg$ & DBSet(Format(miRsAux!FACTURA, "000000"), "T") & "," & DBSet(miRsAux!Fecha2, "F") & "," & DBSet(fRecep, "F") & ")"
                
                        
                    
                





                'select nommacta,dirdatos,despobla,codposta,desprovi,codpais,nifdatos from cuentas
                'Para los cobros     (cojo esta variable)
                '       insert into pagos(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,
                'numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,text41csb,numorden,fecvenci,impvenci)
                
                If Tipointegracion <> 1 Then
                    tCobro = " (" & DBSet(miRsAux!Nommacta, "T") & "," & DBSet(miRsAux!dirdatos, "T", "S")
                    tCobro = tCobro & "," & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S")
                    tCobro = tCobro & "," & DBSet(miRsAux!desProvi, "T", "S") & "," & DBSet(miRsAux!codpais, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T")
                    tCobro = tCobro & "," & DBSet(miRsAux!Serie, "T") & "," & DBSet(miRsAux!FACTURA, "T") & "," & DBSet(miRsAux!Fecha2, "F") & "," & DBSet(miRsAux!cta_cli, "T")
                    tCobro = tCobro & "," & K & "," & DBSet(RecuperaValor(CadenaDesdeOtroForm, 3), "T") & ","
                    
                    'Si me ha indicado IBAN, o no, en la importacion
                    If DBLet(miRsAux!IBAN, "T") = "" Then
                        Aux = DBSet(miRsAux!ibancta, "T", "S")
                    Else
                        Aux = DBSet(miRsAux!IBAN, "T", "S")
                    End If
                    tCobro = tCobro & Aux & ","
                    
                    'Si me ha indicado(o no) txtobserva text33csb text41csb
                    If DBLet(miRsAux!txtcsb, "T") = "" Then
                        Aux = DBSet("Factura : " & miRsAux!FACTURA & " de " & miRsAux!Fecha, "T") & ",null,"
                    Else
                        Aux = miRsAux!txtcsb
                        If Len(Aux) > 80 Then
                            Aux = DBSet(Mid(Aux, 1, 80), "T") & "," & DBSet(Mid(Aux, 81), "T") & ","
                        Else
                            Aux = DBSet(Aux, "T") & ",NULL,"
                        End If
                    End If
                    tCobro = tCobro & Aux
                    
                    CargarCobrosSobreCollectionConSQLInsert ColCob, CStr(K), miRsAux!Fecha, TotalFac, tCobro
                End If
                
                
                
                
                
                
                
                
                Fecha = miRsAux!Fecha
                i = 0
                FACTURA = Cad
                Set ColBases = Nothing
                Set ColBases = New Collection
                
            End If
        End If
        
        'Lineas
      
        i = i + 1
        Cad = "(0," & DBSet(miRsAux!recargo, "T") & "," & Year(Fecha) & ",NULL," 'codccost?
        Cad = Cad & DBSet(miRsAux!ImpIva, "N") & "," & DBSet(miRsAux!recargo, "N", "S") & "," & DBSet(miRsAux!IVA, "N")
        Cad = Cad & "," & DBSet(miRsAux!impventa, "N", "N") & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux!ctaventas, "T") & "," & i
        Cad = Cad & "," & DBSet(Fecha, "F", "N") & ",'" & miRsAux!Serie & "'," & NumRegElim & ")"
        
        ColBases.Add Cad
            
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If FACTURA <> "" Then
        InsertarFacturaProvStandard ColBases, FACTURA, Fecha, Tipointegracion, "1", ColCob
    End If
    
     Msg$ = ""
     Ampliacion = ""
    
    Label1.Caption = ""
    Set Mc = Nothing
    
    
End Sub




'Tipo integracion
'        0 = TODO
'        1 = CONTA
'        2 = TESORERIA
Private Function InsertarFacturaProvStandard(ByRef C As Collection, FacturaC As String, Fecha As Date, Tipointegracion As Byte, Serie_ As String, ByRef Cobros As Collection) As Boolean
Dim B As Boolean
 Dim RT As ADODB.Recordset
    On Error GoTo eInsertarFACTURA
    
    
    InsertarFacturaProvStandard = False
    Conn.BeginTrans

    If Tipointegracion <= 1 Then

        'Ejecutamos los sql
        
        Cad = "INSERT INTO factpro(codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,"
        Cad = Cad & "codconce340,codopera,codmacta,anofactu,codforpa,totbases,totbasesret,"
        Cad = Cad & "totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten"
        Cad = Cad & ",observa ,NUmSerie , Numregis, fecharec, NumFactu, FecFactu,fecliqpr) VALUES " & Msg$
    
        Conn.Execute Cad
        
        'Lineas
        Cad = ""
        For i = 1 To C.Count
            Cad = Cad & ", " & C.Item(i)
        Next i
        

        Cad = "fecharec,numserie,numregis) VALUES " & Mid(Cad, 2)
        Cad = "INSERT INTO factpro_lineas (aplicret,imporec,anofactu,codccost,impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea," & Cad

        Conn.Execute Cad
        
        'Totales
        espera 0.3
    
        
        Cad = " select codigiva, porciva, porcrec, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec "
    
        Cad = Cad & " from factpro_lineas"
        Cad = Cad & " where numserie =  " & DBSet(Serie_, "T") & " and numregis = " & NumRegElim & " and anofactu = " & Year(Fecha)
        Cad = Cad & " group by 1,2,3"
        Cad = Cad & " order by 1,2,3"
        Set RT = New ADODB.Recordset
        RT.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        i = 0
        Cad = ""
        While Not RT.EOF
            i = i + 1
                Cad = Cad & ", (" & DBSet(Serie_, "T") & "," & NumRegElim & "," & DBSet(Fecha, "F") & "," & Year(Fecha)
            Cad = Cad & "," & i & "," & DBSet(RT!Baseimpo, "N") & "," & RT!codigiva & "," & DBSet(RT!porciva, "N", "S")
            Cad = Cad & "," & DBSet(RT!porcrec, "N") & "," & DBSet(RT!Imporiva, "N") & "," & DBSet(RT!imporrec, "N") & ")"
            RT.MoveNext  'rT!
        Wend
        RT.Close
        Set RT = Nothing
        Cad = Mid(Cad, 2)
        
         
        Cad = "insert into factpro_totales (numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)  VALUES " & Cad
        Conn.Execute Cad
        
    
        'Creamos el apunte
         
        B = IntegrarFactura(Serie_, NumRegElim, Fecha)
       
        If Not B Then
            Cad = "Creando asiento. "
            AnyadeEnErrores FacturaC
        End If
    
    
    
    End If
    
    
    'Tesoreria
    If Tipointegracion <> 1 Then
        'insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,
        'numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci,ctabanc1,iban,text33csb)
        Cad = ""
        For i = 1 To Cobros.Count
            Cad = Cad & ", " & Cobros.Item(i)
        Next i
        Cad = Mid(Cad, 2)
        
        Cad = "INSERT INTO pagos(nomprove,domprove,pobprove,cpprove,proprove,codpais,nifprove,numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text1csb,text2csb,numorden,fecefect,impefect) VALUES " & Cad
        
        Conn.Execute Cad
    End If
    
    'Adelante transaccion
    Conn.CommitTrans
    InsertarFacturaProvStandard = True
    
    
    
    Exit Function
eInsertarFACTURA:
    Cad = Err.Description
    Err.Clear
    Conn.RollbackTrans
    
   
    AnyadeEnErrores FacturaC
    
    
    
End Function



Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub







Private Sub CargaEncabezadoSAGE()
Dim clmX As ColumnHeader
Dim B As Boolean

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.Checkboxes = False
    
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Serie"
        clmX.Width = 900 '1500
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Factura"
        clmX.Width = 1500
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Fecha"
        clmX.Width = 1300
            
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Codmacta"
        clmX.Width = 1500
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Nombre"
        clmX.Width = 3900
        
    
        For i = 1 To 4
            Set clmX = ListView1.ColumnHeaders.Add()
            clmX.Text = RecuperaValor("Suplidos|Base|IVA|Total|", CInt(i))
            clmX.Width = 1600
            clmX.Alignment = lvwColumnRight
        Next
    
End Sub



Private Function TraspasaApuSAGE() As Boolean
    On Error GoTo eTraspasaApuSAGE
    TraspasaApuSAGE = False
    

    
    
    FacturaAnterior = "Integracion facturas CSV-SAGE.  " & Now & vbCrLf & vUsu.Nombre
    
    Cad = "INSERT INTO hcabapu(numdiari,fechaent,numasien,obsdiari,feccreacion,usucreacion,desdeaplicacion)"
    Cad = Cad & " SELECT numdiari, fechaent,numasien," & DBSet(FacturaAnterior, "T") & "," & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'Ariconta6' "
    Cad = Cad & " FROM tmpintegrapu WHERE codusu =" & vUsu.Codigo & " GROUP BY numasien,fechaent"
    Conn.Execute Cad
    
    Cad = "INSERT INTO hlinapu(numdiari,fechaent,numasien,linliapu,codmacta,numdocum,codconce,ampconce,timporteD,codccost,timporteH,ctacontr,idcontab) "
    Cad = Cad & " SELECT NumDiari , FechaEnt, NumAsien, Linliapu, codmacta, Numdocum, CodConce, Ampconce, timporteD, codccost, timporteH, ctacontr, idcontab"
    Cad = Cad & " FROM tmpintegrapu WHERE codusu =" & vUsu.Codigo & " ORDER BY numasien,fechaent,linliapu"
    Conn.Execute Cad
    
    
    Cad = "select codigo,texto from tmptesoreriacomun where codusu=" & vUsu.Codigo & " order by codigo"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = miRsAux!TEXTO
        PonerLabel "Sec " & miRsAux!Codigo
        Cad = Replace(Cad, "·", "'")
        Conn.Execute Cad
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    
    
    'Para cada factura creada, updateo su verdadero NUMERO de apunte
    PonerLabel "Ajustando nºasiento factura"
    Cad = "select serie,factura,year(fecha) anyo,txtcsb from tmpintefrafracli  WHERE codusu =" & vUsu.Codigo
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = "UPDATE factcli set numasien = " & miRsAux!txtcsb & " WHERE numserie ='" & miRsAux!Serie & "' AND numfactu =" & miRsAux!FACTURA
        Cad = Cad & " AND anofactu = " & miRsAux!Anyo
        Conn.Execute Cad
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    
    
    TraspasaApuSAGE = True
    
    Exit Function
    
eTraspasaApuSAGE:
    Cad = Err.Description
    Ejecuta "DELETE FROM     tmptesoreriacomun where codusu =" & vUsu.Codigo
    
    Msg = "INSERT INTO  tmptesoreriacomun(codusu,codigo,texto1,texto2,observa1) VALUES (" & vUsu.Codigo & ",1000,'Error',''," & DBSet(Cad, "T") & ")"
    Conn.Execute Msg
    J = 1
    
End Function
