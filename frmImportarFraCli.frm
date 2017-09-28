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
      Left            =   12960
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
      Width           =   4125
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
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Asientos"
         Enabled         =   0   'False
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
         Left            =   13320
         TabIndex        =   6
         Top             =   600
         Width           =   2535
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
         Left            =   9960
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


Private Const NumeroCamposTratarFraCli = 20



Dim strArray() As String
Dim LineaConErrores As Boolean
Dim cad As String
Dim Importe As Currency
Dim FacturaAnterior As String
Dim NF As Integer


Dim NumeroCamposTratarFraPro As Integer


Private Sub cmdAceptar_Click()
    
    ImportarFraCLI

End Sub

Private Sub ImportarFraCLI()
    
    If Me.optVarios(0).Value Then
        
        If ListView1.ListItems.Count > 0 Then
            If MsgBox("¿Continuar con el proceso de importacion?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
        
    Else
        'PRoveedor
        'If Me.cboTipo.ListIndex Then
        CadenaDesdeOtroForm = ""
        frmMensajes.Opcion = 59
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm = "" Then Exit Sub
    
    End If
    
    
    'Proceso realmente dicho
    Screen.MousePointer = vbHourglass
    If Me.optVarios(0).Value Then
        InsertarEnContabilidadFraCli
    Else
    
        InsertarEnContabilidadFraProveedor
        
    End If
        
        If J > 0 Then
            ListView1.Tag = 0
            CargaEncabezado 2
            Set miRsAux = New ADODB.Recordset
            cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
            miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            Me.ListView1.ListItems.Clear
            MsgBox "Proceso finalizado con exito", vbInformation
        End If
    
    Me.cmdAceptar.visible = False
    Set miRsAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancelar_Click()
    If cmdAceptar.visible Then
        If ListView1.ListItems.Count > 0 Then
            cad = "¿Desea cancelar el proceso de importación?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Command1_Click()

    If Text1.Text = "" Then Exit Sub
    If ListView1.ListItems.Count > 0 Then
        'Volver a cargar
        If cmdAceptar.visible Then
            'Importacion anterior con datos correctos.
            'Preguntamos
            cad = "Hay datos correctos pendientes de integrar. " & vbCrLf & "Cancelar proceso  anterior?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    cmdAceptar.visible = False
    ListView1.ListItems.Clear
    ListView1.Tag = 0 'No puede hacer dblclick
    


    Screen.MousePointer = vbHourglass
    If Me.optVarios(0).Value Then
        ImportarFacturasCliente
        
    Else
        If Me.optVarios(1).Value Then ImportacionFraPro
        
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
    cmdAceptar.visible = False
    Label1.Caption = ""


    
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
Dim RC As Byte
    
    'Primer paso. Lectura fichero. Comprobacion basica datos
    PonerLabel "Leyendo fichero"
    RC = ImportarFichFracli
    If RC = 2 Then Exit Sub
        
    If RC = 1 Then
        'Errores en fichero
        'Ha habido errores
        CargaEncabezado 0
    
        'Cargamos datos
        Set miRsAux = New ADODB.Recordset
        cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
   
    cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on cta_cli=cuentas.codmacta where codusu= " & vUsu.Codigo
    cad = cad & " ORDER BY codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        ListView1.ListItems.Add , "K" & i
        If DBLet(miRsAux!Serie, "T") <> "" Then
            
            ListView1.ListItems(i).Text = miRsAux!Serie
            ListView1.ListItems(i).SubItems(1) = Format(miRsAux!FACTURA, "000000")
            ListView1.ListItems(i).SubItems(2) = miRsAux!Fecha
            
            ListView1.ListItems(i).SubItems(3) = miRsAux!Nommacta
            
            
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!impventa, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            ListView1.ListItems(i).SubItems(7) = " " 'separado
            
            ListView1.ListItems(i).SubItems(8) = Format(miRsAux!CalculoImponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(9) = Format(miRsAux!TotalFactura - miRsAux!CalculoImponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(10) = Format(miRsAux!TotalFactura, FormatoImporte)
            
            If vEmpresa.TieneTesoreria Then
                cad = " "
                If DBLet(miRsAux!IBAN, "T") <> "" Or DBLet(miRsAux!txtcsb, "T") <> "" Then cad = "*"
                ListView1.ListItems(i).SubItems(11) = cad
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
Dim b As Boolean

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
        b = True
        If Me.optVarios(1).Value Then
            'Proveedor. De momento para cualquiera de proveedor
            b = False
        End If
        If b Then
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
        Line Input #NF, cad
        
        If NumRegElim = 1 Then
            'Primera linea encabezado?
            If Me.check1.Value = 1 Then cad = ""
        Else
            If InStr(1, String(NumeroCamposTratarFraCli, ";"), cad) > 0 Then cad = "" 'todo puntos y comas
        End If
        
        
        If cad <> "" Then
            'Procesamos linea
            
            strArray = Split(cad, ";")
            
            If UBound(strArray) = NumeroCamposTratarFraCli - 1 Then
                'Falta el ultimo punto y coma
                cad = cad & ";"
                strArray = Split(cad, ";")
            End If
            
            
            If UBound(strArray) <> NumeroCamposTratarFraCli Then
                J = J + 1
                Aux = vUsu.Codigo & "," & J & "," & NumRegElim & ",'Nº campos incorrecto'," & DBSet(cad, "T")
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
        cad = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
        If Val(cad) > 0 Then
            ImportarFichFracli = 1 'Con errores
        Else
            cad = DevuelveDesdeBD("count(*)", "tmpintefrafracli", "codusu", vUsu.Codigo)
            If Val(cad) > 0 Then ImportarFichFracli = 0
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
        If strArray(0) <> "" Or strArray(1) <> "0" Or strArray(2) <> "1900-01-01" Then
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
         If strArray(0) <> "" And strArray(1) <> "0" And strArray(2) <> "1900-01-01" Then FacturaAnterior = strArray(0) & "|" & strArray(1) & "|" & strArray(2) & "|"
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
    '  11       12          13         14       15          16              17          18          19
    'IMP.VENTA  % I.V.A.  MP.IVA    % R.E.    IMP. REC.   TOTAL FACTURA   INTEGRACION  iban     txtcsb
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
    QueError = vUsu.Codigo & "," & J & "," & NumRegElim & "," & DBSet(QueError, "T") & "," & DBSet(cad, "T")
    Conn.Execute "INSERT INTO tmptesoreriacomun (codusu,codigo,texto1,texto2,observa1) VALUES (" & QueError & ")"
    LineaConErrores = True
End Sub



Private Function ComprobacionDatosBDFacturas() As Boolean
Dim FechaMinima As Date

On Error GoTo eComprobacionDatosBD
    ComprobacionDatosBDFacturas = False
    Set miRsAux = New ADODB.Recordset
    
    
    cad = "select min(fecha) minima,max(fecha) from tmpintefrafracli where factura >0 and codusu=" & vUsu.Codigo
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'no puede ser eof
    cad = miRsAux!minima
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
    cad = "select distinct(serie) from tmpintefrafracli where serie<>''  and codusu=" & vUsu.Codigo
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = DevuelveDesdeBD("tiporegi", "contadores", "tiporegi", miRsAux.Fields(0), "T")
        If cad = "" Then
            cad = miRsAux.Fields(0)
            AnyadeEnErrores "No existe contadores"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Cuentas
    cad = "select distinct cta_cli from tmpintefrafracli where cta_cli<>''  and codusu=" & vUsu.Codigo & " and not cta_cli in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cta cliente"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
        
    cad = "select distinct ctaventas from tmpintefrafracli where ctaventas<>''  and codusu=" & vUsu.Codigo & " and not ctaventas in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cuenta ventas"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    'Retencion (si lleva
    If Me.optVarios(0).Value Then
        'Cuenta clientes SI pide trae la cta reten
        cad = "select distinct ctaret from tmpintefrafracli where ctaret<>''  and codusu=" & vUsu.Codigo & " and not ctaret in (select codmacta from cuentas where apudirec='S')"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = miRsAux.Fields(0)
            AnyadeEnErrores "No existe cuenta ventas"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
         
    End If
         
         
     
    
    'Facturas que ya existen
    If Me.optVarios(1).Value Then
        cad = "select factura,cta_cli from tmpintefrafracli where codigo>0  and codusu=" & vUsu.Codigo
        cad = cad & " and (serie,factura,year(fecha),cta_cli ) iN (select numserie,numfactu,anofactu,codmacta from factpro where anofactu>=" & Year(FechaMinima) & " )"
    Else
        cad = "select factura,cta_cli as cta_cli  from tmpintefrafracli where codigo>0  and codusu=" & vUsu.Codigo
        cad = cad & " and (serie,factura,year(fecha)) iN (select numserie,numfactu,anofactu from factcli where anofactu>=" & Year(FechaMinima) & " )"
    End If
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux!FACTURA & "  " & miRsAux!cta_cli
        AnyadeEnErrores "Ya existe factura"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    cad = " select FACTURA,fpago from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
    cad = cad & " and not fpago iN (select codforpa from formapago )"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux!FACTURA
        AnyadeEnErrores "No existe forma de pago " & miRsAux!fpago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si va a meter en cobros, que no existan en la tesoreria
    If Me.optVarios(1).Value Then
        cad = "select factura from tmpintefrafracli where codigo>0  and codusu=" & vUsu.Codigo
        cad = " select serie,factura,fecha from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
        cad = cad & " and  (serie,factura,fecha) iN (select numserie, numfactu, fecfactu from pagos where fecfactu>=" & DBSet(FechaMinima, "F") & " )"
    Else
        cad = " select serie,factura,fecha from tmpintefrafracli where integracion<>1 and factura>0  and codusu=" & vUsu.Codigo
        cad = cad & " and  (serie,factura,fecha) iN (select numserie, numfactu, fecfactu from cobros )"
    End If
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux!FACTURA
        AnyadeEnErrores "Ya existe vencimiento"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    If vParam.autocoste Then
        cad = "select ccoste from tmpintefrafracli where ccoste<>''  and codusu=" & vUsu.Codigo
        cad = cad & " AND  not ccoste iN (select codccost from ccoste )"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = miRsAux!FACTURA
            AnyadeEnErrores "No existe centro de coste: " & miRsAux!ccoste
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    
    'Porcentaje IVA
    'Facturas que ya existen
    cad = "select iva,recargo from tmpintefrafracli where codusu=" & vUsu.Codigo & " GROUP BY 1,2     "
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = DBLet(miRsAux!recargo, "N") 'En tiposiva no puede ser null
        cad = "porcerec=" & cad & " AND porceiva"
        cad = DevuelveDesdeBD("codigiva", "tiposiva", cad, DBSet(miRsAux!IVA, "N"))
        If cad = "" Then
            cad = "IVa: " & miRsAux!IVA & "% Rec: " & DBLet(miRsAux!recar, "T") & "%"
            AnyadeEnErrores "No existe IVA"
            
        Else
        
                cad = "UPDATE tmpintefrafracli SET tipoiva=" & cad & " WHERE codusu =" & vUsu.Codigo
                cad = cad & " AND iva =" & DBSet(miRsAux!IVA, "N") & " AND recargo"
                If IsNull(miRsAux!recargo) Then
                    cad = cad & " is null"
                Else
                    cad = cad & "=" & DBSet(miRsAux!recargo, "N")
                End If
                Conn.Execute cad
        
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si esta duplicado la factura
    cad = "select factura,count(*) from tmpintefrafracli where factura>0 and codusu =" & vUsu.Codigo & " group by 1 having count(*)>1"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = "Factura: " & miRsAux!FACTURA & "   Veces: " & miRsAux.Fields(1)
        AnyadeEnErrores "Datos duplicados"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
                

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

Dim SegundoImporteAuxiliar As Currency
Dim b As Boolean

    On Error GoTo eComprobarTotales
    ComprobarTotales = False


    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    DiferenciaMinimaPermitida = 0
    If Me.optVarios(1).Value Then
        'Proveedores
        If Me.cboTipo.ListIndex = 0 Then DiferenciaMinimaPermitida = 0.2
    End If
    cad = "select * from tmpintefrafracli WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Dim Fin As Boolean
    Fin = False
    While Not miRsAux.EOF
        If miRsAux!Serie = "" Then
            'Es otra base de la misma factura
            
        Else
            cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
            
            If cad <> FACTURA Then
                If FACTURA <> "" Then
                    'Suma correcta
                    SegundoImporteAuxiliar = TotalFac - (BaseImponible + IVA - ImporteRetencion)
                    Ok = True
                    If Abs(SegundoImporteAuxiliar) > DiferenciaMinimaPermitida Then Ok = False
                    
                   
                    cad = "Calculado/fichero:  " & BaseImponible + IVA - ImporteRetencion & " / " & TotalFac
                    If ImporteRetencion <> 0 Then cad = cad & " Retencion Tipo" & TipoRetencion & "  Importe: " & ImporteRetencion
                    
                    If Not Ok Then
                        
                        AnyadeEnErrores "Total factura"
                        
                    Else
                        cad = " AND Serie= " & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & " AND FACTURA =" & Mid(FACTURA, 4)
                        cad = "UPDATE tmpintefrafracli set CalculoImponible=" & DBSet(BaseImponible, "N") & "  WHERE codusu = " & vUsu.Codigo & cad
                        Conn.Execute cad
                    End If
                    cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
                End If
                'Factura NUEVA. Reseteamos importes...
                TipoRetencion = DBLet(miRsAux!tipo_ret, "N")
                ImporteRetencion = 0
                If TipoRetencion <> 0 Then ImporteRetencion = DBLet(miRsAux!impret, "N")
                NumRegElim = miRsAux!FACTURA
                BaseImponible = 0
                IVA = 0
                TotalFac = miRsAux!TotalFactura
                
                FACTURA = cad
                
            End If
        End If
        
        Importe = (miRsAux!impventa * miRsAux!IVA) / 100
        Importe = Round2(Importe, 2)
        SegundoImporteAuxiliar = Importe - miRsAux!ImpIva
        b = False
        If Abs(SegundoImporteAuxiliar) > DiferenciaMinimaPermitida Then b = True
        If b Then
            cad = "IVA calculado/fichero " & miRsAux!IVA & "%: " & Importe & " / " & miRsAux!ImpIva
            AnyadeEnErrores "Calculo  IVA"

        End If
        
        BaseImponible = BaseImponible + miRsAux!impventa
        IVA = IVA + Importe
        
            
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If FACTURA <> "" Then
        'Suma correcta
        SegundoImporteAuxiliar = TotalFac - (BaseImponible + IVA - ImporteRetencion)
        b = Abs(SegundoImporteAuxiliar) > DiferenciaMinimaPermitida
        If b Then
            cad = "Calculado/fichero:  " & BaseImponible + IVA - ImporteRetencion & " / " & TotalFac
            AnyadeEnErrores "Total factura"
        Else
            cad = " AND Serie= " & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & " AND FACTURA =" & Mid(FACTURA, 4)
            cad = "UPDATE tmpintefrafracli set CalculoImponible=" & DBSet(BaseImponible, "N") & "  WHERE codusu = " & vUsu.Codigo & cad
            Conn.Execute cad
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


Private Sub InsertarEnContabilidadFraCli()
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
    
    
    'Valores por defecto
    cad = DevuelveDesdeBD("codigo", "agentes", "1", "1 ORDER BY codigo")
    CadenaDesdeOtroForm = cad & "|"
    cad = DevuelveDesdeBD("codforpa", "formapago", "1", "1 ORDER BY codforpa")
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & cad & "|"
                        
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
    
    cad = "select tmpintefrafracli.*,codpais,nifdatos,desprovi,despobla,codposta,dirdatos,nommacta,cuentas.iban ibancta from tmpintefrafracli left join cuentas ON cta_cli=codmacta"
    cad = cad & " WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    FACTURA = ""
    While Not miRsAux.EOF
        If miRsAux!Serie = "" Then
            'Es otra base de la misma factura
            
        Else
            cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
            
            If cad <> FACTURA Then
                If FACTURA <> "" Then
                   
                    InsertarFACTURA ColBases, FACTURA, Fecha, Tipointegracion, ColCob
                    
                    'Suma correcta
                    cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
                End If
                'Factura NUEVA.
                
                'Haremos los inserts
                'INSERT INTO factcli (codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,totfaccl,trefaccl,totbasesret,totivas,
                'totbases,codagente,dpto,numasien,fecliqcl,retfaccl,cuereten,tiporeten,codopera,observa,numserie,numfactu,fecfactu,codmacta,
                'codforpa,anofactu)
                Tipointegracion = miRsAux!integracion
                
                
                NumRegElim = miRsAux!FACTURA
                BaseImponible = miRsAux!CalculoImponible
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
                FACTURA = cad
                Set ColBases = Nothing
                Set ColBases = New Collection
                
            End If
        End If
        
        'Lineas
        'INSERT INTO factcli_lineas (aplicret,imporec,anofactu,codccost,
        'impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea,fecfactu,numserie,numfactu
        i = i + 1
        cad = "(0," & DBSet(miRsAux!recargo, "T") & "," & Year(Fecha) & ",NULL," 'codccost?
        cad = cad & DBSet(miRsAux!ImpIva, "N") & "," & DBSet(miRsAux!recargo, "N", "S") & "," & DBSet(miRsAux!IVA, "N")
        cad = cad & "," & DBSet(miRsAux!impventa, "N", "N") & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux!ctaventas, "T") & "," & i
        cad = cad & "," & DBSet(Fecha, "F", "N") & "," & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & "," & DBSet(Mid(FACTURA, 4), "N") & ")"
        
        ColBases.Add cad
            
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If FACTURA <> "" Then
        InsertarFACTURA ColBases, FACTURA, Fecha, Tipointegracion, ColCob
    End If
    
     Msg$ = ""
     Ampliacion = ""
    
    
    
    
    
End Sub




'Tipo integracion
'        0 = TODO
'        1 = CONTA
'        2 = TESORERIA
Private Function InsertarFACTURA(ByRef C As Collection, FacturaC As String, Fecha As Date, Tipointegracion As Byte, ByRef Cobros As Collection) As Boolean
Dim b As Boolean
 Dim RT As ADODB.Recordset
    On Error GoTo eInsertarFACTURA
    
    
    InsertarFACTURA = False
    Conn.BeginTrans

    If Tipointegracion <= 1 Then

        'Ejecutamos los sql
        If Me.optVarios(1).Value Then
            cad = "INSERT INTO factpro(codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,"
            cad = cad & "codconce340,codopera,codmacta,anofactu,codforpa,totbases,totbasesret,"
            cad = cad & "totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten"
            cad = cad & ",observa ,NUmSerie , Numregis, fecharec, NumFactu, FecFactu,fecliqpr) VALUES " & Msg$
            
        Else
            'Cliente
            cad = "INSERT INTO factcli (codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,totfaccl,trefaccl,totbasesret,totivas,"
            cad = cad & "totbases,codagente,dpto,numasien,fecliqcl,retfaccl,cuereten,tiporeten,codopera,observa,numserie,numfactu,fecfactu,codmacta,"
            cad = cad & "codforpa,anofactu) VALUES " & Msg$
        End If
        Conn.Execute cad
        
        'Lineas
        cad = ""
        For i = 1 To C.Count
            cad = cad & ", " & C.Item(i)
        Next i
        
       
        If Me.optVarios(1).Value Then
            cad = "fecharec,numserie,numregis) VALUES " & Mid(cad, 2)
            cad = "INSERT INTO factpro_lineas (aplicret,imporec,anofactu,codccost,impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea," & cad
        Else
             cad = "fecfactu,numserie,numfactu) VALUES " & Mid(cad, 2)
            cad = "INSERT INTO factcli_lineas (aplicret,imporec,anofactu,codccost,impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea," & cad
        End If
        Conn.Execute cad
        
        'Totales
        espera 0.3
    
        
        cad = " select codigiva, porciva, porcrec, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec "
        If Me.optVarios(1).Value Then
            cad = cad & " from factpro_lineas"
            cad = cad & " where numserie = " & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & " and numregis = " & NumRegElim & " and anofactu = " & Year(Fecha)
        Else
            cad = cad & " from factcli_lineas"
            cad = cad & " where numserie = " & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & " and numfactu = " & DBSet(Mid(FacturaC, 4), "N") & " and anofactu = " & Year(Fecha)
        End If
        cad = cad & " group by 1,2,3"
        cad = cad & " order by 1,2,3"
        Set RT = New ADODB.Recordset
        RT.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        i = 0
        cad = ""
        While Not RT.EOF
            i = i + 1
            If Me.optVarios(1).Value Then
                cad = cad & ", (" & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & "," & NumRegElim & "," & DBSet(Fecha, "F") & "," & Year(Fecha)
            Else
                cad = cad & ", (" & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & "," & DBSet(Mid(FacturaC, 4), "N") & "," & DBSet(Fecha, "F") & "," & Year(Fecha)
            End If
            cad = cad & "," & i & "," & DBSet(RT!Baseimpo, "N") & "," & RT!codigiva & "," & DBSet(RT!porciva, "N", "S")
            cad = cad & "," & DBSet(RT!porcrec, "N") & "," & DBSet(RT!Imporiva, "N") & "," & DBSet(RT!imporrec, "N") & ")"
            RT.MoveNext  'rT!
        Wend
        RT.Close
        Set RT = Nothing
        cad = Mid(cad, 2)
        
         If Me.optVarios(1).Value Then
            cad = "insert into factpro_totales (numserie,numregis,fecharec,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)  VALUES " & cad
        Else
            cad = "insert into factcli_totales (numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)  VALUES " & cad
        End If
        Conn.Execute cad
        
    
        'Creamos el apunte
         If Me.optVarios(1).Value Then
            b = IntegrarFactura(Trim(Mid(FacturaC, 1, 3)), NumRegElim, Fecha)
        Else
            b = IntegrarFactura(Trim(Mid(FacturaC, 1, 3)), CLng(Mid(FacturaC, 4)), Fecha)
        End If
        If Not b Then
            cad = "Creando asiento. "
            AnyadeEnErrores FacturaC
        End If
    
    
    
    End If
    
    
    'Tesoreria
    If Tipointegracion <> 1 Then
        'insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,
        'numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci,ctabanc1,iban,text33csb)
        cad = ""
        For i = 1 To Cobros.Count
            cad = cad & ", " & Cobros.Item(i)
        Next i
        cad = Mid(cad, 2)
        
        If Me.optVarios(1).Value Then
            cad = "INSERT INTO pagos(nomprove,domprove,pobprove,cpprove,proprove,codpais,nifprove,numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text1csb,text2csb,numorden,fecefect,impefect) VALUES " & cad
        Else
            cad = "insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,text41csb,numorden,fecvenci,impvenci) VALUES " & cad
        End If
        Conn.Execute cad
    End If
    
    'Adelante transaccion
    Conn.CommitTrans
    InsertarFACTURA = True
    'Borramos de la tabla temporal
    
    
    Exit Function
eInsertarFACTURA:
    cad = Err.Description
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



Private Sub imgppal_Click(Index As Integer)
    With frmppal.cd1
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
            cad = "serie =" & DBSet(ListView1.SelectedItem.Text, "T") & " AND  factura =" & ListView1.SelectedItem.SubItems(1) & " AND codusu"
            cad = DevuelveDesdeBD("concat(coalesce(iban,''),'|',coalesce(txtcsb,''),'|')", "tmpintefrafracli", cad, CStr(vUsu.Codigo))
            If cad <> "" Then
                With ListView1.SelectedItem
                    CadenaDesdeOtroForm = "Cliente: " & .SubItems(3) & vbCrLf
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Factura: " & .Text & .SubItems(1) & vbCrLf & vbCrLf
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "IBAN: " & RecuperaValor(cad, 1) & vbCrLf
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "TxtCSB: " & RecuperaValor(cad, 2) & vbCrLf
                End With
                MsgBox CadenaDesdeOtroForm, vbInformation
            End If
        End If
    End If
End Sub

Private Sub optVarios_Click(Index As Integer)
   
    
    If Index = 1 Then
        cboTipo.Clear
        'cboTipo.AddItem "Genérica"
        cboTipo.AddItem "Navarrés"
        cboTipo.ListIndex = 0  'De momento esta por defecto
    End If
     
    Me.Label2.visible = Index = 1
    Me.cboTipo.visible = Index = 1
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
    
    
    If cboTipo.ListIndex = 0 Then ImportacionNavarresFraPro
    
End Sub


'--------------------------------------------------------------------
'TIPO NAvarres
'--------------------------------------------------------------------
Private Sub ImportacionNavarresFraPro()
Dim RC As Byte
Dim b As Boolean

 'Primer paso. Lectura fichero. Comprobacion basica datos
    PonerLabel "Leyendo fichero"
    NumeroCamposTratarFraPro = 12
    RC = ImportarFichFraPronav
    If RC = 2 Then Exit Sub
        
    If RC = 1 Then
        'Errores en fichero
        'Ha habido errores
        CargaEncabezado 0
    
        'Cargamos datos
        Set miRsAux = New ADODB.Recordset
        cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        cad = "select codigo,texto1,texto2,observa1 from tmptesoreriacomun where codusu =" & vUsu.Codigo & " ORDER BY codigo"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    b = True
    If Me.optVarios(1).Value Then b = False
    
    Set miRsAux = New ADODB.Recordset
   
    cad = "select tmpintefrafracli.*,nommacta from tmpintefrafracli left join cuentas on cta_cli=cuentas.codmacta where codusu= " & vUsu.Codigo
    cad = cad & " ORDER BY codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
            If b Then
                ListView1.ListItems(i).SubItems(8) = Format(miRsAux!CalculoImponible, FormatoImporte)
            Else
                ListView1.ListItems(i).SubItems(8) = Format(DBLet(miRsAux!impret, "N"), FormatoImporte)
            End If
            ListView1.ListItems(i).SubItems(9) = Format(miRsAux!TotalFactura - miRsAux!CalculoImponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(10) = Format(miRsAux!TotalFactura, FormatoImporte)
            
            If vEmpresa.TieneTesoreria Then
                cad = " "
                If DBLet(miRsAux!IBAN, "T") <> "" Or DBLet(miRsAux!txtcsb, "T") <> "" Then cad = "*"
                ListView1.ListItems(i).SubItems(11) = cad
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
        Line Input #NF, cad
        
        If NumRegElim = 1 Then
            'Primera linea encabezado?
            If Me.check1.Value = 1 Then cad = ""
        Else
            If InStr(1, String(NumeroCamposTratarFraPro, ";"), cad) > 0 Then cad = "" 'todo puntos y comas
        End If
        
        
        If cad <> "" Then
            'Procesamos linea
            
            strArray = Split(cad, ";")
            
            If UBound(strArray) = NumeroCamposTratarFraPro - 1 Then
                'Falta el ultimo punto y coma
                cad = cad & ";"
                strArray = Split(cad, ";")
            End If
            
            
            If UBound(strArray) <> NumeroCamposTratarFraPro Then
                J = J + 1
                Aux = vUsu.Codigo & "," & J & "," & NumRegElim & ",'Nº campos incorrecto'," & DBSet(cad, "T")
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
        cad = DevuelveDesdeBD("count(*)", "tmptesoreriacomun", "codusu", vUsu.Codigo)
        If Val(cad) > 0 Then
            ImportarFichFraPronav = 1 'Con errores
        Else
            cad = DevuelveDesdeBD("count(*)", "tmpintefrafracli", "codusu", vUsu.Codigo)
            If Val(cad) > 0 Then ImportarFichFraPronav = 0
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
Dim cad As String

    CadenaDesdeOtroForm = "INSERT INTO tmpintefrafracli(codusu,Codigo,serie,factura,cta_cli,fecha,impventa,iva,impiva,CalculoImponible,"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " impret,ctaret,tipo_ret,integracion , IBAN, TotalFactura, txtcsb,ctaventas"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ") VALUES (" & vUsu.Codigo & "," & NumRegElim & ",1" 'serie
    
    'Factura Socio   Fecha   Bruto   % IVA   IVA    Base IRPF   IRPF    Fondo   Cobrado    Cosecha
    '   1   8329    170407  1.044,28    12  125,32  1.169,60    23,41   0      1.146,19    OROGRANDE


    LineaConErrores = False
    NuevaLinea = True
        
    'Si pone algun dato de la factura, debe ponerlos todos
    If strArray(0) = "" Or strArray(1) = "" Or strArray(2) = "" Then AnyadeEnErrores "Campos facturas obligados"
    cad = strArray(2)
    
    strArray(2) = Mid(cad, 5, 2) & "/" & Mid(cad, 3, 2) & "/20" & Mid(cad, 1, 2)
     
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
            If strArray(K) = "8275" Then Stop
            If strArray(K) = "8293" Then Stop
            If strArray(K) = "7529" Then Stop
            
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






Private Sub InsertarEnContabilidadFraProveedor()
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
    cad = DevuelveDesdeBD("codigo", "agentes", "1", "1 ORDER BY codigo")
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & cad & "|"
    
    
    cad = "tipforpa=1 AND nomforpa like '%trans%' AND 1"
    cad = DevuelveDesdeBD("codforpa", "formapago", cad, "1 ORDER BY codforpa")
    
    'FORMA PAGO transferencia
    If cad = "" Then
        cad = "tipforpa=1 AND 1"
        cad = DevuelveDesdeBD("codforpa", "formapago", cad, "1 ORDER BY codforpa")
    End If
    If cad = "" Then
        'Cualqueira
        cad = "tipforpa=3 AND nomforpa like '%trans%' AND 1"
        cad = DevuelveDesdeBD("codforpa", "formapago", cad, "1 ORDER BY codforpa")
    End If
    
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & cad & "|"
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
    
    
    cad = RecuperaValor(CadenaDesdeOtroForm, 3)
    Tipointegracion = 1
    If cad = "1" Then Tipointegracion = 0 'todo
    
    
    
    'codusu,Codigo,serie,factura,fecha,cta_cli,fpago,tipo_operacion,ctaret,impret,tipo_ret,ctaventas,Ccoste,impventa,
    'IVA , ImpIva, recargo, imprecargo, TotalFactura, integracion
    
    cad = "select tmpintefrafracli.*,codpais,nifdatos,desprovi,despobla,codposta,dirdatos,nommacta,cuentas.iban ibancta from tmpintefrafracli left join cuentas ON cta_cli=codmacta"
    cad = cad & " WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    FACTURA = ""
    Set Mc = New Contadores
    
    
    
    
    
    While Not miRsAux.EOF
            cad = Left(miRsAux!Serie & "   ", 3) & Format(miRsAux!FACTURA, "000000")
            
            If Mc.ConseguirContador(1, Actual, False) = 1 Then Exit Sub
                  
            
            'Haremos los inserts
            'INSERT INTO factpro(codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta
            'codconce340,codopera,codmacta,anofactu,
            'codforpa,observa,totbases,totbasesret,totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten
            ',observa NUmSerie , Numregis, fecharec, NumFactu, FecFactu,fecliqpr
            
            
            NumRegElim = Mc.Contador
            BaseImponible = miRsAux!CalculoImponible
            TotalFac = miRsAux!TotalFactura
            IVA = DBLet(miRsAux!ImpIva, "N")
            Msg$ = "(" & DBSet(miRsAux!codpais, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T", "S") & "," & DBSet(miRsAux!desProvi, "T", "S") & ","
            Msg$ = Msg$ & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S") & "," & DBSet(miRsAux!dirdatos, "T", "S") & ","
            Msg$ = Msg$ & DBSet(miRsAux!Nommacta, "T") & ",'X',0," & DBSet(miRsAux!cta_cli, "T") & "," & Year(fRecep)
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
        FACTURA = cad
        Set ColBases = Nothing
        Set ColBases = New Collection
                        
        'Lineas
        'INSERT INTO factpro_lineas (aplicret,imporec,anofactu,codccost,
        'impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea,fecfactu,numserie,numfactu
        i = i + 1
        cad = "(1," & DBSet(miRsAux!recargo, "T") & "," & Year(fRecep) & ",NULL," 'codccost?
        cad = cad & DBSet(miRsAux!ImpIva, "N") & "," & DBSet(miRsAux!recargo, "N", "S") & "," & DBSet(miRsAux!IVA, "N")
        cad = cad & "," & DBSet(miRsAux!impventa, "N", "N") & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux!ctaventas, "T") & "," & i
        cad = cad & "," & DBSet(fRecep, "F", "N") & "," & DBSet(miRsAux!Serie, "T") & "," & NumRegElim & ")"
        
        ColBases.Add cad
            
        If Not InsertarFACTURA(ColBases, FACTURA, fRecep, Tipointegracion, ColCob) Then
            Mc.DevolverContador Mc.TipoContador, Actual, Mc.Contador, False
        End If
        miRsAux.MoveNext
     Wend
     Msg$ = ""
     Ampliacion = ""
    
    
    
    
    
End Sub








Private Sub InsertarFACTURA_proveeedor(ByRef C As Collection, FacturaC As String, Fecha As Date, Tipointegracion As Byte, ByRef Cobros As Collection)
 Dim RT As ADODB.Recordset
    On Error GoTo eInsertarFACTURA
    
    
    
    Conn.BeginTrans

    If Tipointegracion <= 1 Then

        'Ejecutamos los sql
        'INSERT INTO factpro(numserie,numregis,fecharec,numfactu,fecfactu,codconce340,codopera,codmacta,anofactu,
        'codforpa,observa,totbases,totbasesret,totivas,totrecargo,totfacpr,retfacpr,trefacpr,cuereten,tiporeten,
        'fecliqpr,nommacta,dirdatos,codpobla,despobla,desprovi,nifdatos,codpais)
        Conn.Execute cad
        
        'Lineas
        cad = ""
        For i = 1 To C.Count
            cad = cad & ", " & C.Item(i)
        Next i
        
        cad = "fecfactu,numserie,numfactu) VALUES " & Mid(cad, 2)
        
        cad = "INSERT INTO factcli_lineas (aplicret,imporec,anofactu,codccost,impoiva,porcrec,porciva,baseimpo,codigiva,codmacta,numlinea," & cad
        Conn.Execute cad
        
        'Totales
        espera 0.3
    
        
        cad = " select codigiva, porciva, porcrec, sum(baseimpo) baseimpo, sum(coalesce(impoiva,0)) imporiva, sum(coalesce(imporec,0)) imporrec "
        cad = cad & " from factcli_lineas"
        cad = cad & " where numserie = " & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & " and numfactu = " & DBSet(Mid(FacturaC, 4), "N") & " and anofactu = " & Year(Fecha)
        cad = cad & " group by 1,2,3"
        cad = cad & " order by 1,2,3"
        Set RT = New ADODB.Recordset
        RT.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        i = 0
        cad = ""
        While Not RT.EOF
            i = i + 1
            cad = cad & ", (" & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & "," & DBSet(Mid(FacturaC, 4), "N") & "," & DBSet(Fecha, "F") & "," & Year(Fecha)
            cad = cad & "," & i & "," & DBSet(RT!Baseimpo, "N") & "," & RT!codigiva & "," & DBSet(RT!porciva, "N", "S")
            cad = cad & "," & DBSet(RT!porcrec, "N") & "," & DBSet(RT!Imporiva, "N") & "," & DBSet(RT!imporrec, "N") & ")"
            RT.MoveNext  'rT!
        Wend
        RT.Close
        Set RT = Nothing
        cad = Mid(cad, 2)
        cad = "insert into factcli_totales (numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)  VALUES " & cad
        Conn.Execute cad
        
    
        'Creamos el apunte
        If Not IntegrarFactura(Trim(Mid(FacturaC, 1, 3)), CLng(Mid(FacturaC, 4)), Fecha) Then
            cad = "Creando asiento. "
            AnyadeEnErrores FacturaC
        End If
    
    
    
    End If
    
    
    'Tesoreria
    If Tipointegracion <> 1 Then
        'insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,
        'numserie,numfactu,fecfactu,numorden,codmacta,codforpa,fecvenci,impvenci,ctabanc1,iban,text33csb)
        cad = ""
        For i = 1 To Cobros.Count
            cad = cad & ", " & Cobros.Item(i)
        Next i
        cad = Mid(cad, 2)
        cad = "insert into cobros(agente,nomclien, domclien, pobclien, cpclien, proclien, codpais, nifclien,numserie,numfactu,fecfactu,codmacta,codforpa,ctabanc1,iban,text33csb,text41csb,numorden,fecvenci,impvenci) VALUES " & cad
        Conn.Execute cad
    End If
    
    'Adelante transaccion
    Conn.CommitTrans
    
    'Borramos de la tabla temporal
    
    
    Exit Sub
eInsertarFACTURA:
    cad = Err.Description
    Err.Clear
    Conn.RollbackTrans
    
   
    AnyadeEnErrores FacturaC
    
    
    
End Sub

