VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportarFraCli 
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   6045
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
         Left            =   8160
         TabIndex        =   9
         Top             =   360
         Width           =   1785
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Facturas proveedor"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Top             =   360
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
         Top             =   360
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
         Left            =   10800
         TabIndex        =   6
         Top             =   480
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
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   8985
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
         Top             =   840
         Width           =   690
      End
      Begin VB.Image imgppal 
         Height          =   240
         Index           =   0
         Left            =   1080
         Top             =   840
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
Attribute VB_Name = "frmImportarFraCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const NumeroCamposTratar = 17

Dim strArray() As String
Dim LineaConErrores As Boolean
Dim cad As String
Dim Importe As Currency

Dim NF As Integer




Private Sub cmdAceptar_Click()
    If ListView1.ListItems.Count > 0 Then
        If MsgBox("¿Continuar con el proceso de importacion?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    If Me.optVarios(0).Value Then
        
        InsertarEnContabilidadFraCli
        
    End If
End Sub

Private Sub cmdCancelar_Click()
    If cmdAceptar.Visible Then
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
        If cmdAceptar.Visible Then
            'Importacion anterior con datos correctos.
            'Preguntamos
            cad = "Hay datos correctos pendientes de integrar. " & vbCrLf & "Cancelar proceso  anterior?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        End If
    End If
    cmdAceptar.Visible = False
    ListView1.ListItems.Clear
    
    


    Screen.MousePointer = vbHourglass
    If Me.optVarios(0).Value Then
        ImportarFacturasCliente
        
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
    cmdAceptar.Visible = False
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
    If Not ComprobacionDatosBD Then
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
   
    cad = "select tmpfaclin.*,nommacta from tmpfaclin left join cuentas on cta=cuentas.codmacta where codusu= " & vUsu.Codigo
    cad = cad & " ORDER BY codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not miRsAux.EOF
        i = i + 1
        ListView1.ListItems.Add , "K" & i
        If DBLet(miRsAux!NUmSerie, "T") <> "" Then
            
            ListView1.ListItems(i).Text = miRsAux!NUmSerie
            ListView1.ListItems(i).SubItems(1) = Format(miRsAux!NumFac, "000000")
            ListView1.ListItems(i).SubItems(2) = miRsAux!Fecha
            
            ListView1.ListItems(i).SubItems(3) = miRsAux!Nommacta
            
            
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!imponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            ListView1.ListItems(i).SubItems(7) = " " 'separado
            
            ListView1.ListItems(i).SubItems(8) = Format(miRsAux!imponibleant, FormatoImporte)
            ListView1.ListItems(i).SubItems(9) = Format(miRsAux!Total - miRsAux!imponibleant, FormatoImporte)
            ListView1.ListItems(i).SubItems(10) = Format(miRsAux!Total, FormatoImporte)
            
            
            
            
            
        Else
            
            ListView1.ListItems(i).Text = " "
            ListView1.ListItems(i).SubItems(3) = "              IVA " & miRsAux!IVA & "%"
            ListView1.ListItems(i).SubItems(4) = Format(miRsAux!imponible, FormatoImporte)
            ListView1.ListItems(i).SubItems(5) = Format(miRsAux!IVA, FormatoImporte)
            ListView1.ListItems(i).SubItems(6) = Format(miRsAux!ImpIva, FormatoImporte)
            For NF = 1 To 10
                If NF <> 3 And NF <> 4 And NF <> 5 Then ListView1.ListItems(i).SubItems(NF) = " "
            Next
            
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    cmdAceptar.Visible = True
    
    
End Sub


Private Sub CargaEncabezado(LaOpcion As Byte)
Dim clmX As ColumnHeader

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
        clmX.Text = "valore fichero"
        clmX.Width = 7000
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
                
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Serie"
        clmX.Width = 750 '1500
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Factura"
        clmX.Width = 1100
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Fecha"
        clmX.Width = 1300
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Nombre"
        clmX.Width = 4200
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Base"
        clmX.Width = 1100
        clmX.Alignment = lvwColumnRight
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "%Iva"
        clmX.Width = 900
        clmX.Alignment = lvwColumnRight
        
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Iva"
        clmX.Width = 1100
        clmX.Alignment = lvwColumnRight
        
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = " "  'Separador
        clmX.Width = 400
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Total bases"
        clmX.Width = 1500
        clmX.Alignment = lvwColumnRight
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Total iva"
        clmX.Width = 1500
        clmX.Alignment = lvwColumnRight
        
        Set clmX = ListView1.ColumnHeaders.Add()
        clmX.Text = "Total"
        clmX.Width = 1500
        clmX.Alignment = lvwColumnRight
        
    End Select
End Sub

Private Function ImportarFichFracli() As Byte
Dim Aux As String

    On Error GoTo eImportarFicheroFracli
    ImportarFichFracli = 2  'Vacio
    
    NF = -1
    Msg$ = "Serie|Factura|Fecha|Cliente|Forma pago|Tipo opereacion|Cta retencion|Importe retencion|"
    Msg$ = Msg$ & "Tipo retencion|Cta ventas|Importe venta|Porcentaje IVA|Importe IVA|Porcentaje recargo|Impor Recargo|Total factura|INTREGRACION|"
    
    'Preparamos tabla de insercion
    Conn.Execute "DELETE FROM tmpfaclin WHERE codusu = " & vUsu.Codigo
    'errores
    Conn.Execute "DELETE FROM tmptesoreriacomun WHERE codusu = " & vUsu.Codigo
    
    NF = FreeFile
    Open Text1.Text For Input As #NF
    
    
    'tmpfaclin(codusu,codigo,numserie,nomserie,Numfac,Fecha,cta,Cliente,NIF,Imponible,IVA,ImpIVA,Total,retencion,tipoiva,porcrec,recargo,tipoopera,tipoformapago,ctabase,ImponibleAnt,numfactura)
    NumRegElim = 0
    J = 0 'Numero de error en el total de fichero. Secuencial

    While Not EOF(NF)
        NumRegElim = NumRegElim + 1
        Line Input #NF, cad
        
        If NumRegElim = 1 Then
            'Primera linea encabezado?
            If Me.check1.Value = 1 Then cad = ""
        Else
            If InStr(1, String(NumeroCamposTratar, ";"), cad) > 0 Then cad = "" 'todo puntos y comas
        End If
        
        
        If cad <> "" Then
            'Procesamos linea
            
            strArray = Split(cad, ";")
            
            If UBound(strArray) = NumeroCamposTratar - 1 Then
                'Falta el ultimo punto y coma
                cad = cad & ";"
                strArray = Split(cad, ";")
            End If
            
            
            If UBound(strArray) <> NumeroCamposTratar Then
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
            cad = DevuelveDesdeBD("count(*)", "tmpfaclin", "codusu", vUsu.Codigo)
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
    
    CadenaDesdeOtroForm = "INSERT INTO tmpfaclin(codusu,codigo,numserie,Numfac,Fecha,cta,tipoformapago,tipoopera,"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "numfactura,retencion,nif,ctabase,Imponible,IVA,ImpIVA,porcrec"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ",recargo,Total,Cliente) VALUES (" & vUsu.Codigo & "," & NumRegElim
    

    'Vemos los campos.
    ' 0       1         2       3           4           5               6           7       8       9
    'SERIE  FACTURA  FECHA  CTA. CLI.    F.PAGO    TIPO OPERACION   CTA.RET.   IMP.RET   TIPO RET. CTA.VENTAS
    '  10       11          12         13       14          15              16
    'IMP.VENTA  % I.V.A.  MP.IVA    % R.E.    IMP. REC.   TOTAL FACTURA   INTEGRACION
    LineaConErrores = False
    
    For k = 0 To NumeroCamposTratar - 1
        ValidarLinea CByte(k)
    Next k
    
    
    If Not LineaConErrores Then
        
        'Si pone algun dato de la factura, debe ponerlos todos
        If strArray(0) <> "" Or strArray(1) <> "0" Or strArray(2) <> "1900-01-01" Or strArray(3) <> "" Then
            If strArray(0) = "" Or strArray(1) = "0" Or strArray(2) = "1900-01-01" Or strArray(3) = "" Then
                AnyadeEnErrores "Campos facturas. Todos o ninguno"
            End If
        End If
        
       ' 0 = TODO    1 = CONTA  2= TESORERIA
        If Val(strArray(16)) <> 1 Then
            'Queremos que meta el cobro
            If strArray(4) = "" Then
                AnyadeEnErrores "Falta forma de pago"
            End If
        End If
        
        
        'Si indica retencion tiene que indicar el tipo y bicepsversa
        
        If Val(strArray(8)) > 0 Then
            If CCur(strArray(7)) = 0 Then
                
                AnyadeEnErrores "Error importe retencion"
            End If
        Else
            If strArray(7) <> "" Then
                If CCur(strArray(7)) <> 0 Then
                    AnyadeEnErrores "Error tipo retencion"
                End If
            End If
        End If
        
    End If
    
    
    If Not LineaConErrores Then
        'INSERTAMOS EN TMpfaclin
        Conn.Execute CadenaDesdeOtroForm & ")"
        
    
    End If
    
End Sub


Private Sub ValidarLinea(QueCampo As Byte)
Dim ValorSQL As String

    'Vemos los campos.
    ' 0       1         2       3           4           5               6           7       8       9
    'SERIE  FACTURA  FECHA  CTA. CLI.    F.PAGO    TIPO OPERACION   CTA.RET.   IMP.RET   TIPO RET. CTA.VENTAS
    '  10       11          12         13       14          15              16
    'IMP.VENTA  % I.V.A.  MP.IVA    % R.E.    IMP. REC.   TOTAL FACTURA   INTEGRACION
    ValorSQL = "NULL"
    Select Case QueCampo
    'Numerico REQUERIDO
    Case 9, 10, 11, 12
      
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
    Case 1, 4, 7, 13, 14, 15
        
        If QueCampo = 1 Then
            If strArray(QueCampo) = "" Then strArray(QueCampo) = 0
            
        End If
        
        If QueCampo = 1 Or QueCampo = 13 Then
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
    Case 8
        
        
    
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



Private Function ComprobacionDatosBD() As Boolean
Dim FechaMinima As Date

On Error GoTo eComprobacionDatosBD
    ComprobacionDatosBD = False
    Set miRsAux = New ADODB.Recordset
    'Comprobaciones basicas. Que todas las fechas estan en el periodo de contabilidad
    cad = "select min(fecha) minima,max(fecha) from tmpfaclin where numfac >0 and codusu=" & vUsu.Codigo
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
            If miRsAux!minima <= UltimoDiaPeriodoLiquidado Then AnyadeEnErrores "Menor que fecha activa"
        End If
    End If
    miRsAux.Close
    
    
    'Comprobaremos que todas las SERIES estan en contadores
    cad = "select distinct(numserie) from tmpfaclin where numserie<>''  and codusu=" & vUsu.Codigo
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
    cad = "select distinct cta from tmpfaclin where cta<>''  and codusu=" & vUsu.Codigo & " and not cta in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cta cliente"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
        
    cad = "select distinct ctabase from tmpfaclin where ctabase<>''  and codusu=" & vUsu.Codigo & " and not ctabase in (select codmacta from cuentas where apudirec='S')"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux.Fields(0)
        AnyadeEnErrores "No existe cta base"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
     
    
    'Facturas que ya existen
    cad = "select * from tmpfaclin where codigo>0  and codusu=" & vUsu.Codigo
    cad = cad & " and (numserie,numfac,year(fecha)) iN (select numserie,numfactu,anofactu from factcli where anofactu>=" & Year(FechaMinima) & " )"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = miRsAux!NumFac
        AnyadeEnErrores "Ya existe factura"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    


    'Porcentaje IVA
    'Facturas que ya existen
    cad = "select iva,coalesce(porcrec,0) recar from tmpfaclin where codusu=" & vUsu.Codigo & " GROUP BY 1,2     "
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = "porcerec=" & DBSet(miRsAux!recar, "N") & " AND porceiva"
        cad = DevuelveDesdeBD("codigiva", "tiposiva", cad, DBSet(miRsAux!IVA, "N"))
        If cad = "" Then
            cad = "IVa: " & miRsAux!IVA & "% Rec: " & miRsAux!recar & "%"
            AnyadeEnErrores "No existe IVA"
            
        Else
        
                cad = "UPDATE tmpfaclin SET tipoiva=" & cad & " WHERE codusu =" & vUsu.Codigo
                cad = cad & " AND iva =" & DBSet(miRsAux!IVA, "N") & " AND porcrec=" & miRsAux!recar
                Conn.Execute cad
        
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Si esta duplicado la factura
    cad = "select numfac,count(*) from tmpfaclin where numfac>0 and codusu =" & vUsu.Codigo & " group by 1 having count(*)>1"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = "Factura: " & miRsAux!NumFac & "   Veces: " & miRsAux.Fields(1)
        AnyadeEnErrores "Datos duplicados"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
                

    'Si no han habiado errores..  j=0
    ComprobacionDatosBD = J = 0


    If J = 0 Then
        'Pondre en nomserie, la cuenta de ventas
        cad = "select distinct(ctabase) from tmpfaclin where numfac>0 and codusu =" & vUsu.Codigo
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", miRsAux.Fields(0), "T")
            cad = "UPDATE tmpfaclin set nomserie=" & DBSet(cad, "T")
            cad = cad & " WHERE codusu = " & vUsu.Codigo & " AND ctabase=" & DBSet(miRsAux.Fields(0), "T")
            Conn.Execute cad
            miRsAux.MoveNext
        Wend
        miRsAux.Close

    End If

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


    On Error GoTo eComprobarTotales
    ComprobarTotales = False


    Set miRsAux = Nothing
    Set miRsAux = New ADODB.Recordset
    
    
    cad = "select * from tmpfaclin WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Dim Fin As Boolean
    Fin = False
    While Not miRsAux.EOF
        If miRsAux!NUmSerie = "" Then
            'Es otra base de la misma factura
            
        Else
            cad = Left(miRsAux!NUmSerie & "   ", 3) & Format(miRsAux!NumFac, "000000")
            
            If cad <> FACTURA Then
                If FACTURA <> "" Then
                    'Suma correcta
                    Ok = (BaseImponible + IVA - ImporteRetencion) = TotalFac
                   
                    cad = "Calculado/fichero:  " & BaseImponible + IVA - ImporteRetencion & " / " & TotalFac
                    If ImporteRetencion <> 0 Then cad = cad & " Retencion Tipo" & TipoRetencion & "  Importe: " & ImporteRetencion
                    
                    If Not Ok Then
                        
                        AnyadeEnErrores "Total factura"
                        
                    Else
                        cad = " AND numserie= " & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & " AND numfac =" & Mid(FACTURA, 4)
                        cad = "UPDATE tmpfaclin set imponibleant=" & DBSet(BaseImponible, "N") & "  WHERE codusu = " & vUsu.Codigo & cad
                        Conn.Execute cad
                    End If
                    cad = Left(miRsAux!NUmSerie & "   ", 3) & Format(miRsAux!NumFac, "000000")
                End If
                'Factura NUEVA. Reseteamos importes...
                TipoRetencion = DBLet(miRsAux!NIF, "N")
                ImporteRetencion = 0
                If TipoRetencion <> 0 Then ImporteRetencion = DBLet(miRsAux!Retencion, "N")
                NumRegElim = miRsAux!NumFac
                BaseImponible = 0
                IVA = 0
                TotalFac = miRsAux!Total
                
                FACTURA = cad
                
            End If
        End If
        
        Importe = Round((miRsAux!imponible * miRsAux!IVA) / 100, 2)
        If Importe <> miRsAux!ImpIva Then
            cad = "IVA calculado/fichero " & miRsAux!IVA & "%: " & Importe & " / " & miRsAux!ImpIva
            AnyadeEnErrores "Calculo  IVA"

        End If
        
        BaseImponible = BaseImponible + miRsAux!imponible
        IVA = IVA + Importe
        
            
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If FACTURA <> "" Then
        'Suma correcta
        If BaseImponible + IVA <> TotalFac Then
            cad = "Calculado/fichero:  " & BaseImponible + IVA & " / " & TotalFac
            AnyadeEnErrores "Total factura"
        Else
            cad = " AND numserie= " & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & " AND numfac =" & Mid(FACTURA, 4)
            cad = "UPDATE tmpfaclin set imponibleant=" & DBSet(BaseImponible, "N") & "  WHERE codusu = " & vUsu.Codigo & cad
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


    Set miRsAux = New ADODB.Recordset
    
    
    'Valores por defecto
    cad = DevuelveDesdeBD("codigo", "agentes", "1", "1 ORDER BY codigo")
    CadenaDesdeOtroForm = cad & "|"
    cad = DevuelveDesdeBD("codforpa", "formapago", "1", "1 ORDER BY codforpa")
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & cad & "|"
                        
    'Cuenta retencion
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "473000001|"

                        
                        
    'CadenaDesdeOtroForm = agente|codforpa|"
    
    cad = "select tmpfaclin.*,codpais,nifdatos,desprovi,despobla,codposta,dirdatos,nommacta from tmpfaclin left join cuentas ON cta=codmacta"
    cad = cad & " WHERe codusu =" & vUsu.Codigo & " order by codigo"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    FACTURA = ""
    While Not miRsAux.EOF
        If miRsAux!NUmSerie = "" Then
            'Es otra base de la misma factura
            
        Else
            cad = Left(miRsAux!NUmSerie & "   ", 3) & Format(miRsAux!NumFac, "000000")
            
            If cad <> FACTURA Then
                If FACTURA <> "" Then
                   
                    InsertarFACTURA_ ColBases, FACTURA, Fecha
                    
                    'Suma correcta
                    cad = Left(miRsAux!NUmSerie & "   ", 3) & Format(miRsAux!NumFac, "000000")
                End If
                'Factura NUEVA.
                
                'Haremos los inserts
                'INSERT INTO factcli (codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,totfaccl,trefaccl,totbasesret,totivas,
                'totbases,codagente,dpto,numasien,fecliqcl,retfaccl,cuereten,tiporeten,codopera,observa,numserie,numfactu,fecfactu,codmacta,
                'codforpa,anofactu)
                NumRegElim = miRsAux!NumFac
                BaseImponible = miRsAux!imponibleant
                TotalFac = miRsAux!Total
                IVA = TotalFac - BaseImponible + DBLet(miRsAux!Retencion, "N")
                Msg$ = "(" & DBSet(miRsAux!codPAIS, "T", "S") & "," & DBSet(miRsAux!nifdatos, "T", "S") & "," & DBSet(miRsAux!desProvi, "T", "S") & ","
                Msg$ = Msg$ & DBSet(miRsAux!desPobla, "T", "S") & "," & DBSet(miRsAux!codposta, "T", "S") & "," & DBSet(miRsAux!dirdatos, "T", "S") & ","
                'nommacta totfaccl,trefaccl,totbasesret,totivas,totbases,codagente,dpto,numasien,
                Msg$ = Msg$ & DBSet(miRsAux!Nommacta, "T") & "," & DBSet(miRsAux!Total, "N") & "," & DBSet(miRsAux!Retencion, "N") & "," & DBSet(BaseImponible, "N")
                Msg$ = Msg$ & "," & DBSet(IVA, "N") & "," & DBSet(BaseImponible, "N") & ","
                Msg$ = Msg$ & RecuperaValor(CadenaDesdeOtroForm, 1) & ",NULL,NULL," & DBSet(miRsAux!Fecha, "F") & ","
                If DBLet(DBLet(miRsAux!NumFactura, "N"), "N") = 0 Then
                    'No lleva retencion
                    Msg$ = Msg$ & "NULL,NULL,0"
                Else
                    Msg$ = Msg$ & ",'" & RecuperaValor(CadenaDesdeOtroForm, "T") & "'," & DBSet(miRsAux!Retencion, "N", "S") & "," & DBLet(miRsAux!NumFactura, "N") 'tiporeten=numfactura
                End If
                Msg$ = Msg$ & "," & DBLet(miRsAux!tipoopera, "N") & ","
                Msg$ = Msg$ & "'Importacion datos externos. " & Chr(13) & "Usuario: " & DevNombreSQL(vUsu.Nombre) & Chr(13) & "Fecha: " & Now
                Msg$ = Msg$ & "'," & DBSet(miRsAux!NUmSerie, "T") & "," & miRsAux!NumFac & "," & DBSet(miRsAux!Fecha, "F") & "," & DBSet(miRsAux!Cta, "T") & ","
                If DBLet(miRsAux!tipoformapago, "T") = "" Then
                    'No viene forma de pago
                     Msg$ = Msg$ & RecuperaValor(CadenaDesdeOtroForm, 2)
                Else
                     Msg$ = Msg$ & miRsAux!tipoformapago
                End If
                Msg$ = Msg$ & "," & Year(miRsAux!Fecha) & ")"
                
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
        cad = cad & DBSet(miRsAux!ImpIva, "N") & "," & DBSet(miRsAux!porcrec, "N", "S") & "," & DBSet(miRsAux!IVA, "N")
        cad = cad & "," & DBSet(miRsAux!imponible, "N", "N") & "," & DBSet(miRsAux!TipoIva, "N") & "," & DBSet(miRsAux!CtaBase, "T") & "," & i
        cad = cad & "," & DBSet(Fecha, "F", "N") & "," & DBSet(Trim(Mid(FACTURA, 1, 3)), "T") & "," & DBSet(Mid(FACTURA, 4), "N") & ")"
        
        ColBases.Add cad
            
        'Siguiente
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If FACTURA <> "" Then
        InsertarFACTURA_ ColBases, FACTURA, Fecha
    End If
    

    
End Sub




Private Sub InsertarFACTURA_(ByRef C As Collection, FacturaC As String, Fecha As Date)
 Dim rT As ADODB.Recordset
    On Error GoTo eInsertarFACTURA
    
    
    
     Conn.BeginTrans

    'Ejecutamos los sql
    cad = "INSERT INTO factcli (codpais,nifdatos,desprovi,despobla,codpobla,dirdatos,nommacta,totfaccl,trefaccl,totbasesret,totivas,"
    cad = cad & "totbases,codagente,dpto,numasien,fecliqcl,retfaccl,cuereten,tiporeten,codopera,observa,numserie,numfactu,fecfactu,codmacta,"
    cad = cad & "codforpa,anofactu) VALUES " & Msg$
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
    Set rT = New ADODB.Recordset
    rT.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    i = 0
    cad = ""
    While Not rT.EOF
        i = i + 1
        cad = cad & ", (" & DBSet(Trim(Mid(FacturaC, 1, 3)), "T") & "," & DBSet(Mid(FacturaC, 4), "N") & "," & DBSet(Fecha, "F") & "," & Year(Fecha)
        cad = cad & "," & i & "," & DBSet(rT!Baseimpo, "N") & "," & rT!Codigiva & "," & DBSet(rT!porciva, "N", "S")
        cad = cad & "," & DBSet(rT!porcrec, "N") & "," & DBSet(rT!imporiva, "N") & "," & DBSet(rT!imporrec, "N") & ")"
        rT.MoveNext  'rT!
    Wend
    rT.Close
    Set rT = Nothing
    cad = Mid(cad, 2)
    cad = "insert into factcli_totales (numserie,numfactu,fecfactu,anofactu,numlinea,baseimpo,codigiva,porciva,porcrec,impoiva,imporec)  VALUES " & cad
    Conn.Execute cad
    
    
    
    
    'Adelante transaccion
    Conn.CommitTrans
    
    
    
    
    Exit Sub
eInsertarFACTURA:
    MuestraError Err.Number, Err.Description
    Conn.RollbackTrans
End Sub


'Private Function IntegrarFactura() As Boolean
'Dim SqlLog As String
'
'    IntegrarFactura = False
'
'    SqlLog = "Factura : " & Text1(2).Text & " " & Text1(0).Text & " de fecha " & Text1(1).Text
'    SqlLog = SqlLog & vbCrLf & "Línea   : " & DBLet(Me.AdoAux(1).Recordset!NumLinea, "N")
'    SqlLog = SqlLog & vbCrLf & "Cuenta  : " & DBLet(Me.AdoAux(1).Recordset!codmacta, "T") & " " & DBLet(Me.AdoAux(1).Recordset!Nommacta, "T")
'    SqlLog = SqlLog & vbCrLf & "Importe : " & DBLet(Me.AdoAux(1).Recordset!Baseimpo, "N")
'
'
'
'    With frmActualizar
'        .OpcionActualizar = 6
'        'NumAsiento     --> CODIGO FACTURA
'        'NumDiari       --> AÑO FACTURA
'        'NUmSerie       --> SERIE DE LA FACTURA
'        'FechaAsiento   --> Fecha factura
'        'FechaAnterior  --> Fecha Anterior de la Factura (ahora no se borra la cabecera del asiento)
'        .NumFac = CLng(Text1(0).Text)
'        .NumDiari = CInt(Text1(14).Text)
'        .NUmSerie = Text1(2).Text
'        .FechaAsiento = Text1(1).Text
'        .FechaAnterior = FecFactuAnt
'        .SqlLog = SqlLog
'        If Numasien2 < 0 Then
'
'            If Not Text1(8).Enabled Then
'                If Text1(8).Text <> "" Then
'                    Numasien2 = Text1(8).Text
'                End If
'            End If
'
'        End If
'        If NumDiario <= 0 Then NumDiario = vParam.numdiacl
'        .DiarioFacturas = NumDiario
'        .NumAsiento = Numasien2
'        .Show vbModal
'
'        If AlgunAsientoActualizado Then IntegrarFactura = True
'
'        Screen.MousePointer = vbHourglass
'        Me.Refresh
'    End With
'
'End Function





Private Sub imgppal_Click(Index As Integer)
    With frmppal.cd1
        .CancelError = False
        .Filter = "*.csv|*.csv"
        .FilterIndex = 1
        .DialogTitle = "Importar facturas proveedor"
        .ShowOpen
        If .FileName <> "" Then Text1.Text = .FileName
    End With
End Sub


