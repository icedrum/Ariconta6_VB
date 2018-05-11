VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmActualizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar diario"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmActualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame1Asiento 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4815
      Begin ComCtl2.Animation Animation1 
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   241
         FullHeight      =   49
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label9 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   210
         TabIndex        =   5
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lblAsiento 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public OpcionActualizar As Byte
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Si el asiento es de una factura entonces NUMSERIE tendra "FRACLI" o "FRAPRO"
    ' con lo cual habra que poner su factura asociada a NULL
    
    '4.- Si es para enviar datos a impresora
    '5.- Actualiza mas de 1 asiento
    
    '6.- Integra 1 factura
    '7.- Elimina factura integrada . DesINTEGRA   . C L I E N T E S
    '8.- Integra 1 factura PROVEEDORES
    '9.- Elimina factura integrada . Desintegra. P R O V E E D O R E S
    
    '10 .- Integracion masiva facturas clientes
    '11 .- Integracion masiva facturas Proveedores
    
    
        
Public NumAsiento As Long
Public FechaAsiento As Date
Public NumDiari As Integer
Public NUmSerie As String
Public NumFac As Long
Public FechaAnterior As Date
Public Proveedor As String
Public FACTURA As String
Public FechaFactura As Date

Public DentroBeginTrans As Boolean

'Nuevo. 17 Cotubre de 2005
'-------------------------
'  Los clientes que facturan con mas de un diario, las facturas SIEMPRE
'  van al diaro de parametros, con lo cual ES una cagada
Public DiarioFacturas As Integer


Public SqlLog As String

Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private Cuenta As String
Private ImporteD As Currency
Private ImporteH As Currency
Private CCost As String
'Y estas son privadas
Private Mes As Integer
Private Anyo As Integer
Dim Fecha As String  'TENDRA la fecha ya formateada en yyy-mm-dd
Dim PrimeraVez As Boolean
Dim SQL As String
Dim Rs As Recordset

Dim INC As Long

Dim NE As Integer
Dim ErroresAbiertos As Boolean
Dim NumErrores As Long

Dim ItmX As ListItem  'Para mostra errores masivos

Private Sub AñadeError(ByRef Mensaje As String)
On Error Resume Next
'Escribimos en el fichero
If Not ErroresAbiertos Then
    NE = FreeFile
    ErroresAbiertos = True
    Open App.Path & "\ErrActua.txt" For Output As NE
    If Err.Number <> 0 Then
        MsgBox " Error abriendo fichero errores", vbExclamation
        Err.Clear
    End If
End If
Print #NE, Mensaje
If Err.Number <> 0 Then
    Err.Clear
    NumErrores = -20000
Else
    NumErrores = NumErrores + 1
End If
End Sub



Private Function CadenaImporte(VaAlDebe As Boolean, ByRef Importe As Currency, ElImporteEsCero As Boolean) As String
Dim CadImporte As String

'Si va al debe, pero el importe es negativo entonces va al haber a no ser que la contabilidad admita importes negativos
    If Importe < 0 Then
        If Not vParam.abononeg Then
            VaAlDebe = Not VaAlDebe
            Importe = Abs(Importe)
        End If
    End If
    ElImporteEsCero = (Importe = 0)
    CadImporte = TransformaComasPuntos(CStr(Importe))
    If VaAlDebe Then
        CadenaImporte = CadImporte & ",NULL"
    Else
        CadenaImporte = "NULL," & CadImporte
    End If
End Function

Private Sub CargaProgres(Valor As Integer)
Me.ProgressBar1.Max = Valor
Me.ProgressBar1.Value = 0
End Sub



Private Sub IncrementaProgres(Veces As Integer)
On Error Resume Next
Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * INC)
If Err.Number <> 0 Then
    Err.Clear
    ProgressBar1.Value = 0
End If

End Sub





'Eliminar factura con asiento
Private Function EliminaFacturaConAsiento()
Dim Donde As String
Dim bol As Boolean
Dim LEtra As String
Dim Mc As Contadores
Dim Contabilizada As String

    On Error GoTo EEliminaFacturaConAsiento
    'Sabemos que
    'numasiento     --> Nº aseinto
    'numfac         --> CODIGO FACTURA
    'NumDiari       --> ATENCION -> Nº de diario, no como al integrar
    'FechaAsiento   --> Fecha asiento
    'NUmSerie       --> SERIE DE LA FACTURA  y el año (sep. con pipes)

    'Obtenemos el mes y el año
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Aqui bloquearemos
    Conn.BeginTrans
    
    'Eliminamos factura
    LEtra = RecuperaValor(NUmSerie, 1)
    If Me.OpcionActualizar = 7 Then
        '-------------------------------------------------------------
        '               C L I E N T E S
        '-------------------------------------------------------------
        SQL = " WHERE numserie = '" & LEtra & "'"
        SQL = SQL & " AND numfactu = " & NumFac
        SQL = SQL & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factcli_lineas " & SQL
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factcli_totales " & SQL
        Conn.Execute Cuenta
        
        
        Contabilizada = "select count(*) from cobros where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F") & " and impcobro <> 0 and not impcobro is null "
        
        If TotalRegistros(Contabilizada) <> 0 Then
            MsgBox "Hay cobros que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
        Else
            ' cobro de la factura
            Donde = "Cobro factura"
            
            Cuenta = "DELETE from cobros where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F")
            Conn.Execute Cuenta
        End If
        
        'La factura
        Donde = "Cabecera factura"
        Cuenta = "DELETE from factcli " & SQL
        Conn.Execute Cuenta

    Else
        '-------------------------------------------------------------
        '       P R O V E E D O R E S
        '-------------------------------------------------------------
        SQL = " WHERE numserie = '" & LEtra & "'"
        SQL = SQL & " AND numregis = " & NumFac
        SQL = SQL & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factpro_lineas " & SQL
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factpro_totales " & SQL
        Conn.Execute Cuenta
        
        Contabilizada = "select count(*) from pagos where numserie = " & DBSet(LEtra, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfactu = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(FechaFactura, "F") & " and imppagad <> 0 and not imppagad is null "
        
        If TotalRegistros(Contabilizada) <> 0 Then
            MsgBox "Hay pagos que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
        Else
            ' cobro de la factura
            Donde = "Pago factura"
            
            Cuenta = "DELETE from pagos where numserie = " & DBSet(LEtra, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfactu = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(FechaFactura, "F")
            Conn.Execute Cuenta
        End If
        
        'La factura
        Donde = "Cabecera factura"
        Cuenta = "DELETE from factpro " & SQL
        Conn.Execute Cuenta
        LEtra = RecuperaValor(NUmSerie, 1) '"1"
    End If

    bol = DesActualizaElASiento(Donde)

EEliminaFacturaConAsiento:
        If Err.Number <> 0 Then
            SQL = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            MuestraError Err.Number, SQL, Err.Description
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            
            'Intentamos devolver el contador
            If FechaAsiento >= vParam.fechaini Then
                Set Mc = New Contadores
                Mc.DevolverContador LEtra, (FechaAsiento <= vParam.fechafin), NumFac
                Set Mc = Nothing
            End If
            
            
            'INSERTO EN LOG
            Mes = 6
            
            
            If Me.OpcionActualizar <> 7 Then
                Mes = 10   'FRAPRO
                LEtra = ""
                
                vLog.Insertar 10, vUsu, SqlLog
            Else
                vLog.Insertar 6, vUsu, SqlLog
            End If
            
            
            EliminaFacturaConAsiento = True
            AlgunAsientoActualizado = True
        Else
            Conn.RollbackTrans
        End If
    
End Function



    






Private Sub Form_Activate()
Dim bol As Boolean
If PrimeraVez Then
    PrimeraVez = False
    Me.Refresh
    bol = False
    Select Case OpcionActualizar
    Case 2, 3
        DesActualizaAsiento
        bol = True
    Case 6, 8
        'Integramos la factura (Dependera del opcion si es de clientes o de proveedores
        IntegraFactura
        bol = True
    Case 7, 9
         'Integramos la factura (Dependera del opcion si es de clientes o de proveedores
        EliminaFacturaConAsiento
        bol = True
        
        
    Case 16
        'Insertar Asiento en el hco
        
    End Select
    If bol Then Unload Me
End If
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim B As Boolean

    Me.Icon = frmppal.Icon


    ErroresAbiertos = False
    Limpiar Me
    PrimeraVez = True
    
    Select Case OpcionActualizar
    Case 1, 2, 3
        Label1.Caption = "Nº Asiento"
        Me.lblAsiento.Caption = NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 1 Then
            Label9.Caption = "Actualizar"
        Else
            Label9.Caption = "Modi/Eliminar"
        End If
        'Tamaño
        Me.Height = 3000
        B = True
            
    Case 6, 7, 8, 9
        '// Estamos en Facturas
        Label1.Caption = "Nº factura"
        If OpcionActualizar < 8 Then
            Label1.Caption = Label1.Caption & " Cliente"
        Else
            Label1.Caption = Label1.Caption & " Proveedor"
        End If
        Me.lblAsiento.Caption = NUmSerie & NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 6 Or OpcionActualizar = 8 Then
            Label9.Caption = "Integrar Factura"
        Else
            Label9.Caption = "Eliminar Factura"
        End If
        Me.Caption = "Actualizar facturas"
        'Tamaño
        Me.Height = 3315
        B = True
        
    End Select
    Me.frame1Asiento.visible = B
    Me.Animation1.visible = B
End Sub



Private Function IntegraFactura() As Boolean
Dim B As Boolean
Dim Donde As String
Dim vConta As Contadores

Dim TipoConce As String
On Error GoTo EIntegraFactura
    
    IntegraFactura = False
    
    If Not DentroBeginTrans Then Conn.BeginTrans
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    
    'Vemos si estamos intentato forzar numero de asiento
    If NumAsiento > 0 Then
        'Primero que nada obtendremos el contador
        If AsientoExiste Then
            MsgBox "Ya existe el asiento con la numeración: " & NumAsiento & " " & FechaAsiento & " " & NumDiari, vbExclamation
            'Vamoa al final del proceso de esta factura
            GoTo EIntegraFactura
        End If
    Else
        Donde = "Conseguir contador"
        Set vConta = New Contadores
        If vConta.ConseguirContador("0", (FechaAsiento <= vParam.fechafin), True) = 1 Then
            MsgBox "Error consiguiendo contador asiento", vbExclamation
            'Vamoa al final del proceso de esta factura
            GoTo EIntegraFactura
        End If
        
        If Not vConta.YaExisteContador((FechaAsiento <= vParam.fechafin), vParam.fechafin, (OpcionActualizar < 10)) Then
            If OpcionActualizar > 9 Then InsertaError "Error contadores asiento: " & vConta.Contador
            GoTo EIntegraFactura
        End If
        NumAsiento = vConta.Contador
        Set vConta = Nothing
    End If
    
    'Actualizamos los datos
    If OpcionActualizar = 6 Or OpcionActualizar = 10 Then
        B = IntegraLaFactura(Donde)
    Else
        B = IntegraLaFacturaProv(Donde)
    End If
    
EIntegraFactura:
    If Err.Number <> 0 Then
        If OpcionActualizar > 9 Then
            'Esta actualizando varias a la vez
            InsertaError Donde & " - " & Err.Description
        Else
            MuestraError Err.Number, "Integra factura(I)" & vbCrLf & Donde
        End If
        Err.Clear
        B = False
    End If
    If B Then
        
        If OpcionActualizar > 9 Then
            'Actualizando desde/hasta y ha ido bien. La meto al LOG
            vLog.AnyadeTextoDatosDes NUmSerie & Format(NumFac, "000000")
            'If OpcionActualizar = 10 Then
            '    'FRACLI
        End If
    End If
    IntegraFactura = B
    AlgunAsientoActualizado = B
    
    If Not DentroBeginTrans Then
        If B Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
        End If
    End If
End Function

Private Function IntegraLaFactura(ByRef A_Donde As String) As Boolean
Dim Cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim Amplia2 As String
Dim DocConcAmp As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean
Dim PrimeraContrapartida As String
    
    Dim SqlIva As String
    Dim RsIvas As ADODB.Recordset

    IntegraLaFactura = False
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'NumDiari       --> AÑO FACTURA
    'NUmSerie       --> SERIE DE LA FACTURA
    'FechaAsiento   --> Fecha factura
    'FecFactuAnt    --> FecFactura Anterior
    
    'Obtenemos los datos de la factura
    A_Donde = "Leyendo datos factura"
    Set RF = New ADODB.Recordset
    SQL = "SELECT * FROM factcli"
    SQL = SQL & " WHERE numserie='" & NUmSerie
    SQL = SQL & "' AND numfactu= " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & SQL, vbExclamation
        RF.Close
        Exit Function
    End If
    
 
    SQL = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(SQL) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        SQL = "UPDATE hcabapu SET "
        SQL = SQL & " fechaent = " & DBSet(Fecha, "F")
        SQL = SQL & " where numdiari = " & DBSet(DiarioFacturas, "N")
        SQL = SQL & " and fechaent = " & DBSet(FechaAnterior, "F")
        SQL = SQL & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute SQL
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        SQL = SQL & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        SQL = SQL & ","
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        Cad = DBLet(RF!observa, "T")
        If Cad = "" Then
            Cad = "NULL,"
        Else
            Cad = "'" & DevNombreSQL(Cad) & "',"
        End If
        Cad = Cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6:Contab. Fra cli " & NUmSerie & Format(NumFac, "0000000") & " " & Fecha & "')"
        
        
        SQL = SQL & Cad
        Conn.Execute SQL
    End If
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    Cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    Cad = Cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    Cad = Cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    
    
    A_Donde = "Linea cliente"
    '-------------------------------------------------------------------
    'LINEA Cliente
    SQL = Mes & ",'" & RF!codmacta & "',"
    
    DocConcAmp = "'" & NUmSerie & Format(NumFac, "0000000") & "'," & vParam.concefcl & ",'"
    
    
    'Ampliacion segun parametros
    Select Case vParam.nctafact
    Case 1
        If RF!totfaccl < 0 Then
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasCli, 2)
        Else
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasCli, 1)
        End If
        Cad2 = Cad2 & " " & NUmSerie & Format(NumFac, "0000000")
    Case 2
        Cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        Cad2 = DBLet(RF!observa)
    End Select
    
    '   Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = Cad2
    If vParam.CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(True, PrimeraContrapartida)
        If Cad3 <> "" Then
            If Len(Amplia2) > 21 Then Amplia2 = Mid(Amplia2, 1, 21)
            Amplia2 = Amplia2 & " [" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    SQL = SQL & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & Cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    'Esta variable sirve para las demas
    ImporteNegativo = (DBLet(RF!totfaccl, "N") < 0)
    
    'Importes, atencion importes negativos
    '  antes --> Cad2 = CadenaImporte(ImporteNegativo, True, RF!totfaccl)
    Cad2 = CadenaImporte(True, DBLet(RF!totfaccl, "N"), Importe0)
    SQL = SQL & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        SQL = SQL & "'" & PrimeraContrapartida & "'"
    Else
        SQL = SQL & "NULL"
    End If
    SQL = SQL & ",'FRACLI',0)"
    
    
    Conn.Execute Cad & SQL
    Mes = Mes + 1 'Es el contador de lineaapunteshco
    
    ' cuentas de iva ahora se sacan de las tablas de totales
    SqlIva = "select * from factcli_totales "
    SqlIva = SqlIva & " WHERE numserie='" & NUmSerie
    SqlIva = SqlIva & "' AND numfactu= " & NumFac
    SqlIva = SqlIva & " AND anofactu=" & NumDiari
    SqlIva = SqlIva & " order by numlinea "
    
    Set RsIvas = New ADODB.Recordset
    RsIvas.Open SqlIva, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RsIvas.EOF
        Cad3 = "cuentarr"
        Cad2 = DevuelveDesdeBD("cuentare", "tiposiva", "codigiva", DBLet(RsIvas!codigiva, "N"), "N", Cad3)
        If Cad2 <> "" Then
        
            SQL = Mes & ",'" & Cad2 & "'," & DocConcAmp
            Cad2 = CadenaImporte(False, DBLet(RsIvas!Impoiva, "N"), Importe0)
            SQL = SQL & "," & Cad2 & ","
            SQL = SQL & "NULL,'" & DBLet(RF!codmacta, "T") & "','FRACLI',0)"
            'dependiendo de si ContabilizarAptIva0 = 1 se contabiliza o no el iva
            If Importe0 Then
                If vParam.ContabApteIva0 Then
                    Conn.Execute Cad & SQL
                    Mes = Mes + 1
                End If
            Else
                Conn.Execute Cad & SQL
                Mes = Mes + 1
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                     SQL = Mes & "," & Cad3 & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, DBLet(RsIvas!ImpoRec, "N"), Importe0)
                    SQL = SQL & "," & Cad2 & ","
                    SQL = SQL & "NULL,'" & DBLet(RF!codmacta, "T") & "','FRACLI',0)"
                    If Not Importe0 Then
                        Conn.Execute Cad & SQL
                        Mes = Mes + 1
                    End If
            End If
        Else
            MsgBox "Error leyendo TIPO de IVA: " & DBLet(RsIvas!codigiva, "N"), vbExclamation
            RF.Close
            Exit Function
        End If
    
        RsIvas.MoveNext
    Wend
    Set RsIvas = Nothing
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    If Not IsNull(RF!cuereten) Then
        SQL = Mes & ",'" & DBLet(RF!cuereten, "T") & "'," & DocConcAmp
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(True, DBLet(RF!trefaccl, "T"), Importe0)
        SQL = SQL & "," & Cad2 & ","
        SQL = SQL & "NULL,NULL,'FRACLI',0)"
       
        Conn.Execute Cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
    End If
    
    
    IncrementaProgres 2
    
    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    Cad2 = RF!codmacta
    ImporteD = DBLet(RF!totfaccl, "N")
    
    
    'Cerramos el RF
    Cuenta = RF!codmacta
    RF.Close
    
    
    
    A_Donde = "Leyendo lineas factura"
'    SQL = "Select factcli_lineas.* , cuentas.codmacta FROM factcli_lineas,Cuentas "
    SQL = "Select cuentas.codmacta, factcli_lineas.codccost, sum(factcli_lineas.baseimpo) baseimpo FROM factcli_lineas,cuentas "
    SQL = SQL & " WHERE numserie='" & NUmSerie
    SQL = SQL & "' AND numfactu= " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    SQL = SQL & " AND factcli_lineas.codmacta = cuentas.codmacta"
    SQL = SQL & " group by 1,2 "
    SQL = SQL & " order by 1,2 "
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    Cad2 = ""
    A_Donde = "Procesando lineas"
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        SQL = Mes & ",'" & DBLet(RF!codmacta, "T") & "'," & DocConcAmp
        Cad2 = CadenaImporte(False, DBLet(RF!Baseimpo, "N"), Importe0)
        SQL = SQL & "," & Cad2 & ","
        If IsNull(RF!codccost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!codccost & "'"
        End If
        
        SQL = SQL & Cad2 & ",'" & DBLet(Cuenta, "T") & "','FRACLI',0)"
    
        Conn.Execute Cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
        
        'Siguiente
        IncrementaProgres 1
        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If Not vParam.CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            SQL = "UPDATE factcli_lineas SET codmacta ='" & PrimeraContrapartida & "'"
            SQL = SQL & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            SQL = SQL & " AND numlinea =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL SQL  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
        
    
    
    
    'Actualimos en factura, el nº de asiento
    SQL = "UPDATE factcli SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    SQL = SQL & " WHERE numserie='" & NUmSerie
    SQL = SQL & "' AND numfactu= " & NumFac
    SQL = SQL & " AND anofactu= " & NumDiari
    Conn.Execute SQL
    
    'Para los saldos ponemos el numero de asiento donde toca
    '
    A_Donde = "Saldos factura"
    NumDiari = vParam.numdiacl
    NumDiari = DiarioFacturas
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    
    
    IntegraLaFactura = True
End Function



'////////////////////////////////////////////////////////////////////
'
'           Facturas proveedores
Private Function IntegraLaFacturaProv(ByRef A_Donde As String) As Boolean
Dim Cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim DocConcAmp As String
Dim Amplia2 As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean 'Para saber si el importe es 0
Dim PrimeraContrapartida As String  'Si hay solo una linea entonces la pondremos como contrapartida de la primera base


'Modificacion de 31 Enero 2005
'-------------------------------------
'-------------------------------------
Dim ColumnaIVA As String
Dim TipoDIva As Byte
    
    Dim SqlIva As String
    Dim RsIvas As ADODB.Recordset

    IntegraLaFacturaProv = False
    
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'NumDiari       --> AÑO FACTURA
    'FechaAsiento   --> Fecha factura
    
    
    'Obtenemos los datos de la factura
    A_Donde = "Leyendo datos factura"
    Set RF = New ADODB.Recordset
    SQL = "SELECT * FROM factpro"
    SQL = SQL & " WHERE numregis = " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    SQL = SQL & " AND numserie=" & DBSet(NUmSerie, "T")
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & SQL, vbExclamation
        RF.Close
        Exit Function
    End If
    
    SQL = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(SQL) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        SQL = "UPDATE hcabapu SET "
        SQL = SQL & " fechaent = " & DBSet(Fecha, "F")
        SQL = SQL & " where numdiari = " & DBSet(DiarioFacturas, "N")
        SQL = SQL & " and fechaent = " & DBSet(FechaAnterior, "F")
        SQL = SQL & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute SQL
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        SQL = SQL & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        Cad = DBLet(RF!observa, "T")
        If Cad = "" Then
            Cad = "NULL,"
        Else
            Cad = "'" & DevNombreSQL(Cad) & "',"
        End If
        
        Cad = Cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Fra.Pro Reg:" & Format(NumFac, "000000") & " " & Fecha & "')"
        
        SQL = SQL & "," & Cad
        
        Conn.Execute SQL
        
    End If
    
    
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    Cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    Cad = Cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    Cad = Cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    PrimeraContrapartida = ""
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RF!totfacpr < 0)
    A_Donde = "Linea proveedor"
    '-------------------------------------------------------------------
    'LINEA Proveedor
    SQL = Mes & ",'" & RF!codmacta & "',"
    
    'Documento "numdocum"
    If vParam.CodiNume = 1 Then
        Cad2 = Format(NumFac, "0000000000")
    Else
        Cad2 = DBLet(RF!NumFactu)
    End If
    

    DocConcAmp = "'" & Cad2 & "'," & vParam.concefpr & ",'"
    
    
    'Ampliacion segun parametros
    Select Case vParam.nctafact
    Case 1
        If RF!totfacpr < 0 Then
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 2)
        Else
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 1)
        End If
        Cad2 = Cad2 & " " & DevNombreSQL(RF!NumFactu)
        
        Cad2 = Cad2 & " (" & Format(RF!FecFactu, "ddmmyy") & ")"
    Case 2
        Cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        Cad2 = DBLet(RF!observa)
    End Select
    
        
    
    'Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = Cad2
    If vParam.CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(False, PrimeraContrapartida)
        If Cad3 <> "" Then
            If Len(Amplia2) > 26 Then Amplia2 = Mid(Amplia2, 1, 26)
            Amplia2 = Amplia2 & "[" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    SQL = SQL & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & Cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    
    'Importes, atencion importes negativos
    Cad2 = CadenaImporte(False, RF!totfacpr, Importe0)
    SQL = SQL & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        SQL = SQL & "'" & PrimeraContrapartida & "'"
    Else
        SQL = SQL & "NULL"
    End If
    SQL = SQL & ",'FRAPRO',0)"
    
    Conn.Execute Cad & SQL
    Mes = Mes + 1 'Es el contador de lineaapunteshco
    
    ' cuentas de iva ahora se sacan de las tablas de totales
    SqlIva = "select * from factpro_totales "
    SqlIva = SqlIva & " WHERE numserie='" & NUmSerie
    SqlIva = SqlIva & "' AND numregis= " & NumFac
    SqlIva = SqlIva & " AND anofactu=" & NumDiari
    SqlIva = SqlIva & " order by numlinea "
    
    
    Dim EsSujetoPasivo As Boolean
    Dim EsImportacion As Boolean
    
    EsImportacion = (DBLet(RF!CodOpera, "N") = 2)
    EsSujetoPasivo = ((DBLet(RF!CodOpera, "N") = 1) Or (DBLet(RF!CodOpera, "N") = 4))
    
    Set RsIvas = New ADODB.Recordset
    RsIvas.Open SqlIva, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RsIvas.EOF
        TipoDIva = DevuelveValor("select tipodiva from tiposiva where codigiva = " & DBSet(RsIvas!codigiva, "N"))
        If TipoDIva = 1 Then
            'Es iva NO deducible
            ColumnaIVA = "cuentasn"
        Else
            ColumnaIVA = "cuentaso"   'La normal
        End If
        
        Cad3 = "cuentasr"
        Cad2 = DevuelveDesdeBD(ColumnaIVA, "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
        If Cad2 <> "" Then
            SQL = Mes & ",'" & Cad2 & "'," & DocConcAmp
            Cad2 = CadenaImporte(True, RsIvas!Impoiva, Importe0)
            SQL = SQL & "," & Cad2 & ","
            SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
            
            If Importe0 Then
                If vParam.ContabApteIva0 Then
                    If Not EsImportacion Then
                        Conn.Execute Cad & SQL
                        Mes = Mes + 1
                    End If
                End If
            Else
                If Not EsImportacion Then
                    Conn.Execute Cad & SQL
                    Mes = Mes + 1
                End If
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                SQL = Mes & "," & Cad3 & "," & DocConcAmp
                'Importes, atencion importes negativos
                Cad2 = CadenaImporte(True, RsIvas!ImpoRec, Importe0)
                SQL = SQL & "," & Cad2 & ","
                SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                If Not Importe0 Then
                    Conn.Execute Cad & SQL
                    Mes = Mes + 1
                End If
            End If
            
            If EsSujetoPasivo Then
                Cad3 = "cuentarr"
                Cad2 = DevuelveDesdeBD("cuentare", "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
                
                Cad3 = Cad2 & "|" & Cad3 & "|"
                
                
                SQL = Mes & ",'" & RecuperaValor(Cad3, 1) & "'," & DocConcAmp
                Cad2 = CadenaImporte(False, RsIvas!Impoiva, Importe0)
                SQL = SQL & "," & Cad2 & ","
                SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                'If Not Importe0 Then
                    Conn.Execute Cad & SQL
                    Mes = Mes + 1
                'End If
               
                If Not IsNull(RsIvas!ImpoRec) Then
                     SQL = Mes & "," & RecuperaValor(Cad3, 2) & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, RsIvas!ImpoRec, Importe0)
                    SQL = SQL & "," & Cad2 & ","
                    SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        Conn.Execute Cad & SQL
                        Mes = Mes + 1
                    End If
                End If
            End If
            
        Else
            MsgBox "Error leyendo TIPO de IVA: " & RsIvas!codigiva, vbExclamation
            RF.Close
            Exit Function
        End If
    
        RsIvas.MoveNext
    Wend
    Set RsIvas = Nothing
    
    '-------------------------------------
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    If Not IsNull(RF!cuereten) Then
        SQL = Mes & ",'" & RF!cuereten & "'," & DocConcAmp
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(False, DBLet(RF!trefacpr, "N"), Importe0)
        SQL = SQL & "," & Cad2 & ","
        SQL = SQL & "NULL,NULL,'FRAPRO',0)"
       
        Conn.Execute Cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
    End If
    
    
    IncrementaProgres 2
    
    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    Cad2 = RF!codmacta
    ImporteD = RF!totfacpr
    
    
    
    'Cerramos el RF
    Cuenta = RF!codmacta
    RF.Close
    
    
    
    A_Donde = "Leyendo lineas factura"
    SQL = "Select factpro_lineas.codmacta, factpro_lineas.codccost, sum(factpro_lineas.baseimpo) baseimpo  FROM factpro_lineas "
    SQL = SQL & " WHERE numregis= " & NumFac
    SQL = SQL & " AND numserie= " & DBSet(NUmSerie, "T")
    SQL = SQL & " AND anofactu=" & NumDiari
    SQL = SQL & " GROUP BY 1,2 "
    SQL = SQL & " ORDER BY 1,2 "
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    A_Donde = "Procesando lineas"
    Cad2 = ""
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        SQL = Mes & ",'" & RF!codmacta & "'," & DocConcAmp
        Cad2 = CadenaImporte(True, RF!Baseimpo, Importe0)
        SQL = SQL & "," & Cad2 & ","
        If IsNull(RF!codccost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!codccost & "'"
        End If
        
        SQL = SQL & Cad2 & ",'" & Cuenta & "','FRAPRO',0)"
    
        Conn.Execute Cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
        
        'Siguiente
        IncrementaProgres 1
        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If Not vParam.CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            SQL = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
            SQL = SQL & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            SQL = SQL & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL SQL  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
    
    'Actualimos en factura, el nº de asiento
    SQL = "UPDATE factpro SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    SQL = SQL & " WHERE  numregis = " & NumFac
    SQL = SQL & " AND numserie= " & DBSet(NUmSerie, "T")
    SQL = SQL & " AND anofactu=" & NumDiari
    Conn.Execute SQL
    
    'Para los saldos ponemos el numero de asiento donde toca
    '
    A_Donde = "Saldos factura"
    NumDiari = DiarioFacturas
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    
    IntegraLaFacturaProv = True
End Function



Private Function AsientoExiste() As Boolean
    AsientoExiste = True
    SQL = "SELECT numdiari from hcabapu"
    SQL = SQL & " WHERE numdiari =" & NumDiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & NumAsiento
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs.EOF Then AsientoExiste = False
    Rs.Close
    Set Rs = Nothing
End Function






'-------------------------------------------------------
'-------------------------------------------------------
'ANALITICA
'-------------------------------------------------------
'-------------------------------------------------------




Private Function BorrarASiento(BorrarCabecera As Boolean) As Boolean

On Error GoTo EBorrarASiento
    BorrarASiento = False
    
    'Borramos las lineas
    SQL = "Delete from hlinapu"
    SQL = SQL & " WHERE numasien = " & NumAsiento
    SQL = SQL & " AND numdiari = " & NumDiari
    SQL = SQL & " AND fechaent=" & DBSet(FechaAnterior, "F")
    Conn.Execute SQL
    
    If BorrarCabecera Then
        'La cabecera
        SQL = "Delete from hcabapu"
        SQL = SQL & " WHERE numdiari =" & NumDiari
        SQL = SQL & " AND fechaent=" & DBSet(FechaAnterior, "F")
        SQL = SQL & " AND numasien=" & NumAsiento
        
        Conn.Execute SQL
    Else
        'Actualizamos la fecha de la cabecera
        SQL = "Update hcabapu"
        SQL = SQL & " set fechaent = " & DBSet(Fecha, "F")
        SQL = SQL & " WHERE numdiari =" & NumDiari
        SQL = SQL & " AND fechaent=" & DBSet(FechaAnterior, "F")
        SQL = SQL & " AND numasien=" & NumAsiento
    
        Conn.Execute SQL
    End If
    
    BorrarASiento = True
    Exit Function
EBorrarASiento:
    Err.Clear
    
End Function

Private Sub ObtenFoco(ByRef T As TextBox)
T.SelStart = 0
T.SelLength = Len(T.Text)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If NumErrores > 0 Then CerrarFichero
End Sub

Private Sub CerrarFichero()
On Error Resume Next
If NE = 0 Then Exit Sub
Close #NE
If Err.Number <> 0 Then Err.Clear
End Sub








Private Function InsertaError(ByRef CADENA As String)
Dim vS As String
    'Insertamos en errores
    'Esta lo tratamos con error especifico
    
    On Error Resume Next

    If OpcionActualizar < 10 Then
        'Insertamos error para ASIENTOS
        vS = NumDiari & "|"
        vS = vS & Fecha & "|"
        vS = vS & NumAsiento & "|"
        vS = vS & CADENA & "|"
    
    Else
        vS = NUmSerie & " " & NumFac & "|"
        vS = vS & FechaAsiento & "|"
        vS = vS & CADENA & "|"
    End If
    AñadeError vS
    
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error." & vbCrLf & Err.Description & vbCrLf & vS
        Err.Clear
    End If
End Function


Private Function DesActualizaAsiento() As Boolean
    Dim bol As Boolean
    Dim Donde As String
    On Error GoTo EDesActualizaAsiento
    
    
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Obtenemos el mes y el año
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Comprobamos que no existe en APUNTES
    'Obviamente solo comprobamos si vamos a insertar
    'en apuntes
    If Me.OpcionActualizar = 3 Then
        If AsientoExiste Then Exit Function
    End If
    'Aqui bloquearemos
    
    Conn.BeginTrans
    
    bol = DesActualizaElASiento(Donde)
    
EDesActualizaAsiento:
        If Err.Number <> 0 Then
            SQL = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            MuestraError Err.Number, SQL, Err.Description
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            espera 0.2
            DesActualizaAsiento = True
            AlgunAsientoActualizado = True
        Else
            Conn.RollbackTrans
        End If
End Function


Private Function DesActualizaElASiento(ByRef A_Donde As String) As Boolean

    '2  .- Desactualiza pero NO insertes en apuntes
    '      Si viene FRACLI o FRAPROV habrá que volver
    '3  .- Desactualizar asiento desde hco
        


    DesActualizaElASiento = False
    
    Select Case Me.OpcionActualizar
    
    Case 2
        If NUmSerie = "FRACLI" Or NUmSerie = "FRAPRO" Then
            A_Donde = "Desvinculando facturas"
            If Not DesvincularFactura(NUmSerie = "FRACLI") Then Exit Function
            IncrementaProgres 1
        End If
    End Select
    
    
    'Borramos cabeceras y lineas del asiento
    A_Donde = "Borrar cabeceras y lineas en historico"
    
    If OpcionActualizar = 2 Then
        If Not BorrarASiento(False) Then Exit Function
    Else
        If Not BorrarASiento(True) Then Exit Function
    End If
    
    IncrementaProgres 2
    DesActualizaElASiento = True
End Function

Private Function DesvincularFactura(Clientes As Boolean) As Boolean
On Error Resume Next
    Set Rs = New ADODB.Recordset
    If Clientes Then
        CCost = "factcli"
    Else
        CCost = "factpro"
    End If
    SQL = "Select * From " & CCost
    SQL = SQL & " WHERE numasien=" & NumAsiento
    SQL = SQL & " AND numdiari = " & NumDiari
    SQL = SQL & " AND fechaent = '" & Fecha & "'"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        
        SQL = "UPDATE " & CCost & " SET numasien=NULL, fechaent=NULL, numdiari=NULL"
        If Clientes Then
            SQL = SQL & " WHERE numfactu = " & Rs!codfaccl
            SQL = SQL & " AND anofaccl =" & Rs!anofaccl
            SQL = SQL & " AND numserie = '" & Rs!NUmSerie & "'"
        Else
            'proveedores
            SQL = SQL & " WHERE numregis = " & Rs!Numregis
            SQL = SQL & " AND anofactu =" & Rs!anofactu
        End If
        Conn.Execute SQL
    End If
    If Err.Number <> 0 Then
        DesvincularFactura = False
        MuestraError Err.Number, "Desvincular factura"
    Else
        DesvincularFactura = True
    End If
End Function



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub BorrarArchivoTemporal()
On Error Resume Next
If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero temporal"
End Sub





Private Function DevuelveCentroCosteFactura(Cliente As Boolean, LaPrimeraContrapartida As String) As String
Dim R As ADODB.Recordset
Dim SQL As String
    DevuelveCentroCosteFactura = ""
    If Cliente Then
        
        SQL = "SELECT codccost,numlinea,baseimpo codtbase FROM factcli_lineas"
        SQL = SQL & " WHERE numserie='" & NUmSerie
        SQL = SQL & "' AND numfactu= " & NumFac
        SQL = SQL & " AND anofactu=" & NumDiari
        SQL = SQL & " AND not (codccost is null)"   'El primero k devuelva
        SQL = SQL & " ORDER BY numlinea"
    Else
        SQL = "SELECT codccost,numlinea,baseimpo codtbase FROM factpro_lineas"
        SQL = SQL & " WHERE numregis = " & NumFac
        SQL = SQL & " AND anofactu=" & NumDiari
        SQL = SQL & " AND numserie=" & DBSet(NUmSerie, "T")
        SQL = SQL & " AND not (codccost is null)"   'El primero k devuelva
        SQL = SQL & " ORDER BY numlinea"
    End If
    
    
    Set R = New ADODB.Recordset
    R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
        If Not IsNull(R.Fields(0)) Then DevuelveCentroCosteFactura = R.Fields(0)
        LaPrimeraContrapartida = R!codtbase
        R.MoveNext
        If Not R.EOF Then LaPrimeraContrapartida = ""
    End If
    R.Close
    Set R = Nothing
End Function








