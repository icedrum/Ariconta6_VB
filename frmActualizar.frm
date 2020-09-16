VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
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
         TabIndex        =   4
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

Private Const maxAmpliaci = 46
Private Longi As Integer
Private AuxAmpliCta As String

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
Dim Sql As String
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
        Sql = " WHERE numserie = '" & LEtra & "'"
        Sql = Sql & " AND numfactu = " & NumFac
        Sql = Sql & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factcli_lineas " & Sql
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factcli_totales " & Sql
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
        Cuenta = "DELETE from factcli " & Sql
        Conn.Execute Cuenta

    Else
        '-------------------------------------------------------------
        '       P R O V E E D O R E S
        '-------------------------------------------------------------
        Sql = " WHERE numserie = '" & LEtra & "'"
        Sql = Sql & " AND numregis = " & NumFac
        Sql = Sql & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factpro_lineas " & Sql
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factpro_totales " & Sql
        Conn.Execute Cuenta
        
        
        Donde = "Documentos asociados"
        Cuenta = "DELETE from factpro_fichdocs " & Sql
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
        Cuenta = "DELETE from factpro " & Sql
        Conn.Execute Cuenta
        LEtra = RecuperaValor(NUmSerie, 1) '"1"
    End If

    bol = DesActualizaElASiento(Donde)

EEliminaFacturaConAsiento:
        If Err.Number <> 0 Then
            Sql = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            Sql = Sql & Donde
            MuestraError Err.Number, Sql, Err.Description
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
   ' Me.Animation1.visible = B
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
            MuestraError Err.Number, "Integra factura(I)" & vbCrLf & Donde & Err.Description
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
Dim cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim Amplia2 As String
Dim DocConcAmp As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean
Dim PrimeraContrapartida As String
Dim TipoDeIva_ As Integer
Dim B As Boolean

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
    Sql = "SELECT * FROM factcli"
    Sql = Sql & " WHERE numserie='" & NUmSerie
    Sql = Sql & "' AND numfactu= " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & Sql, vbExclamation
        RF.Close
        Exit Function
    End If
    
 
    Sql = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(Sql) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        Sql = "UPDATE hcabapu SET "
        Sql = Sql & " fechaent = " & DBSet(Fecha, "F")
        Sql = Sql & " where numdiari = " & DBSet(DiarioFacturas, "N")
        Sql = Sql & " and fechaent = " & DBSet(FechaAnterior, "F")
        Sql = Sql & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute Sql
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        Sql = Sql & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        Sql = Sql & ","
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        cad = DBLet(RF!observa, "T")
        If cad = "" Then
            cad = "NULL,"
        Else
            cad = "'" & DevNombreSQL(cad) & "',"
        End If
        cad = cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6:Contab. Fra cli " & NUmSerie & Format(NumFac, "0000000") & " " & Fecha & "')"
        
        
        Sql = Sql & cad
        Conn.Execute Sql
    End If
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    
    
    A_Donde = "Linea cliente"
    '-------------------------------------------------------------------
    'LINEA Cliente
    Sql = Mes & ",'" & RF!codmacta & "',"
    
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
            Longi = Len(Amplia2) + 7 'que ocupará con el CC
            Longi = maxAmpliaci - Longi
            If Longi < 0 Then Amplia2 = Mid(Amplia2, 1, maxAmpliaci - 7)
            Amplia2 = Amplia2 & " [" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    Sql = Sql & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & Cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    'Esta variable sirve para las demas
    ImporteNegativo = (DBLet(RF!totfaccl, "N") < 0)
    
    'Importes, atencion importes negativos
    '  antes --> Cad2 = CadenaImporte(ImporteNegativo, True, RF!totfaccl)
    Cad2 = CadenaImporte(True, DBLet(RF!totfaccl, "N"), Importe0)
    Sql = Sql & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'FRACLI',0)"
    
    
    Conn.Execute cad & Sql
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
        
        'Enero20
        Sql = "tipodiva"   '4:Suplido
        Cad2 = DevuelveDesdeBD3("cuentare", "tiposiva", "codigiva", DBLet(RsIvas!codigiva, "N"), "N", Cad3, Sql)
        TipoDeIva_ = Val(Sql)
        Sql = ""
        
        
        
        If Cad2 <> "" Then
        
            Sql = Mes & ",'" & Cad2 & "'," & DocConcAmp
            Cad2 = CadenaImporte(False, DBLet(RsIvas!Impoiva, "N"), Importe0)
            Sql = Sql & "," & Cad2 & ","
            Sql = Sql & "NULL,'" & DBLet(RF!codmacta, "T") & "','FRACLI',0)"
            'dependiendo de si ContabilizarAptIva0 = 1 se contabiliza o no el iva
            If Importe0 Then
                B = vParam.ContabApteIva0
                If TipoDeIva_ = 4 Then B = False 'Suplidos NO mete nada
                If B Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                End If
            Else
                Conn.Execute cad & Sql
                Mes = Mes + 1
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                     Sql = Mes & "," & Cad3 & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, DBLet(RsIvas!ImpoRec, "N"), Importe0)
                    Sql = Sql & "," & Cad2 & ","
                    Sql = Sql & "NULL,'" & DBLet(RF!codmacta, "T") & "','FRACLI',0)"
                    If Not Importe0 Then
                        Conn.Execute cad & Sql
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
        Sql = Mes & ",'" & DBLet(RF!cuereten, "T") & "'," & DocConcAmp
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(True, DBLet(RF!trefaccl, "T"), Importe0)
        Sql = Sql & "," & Cad2 & ","
        'Antes Nov 2019
        'SQL = SQL & "NULL,NULL,'FRACLI',0)"
        Sql = Sql & "NULL," & DBSet(RF!codmacta, "T")
        Sql = Sql & ",'FRACLI',0)"
       
       
        Conn.Execute cad & Sql
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
    Sql = "Select cuentas.codmacta, factcli_lineas.codccost, sum(factcli_lineas.baseimpo) baseimpo FROM factcli_lineas,cuentas "
    Sql = Sql & " WHERE numserie='" & NUmSerie
    Sql = Sql & "' AND numfactu= " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    Sql = Sql & " AND factcli_lineas.codmacta = cuentas.codmacta"
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    Cad2 = ""
    A_Donde = "Procesando lineas"
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        Sql = Mes & ",'" & DBLet(RF!codmacta, "T") & "'," & DocConcAmp
        Cad2 = CadenaImporte(False, DBLet(RF!Baseimpo, "N"), Importe0)
        Sql = Sql & "," & Cad2 & ","
        If IsNull(RF!CodCCost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!CodCCost & "'"
        End If
        
        Sql = Sql & Cad2 & ",'" & DBLet(Cuenta, "T") & "','FRACLI',0)"
    
        Conn.Execute cad & Sql
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
            Sql = "UPDATE factcli_lineas SET codmacta ='" & PrimeraContrapartida & "'"
            Sql = Sql & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            Sql = Sql & " AND numlinea =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL Sql  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
        
    
    
    
    'Actualimos en factura, el nº de asiento
    Sql = "UPDATE factcli SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    Sql = Sql & " WHERE numserie='" & NUmSerie
    Sql = Sql & "' AND numfactu= " & NumFac
    Sql = Sql & " AND anofactu= " & NumDiari
    Conn.Execute Sql
    
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
Dim cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim Ampliacion As String
Dim DocumtoYcodConce As String
Dim DocConcAmp_ As String   'dcumento, codconce, ampliaci
Dim numeroFacturaParaObservacionesLinea As String
Dim Amplia2 As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean 'Para saber si el importe es 0
Dim PrimeraContrapartida As String  'Si hay solo una linea entonces la pondremos como contrapartida de la primera base
Dim TipoDeIva_ As Integer
Dim B As Boolean

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
    Sql = "SELECT * FROM factpro"
    Sql = Sql & " WHERE numregis = " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    Sql = Sql & " AND numserie=" & DBSet(NUmSerie, "T")
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & Sql, vbExclamation
        RF.Close
        Exit Function
    End If
    
    Sql = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(Sql) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        Sql = "UPDATE hcabapu SET "
        Sql = Sql & " fechaent = " & DBSet(Fecha, "F")
        Sql = Sql & " where numdiari = " & DBSet(DiarioFacturas, "N")
        Sql = Sql & " and fechaent = " & DBSet(FechaAnterior, "F")
        Sql = Sql & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute Sql
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        Sql = Sql & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        cad = DBLet(RF!observa, "T")
        If cad = "" Then
            cad = "NULL,"
        Else
            cad = "'" & DevNombreSQL(cad) & "',"
        End If
        
        cad = cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Fra.Pro Reg:" & Format(NumFac, "000000") & " " & Fecha & "')"
        
        Sql = Sql & "," & cad
        
        Conn.Execute Sql
        
    End If
    
    
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    PrimeraContrapartida = ""
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RF!totfacpr < 0)
    A_Donde = "Linea proveedor"
    '-------------------------------------------------------------------
    'LINEA Proveedor
    Sql = Mes & ",'" & RF!codmacta & "',"
    
    'Documento "numdocum"
    If vParam.CodiNume = 1 Then
        Cad2 = Format(NumFac, "0000000000")
    Else
        Cad2 = DBLet(RF!numfactu)
    End If
    
    DocumtoYcodConce = "'" & Cad2 & "'," & vParam.concefpr & ","
    
    
    'Ampliacion segun parametros
    Select Case vParam.nctafact
    Case 1
        If RF!totfacpr < 0 Then
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 2)
        Else
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 1)
        End If
        Cad2 = Cad2 & " " & DevNombreSQL(RF!numfactu)
        
        Cad2 = Cad2 & " (" & Format(RF!FecFactu, "ddmmyy") & ")"
    Case 2
        Cad2 = DevNombreSQL(DBLet(RF!Nommacta))
        
        
    Case Else
        Cad2 = DBLet(RF!observa)
    End Select
    
    
    If vParam.ObsrvacionesLineasFrapro Then numeroFacturaParaObservacionesLinea = RF!numfactu
    
    
    'Agosto 2020
    'Si no es numero fac, y parametro añadir factura
    If vParam.nctafact <> 1 Then
        If vParam.AnyadeNumFaproAmpliacApunte Then
    
            'Private Const maxAmpliaci = 50
            AuxAmpliCta = "Nº" & RF!numfactu
            Longi = 0
            If vParam.CCenFacturas Then Longi = 7
            Longi = maxAmpliaci - Len(AuxAmpliCta) - 1 - Longi
            If Longi > 0 Then Cad2 = Mid(Cad2, 1, Longi)
            Cad2 = Cad2 & " " & AuxAmpliCta
            Cad2 = DevNombreSQL(Cad2)
            
        End If
    End If
        
        
    
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
    A_Donde = "Linea proveedor"
    
    
    Sql = Sql & DocumtoYcodConce & DBSet(Amplia2, "T")
    DocConcAmp_ = DocumtoYcodConce & DBSet(Cad2, "T")   'DocConcAmp Sirve para el IVA
    
    
    'Importes, atencion importes negativos
    Cad2 = CadenaImporte(False, RF!totfacpr, Importe0)
    Sql = Sql & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'FRAPRO',0)"
    
    Conn.Execute cad & Sql
    Mes = Mes + 1 'Es el contador de lineaapunteshco
    
    ' cuentas de iva ahora se sacan de las tablas de totales
    SqlIva = "select * from factpro_totales "
    SqlIva = SqlIva & " WHERE numserie='" & NUmSerie
    SqlIva = SqlIva & "' AND numregis= " & NumFac
    SqlIva = SqlIva & " AND anofactu=" & NumDiari
    SqlIva = SqlIva & " order by numlinea "
    
    
    Dim Apunte_472_477 As Boolean   'Ante sEsSujetoPasivo:  INTRACOMUNITARIA  - SUJETO PASIVO
    Dim EsImportacion As Boolean
    
    EsImportacion = (DBLet(RF!CodOpera, "N") = 2)
    Apunte_472_477 = ((DBLet(RF!CodOpera, "N") = 1) Or (DBLet(RF!CodOpera, "N") = 4))   ' Or (DBLet(RF!CodOpera, "N") = 6)    'DUA hace apunte normal
    
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
        
        Cad3 = "cuentaso"
        'Cad2 = DevuelveDesdeBD(ColumnaIVA, "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
        Sql = "tipodiva"   '4:Suplido
        Cad2 = DevuelveDesdeBD3(ColumnaIVA, "tiposiva", "codigiva", DBLet(RsIvas!codigiva, "N"), "N", Cad3, Sql)
        TipoDeIva_ = Val(Sql)
        Sql = ""
        
        
        
        If Cad2 <> "" Then
            Sql = Mes & ",'" & Cad2 & "'," & DocConcAmp_
            Cad2 = CadenaImporte(True, RsIvas!Impoiva, Importe0)
            Sql = Sql & "," & Cad2 & ","
            Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
            
            If Importe0 Then
                
                
                
                B = False
                If vParam.ContabApteIva0 Then
                    If Not EsImportacion Then
                        B = TipoDeIva_ <> 4
                    End If
                End If
                    
                If B Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                End If
            Else
                If Not EsImportacion Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                End If
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                Sql = Mes & "," & Cad3 & "," & DocConcAmp_
                'Importes, atencion importes negativos
                Cad2 = CadenaImporte(True, RsIvas!ImpoRec, Importe0)
                Sql = Sql & "," & Cad2 & ","
                Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                If Not Importe0 Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                End If
            End If
            
            If Apunte_472_477 Then
                Cad3 = "cuentarr"
                Cad2 = DevuelveDesdeBD("cuentare", "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
                
                Cad3 = Cad2 & "|" & Cad3 & "|"
                
                
                Sql = Mes & ",'" & RecuperaValor(Cad3, 1) & "'," & DocConcAmp_
                Cad2 = CadenaImporte(False, RsIvas!Impoiva, Importe0)
                Sql = Sql & "," & Cad2 & ","
                Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                'If Not Importe0 Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                'End If
               
                If Not IsNull(RsIvas!ImpoRec) Then
                     Sql = Mes & "," & RecuperaValor(Cad3, 2) & "," & DocConcAmp_
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, RsIvas!ImpoRec, Importe0)
                    Sql = Sql & "," & Cad2 & ","
                    Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        Conn.Execute cad & Sql
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
        Sql = Mes & ",'" & RF!cuereten & "'," & DocConcAmp_
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(False, DBLet(RF!trefacpr, "N"), Importe0)
        Sql = Sql & "," & Cad2 & ","
        'Antes Nov19
        'SQL = SQL & "NULL,NULL,'FRAPRO',0)"
        
        Sql = Sql & "NULL," & DBSet(RF!codmacta, "T")
        Sql = Sql & ",'FRAPRO',0)"
        
       
        Conn.Execute cad & Sql
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
    If vParam.ObsrvacionesLineasFrapro Then   'septiembre2020
        Sql = "Select factpro_lineas.codmacta, factpro_lineas.codccost,  baseimpo ,observalin FROM factpro_lineas "
    Else
        Sql = "Select factpro_lineas.codmacta, factpro_lineas.codccost, sum(factpro_lineas.baseimpo) baseimpo  FROM factpro_lineas "
    End If
    'where
    Sql = Sql & " WHERE numregis= " & NumFac
    Sql = Sql & " AND numserie= " & DBSet(NUmSerie, "T")
    Sql = Sql & " AND anofactu=" & NumDiari
    If Not vParam.ObsrvacionesLineasFrapro Then Sql = Sql & " GROUP BY 1,2 "
    Sql = Sql & " ORDER BY 1,2 "
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    A_Donde = "Procesando lineas"
    Cad2 = ""
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        Sql = ""
        If vParam.ObsrvacionesLineasFrapro Then Sql = DBLet(RF!observalin, "T")
        
        
        If Sql <> "" Then 'Tiene
            'Añade numero factura al final de la ampliacion
            If vParam.AnyadeNumFaproAmpliacApunte Then
    
                'Private Const maxAmpliaci = 50
                AuxAmpliCta = numeroFacturaParaObservacionesLinea
                Longi = maxAmpliaci - Len("Nº" & AuxAmpliCta) - 1
                If Longi > 0 Then Sql = Mid(Sql, 1, Longi)
                Sql = Sql & " Nº" & AuxAmpliCta
            
            End If
        
        
        
        
            Sql = Mes & ",'" & RF!codmacta & "'," & DocumtoYcodConce & DBSet(Sql, "T")
        Else
            'lo que habia
            Sql = Mes & ",'" & RF!codmacta & "'," & DocConcAmp_
        End If
        Cad2 = CadenaImporte(True, RF!Baseimpo, Importe0)
        Sql = Sql & "," & Cad2 & ","
        If IsNull(RF!CodCCost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!CodCCost & "'"
        End If
        
        Sql = Sql & Cad2 & ",'" & Cuenta & "','FRAPRO',0)"
    
        Conn.Execute cad & Sql
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
            Sql = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
            Sql = Sql & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            Sql = Sql & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL Sql  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
    
    'Actualimos en factura, el nº de asiento
    Sql = "UPDATE factpro SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    Sql = Sql & " ,fecregcontable = fecregcontable"
    Sql = Sql & " WHERE  numregis = " & NumFac
    Sql = Sql & " AND numserie= " & DBSet(NUmSerie, "T")
    Sql = Sql & " AND anofactu=" & NumDiari
    Conn.Execute Sql
    
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
    Sql = "SELECT numdiari from hcabapu"
    Sql = Sql & " WHERE numdiari =" & NumDiari
    Sql = Sql & " AND fechaent='" & Fecha & "'"
    Sql = Sql & " AND numasien=" & NumAsiento
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
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
    Sql = "Delete from hlinapu"
    Sql = Sql & " WHERE numasien = " & NumAsiento
    Sql = Sql & " AND numdiari = " & NumDiari
    Sql = Sql & " AND fechaent=" & DBSet(FechaAnterior, "F")
    Conn.Execute Sql
    
    If BorrarCabecera Then
        'La cabecera
        Sql = "Delete from hcabapu"
        Sql = Sql & " WHERE numdiari =" & NumDiari
        Sql = Sql & " AND fechaent=" & DBSet(FechaAnterior, "F")
        Sql = Sql & " AND numasien=" & NumAsiento
        
        Conn.Execute Sql
    Else
        'Actualizamos la fecha de la cabecera
        Sql = "Update hcabapu"
        Sql = Sql & " set fechaent = " & DBSet(Fecha, "F")
        Sql = Sql & " WHERE numdiari =" & NumDiari
        Sql = Sql & " AND fechaent=" & DBSet(FechaAnterior, "F")
        Sql = Sql & " AND numasien=" & NumAsiento
    
        Conn.Execute Sql
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








Private Function InsertaError(ByRef Cadena As String)
Dim vS As String
    'Insertamos en errores
    'Esta lo tratamos con error especifico
    
    On Error Resume Next

    If OpcionActualizar < 10 Then
        'Insertamos error para ASIENTOS
        vS = NumDiari & "|"
        vS = vS & Fecha & "|"
        vS = vS & NumAsiento & "|"
        vS = vS & Cadena & "|"
    
    Else
        vS = NUmSerie & " " & NumFac & "|"
        vS = vS & FechaAsiento & "|"
        vS = vS & Cadena & "|"
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
            Sql = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            Sql = Sql & Donde
            MuestraError Err.Number, Sql, Err.Description
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
    Sql = "Select * From " & CCost
    Sql = Sql & " WHERE numasien=" & NumAsiento
    Sql = Sql & " AND numdiari = " & NumDiari
    Sql = Sql & " AND fechaent = '" & Fecha & "'"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        
        Sql = "UPDATE " & CCost & " SET numasien=NULL, fechaent=NULL, numdiari=NULL"
        If Clientes Then
            Sql = Sql & " WHERE numfactu = " & Rs!codfaccl
            Sql = Sql & " AND anofaccl =" & Rs!anofaccl
            Sql = Sql & " AND numserie = '" & Rs!NUmSerie & "'"
        Else
            'proveedores
            Sql = Sql & " ,fecregcontable=fecregcontable"
            Sql = Sql & " WHERE numregis = " & Rs!Numregis
            Sql = Sql & " AND anofactu =" & Rs!Anofactu
        End If
        Conn.Execute Sql
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
Dim Sql As String
    DevuelveCentroCosteFactura = ""
    If Cliente Then
        
        'Sql = "SELECT codccost,numlinea,baseimpo codtbase FROM factcli_lineas"
        Sql = "SELECT codccost,numlinea,codmacta FROM factcli_lineas"
        Sql = Sql & " WHERE numserie='" & NUmSerie
        Sql = Sql & "' AND numfactu= " & NumFac
        Sql = Sql & " AND anofactu=" & NumDiari
        Sql = Sql & " AND not (codccost is null)"   'El primero k devuelva
        Sql = Sql & " ORDER BY numlinea"
    Else
        Sql = "SELECT codccost,numlinea,codmacta FROM factpro_lineas"
        Sql = Sql & " WHERE numregis = " & NumFac
        Sql = Sql & " AND anofactu=" & NumDiari
        Sql = Sql & " AND numserie=" & DBSet(NUmSerie, "T")
        Sql = Sql & " AND not (codccost is null)"   'El primero k devuelva
        Sql = Sql & " ORDER BY numlinea"
    End If
    
    
    Set R = New ADODB.Recordset
    R.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
        If Not IsNull(R.Fields(0)) Then DevuelveCentroCosteFactura = R.Fields(0)
        LaPrimeraContrapartida = DBLet(R!codmacta, "T")
        R.MoveNext
        If Not R.EOF Then LaPrimeraContrapartida = ""
    End If
    R.Close
    Set R = Nothing
End Function








