VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESTransferenciasCont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmTESTransferenciasCont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameContabilRem2 
      Height          =   4215
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5535
      Begin VB.CheckBox chkAgrupaCancelacion 
         Caption         =   "Agrupa cancelacion"
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
         Left            =   450
         TabIndex        =   6
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CommandButton cmdContabRemesa 
         Caption         =   "Contabilizar"
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
         Left            =   2490
         TabIndex        =   4
         Top             =   3600
         Width           =   1425
      End
      Begin VB.TextBox txtImporte 
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
         Left            =   2640
         TabIndex        =   3
         Text            =   "Text4"
         Top             =   2520
         Width           =   1395
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
         Index           =   10
         Left            =   2640
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1920
         Width           =   1365
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   8
         Left            =   4020
         TabIndex        =   5
         Top             =   3600
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Gastos banco (�)"
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
         Index           =   0
         Left            =   450
         TabIndex        =   8
         Top             =   2490
         Width           =   2070
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Contable"
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
         Left            =   450
         TabIndex        =   7
         Top             =   1950
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   10
         Left            =   2280
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "CONTABILIZAR REMESA"
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
         Height          =   1245
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmTESTransferenciasCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '8.- Contabilizar remesa
        
    
    
Public SubTipo As Byte

    'Para la opcion 22
    '   Remesas cancelacion cliente.
    '       1:  Efectos
    '       2: Talones pagares
    
'Febrero 2010
'Cuando pago proveedores con un talon, y le he indicado el numero
Public NumeroDocumento As String
Public vTextos As String

Public Cobros As Boolean
Public TipoTrans As Byte ' 0=transferencia de abonos
                         ' 1=transferencias de pagos
                         ' 2=pagos domiciliados
                         ' 3=confirming
    
Public ImporteGastosTarjeta_ As Currency   'Para cuando viene de recepciondocumentos pondre el importe que le falta
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1


Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Integer
Dim IT As ListItem  'Comun
Dim PrimeraVez As Boolean
Dim Cancelado As Boolean
Dim CuentasCC As String
Dim impo As Currency
Dim riesgo As Currency
Dim Tipo As Integer
Dim ContabTransfer As Boolean
Dim Fecha As Date
Dim FechaAsiento As Date
Private vp As Ctipoformapago
Private SubItemVto As Integer
Dim DescripcionTransferencia As String
Dim NumeroTalonPagere As String


Private GastosTransferencia As Currency

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 21 Or Index = 25 Or Index = 31 Then CadenaDesdeOtroForm = "" 'ME garantizo =""
    If Index = 31 Then
        If MsgBox("�Cancelar el proceso?", vbQuestion + vbYesNo) = vbYes Then SubTipo = 0
    End If
    Unload Me
End Sub




Private Sub cmdContabRemesa_Click()
Dim B As Boolean
Dim Importe As Currency
Dim CC As String
Dim Opt As Byte
Dim AgrupaCance As Boolean
Dim ContabilizacionEspecialNorma19 As Boolean


'Dim ImporteEnRecepcion As Currency
'Dim TalonPagareBeneficios As String
    Sql = ""
    
    If Text1(10).Text = "" Then Sql = "Ponga la fecha de abono"
    
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(10).Text), True) > 1 Then Exit Sub
    
    
    'Ahora miramos la remesa. En que sitaucion , y de que tipo es
    Sql = "Select * from transferencias where codigo =" & RecuperaValor(NumeroDocumento, 1)
    Sql = Sql & " AND anyo =" & RecuperaValor(NumeroDocumento, 2)
    If Cobros Then
        Sql = Sql & " and tipotrans = 1 "
    Else
        Sql = Sql & " and tipotrans = 0 "
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Rs.EOF Then
        Select Case TipoTrans
            Case 0, 1
                MsgBox "Ninguna transferencia con esos valores", vbExclamation
            Case 2
                MsgBox "Ning�n pago domiciliado con esos valores", vbExclamation
            Case 3
                MsgBox "Ning�n confirming con esos valores", vbExclamation
        End Select
        Rs.Close
        Set Rs = Nothing
        Exit Sub

    End If
    
    'Tiene valor
    Sql = ""
    B = AdelanteConLaTransferencia()
    ContabilizacionEspecialNorma19 = False
    
    If Cobros Then
    
        Sql = "Select cobros.codmacta,nomclien,fecbloq from cobros,cuentas where cobros.codmacta = cuentas.codmacta"
        Sql = Sql & " and transfer =" & RecuperaValor(NumeroDocumento, 1)
        Sql = Sql & " AND anyorem =" & RecuperaValor(NumeroDocumento, 2)
        Sql = Sql & " AND fecbloq <='" & Format(Text1(10).Text, FormatoFecha) & "' GROUP BY 1"
        
    Else
        Sql = "Select pagos.codmacta,nomprove nomclien,fecbloq from pagos,cuentas where pagos.codmacta = cuentas.codmacta"
        Sql = Sql & " and nrodocum =" & RecuperaValor(NumeroDocumento, 1)
        Sql = Sql & " AND anyodocum =" & RecuperaValor(NumeroDocumento, 2)
        Sql = Sql & " AND fecbloq <='" & Format(Text1(10).Text, FormatoFecha) & "' GROUP BY 1"
    
    
    
    End If
        
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not miRsAux.EOF
        Sql = Sql & miRsAux!codmacta & Space(10) & miRsAux!FecBloq & Space(10) & miRsAux!nomclien & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If Sql <> "" Then
        CC = "Cuenta          Fec. bloqueo           Nombre" & vbCrLf & String(80, "-") & vbCrLf
        CC = "Cuentas bloqueadas" & vbCrLf & vbCrLf & CC & Sql
        MsgBox CC, vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Sub
    End If
       
       
       
    'Bloqueariamos la opcion de modificar esa remesa
        
    Importe = TextoAimporte(txtImporte(0).Text)

    'Tiene gastos. Falta ver si tiene la cuenta de gastos configurada. ASi como
    'si es analitica, el CC asociado
    CC = ""
    If vParam.autocoste Then CC = "codccost"
        
    Sql = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", Rs!codmacta, "T", CC)
    If Sql = "" Then
        MsgBox "Falta configurar la cuenta de gastos del banco:" & Rs!codmacta, vbExclamation
        Set Rs = Nothing
        Exit Sub
    End If
    
    If vParam.autocoste Then
        If CC = "" Then
            MsgBox "Necesita asignar centro de coste a la cuenta de gastos del banco: " & Rs!codmacta, vbExclamation
            Set Rs = Nothing
            Exit Sub
        End If
    End If
    
    Sql = Sql & "|" & CC & "|"
      
      
      'A�ado, si tiene, la cuenta de ingresos
      CC = DevuelveDesdeBD("ctaingreso", "bancos", "codmacta", Rs!codmacta, "T")
      If CC = "" Then
          If Importe > 0 Then
              MsgBox "Falta configurar la cuenta de ingresos del banco:" & Rs!codmacta, vbExclamation
              Set Rs = Nothing
              Exit Sub
          End If
      End If
      
      Sql = Sql & CC & "|"   'La
      

    Sql = Rs!codmacta & "|" & Sql
    
    
    'Contab. remesa. Si es talon/pagare vamos a comprobar si hay diferencias entre el importe del documento
    'y el total de lineas
    B = False    'Si ya se ha hecho la pregunta no la volveremos a repetir
    'TalonPagareBeneficios = ""    'Solo para TAL/PAG y si hay importe beneficios etc

    
    'Pregunta conbilizacion
    If Not B Then   'Si no hemos hecho la pregunta en otro sitio la hacemos ahora
        Select Case Opcion
        Case 8
            CC = "Va a abonar"
        Case 22
            CC = "Procede a realizar la cancelacion del cliente de"
        Case 23
            CC = "Procede a realizar la confirmacion de"
        End Select
        Select Case TipoTrans
            Case 0, 1
                CC = CC & " la transferencia: " & Rs!Codigo & " / " & Rs!Anyo & vbCrLf & vbCrLf
            Case 2
                CC = CC & " el pago domiciliado: " & Rs!Codigo & " / " & Rs!Anyo & vbCrLf & vbCrLf
            Case 3
                CC = CC & " el confirming: " & Rs!Codigo & " / " & Rs!Anyo & vbCrLf & vbCrLf
        End Select
        CC = CC & Space(30) & "�Continuar?"
        If SubTipo = 2 Then
            If Val(Rs!Tiporem) = 3 Then
                CC = "Tal�n" & vbCrLf & CC
            Else
                CC = "Pagar�" & vbCrLf & CC
            End If
            CC = "Tipo: " & CC
        End If
    
        If MsgBox(CC, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Para llevarlos a hco
    Conn.Execute "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    
        
    
    'CONTABILIZACION    ABONO REMESA
    
    'NORMA 19
    '------------------------------------
    
    'Contabilizaremos la remesa
    Conn.BeginTrans
    
    'mayo 2012
    
    B = HacerNuevaContabilizacion
    
    'si se contabiliza entonces updateo y la pongo en
    'situacion Q. Contabilizada a falta de devueltos ,
    If B Then
        Conn.CommitTrans
        'AQUI updateamos el registro pq es una tabla myisam
        'y no debemos meterla en la transaccion
        
        HaHabidoCambios = True
        
        
        Sql = "UPDATE transferencias SET"
        Sql = Sql & " situacion= 'Q'"
        Sql = Sql & " WHERE codigo=" & RecuperaValor(NumeroDocumento, 1)
        Sql = Sql & " and anyo=" & RecuperaValor(NumeroDocumento, 2)
        

        If Not Ejecuta(Sql) Then MsgBox "Error actualizando tabla transferencias.", vbExclamation
        
        If Cobros Then
            Sql = "update cobros set siturem = 'Q', situacion = 1 "
            Sql = Sql & " WHERE transfer=" & RecuperaValor(NumeroDocumento, 1)
            Sql = Sql & " and anyorem=" & RecuperaValor(NumeroDocumento, 2)
        
            If Not Ejecuta(Sql) Then MsgBox "Error actualizando tabla cobros.", vbExclamation
        
        Else
            Sql = "update pagos set situdocum = 'Q', situacion = 1 "
            Sql = Sql & " WHERE nrodocum=" & RecuperaValor(NumeroDocumento, 1)
            Sql = Sql & " and anyodocum=" & RecuperaValor(NumeroDocumento, 2)
            
            If Not Ejecuta(Sql) Then MsgBox "Error actualizando tabla pagos.", vbExclamation
            
        End If
        
        
        'Ahora actualizamos los registros que estan en tmpactualziar
        Screen.MousePointer = vbDefault
        'Cerramos
        'RS.Close
        Unload Me
        Exit Sub
    Else
        TirarAtrasTransaccion
    End If


    
    
    Rs.Close
    Set Rs = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Function HacerNuevaContabilizacion() As Boolean



    On Error GoTo EHacer
    HacerNuevaContabilizacion = False
    
    Tipo = 1
    
    ContabTransfer = True
    
    If Cobros Then
        GastosTransferencia = DBSet(txtImporte(0).Text, "N") * (-1)
    Else
        GastosTransferencia = DBSet(txtImporte(0).Text, "N")
    End If
    
    
    'Si el parametro dice k van todos en el mismo asiento, pues eso, todos en el mismo asiento
    'Primero leemos la forma de pago, el tipo perdon
    Set vp = New Ctipoformapago
    
    Dim cad As String
    
    
    'en vtextos, en el 3 tenemos la forpa
    If vp.Leer(vbTransferencia) = 1 Then
        'ERROR GRAVE LEYENDO LA FORMA DE PAGO
        Screen.MousePointer = vbDefault
        Set vp = Nothing
        End
    End If
    
    
    
    'Paso1. Meto todos los seleccionados en una tabla
    If Not InsertarPagosEnTemporal2 Then Exit Function
    
    
    
    'Paso 2
    'Compruebo que los vtos a cobrar no tienen ni la cuenta bloqueada, ni,
    'si contabilizo por fecha de bloqueo, alguna de los vencimienotos
    'esta fuera del de fechas
    If Not ComprobarCuentasBloquedasYFechasVencimientos Then Exit Function
    
    
    
    'Contabilizo desde la tabla. Asi puedo agrupar mejor
    ContablizaDesdeTmp
    
    HacerNuevaContabilizacion = True
    
    
    Exit Function
EHacer:
    MuestraError Err.Number, "Contabilizando"
End Function

Private Function InsertarPagosEnTemporal2() As Boolean
Dim C As String
Dim Aux As String
Dim J As Long
Dim FechaContab As Date
Dim FechaFinEjercicios As Date
Dim vGasto As Currency

Dim Sql1 As String
Dim Rs As ADODB.Recordset
Dim cad As String

    InsertarPagosEnTemporal2 = False
    
    C = " WHERE codusu =" & vUsu.Codigo
    Conn.Execute "DELETE FROM tmpfaclin" & C


    'Fechas fin ejercicios
    FechaFinEjercicios = DateAdd("yyyy", 1, vParam.fechafin)

    If Cobros Then
        Sql1 = "select * from cobros where transfer = " & DBSet(RecuperaValor(NumeroDocumento, 1), "N") & " and anyorem = " & RecuperaValor(NumeroDocumento, 2)
    Else
        Sql1 = "select * from pagos where nrodocum = " & DBSet(RecuperaValor(NumeroDocumento, 1), "N") & " and anyodocum = " & RecuperaValor(NumeroDocumento, 2)
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    

     'codusu,j,FechaPosibleVto,FechaVto,Cta,SerieFactura|Fechafac|,ctacobro,IMpoorte,gastos)
     'NUEVO. Febrero 2010.
     'Llevar serie, fecha y NUMORDEN
     'codusu,j,FechaPosibleVto,FechaVto,Cta,SerieFactura|Fechafac|numorden|,ctacobro,IMpoorte,gastos)
    Aux = "INSERT INTO tmpfaclin (codusu, codigo, Fecha,Numfactura, cta, Cliente, NIF, Imponible,  Total) "
    Aux = Aux & "VALUES (" & vUsu.Codigo & ","
    
    J = 0
    While Not Rs.EOF
        J = J + 1
            
        C = J & ",'"
        'Si la fecha de contabilizacion esta fuera de ejercicios
        If Cobros Then
            FechaContab = DBLet(Rs!FecVenci, "F")
        Else
            FechaContab = DBLet(Rs!fecefect, "F")
        End If
            

        C = C & Format(FechaContab, FormatoFecha) & "','"
        
        '-----------------------------------------------------
        'Fecha de contabilizacion
            'La fecha de contabilizacion es la del text
        FechaContab = CDate(Text1(10).Text)
        'MEto la fecha de contabilizaccion
        C = C & Format(FechaContab, FormatoFecha) & "','"
        'Cuenta contable
        C = C & DBLet(Rs!codmacta, "T") & "','"
        'Serie factura |FECHAfactura|
        'Neuvo febrero 2008 Serie factura |FECHAfactura|numvto|
        If Cobros Then
            C = C & DBLet(Rs!NUmSerie, "T") & "|" & DBLet(Rs!NumFactu, "N") & "|" & DBLet(Rs!FecFactu, "F") & "|" & DBLet(Rs!numorden, "N")
        Else
            C = C & DBLet(Rs!NUmSerie, "T") & "|" & DBLet(Rs!NumFactu, "T") & "|" & DBLet(Rs!FecFactu, "F") & "|" & DBLet(Rs!numorden, "N")
        End If
'        Else
'            C = C & DevNombreSQL(ListView1.ListItems(J).Text) & "|" & ListView1.ListItems(J).SubItems(1) & "|" & ListView1.ListItems(J).SubItems(3)
'        End If
        C = C & "|','',"
        '###AQUI
        
        'Dinerito
        'riesgo es GASTO
        If Cobros Then
            impo = DBLet(Rs!ImpVenci, "N")
        Else
            impo = DBLet(Rs!ImpEfect, "N")
        End If
        
        If Cobros Then
            riesgo = ImporteFormateado(DBLet(Rs!Gastos, "N"))
        Else
            riesgo = 0
        End If
        impo = impo - riesgo
        C = C & TransformaComasPuntos(CStr(impo)) & "," & TransformaComasPuntos(CStr(riesgo)) & ")"
        

        'Lo meto en la BD
        C = Aux & C
        Conn.Execute C
    
        Rs.MoveNext
    
    Wend

    
    'Gastos contabilizacion transferencia
    If GastosTransferencia <> 0 Then
            J = J + 1
    
            'aqui ira los gastos asociados a la transferencia
            'Hay que ver los lados
            
            'Cad = DevuelveDesdeBD("ctagastostarj", "ctabancaria", "codmacta", Text3(1).Tag, "T")
            cad = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", RecuperaValor(NumeroDocumento, 4), "T")
            
            FechaContab = CDate(Text1(10).Text)
            C = "'" & Format(FechaContab, FormatoFecha) & "'"
            C = C & "," & C
            C = J & "," & C & ",'" & cad & "','"
            'Serie factura |FECHAfactura| ----> pondre: "gastos" | fecha contab
            C = C & "TRA" & Format(RecuperaValor(NumeroDocumento, 1), "0000000") & "|" & FechaContab & "|','" & cad & "',"
            'Dinerito
            'riesgo es GASTO
            impo = GastosTransferencia
            C = C & TransformaComasPuntos(CStr(impo)) & ",0)"
            C = Aux & C
            Conn.Execute C
        
    End If
    
    InsertarPagosEnTemporal2 = True
    
    

End Function


Private Sub ContablizaDesdeTmp()
Dim Sql As String
Dim ContraPartidaPorLinea As Boolean
Dim UnAsientoPorCuenta As Boolean
Dim PonerCuentaGenerica As Boolean
Dim AgrupaCuenta As Boolean
Dim Rs As ADODB.Recordset
Dim MiCon As Contadores
Dim CampoCuenta As String
Dim CampoFecha As String
Dim GeneraAsiento As Boolean
Dim CierraAsiento As Boolean
Dim NumLinea As Integer
Dim ImpBanco As Currency
Dim NumVtos As Integer
Dim GastosTransDescontados As Boolean
Dim LineaUltima As Integer

Dim cad As String

    'Valores por defecto
    ContraPartidaPorLinea = False
    UnAsientoPorCuenta = False
    PonerCuentaGenerica = False
    AgrupaCuenta = False
    CampoFecha = "numfactura" '"numfac"
    GastosTransDescontados = False 'por lo que pueda pasar
    
    'Si va agrupado por cta
    If Tipo = 1 And ContabTransfer Then
        
        'Si lleva GastosTransferencia entonce AGRUPAMOS banco
        If GastosTransferencia <> 0 Then
            
            'gastos tramtiaacion transferenca descontados importe
            Sql = DevuelveDesdeBD("GastTransDescontad", "bancos", "codmacta", RecuperaValor(NumeroDocumento, 4), "T")
            GastosTransDescontados = Sql = "1"
            
            AgrupaCuenta = False
        
        End If
    End If
    
    If PonerCuentaGenerica Then
        CampoCuenta = "NIF"
    Else
        CampoCuenta = "cta"
    End If
    'EL SQL lo empezamos aquin
    Sql = CampoCuenta & " AS cliprov,"
    'Selecciona
    Sql = "select count(*) as numvtos,codigo,numfactura,fecha,cliente," & Sql & "sum(imponible) as importe,sum(total) as gastos from tmpfaclin"
    Sql = Sql & " where codusu =" & vUsu.Codigo & " GROUP BY "
    cad = ""
    If AgrupaCuenta Then
       If PonerCuentaGenerica Then
            cad = "nif" 'La columna NIF lleva los datos de la cuenta generica
        Else
            cad = "cta"
        End If
        'Como estamos agrupando por cuenta, marcaremos tb la fecha
        'Ya que si tienen fechas distintas son apuntes distintos
        cad = cad & "," & CampoFecha
    End If
    
    'Si no agrupo por nada agrupare por codigo(es decir como si no agrupara)
    If cad = "" Then cad = "codigo"
    
    'La ordenacion
    cad = cad & " ORDER BY " & CampoFecha
    If Not PonerCuentaGenerica Then cad = cad & ",cta"
        
    
    'Tanto si agrupamos por cuenta (Generica o no)
    'el recodset tendra las lineas que habra que insertar en/los apuntes(s)
    '
    'Es decir. Que si agrupo no tengo que ir moviendome por el recodset mirando a ver si
    'las cuentas son iguales.
    'Ya que al hacer group by ya lo estaran
    cad = Sql & cad
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Inicializamos variables
    Fecha = CDate("01/01/1900")
    GeneraAsiento = False
    While Not Rs.EOF
        'Comprobaciones iniciales
        If UnAsientoPorCuenta Then
            'Para cada linea ira su asiento
            GeneraAsiento = True
            CierraAsiento = True
            If Fecha < CDate("01/01/1950") Then CierraAsiento = False
            Fecha = CDate(Rs.Fields(CampoFecha))
        Else
            'Veremos en funcion de la fecha
            GeneraAsiento = False
            If CDate(Rs.Fields(CampoFecha)) = Fecha Then
                'Estamos en la misma fecha. Luego sera el mismo asiento
                'Excepto que asi no lo digan las variables
                If Not PonerCuentaGenerica Then
                    If UnAsientoPorCuenta Then
                        GeneraAsiento = True
                        If Fecha < CDate("01/01/1950") Then CierraAsiento = True
                    End If
                End If
                        
            Else
                'Fechas distintas.
                GeneraAsiento = True
                CierraAsiento = True
                If Fecha < CDate("01/01/1950") Then CierraAsiento = False
        
                Fecha = CDate(Rs.Fields(CampoFecha))
            End If
        End If 'de aseinto por cuenta
        
        
        'Si tengo que cerrar el asiento anterior
        If CierraAsiento Then
            'Tirar atras el RS
            If Not ContraPartidaPorLinea Then
                Rs.MovePrevious
                Fecha = CDate(Rs.Fields(CampoFecha))  'Para la fecha de asiento
                impo = ImpBanco
                'Generamos las lineas de apunte que faltan
                InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
                
                'Inserto para que actalice             3: Opcion para INSERT INTO tmpactualizar
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
                
                'Reestauramos variables
                NumVtos = 0
                'Ponemos la variable
                CierraAsiento = False
                'Volvemos el RS al sitio
                Rs.MoveNext
                Fecha = CDate(Rs.Fields(CampoFecha))
            Else
                'Inserto para que actalice             3: Opcion para INSERT INTO tmpactualizar
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
            End If
        End If
 
        
        'Si genero asiento
        If GeneraAsiento Then
            If MiCon Is Nothing Then Set MiCon = New Contadores
            MiCon.ConseguirContador "0", Fecha <= vParam.fechafin, True
                        
            'Genero la cabecera
            InsertarEnAsientosDesdeTemp Rs, MiCon, 0, NumLinea, NumVtos
            
            NumLinea = 1
            ImpBanco = 0
            'Reservo la primera linea para el banco
            If GastosTransferencia <> 0 Then
                NumLinea = 2
                If Not GastosTransDescontados Then
                    If Cobros Then
                        ImpBanco = -GastosTransferencia
                    Else
                        ImpBanco = -GastosTransferencia
                    End If
                End If
            End If
            
            riesgo = 0
        End If
        
    
        'Para el cobro /pago  que tendremos en la fila actual del recordset
        impo = Rs!Importe
        InsertarEnAsientosDesdeTemp Rs, MiCon, 1, NumLinea, Rs!NumVtos
    
        If Cobros Then
            riesgo = riesgo + Rs!Gastos
        Else
            riesgo = 0
        End If
        ImpBanco = ImpBanco + Rs!Importe
        NumLinea = NumLinea + 1
        
        'Si tengo que generar la contrapartida
        If ContraPartidaPorLinea Then
            NumVtos = Rs!NumVtos
            InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
            NumLinea = NumLinea + 1
            ImpBanco = 0
            riesgo = 0
        Else
            NumVtos = NumVtos + Rs!NumVtos
        End If
        
        'Nos movemos
        Rs.MoveNext
        
        
        If Rs.EOF Then
            
            If Not ContraPartidaPorLinea Then
                
                'Era la ultima linea.
                Rs.MovePrevious
                
                LineaUltima = NumLinea
                
                'Cierro el apunte, del banco
                'Si fuera una transferenicia con gastos descontados, me he dejado el numlinea=1
                'si no, no hago nada
                If GastosTransferencia <> 0 Then
                    If Not GastosTransDescontados Then NumLinea = 1
                End If
                impo = ImpBanco
                InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
    
                If GastosTransferencia <> 0 Then
                    If Not GastosTransDescontados Then
                        NumLinea = LineaUltima + 1
                
                        impo = GastosTransferencia
                        
                        InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, NumVtos
                    End If
                End If
    
    
                'CIERRO EL APUNTE
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
                
                'Y vuelvo a ponerlo ande tocaba. Para que se salga del bucle
                Rs.MoveNext
                
            Else
                'Cada linea de asiento tiene su banco
                'Faltara insertarlo en tmpactualizar
                InsertarEnAsientosDesdeTemp Rs, MiCon, 3, NumLinea, NumVtos
            End If
        End If
    Wend
    Rs.Close
    
    
    
    
    'Si es cobro por efectivo y me indica que lo llevo al banco
    'entoces generare dos lineas mas que sera el total del banco contra el total
    'la cuenta del banco donde lo llevamos
    ' EN ImporteGastosTarjeta llevo el banco donde llevo la pasta en efectivo
    
    If Cobros And Tipo = 0 And ImporteGastosTarjeta_ > 0 Then
        'Cuadramos el apunte.
        'Para ello guardamos unos valores que reestableceremos despues
        ImporteGastosTarjeta_ = CCur(Sql)
        UnAsientoPorCuenta = vParam.abononeg
        vParam.abononeg = False
        
        On Error Resume Next    'Por no llevarme todas las variables otra funcion
        AgrupaCuenta = False
        
        
        cad = " select sum(imponible-total),'" & CStr(ImporteGastosTarjeta_) & "' as cliprov, 'LLEV.BANCO||' as cliente"
        cad = cad & " from tmpfaclin WHERE codusu = " & vUsu.Codigo & " group by codusu"
        Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Err.Number = 0 Then
            If Not Rs.EOF Then
                impo = Rs.Fields(0)
                NumLinea = NumLinea + 1
                InsertarEnAsientosDesdeTemp Rs, MiCon, 1, NumLinea, 1
                
                If Err.Number = 0 Then
                
                    NumLinea = NumLinea + 1
                    InsertarEnAsientosDesdeTemp Rs, MiCon, 2, NumLinea, 1
                    
                    If Err.Number <> 0 Then
                        MuestraError Err.Number, "Cuadre llevar banco"
                        AgrupaCuenta = True
                    End If
                Else
                    'Error
                    AgrupaCuenta = True
                End If
            End If
            Rs.Close
        Else
            AgrupaCuenta = True
        End If
        

        ImporteGastosTarjeta_ = 0
        vParam.abononeg = UnAsientoPorCuenta
        On Error GoTo 0
        If AgrupaCuenta Then
            'Se ha producido un error
            'Provoco uno para que no siga la contabilizacion
            impo = 1 / 0
        End If
    End If
    
    Set Rs = Nothing
    
    
End Sub






'----------------------------------------------------------
'   A partir de la tabla tmp
'   Se que cuentas hay y los vencimientos.Por lo tanto, comprobare
'   que si la fechas estan fuera de ejercicios o de ambito
'   y si hay cuentas bloquedas
Private Function ComprobarCuentasBloquedasYFechasVencimientos() As Boolean
Dim cad As String

    ComprobarCuentasBloquedasYFechasVencimientos = False
    On Error GoTo EComprobarCuentasBloquedasYFechasVencimientos
    Set Rs = New ADODB.Recordset
    

    cad = "select codmacta,nommacta,numfac,fecha,fecbloq,cliente from tmpfaclin,cuentas where codusu=" & vUsu.Codigo & " and cta=codmacta and not (fecbloq is null )"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not Rs.EOF
        If CDate(Rs!NumFac) > Rs!FecBloq Then cad = cad & Rs!codmacta & "    " & Rs!FecBloq & "     " & Format(Rs!NumFac, "dd/mm/yyyy") & Space(15) & RecuperaValor(Rs!Cliente, 1) & vbCrLf
        Rs.MoveNext
    Wend
    Rs.Close


    If cad <> "" Then
        cad = vbCrLf & String(90, "-") & vbCrLf & cad
        cad = "Cta           Fec. Bloq            Fecha contab         Factura" & cad
        cad = "Cuentas bloqueadas: " & vbCrLf & vbCrLf & vbCrLf & cad
        MsgBox cad, vbExclamation
    Else
        ComprobarCuentasBloquedasYFechasVencimientos = True
    End If
EComprobarCuentasBloquedasYFechasVencimientos:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function








Private Function AdelanteConLaTransferencia() As Boolean
Dim C As String

    AdelanteConLaTransferencia = False
    Sql = ""
    
    If Rs!Situacion = "A" Then
        Select Case TipoTrans
            Case 0, 1
                Sql = "Transferencia abierta. Sin llevar al banco."
            Case 2
                Sql = "Pago domiciliado abierto. Sin llevar al banco."
            Case 3
                Sql = "Confirming abierto. Sin llevar al banco."
        End Select
    End If
    
    'Ya contabilizada
    If Rs!Situacion = "Q" Then
        Select Case TipoTrans
            Case 0, 1
                Sql = "Transferencia abonada."
            Case 2
                Sql = "Pago domiciliado abonado."
            Case 3
                Sql = "Confirming abonado."
        End Select
    End If
    If Sql <> "" Then Exit Function
    
    
    AdelanteConLaTransferencia = Sql = ""
    
End Function

Private Function SugerirCodigoSiguienteTransferencia() As String
    
    Sql = "Select Max(codigo) from stransfer"
    If SubTipo = 0 Then Sql = Sql & "cob"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, , , adCmdText
    Sql = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            Sql = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    SugerirCodigoSiguienteTransferencia = Sql
End Function




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    Limpiar Me
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    
    
    'Cago los iconos
    CargaImagenesAyudas Me.Image1, 2
    
    FrameContabilRem2.Visible = False
    
    Select Case Opcion
    Case 8, 22, 23
        'Utilizare el mismo FRAM para
        '   8.- Contabilizar / Abono remesa
        '   22- Cancelacion cliente
        '   23- Confirmacion remesa
        '  TANTO DE EFECTOS como de talones pagares
        FrameContabilRem2.Visible = True
        
        Caption = "Transferencia"
        chkAgrupaCancelacion.Visible = False
        
        Sql = "Abono transferencia"
        CuentasCC = "Contabilizar"
        
        Label5(2).Caption = Sql
        cmdContabRemesa.Caption = CuentasCC
        
        If Opcion = 8 Then
            If Cobros Then
                Me.Caption = "Abono transferencia"
            Else
                Select Case TipoTrans
                    Case 0, 1
                        Me.Caption = "Contabilizaci�n Transferencia"
                        Label5(2).Caption = "Transferencia : "
                    Case 2
                        Me.Caption = "Contabilizaci�n Pago Domiciliado"
                        Label5(2).Caption = "Pago Domiciliado : "
                    Case 3
                        Me.Caption = "Contabilizaci�n Confirming"
                        Label5(2).Caption = "Confirming : "
                End Select
            End If
            Label5(2).Caption = Label5(2).Caption & RecuperaValor(NumeroDocumento, 1) & "/" & RecuperaValor(NumeroDocumento, 2) & vbCrLf & " Banco : " & RecuperaValor(NumeroDocumento, 4) & vbCrLf & " Importe: " & RecuperaValor(NumeroDocumento, 5)
        End If
        
        CuentasCC = ""
        'Los gastos solo van en la contabilizacion
        Label3(0).Visible = Opcion = 8
        txtImporte(0).Visible = Opcion = 8
        
        
        W = FrameContabilRem2.Width
        H = FrameContabilRem2.Height
    End Select
    
    
    Me.Height = H + 360
    Me.Width = W + 90
    
    H = Opcion
    Me.cmdCancelar(H).Cancel = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumeroDocumento = "" 'Para reestrablecerlo siempre
End Sub



Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(10).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    Sql = RecuperaValor(CadenaSeleccion, 1)
End Sub


Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    Image1(10).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
    If Text1(Index).Text <> "" Then PonerFoco Text1(Index)
End Sub


Private Sub PonerFoco(ByRef o As Object)
    On Error Resume Next
    o.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ObtenerFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Sub KEYpress(ByRef KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
        Text1(Index).Text = ""
        PonerFoco Text1(Index)
    End If
    
End Sub



Private Sub txtImporte_GotFocus(Index As Integer)
    With txtImporte(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
Dim Valor
 
    txtImporte(Index).Text = Trim(txtImporte(Index))
    If txtImporte(Index).Text = "" Then Exit Sub
    

    If Not EsNumerico(txtImporte(Index).Text) Then
        txtImporte(Index).Text = ""
        Exit Sub
    End If

    If InStr(1, txtImporte(Index).Text, ",") > 0 Then
        Valor = ImporteFormateado(txtImporte(Index).Text)
    Else
        Valor = CCur(TransformaPuntosComas(txtImporte(Index).Text))
    End If
    txtImporte(Index).Text = Format(Valor, FormatoImporte)
        
End Sub

Private Sub PonerCuentasCC()

    CuentasCC = ""
    If vParam.autocoste Then
        Sql = "Select * from parametros"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        CuentasCC = "|" & miRsAux!grupogto & "|" & miRsAux!grupovta & "|"
        miRsAux.Close
        Set miRsAux = Nothing
    End If
End Sub





Private Function InsertarEnAsientosDesdeTemp(ByRef RS1 As ADODB.Recordset, ByRef m As Contadores, Cabecera As Byte, ByRef NumLine As Integer, NumVtos As Integer, Optional VienedeGastos As Boolean)
Dim Sql As String
Dim Ampliacion As String
Dim Debe As Boolean
Dim Conce As Integer
Dim TipoAmpliacion As Integer
Dim PonerContrPartida As Boolean
Dim Aux As String
Dim ImporteInterno As Currency
Dim TipForpa As Byte
    
    
    ImporteInterno = impo
    
    'LaUltimaAmpliacion  --> Servira pq si en parametros esta marcado un apunte por movimiento, o solo metemos
    '                        un unico pagao/cobro, repetiremos numdocum, textoampliacion
    
    'El diario

    FechaAsiento = Fecha
    If Cobros Then
        Ampliacion = vp.diaricli
    Else
        Ampliacion = vp.diaripro
    End If
    
    If Cabecera = 0 Then
        'La cabecera
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
        Sql = Sql & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador
        Sql = Sql & ",  '"
        Sql = Sql & "Generado desde Tesorer�a el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre
        
        Sql = Sql & "',"
        If Cobros Then
            Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Transf.Abonos'"
        Else
            Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Transf.Pagos'"
        End If

        
        Sql = Sql & ")"
        NumLine = 0
     
    Else
        If Cabecera < 3 Then
            'Lineas de apuntes o cabecera.
            'Comparten el principio
             Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
             If Cobros Then
                Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfaccl, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
             Else
                Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfacpr, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
             End If
             Sql = Sql & "VALUES (" & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador & "," & NumLine & ",'"
             
             '1:  Asiento para el VTO
             If Cabecera = 1 Then
                 'codmacta
                 'Si agrupa la cuenta entonces
                 Sql = Sql & RS1!cliprov & "','"
                 
                 
                 'numdocum: la factura
                 If NumVtos > 1 Then
                    Ampliacion = "Vtos: " & NumVtos
                 Else
                    Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2))
                 End If
                 Sql = Sql & Ampliacion & "',"
                
                
                 'Veamos si va al debe, al haber, si ponemos concepto debe / haber etc eyc
                 If Cobros Then
                    'CLIENTES
                    If ImporteInterno < 0 Then
                       If vParam.abononeg Then
                           Debe = False
                       Else
                           'Va al debe pero cambiado de signo
                           Debe = True
                           ImporteInterno = Abs(ImporteInterno)
                       End If
                    Else
                       Debe = False
                    End If
                    If Debe Then
                        Conce = vp.condecli
                        TipoAmpliacion = vp.ampdecli
                        PonerContrPartida = vp.ctrdecli = 1
                    Else
                        Conce = vp.conhacli
                        TipoAmpliacion = vp.amphacli
                        PonerContrPartida = vp.ctrhacli = 1
                    End If
                 
                 
                 Else
                    'PROVEEDORES
                    If ImporteInterno < 0 Then
                       If vParam.abononeg Then
                           Debe = True
                       Else
                           'Va al debe pero cambiado de signo
                           Debe = False
                           ImporteInterno = Abs(ImporteInterno)
                       End If
                    Else
                       Debe = True
                    End If
                    If Debe Then
                        Conce = vp.condepro
                        TipoAmpliacion = vp.ampdepro
                        PonerContrPartida = vp.ctrdepro = 1
                    Else
                        Conce = vp.conhapro
                        TipoAmpliacion = vp.amphapro
                        PonerContrPartida = vp.ctrhapro = 1
                    End If
                     
                 End If
                
                
                 Sql = Sql & Conce & ","
                 
                 'AMPLIACION
                 Ampliacion = ""
                


                Select Case TipoAmpliacion
                Case 0, 1
                   If TipoAmpliacion = 1 Then Ampliacion = Ampliacion & vp.siglas & " "
                   Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                
                Case 2
                
                   Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 3)
                
                Case 3
                    'NUEVA AMPLIC
                    Ampliacion = DescripcionTransferencia
                Case 4
                    'Estamos en la amplicacion del cliente. Es una tonteria tener esta opcion marcada, pero bien
                    Ampliacion = RecuperaValor(vTextos, 3)
                    Ampliacion = Mid(Ampliacion, InStr(1, Ampliacion, "-") + 1)
                Case 5
                    'Si hubiera que especificar mas el documento
'                    If Tipo = vbTalon Then
'                        AUX = "TAL N�"
'                    Else
'                        AUX = "PAG N�"
'                    End If
'
                
                    If Cobros Then
                        'Veo la el camporefencia de ese talon
                        'Antes cogiamos numero fra
                        'ahora contrapar
                        Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)  'Num tal pag
                        If False Then
                            
                            Ampliacion = "numserie = '" & RecuperaValor(RS1!Cliente, 1) & "' AND RecuperaValor(RS1!Cliente, 2)"
                            Ampliacion = Ampliacion & " AND numorden = " & RecuperaValor(RS1!Cliente, 4) & " AND fecfactu "
                            Ampliacion = DevuelveDesdeBD("reftalonpag", "cobros", Ampliacion, Format(RecuperaValor(RS1!Cliente, 3), FormatoFecha), "F")
                            
                        Else
                            'Es numero tal pag + ctrpar
                            DescripcionTransferencia = RecuperaValor(vTextos, 2)
                            DescripcionTransferencia = Mid(DescripcionTransferencia, InStr(1, DescripcionTransferencia, "-") + 1)
                            Ampliacion = Ampliacion & " " & DescripcionTransferencia
                            DescripcionTransferencia = ""
                        End If
                        If Ampliacion = "" Then
                            Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                        Else
                            Ampliacion = " N�Doc: " & Ampliacion
                        End If
                    Else
                        If NumeroTalonPagere = "" Then
                            Ampliacion = ""
                        Else
                            'Cta banco
                            Ampliacion = RecuperaValor(vTextos, 2)
                            Ampliacion = Mid(Ampliacion, InStr(1, Ampliacion, "-") + 1)
                            'Numero tal/pag
                        
                            Ampliacion = NumeroTalonPagere & " " & Ampliacion
                        
                        End If
                        
                        If Ampliacion = "" Then
                            Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                        Else
                            Ampliacion = "N�Doc: " & Ampliacion
                        End If
                    End If
                    
                End Select
                   
                If NumVtos > 1 Then
                    'TIENE MAS DE UN VTO. No puedo ponerlo en la ampliacion
                    Ampliacion = "Vtos: " & NumVtos
                End If
                
                 'Le concatenamos el texto del concepto para el asiento -ampliacion
                 Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce)) & " "
                 'Para la ampliacion de n�tal + ctrapar NO pongo la ampliacion del concepto
                 If TipoAmpliacion = 5 Then Aux = ""
                 Ampliacion = Aux & Ampliacion
                 If Len(Ampliacion) > 30 Then Ampliacion = Mid(Ampliacion, 1, 30)
                
                 Sql = Sql & "'" & DevNombreSQL(Ampliacion) & "',"
                 
                 
                 If Debe Then
                    Sql = Sql & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                 Else
                    Sql = Sql & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                 End If
             
                'CENTRO DE COSTE
                Sql = Sql & "NULL,"
                
                'SI pone contrapardida
                If PonerContrPartida Then
                   Sql = Sql & "'" & RecuperaValor(NumeroDocumento, 4) & "',"
                Else
                   Sql = Sql & "NULL,"
                End If
            
             
            Else
                    '----------------------------------------------------
                    'Cierre del asiento con el total contra banco o caja
                    '----------------------------------------------------
                    'codmacta
                    Sql = Sql & RecuperaValor(NumeroDocumento, 4) & "','"
                     
  
                    PonerContrPartida = False
                    If NumVtos = 1 Then
                        PonerContrPartida = True
                    Else
                        PonerContrPartida = False
                    End If
                       
                    If PonerContrPartida Then
                       Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2))
                    Else
                       
                       Ampliacion = ""
                    End If
                     
                    Sql = Sql & Ampliacion & "',"
                   
                    
                    If Cobros Then
                        '----------------------------------------------------------------------
                        If ImporteInterno < 0 Then
                           If vParam.abononeg Then
                               Debe = True
                           Else
                               'Va al debe pero cambiado de signo
                               Debe = False
                               ImporteInterno = Abs(ImporteInterno)
                           End If
                        Else
                           Debe = True
                        End If
                                   
                        
                        'COmo el banco o caja, siempre van al reves (Su abono es nuetro pago..)
                        If Not Debe Then
                            Conce = vp.condecli
                            TipoAmpliacion = vp.ampdecli
                        Else
                            Conce = vp.conhacli
                            TipoAmpliacion = vp.amphacli
                        End If
                        
                     Else
                        'PROVEEDORES
                        If ImporteInterno < 0 Then
                           If vParam.abononeg Then
                               Debe = False
                           Else
                               'Va al debe pero cambiado de signo
                               Debe = True
                               ImporteInterno = Abs(ImporteInterno)
                           End If
                        Else
                           Debe = False
                        End If
                        
                        If Not Debe Then
                            Conce = vp.condepro
                            TipoAmpliacion = vp.ampdepro
                        Else
                            Conce = vp.conhapro
                            TipoAmpliacion = vp.amphapro
                        End If
                     End If
                     
                        
                     
                     
                
                     Sql = Sql & Conce & ","
                     'AMPLIACION
                     'AMPLIACION
                     Ampliacion = ""
                     
                     'Si estoy contabilizando pag de UN unico proveedor entonces NumeroTalonPageretendra valor
                     If NumVtos > 1 And NumeroTalonPagere <> "" Then NumVtos = 1
                        
                     
                     If NumVtos = 1 Then
                    
                        Select Case TipoAmpliacion
                        Case 0, 1
                           If TipoAmpliacion = 1 Then Ampliacion = Ampliacion & vp.siglas & " "
                           Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                        
                        Case 2
                        
                           Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 3)
                        
                        Case 3
                            'NUEVA AMPLIC
                             Ampliacion = DescripcionTransferencia
                        Case 4, 5
                            'Nombre ctrpartida
                            Ampliacion = CStr(DBLet(RS1!cliprov, "T"))
                            Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Ampliacion, "T")
                            DescripcionTransferencia = Ampliacion
                            If Cobros Then
                                
                                'Veo la el camporefencia de ese talon
                                Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Ampliacion = "numserie = '" & RecuperaValor(RS1!Cliente, 1) & "' AND numfaccl = " & RecuperaValor(RS1!Cliente, 2)
                                Ampliacion = Ampliacion & " AND numorden = " & RecuperaValor(RS1!Cliente, 4) & " AND fecfactu "
                                Ampliacion = DevuelveDesdeBD("reftalonpag", "hlinapu", Ampliacion, Format(RecuperaValor(RS1!Cliente, 3), FormatoFecha), "F")
                                
                                If Ampliacion = "" Then
                                    Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Else
                                    Ampliacion = " N�Doc: " & Ampliacion
                                End If
                                Ampliacion = Ampliacion & " " & DescripcionTransferencia
     
                            Else
                                
                                Ampliacion = NumeroTalonPagere
                                If Ampliacion = "" Then
                                    Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Else
                                    Ampliacion = "N�Doc: " & Ampliacion
                                End If
                            End If
                          
                            Ampliacion = Ampliacion & " " & DescripcionTransferencia
                            DescripcionTransferencia = ""
                          
                          
                        End Select
                    Else
                        'Ma de un VTO.  Si no
                        If vp.tipoformapago = vbTransferencia Then
                            'SI es transferencia
                            'If TipoAmpliacion = 3 Then Ampliacion = DescripcionTransferencia
                            Ampliacion = DescripcionTransferencia
                        
                        End If
                    End If
                    
                     Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce))
                     Aux = Aux & " "
                     'Para la ampliacion de n�tal + ctrapar NO pongo la ampliacion del concepto
                     If TipoAmpliacion = 5 Then Aux = ""
                     Ampliacion = Trim(Aux & Ampliacion)
                     If Len(Ampliacion) > 30 Then Ampliacion = Mid(Ampliacion, 1, 30)
                    
                     Sql = Sql & "'" & DevNombreSQL(Ampliacion) & "',"
        
                         
                     If Debe Then
                        Sql = Sql & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                     Else
                        Sql = Sql & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                     End If
                 
                     'CENTRO DE COSTE
                     Sql = Sql & "NULL,"
                    
                     'SI pone contrapardida
                     If PonerContrPartida Then
                        Sql = Sql & "'" & RS1!cliprov & "',"
                     Else
                        Sql = Sql & "NULL,"
                     End If
                 
            End If
            
            'Trozo comun
            '------------------------
            'IdContab
            If Cobros Then
                Sql = Sql & "'COBROS',"
            Else
                Sql = Sql & "'PAGOS',"
            End If
            
            'Punteado
            Sql = Sql & "0,"
            
            If Cabecera = 1 And Mid(RS1!Cliente, 1, 3) <> "TRA" Then
            
                '--TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", RS!codforpa, "N")
                Select Case TipoTrans
                    Case 0, 1
                        TipForpa = vbTransferencia
                    Case 2, 3
                        TipForpa = vbConfirming
                End Select
            
                ' nuevos campos de la factura
                'numSerie , numfacpr, FecFactu, numorden, TipForpa, reftalonpag, bancotalonpag
                Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 2), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & ","
                Sql = Sql & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & "," & DBSet(TipForpa, "N") & ","
                
                Dim SqlBanco As String
                Dim RsBanco As ADODB.Recordset
                
                SqlBanco = "select reftalonpag, bancotalonpag from tmpcobros2 where codusu = " & vUsu.Codigo
                SqlBanco = SqlBanco & " and numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T")
                SqlBanco = SqlBanco & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "N")
                SqlBanco = SqlBanco & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F")
                SqlBanco = SqlBanco & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
                SqlBanco = SqlBanco & " and codmacta = " & DBSet(RS1!cliprov, "T")
        
                Set RsBanco = New ADODB.Recordset
                RsBanco.Open SqlBanco, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsBanco.EOF Then
                    Sql = Sql & DBSet(RsBanco.Fields(0), "T") & "," & DBSet(RsBanco.Fields(1), "T") & ")"
                Else
                    Sql = Sql & ValorNulo & "," & ValorNulo & ")"
                End If
                Set RsBanco = Nothing
                
            Else
                Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            End If
                 
             
        End If 'De cabecera menor que 3, es decir : 1y 2
    
    
    End If
    
    'Ejecutamos si:
    '   Cabecera=0 o 1
    '   Cabecera=2 y impo=0.  Esto sginifica que estamos desbloqueando el apunte e insertandolo para pasarlo a hco
    Debe = True
    If Cabecera = 3 Then Debe = False
    If Debe Then Conn.Execute Sql
    
    If Debe Then
        '++monica
'        If Cobros And Cabecera = 1 And Not VienedeGastos Then
'
'            Dim Situacion As Byte
'
'            Situacion = 1
'
'            SQL = "update cobros set impcobro = coalesce(impcobro,0) + " & DBSet(ImporteInterno, "N")
'            SQL = SQL & " ,fecultco = " & DBSet(FechaAsiento, "F")
'            SQL = SQL & ", situacion = " & DBSet(Situacion, "N")
'            SQL = SQL & " where numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "N")
'            SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
'
'            Conn.Execute SQL
'
'        ' en tmppendientes metemos la clave primaria de cobros_recibidos y el importe en letra
'                                                          'importe=nro factura,   codforpa=linea de cobros_realizados
'            SQL = "insert into tmppendientes (codusu,serie_cta,importe,fecha,numorden, observa) values ("
'            SQL = SQL & vUsu.Codigo & "," & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," 'numserie
'            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 2), "N") & "," 'numfactu
'            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & "," 'fecfactu
'            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & "," 'numorden
'            SQL = SQL & DBSet(EscribeImporteLetra(ImporteFormateado(CStr(ImporteInterno))), "T") & ") "
'
'            Conn.Execute SQL
'
'        End If
    
    End If
    
    
    
    
    '-------------------------------------------------------------------
    'Si es apunte de banco, y hay gastos
    If Cabecera = 2 Then
        'SOOOOLO COBROS
        If Cobros And riesgo > 0 Then
                     
             Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
             Sql = Sql & " timporteH,  ctacontr,codccost, idcontab, punteada) "
             Sql = Sql & "VALUES (" & vp.diaricli & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador & ","
             
             Ampliacion = DevuelveDesdeBD("ctaingreso", "bancos", "codmacta", RecuperaValor(NumeroDocumento, 4), "T")
             If Ampliacion = "" Then
                MsgBox "Cta ingreso bancario MAL configurada. Se utilizara la misma del banco", vbExclamation
                Ampliacion = RecuperaValor(NumeroDocumento, 4)
            End If
            'linea,numdocum,codconce  amconce
            For Conce = 1 To 2
                NumLine = NumLine + 1
                Aux = NumLine & ",'"
                If Conce = 1 Then
                    Aux = Aux & RecuperaValor(NumeroDocumento, 4)
                Else
                    Aux = Aux & Ampliacion
                End If
                Aux = Aux & "',''," & vp.condecli & ",'" & DevNombreSQL(DevuelveDesdeBD("nomconce", "conceptos", "codconce", vp.condecli)) & "',"
                If Conce = 1 Then
                    Aux = Aux & TransformaComasPuntos(CStr(riesgo)) & ",NULL"
                Else
                    Aux = Aux & "NULL," & TransformaComasPuntos(CStr(riesgo))
                End If
                If Conce = 2 Then
                    Aux = Aux & ",'" & RecuperaValor(NumeroDocumento, 4)
                Else
                    Aux = Aux & ",'" & Ampliacion
                End If
                Aux = Aux & "',"
                'CC
                If Conce = 1 Then
                    Aux = Aux & "NULL"
                Else
                    If vParam.autocoste Then
                        Ampliacion = DevuelveDesdeBD("codccost", "bancos", "codmacta", RecuperaValor(NumeroDocumento, 4), "T")
                        If Ampliacion = "" Then
                            Ampliacion = "NULL"
                        Else
                            Ampliacion = "'" & Ampliacion & "'"
                        End If
                    Else
                        'NO LLEVA ANALITICA
                        Ampliacion = "NULL"
                    End If
                    Aux = Aux & Ampliacion
                End If
                If Cobros Then
                    Aux = Aux & ",'COBROS',0)"
                Else
                    Aux = Aux & ",'PAGOS',0)"
                End If
                
                Aux = Sql & Aux
                Ejecuta Aux
            Next Conce
        End If
    End If
    
    
End Function


