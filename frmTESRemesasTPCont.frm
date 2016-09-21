VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESRemesasTPCont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmTESRemesasTPCont.frx":0000
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
Attribute VB_Name = "frmTESRemesasTPCont"
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
    Sql = "Select * from remesas where codigo =" & RecuperaValor(NumeroDocumento, 1)
    Sql = Sql & " AND anyo =" & RecuperaValor(NumeroDocumento, 2)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Rs.EOF Then
        MsgBox "Ninguna remesa con esos valores", vbExclamation
        Rs.Close
        Set Rs = Nothing
        Exit Sub

    End If
    
    'Tiene valor
    Sql = ""
    B = AdelanteConLaRemesa()
    ContabilizacionEspecialNorma19 = False
    If B Then
        'Si es norma19 y tiene le parametro de contabilizacion por fecha comprobaremos la fecha de los vtos
        If Opcion = 8 Then
        
            'Se podrian agrupar los IFs, pero asi de momento me entero mas
        
            'Para RECIBOS BANCARIOS SOLO
            If DBLet(Rs!Tiporem, "N") = 1 Then
                If vParamT.Norma19xFechaVto Then
                    If Not IsNull(Rs!Tipo) Then
                        If Rs!Tipo = 0 Then
                            'NORMA 19
                            'Contbiliza por fecha VTO
                            'Comprobaremos que toooodos estan en fecha ejercicio
                            Sql = ComprobacionFechasRemesaN19PorVto
                            If Sql <> "" Then Sql = "-Comprobando fechas remesas N19" & vbCrLf & Sql
                            
                            
                            If txtImporte(0).Text <> "" Then Sql = Sql & vbCrLf & "N19 no permite gastos bancario"
                            
                            
                            If Sql <> "" Then
                                B = False
                            Else
                                ContabilizacionEspecialNorma19 = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    End If

    If Not B Then
        If Sql = "" Then Sql = "Error y punto"
        Rs.Close
        Set Rs = Nothing
        MsgBox Sql, vbExclamation
        Exit Sub
    End If
    Sql = "Select cobros.codmacta,nomclien,fecbloq from cobros,cuentas where cobros.codmacta = cuentas.codmacta"
    Sql = Sql & " and  codrem =" & RecuperaValor(NumeroDocumento, 1)
    Sql = Sql & " AND anyorem =" & RecuperaValor(NumeroDocumento, 2)
    Sql = Sql & " AND fecbloq <='" & Format(Text1(10).Text, FormatoFecha) & "' GROUP BY 1"
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
        CC = CC & " la remesa: " & Rs!Codigo & " / " & Rs!Anyo & vbCrLf & vbCrLf
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
    If ContabilizacionEspecialNorma19 Then
        'Utiliza Morales
        'Es para contabilizar los recibos por fecha de vto
        
        B = ContabNorma19PorFechaVto(Rs!Codigo, Rs!Anyo, Sql)
    Else
        'Toooodas las demas opciones estan aqui
    
                                'Efecto(1),pagare(2),talon(3)
        B = ContabilizarRecordsetRemesa(Rs!Tiporem, DBLet(Rs!Tipo, "N") = 0, Rs!Codigo, Rs!Anyo, Sql, CDate(Text1(10).Text), Importe)
    
    End If
    
    'si se contabiliza entonces updateo y la pongo en
    'situacion Q. Contabilizada a falta de devueltos ,
    If B Then
        Conn.CommitTrans
        'AQUI updateamos el registro pq es una tabla myisam
        'y no debemos meterla en la transaccion
        
        HaHabidoCambios = True
        
        
        Sql = "UPDATE remesas SET"
        Sql = Sql & " situacion= 'Q'"
        Sql = Sql & " WHERE codigo=" & Rs!Codigo
        Sql = Sql & " and anyo=" & Rs!Anyo

        If Not Ejecuta(Sql) Then MsgBox "Error actualizando tabla remesa.", vbExclamation
        
        
        'Ahora actualizamos los registros que estan en tmpactualziar
        Screen.MousePointer = vbDefault
        'Cerramos
        Rs.Close
        Unload Me
        Exit Sub
    Else
        TirarAtrasTransaccion
    End If

    
    
    
    
    Rs.Close
    Set Rs = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Function AdelanteConLaRemesa() As Boolean
Dim C As String

    AdelanteConLaRemesa = False
    Sql = ""
    
    'Efectos eliminados
    If Rs!Situacion = "Z" Or Rs!Situacion = "Y" Then Sql = "Efectos eliminados"
    
    'abierta sin llevar a banco. Esto solo es valido para las de efectos
    If SubTipo = 1 Then
        If Rs!Situacion = "A" Then Sql = "Remesa abierta. Sin llevar al banco."
    
    End If
    'Ya contabilizada
    If Rs!Situacion = "Q" Then Sql = "Remesa abonada."
    
    If Sql <> "" Then Exit Function
    
    
    If Opcion = 8 Then
        'COntbilizar / abonar remesa
        '---------------------------------------------------------------------------
        If SubTipo = 1 Then
            'Febrero 2009
            'Ahora toooodas las remesas se hace lo mismmo
            ' De llevada a banco a cancelar cliente. De cancelar a abonar y de abonar a eliminar. NO
            'hay distinciones entre remesas. Para podrer abonar una remesa esta tiene que estar cancelada
            
        Else
            If Rs!Tiporem = 2 And vParamT.PagaresCtaPuente Then
                If Rs!Situacion <> "F" Then Sql = "La remesa NO puede abonarse. Falta cancelaci�n "
            End If
            
            If Rs!Tiporem = 3 And vParamT.TalonesCtaPuente Then
                If Rs!Situacion <> "F" Then Sql = "La remesa NO puede abonarse. Falta cancelaci�n "
            End If
        End If
        
            
    Else
       'Vamos a proceder al proceso de generacion cancelacion  /* CANCELACION */
       If SubTipo = 1 Then
       End If
       
       'Para elos tipos 1,2
       If Opcion = 22 Then
            'Cancelacion cliente
            'Para los efectos, tiene que estar generado soporte. Para talones/pagares no es obligado
            If SubTipo = 1 Then
                If Rs!Situacion <> "B" Then Sql = "Para cancelar la remesa deberia esta en situaci�n 'Soporte generado'"
            Else
                If Rs!Situacion = "F" Then Sql = "Remesa YA cancelada"
            End If
        Else
            'Febrero 2009
            'No hay confirmacion
            Sql = "Opci�n de confirmacion NO es v�lida"
            'Confirmacion
            'If Rs!situacion <> "F" Then SQL = "Para confirmar la remesa esta deberia estar 'Cancelacion cliente'"
       End If
       
       
       'Si hasta aqui esta bien:
       'Compruebo que tiene configurado en parametros
       If Sql = "" Then
            'Comprobamos si esta bien configurada
            '
            If SubTipo = 1 Then
                If Opcion = 22 Then
                    'SQL = "4310"
                    Sql = "RemesaCancelacion"
                Else
                    Sql = "RemesaConfirmacion"
                End If
                Sql = DevuelveDesdeBD(Sql, "paramtesor", "codigo", "1")
                If Sql = "" Then
                    Sql = "Falta configurar par�metros cuentas confirmaci�n/cancelaci�n remesa. "
                Else
                    'OK. Esta configurado
                    Sql = ""
                End If
                    
            Else
                'talones pagares
                'Veremos si esta configurado(y bien configurado) para el proceso
                If Rs!Tiporem = 2 Then
                    'Pagare
                    C = "contapagarepte"
                ElseIf Rs!Tiporem = 3 Then
                    'Talones
                    C = "contatalonpte"
                Else
                    'NO DEBIA HABERSE METIDO AQUI
                    C = ""
                    
                End If
                If C = "" Then
                    Sql = "Error validando tipo de remesa"
                    
                Else
                    C = DevuelveDesdeBD(C, "paramtesor", "codigo", 1)
                    If C = "" Then C = "0"
                    If Val(C) = 0 Then
                        Sql = "Falta configurar la aplicacion para las remesas de talones / pagares"
                    Else
                        Sql = ""
                    End If
                End If
            End If
       End If
    End If
    AdelanteConLaRemesa = Sql = ""
    
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
        
        Caption = "Remesas"
        If SubTipo = 1 Then
            Caption = Caption & " EFECTOS"
        Else
            Caption = Caption & " talones/pagar�s"
        End If
        chkAgrupaCancelacion.Visible = False
        
        If Opcion = 8 Then
            Sql = "Abono remesa"
            CuentasCC = "Contabilizar"
        Else
        
            If Opcion = 22 Then
            
                Sql = DevuelveDesdeBD("RemesaCancelacion", "paramtesor", "codigo", "1", "N")
                chkAgrupaCancelacion.Visible = Len(Sql) = vEmpresa.DigitosUltimoNivel
                Sql = "Cancelacion cliente"
                CuentasCC = "Can. cliente"
            Else
                Sql = "Confirmacion remesa"
                CuentasCC = "Confirmar"
            End If
            
        End If
        Label5(2).Caption = Sql
        cmdContabRemesa.Caption = CuentasCC
        
        If Opcion = 8 Then
            Me.Caption = "Abono remesa"
            Label5(2).Caption = "Remesa : " & RecuperaValor(NumeroDocumento, 1) & "/" & RecuperaValor(NumeroDocumento, 2) & vbCrLf & " Banco : " & RecuperaValor(NumeroDocumento, 4) & vbCrLf & " Importe: " & RecuperaValor(NumeroDocumento, 5)
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

Private Sub EliminarEnRecepcionDocumentos()
Dim CtaPte As Boolean
Dim J As Integer
Dim CualesEliminar As String
On Error GoTo EEliminarEnRecepcionDocumentos

    'Comprobaremos si hay datos
    
        'Si no lleva cuenta puente, no hace falta que este contabilizada
        'Es decir. Solo mirare contabilizados si llevo ctapuente
        CuentasCC = ""
        CualesEliminar = ""
        J = 0
        For I = 0 To 1
            ' contatalonpte
            Sql = "pagarecta"
            If I = 1 Then Sql = "contatalonpte"
            CtaPte = (DevuelveDesdeBD(Sql, "paramtesor", "codigo", "1") = "1")
            
            'Repetiremos el proceso dos veces
            Sql = "Select * from scarecepdoc where fechavto<='" & Format(Text1(17).Text, FormatoFecha) & "'"
            Sql = Sql & " AND   talon = " & I
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                    'Si lleva cta puente habra que ver si esta contbilizada
                    J = 0
                    If CtaPte Then
                        If Val(Rs!Contabilizada) = 0 Then
                            'Veo si tiene lineas. S
                            Sql = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(Rs!Codigo))
                            If Sql = "" Then Sql = "0"
                            If Val(Sql) > 0 Then
                                CuentasCC = CuentasCC & Rs!Codigo & " - No contabilizada" & vbCrLf
                                J = 1
                            End If
                        End If
                    End If
                    If J = 0 Then
                        'Si va benee
                        If Val(DBLet(Rs!llevadobanco, "N")) = 0 Then
                            Sql = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(Rs!Codigo))
                            If Sql = "" Then Sql = "0"
                            If Val(Sql) > 0 Then
                                CuentasCC = CuentasCC & Rs!Codigo & " - Sin llevar a banco" & vbCrLf
                                J = 1
                            End If
                    
                        End If
                    End If
                    'Esta la borraremos
                    If J = 0 Then CualesEliminar = CualesEliminar & ", " & Rs!Codigo
                    
                    Rs.MoveNext
            Wend
            Rs.Close
            
            
            
        Next I
        
        

        
        If CualesEliminar = "" Then
            'No borraremos ninguna
            If CuentasCC <> "" Then
                CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
                MsgBox CuentasCC, vbExclamation
                
            End If
            Exit Sub
        End If
            
        
        
        'Si k hay para borrar
        CualesEliminar = Mid(CualesEliminar, 2)
        J = 1
        Sql = "X"
        Do
            I = InStr(J, CualesEliminar, ",")
            If I > 0 Then
                J = I + 1
                Sql = Sql & "X"
            End If
        Loop Until I = 0
        
        Sql = "Va a eliminar " & Len(Sql) & " registros de la recepcion de documentos." & vbCrLf & vbCrLf & vbCrLf
        If CuentasCC <> "" Then CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
        Sql = Sql & vbCrLf & CuentasCC
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
            Sql = "DELETE from slirecepdoc where id in (" & CualesEliminar & ")"
            Conn.Execute Sql
            
            Sql = "DELETE from scarecepdoc where codigo in (" & CualesEliminar & ")"
            Conn.Execute Sql
    
        End If

    Exit Sub
EEliminarEnRecepcionDocumentos:
    MuestraError Err.Number, Err.Description
End Sub




Private Function ComprobacionFechasRemesaN19PorVto() As String
Dim Aux As String

    ComprobacionFechasRemesaN19PorVto = ""
    Aux = "anyorem = " & Rs!Anyo & " AND codrem "
    Aux = DevuelveDesdeBD("min(fecvenci)", "cobros", Aux, Rs!Codigo)
    If Aux = "" Then
        ComprobacionFechasRemesaN19PorVto = "Error fechas vto"
    Else
        If CDate(Aux) < vParam.fechaini Then
            ComprobacionFechasRemesaN19PorVto = "Vtos con fecha menor que inicio de ejercicio"
        End If
    End If
    If ComprobacionFechasRemesaN19PorVto <> "" Then Exit Function
    
    ComprobacionFechasRemesaN19PorVto = ""
    Aux = "anyorem = " & Rs!Anyo & " AND codrem "
    Aux = DevuelveDesdeBD("max(fecvenci)", "cobros", Aux, Rs!Codigo)
    If Aux = "" Then
        ComprobacionFechasRemesaN19PorVto = "Error fechas vto"
        Exit Function
    End If
    If CDate(Aux) > DateAdd("yyyy", 1, vParam.fechafin) Then ComprobacionFechasRemesaN19PorVto = "Vtos con fecha mayor que fin de ejercicio"
    
    
    
End Function



