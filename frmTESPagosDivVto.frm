VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESPagosDivVto 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmTESPagosDivVto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDividVto 
      Height          =   3855
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtcodigo 
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
         Index           =   3
         Left            =   2160
         TabIndex        =   3
         Top             =   2700
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   2
         Left            =   2190
         TabIndex        =   2
         Top             =   2280
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2190
         TabIndex        =   0
         Tag             =   "Nº asiento|N|S|0||hcabapu|numasien|####0|S|"
         Top             =   1350
         Width           =   1365
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   2190
         TabIndex        =   1
         Top             =   1800
         Width           =   1365
      End
      Begin VB.CommandButton cmdDivVto 
         Caption         =   "Aceptar"
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
         Left            =   3000
         TabIndex        =   4
         Top             =   3300
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
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
         Index           =   27
         Left            =   4200
         TabIndex        =   5
         Top             =   3300
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Días resto Vtos."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   450
         TabIndex        =   13
         Top             =   2760
         Width           =   1680
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmTESPagosDivVto.frx":000C
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha 1er Vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   450
         TabIndex        =   12
         Top             =   2280
         Width           =   1410
      End
      Begin VB.Label Label4 
         Caption         =   "Nº Vencimientos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   450
         TabIndex        =   11
         Top             =   1350
         Width           =   1650
      End
      Begin VB.Label Label4 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   450
         TabIndex        =   10
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "euros"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   62
         Left            =   3660
         TabIndex        =   9
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   57
         Left            =   240
         TabIndex        =   8
         Top             =   660
         Width           =   5040
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   56
         Left            =   240
         TabIndex        =   7
         Top             =   330
         Width           =   5040
      End
   End
End
Attribute VB_Name = "frmTESPagosDivVto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Public Opcion As Byte
    '27.-  Divide el vencimiento en dos vtos a partir del importe introducido en el text
    
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1

Dim Sql As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim i As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Dim ParaElLog As String

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 20 Or Index = 23 Or Index >= 26 Then
        CadenaDesdeOtroForm = "" 'Por si acaso. Tiene que devolve "" para que no haga nada
    End If
    Unload Me
End Sub



Private Sub cmdDivVto_Click()
Dim Im As Currency
Dim vImpvto As Currency
Dim vVtos As Integer
Dim vTotal As Currency
Dim J As Integer
Dim k As Integer
Dim ImportePagado As Currency
Dim vFecVenci As Date
Dim FecVenci As Date

Dim Dias As Integer

    On Error GoTo ecmdDivVto


    'Dividira el vto en dos. En uno dejara el importe que solicita y en el otro el resto
    'Los gastos s quedarian en uno asi como el cobrado si diera lugar
    
    ' controles
    
    
    RC = RecuperaValor(CadenaDesdeOtroForm, 3)
    Importe = CCur(RC)
        
    vImpvto = 0
    vVtos = 0
    If txtCodigo(1).Text <> "" Then vImpvto = ImporteSinFormato(ComprobarCero(txtCodigo(1).Text))
    If txtCodigo(0).Text <> "" Then vVtos = CInt(ComprobarCero(txtCodigo(0).Text))
        
    If vImpvto = 0 And vVtos = 0 Then
        MsgBox "Debe introducir el importe o el nro de vencimientos o ambos. Revise.", vbExclamation
        PonFoco txtCodigo(0)
        Exit Sub
    End If
    
    ' debe introducir la fecha del primer vto, viene cargada
    If txtCodigo(2).Text = "" Then
        MsgBox "Debe introducir la fecha del primer vencimiento", vbExclamation
        PonFoco txtCodigo(2)
        Exit Sub
    End If
    
    
    If txtCodigo(3).Text = "" Then
        If MsgBox("No ha puesto valor en el campo de días de resto de vencimientos. " & vbCrLf & vbCrLf & "¿ Desea continuar ?" & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    
    
    ' me ponen importe
    If vImpvto <> 0 Then
        If Importe < vImpvto Then
            MsgBox "El importe del vencimiento es inferior del importe a dividir. Revise", vbExclamation
            PonFoco txtCodigo(0)
            Exit Sub
        End If
        ' me ponen nro de vtos
        If vVtos <> 0 Then
            If Importe - Round(vImpvto * (vVtos - 1), 2) < 0 Then
                MsgBox "Es imposible dividir el vencimiento en " & vVtos & " vencimientos de " & Format(vImpvto, "###,###,##0.00") & " euros.", vbExclamation
                PonFoco txtCodigo(0)
                Exit Sub
            End If
            If vVtos = 1 And vImpvto <> Importe Then
                MsgBox "No podemos dejar el vencimiento con menos importe del original. Revise.", vbExclamation
                PonFoco txtCodigo(0)
                Exit Sub
            End If
        End If
    End If
    
    If vImpvto = 0 Then
        vImpvto = Round(Importe / vVtos, 2)
    End If
    
    If vVtos = 0 Then
        vVtos = Round(Importe / vImpvto, 0)
    End If
    
    Conn.BeginTrans
    
    Sql = ""
    If Sql = "" Then
        Set Rs = New ADODB.Recordset
        
        
        'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        i = -1
        RC = "Select max(numorden) from pagos WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        Rs.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs.EOF Then
            Sql = "Error. Vencimiento NO encontrado: " & CadenaDesdeOtroForm
        Else
            i = Rs.Fields(0) '+ 1
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
        PonFoco txtCodigo(1)
        Exit Sub
        
    Else
        Sql = "¿Desea desdoblar el vencimiento en los indicados?" 'uno de : " & Im & " euros?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    Dias = txtCodigo(3).Text

    
    FecVenci = CDate(txtCodigo(2))
    vFecVenci = FecVenci
    'OK.  a desdoblar
    vTotal = 0
    k = i + 1
    For J = 1 To vVtos - 1
    
        vTotal = vTotal + vImpvto
    
        vFecVenci = DateAdd("d", DBLet(Dias, "N"), vFecVenci)
        
    
        Sql = "INSERT INTO pagos (numorden,impefect,fecultpa,imppagad,emitdocum,"
        Sql = Sql & "numserie,numfactu,fecfactu,codmacta,codforpa,fecefect,ctabanc1,"
        Sql = Sql & "text1csb,text2csb,"
        Sql = Sql & "observa,nomprove,domprove,pobprove,cpprove,proprove,codpais,nifprove,iban,codusu) "
        'Valores
        Sql = Sql & " SELECT " & k & "," & TransformaComasPuntos(CStr(vImpvto)) & ",NULL,NULL,0,"
        Sql = Sql & "numserie,numfactu,fecfactu,codmacta,codforpa,"
        Sql = Sql & DBSet(vFecVenci, "F") & ","
        Sql = Sql & "ctabanc1,text1csb,text2csb,"
        'text83csb`,
        Sql = Sql & "observa,nomprove,domprove,pobprove,cpprove,proprove,codpais,nifprove,iban "
        Sql = Sql & "," & DBSet(vUsu.Id, "N")
        Sql = Sql & " FROM pagos WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        Sql = Sql & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
'        Sql = Sql & " and codmacta = "
    
        Conn.Execute Sql
    
        k = k + 1
    
    Next J
    
    
    ' actualizamos el primer vencimiento
    vTotal = vTotal + vImpvto
        
    Sql = "update pagos set impefect = coalesce(imppagad,0) + " & DBSet(vImpvto, "N")
    Sql = Sql & ", fecefect = " & DBSet(FecVenci, "F")
    
    Sql = Sql & " WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
    Sql = Sql & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
    
    Conn.Execute Sql
    
    ' en el ultimo dejamos la diferencia
    If vTotal <> Importe Then
        Sql = "update pagos set impefect = impefect + " & DBSet(Importe - vTotal, "N")
        
        Sql = Sql & " WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        Sql = Sql & " AND numorden = " & DBSet(k - 1, "N")
        
        Conn.Execute Sql
    End If
    
    'Insertamos el LOG
    ParaElLog = "Dividir Vto.Fra.: " & Me.Label4(57).Caption & vbCrLf
    ParaElLog = ParaElLog & "Proveedor         : " & Me.Label4(56).Caption & vbCrLf
    ParaElLog = ParaElLog & "Nro.Vencimientos: " & txtCodigo(0).Text & vbCrLf
    ParaElLog = ParaElLog & "Importe Vto     : " & txtCodigo(1).Text & vbCrLf
    ParaElLog = ParaElLog & "Fecha primer Vto: " & txtCodigo(2).Text & vbCrLf
    ParaElLog = ParaElLog & "Día Resto Vtos  : " & txtCodigo(3).Text & vbCrLf
    
    vLog.Insertar 1, vUsu, ParaElLog
    ParaElLog = ""
    
    
ecmdDivVto:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Dividir vencimientos", Err.Description
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & k & "|"
        MsgBox "Proceso realizado correctamente", vbExclamation
        Unload Me
    End If
    
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmppal.Icon
    
    'Limpiamos el tag
    PrimeraVez = True
    
    FrameDividVto.Visible = False
    
    CommitConexion
    
    Select Case Opcion
        Case 27
                    'CadenaDesdeOtroForm. Pipes
            '           1.- cadenaSQL numfac,numsere,fecfac
            '           2.- Numero vto
            '           3.- Importe maximo
            H = FrameDividVto.Height + 120
            W = FrameDividVto.Width
            FrameDividVto.Visible = True
            Me.Caption = "Dividir Vencimiento"
    End Select
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    i = Opcion
    If Opcion = 13 Or i = 43 Or i = 44 Then i = 11
    
    'Aseguradas
    Me.cmdCancelar(i).Cancel = True
    
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtCodigo(2).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgFecha_Click(Index As Integer)
    'Fecha de primer vencimiento
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtCodigo(2).Text <> "" Then frmF.Fecha = CDate(txtCodigo(2).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtCodigo(2)

End Sub

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtcodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub txtcodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim B As Boolean

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'nro de vtos
            PonerFormatoEntero txtCodigo(Index)
            
            If txtCodigo(0).Text <> "" Then
                txtCodigo(1).Text = Format(Round(ImporteSinFormato(ComprobarCero(txtCodigo(1).Text)) / txtCodigo(0), 2), "###,###,##0.00")
            End If
            
        Case 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 1 'Importe
            PonerFormatoDecimal txtCodigo(Index), 3
            
    End Select
End Sub

Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub
