VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESCobrosDivVto 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmTESCobrosDivVto.frx":0000
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
         Picture         =   "frmTESCobrosDivVto.frx":000C
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
         Top             =   1770
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
Attribute VB_Name = "frmTESCobrosDivVto"
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

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim cad As String
Dim CONT As Long
Dim I As Integer
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
    If txtcodigo(1).Text <> "" Then vImpvto = ImporteSinFormato(ComprobarCero(txtcodigo(1).Text))
    If txtcodigo(0).Text <> "" Then vVtos = CInt(ComprobarCero(txtcodigo(0).Text))
        
    If vImpvto = 0 And vVtos = 0 Then
        MsgBox "Debe introducir el importe o el nro de vencimientos o ambos. Revise.", vbExclamation
        PonFoco txtcodigo(0)
        Exit Sub
    End If
    
    ' debe introducir la fecha del primer vto, viene cargada
    If txtcodigo(2).Text = "" Then
        MsgBox "Debe introducir la fecha del primer vencimiento", vbExclamation
        PonFoco txtcodigo(2)
        Exit Sub
    End If
    
    
    If txtcodigo(3).Text = "" Then
        If MsgBox("No ha puesto valor en el campo de días de resto de vencimientos. " & vbCrLf & vbCrLf & "¿ Desea continuar ?" & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    
    
    ' me ponen importe
    If vImpvto <> 0 Then
        If Importe < vImpvto Then
            MsgBox "El importe del vencimiento es inferior del importe a dividir. Revise", vbExclamation
            PonFoco txtcodigo(0)
            Exit Sub
        End If
        ' me ponen nro de vtos
        If vVtos <> 0 Then
            If Importe - Round(vImpvto * (vVtos - 1), 2) < 0 Then
                MsgBox "Es imposible dividir el vencimiento en " & vVtos & " vencimientos de " & Format(vImpvto, "###,###,##0.00") & " euros.", vbExclamation
                PonFoco txtcodigo(0)
                Exit Sub
            End If
            If vVtos = 1 And vImpvto <> Importe Then
                MsgBox "No podemos dejar el vencimiento con menos importe del original. Revise.", vbExclamation
                PonFoco txtcodigo(0)
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
    
    SQL = ""
    If SQL = "" Then
        Set RS = New ADODB.Recordset
        
        
        'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        I = -1
        RC = "Select max(numorden) from cobros WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            SQL = "Error. Vencimiento NO encontrado: " & CadenaDesdeOtroForm
        Else
            I = RS.Fields(0) '+ 1
        End If
        RS.Close
        Set RS = Nothing
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        PonFoco txtcodigo(1)
        Exit Sub
        
    Else
        SQL = "¿Desea desdoblar el vencimiento en los indicados?" 'uno de : " & Im & " euros?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    Dias = txtcodigo(3).Text

    
    FecVenci = CDate(txtcodigo(2))
    vFecVenci = FecVenci
    'OK.  a desdoblar
    vTotal = 0
    k = I + 1
    For J = 1 To vVtos - 1
    
        vTotal = vTotal + vImpvto
    
        vFecVenci = DateAdd("d", DBLet(Dias, "N"), vFecVenci)
        
    
        SQL = "INSERT INTO cobros (`numorden`,`gastos`,impvenci,`fecultco`,`impcobro`,`recedocu`,"
        SQL = SQL & "`tiporem`,`codrem`,`anyorem`,`siturem`,"
        SQL = SQL & "`numserie`,`numfactu`,`fecfactu`,`codmacta`,`codforpa`,`fecvenci`,`ctabanc1`,"
        SQL = SQL & "`text33csb`,`text41csb`,`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,"
        SQL = SQL & "`noremesar`,`observa`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,`codpais`,`nifclien`,iban, codusu) "
        'Valores
        SQL = SQL & " SELECT " & k & ",NULL," & TransformaComasPuntos(CStr(vImpvto)) & ",NULL,NULL,0,"
        SQL = SQL & "NULL,NULL,NULL,NULL,"
        SQL = SQL & "`numserie`,`numfactu`,`fecfactu`,`codmacta`,`codforpa`,"
        SQL = SQL & DBSet(vFecVenci, "F") & ","
        SQL = SQL & "`ctabanc1`,`text33csb`,`text41csb`,"
        'text83csb`,
        SQL = SQL & "`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,`noremesar`,`observa`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,`codpais`,`nifclien`,iban "
        SQL = SQL & "," & DBSet(vUsu.Id, "N")
        SQL = SQL & " FROM "
        SQL = SQL & " cobros WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        SQL = SQL & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
    
        Conn.Execute SQL
    
        k = k + 1
    
    Next J
    
    
    ' actualizamos el primer vencimiento
    vTotal = vTotal + vImpvto
        
    SQL = "update cobros set impvenci = coalesce(impcobro,0) + " & DBSet(vImpvto, "N")
    SQL = SQL & ", fecvenci = " & DBSet(FecVenci, "F")
    
    SQL = SQL & " WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
    SQL = SQL & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
    
    Conn.Execute SQL
    
    ' en el ultimo dejamos la diferencia
    If vTotal <> Importe Then
        SQL = "update cobros set impvenci = impvenci + " & DBSet(Importe - vTotal, "N")
        
        SQL = SQL & " WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        SQL = SQL & " AND numorden = " & DBSet(k - 1, "N")
        
        Conn.Execute SQL
    End If
    
    'Insertamos el LOG
    ParaElLog = "Dividir Vto.Fra.: " & Me.Label4(57).Caption & vbCrLf
    ParaElLog = ParaElLog & "Cliente         : " & Me.Label4(56).Caption & vbCrLf
    ParaElLog = ParaElLog & "Nro.Vencimientos: " & txtcodigo(0).Text & vbCrLf
    ParaElLog = ParaElLog & "Importe Vto     : " & txtcodigo(1).Text & vbCrLf
    ParaElLog = ParaElLog & "Fecha primer Vto: " & txtcodigo(2).Text & vbCrLf
    ParaElLog = ParaElLog & "Día Resto Vtos  : " & txtcodigo(3).Text & vbCrLf
    
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
    Me.Icon = frmPpal.Icon
    
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
    
    I = Opcion
    If Opcion = 13 Or I = 43 Or I = 44 Then I = 11
    
    'Aseguradas
    Me.cmdCancelar(I).Cancel = True
    
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtcodigo(2).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgFecha_Click(Index As Integer)
    'Fecha de primer vencimiento
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtcodigo(2).Text <> "" Then frmF.Fecha = CDate(txtcodigo(2).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtcodigo(2)

End Sub

Private Sub txtcodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
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
Dim cad As String, cadTipo As String 'tipo cliente
Dim B As Boolean

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'nro de vtos
            PonerFormatoEntero txtcodigo(Index)
            
            If txtcodigo(0).Text <> "" Then
                txtcodigo(1).Text = Format(Round(ImporteSinFormato(ComprobarCero(txtcodigo(1).Text)) / txtcodigo(0), 2), "###,###,##0.00")
            End If
            
        Case 2 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 1 'Importe
            PonerFormatoDecimal txtcodigo(Index), 3
            
    End Select
End Sub

Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub
