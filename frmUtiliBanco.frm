VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtiliBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación Norma 43"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   Icon            =   "frmUtiliBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   8175
      Begin VB.CheckBox chkElimmFich 
         Caption         =   "Eliminar fichero "
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   540
         Width           =   7815
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar"
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   2
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CommandButton cmdImportar 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   6960
         TabIndex        =   1
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   525
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   840
         Picture         =   "frmUtiliBanco.frx":000C
         Top             =   240
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7080
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   6735
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton Command1 
         Caption         =   "&Volver"
         Height          =   375
         Index           =   1
         Left            =   10320
         TabIndex        =   8
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Integrar"
         Height          =   375
         Index           =   0
         Left            =   9120
         TabIndex        =   7
         Top             =   6240
         Width           =   1095
      End
      Begin VB.TextBox txtDatos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5955
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Text            =   "frmUtiliBanco.frx":685E
         Top             =   180
         Width           =   11235
      End
   End
End
Attribute VB_Name = "frmUtiliBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 313

Public opcion As Byte
    ' 0.- Cargar fichero de datos NORMA 43

Dim SQL As String
Dim NF As Integer
Dim FicheroPpal As String

'Para el procesamiento de datos desde el fichero
Dim Cta As String
Dim Saldo As Currency
Dim Importe As Currency
Dim Rs As ADODB.Recordset
Dim Cad As String

Private Sub cmdImportar_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then
        MsgBox "Debes poner el nombre de archivo?", vbExclamation
        Exit Sub
    End If
    If Dir(Text1.Text, vbArchive) = "" Then
        MsgBox "Fichero NO existe", vbExclamation
        Exit Sub
    End If
    'Borramos los temporales
    SQL = "Delete from Usuarios.wnorma43 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    Screen.MousePointer = vbHourglass
    If ProcesarFichero Then
        NumRegElim = 1
        'Ahora procesamos los datos
        ProcesarDatos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerModo(vModo As Byte)
    Select Case vModo
    Case 0
        'Primer frame
        Frame1.Visible = True
        Frame2.Visible = False
        Me.Width = Frame1.Width
        Me.Height = Frame1.Height
    Case 1
        Frame2.Visible = True
        Frame1.Visible = False
        Me.Width = Frame2.Width
        Me.Height = Frame2.Height
    End Select
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 150
    Me.Refresh
End Sub


Private Function ProcesarFichero() As Boolean
Dim Fin As Boolean
Dim Cad As String

On Error GoTo EProcesarFichero
    'Abrimos el fichero para lectura
    ProcesarFichero = False
    NF = FreeFile
    FicheroPpal = "|"
    Open Text1.Text For Input As #NF
    While Not EOF(NF)
        Line Input #NF, SQL
        If SQL <> "" Then
                                        'Separador de lineas
            FicheroPpal = FicheroPpal & SQL & "|"
        End If
    Wend
    Close #NF
    ProcesarFichero = True
    Exit Function
EProcesarFichero:
    MuestraError Err.Number
End Function


Private Sub ProcesarDatos()
Dim i As Long
Dim Cont As Long
Dim NF As Long
Dim Linea As String
Dim Fichero As String
Dim Primer23 As Boolean
Dim Num22 As Integer  'Para conrolar los asientos k se han realizado
Dim Ampliacion As String
Dim RegistroInsertado As Boolean
Dim Comienzo As Long   'Para cuando vienen varios bancos
Dim Fecha As String   'Fecha importacion datos

Dim ContadorMYSQL As Integer
Dim ContadorRegistrosBanco As Integer

    'Vemos cuantas cuentas trae el extracto
    i = 0
    Cont = 0
    Do
        NF = i + 1
        i = InStr(NF, FicheroPpal, "|11")  'los registros empiezan por 11 para las cuentas
        If i > 0 Then Cont = Cont + 1
    Loop Until i = 0
        
    If Cont = 0 Then
        MsgBox "Error en el fichero. No se ha encontrado registro 11", vbExclamation
        Exit Sub
    End If

    
    
    txtDatos.Text = ""
    Comienzo = 2
    ContadorMYSQL = 1
    ContadorRegistrosBanco = 0
    Cta = ""
    'Ya sabemos cuantas cont hay k tratar
    For i = 1 To Cont
        If i <> Cont Then
            Linea = "|11"
            'Hay mas de un |11 o cuenta bancaria
        Else
            'Una unica cta bancaria en este fichero
            Linea = "|88"
        End If
        
        NF = InStr(Comienzo, FicheroPpal, Linea)
        If NF = 0 Then
            MsgBox "imposible situar datos."
            Exit Sub
        End If
        
        Fichero = Mid(FicheroPpal, Comienzo, NF - 1)
        
        Comienzo = NF + 1
                
        'Fecha
        Fecha = ""
        Linea = Mid(Fichero, 31, 2) & "/" & Mid(Fichero, 29, 2) & "/" & Mid(Fichero, 27, 2)
        If IsDate(Linea) Then
            Fecha = "Fecha: " & Space(18) & Format(Linea, "dd/mm/yyyy")
        Else
            Fecha = "Fecha: " & Space(18) & "Error obteniendo fecha"
        End If
        Fecha = Fecha & vbCrLf
                
        'ANTES
        'NF = InStr(1, Fichero, "|22") 'Es el fin de la primera linea
        NF = InStr(1, Fichero, "|") 'Es el fin de la primera linea
        
        'Primara linea, la de la cuenta
        Linea = Mid(Fichero, 1, NF - 1) 'pq quitamos el pipe del principio y del final
        
        'De la primera linea obtenemos el numero de cuenta
        Ampliacion = Cta
        FijarCtaContable (Linea)
        If Ampliacion <> Cta Then
            If Ampliacion <> "" Then
                'HA CAMBIADO DE CUENTA DEEEENTRO DEL MISMO Fichero
                ContadorRegistrosBanco = 0
            End If
        End If
        
        If Cta = "" Then
            
            MsgBox "Error obteniendo la cuenta contable asociada. Linea: " & Linea, vbExclamation
            Exit Sub
        Else
            SQL = ""
            If ContadorRegistrosBanco = 0 Then
                If txtDatos.Text <> "" Then txtDatos.Text = txtDatos.Text & SQL & vbCrLf
                For NF = 1 To 98
                    SQL = SQL & "="
                Next NF
                txtDatos.Text = txtDatos.Text & SQL & vbCrLf
                SQL = Mid(Linea, 3, 4) & " " & Mid(Linea, 7, 4) & " ** " & Mid(Linea, 11, 10)
                txtDatos.Text = txtDatos.Text & "Cuenta bancaria: " & SQL & vbCrLf
                Fecha = Fecha & "Cuenta bancaria:   " & SQL & vbCrLf
                txtDatos.Text = txtDatos.Text & "Cuenta contable:   " & Cta & vbCrLf
                Fecha = Fecha & "Cuenta contable:    " & Cta & vbCrLf
                txtDatos.Text = txtDatos.Text & "Linea  F.Opercion   F.Valor         Debe            Haber          Concepto" & vbCrLf
                SQL = ""
                For NF = 1 To 98
                    SQL = SQL & "-"
                Next NF
                txtDatos.Text = txtDatos.Text & SQL & vbCrLf
            Else
                'Es otro trozo de fichero 11| pero de la misma cuenta
                txtDatos.Text = txtDatos.Text & String(98, "=") & vbCrLf
            End If
        End If
        
        'Fijaremos el saldo incial
        SQL = Mid(Linea, 34, 14)
        If Not IsNumeric(SQL) Then
            MsgBox "Error. Se esperaba un importe: " & SQL, vbExclamation
            Exit Sub
        End If
        Saldo = Val(SQL) / 100
        
        'ANTES 25 Noviembre
        'Se trabaja al reves
        'Signo del saldo
        'If Mid(LINEA, 33, 1) = "2" Then Saldo = Saldo * -1
        If Mid(Linea, 33, 1) = "1" Then Saldo = Saldo * -1
        
        NF = InStr(1, Fichero, "|") 'Es el fin de la primera linea
        Fichero = Mid(Fichero, NF + 1) '+1 y le quito el pipe
        
        RegistroInsertado = False
        Ampliacion = ""
        Num22 = 0
        'Ya tenemos los primeros datos. Ahora a por los apuntes
        Do
            NF = InStr(1, Fichero, "|")
            Linea = Mid(Fichero, 1, NF - 1)
            Fichero = Mid(Fichero, NF + 1)
            
            SQL = Mid(Linea, 1, 2)
          
            
            If SQL = "22" Then
                If Num22 > 0 Then
                    If Not RegistroInsertado Then
                        If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                    End If
                End If
            
                'Primera parte de la linea de apunte
                If Not ProcesaLineaASiento(Linea, Ampliacion) Then Exit Sub
                RegistroInsertado = False
                Primer23 = True
                Num22 = Num22 + 1
            Else
                If SQL = "23" Then
                    If Primer23 Then
                        Primer23 = False
                        'Insertaremos
                        Ampliacion = ProcesaAmpliacion2(Linea)
                        If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                        RegistroInsertado = True
                        'txtDatos.Text = txtDatos.Text & vbCrLf & vbCrLf
                    End If
                    
                    
                Else
                    If SQL = "33" Then
                        If Not RegistroInsertado Then
                            If Num22 > 0 Then
                                If Not InsertarRegistro(Ampliacion, ContadorMYSQL, ContadorRegistrosBanco) Then Exit Sub
                            End If
                        End If
                        'Fin CTA. Hacer comprobaciones
                        
                        If Not HacerComprobaciones(Linea, ContadorRegistrosBanco, ContadorMYSQL) Then
                            Exit Sub
                        End If
                        Fichero = ""
                       
                    Else
                        'Cualquier otro caso no esta tratado
                        Fichero = ""
                    End If
                End If
            End If
        Loop Until Fichero = ""
        'Kitamos de ppal el valor
    Next i
    
    'Si llega aqui es k ha ido bien.Si no inserta nada, NO muestro los datos
    If ContadorMYSQL > 1 Then PonerModo 1
End Sub






Private Sub FijarCtaContable(ByRef Lin As String)
    SQL = "Select codmacta from ctabancaria"
    SQL = SQL & " where Entidad = " & Mid(Lin, 3, 4)
    SQL = SQL & " AND oficina = " & Mid(Lin, 7, 4)
    SQL = SQL & " AND ctabanco = '" & Mid(Lin, 11, 10) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cta = ""
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Cta = Rs.Fields(0)
    End If
    Rs.Close
    Set Rs = Nothing
    If Cta = "" Then
        SQL = "Fichero pertenece a la cuenta bancaria:  " & Mid(Lin, 3, 4) & "  " & Mid(Lin, 7, 4) & " ** " & Mid(Lin, 11, 10) & vbCrLf
        SQL = SQL & vbCrLf & "No esta asociada a ninguna cuenta contable."
        MsgBox SQL, vbExclamation
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim SQ As String

    If Index = 1 Then
        PonerModo 0
        Exit Sub
    End If
    
    'Comprobaremos que hay datos para traspasar
    If txtDatos.Text = "" Then
        MsgBox "Datos vacios", vbExclamation
        Exit Sub
    End If
    
    'COntamos los saltos de linea
    NumRegElim = 1
    SQ = txtDatos.Text
    NF = 0
    Do
        NumRegElim = InStr(1, SQ, vbCrLf)
        If NumRegElim > 0 Then
            SQ = Mid(SQ, NumRegElim + 2)  'vbcrlf son DOS caracteres
            NF = NF + 1
            If NF > 5 Then NumRegElim = 0 'Hay mas lineas que las del encabezado
        End If
    Loop Until NumRegElim = 0
    'Fichero comprobacion de saldos
    If NF <= 5 Then
        txtDatos.Text = ""
        If chkElimmFich.Value = 1 Then
            If Dir(Text1.Text, vbArchive) <> "" Then Kill Text1.Text
        End If
        Exit Sub
    End If
    'Comprobamos que no existen datos entre las fechas
    Screen.MousePointer = vbHourglass
    SQ = ""
    Set Rs = New ADODB.Recordset
    SQL = "Select min(fecopera) from Usuarios.wnorma43 where codusu = " & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then SQ = " fecopera >='" & Format(Rs.Fields(0), FormatoFecha) & "'"
    End If
    Rs.Close
    SQL = "Select max(fecopera) from Usuarios.wnorma43 where codusu = " & vUsu.Codigo
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then SQ = SQ & " and fecopera <='" & Format(Rs.Fields(0), FormatoFecha) & "'"
    End If
    Rs.Close
    SQL = "Select count(*) from norma43 where " & SQ
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Not Rs.EOF Then
        NF = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Set Rs = Nothing
    
    If NF > 0 Then
        SQL = "Se han encontrado datos entre las fechas importadas." & vbCrLf
        SQL = SQL & "( " & SQ & " )" & vbCrLf & vbCrLf
        SQL = SQL & "Puede duplicar los datos. ¿ Desea continuar ? " & vbCrLf
        If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If MsgBox("¿Los datos serán importados. ¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
    End If
    
    'Haremos la insercion del registro del banco
    If BloqueoManual(True, "norma43", "clave") Then
        InsertarHcoBanco
        BloqueoManual False, "norma43", ""
        Command1(0).Enabled = False
    Else
        MsgBox "Tabla bloqueada por otro usuario.", vbExclamation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    PonerModo 0
    Top = 600
    Left = 600
End Sub

Private Sub Image1_Click()

    cd1.CancelError = False
    cd1.DialogTitle = "Archivo banco NORMA 43"
    cd1.ShowOpen
    If cd1.FileName <> "" Then Text1.Text = cd1.FileName
    
End Sub




'Metere en CadenaDesdeOtroForm, empipado
' Fecha operacion, fecha valor, importeDebe, importe haber, numdocum
Private Function ProcesaLineaASiento(ByRef Lin As String, vAmpliacion As String) As Boolean
Dim Debe As Boolean


    ProcesaLineaASiento = False
    CadenaDesdeOtroForm = ""
    'Fecha operacion
    Cad = Mid(Lin, 11, 6)
    Cad = "20" & Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5, 2)
    If Not IsDate(Cad) Then
        MsgBox "Formato fecha incorrecto", vbExclamation
        Exit Function
    End If
    CadenaDesdeOtroForm = Format(Cad, FormatoFecha) & "|"
    
    'Fecha valor
    Cad = Mid(Lin, 17, 6)
    Cad = "20" & Mid(Cad, 1, 2) & "/" & Mid(Cad, 3, 2) & "/" & Mid(Cad, 5, 2)
    If Not IsDate(Cad) Then
        MsgBox "Formato fecha incorrecto", vbExclamation
        Exit Function
    End If
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(Cad, FormatoFecha) & "|"
    
    
    'Importe
    Cad = Mid(Lin, 28, 1)
    Debe = Cad = "1"
    Cad = Mid(Lin, 29, 14)
    If Not IsNumeric(Cad) Then
        MsgBox "Importe registro 22 incorrecto: " & Cad, vbExclamation
        Exit Function
    End If
    Importe = Val(Cad) / 100
    Cad = TransformaComasPuntos(CStr(Importe))
    
    'Importe debe / haber
    If Debe Then
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Cad & "|NULL|"
    Else
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "NULL|" & Cad & "|"
    End If
    
    
    'Posible ampliacion
    If Len(Lin) > 53 Then
        vAmpliacion = Trim(Mid(Lin, 53))
        If Len(vAmpliacion) > 30 Then vAmpliacion = Mid(vAmpliacion, 1, 30)
    Else
        vAmpliacion = ""
    End If
    
  '  'Para el arrastrado
  '  'Esto va al reves de la contbiliad, ya k trabajamos con la cuenta del banoc
  '  'ANTES del 25 de Novi
    If Not Debe Then Importe = Importe * -1
  '  If Debe Then Importe = Importe * -1
    'Num docum
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Mid(Lin, 43, 10)
    ProcesaLineaASiento = True
End Function

Private Function ProcesaAmpliacion2(miLinea As String) As String
Dim CADENA As String
Dim C2 As String
Dim Blanco As Boolean
Dim i As Integer

    CADENA = ""
    Blanco = False
    For i = 5 To Len(miLinea)
        C2 = Mid(miLinea, i, 1)
        If C2 = " " Then
             If Not Blanco Then
                CADENA = CADENA & C2
                Blanco = True
            End If
        Else
            Blanco = False
            CADENA = CADENA & C2
        End If
    Next i
    If Len(CADENA) > 30 Then CADENA = Mid(CADENA, 1, 30)
    ProcesaAmpliacion2 = CADENA
End Function


Private Function InsertarRegistro(Ampliacion As String, ByRef ContadorMYSQL As Integer, ByRef ContadorRegistrosDeUnBanco As Integer) As Boolean
Dim vSQL As String
Dim L As String

    On Error GoTo EProcesaAmpliacion
    InsertarRegistro = False
        
    vSQL = "INSERT INTO Usuarios.wnorma43 (codusu,orden, codmacta, fecopera,"
    vSQL = vSQL & "fecvalor, importeD, importeH,  concepto,"
    vSQL = vSQL & "numdocum, saldo) VALUES (" & vUsu.Codigo & "," & ContadorMYSQL & ",'"
    'Numero de apunte
    txtDatos.Text = txtDatos.Text & Right("     " & NumRegElim, 5)
    'Fecha operacion
    L = RecuperaValor(CadenaDesdeOtroForm, 1)
    txtDatos.Text = txtDatos.Text & "  " & Format(L, "dd/mm/yyyy")
    vSQL = vSQL & Cta & "','" & L
    'Fc Valor
    L = RecuperaValor(CadenaDesdeOtroForm, 2)
    txtDatos.Text = txtDatos.Text & " " & Format(L, "dd/mm/yyyy")
    vSQL = vSQL & "','" & L
    'Importe DEBE/HABER
    vSQL = vSQL & "'," & RecuperaValor(CadenaDesdeOtroForm, 3)
    L = RecuperaValor(CadenaDesdeOtroForm, 3)
    NF = 0
    If L = "NULL" Then
        NF = 1
        L = RecuperaValor(CadenaDesdeOtroForm, 4)
    End If
    
    L = TransformaPuntosComas(L)
    L = Format(L, FormatoImporte)
    Cad = "              "
    If NF = 0 Then
        'Debe
        txtDatos.Text = txtDatos.Text & "  " & Right("              " & L, 14) & "    " & Cad
    Else
        txtDatos.Text = txtDatos.Text & "  " & Cad & "    " & Right("              " & L, 14)
    End If
    vSQL = vSQL & "," & RecuperaValor(CadenaDesdeOtroForm, 4)
    
    'El concepto lo saco de la linea de aqui
    Cad = DevNombreSQL(Trim(Ampliacion))  '30 como mucho
    vSQL = vSQL & ",'" & Cad & "',"
    txtDatos.Text = txtDatos.Text & "    " & Ampliacion & vbCrLf
        
    'NumDocum
    vSQL = vSQL & "'" & RecuperaValor(CadenaDesdeOtroForm, 5) & "'"
    Saldo = Saldo - Importe
    Cad = TransformaComasPuntos(CStr(Saldo))
    vSQL = vSQL & "," & Cad & ")"
    'Para la BD
    ContadorMYSQL = ContadorMYSQL + 1
    
    'Para comprobar los regisitros
    ContadorRegistrosDeUnBanco = ContadorRegistrosDeUnBanco + 1
    'El que habia.
    NumRegElim = NumRegElim + 1 'Contador mas uno
    Conn.Execute vSQL
    
    InsertarRegistro = True
    Exit Function
EProcesaAmpliacion:
    MuestraError Err.Number, Err.Description & vbCrLf & vSQL
       
End Function




Private Function HacerComprobaciones(ByRef Lin As String, ContadorRegistrosBanco As Integer, TotalRegistrosInsertados As Integer) As Boolean
Dim Ok As Boolean
Dim InsercionesActuales As Integer
    Set Rs = New ADODB.Recordset
    HacerComprobaciones = False
    InsercionesActuales = NumRegElim - 1
    Cad = "Select max(orden) from Usuarios.wnorma43 where codusu =" & vUsu.Codigo
    Cad = Cad & " AND codmacta ='" & Cta & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NF = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then NF = Rs.Fields(0)
    End If
    Rs.Close
    
    'Numero de lineas insertadas
    Ok = False
    'Total registros en BD
    If NF = ContadorRegistrosBanco Then
        'Coinciden los contadores de insercion parcial
        
        NF = Val(Mid(Lin, 21, 5)) + Val(Mid(Lin, 40, 5))
        If NF = NumRegElim - 1 Then Ok = True
    End If
    If Not Ok Then
        'Error en contadores de registros
        MsgBox "Error en contadores de registo", vbExclamation
        NumRegElim = 0
    End If
    
    
    
    If NumRegElim > 0 Then
        'Obtengo la suma de importes
        Cad = "Select sum(importeD)as debe,sum(importeH) as haber,sum(importeD)-sum(importeH) from Usuarios.wnorma43 where codusu = " & vUsu.Codigo
        Cad = Cad & " AND codmacta ='" & Cta & "'"
        'Enero 2009.
        'Estamos admitiendo ficheros que , aun siendo de la misma cuenta, tran mas de una entrada 11| (cabecera de cuenta
        NF = ContadorRegistrosBanco - InsercionesActuales
        Cad = Cad & " AND orden >" & NF
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If Not Rs.EOF Then
            Cad = CStr(Val(Mid(Lin, 26, 14)) / 100)
            CadenaDesdeOtroForm = DBLet(Rs.Fields(0), "N")
            Ok = (Cad = CadenaDesdeOtroForm)
            If Ok Then
                Cad = CStr(Val(Mid(Lin, 45, 14)) / 100)
                CadenaDesdeOtroForm = DBLet(Rs.Fields(1), "N")
                Ok = (Cad = CadenaDesdeOtroForm)
            End If
            If Ok Then
                Importe = Val(Mid(Lin, 60, 14)) / 100
                If Mid(Lin, 59, 1) = "2" Then Importe = Importe * -1
                
                If ContadorRegistrosBanco = 0 Then
                    Cad = "Fichero de comprobación de saldos: " & vbCrLf & vbCrLf
                    Cad = Cad & "Saldo: " & CStr(Importe)
                    Cad = Cad & vbCrLf & vbCrLf & vbCrLf
                    Cad = Cad & "¿Desea eliminar el archivo?"
                    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                        If Dir(Text1.Text, vbArchive) <> "" Then
                            Kill Text1.Text
                            Text1.Text = ""
                        End If
                    End If
                End If
                
            End If
        End If
        Rs.Close
        If Ok Then
            NumRegElim = 1
        Else
            NumRegElim = 0
        End If
    End If
    
    'Si llegamos aqui y numregelim>0 esta bien
    If NumRegElim > 0 Then HacerComprobaciones = True
    Set Rs = Nothing
    
End Function


Private Sub InsertarHcoBanco()
Dim Codigo As Long
    
    Set Rs = New ADODB.Recordset
    Codigo = 0
    SQL = "Select max(codigo) from norma43"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        Codigo = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close
    Codigo = Codigo + 1
    
    SQL = "Select * from Usuarios.wnorma43 where codusu = " & vUsu.Codigo & " ORDER By Orden"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'Cadena de insercion
    SQL = "INSERT INTO norma43 (codigo, codmacta, fecopera, fecvalor, importeD,"
    SQL = SQL & "importeH, concepto, numdocum, saldo, punteada) VALUES ("
    While Not Rs.EOF
        Cad = Codigo & ",'" & Rs!codmacta & "','" & Format(Rs!fecopera, FormatoFecha)
        Cad = Cad & "','" & Format(Rs!fecvalor, FormatoFecha) & "',"
        If IsNull(Rs!ImporteD) Then
            Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Rs!ImporteH))
        Else
            Cad = Cad & TransformaComasPuntos(CStr(Rs!ImporteD)) & ",NULL"
        End If
        Cad = Cad & ",'" & DevNombreSQL(DBLet(Rs!Concepto)) & "','" & Rs!numdocum & "',"
        Cad = Cad & TransformaComasPuntos(CStr(Rs!Saldo)) & ",0);"
        Cad = SQL & Cad
        'Ejecutamos SQL
        Conn.Execute Cad
        Codigo = Codigo + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
    'Ahora deberiamos eliminar el archivo
    If chkElimmFich.Value = 1 Then
        If Dir(Text1.Text, vbArchive) <> "" Then Kill Text1.Text
         MsgBox "Importación finalizada", vbInformation
    Else
        MsgBox "Proceso finalizado. El fichero NO será eliminado", vbExclamation
    End If
End Sub
