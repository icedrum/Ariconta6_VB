VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBalancesCopy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmBalancesCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelarAccion 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   4590
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCopyBalan 
      ForeColor       =   &H00800000&
      Height          =   4065
      Left            =   -30
      TabIndex        =   1
      Top             =   -60
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chkCopyBalan 
         Caption         =   "Copiar las cuentas "
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
         Left            =   720
         TabIndex        =   11
         Top             =   3450
         Width           =   2175
      End
      Begin VB.CommandButton cmdCopyBalan 
         Caption         =   "&Aceptar"
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
         Left            =   3600
         TabIndex        =   10
         Top             =   3330
         Width           =   975
      End
      Begin VB.CommandButton cmdCanListExtr 
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
         Index           =   57
         Left            =   4800
         TabIndex        =   8
         Top             =   3330
         Width           =   975
      End
      Begin VB.TextBox TextDescBalance 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2490
         Width           =   4065
      End
      Begin VB.TextBox txtNumBal 
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
         Left            =   720
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2490
         Width           =   885
      End
      Begin VB.TextBox TextDescBalance 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1410
         Width           =   4065
      End
      Begin VB.TextBox txtNumBal 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   2
         Left            =   720
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1410
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Origen"
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
         Height          =   240
         Index           =   117
         Left            =   120
         TabIndex        =   9
         Top             =   2130
         Width           =   720
      End
      Begin VB.Image ImgNumBal 
         Height          =   240
         Index           =   3
         Left            =   360
         Top             =   2490
         Width           =   240
      End
      Begin VB.Label Label17 
         Caption         =   "Copiar balances configurables"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   330
         Width           =   4875
      End
      Begin VB.Image ImgNumBal 
         Height          =   240
         Index           =   2
         Left            =   360
         Top             =   1410
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DESTINO"
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
         Height          =   240
         Index           =   116
         Left            =   120
         TabIndex        =   4
         Top             =   1050
         Width           =   945
      End
   End
   Begin VB.Menu mnP1 
      Caption         =   "p1"
      Visible         =   0   'False
      Begin VB.Menu mnPrueba 
         Caption         =   "Prueba F1"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmBalancesCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    
    ' 57 .- Copiar balance configurables
    
Public EjerciciosCerrados As Boolean
    'En algunos informes me servira para utilizar unas tablas u otras
Public Legalizacion As String   'Datos para la legalizacion
    
    
Public ConAsiento As Boolean
    
    
Dim Tablas As String
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBal As frmBasico
Attribute frmBal.VB_VarHelpID = -1


Dim SQL As String
Dim RC As String
Dim Rs As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim Cont As Long
Dim i As Integer

Dim Importe As Currency

'Para los balcenes frameBalance
' Cuando este trbajando con cerrado
' Para poder sbaer cuando empezaba el año del ejercicio a listar
Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date


Dim HanPulsadoSalir As Boolean

'Para cancelar
Dim PulsadoCancelar As Boolean


Private Sub chkCopyBalan_KeyPress(KeyAscii As Integer)
    ListadoKEYpress KeyAscii
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub cmdCanListExtr_Click(Index As Integer)
    If Me.cmdCancelarAccion.Visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub

Private Sub cmdCopyBalan_Click()
    If txtNumBal(3).Text = "" Then
        MsgBox "Seleccione el balance origen", vbExclamation
        Exit Sub
    End If
    
    
    Cad = "Va a copiar los datos del balance: " & vbCrLf & vbCrLf
    Cad = Cad & txtNumBal(3).Text & " - " & Me.TextDescBalance(3).Text & vbCrLf
    Cad = Cad & " sobre " & vbCrLf
    Cad = Cad & txtNumBal(2).Text & " - " & Me.TextDescBalance(2).Text & vbCrLf
    Cad = Cad & vbCrLf & vbCrLf & "Los datos del balance destino seran eliminados"
    Cad = Cad & vbCrLf & vbCrLf & "¿Desea continuar?"
    If MsgBox(Cad, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    
    SQL = "aparece"
    Cad = DevuelveDesdeBD("perdidas", "balances", "numbalan", txtNumBal(3).Text, "N", SQL)
    If Cad = "" Then
        MsgBox "Error leyendo datos: " & txtNumBal(3).Text
        Exit Sub
    End If
    
    Cad = "UPDATE balances SET perdidas=" & Cad & ",Aparece= " & SQL & " WHERE numbalan=" & txtNumBal(2).Text
    Conn.Execute Cad
    
    Cad = "DELETE FROM balances_ctas WHERE numbalan=" & txtNumBal(2).Text
    Conn.Execute Cad
    Cad = "DELETE FROM balances_texto WHERE numbalan=" & txtNumBal(2).Text
    Conn.Execute Cad
    Cad = "INSERT INTO balances_texto (NumBalan, Pasivo, codigo, padre, Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita, A_Cero, Pintar, LibroCD)"
    Cad = Cad & " SELECT " & txtNumBal(2).Text & ", Pasivo, codigo, padre, Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita, A_Cero, Pintar, LibroCD FROM"
    Cad = Cad & " balances_texto WHERE numbalan = " & txtNumBal(3).Text
    Conn.Execute Cad
    
    
    If Me.chkCopyBalan.Value = 1 Then
        'COpio los datos tb
       ' NumBalan, Pasivo, codigo, codmacta, tipsaldo, Resta
        Cad = "INSERT INTO balances_ctas ( NumBalan, Pasivo, codigo, codmacta, tipsaldo, Resta)"
        Cad = Cad & " SELECT " & txtNumBal(2).Text & ", Pasivo, codigo, codmacta, tipsaldo, Resta FROM"
        Cad = Cad & " balances_ctas WHERE numbalan = " & txtNumBal(3).Text
        Conn.Execute Cad
    End If
    Unload Me
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


    




Private Sub Form_Activate()

    

    If PrimeraVez Then
        PrimeraVez = False
        CommitConexion
        'Ponemos el foco
    End If
        Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Single
Dim W As Single

    Me.Icon = frmPpal.Icon

    Screen.MousePointer = vbHourglass
    PrimeraVez = True
    Limpiar Me
    
    
    For i = 2 To 3
        Me.ImgNumBal(i).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next i
    
    
    Select Case Opcion
    Case 57
        txtNumBal(2).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
        TextDescBalance(2).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
        H = FrameCopyBalan.Height
        W = FrameCopyBalan.Width
        FrameCopyBalan.Visible = True
    End Select
    HanPulsadoSalir = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    i = Opcion
    If Opcion = 23 Then i = 22
    If Opcion = 26 Or Opcion = 27 Or Opcion = 39 Or Opcion = 40 Then i = 25
    If Opcion = 51 Then i = 50
    If Opcion = 41 Then i = 5
    If Opcion = 52 Then i = 13
    If Opcion = 53 Then i = 8
    If Opcion = 56 Then i = 55
    
    'Legalizacion
    HanPulsadoSalir = True
    
    If Opcion < 32 Or Opcion > 38 Then
    
        Me.cmdCanListExtr(i).Cancel = True
        
        'Ajustaremos el boton para cancelar algunos de los listados k mas puedan costar
        AjustaBotonCancelarAccion
        cmdCancelarAccion.Visible = False
        cmdCancelarAccion.ZOrder 0
    
    End If
    Me.Width = W + 240
    Me.Height = H + 400
    
    'Añadimos ejercicios cerrados
    If EjerciciosCerrados Then Caption = Caption & "    EJERC. TRASPASADOS"
End Sub

Private Sub AjustaBotonCancelarAccion()
On Error GoTo EAj
    Me.cmdCancelarAccion.Top = cmdCanListExtr(i).Top
    Me.cmdCancelarAccion.Left = cmdCanListExtr(i).Left + 60
    cmdCancelarAccion.Width = cmdCanListExtr(i).Width
    cmdCancelarAccion.Height = cmdCanListExtr(i).Height + 30
    Exit Sub
EAj:
    MuestraError Err.Number, "Ajuste BOTON cancelar"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not HanPulsadoSalir Then Cancel = 1
    Legalizacion = ""
End Sub




Private Sub frmBal_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtNumBal(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    TextDescBalance(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub




Private Sub ImgNumBal_Click(Index As Integer)
    Screen.MousePointer = vbHourglass

    RC = Index
    
    Set frmBal = New frmBasico
    AyudaBalances frmBal
    Set frmBal = Nothing

    Screen.MousePointer = vbDefault
End Sub






Private Sub mnPrueba_Click()
    MsgBox "prueab"
End Sub



Private Sub txtNumBal_GotFocus(Index As Integer)
    PonFoco txtNumBal(Index)
End Sub


'++
Private Sub txtNumBal_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYBusqueda2 KeyAscii, 0
            Case 1:  KEYBusqueda2 KeyAscii, 1
        End Select
    Else
        ListadoKEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda2(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    ImgNumBal_Click (indice)
End Sub
'++






Private Sub txtNumBal_LostFocus(Index As Integer)
    SQL = ""
    With txtNumBal(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "Numero de balance debe de ser numérico: " & .Text, vbExclamation
                .Text = ""
            Else
                SQL = DevuelveDesdeBD("nombalan", "balances", "numbalan", .Text)
                If SQL = "" Then
                    MsgBox "El balance " & .Text & " NO existe", vbExclamation
                    .Text = ""
                End If
            End If
        End If
    End With
    TextDescBalance(Index).Text = SQL
End Sub


Private Sub PonerFoco(ByRef T As Object)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Sub ListadoKEYpress(ByRef KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


